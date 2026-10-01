param(
    [Parameter(Mandatory = $true)][string]$WorkbookPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$fixtureRoot = Join-Path ([IO.Path]::GetTempPath()) ('WorkbookSources-test-' + [Guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($fixtureRoot) | Out-Null
$fixturePath = Join-Path $fixtureRoot 'SourceContract.xlsm'
[IO.File]::Copy([IO.Path]::GetFullPath($WorkbookPath), $fixturePath)
$sourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\vba'))
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($fixturePath)
    for ($index = $book.VBProject.VBComponents.Count; $index -ge 1; $index--) {
        $component = $book.VBProject.VBComponents.Item($index)
        if ($component.Type -eq 100) {
            if ($component.CodeModule.CountOfLines -gt 0) { $component.CodeModule.DeleteLines(1, $component.CodeModule.CountOfLines) }
        }
        else { $book.VBProject.VBComponents.Remove($component) }
    }
    foreach ($file in Get-ChildItem -LiteralPath $sourceRoot -Recurse -Filter '*.vba') {
        $source = [IO.File]::ReadAllText($file.FullName, [Text.Encoding]::UTF8)
        $nameMatch = [regex]::Match($source, '(?m)^Attribute VB_Name = "([^"]+)"')
        $name = if ($nameMatch.Success) { $nameMatch.Groups[1].Value } else { $file.BaseName -replace '\.cls$', '' }
        $lines = New-Object 'System.Collections.Generic.List[string]'
        $insideHeader = $false
        foreach ($line in ($source -split '\r?\n')) {
            if ($line.Trim() -eq 'VERSION 1.0 CLASS') { $insideHeader = $true; continue }
            if ($insideHeader) {
                if ($line.Trim() -eq 'END') { $insideHeader = $false }
                continue
            }
            if ($line.Trim().StartsWith('Attribute ')) { continue }
            $lines.Add($line)
        }
        if ($file.Name -eq 'ThisWorkbook.vba') {
            $component = $book.VBProject.VBComponents.Item($book.CodeName)
        }
        else {
            $type = if ($file.Name.EndsWith('.cls.vba')) { 2 } else { 1 }
            $component = $book.VBProject.VBComponents.Add($type)
            $component.Name = $name
        }
        $component.CodeModule.AddFromString(($lines -join [Environment]::NewLine))
    }
    $probe = $book.VBProject.VBComponents.Add(1)
    $probe.Name = 'ex_ContractProbe'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Public Function Run() As Boolean
    Dim context As Object
    If Not ex_RuntimeLifecycle.fn_TryEnter(context) Then Exit Function
    ex_RuntimeLifecycle.fn_Leave context
    If CLng(context("ActiveCalls")) <> 0 Then Exit Function
    context("StopRequested") = True
    context("Phase") = "Requested"
    Run = ex_RuntimeLifecycle.fn_PrepareReload()
End Function
'@)
    $book.Save()
    $book.Close($false)
    $book = $null
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($fixturePath)
    if (-not $excel.Run("'SourceContract.xlsm'!ex_ContractProbe.Run")) {
        throw 'The real project lifecycle contract failed.'
    }
    Write-Output 'PASS: real project sources imported into a copy; entry accounting and shutdown contract execute successfully.'
    Write-Output "Fixture preserved: $fixtureRoot"
}
finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
}