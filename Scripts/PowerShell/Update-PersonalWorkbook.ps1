param(
    [string]$SourcePath = (Join-Path $PSScriptRoot '..\..\MacrosExcel\PERSONAL')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$vbextCtStdModule = 1
$vbextCtClassModule = 2
$vbextCtDocument = 100

function Get-ComponentName {
    param([System.IO.FileInfo]$File)

    if ($File.Name -eq 'ThisWorkbook.vba') {
        return 'ThisWorkbook'
    }
    $attribute = Select-String -LiteralPath $File.FullName -Pattern '^Attribute VB_Name = "([^"]+)"' |
        Select-Object -First 1
    if ($attribute) {
        return $attribute.Matches[0].Groups[1].Value
    }
    return [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
}

function Get-ImportText {
    param([System.IO.FileInfo]$File)

    $lines = Get-Content -LiteralPath $File.FullName
    $lines | Where-Object {
        $_ -notmatch '^VERSION ' -and $_ -notmatch '^Attribute VB_'
    }
}

function Set-ComponentCode {
    param(
        [object]$Component,
        [string[]]$Lines
    )

    $codeModule = $Component.CodeModule
    if ($codeModule.CountOfLines -gt 0) {
        $codeModule.DeleteLines(1, $codeModule.CountOfLines)
    }
    if ($Lines.Count -gt 0) {
        $codeModule.AddFromString(($Lines -join [Environment]::NewLine))
    }
}

if (-not (Test-Path -LiteralPath $SourcePath -PathType Container)) {
    throw "PERSONAL source folder was not found: $SourcePath"
}

$excel = $null
try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
}
catch {
    throw 'Open Excel and load PERSONAL.XLSB before running this script.'
}

$personalWorkbook = @($excel.Workbooks | Where-Object {
    $_.Name -ieq 'PERSONAL.XLSB'
}) | Select-Object -First 1
if ($null -eq $personalWorkbook) {
    throw 'PERSONAL.XLSB is not loaded in the active Excel instance.'
}

$sourceFiles = Get-ChildItem -LiteralPath $SourcePath -File |
    Where-Object { $_.Extension -in '.vba', '.cls' } |
    Sort-Object Name
if ($sourceFiles.Count -eq 0) {
    throw "No VBA source files were found: $SourcePath"
}

$updatedCount = 0
foreach ($sourceFile in $sourceFiles) {
    $componentName = Get-ComponentName $sourceFile
    $component = $null
    try {
        $component = $personalWorkbook.VBProject.VBComponents.Item($componentName)
    }
    catch {
        $component = $null
    }

    $sourceLines = @(Get-ImportText $sourceFile)
    if ($null -ne $component) {
        Set-ComponentCode $component $sourceLines
    }
    else {
        if ($sourceFile.Name -eq 'ThisWorkbook.vba') {
            throw 'ThisWorkbook source exists, but PERSONAL.XLSB has no ThisWorkbook component.'
        }
        $componentType = if ($sourceFile.Name.StartsWith('cls_')) {
            $vbextCtClassModule
        }
        else {
            $vbextCtStdModule
        }
        $component = $personalWorkbook.VBProject.VBComponents.Add($componentType)
        $component.Name = $componentName
        Set-ComponentCode $component $sourceLines
    }
    $updatedCount++
}

$personalWorkbook.Save()
Write-Host "Updated PERSONAL.XLSB components: $updatedCount"