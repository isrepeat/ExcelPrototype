Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Join-Path ([IO.Path]::GetTempPath()) ('WorkbookModules-test-' + [Guid]::NewGuid().ToString('N'))
$source = Join-Path $root 'vba'
[IO.Directory]::CreateDirectory((Join-Path $source '[0] classes')) | Out-Null
$utf8 = New-Object Text.UTF8Encoding($true)
$updater = Join-Path $PSScriptRoot 'Update-WorkbookModules.ps1'
function Write-Fixture([string]$Name, [string]$Text) {
    [IO.File]::WriteAllText((Join-Path $source $Name), $Text.TrimEnd(), $utf8)
}
function Assert-Equal($Expected, $Actual, [string]$Label) {
    if ($Expected -cne $Actual) { throw "${Label}: expected '$Expected', got '$Actual'" }
}
Write-Fixture 'modules.json' '{"Fixture":["Main.vba","[[]0] classes/*.cls.vba","Native.bas","ThisWorkbook.vba"]}'
Write-Fixture 'Main.vba' @'
Option Explicit
Public Function ReadText() As String
    Dim sample As SampleClass
    Set sample = New SampleClass
    ReadText = "Привет " & sample.Value & Native.Value
End Function
'@
Write-Fixture '[0] classes/SampleClass.cls.vba' @'
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "SampleClass"
Option Explicit
Public Property Get Value() As String
    Value = "class"
End Property
'@
Write-Fixture 'Native.bas' @'
Attribute VB_Name = "Native"
Option Explicit
Public Function Value() As String
    Value = " native"
End Function
'@
Write-Fixture 'ThisWorkbook.vba' 'Option Explicit'
$nativePath = Join-Path $source 'Native.bas'
[IO.File]::WriteAllText($nativePath, [IO.File]::ReadAllText($nativePath), [Text.Encoding]::ASCII)
$formExcel = New-Object -ComObject Excel.Application
$formBook = $null
try {
    $formExcel.Visible = $false
    $formExcel.DisplayAlerts = $false
    $formExcel.EnableEvents = $false
    $formBook = $formExcel.Workbooks.Add()
    $form = $formBook.VBProject.VBComponents.Add(3)
    $form.Name = 'FixtureForm'
    $form.Export((Join-Path $source 'FixtureForm.frm'))
}
finally {
    if ($null -ne $formBook) { $formBook.Close($false) }
    $formExcel.Quit()
}
Write-Fixture 'modules.json' '{"Fixture":["Main.vba","[[]0] classes/*.cls.vba","Native.bas","FixtureForm.frm","ThisWorkbook.vba"]}'
$bookPath = Join-Path $root 'Fixture.xlsm'
& $updater -Create -WorkbookPath $bookPath -VbaFolderPath $source -Profile Fixture
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $book = $excel.Workbooks.Open($bookPath)
    Assert-Equal 'Привет class native' ($excel.Run("'Fixture.xlsm'!Main.ReadText")) 'Unicode, class and native import'
    Assert-Equal 3 $book.VBProject.VBComponents.Item('FixtureForm').Type 'Native form import'
    $book.Worksheets.Item(1).Name = 'Input'
    $book.VBProject.VBComponents.Item($book.Worksheets.Item(1).CodeName).CodeModule.AddFromString('Public Sub OldEvent(): End Sub')
    $obsolete = $book.VBProject.VBComponents.Add(1)
    $obsolete.Name = 'Obsolete'
    $book.Save()
    $book.Close($false)
    $book = $null
    $excel.Quit()
    $excel = $null

    Write-Fixture 'ws_Input.vba' 'Public Function NewEvent() As Boolean: NewEvent = True: End Function'
    Write-Fixture 'modules.json' '{"Fixture":["Main.vba","[[]0] classes/*.cls.vba","Native.bas","FixtureForm.frm","ThisWorkbook.vba","ws_Input.vba"]}'
    & $updater -WorkbookPath $bookPath -VbaFolderPath $source -Profile Fixture
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $book = $excel.Workbooks.Open($bookPath)
    $names = @($book.VBProject.VBComponents | ForEach-Object { $_.Name })
    if ($names -contains 'Obsolete') { throw 'Obsolete module survived replacement.' }
    $code = $book.VBProject.VBComponents.Item($book.Worksheets.Item('Input').CodeName).CodeModule
    if ($code.Lines(1, $code.CountOfLines) -notmatch 'NewEvent' -or $code.Lines(1, $code.CountOfLines) -match 'OldEvent') {
        throw 'Document module was not replaced by CodeName.'
    }
    $book.Close($false)
    $book = $null
    $excel.Quit()
    $excel = $null
    $originalHash = (Get-FileHash -LiteralPath $bookPath).Hash
    Write-Fixture 'modules.json' '{"Fixture":["missing.vba"]}'
    $rejected = $false
    try { & $updater -WorkbookPath $bookPath -VbaFolderPath $source -Profile Fixture }
    catch { $rejected = $true }
    Assert-Equal $true $rejected 'Missing source rejected'
    Assert-Equal $originalHash (Get-FileHash -LiteralPath $bookPath).Hash 'Preflight did not alter workbook'
    if (@(Get-ChildItem -LiteralPath (Join-Path $root '.backup') -Directory -Filter 'modules-*').Count -ne 1) { throw 'Backup was not created.' }
    # Verify that vba is resolved relative to the script regardless of the current directory.
    Copy-Item -LiteralPath $updater -Destination $root
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'WorkbookModules.ps1') -Destination $root
    Write-Fixture 'modules.json' '{"Fixture":["Main.vba"]}'
    $fallbackPlan = @(& (Join-Path $root 'Update-WorkbookModules.ps1') -Profile Fixture -PlanOnly)
    Assert-Equal 'Main' $fallbackPlan[0].Name 'Adjacent vba fallback'
    & $updater -WorkbookPath $bookPath -Mode Clear
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $book = $excel.Workbooks.Open($bookPath)
    foreach ($component in $book.VBProject.VBComponents) {
        if ($component.Type -ne 100 -or $component.CodeModule.CountOfLines -ne 0) { throw 'Clear left VBA code behind.' }
    }
    Write-Output "PASS: schema masks, Unicode, class headers, native modules/forms, document mapping, replacement, backup, preflight, adjacent vba, clear. Fixtures: $root"
}
finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
}