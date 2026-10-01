Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'WorkbookModules.ps1')
$root = Join-Path ([IO.Path]::GetTempPath()) ('SourceEncoding-' + [Guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($root) | Out-Null
$path = Join-Path $root 'Source.vba'
$utf8 = New-Object Text.UTF8Encoding($false, $true)
function Assert-Rejected([scriptblock]$Action) {
    $rejected = $false
    try { & $Action | Out-Null } catch { $rejected = $true }
    if (-not $rejected) { throw 'Invalid source was accepted.' }
}
$sample = 'Русский текст; Українська: ї є ґ; 😀'
[IO.File]::WriteAllText($path, $sample, $utf8)
if ((Read-WorkbookSourceText $path) -cne $sample) { throw 'UTF-8 round trip failed.' }
Assert-Rejected { Read-WorkbookSourceText $path $true }
foreach ($bytes in @(@(0xC0, 0xAF), @(0xED, 0xA0, 0x80), @(0xF4, 0x90, 0x80, 0x80), @(0xE2, 0x82))) {
    [IO.File]::WriteAllBytes($path, [byte[]]$bytes)
    Assert-Rejected { Read-WorkbookSourceText $path }
}
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $component = $book.VBProject.VBComponents.Add(1)
    $component.Name = 'UnicodeProbe'
    $source = 'Public Function ReadText() As String' + "`r`n" + 'ReadText = "' + $sample + '"' + "`r`n" + 'End Function' + "`r`n" + "' English comment"
    $source += "`r`n" + 'Public Function ErrorText() As String' + "`r`n" + 'On Error GoTo EH' + "`r`n" + 'Err.Raise vbObjectError + 1, , "' + $sample + '"' + "`r`n" + 'EH: ErrorText = Err.Description' + "`r`n" + 'End Function'
    $comment = "' Кириллический комментарий"
    if ((ConvertTo-WorkbookModuleCode $comment) -cne $comment) { throw 'Comment was modified.' }
    $code = ConvertTo-WorkbookModuleCode $source
    if ($code -match '[^\x00-\x7F]') { throw 'Prepared code is not ASCII.' }
    $component.CodeModule.AddFromString($code)
    if ($excel.Run("'" + $book.Name + "'!UnicodeProbe.ReadText") -cne $sample) { throw 'VBA Unicode execution failed.' }
    if ($excel.Run("'" + $book.Name + "'!UnicodeProbe.ErrorText") -cne $sample) { throw 'Err.Description Unicode failed.' }
    $excel.StatusBar = $sample
    if ($excel.StatusBar -cne $sample) { throw 'StatusBar Unicode failed.' }
    $logPath = Join-Path $root 'diagnostic.log'
    $fileSystem = New-Object -ComObject Scripting.FileSystemObject
    $log = $fileSystem.OpenTextFile($logPath, 8, $true, -1)
    $log.WriteLine($sample)
    $log.Close()
    if ([IO.File]::ReadAllText($logPath, [Text.Encoding]::Unicode).TrimEnd() -cne $sample) { throw 'Unicode log round trip failed.' }
    $updaterComponent = $book.VBProject.VBComponents.Add(1)
    $updaterComponent.Name = 'ex_WorkbookUpdater'
    $updaterSource = [IO.File]::ReadAllText((Join-Path $PSScriptRoot '../../PROJECTS_EXCEL/WorkbookUpdater/ex_WorkbookUpdater.vba'), $utf8)
    $updaterComponent.CodeModule.AddFromString((ConvertTo-WorkbookModuleCode $updaterSource))
    $updaterComponent.CodeModule.AddFromString(@"
' namespace Test {
Public Function ValidateEncodingProbe(ByVal path As String, ByVal nativeSource As Boolean) As Boolean
    On Error GoTo Rejected
    private_ValidateSourceEncoding path, nativeSource
    ValidateEncodingProbe = True
Rejected:
End Function
Public Function PrepareEncodingProbe(ByVal text As String) As String
    PrepareEncodingProbe = private_VbaReload_PrepareSourceForVbe(text)
End Function
' } // namespace Test
"@)
    $prefix = "'" + $book.Name + "'!ex_WorkbookUpdater."
    [IO.File]::WriteAllText($path, $sample, $utf8)
    if (-not $excel.Run($prefix + 'ValidateEncodingProbe', $path, $false)) { throw 'Valid UTF-8 rejected by XLAM.' }
    if ($excel.Run($prefix + 'ValidateEncodingProbe', $path, $true)) { throw 'Non-ASCII native accepted by XLAM.' }
    foreach ($bytes in @(@(0xC0, 0xAF), @(0xED, 0xA0, 0x80), @(0xF4, 0x90, 0x80, 0x80), @(0xE2, 0x82))) {
        [IO.File]::WriteAllBytes($path, [byte[]]$bytes)
        if ($excel.Run($prefix + 'ValidateEncodingProbe', $path, $false)) { throw 'Invalid UTF-8 accepted by XLAM.' }
    }
    if ($excel.Run($prefix + 'PrepareEncodingProbe', $comment).TrimEnd() -cne $comment) { throw 'XLAM comment was modified.' }
    $prepared = $excel.Run($prefix + 'PrepareEncodingProbe', $source)
    if ($prepared -match '[^\x00-\x7F]') { throw 'XLAM preparation is not ASCII.' }
    $component.CodeModule.DeleteLines(1, $component.CodeModule.CountOfLines)
    $component.CodeModule.AddFromString($prepared)
    if ($excel.Run("'" + $book.Name + "'!UnicodeProbe.ReadText") -cne $sample) { throw 'XLAM Unicode execution failed.' }
    'PASS: strict UTF-8, native rejection, ASCII preparation, VBA, status bar and UTF-16 log round trip'
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}