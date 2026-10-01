param(
    [Parameter(Mandatory = $true)][string]$WorkbookPath,
    [Parameter(Mandatory = $true)][string]$VbaFolderPath,
    [int]$CloseWaitSeconds = 30,
    [string]$LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$targetPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$lockPath = Join-Path ([IO.Path]::GetDirectoryName($targetPath)) ('~$' + [IO.Path]::GetFileName($targetPath))
$deadline = [DateTime]::UtcNow.AddSeconds($CloseWaitSeconds)
while (Test-Path -LiteralPath $lockPath) {
    if ([DateTime]::UtcNow -ge $deadline) { throw "The workbook was not closed within $CloseWaitSeconds seconds: $targetPath" }
    Start-Sleep -Milliseconds 250
}
$targetExcel = $null
try { $targetExcel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') }
catch { $targetExcel = $null }
$output = & (Join-Path $PSScriptRoot '..\..\Scripts\PowerShell\Update-WorkbookModules.ps1') `
    -WorkbookPath $targetPath -VbaFolderPath $VbaFolderPath
$output | Write-Output
if ($LogPath) {
    [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($LogPath))) | Out-Null
    $encoding = if ([IO.Path]::GetFileName($LogPath) -ieq 'diagnostic.log') { [Text.Encoding]::Unicode } else { [Text.Encoding]::UTF8 }
    [IO.File]::AppendAllText($LogPath, ($output -join [Environment]::NewLine) + [Environment]::NewLine, $encoding)
}
if ($null -ne $targetExcel) { $targetExcel.Workbooks.Open($targetPath) | Out-Null }
else { Start-Process -FilePath 'excel.exe' -ArgumentList ('"' + $targetPath + '"') -WindowStyle Hidden }