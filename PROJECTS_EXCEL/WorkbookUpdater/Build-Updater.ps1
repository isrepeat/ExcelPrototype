param([string]$OutputPath = (Join-Path $PSScriptRoot 'WorkbookUpdater.xlam'))

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
& (Join-Path $PSScriptRoot '..\..\Scripts\PowerShell\Update-WorkbookModules.ps1') `
    -Create -AsAddin -WorkbookPath $OutputPath -VbaFolderPath $PSScriptRoot -Profile WorkbookUpdater