param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'WorkbookUpdater.xlam'),
    [switch]$UpdateExisting
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$staging = Join-Path ([IO.Path]::GetTempPath()) ('WorkbookUpdaterSources-' + [guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($staging) | Out-Null
try {
    Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.vba' | Copy-Item -Destination $staging
    $schema = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'modules.json') | ConvertFrom-Json
    $schema.WorkbookUpdater += 'ex_ConfigImports.vba'
    [IO.File]::WriteAllText((Join-Path $staging 'modules.json'), ($schema | ConvertTo-Json))
    # Package the shared resolver with the standalone Excel reader.
    $sourcePath = Join-Path $PSScriptRoot '../vba/common/modules/ex_ConfigImports.vba'
    $source = [IO.File]::ReadAllText($sourcePath).Replace('#Const CONFIG_QUERY_AVAILABLE = True', '#Const CONFIG_QUERY_AVAILABLE = False')
    [IO.File]::WriteAllText((Join-Path $staging 'ex_ConfigImports.vba'), $source)
    $installer = Join-Path $PSScriptRoot '..\..\Scripts\PowerShell\Update-WorkbookModules.ps1'
    if ($UpdateExisting) {
        & $installer -WorkbookPath $OutputPath -VbaFolderPath $staging -Profile WorkbookUpdater
    }
    else {
        & $installer -Create -AsAddin -WorkbookPath $OutputPath -VbaFolderPath $staging -Profile WorkbookUpdater
    }
}
finally {
    # Only remove files created in this unique staging directory.
    foreach ($name in @('modules.json', 'ex_ConfigImports.vba', 'ex_ConfigRefresh.vba', 'ex_WindowsUi.vba', 'ex_WorkbookUpdater.vba', 'ThisWorkbook.vba')) {
        $file = Join-Path $staging $name
        if (Test-Path -LiteralPath $file) { Remove-Item -LiteralPath $file }
    }
    [IO.Directory]::Delete($staging, $false)
}