param(
    [string]$WorkbookName,
    [string]$WorkbookPath,
    [string]$VbaFolderPath = (Join-Path $PSScriptRoot 'vba'),
    [string]$Profile,
    [ValidateSet('Update', 'Clear')][string]$Mode = 'Update',
    [switch]$Create,
    [switch]$AsAddin,
    [string]$CanUpdateMacro,
    [string]$PrepareMacro,
    [string]$InitializeMacro,
    [switch]$PlanOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'WorkbookModules.ps1')

if ($Create -and ([string]::IsNullOrWhiteSpace($WorkbookPath) -or $WorkbookName)) {
    throw 'Create requires WorkbookPath and cannot be combined with WorkbookName.'
}
if ($AsAddin -and -not $Create) { throw 'AsAddin is only supported with Create.' }
if ($Create -and $Mode -ne 'Update') { throw 'Create requires Update mode.' }
if ($WorkbookName -and $WorkbookPath) { throw 'Specify either WorkbookName or WorkbookPath.' }
if (-not $PlanOnly -and -not $WorkbookName -and -not $WorkbookPath) {
    throw 'Specify WorkbookName for a loaded workbook or WorkbookPath for a closed workbook.'
}

$excel = $null
$book = $null
$ownsExcel = $false
$previousEvents = $null
$previousScreenUpdating = $null
$previousCalculation = $null
$backupPath = $null
try {
    # Проверяем схему до подключения к Excel, если профиль задан явно.
    $plan = $null
    if ($Mode -eq 'Update' -and $Profile) { $plan = @(Get-WorkbookModulePlan $VbaFolderPath $Profile) }
    if ($PlanOnly) {
        if (-not $Profile -or $Mode -ne 'Update') { throw 'PlanOnly requires Profile and Update mode.' }
        $plan | Select-Object Name, Type, Document, Path
        return
    }
    if ($WorkbookName) {
        $loadedTarget = Find-LoadedExcelWorkbook $WorkbookName
        $excel = $loadedTarget.Excel
        $book = $loadedTarget.Workbook
    }
    else {
        $WorkbookPath = [IO.Path]::GetFullPath($WorkbookPath)
        if ($Create) {
            if (Test-Path -LiteralPath $WorkbookPath) { throw "Output already exists: $WorkbookPath" }
            if (-not (Test-Path -LiteralPath ([IO.Path]::GetDirectoryName($WorkbookPath)) -PathType Container)) {
                throw "Output folder does not exist: $WorkbookPath"
            }
        }
        else {
            if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) { throw "Workbook not found: $WorkbookPath" }
            $lockPath = Join-Path ([IO.Path]::GetDirectoryName($WorkbookPath)) ('~$' + [IO.Path]::GetFileName($WorkbookPath))
            if (Test-Path -LiteralPath $lockPath) { throw 'Close the workbook first, or use WorkbookName to target the loaded workbook.' }
        }
        $excel = New-Object -ComObject Excel.Application
        $ownsExcel = $true
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AutomationSecurity = 3
    }
    $previousEvents = $excel.EnableEvents
    $previousScreenUpdating = $excel.ScreenUpdating
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    if ($null -eq $book) {
        if ($Create) { $book = $excel.Workbooks.Add() }
        else { $book = $excel.Workbooks.Open($WorkbookPath, 0, $false) }
    }
    $hasVisibleWindow = $false
    foreach ($bookWindow in $book.Windows) {
        if ($bookWindow.Visible) { $hasVisibleWindow = $true }
    }
    if (-not $book.IsAddin -and $hasVisibleWindow) {
        $previousCalculation = $excel.Calculation
        $excel.Calculation = -4135
    }
    if ($book.ReadOnly) { throw 'The workbook is read-only.' }
    if ($book.VBProject.Protection -ne 0) { throw 'The VBA project is protected.' }
    if ($book.VBProject.Mode -ne 2) { throw 'The VBA project must be in Design mode.' }
    if ($Mode -eq 'Update' -and $null -eq $plan) {
        $Profile = Get-WorkbookProfile $book
        $plan = @(Get-WorkbookModulePlan $VbaFolderPath $Profile)
    }
    # Все документные компоненты проверяются до удаления любого кода.
    if ($Mode -eq 'Update') { Assert-WorkbookModulePlan $book $plan }
    if ($CanUpdateMacro) {
        $macro = "'" + $book.Name.Replace("'", "''") + "'!" + $CanUpdateMacro
        if ($excel.Run($macro) -ne $true) { throw 'The workbook lifecycle rejected the update.' }
    }
    if (-not $Create) {
        $backupPath = Join-Path $book.Path ('.modules-backup-' + [Guid]::NewGuid().ToString('N'))
        [IO.Directory]::CreateDirectory($backupPath) | Out-Null
        $backupPath = Join-Path $backupPath $book.Name
        $book.SaveCopyAs($backupPath)
        Write-Output "Backup: $backupPath"
    }
    if ($PrepareMacro) {
        $macro = "'" + $book.Name.Replace("'", "''") + "'!" + $PrepareMacro
        $excel.Run($macro) | Out-Null
    }
    Set-WorkbookModulePlan $book $plan $Mode
    if ($InitializeMacro) {
        $macro = "'" + $book.Name.Replace("'", "''") + "'!" + $InitializeMacro
        $initializeResult = $excel.Run($macro)
        if ($initializeResult -is [bool] -and -not $initializeResult) { throw "Initialization returned False: $InitializeMacro" }
    }
    if ($Create) {
        if ($AsAddin) {
            if ([IO.Path]::GetExtension($WorkbookPath) -ine '.xlam') { throw 'AsAddin requires an .xlam output.' }
            $excel.Calculation = $previousCalculation
            $previousCalculation = $null
            $book.IsAddin = $true
            $book.SaveAs($WorkbookPath, 55)
        }
        else {
            if ([IO.Path]::GetExtension($WorkbookPath) -ine '.xlsm') { throw 'Create requires an .xlsm output unless AsAddin is specified.' }
            $book.SaveAs($WorkbookPath, 52)
        }
    }
    else { $book.Save() }
    $completedPath = if ($Create) { $WorkbookPath } else { $book.FullName }
    Write-Output "Completed: $completedPath; mode=$Mode; profile=$Profile"
}
catch {
    if ($backupPath) { Write-Warning "Update failed. Automatic rollback is not performed. Backup: $backupPath" }
    throw
}
finally {
    if ($null -ne $excel) {
        try {
        if ($null -ne $previousCalculation) { $excel.Calculation = $previousCalculation }
        if ($null -ne $previousScreenUpdating) { $excel.ScreenUpdating = $previousScreenUpdating }
        if ($null -ne $previousEvents) { $excel.EnableEvents = $previousEvents }
        }
        finally {
        if ($ownsExcel) {
            if ($null -ne $book) {
                $book.Close($false)
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
                $book = $null
            }
            $excel.Quit()
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
            $excel = $null
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
        }
    }
}