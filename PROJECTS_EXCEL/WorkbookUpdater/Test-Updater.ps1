param([string]$UpdaterPath = (Join-Path $PSScriptRoot 'WorkbookUpdater.xlam'))

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$fixtureRoot = Join-Path ([IO.Path]::GetTempPath()) ('WorkbookUpdater-test-' + [Guid]::NewGuid().ToString('N'))
$sourceFolder = Join-Path $fixtureRoot 'vba'
[IO.Directory]::CreateDirectory($sourceFolder) | Out-Null
[IO.Directory]::CreateDirectory((Join-Path $fixtureRoot 'ui')) | Out-Null
$utf8 = New-Object Text.UTF8Encoding($false)
$excel = $null
$book = $null
$otherBook = $null
$updaterBook = $null

function Write-Source([string]$Name, [string]$Code) {
    [IO.File]::WriteAllText((Join-Path $sourceFolder ($Name + '.vba')), $Code.TrimEnd(), $utf8)
}

function Assert-Equal($Expected, $Actual, [string]$Message) {
    if ($Expected -ne $Actual) { throw "${Message}: expected '$Expected', got '$Actual'." }
}

function Wait-Result([string]$Expected) {
    $deadline = [DateTime]::UtcNow.AddSeconds(15)
    do {
        Start-Sleep -Milliseconds 300
        try {
            $result = $excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_LastResult")
        }
        catch {
            if ([DateTime]::UtcNow -ge $deadline) {
                throw "Excel cannot execute the status macro; project mode=$($book.VBProject.Mode); $($_.Exception.Message)"
            }
            $result = 'Pending'
        }
        if ($result -eq $Expected) { return }
        if ($result -eq 'Faulted') {
            throw $excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_LastError")
        }
    } while ([DateTime]::UtcNow -lt $deadline)
    $detail = $excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_LastError")
    throw "Timed out waiting for '$Expected'; last result '$result'; $detail."
}

try {
    $lifecycle = [IO.File]::ReadAllText((Join-Path $PSScriptRoot '..\vba\common\modules\ex_RuntimeLifecycle.vba'))
    Write-Source 'ex_RuntimeLifecycle' $lifecycle
    $callbacks = [IO.File]::ReadAllText((Join-Path $PSScriptRoot '..\vba\common\modules\ex_WorkbookCallbacks.vba'))
    Write-Source 'ex_WorkbookCallbacks' $callbacks
    foreach ($name in @('ex_AppHotkeys', 'ex_UiPageManager', 'ex_UiBindings', 'ex_UiRuntime', 'ex_UiElementFactory', 'ex_UiControlFactory')) {
        Write-Source $name @'
Option Explicit
Public Sub fn_Module_Dispose()
End Sub
Public Sub fn_PrepareReload()
End Sub
Public Sub fn_Activate()
End Sub
'@
    }
    Write-Source 'ex_Core' @'
Option Explicit
Public Function fn_Diagnostic_Flush() As Boolean
    fn_Diagnostic_Flush = True
End Function
Public Function fn_TryGetWorkbookConfigValue( _
    ByVal key As String, _
    ByRef value As String _
) As Boolean
    Dim row As ListRow

    For Each row In ThisWorkbook.Worksheets(1).ListObjects("tbConfig").ListRows
        If row.Range.Cells(1, 1).Value2 = key Then
            value = row.Range.Cells(1, 2).Value2
            fn_TryGetWorkbookConfigValue = True
            Exit Function
        End If
    Next row
End Function
Public Sub fn_Module_Dispose()
End Sub
'@
    Write-Source 'ex_RuntimePaths' @'
Option Explicit
Public Sub fn_SetUiFolder(ByVal value As String)
End Sub
Public Sub fn_Module_Dispose()
End Sub
'@
    Write-Source 'ex_PersonalEventBuilder' @'
Option Explicit
Private m_initialized As Boolean
Public Function fn_Initialize() As Boolean
    m_initialized = True
    fn_Initialize = True
End Function
Public Function IsInitialized() As Boolean
    IsInitialized = m_initialized
End Function
'@
    Write-Source 'ThisWorkbook' 'Option Explicit'
    Write-Source 'ex_Probe' @'
Option Explicit
Private m_lease As Object
Public Sub Hold()
    If Not ex_RuntimeLifecycle.fn_TryEnter(m_lease) Then Err.Raise 5
End Sub
Public Sub Release()
    ex_RuntimeLifecycle.fn_Leave m_lease
    Set m_lease = Nothing
End Sub
Public Function CanEnter() As Boolean
    Dim context As Object
    CanEnter = ex_RuntimeLifecycle.fn_TryEnter(context)
    If CanEnter Then ex_RuntimeLifecycle.fn_Leave context
End Function
Public Function Phase() As String
    Dim context As Object
    Set context = ex_RuntimeLifecycle.fn_Context()
    Phase = context("Phase")
End Function
Public Function Generation() As Long
    Dim context As Object
    Set context = ex_RuntimeLifecycle.fn_Context()
    Generation = CLng(context("Generation"))
End Function
Public Function Value() As Long
    Value = 1
End Function
'@
    [IO.File]::WriteAllText((Join-Path $sourceFolder 'modules.json'), '{"Test":["*.vba"]}', $utf8)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $updaterBook = $excel.Workbooks.Open([IO.Path]::GetFullPath($UpdaterPath))
    $book = $excel.Workbooks.Add()
    $sheet = $book.Worksheets.Item(1)
    $sheet.Cells.Item(1,1).Value2 = 'Key'
    $sheet.Cells.Item(1,2).Value2 = 'Value'
    $sheet.Cells.Item(2,1).Value2 = 'ThisWorkbook::id'
    $sheet.Cells.Item(2,2).Value2 = 'Test'
    $sheet.Cells.Item(3,1).Value2 = 'ThisWorkbook::vbaPath'
    $sheet.Cells.Item(3,2).Value2 = 'vba'
    $sheet.Cells.Item(4,1).Value2 = 'ThisWorkbook::uiPath'
    $sheet.Cells.Item(4,2).Value2 = 'ui'
    $sheet.Cells.Item(5,1).Value2 = 'ThisWorkbook::initializer'
    $sheet.Cells.Item(5,2).Value2 = 'ex_PersonalEventBuilder.fn_Initialize'
    $table = $sheet.ListObjects.Add(1, $sheet.Range('A1:B5'), $null, 1)
    $table.Name = 'tbConfig'
    $book.SaveAs((Join-Path $fixtureRoot 'Target.xlsm'), 52)
    $legacy = $book.VBProject.VBComponents.Add(1)
    $legacy.Name = 'LegacyCode'
    $legacy.CodeModule.AddFromString('Public Sub LegacyEntry(): End Sub')
    Assert-Equal $false ($excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_RequestReload", $book, $false)) 'Legacy code without lifecycle is rejected'
    Assert-Equal 1 $legacy.CodeModule.CountOfLines 'Rejected legacy code remains intact'
    $book.VBProject.VBComponents.Remove($legacy)
    $book.Names.Add('_RuntimeReloadBlocked', '=TRUE') | Out-Null
    Assert-Equal $false ($excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_RequestReload", $book, $false)) 'Blocked empty project is rejected'
    $book.Names.Item('_RuntimeReloadBlocked').Delete()
    Assert-Equal $true ($excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_RequestReload", $book, $false)) 'Empty project initial request accepted'
    $excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_CancelPending", $book) | Out-Null
    Assert-Equal 'Cancelled' ($excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_LastResult")) 'Initial request cancelled'
    Assert-Equal $true ($excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_RequestReload", $book, $false)) 'Initial request can be repeated after cancellation'
    Wait-Result 'Completed'
    Assert-Equal $true ($excel.Run("'Target.xlsm'!ex_PersonalEventBuilder.IsInitialized")) 'Target module state survives initial callback return'
    Assert-Equal $true ($excel.Run("'Target.xlsm'!ex_Probe.CanEnter")) 'Initial installation initializes the runtime'
    Assert-Equal 1 ($excel.Run("'Target.xlsm'!ex_Probe.Generation")) 'Initial generation established'
    $otherBook = $excel.Workbooks.Add()
    $baselineBookCount = $excel.Workbooks.Count
    $context = $excel.Run("'Target.xlsm'!ex_RuntimeLifecycle.fn_Context")
    $excel.Run("'Target.xlsm'!ex_Probe.Hold") | Out-Null
    Assert-Equal $true ($excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_RequestReload", $book, $false)) 'Request accepted'
    Assert-Equal $false ($excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_CanUnload")) 'Pending updater cannot unload'
    Assert-Equal $false ($excel.Run("'Target.xlsm'!ex_Probe.CanEnter")) 'New calls blocked'
    $otherBook.Activate()
    Start-Sleep -Milliseconds 2200
    Assert-Equal 'Pending' ($excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_LastResult")) 'Waits for active calls'
    $probePath = Join-Path $sourceFolder 'ex_Probe.vba'
    $probeSource = [IO.File]::ReadAllText($probePath)
    [IO.File]::WriteAllText($probePath, $probeSource.Replace('Value = 1', 'Value = 99'), $utf8)
    $excel.Run("'Target.xlsm'!ex_Probe.Release") | Out-Null
    Wait-Result 'Completed'
    Assert-Equal 1 ($excel.Run("'Target.xlsm'!ex_Probe.Value")) 'Imports captured source snapshot'
    Assert-Equal $true ($excel.Run("'Target.xlsm'!ex_PersonalEventBuilder.IsInitialized")) 'Target module state survives reload callback return'
    Assert-Equal 2 ($excel.Run("'Target.xlsm'!ex_Probe.Generation")) 'Preserves external context'
    Assert-Equal $true ($excel.Run("'Target.xlsm'!ex_Probe.CanEnter")) 'Reopens entry gate'
    Assert-Equal $true ($excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_CanUnload")) 'Completed updater can unload'
    Assert-Equal $baselineBookCount $excel.Workbooks.Count 'Does not close target or other workbook'
    Assert-Equal $false $excel.EnableEvents 'Restores original event setting'
    $backups = @(Get-ChildItem -LiteralPath (Join-Path $fixtureRoot '.backup') -Directory -Filter 'reload-*')
    $targetBackups = @($backups | Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'Target.xlsm') })
    Assert-Equal $true ($targetBackups.Count -ge 2) 'Creates initial-install and reload backups'
    Assert-Equal $true ($excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_RequestReload", $book, $false)) 'Second request accepted'
    $excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_CancelPending", $book) | Out-Null
    Assert-Equal 'Cancelled' ($excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_LastResult")) 'Cancels exact OnTime callback'
    Assert-Equal $true ($excel.Run("'Target.xlsm'!ex_Probe.CanEnter")) 'Cancellation restores gate'
    Start-Sleep -Milliseconds 1500
    Assert-Equal 'Cancelled' ($excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_LastResult")) 'Cancelled callback does not run'
    $excel.Run("'Target.xlsm'!ex_Probe.Hold") | Out-Null
    Assert-Equal $true ($excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_RequestClear", $book, $false)) 'Safe clear request accepted'
    Start-Sleep -Milliseconds 1500
    Assert-Equal 'Pending' ($excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_LastResult")) 'Clear waits for active calls'
    $excel.Run("'Target.xlsm'!ex_Probe.Release") | Out-Null
    Wait-Result 'Cleared'
    foreach ($component in $book.VBProject.VBComponents) {
        Assert-Equal 100 $component.Type 'Clear retains only document components'
        Assert-Equal 0 $component.CodeModule.CountOfLines 'Clear removes document code'
    }
    Assert-Equal $true ($excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_CanUnload")) 'Cleared updater can unload'
    Assert-Equal $true ($excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_RequestReload", $book, $false)) 'Cleared workbook accepts initial installation again'
    Wait-Result 'Completed'
    Assert-Equal $true ($excel.Run("'Target.xlsm'!ex_PersonalEventBuilder.IsInitialized")) 'Runtime recreated after clear and install'
    Write-Source 'ex_PersonalEventBuilder' @'
Option Explicit
Public Function fn_Initialize() As Boolean
    fn_Initialize = False
End Function
'@
    Assert-Equal $true ($excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_RequestReload", $book, $false)) 'Failure test request accepted'
    $deadline = [DateTime]::UtcNow.AddSeconds(15)
    do {
        Start-Sleep -Milliseconds 300
        $result = $excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_LastResult")
    } while ($result -eq 'Pending' -and [DateTime]::UtcNow -lt $deadline)
    Assert-Equal 'Faulted' $result 'Initialization failure blocks runtime'
    Assert-Equal $false ($excel.Run("'Target.xlsm'!ex_Probe.CanEnter")) 'Fault gate stays closed'
    $book.Names.Item('_RuntimeReloadBlocked') | Out-Null
    Assert-Equal $false $excel.EnableEvents 'Failure restores global event setting'
    Write-Output 'PASS: empty-project installation, safe clear and reinstall, legacy/blocked rejection, quiescence, entry gate, source snapshot, captured target, context survival, backup, cancellation, initialization failure.'
    Write-Output "Fixture preserved: $fixtureRoot"
}
finally {
    if ($null -ne $excel) {
        $excel.EnableEvents = $false
        if ($null -ne $updaterBook -and $null -ne $book) {
            try { $excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_CancelPending", $book) | Out-Null } catch { }
        }
        if ($null -ne $book) { $book.Close($false) }
        if ($null -ne $otherBook) { $otherBook.Close($false) }
        if ($null -ne $updaterBook) { $updaterBook.Close($false) }
        $excel.Quit()
    }
}