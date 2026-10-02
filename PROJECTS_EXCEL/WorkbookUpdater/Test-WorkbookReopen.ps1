param(
    [Parameter(Mandatory = $true)][string]$WorkbookPath,
    [string]$UpdaterPath,
    [string]$Root,
    [ValidateSet('Prepare','Hot','Reopen','AutoOpen')][string]$Mode
)
$ErrorActionPreference = 'Stop'
if (-not $UpdaterPath) { $UpdaterPath = Join-Path $PSScriptRoot 'WorkbookUpdater.xlam' }
$WorkbookPath = [IO.Path]::GetFullPath($WorkbookPath)
$UpdaterPath = [IO.Path]::GetFullPath($UpdaterPath)
if (-not $Mode) {
    $Root = Join-Path ([IO.Path]::GetTempPath()) ('WorkbookReopen-' + [Guid]::NewGuid().ToString('N'))
    [IO.Directory]::CreateDirectory((Join-Path $Root 'book')) | Out-Null
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot '../vba') -Destination $Root -Recurse
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot '../ui') -Destination $Root -Recurse
    Copy-Item -LiteralPath $WorkbookPath -Destination (Join-Path $Root 'book/Fixture.xlsm')
    $probe = @"
Option Explicit
' namespace Test {
Public Sub UpdatePage()
    Dim command As obj_UiCommand
    If Not ex_UiBindings.fn_TryGetCommand("btn_UpdatePage", command) Then Err.Raise 5
    command.Execute
End Sub
Public Sub Generate()
    Dim command As obj_UiCommand
    Dim rawTable As obj_UiRawTable
    Dim returnedTable As obj_UiRawTable
    Dim source As obj_IUiTableSource
    Dim values(1 To 1, 1 To 1) As Variant
    values(1, 1) = "Probe"
    Set rawTable = New obj_UiRawTable
    If Not rawTable.Initialize(values, Array("Header"), "Probe") Then Err.Raise 5
    Set source = rawTable
    Set returnedTable = source.GetTable(1)
    If Not (returnedTable Is rawTable) Then Err.Raise 5
    If Not ex_UiBindings.fn_TryGetCommand("btn_GenerateTables", command) Then Err.Raise 5
    command.Execute
    If ThisWorkbook.Worksheets("MainPage").UsedRange.Find("Candidate 10.3") Is Nothing Then Err.Raise 5
End Sub
' } // namespace Test
"@
    [IO.File]::WriteAllText((Join-Path $Root 'vba/common/modules/ex_ReopenProbe.vba'), $probe.TrimEnd(), [Text.UTF8Encoding]::new($false))
    $shell = Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe'
    $step = 0
    foreach ($phase in @('Prepare','Hot','Reopen','AutoOpen','Reopen','Hot','Reopen')) {
        $step++
        $stdout = Join-Path $Root "$step-$phase.out.txt"
        $stderr = Join-Path $Root "$step-$phase.err.txt"
        $arguments = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -WorkbookPath "{1}" -UpdaterPath "{2}" -Root "{3}" -Mode {4}' -f $PSCommandPath, $WorkbookPath, ([IO.Path]::GetFullPath($UpdaterPath)), $Root, $phase
        $process = Start-Process -FilePath $shell -ArgumentList $arguments -WindowStyle Hidden -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
        $processHandle = $process.Handle
        if (-not $process.WaitForExit(45000)) {
            # Only terminate the Excel process created and identified by this worker.
            $output = [IO.File]::ReadAllText($stdout)
            if ($output -match 'Excel PID=(\d+)') { Stop-Process -Id ([int]$Matches[1]) -ErrorAction SilentlyContinue }
            Stop-Process -Id $process.Id -ErrorAction SilentlyContinue
            throw "Reopen regression timed out at $phase. Evidence: $Root"
        }
        $exitCode = $process.ExitCode
        Get-Content -LiteralPath $stdout
        if ($exitCode -ne 0) { throw ([IO.File]::ReadAllText($stderr) + " Evidence: $Root") }
    }
    Write-Output "PASS: hot reload, save before Update Page, fresh Excel reopen, Buffered persistence, single/list table sources, ten generated tables and repeated reload. Evidence: $Root"
    return
}
$path = Join-Path $Root 'book/Fixture.xlsm'
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $previousAutoRecover = $excel.AutoRecover.Enabled
    $excel.AutoRecover.Enabled = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = ($Mode -eq 'AutoOpen')
    $excel.AutomationSecurity = $(if ($Mode -eq 'Prepare') {3} else {1})
    Add-Type 'using System; using System.Runtime.InteropServices; public class ReopenPid { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p); }'
    [uint32]$excelPid = 0
    [void][ReopenPid]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$excelPid)
    "Excel PID=$excelPid Mode=$Mode"
    "Opening $path"
    if ($Mode -eq 'AutoOpen') {
        $personalPath = Join-Path $env:APPDATA 'Microsoft/Excel/XLSTART/PERSONAL.XLSB'
        if (Test-Path -LiteralPath $personalPath) { $personal = $excel.Workbooks.Open($personalPath) }
    }
    $book = $excel.Workbooks.Open($path,0,$false)
    if ($book.ReadOnly) { throw 'Fixture is read-only.' }
    'Opened'
    if ($Mode -eq 'Prepare') {
        foreach ($sheet in $book.Worksheets) {
            foreach ($table in $sheet.ListObjects) {
                if ($table.Name -ne 'tbConfig') {continue}
                $col = $table.ListColumns.Item('Key').Index
                foreach ($row in $table.ListRows) {
                    switch ([string]$row.Range.Cells.Item(1,$col).Value2) {
                        'ThisWorkbook::vbaPath' {$row.Range.Cells.Item(1,$col+1).Value2='..\vba'}
                        'ThisWorkbook::uiPath' {$row.Range.Cells.Item(1,$col+1).Value2='..\ui'}
                        'ThisWorkbook::logPath' {$row.Range.Cells.Item(1,$col+1).Value2=$Root}
                        'ThisWorkbook::logMode' {$row.Range.Cells.Item(1,$col+1).Value2='Immediate'}
                    }
                }
            }
        }
        . (Join-Path $PSScriptRoot '../../Scripts/PowerShell/WorkbookModules.ps1')
        Set-WorkbookModulePlan $book @() 'Clear'
        $book.Save()
    } elseif ($Mode -eq 'Hot') {
        $addon = $excel.Workbooks.Open($UpdaterPath)
        if (-not $excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_RequestReload",$book,$false)) {throw 'Reload rejected'}
        $deadline = [DateTime]::UtcNow.AddSeconds(30)
        do {
            Start-Sleep -Milliseconds 300
            $result = $excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_LastResult")
            if ($result -eq 'Faulted') {throw $excel.Run("'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_LastError")}
        } while ($result -ne 'Completed' -and [DateTime]::UtcNow -lt $deadline)
        "Reload status=$result"
        if ($result -ne 'Completed') {throw 'Reload timeout'}
        $book.Save()
        'Saved after import'
        $book.Activate()
        $book.Worksheets.Item('MainPage').Activate()
        $excel.Run("'Fixture.xlsm'!ex_ReopenProbe.UpdatePage")
        'Updated page'
    } else {
        $book.Activate()
        $book.Worksheets.Item('MainPage').Activate()
        if ($Mode -ne 'AutoOpen' -and -not $excel.Run("'Fixture.xlsm'!ex_PersonalEventBuilder.fn_Initialize")) {throw 'Initialization failed'}
        'Initialized'
        if (-not $excel.Run("'Fixture.xlsm'!ex_Core.fn_Diagnostic_SetMode",'Buffered')) { throw 'Mode change failed' }
        $excel.Run("'Fixture.xlsm'!ex_ReopenProbe.Generate")
        'Generated and verified ten tables'
        $book.Save()
        'Saved after initialization'
        $excel.Run("'Fixture.xlsm'!ex_ReopenProbe.UpdatePage")
        'Updated page'
    }
    $excel.EnableEvents = $true
    $book.Close($false)
    $book=$null
    'PASS'
} finally {
    if ($null -ne $book) {try {$book.Close($false)} catch {}}
    if ($null -ne $excel) {try {$excel.AutoRecover.Enabled = $previousAutoRecover; $excel.Quit()} catch {}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}