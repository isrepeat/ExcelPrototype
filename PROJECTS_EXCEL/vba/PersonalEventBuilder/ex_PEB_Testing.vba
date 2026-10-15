Attribute VB_Name = "ex_PEB_Testing"
Option Explicit

Private m_testUi As obj_PEB_TestUi
Private m_tickAt As Date
Private m_tickPending As Boolean
Private m_tickProcedure As String
Private m_advancing As Boolean

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    fn_CancelTick
    If Not m_testUi Is Nothing Then m_testUi.Dispose
    Set m_testUi = Nothing
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_IsRunning() As Boolean
    If Not m_testUi Is Nothing Then fn_IsRunning = m_testUi.IsRunning
End Function

Public Function fn_ReadCatalog( _
    ByRef catalog As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim executor As obj_AdoQueryExecutor
    Dim source As obj_TableSource
    Dim query As obj_TableQuery
    Dim data As obj_DataTable
    Dim cases As Object
    Dim item As Object
    Dim numbers As Object
    Dim steps As Collection
    Dim ordered As Collection
    Dim kind As Variant
    Dim id As String
    Dim number As String
    Dim index As Long
    Dim row As Long
    Dim started As Double

    started = VBA.Timer
    On Error GoTo Failed
    Set catalog = New Collection
    Set cases = VBA.CreateObject("Scripting.Dictionary")
    Set executor = New obj_AdoQueryExecutor
    If Not executor.Initialize() Then Err.Raise 5, , "Cannot initialize catalog reader."
    For Each kind In VBA.Array("cases", "steps")
        Set source = New obj_TableSource
        If Not source.Initialize() Then Err.Raise 5, , "Cannot initialize catalog source."
        source.WorkbookPath = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("test.scenarios.path"))
        source.SheetName = ex_PEB_Export.fn_Config("test.scenarios." & kind & ".sheet")
        source.RangeAddress = ex_PEB_Export.fn_Config("test.scenarios." & kind & ".readRange")
        Set query = New obj_TableQuery
        If Not query.Initialize() Then Err.Raise 5, , "Cannot initialize catalog query."
        query.AddColumn ex_PEB_Export.fn_Config("test.scenarios.headers.id")
        If kind = "cases" Then
            query.AddColumn ex_PEB_Export.fn_Config("test.scenarios.headers.name")
            query.AddColumn ex_PEB_Export.fn_Config("test.scenarios.headers.ruleStatus")
            query.AddColumn ex_PEB_Export.fn_Config("test.scenarios.headers.catalogOperation")
            query.AddColumn ex_PEB_Export.fn_Config("test.scenarios.headers.catalogEvent")
        Else
            query.AddColumn ex_PEB_Export.fn_Config("test.scenarios.headers.step")
        End If
        If Not executor.TryExecute(source, query, data, diagnostic) Then GoTo CleanExit
        For row = 1 To data.RowCount
            id = ex_PEB_Export.fn_Text(data.ValueAt(row, 1))
            If VBA.Len(id) > 0 Then
                If kind = "cases" Then
                    If cases.Exists(id) Then Err.Raise 5, , "Duplicate scenario identity: " & id
                    Set item = VBA.CreateObject("Scripting.Dictionary")
                    item("Id") = id
                    item("Name") = ex_PEB_Export.fn_Text(data.ValueAt(row, 2))
                    item("RuleStatus") = ex_PEB_Export.fn_Text(data.ValueAt(row, 3))
                    item("Operation") = ex_PEB_Export.fn_Text(data.ValueAt(row, 4))
                    item("EventLabel") = ex_PEB_Export.fn_Text(data.ValueAt(row, 5))
                    If VBA.Len(item("Operation")) = 0 Or VBA.Len(item("EventLabel")) = 0 Then Err.Raise 5, , "Scenario classification is missing: " & id
                    If item("RuleStatus") <> "Defined" And item("RuleStatus") <> "PendingRule" Then Err.Raise 5, , "Invalid scenario rule status: " & id
                    Set numbers = VBA.CreateObject("Scripting.Dictionary")
                    Set item("Numbers") = numbers
                    cases.Add id, item
                    catalog.Add item
                Else
                    If Not cases.Exists(id) Then Err.Raise 5, , "Step references an unknown scenario: " & id
                    number = ex_PEB_Export.fn_Text(data.ValueAt(row, 2))
                    If Not VBA.IsNumeric(number) Then Err.Raise 5, , "Invalid scenario step: " & id
                    If VBA.CDbl(number) < 1 Or VBA.CDbl(number) <> VBA.Fix(VBA.CDbl(number)) Then Err.Raise 5, , "Invalid scenario step: " & id
                    Set numbers = cases(id)("Numbers")
                    number = VBA.CStr(VBA.CLng(number))
                    If numbers.Exists(number) Then Err.Raise 5, , "Duplicate scenario step: " & id
                    numbers.Add number, True
                End If
            End If
        Next row
        data.Dispose
        Set data = Nothing
        query.Dispose
        Set query = Nothing
        source.Dispose
        Set source = Nothing
    Next kind
    For Each item In catalog
        Set numbers = item("Numbers")
        If numbers.Count = 0 Then Err.Raise 5, , "Scenario has no steps: " & item("Id")
        Set steps = New Collection
        For index = 1 To numbers.Count
            If Not numbers.Exists(VBA.CStr(index)) Then Err.Raise 5, , "Scenario steps are not consecutive: " & item("Id")
            steps.Add index
        Next index
        Set item("Steps") = steps
        item.Remove "Numbers"
    Next item
    Set ordered = New Collection
    For Each kind In VBA.Array("DepartureType", "ArrivalType", "TestCustomOperation")
        For Each item In catalog
            If item("Operation") = ex_PEB_Export.fn_Config("text." & kind) Then ordered.Add item
        Next item
    Next kind
    If ordered.Count <> catalog.Count Then Err.Raise 5, , "Unknown operation in scenario catalog."
    Set catalog = ordered
    fn_ReadCatalog = True
CleanExit:
    If Not data Is Nothing Then data.Dispose
    If Not query Is Nothing Then query.Dispose
    If Not source Is Nothing Then source.Dispose
    If Not executor Is Nothing Then executor.Dispose
    ex_Core.fn_Diagnostic_WritePerf "TestCatalog.Read | Success=" & VBA.CStr(fn_ReadCatalog), started
    Exit Function
Failed:
    diagnostic = VBA.Err.Description
    Resume CleanExit
End Function

Public Function fn_Open() As Boolean
    If m_testUi Is Nothing Then
        Set m_testUi = New obj_PEB_TestUi
        If Not m_testUi.Initialize() Then Exit Function
    End If
    fn_Open = m_testUi.OpenPage()
End Function

Public Sub fn_RequestStop()
    If m_testUi Is Nothing Then Exit Sub
    If Not m_testUi.IsRunning Then Exit Sub
    m_testUi.RequestStop
    If m_advancing Then Exit Sub
    fn_CancelTick
    fn_ScheduleTick True
End Sub

Public Sub fn_ScheduleTick(Optional ByVal immediate As Boolean = False)
    Dim delay As Long

    If m_advancing Then Exit Sub
    If m_testUi Is Nothing Or m_tickPending Then Exit Sub
    If Not m_testUi.IsRunning Then Exit Sub
    delay = ex_PEB_Export.fn_Count("test.runner.tickSeconds")
    If delay < 1 Then Err.Raise 5, , "test.runner.tickSeconds must be positive."
    If immediate Then delay = 0
    m_tickAt = VBA.Now + VBA.TimeSerial(0, 0, delay)
    m_tickProcedure = private_Callback()
    Application.OnTime m_tickAt, m_tickProcedure
    m_tickPending = True
End Sub

Public Sub fn_PumpUi()
    Dim previousScreen As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    If Not m_advancing Then Exit Sub
    previousScreen = Application.ScreenUpdating
    On Error GoTo Failed
    Application.ScreenUpdating = True
    DoEvents
CleanExit:
    Application.ScreenUpdating = previousScreen
    If errorNumber <> 0 Then Err.Raise errorNumber, , errorDescription
    Exit Sub
Failed:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    Resume CleanExit
End Sub

Public Sub fn_ReportPhase(ByVal key As String)
    If m_testUi Is Nothing Then Exit Sub
    If Not m_testUi.IsRunning Then Exit Sub
    m_testUi.ReportPhase key
    private_FlushMain
    fn_PumpUi
End Sub

Public Function fn_StopRequested() As Boolean
    If m_testUi Is Nothing Then
        fn_StopRequested = True
    Else
        fn_StopRequested = m_testUi.StopRequested
    End If
End Function

Public Sub fn_CancelTick()
    If Not m_tickPending Then Exit Sub
    If m_advancing Then Exit Sub
    On Error Resume Next
    Application.OnTime m_tickAt, m_tickProcedure, , False
    On Error GoTo 0
    m_tickPending = False
End Sub

Public Sub fn_Tick()
    Dim runtimeContext As Object
    Dim previousEvents As Boolean
    Dim previousScreen As Boolean
    Dim entered As Boolean
    Dim more As Boolean
    Dim diagnostic As String

    If m_advancing Then Exit Sub
    If Not m_tickPending Then Exit Sub
    fn_CancelTick
    If m_testUi Is Nothing Then Exit Sub
    previousEvents = Application.EnableEvents
    previousScreen = Application.ScreenUpdating
    On Error GoTo EH
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then
        fn_Module_Dispose
        Exit Sub
    End If
    entered = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    m_advancing = True
    more = m_testUi.Advance()
    private_FlushMain
CleanExit:
    m_advancing = False
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreen
    If entered Then
        ex_RuntimeLifecycle.fn_Leave runtimeContext
        entered = False
    End If
    If more Then fn_ScheduleTick m_testUi.StopRequested
    If VBA.Len(diagnostic) > 0 Then ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation, ex_PEB_Export.fn_Config("text.TestTitle")
    Exit Sub
EH:
    diagnostic = VBA.Err.Description
    more = False
    fn_Module_Dispose
    Resume CleanExit
End Sub

Public Sub fn_Click()
    Dim runtimeContext As Object
    Dim caller As String
    Dim entered As Boolean
    Dim previousEvents As Boolean

    previousEvents = Application.EnableEvents
    On Error GoTo EH
    If Not Application.ActiveWorkbook Is ThisWorkbook Then Exit Sub
    Application.EnableEvents = False
    caller = VBA.CStr(Application.Caller)
    If caller = "pebt_Stop" Then
        fn_RequestStop
        private_FlushMain
        GoTo CleanExit
    End If
    If m_testUi Is Nothing Then
        fn_Open
        GoTo CleanExit
    End If
    If m_testUi.IsRunning Then GoTo CleanExit
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then GoTo CleanExit
    entered = True
    Select Case caller
        Case "pebt_All"
            m_testUi.SelectAll True
        Case "pebt_None"
            m_testUi.SelectAll False
        Case "pebt_Run"
            m_testUi.RunSelected
        Case "pebt_Refresh"
            m_testUi.OpenPage
        Case Else
            Err.Raise 5, , "Unknown test UI command: " & caller
    End Select
    private_FlushMain
CleanExit:
    Application.EnableEvents = previousEvents
    If entered Then ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH:
    ex_WindowsUi.fn_ShowMessage VBA.Err.Description, VBA.vbExclamation, ex_PEB_Export.fn_Config("text.TestTitle")
    Resume CleanExit
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_Callback() As String
    private_Callback = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!ex_PEB_Testing.fn_Tick"
End Function

Private Sub private_FlushMain()
    Dim sheet As Worksheet
    Dim context As obj_UiRenderContext
    Dim diagnostic As String

    Set sheet = ThisWorkbook.Worksheets(ex_PEB_Export.fn_Config("test.ui.mainSheet"))
    If ex_UiRuntime.fn_TryGetContext(sheet, context) Then
        If Not context.FlushLayout(diagnostic) Then Err.Raise 5, , diagnostic
    End If
End Sub