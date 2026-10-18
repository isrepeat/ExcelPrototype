Attribute VB_Name = "ex_PEB_Testing"
Option Explicit

Private m_testController As obj_PEB_PgTestsController
Private m_tickAt As Date
Private m_tickPending As Boolean
Private m_tickProcedure As String
Private m_advancing As Boolean
Private m_lastPhaseFlush As Double

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    fn_CancelTick
    Set m_testController = Nothing
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_IsRunning() As Boolean
    If Not m_testController Is Nothing Then
        fn_IsRunning = m_testController.IsRunning
    End If
End Function

Public Function fn_RegistryValues(ByVal kind As String) As Variant
    Dim values As Variant
    Dim value As Variant
    Dim seen As Object

    If kind = "ModeId" Then
        values = ex_PEB_EventRoutes.fn_Ids()
    Else
        values = VBA.Split(ex_PEB_Export.fn_Config("test.registry." & kind & ".values"), "|")
    End If
    Set seen = VBA.CreateObject("Scripting.Dictionary")
    For Each value In values
        If VBA.Len(value) = 0 Or value <> VBA.Trim$(value) Then Err.Raise 5, , "Invalid registry value: " & kind
        If seen.Exists(value) Then Err.Raise 5, , "Duplicate registry value: " & kind & "/" & value
        seen.Add value, True
    Next value
    fn_RegistryValues = values
End Function

Public Function fn_RegistryContains( _
    ByVal kind As String, _
    ByVal value As String _
) As Boolean
    Dim candidate As Variant

    For Each candidate In fn_RegistryValues(kind)
        If VBA.StrComp(candidate, value, VBA.vbBinaryCompare) = 0 Then
            fn_RegistryContains = True
            Exit Function
        End If
    Next candidate
End Function

Public Function fn_ResolveTestMode(ByVal templateId As String) As String
    If Not fn_RegistryContains("ModeId", templateId) Then
        Err.Raise 5, , "Unknown XML business event: " & templateId
    End If
    fn_ResolveTestMode = templateId
End Function

Public Function fn_ReadCatalog( _
    ByRef catalog As Collection, _
    ByRef diagnostic As String, _
    Optional ByVal suite As String = "Basic" _
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
    Dim stepIds As Object
    Dim stepId As String
    Dim action As String
    Dim allCases As Object
    Dim stepKind As String

    started = VBA.Timer
    On Error GoTo Failed
    Set catalog = New Collection
    Set cases = VBA.CreateObject("Scripting.Dictionary")
    Set allCases = VBA.CreateObject("Scripting.Dictionary")
    Set stepIds = VBA.CreateObject("Scripting.Dictionary")
    stepKind = "steps"
    If suite = "Grouping" Then
        stepKind = "groupingSteps"
    ElseIf suite <> "Basic" Then
        Err.Raise 5, , "Unknown test suite: " & suite
    End If
    Set executor = New obj_AdoQueryExecutor
    If Not executor.Initialize() Then Err.Raise 5, , "Cannot initialize catalog reader."
    For Each kind In VBA.Array("cases", stepKind)
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
            query.AddColumn ex_PEB_Export.fn_Config("test.scenarios.headers.suite")
        Else
            query.AddColumn ex_PEB_Export.fn_Config("test.scenarios.headers.step")
            query.AddColumn ex_PEB_Export.fn_Config("test.scenarios.headers.stepId")
            query.AddColumn ex_PEB_Export.fn_Config("test.scenarios.headers.action")
        End If
        If Not executor.TryExecute(source, query, data, diagnostic) Then GoTo CleanExit
        For row = 1 To data.RowCount
            id = ex_PEB_Export.fn_Text(data.ValueAt(row, 1))
            If VBA.Len(id) > 0 Then
                If kind = "cases" Then
                    If allCases.Exists(id) Then
                        Err.Raise 5, , "Duplicate scenario identity: " & id
                    End If
                    allCases.Add id, ex_PEB_Export.fn_Text(data.ValueAt(row, 6))
                    If allCases(id) <> "Basic" And allCases(id) <> "Grouping" Then
                        Err.Raise 5, , "Unknown scenario suite: " & id
                    End If
                    If allCases(id) <> suite Then
                        GoTo NextCatalogRow
                    End If
                    If cases.Exists(id) Then Err.Raise 5, , "Duplicate scenario identity: " & id
                    Set item = VBA.CreateObject("Scripting.Dictionary")
                    item("Id") = id
                    item("Name") = ex_PEB_Export.fn_Text(data.ValueAt(row, 2))
                    item("RuleStatus") = ex_PEB_Export.fn_Text(data.ValueAt(row, 3))
                    item("Operation") = ex_PEB_Export.fn_Text(data.ValueAt(row, 4))
                    item("EventLabel") = ex_PEB_Export.fn_Text(data.ValueAt(row, 5))
                    If VBA.Len(item("Operation")) = 0 Or VBA.Len(item("EventLabel")) = 0 Then Err.Raise 5, , "Scenario classification is missing: " & id
                    If Not fn_RegistryContains("RuleStatus", item("RuleStatus")) Then Err.Raise 5, , "Invalid scenario rule status: " & id
                    If Not fn_RegistryContains("Operation", item("Operation")) Then Err.Raise 5, , "Invalid scenario operation: " & id
                    Set numbers = VBA.CreateObject("Scripting.Dictionary")
                    Set item("Numbers") = numbers
                    cases.Add id, item
                    catalog.Add item
                Else
                    If Not allCases.Exists(id) Then Err.Raise 5, , "Step references an unknown scenario: " & id
                    If Not cases.Exists(id) Then
                        Err.Raise 5, , "Step belongs to a different test suite: " & id
                    End If
                    number = ex_PEB_Export.fn_Text(data.ValueAt(row, 2))
                    If Not VBA.IsNumeric(number) Then Err.Raise 5, , "Invalid scenario step: " & id
                    If VBA.CDbl(number) < 1 Or VBA.CDbl(number) <> VBA.Fix(VBA.CDbl(number)) Then Err.Raise 5, , "Invalid scenario step: " & id
                    Set numbers = cases(id)("Numbers")
                    number = VBA.CStr(VBA.CLng(number))
                    If numbers.Exists(number) Then Err.Raise 5, , "Duplicate scenario step: " & id
                    stepId = ex_PEB_Export.fn_Text(data.ValueAt(row, 3))
                    If VBA.Len(stepId) = 0 Or stepIds.Exists(stepId) Then Err.Raise 5, , "Missing or duplicate StepId: " & stepId
                    action = ex_PEB_Export.fn_Text(data.ValueAt(row, 4))
                    If suite = "Grouping" Then
                        If action <> "GroupAdd" And action <> "GroupEdit" Then
                            Err.Raise 5, , "Unknown grouping action: " & stepId
                        End If
                    ElseIf Not fn_RegistryContains("Action", action) Then
                        Err.Raise 5, , "Unknown test action: " & stepId
                    End If
                    stepIds.Add stepId, True
                    numbers.Add number, stepId
                End If
            End If
NextCatalogRow:
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
            steps.Add numbers(VBA.CStr(index))
        Next index
        Set item("Steps") = steps
        item.Remove "Numbers"
    Next item
    Set ordered = New Collection
    For Each kind In VBA.Array("DepartureType", "ArrivalType", "TestCustomOperation", "TestPreviewOperation")
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

Public Sub fn_Attach(ByVal controller As obj_PEB_PgTestsController)
    If controller Is Nothing Then
        Err.Raise 5, , "Test page controller is missing."
    End If
    If Not m_testController Is Nothing Then
        If Not m_testController Is controller Then
            Err.Raise 5, , "Another test page controller is attached."
        End If
    End If
    Set m_testController = controller
End Sub

Public Sub fn_Detach(ByVal controller As obj_PEB_PgTestsController)
    If m_testController Is Nothing Then
        Exit Sub
    End If
    If Not m_testController Is controller Then
        Exit Sub
    End If
    fn_CancelTick
    Set m_testController = Nothing
End Sub

Public Sub fn_RequestStop()
    If m_testController Is Nothing Then
        Exit Sub
    End If
    If Not m_testController.IsRunning Then
        Exit Sub
    End If
    m_testController.RequestStop
    If m_advancing Then Exit Sub
    fn_CancelTick
    fn_ScheduleTick True
End Sub

Public Sub fn_ScheduleTick(Optional ByVal immediate As Boolean = False)
    Dim delay As Long

    If m_advancing Then Exit Sub
    If m_testController Is Nothing Or m_tickPending Then
        Exit Sub
    End If
    If Not m_testController.IsRunning Then
        Exit Sub
    End If
    delay = ex_PEB_Export.fn_Count("test.runner.tickSeconds")
    If delay < 0 Then Err.Raise 5, , "test.runner.tickSeconds must not be negative."
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
    ex_UiBusy.fn_Refresh
    If errorNumber <> 0 Then Err.Raise errorNumber, , errorDescription
    Exit Sub
Failed:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    Resume CleanExit
End Sub

Public Sub fn_ReportPhase( _
    ByVal key As String, _
    Optional ByVal completedUnits As Long = 0, _
    Optional ByVal totalUnits As Long = 0 _
)
    Dim elapsed As Double
    Dim nowAt As Double

    If m_testController Is Nothing Then
        Exit Sub
    End If
    If Not m_testController.IsRunning Then
        Exit Sub
    End If
    m_testController.ReportPhase key, completedUnits, totalUnits
    nowAt = VBA.Timer
    elapsed = nowAt - m_lastPhaseFlush
    If elapsed < 0 Then elapsed = elapsed + 86400#
    If totalUnits = 0 Then
        If elapsed * 1000# < ex_PEB_Export.fn_Count("test.runner.uiRefreshMilliseconds") Then Exit Sub
    End If
    m_lastPhaseFlush = nowAt
    private_FlushMain
    fn_PumpUi
End Sub

Public Function fn_StopRequested() As Boolean
    If m_testController Is Nothing Then
        fn_StopRequested = True
    Else
        fn_StopRequested = m_testController.StopRequested
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
    Dim started As Double
    Dim elapsed As Double
    Dim budget As Long
    Dim configEntered As Boolean
    Dim busyEntered As Boolean

    If m_advancing Then Exit Sub
    If Not m_tickPending Then Exit Sub
    fn_CancelTick
    If m_testController Is Nothing Then
        Exit Sub
    End If
    previousEvents = Application.EnableEvents
    previousScreen = Application.ScreenUpdating
    On Error GoTo EH
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then
        fn_Module_Dispose
        Exit Sub
    End If
    entered = True
    ' Scheduled callbacks run after the initiating click has released its busy scope.
    busyEntered = ex_UiBusy.fn_Begin()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    m_advancing = True
    ex_PEB_Export.fn_BeginConfigSnapshot
    configEntered = True
    budget = ex_PEB_Export.fn_Count("test.runner.batchMilliseconds")
    If budget < 1 Then Err.Raise 5, , "test.runner.batchMilliseconds must be positive."
    started = VBA.Timer
    Do
        more = m_testController.Advance()
        If Not more Then Exit Do
        fn_PumpUi
        If m_testController.StopRequested Then
            Exit Do
        End If
        elapsed = VBA.Timer - started
        If elapsed < 0 Then elapsed = elapsed + 86400#
    Loop While elapsed * 1000# < budget
    private_FlushMain
CleanExit:
    If configEntered Then ex_PEB_Export.fn_EndConfigSnapshot
    m_advancing = False
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreen
    If busyEntered Then ex_UiBusy.fn_End
    If entered Then
        ex_RuntimeLifecycle.fn_Leave runtimeContext
        entered = False
    End If
    If more Then
        fn_ScheduleTick m_testController.StopRequested
    End If
    If VBA.Len(diagnostic) > 0 Then ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation, ex_PEB_Export.fn_Config("text.TestTitle")
    Exit Sub
EH:
    diagnostic = VBA.Err.Description
    more = False
    If Not m_testController Is Nothing Then
        m_testController.FailRun diagnostic
        diagnostic = VBA.vbNullString
    End If
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