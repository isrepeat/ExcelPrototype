VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PEB_TestRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_session As obj_PEB_TestScenarios
Private m_exportService As obj_PEB_ExportService
Private m_catalog As Collection
Private m_queue As Collection
Private m_completed As Long
Private m_runId As String
Private m_currentScenario As String
Private m_failedSequence As Boolean
Private m_fatal As Boolean
Private m_cancelled As Boolean
Private m_before As Object
Private m_after As Object
Private m_values As Object

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Properties
' //
Public Property Get Catalog() As Collection
    Set Catalog = m_catalog
End Property

Public Property Get Total() As Long
    If Not m_queue Is Nothing Then Total = m_queue.Count
End Property

Public Property Get Completed() As Long
    Completed = m_completed
End Property

Public Property Get Finished() As Boolean
    If Not m_queue Is Nothing Then Finished = (m_completed >= m_queue.Count)
End Property

' //
' // API
' //
Public Function Initialize(ByRef diagnostic As String) As Boolean
    If m_isInitialized Or m_isDisposed Then Exit Function
    On Error GoTo EH
    Set m_session = New obj_PEB_TestScenarios
    If Not m_session.Initialize(diagnostic) Then GoTo Failed
    If Not m_session.TryCatalog(m_catalog, diagnostic) Then GoTo Failed
    Set m_exportService = New obj_PEB_ExportService
    If Not m_exportService.Initialize() Then Err.Raise 5, , "Cannot initialize the test export service."
    m_isInitialized = True
    Initialize = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
Failed:
    Me.Dispose
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
    If Not m_exportService Is Nothing Then m_exportService.Dispose
    If Not m_session Is Nothing Then m_session.Dispose
    Set m_exportService = Nothing
    Set m_session = Nothing
    Set m_catalog = Nothing
    Set m_queue = Nothing
    Set m_before = Nothing
    Set m_after = Nothing
    Set m_values = Nothing
End Sub

Public Function TryBegin( _
    ByVal selected As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim scenario As Object
    Dim testStep As Object
    Dim item As Object
    Dim known As Object
    Dim key As Variant

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Or Not m_queue Is Nothing Then Err.Raise 5, , "Test runner is not ready."
    If selected.Count = 0 Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.TestsSelectRequired")
    Set m_queue = New Collection
    Set known = VBA.CreateObject("Scripting.Dictionary")
    m_runId = VBA.Format$(VBA.Now, "yyyymmdd_hhnnss") & "_" & VBA.CStr(VBA.CLng(VBA.Timer * 1000#) Mod 1000)
    For Each scenario In m_catalog
        known.Add scenario("Id"), True
        If selected.Exists(scenario("Id")) Then
            For Each testStep In scenario("Steps")
                Set item = VBA.CreateObject("Scripting.Dictionary")
                item("ScenarioId") = scenario("Id")
                item("ScenarioName") = scenario("Name")
                item("RuleStatus") = scenario("RuleStatus")
                Set item("Step") = testStep
                m_queue.Add item
            Next testStep
        End If
    Next scenario
    For Each key In selected.Keys
        If Not known.Exists(key) Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.TestsRefreshRequired")
    Next key
    TryBegin = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
End Function

Public Sub RequestStop()
    m_cancelled = True
End Sub

Public Function TryAdvance( _
    ByRef report As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim item As Object
    Dim testStep As Object
    Dim reason As String
    Dim archiveDiagnostic As String

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Or m_queue Is Nothing Then Err.Raise 5, , "Test queue is not initialized."
    If m_completed >= m_queue.Count Then Err.Raise 5, , "Test queue is already complete."
    If ex_PEB_Export.fn_Config("environment.name") <> "TEST" Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.TestOnly")
    Set item = m_queue(m_completed + 1)
    Set testStep = item("Step")
    If item("ScenarioId") <> m_currentScenario Then
        m_currentScenario = item("ScenarioId")
        m_failedSequence = False
    End If
    Set report = private_Report(m_runId, item("ScenarioId"), testStep)
    report("ScenarioName") = item("ScenarioName")
    report("RuleStatus") = item("RuleStatus")
    report("ModeId") = testStep("ModeId")
    report("EventName") = testStep("EventName")
    report("PersonId") = testStep("PersonId")
    report("Operation") = ex_PEB_Export.fn_Config("export.Movement.profiles." & _
        ex_PEB_Export.fn_Config("export.Movement." & testStep("ModeId") & ".profile") & ".operation")
    report("Timestamp") = VBA.Now
    Set m_before = Nothing
    Set m_after = Nothing
    Set m_values = Nothing
    If m_fatal Then
        reason = ex_PEB_Export.fn_Config("text.TestsAborted")
    ElseIf m_cancelled Then
        reason = ex_PEB_Export.fn_Config("text.TestsCancelled")
    ElseIf m_failedSequence Then
        reason = ex_PEB_Export.fn_Config("text.TestsPreviousFailed")
    ElseIf item("RuleStatus") = "PendingRule" Or testStep("Expected") = "Pending" Then
        reason = ex_PEB_Export.fn_Config("text.TestsPendingRule")
    End If
    If VBA.Len(reason) > 0 Then
        report("Result") = "SKIP"
        report("Actual") = "Skipped"
        report("Details") = reason
    Else
        If Not private_RunStep(item("ScenarioId"), testStep, report, diagnostic) Then
            report("Result") = "ERROR"
            report("Actual") = "Error"
            report("ActualCode") = "RunnerError"
            report("Details") = diagnostic
            m_fatal = True
        End If
        m_failedSequence = (report("Result") <> "PASS")
    End If
    report("TotalSteps") = m_queue.Count
    report("CompletedSteps") = m_completed + 1
    report("RunState") = "Running"
    If m_completed + 1 = m_queue.Count Then report("RunState") = "Completed"
    If m_cancelled Then report("RunState") = "Cancelled"
    If m_fatal Then report("RunState") = "Aborted"
    If Not m_session.TryArchive(report, m_before, m_after, testStep("Input"), m_values, archiveDiagnostic) Then
        report("Result") = "ERROR"
        report("Details") = report("Details") & " " & archiveDiagnostic
        diagnostic = archiveDiagnostic
        Exit Function
    End If
    ex_PEB_Testing.fn_ReportPhase "Archive"
    m_completed = m_completed + 1
    TryAdvance = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    If Not report Is Nothing Then
        report("Result") = "ERROR"
        report("Actual") = "Error"
        report("ActualCode") = "RunnerError"
        report("Details") = diagnostic
    End If
End Function

' //
' // Private
' //
Private Function private_Report( _
    ByVal runId As String, _
    ByVal scenarioId As String, _
    ByVal testStep As Object _
) As Object
    Dim report As Object
    Dim field As Variant

    Set report = VBA.CreateObject("Scripting.Dictionary")
    For Each field In VBA.Split(ex_PEB_Export.fn_Config("test.ui.resultFields"), "|")
        report(field) = VBA.vbNullString
    Next field
    report("RunId") = runId
    report("ScenarioId") = scenarioId
    report("StepId") = testStep("Id")
    report("Expected") = testStep("Expected")
    report("ExpectedCode") = private_ExpectedCode(VBA.CStr(testStep("Code")))
    report("ExpectedOpen") = testStep("OpenCount")
    Set private_Report = report
End Function

Private Function private_RunStep( _
    ByVal scenarioId As String, _
    ByVal testStep As Object, _
    ByVal report As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim binding As obj_UiBindingContext
    Dim exporter As obj_PEB_ExptrMovement
    Dim updates As Object
    Dim values As Object
    Dim before As Object
    Dim after As Object
    Dim key As Variant
    Dim modeId As String
    Dim direction As String
    Dim status As String
    Dim actualCode As String
    Dim expectedCode As String
    Dim actualDecision As String
    Dim exportDiagnostic As String
    Dim checkDiagnostic As String
    Dim success As Boolean
    Dim unchanged As Boolean
    Dim assertion As Boolean
    Dim started As Double
    Dim elapsed As Double

    On Error GoTo EH
    started = VBA.Timer
    If ex_PEB_Testing.fn_StopRequested() Then GoTo Cancelled
    If Not m_session.TryPrepare(scenarioId, testStep("Id"), testStep("Id") = 1, modeId, updates, status, diagnostic) Then GoTo Cleanup
    ex_PEB_Testing.fn_ReportPhase "Prepare"
    ex_PEB_Testing.fn_PumpUi
    If ex_PEB_Testing.fn_StopRequested() Then GoTo Cancelled
    If Not m_session.TrySnapshot(VBA.CStr(updates("Form.PersonId")), before, diagnostic) Then GoTo Cleanup
    ex_PEB_Testing.fn_ReportPhase "Snapshot"
    ex_PEB_Testing.fn_PumpUi
    If ex_PEB_Testing.fn_StopRequested() Then GoTo Cancelled
    Set binding = New obj_UiBindingContext
    If Not binding.Initialize() Then Err.Raise 5, , "Cannot create isolated test form bindings."
    For Each key In VBA.Split(ex_PEB_Export.fn_Config("export.fields"), "|")
        If Not binding.SetValue("Form", VBA.CStr(key), VBA.vbNullString) Then Err.Raise 5, , "Cannot register test form field: " & key
    Next key
    For Each key In updates.Keys
        If Not binding.SetValue("Form", VBA.Mid$(key, 6), updates(key)) Then Err.Raise 5, , "Cannot set test form field: " & key
    Next key
    If Not binding.SetValue("Form", "ModeId", modeId) Then Err.Raise 5, , "Cannot set test mode."
    direction = ex_PEB_Export.fn_Config("modes." & modeId & ".direction")
    If Not binding.SetValue("Form", "ModeDirection", direction) Then Err.Raise 5, , "Cannot set test direction."
    If Not binding.SetValue("Form", "EventType", ex_PEB_Export.fn_Config("text." & direction & "Type")) Then Err.Raise 5, , "Cannot set test event type."
    If Not binding.SetValue("Form", "ModeForm", ex_PEB_Export.fn_Config("modes." & modeId & ".form")) Then Err.Raise 5, , "Cannot set test form layout."
    If Not m_exportService.TryPrepare(binding, "Movement", values, diagnostic) Then GoTo Cleanup
    ex_PEB_Testing.fn_ReportPhase "Inputs"
    If Not m_session.TryCommit(diagnostic, False) Then GoTo Cleanup
    If Not m_session.TrySnapshot(VBA.CStr(values("PersonId")), before, diagnostic) Then GoTo Cleanup
    ex_PEB_Testing.fn_ReportPhase "Commit"
    ex_PEB_Testing.fn_PumpUi
    If ex_PEB_Testing.fn_StopRequested() Then GoTo Cancelled

    ' Exercise the same exporter as the PEB button, in the validated private TEST workbook.
    Set exporter = New obj_PEB_ExptrMovement
    If Not exporter.Initialize() Then Err.Raise 5, , "Cannot initialize the Movement exporter."
    If ex_PEB_Testing.fn_StopRequested() Then GoTo Cancelled
    success = exporter.TryExport(values, exportDiagnostic, actualCode, m_session.TargetWorkbook)
    ex_PEB_Testing.fn_ReportPhase "Export"
    If Not m_session.TrySnapshot(VBA.CStr(values("PersonId")), after, diagnostic, True) Then GoTo Cleanup
    ex_PEB_Testing.fn_ReportPhase "AfterSnapshot"
    unchanged = private_Unchanged(before, after)
    actualDecision = "Block"
    If success Then
        actualDecision = "Allow"
    ElseIf actualCode = "DuplicateEvent" And unchanged Then
        actualDecision = "NoChange"
    End If
    expectedCode = private_ExpectedCode(VBA.CStr(testStep("Code")))
    assertion = (actualDecision = testStep("Expected") And actualCode = expectedCode)
    If Not VBA.IsNumeric(testStep("OpenCount")) Then Err.Raise 5, , "ExpectedOpenCount must be numeric: " & scenarioId
    assertion = assertion And (VBA.CDbl(testStep("OpenCount")) = VBA.CDbl(after("OpenCount")))
    If success Then
        assertion = private_ValidateMutation(before, after, values, checkDiagnostic) And assertion
    Else
        assertion = unchanged And assertion
        If Not unchanged Then checkDiagnostic = ex_PEB_Export.fn_Config("text.TestsUnexpectedMutation")
    End If
    report("Actual") = actualDecision
    report("ActualCode") = actualCode
    report("ActualOpen") = after("OpenCount")
    report("Result") = "FAIL"
    If assertion Then report("Result") = "PASS"
    report("Details") = exportDiagnostic
    If VBA.Len(checkDiagnostic) > 0 Then report("Details") = report("Details") & " " & checkDiagnostic
    If Not assertion And VBA.Len(report("Details")) = 0 Then report("Details") = ex_PEB_Export.fn_Config("text.TestsMismatch")
    ex_PEB_Testing.fn_ReportPhase "Verify"
    private_RunStep = True
    GoTo Cleanup
Cancelled:
    m_cancelled = True
    report("Result") = "SKIP"
    report("Actual") = "Skipped"
    report("Details") = ex_PEB_Export.fn_Config("text.TestsCancelled")
    private_RunStep = True
Cleanup:
    elapsed = VBA.Timer - started
    If elapsed < 0 Then elapsed = elapsed + 86400#
    report("Milliseconds") = VBA.CLng(elapsed * 1000#)
    Set m_before = before
    Set m_after = after
    Set m_values = values
    If Not exporter Is Nothing Then exporter.Dispose
    If Not binding Is Nothing Then binding.Dispose
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function

Private Function private_ExpectedCode(ByVal code As String) As String
    Dim mapping As Variant
    Dim pair As Variant

    private_ExpectedCode = code
    For Each mapping In VBA.Split(ex_PEB_Export.fn_Config("test.runner.codeAliases"), "|")
        pair = VBA.Split(mapping, "=", 2)
        If UBound(pair) <> 1 Then Err.Raise 5, , "Invalid test result code alias."
        If pair(0) = code Then private_ExpectedCode = pair(1)
    Next mapping
End Function

Private Function private_Unchanged( _
    ByVal before As Object, _
    ByVal after As Object _
) As Boolean
    Dim oldRows As Collection
    Dim newRows As Collection
    Dim index As Long

    Set oldRows = before("Rows")
    Set newRows = after("Rows")
    If before("Integrity") <> after("Integrity") Then Exit Function
    If oldRows.Count <> newRows.Count Then Exit Function
    For index = 1 To oldRows.Count
        If oldRows(index)("_signature") <> newRows(index)("_signature") Then Exit Function
    Next index
    private_Unchanged = True
End Function

Private Function private_ValidateMutation( _
    ByVal before As Object, _
    ByVal after As Object, _
    ByVal values As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim oldRows As Collection
    Dim newRows As Collection
    Dim changedRow As Object
    Dim prefix As String
    Dim operation As String
    Dim index As Long
    Dim changed As Long
    Dim field As String
    Dim header As String
    Dim expected As Variant
    Dim expectedText As String
    Dim oldChangedRow As Object
    Dim allowed As Object
    Dim key As Variant

    Set oldRows = before("Rows")
    Set newRows = after("Rows")
    prefix = "export.Movement.profiles." & ex_PEB_Export.fn_Config("export.Movement." & values("ModeId") & ".profile")
    operation = ex_PEB_Export.fn_Config(prefix & ".operation")
    If operation = "append" Then
        If newRows.Count <> oldRows.Count + 1 Then GoTo Mismatch
        For index = 1 To oldRows.Count
            If oldRows(index)("_signature") <> newRows(index)("_signature") Then GoTo Mismatch
        Next index
        Set changedRow = newRows(newRows.Count)
        If after("OpenCount") <> before("OpenCount") + 1 Then GoTo Mismatch
    ElseIf operation = "close" Then
        If newRows.Count <> oldRows.Count Then GoTo Mismatch
        For index = 1 To oldRows.Count
            If oldRows(index)("_signature") <> newRows(index)("_signature") Then
                changed = changed + 1
                Set changedRow = newRows(index)
                Set oldChangedRow = oldRows(index)
            End If
        Next index
        If changed <> 1 Then GoTo Mismatch
        If after("OpenCount") <> before("OpenCount") - 1 Then GoTo Mismatch
        For index = 1 To ex_PEB_Export.fn_Count(prefix & ".join.count")
            field = ex_PEB_Export.fn_Config(prefix & ".join." & VBA.CStr(index) & ".field")
            header = ex_PEB_Export.fn_Config(prefix & ".join." & VBA.CStr(index) & ".header")
            If Not changedRow.Exists(header) Then GoTo Mismatch
            If changedRow(header) <> ex_PEB_Export.fn_Text(values(field)) Then GoTo Mismatch
        Next index
    Else
        Err.Raise 5, , "Unknown Movement test operation: " & operation
    End If
    Set allowed = VBA.CreateObject("Scripting.Dictionary")
    For index = 1 To ex_PEB_Export.fn_Count(prefix & ".fields.count")
        field = ex_PEB_Export.fn_Config(prefix & ".fields." & VBA.CStr(index) & ".field")
        header = ex_PEB_Export.fn_Config(prefix & ".fields." & VBA.CStr(index) & ".header")
        allowed(header) = True
        expected = values(field)
        If VBA.VarType(expected) = VBA.vbDate Then
            expectedText = VBA.CStr(VBA.CDbl(expected))
        Else
            expectedText = ex_PEB_Export.fn_Text(expected)
        End If
        If Not changedRow.Exists(header) Then GoTo Mismatch
        If changedRow(header) <> expectedText Then GoTo Mismatch
    Next index
    If Not oldChangedRow Is Nothing Then
        For Each key In oldChangedRow.Keys
            If key <> "_signature" And key <> "_cells" And Not allowed.Exists(key) Then
                If Not changedRow.Exists(key) Then GoTo Mismatch
                If oldChangedRow(key) <> changedRow(key) Then GoTo Mismatch
            End If
        Next key
    End If
    private_ValidateMutation = True
    Exit Function
Mismatch:
    diagnostic = ex_PEB_Export.fn_Config("text.TestsMutationMismatch")
End Function