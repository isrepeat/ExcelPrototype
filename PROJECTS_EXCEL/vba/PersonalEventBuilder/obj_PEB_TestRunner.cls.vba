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
Private m_testWordSession As obj_PEB_TestWordSession

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

Public Property Get WordDocumentPath() As String
    If Not m_testWordSession Is Nothing Then WordDocumentPath = m_testWordSession.OutputPath
End Property

' //
' // API
' //
Public Function Initialize(ByRef diagnostic As String) As Boolean
    If m_isInitialized Or m_isDisposed Then Exit Function
    On Error GoTo EH
    Set m_session = New obj_PEB_TestScenarios
    If Not m_session.Initialize(diagnostic) Then GoTo Failed
    ex_PEB_Testing.fn_ReportPhase "Catalog"
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
    If Not m_testWordSession Is Nothing Then m_testWordSession.Dispose
    Set m_testWordSession = Nothing
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

Public Function TryCheckpoint(ByRef diagnostic As String) As Boolean
    Dim wordDiagnostic As String
    Dim success As Boolean
    Dim runState As String

    If Not m_isInitialized Or m_isDisposed Then
        TryCheckpoint = True
        Exit Function
    End If
    success = m_session.TryCheckpoint(diagnostic)
    If Not m_testWordSession Is Nothing Then
        runState = "Interrupted"
        If Me.Finished Then runState = "Completed"
        If m_cancelled Then runState = "Cancelled"
        If m_fatal Then runState = "Aborted"
        If Not m_testWordSession.TryFinish(runState, wordDiagnostic) Then
            diagnostic = diagnostic & " " & wordDiagnostic
            success = False
        End If
    End If
    TryCheckpoint = success
End Function

Public Function TryBegin( _
    ByVal selected As Object, _
    ByRef diagnostic As String, _
    Optional ByVal createWordReport As Boolean = False _
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
                If VBA.VarType(selected(scenario("Id"))) <> VBA.vbBoolean Then Err.Raise 5, , "Invalid preview selection."
                item("PreviewEnabled") = VBA.CBool(selected(scenario("Id")))
                Set item("Step") = testStep
                m_queue.Add item
            Next testStep
        End If
    Next scenario
    For Each key In selected.Keys
        If Not known.Exists(key) Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.TestsReopenRequired")
    Next key
    If createWordReport Then
        Set m_testWordSession = New obj_PEB_TestWordSession
        If Not m_testWordSession.Initialize(m_runId, diagnostic) Then
            Set m_testWordSession = Nothing
            Exit Function
        End If
    End If
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
    Dim checkpoint As Boolean
    Dim nextItem As Object

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
    report("ModeId") = testStep("FixtureModeId")
    report("EventName") = testStep("EventName")
    report("PersonId") = testStep("PersonId")
    report("Action") = testStep("Action")
    report("PreviewEnabled") = item("PreviewEnabled")
    If Not report("PreviewEnabled") Then report("PreviewLink") = VBA.vbNullString
    report("StepOrder") = testStep("Order")
    report("Operation") = "preview"
    If testStep("Action") = "ExportMovement" Or testStep("Action") = "ExportEvent" Then
        report("Operation") = ex_PEB_Export.fn_Config("export.Movement.profiles." & _
            ex_PEB_EventRoutes.fn_Value(testStep("ModeId"), "movement.profile") & ".operation")
        report("StateCheck") = testStep("StateCheck")
    End If
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
        If testStep("Action") = "PreviewWord" Then
            report("WordResult") = "SKIP"
        Else
            report("MovementResult") = "SKIP"
            If testStep("Action") = "ExportEvent" Then
                report("WordResult") = "SKIP"
            End If
        End If
        report("Actual") = "Skipped"
        report("Details") = reason
    Else
        If Not private_RunStep(item("ScenarioId"), testStep, report, diagnostic) Then
            report("Result") = "ERROR"
            If testStep("Action") = "PreviewWord" Then
                report("WordResult") = "ERROR"
            Else
                report("MovementResult") = "ERROR"
                If testStep("Action") = "ExportEvent" Then
                    report("WordResult") = "ERROR"
                End If
            End If
            report("Actual") = "Error"
            report("ActualCode") = "RunnerError"
            report("Details") = diagnostic
            m_fatal = True
        End If
        m_failedSequence = (report("Result") <> "PASS")
    End If
    If report("Result") <> "PASS" Then m_failedSequence = True
    report("TotalSteps") = m_queue.Count
    report("CompletedSteps") = m_completed + 1
    report("RunState") = "Running"
    If m_completed + 1 = m_queue.Count Then report("RunState") = "Completed"
    If m_cancelled Then report("RunState") = "Cancelled"
    If m_fatal Then report("RunState") = "Aborted"
    If Not m_testWordSession Is Nothing Then
        If Not m_testWordSession.TryAppend(report, diagnostic) Then
            report("Result") = "ERROR"
            report("Details") = report("Details") & " " & diagnostic
            Exit Function
        End If
    End If
    checkpoint = m_cancelled Or m_fatal Or m_completed + 1 = m_queue.Count
    If Not checkpoint Then
        Set nextItem = m_queue(m_completed + 2)
        checkpoint = (nextItem("ScenarioId") <> item("ScenarioId"))
    End If
    If Not m_session.TryArchive(report, m_before, m_after, testStep("Input"), m_values, archiveDiagnostic, checkpoint) Then
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
    report("PreviewText") = ex_PEB_Export.fn_Config("text.TestPreviewNotRun")
    report("PreviewLink") = ex_PEB_Export.fn_Config("text.TestPreviewLink")
    report("MovementResult") = "NotApplicable"
    report("WordResult") = "NotApplicable"
    report("WordCreateMs") = VBA.vbNullString
    report("WordInsertMs") = VBA.vbNullString
    report("WordVerifyMs") = VBA.vbNullString
    report("WordAppendMs") = VBA.vbNullString
    report("WordDocumentPath") = VBA.vbNullString
    If testStep("Action") = "PreviewWord" Then
        report("WordResult") = "NotRun"
        report("ExpectedPreviewText") = testStep("ExpectedText")
        report("ExpectedPreviewError") = testStep("ExpectedError")
    Else
        report("MovementResult") = "NotRun"
        If testStep("Action") = "ExportEvent" Then
            report("WordResult") = "NotRun"
            report("ExpectedPreviewText") = testStep("WordExpectation")("ExpectedText")
            report("ExpectedPreviewError") = testStep("WordExpectation")("ExpectedError")
        End If
    End If
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
    Dim previewValues As Object
    Dim previewDiagnostic As String
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
    Dim wordReport As Object
    Dim wordPassed As Boolean
    Dim wordAfter As Object

    On Error GoTo EH
    started = VBA.Timer
    If ex_PEB_Testing.fn_StopRequested() Then GoTo Cancelled
    If Not m_session.TryPrepare(scenarioId, testStep("Order"), testStep("Order") = 1, modeId, updates, status, diagnostic) Then GoTo Cleanup
    ex_PEB_Testing.fn_ReportPhase "Prepare"
    If ex_PEB_Testing.fn_StopRequested() Then GoTo Cancelled
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
    direction = ex_PEB_EventRoutes.fn_Value(modeId, "direction")
    If Not binding.SetValue("Form", "ModeDirection", direction) Then Err.Raise 5, , "Cannot set test direction."
    If Not binding.SetValue("Form", "EventType", ex_PEB_Export.fn_Config("text." & direction & "Type")) Then Err.Raise 5, , "Cannot set test event type."
    If Not binding.SetValue("Form", "ModeForm", ex_PEB_EventRoutes.fn_Value(modeId, "form")) Then
        Err.Raise 5, , "Cannot set test form layout."
    End If
    If testStep("Action") = "PreviewWord" Then
        ex_PEB_Testing.fn_ReportPhase "Inputs"
        If Not m_session.TryCommit(diagnostic, False) Then GoTo Cleanup
        If Not m_session.TrySnapshot(VBA.CStr(updates("Form.PersonId")), before, diagnostic) Then GoTo Cleanup
        ex_PEB_Testing.fn_ReportPhase "Commit"
        If ex_PEB_Testing.fn_StopRequested() Then GoTo Cancelled
        assertion = private_ValidateWord(binding, testStep, report, values)
        ex_PEB_Testing.fn_ReportPhase "Export"
        If Not m_session.TrySnapshot(VBA.CStr(updates("Form.PersonId")), after, diagnostic) Then GoTo Cleanup
        ex_PEB_Testing.fn_ReportPhase "AfterSnapshot"
        If Not private_Unchanged(before, after) Then
            assertion = False
            report("Details") = report("Details") & " " & ex_PEB_Export.fn_Config("text.TestsUnexpectedMutation")
        End If
        report("Result") = "FAIL"
        If assertion Then report("Result") = "PASS"
        report("WordResult") = report("Result")
        ex_PEB_Testing.fn_ReportPhase "Verify"
        private_RunStep = True
        GoTo Cleanup
    End If
    If Not m_exportService.TryPrepare(binding, "Movement", values, diagnostic) Then GoTo Cleanup
    ex_PEB_Testing.fn_ReportPhase "Inputs"
    If (report("PreviewEnabled") Or Not m_testWordSession Is Nothing) And testStep("Action") <> "ExportEvent" Then
        If m_exportService.TryPrepare(binding, "Word", previewValues, previewDiagnostic, False) Then
            report("PreviewText") = VBA.Replace$(VBA.CStr(previewValues("WordText")), VBA.vbCrLf, VBA.vbLf)
            If Not m_testWordSession Is Nothing Then
                report("WordResult") = "Omitted"
                If previewValues("WordOutcome") = "Success" Then
                    report("WordResult") = "PASS"
                    If Not m_testWordSession.TryExercise(previewValues, previewDiagnostic) Then
                        report("WordResult") = "FAIL"
                        report("PreviewDiagnostic") = previewDiagnostic
                    End If
                End If
            End If
        Else
            report("PreviewText") = ex_PEB_Export.fn_Config("text.TestPreviewError") & VBA.vbLf & previewDiagnostic
            If Not m_testWordSession Is Nothing Then
                report("WordResult") = "FAIL"
                report("PreviewDiagnostic") = previewDiagnostic
            End If
        End If
    End If
    If Not m_session.TryCommit(diagnostic, False) Then GoTo Cleanup
    If Not m_session.TrySnapshot(VBA.CStr(values("PersonId")), before, diagnostic) Then GoTo Cleanup
    ex_PEB_Testing.fn_ReportPhase "Commit"
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
    If testStep("Action") = "ExportEvent" Then
        Set wordReport = VBA.CreateObject("Scripting.Dictionary")
        For Each key In report.Keys
            wordReport(key) = report(key)
        Next key
        wordPassed = private_ValidateWord(binding, testStep("WordExpectation"), wordReport, previewValues)
        If Not m_session.TrySnapshot(VBA.CStr(values("PersonId")), wordAfter, diagnostic) Then
            GoTo Cleanup
        End If
        If Not private_Unchanged(after, wordAfter) Then
            wordPassed = False
            wordReport("Details") = wordReport("Details") & " " & ex_PEB_Export.fn_Config("text.TestsUnexpectedMutation")
        End If
        Set after = wordAfter
        report("WordResult") = "FAIL"
        If wordPassed Then
            report("WordResult") = "PASS"
        End If
        report("PreviewText") = wordReport("PreviewText")
        report("PreviewDiagnostic") = wordReport("PreviewDiagnostic")
    End If
    If Not VBA.IsNumeric(testStep("OpenCount")) Then Err.Raise 5, , "ExpectedOpenCount must be numeric: " & scenarioId
    assertion = assertion And (VBA.CDbl(testStep("OpenCount")) = VBA.CDbl(after("OpenCount")))
    If success Then
        assertion = private_ValidateMutation(before, after, values, checkDiagnostic) And assertion
    End If
    If Not success Or testStep("StateCheck") = "Unchanged" Then
        assertion = unchanged And assertion
        If Not unchanged Then private_AddDetail checkDiagnostic, ex_PEB_Export.fn_Config("text.TestsUnexpectedMutation")
    End If
    If testStep("StateCheck") = "Snapshot" Then
        assertion = private_ValidateRows(after, testStep, checkDiagnostic) And assertion
    End If
    report("Actual") = actualDecision
    report("ActualCode") = actualCode
    report("ActualOpen") = after("OpenCount")
    If actualDecision <> testStep("Expected") Then
        private_AddDetail checkDiagnostic, private_Mismatch(testStep("Id"), "РУХ", "Decision", testStep("Expected"), actualDecision)
    End If
    If actualCode <> expectedCode Then
        private_AddDetail checkDiagnostic, private_Mismatch(testStep("Id"), "РУХ", "Code", expectedCode, actualCode)
    End If
    If VBA.CDbl(testStep("OpenCount")) <> VBA.CDbl(after("OpenCount")) Then
        private_AddDetail checkDiagnostic, private_Mismatch(testStep("Id"), "РУХ", "OpenCount", VBA.CStr(testStep("OpenCount")), VBA.CStr(after("OpenCount")))
    End If
    report("MovementResult") = "FAIL"
    If assertion Then
        report("MovementResult") = "PASS"
    End If
    If Not wordReport Is Nothing Then
        assertion = assertion And wordPassed
    End If
    If report("WordResult") = "FAIL" Then assertion = False
    report("Result") = "FAIL"
    If assertion Then report("Result") = "PASS"
    report("Details") = exportDiagnostic
    If testStep("Action") = "ExportMovement" And report("WordResult") = "FAIL" Then
        report("Details") = report("Details") & " " & report("PreviewDiagnostic")
    End If
    If Not wordReport Is Nothing Then
        If Not wordPassed Then
            report("Details") = report("Details") & " " & wordReport("Details")
        End If
    End If
    If VBA.Len(checkDiagnostic) > 0 Then report("Details") = report("Details") & " " & checkDiagnostic
    If Not assertion And VBA.Len(report("Details")) = 0 Then report("Details") = ex_PEB_Export.fn_Config("text.TestsMismatch")
    ex_PEB_Testing.fn_ReportPhase "Verify"
    private_RunStep = True
    GoTo Cleanup
Cancelled:
    m_cancelled = True
    report("Result") = "SKIP"
    If testStep("Action") = "PreviewWord" Then
        report("WordResult") = "SKIP"
    Else
        report("MovementResult") = "SKIP"
        If testStep("Action") = "ExportEvent" Then
            report("WordResult") = "SKIP"
        End If
    End If
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

Private Function private_ValidateRows( _
    ByVal state As Object, _
    ByVal testStep As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim expectedRows As Collection
    Dim actualRows As Collection
    Dim expectedRow As Object
    Dim actualRow As Object
    Dim actualCells As Object
    Dim actualByKey As Object
    Dim matched As Object
    Dim key As String
    Dim field As Variant
    Dim result As Boolean

    Set expectedRows = testStep("ExpectedRows")
    Set actualRows = state("Rows")
    Set actualByKey = VBA.CreateObject("Scripting.Dictionary")
    Set matched = VBA.CreateObject("Scripting.Dictionary")
    result = True
    If actualRows.Count <> expectedRows.Count Then
        private_AddDetail diagnostic, private_Mismatch(testStep("Id"), "Rows", "Count", VBA.CStr(expectedRows.Count), VBA.CStr(actualRows.Count))
        result = False
    End If
    For Each actualRow In actualRows
        Set actualCells = actualRow("_cells")
        key = private_RowIdentity(actualCells)
        If actualByKey.Exists(key) Then
            private_AddDetail diagnostic, private_Mismatch(testStep("Id"), "Rows", "Identity", "Unique", key)
            result = False
        Else
            actualByKey.Add key, actualCells
        End If
    Next actualRow
    For Each expectedRow In expectedRows
        key = private_RowIdentity(expectedRow)
        If Not actualByKey.Exists(key) Then
            private_AddDetail diagnostic, private_Mismatch(testStep("Id"), expectedRow("RecordId"), "Identity", key, "Missing")
            result = False
        Else
            matched(key) = True
            Set actualCells = actualByKey(key)
            For Each field In expectedRow.Keys
                If field <> "RecordId" Then
                    If Not private_SameValue(expectedRow(field), actualCells(field)) Then
                        private_AddDetail diagnostic, private_Mismatch(testStep("Id"), expectedRow("RecordId"), VBA.CStr(field), _
                            private_ValueText(expectedRow(field)), private_ValueText(actualCells(field)))
                        If VBA.TypeName(expectedRow(field)) <> VBA.TypeName(actualCells(field)) Then
                            private_AddDetail diagnostic, private_Mismatch(testStep("Id"), expectedRow("RecordId"), "Type(" & field & ")", _
                                VBA.TypeName(expectedRow(field)), VBA.TypeName(actualCells(field)))
                        End If
                        result = False
                    End If
                End If
            Next field
        End If
    Next expectedRow
    For Each actualRow In actualRows
        Set actualCells = actualRow("_cells")
        key = private_RowIdentity(actualCells)
        If Not matched.Exists(key) Then
            private_AddDetail diagnostic, private_Mismatch(testStep("Id"), "Rows", "Identity", "Absent", key)
            result = False
        End If
    Next actualRow
    private_ValidateRows = result
End Function

Private Function private_ValidateWord( _
    ByVal binding As obj_UiBindingContext, _
    ByVal testStep As Object, _
    ByVal report As Object, _
    ByRef values As Object _
) As Boolean
    Dim diagnostic As String
    Dim success As Boolean
    Dim expected As String
    Dim actual As String
    Dim position As Long
    Dim line As Long
    Dim column As Long
    Dim index As Long
    Dim details As String

    success = m_exportService.TryPrepare(binding, "Word", values, diagnostic, False)
    report("Actual") = "Error"
    report("ActualCode") = "PreviewError"
    report("PreviewDiagnostic") = diagnostic
    report("PreviewText") = VBA.vbNullString
    If success Then
        report("Actual") = values("WordOutcome")
        report("ActualCode") = VBA.vbNullString
        actual = private_NormalizeNewlines(VBA.CStr(values("WordText")))
        report("PreviewText") = actual
        If Not m_testWordSession Is Nothing And values("WordOutcome") = "Success" Then
            If Not m_testWordSession.TryExercise(values, diagnostic) Then
                report("PreviewDiagnostic") = diagnostic
                report("Details") = diagnostic
                Exit Function
            End If
        End If
    End If
    If report("Actual") <> testStep("Expected") Then
        report("Details") = private_Mismatch(testStep("Id"), "WORD", "Outcome", testStep("Expected"), report("Actual"))
        If Not success Then
            details = report("Details")
            private_AddDetail details, diagnostic
            report("Details") = details
        End If
        Exit Function
    End If
    If success Then
        expected = private_NormalizeNewlines(VBA.CStr(testStep("ExpectedText")))
    Else
        expected = private_NormalizeNewlines(VBA.CStr(testStep("ExpectedError")))
        actual = private_NormalizeNewlines(diagnostic)
    End If
    If VBA.StrComp(expected, actual, VBA.vbBinaryCompare) = 0 Then
        private_ValidateWord = True
        Exit Function
    End If
    position = 1
    Do While position <= VBA.Len(expected) And position <= VBA.Len(actual)
        If VBA.Mid$(expected, position, 1) <> VBA.Mid$(actual, position, 1) Then Exit Do
        position = position + 1
    Loop
    line = 1
    column = 1
    For index = 1 To position - 1
        If VBA.Mid$(expected, index, 1) = VBA.vbLf Then
            line = line + 1
            column = 1
        Else
            column = column + 1
        End If
    Next index
    report("Details") = private_Mismatch(testStep("Id"), "WORD", "Text " & VBA.CStr(line) & ":" & VBA.CStr(column), _
        VBA.Mid$(expected, position, 80), VBA.Mid$(actual, position, 80))
End Function

Private Function private_RowIdentity(ByVal row As Object) As String
    Dim header As Variant
    Dim text As String

    For Each header In VBA.Split(ex_PEB_Export.fn_Config("test.expected.rowKeys"), "|")
        text = private_ValueText(row(header))
        private_RowIdentity = private_RowIdentity & VBA.CStr(VBA.Len(text)) & ":" & text & "|"
    Next header
End Function

Private Function private_SameValue( _
    ByVal expected As Variant, _
    ByVal actual As Variant _
) As Boolean
    If VBA.IsError(expected) Or VBA.IsError(actual) Then Exit Function
    If VBA.Len(private_ValueText(expected)) = 0 And VBA.Len(private_ValueText(actual)) = 0 Then
        private_SameValue = True
    ElseIf VBA.VarType(expected) = VBA.vbString Or VBA.VarType(actual) = VBA.vbString Then
        If VBA.VarType(expected) <> VBA.VarType(actual) Then Exit Function
        private_SameValue = (VBA.StrComp(VBA.CStr(expected), VBA.CStr(actual), VBA.vbBinaryCompare) = 0)
    ElseIf VBA.VarType(expected) = VBA.vbBoolean Or VBA.VarType(actual) = VBA.vbBoolean Then
        If VBA.VarType(expected) <> VBA.VarType(actual) Then Exit Function
        private_SameValue = (expected = actual)
    ElseIf VBA.IsNumeric(expected) And VBA.IsNumeric(actual) Then
        private_SameValue = (VBA.CDbl(expected) = VBA.CDbl(actual))
    End If
End Function

Private Function private_ValueText(ByVal value As Variant) As String
    If VBA.IsEmpty(value) Or VBA.IsNull(value) Then Exit Function
    private_ValueText = VBA.CStr(value)
End Function

Private Function private_NormalizeNewlines(ByVal text As String) As String
    private_NormalizeNewlines = VBA.Replace$(VBA.Replace$(text, VBA.vbCrLf, VBA.vbLf), VBA.vbCr, VBA.vbLf)
End Function

Private Function private_Mismatch( _
    ByVal stepId As String, _
    ByVal recordId As String, _
    ByVal field As String, _
    ByVal expected As String, _
    ByVal actual As String _
) As String
    Dim text As String

    text = ex_PEB_Export.fn_Config("text.TestsFieldMismatch")
    text = VBA.Replace$(text, "{step}", stepId)
    text = VBA.Replace$(text, "{record}", recordId)
    text = VBA.Replace$(text, "{field}", field)
    text = VBA.Replace$(text, "{expected}", expected)
    private_Mismatch = VBA.Replace$(text, "{actual}", actual)
End Function

Private Sub private_AddDetail( _
    ByRef details As String, _
    ByVal text As String _
)
    If VBA.Len(details) > 0 Then details = details & VBA.vbLf
    details = VBA.Left$(details & text, 30000)
End Sub

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
    prefix = "export.Movement.profiles." & values("MovementPlanId")
    operation = ex_PEB_Export.fn_Config(prefix & ".operation")
    If operation = "transition" Then
        private_ValidateMutation = private_ValidateTransition(before, after, values, prefix, diagnostic)
        Exit Function
    End If
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

Private Function private_ValidateTransition( _
    ByVal before As Object, _
    ByVal after As Object, _
    ByVal values As Object, _
    ByVal prefix As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim oldRows As Collection
    Dim newRows As Collection
    Dim oldRow As Object
    Dim newRow As Object
    Dim changed As Long
    Dim index As Long
    Dim mappingIndex As Long
    Dim mappingPrefix As String
    Dim field As String
    Dim header As String
    Dim key As Variant
    Dim allowed As Object

    Set oldRows = before("Rows")
    Set newRows = after("Rows")
    If newRows.Count <> oldRows.Count + 1 Or before("OpenCount") <> after("OpenCount") Then
        GoTo Mismatch
    End If
    mappingPrefix = "export.Movement.profiles." & ex_PEB_Export.fn_Config(prefix & ".closeProfile") & ".fields"
    Set allowed = VBA.CreateObject("Scripting.Dictionary")
    For mappingIndex = 1 To ex_PEB_Export.fn_Count(mappingPrefix & ".count")
        header = ex_PEB_Export.fn_Config(mappingPrefix & "." & VBA.CStr(mappingIndex) & ".header")
        allowed(header) = True
    Next mappingIndex
    For index = 1 To oldRows.Count
        Set oldRow = oldRows(index)
        Set newRow = newRows(index)
        If oldRow("_signature") <> newRow("_signature") Then
            changed = changed + 1
            For mappingIndex = 1 To ex_PEB_Export.fn_Count(prefix & ".join.count")
                header = ex_PEB_Export.fn_Config(prefix & ".join." & VBA.CStr(mappingIndex) & ".header")
                field = ex_PEB_Export.fn_Config(prefix & ".join." & VBA.CStr(mappingIndex) & ".field")
                If oldRow(header) <> ex_PEB_Export.fn_Text(values(field)) Then
                    GoTo Mismatch
                End If
            Next mappingIndex
            For Each key In oldRow.Keys
                If key <> "_signature" And key <> "_cells" And Not allowed.Exists(key) Then
                    If oldRow(key) <> newRow(key) Then
                        GoTo Mismatch
                    End If
                End If
            Next key
            For mappingIndex = 1 To ex_PEB_Export.fn_Count(mappingPrefix & ".count")
                header = ex_PEB_Export.fn_Config(mappingPrefix & "." & VBA.CStr(mappingIndex) & ".header")
                field = ex_PEB_Export.fn_Config(mappingPrefix & "." & VBA.CStr(mappingIndex) & ".field")
                If Not private_MappedValueMatches(values(field), newRow("_cells")(header)) Then
                    GoTo Mismatch
                End If
            Next mappingIndex
        End If
    Next index
    If changed <> 1 Then
        GoTo Mismatch
    End If
    Set newRow = newRows(newRows.Count)
    For mappingIndex = 1 To ex_PEB_Export.fn_Count(prefix & ".fields.count")
        header = ex_PEB_Export.fn_Config(prefix & ".fields." & VBA.CStr(mappingIndex) & ".header")
        field = ex_PEB_Export.fn_Config(prefix & ".fields." & VBA.CStr(mappingIndex) & ".field")
        If Not private_MappedValueMatches(values(field), newRow("_cells")(header)) Then
            GoTo Mismatch
        End If
    Next mappingIndex
    private_ValidateTransition = True
    Exit Function
Mismatch:
    diagnostic = ex_PEB_Export.fn_Config("text.TestsMutationMismatch")
End Function

Private Function private_MappedValueMatches( _
    ByVal expected As Variant, _
    ByVal actual As Variant _
) As Boolean
    If VBA.VarType(expected) = VBA.vbDate Then
        expected = VBA.CDbl(expected)
    End If
    private_MappedValueMatches = private_SameValue(expected, actual)
End Function