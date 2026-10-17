VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PEB_PgTestsController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_pageBase As obj_PageBase
Private m_sheet As Worksheet
Private m_catalog As Collection
Private m_catalogRecords As Collection
Private m_catalogTable As obj_UiRawTable
Private m_selected As Object
Private m_previews As Object
Private m_runSelection As Object
Private m_runner As obj_PEB_TestRunner
Private m_reports As Collection
Private m_allSelected As Boolean
Private m_allPreviews As Boolean
Private m_running As Boolean
Private m_stop As Boolean
Private m_completed As Long
Private m_phasePercent As Double
Private m_total As Long
Private m_pass As Long
Private m_fail As Long
Private m_skip As Long
Private m_error As Long
Private m_previousStatus As Variant
Private m_status As String
Private m_writtenStatus As String
Private m_ownsStatus As Boolean
Private m_environmentSwitchPending As Boolean

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
Public Property Get IsRunning() As Boolean
    IsRunning = m_running
End Property

Public Property Get StopRequested() As Boolean
    StopRequested = m_stop
End Property

' //
' // API
' //
Public Function Initialize(ByVal pageBase As obj_PageBase) As Boolean
    If m_isInitialized Or m_isDisposed Or pageBase Is Nothing Then
        Exit Function
    End If
    Set m_pageBase = pageBase
    Set m_sheet = ThisWorkbook.Worksheets(ex_PEB_Export.fn_Config("test.ui.mainSheet"))
    Set m_selected = VBA.CreateObject("Scripting.Dictionary")
    Set m_previews = VBA.CreateObject("Scripting.Dictionary")
    Set m_reports = New Collection
    Set m_catalogRecords = New Collection
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    m_stop = True
    m_running = False
    ex_PEB_Testing.fn_Detach Me
    If Not m_runner Is Nothing Then
        m_runner.Dispose
    End If
    Set m_runner = Nothing
    private_RestoreStatus
    Set m_pageBase = Nothing
    Set m_sheet = Nothing
    Set m_catalog = Nothing
    Set m_catalogRecords = Nothing
    Set m_catalogTable = Nothing
    Set m_selected = Nothing
    Set m_previews = Nothing
    Set m_runSelection = Nothing
    Set m_reports = Nothing
End Sub

Public Function TryLoadCatalog(ByRef diagnostic As String) As Boolean
    Dim scenario As Object
    Dim record As Object
    Dim retained As Object

    On Error GoTo EH
    private_Guard
    If m_running Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.TestsBusy")
    End If
    If Not ex_PEB_Testing.fn_ReadCatalog(m_catalog, diagnostic) Then
        Exit Function
    End If
    If m_catalog.Count > ex_PEB_Export.fn_Count("test.ui.maxScenarios") Then
        Err.Raise 5, , "Scenario catalog exceeds the UI limit."
    End If
    Set retained = VBA.CreateObject("Scripting.Dictionary")
    Set m_catalogRecords = New Collection
    For Each scenario In m_catalog
        If m_selected.Exists(scenario("Id")) Then
            retained.Add scenario("Id"), True
        End If
        If Not m_previews.Exists(scenario("Id")) Then
            m_previews.Add scenario("Id"), True
        End If
        Set record = VBA.CreateObject("Scripting.Dictionary")
        record("ScenarioId") = scenario("Id")
        record("Operation") = scenario("Operation")
        record("EventLabel") = scenario("EventLabel")
        record("StepCount") = scenario("Steps").Count
        record("RuleStatus") = scenario("RuleStatus")
        m_catalogRecords.Add record
    Next scenario
    Set m_selected = retained
    private_RefreshSelection True
    private_DrawResults
    TryLoadCatalog = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
End Function

Public Function SelectionChangedHandler() As Boolean
    Dim record As Object
    Dim id As String
    Dim scenario As Object

    private_Guard
    If m_environmentSwitchPending Then
        Exit Function
    End If
    If Not m_running Then
        If private_CheckboxValue("AllC") <> m_allSelected Then
            private_SelectAll private_CheckboxValue("AllC")
        ElseIf private_CheckboxValue("AllP") <> m_allPreviews Then
            For Each scenario In m_catalog
                m_previews(scenario("Id")) = private_CheckboxValue("AllP")
            Next scenario
        Else
            For Each record In m_catalogRecords
                id = record("ScenarioId")
                If VBA.VarType(record("Selected")) <> VBA.vbBoolean Or VBA.VarType(record("PreviewEnabled")) <> VBA.vbBoolean Then
                    Err.Raise 5, , "Invalid scenario selection."
                End If
                If m_selected.Exists(id) Then
                    m_selected.Remove id
                End If
                If record("Selected") Then
                    m_selected.Add id, True
                End If
                m_previews(id) = record("PreviewEnabled")
            Next record
        End If
    End If
    private_RefreshSelection
    SelectionChangedHandler = True
End Function

Public Sub RunSelected()
    Dim key As Variant
    Dim scenario As Object
    Dim diagnostic As String

    On Error GoTo EH
    private_Guard
    If m_running Then
        Exit Sub
    End If
    If m_environmentSwitchPending Then
        Exit Sub
    End If
    If m_selected.Count = 0 Then
        ex_WindowsUi.fn_ShowMessage ex_PEB_Export.fn_Config("text.TestsSelectRequired"), _
            VBA.vbExclamation, ex_PEB_Export.fn_Config("text.TestTitle")
        Exit Sub
    End If
    Set m_runSelection = VBA.CreateObject("Scripting.Dictionary")
    For Each key In m_selected.Keys
        m_runSelection.Add key, VBA.CBool(m_previews(key))
    Next key
    Set m_reports = New Collection
    m_running = True
    m_stop = False
    m_completed = 0
    m_phasePercent = 0
    m_total = 0
    For Each scenario In m_catalog
        If m_runSelection.Exists(scenario("Id")) Then
            m_total = m_total + scenario("Steps").Count
        End If
    Next scenario
    m_pass = 0
    m_fail = 0
    m_skip = 0
    m_error = 0
    If Not m_ownsStatus Then
        m_previousStatus = Application.StatusBar
    ElseIf VBA.CStr(Application.StatusBar) <> m_writtenStatus Then
        m_previousStatus = Application.StatusBar
    End If
    m_writtenStatus = VBA.vbNullString
    m_ownsStatus = True
    private_SetRunning True
    private_DrawResults
    private_Progress ex_PEB_Export.fn_Config("text.TestsRunning")
    ex_PEB_Testing.fn_ScheduleTick
    Exit Sub
EH:
    diagnostic = VBA.Err.Description
    private_Finish diagnostic
End Sub

Public Sub RequestStop()
    If Not m_running Then
        Exit Sub
    End If
    m_stop = True
    If Not m_runner Is Nothing Then
        m_runner.RequestStop
    End If
    private_Progress ex_PEB_Export.fn_Config("text.TestsStopping")
End Sub

Public Sub FailRun(ByVal diagnostic As String)
    If Not m_isInitialized Or m_isDisposed Then
        Exit Sub
    End If
    private_Finish diagnostic
End Sub

Public Sub ReportPhase(ByVal key As String)
    Dim percent As Double

    If Not m_running Then
        Exit Sub
    End If
    percent = 100# * ex_PEB_Export.fn_Count("test.progress." & key, 8) / ex_PEB_Export.fn_Count("test.progress.stageCount", 8)
    If percent > m_phasePercent Then
        m_phasePercent = percent
    End If
    If m_stop Then
        private_Progress ex_PEB_Export.fn_Config("text.TestsStopping")
    Else
        private_Progress ex_PEB_Export.fn_Config("text.TestPhase." & key)
    End If
End Sub

Public Function Advance() As Boolean
    Dim report As Object
    Dim diagnostic As String
    Dim success As Boolean

    On Error GoTo EH
    If Not m_running Or m_isDisposed Then
        Exit Function
    End If
    private_Guard
    If m_runner Is Nothing Then
        Set m_runner = New obj_PEB_TestRunner
        If Not m_runner.Initialize(diagnostic) Then
            GoTo Failed
        End If
        If Not m_runner.TryBegin(m_runSelection, diagnostic) Then
            GoTo Failed
        End If
        m_total = m_runner.Total
        If m_stop Then
            m_runner.RequestStop
        End If
        private_Progress ex_PEB_Export.fn_Config("text.TestsRunning")
        Advance = True
        Exit Function
    End If
    If m_stop Then
        m_runner.RequestStop
    End If
    success = m_runner.TryAdvance(report, diagnostic)
    If Not report Is Nothing Then
        Me.PublishResult report
    End If
    If m_runner.Completed > m_completed Then
        m_phasePercent = 0
    End If
    m_completed = m_runner.Completed
    private_Progress ex_PEB_Export.fn_Config("text.TestsRunning")
    If Not success Then
        GoTo Failed
    End If
    If m_runner.Finished Then
        private_Finish VBA.vbNullString
    Else
        Advance = True
    End If
    Exit Function
EH:
    diagnostic = VBA.Err.Description
Failed:
    private_Finish diagnostic
End Function

Public Sub PublishResult(ByVal report As Object)
    m_reports.Add report
    Select Case report("Result")
        Case "PASS"
            m_pass = m_pass + 1
        Case "FAIL"
            m_fail = m_fail + 1
        Case "SKIP"
            m_skip = m_skip + 1
        Case "ERROR"
            m_error = m_error + 1
    End Select
    private_DrawResults
End Sub

Public Function NavigatePreview(ByVal item As Object) As Boolean
    Dim index As Long
    Dim previewIndex As Long
    Dim context As obj_UiRenderContext
    Dim area As Range
    Dim report As Object

    private_Guard
    If item Is Nothing Then
        Exit Function
    End If
    For index = 1 To m_reports.Count
        Set report = m_reports(index)
        If report("PreviewEnabled") Then
            previewIndex = previewIndex + 1
        End If
        If VBA.CStr(report("ScenarioId")) = VBA.CStr(item("ScenarioId")) And _
                VBA.CStr(report("StepId")) = VBA.CStr(item("StepId")) Then Exit For
    Next index
    If index > m_reports.Count Then
        Exit Function
    End If
    If Not report("PreviewEnabled") Then
        Exit Function
    End If
    If Not ex_UiRuntime.fn_TryGetContext(m_sheet, context) Then
        Exit Function
    End If
    If Not context.TryGetItemArea("TestDocuments", previewIndex, area) Then
        Exit Function
    End If
    Application.Goto area.Cells(1, 1), True
    NavigatePreview = True
End Function

Public Function TestRunHandler() As Boolean
    Me.RunSelected
    TestRunHandler = True
End Function

Public Function CancelTestsHandler() As Boolean
    ex_PEB_Testing.fn_RequestStop
    CancelTestsHandler = True
End Function

Public Function TestPreviewHandler(ByVal item As Object) As Boolean
    TestPreviewHandler = Me.NavigatePreview(item)
End Function

Public Function EnvironmentSwitchHandler() As Boolean
    Dim environment As String
    Dim target As String
    Dim configPath As String
    Dim diagnostic As String
    Dim status As String

    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If ex_PEB_Testing.fn_IsRunning() Then
        EnvironmentSwitchHandler = private_SetActionStatus(ex_PEB_Export.fn_Config("text.TestsBusy"), True)
        Exit Function
    End If
    If m_environmentSwitchPending Then
        EnvironmentSwitchHandler = True
        Exit Function
    End If
    On Error GoTo EH_ENVIRONMENT
    environment = ex_PEB_Export.fn_Config("environment.name")
    Select Case environment
        Case "PROD"
            target = "TEST"
        Case "TEST"
            target = "PROD"
        Case Else
            Err.Raise 5, , "Unknown PEB environment: " & environment
    End Select
    configPath = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("environment.config." & target))
    status = VBA.Replace$(ex_PEB_Export.fn_Config("text.EnvironmentSwitchPending"), "{environment}", target)
    If Not ex_ConfigRefresh.fn_TryUpdate(diagnostic, configPath) Then
        GoTo FailedEnvironment
    End If
    m_environmentSwitchPending = True
    ' Reinitialize after this handler returns; never dispose an executing UI control.
    If Not ex_AppCommands.fn_RequestReinitialize() Then
        Err.Raise 5, , "Cannot schedule PEB environment initialization."
    End If
    EnvironmentSwitchHandler = private_SetActionStatus(status, False)
    Exit Function
EH_ENVIRONMENT:
    diagnostic = VBA.Err.Description
    m_environmentSwitchPending = False
FailedEnvironment:
    private_SetActionStatus diagnostic, True
    ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation, "PersonalEventBuilder"
End Function

' //
' // Private
' //

Private Sub private_SelectAll(ByVal selected As Boolean)
    Dim scenario As Object

    Set m_selected = VBA.CreateObject("Scripting.Dictionary")
    For Each scenario In m_catalog
        If selected Then
            m_selected.Add scenario("Id"), True
        End If
    Next scenario
End Sub

Private Function private_CheckboxValue(ByVal key As String) As Boolean
    Dim value As Variant
    Dim resolved As Object
    Dim isObject As Boolean

    If Not m_pageBase.BindingContext.TryGetValue("TestSelection", key, value, resolved, isObject) Then
        Err.Raise 5, , "Missing test selection binding: " & key
    End If
    If isObject Or VBA.VarType(value) <> VBA.vbBoolean Then
        Err.Raise 5, , "Invalid test selection binding: " & key
    End If
    private_CheckboxValue = VBA.CBool(value)
End Function

Private Sub private_Guard()
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Test UI is not initialized."
    End If
    If ex_PEB_Export.fn_Config("environment.name") <> "TEST" Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.TestOnly")
    End If
End Sub

Private Sub private_Finish(ByVal diagnostic As String)
    Dim text As String
    Dim checkpointDiagnostic As String

    ex_PEB_Testing.fn_CancelTick
    If Not m_runner Is Nothing Then
        If Not m_runner.TryCheckpoint(checkpointDiagnostic) Then
            diagnostic = diagnostic & " " & checkpointDiagnostic
        End If
    End If
    If Not m_runner Is Nothing Then
        m_runner.Dispose
    End If
    Set m_runner = Nothing
    m_running = False
    private_SetRunning False
    text = ex_PEB_Export.fn_Config("text.TestsFinished")
    If m_stop Then
        text = ex_PEB_Export.fn_Config("text.TestsStopped")
    End If
    If VBA.Len(diagnostic) > 0 Then
        text = ex_PEB_Export.fn_Config("text.TestsRunError")
        If m_error = 0 Then
            m_error = 1
        End If
    End If
    text = VBA.Replace$(text, "{summary}", private_Summary())
    If VBA.Len(diagnostic) > 0 Then
        text = text & " " & diagnostic
    End If
    private_Progress text
    If Not Application.ActiveWorkbook Is ThisWorkbook Then
        private_RestoreStatus
    End If
    If VBA.Len(diagnostic) > 0 Then
        ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation, ex_PEB_Export.fn_Config("text.TestTitle")
    End If
End Sub

Private Sub private_RestoreStatus()
    If Not m_ownsStatus Then
        Exit Sub
    End If
    If VBA.Len(m_writtenStatus) > 0 Then
        If VBA.CStr(Application.StatusBar) = m_writtenStatus Then
            Application.StatusBar = m_previousStatus
        End If
    End If
    m_ownsStatus = False
End Sub

Private Sub private_SetRunning(ByVal running As Boolean)
    If Not m_pageBase.BindingContext.SetValue("Settings", "TestsRunning", running) Then
        Err.Raise 5, , "Cannot update test runner state."
    End If
End Sub

Private Sub private_Progress(ByVal state As String)
    Dim percent As Long
    Dim style As String

    If m_total > 0 Then
        percent = VBA.CLng(VBA.Fix((100# * m_completed + m_phasePercent) / m_total))
    End If
    m_status = ex_PEB_Export.fn_Config("text.TestsProgress")
    m_status = VBA.Replace$(m_status, "{percent}", VBA.CStr(percent))
    m_status = VBA.Replace$(m_status, "{completed}", VBA.CStr(m_completed))
    m_status = VBA.Replace$(m_status, "{total}", VBA.CStr(m_total))
    m_status = VBA.Replace$(m_status, "{state}", state)
    If Application.ActiveWorkbook Is ThisWorkbook Then
        Application.StatusBar = m_status
        m_writtenStatus = m_status
    End If
    If Not m_pageBase.BindingContext.SetValue("Form", "ActionStatus", m_status) Then
        Err.Raise 5, , "Cannot update test progress."
    End If
    style = "actionSuccess"
    If m_fail > 0 Or m_error > 0 Then
        style = "actionError"
    End If
    If Not m_pageBase.BindingContext.SetValue("Form", "ActionStatusStyle", style) Then
        Err.Raise 5, , "Cannot update test progress style."
    End If
End Sub

Private Function private_Summary() As String
    Dim text As String

    text = ex_PEB_Export.fn_Config("text.TestsSummary")
    text = VBA.Replace$(text, "{pass}", VBA.CStr(m_pass))
    text = VBA.Replace$(text, "{fail}", VBA.CStr(m_fail))
    text = VBA.Replace$(text, "{skip}", VBA.CStr(m_skip))
    private_Summary = VBA.Replace$(text, "{error}", VBA.CStr(m_error))
End Function

Private Sub private_RefreshSelection(Optional ByVal bindCatalog As Boolean = False)
    Dim record As Object
    Dim previewCount As Long
    Dim id As String
    Dim fields As Variant
    Dim index As Long
    Dim fieldIndex As Long

    For Each record In m_catalogRecords
        id = record("ScenarioId")
        record("Selected") = m_selected.Exists(id)
        record("PreviewEnabled") = VBA.CBool(m_previews(id))
        If record("PreviewEnabled") Then
            previewCount = previewCount + 1
        End If
    Next record
    m_allSelected = (m_catalog.Count > 0 And m_selected.Count = m_catalog.Count)
    m_allPreviews = (m_catalog.Count > 0 And previewCount = m_catalog.Count)
    If Not m_pageBase.BindingContext.SetValue("TestSelection", "AllC", m_allSelected) Then
        Err.Raise 5, , "Cannot refresh master selection."
    End If
    If Not m_pageBase.BindingContext.SetValue("TestSelection", "AllP", m_allPreviews) Then
        Err.Raise 5, , "Cannot refresh master preview selection."
    End If
    If bindCatalog Then
        private_BindCatalog
    Else
        If m_catalogTable Is Nothing Then
            Err.Raise 5, , "Test catalog table is missing."
        End If
        fields = VBA.Split(ex_PEB_Export.fn_Config("test.ui.catalogFields"), "|")
        For index = 1 To m_catalogRecords.Count
            For fieldIndex = 0 To UBound(fields)
                If Not m_catalogTable.SetValueAt(index, fieldIndex + 1, m_catalogRecords(index)(fields(fieldIndex))) Then
                    Err.Raise 5, , "Cannot update scenario table values."
                End If
            Next fieldIndex
        Next index
        private_MarkDirty "TestCatalog", False
    End If
End Sub

Private Sub private_BindCatalog()
    Dim rawTable As obj_UiRawTable
    Dim fields As Variant
    Dim captions As Variant
    Dim values As Variant
    Dim index As Long
    Dim fieldIndex As Long

    fields = VBA.Split(ex_PEB_Export.fn_Config("test.ui.catalogFields"), "|")
    captions = VBA.Split(ex_PEB_Export.fn_Config("test.ui.catalogCaptions"), "|")
    If UBound(fields) <> UBound(captions) Then
        Err.Raise 5, , "Test catalog fields differ."
    End If
    Set rawTable = New obj_UiRawTable
    If m_catalogRecords.Count > 0 Then
        ReDim values(1 To m_catalogRecords.Count, 1 To UBound(fields) + 1)
        For index = 1 To m_catalogRecords.Count
            For fieldIndex = 0 To UBound(fields)
                values(index, fieldIndex + 1) = m_catalogRecords(index)(fields(fieldIndex))
            Next fieldIndex
        Next index
        If Not rawTable.Initialize(values, captions, VBA.vbNullString, fields) Then
            Err.Raise 5, , "Cannot initialize scenario table data."
        End If
        If Not rawTable.SetRecords(m_catalogRecords) Then
            Err.Raise 5, , "Cannot attach scenario records."
        End If
    Else
        If Not rawTable.InitializeEmpty(captions, VBA.vbNullString, fields) Then
            Err.Raise 5, , "Cannot initialize empty scenario table."
        End If
    End If
    If Not m_pageBase.BindingContext.SetObject("Data", "TestCatalog", rawTable) Then
        Err.Raise 5, , "Cannot bind scenario table."
    End If
    Set m_catalogTable = rawTable
    private_MarkDirty "TestCatalog"
End Sub

Private Sub private_MarkDirty( _
    ByVal name As String, _
    Optional ByVal layoutChanged As Boolean = True _
)
    Dim context As obj_UiRenderContext

    If Not ex_UiRuntime.fn_TryGetContext(m_sheet, context) Then
        Exit Sub
    End If
    If Not context.BindingContext Is m_pageBase.BindingContext Then
        Exit Sub
    End If
    context.MarkDirty name, layoutChanged
End Sub

Private Sub private_DrawResults()
    Dim rawTable As obj_UiRawTable
    Dim fields As Variant
    Dim captions As Variant
    Dim values As Variant
    Dim report As Object
    Dim index As Long
    Dim fieldIndex As Long
    Dim rowCount As Long

    If m_sheet Is Nothing Then
        Exit Sub
    End If
    fields = VBA.Split(ex_PEB_Export.fn_Config("test.panel.fields"), "|")
    captions = VBA.Split(ex_PEB_Export.fn_Config("test.panel.captions"), "|")
    If UBound(fields) <> UBound(captions) Then
        Err.Raise 5, , "Test panel fields differ."
    End If
    rowCount = m_reports.Count
    Set rawTable = New obj_UiRawTable
    If rowCount > 0 Then
        ReDim values(1 To rowCount, 1 To UBound(fields) + 1)
        For index = 1 To m_reports.Count
            Set report = m_reports(index)
            For fieldIndex = 0 To UBound(fields)
                values(index, fieldIndex + 1) = ex_PEB_Export.fn_Text(report(fields(fieldIndex)))
            Next fieldIndex
        Next index
        If Not rawTable.Initialize(values, captions, VBA.vbNullString, fields) Then
            Err.Raise 5, , "Cannot initialize test result data."
        End If
        If Not rawTable.SetRecords(m_reports) Then
            Err.Raise 5, , "Cannot attach result row records."
        End If
    Else
        If Not rawTable.InitializeEmpty(captions, VBA.vbNullString, fields) Then
            Err.Raise 5, , "Cannot initialize empty test results."
        End If
    End If
    private_BindPreviews
    If Not m_pageBase.BindingContext.SetObject("Data", "TestResults", rawTable) Then
        Err.Raise 5, , "Cannot bind test results."
    End If
    private_MarkDirty "TestResults"
End Sub

Private Sub private_BindPreviews()
    Dim previews As Collection
    Dim report As Object
    Dim item As Object
    Dim text As String

    Set previews = New Collection
    For Each report In m_reports
        If report("PreviewEnabled") Then
            Set item = VBA.CreateObject("Scripting.Dictionary")
            text = ex_PEB_Export.fn_Config("text.TestPreviewHeading")
            text = VBA.Replace$(text, "{scenario}", VBA.CStr(report("ScenarioId")))
            text = VBA.Replace$(text, "{step}", VBA.CStr(report("StepId")))
            text = VBA.Replace$(text, "{event}", VBA.CStr(report("EventName")))
            text = text & VBA.vbLf & VBA.vbLf & VBA.CStr(report("PreviewText"))
            If report.Exists("PreviewDiagnostic") Then
                If VBA.Len(report("PreviewDiagnostic")) > 0 Then
                    text = text & VBA.vbLf & ex_PEB_Export.fn_Config("text.TestPreviewError") & VBA.vbLf & report("PreviewDiagnostic")
                End If
            End If
            item.Add "Text", text
            previews.Add item
        End If
    Next report
    If Not m_pageBase.BindingContext.SetObject("Data", "TestPreviews", previews) Then
        Err.Raise 5, , "Cannot bind test previews."
    End If
    private_MarkDirty "TestDocuments"
End Sub

Private Function private_SetActionStatus( _
    ByVal text As String, _
    ByVal isError As Boolean _
) As Boolean
    Dim updates As Object
    Dim diagnostic As String

    Set updates = VBA.CreateObject("Scripting.Dictionary")
    updates("Form.ActionStatus") = text
    updates("Form.ActionStatusStyle") = "actionSuccess"
    If isError Then
        updates("Form.ActionStatusStyle") = "actionError"
    End If
    private_SetActionStatus = m_pageBase.BindingContext.TryApplyValues(updates, diagnostic)
End Function