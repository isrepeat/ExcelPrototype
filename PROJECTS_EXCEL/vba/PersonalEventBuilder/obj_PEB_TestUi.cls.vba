VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PEB_TestUi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_sheet As Worksheet
Private m_checkboxes As Collection
Private m_selectionCommand As obj_UiCommand
Private m_catalog As Collection
Private m_selected As Object
Private m_runSelection As Object
Private m_runner As obj_PEB_TestRunner
Private m_reports As Collection
Private m_running As Boolean
Private m_stop As Boolean
Private m_baseRow As Long
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
Public Function Initialize() As Boolean
    If m_isInitialized Or m_isDisposed Then Exit Function
    Set m_selected = VBA.CreateObject("Scripting.Dictionary")
    Set m_reports = New Collection
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    Dim cleanupError As String
    Dim index As Long
    Dim previousScreen As Boolean

    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
    m_stop = True
    On Error Resume Next
    private_ClearCheckboxes
    If Not m_selectionCommand Is Nothing Then m_selectionCommand.Dispose
    Set m_selectionCommand = Nothing
    ex_PEB_Testing.fn_CancelTick
    If Not m_runner Is Nothing Then m_runner.Dispose
    Set m_runner = Nothing
    m_running = False
    private_SetRunning False
    private_RestoreStatus
    previousScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    If Not m_sheet Is Nothing Then
        For index = m_sheet.Shapes.Count To 1 Step -1
            If VBA.Left$(m_sheet.Shapes(index).Name, 5) = "pebt_" Then m_sheet.Shapes(index).Delete
        Next index
    End If
    Application.ScreenUpdating = previousScreen
    Set m_sheet = Nothing
    Set m_selected = Nothing
    Set m_runSelection = Nothing
    Set m_catalog = Nothing
    Set m_reports = Nothing
    cleanupError = VBA.Err.Description
    If VBA.Len(cleanupError) > 0 Then ex_Core.fn_Diagnostic_WriteLog "TEST_UI_DISPOSE_ERROR | " & cleanupError
    On Error GoTo 0
End Sub

Public Function OpenPage() As Boolean
    Dim context As obj_UiRenderContext
    Dim rawTable As obj_UiRawTable
    Dim values As Variant
    Dim retained As Object
    Dim scenario As Object
    Dim diagnostic As String
    Dim previousScreen As Boolean
    Dim previousEvents As Boolean
    Dim index As Long
    Dim rowIndex As Long
    Dim captions As Variant

    private_Guard
    If m_running Then Exit Function
    On Error GoTo EH_READ
    If Not ex_PEB_Testing.fn_ReadCatalog(m_catalog, diagnostic) Then Err.Raise 5, , diagnostic
    If m_catalog.Count > ex_PEB_Export.fn_Count("test.ui.maxScenarios") Then Err.Raise 5, , "Scenario catalog exceeds the UI limit."
    Set retained = VBA.CreateObject("Scripting.Dictionary")
    For Each scenario In m_catalog
        If m_selected.Exists(scenario("Id")) Then retained.Add scenario("Id"), True
    Next scenario
    Set m_selected = retained
    Set m_sheet = ThisWorkbook.Worksheets(ex_PEB_Export.fn_Config("test.ui.mainSheet"))
    m_baseRow = ex_PEB_Export.fn_Count("test.panel.row")
    previousScreen = Application.ScreenUpdating
    previousEvents = Application.EnableEvents
    On Error GoTo EH_RENDER
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    private_ClearCheckboxes
    Set m_checkboxes = New Collection
    For index = m_sheet.Shapes.Count To 1 Step -1
        If VBA.Left$(m_sheet.Shapes(index).Name, 5) = "pebt_" Then m_sheet.Shapes(index).Delete
    Next index
    captions = VBA.Split(ex_PEB_Export.fn_Config("test.ui.catalogCaptions"), "|")
    If UBound(captions) <> 5 Then Err.Raise 5, , "Test catalog captions must contain six columns."
    Set rawTable = New obj_UiRawTable
    If m_catalog.Count > 0 Then
        ReDim values(1 To m_catalog.Count, 1 To 6)
        For index = 1 To m_catalog.Count
            Set scenario = m_catalog(index)
            values(index, 1) = VBA.vbNullString
            values(index, 2) = scenario("Id")
            values(index, 3) = scenario("Operation")
            values(index, 4) = scenario("EventLabel")
            values(index, 5) = scenario("Steps").Count
            values(index, 6) = scenario("RuleStatus")
        Next index
        If Not rawTable.Initialize(values, captions) Then Err.Raise 5, , "Cannot initialize scenario table data."
    Else
        If Not rawTable.InitializeEmpty(captions) Then Err.Raise 5, , "Cannot initialize empty scenario table."
    End If
    If Not ex_UiRuntime.fn_TryGetContext(m_sheet, context) Then Err.Raise 5, , "Test page render context is missing."
    If Not context.BindingContext.SetObject("Data", "TestCatalog", rawTable) Then Err.Raise 5, , "Cannot bind scenario table data."
    context.MarkDirty "TestCatalog", True
    If Not context.FlushLayout(diagnostic) Then Err.Raise 5, , diagnostic
    rowIndex = m_baseRow + 5
    For index = 1 To m_catalog.Count
        Set scenario = m_catalog(index)
        private_CreateCheckbox index, rowIndex, m_selected.Exists(scenario("Id"))
        rowIndex = rowIndex + 1
    Next index
    ex_Core.fn_Diagnostic_WriteLog "TEST_UI_STAGE | Name=DrawResults"
    ex_Core.fn_Diagnostic_Flush
    private_DrawResults
    m_sheet.Activate
    Application.Goto m_sheet.Cells(1, 1), True
    OpenPage = True
    ex_Core.fn_Diagnostic_WriteLog "TEST_UI_OPEN_COMPLETED"
    ex_Core.fn_Diagnostic_Flush
CleanRender:
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreen
    On Error GoTo 0
    If VBA.Len(diagnostic) > 0 Then Err.Raise 5, , diagnostic
    Exit Function
EH_RENDER:
    diagnostic = VBA.Err.Description
    ex_Core.fn_Diagnostic_WriteLog "TEST_UI_RENDER_ERROR | " & diagnostic
    ex_Core.fn_Diagnostic_Flush
    Resume CleanRender
EH_READ:
    diagnostic = VBA.Err.Description
    Err.Raise 5, , diagnostic
End Function

Public Function SelectionChanged() As Boolean
    Dim index As Long
    Dim context As obj_UiRenderContext
    Dim value As Variant
    Dim resolved As Object
    Dim isObject As Boolean
    Dim id As String

    private_Guard
    If Not ex_UiRuntime.fn_TryGetContext(m_sheet, context) Then Exit Function
    If Not m_running Then
        For index = 1 To m_catalog.Count
            If Not context.BindingContext.TryGetValue("TestSelection", "C" & VBA.CStr(index), value, resolved, isObject) Then Exit Function
            If isObject Or VBA.VarType(value) <> VBA.vbBoolean Then Exit Function
            id = m_catalog(index)("Id")
            If m_selected.Exists(id) Then m_selected.Remove id
            If VBA.CBool(value) Then m_selected.Add id, True
        Next index
    End If
    private_RefreshCheckboxes
    SelectionChanged = True
End Function

Public Sub SelectAll(ByVal selected As Boolean)
    Dim scenario As Object
    Dim index As Long

    private_Guard
    If m_running Or m_catalog Is Nothing Then Exit Sub
    Set m_selected = VBA.CreateObject("Scripting.Dictionary")
    For index = 1 To m_catalog.Count
        Set scenario = m_catalog(index)
        If selected Then
            m_selected.Add scenario("Id"), True
        End If
    Next index
    private_RefreshCheckboxes
End Sub

Public Sub RunSelected()
    Dim key As Variant
    Dim scenario As Object
    Dim diagnostic As String

    private_Guard
    If m_running Then Exit Sub
    If m_selected.Count = 0 Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.TestsSelectRequired")
    On Error GoTo EH
    Set m_runSelection = VBA.CreateObject("Scripting.Dictionary")
    For Each key In m_selected.Keys
        m_runSelection.Add key, True
    Next key
    Set m_reports = New Collection
    m_running = True
    m_stop = False
    m_completed = 0
    m_phasePercent = 0
    m_total = 0
    For Each scenario In m_catalog
        If m_runSelection.Exists(scenario("Id")) Then m_total = m_total + scenario("Steps").Count
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
    If Not m_running Then Exit Sub
    m_stop = True
    If Not m_runner Is Nothing Then m_runner.RequestStop
    private_Progress ex_PEB_Export.fn_Config("text.TestsStopping")
End Sub

Public Sub ReportPhase(ByVal key As String)
    Dim percent As Double

    If Not m_running Then Exit Sub
    percent = 100# * ex_PEB_Export.fn_Count("test.progress." & key, 8) / ex_PEB_Export.fn_Count("test.progress.stageCount", 8)
    If percent > m_phasePercent Then m_phasePercent = percent
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
    If Not m_running Or m_isDisposed Then Exit Function
    private_Guard
    If m_runner Is Nothing Then
        Set m_runner = New obj_PEB_TestRunner
        If Not m_runner.Initialize(diagnostic) Then GoTo Failed
        If Not m_runner.TryBegin(m_runSelection, diagnostic) Then GoTo Failed
        m_total = m_runner.Total
        If m_stop Then m_runner.RequestStop
        private_Progress ex_PEB_Export.fn_Config("text.TestsRunning")
        Advance = True
        Exit Function
    End If
    If m_stop Then m_runner.RequestStop
    success = m_runner.TryAdvance(report, diagnostic)
    If Not report Is Nothing Then Me.PublishResult report
    If m_runner.Completed > m_completed Then m_phasePercent = 0
    m_completed = m_runner.Completed
    private_Progress ex_PEB_Export.fn_Config("text.TestsRunning")
    If Not success Then GoTo Failed
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
    Dim context As obj_UiRenderContext
    Dim area As Range
    Dim report As Object

    private_Guard
    If item Is Nothing Then Exit Function
    For index = 1 To m_reports.Count
        Set report = m_reports(index)
        If VBA.CStr(report("ScenarioId")) = VBA.CStr(item("ScenarioId")) And _
                VBA.CStr(report("StepId")) = VBA.CStr(item("StepId")) Then Exit For
    Next index
    If index > m_reports.Count Then Exit Function
    If Not ex_UiRuntime.fn_TryGetContext(m_sheet, context) Then Exit Function
    If Not context.TryGetItemArea("TestDocuments", index, area) Then Exit Function
    Application.Goto area.Cells(1, 1), True
    NavigatePreview = True
End Function

' //
' // Private
' //
Private Sub private_Guard()
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Test UI is not initialized."
    If ex_PEB_Export.fn_Config("environment.name") <> "TEST" Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.TestOnly")
End Sub

Private Sub private_Finish(ByVal diagnostic As String)
    Dim text As String

    ex_PEB_Testing.fn_CancelTick
    If Not m_runner Is Nothing Then m_runner.Dispose
    Set m_runner = Nothing
    m_running = False
    private_SetRunning False
    text = ex_PEB_Export.fn_Config("text.TestsFinished")
    If m_stop Then text = ex_PEB_Export.fn_Config("text.TestsStopped")
    If VBA.Len(diagnostic) > 0 Then
        text = ex_PEB_Export.fn_Config("text.TestsRunError")
        If m_error = 0 Then m_error = 1
    End If
    text = VBA.Replace$(text, "{summary}", private_Summary())
    If VBA.Len(diagnostic) > 0 Then text = text & " " & diagnostic
    private_Progress text
    If Not Application.ActiveWorkbook Is ThisWorkbook Then private_RestoreStatus
    If VBA.Len(diagnostic) > 0 Then ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation, ex_PEB_Export.fn_Config("text.TestTitle")
End Sub

Private Sub private_RestoreStatus()
    If Not m_ownsStatus Then Exit Sub
    If VBA.Len(m_writtenStatus) > 0 Then
        If VBA.CStr(Application.StatusBar) = m_writtenStatus Then Application.StatusBar = m_previousStatus
    End If
    m_ownsStatus = False
End Sub

Private Sub private_SetRunning(ByVal running As Boolean)
    Dim context As obj_UiRenderContext

    If m_sheet Is Nothing Then Exit Sub
    If ex_UiRuntime.fn_TryGetContext(m_sheet, context) Then
        If Not context.BindingContext.SetValue("Settings", "TestsRunning", running) Then Err.Raise 5, , "Cannot update test runner state."
    End If
End Sub

Private Sub private_Progress(ByVal state As String)
    Dim context As obj_UiRenderContext
    Dim percent As Long
    Dim style As String

    If m_total > 0 Then percent = VBA.CLng(VBA.Fix((100# * m_completed + m_phasePercent) / m_total))
    m_status = ex_PEB_Export.fn_Config("text.TestsProgress")
    m_status = VBA.Replace$(m_status, "{percent}", VBA.CStr(percent))
    m_status = VBA.Replace$(m_status, "{completed}", VBA.CStr(m_completed))
    m_status = VBA.Replace$(m_status, "{total}", VBA.CStr(m_total))
    m_status = VBA.Replace$(m_status, "{state}", state)
    If Application.ActiveWorkbook Is ThisWorkbook Then
        Application.StatusBar = m_status
        m_writtenStatus = m_status
    End If
    If ex_UiRuntime.fn_TryGetContext(m_sheet, context) Then
        If Not context.BindingContext.SetValue("Form", "ActionStatus", m_status) Then Err.Raise 5, , "Cannot update test progress."
        style = "actionSuccess"
        If m_fail > 0 Or m_error > 0 Then style = "actionError"
        If Not context.BindingContext.SetValue("Form", "ActionStatusStyle", style) Then Err.Raise 5, , "Cannot update test progress style."
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

Private Sub private_CreateCheckbox( _
    ByVal index As Long, _
    ByVal row As Long, _
    ByVal selected As Boolean _
)
    Dim context As obj_UiRenderContext
    Dim document As Object
    Dim definition As Object
    Dim control As obj_IUiControl
    Dim diagnostic As String
    Dim rows As Long
    Dim columns As Long

    If Not ex_UiRuntime.fn_TryGetContext(m_sheet, context) Then Err.Raise 5, , "Test page render context is missing."
    If index = 1 Then
        If Not m_selectionCommand Is Nothing Then m_selectionCommand.Dispose
        Set m_selectionCommand = New obj_UiCommand
        If Not m_selectionCommand.Initialize(Me, "SelectionChanged") Then Err.Raise 5, , "Cannot initialize test selection command."
        If Not context.BindingContext.SetObject("Commands", "TestSelectionChanged", m_selectionCommand) Then Err.Raise 5, , "Cannot bind test selection command."
        If Not context.BindingContext.SetValue("TestSelection", "Caption", VBA.vbNullString) Then Err.Raise 5, , "Cannot bind test checkbox caption."
    End If
    If Not context.BindingContext.SetValue("TestSelection", "C" & VBA.CStr(index), selected) Then Err.Raise 5, , "Cannot bind test checkbox value."
    Set document = VBA.CreateObject("Msxml2.DOMDocument.6.0")
    Set definition = document.createElement("control")
    definition.setAttribute "name", "pebt_C" & VBA.CStr(index)
    definition.setAttribute "type", "checkbox"
    definition.setAttribute "row", "1"
    definition.setAttribute "column", "1"
    definition.setAttribute "style", "testCheckbox"
    definition.setAttribute "value", "{Binding Path=TestSelection.C" & VBA.CStr(index) & "}"
    definition.setAttribute "caption", "{Binding Path=TestSelection.Caption}"
    definition.setAttribute "checkedGlyph", "{Binding Path=Text.CheckboxChecked}"
    definition.setAttribute "uncheckedGlyph", "{Binding Path=Text.CheckboxUnchecked}"
    definition.setAttribute "onChange", "{Binding Path=Commands.TestSelectionChanged}"
    Set control = ex_UiControlFactory.fn_Create(definition)
    If control Is Nothing Then Err.Raise 5, , "Checkbox control factory is unavailable."
    m_checkboxes.Add control
    If Not control.Configure(definition, context, VBA.vbNullString, diagnostic) Then Err.Raise 5, , diagnostic
    If Not control.Measure(rows, columns, diagnostic) Then Err.Raise 5, , diagnostic
    If Not control.Arrange(row, 2, diagnostic) Then Err.Raise 5, , diagnostic
    If Not control.Render(diagnostic) Then Err.Raise 5, , diagnostic
End Sub

Private Sub private_RefreshCheckboxes()
    Dim context As obj_UiRenderContext
    Dim target As obj_IUiBindingTarget
    Dim index As Long

    If m_checkboxes Is Nothing Then Exit Sub
    If Not ex_UiRuntime.fn_TryGetContext(m_sheet, context) Then Err.Raise 5, , "Test page render context is missing."
    For index = 1 To m_checkboxes.Count
        If Not context.BindingContext.SetValue("TestSelection", "C" & VBA.CStr(index), m_selected.Exists(m_catalog(index)("Id"))) Then Err.Raise 5, , "Cannot refresh test selection."
        Set target = m_checkboxes(index)
        If Not target.RefreshBindings(context) Then Err.Raise 5, , "Cannot refresh test checkbox."
    Next index
End Sub

Private Sub private_ClearCheckboxes()
    Dim context As obj_UiRenderContext
    Dim control As obj_IUiControl
    Dim index As Long
    Dim shapeName As String

    If m_checkboxes Is Nothing Then Exit Sub
    ex_UiRuntime.fn_TryGetContext m_sheet, context
    For index = 1 To m_checkboxes.Count
        shapeName = "chk_pebt_C" & VBA.CStr(index)
        If Not context Is Nothing Then context.Router.UnregisterShape shapeName
        Set control = m_checkboxes(index)
        control.Dispose
    Next index
    For index = m_sheet.Shapes.Count To 1 Step -1
        If VBA.Left$(m_sheet.Shapes(index).Name, 10) = "chk_pebt_C" Then m_sheet.Shapes(index).Delete
    Next index
    Set m_checkboxes = Nothing
End Sub


Private Sub private_DrawResults()
    Dim context As obj_UiRenderContext
    Dim rawTable As obj_UiRawTable
    Dim fields As Variant
    Dim captions As Variant
    Dim values As Variant
    Dim report As Object
    Dim index As Long
    Dim fieldIndex As Long
    Dim rowCount As Long
    Dim diagnostic As String

    If m_sheet Is Nothing Then Exit Sub
    fields = VBA.Split(ex_PEB_Export.fn_Config("test.panel.fields"), "|")
    captions = VBA.Split(ex_PEB_Export.fn_Config("test.panel.captions"), "|")
    If UBound(fields) <> UBound(captions) Then Err.Raise 5, , "Test panel fields differ."
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
        If Not rawTable.Initialize(values, captions, VBA.vbNullString, fields) Then Err.Raise 5, , "Cannot initialize test result data."
        If Not rawTable.SetRecords(m_reports) Then Err.Raise 5, , "Cannot attach result row records."
    Else
        If Not rawTable.InitializeEmpty(captions, VBA.vbNullString, fields) Then Err.Raise 5, , "Cannot initialize empty test results."
    End If
    If Not ex_UiRuntime.fn_TryGetContext(m_sheet, context) Then Err.Raise 5, , "Test page render context is missing."
    private_BindPreviews context
    If Not context.BindingContext.SetObject("Data", "TestResults", rawTable) Then Err.Raise 5, , "Cannot bind test results."
    context.MarkDirty "TestResults", True
    If Not context.FlushUpdates(diagnostic) Then Err.Raise 5, , diagnostic
End Sub

Private Sub private_BindPreviews(ByVal context As obj_UiRenderContext)
    Dim previews As Collection
    Dim report As Object
    Dim item As Object
    Dim text As String

    Set previews = New Collection
    For Each report In m_reports
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
    Next report
    If Not context.BindingContext.SetObject("Data", "TestPreviews", previews) Then Err.Raise 5, , "Cannot bind test previews."
    context.MarkDirty "TestDocuments", True
End Sub