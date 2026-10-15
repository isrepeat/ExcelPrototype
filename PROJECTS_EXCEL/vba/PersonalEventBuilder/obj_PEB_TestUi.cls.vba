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
Private m_catalog As Collection
Private m_selected As Object
Private m_runSelection As Object
Private m_runner As obj_PEB_TestRunner
Private m_reports As Collection
Private m_running As Boolean
Private m_stop As Boolean
Private m_baseRow As Long
Private m_resultRow As Long
Private m_endRow As Long
Private m_completed As Long
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

    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
    m_stop = True
    On Error Resume Next
    ex_PEB_Testing.fn_CancelTick
    If Not m_runner Is Nothing Then m_runner.Dispose
    Set m_runner = Nothing
    m_running = False
    private_SetRunning False
    private_RestoreStatus
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
    Dim reader As obj_PEB_TestRunner
    Dim retained As Object
    Dim scenario As Object
    Dim diagnostic As String
    Dim previousScreen As Boolean
    Dim previousEvents As Boolean
    Dim index As Long
    Dim rowIndex As Long
    Dim checkbox As Shape
    Dim captions As Variant

    private_Guard
    If m_running Then Exit Function
    On Error GoTo EH_READ
    Set reader = New obj_PEB_TestRunner
    If Not reader.Initialize(diagnostic) Then Err.Raise 5, , diagnostic
    Set m_catalog = reader.Catalog
    reader.Dispose
    Set reader = Nothing
    If m_catalog.Count > ex_PEB_Export.fn_Count("test.ui.maxScenarios") Then Err.Raise 5, , "Scenario catalog exceeds the UI limit."
    Set retained = VBA.CreateObject("Scripting.Dictionary")
    For Each scenario In m_catalog
        If m_selected.Exists(scenario("Id")) Then retained.Add scenario("Id"), True
    Next scenario
    Set m_selected = retained
    Set m_sheet = ThisWorkbook.Worksheets(ex_PEB_Export.fn_Config("test.ui.mainSheet"))
    m_baseRow = ex_PEB_Export.fn_Count("test.panel.row")
    index = ex_PEB_Export.fn_Count("lookup.Personnel.maxResults") + ex_PEB_Export.fn_Count("test.panel.candidateOffset")
    If m_baseRow < index Then m_baseRow = index
    index = ex_PEB_Export.fn_Count("lookup.Hospital.maxResults") + ex_PEB_Export.fn_Count("test.panel.candidateOffset")
    If m_baseRow < index Then m_baseRow = index
    previousScreen = Application.ScreenUpdating
    previousEvents = Application.EnableEvents
    On Error GoTo EH_RENDER
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    If m_endRow >= m_baseRow Then m_sheet.Range(m_sheet.Cells(m_baseRow, 2), m_sheet.Cells(m_endRow, 17)).Clear
    For index = m_sheet.Shapes.Count To 1 Step -1
        If VBA.Left$(m_sheet.Shapes(index).Name, 5) = "pebt_" Then m_sheet.Shapes(index).Delete
    Next index
    m_resultRow = m_baseRow + m_catalog.Count + 7
    m_endRow = m_resultRow + ex_PEB_Export.fn_Count("test.panel.displayRows")
    m_sheet.Range(m_sheet.Cells(m_baseRow, 2), m_sheet.Cells(m_endRow, 17)).UnMerge
    m_sheet.Range(m_sheet.Cells(m_baseRow, 2), m_sheet.Cells(m_endRow, 17)).NumberFormat = "@"
    private_Cell m_baseRow, 2, 16, ex_PEB_Export.fn_Config("text.TestsPageTitle")
    private_Button "All", m_baseRow + 1, 2, 2, "TestsSelectAll"
    private_Button "None", m_baseRow + 1, 4, 2, "TestsSelectNone"
    private_Button "Run", m_baseRow + 1, 6, 4, "TestsRun"
    private_Button "Stop", m_baseRow + 1, 10, 4, "TestsStop"
    private_Button "Refresh", m_baseRow + 1, 14, 4, "TestsRefresh"
    private_Cell m_baseRow + 2, 2, 16, ex_PEB_Export.fn_Config("text.TestsHelp")
    rowIndex = m_baseRow + 3
    captions = VBA.Split(ex_PEB_Export.fn_Config("test.ui.catalogCaptions"), "|")
    If UBound(captions) <> 4 Then Err.Raise 5, , "Test catalog captions must contain five columns."
    private_Cell rowIndex, 2, 1, VBA.CStr(captions(0))
    private_Cell rowIndex, 3, 1, VBA.CStr(captions(1))
    private_Cell rowIndex, 4, 8, VBA.CStr(captions(2))
    private_Cell rowIndex, 12, 2, VBA.CStr(captions(3))
    private_Cell rowIndex, 14, 4, VBA.CStr(captions(4))
    rowIndex = rowIndex + 1
    For index = 1 To m_catalog.Count
        Set scenario = m_catalog(index)
        private_Cell rowIndex, 3, 1, scenario("Id")
        private_Cell rowIndex, 4, 8, scenario("Name")
        private_Cell rowIndex, 12, 2, VBA.CStr(scenario("Steps").Count)
        private_Cell rowIndex, 14, 4, scenario("RuleStatus")
        Set checkbox = m_sheet.Shapes.AddFormControl(xlCheckBox, m_sheet.Cells(rowIndex, 2).Left + 3, _
            m_sheet.Cells(rowIndex, 2).Top + 3, m_sheet.Cells(rowIndex, 2).Width - 6, m_sheet.Cells(rowIndex, 2).Height - 6)
        checkbox.Name = "pebt_C" & VBA.CStr(index)
        checkbox.TextFrame.Characters.Text = VBA.vbNullString
        checkbox.Placement = xlMoveAndSize
        checkbox.OnAction = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!ex_PEB_Testing.fn_Click"
        checkbox.ControlFormat.Value = xlOff
        If m_selected.Exists(scenario("Id")) Then checkbox.ControlFormat.Value = xlOn
        rowIndex = rowIndex + 1
    Next index
    private_Cell m_resultRow - 1, 2, 16, ex_PEB_Export.fn_Config("text.TestsResults")
    private_ResultHeader
    private_DrawResults
    m_sheet.Activate
    Application.Goto m_sheet.Cells(m_baseRow, 2), True
    OpenPage = True
CleanRender:
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreen
    If VBA.Len(diagnostic) > 0 Then Err.Raise 5, , diagnostic
    Exit Function
EH_RENDER:
    diagnostic = VBA.Err.Description
    Resume CleanRender
EH_READ:
    diagnostic = VBA.Err.Description
    If Not reader Is Nothing Then reader.Dispose
    Err.Raise 5, , diagnostic
End Function

Public Sub ToggleCase(ByVal index As Long)
    Dim id As String

    private_Guard
    If m_running Or m_catalog Is Nothing Then Exit Sub
    If index < 1 Or index > m_catalog.Count Then Err.Raise 5, , "Invalid scenario checkbox."
    id = m_catalog(index)("Id")
    If m_selected.Exists(id) Then m_selected.Remove id
    If m_sheet.Shapes("pebt_C" & VBA.CStr(index)).ControlFormat.Value = xlOn Then m_selected.Add id, True
End Sub

Public Sub SelectAll(ByVal selected As Boolean)
    Dim scenario As Object
    Dim index As Long

    private_Guard
    If m_running Or m_catalog Is Nothing Then Exit Sub
    Set m_selected = VBA.CreateObject("Scripting.Dictionary")
    For index = 1 To m_catalog.Count
        Set scenario = m_catalog(index)
        m_sheet.Shapes("pebt_C" & VBA.CStr(index)).ControlFormat.Value = xlOff
        If selected Then
            m_selected.Add scenario("Id"), True
            m_sheet.Shapes("pebt_C" & VBA.CStr(index)).ControlFormat.Value = xlOn
        End If
    Next index
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
    Dim index As Long

    If m_sheet Is Nothing Then Exit Sub
    If ex_UiRuntime.fn_TryGetContext(m_sheet, context) Then
        If Not context.BindingContext.SetValue("Settings", "TestsRunning", running) Then Err.Raise 5, , "Cannot update test runner state."
    End If
    If Not m_catalog Is Nothing Then
        For index = 1 To m_catalog.Count
            m_sheet.Shapes("pebt_C" & VBA.CStr(index)).ControlFormat.Enabled = Not running
        Next index
    End If
End Sub

Private Sub private_Progress(ByVal state As String)
    Dim context As obj_UiRenderContext
    Dim percent As Long
    Dim style As String

    If m_total > 0 Then percent = VBA.CLng(VBA.Fix(100# * m_completed / m_total))
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

Private Sub private_Cell( _
    ByVal row As Long, _
    ByVal column As Long, _
    ByVal span As Long, _
    ByVal value As String _
)
    Dim area As Range

    Set area = m_sheet.Cells(row, column).Resize(1, span)
    area.UnMerge
    If span > 1 Then area.Merge
    area.NumberFormat = "@"
    area.Cells(1, 1).Value2 = value
    area.WrapText = True
    area.VerticalAlignment = xlCenter
    m_sheet.Rows(row).RowHeight = ex_PEB_Export.fn_Count("test.ui.rowHeight")
End Sub

Private Sub private_Button( _
    ByVal name As String, _
    ByVal row As Long, _
    ByVal column As Long, _
    ByVal span As Long, _
    ByVal caption As String _
)
    Dim area As Range
    Dim shape As Shape

    Set area = m_sheet.Cells(row, column).Resize(1, span)
    Set shape = m_sheet.Shapes.AddShape(msoShapeRectangle, area.Left, area.Top, area.Width, area.Height)
    shape.Name = "pebt_" & name
    shape.TextFrame2.TextRange.Text = ex_PEB_Export.fn_Config("text." & caption)
    shape.OnAction = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!ex_PEB_Testing.fn_Click"
    shape.Placement = xlMoveAndSize
End Sub

Private Sub private_ResultHeader()
    Dim captions As Variant
    Dim spans As Variant
    Dim index As Long
    Dim column As Long

    captions = VBA.Split(ex_PEB_Export.fn_Config("test.panel.captions"), "|")
    spans = VBA.Split(ex_PEB_Export.fn_Config("test.panel.spans"), "|")
    If UBound(captions) <> UBound(spans) Then Err.Raise 5, , "Test panel columns differ."
    column = 2
    For index = 0 To UBound(captions)
        private_Cell m_resultRow, column, VBA.CLng(spans(index)), VBA.CStr(captions(index))
        column = column + VBA.CLng(spans(index))
    Next index
    If column <> 18 Then Err.Raise 5, , "Test panel must occupy columns B:Q."
End Sub

Private Sub private_DrawResults()
    Dim fields As Variant
    Dim spans As Variant
    Dim report As Object
    Dim index As Long
    Dim fieldIndex As Long
    Dim first As Long
    Dim row As Long
    Dim column As Long

    If m_sheet Is Nothing Or m_resultRow = 0 Then Exit Sub
    fields = VBA.Split(ex_PEB_Export.fn_Config("test.panel.fields"), "|")
    spans = VBA.Split(ex_PEB_Export.fn_Config("test.panel.spans"), "|")
    If UBound(fields) <> UBound(spans) Then Err.Raise 5, , "Test panel fields differ."
    m_sheet.Range(m_sheet.Cells(m_resultRow + 1, 2), m_sheet.Cells(m_endRow, 17)).ClearContents
    first = m_reports.Count - ex_PEB_Export.fn_Count("test.panel.displayRows") + 1
    If first < 1 Then first = 1
    row = m_resultRow + 1
    For index = first To m_reports.Count
        Set report = m_reports(index)
        column = 2
        For fieldIndex = 0 To UBound(fields)
            private_Cell row, column, VBA.CLng(spans(fieldIndex)), ex_PEB_Export.fn_Text(report(fields(fieldIndex)))
            column = column + VBA.CLng(spans(fieldIndex))
        Next fieldIndex
        row = row + 1
    Next index
End Sub