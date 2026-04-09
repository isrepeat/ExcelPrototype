Attribute VB_Name = "ex_SheetRenderer"
Option Explicit

Private Const UI_NS As String = "urn:excelprototype:profiles"
Private Const SHEET_UI_BASE_REL_PATH As String = "ui\"
Private Const SHEET_UI_FILE_SUFFIX As String = "UI.xml"

Private g_LastRenderWorkbook As Workbook
Private g_LastRenderWorksheet As Worksheet
Private g_LastRenderUiPath As String
Private g_IsRendering As Boolean

Public Sub m_RenderWorksheet(ByVal ws As Worksheet, Optional ByVal wsUiPath As String = vbNullString)
    Dim app As Application
    Dim wb As Workbook
    Dim resolvedWsUiPath As String
    Dim wsUiDoc As Object
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevCalculation As XlCalculation
    Dim prevStatusBar As Variant
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    If ws Is Nothing Then
        MsgBox "PrototypeNew: worksheet is not specified.", vbExclamation
        Exit Sub
    End If
    If g_IsRendering Then Exit Sub

    Set wb = ws.Parent
    If wb Is Nothing Then
        MsgBox "PrototypeNew: workbook is not specified.", vbExclamation
        Exit Sub
    End If

    g_IsRendering = True
    Set app = Application
    mp_EnterFastRenderMode app, prevScreenUpdating, prevEnableEvents, prevDisplayAlerts, prevCalculation, prevStatusBar

    On Error GoTo EH_RENDER

    resolvedWsUiPath = mp_ResolveWsUiPath(ws, wsUiPath)
    Set g_LastRenderWorkbook = wb
    Set g_LastRenderWorksheet = ws
    g_LastRenderUiPath = resolvedWsUiPath

    Set wsUiDoc = ex_XmlCore.m_LoadDomByRelativePath( _
        wb, _
        resolvedWsUiPath, _
        "PrototypeNew: page UI file was not found: ", _
        "PrototypeNew: failed to parse page UI file: ", _
        UI_NS)
    If wsUiDoc Is Nothing Then GoTo Cleanup

    ex_ControlPartsRuntime.m_ResetControlParts
    ex_InlineTextRuntime.m_ResetInlineRuns
    If Not ex_XmlLayoutEngine.m_RenderPageLayout(wb, ws, wsUiDoc) Then GoTo Cleanup
    If Not ex_StylePipelineEngine.m_ApplyPageStyles(ws, wsUiDoc) Then GoTo Cleanup
    If Not ex_InlineTextRuntime.m_ApplyInlineRuns(ws) Then GoTo Cleanup

Cleanup:
    mp_LeaveFastRenderMode app, prevScreenUpdating, prevEnableEvents, prevDisplayAlerts, prevCalculation, prevStatusBar
    g_IsRendering = False
    Exit Sub

EH_RENDER:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description

    mp_LeaveFastRenderMode app, prevScreenUpdating, prevEnableEvents, prevDisplayAlerts, prevCalculation, prevStatusBar
    g_IsRendering = False
    MsgBox "PrototypeNew: render failed: [" & errSource & " #" & CStr(errNumber) & "] " & errDescription, vbExclamation
End Sub

Public Function m_TryRerenderLastRenderedPage(Optional ByVal reason As String = vbNullString) As Boolean
    Dim ws As Worksheet

    If g_IsRendering Then Exit Function

    On Error GoTo EH_CONTEXT
    Set ws = g_LastRenderWorksheet
    On Error GoTo 0

    If ws Is Nothing Then Exit Function
    If Len(Trim$(g_LastRenderUiPath)) = 0 Then Exit Function

    m_RenderWorksheet ws, g_LastRenderUiPath
    m_TryRerenderLastRenderedPage = True
    Exit Function

EH_CONTEXT:
    Set g_LastRenderWorkbook = Nothing
    Set g_LastRenderWorksheet = Nothing
    g_LastRenderUiPath = vbNullString
End Function

Public Sub m_ApplyWorksheetStyleStage( _
    ByVal ws As Worksheet, _
    ByVal stageName As String, _
    Optional ByVal wsUiPath As String = vbNullString _
)
    Dim wb As Workbook
    Dim resolvedWsUiPath As String
    Dim wsUiDoc As Object

    If ws Is Nothing Then
        MsgBox "PrototypeNew: worksheet is not specified.", vbExclamation
        Exit Sub
    End If

    stageName = Trim$(stageName)
    If Len(stageName) = 0 Then
        MsgBox "PrototypeNew: style stage name is required.", vbExclamation
        Exit Sub
    End If

    Set wb = ws.Parent
    If wb Is Nothing Then
        MsgBox "PrototypeNew: workbook is not specified.", vbExclamation
        Exit Sub
    End If

    resolvedWsUiPath = mp_ResolveWsUiPath(ws, wsUiPath)

    Set wsUiDoc = ex_XmlCore.m_LoadDomByRelativePath( _
        wb, _
        resolvedWsUiPath, _
        "PrototypeNew: page UI file was not found: ", _
        "PrototypeNew: failed to parse page UI file: ", _
        UI_NS)
    If wsUiDoc Is Nothing Then Exit Sub

    If Not ex_StylePipelineEngine.m_ApplyPageStyleStage(ws, wsUiDoc, stageName) Then Exit Sub
End Sub

Private Function mp_ResolveWsUiPath(ByVal ws As Worksheet, ByVal wsUiPath As String) As String
    wsUiPath = Trim$(wsUiPath)
    If Len(wsUiPath) > 0 Then
        mp_ResolveWsUiPath = wsUiPath
        Exit Function
    End If

    mp_ResolveWsUiPath = SHEET_UI_BASE_REL_PATH & ws.Name & SHEET_UI_FILE_SUFFIX
End Function

Private Sub mp_EnterFastRenderMode( _
    ByVal app As Application, _
    ByRef prevScreenUpdating As Boolean, _
    ByRef prevEnableEvents As Boolean, _
    ByRef prevDisplayAlerts As Boolean, _
    ByRef prevCalculation As XlCalculation, _
    ByRef prevStatusBar As Variant _
)
    If app Is Nothing Then Exit Sub

    prevScreenUpdating = app.ScreenUpdating
    prevEnableEvents = app.EnableEvents
    prevDisplayAlerts = app.DisplayAlerts
    prevCalculation = app.Calculation
    prevStatusBar = app.StatusBar

    app.ScreenUpdating = False
    app.EnableEvents = False
    app.DisplayAlerts = False
    app.Calculation = xlCalculationManual
    app.StatusBar = "PrototypeNew: rendering UI..."
End Sub

Private Sub mp_LeaveFastRenderMode( _
    ByVal app As Application, _
    ByVal prevScreenUpdating As Boolean, _
    ByVal prevEnableEvents As Boolean, _
    ByVal prevDisplayAlerts As Boolean, _
    ByVal prevCalculation As XlCalculation, _
    ByVal prevStatusBar As Variant _
)
    If app Is Nothing Then Exit Sub

    On Error Resume Next
    app.ScreenUpdating = prevScreenUpdating
    app.EnableEvents = prevEnableEvents
    app.DisplayAlerts = prevDisplayAlerts
    app.Calculation = prevCalculation
    app.StatusBar = prevStatusBar
    On Error GoTo 0
End Sub
