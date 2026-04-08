Attribute VB_Name = "ex_SheetRenderer"
Option Explicit

Private Const UI_NS As String = "urn:excelprototype:profiles"
Private Const SHEET_UI_BASE_REL_PATH As String = "ui\"
Private Const SHEET_UI_FILE_SUFFIX As String = "UI.xml"

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

    If ws Is Nothing Then
        MsgBox "PrototypeNew: worksheet is not specified.", vbExclamation
        Exit Sub
    End If

    Set wb = ws.Parent
    If wb Is Nothing Then
        MsgBox "PrototypeNew: workbook is not specified.", vbExclamation
        Exit Sub
    End If

    Set app = Application
    mp_EnterFastRenderMode app, prevScreenUpdating, prevEnableEvents, prevDisplayAlerts, prevCalculation, prevStatusBar

    On Error GoTo EH_RENDER

    resolvedWsUiPath = mp_ResolveWsUiPath(ws, wsUiPath)

    Set wsUiDoc = ex_XmlCore.m_LoadDomByRelativePath( _
        wb, _
        resolvedWsUiPath, _
        "PrototypeNew: page UI file was not found: ", _
        "PrototypeNew: failed to parse page UI file: ", _
        UI_NS)
    If wsUiDoc Is Nothing Then GoTo Cleanup

    If Not ex_XmlLayoutEngine.m_RenderPageLayout(wb, ws, wsUiDoc) Then GoTo Cleanup
    If Not ex_StylePipelineEngine.m_ApplyPageStyles(ws, wsUiDoc) Then GoTo Cleanup

Cleanup:
    mp_LeaveFastRenderMode app, prevScreenUpdating, prevEnableEvents, prevDisplayAlerts, prevCalculation, prevStatusBar
    Exit Sub

EH_RENDER:
    mp_LeaveFastRenderMode app, prevScreenUpdating, prevEnableEvents, prevDisplayAlerts, prevCalculation, prevStatusBar
    MsgBox "PrototypeNew: render failed: [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description, vbExclamation
End Sub

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
