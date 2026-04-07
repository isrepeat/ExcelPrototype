Attribute VB_Name = "ex_ControlRuntime"
Option Explicit

Private Const UI_NS As String = "urn:excelprototype:profiles"
Private Const SHEET_UI_BASE_REL_PATH As String = "ui\"
Private Const SHEET_UI_FILE_SUFFIX As String = "UI.xml"
Private Const CONTROL_UI_BASE_REL_PATH As String = "vba\classes\"
Private Const CONTROL_UI_FILE_PREFIX As String = "obj_"
Private Const CONTROL_UI_FILE_SUFFIX As String = "ControlUI.xml"

Public Sub m_TEST_RenderDevUI()
    Dim wb As Workbook
    Dim activeSheetObj As Object
    Dim ws As Worksheet

    Set wb = ThisWorkbook
    If wb Is Nothing Then
        MsgBox "PrototypeNew: workbook is not specified.", vbExclamation
        Exit Sub
    End If

    Set activeSheetObj = wb.ActiveSheet
    If activeSheetObj Is Nothing Then
        MsgBox "PrototypeNew: active sheet is not specified.", vbExclamation
        Exit Sub
    End If

    If Not TypeOf activeSheetObj Is Worksheet Then
        MsgBox "PrototypeNew: active sheet is not a worksheet.", vbExclamation
        Exit Sub
    End If

    Set ws = activeSheetObj
    m_RenderWorksheet ws, "ui\DevUI.xml"
End Sub

Public Sub m_RenderWorksheet(ByVal ws As Worksheet, Optional ByVal wsUiPath As String = vbNullString)
    Dim wb As Workbook
    Dim resolvedWsUiPath As String

    If ws Is Nothing Then
        MsgBox "PrototypeNew: worksheet is not specified.", vbExclamation
        Exit Sub
    End If

    Set wb = ws.Parent
    If wb Is Nothing Then
        MsgBox "PrototypeNew: workbook is not specified.", vbExclamation
        Exit Sub
    End If

    resolvedWsUiPath = mp_ResolveWsUiPath(ws, wsUiPath)
    mp_RenderLayout wb, ws, resolvedWsUiPath
End Sub

Private Sub mp_RenderLayout( _
    ByVal wb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal wsUiRelPath As String _
)
    Dim devUiDoc As Object
    Dim layoutControlNodes As Object
    Dim layoutControlNode As Object
    Dim layoutControlName As String
    Dim controlType As String
    Dim typeRoot As String
    Dim controlUiRelPath As String
    Dim runtimeControlNode As Object
    Dim control As obj_IControl

    Set devUiDoc = ex_XmlCore.m_LoadDomByRelativePath( _
        wb, _
        wsUiRelPath, _
        "PrototypeNew: page UI file was not found: ", _
        "PrototypeNew: failed to parse page UI file: ", _
        UI_NS)
    If devUiDoc Is Nothing Then Exit Sub

    Set layoutControlNodes = devUiDoc.selectNodes("/p:uiDefinition/p:layout//p:control")
    If layoutControlNodes Is Nothing Then
        MsgBox "PrototypeNew: invalid page UI format, controls path not found.", vbExclamation
        Exit Sub
    End If
    If layoutControlNodes.Length = 0 Then
        MsgBox "PrototypeNew: no controls found in page UI '" & wsUiRelPath & "'.", vbExclamation
        Exit Sub
    End If

    For Each layoutControlNode In layoutControlNodes
        layoutControlName = Trim$(ex_XmlCore.m_NodeAttrText(layoutControlNode, "name"))
        controlType = Trim$(ex_XmlCore.m_NodeAttrText(layoutControlNode, "type"))
        typeRoot = mp_NormalizeTypeRoot(controlType)

        If Len(layoutControlName) = 0 Then
            MsgBox "PrototypeNew: page control is missing required attribute 'name'.", vbExclamation
            Exit Sub
        End If
        If Len(controlType) = 0 Then
            MsgBox "PrototypeNew: page control '" & layoutControlName & "' is missing required attribute 'type'.", vbExclamation
            Exit Sub
        End If
        If Len(typeRoot) = 0 Then
            MsgBox "PrototypeNew: page control '" & layoutControlName & "' has invalid type '" & controlType & "'.", vbExclamation
            Exit Sub
        End If

        controlUiRelPath = mp_ResolveControlUiRelPathByTypeRoot(typeRoot)

        Set runtimeControlNode = mp_LoadControlNodeFromControlUi(wb, controlUiRelPath, layoutControlName, ws.Name)
        If runtimeControlNode Is Nothing Then Exit Sub

        Set control = ex_ControlFactory.m_CreateControlByTypeRoot(typeRoot)
        If control Is Nothing Then Exit Sub

        control.Configure runtimeControlNode
        control.Render wb
    Next layoutControlNode
End Sub

Private Function mp_LoadControlNodeFromControlUi( _
    ByVal wb As Workbook, _
    ByVal controlUiRelPath As String, _
    ByVal controlName As String, _
    ByVal sheetName As String _
) As Object
    Dim uiDoc As Object
    Dim escapedName As String
    Dim xPath As String

    Set uiDoc = ex_XmlCore.m_LoadDomByRelativePath( _
        wb, _
        controlUiRelPath, _
        "PrototypeNew: control UI file was not found: ", _
        "PrototypeNew: failed to parse control UI file: ", _
        UI_NS)
    If uiDoc Is Nothing Then Exit Function

    escapedName = ex_XmlCore.m_XPathLiteral(controlName)
    xPath = "/p:uiDefinition/p:layout//p:control[@name=" & escapedName & "]"
    Set mp_LoadControlNodeFromControlUi = uiDoc.selectSingleNode(xPath)

    If mp_LoadControlNodeFromControlUi Is Nothing Then
        Set mp_LoadControlNodeFromControlUi = uiDoc.selectSingleNode("/p:uiDefinition/p:layout//p:control[1]")
    End If

    If mp_LoadControlNodeFromControlUi Is Nothing Then
        MsgBox "PrototypeNew: control node not found in UI file '" & controlUiRelPath & "'.", vbExclamation
        Exit Function
    End If

    On Error Resume Next
    mp_LoadControlNodeFromControlUi.setAttribute "name", controlName
    If Len(Trim$(sheetName)) > 0 Then
        mp_LoadControlNodeFromControlUi.setAttribute "sheet", sheetName
    End If
    On Error GoTo 0
End Function

Private Function mp_ResolveWsUiPath(ByVal ws As Worksheet, ByVal wsUiPath As String) As String
    wsUiPath = Trim$(wsUiPath)
    If Len(wsUiPath) > 0 Then
        mp_ResolveWsUiPath = wsUiPath
        Exit Function
    End If

    mp_ResolveWsUiPath = SHEET_UI_BASE_REL_PATH & ws.Name & SHEET_UI_FILE_SUFFIX
End Function

Private Function mp_GetWorksheetByName(ByVal wb As Workbook, ByVal wsName As String) As Worksheet
    On Error Resume Next
    Set mp_GetWorksheetByName = wb.Worksheets(wsName)
    On Error GoTo 0
End Function

Private Function mp_NormalizeTypeRoot(ByVal controlType As String) As String
    mp_NormalizeTypeRoot = Trim$(controlType)
End Function

Private Function mp_ResolveControlUiRelPathByTypeRoot(ByVal typeRoot As String) As String
    mp_ResolveControlUiRelPathByTypeRoot = _
        CONTROL_UI_BASE_REL_PATH & CONTROL_UI_FILE_PREFIX & typeRoot & CONTROL_UI_FILE_SUFFIX
End Function
