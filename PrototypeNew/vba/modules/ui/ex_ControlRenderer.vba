Attribute VB_Name = "ex_ControlRenderer"
Option Explicit

Private Const UI_NS As String = "urn:excelprototype:profiles"
Private Const CONTROL_UI_BASE_REL_PATH As String = "vba\classes\"
Private Const CONTROL_UI_FILE_PREFIX As String = "obj_"
Private Const CONTROL_UI_FILE_SUFFIX As String = "ControlUI.xml"
Private Const MAX_TEMPLATE_RECURSION_DEPTH As Long = 12

Public Function m_RenderControl( _
    ByVal wb As Workbook, _
    ByVal ws As Worksheet, _
    ByVal layoutControlNode As Object, _
    Optional ByVal recursionDepth As Long = 0, _
    Optional ByVal layoutRowStart As Long = 0, _
    Optional ByVal layoutColStart As Long = 0, _
    Optional ByVal layoutRowEnd As Long = 0, _
    Optional ByVal layoutColEnd As Long = 0 _
) As Boolean
    Dim layoutControlName As String
    Dim controlType As String
    Dim typeRoot As String
    Dim controlUiRelPath As String
    Dim runtimeControlNode As Object
    Dim control As obj_IControl

    If recursionDepth > MAX_TEMPLATE_RECURSION_DEPTH Then
        MsgBox "PrototypeNew: control template recursion depth exceeded for node.", vbExclamation
        Exit Function
    End If

    If wb Is Nothing Then
        MsgBox "PrototypeNew: workbook is not specified.", vbExclamation
        Exit Function
    End If
    If ws Is Nothing Then
        MsgBox "PrototypeNew: worksheet is not specified.", vbExclamation
        Exit Function
    End If
    If layoutControlNode Is Nothing Then
        MsgBox "PrototypeNew: control node is not specified.", vbExclamation
        Exit Function
    End If

    layoutControlName = Trim$(ex_XmlCore.m_NodeAttrText(layoutControlNode, "name"))
    controlType = Trim$(ex_XmlCore.m_NodeAttrText(layoutControlNode, "type"))
    typeRoot = mp_NormalizeTypeRoot(controlType)

    If Len(layoutControlName) = 0 Then
        MsgBox "PrototypeNew: page control is missing required attribute 'name'.", vbExclamation
        Exit Function
    End If
    If Len(controlType) = 0 Then
        MsgBox "PrototypeNew: page control '" & layoutControlName & "' is missing required attribute 'type'.", vbExclamation
        Exit Function
    End If
    If Len(typeRoot) = 0 Then
        MsgBox "PrototypeNew: page control '" & layoutControlName & "' has invalid type '" & controlType & "'.", vbExclamation
        Exit Function
    End If

    Set control = ex_ControlFactory.m_CreateControlByTypeRoot(typeRoot)
    If control Is Nothing Then Exit Function

    controlUiRelPath = mp_ResolveControlUiRelPathByTypeRoot(typeRoot)
    Set runtimeControlNode = mp_LoadControlNodeFromControlUi( _
        wb, controlUiRelPath, layoutControlNode, control, layoutControlName, typeRoot)
    If runtimeControlNode Is Nothing Then Exit Function

    mp_ApplyRuntimeLayoutBounds runtimeControlNode, ws.Name, layoutRowStart, layoutColStart, layoutRowEnd, layoutColEnd

    control.Configure runtimeControlNode
    control.Render wb

    If Not ex_XmlLayoutEngine.m_RenderTemplateChildren( _
        wb, ws, runtimeControlNode, recursionDepth + 1, _
        layoutRowStart, layoutColStart, layoutRowEnd, layoutColEnd) Then Exit Function

    m_RenderControl = True
End Function

Private Function mp_LoadControlNodeFromControlUi( _
    ByVal wb As Workbook, _
    ByVal controlUiRelPath As String, _
    ByVal layoutControlNode As Object, _
    ByVal control As obj_IControl, _
    ByVal controlName As String, _
    ByVal typeRoot As String _
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
        MsgBox "PrototypeNew: control template has no <control> node in UI file '" & controlUiRelPath & "'.", vbExclamation
        Exit Function
    End If

    If Not mp_ApplyLayoutControlOverridesByContract( _
        mp_LoadControlNodeFromControlUi, layoutControlNode, control, controlName, typeRoot) Then
        Set mp_LoadControlNodeFromControlUi = Nothing
        Exit Function
    End If

    On Error Resume Next
    mp_LoadControlNodeFromControlUi.setAttribute "name", controlName
    On Error GoTo 0
End Function

Private Function mp_ApplyLayoutControlOverridesByContract( _
    ByVal runtimeControlNode As Object, _
    ByVal layoutControlNode As Object, _
    ByVal control As obj_IControl, _
    ByVal controlName As String, _
    ByVal typeRoot As String _
) As Boolean
    Dim layoutAttrs As Object
    Dim attrNode As Object
    Dim attrName As String

    If runtimeControlNode Is Nothing Then Exit Function
    If layoutControlNode Is Nothing Then Exit Function
    If control Is Nothing Then Exit Function

    Set layoutAttrs = layoutControlNode.selectNodes("@*")
    If layoutAttrs Is Nothing Then Exit Function

    For Each attrNode In layoutAttrs
        attrName = CStr(attrNode.nodeName)

        If mp_IsLayoutAttribute(attrName) Then GoTo ContinueLoop

        If Not ex_ControlAttributeContracts.m_IsSupportedControlAttribute(control, attrName) Then
            MsgBox "PrototypeNew: attribute '" & attrName & "' is not supported by control '" & controlName & "' of type '" & typeRoot & "'.", vbExclamation
            Exit Function
        End If

        On Error Resume Next
        runtimeControlNode.setAttribute attrName, CStr(attrNode.Text)
        If Err.Number <> 0 Then
            MsgBox "PrototypeNew: failed to apply attribute '" & attrName & "' to control '" & controlName & "': " & Err.Description, vbExclamation
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0

ContinueLoop:
    Next attrNode

    mp_ApplyLayoutControlOverridesByContract = True
End Function

Private Function mp_IsLayoutAttribute(ByVal attrName As String) As Boolean
    Select Case LCase$(Trim$(attrName))
        Case "at", "spancells", "spanrows"
            mp_IsLayoutAttribute = True
    End Select
End Function

Private Sub mp_ApplyRuntimeLayoutBounds( _
    ByVal runtimeControlNode As Object, _
    ByVal sheetName As String, _
    ByVal layoutRowStart As Long, _
    ByVal layoutColStart As Long, _
    ByVal layoutRowEnd As Long, _
    ByVal layoutColEnd As Long _
)
    runtimeControlNode.setAttribute "__layoutSheet", sheetName

    If layoutRowStart > 0 Then runtimeControlNode.setAttribute "__layoutRowStart", CStr(layoutRowStart)
    If layoutColStart > 0 Then runtimeControlNode.setAttribute "__layoutColStart", CStr(layoutColStart)
    If layoutRowEnd > 0 Then runtimeControlNode.setAttribute "__layoutRowEnd", CStr(layoutRowEnd)
    If layoutColEnd > 0 Then runtimeControlNode.setAttribute "__layoutColEnd", CStr(layoutColEnd)
End Sub

Private Function mp_NormalizeTypeRoot(ByVal controlType As String) As String
    mp_NormalizeTypeRoot = Trim$(controlType)
End Function

Private Function mp_ResolveControlUiRelPathByTypeRoot(ByVal typeRoot As String) As String
    mp_ResolveControlUiRelPathByTypeRoot = _
        CONTROL_UI_BASE_REL_PATH & CONTROL_UI_FILE_PREFIX & typeRoot & CONTROL_UI_FILE_SUFFIX
End Function
