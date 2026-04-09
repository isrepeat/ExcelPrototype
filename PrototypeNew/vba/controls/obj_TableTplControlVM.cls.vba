VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_TableTplControlVM"
Option Explicit
Implements obj_IControl

Private m_ControlName As String
Private m_ItemsSourceRaw As String
Private m_IsConfigured As Boolean

Private Sub obj_IControl_Configure(ByVal controlNode As Object)
    Dim listNode As Object

    m_IsConfigured = False

    If controlNode Is Nothing Then
        MsgBox "TableTpl: control node is not specified.", vbExclamation
        Exit Sub
    End If

    m_ControlName = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "name"))
    If Len(m_ControlName) = 0 Then m_ControlName = "tableTpl"

    m_ItemsSourceRaw = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "itemsSource"))
    If Len(m_ItemsSourceRaw) = 0 Then
        MsgBox "TableTpl: itemsSource is not specified for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    Set listNode = mp_FindFirstChildListNode(controlNode)
    If listNode Is Nothing Then
        MsgBox "TableTpl: primitive table layout must contain root <list>.", vbExclamation
        Exit Sub
    End If

    listNode.setAttribute "itemsSource", m_ItemsSourceRaw
    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render(ByVal wb As Workbook)
    If Not m_IsConfigured Then
        MsgBox "TableTpl: control '" & m_ControlName & "' is not configured.", vbExclamation
        Exit Sub
    End If

    ' No-op by design: primitive nodes are rendered by ex_XmlLayoutEngine.m_RenderTemplateChildren.
End Sub

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case LCase$(Trim$(attrName))
        Case "itemssource"
            obj_IControl_SupportsAttribute = True
    End Select
End Function

Private Function mp_FindFirstChildListNode(ByVal parentNode As Object) As Object
    Dim childNode As Object

    If parentNode Is Nothing Then Exit Function

    For Each childNode In parentNode.ChildNodes
        If childNode.NodeType <> 1 Then GoTo ContinueLoop
        If StrComp(LCase$(CStr(childNode.baseName)), "list", vbBinaryCompare) = 0 Then
            Set mp_FindFirstChildListNode = childNode
            Exit Function
        End If
ContinueLoop:
    Next childNode
End Function
