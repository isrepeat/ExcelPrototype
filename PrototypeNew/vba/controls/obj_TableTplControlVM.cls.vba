VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_TableTplControlVM"
Option Explicit
Implements obj_IControl

Private m_ControlBase As obj_ControlBase
Private m_ControlName As String
Private m_ItemsSourceRaw As String
Private m_IsConfigured As Boolean

' //
' // Interface
' //
Private Sub obj_IControl_Configure(ByVal page As obj_PageBase, ByVal controlNode As Object)
    Dim listNode As Object

    m_IsConfigured = False
    Set m_ControlBase = Nothing

    Set m_ControlBase = New obj_ControlBase
    If Not m_ControlBase.Configure(page, controlNode, "TableTpl", "tableTpl", m_ControlName) Then Exit Sub

    m_ItemsSourceRaw = VBA.Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "itemsSource"))
    If VBA.Len(m_ItemsSourceRaw) = 0 Then
        VBA.MsgBox "TableTpl: itemsSource is not specified for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    Set listNode = private_FindFirstChildListNode(controlNode)
    If listNode Is Nothing Then
        VBA.MsgBox "TableTpl: primitive table layout must contain root <list>.", VBA.vbExclamation
        Exit Sub
    End If

    listNode.setAttribute "itemsSource", m_ItemsSourceRaw
    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render()
    If Not m_IsConfigured Then
        VBA.MsgBox "TableTpl: control '" & m_ControlName & "' is not configured.", VBA.vbExclamation
        Exit Sub
    End If

    ' No-op by design: primitive nodes are rendered by ex_XmlLayoutEngine.m_RenderTemplateChildren.
End Sub

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(attrName))
        Case "itemssource"
            obj_IControl_SupportsAttribute = True
    End Select
End Function

' //
' // API
' //
' (No public API yet.)
'
' //
' // Internal
' //
Private Function private_FindFirstChildListNode(ByVal parentNode As Object) As Object
    Dim childNode As Object

    If parentNode Is Nothing Then Exit Function

    For Each childNode In parentNode.ChildNodes
        If childNode.NodeType <> 1 Then GoTo ContinueLoop
        If VBA.StrComp(VBA.LCase$(VBA.CStr(childNode.baseName)), "list", VBA.vbBinaryCompare) = 0 Then
            Set private_FindFirstChildListNode = childNode
            Exit Function
        End If
ContinueLoop:
    Next childNode
End Function
