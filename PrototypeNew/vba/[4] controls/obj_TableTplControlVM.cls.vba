VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_TableTplControlVM"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean
Implements obj_IControl

Private m_ControlBase As obj_ControlBase
Private m_ControlName As String
Private m_ItemsSourceRaw As String
Private m_IsConfigured As Boolean
Private m_Page As obj_IPage

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Dispose
    On Error GoTo 0
End Sub

' //
' // Interface
' //
Private Sub obj_IControl_Configure(ByVal controlNode As Object)
    Dim listNode As Object
    Dim pageBase As obj_PageBase

    m_IsConfigured = False
    Set m_ControlBase = Nothing

    Set pageBase = m_Page.GetPageBase()
    Set m_ControlBase = New obj_ControlBase
    If Not m_ControlBase.Initialize(m_Page) Then Exit Sub
    If Not m_ControlBase.Configure(pageBase, controlNode, "TableTpl", "tableTpl", m_ControlName) Then Exit Sub

    m_ItemsSourceRaw = VBA.Trim$(ex_XmlCore.fn_NodeAttrText(controlNode, "itemsSource"))
    If VBA.Len(m_ItemsSourceRaw) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableTpl: itemsSource is not specified for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    Set listNode = private_FindFirstChildListNode(controlNode)
    If listNode Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableTpl: primitive table layout must contain root <list>."
#End If
        Exit Sub
    End If

    listNode.setAttribute "itemsSource", m_ItemsSourceRaw
    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render()
    If Not m_IsConfigured Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableTpl: control '" & m_ControlName & "' is not configured."
#End If
        Exit Sub
    End If

    ' No-op by design: primitive nodes are rendered by ex_XmlLayoutEngine.fn_RenderTemplateChildren.
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
Public Function Initialize(ByVal page As obj_IPage) As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    m_IsDisposed = False
    m_IsConfigured = False
    Set m_Page = page
    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Err.Clear
    Set m_ControlBase = Nothing
    Set m_Page = Nothing
    On Error GoTo 0
End Sub

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
