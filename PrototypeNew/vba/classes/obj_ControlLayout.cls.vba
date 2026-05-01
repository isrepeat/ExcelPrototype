VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ControlLayout"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean

Private m_StyleName As String
Private m_LayoutSheetName As String
Private m_RowStart As Long
Private m_ColStart As Long
Private m_RowEnd As Long
Private m_ColEnd As Long

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
End Sub
Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Dispose
    On Error GoTo 0
End Sub

' //
' // API
' //
Public Function Initialize() As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    Initialize = True
End Function
Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    On Error GoTo 0
End Sub

' Callstack[1]: rt_PageManager.m_RenderPage -> page.Render -> obj_PageBase.Render -> ex_XmlLayoutEngine.m_RenderNode -> ex_LayoutControlRenderer.m_Render -> obj_ButtonControlVM.obj_IControl_Configure -> m_Layout.TryReadFromNode -> obj_ControlLayout.TryReadFromNode
' Callstack[2]: rt_PageManager.m_RenderPage -> page.Render -> obj_PageBase.Render -> ex_XmlLayoutEngine.m_RenderNode -> ex_LayoutControlRenderer.m_Render -> obj_BannerControlVM.obj_IControl_Configure -> m_Layout.TryReadFromNode -> obj_ControlLayout.TryReadFromNode
' Callstack[3]: rt_PageManager.m_RenderPage -> page.Render -> obj_PageBase.Render -> ex_XmlLayoutEngine.m_RenderNode -> ex_LayoutControlRenderer.m_Render -> obj_LabelControlVM.obj_IControl_Configure -> m_Layout.TryReadFromNode -> obj_ControlLayout.TryReadFromNode
' Callstack[4]: rt_PageManager.m_RenderPage -> page.Render -> obj_PageBase.Render -> ex_XmlLayoutEngine.m_RenderNode -> ex_LayoutControlRenderer.m_Render -> control.Configure(obj_IControl) -> m_Layout.TryReadFromNode -> obj_ControlLayout.TryReadFromNode
' Callstack[5]: rt_PageManager.m_RenderPage -> page.Render -> obj_PageBase.Render -> ex_XmlLayoutEngine.m_RenderNode -> ex_LayoutControlRenderer.m_Render -> obj_ConfigControlVM.obj_IControl_Configure -> m_Layout.TryReadFromNode -> obj_ControlLayout.TryReadFromNode
' Callstack[6]: rt_PageManager.m_RenderPage -> page.Render -> obj_PageBase.Render -> ex_XmlLayoutEngine.m_RenderNode -> ex_LayoutControlRenderer.m_Render -> obj_SelectControlVM.obj_IControl_Configure -> m_Layout.TryReadFromNode -> obj_ControlLayout.TryReadFromNode
Public Function TryReadFromNode( _
    ByVal controlNode As Object, _
    ByVal controlTypeLabel As String, _
    ByVal controlName As String, _
    Optional ByVal styleAttrName As String = "style" _
) As Boolean
    If controlNode Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError controlTypeLabel & ": control node is not specified."
#End If
        Exit Function
    End If

    m_StyleName = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, styleAttrName)))

    m_LayoutSheetName = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "__layoutSheetName")))
    If VBA.Len(m_LayoutSheetName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError controlTypeLabel & ": runtime layout sheet is missing for control '" & controlName & "'."
#End If
        Exit Function
    End If

    If Not private_TryReadLayoutLongAttr(controlNode, "__layoutRowStart", controlTypeLabel, controlName, m_RowStart) Then Exit Function
    If Not private_TryReadLayoutLongAttr(controlNode, "__layoutColStart", controlTypeLabel, controlName, m_ColStart) Then Exit Function
    If Not private_TryReadLayoutLongAttr(controlNode, "__layoutRowEnd", controlTypeLabel, controlName, m_RowEnd) Then Exit Function
    If Not private_TryReadLayoutLongAttr(controlNode, "__layoutColEnd", controlTypeLabel, controlName, m_ColEnd) Then Exit Function

    If Not private_TryValidateLayoutBounds(controlTypeLabel, controlName) Then Exit Function

    TryReadFromNode = True
End Function

' Callstack[1]: rt_Snapshots.m_RestorePageSnapshots -> serializablePage.TryDeserializeSnapshot(obj_PageMain) -> obj_PageMain.obj_IPage_Render -> obj_PageMain.private_TryRestorePendingControlSnapshots -> obj_PageBase.TryRestoreSerializableControlSnapshots -> serializableControl.TryDeserializeSnapshot(obj_ButtonControlVM) -> m_Layout.TryReadFromRuntimeValues -> obj_ControlLayout.TryReadFromRuntimeValues
' Callstack[2]: rt_Snapshots.m_RestorePageSnapshots -> serializablePage.TryDeserializeSnapshot(obj_PageMain) -> obj_PageMain.obj_IPage_Render -> obj_PageMain.private_TryRestorePendingControlSnapshots -> obj_PageBase.TryRestoreSerializableControlSnapshots -> serializableControl.TryDeserializeSnapshot(obj_SelectControlVM) -> m_Layout.TryReadFromRuntimeValues -> obj_ControlLayout.TryReadFromRuntimeValues
Public Function TryReadFromRuntimeValues( _
    ByVal controlTypeLabel As String, _
    ByVal controlName As String, _
    ByVal layoutSheetName As String, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long, _
    Optional ByVal styleName As String = VBA.vbNullString _
) As Boolean
    m_StyleName = VBA.Trim$(VBA.CStr(styleName))
    m_LayoutSheetName = VBA.Trim$(VBA.CStr(layoutSheetName))
    m_RowStart = rowStart
    m_ColStart = colStart
    m_RowEnd = rowEnd
    m_ColEnd = colEnd

    If VBA.Len(m_LayoutSheetName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError controlTypeLabel & ": runtime layout sheet is missing for control '" & controlName & "'."
#End If
        Exit Function
    End If

    If Not private_TryValidateLayoutBounds(controlTypeLabel, controlName) Then Exit Function
    TryReadFromRuntimeValues = True
End Function

Public Property Get StyleName() As String
    StyleName = m_StyleName
End Property

Public Property Get LayoutSheetName() As String
    LayoutSheetName = m_LayoutSheetName
End Property

Public Property Get RowStart() As Long
    RowStart = m_RowStart
End Property

Public Property Get ColStart() As Long
    ColStart = m_ColStart
End Property

Public Property Get RowEnd() As Long
    RowEnd = m_RowEnd
End Property

Public Property Get ColEnd() As Long
    ColEnd = m_ColEnd
End Property

' //
' // Internal
' //
Private Function private_TryValidateLayoutBounds(ByVal controlTypeLabel As String, ByVal controlName As String) As Boolean
    If m_RowStart <= 0 Or m_ColStart <= 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError controlTypeLabel & ": invalid row/column start for control '" & controlName & "'."
#End If
        Exit Function
    End If
    If m_RowEnd < m_RowStart Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError controlTypeLabel & ": invalid spanRows range for control '" & controlName & "'."
#End If
        Exit Function
    End If
    If m_ColEnd < m_ColStart Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError controlTypeLabel & ": invalid spanColls range for control '" & controlName & "'."
#End If
        Exit Function
    End If

    private_TryValidateLayoutBounds = True
End Function

Private Function private_TryReadLayoutLongAttr( _
    ByVal controlNode As Object, _
    ByVal attrName As String, _
    ByVal controlTypeLabel As String, _
    ByVal controlName As String, _
    ByRef outValue As Long _
) As Boolean
    Dim rawText As String

    rawText = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, attrName)))
    If VBA.Len(rawText) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError controlTypeLabel & ": runtime layout attribute '" & attrName & "' is missing for control '" & controlName & "'."
#End If
        Exit Function
    End If
    If Not VBA.IsNumeric(rawText) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError controlTypeLabel & ": runtime layout attribute '" & attrName & "' must be numeric for control '" & controlName & "'."
#End If
        Exit Function
    End If

    outValue = VBA.CLng(rawText)
    private_TryReadLayoutLongAttr = True
End Function

