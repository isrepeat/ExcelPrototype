VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ControlBase"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean

Private m_PageBase As obj_PageBase
Private m_DataContext As Object

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
    Set m_PageBase = Nothing
    Set m_DataContext = Nothing
    On Error GoTo 0
End Sub

' Callstack[1]: VBA.ImmediateWindow -> obj_ControlBase.Reset
Public Sub Reset()
    Set m_PageBase = Nothing
    Set m_DataContext = Nothing
End Sub

' Callstack[1]: rt_PageManager.m_RenderPage -> page.Render -> obj_PageBase.Render -> ex_XmlLayoutEngine.m_RenderNode -> ex_LayoutControlRenderer.m_Render -> control.Configure(obj_IControl) -> obj_ButtonControlVM.obj_IControl_Configure -> m_Base.Configure -> obj_ControlBase.Configure
' Callstack[2]: rt_PageManager.m_RenderPage -> page.Render -> obj_PageBase.Render -> ex_XmlLayoutEngine.m_RenderNode -> ex_LayoutControlRenderer.m_Render -> control.Configure(obj_IControl) -> obj_BannerControlVM.obj_IControl_Configure -> m_Base.Configure -> obj_ControlBase.Configure
' Callstack[3]: rt_PageManager.m_RenderPage -> page.Render -> obj_PageBase.Render -> ex_XmlLayoutEngine.m_RenderNode -> ex_LayoutControlRenderer.m_Render -> control.Configure(obj_IControl) -> obj_LabelControlVM.obj_IControl_Configure -> m_Base.Configure -> obj_ControlBase.Configure
' Callstack[4]: rt_PageManager.m_RenderPage -> page.Render -> obj_PageBase.Render -> ex_XmlLayoutEngine.m_RenderNode -> ex_LayoutControlRenderer.m_Render -> control.Configure(obj_IControl) -> m_Base.Configure -> obj_ControlBase.Configure
' Callstack[5]: rt_PageManager.m_RenderPage -> page.Render -> obj_PageBase.Render -> ex_XmlLayoutEngine.m_RenderNode -> ex_LayoutControlRenderer.m_Render -> control.Configure(obj_IControl) -> obj_ConfigControlVM.obj_IControl_Configure -> m_Base.Configure -> obj_ControlBase.Configure
' Callstack[6]: rt_PageManager.m_RenderPage -> page.Render -> obj_PageBase.Render -> ex_XmlLayoutEngine.m_RenderNode -> ex_LayoutControlRenderer.m_Render -> control.Configure(obj_IControl) -> obj_SelectControlVM.obj_IControl_Configure -> m_Base.Configure -> obj_ControlBase.Configure
' Callstack[7]: rt_PageManager.m_RenderPage -> page.Render -> obj_PageBase.Render -> ex_XmlLayoutEngine.m_RenderNode -> ex_LayoutControlRenderer.m_Render -> control.Configure(obj_IControl) -> obj_TableListControlVM.obj_IControl_Configure -> m_Base.Configure -> obj_ControlBase.Configure
' Callstack[8]: rt_PageManager.m_RenderPage -> page.Render -> obj_PageBase.Render -> ex_XmlLayoutEngine.m_RenderNode -> ex_LayoutControlRenderer.m_Render -> control.Configure(obj_IControl) -> obj_TableSingleControlVM.obj_IControl_Configure -> m_Base.Configure -> obj_ControlBase.Configure
' Callstack[9]: rt_PageManager.m_RenderPage -> page.Render -> obj_PageBase.Render -> ex_XmlLayoutEngine.m_RenderNode -> ex_LayoutControlRenderer.m_Render -> control.Configure(obj_IControl) -> obj_TableTplControlVM.obj_IControl_Configure -> m_Base.Configure -> obj_ControlBase.Configure
Public Function Configure( _
    ByVal page As obj_PageBase, _
    ByVal controlNode As Object, _
    ByVal controlTypeName As String, _
    ByVal defaultControlName As String, _
    ByRef outControlName As String _
) As Boolean
    Dim resolvedName As String
    Dim dataContextRaw As String
    Dim dataContext As Object

    outControlName = VBA.Trim$(defaultControlName)
    Set m_DataContext = Nothing

    If page Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError controlTypeName & ": page is not specified for control configure."
#End If
        Exit Function
    End If
    Set m_PageBase = page

    If controlNode Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError controlTypeName & ": control node is not specified."
#End If
        Exit Function
    End If

    resolvedName = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "name")))
    If VBA.Len(resolvedName) = 0 Then
        resolvedName = VBA.Trim$(defaultControlName)
    End If
    outControlName = resolvedName

    dataContextRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "dataContext")))
    If VBA.Len(dataContextRaw) > 0 Then
        If page.RuntimeSources Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.m_Diagnostic_LogError controlTypeName & ": runtime sources are not specified for control '" & outControlName & "'."
#End If
            Exit Function
        End If
        If Not ex_RuntimeSourceResolver.m_TryResolveObjectSource(page.RuntimeSources, dataContextRaw, dataContext, False) Then Exit Function
        If dataContext Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.m_Diagnostic_LogError controlTypeName & ": dataContext resolved to empty object for control '" & outControlName & "'."
#End If
            Exit Function
        End If
        Set m_DataContext = dataContext
    End If

    Configure = True
End Function

Public Property Get PageBase() As obj_PageBase
    Set PageBase = m_PageBase
End Property

Public Property Get DataContext() As Object
    Set DataContext = m_DataContext
End Property

