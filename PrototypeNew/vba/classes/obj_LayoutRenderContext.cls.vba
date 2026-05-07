VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_LayoutRenderContext"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean

Private Const LIST_RUNTIME_KEY_PREFIX As String = "__list_runtime_"
Private Const OBJECT_RUNTIME_KEY_PREFIX As String = "__object_runtime_"

Private m_PageBase As obj_PageBase
Private m_Worksheet As Worksheet
Private m_Workbook As Workbook
Private m_RunToken As String
Private m_ListRuntimeSeed As Long
Private m_ObjectRuntimeSeed As Long
Private m_ObjectRenderSuffixSeed As Long

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
' // API
' //
Public Function Initialize(ByVal pageBase As obj_PageBase) As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    Set m_PageBase = Nothing
    Set m_Worksheet = Nothing
    Set m_Workbook = Nothing
    m_RunToken = VBA.vbNullString
    m_ListRuntimeSeed = 0
    m_ObjectRuntimeSeed = 0
    m_ObjectRenderSuffixSeed = 0

    If pageBase Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: page base is not specified."
#End If
        Exit Function
    End If

    Set m_Worksheet = pageBase.Worksheet
    If m_Worksheet Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: worksheet is not specified."
#End If
        Exit Function
    End If

    Set m_Workbook = m_Worksheet.Parent
    If m_Workbook Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: workbook is not specified."
#End If
        Exit Function
    End If

    Set m_PageBase = pageBase
    m_RunToken = private_CreateRunToken()
    Initialize = True
End Function
Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Set m_PageBase = Nothing
    Set m_Worksheet = Nothing
    Set m_Workbook = Nothing
    On Error GoTo 0
End Sub

' Callstack[1]: obj_PageBase.Render -> renderCtx.Initialize(Me) -> obj_LayoutRenderContext.Initialize
' Callstack[2]: ex_ControlRefreshRuntime.fn_TryRefreshStaticControl -> renderCtx.Initialize(pageBase) -> obj_LayoutRenderContext.Initialize

Public Property Get PageBase() As obj_PageBase
    Set PageBase = m_PageBase
End Property

Public Property Get Worksheet() As Worksheet
    Set Worksheet = m_Worksheet
End Property

Public Property Get Workbook() As Workbook
    Set Workbook = m_Workbook
End Property

Public Property Get RunToken() As String
    RunToken = m_RunToken
End Property

' Callstack[1]: ex_LayoutListRenderer.private_RegisterRuntimeListItemsSourceKey -> renderCtx.NextListRuntimeSourceKey -> obj_LayoutRenderContext.NextListRuntimeSourceKey
' Callstack[2]: ex_LayoutItemControlRenderer.private_RegisterRuntimeListItemsSourceKey -> renderCtx.NextListRuntimeSourceKey -> obj_LayoutRenderContext.NextListRuntimeSourceKey
Public Function NextListRuntimeSourceKey() As String
    m_ListRuntimeSeed = m_ListRuntimeSeed + 1
    NextListRuntimeSourceKey = LIST_RUNTIME_KEY_PREFIX & m_RunToken & "_" & VBA.CStr(m_ListRuntimeSeed)
End Function

' Callstack[1]: ex_LayoutListRenderer.private_RegisterRuntimeObjectSourceKey -> renderCtx.NextObjectRuntimeSourceKey -> obj_LayoutRenderContext.NextObjectRuntimeSourceKey
' Callstack[2]: ex_LayoutItemControlRenderer.private_RegisterRuntimeObjectSourceKey -> renderCtx.NextObjectRuntimeSourceKey -> obj_LayoutRenderContext.NextObjectRuntimeSourceKey
Public Function NextObjectRuntimeSourceKey() As String
    m_ObjectRuntimeSeed = m_ObjectRuntimeSeed + 1
    NextObjectRuntimeSourceKey = OBJECT_RUNTIME_KEY_PREFIX & m_RunToken & "_" & VBA.CStr(m_ObjectRuntimeSeed)
End Function

' Callstack[1]: ex_LayoutItemControlRenderer.fn_Render -> renderCtx.NextObjectRenderSuffix -> obj_LayoutRenderContext.NextObjectRenderSuffix
Public Function NextObjectRenderSuffix() As Long
    m_ObjectRenderSuffixSeed = m_ObjectRenderSuffixSeed + 1
    NextObjectRenderSuffix = m_ObjectRenderSuffixSeed
End Function

' //
' // Internal
' //
Private Function private_CreateRunToken() As String
    Static s_RunSerial As Long

    s_RunSerial = s_RunSerial + 1
    private_CreateRunToken = VBA.Format$(VBA.Now, "yyyymmddhhnnss") & "_" & VBA.CStr(VBA.CLng(VBA.Timer * 1000)) & "_" & VBA.CStr(s_RunSerial)
End Function

