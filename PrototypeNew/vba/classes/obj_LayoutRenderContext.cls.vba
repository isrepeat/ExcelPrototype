VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_LayoutRenderContext"
Option Explicit

Private Const LIST_RUNTIME_KEY_PREFIX As String = "__list_runtime_"
Private Const OBJECT_RUNTIME_KEY_PREFIX As String = "__object_runtime_"

Private m_PageBase As obj_PageBase
Private m_Worksheet As Worksheet
Private m_Workbook As Workbook
Private m_RunToken As String
Private m_ListRuntimeSeed As Long
Private m_ObjectRuntimeSeed As Long
Private m_ObjectRenderSuffixSeed As Long

' //
' // API
' //
' Callstack[1]: obj_PageBase.Render -> renderCtx.Initialize(Me) -> obj_LayoutRenderContext.Initialize
' Callstack[2]: ex_ControlRefreshRuntime.m_TryRefreshStaticControl -> renderCtx.Initialize(pageBase) -> obj_LayoutRenderContext.Initialize
Public Function Initialize(ByVal pageBase As obj_PageBase) As Boolean
    Set m_PageBase = Nothing
    Set m_Worksheet = Nothing
    Set m_Workbook = Nothing
    m_RunToken = VBA.vbNullString
    m_ListRuntimeSeed = 0
    m_ObjectRuntimeSeed = 0
    m_ObjectRenderSuffixSeed = 0

    If pageBase Is Nothing Then
        VBA.MsgBox "PrototypeNew: page base is not specified.", VBA.vbExclamation
        Exit Function
    End If

    Set m_Worksheet = pageBase.Worksheet
    If m_Worksheet Is Nothing Then
        VBA.MsgBox "PrototypeNew: worksheet is not specified.", VBA.vbExclamation
        Exit Function
    End If

    Set m_Workbook = m_Worksheet.Parent
    If m_Workbook Is Nothing Then
        VBA.MsgBox "PrototypeNew: workbook is not specified.", VBA.vbExclamation
        Exit Function
    End If

    Set m_PageBase = pageBase
    m_RunToken = private_CreateRunToken()
    Initialize = True
End Function

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

' Callstack[1]: ex_LayoutItemControlRenderer.m_Render -> renderCtx.NextObjectRenderSuffix -> obj_LayoutRenderContext.NextObjectRenderSuffix
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
