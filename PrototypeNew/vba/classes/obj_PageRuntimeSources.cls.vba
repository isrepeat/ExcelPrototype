VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PageRuntimeSources"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean

Private m_ItemsSourceMap As Object
Private m_ObjectSourceMap As Object

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
    Set m_ItemsSourceMap = Nothing
    Set m_ObjectSourceMap = Nothing
    On Error GoTo 0
End Sub

' Callstack[1]: obj_PageMain.private_PrepareDemoConfigRuntime -> m_Base.RuntimeSources.ResetItemsSources -> obj_PageRuntimeSources.ResetItemsSources
' Callstack[2]: obj_PageMain.private_RegisterDemoTableItems -> m_Base.RuntimeSources.ResetItemsSources -> obj_PageRuntimeSources.ResetItemsSources
' Callstack[3]: obj_PageMain.private_RegisterDemoSingleTableItems -> m_Base.RuntimeSources.ResetItemsSources -> obj_PageRuntimeSources.ResetItemsSources
' Callstack[4]: obj_PageMain.private_RegisterDemoTablePartStylesItems -> m_Base.RuntimeSources.ResetItemsSources -> obj_PageRuntimeSources.ResetItemsSources
' Callstack[5]: ex_Test.private_ResetItemsSources -> pageBase.RuntimeSources.ResetItemsSources -> obj_PageRuntimeSources.ResetItemsSources
Public Sub ResetItemsSources()
    Set m_ItemsSourceMap = Nothing
End Sub

' Callstack[1]: obj_PageMain.private_PrepareDemoConfigRuntime -> m_Base.RuntimeSources.ResetObjectSources -> obj_PageRuntimeSources.ResetObjectSources
' Callstack[2]: obj_PageMain.private_RegisterDemoTableItems -> m_Base.RuntimeSources.ResetObjectSources -> obj_PageRuntimeSources.ResetObjectSources
' Callstack[3]: obj_PageMain.private_RegisterDemoSingleTableItems -> m_Base.RuntimeSources.ResetObjectSources -> obj_PageRuntimeSources.ResetObjectSources
' Callstack[4]: obj_PageMain.private_RegisterDemoTablePartStylesItems -> m_Base.RuntimeSources.ResetObjectSources -> obj_PageRuntimeSources.ResetObjectSources
' Callstack[5]: ex_Test.private_ResetObjectSources -> pageBase.RuntimeSources.ResetObjectSources -> obj_PageRuntimeSources.ResetObjectSources
Public Sub ResetObjectSources()
    Set m_ObjectSourceMap = Nothing
End Sub

' Callstack[1]: obj_PageMain.private_RegisterDemoTableItems -> m_Base.RuntimeSources.SetItemsSource -> obj_PageRuntimeSources.SetItemsSource
' Callstack[2]: obj_PageMain.private_RegisterDemoSingleTableItems -> m_Base.RuntimeSources.SetItemsSource -> obj_PageRuntimeSources.SetItemsSource
' Callstack[3]: obj_PageMain.private_RegisterDemoTablePartStylesItems -> m_Base.RuntimeSources.SetItemsSource -> obj_PageRuntimeSources.SetItemsSource
' Callstack[4]: ex_Test.private_TrySetItemsSource -> pageBase.RuntimeSources.SetItemsSource -> obj_PageRuntimeSources.SetItemsSource
' Callstack[5]: ex_LayoutListRenderer.private_RegisterRuntimeListItemsSourceKey -> renderCtx.PageBase.RuntimeSources.SetItemsSource -> obj_PageRuntimeSources.SetItemsSource
' Callstack[6]: ex_LayoutItemControlRenderer.private_RegisterRuntimeListItemsSourceKey -> renderCtx.PageBase.RuntimeSources.SetItemsSource -> obj_PageRuntimeSources.SetItemsSource
Public Function SetItemsSource(ByVal itemsSourceKey As String, ByVal items As Collection) As Boolean
    Dim normalizedKey As String

    normalizedKey = VBA.LCase$(VBA.Trim$(itemsSourceKey))
    If VBA.Len(normalizedKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: itemsSource key is empty."
#End If
        Exit Function
    End If
    If items Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: itemsSource collection is not specified for key '" & normalizedKey & "'."
#End If
        Exit Function
    End If

    private_EnsureItemsSourceMap
    Set m_ItemsSourceMap(normalizedKey) = items

    SetItemsSource = True
End Function

' Callstack[1]: ex_Test.private_TrySetObjectSource -> pageBase.RuntimeSources.SetObjectSource -> obj_PageRuntimeSources.SetObjectSource
' Callstack[2]: ex_LayoutListRenderer.private_RegisterRuntimeObjectSourceKey -> renderCtx.PageBase.RuntimeSources.SetObjectSource -> obj_PageRuntimeSources.SetObjectSource
' Callstack[3]: ex_LayoutItemControlRenderer.private_RegisterRuntimeObjectSourceKey -> renderCtx.PageBase.RuntimeSources.SetObjectSource -> obj_PageRuntimeSources.SetObjectSource
Public Function SetObjectSource(ByVal objectSourceKey As String, ByVal sourceObject As Object) As Boolean
    Dim normalizedKey As String

    normalizedKey = VBA.LCase$(VBA.Trim$(objectSourceKey))
    If VBA.Len(normalizedKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: objectSource key is empty."
#End If
        Exit Function
    End If
    If sourceObject Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: objectSource object is not specified for key '" & normalizedKey & "'."
#End If
        Exit Function
    End If

    private_EnsureObjectSourceMap
    Set m_ObjectSourceMap(normalizedKey) = sourceObject

    SetObjectSource = True
End Function

' Callstack[1]: ex_Test.private_TryRemoveObjectSource -> pageBase.RuntimeSources.RemoveObjectSource -> obj_PageRuntimeSources.RemoveObjectSource
Public Function RemoveObjectSource(ByVal objectSourceKey As String) As Boolean
    Dim normalizedKey As String

    normalizedKey = VBA.LCase$(VBA.Trim$(objectSourceKey))
    If VBA.Len(normalizedKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: objectSource key is empty."
#End If
        Exit Function
    End If

    private_EnsureObjectSourceMap
    If m_ObjectSourceMap.Exists(normalizedKey) Then
        m_ObjectSourceMap.Remove normalizedKey
    End If

    RemoveObjectSource = True
End Function

Public Function TryGetItemsSourceByKey( _
    ByVal itemsSourceKey As String, _
    ByRef outItems As Collection, _
    Optional ByVal allowMissing As Boolean = False _
) As Boolean
    Dim normalizedKey As String

    Set outItems = Nothing
    normalizedKey = VBA.LCase$(VBA.Trim$(itemsSourceKey))

    If VBA.Len(normalizedKey) = 0 Then
        If allowMissing Then
            TryGetItemsSourceByKey = True
            Exit Function
        End If
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: itemsSource key is empty."
#End If
        Exit Function
    End If

    private_EnsureItemsSourceMap
    If m_ItemsSourceMap.Exists(normalizedKey) Then
        Set outItems = m_ItemsSourceMap(normalizedKey)
        TryGetItemsSourceByKey = True
        Exit Function
    End If

    If allowMissing Then
        TryGetItemsSourceByKey = True
        Exit Function
    End If

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.m_Diagnostic_LogError "PrototypeNew: itemsSource '" & normalizedKey & "' is not registered in page runtime map."
#End If
End Function

Public Function TryGetObjectSourceByKey( _
    ByVal objectSourceKey As String, _
    ByRef outObject As Object, _
    Optional ByVal allowMissing As Boolean = False _
) As Boolean
    Dim normalizedKey As String

    Set outObject = Nothing
    normalizedKey = VBA.LCase$(VBA.Trim$(objectSourceKey))

    If VBA.Len(normalizedKey) = 0 Then
        If allowMissing Then
            TryGetObjectSourceByKey = True
            Exit Function
        End If
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PrototypeNew: objectSource key is empty."
#End If
        Exit Function
    End If

    private_EnsureObjectSourceMap
    If m_ObjectSourceMap.Exists(normalizedKey) Then
        Set outObject = m_ObjectSourceMap(normalizedKey)
        TryGetObjectSourceByKey = True
        Exit Function
    End If

    If allowMissing Then
        TryGetObjectSourceByKey = True
        Exit Function
    End If

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.m_Diagnostic_LogError "PrototypeNew: objectSource '" & normalizedKey & "' is not registered in page runtime map."
#End If
End Function

Public Property Get ItemsSourceMap() As Object
    private_EnsureItemsSourceMap
    Set ItemsSourceMap = m_ItemsSourceMap
End Property

Public Property Get ObjectSourceMap() As Object
    private_EnsureObjectSourceMap
    Set ObjectSourceMap = m_ObjectSourceMap
End Property

' //
' // Internal
' //
Private Sub private_EnsureItemsSourceMap()
    If Not m_ItemsSourceMap Is Nothing Then Exit Sub

    Set m_ItemsSourceMap = VBA.CreateObject("Scripting.Dictionary")
    m_ItemsSourceMap.CompareMode = 1
End Sub

Private Sub private_EnsureObjectSourceMap()
    If Not m_ObjectSourceMap Is Nothing Then Exit Sub

    Set m_ObjectSourceMap = VBA.CreateObject("Scripting.Dictionary")
    m_ObjectSourceMap.CompareMode = 1
End Sub

