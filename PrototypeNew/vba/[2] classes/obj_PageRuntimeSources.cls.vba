VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PageRuntimeSources"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

' Префикс для временных ключей itemsSource, которые создают layout-рендереры
' когда {Binding ...} резолвится в Collection внутри list/item шаблонов.
Private Const LIST_RUNTIME_KEY_PREFIX As String = "__list_runtime_"
' Префикс для временных ключей objectSource, которые создают layout-рендереры
' когда {Binding ...} резолвится в Object внутри list/item шаблонов.
Private Const OBJECT_RUNTIME_KEY_PREFIX As String = "__object_runtime_"

Private m_Page As obj_IPage
Private m_ItemsSourceMap As Object
Private m_ObjectSourceMap As Object
Private m_IsDisposed As Boolean

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
Public Function Initialize(ByVal page As obj_IPage) As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    Set m_Page = page
    Set m_ItemsSourceMap = Nothing
    Set m_ObjectSourceMap = Nothing
    m_IsDisposed = False
    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Set m_Page = Nothing
    Set m_ItemsSourceMap = Nothing
    Set m_ObjectSourceMap = Nothing
    On Error GoTo 0
End Sub

' Устаревший API: полный сброс карты items.
' Предпочтительна точечная очистка через RemoveItemsSource/RemoveTemporaryItemsSources.
Public Sub ResetItemsSources( _
    Optional ByVal notifyChange As Boolean = False _
)
    Dim hadItems As Boolean

    hadItems = Not m_ItemsSourceMap Is Nothing
    Set m_ItemsSourceMap = Nothing
    If notifyChange And hadItems Then
        Call private_NotifyRuntimeChanged("itemsSource:reset")
    End If
End Sub

' Устаревший API: полный сброс карты objects.
' Предпочтительна точечная очистка через RemoveObjectSource/RemoveTemporaryObjectsSources.
Public Sub ResetObjectSources( _
    Optional ByVal notifyChange As Boolean = False _
)
    Dim hadObjects As Boolean

    hadObjects = Not m_ObjectSourceMap Is Nothing
    Set m_ObjectSourceMap = Nothing
    If notifyChange And hadObjects Then
        Call private_NotifyRuntimeChanged("objectSource:reset")
    End If
End Sub

' Callstack[1]: obj_PageMain.private_RegisterDemoTableItems -> m_Base.RuntimeSources.SetItemsSource -> obj_PageRuntimeSources.SetItemsSource
' Callstack[2]: obj_PageMain.private_RegisterDemoSingleTableItems -> m_Base.RuntimeSources.SetItemsSource -> obj_PageRuntimeSources.SetItemsSource
' Callstack[3]: obj_PageMain.private_RegisterDemoTablePartStylesItems -> m_Base.RuntimeSources.SetItemsSource -> obj_PageRuntimeSources.SetItemsSource
' Callstack[4]: ex_Test.private_TrySetItemsSource -> pageBase.RuntimeSources.SetItemsSource -> obj_PageRuntimeSources.SetItemsSource
' Callstack[5]: ex_LayoutListRenderer.private_RegisterRuntimeListItemsSourceKey -> renderCtx.PageBase.RuntimeSources.SetItemsSource -> obj_PageRuntimeSources.SetItemsSource
' Callstack[6]: ex_LayoutItemControlRenderer.private_RegisterRuntimeListItemsSourceKey -> renderCtx.PageBase.RuntimeSources.SetItemsSource -> obj_PageRuntimeSources.SetItemsSource
Public Function SetItemsSource( _
    ByVal itemsSourceKey As String, _
    ByVal items As Collection, _
    Optional ByVal notifyChange As Boolean = False _
) As Boolean
    Dim normalizedKey As String

    normalizedKey = VBA.LCase$(VBA.Trim$(itemsSourceKey))
    If VBA.Len(normalizedKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: itemsSource key is empty."
#End If
        Exit Function
    End If
    If items Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: itemsSource collection is not specified for key '" & normalizedKey & "'."
#End If
        Exit Function
    End If

    private_EnsureItemsSourceMap
    Set m_ItemsSourceMap(normalizedKey) = items

    If notifyChange Then
        If Not private_NotifyRuntimeChanged("itemsSource:" & normalizedKey) Then Exit Function
    End If

    SetItemsSource = True
End Function

Public Function RemoveItemsSource( _
    ByVal itemsSourceKey As String, _
    Optional ByVal notifyChange As Boolean = False _
) As Boolean
    Dim normalizedKey As String
    Dim removed As Boolean

    normalizedKey = VBA.LCase$(VBA.Trim$(itemsSourceKey))
    If VBA.Len(normalizedKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: itemsSource key is empty."
#End If
        Exit Function
    End If

    private_EnsureItemsSourceMap
    If m_ItemsSourceMap.Exists(normalizedKey) Then
        m_ItemsSourceMap.Remove normalizedKey
        removed = True
    End If

    If notifyChange And removed Then
        If Not private_NotifyRuntimeChanged("itemsSource:remove:" & normalizedKey) Then Exit Function
    End If

    RemoveItemsSource = True
End Function

Public Function RemoveItemsSourcesByPrefix( _
    ByVal itemsSourceKeyPrefix As String, _
    Optional ByVal notifyChange As Boolean = False _
) As Boolean
    Dim normalizedPrefix As String
    Dim keyObj As Variant
    Dim keysToRemove As Collection
    Dim removedCount As Long

    normalizedPrefix = VBA.LCase$(VBA.Trim$(itemsSourceKeyPrefix))
    If VBA.Len(normalizedPrefix) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: itemsSource prefix is empty."
#End If
        Exit Function
    End If

    private_EnsureItemsSourceMap
    Set keysToRemove = New Collection

    For Each keyObj In m_ItemsSourceMap.Keys
        If ex_Helpers.fn_TextStartsWith(VBA.CStr(keyObj), normalizedPrefix) Then
            keysToRemove.Add VBA.CStr(keyObj)
        End If
    Next keyObj

    For Each keyObj In keysToRemove
        m_ItemsSourceMap.Remove VBA.CStr(keyObj)
        removedCount = removedCount + 1
    Next keyObj

    If notifyChange And removedCount > 0 Then
        If Not private_NotifyRuntimeChanged("itemsSource:prefix:" & normalizedPrefix) Then Exit Function
    End If

    RemoveItemsSourcesByPrefix = True
End Function

' Удаляет все временные list itemsSource, зарегистрированные layout-рендерерами
' под ключами с префиксом "__list_runtime_...".
Public Function RemoveTemporaryItemsSources( _
    Optional ByVal notifyChange As Boolean = False _
) As Boolean
    RemoveTemporaryItemsSources = RemoveItemsSourcesByPrefix(LIST_RUNTIME_KEY_PREFIX, notifyChange)
End Function

' Callstack[1]: ex_Test.private_TrySetObjectSource -> pageBase.RuntimeSources.SetObjectSource -> obj_PageRuntimeSources.SetObjectSource
' Callstack[2]: ex_LayoutListRenderer.private_RegisterRuntimeObjectSourceKey -> renderCtx.PageBase.RuntimeSources.SetObjectSource -> obj_PageRuntimeSources.SetObjectSource
' Callstack[3]: ex_LayoutItemControlRenderer.private_RegisterRuntimeObjectSourceKey -> renderCtx.PageBase.RuntimeSources.SetObjectSource -> obj_PageRuntimeSources.SetObjectSource
Public Function SetObjectSource( _
    ByVal objectSourceKey As String, _
    ByVal sourceObject As Object, _
    Optional ByVal notifyChange As Boolean = False _
) As Boolean
    Dim normalizedKey As String

    normalizedKey = VBA.LCase$(VBA.Trim$(objectSourceKey))
    If VBA.Len(normalizedKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: objectSource key is empty."
#End If
        Exit Function
    End If
    If sourceObject Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: objectSource object is not specified for key '" & normalizedKey & "'."
#End If
        Exit Function
    End If

    private_EnsureObjectSourceMap
    Set m_ObjectSourceMap(normalizedKey) = sourceObject

    If notifyChange Then
        If Not private_NotifyRuntimeChanged("objectSource:" & normalizedKey) Then Exit Function
    End If

    SetObjectSource = True
End Function

' Callstack[1]: ex_Test.private_TryRemoveObjectSource -> pageBase.RuntimeSources.RemoveObjectSource -> obj_PageRuntimeSources.RemoveObjectSource
Public Function RemoveObjectSource( _
    ByVal objectSourceKey As String, _
    Optional ByVal notifyChange As Boolean = False _
) As Boolean
    Dim normalizedKey As String
    Dim removed As Boolean

    normalizedKey = VBA.LCase$(VBA.Trim$(objectSourceKey))
    If VBA.Len(normalizedKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: objectSource key is empty."
#End If
        Exit Function
    End If

    private_EnsureObjectSourceMap
    If m_ObjectSourceMap.Exists(normalizedKey) Then
        m_ObjectSourceMap.Remove normalizedKey
        removed = True
    End If

    If notifyChange And removed Then
        If Not private_NotifyRuntimeChanged("objectSource:remove:" & normalizedKey) Then Exit Function
    End If

    RemoveObjectSource = True
End Function

Public Function RemoveObjectSourcesByPrefix( _
    ByVal objectSourceKeyPrefix As String, _
    Optional ByVal notifyChange As Boolean = False _
) As Boolean
    Dim normalizedPrefix As String
    Dim keyObj As Variant
    Dim keysToRemove As Collection
    Dim removedCount As Long

    normalizedPrefix = VBA.LCase$(VBA.Trim$(objectSourceKeyPrefix))
    If VBA.Len(normalizedPrefix) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: objectSource prefix is empty."
#End If
        Exit Function
    End If

    private_EnsureObjectSourceMap
    Set keysToRemove = New Collection

    For Each keyObj In m_ObjectSourceMap.Keys
        If ex_Helpers.fn_TextStartsWith(VBA.CStr(keyObj), normalizedPrefix) Then
            keysToRemove.Add VBA.CStr(keyObj)
        End If
    Next keyObj

    For Each keyObj In keysToRemove
        m_ObjectSourceMap.Remove VBA.CStr(keyObj)
        removedCount = removedCount + 1
    Next keyObj

    If notifyChange And removedCount > 0 Then
        If Not private_NotifyRuntimeChanged("objectSource:prefix:" & normalizedPrefix) Then Exit Function
    End If

    RemoveObjectSourcesByPrefix = True
End Function

' Удаляет все временные objectSource, зарегистрированные layout-рендерерами
' под ключами с префиксом "__object_runtime_...".
Public Function RemoveTemporaryObjectsSources( _
    Optional ByVal notifyChange As Boolean = False _
) As Boolean
    RemoveTemporaryObjectsSources = RemoveObjectSourcesByPrefix(OBJECT_RUNTIME_KEY_PREFIX, notifyChange)
End Function

' Совместимый алиас с сохраненной опечаткой.
Public Function RemoveTeporaryObjectsSources( _
    Optional ByVal notifyChange As Boolean = False _
) As Boolean
    RemoveTeporaryObjectsSources = RemoveTemporaryObjectsSources(notifyChange)
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
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: itemsSource key is empty."
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
    ex_Core.fn_Diagnostic_LogError "PrototypeNew: itemsSource '" & normalizedKey & "' is not registered in page runtime map."
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
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: objectSource key is empty."
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
    ex_Core.fn_Diagnostic_LogError "PrototypeNew: objectSource '" & normalizedKey & "' is not registered in page runtime map."
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

Private Function private_NotifyRuntimeChanged(ByVal reason As String) As Boolean
    reason = VBA.Trim$(reason)
    If VBA.Len(reason) = 0 Then reason = "runtimeSources:changed"

    private_NotifyRuntimeChanged = rt_PageManager.fn_RenderPage(m_Page, reason)
End Function
