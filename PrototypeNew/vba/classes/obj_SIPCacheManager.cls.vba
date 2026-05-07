VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_SIPCacheManager"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean

Private Const CACHE_ENTRY_STAMP_KEY As String = "Stamp"
Private Const CACHE_ENTRY_ITEMS_KEY As String = "Items"

Private m_ProvidersByKey As Object
Private m_CacheByProviderKey As Object

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
Public Function Initialize() As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    private_EnsureProvidersMap
    private_EnsureCacheMap
    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Set m_ProvidersByKey = Nothing
    Set m_CacheByProviderKey = Nothing
    On Error GoTo 0
End Sub

Public Function RegisterProvider( _
    ByVal provider As obj_ISelectItemsSourceProvider, _
    Optional ByVal replaceExisting As Boolean = True _
) As Boolean
    Dim normalizedKey As String

    If provider Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "SelectItemsSourceCacheManager: provider is not specified."
#End If
        Exit Function
    End If

    normalizedKey = private_NormalizeKey(provider.GetProviderKey())
    If VBA.Len(normalizedKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "SelectItemsSourceCacheManager: provider key is empty."
#End If
        Exit Function
    End If

    private_EnsureProvidersMap
    private_EnsureCacheMap

    If m_ProvidersByKey.Exists(normalizedKey) Then
        If Not replaceExisting Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "SelectItemsSourceCacheManager: provider '" & normalizedKey & "' is already registered."
#End If
            Exit Function
        End If
    End If

    Set m_ProvidersByKey(normalizedKey) = provider
    If m_CacheByProviderKey.Exists(normalizedKey) Then m_CacheByProviderKey.Remove normalizedKey
    RegisterProvider = True
End Function

Public Function UnregisterProvider(ByVal providerKey As String) As Boolean
    Dim normalizedKey As String

    normalizedKey = private_NormalizeKey(providerKey)
    If VBA.Len(normalizedKey) = 0 Then Exit Function

    private_EnsureProvidersMap
    private_EnsureCacheMap

    If m_ProvidersByKey.Exists(normalizedKey) Then m_ProvidersByKey.Remove normalizedKey
    If m_CacheByProviderKey.Exists(normalizedKey) Then m_CacheByProviderKey.Remove normalizedKey
    UnregisterProvider = True
End Function

Public Function HasProvider(ByVal providerKey As String) As Boolean
    Dim normalizedKey As String

    normalizedKey = private_NormalizeKey(providerKey)
    If VBA.Len(normalizedKey) = 0 Then Exit Function

    private_EnsureProvidersMap
    HasProvider = m_ProvidersByKey.Exists(normalizedKey)
End Function

Public Function ResetCacheByProviderKey(ByVal providerKey As String) As Boolean
    Dim normalizedKey As String

    normalizedKey = private_NormalizeKey(providerKey)
    If VBA.Len(normalizedKey) = 0 Then Exit Function

    private_EnsureCacheMap
    If m_CacheByProviderKey.Exists(normalizedKey) Then m_CacheByProviderKey.Remove normalizedKey
    ResetCacheByProviderKey = True
End Function

Public Sub ResetAllCaches()
    Set m_CacheByProviderKey = Nothing
End Sub

Public Function TryResolveItemsByProviderKey( _
    ByVal providerKey As String, _
    ByRef outItems As Collection, _
    Optional ByRef outUsedCache As Boolean = False, _
    Optional ByVal allowMissing As Boolean = False _
) As Boolean
    Dim normalizedKey As String
    Dim provider As obj_ISelectItemsSourceProvider
    Dim currentStamp As String
    Dim cacheEntry As Object
    Dim cachedStamp As String
    Dim builtItems As Collection

    Set outItems = Nothing
    outUsedCache = False

    normalizedKey = private_NormalizeKey(providerKey)
    If VBA.Len(normalizedKey) = 0 Then
        If allowMissing Then
            TryResolveItemsByProviderKey = True
            Exit Function
        End If
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "SelectItemsSourceCacheManager: provider key is empty."
#End If
        Exit Function
    End If

    private_EnsureProvidersMap
    private_EnsureCacheMap

    If Not m_ProvidersByKey.Exists(normalizedKey) Then
        If allowMissing Then
            TryResolveItemsByProviderKey = True
            Exit Function
        End If
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "SelectItemsSourceCacheManager: provider '" & normalizedKey & "' is not registered."
#End If
        Exit Function
    End If
    Set provider = m_ProvidersByKey(normalizedKey)
    If provider Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "SelectItemsSourceCacheManager: provider '" & normalizedKey & "' is not initialized."
#End If
        Exit Function
    End If

    ' Stamp — "версия данных" от provider.
    ' Если stamp не поменялся, возвращаем прежнюю collection из cache.
    If Not provider.TryGetCurrentStamp(currentStamp) Then Exit Function
    currentStamp = VBA.Trim$(currentStamp)
    If VBA.Len(currentStamp) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "SelectItemsSourceCacheManager: provider '" & normalizedKey & "' returned empty stamp."
#End If
        Exit Function
    End If

    ' Быстрый путь: cache-hit (stamp совпал) -> rebuild не нужен.
    If m_CacheByProviderKey.Exists(normalizedKey) Then
        Set cacheEntry = m_CacheByProviderKey(normalizedKey)
        If Not cacheEntry Is Nothing Then
            cachedStamp = VBA.CStr(cacheEntry(CACHE_ENTRY_STAMP_KEY))
            If VBA.StrComp(cachedStamp, currentStamp, VBA.vbBinaryCompare) = 0 Then
                If VBA.IsObject(cacheEntry(CACHE_ENTRY_ITEMS_KEY)) Then
                    If VBA.StrComp(VBA.TypeName(cacheEntry(CACHE_ENTRY_ITEMS_KEY)), "Collection", VBA.vbTextCompare) = 0 Then
                        Set outItems = cacheEntry(CACHE_ENTRY_ITEMS_KEY)
                        If Not outItems Is Nothing Then
                            outUsedCache = True
                            TryResolveItemsByProviderKey = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    End If

    ' Медленный путь: cache-miss -> provider пересобирает items.
    If Not provider.TryBuildItems(builtItems) Then Exit Function
    If builtItems Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "SelectItemsSourceCacheManager: provider '" & normalizedKey & "' returned empty items collection."
#End If
        Exit Function
    End If

    ' Обновляем cache новой парой (stamp + items).
    Set cacheEntry = VBA.CreateObject("Scripting.Dictionary")
    cacheEntry.CompareMode = 1
    cacheEntry(CACHE_ENTRY_STAMP_KEY) = currentStamp
    Set cacheEntry(CACHE_ENTRY_ITEMS_KEY) = builtItems
    Set m_CacheByProviderKey(normalizedKey) = cacheEntry

    Set outItems = builtItems
    TryResolveItemsByProviderKey = True
End Function

' //
' // Internal
' //
Private Sub private_EnsureProvidersMap()
    If Not m_ProvidersByKey Is Nothing Then Exit Sub
    Set m_ProvidersByKey = VBA.CreateObject("Scripting.Dictionary")
    m_ProvidersByKey.CompareMode = 1
End Sub

Private Sub private_EnsureCacheMap()
    If Not m_CacheByProviderKey Is Nothing Then Exit Sub
    Set m_CacheByProviderKey = VBA.CreateObject("Scripting.Dictionary")
    m_CacheByProviderKey.CompareMode = 1
End Sub

Private Function private_NormalizeKey(ByVal keyText As String) As String
    private_NormalizeKey = VBA.LCase$(VBA.Trim$(keyText))
End Function
