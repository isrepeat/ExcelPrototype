Attribute VB_Name = "ex_CacheRuntime"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False

Private g_NamespaceCaches As Object

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:ex_CacheRuntime.fn_Module_Dispose"
#End If
    On Error Resume Next
    Set g_NamespaceCaches = Nothing
    On Error GoTo 0
End Sub

' //
' // API
' //
Public Function fn_Cache_SetValue( _
    ByVal namespaceKey As String, _
    ByVal itemKey As String, _
    ByVal valueIn As Variant _
) As Boolean
    Dim cacheMap As Object
    Dim normalizedItemKey As String

    normalizedItemKey = private_NormalizeKey(itemKey)
    If VBA.Len(normalizedItemKey) = 0 Then Exit Function

    Set cacheMap = private_GetNamespaceCache(namespaceKey, True)
    If cacheMap Is Nothing Then Exit Function

    cacheMap(normalizedItemKey) = valueIn
    fn_Cache_SetValue = True
End Function

Public Function fn_Cache_TryGetValue( _
    ByVal namespaceKey As String, _
    ByVal itemKey As String, _
    ByRef outValue As Variant _
) As Boolean
    Dim cacheMap As Object
    Dim normalizedItemKey As String

    outValue = Empty
    normalizedItemKey = private_NormalizeKey(itemKey)
    If VBA.Len(normalizedItemKey) = 0 Then Exit Function

    Set cacheMap = private_GetNamespaceCache(namespaceKey, False)
    If cacheMap Is Nothing Then Exit Function
    If Not cacheMap.Exists(normalizedItemKey) Then Exit Function

    outValue = cacheMap(normalizedItemKey)
    fn_Cache_TryGetValue = True
End Function

Public Function fn_Cache_SetObject( _
    ByVal namespaceKey As String, _
    ByVal itemKey As String, _
    ByVal valueObject As Object _
) As Boolean
    Dim cacheMap As Object
    Dim normalizedItemKey As String

    normalizedItemKey = private_NormalizeKey(itemKey)
    If VBA.Len(normalizedItemKey) = 0 Then Exit Function
    If valueObject Is Nothing Then Exit Function

    Set cacheMap = private_GetNamespaceCache(namespaceKey, True)
    If cacheMap Is Nothing Then Exit Function

    Set cacheMap(normalizedItemKey) = valueObject
    fn_Cache_SetObject = True
End Function

Public Function fn_Cache_TryGetObject( _
    ByVal namespaceKey As String, _
    ByVal itemKey As String, _
    ByRef outObject As Object _
) As Boolean
    Dim cacheMap As Object
    Dim normalizedItemKey As String
    Dim rawValue As Variant

    Set outObject = Nothing
    normalizedItemKey = private_NormalizeKey(itemKey)
    If VBA.Len(normalizedItemKey) = 0 Then Exit Function

    Set cacheMap = private_GetNamespaceCache(namespaceKey, False)
    If cacheMap Is Nothing Then Exit Function
    If Not cacheMap.Exists(normalizedItemKey) Then Exit Function

    rawValue = cacheMap(normalizedItemKey)
    If Not VBA.IsObject(rawValue) Then Exit Function
    Set outObject = rawValue
    fn_Cache_TryGetObject = Not outObject Is Nothing
End Function

Public Function fn_Cache_Remove(ByVal namespaceKey As String, ByVal itemKey As String) As Boolean
    Dim cacheMap As Object
    Dim normalizedItemKey As String

    normalizedItemKey = private_NormalizeKey(itemKey)
    If VBA.Len(normalizedItemKey) = 0 Then Exit Function

    Set cacheMap = private_GetNamespaceCache(namespaceKey, False)
    If cacheMap Is Nothing Then
        fn_Cache_Remove = True
        Exit Function
    End If

    If cacheMap.Exists(normalizedItemKey) Then cacheMap.Remove normalizedItemKey
    fn_Cache_Remove = True
End Function

Public Function fn_Cache_ClearNamespace(ByVal namespaceKey As String) As Boolean
    Dim normalizedNamespaceKey As String

    normalizedNamespaceKey = private_NormalizeNamespaceKey(namespaceKey)
    private_EnsureStorage
    If g_NamespaceCaches.Exists(normalizedNamespaceKey) Then g_NamespaceCaches.Remove normalizedNamespaceKey
    fn_Cache_ClearNamespace = True
End Function

Public Function fn_Cache_ClearAll() As Boolean
    private_EnsureStorage
    g_NamespaceCaches.RemoveAll
    fn_Cache_ClearAll = True
End Function

' //
' // Internal
' //
Private Sub private_EnsureStorage()
    If g_NamespaceCaches Is Nothing Then
        Set g_NamespaceCaches = ex_Helpers.fn_CreateDictionaryTextCompare()
    End If
End Sub

Private Function private_GetNamespaceCache( _
    ByVal namespaceKey As String, _
    ByVal createIfMissing As Boolean _
) As Object
    Dim normalizedNamespaceKey As String
    Dim cacheMap As Object

    normalizedNamespaceKey = private_NormalizeNamespaceKey(namespaceKey)
    private_EnsureStorage

    If g_NamespaceCaches.Exists(normalizedNamespaceKey) Then
        Set private_GetNamespaceCache = g_NamespaceCaches(normalizedNamespaceKey)
        Exit Function
    End If

    If Not createIfMissing Then Exit Function

    Set cacheMap = ex_Helpers.fn_CreateDictionaryTextCompare()
    Set g_NamespaceCaches(normalizedNamespaceKey) = cacheMap
    Set private_GetNamespaceCache = cacheMap
End Function

Private Function private_NormalizeNamespaceKey(ByVal keyText As String) As String
    private_NormalizeNamespaceKey = private_NormalizeKey(keyText)
    If VBA.Len(private_NormalizeNamespaceKey) = 0 Then private_NormalizeNamespaceKey = "__default__"
End Function

Private Function private_NormalizeKey(ByVal keyText As String) As String
    private_NormalizeKey = VBA.LCase$(VBA.Trim$(VBA.CStr(keyText)))
End Function
