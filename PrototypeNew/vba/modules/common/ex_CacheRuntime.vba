Attribute VB_Name = "ex_CacheRuntime"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False

' Централизованное in-memory хранилище кеша (словарь словарей).
' Псевдо-тип в терминах шаблонов C++:
'   Dictionary<namespaceKey, Dictionary<itemKey, valueVariant>>
' где:
'   namespaceKey = логическая область кеша (например, модуль/сценарий),
'   itemKey      = ключ внутри области,
'   valueVariant = сохраненное значение (Variant/Object).
' Кеш живет только в рамках текущего runtime-сеанса Excel.
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
' Сохраняет Variant-значение по паре ключей:
' namespaceKey - логическая область кеша (модуль/сценарий),
' itemKey      - ключ внутри этой области.
Public Function fn_SetValue( _
    ByVal namespaceKey As String, _
    ByVal itemKey As String, _
    ByVal valueVariant As Variant _
) As Boolean
    Dim cacheMap As Object
    Dim normalizedItemKey As String

    normalizedItemKey = private_NormalizeKey(itemKey)
    If VBA.Len(normalizedItemKey) = 0 Then Exit Function

    ' createIfMissing=True: если namespace еще не существует, создаем его.
    Set cacheMap = private_GetNamespaceCache(namespaceKey, True)
    If cacheMap Is Nothing Then Exit Function

    cacheMap(normalizedItemKey) = valueVariant
    fn_SetValue = True
End Function

' Пытается прочитать Variant-значение из кеша.
' Возвращает True только если ключ найден.
Public Function fn_TryGetValue( _
    ByVal namespaceKey As String, _
    ByVal itemKey As String, _
    ByRef outVariant As Variant _
) As Boolean
    Dim cacheMap As Object
    Dim normalizedItemKey As String

    outVariant = Empty
    normalizedItemKey = private_NormalizeKey(itemKey)
    If VBA.Len(normalizedItemKey) = 0 Then Exit Function

    Set cacheMap = private_GetNamespaceCache(namespaceKey, False)
    If cacheMap Is Nothing Then Exit Function
    If Not cacheMap.Exists(normalizedItemKey) Then Exit Function

    outVariant = cacheMap(normalizedItemKey)
    fn_TryGetValue = True
End Function

' Сохраняет объект в кеше.
' Используется отдельно от fn_SetValue, чтобы корректно работать через Set.
Public Function fn_SetObject( _
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
    fn_SetObject = True
End Function

' Пытается прочитать объект из кеша.
' Возвращает False, если ключ отсутствует или значение не является объектом.
Public Function fn_TryGetObject( _
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
    fn_TryGetObject = Not outObject Is Nothing
End Function

' Удаляет конкретный itemKey в рамках namespace.
' Операция идемпотентна: если ключа/namespace нет, считаем удаление успешным.
Public Function fn_Remove(ByVal namespaceKey As String, ByVal itemKey As String) As Boolean
    Dim cacheMap As Object
    Dim normalizedItemKey As String

    normalizedItemKey = private_NormalizeKey(itemKey)
    If VBA.Len(normalizedItemKey) = 0 Then Exit Function

    Set cacheMap = private_GetNamespaceCache(namespaceKey, False)
    If cacheMap Is Nothing Then
        fn_Remove = True
        Exit Function
    End If

    If cacheMap.Exists(normalizedItemKey) Then cacheMap.Remove normalizedItemKey
    fn_Remove = True
End Function

' Полностью очищает один namespace (все ключи внутри него).
Public Function fn_ClearNamespace(ByVal namespaceKey As String) As Boolean
    Dim normalizedNamespaceKey As String

    normalizedNamespaceKey = private_NormalizeNamespaceKey(namespaceKey)
    private_EnsureStorage
    If g_NamespaceCaches.Exists(normalizedNamespaceKey) Then g_NamespaceCaches.Remove normalizedNamespaceKey
    fn_ClearNamespace = True
End Function

' Полная очистка всех namespace и всех кеш-ключей.
Public Function fn_ClearAll() As Boolean
    private_EnsureStorage
    g_NamespaceCaches.RemoveAll
    fn_ClearAll = True
End Function

' //
' // Internal
' //
' Ленивая инициализация root-хранилища namespace-ов.
Private Sub private_EnsureStorage()
    If g_NamespaceCaches Is Nothing Then
        Set g_NamespaceCaches = ex_Helpers.fn_CreateDictionaryTextCompare()
    End If
End Sub

' Возвращает словарь namespace.
' createIfMissing=False -> только чтение.
' createIfMissing=True  -> при отсутствии namespace создается автоматически.
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

' Нормализация namespace:
' trim + lower-case, пустой ключ заменяем на служебный __default__.
Private Function private_NormalizeNamespaceKey(ByVal keyText As String) As String
    private_NormalizeNamespaceKey = private_NormalizeKey(keyText)
    If VBA.Len(private_NormalizeNamespaceKey) = 0 Then private_NormalizeNamespaceKey = "__default__"
End Function

' Единая нормализация ключей для регистронезависимого кеша.
Private Function private_NormalizeKey(ByVal keyText As String) As String
    private_NormalizeKey = VBA.LCase$(VBA.Trim$(VBA.CStr(keyText)))
End Function
