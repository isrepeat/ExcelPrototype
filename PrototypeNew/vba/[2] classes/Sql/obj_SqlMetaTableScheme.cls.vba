VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_SqlMetaTableScheme"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private m_IsDisposed As Boolean
Private m_SqlMetaRowItems As list__obj_SqlMetaRowItem
' Быстрый индекс строк по значениям ItemIdentity.
' Это map:
'   lookupKey As String -> list__obj_SqlMetaRowItem
' Где lookupKey строится из пары "<identityKey>|<identityValue>".
' Примеры содержимого:
'   "row_kind|Meta"    -> list(metaRow1, metaRow2, metaRow3)
'   "position|Before"  -> list(row5, row8)
'   "is_virtual|False" -> list(row1, row2, row3, ...)
Private m_mapLookupKeyToListSqlMetaRowItem As Object

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
    Set m_SqlMetaRowItems = New list__obj_SqlMetaRowItem
    Set m_mapLookupKeyToListSqlMetaRowItem = ex_Helpers.fn_CreateDictionaryTextCompare()
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Me.Dispose
    On Error GoTo 0
End Sub

' //
' // Properties
' //
Public Property Get Count() As Long
    If m_SqlMetaRowItems Is Nothing Then Exit Property
    Count = m_SqlMetaRowItems.Count
End Property

Public Property Get SqlMetaRowItems() As list__obj_SqlMetaRowItem
    ' Основное хранилище остается последовательным списком.
    ' Порядок элементов здесь совпадает с порядком AddRow, то есть обычно
    ' с порядком прохода по исходной таблице сверху вниз.
    If m_SqlMetaRowItems Is Nothing Then Set m_SqlMetaRowItems = New list__obj_SqlMetaRowItem
    Set SqlMetaRowItems = m_SqlMetaRowItems
End Property

' //
' // API
' //
Public Function Initialize() As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    m_IsDisposed = False
    Set m_SqlMetaRowItems = New list__obj_SqlMetaRowItem
    Set m_mapLookupKeyToListSqlMetaRowItem = ex_Helpers.fn_CreateDictionaryTextCompare()
    Initialize = Not m_mapLookupKeyToListSqlMetaRowItem Is Nothing
End Function

Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True

    On Error Resume Next
    Set m_SqlMetaRowItems = Nothing
    Set m_mapLookupKeyToListSqlMetaRowItem = Nothing
    On Error GoTo 0
End Sub

Public Function AddRow( _
    ByVal sourceRow As obj_Row, _
    ByVal itemIdentity As obj_ItemIdentity, _
    Optional ByRef outSqlMetaRowItem As obj_SqlMetaRowItem _
) As Boolean
    Dim sqlMetaRowItem As obj_SqlMetaRowItem
    Dim normalizedIdentity As obj_ItemIdentity

    Set outSqlMetaRowItem = Nothing

    ' AddRow сохраняет не "голую" строку, а пару:
    '   SourceRow     - исходная/сгенерированная строка;
    '   ItemIdentity  - набор признаков этой строки.
    ' Именно ItemIdentity отвечает на вопросы вида:
    ' "это meta-строка?", "позиция Before или After?",
    ' "строка виртуальная?", "какой owner сейчас активен?".
    If itemIdentity Is Nothing Then
        Set normalizedIdentity = New obj_ItemIdentity
    Else
        ' Клонируем identity, чтобы вызывающий код мог дальше менять свой объект
        ' и случайно не переписал уже сохраненные признаки строки.
        Set normalizedIdentity = itemIdentity.Clone
    End If
    If normalizedIdentity Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "SqlMetaTableScheme.AddRow: normalized identity is Nothing."
#End If
        Exit Function
    End If

    Set sqlMetaRowItem = New obj_SqlMetaRowItem
    If sqlMetaRowItem Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "SqlMetaTableScheme.AddRow: failed to create obj_SqlMetaRowItem."
#End If
        Exit Function
    End If
    ' obj_SqlMetaRowItem внутри тоже клонирует SourceRow и ItemIdentity.
    ' В итоге схема хранит собственный снимок данных на момент AddRow.
    If Not sqlMetaRowItem.Initialize(sourceRow, normalizedIdentity) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "SqlMetaTableScheme.AddRow: SqlMetaRowItem.Initialize failed."
#End If
        Exit Function
    End If

    If m_SqlMetaRowItems Is Nothing Then Set m_SqlMetaRowItems = New list__obj_SqlMetaRowItem
    If Not m_SqlMetaRowItems.Add(sqlMetaRowItem) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "SqlMetaTableScheme.AddRow: failed to add item to main list."
#End If
        Exit Function
    End If

    ' После добавления в основной список строим вторичный индекс.
    ' Он нужен, чтобы BuildResult мог быстро доставать группы строк по признакам,
    ' не проходя весь список каждый раз вручную.
    If Not private_IndexRowItem(sqlMetaRowItem) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "SqlMetaTableScheme.AddRow: failed to index row item."
#End If
        Exit Function
    End If

    Set outSqlMetaRowItem = sqlMetaRowItem
    AddRow = True
End Function

Public Function TryGetSqlMetaRowItemsByIdentityValue( _
    ByVal identityKey As String, _
    ByVal expectedValue As Variant, _
    ByRef outSqlMetaRowItems As list__obj_SqlMetaRowItem _
) As Boolean
    Dim lookupKey As String
    Dim sqlMetaRowItem As obj_SqlMetaRowItem
    Dim indexedSqlMetaRowItems As list__obj_SqlMetaRowItem
    Dim i As Long

    Set outSqlMetaRowItems = New list__obj_SqlMetaRowItem
    identityKey = VBA.Trim$(identityKey)
    If VBA.Len(identityKey) = 0 Then
        ' Пустой ключ считаем корректным "ничего не найдено".
        ' Ошибка здесь не нужна: вызывающий код просто получит пустой список.
        TryGetSqlMetaRowItemsByIdentityValue = True
        Exit Function
    End If

    ' expectedValue - это ожидаемое значение для указанного identityKey.
    ' Например:
    '   identityKey    = "RowKind"
    '   expectedValue  = "Meta"
    ' Тогда метод вернет все obj_SqlMetaRowItem, у которых:
    '   ItemIdentity("RowKind") = "Meta"
    ' Пара key+value важна: одного ключа мало, потому что под одним ключом
    ' могут быть разные значения: RowKind=Owner, RowKind=Meta, RowKind=Common.
    lookupKey = private_BuildIdentityLookupKey(identityKey, expectedValue)
    If VBA.Len(lookupKey) = 0 Then
        TryGetSqlMetaRowItemsByIdentityValue = True
        Exit Function
    End If

    If m_mapLookupKeyToListSqlMetaRowItem Is Nothing Then
        TryGetSqlMetaRowItemsByIdentityValue = True
        Exit Function
    End If
    If Not m_mapLookupKeyToListSqlMetaRowItem.Exists(lookupKey) Then
        TryGetSqlMetaRowItemsByIdentityValue = True
        Exit Function
    End If

    Set indexedSqlMetaRowItems = m_mapLookupKeyToListSqlMetaRowItem(lookupKey)
    If indexedSqlMetaRowItems Is Nothing Then
        TryGetSqlMetaRowItemsByIdentityValue = True
        Exit Function
    End If

    ' В индекс кладутся сами sqlMetaRowItem-объекты. Здесь мы переносим найденные
    ' элементы в типизированный список, чтобы внешний код не работал
    ' напрямую со словарем индекса.
    For i = 1 To indexedSqlMetaRowItems.Count
        Set sqlMetaRowItem = indexedSqlMetaRowItems.Item(i)
        If sqlMetaRowItem Is Nothing Then GoTo ContinueItem
        If Not outSqlMetaRowItems.Add(sqlMetaRowItem) Then Exit Function
ContinueItem:
    Next i

    TryGetSqlMetaRowItemsByIdentityValue = True
End Function

Public Function TryGetRowsByIdentityValue( _
    ByVal identityKey As String, _
    ByVal expectedValue As Variant, _
    ByRef outRows As list__obj_Row _
) As Boolean
    Dim sqlMetaRowItems As list__obj_SqlMetaRowItem
    Dim sqlMetaRowItem As obj_SqlMetaRowItem
    Dim sourceRow As obj_Row
    Dim i As Long

    Set outRows = New list__obj_Row
    ' Упрощенная обертка над TryGetSqlMetaRowItemsByIdentityValue.
    ' Если потребителю нужны только строки, а не их identity,
    ' этот метод достает SourceRow из каждого найденного sqlMetaRowItem.
    If Not Me.TryGetSqlMetaRowItemsByIdentityValue(identityKey, expectedValue, sqlMetaRowItems) Then Exit Function
    If sqlMetaRowItems Is Nothing Then
        TryGetRowsByIdentityValue = True
        Exit Function
    End If

    For i = 1 To sqlMetaRowItems.Count
        Set sqlMetaRowItem = sqlMetaRowItems.Item(i)
        If sqlMetaRowItem Is Nothing Then GoTo ContinueRowItem
        Set sourceRow = sqlMetaRowItem.SourceRow
        If sourceRow Is Nothing Then GoTo ContinueRowItem
        If Not outRows.Add(sourceRow) Then Exit Function
ContinueRowItem:
    Next i

    TryGetRowsByIdentityValue = True
End Function

Public Sub ClearRows()
    ' Очищаем сразу оба хранилища:
    '   m_SqlMetaRowItems                  - последовательный список;
    '   m_mapLookupKeyToListSqlMetaRowItem - быстрый индекс по признакам.
    Set m_SqlMetaRowItems = New list__obj_SqlMetaRowItem
    Set m_mapLookupKeyToListSqlMetaRowItem = ex_Helpers.fn_CreateDictionaryTextCompare()
End Sub

' //
' // Internal
' //
Private Function private_IndexRowItem(ByVal sqlMetaRowItem As obj_SqlMetaRowItem) As Boolean
    Dim itemIdentity As obj_ItemIdentity
    Dim identityKeys As list__obj_String
    Dim identityKey As String
    Dim valueCandidate As Variant
    Dim lookupKey As String
    Dim indexedSqlMetaRowItems As list__obj_SqlMetaRowItem
    Dim i As Long

    If sqlMetaRowItem Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "SqlMetaTableScheme.IndexRowItem: row item is Nothing."
#End If
        Exit Function
    End If
    Set itemIdentity = sqlMetaRowItem.ItemIdentity
    If itemIdentity Is Nothing Then
        ' Строка без identity все равно может жить в m_SqlMetaRowItems,
        ' но индексировать ее не по чему.
        private_IndexRowItem = True
        Exit Function
    End If

    If m_mapLookupKeyToListSqlMetaRowItem Is Nothing Then
        Set m_mapLookupKeyToListSqlMetaRowItem = ex_Helpers.fn_CreateDictionaryTextCompare()
        If m_mapLookupKeyToListSqlMetaRowItem Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "SqlMetaTableScheme.IndexRowItem: failed to create lookup map."
#End If
            Exit Function
        End If
    End If

    Set identityKeys = itemIdentity.Keys
    If identityKeys Is Nothing Then
        private_IndexRowItem = True
        Exit Function
    End If

    ' Один sqlMetaRowItem индексируется по всем своим признакам.
    ' Если identity содержит:
    '   RowKind=Meta
    '   Position=Before
    '   IsVirtual=False
    ' то один и тот же sqlMetaRowItem попадет сразу в три группы индекса:
    '   "rowkind|Meta"
    '   "position|Before"
    '   "isvirtual|False"
    ' Это позволяет потом выбирать строки с разных точек зрения.
    For i = 1 To identityKeys.Count
        identityKey = identityKeys.Item(i)
        If Not itemIdentity.TryGetValue(identityKey, valueCandidate) Then GoTo ContinueKey
        lookupKey = private_BuildIdentityLookupKey(identityKey, valueCandidate)
        If VBA.Len(lookupKey) = 0 Then GoTo ContinueKey

        If m_mapLookupKeyToListSqlMetaRowItem.Exists(lookupKey) Then
            Set indexedSqlMetaRowItems = m_mapLookupKeyToListSqlMetaRowItem(lookupKey)
        Else
            ' Для каждой пары key+value храним list__obj_SqlMetaRowItem
            ' со всеми строками, которые имеют такой признак.
            Set indexedSqlMetaRowItems = New list__obj_SqlMetaRowItem
            m_mapLookupKeyToListSqlMetaRowItem.Add lookupKey, indexedSqlMetaRowItems
        End If

        If indexedSqlMetaRowItems Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "SqlMetaTableScheme.IndexRowItem: indexed list is Nothing for lookupKey='" & lookupKey & "'."
#End If
            Exit Function
        End If
        If Not indexedSqlMetaRowItems.Add(sqlMetaRowItem) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "SqlMetaTableScheme.IndexRowItem: failed to add item to indexed list lookupKey='" & lookupKey & "'."
#End If
            Exit Function
        End If
ContinueKey:
    Next i

    private_IndexRowItem = True
End Function

Private Function private_BuildIdentityLookupKey(ByVal identityKey As String, ByVal identityValue As Variant) As String
    Dim keyText As String
    Dim valueText As String

    keyText = VBA.LCase$(VBA.Trim$(VBA.CStr(identityKey)))
    If VBA.Len(keyText) = 0 Then Exit Function

    ' Нормализованный lookupKey должен строиться одинаково и при записи
    ' в индекс, и при чтении через expectedValue. Поэтому поиск:
    '   TryGetSqlMetaRowItemsByIdentityValue("RowKind", "Meta", ...)
    ' попадет ровно в ту же ячейку индекса, что была создана для
    ' ItemIdentity.Add("RowKind", "Meta").
    valueText = private_NormalizeIdentityValue(identityValue)
    private_BuildIdentityLookupKey = keyText & "|" & valueText
End Function

Private Function private_NormalizeIdentityValue(ByVal identityValue As Variant) As String
    ' Значения identity приводим к тексту для ключа словаря.
    ' Для объектов не пытаемся сериализовать внутреннее состояние:
    ' индексируем только факт Nothing или тип объекта.
    If VBA.IsObject(identityValue) Then
        If identityValue Is Nothing Then
            private_NormalizeIdentityValue = "<nothing>"
            Exit Function
        End If

        private_NormalizeIdentityValue = "<object:" & VBA.TypeName(identityValue) & ">"
        Exit Function
    End If

    private_NormalizeIdentityValue = VBA.Trim$(VBA.CStr(identityValue))
End Function
