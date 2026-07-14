VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ExtWorkbookQuery"
Option Explicit

' Описывает запрос к табличному диапазону внешней Excel-книги без привязки
' к способу чтения источника. Engine сам выберет реализацию:
' - открытая книга: Worksheet/Range и фильтрация загруженных массивов;
' - закрытая книга: ACE/ADO SQL с HDR=YES.
' Структурированный контракт намеренный: вызывающий код задает выбираемые
' колонки, AND-условия, направление поиска и лимит строк, поэтому обе
' реализации обязаны возвращать одинаковый obj_TableDynamic без разбора raw SQL.
Private m_SourcePath As String
Private m_TableRef As String
Private m_KeyColumn As String
Private m_KeyValue As String
Private m_SelectColumns As Collection
Private m_Conditions As Collection
Private m_NormalizeKey As Boolean
Private m_ReverseOrder As Boolean
Private m_MaxRows As Long

Private Sub Class_Initialize()
    ' Наиболее частый сценарий — получить одну строку по нормализованному ключу.
    ' Вызывающий код явно меняет эти значения для проверки дублей или поиска
    ' последней физической записи в журнале Movement.
    Set m_SelectColumns = New Collection
    Set m_Conditions = New Collection
    m_NormalizeKey = True
    m_MaxRows = 1
End Sub

Public Property Get SourcePath() As String
    SourcePath = m_SourcePath
End Property

Public Property Let SourcePath(ByVal value As String)
    m_SourcePath = VBA.Trim$(value)
End Property

Public Property Get TableRef() As String
    TableRef = m_TableRef
End Property

Public Property Let TableRef(ByVal value As String)
    ' Формат совместим с ADO, например [Посади$A1:E12000]. Open-workbook
    ' backend разбирает ту же ссылку на имя листа и границы Range.
    m_TableRef = VBA.Trim$(value)
End Property

Public Property Get KeyColumn() As String
    KeyColumn = m_KeyColumn
End Property

Public Property Let KeyColumn(ByVal value As String)
    m_KeyColumn = VBA.Trim$(value)
End Property

Public Property Get KeyValue() As String
    KeyValue = m_KeyValue
End Property

Public Property Let KeyValue(ByVal value As String)
    m_KeyValue = VBA.Trim$(value)
End Property

Public Property Get NormalizeKey() As Boolean
    NormalizeKey = m_NormalizeKey
End Property

Public Property Let NormalizeKey(ByVal value As Boolean)
    ' При True обе реализации одинаково игнорируют регистр, крайние пробелы,
    ' NBSP, переносы строк и табуляцию в ключевой колонке.
    m_NormalizeKey = value
End Property

Public Property Get ReverseOrder() As Boolean
    ReverseOrder = m_ReverseOrder
End Property

Public Property Let ReverseOrder(ByVal value As Boolean)
    ' ReverseOrder используется, когда "последняя" означает последнюю
    ' физическую подходящую строку диапазона, а не сортировку по значению поля.
    m_ReverseOrder = value
End Property

Public Property Get MaxRows() As Long
    MaxRows = m_MaxRows
End Property

Public Property Let MaxRows(ByVal value As Long)
    ' Ноль означает отсутствие лимита. Положительный лимит применяется после
    ' учета ReverseOrder, поэтому ReverseOrder=True/MaxRows=1 дает последнюю строку.
    If value < 0 Then value = 0
    m_MaxRows = value
End Property

Public Property Get SelectColumns() As Collection
    Dim result As Collection
    Dim item As Variant

    Set result = New Collection
    For Each item In m_SelectColumns
        result.Add VBA.CStr(item)
    Next item
    Set SelectColumns = result
End Property

Public Property Get Conditions() As Collection
    Dim result As Collection
    Dim condition As obj_ExtWorkbookCondition

    Set result = New Collection
    For Each condition In m_Conditions
        result.Add condition
    Next condition
    Set Conditions = result
End Property

Public Function AddSelectColumn(ByVal columnName As String) As Boolean
    ' Порядок добавления колонок становится порядком ячеек результирующей строки.
    columnName = VBA.Trim$(columnName)
    If VBA.Len(columnName) = 0 Then Exit Function
    m_SelectColumns.Add columnName
    AddSelectColumn = True
End Function

' Добавляет дополнительный AND-предикат. Старые KeyColumn/KeyValue при этом
' не отключаются, а становятся первым Equals-условием для совместимости.
Public Function AddCondition( _
    ByVal columnName As String, _
    ByVal operatorValue As en_ExtWorkbookQueryOp, _
    Optional ByVal conditionValue As String = VBA.vbNullString, _
    Optional ByVal normalizeValue As Boolean = True _
) As Boolean
    Dim condition As obj_ExtWorkbookCondition
    Dim validationError As String

    Set condition = New obj_ExtWorkbookCondition
    condition.ColumnName = columnName
    condition.Operation = operatorValue
    condition.Value = conditionValue
    condition.NormalizeValue = normalizeValue
    If Not condition.TryValidate(validationError) Then Exit Function

    m_Conditions.Add condition
    AddCondition = True
End Function

' Возвращает полный набор фильтров для engine. Legacy key конвертируется в
' обычное Equals-условие, поэтому backend больше не содержит отдельной логики
' для старого и нового API.
Public Function BuildEffectiveConditions() As Collection
    Dim result As Collection
    Dim condition As obj_ExtWorkbookCondition
    Dim legacyCondition As obj_ExtWorkbookCondition

    Set result = New Collection
    If VBA.Len(m_KeyColumn) > 0 Then
        Set legacyCondition = New obj_ExtWorkbookCondition
        legacyCondition.ColumnName = m_KeyColumn
        legacyCondition.Operation = en_ExtWorkbookQueryOp.ExtQueryOpEquals
        legacyCondition.Value = m_KeyValue
        legacyCondition.NormalizeValue = m_NormalizeKey
        result.Add legacyCondition
    End If
    For Each condition In m_Conditions
        result.Add condition
    Next condition
    Set BuildEffectiveConditions = result
End Function

Public Function TryValidate(ByRef outError As String) As Boolean
    outError = VBA.vbNullString
    If VBA.Len(m_SourcePath) = 0 Then
        outError = "SourcePath is empty."
    ElseIf VBA.Len(m_TableRef) = 0 Then
        outError = "TableRef is empty."
    ElseIf m_SelectColumns Is Nothing Then
        outError = "At least one selected column is required."
    ElseIf m_SelectColumns.Count = 0 Then
        outError = "At least one selected column is required."
    Else
        If Not private_TryValidateConditions(outError) Then Exit Function
        TryValidate = True
    End If
End Function

Private Function private_TryValidateConditions(ByRef outError As String) As Boolean
    Dim condition As obj_ExtWorkbookCondition
    Dim conditionError As String

    If m_Conditions Is Nothing Then
        outError = "Conditions collection is not initialized."
        Exit Function
    End If
    For Each condition In m_Conditions
        If condition Is Nothing Then
            outError = "Conditions collection contains an empty item."
            Exit Function
        End If
        If Not condition.TryValidate(conditionError) Then
            outError = conditionError
            Exit Function
        End If
    Next condition
    private_TryValidateConditions = True
End Function
