VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ExtWorkbookQueryEngine"
Option Explicit

' Единая точка чтения внешних Excel-таблиц. Провайдеры не проверяют состояние
' книги и не содержат параллельные SQL/Worksheet реализации: они формируют
' obj_ExtWorkbookQuery и всегда получают obj_TableDynamic.
Private m_IsDisposed As Boolean
Private m_Connections As Object

Private Const ERROR_TITLE As String = "PrototypeNew / external workbook query"

Private Sub Class_Initialize()
    Set m_Connections = VBA.CreateObject("Scripting.Dictionary")
    m_Connections.CompareMode = 1
End Sub

Private Sub Class_Terminate()
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Me.Dispose
    On Error GoTo 0
End Sub

Public Function Initialize() As Boolean
    m_IsDisposed = False
    If m_Connections Is Nothing Then
        Set m_Connections = VBA.CreateObject("Scripting.Dictionary")
        m_Connections.CompareMode = 1
    End If
    Initialize = True
End Function

Public Sub Dispose()
    Dim key As Variant
    Dim conn As Object

    ' Закрываем только ADO-соединения с закрытыми источниками. Для открытых книг
    ' engine не владеет Workbook/Worksheet и поэтому не закрывает их.
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    If Not m_Connections Is Nothing Then
        For Each key In m_Connections.Keys
            Set conn = m_Connections(key)
            If Not conn Is Nothing Then If conn.State <> 0 Then conn.Close
            Set m_Connections(key) = Nothing
        Next key
        m_Connections.RemoveAll
    End If
    Set m_Connections = Nothing
    On Error GoTo 0
End Sub

' Выполняет один и тот же структурированный запрос двумя способами:
' открытая в текущем экземпляре Excel книга читается из живого Worksheet,
' закрытая — через ACE/ADO без запуска дополнительного экземпляра Excel.
' Выбор делается по полному пути, а не только по имени файла. Это исключает
' случайное чтение одноименной книги из другого каталога.
Public Function TryExecute( _
    ByVal query As obj_ExtWorkbookQuery, _
    ByRef outTable As obj_TableDynamic _
) As Boolean
    Dim validationError As String
    Dim openWorkbook As Workbook

    Set outTable = Nothing
    If m_IsDisposed Then Exit Function
    If query Is Nothing Then
        VBA.MsgBox "PrototypeNew: external workbook query is not specified.", VBA.vbExclamation, ERROR_TITLE
        Exit Function
    End If
    If Not query.TryValidate(validationError) Then
        VBA.MsgBox "PrototypeNew: invalid external workbook query. " & validationError, VBA.vbExclamation, ERROR_TITLE
        Exit Function
    End If
    If VBA.Len(VBA.Dir$(query.SourcePath)) = 0 Then
        VBA.MsgBox "PrototypeNew: external workbook was not found: " & query.SourcePath, VBA.vbExclamation, ERROR_TITLE
        Exit Function
    End If

    Set openWorkbook = private_FindOpenWorkbookByPath(query.SourcePath)
    If Not openWorkbook Is Nothing Then
        ' Книга могла быть открыта пользователем после предыдущего SQL-запроса.
        ' Освобождаем старый ADO handle и читаем актуальные, в том числе еще не
        ' сохраненные, значения непосредственно из Worksheet.
        private_DropConnection query.SourcePath
        TryExecute = private_TryExecuteOpenWorkbook(query, openWorkbook, outTable)
    Else
        TryExecute = private_TryExecuteAdo(query, outTable)
    End If
End Function

Private Function private_TryExecuteAdo( _
    ByVal query As obj_ExtWorkbookQuery, _
    ByRef outTable As obj_TableDynamic _
) As Boolean
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim selectColumns As Collection
    Dim conditions As Collection
    Dim whereClause As String
    Dim recordsetData As Variant
    Dim recordIndex As Long
    Dim startIndex As Long
    Dim endIndex As Long
    Dim stepValue As Long
    Dim resultCount As Long
    Dim i As Long

    ' Closed-workbook backend. SQL получает только заказанные колонки и строки,
    ' поэтому закрытую книгу не требуется открывать через Excel object model.
    ' HDR=YES означает, что KeyColumn/SelectColumns адресуются по заголовкам.
    On Error GoTo EH
    Set selectColumns = query.SelectColumns
    Set conditions = query.BuildEffectiveConditions
    If Not private_TryGetConnection(query.SourcePath, conn) Then Exit Function

    sql = "SELECT "
    If Not query.ReverseOrder And query.MaxRows > 0 Then sql = sql & "TOP " & VBA.CStr(query.MaxRows) & " "
    whereClause = private_BuildWhereClause(conditions)
    If query.SelectAllColumns Then
        sql = sql & "*"
    Else
        sql = sql & private_BuildSelectClause(selectColumns)
    End If
    sql = sql & " FROM " & query.TableRef
    If VBA.Len(whereClause) > 0 Then sql = sql & " WHERE " & whereClause

    Set rs = VBA.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 0, 1
    If query.SelectAllColumns Then
        Set selectColumns = private_BuildRecordsetFieldNames(rs)
        If selectColumns Is Nothing Then GoTo CleanupFail
        If selectColumns.Count = 0 Then GoTo CleanupFail
    End If
    If Not rs.EOF Then recordsetData = rs.GetRows

    If Not private_TryCreateResultTable(selectColumns, outTable) Then GoTo CleanupFail
    If IsEmpty(recordsetData) Then
        private_TryExecuteAdo = True
        GoTo CleanupDone
    End If

    ' У ACE нет надежной служебной колонки с физическим номером строки.
    ' Поэтому для обратного поиска читаем совпадения в их исходном порядке,
    ' разворачиваем recordset в памяти и только затем применяем MaxRows.
    If query.ReverseOrder Then
        startIndex = UBound(recordsetData, 2)
        endIndex = LBound(recordsetData, 2)
        stepValue = -1
    Else
        startIndex = LBound(recordsetData, 2)
        endIndex = UBound(recordsetData, 2)
        stepValue = 1
    End If

    For recordIndex = startIndex To endIndex Step stepValue
        If Not private_TryPushRecordsetRow(outTable, recordsetData, recordIndex, selectColumns.Count) Then GoTo CleanupFail
        resultCount = resultCount + 1
        If query.MaxRows > 0 And resultCount >= query.MaxRows Then Exit For
    Next recordIndex

    private_TryExecuteAdo = True
    GoTo CleanupDone

CleanupFail:
    Set outTable = Nothing
CleanupDone:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> 0 Then rs.Close
    Set rs = Nothing
    Set conn = Nothing
    On Error GoTo 0
    Exit Function

EH:
    private_DropConnection query.SourcePath
    ex_Core.fn_Diagnostic_LogError "external-query:ado-failed workbook='" & query.SourcePath & _
        "' tableRef='" & query.TableRef & "' sql='" & VBA.Replace$(sql, "'", "''") & _
        "' error='" & VBA.Replace$(Err.Description, "'", "''") & "'"
    VBA.MsgBox "PrototypeNew: external workbook SQL query failed." & _
        VBA.vbCrLf & "Workbook: " & query.SourcePath & _
        VBA.vbCrLf & "Range: " & query.TableRef & _
        VBA.vbCrLf & "Error: " & Err.Description, VBA.vbExclamation, ERROR_TITLE
    Resume CleanupFail
End Function

Private Function private_TryExecuteOpenWorkbook( _
    ByVal query As obj_ExtWorkbookQuery, _
    ByVal sourceWorkbook As Workbook, _
    ByRef outTable As obj_TableDynamic _
) As Boolean
    Dim sheetName As String
    Dim startMarker As String
    Dim endMarker As String
    Dim sourceSheet As Worksheet
    Dim startCell As Range
    Dim configuredEndCell As Range
    Dim headerRange As Range
    Dim headerValues As Variant
    Dim headerMap As Object
    Dim selectColumns As Collection
    Dim conditions As Collection
    Dim selectedColumnIndexes() As Long
    Dim conditionColumnIndexes() As Long
    Dim conditionOperations() As Long
    Dim conditionExpectedValues() As String
    Dim conditionNormalizeFlags() As Boolean
    Dim effectiveEndColumn As Long
    Dim effectiveLastRow As Long
    Dim usedLastRow As Long
    Dim usedLastColumn As Long
    Dim selectedValues() As Variant
    Dim conditionValues() As Variant
    Dim matchedRowOffsets As Collection
    Dim matchedRowOffset As Variant
    Dim matchedRowValues As Variant
    Dim rowOffset As Long
    Dim rowStart As Long
    Dim rowEnd As Long
    Dim rowStep As Long
    Dim resultCount As Long
    Dim i As Long
    Dim conditionCount As Long
    Dim rowMatches As Boolean
    Dim useSingleReverseEqualsFastPath As Boolean
    Dim fastConditionMatrix As Variant
    Dim fastActualValue As String
    Dim fastCompareMode As VbCompareMethod
    Dim headerText As String
    Dim condition As obj_ExtWorkbookCondition
    Dim deferSelectedValuesRead As Boolean
    Dim selectedFirstColumn As Long
    Dim selectedLastColumn As Long
    Dim matchedSourceRow As Long

    ' Open-workbook backend повторяет семантику SQL без ADO: первая строка
    ' TableRef считается строкой заголовков, а данные читаются со следующей.
    ' Работаем через массивы Value2, а не перебираем Cells по одной — это важно
    ' для больших Movement/ШПО диапазонов.
    On Error GoTo EH
    If Not private_TryParseTableRef(query.TableRef, sheetName, startMarker, endMarker) Then
        VBA.MsgBox "PrototypeNew: cannot parse external workbook range: " & query.TableRef, VBA.vbExclamation, ERROR_TITLE
        Exit Function
    End If
    If Not private_TryGetWorksheet(sourceWorkbook, sheetName, sourceSheet) Then Exit Function

    Set startCell = sourceSheet.Range(startMarker)
    Set configuredEndCell = sourceSheet.Range(endMarker)
    ' Конфиг задает максимальные допустимые границы таблицы. UsedRange сужает их,
    ' чтобы не загружать до миллиона пустых строк открытого листа.
    usedLastRow = sourceSheet.UsedRange.Row + sourceSheet.UsedRange.Rows.Count - 1
    usedLastColumn = sourceSheet.UsedRange.Column + sourceSheet.UsedRange.Columns.Count - 1
    effectiveLastRow = configuredEndCell.Row
    If usedLastRow < effectiveLastRow Then effectiveLastRow = usedLastRow
    effectiveEndColumn = configuredEndCell.Column
    If usedLastColumn < effectiveEndColumn Then effectiveEndColumn = usedLastColumn
    If effectiveLastRow < startCell.Row Then effectiveLastRow = startCell.Row
    If effectiveEndColumn < startCell.Column Then effectiveEndColumn = startCell.Column

    Set headerRange = sourceSheet.Range( _
        sourceSheet.Cells(startCell.Row, startCell.Column), _
        sourceSheet.Cells(startCell.Row, effectiveEndColumn))
    headerValues = headerRange.Value2
    Set headerMap = VBA.CreateObject("Scripting.Dictionary")
    headerMap.CompareMode = 1
    For i = 1 To headerRange.Columns.Count
        headerText = private_NormalizeHeader(private_MatrixValue(headerValues, 1, i))
        If VBA.Len(headerText) > 0 Then
            If Not headerMap.Exists(headerText) Then headerMap.Add headerText, startCell.Column + i - 1
        End If
    Next i

    Set selectColumns = query.SelectColumns
    If query.SelectAllColumns Then
        Set selectColumns = New Collection
        For i = 1 To headerRange.Columns.Count
            headerText = private_NormalizeHeader(private_MatrixValue(headerValues, 1, i))
            If VBA.Len(headerText) > 0 Then selectColumns.Add headerText
        Next i
        If selectColumns.Count = 0 Then
            VBA.MsgBox "PrototypeNew: external workbook table has no column headers.", _
                VBA.vbExclamation, ERROR_TITLE
            Exit Function
        End If
    End If
    Set conditions = query.BuildEffectiveConditions
    conditionCount = conditions.Count
    ReDim selectedColumnIndexes(1 To selectColumns.Count)
    ReDim selectedValues(1 To selectColumns.Count)
    For i = 1 To selectColumns.Count
        If Not private_TryResolveHeaderColumn(headerMap, VBA.CStr(selectColumns.Item(i)), selectedColumnIndexes(i)) Then Exit Function
        If selectedFirstColumn = 0 Or selectedColumnIndexes(i) < selectedFirstColumn Then
            selectedFirstColumn = selectedColumnIndexes(i)
        End If
        If selectedColumnIndexes(i) > selectedLastColumn Then
            selectedLastColumn = selectedColumnIndexes(i)
        End If
    Next i
    If conditionCount > 0 Then
        ReDim conditionColumnIndexes(1 To conditionCount)
        ReDim conditionValues(1 To conditionCount)
        ReDim conditionOperations(1 To conditionCount)
        ReDim conditionExpectedValues(1 To conditionCount)
        ReDim conditionNormalizeFlags(1 To conditionCount)
        For i = 1 To conditionCount
            Set condition = conditions.Item(i)
            If Not private_TryResolveHeaderColumn(headerMap, condition.ColumnName, conditionColumnIndexes(i)) Then Exit Function
            ' Объекты условий и Collection используются только на этапе подготовки.
            ' В горячем цикле по строкам остаются обычные массивы примитивов,
            ' поэтому стоимость object-property/Collection.Item не умножается
            ' на количество строк внешней таблицы.
            conditionOperations(i) = VBA.CLng(condition.Operation)
            conditionNormalizeFlags(i) = VBA.CBool(condition.NormalizeValue)
            conditionExpectedValues(i) = VBA.CStr(condition.Value)
            If conditionNormalizeFlags(i) Then
                conditionExpectedValues(i) = private_NormalizeKey(conditionExpectedValues(i))
            End If
        Next i
    End If

    If Not private_TryCreateResultTable(selectColumns, outTable) Then Exit Function
    If effectiveLastRow <= startCell.Row Then
        private_TryExecuteOpenWorkbook = True
        Exit Function
    End If

    ' Сначала загружаем только колонки условий. Для SELECT * полные данные
    ' открытого листа читаются позднее и только по совпавшим строкам: история
    ' Movement не должна загружать 23 колонки для всего диапазона до 20000.
    If conditionCount > 0 Then
        For i = 1 To conditionCount
            conditionValues(i) = sourceSheet.Range( _
                sourceSheet.Cells(startCell.Row + 1, conditionColumnIndexes(i)), _
                sourceSheet.Cells(effectiveLastRow, conditionColumnIndexes(i))).Value2
        Next i
    End If
    ' Отложенное чтение имеет смысл только при наличии фильтра: без условий
    ' совпадает каждая строка, и множество точечных обращений к Worksheet было
    ' бы медленнее прежней пакетной загрузки колонок. SelectAllColumns отделяет
    ' широкий preview Movement от обычных узких запросов экспортной валидации.
    deferSelectedValuesRead = query.SelectAllColumns And conditionCount > 0
    If deferSelectedValuesRead Then
        Set matchedRowOffsets = New Collection
    Else
        For i = 1 To selectColumns.Count
            selectedValues(i) = sourceSheet.Range( _
                sourceSheet.Cells(startCell.Row + 1, selectedColumnIndexes(i)), _
                sourceSheet.Cells(effectiveLastRow, selectedColumnIndexes(i))).Value2
        Next i
    End If

    If query.ReverseOrder Then
        rowStart = effectiveLastRow - startCell.Row
        rowEnd = 1
        rowStep = -1
    Else
        rowStart = 1
        rowEnd = effectiveLastRow - startCell.Row
        rowStep = 1
    End If

    ' Частый запрос "последняя физическая строка, где column = value" не
    ' проходит через универсальный dispatcher условий. Он применим к любой
    ' книге/колонке и включается только по структуре запроса, без знания о
    ' Movement или ИПН. Первый найденный снизу элемент сразу завершает поиск.
    useSingleReverseEqualsFastPath = False
    ' VBA And не short-circuit, поэтому к массиву conditionOperations можно
    ' обращаться только после отдельной проверки conditionCount.
    If conditionCount = 1 Then
        useSingleReverseEqualsFastPath = _
            (conditionOperations(1) = en_ExtWorkbookQueryOp.ExtQueryOpEquals) _
            And query.ReverseOrder _
            And (query.MaxRows = 1)
    End If

    If useSingleReverseEqualsFastPath Then
        fastConditionMatrix = conditionValues(1)
        If conditionNormalizeFlags(1) Then
            fastCompareMode = VBA.vbBinaryCompare
        Else
            fastCompareMode = VBA.vbTextCompare
        End If

        For rowOffset = rowStart To rowEnd Step rowStep
            If IsArray(fastConditionMatrix) Then
                fastActualValue = private_SafeText(fastConditionMatrix(rowOffset, 1))
            Else
                fastActualValue = private_SafeText(fastConditionMatrix)
            End If
            If conditionNormalizeFlags(1) Then fastActualValue = private_NormalizeKey(fastActualValue)

            If VBA.StrComp(fastActualValue, conditionExpectedValues(1), fastCompareMode) = 0 Then
                If deferSelectedValuesRead Then
                    ' rowOffset считается от первой строки данных, а не от
                    ' абсолютной строки листа. Это позволяет отложить Worksheet
                    ' read, не удерживая Range на потенциально изменяемой книге.
                    matchedRowOffsets.Add rowOffset
                Else
                    If Not private_TryPushWorksheetRow(outTable, selectedValues, rowOffset, selectColumns.Count) Then
                        Set outTable = Nothing
                        Exit Function
                    End If
                End If
                resultCount = 1
                Exit For
            End If
        Next rowOffset
    Else
        For rowOffset = rowStart To rowEnd Step rowStep
            rowMatches = True
            If conditionCount > 0 Then
                rowMatches = private_RowMatchesPreparedConditions( _
                    conditionValues, _
                    conditionOperations, _
                    conditionExpectedValues, _
                    conditionNormalizeFlags, _
                    conditionCount, _
                    rowOffset)
            End If
            If rowMatches Then
                If deferSelectedValuesRead Then
                    ' Сохраняем offsets в фактическом порядке scan. Поэтому
                    ' последующее чтение не меняет семантику ReverseOrder.
                    matchedRowOffsets.Add rowOffset
                Else
                    If Not private_TryPushWorksheetRow(outTable, selectedValues, rowOffset, selectColumns.Count) Then
                        Set outTable = Nothing
                        Exit Function
                    End If
                End If
                resultCount = resultCount + 1
                If query.MaxRows > 0 And resultCount >= query.MaxRows Then Exit For
            End If
        Next rowOffset
    End If

    If deferSelectedValuesRead Then
        For Each matchedRowOffset In matchedRowOffsets
            matchedSourceRow = startCell.Row + VBA.CLng(matchedRowOffset)
            ' Одна найденная строка читается одним COM-вызовом. Берём
            ' прямоугольник от первой до последней выбранной колонки, а helper
            ' ниже восстановит точный SELECT-порядок и пропустит колонки без
            ' заголовков, если между ними есть физические разрывы.
            matchedRowValues = sourceSheet.Range( _
                sourceSheet.Cells(matchedSourceRow, selectedFirstColumn), _
                sourceSheet.Cells(matchedSourceRow, selectedLastColumn)).Value2
            If Not private_TryPushWorksheetRangeRow( _
                outTable, _
                matchedRowValues, _
                selectedColumnIndexes, _
                selectedFirstColumn, _
                selectColumns.Count) Then
                Set outTable = Nothing
                Exit Function
            End If
        Next matchedRowOffset
    End If

    private_TryExecuteOpenWorkbook = True
    Exit Function

EH:
    ex_Core.fn_Diagnostic_LogError "external-query:open-workbook-failed workbook='" & query.SourcePath & _
        "' tableRef='" & query.TableRef & "' error='" & VBA.Replace$(Err.Description, "'", "''") & "'"
    VBA.MsgBox "PrototypeNew: failed to query an open external workbook." & _
        VBA.vbCrLf & "Workbook: " & query.SourcePath & _
        VBA.vbCrLf & "Range: " & query.TableRef & _
        VBA.vbCrLf & "Error: " & Err.Description, VBA.vbExclamation, ERROR_TITLE
End Function

Private Function private_TryCreateResultTable( _
    ByVal selectColumns As Collection, _
    ByRef outTable As obj_TableDynamic _
) As Boolean
    Dim tableObj As obj_TableDynamic
    Dim columnObj As obj_Column
    Dim i As Long

    Set tableObj = New obj_TableDynamic
    tableObj.SectionTitle = "External Query Result"
    For i = 1 To selectColumns.Count
        Set columnObj = New obj_Column
        columnObj.Position = i
        columnObj.Name = VBA.CStr(selectColumns.Item(i))
        If Not tableObj.PushColumn(columnObj) Then Exit Function
    Next i
    Set outTable = tableObj
    private_TryCreateResultTable = True
End Function

Private Function private_BuildRecordsetFieldNames(ByVal rs As Object) As Collection
    Dim result As Collection
    Dim fieldIndex As Long

    If rs Is Nothing Then Exit Function
    Set result = New Collection
    For fieldIndex = 0 To rs.Fields.Count - 1
        result.Add VBA.CStr(rs.Fields(fieldIndex).Name)
    Next fieldIndex
    Set private_BuildRecordsetFieldNames = result
End Function

Private Function private_TryPushRecordsetRow( _
    ByVal tableObj As obj_TableDynamic, _
    ByVal recordsetData As Variant, _
    ByVal recordIndex As Long, _
    ByVal columnCount As Long _
) As Boolean
    Dim rowObj As obj_Row
    Dim i As Long

    Set rowObj = New obj_Row
    For i = 1 To columnCount
        rowObj.PushCellRaw private_SafeText(recordsetData(i - 1, recordIndex))
    Next i
    private_TryPushRecordsetRow = tableObj.PushRow(rowObj)
End Function

Private Function private_TryPushWorksheetRow( _
    ByVal tableObj As obj_TableDynamic, _
    ByRef selectedValues() As Variant, _
    ByVal rowOffset As Long, _
    ByVal columnCount As Long _
) As Boolean
    Dim rowObj As obj_Row
    Dim i As Long

    Set rowObj = New obj_Row
    For i = 1 To columnCount
        rowObj.PushCellRaw private_SafeText(private_MatrixValue(selectedValues(i), rowOffset, 1))
    Next i
    private_TryPushWorksheetRow = tableObj.PushRow(rowObj)
End Function

Private Function private_TryPushWorksheetRangeRow( _
    ByVal tableObj As obj_TableDynamic, _
    ByVal rowValues As Variant, _
    ByRef selectedColumnIndexes() As Long, _
    ByVal firstSourceColumn As Long, _
    ByVal columnCount As Long _
) As Boolean
    Dim rowObj As obj_Row
    Dim i As Long
    Dim relativeColumnIndex As Long

    If tableObj Is Nothing Then Exit Function
    Set rowObj = New obj_Row
    For i = 1 To columnCount
        ' Индексы query абсолютны относительно Worksheet, тогда как rowValues
        ' начинается с firstSourceColumn. Перевод сохраняет порядок selectColumns
        ' и поддерживает непоследовательные физические колонки.
        relativeColumnIndex = selectedColumnIndexes(i) - firstSourceColumn + 1
        rowObj.PushCellRaw private_SafeText( _
            private_MatrixValue(rowValues, 1, relativeColumnIndex))
    Next i
    private_TryPushWorksheetRangeRow = tableObj.PushRow(rowObj)
End Function

Private Function private_BuildSelectClause(ByVal selectColumns As Collection) As String
    Dim parts As String
    Dim i As Long

    For i = 1 To selectColumns.Count
        If VBA.Len(parts) > 0 Then parts = parts & ", "
        parts = parts & private_QuoteIdentifier(VBA.CStr(selectColumns.Item(i)))
    Next i
    private_BuildSelectClause = parts
End Function

Private Function private_BuildWhereClause(ByVal conditions As Collection) As String
    Dim condition As obj_ExtWorkbookCondition
    Dim conditionSql As String
    Dim result As String

    For Each condition In conditions
        conditionSql = private_BuildConditionSql(condition)
        If VBA.Len(conditionSql) = 0 Then Exit Function
        If VBA.Len(result) > 0 Then result = result & " AND "
        result = result & "(" & conditionSql & ")"
    Next condition
    private_BuildWhereClause = result
End Function

Private Function private_BuildConditionSql(ByVal condition As obj_ExtWorkbookCondition) As String
    Dim valueExpression As String
    Dim expectedValue As String
    Dim likeValue As String

    valueExpression = private_BuildConditionValueExpression(condition.ColumnName, condition.NormalizeValue)
    expectedValue = condition.Value
    If condition.NormalizeValue Then expectedValue = private_NormalizeKey(expectedValue)

    Select Case condition.Operation
        Case en_ExtWorkbookQueryOp.ExtQueryOpEquals
            private_BuildConditionSql = valueExpression & " = " & private_SqlTextLiteral(expectedValue)
        Case en_ExtWorkbookQueryOp.ExtQueryOpNotEquals
            private_BuildConditionSql = valueExpression & " <> " & private_SqlTextLiteral(expectedValue)
        Case en_ExtWorkbookQueryOp.ExtQueryOpContains
            likeValue = "%" & private_EscapeLikeValue(expectedValue) & "%"
            private_BuildConditionSql = valueExpression & " LIKE " & private_SqlTextLiteral(likeValue)
        Case en_ExtWorkbookQueryOp.ExtQueryOpStartsWith
            likeValue = private_EscapeLikeValue(expectedValue) & "%"
            private_BuildConditionSql = valueExpression & " LIKE " & private_SqlTextLiteral(likeValue)
        Case en_ExtWorkbookQueryOp.ExtQueryOpEndsWith
            likeValue = "%" & private_EscapeLikeValue(expectedValue)
            private_BuildConditionSql = valueExpression & " LIKE " & private_SqlTextLiteral(likeValue)
        Case en_ExtWorkbookQueryOp.ExtQueryOpIsEmpty
            private_BuildConditionSql = valueExpression & " = ''"
        Case en_ExtWorkbookQueryOp.ExtQueryOpIsNotEmpty
            private_BuildConditionSql = valueExpression & " <> ''"
        Case en_ExtWorkbookQueryOp.ExtQueryOpGreaterThan
            private_BuildConditionSql = valueExpression & " > " & private_SqlTextLiteral(expectedValue)
        Case en_ExtWorkbookQueryOp.ExtQueryOpLessThan
            private_BuildConditionSql = valueExpression & " < " & private_SqlTextLiteral(expectedValue)
    End Select
End Function

Private Function private_BuildConditionValueExpression( _
    ByVal columnName As String, _
    ByVal normalizeValue As Boolean _
) As String
    Dim quotedColumn As String
    Dim safeExpression As String

    quotedColumn = private_QuoteIdentifier(columnName)
    safeExpression = "CStr(IIf(IsNull(" & quotedColumn & "), '', " & quotedColumn & "))"
    If normalizeValue Then
        private_BuildConditionValueExpression = _
            "LCase(Trim(Replace(Replace(Replace(Replace(" & safeExpression & _
            ", Chr(160), ' '), Chr(13), ''), Chr(10), ''), Chr(9), '')))"
    Else
        private_BuildConditionValueExpression = safeExpression
    End If
End Function

Private Function private_EscapeLikeValue(ByVal valueText As String) As String
    valueText = VBA.Replace$(valueText, "[", "[[]")
    valueText = VBA.Replace$(valueText, "%", "[%]")
    valueText = VBA.Replace$(valueText, "_", "[_]")
    private_EscapeLikeValue = valueText
End Function

Private Function private_TryGetConnection(ByVal sourcePath As String, ByRef outConnection As Object) As Boolean
    Dim cacheKey As String

    ' Соединение переиспользуется в течение жизни provider/engine, чтобы серия
    ' запросов склонений не платила стоимость открытия одного XLSX многократно.
    Set outConnection = Nothing
    cacheKey = private_NormalizePath(sourcePath)
    On Error GoTo EH
    If m_Connections.Exists(cacheKey) Then
        Set outConnection = m_Connections(cacheKey)
        If outConnection.State = 0 Then outConnection.Open private_BuildConnectionString(sourcePath)
    Else
        Set outConnection = VBA.CreateObject("ADODB.Connection")
        outConnection.Open private_BuildConnectionString(sourcePath)
        Set m_Connections(cacheKey) = outConnection
    End If
    private_TryGetConnection = True
    Exit Function
EH:
    private_DropConnection sourcePath
    VBA.MsgBox "PrototypeNew: failed to open external workbook data source." & _
        VBA.vbCrLf & "Workbook: " & sourcePath & _
        VBA.vbCrLf & "Error: " & Err.Description, VBA.vbExclamation, ERROR_TITLE
End Function

Private Sub private_DropConnection(ByVal sourcePath As String)
    Dim cacheKey As String
    Dim conn As Object

    If m_Connections Is Nothing Then Exit Sub
    cacheKey = private_NormalizePath(sourcePath)
    If Not m_Connections.Exists(cacheKey) Then Exit Sub
    On Error Resume Next
    Set conn = m_Connections(cacheKey)
    If Not conn Is Nothing Then If conn.State <> 0 Then conn.Close
    Set m_Connections(cacheKey) = Nothing
    m_Connections.Remove cacheKey
    On Error GoTo 0
End Sub

Private Function private_FindOpenWorkbookByPath(ByVal sourcePath As String) As Workbook
    Dim workbookObj As Workbook
    Dim normalizedPath As String

    normalizedPath = private_NormalizePath(sourcePath)
    For Each workbookObj In Application.Workbooks
        If VBA.StrComp(private_NormalizePath(workbookObj.FullName), normalizedPath, VBA.vbTextCompare) = 0 Then
            Set private_FindOpenWorkbookByPath = workbookObj
            Exit Function
        End If
    Next workbookObj
End Function

Private Function private_TryGetWorksheet( _
    ByVal workbookObj As Workbook, _
    ByVal sheetName As String, _
    ByRef outWorksheet As Worksheet _
) As Boolean
    Set outWorksheet = Nothing
    On Error Resume Next
    Set outWorksheet = workbookObj.Worksheets(sheetName)
    On Error GoTo 0
    If outWorksheet Is Nothing Then
        VBA.MsgBox "PrototypeNew: worksheet '" & sheetName & "' was not found in " & workbookObj.Name & ".", VBA.vbExclamation, ERROR_TITLE
        Exit Function
    End If
    private_TryGetWorksheet = True
End Function

Private Function private_TryParseTableRef( _
    ByVal tableRef As String, _
    ByRef outSheetName As String, _
    ByRef outStartMarker As String, _
    ByRef outEndMarker As String _
) As Boolean
    Dim body As String
    Dim dollarPos As Long
    Dim colonPos As Long

    body = VBA.Trim$(tableRef)
    If VBA.Left$(body, 1) = "[" And VBA.Right$(body, 1) = "]" Then body = VBA.Mid$(body, 2, VBA.Len(body) - 2)
    body = VBA.Replace$(body, "]]", "]")
    dollarPos = VBA.InStrRev(body, "$")
    colonPos = VBA.InStr(dollarPos + 1, body, ":")
    If dollarPos <= 1 Or colonPos <= dollarPos + 1 Then Exit Function
    outSheetName = VBA.Left$(body, dollarPos - 1)
    outStartMarker = VBA.Mid$(body, dollarPos + 1, colonPos - dollarPos - 1)
    outEndMarker = VBA.Mid$(body, colonPos + 1)
    private_TryParseTableRef = (VBA.Len(outSheetName) > 0 And VBA.Len(outStartMarker) > 0 And VBA.Len(outEndMarker) > 0)
End Function

Private Function private_TryResolveHeaderColumn( _
    ByVal headerMap As Object, _
    ByVal headerName As String, _
    ByRef outColumnIndex As Long _
) As Boolean
    Dim normalizedHeader As String

    normalizedHeader = private_NormalizeHeader(headerName)
    If headerMap.Exists(normalizedHeader) Then
        outColumnIndex = VBA.CLng(headerMap(normalizedHeader))
        private_TryResolveHeaderColumn = True
        Exit Function
    End If
    VBA.MsgBox "PrototypeNew: column '" & headerName & "' was not found in the open external workbook.", VBA.vbExclamation, ERROR_TITLE
End Function

Private Function private_RowMatchesPreparedConditions( _
    ByRef conditionValues() As Variant, _
    ByRef conditionOperations() As Long, _
    ByRef conditionExpectedValues() As String, _
    ByRef conditionNormalizeFlags() As Boolean, _
    ByVal conditionCount As Long, _
    ByVal rowOffset As Long _
) As Boolean
    Dim i As Long

    ' Все свойства conditions уже скомпилированы в массивы до начала scan.
    ' Здесь нет обращений к Collection или obj_ExtWorkbookCondition.
    For i = 1 To conditionCount
        If Not private_PreparedConditionMatches( _
            private_MatrixValue(conditionValues(i), rowOffset, 1), _
            conditionOperations(i), _
            conditionExpectedValues(i), _
            conditionNormalizeFlags(i)) Then Exit Function
    Next i
    private_RowMatchesPreparedConditions = True
End Function

Private Function private_PreparedConditionMatches( _
    ByVal rawValue As Variant, _
    ByVal operationValue As Long, _
    ByVal expectedValue As String, _
    ByVal normalizeValue As Boolean _
) As Boolean
    Dim actualValue As String
    Dim compareMode As VbCompareMethod
    Dim compareResult As Long

    actualValue = private_SafeText(rawValue)
    If normalizeValue Then
        actualValue = private_NormalizeKey(actualValue)
        compareMode = VBA.vbBinaryCompare
    Else
        compareMode = VBA.vbTextCompare
    End If
    compareResult = VBA.StrComp(actualValue, expectedValue, compareMode)

    Select Case operationValue
        Case en_ExtWorkbookQueryOp.ExtQueryOpEquals
            private_PreparedConditionMatches = (compareResult = 0)
        Case en_ExtWorkbookQueryOp.ExtQueryOpNotEquals
            private_PreparedConditionMatches = (compareResult <> 0)
        Case en_ExtWorkbookQueryOp.ExtQueryOpContains
            private_PreparedConditionMatches = (VBA.InStr(1, actualValue, expectedValue, compareMode) > 0)
        Case en_ExtWorkbookQueryOp.ExtQueryOpStartsWith
            private_PreparedConditionMatches = (VBA.StrComp(VBA.Left$(actualValue, VBA.Len(expectedValue)), expectedValue, compareMode) = 0)
        Case en_ExtWorkbookQueryOp.ExtQueryOpEndsWith
            private_PreparedConditionMatches = (VBA.StrComp(VBA.Right$(actualValue, VBA.Len(expectedValue)), expectedValue, compareMode) = 0)
        Case en_ExtWorkbookQueryOp.ExtQueryOpIsEmpty
            private_PreparedConditionMatches = (VBA.Len(actualValue) = 0)
        Case en_ExtWorkbookQueryOp.ExtQueryOpIsNotEmpty
            private_PreparedConditionMatches = (VBA.Len(actualValue) > 0)
        Case en_ExtWorkbookQueryOp.ExtQueryOpGreaterThan
            private_PreparedConditionMatches = (compareResult > 0)
        Case en_ExtWorkbookQueryOp.ExtQueryOpLessThan
            private_PreparedConditionMatches = (compareResult < 0)
    End Select
End Function

Private Function private_NormalizeKey(ByVal valueText As String) As String
    valueText = VBA.Replace$(valueText, VBA.ChrW$(160), " ")
    valueText = VBA.Replace$(valueText, VBA.vbCr, VBA.vbNullString)
    valueText = VBA.Replace$(valueText, VBA.vbLf, VBA.vbNullString)
    valueText = VBA.Replace$(valueText, VBA.vbTab, VBA.vbNullString)
    private_NormalizeKey = VBA.LCase$(VBA.Trim$(valueText))
End Function

Private Function private_NormalizeHeader(ByVal valueText As Variant) As String
    private_NormalizeHeader = VBA.LCase$(VBA.Trim$(private_SafeText(valueText)))
End Function

Private Function private_MatrixValue(ByVal matrix As Variant, ByVal rowIndex As Long, ByVal columnIndex As Long) As Variant
    If IsArray(matrix) Then
        private_MatrixValue = matrix(rowIndex, columnIndex)
    ElseIf rowIndex = 1 And columnIndex = 1 Then
        private_MatrixValue = matrix
    Else
        private_MatrixValue = Empty
    End If
End Function

Private Function private_SafeText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    private_SafeText = VBA.CStr(value)
End Function

Private Function private_QuoteIdentifier(ByVal valueText As String) As String
    private_QuoteIdentifier = "[" & VBA.Replace$(VBA.Trim$(valueText), "]", "]]") & "]"
End Function

Private Function private_SqlTextLiteral(ByVal valueText As String) As String
    private_SqlTextLiteral = "'" & VBA.Replace$(valueText, "'", "''") & "'"
End Function

Private Function private_NormalizePath(ByVal sourcePath As String) As String
    private_NormalizePath = VBA.LCase$(VBA.Replace$(VBA.Trim$(sourcePath), "/", "\"))
End Function

Private Function private_BuildConnectionString(ByVal sourcePath As String) As String
    Dim extensionText As String
    Dim propertiesText As String

    extensionText = VBA.LCase$(VBA.Mid$(sourcePath, VBA.InStrRev(sourcePath, ".") + 1))
    Select Case extensionText
        Case "xls": propertiesText = "Excel 8.0"
        Case "xlsm": propertiesText = "Excel 12.0 Macro"
        Case "xlsb": propertiesText = "Excel 12.0"
        Case Else: propertiesText = "Excel 12.0 Xml"
    End Select
    ' HDR=YES сохраняет одинаковый контракт имен колонок с Worksheet backend.
    ' IMEX/ImportMixedTypes уменьшают риск потери текстовых ИПН в mixed-колонках.
    propertiesText = propertiesText & ";HDR=YES;IMEX=1;ReadOnly=True;TypeGuessRows=0;ImportMixedTypes=Text;MAXSCANROWS=0"
    private_BuildConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sourcePath & _
        ";Extended Properties=""" & propertiesText & """;"
End Function
