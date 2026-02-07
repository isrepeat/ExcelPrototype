Attribute VB_Name = "ex_TableComparer"
Option Explicit

Public Function CompareSheets( _
    ByVal oldSheetName As String, _
    ByVal newSheetName As String, _
    ByVal keyColumnName As String _
) As Variant
    Dim wsOld As Worksheet
    Dim wsNew As Worksheet
    
    Set wsOld = ThisWorkbook.Worksheets("g_" & oldSheetName)
    Set wsNew = ThisWorkbook.Worksheets("g_" & newSheetName)
    
    CompareSheets = CompareRanges( _
        wsOld.UsedRange, _
        wsNew.UsedRange, _
        keyColumnName _
    )
End Function

Private Function CompareRanges( _
    ByVal oldRange As Range, _
    ByVal newRange As Range, _
    ByVal keyColumnName As String _
) As Variant
    ' ------------------------------------------------------------------
    ' Сравнивает два диапазона (Old / New) по ключевому столбцу и формирует
    ' выходной 2D-массив с колонками:
    '   [Key, Status, <все колонки New в порядке заголовков>]
    '
    ' Основная логика:
    ' 1) Считать заголовки и найти индекс ключевой колонки в каждом листе
    ' 2) Построить словари строк: ключ -> массив строковых значений
    ' 3) Выделить выходной массив заранее (заголовок + все ключи из New +
    '    отсутствующие в New ключи из Old)
    ' 4) Перебрать ключи New: пометить как Added / Changed / OK и записать значения
    ' 5) Перебрать ключи Old, отсутствующие в New: пометить как Removed
    ' 6) Вернуть готовый массив для записи в лист Result
    ' ------------------------------------------------------------------
    Dim oldHeaders As Variant
    Dim newHeaders As Variant
    
    ' Индексы ключевых колонок (порядок в массивах header'ов)
    Dim oldKeyCol As Long
    Dim newKeyCol As Long
    
    ' Словари строк (ключ -> массив значений)
    Dim oldDict As Object
    Dim newDict As Object
    
    ' Выходной массив и вспомогательные переменные
    Dim outData() As Variant
    Dim outRow As Long
    Dim outColCount As Long
    
    Dim i As Long
    Dim keyValue As String
    Dim key As Variant
    Dim rowArr As Variant
    
    oldHeaders = ReadHeaderRow(oldRange)
    newHeaders = ReadHeaderRow(newRange)
    
    oldKeyCol = FindHeaderIndex(oldHeaders, keyColumnName)
    newKeyCol = FindHeaderIndex(newHeaders, keyColumnName)
    
    If oldKeyCol = 0 Or newKeyCol = 0 Then
        Err.Raise vbObjectError + 1, "CompareRanges", "Key column not found: " & keyColumnName
    End If
    
    Set oldDict = BuildRowDict(oldRange, oldKeyCol)
    Set newDict = BuildRowDict(newRange, newKeyCol)
    
    ' ------------------------------------------------------------------
    ' Выделение выходного массива ОДИН раз (ReDim Preserve не работает для 2D)
    ' Количество строк:
    '   1 - заголовок
    '   + все ключи из New (OK/Changed/Added)
    '   + ключи из Old, отсутствующие в New (Removed)
    ' ------------------------------------------------------------------
    Dim totalRows As Long
    
    totalRows = 1
    
    For Each key In newDict.Keys
        totalRows = totalRows + 1
    Next key
    
    For Each key In oldDict.Keys
        If Not newDict.Exists(CStr(key)) Then
            totalRows = totalRows + 1
        End If
    Next key
    
    outColCount = 2 + UBound(newHeaders)
    ReDim outData(1 To totalRows, 1 To outColCount)
    
    ' Строка заголовка
    outData(1, 1) = keyColumnName
    outData(1, 2) = "Status"
    
    For i = 1 To UBound(newHeaders)
        outData(1, 2 + i) = newHeaders(i)
    Next i
    
    outRow = 1
    
    ' ------------------------------------------------------------------
    ' Добавлено / Изменено / OK (перебор New)
    ' ------------------------------------------------------------------
    For Each key In newDict.Keys
        keyValue = CStr(key)
        outRow = outRow + 1
        
        outData(outRow, 1) = keyValue
        
        If Not oldDict.Exists(keyValue) Then
            outData(outRow, 2) = "Added"
        Else
            If RowsAreDifferent(oldDict(keyValue), newDict(keyValue)) Then
                outData(outRow, 2) = "Changed"
            Else
                outData(outRow, 2) = "OK"
            End If
        End If
        
        rowArr = newDict(keyValue)
        For i = 1 To UBound(rowArr)
            outData(outRow, 2 + i) = rowArr(i)
        Next i
    Next key
    
    ' ------------------------------------------------------------------
    ' Удалено (перебор ключей Old, отсутствующих в New)
    ' ------------------------------------------------------------------
    For Each key In oldDict.Keys
        keyValue = CStr(key)
        
        If Not newDict.Exists(keyValue) Then
            outRow = outRow + 1
            
            outData(outRow, 1) = keyValue
            outData(outRow, 2) = "Removed"
            
            rowArr = oldDict(keyValue)
            For i = 1 To UBound(rowArr)
                outData(outRow, 2 + i) = rowArr(i)
            Next i
        End If
    Next key
    
    CompareRanges = outData
End Function

' -----------------------------------------------------------------------------
' Помощники
' -----------------------------------------------------------------------------
Private Function ReadHeaderRow(ByVal dataRange As Range) As Variant
    Dim colCount As Long
    Dim headers() As Variant
    Dim c As Long
    
    colCount = dataRange.Columns.Count
    ReDim headers(1 To colCount)
    
    For c = 1 To colCount
        headers(c) = CStr(dataRange.Cells(1, c).Value)
    Next c
    
    ReadHeaderRow = headers
End Function

Private Function FindHeaderIndex(ByVal headers As Variant, ByVal name As String) As Long
    Dim i As Long
    
    For i = LBound(headers) To UBound(headers)
        If StrComp(CStr(headers(i)), name, vbTextCompare) = 0 Then
            FindHeaderIndex = i
            Exit Function
        End If
    Next i
    
    FindHeaderIndex = 0
End Function

Private Function BuildRowDict(ByVal dataRange As Range, ByVal keyCol As Long) As Object
    Dim dict As Object
    Dim r As Long
    Dim rowCount As Long
    Dim colCount As Long
    Dim effectiveColCount As Long
    
    Dim keyValue As String
    Dim rowArr() As Variant
    Dim c As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    rowCount = dataRange.Rows.Count
    colCount = dataRange.Columns.Count
    
    ' Определяем последний ненулевой столбец заголовков, чтобы избежать лишних пустых столбцов в UsedRange
    effectiveColCount = 0
    For c = colCount To 1 Step -1
        If Len(CStr(dataRange.Cells(1, c).Value)) > 0 Then
            effectiveColCount = c
            Exit For
        End If
    Next c
    If effectiveColCount = 0 Then effectiveColCount = colCount
    
    For r = 2 To rowCount
        keyValue = CStr(dataRange.Cells(r, keyCol).Value)
        
        If Len(keyValue) > 0 Then
            ReDim rowArr(1 To effectiveColCount)
            
            For c = 1 To effectiveColCount
                rowArr(c) = CStr(dataRange.Cells(r, c).Value)
            Next c
            
            dict(keyValue) = rowArr
        End If
    Next r
    
    Set BuildRowDict = dict
End Function

Private Function RowsAreDifferent(ByVal oldRow As Variant, ByVal newRow As Variant) As Boolean
    Dim i As Long
    
    ' Если массивы разной длины — считаем, что строки отличаются.
    ' Это покрывает случаи, когда одна из строк содержит дополнительные
    ' столбцы (например, из-за лишних заголовков в UsedRange).
    If UBound(oldRow) <> UBound(newRow) Then
        RowsAreDifferent = True
        Exit Function
    End If
    
    For i = LBound(oldRow) To UBound(oldRow)
        If CStr(oldRow(i)) <> CStr(newRow(i)) Then
            RowsAreDifferent = True
            Exit Function
        End If
    Next i
    
    RowsAreDifferent = False
End Function

' Диагностический хелпер - печатает значения строки для заданного ключа в Immediate Window
Public Sub DebugCompareKey(ByVal keyValue As String, ByVal oldSheetName As String, ByVal newSheetName As String)
    Dim wsOld As Worksheet
    Dim wsNew As Worksheet
    Dim oldDict As Object
    Dim newDict As Object
    Dim rowArr As Variant
    Dim i As Long
    
    Set wsOld = ThisWorkbook.Worksheets("g_" & oldSheetName)
    Set wsNew = ThisWorkbook.Worksheets("g_" & newSheetName)
    
    Set oldDict = BuildRowDict(wsOld.UsedRange, FindHeaderIndex(ReadHeaderRow(wsOld.UsedRange), "Id"))
    Set newDict = BuildRowDict(wsNew.UsedRange, FindHeaderIndex(ReadHeaderRow(wsNew.UsedRange), "Id"))
    
    Debug.Print "DebugCompareKey: key=" & keyValue
    If oldDict.Exists(keyValue) Then
        rowArr = oldDict(keyValue)
        Debug.Print "Old row:";
        For i = LBound(rowArr) To UBound(rowArr)
            Debug.Print " [" & i & "]='" & rowArr(i) & "'";
        Next i
        Debug.Print ""
    Else
        Debug.Print "Old row: MISSING"
    End If
    
    If newDict.Exists(keyValue) Then
        rowArr = newDict(keyValue)
        Debug.Print "New row:";
        For i = LBound(rowArr) To UBound(rowArr)
            Debug.Print " [" & i & "]='" & rowArr(i) & "'";
        Next i
        Debug.Print ""
    Else
        Debug.Print "New row: MISSING"
    End If
    
    Debug.Print "RowsAreDifferent=" & RowsAreDifferent(oldDict(keyValue), newDict(keyValue))
End Sub
