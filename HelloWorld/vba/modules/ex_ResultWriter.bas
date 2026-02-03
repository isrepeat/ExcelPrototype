Attribute VB_Name = "ex_ResultWriter"
Option Explicit

' ============================================================================
' ex_ResultWriter
' ----------------------------------------------------------------------------
' Модуль отображения результата.
'
' Отвечает за:
' - создание / получение листа Result
' - очистку листа
' - запись таблицы данных
' - форматирование (шапка, автоширина, фильтры, Freeze Panes)
' - подсветку строк по статусу
'
' НЕ занимается получением данных.
' ============================================================================

Public Sub WriteTableToResultSheet(ByVal tableData As Variant)
    Dim ws As Worksheet
    Dim rowCount As Long
    Dim colCount As Long
    Dim targetRange As Range
    
    ' Получаем лист Result или создаём его, если его нет
    Set ws = GetOrCreateWorksheet("Result")
    
    ' Полностью очищаем лист
    ws.Cells.Clear
    
    ' Определяем размер таблицы
    rowCount = UBound(tableData, 1)
    colCount = UBound(tableData, 2)
    
    ' Записываем данные одним присваиванием (быстро)
    Set targetRange = ws.Range(ws.Cells(1, 1), ws.Cells(rowCount, colCount))
    targetRange.Value = tableData
    
    ' Базовое форматирование таблицы
    FormatAsTable _
        ws, _
        rowCount, _
        colCount
    
    ' Подсветка строк по колонке Status
    ApplyStatusHighlight _
        ws, _
        rowCount, _
        colCount
End Sub

' ----------------------------------------------------------------------------
' Получает лист по имени или создаёт новый, если его нет
' ----------------------------------------------------------------------------
Private Function GetOrCreateWorksheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            Set GetOrCreateWorksheet = ws
            Exit Function
        End If
    Next ws
    
    ' Если лист не найден — создаём новый
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = sheetName
    
    Set GetOrCreateWorksheet = ws
End Function

' ----------------------------------------------------------------------------
' Приводит диапазон к "табличному" виду:
' - жирная шапка
' - автоширина колонок
' - автофильтр
' - закрепление первой строки
' ----------------------------------------------------------------------------
Private Sub FormatAsTable( _
    ByVal ws As Worksheet, _
    ByVal rowCount As Long, _
    ByVal colCount As Long _
)
    Dim headerRange As Range
    Dim allRange As Range
    
    Set headerRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount))
    Set allRange = ws.Range(ws.Cells(1, 1), ws.Cells(rowCount, colCount))
    
    headerRange.Font.Bold = True
    
    allRange.EntireColumn.AutoFit
    allRange.AutoFilter
    
    ' Закрепляем строку заголовков
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    
    allRange.HorizontalAlignment = xlCenter
    headerRange.HorizontalAlignment = xlCenter
End Sub

' ----------------------------------------------------------------------------
' Подсвечивает строки по значению в колонке "Status"
' ----------------------------------------------------------------------------
Private Sub ApplyStatusHighlight( _
    ByVal ws As Worksheet, _
    ByVal rowCount As Long, _
    ByVal colCount As Long _
)
    Dim statusCol As Long
    Dim r As Long
    Dim statusValue As String
    Dim rowRange As Range
    
    ' Находим индекс колонки Status
    statusCol = FindColumnIndex(ws, colCount, "Status")
    If statusCol = 0 Then
        Exit Sub
    End If
    
    ' Проходим по строкам данных
    For r = 2 To rowCount
        statusValue = CStr(ws.Cells(r, statusCol).Value)
        Set rowRange = ws.Range(ws.Cells(r, 1), ws.Cells(r, colCount))
        
        If statusValue = "Added" Then
            rowRange.Interior.Color = RGB(198, 239, 206)
        ElseIf statusValue = "Changed" Then
            rowRange.Interior.Color = RGB(255, 235, 156)
        ElseIf statusValue = "Removed" Then
            rowRange.Interior.Color = RGB(255, 199, 206)
        ElseIf statusValue = "Error" Then
            rowRange.Interior.Color = RGB(244, 176, 132)
        Else
            ' Для остальных статусов — без заливки
            rowRange.Interior.Pattern = xlNone
        End If
    Next r
End Sub

' ----------------------------------------------------------------------------
' Ищет колонку по имени заголовка
' ----------------------------------------------------------------------------
Private Function FindColumnIndex( _
    ByVal ws As Worksheet, _
    ByVal colCount As Long, _
    ByVal headerName As String _
) As Long
    Dim c As Long
    Dim v As String
    
    For c = 1 To colCount
        v = CStr(ws.Cells(1, c).Value)
        If LCase$(v) = LCase$(headerName) Then
            FindColumnIndex = c
            Exit Function
        End If
    Next c
    
    FindColumnIndex = 0
End Function