Attribute VB_Name = "ex_ResultWriter"
Option Explicit

' =============================================================================
' ex_ResultWriter
' =============================================================================
' Назначение:
'   Отрисовать результат сравнения на листе "Result" с тёмной темой интерфейса.
'
' Обязанности:
'   - Создать / получить лист "Result"
'   - Очистить старое содержимое и записать 2D-массив Variant в ячейки
'   - Применить базовое форматирование (заголовок, фильтры, автоширина, закрепление)
'   - Применить тёмный фон для видимой области (с дополнительным запасом)
'   - Подсветить ТОЛЬКО 3 статуса цветом строки:
'         Added   -> зелёный
'         Changed -> фиолетовый
'         Removed -> красный
'     Любой другой статус (OK / Error) остаётся с тёмным фоном.
'   - Нарисовать сеточные границы (как "All Borders") на тёмном фоне
'
' Примечание о ограничениях Excel:
'   Excel не имеет фонового цвета листа как у UI-канвы.
'   Мы симулируем это заполнением большого диапазона и ограничением области прокрутки.
' =============================================================================

Public Sub WriteTableToResultSheet(ByVal tableData As Variant)
    ' Записывает 2D-массив результата сравнения на лист `g_Result` и применяет
    ' оформление и тему. Порядок действий:
    ' 1) Получить/создать лист
    ' 2) Очистить старое содержимое и область прокрутки
    ' 3) Записать весь массив за одну операцию (быстро)
    ' 4) Применить базовое форматирование таблицы (шрифты, заголовок, фильтры,
    '    авторазмер колонок, закрепление)
    ' 5) Применить тёмную тему и подсветку по статусу (Added/Changed/Removed)
    Dim ws As Worksheet
    Dim rowCount As Long
    Dim colCount As Long
    Dim targetRange As Range
    
    ' Получить или создать лист Result
    Set ws = GetOrCreateWorksheet("Result")
    
    ' Очистить предыдущее содержимое и сбросить область прокрутки
    ws.Cells.Clear
    ws.ScrollArea = ""
    
    ' Определить размер таблицы по 2D-массиву
    rowCount = UBound(tableData, 1)
    colCount = UBound(tableData, 2)
    
    ' Записать данные за одну операцию (быстро)
    Set targetRange = ws.Range(ws.Cells(1, 1), ws.Cells(rowCount, colCount))
    targetRange.Value = tableData
    
    ' Базовое форматирование таблицы (шрифты, заголовок, фильтры, закрепление)
    FormatAsTable _
        ws, _
        rowCount, _
        colCount
    
    ' Применить общую тёмную тему + подсветку по статусу
    ex_SheetTheme.ApplyDarkThemeToSheet _
        ws, _
        True
End Sub


Private Function GetOrCreateWorksheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim fullName As String
    
    ' Добавить префикс g_ к имени листа
    fullName = "g_" & sheetName
    
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, fullName, vbTextCompare) = 0 Then
            Set GetOrCreateWorksheet = ws
            Exit Function
        End If
    Next ws
    
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = fullName
    Call ex_ApplyDefaultSheetView(ws)
    
    Set GetOrCreateWorksheet = ws
End Function

Private Sub FormatAsTable( _
    ByVal ws As Worksheet, _
    ByVal rowCount As Long, _
    ByVal colCount As Long _
)
    Dim headerRange As Range
    Dim allRange As Range
    
    ' headerRange — заголовочная строка, allRange — весь диапазон таблицы
    Set headerRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount))
    Set allRange = ws.Range(ws.Cells(1, 1), ws.Cells(rowCount, colCount))
    
    ' Общие настройки шрифта для таблицы
    allRange.Font.Name = "Segoe UI"
    allRange.Font.Size = 10
    
    ' Выделить заголовок жирным
    headerRange.Font.Bold = True
    
    ' Выравнивание: заголовок по центру, остальные ячейки тоже
    allRange.HorizontalAlignment = xlCenter
    headerRange.HorizontalAlignment = xlCenter
    
    ' Авторазмер колонок и включение фильтров для удобства
    allRange.EntireColumn.AutoFit
    allRange.AutoFilter
    
    ' Показать лист, выбрать первую строку данных и закрепить панель
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub ApplyDarkSheetBackground( _
    ByVal ws As Worksheet, _
    ByVal rowCount As Long, _
    ByVal colCount As Long _
)
    Dim visibleRange As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim bgRange As Range
    
    ws.Activate
    Set visibleRange = ActiveWindow.VisibleRange
    
    lastRow = visibleRange.Row + visibleRange.Rows.Count - 1 + 200
    lastCol = visibleRange.Column + visibleRange.Columns.Count - 1 + 30
    
    If lastRow < rowCount + 200 Then
        lastRow = rowCount + 200
    End If
    
    If lastCol < colCount + 10 Then
        lastCol = colCount + 10
    End If
    
    Set bgRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    bgRange.Interior.Pattern = xlSolid
    bgRange.Interior.Color = RGB(30, 30, 30)
    bgRange.Font.Color = RGB(235, 235, 235)
    
    ActiveWindow.DisplayGridlines = False
    ws.ScrollArea = bgRange.Address
End Sub

Private Sub ApplyStatusHighlight( _
    ByVal ws As Worksheet, _
    ByVal rowCount As Long, _
    ByVal colCount As Long _
)
    Dim statusCol As Long
    Dim r As Long
    Dim statusValue As String
    Dim rowRange As Range
    
    statusCol = FindColumnIndex(ws, colCount, "Status")
    If statusCol = 0 Then
        Exit Sub
    End If
    
    For r = 2 To rowCount
        statusValue = CStr(ws.Cells(r, statusCol).Value)
        Set rowRange = ws.Range(ws.Cells(r, 1), ws.Cells(r, colCount))
        
        Select Case LCase$(statusValue)
            Case "added"
                rowRange.Interior.Pattern = xlSolid
                rowRange.Interior.Color = RGB(46, 125, 50)
            Case "changed"
                rowRange.Interior.Pattern = xlSolid
                rowRange.Interior.Color = RGB(123, 31, 162)
            Case "removed"
                rowRange.Interior.Pattern = xlSolid
                rowRange.Interior.Color = RGB(183, 28, 28)
            Case Else
                rowRange.Interior.Pattern = xlSolid
                rowRange.Interior.Color = RGB(30, 30, 30)
        End Select
        
        rowRange.Font.Color = RGB(235, 235, 235)
    Next r
End Sub

Private Sub ApplyAllBordersToRange(ByVal targetRange As Range)
    With targetRange
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
        
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideHorizontal).Weight = xlThin
        
        .Borders.Color = RGB(80, 80, 80)
    End With
End Sub

Private Function FindColumnIndex( _
    ByVal ws As Worksheet, _
    ByVal colCount As Long, _
    ByVal headerName As String _
) As Long
    Dim c As Long
    Dim v As String
    
    For c = 1 To colCount
        v = CStr(ws.Cells(1, c).Value)
        If StrComp(v, headerName, vbTextCompare) = 0 Then
            FindColumnIndex = c
            Exit Function
        End If
    Next c
    
    FindColumnIndex = 0
End Function
