Option Explicit

' Экспортирует четыре выбранных столбца Excel в новую таблицу Word.
' Ожидаемый порядок: в/ч, населённый пункт/район, область, номер отправления.
' Выделение может содержать строку заголовков — она будет пропущена автоматически.
' Повторно заполненные реквизиты всегда начинают новую строку Word.
Public Sub ExportDraftToWordFormatted()
    Const WD_ORIENT_PORTRAIT As Long = 0
    Const WD_ALIGN_PARAGRAPH_CENTER As Long = 1
    Const WD_ALIGN_VERTICAL_CENTER As Long = 1
    Const WD_AUTO_FIT_FIXED As Long = 0
    Const WD_ROW_HEIGHT_AT_LEAST As Long = 1
    Const RECIPIENT_PREFIX As String = "Військова частина A"
    Const DISPATCH_TYPE As String = "Службовий лист"

    Dim sourceRange As Range
    Dim groups As Object
    Dim groupData As Variant
    Dim currentUnit As String
    Dim currentPlace As String
    Dim currentRegion As String
    Dim shipmentNumber As String
    Dim groupKey As String
    Dim groupCount As Long
    Dim startsNewGroup As Boolean
    Dim pendingNewGroup As Boolean
    Dim rowIndex As Long
    Dim firstDataRow As Long
    Dim skippedRows As Long

    Dim wordApp As Object
    Dim wordDoc As Object
    Dim wordTable As Object
    Dim wordRange As Object
    Dim dataRow As Long

    On Error GoTo ExportError

    Set sourceRange = ResolveSourceRange(ActiveSheet)
    If sourceRange Is Nothing Then
        MsgBox "Не удалось найти исходную таблицу на активном листе." & vbCrLf & _
               "Ожидаются четыре соседних столбца: в/ч, населённый пункт, область и номер отправления." & vbCrLf & _
               "При необходимости выделите эти четыре столбца вручную и повторите экспорт.", _
               vbExclamation, "Экспорт в Word"
        Exit Sub
    End If

    firstDataRow = 1
    If IsSourceHeaderRow(sourceRange) Then firstDataRow = 2

    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbTextCompare

    For rowIndex = firstDataRow To sourceRange.Rows.Count
        startsNewGroup = _
            CleanCellText(sourceRange.Cells(rowIndex, 1).Value2) <> vbNullString Or _
            CleanCellText(sourceRange.Cells(rowIndex, 2).Value2) <> vbNullString Or _
            CleanCellText(sourceRange.Cells(rowIndex, 3).Value2) <> vbNullString

        If startsNewGroup Then
            currentUnit = CleanCellText(sourceRange.Cells(rowIndex, 1).Value2)
            currentPlace = CleanCellText(sourceRange.Cells(rowIndex, 2).Value2)
            currentRegion = CleanCellText(sourceRange.Cells(rowIndex, 3).Value2)
            pendingNewGroup = True
        End If

        shipmentNumber = CleanCellText(sourceRange.Cells(rowIndex, 4).Value2)
        If shipmentNumber <> vbNullString Then
            If currentUnit = vbNullString Or currentPlace = vbNullString Then
                MsgBox "Строка " & sourceRange.Cells(rowIndex, 4).Row & _
                       ": для номера отправления отсутствуют в/ч или населённый пункт.", _
                       vbExclamation, "Экспорт в Word"
                Exit Sub
            End If

            If pendingNewGroup Or groupCount = 0 Then
                groupCount = groupCount + 1
                groupKey = CStr(groupCount)
                groupData = Array(currentUnit, currentPlace, currentRegion, FormatShipmentNumber(shipmentNumber))
                groups.Add groupKey, groupData
                pendingNewGroup = False
            Else
                groupKey = CStr(groupCount)
                groupData = groups(groupKey)
                groupData(3) = groupData(3) & vbCr & FormatShipmentNumber(shipmentNumber)
                groups(groupKey) = groupData
            End If
        ElseIf currentUnit <> vbNullString Or currentPlace <> vbNullString Or currentRegion <> vbNullString Then
            skippedRows = skippedRows + 1
        End If
    Next rowIndex

    If groups.Count = 0 Then
        MsgBox "В четвёртом столбце выделенного диапазона не найдено ни одного номера отправления.", _
               vbExclamation, "Экспорт в Word"
        Exit Sub
    End If

    Set wordApp = GetWordApplication()
    If wordApp Is Nothing Then
        MsgBox "Не удалось запустить Microsoft Word. Проверьте, что Word установлен, и повторите экспорт.", _
               vbCritical, "Экспорт в Word"
        Exit Sub
    End If

    Set wordDoc = wordApp.Documents.Add
    wordDoc.PageSetup.Orientation = WD_ORIENT_PORTRAIT
    wordDoc.PageSetup.TopMargin = wordApp.CentimetersToPoints(1)
    wordDoc.PageSetup.BottomMargin = wordApp.CentimetersToPoints(1)
    wordDoc.PageSetup.LeftMargin = wordApp.CentimetersToPoints(1)
    wordDoc.PageSetup.RightMargin = wordApp.CentimetersToPoints(1)

    Set wordRange = wordDoc.Range(0, 0)
    Set wordTable = wordDoc.Tables.Add(wordRange, groups.Count + 2, 10)
    wordTable.AllowAutoFit = False
    wordTable.AutoFitBehavior WD_AUTO_FIT_FIXED
    wordTable.Borders.Enable = True
    wordTable.Range.Font.Name = "Times New Roman"
    wordTable.Range.Font.Size = 8
    wordTable.Range.ParagraphFormat.SpaceAfter = 0
    wordTable.Range.ParagraphFormat.SpaceBefore = 0
    wordTable.Range.ParagraphFormat.LineSpacingRule = 0
    wordTable.Range.Cells.VerticalAlignment = WD_ALIGN_VERTICAL_CENTER

    FillHeaders wordTable
    ApplyColumnWidths wordTable, wordApp

    dataRow = 3
    For rowIndex = 1 To groupCount
        groupData = groups(CStr(rowIndex))
        PutCellText wordTable, dataRow, 1, CStr(dataRow - 2) & "."
        PutCellText wordTable, dataRow, 2, BuildDestination(CStr(groupData(1)), CStr(groupData(2)))
        PutCellText wordTable, dataRow, 3, RECIPIENT_PREFIX & CStr(groupData(0))
        PutCellText wordTable, dataRow, 4, DISPATCH_TYPE
        PutCellText wordTable, dataRow, 5, CStr(groupData(3))
        PutCellText wordTable, dataRow, 6, "н/в"
        PutCellText wordTable, dataRow, 7, vbNullString
        PutCellText wordTable, dataRow, 8, vbNullString
        PutCellText wordTable, dataRow, 9, vbNullString
        PutCellText wordTable, dataRow, 10, vbNullString

        wordTable.Rows(dataRow).HeightRule = WD_ROW_HEIGHT_AT_LEAST
        wordTable.Cell(dataRow, 1).Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER
        wordTable.Cell(dataRow, 6).Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER
        dataRow = dataRow + 1
    Next rowIndex

    wordTable.Rows(1).HeadingFormat = True
    wordTable.Rows(2).HeadingFormat = True
    wordTable.Rows(1).Range.Bold = True
    wordTable.Rows(2).Range.Bold = True
    wordTable.Rows(1).Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER
    wordTable.Rows(2).Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER

    wordApp.Visible = True
    wordDoc.Activate

    If skippedRows > 0 Then
        MsgBox "Таблица Word создана. Строк без номера отправления пропущено: " & skippedRows & ".", _
               vbInformation, "Экспорт в Word"
    End If
    Exit Sub

ExportError:
    If Not wordApp Is Nothing Then wordApp.Visible = True
    MsgBox "Не удалось сформировать таблицу Word." & vbCrLf & _
           "Ошибка " & Err.Number & ": " & Err.Description, vbCritical, "Экспорт в Word"
End Sub

' Сначала использует корректное ручное выделение, иначе ищет строку заголовков
' на активном листе. Это позволяет запускать экспорт обычной кнопкой Shape.
Private Function ResolveSourceRange(ByVal sourceSheet As Worksheet) As Range
    Dim selectedRange As Range
    Dim usedCell As Range
    Dim headerCell As Range
    Dim candidateRange As Range
    Dim lastDataRow As Long
    Dim columnIndex As Long
    Dim candidateLastRow As Long

    If TypeName(Selection) = "Range" Then
        Set selectedRange = Selection
        If selectedRange.Parent Is sourceSheet Then
            If selectedRange.Areas.Count = 1 And selectedRange.Columns.Count = 4 And _
               selectedRange.Rows.Count > 1 Then
                Set ResolveSourceRange = selectedRange
                Exit Function
            End If
        End If
    End If

    For Each usedCell In sourceSheet.UsedRange.Cells
        If IsUnitHeader(CleanCellText(usedCell.Value2)) Then
            If usedCell.Column <= sourceSheet.Columns.Count - 3 Then
                Set candidateRange = sourceSheet.Range(usedCell, usedCell.Offset(0, 3))
                If IsSourceHeaderRow(candidateRange) Then
                    If headerCell Is Nothing Then
                        Set headerCell = usedCell
                    Else
                        MsgBox "На листе найдено несколько подходящих заголовков «В/ч»." & vbCrLf & _
                               "Выделите нужные четыре столбца вручную вместе со строками данных.", _
                               vbExclamation, "Экспорт в Word"
                        Exit Function
                    End If
                End If
            End If
        End If
    Next usedCell

    If headerCell Is Nothing Then Exit Function

    lastDataRow = headerCell.Row
    For columnIndex = headerCell.Column To headerCell.Column + 3
        candidateLastRow = sourceSheet.Cells(sourceSheet.Rows.Count, columnIndex).End(xlUp).Row
        If candidateLastRow > lastDataRow Then lastDataRow = candidateLastRow
    Next columnIndex

    If lastDataRow <= headerCell.Row Then Exit Function
    Set ResolveSourceRange = sourceSheet.Range( _
        headerCell, sourceSheet.Cells(lastDataRow, headerCell.Column + 3))
End Function

Private Function IsUnitHeader(ByVal cellText As String) As Boolean
    Dim normalizedText As String

    normalizedText = LCase$(Replace(Replace(Trim$(cellText), " ", vbNullString), ".", vbNullString))
    IsUnitHeader = (normalizedText = "в/ч") Or _
                   (normalizedText = "вч") Or _
                   (normalizedText = "в/ч№") Or _
                   (normalizedText = "вч№")
End Function

Private Function IsSourceHeaderRow(ByVal sourceRange As Range) As Boolean
    Dim headerText As String

    headerText = LCase$(CleanCellText(sourceRange.Cells(1, 4).Value2))
    IsSourceHeaderRow = (InStr(1, headerText, "номер", vbTextCompare) > 0) Or _
                        (InStr(1, headerText, "відправ", vbTextCompare) > 0) Or _
                        (InStr(1, headerText, "отправ", vbTextCompare) > 0)
End Function

Private Function CleanCellText(ByVal value As Variant) As String
    Dim result As String

    If IsError(value) Or IsEmpty(value) Then Exit Function
    result = CStr(value)
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")
    result = Replace(result, ChrW$(160), " ")
    CleanCellText = Trim$(result)
End Function

Private Function FormatShipmentNumber(ByVal shipmentNumber As String) As String
    Dim result As String

    result = Trim$(shipmentNumber)
    If Left$(result, 1) = "№" Then result = Trim$(Mid$(result, 2))
    If Left$(result, 5) = "1656/" Then result = Mid$(result, 6)

    FormatShipmentNumber = "№" & ChrW$(160) & "1656/" & result
End Function

Private Function BuildDestination(ByVal place As String, ByVal region As String) As String
    If region = vbNullString Then
        BuildDestination = place
    Else
        BuildDestination = place & vbCr & region
    End If
End Function

Private Function GetWordApplication() As Object
    Dim wordApp As Object

    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    On Error GoTo 0

    If wordApp Is Nothing Then
        On Error Resume Next
        Set wordApp = CreateObject("Word.Application")
        On Error GoTo 0
    End If

    Set GetWordApplication = wordApp
End Function

Private Sub FillHeaders(ByVal wordTable As Object)
    Dim headers As Variant
    Dim columnNumbers As Variant
    Dim columnIndex As Long

    headers = Array( _
        "№№ з/п", _
        "Куди (пункт призначення, район, область)", _
        "Кому (пункт найменування адресата)", _
        "Вид відправлення", _
        "Номери, вказані на відправленнях", _
        "Важливість", _
        "Маса", _
        "Сума плати, грн", _
        "Сума плати, коп", _
        "Присвоєні номери")
    columnNumbers = Array("1", "2", "3", "4", "5", "6", "6а", "6б", "6в", "7")

    For columnIndex = 1 To 10
        PutCellText wordTable, 1, columnIndex, CStr(headers(columnIndex - 1))
        PutCellText wordTable, 2, columnIndex, CStr(columnNumbers(columnIndex - 1))
    Next columnIndex
End Sub

Private Sub ApplyColumnWidths(ByVal wordTable As Object, ByVal wordApp As Object)
    Dim widths As Variant
    Dim columnIndex As Long

    widths = Array(0.55, 3.1, 2.45, 1.75, 2.95, 0.7, 0.7, 0.85, 0.75, 1.15)
    For columnIndex = 1 To 10
        wordTable.Columns(columnIndex).Width = wordApp.CentimetersToPoints(CDbl(widths(columnIndex - 1)))
    Next columnIndex
End Sub

Private Sub PutCellText(ByVal wordTable As Object, ByVal rowIndex As Long, _
                        ByVal columnIndex As Long, ByVal text As String)
    wordTable.Cell(rowIndex, columnIndex).Range.Text = text
End Sub
