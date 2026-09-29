Option Explicit

Private Const EVENT_AWOL As String = "Самовільне залишення частини"
Private Const EVENT_BUSINESS_TRIP As String = "Відрядження"
Private Const EVENT_RETURN_OTHER_UNIT As String = "Повернення до іншої частини"
Private Const RESULT_SHEET_NAME As String = "Повернення_до_іншої_частини"

Private Const HEADER_FULL_NAME As String = "ПІБ"
Private Const HEADER_TAX_ID As String = "ІПН"
Private Const HEADER_EVENT As String = "Подія"
Private Const HEADER_FROM_PROVISION As String = "З продовольчого"
Private Const HEADER_DEPARTURE As String = "Вибуття"
Private Const HEADER_ARRIVAL As String = "Прибуття"
Private Const HEADER_TO_PROVISION As String = "На продовольче"
Private Const HEADER_DEPARTURE_TERM As String = "Термін вибуття"
Private Const HEADER_PLANNED_ARRIVAL As String = "Планове прибуття"
Private Const HEADER_DEPARTURE_ORDER As String = "Наказ вибуття"
Private Const HEADER_ARRIVAL_ORDER As String = "Наказ прибуття"
Private Const HEADER_DEPARTURE_REASON As String = "Підстава вибуття"
Private Const HEADER_ARRIVAL_REASON As String = "Підстава прибуття"

Private Const RESULT_COL_SOURCE_AWOL As Long = 1
Private Const RESULT_COL_SOURCE_BUSINESS_TRIP As Long = 2
Private Const RESULT_COL_GAP_DAYS As Long = 3
Private Const RESULT_COL_STATUS As Long = 4
Private Const RESULT_TABLE_COL_OFFSET As Long = 4

Private Const RESULT_STATUS_CREATED As String = "Created"
Private Const RESULT_STATUS_INVALID_DATES As String = "Invalid dates"

Public Sub fn_FindReturnToOtherUnitEvents()
    Dim tbl As ListObject
    Dim resultWs As Worksheet
    Dim people As Object
    Dim personKey As Variant
    Dim rowsByPerson As Collection
    Dim visibleRows As Collection
    Dim oneRow As Range
    Dim i As Long
    Dim resultRow As Long
    Dim awolRow As Range
    Dim tripRow As Range
    Dim awolArrivalDate As Variant
    Dim tripDepartureDate As Variant
    Dim gapDays As Long
    Dim createdCount As Long
    Dim skippedCount As Long
    
    Set tbl = private_GetSourceTable()
    
    If tbl Is Nothing Then
        VBA.MsgBox "Source table was not found. Select any cell inside the required table and run the macro again.", VBA.vbExclamation
        Exit Sub
    End If
    
    If tbl.DataBodyRange Is Nothing Then
        VBA.MsgBox "The selected table does not contain data rows.", VBA.vbExclamation
        Exit Sub
    End If
    
    If Not private_ValidateRequiredColumns(tbl) Then Exit Sub
    
    Set visibleRows = private_GetVisibleTableRows(tbl)
    
    If visibleRows.Count < 2 Then
        VBA.MsgBox "At least two visible table rows are required for analysis.", VBA.vbInformation
        Exit Sub
    End If
    
    Set people = VBA.CreateObject("Scripting.Dictionary")
    
    For Each oneRow In visibleRows
        personKey = private_GetPersonKey(tbl, oneRow)
        
        If VBA.Len(VBA.CStr(personKey)) > 0 Then
            If Not people.Exists(personKey) Then
                Set people(personKey) = New Collection
            End If
            
            people(personKey).Add oneRow
        End If
    Next oneRow
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set resultWs = private_PrepareResultSheet(tbl)
    resultRow = 2
    
    For Each personKey In people.Keys
        Set rowsByPerson = people(personKey)
        
        If rowsByPerson.Count >= 2 Then
            For i = 1 To rowsByPerson.Count - 1
                Set awolRow = rowsByPerson(i)
                Set tripRow = rowsByPerson(i + 1)
                
                If private_IsEventRow(tbl, awolRow, EVENT_AWOL) _
                   And private_IsEventRow(tbl, tripRow, EVENT_BUSINESS_TRIP) Then
                    
                    awolArrivalDate = private_GetCellValue(tbl, awolRow, HEADER_ARRIVAL)
                    tripDepartureDate = private_GetCellValue(tbl, tripRow, HEADER_DEPARTURE)
                    
                    If VBA.IsDate(awolArrivalDate) And VBA.IsDate(tripDepartureDate) Then
                        gapDays = VBA.DateDiff("d", VBA.DateValue(VBA.CDate(awolArrivalDate)), VBA.DateValue(VBA.CDate(tripDepartureDate)))
                        
                        If gapDays > 0 Then
                            private_WriteGeneratedRow tbl, resultWs, resultRow, awolRow, tripRow, VBA.CDate(awolArrivalDate), VBA.CDate(tripDepartureDate), gapDays
                            resultRow = resultRow + 1
                            createdCount = createdCount + 1
                        End If
                    Else
                        private_WriteInvalidDateRow tbl, resultWs, resultRow, awolRow, tripRow, awolArrivalDate, tripDepartureDate
                        resultRow = resultRow + 1
                        skippedCount = skippedCount + 1
                    End If
                End If
            Next i
        End If
    Next personKey
    
    If resultRow > 2 Then
        private_FormatResultSheet resultWs, tbl, resultRow - 1
    Else
        private_FormatEmptyResultSheet resultWs, tbl
    End If

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    VBA.MsgBox "Done. Rows created for copying: " & createdCount & "." & _
           VBA.IIf(skippedCount > 0, VBA.vbCrLf & "Rows skipped because of invalid dates: " & skippedCount, VBA.vbNullString), VBA.vbInformation
End Sub

Private Function private_GetSourceTable() As ListObject
    On Error Resume Next
    Set private_GetSourceTable = Application.ActiveCell.ListObject
    On Error GoTo 0
    
    If private_GetSourceTable Is Nothing Then
        If Application.ActiveSheet.ListObjects.Count = 1 Then
            Set private_GetSourceTable = Application.ActiveSheet.ListObjects(1)
        End If
    End If
End Function

Private Function private_ValidateRequiredColumns(ByVal tbl As ListObject) As Boolean
    Dim requiredHeaders As Variant
    Dim header As Variant
    Dim missing As String
    
    requiredHeaders = VBA.Array( _
        HEADER_FULL_NAME, _
        HEADER_TAX_ID, _
        HEADER_EVENT, _
        HEADER_FROM_PROVISION, _
        HEADER_DEPARTURE, _
        HEADER_ARRIVAL, _
        HEADER_TO_PROVISION, _
        HEADER_DEPARTURE_ORDER, _
        HEADER_ARRIVAL_ORDER, _
        HEADER_DEPARTURE_REASON, _
        HEADER_ARRIVAL_REASON _
    )
    
    For Each header In requiredHeaders
        If private_GetColumnIndex(tbl, VBA.CStr(header)) = 0 Then
            missing = missing & VBA.IIf(VBA.Len(missing) > 0, ", ", VBA.vbNullString) & """" & VBA.CStr(header) & """"
        End If
    Next header
    
    If VBA.Len(missing) > 0 Then
        VBA.MsgBox "Required table columns were not found: " & missing & ".", VBA.vbExclamation
        private_ValidateRequiredColumns = False
    Else
        private_ValidateRequiredColumns = True
    End If
End Function

Private Function private_GetVisibleTableRows(ByVal tbl As ListObject) As Collection
    Dim rows As New Collection
    Dim oneRow As Range
    
    For Each oneRow In tbl.DataBodyRange.Rows
        If Not oneRow.EntireRow.Hidden Then rows.Add oneRow
    Next oneRow
    
    Set private_GetVisibleTableRows = rows
End Function

Private Function private_GetPersonKey(ByVal tbl As ListObject, ByVal oneRow As Range) As String
    Dim ipn As String
    Dim pib As String
    
    ipn = private_NormalizeText(VBA.CStr(private_GetCellValue(tbl, oneRow, HEADER_TAX_ID)))
    pib = private_NormalizeText(VBA.CStr(private_GetCellValue(tbl, oneRow, HEADER_FULL_NAME)))
    
    If VBA.Len(ipn) > 0 Then
        private_GetPersonKey = ipn
    Else
        private_GetPersonKey = pib
    End If
End Function

Private Function private_PrepareResultSheet(ByVal tbl As ListObject) As Worksheet
    Dim ws As Worksheet
    Dim col As Long
    
    Set ws = private_GetOrCreateSheet(RESULT_SHEET_NAME)
    ws.Cells.Clear
    
    private_ApplyDarkBaseStyle ws
    
    ws.Cells(1, RESULT_COL_SOURCE_AWOL).Value = "SourceRow_СЗЧ"
    ws.Cells(1, RESULT_COL_SOURCE_BUSINESS_TRIP).Value = "SourceRow_Відрядження"
    ws.Cells(1, RESULT_COL_GAP_DAYS).Value = "GapDays"
    ws.Cells(1, RESULT_COL_STATUS).Value = "Status"
    
    For col = 1 To tbl.ListColumns.Count
        ws.Cells(1, col + RESULT_TABLE_COL_OFFSET).Value = tbl.HeaderRowRange.Cells(1, col).Value
    Next col
    
    Set private_PrepareResultSheet = ws
End Function

Private Function private_GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set private_GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If private_GetOrCreateSheet Is Nothing Then
        Set private_GetOrCreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        private_GetOrCreateSheet.Name = sheetName
    End If
End Function

Private Sub private_WriteGeneratedRow( _
    ByVal tbl As ListObject, _
    ByVal ws As Worksheet, _
    ByVal targetRow As Long, _
    ByVal awolRow As Range, _
    ByVal tripRow As Range, _
    ByVal awolArrivalDate As Date, _
    ByVal tripDepartureDate As Date, _
    ByVal gapDays As Long)

    Dim col As Long
    Dim header As String
    Dim valueToWrite As Variant
    Dim tripDepartureOrder As Variant
    Dim tripDepartureReason As Variant
    
    tripDepartureOrder = private_GetCellValue(tbl, tripRow, HEADER_DEPARTURE_ORDER)
    tripDepartureReason = private_GetCellValue(tbl, tripRow, HEADER_DEPARTURE_REASON)
    
    ws.Cells(targetRow, RESULT_COL_SOURCE_AWOL).Value = awolRow.Row
    ws.Cells(targetRow, RESULT_COL_SOURCE_BUSINESS_TRIP).Value = tripRow.Row
    ws.Cells(targetRow, RESULT_COL_GAP_DAYS).Value = gapDays
    ws.Cells(targetRow, RESULT_COL_STATUS).Value = RESULT_STATUS_CREATED
    
    For col = 1 To tbl.ListColumns.Count
        header = VBA.CStr(tbl.HeaderRowRange.Cells(1, col).Value)
        valueToWrite = tripRow.Cells(1, col).Value
        
        Select Case private_NormalizeHeader(header)
            Case private_NormalizeHeader(HEADER_EVENT)
                valueToWrite = EVENT_RETURN_OTHER_UNIT
            
            Case private_NormalizeHeader(HEADER_FROM_PROVISION)
                valueToWrite = awolArrivalDate
            
            Case private_NormalizeHeader(HEADER_DEPARTURE)
                valueToWrite = awolArrivalDate
            
            Case private_NormalizeHeader(HEADER_ARRIVAL)
                valueToWrite = tripDepartureDate
            
            Case private_NormalizeHeader(HEADER_TO_PROVISION)
                valueToWrite = tripDepartureDate
            
            Case private_NormalizeHeader(HEADER_DEPARTURE_ORDER)
                valueToWrite = tripDepartureOrder
            
            Case private_NormalizeHeader(HEADER_ARRIVAL_ORDER)
                valueToWrite = tripDepartureOrder
            
            Case private_NormalizeHeader(HEADER_DEPARTURE_REASON)
                valueToWrite = tripDepartureReason
            
            Case private_NormalizeHeader(HEADER_ARRIVAL_REASON)
                valueToWrite = tripDepartureReason
        End Select
        
        ws.Cells(targetRow, col + RESULT_TABLE_COL_OFFSET).Value = valueToWrite
    Next col
End Sub

Private Sub private_WriteInvalidDateRow( _
    ByVal tbl As ListObject, _
    ByVal ws As Worksheet, _
    ByVal targetRow As Long, _
    ByVal awolRow As Range, _
    ByVal tripRow As Range, _
    ByVal awolArrivalDate As Variant, _
    ByVal tripDepartureDate As Variant)

    Dim col As Long
    Dim header As String
    Dim valueToWrite As Variant
    
    ws.Cells(targetRow, RESULT_COL_SOURCE_AWOL).Value = awolRow.Row
    ws.Cells(targetRow, RESULT_COL_SOURCE_BUSINESS_TRIP).Value = tripRow.Row
    ws.Cells(targetRow, RESULT_COL_GAP_DAYS).Value = VBA.vbNullString
    ws.Cells(targetRow, RESULT_COL_STATUS).Value = RESULT_STATUS_INVALID_DATES
    
    For col = 1 To tbl.ListColumns.Count
        header = VBA.CStr(tbl.HeaderRowRange.Cells(1, col).Value)
        valueToWrite = tripRow.Cells(1, col).Value
        
        If private_NormalizeHeader(header) = private_NormalizeHeader(HEADER_FULL_NAME) Then
            valueToWrite = private_FirstNonEmptyValue(tbl, awolRow, tripRow, HEADER_FULL_NAME)
        ElseIf private_NormalizeHeader(header) = private_NormalizeHeader(HEADER_TAX_ID) Then
            valueToWrite = private_FirstNonEmptyValue(tbl, awolRow, tripRow, HEADER_TAX_ID)
        ElseIf private_NormalizeHeader(header) = private_NormalizeHeader(HEADER_EVENT) Then
            valueToWrite = RESULT_STATUS_INVALID_DATES
        ElseIf private_NormalizeHeader(header) = private_NormalizeHeader(HEADER_ARRIVAL) Then
            valueToWrite = awolArrivalDate
        ElseIf private_NormalizeHeader(header) = private_NormalizeHeader(HEADER_DEPARTURE) Then
            valueToWrite = tripDepartureDate
        End If
        
        ws.Cells(targetRow, col + RESULT_TABLE_COL_OFFSET).Value = valueToWrite
    Next col
End Sub

Private Sub private_FormatResultSheet(ByVal ws As Worksheet, ByVal tbl As ListObject, ByVal lastRow As Long)
    Dim lastCol As Long
    Dim dateHeaders As Variant
    Dim header As Variant
    Dim colIndex As Long
    Dim rng As Range
    
    lastCol = tbl.ListColumns.Count + RESULT_TABLE_COL_OFFSET
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    private_ApplyDarkRangeStyle rng
    private_ApplyDarkHeaderStyle ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
    
    rng.AutoFilter
    ws.Columns.AutoFit
    
    dateHeaders = VBA.Array(HEADER_FROM_PROVISION, HEADER_DEPARTURE, HEADER_DEPARTURE_TERM, HEADER_PLANNED_ARRIVAL, HEADER_ARRIVAL, HEADER_TO_PROVISION)
    
    For Each header In dateHeaders
        colIndex = private_FindResultColumn(ws, VBA.CStr(header))
        If colIndex > 0 Then
            ws.Range(ws.Cells(2, colIndex), ws.Cells(lastRow, colIndex)).NumberFormat = "dd.mm.yyyy"
        End If
    Next header
    
    private_HighlightInvalidDateRows ws, lastRow, lastCol
    private_FreezeTopRow ws
End Sub

Private Sub private_FormatEmptyResultSheet(ByVal ws As Worksheet, ByVal tbl As ListObject)
    Dim lastCol As Long
    
    lastCol = tbl.ListColumns.Count + RESULT_TABLE_COL_OFFSET
    
    private_ApplyDarkRangeStyle ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
    private_ApplyDarkHeaderStyle ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
    
    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).AutoFilter
    ws.Columns.AutoFit
    
    private_FreezeTopRow ws
End Sub

Private Sub private_ApplyDarkBaseStyle(ByVal ws As Worksheet)
    With ws.Cells
        .Interior.Color = VBA.RGB(31, 31, 31)
        .Font.Color = VBA.RGB(255, 255, 255)
        .Font.Name = "Calibri"
        .Font.Size = 11
    End With
End Sub

Private Sub private_ApplyDarkRangeStyle(ByVal rng As Range)
    With rng
        .Interior.Color = VBA.RGB(31, 31, 31)
        .Font.Color = VBA.RGB(255, 255, 255)
        .Borders.LineStyle = Excel.xlContinuous
        .Borders.Color = VBA.RGB(70, 70, 70)
        .Borders.Weight = Excel.xlThin
    End With
End Sub

Private Sub private_ApplyDarkHeaderStyle(ByVal rng As Range)
    With rng
        .Interior.Color = VBA.RGB(0, 0, 0)
        .Font.Color = VBA.RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = Excel.xlCenter
        .VerticalAlignment = Excel.xlCenter
    End With
End Sub

Private Sub private_FreezeTopRow(ByVal ws As Worksheet)
    ws.Activate
    Application.ActiveWindow.FreezePanes = False
    Application.ActiveWindow.SplitRow = 1
    Application.ActiveWindow.FreezePanes = True
End Sub

Private Sub private_HighlightInvalidDateRows(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    Dim rowIndex As Long
    
    For rowIndex = 2 To lastRow
        If ws.Cells(rowIndex, RESULT_COL_STATUS).Value = RESULT_STATUS_INVALID_DATES Then
            With ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, lastCol))
                .Interior.Color = VBA.RGB(120, 0, 0)
                .Font.Color = VBA.RGB(255, 255, 255)
            End With
        End If
    Next rowIndex
End Sub

Private Function private_IsEventRow(ByVal tbl As ListObject, ByVal oneRow As Range, ByVal eventName As String) As Boolean
    private_IsEventRow = _
        private_NormalizeText(VBA.CStr(private_GetCellValue(tbl, oneRow, HEADER_EVENT))) = private_NormalizeText(eventName)
End Function

Private Function private_FirstNonEmptyValue(ByVal tbl As ListObject, ByVal firstRow As Range, ByVal secondRow As Range, ByVal header As String) As Variant
    Dim firstValue As Variant
    Dim secondValue As Variant
    
    firstValue = private_GetCellValue(tbl, firstRow, header)
    secondValue = private_GetCellValue(tbl, secondRow, header)
    
    If VBA.Len(private_NormalizeText(VBA.CStr(firstValue))) > 0 Then
        private_FirstNonEmptyValue = firstValue
    Else
        private_FirstNonEmptyValue = secondValue
    End If
End Function

Private Function private_GetCellValue(ByVal tbl As ListObject, ByVal oneRow As Range, ByVal header As String) As Variant
    Dim colIndex As Long
    
    colIndex = private_GetColumnIndex(tbl, header)
    
    If colIndex > 0 Then
        private_GetCellValue = oneRow.Cells(1, colIndex).Value
    Else
        private_GetCellValue = Empty
    End If
End Function

Private Function private_GetColumnIndex(ByVal tbl As ListObject, ByVal header As String) As Long
    Dim col As Long
    
    For col = 1 To tbl.ListColumns.Count
        If private_NormalizeHeader(VBA.CStr(tbl.HeaderRowRange.Cells(1, col).Value)) = private_NormalizeHeader(header) Then
            private_GetColumnIndex = col
            Exit Function
        End If
    Next col
End Function

Private Function private_FindResultColumn(ByVal ws As Worksheet, ByVal header As String) As Long
    Dim found As Range
    
    Set found = ws.Rows(1).Find(What:=header, LookIn:=Excel.xlValues, LookAt:=Excel.xlWhole, MatchCase:=False)
    
    If Not found Is Nothing Then
        private_FindResultColumn = found.Column
    End If
End Function

Private Function private_NormalizeHeader(ByVal text As String) As String
    private_NormalizeHeader = private_NormalizeText(VBA.Replace(text, VBA.Chr$(10), " "))
End Function

Private Function private_NormalizeText(ByVal text As String) As String
    text = VBA.Replace(text, VBA.vbCr, " ")
    text = VBA.Replace(text, VBA.vbLf, " ")
    text = VBA.Replace(text, VBA.Chr$(160), " ")
    text = VBA.Trim(text)
    
    Do While VBA.InStr(text, "  ") > 0
        text = VBA.Replace(text, "  ", " ")
    Loop
    
    private_NormalizeText = VBA.LCase$(text)
End Function
