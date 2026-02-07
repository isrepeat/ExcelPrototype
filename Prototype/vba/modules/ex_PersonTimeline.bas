Option Explicit

Public Sub m_ShowPersonTimeline_UI()

    Dim fio As String
    ' Ввод ФИО через UI закомментирован — читаем из конфига.
    ' fio = InputBox("Enter Full Name (exact match):", "Timeline by Full Name")

    fio = Trim$(ex_Config.m_GetConfigValue("PersonFIO", vbNullString))
    If fio = vbNullString Then
        ' При отсутствии в конфиге можно раскомментировать fallback ниже:
        ' fio = Trim$(InputBox("Enter Full Name (exact match):", "Timeline by Full Name"))
        Exit Sub
    End If

    m_ShowPersonTimeline fio

End Sub

Public Sub m_ShowPersonTimeline(ByVal fio As String)

    On Error GoTo EH

    ' 1) Load external tables into internal sheets based on dev config
    ex_SourceLoader.m_LoadStateEventsFromConfigToInternalSheets

    ' 2) Get internal sheets
    Dim wsState As Worksheet
    Dim wsEvents As Worksheet

    Set wsState = ThisWorkbook.Worksheets("g_State")
    Set wsEvents = ThisWorkbook.Worksheets("g_Events")

    ' 3) Create output sheet
    Dim wsOut As Worksheet
    Set wsOut = mp_CreateOrClearSheet("g_PersonTimeline")

    wsOut.Activate
    ActiveWindow.Zoom = 115

    Dim rowIndex As Long
    rowIndex = 1

    rowIndex = mp_WriteHeader(wsOut, fio, rowIndex)
    rowIndex = mp_WriteStateCard_FromSheet(wsOut, wsState, fio, rowIndex + 1)
    rowIndex = mp_WriteEventsTimeline_FromSheet(wsOut, wsEvents, fio, rowIndex + 2)

    wsOut.Columns.AutoFit
    Exit Sub

EH:
    MsgBox "Error: " & Err.Description, vbExclamation, "m_ShowPersonTimeline"

End Sub

' ========================================================
' Output Functions
' ========================================================

Private Function mp_WriteHeader(ByVal ws As Worksheet, ByVal fio As String, ByVal rowIndex As Long) As Long

    ws.Cells(rowIndex, 1).Value = "Timeline by Full Name"
    ws.Cells(rowIndex, 2).Value = fio

    ws.Cells(rowIndex, 1).Font.Bold = True
    ws.Cells(rowIndex, 2).Font.Bold = True

    mp_WriteHeader = rowIndex

End Function

Private Function mp_WriteStateCard_FromSheet(ByVal wsOut As Worksheet, ByVal wsState As Worksheet, ByVal fio As String, ByVal rowIndex As Long) As Long

    Dim colFio As Long
    Dim colBirth As Long
    Dim colCity As Long
    Dim colPhone As Long

    colFio = mp_FindHeaderColumn(wsState, 1, "FIO")
    colBirth = mp_FindHeaderColumn(wsState, 1, "BirthDate")
    colCity = mp_FindHeaderColumn(wsState, 1, "City")
    colPhone = mp_FindHeaderColumn(wsState, 1, "Phone")

    If colFio = 0 Then
        Err.Raise vbObjectError + 601, "ex_PersonTimeline", "g_State: column 'FIO' not found"
    End If

    wsOut.Cells(rowIndex, 1).Value = "State"
    wsOut.Cells(rowIndex, 1).Font.Bold = True

    Dim foundRow As Long
    foundRow = mp_FindRowByKey(wsState, colFio, fio)

    Dim r As Long
    r = rowIndex + 1

    wsOut.Cells(r, 1).Value = "FIO"
    wsOut.Cells(r, 2).Value = fio

    wsOut.Cells(r + 1, 1).Value = "BirthDate"
    wsOut.Cells(r + 2, 1).Value = "City"
    wsOut.Cells(r + 3, 1).Value = "Phone"

    wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r + 3, 1)).Font.Bold = True

    If foundRow = 0 Then
        wsOut.Cells(r + 1, 2).Value = "(not found in TableState)"
        mp_WriteStateCard_FromSheet = r + 3
        Exit Function
    End If

    If colBirth > 0 Then
        wsOut.Cells(r + 1, 2).Value = "'" & CStr(wsState.Cells(foundRow, colBirth).Value)
    End If

    If colCity > 0 Then
        wsOut.Cells(r + 2, 2).Value = "'" & CStr(wsState.Cells(foundRow, colCity).Value)
    End If

    If colPhone > 0 Then
        wsOut.Cells(r + 3, 2).Value = "'" & CStr(wsState.Cells(foundRow, colPhone).Value)
    End If

    mp_WriteStateCard_FromSheet = r + 3

End Function

Private Function mp_WriteEventsTimeline_FromSheet(ByVal wsOut As Worksheet, ByVal wsEvents As Worksheet, ByVal fio As String, ByVal rowIndex As Long) As Long

    Dim colFio As Long
    Dim colDate As Long
    Dim colType As Long
    Dim colDept As Long
    Dim colPos As Long
    Dim colSalary As Long
    Dim colRecordNo As Long

    colFio = mp_FindHeaderColumn(wsEvents, 1, "FIO")
    colDate = mp_FindHeaderColumn(wsEvents, 1, "EventDate")
    colType = mp_FindHeaderColumn(wsEvents, 1, "EventType")
    colDept = mp_FindHeaderColumn(wsEvents, 1, "Department")
    colPos = mp_FindHeaderColumn(wsEvents, 1, "Position")
    colSalary = mp_FindHeaderColumn(wsEvents, 1, "Salary")
    colRecordNo = mp_FindHeaderColumn(wsEvents, 1, "RecordNo")

    If colFio = 0 Then
        Err.Raise vbObjectError + 602, "ex_PersonTimeline", "g_Events: column 'FIO' not found"
    End If

    wsOut.Cells(rowIndex, 1).Value = "Events (Timeline)"
    wsOut.Cells(rowIndex, 1).Font.Bold = True

    Dim outTop As Long
    outTop = rowIndex + 1

    wsOut.Cells(outTop, 1).Value = "RecordNo"
    wsOut.Cells(outTop, 2).Value = "EventDate"
    wsOut.Cells(outTop, 3).Value = "EventType"
    wsOut.Cells(outTop, 4).Value = "Department"
    wsOut.Cells(outTop, 5).Value = "Position"
    wsOut.Cells(outTop, 6).Value = "Salary"

    wsOut.Range(wsOut.Cells(outTop, 1), wsOut.Cells(outTop, 6)).Font.Bold = True

    Dim lastRow As Long
    lastRow = wsEvents.Cells(wsEvents.Rows.Count, colFio).End(xlUp).Row

    Dim outRow As Long
    outRow = outTop + 1

    Dim r As Long
    For r = 2 To lastRow

        If CStr(wsEvents.Cells(r, colFio).Value) = fio Then

            If colRecordNo > 0 Then
                wsOut.Cells(outRow, 1).Value = "'" & CStr(wsEvents.Cells(r, colRecordNo).Value)
            End If

            If colDate > 0 Then
                wsOut.Cells(outRow, 2).Value = "'" & CStr(wsEvents.Cells(r, colDate).Value)
            End If

            If colType > 0 Then
                wsOut.Cells(outRow, 3).Value = "'" & CStr(wsEvents.Cells(r, colType).Value)
            End If

            If colDept > 0 Then
                wsOut.Cells(outRow, 4).Value = "'" & CStr(wsEvents.Cells(r, colDept).Value)
            End If

            If colPos > 0 Then
                wsOut.Cells(outRow, 5).Value = "'" & CStr(wsEvents.Cells(r, colPos).Value)
            End If

            If colSalary > 0 Then
                wsOut.Cells(outRow, 6).Value = "'" & CStr(wsEvents.Cells(r, colSalary).Value)
            End If

            outRow = outRow + 1

        End If

    Next r

    If outRow = outTop + 1 Then
        wsOut.Cells(outTop + 1, 1).Value = "(no events found for this FIO)"
        mp_WriteEventsTimeline_FromSheet = outTop + 1
        Exit Function
    End If

    If colRecordNo > 0 Then
        mp_SortRangeByFirstColumn wsOut, wsOut.Range(wsOut.Cells(outTop, 1), wsOut.Cells(outRow - 1, 6))
    End If

    mp_WriteEventsTimeline_FromSheet = outRow - 1

End Function

Private Sub mp_SortRangeByFirstColumn(ByVal ws As Worksheet, ByVal rng As Range)

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=rng.Columns(1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange rng
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

End Sub

' ========================================================
' Helper Functions
' ========================================================

Private Function mp_CreateOrClearSheet(ByVal sheetName As String) As Worksheet

    Dim ws As Worksheet

    Dim fullName As String

    ' Ensure sheet names created by code use g_ prefix
    If Left$(sheetName, 2) = "g_" Then
        fullName = sheetName
    Else
        fullName = "g_" & sheetName
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(fullName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = fullName
        Call m_ApplyDefaultSheetView(ws)
    Else
        ws.Cells.Clear
    End If

    ' Apply dark theme so timeline matches g_State appearance
    ex_SheetTheme.m_ApplyDarkThemeToSheet ws

    ' Format all cells as text
    ws.Cells.NumberFormat = "@"

    Set mp_CreateOrClearSheet = ws

End Function

Private Function mp_BuildTimelineSheetName(ByVal fio As String) As String

    Dim safeName As String
    Dim prefix As String
    Dim maxLen As Long
    Dim i As Long
    Dim char As String

    prefix = "Timeline_"
    maxLen = 31 - Len(prefix)  ' Excel limit: 31 chars for sheet names

    safeName = fio

    ' Заменить все невалидные символы Excel на подчёркивание
    safeName = Replace$(safeName, ":", "_")
    safeName = Replace$(safeName, "\", "_")
    safeName = Replace$(safeName, "/", "_")
    safeName = Replace$(safeName, "?", "_")
    safeName = Replace$(safeName, "*", "_")
    safeName = Replace$(safeName, "[", "_")
    safeName = Replace$(safeName, "]", "_")
    safeName = Replace$(safeName, "'", "_")  ' одинарная кавычка
    safeName = Replace$(safeName, """", "_")  ' двойная кавычка
    safeName = Replace$(safeName, ";", "_")

    ' Удалить повторяющиеся подчёркивания
    Do While InStr(safeName, "__") > 0
        safeName = Replace$(safeName, "__", "_")
    Loop

    safeName = Trim$(safeName)
    ' Удалить ведущие и хвостовые подчёркивания
    While Left$(safeName, 1) = "_"
        safeName = Mid$(safeName, 2)
    Wend
    While Right$(safeName, 1) = "_"
        safeName = Left$(safeName, Len(safeName) - 1)
    Wend

    safeName = Trim$(safeName)
    If Len(safeName) = 0 Then
        safeName = "Timeline"
    ElseIf Len(safeName) > maxLen Then
        safeName = Left$(safeName, maxLen)
    End If

    mp_BuildTimelineSheetName = prefix & safeName

End Function

Private Function mp_FindHeaderColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerName As String) As Long

    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To lastCol
        If StrComp(CStr(ws.Cells(headerRow, c).Value), headerName, vbTextCompare) = 0 Then
            mp_FindHeaderColumn = c
            Exit Function
        End If
    Next c

    mp_FindHeaderColumn = 0

End Function

Private Function mp_FindRowByKey(ByVal ws As Worksheet, ByVal keyCol As Long, ByVal keyValue As String) As Long

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        If CStr(ws.Cells(r, keyCol).Value) = keyValue Then
            mp_FindRowByKey = r
            Exit Function
        End If
    Next r

    mp_FindRowByKey = 0

End Function
