Option Explicit

Public Sub m_ShowPersonTimeline_UI()

    Dim fio As String

    fio = Trim$(ex_Config.m_GetConfigValue("PersonFIO", vbNullString))
    If fio = vbNullString Then
        Exit Sub
    End If

    m_ShowPersonTimeline fio

End Sub

Public Sub m_ShowPersonTimeline(ByVal fio As String)

    On Error GoTo EH

    mp_LogInit
    mp_Log "Timeline", "Start: fio='" & fio & "'"

    ' 1) Load external tables into internal sheets based on dev config
    ex_SourceLoader.m_LoadStateEventsFromConfigToInternalSheets

    ' 2) Get internal sheets
    Dim wsState As Worksheet
    Dim wsEvents As Worksheet

    Set wsState = ThisWorkbook.Worksheets("g_State")
    Set wsEvents = ThisWorkbook.Worksheets("g_Events")

    mp_Log "Timeline", "Sheets: g_State='" & wsState.Name & "', g_Events='" & wsEvents.Name & "'"
    mp_Log "Timeline", "g_Events UsedRange=" & wsEvents.UsedRange.Address
    mp_Log "Timeline", "g_Events Header row 1: " & mp_DebugHeadersRow(wsEvents, 1)

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
    mp_Log "Timeline", "Done"
    Exit Sub

EH:
    mp_Log "ERROR", "Err=" & CStr(Err.Number) & " Desc='" & Err.Description & "'"
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

    Dim stateLayout As Variant
    stateLayout = mp_GetFieldIdList("Layout.State")

    If mp_IsEmptyVariantArray(stateLayout) Then
        Err.Raise vbObjectError + 601, "ex_PersonTimeline", "Layout.State is empty"
    End If

    Dim keyFieldId As String
    keyFieldId = Trim$(ex_Config.m_GetConfigValue("KeyField.State", "State.FIO"))

    Dim mapKey As String
    mapKey = "Map." & keyFieldId

    Dim headerName As String
    headerName = Trim$(ex_Config.m_GetConfigValue(mapKey, vbNullString))

    mp_Log "State", "KeyField.State='" & keyFieldId & "'"
    mp_Log "State", "Config '" & mapKey & "'='" & headerName & "' (len=" & Len(headerName) & ")"
    mp_Log "State", "Header row 1: " & mp_DebugHeadersRow(wsState, 1)

    Dim colKey As Long
    colKey = mp_TryGetColumnByFieldId(wsState, 1, keyFieldId)

    mp_Log "State", "mp_TryGetColumnByFieldId(row=1, fieldId='" & keyFieldId & "') => " & CStr(colKey)

    If colKey = 0 Then
        Err.Raise vbObjectError + 602, "ex_PersonTimeline", "g_State: key column not found or not mapped for " & keyFieldId
    End If

    wsOut.Cells(rowIndex, 1).Value = "State"
    wsOut.Cells(rowIndex, 1).Font.Bold = True

    Dim foundRow As Long
    foundRow = mp_FindRowByKey(wsState, colKey, fio)

    mp_Log "State", "FindRowByKey => " & CStr(foundRow)

    Dim outRow As Long
    outRow = rowIndex + 1

    Dim i As Long
    For i = LBound(stateLayout) To UBound(stateLayout)

        Dim fieldId As String
        fieldId = CStr(stateLayout(i))

        Dim labelText As String
        labelText = mp_GetLabel(fieldId)

        wsOut.Cells(outRow, 1).Value = labelText
        wsOut.Cells(outRow, 1).Font.Bold = True

        If StrComp(fieldId, keyFieldId, vbTextCompare) = 0 Then
            wsOut.Cells(outRow, 2).Value = "'" & fio
            outRow = outRow + 1
            GoTo ContinueLoop
        End If

        If foundRow = 0 Then
            wsOut.Cells(outRow, 2).Value = "(not found in TableState)"
            outRow = outRow + 1
            GoTo ContinueLoop
        End If

        Dim colVal As Long
        colVal = mp_TryGetColumnByFieldId(wsState, 1, fieldId)

        If colVal = 0 Then
            wsOut.Cells(outRow, 2).Value = "(column not mapped)"
        Else
            wsOut.Cells(outRow, 2).Value = "'" & CStr(wsState.Cells(foundRow, colVal).Value)
        End If

        outRow = outRow + 1

ContinueLoop:
    Next i

    mp_WriteStateCard_FromSheet = outRow - 1

End Function

Private Function mp_WriteEventsTimeline_FromSheet(ByVal wsOut As Worksheet, ByVal wsEvents As Worksheet, ByVal fio As String, ByVal rowIndex As Long) As Long

    Dim eventsLayout As Variant
    eventsLayout = mp_GetFieldIdList("Layout.Events")

    If mp_IsEmptyVariantArray(eventsLayout) Then
        Err.Raise vbObjectError + 610, "ex_PersonTimeline", "Layout.Events is empty"
    End If

    mp_Log "Events", "Layout.Events count=" & CStr(UBound(eventsLayout) - LBound(eventsLayout) + 1)

    Dim keyFieldId As String
    keyFieldId = Trim$(ex_Config.m_GetConfigValue("KeyField.Events", "Events.FIO"))

    mp_Log "Events", "KeyField.Events='" & keyFieldId & "'"

    Dim mapKey As String
    mapKey = "Map." & keyFieldId

    Dim headerName As String
    headerName = Trim$(ex_Config.m_GetConfigValue(mapKey, vbNullString))

    mp_Log "Events", "Config '" & mapKey & "'='" & headerName & "' (len=" & Len(headerName) & ")"
    mp_Log "Events", "Header row 1: " & mp_DebugHeadersRow(wsEvents, 1)

    Dim colKey As Long
    colKey = mp_TryGetColumnByFieldId(wsEvents, 1, keyFieldId)

    mp_Log "Events", "mp_TryGetColumnByFieldId(row=1, fieldId='" & keyFieldId & "') => " & CStr(colKey)

    If colKey = 0 Then
        Err.Raise vbObjectError + 611, "ex_PersonTimeline", "g_Events: key column not found or not mapped for " & keyFieldId
    End If

    wsOut.Cells(rowIndex, 1).Value = "Events (Timeline)"
    wsOut.Cells(rowIndex, 1).Font.Bold = True

    Dim outTop As Long
    outTop = rowIndex + 1

    ' Header row
    Dim i As Long
    For i = LBound(eventsLayout) To UBound(eventsLayout)
        wsOut.Cells(outTop, 1 + i).Value = mp_GetLabel(CStr(eventsLayout(i)))
        wsOut.Cells(outTop, 1 + i).Font.Bold = True
    Next i

    ' Output rows
    Dim lastRow As Long
    lastRow = wsEvents.Cells(wsEvents.Rows.Count, colKey).End(xlUp).Row

    Dim outRow As Long
    outRow = outTop + 1

    Dim r As Long
    For r = 2 To lastRow

        If CStr(wsEvents.Cells(r, colKey).Value) = fio Then

            For i = LBound(eventsLayout) To UBound(eventsLayout)

                Dim fieldId As String
                fieldId = CStr(eventsLayout(i))

                Dim colVal As Long
                colVal = mp_TryGetColumnByFieldId(wsEvents, 1, fieldId)

                If colVal > 0 Then
                    wsOut.Cells(outRow, 1 + i).Value = "'" & CStr(wsEvents.Cells(r, colVal).Value)
                Else
                    wsOut.Cells(outRow, 1 + i).Value = vbNullString
                End If

            Next i

            outRow = outRow + 1

        End If

    Next r

    If outRow = outTop + 1 Then
        wsOut.Cells(outTop + 1, 1).Value = "(no events found for this person)"
        mp_WriteEventsTimeline_FromSheet = outTop + 1
        Exit Function
    End If

    ' Optional sort (only if SortField.Events is present in layout)
    Dim sortFieldId As String
    sortFieldId = Trim$(ex_Config.m_GetConfigValue("SortField.Events", vbNullString))

    If Len(sortFieldId) > 0 Then

        Dim sortIndex As Long
        sortIndex = -1

        For i = LBound(eventsLayout) To UBound(eventsLayout)
            If StrComp(CStr(eventsLayout(i)), sortFieldId, vbTextCompare) = 0 Then
                sortIndex = i
                Exit For
            End If
        Next i

        If sortIndex >= 0 Then
            Dim sortRange As Range
            Set sortRange = wsOut.Range( _
                wsOut.Cells(outTop, 1), _
                wsOut.Cells(outRow - 1, 1 + UBound(eventsLayout)) _
            )

            mp_SortRangeByColumnIndex wsOut, sortRange, 1 + sortIndex
        End If

    End If

    mp_WriteEventsTimeline_FromSheet = outRow - 1

End Function

' ========================================================
' Helper Functions
' ========================================================

Private Function mp_CreateOrClearSheet(ByVal sheetName As String) As Worksheet

    Dim ws As Worksheet
    Dim fullName As String

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

    ex_SheetTheme.m_ApplyDarkThemeToSheet ws
    ws.Cells.NumberFormat = "@"

    Set mp_CreateOrClearSheet = ws

End Function

Private Function mp_NormalizeHeader(ByVal s As String) As String

    s = CStr(s)
    s = Replace$(s, ChrW(160), " ")
    s = Trim$(s)
    s = LCase$(s)

    mp_NormalizeHeader = s

End Function

Private Function mp_FindHeaderColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerName As String) As Long

    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To lastCol
        If mp_NormalizeHeader(ws.Cells(headerRow, c).Value) = mp_NormalizeHeader(headerName) Then
            mp_FindHeaderColumn = c
            Exit Function
        End If
    Next c

    mp_FindHeaderColumn = 0

End Function

Private Function mp_IsEmptyVariantArray(ByVal arr As Variant) As Boolean

    On Error GoTo EH

    If IsArray(arr) = False Then
        mp_IsEmptyVariantArray = True
        Exit Function
    End If

    If LBound(arr) > UBound(arr) Then
        mp_IsEmptyVariantArray = True
        Exit Function
    End If

    mp_IsEmptyVariantArray = False
    Exit Function

EH:
    mp_IsEmptyVariantArray = True

End Function

Private Function mp_GetFieldIdList(ByVal configKey As String) As Variant

    Dim raw As String
    raw = Trim$(ex_Config.m_GetConfigValue(configKey, vbNullString))

    If Len(raw) = 0 Then
        mp_GetFieldIdList = Array()
        Exit Function
    End If

    raw = Replace$(raw, vbCr, vbNullString)
    raw = Replace$(raw, vbLf, vbNullString)

    raw = Replace$(raw, ",", ";")

    Dim parts() As String
    parts = Split(raw, ";")

    Dim cleaned() As String
    Dim i As Long
    Dim n As Long
    n = 0

    ReDim cleaned(0 To UBound(parts))

    For i = LBound(parts) To UBound(parts)

        Dim item As String
        item = Trim$(parts(i))

        If Len(item) > 0 Then
            cleaned(n) = item
            n = n + 1
        End If

    Next i

    If n = 0 Then
        mp_GetFieldIdList = Array()
        Exit Function
    End If

    ReDim Preserve cleaned(0 To n - 1)
    mp_GetFieldIdList = cleaned

End Function

Private Function mp_GetMappedSourceHeader(ByVal fieldId As String) As String

    Dim k As String
    k = "Map." & fieldId

    mp_GetMappedSourceHeader = Trim$(ex_Config.m_GetConfigValue(k, vbNullString))

End Function

Private Function mp_GetLabel(ByVal fieldId As String) As String

    Dim k As String
    k = "Label." & fieldId

    Dim lbl As String
    lbl = Trim$(ex_Config.m_GetConfigValue(k, vbNullString))

    If Len(lbl) > 0 Then
        mp_GetLabel = lbl
        Exit Function
    End If

    Dim p As Long
    p = InStrRev(fieldId, ".")
    If p > 0 Then
        mp_GetLabel = Mid$(fieldId, p + 1)
    Else
        mp_GetLabel = fieldId
    End If

End Function

Private Function mp_TryGetColumnByFieldId(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal fieldId As String) As Long

    Dim headerName As String
    headerName = mp_GetMappedSourceHeader(fieldId)

    If Len(headerName) = 0 Then
        mp_TryGetColumnByFieldId = 0
        Exit Function
    End If

    mp_TryGetColumnByFieldId = mp_FindHeaderColumn(ws, headerRow, headerName)

End Function

Private Sub mp_SortRangeByColumnIndex(ByVal ws As Worksheet, ByVal rng As Range, ByVal columnIndex As Long)

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=rng.Columns(columnIndex), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange rng
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

End Sub

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

' ========================================================
' Logging (g_Log sheet)
' ========================================================

Private Sub mp_LogInit()

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("g_Log")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "g_Log"
    Else
        ws.Cells.Clear
    End If

    ws.Cells(1, 1).Value = "Time"
    ws.Cells(1, 2).Value = "Module"
    ws.Cells(1, 3).Value = "Message"
    ws.Rows(1).Font.Bold = True

    ws.Columns(1).ColumnWidth = 22
    ws.Columns(2).ColumnWidth = 18
    ws.Columns(3).ColumnWidth = 140

    ex_SheetTheme.m_ApplyDarkThemeToSheet ws
    ws.Cells.NumberFormat = "@"

End Sub

Private Sub mp_Log(ByVal moduleName As String, ByVal message As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("g_Log")

    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = moduleName
    ws.Cells(r, 3).Value = message

End Sub

Private Function mp_DebugHeadersRow(ByVal ws As Worksheet, ByVal headerRow As Long) As String

    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    Dim s As String

    For c = 1 To lastCol
        s = s & "[" & c & "]{" & Len(CStr(ws.Cells(headerRow, c).Value)) & "}='" & CStr(ws.Cells(headerRow, c).Value) & "' "
    Next c

    mp_DebugHeadersRow = s

End Function
