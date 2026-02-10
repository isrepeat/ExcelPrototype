Attribute VB_Name = "ex_PersonTimeline"
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

    ' 4) Render content
    wsOut.Activate
    ActiveWindow.Zoom = 115

    Dim rowIndex As Long
    rowIndex = 1

    Dim mode As OutputMode
    mode = ex_Settings.m_GetOutputMode()

    mp_Log "Timeline", "Mode=" & ex_Settings.m_GetOutputModeDisplay()

    rowIndex = mp_WriteHeader(wsOut, fio, rowIndex, mode)
    rowIndex = rowIndex + 1

    Select Case mode
        Case PersonTimeline
            rowIndex = mp_WriteStateCard_FromSheet(wsOut, wsState, fio, rowIndex + 1)
            rowIndex = mp_WriteEventsTimeline_FromSheet(wsOut, wsEvents, fio, rowIndex + 2)
        Case StateTableOnly
            rowIndex = mp_WriteStateCard_FromSheet(wsOut, wsState, fio, rowIndex + 1)
        Case EventsTableOnly
            rowIndex = mp_WriteEventsTimeline_FromSheet(wsOut, wsEvents, fio, rowIndex + 1)
        Case Else
            rowIndex = mp_WriteStateCard_FromSheet(wsOut, wsState, fio, rowIndex + 1)
            rowIndex = mp_WriteEventsTimeline_FromSheet(wsOut, wsEvents, fio, rowIndex + 2)
    End Select

    wsOut.Columns.AutoFit
    ex_SheetTheme.m_ApplyDarkThemeToSheet wsOut
    mp_Log "Timeline", "Done"
    Exit Sub

EH:
    mp_Log "ERROR", "Err=" & CStr(Err.Number) & " Desc='" & Err.Description & "'"
    MsgBox "Error: " & Err.Description, vbExclamation, "m_ShowPersonTimeline"

End Sub

Private Function mp_WriteHeader(ByVal ws As Worksheet, ByVal fio As String, ByVal rowIndex As Long, ByVal mode As OutputMode) As Long

    Dim title As String
    Select Case mode
        Case PersonTimeline
            title = "Timeline by Full Name"
        Case StateTableOnly
            title = "State by Full Name"
        Case EventsTableOnly
            title = "Events by Full Name"
        Case Else
            title = "Timeline by Full Name"
    End Select

    ws.Cells(rowIndex, 1).Value = title
    ws.Cells(rowIndex, 2).Value = fio

    ws.Cells(rowIndex, 1).Font.Bold = True
    ws.Cells(rowIndex, 2).Font.Bold = True

    mp_WriteHeader = rowIndex

End Function

Private Function mp_WriteStateCard_FromSheet(ByVal wsOut As Worksheet, ByVal wsState As Worksheet, ByVal fio As String, ByVal rowIndex As Long) As Long

    Dim stateLayout As Variant
    stateLayout = mp_GetFieldIdList("Model.State.Fields")

    If mp_IsEmptyVariantArray(stateLayout) Then
        Err.Raise vbObjectError + 1000, "ex_PersonTimeline", "Model.State.Fields is empty"
    End If

    Dim keyFieldId As String
    keyFieldId = Trim$(ex_Config.m_GetConfigValue("Model.State.Key", "state_FIO"))

    Dim keyHeaderName As String
    keyHeaderName = mp_GetMappedSourceHeader(keyFieldId)

    If Len(keyHeaderName) = 0 Then
        Err.Raise vbObjectError + 1001, "ex_PersonTimeline", "Map is missing for Model.State.Key: '" & keyFieldId & "'"
    End If

    Dim keyColIndex As Long
    keyColIndex = mp_FindHeaderColumn(wsState, 1, keyHeaderName)

    If keyColIndex <= 0 Then
        Err.Raise vbObjectError + 1002, "ex_PersonTimeline", "State key header not found: '" & keyHeaderName & "'"
    End If

    Dim foundRow As Long
    foundRow = mp_FindRowByKey(wsState, keyColIndex, fio, 2)

    If foundRow <= 0 Then
        Err.Raise vbObjectError + 1003, "ex_PersonTimeline", "State row not found for fio='" & fio & "'"
    End If

    Dim i As Long
    For i = LBound(stateLayout) To UBound(stateLayout)

        Dim fieldId As String
        fieldId = Trim$(CStr(stateLayout(i)))

        If Len(fieldId) = 0 Then
            GoTo ContinueLoop
        End If

        Dim fieldLabel As String
        fieldLabel = mp_GetLabel(fieldId)

        Dim colIndex As Long
        colIndex = mp_TryGetColumnByFieldId(wsState, 1, fieldId)

        wsOut.Cells(rowIndex, 1).Value = fieldLabel

        If colIndex > 0 Then
            wsOut.Cells(rowIndex, 2).Value = wsState.Cells(foundRow, colIndex).Value
        Else
            wsOut.Cells(rowIndex, 2).Value = "(missing column)"
        End If

        rowIndex = rowIndex + 1

ContinueLoop:
    Next i

    mp_WriteStateCard_FromSheet = rowIndex

End Function

Private Function mp_WriteEventsTimeline_FromSheet(ByVal wsOut As Worksheet, ByVal wsEvents As Worksheet, ByVal fio As String, ByVal rowIndex As Long) As Long

    Dim eventsLayout As Variant
    eventsLayout = mp_GetFieldIdList("Model.Events.Fields")

    If mp_IsEmptyVariantArray(eventsLayout) Then
        Err.Raise vbObjectError + 1100, "ex_PersonTimeline", "Model.Events.Fields is empty"
    End If

    Dim keyFieldId As String
    keyFieldId = Trim$(ex_Config.m_GetConfigValue("Model.Events.Key", "events_FIO"))

    Dim keyHeaderName As String
    keyHeaderName = mp_GetMappedSourceHeader(keyFieldId)

    If Len(keyHeaderName) = 0 Then
        Err.Raise vbObjectError + 1101, "ex_PersonTimeline", "Map is missing for Model.Events.Key: '" & keyFieldId & "'"
    End If

    Dim keyColIndex As Long
    keyColIndex = mp_FindHeaderColumn(wsEvents, 1, keyHeaderName)

    If keyColIndex <= 0 Then
        Err.Raise vbObjectError + 1102, "ex_PersonTimeline", "Events key header not found: '" & keyHeaderName & "'"
    End If

    Dim outHeaderRow As Long
    outHeaderRow = rowIndex

    Dim i As Long
    For i = LBound(eventsLayout) To UBound(eventsLayout)

        Dim fieldId As String
        fieldId = Trim$(CStr(eventsLayout(i)))

        wsOut.Cells(outHeaderRow, 1 + (i - LBound(eventsLayout))).Value = mp_GetLabel(fieldId)

    Next i

    Dim outDataRow As Long
    outDataRow = outHeaderRow + 1

    Dim lastRow As Long
    lastRow = wsEvents.Cells(wsEvents.Rows.Count, keyColIndex).End(xlUp).row

    Dim r As Long
    For r = 2 To lastRow

        Dim rowFio As String
        rowFio = CStr(wsEvents.Cells(r, keyColIndex).Value)

        If StrComp(Trim$(rowFio), fio, vbTextCompare) <> 0 Then
            GoTo ContinueRow
        End If

        For i = LBound(eventsLayout) To UBound(eventsLayout)

            Dim colIndex As Long
            colIndex = mp_TryGetColumnByFieldId(wsEvents, 1, Trim$(CStr(eventsLayout(i))))

            If colIndex > 0 Then
                wsOut.Cells(outDataRow, 1 + (i - LBound(eventsLayout))).Value = wsEvents.Cells(r, colIndex).Value
            Else
                wsOut.Cells(outDataRow, 1 + (i - LBound(eventsLayout))).Value = "(missing column)"
            End If

        Next i

        outDataRow = outDataRow + 1

ContinueRow:
    Next r

    If outDataRow = outHeaderRow + 1 Then
        wsOut.Cells(outDataRow, 1).Value = "(no events found for this person)"
        mp_WriteEventsTimeline_FromSheet = outDataRow + 1
        Exit Function
    End If

    Dim sortFieldId As String
    sortFieldId = Trim$(ex_Config.m_GetConfigValue("Model.Events.Sort", vbNullString))

    If Len(sortFieldId) > 0 Then

        Dim sortOutCol As Long
        sortOutCol = -1

        For i = LBound(eventsLayout) To UBound(eventsLayout)
            If StrComp(Trim$(CStr(eventsLayout(i))), sortFieldId, vbTextCompare) = 0 Then
                sortOutCol = 1 + (i - LBound(eventsLayout))
                Exit For
            End If
        Next i

        If sortOutCol > 0 Then
            mp_SortRangeByColumnIndex wsOut, outHeaderRow, outDataRow - 1, 1, (UBound(eventsLayout) - LBound(eventsLayout) + 1), sortOutCol
        Else
            mp_Log "Events", "Sort ignored: '" & sortFieldId & "' is not in Model.Events.Fields"
        End If

    End If

    mp_WriteEventsTimeline_FromSheet = outDataRow + 1

End Function

Private Function mp_CreateOrClearSheet(ByVal sheetName As String) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    ex_SheetTheme.m_ApplyDarkThemeToSheet ws
    ws.Cells.NumberFormat = "@"

    Set mp_CreateOrClearSheet = ws

End Function

Private Function mp_NormalizeHeader(ByVal s As String) As String

    mp_NormalizeHeader = LCase$(Trim$(s))

End Function

Private Function mp_FindHeaderColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerName As String) As Long

    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    Dim normalizedNeedle As String
    normalizedNeedle = mp_NormalizeHeader(headerName)

    Dim c As Long
    For c = 1 To lastCol
        If mp_NormalizeHeader(CStr(ws.Cells(headerRow, c).Value)) = normalizedNeedle Then
            mp_FindHeaderColumn = c
            Exit Function
        End If
    Next c

    mp_FindHeaderColumn = -1

End Function

Private Function mp_IsEmptyVariantArray(ByVal v As Variant) As Boolean

    On Error GoTo EH

    If IsArray(v) = False Then
        mp_IsEmptyVariantArray = True
        Exit Function
    End If

    If UBound(v) < LBound(v) Then
        mp_IsEmptyVariantArray = True
        Exit Function
    End If

    If (UBound(v) = LBound(v)) Then
        If Len(Trim$(CStr(v(LBound(v))))) = 0 Then
            mp_IsEmptyVariantArray = True
            Exit Function
        End If
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

    raw = Replace$(raw, ",", ";")

    Dim parts As Variant
    parts = Split(raw, ";")

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        parts(i) = Trim$(CStr(parts(i)))
    Next i

    mp_GetFieldIdList = parts

End Function

Private Function mp_GetMappedSourceHeader(ByVal fieldId As String) As String

    Dim k As String
    k = "Map." & fieldId

    Dim raw As String
    raw = Trim$(ex_Config.m_GetConfigValue(k, vbNullString))

    If Len(raw) = 0 Then
        mp_GetMappedSourceHeader = vbNullString
        Exit Function
    End If

    Dim p As Long
    p = InStr(1, raw, "|", vbBinaryCompare)

    If p > 0 Then
        mp_GetMappedSourceHeader = Trim$(Left$(raw, p - 1))
    Else
        mp_GetMappedSourceHeader = raw
    End If

End Function

Private Function mp_GetLabel(ByVal fieldId As String) As String

    Dim k As String
    k = "Map." & fieldId

    Dim raw As String
    raw = Trim$(ex_Config.m_GetConfigValue(k, vbNullString))

    If Len(raw) > 0 Then

        Dim p As Long
        p = InStr(1, raw, "|", vbBinaryCompare)

        If p > 0 Then

            Dim lbl As String
            lbl = Trim$(Mid$(raw, p + 1))

            If Len(lbl) > 0 Then
                mp_GetLabel = lbl
                Exit Function
            End If

        End If

    End If

    ' Fallback: use the field id itself (prefer the suffix after "state_" / "events_")
    Dim u As Long
    u = InStr(1, fieldId, "_", vbBinaryCompare)

    If u > 0 Then
        mp_GetLabel = Mid$(fieldId, u + 1)
    Else
        mp_GetLabel = fieldId
    End If

End Function

Private Function mp_TryGetColumnByFieldId(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal fieldId As String) As Long

    Dim headerName As String
    headerName = mp_GetMappedSourceHeader(fieldId)

    If Len(headerName) = 0 Then
        mp_TryGetColumnByFieldId = -1
        Exit Function
    End If

    mp_TryGetColumnByFieldId = mp_FindHeaderColumn(ws, headerRow, headerName)

End Function

Private Sub mp_SortRangeByColumnIndex(ByVal ws As Worksheet, ByVal topRow As Long, ByVal bottomRow As Long, ByVal leftCol As Long, ByVal rightCol As Long, ByVal sortColRelative As Long)

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(topRow, leftCol), ws.Cells(bottomRow, rightCol))

    rng.Sort Key1:=ws.Cells(topRow + 1, leftCol + sortColRelative - 1), Order1:=xlAscending, Header:=xlYes

End Sub

Private Function mp_FindRowByKey(ByVal ws As Worksheet, ByVal keyColIndex As Long, ByVal keyValue As String, ByVal startRow As Long) As Long

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, keyColIndex).End(xlUp).row

    Dim r As Long
    For r = startRow To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, keyColIndex).Value)), keyValue, vbTextCompare) = 0 Then
            mp_FindRowByKey = r
            Exit Function
        End If
    Next r

    mp_FindRowByKey = -1

End Function

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

    ' Apply dark theme to the log sheet
    ex_SheetTheme.m_ApplyDarkThemeToSheet ws

End Sub

Private Sub mp_Log(ByVal category As String, ByVal message As String)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("g_Log")

    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1

    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = category
    ws.Cells(r, 3).Value = message

End Sub

Private Function mp_DebugHeadersRow(ByVal ws As Worksheet, ByVal headerRow As Long) As String

    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    Dim s As String
    s = ""

    Dim c As Long
    For c = 1 To lastCol
        If Len(s) > 0 Then
            s = s & "; "
        End If
        s = s & CStr(ws.Cells(headerRow, c).Value)
    Next c

    mp_DebugHeadersRow = s

End Function
