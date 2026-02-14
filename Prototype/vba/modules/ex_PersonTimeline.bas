Attribute VB_Name = "ex_PersonTimeline"
Option Explicit

Private Const DEV_CONFIG_TABLE_NAME As String = "tblDevConfig"
Private Const DEV_MARKER_SYMBOL As String = "#"
Private Const DEV_COL_MARKER As Long = 1
Private Const DEV_COL_KEY As Long = 2
Private Const DEV_COL_VALUE As Long = 3

Public Sub m_ShowPersonTimeline_UI()

    Dim fio As String

    fio = Trim$(ex_Config.m_GetConfigValue("Context.PersonValue", vbNullString))
    If Len(fio) = 0 Then
        fio = Trim$(ex_Config.m_GetConfigValue("PersonFIO", vbNullString))
    End If
    If Len(fio) = 0 Then
        MsgBox "Config key 'Context.PersonValue' is empty.", vbExclamation, "m_ShowPersonTimeline_UI"
        Exit Sub
    End If

    m_ShowPersonTimeline fio

End Sub

Public Sub m_ShowPersonTimeline(ByVal fio As String)

    On Error GoTo EH

    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevEnableEvents As Boolean

    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    prevEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Dim cfg As Object
    Set cfg = mp_LoadConfigDictionary()

    Dim mode As OutputMode
    mode = ex_Settings.m_GetOutputMode()

    Dim outputAliases As Variant
    outputAliases = mp_GetListRequired(cfg, "Output.Tables")

    Dim wbCache As Object
    Set wbCache = CreateObject("Scripting.Dictionary")
    wbCache.CompareMode = 1

    Dim wsOut As Worksheet
    Set wsOut = mp_CreateOrClearSheet("g_PersonTimeline")

    wsOut.Activate
    ActiveWindow.Zoom = 115

    Dim rowIndex As Long
    rowIndex = 1
    rowIndex = mp_WriteHeader(wsOut, fio, rowIndex, mode)
    rowIndex = rowIndex + 1

    Dim renderedCount As Long
    renderedCount = 0

    Dim i As Long
    For i = LBound(outputAliases) To UBound(outputAliases)
        Dim tableAlias As String
        tableAlias = Trim$(CStr(outputAliases(i)))
        If Len(tableAlias) = 0 Then
            GoTo ContinueAlias
        End If

        Dim sourceAlias As String
        sourceAlias = mp_FindSourceAliasForTable(cfg, tableAlias)

        Dim tableType As String
        tableType = LCase$(mp_GetCfgRequired(cfg, sourceAlias & ".Table[" & tableAlias & "].Type"))

        If mode = StateTableOnly And tableType <> "state" Then
            GoTo ContinueAlias
        End If
        If mode = EventsTableOnly And tableType <> "events" Then
            GoTo ContinueAlias
        End If

        If tableType <> "state" And tableType <> "events" Then
            Err.Raise vbObjectError + 1301, "ex_PersonTimeline", _
                "Unsupported table type for alias '" & tableAlias & "': " & tableType
        End If

        Dim tableName As String
        tableName = mp_GetCfgRequired(cfg, sourceAlias & ".Table[" & tableAlias & "].Name")

        Dim wb As Workbook
        Set wb = mp_GetWorkbookForSource(wbCache, cfg, sourceAlias)

        Dim lo As ListObject
        Set lo = mp_FindListObjectByName(wb, tableName)
        If lo Is Nothing Then
            Err.Raise vbObjectError + 1302, "ex_PersonTimeline", _
                "Table '" & tableName & "' for alias '" & tableAlias & "' was not found in source '" & sourceAlias & "'."
        End If

        If tableType = "state" Then
            rowIndex = mp_WriteStateCardGeneric(wsOut, lo, fio, rowIndex, cfg, sourceAlias, tableAlias)
            rowIndex = rowIndex + 1
        Else
            If mode = PersonTimeline And renderedCount > 0 Then
                wsOut.Cells(rowIndex, 1).Value = "Events [" & tableAlias & "]"
                wsOut.Cells(rowIndex, 1).Font.Bold = True
                rowIndex = rowIndex + 1
            End If
            rowIndex = mp_WriteEventsGeneric(wsOut, lo, fio, rowIndex, cfg, sourceAlias, tableAlias)
            rowIndex = rowIndex + 1
        End If

        renderedCount = renderedCount + 1

ContinueAlias:
    Next i

    If renderedCount = 0 Then
        Err.Raise vbObjectError + 1303, "ex_PersonTimeline", _
            "No tables were rendered for mode '" & ex_Settings.m_GetOutputModeDisplay() & "'. Check Output.Tables and table Type."
    End If

    wsOut.Columns.AutoFit
    ex_SheetTheme.m_ApplyDarkThemeToSheet wsOut

    mp_CloseWorkbooks wbCache

    Application.EnableEvents = prevEnableEvents
    Application.DisplayAlerts = prevDisplayAlerts
    Application.ScreenUpdating = prevScreenUpdating

    Exit Sub

EH:
    On Error Resume Next
    mp_CloseWorkbooks wbCache
    Application.EnableEvents = prevEnableEvents
    Application.DisplayAlerts = prevDisplayAlerts
    Application.ScreenUpdating = prevScreenUpdating
    On Error GoTo 0
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

Private Function mp_WriteStateCardGeneric( _
    ByVal wsOut As Worksheet, _
    ByVal lo As ListObject, _
    ByVal fio As String, _
    ByVal rowIndex As Long, _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String _
) As Long

    Dim fields As Variant
    fields = mp_GetListRequired(cfg, sourceAlias & ".Table[" & tableAlias & "].FieldsAliases")

    Dim keyAlias As String
    keyAlias = mp_GetCfgRequired(cfg, sourceAlias & ".Table[" & tableAlias & "].Key")

    Dim keyHeader As String
    keyHeader = mp_GetMappedSourceHeader(cfg, sourceAlias, tableAlias, keyAlias)

    Dim keyCol As Long
    keyCol = mp_FindHeaderColumnInTable(lo, keyHeader)
    If keyCol <= 0 Then
        Err.Raise vbObjectError + 1310, "ex_PersonTimeline", _
            "State key header not found: '" & keyHeader & "' (alias '" & tableAlias & "')."
    End If

    Dim foundRow As Long
    foundRow = mp_FindDataRowByKeyInTable(lo, keyCol, fio)
    If foundRow <= 0 Then
        Err.Raise vbObjectError + 1311, "ex_PersonTimeline", _
            "State row not found for person '" & fio & "' in table alias '" & tableAlias & "'."
    End If

    Dim i As Long
    For i = LBound(fields) To UBound(fields)
        Dim fieldAlias As String
        fieldAlias = Trim$(CStr(fields(i)))
        If Len(fieldAlias) = 0 Then GoTo ContinueField

        Dim labelText As String
        labelText = mp_GetLabel(cfg, sourceAlias, tableAlias, fieldAlias)

        Dim colIndex As Long
        colIndex = mp_TryGetTableColumnByFieldAlias(lo, cfg, sourceAlias, tableAlias, fieldAlias)

        wsOut.Cells(rowIndex, 1).Value = labelText

        If colIndex > 0 Then
            wsOut.Cells(rowIndex, 2).Value = lo.DataBodyRange.Cells(foundRow, colIndex).Value
        Else
            wsOut.Cells(rowIndex, 2).Value = "(missing column)"
        End If

        rowIndex = rowIndex + 1

ContinueField:
    Next i

    mp_WriteStateCardGeneric = rowIndex

End Function

Private Function mp_WriteEventsGeneric( _
    ByVal wsOut As Worksheet, _
    ByVal lo As ListObject, _
    ByVal fio As String, _
    ByVal rowIndex As Long, _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String _
) As Long

    Dim fields As Variant
    fields = mp_GetListRequired(cfg, sourceAlias & ".Table[" & tableAlias & "].FieldsAliases")

    Dim keyAlias As String
    keyAlias = mp_GetCfgRequired(cfg, sourceAlias & ".Table[" & tableAlias & "].Key")

    Dim keyHeader As String
    keyHeader = mp_GetMappedSourceHeader(cfg, sourceAlias, tableAlias, keyAlias)

    Dim keyCol As Long
    keyCol = mp_FindHeaderColumnInTable(lo, keyHeader)
    If keyCol <= 0 Then
        Err.Raise vbObjectError + 1320, "ex_PersonTimeline", _
            "Events key header not found: '" & keyHeader & "' (alias '" & tableAlias & "')."
    End If

    Dim outHeaderRow As Long
    outHeaderRow = rowIndex

    Dim i As Long
    For i = LBound(fields) To UBound(fields)
        Dim fieldAlias As String
        fieldAlias = Trim$(CStr(fields(i)))
        wsOut.Cells(outHeaderRow, 1 + (i - LBound(fields))).Value = mp_GetLabel(cfg, sourceAlias, tableAlias, fieldAlias)
    Next i

    Dim colIndexes() As Long
    ReDim colIndexes(LBound(fields) To UBound(fields))

    For i = LBound(fields) To UBound(fields)
        colIndexes(i) = mp_TryGetTableColumnByFieldAlias(lo, cfg, sourceAlias, tableAlias, Trim$(CStr(fields(i))))
    Next i

    Dim outDataRow As Long
    outDataRow = outHeaderRow + 1

    If lo.DataBodyRange Is Nothing Then
        wsOut.Cells(outDataRow, 1).Value = "(no events found for this person)"
        mp_WriteEventsGeneric = outDataRow + 1
        Exit Function
    End If

    Dim r As Long
    For r = 1 To lo.DataBodyRange.Rows.Count
        If StrComp(Trim$(CStr(lo.DataBodyRange.Cells(r, keyCol).Value)), fio, vbTextCompare) <> 0 Then
            GoTo ContinueRow
        End If

        For i = LBound(fields) To UBound(fields)
            If colIndexes(i) > 0 Then
                wsOut.Cells(outDataRow, 1 + (i - LBound(fields))).Value = lo.DataBodyRange.Cells(r, colIndexes(i)).Value
            Else
                wsOut.Cells(outDataRow, 1 + (i - LBound(fields))).Value = "(missing column)"
            End If
        Next i

        outDataRow = outDataRow + 1

ContinueRow:
    Next r

    If outDataRow = outHeaderRow + 1 Then
        wsOut.Cells(outDataRow, 1).Value = "(no events found for this person)"
        mp_WriteEventsGeneric = outDataRow + 1
        Exit Function
    End If

    Dim sortAlias As String
    sortAlias = mp_GetCfgOptional(cfg, sourceAlias & ".Table[" & tableAlias & "].Sort", vbNullString)

    If Len(sortAlias) > 0 Then
        Dim sortOutCol As Long
        sortOutCol = -1

        For i = LBound(fields) To UBound(fields)
            If StrComp(Trim$(CStr(fields(i))), sortAlias, vbTextCompare) = 0 Then
                sortOutCol = 1 + (i - LBound(fields))
                Exit For
            End If
        Next i

        If sortOutCol > 0 Then
            mp_SortRangeByColumnIndex wsOut, outHeaderRow, outDataRow - 1, 1, (UBound(fields) - LBound(fields) + 1), sortOutCol
        End If
    End If

    mp_WriteEventsGeneric = outDataRow + 1

End Function

Private Function mp_LoadConfigDictionary() As Object

    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = ws_Dev

    On Error Resume Next
    Set tbl = ws.ListObjects(DEV_CONFIG_TABLE_NAME)
    On Error GoTo 0

    If tbl Is Nothing Then
        Err.Raise vbObjectError + 1330, "ex_PersonTimeline", _
            "Config table '" & DEV_CONFIG_TABLE_NAME & "' was not found on sheet '" & ws.Name & "'."
    End If

    If tbl.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 1331, "ex_PersonTimeline", _
            "Config table '" & DEV_CONFIG_TABLE_NAME & "' has no data rows."
    End If

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    Dim dataRange As Range
    Set dataRange = tbl.DataBodyRange

    Dim r As Long
    For r = 1 To dataRange.Rows.Count
        Dim markerText As String
        markerText = Trim$(CStr(dataRange.Cells(r, DEV_COL_MARKER).Value))
        If StrComp(markerText, DEV_MARKER_SYMBOL, vbTextCompare) = 0 Then
            GoTo ContinueRow
        End If

        Dim keyText As String
        keyText = Trim$(CStr(dataRange.Cells(r, DEV_COL_KEY).Value))
        If Len(keyText) = 0 Then
            GoTo ContinueRow
        End If

        dict(keyText) = CStr(dataRange.Cells(r, DEV_COL_VALUE).Value)

ContinueRow:
    Next r

    Set mp_LoadConfigDictionary = dict

End Function

Private Function mp_FindSourceAliasForTable(ByVal cfg As Object, ByVal tableAlias As String) As String

    Dim sourceAliases As Variant
    sourceAliases = mp_GetSourceAliases(cfg)

    Dim found As String
    Dim i As Long

    For i = LBound(sourceAliases) To UBound(sourceAliases)
        Dim src As String
        src = CStr(sourceAliases(i))

        Dim listKey As String
        listKey = "Source." & src & ".TablesAliases"

        Dim aliases As Variant
        aliases = mp_GetListRequired(cfg, listKey)

        If mp_ArrayContainsText(aliases, tableAlias) Then
            If Len(found) > 0 Then
                Err.Raise vbObjectError + 1340, "ex_PersonTimeline", _
                    "Table alias '" & tableAlias & "' is declared in multiple sources: '" & found & "' and '" & src & "'."
            End If
            found = src
        End If
    Next i

    If Len(found) = 0 Then
        Err.Raise vbObjectError + 1341, "ex_PersonTimeline", _
            "Table alias '" & tableAlias & "' is not declared in any Source.*.TablesAliases."
    End If

    mp_FindSourceAliasForTable = found

End Function

Private Function mp_GetSourceAliases(ByVal cfg As Object) As Variant

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    Dim key As Variant
    For Each key In cfg.Keys
        Dim k As String
        k = CStr(key)

        If LCase$(Left$(k, 7)) = "source." Then
            Dim p As Long
            p = InStr(8, k, ".", vbBinaryCompare)
            If p > 8 Then
                Dim srcAlias As String
                srcAlias = Mid$(k, 8, p - 8)
                If Len(srcAlias) > 0 Then
                    d(srcAlias) = srcAlias
                End If
            End If
        End If
    Next key

    If d.Count = 0 Then
        Err.Raise vbObjectError + 1350, "ex_PersonTimeline", "No Source.* keys found in config."
    End If

    Dim arr() As String
    ReDim arr(0 To d.Count - 1)

    Dim i As Long
    i = 0
    For Each key In d.Keys
        arr(i) = CStr(key)
        i = i + 1
    Next key

    mp_GetSourceAliases = arr

End Function

Private Function mp_GetWorkbookForSource(ByVal wbCache As Object, ByVal cfg As Object, ByVal sourceAlias As String) As Workbook

    If wbCache.Exists(sourceAlias) Then
        Set mp_GetWorkbookForSource = wbCache(sourceAlias)
        Exit Function
    End If

    Dim fileKey As String
    fileKey = "Source." & sourceAlias & ".FilePath"

    Dim sourcePath As String
    sourcePath = mp_ResolvePathLocal(mp_GetCfgRequired(cfg, fileKey))

    If Dir(sourcePath) = vbNullString Then
        Err.Raise vbObjectError + 1360, "ex_PersonTimeline", "Source file not found: " & sourcePath
    End If

    Dim wb As Workbook
    Set wb = Workbooks.Open(Filename:=sourcePath, ReadOnly:=True, UpdateLinks:=0)

    On Error Resume Next
    wb.Windows(1).Visible = False
    On Error GoTo 0

    wbCache.Add sourceAlias, wb
    Set mp_GetWorkbookForSource = wb

End Function

Private Sub mp_CloseWorkbooks(ByVal wbCache As Object)

    If wbCache Is Nothing Then Exit Sub

    On Error Resume Next
    Dim key As Variant
    For Each key In wbCache.Keys
        Dim wb As Workbook
        Set wb = wbCache(key)
        If Not wb Is Nothing Then
            wb.Close SaveChanges:=False
        End If
    Next key
    wbCache.RemoveAll
    On Error GoTo 0

End Sub

Private Function mp_GetCfgRequired(ByVal cfg As Object, ByVal keyName As String) As String

    If Not cfg.Exists(keyName) Then
        Err.Raise vbObjectError + 1370, "ex_PersonTimeline", "Missing config key: " & keyName
    End If

    Dim valueText As String
    valueText = Trim$(CStr(cfg(keyName)))

    If Len(valueText) = 0 Then
        Err.Raise vbObjectError + 1371, "ex_PersonTimeline", "Empty config value: " & keyName
    End If

    mp_GetCfgRequired = valueText

End Function

Private Function mp_GetCfgOptional(ByVal cfg As Object, ByVal keyName As String, ByVal defaultValue As String) As String

    If cfg.Exists(keyName) Then
        mp_GetCfgOptional = Trim$(CStr(cfg(keyName)))
    Else
        mp_GetCfgOptional = defaultValue
    End If

End Function

Private Function mp_GetListRequired(ByVal cfg As Object, ByVal keyName As String) As Variant

    Dim raw As String
    raw = mp_GetCfgRequired(cfg, keyName)
    mp_GetListRequired = mp_SplitList(raw)

    If mp_IsEmptyVariantArray(mp_GetListRequired) Then
        Err.Raise vbObjectError + 1380, "ex_PersonTimeline", "List is empty for config key: " & keyName
    End If

End Function

Private Function mp_SplitList(ByVal raw As String) As Variant

    raw = Trim$(raw)
    If Len(raw) = 0 Then
        mp_SplitList = Array()
        Exit Function
    End If

    raw = Replace$(raw, ",", ";")

    Dim parts As Variant
    parts = Split(raw, ";")

    Dim count As Long
    count = 0

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(CStr(parts(i)))) > 0 Then
            count = count + 1
        End If
    Next i

    If count = 0 Then
        mp_SplitList = Array()
        Exit Function
    End If

    Dim out() As String
    ReDim out(0 To count - 1)

    Dim j As Long
    j = 0
    For i = LBound(parts) To UBound(parts)
        Dim token As String
        token = Trim$(CStr(parts(i)))
        If Len(token) > 0 Then
            out(j) = token
            j = j + 1
        End If
    Next i

    mp_SplitList = out

End Function

Private Function mp_ArrayContainsText(ByVal values As Variant, ByVal needle As String) As Boolean

    If mp_IsEmptyVariantArray(values) Then Exit Function

    Dim i As Long
    For i = LBound(values) To UBound(values)
        If StrComp(Trim$(CStr(values(i))), Trim$(needle), vbTextCompare) = 0 Then
            mp_ArrayContainsText = True
            Exit Function
        End If
    Next i

End Function

Private Function mp_GetMappedSourceHeader( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAlias As String _
) As String

    Dim raw As String
    raw = mp_GetCfgRequired(cfg, sourceAlias & ".Table[" & tableAlias & "].Map[" & fieldAlias & "]")

    Dim p As Long
    p = InStr(1, raw, "|", vbBinaryCompare)

    If p > 0 Then
        mp_GetMappedSourceHeader = Trim$(Left$(raw, p - 1))
    Else
        mp_GetMappedSourceHeader = Trim$(raw)
    End If

    If Len(mp_GetMappedSourceHeader) = 0 Then
        Err.Raise vbObjectError + 1390, "ex_PersonTimeline", _
            "Mapped source header is empty for " & sourceAlias & ".Table[" & tableAlias & "].Map[" & fieldAlias & "]"
    End If

End Function

Private Function mp_GetLabel( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAlias As String _
) As String

    Dim raw As String
    raw = mp_GetCfgRequired(cfg, sourceAlias & ".Table[" & tableAlias & "].Map[" & fieldAlias & "]")

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

    mp_GetLabel = mp_GetMappedSourceHeader(cfg, sourceAlias, tableAlias, fieldAlias)

End Function

Private Function mp_TryGetTableColumnByFieldAlias( _
    ByVal lo As ListObject, _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAlias As String _
) As Long

    Dim headerName As String
    headerName = mp_GetMappedSourceHeader(cfg, sourceAlias, tableAlias, fieldAlias)

    mp_TryGetTableColumnByFieldAlias = mp_FindHeaderColumnInTable(lo, headerName)

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

Private Function mp_FindListObjectByName(ByVal wbSrc As Workbook, ByVal tableName As String) As ListObject

    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In wbSrc.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set mp_FindListObjectByName = lo
                Exit Function
            End If
        Next lo
    Next ws

    Set mp_FindListObjectByName = Nothing

End Function

Private Function mp_FindHeaderColumnInTable(ByVal lo As ListObject, ByVal headerName As String) As Long

    Dim normalizedNeedle As String
    normalizedNeedle = mp_NormalizeHeader(headerName)

    Dim c As Long
    For c = 1 To lo.ListColumns.Count
        If mp_NormalizeHeader(CStr(lo.HeaderRowRange.Cells(1, c).Value)) = normalizedNeedle Then
            mp_FindHeaderColumnInTable = c
            Exit Function
        End If
    Next c

    mp_FindHeaderColumnInTable = -1

End Function

Private Function mp_FindDataRowByKeyInTable(ByVal lo As ListObject, ByVal keyColIndex As Long, ByVal keyValue As String) As Long

    If lo.DataBodyRange Is Nothing Then
        mp_FindDataRowByKeyInTable = -1
        Exit Function
    End If

    Dim r As Long
    For r = 1 To lo.DataBodyRange.Rows.Count
        If StrComp(Trim$(CStr(lo.DataBodyRange.Cells(r, keyColIndex).Value)), keyValue, vbTextCompare) = 0 Then
            mp_FindDataRowByKeyInTable = r
            Exit Function
        End If
    Next r

    mp_FindDataRowByKeyInTable = -1

End Function

Private Sub mp_SortRangeByColumnIndex(ByVal ws As Worksheet, ByVal topRow As Long, ByVal bottomRow As Long, ByVal leftCol As Long, ByVal rightCol As Long, ByVal sortColRelative As Long)

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(topRow, leftCol), ws.Cells(bottomRow, rightCol))

    rng.Sort Key1:=ws.Cells(topRow + 1, leftCol + sortColRelative - 1), Order1:=xlAscending, Header:=xlYes

End Sub

Private Function mp_ResolvePathLocal(ByVal inputPath As String) As String

    Dim basePath As String

    inputPath = Trim$(inputPath)
    If Len(inputPath) = 0 Then Exit Function

    If Left$(inputPath, 2) = "\\" Or InStr(1, inputPath, ":\", vbTextCompare) > 0 Then
        mp_ResolvePathLocal = inputPath
        Exit Function
    End If

    basePath = ThisWorkbook.Path
    If Len(basePath) = 0 Then
        mp_ResolvePathLocal = inputPath
        Exit Function
    End If

    If Right$(basePath, 1) <> "\" Then
        basePath = basePath & "\"
    End If

    mp_ResolvePathLocal = basePath & inputPath

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

    mp_IsEmptyVariantArray = False
    Exit Function

EH:
    mp_IsEmptyVariantArray = True

End Function
