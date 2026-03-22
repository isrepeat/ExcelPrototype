Attribute VB_Name = "ex_ModeReportCreation"
Option Explicit

Private Const RESULT_SHEET_NAME As String = "g_ReportCreation"
Private Const KIND_HEADER As String = "header"
Private Const KIND_SECTION As String = "section"
Private Const KIND_CONTENT As String = "content"
Private Const KIND_OWNER_DIVIDER As String = "ownerdivider"
Private Const BATCH_RUNTIME_VAR As String = "__Batch"
Private Const UI_WARNING_TITLE As String = "WARNING: Mode UI is not loaded"
Private Const UI_WARNING_TEXT As String = "Mode-specific UI config for ReportCreation is unavailable. Control panel/view-zone settings were not applied."

Private Const CONFIG_KEYS_COLLECTION_KEY As String = "KeysCollection"
' Явный alias для member-итерации в DSL: row.SrcTableRows -> <Source>.Sheet[<Table>]
Private Const SRC_TABLE_ROWS_ALIAS As String = "SrcTableRows"

Public Sub m_RunKeysCollectionReport()
    Dim cfg As Object
    Dim pipelineInput As Object

    Set cfg = ex_ConfigProvider.m_LoadConfigDictionary("ex_ModeReportCreation", 6006, 6007)
    Set pipelineInput = mp_CreatePipelineInputForReport(cfg)
    ex_ModePipeline.m_RunModePipeline cfg, "ex_ModeReportCreation.m_RunMode", pipelineInput, False
End Sub

Private Function mp_CreatePipelineInputForReport(ByVal cfg As Object) As Object
    Dim rawKeys As String
    Dim explicitKeys As Collection
    Dim pipelineInput As obj_ScriptIOPayload

    rawKeys = mp_GetOptionalConfigFromDict(cfg, CONFIG_KEYS_COLLECTION_KEY, vbNullString)
    If Len(Trim$(rawKeys)) > 0 Then
        Set explicitKeys = mp_ResolveExplicitKeys(rawKeys, CONFIG_KEYS_COLLECTION_KEY)
    Else
        Set explicitKeys = New Collection
    End If

    Set pipelineInput = New obj_ScriptIOPayload
    If Not explicitKeys Is Nothing Then
        pipelineInput.m_SetObject CONFIG_KEYS_COLLECTION_KEY, explicitKeys
    End If

    Set mp_CreatePipelineInputForReport = pipelineInput
End Function

Private Function mp_ResolveExplicitKeys( _
    ByVal rawKeysText As String, _
    Optional ByVal configKeyName As String = CONFIG_KEYS_COLLECTION_KEY _
) As Collection
    Dim parts() As String
    Dim i As Long
    Dim keyValue As String
    Dim seen As Object
    Dim result As Collection

    rawKeysText = Trim$(CStr(rawKeysText))
    If Len(rawKeysText) = 0 Then
        Err.Raise vbObjectError + 6101, "ex_ModeReportCreation", _
            "Config key '" & Trim$(configKeyName) & "' is empty. Provide keys separated by ';'."
    End If

    parts = Split(rawKeysText, ";")
    Set result = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1

    For i = LBound(parts) To UBound(parts)
        keyValue = Trim$(parts(i))
        If Len(keyValue) > 0 Then
            If Not seen.Exists(keyValue) Then
                seen.Add keyValue, True
                result.Add keyValue
            End If
        End If
    Next i

    If result.Count = 0 Then
        Err.Raise vbObjectError + 6102, "ex_ModeReportCreation", _
            "Config key '" & Trim$(configKeyName) & "' contains no valid keys after parsing."
    End If

    Set mp_ResolveExplicitKeys = result
End Function

Public Function m_RunMode(ByVal cfg As Object, ByVal modeInput As Object, ByVal preProcessContext As Object) As Object
    Dim keys As Collection
    Dim wsOut As Worksheet
    Dim headerRows As Collection
    Dim sectionRows As Collection
    Dim contentRows As Collection
    Dim ownerDividerRows As Collection
    Dim rowKindRanges As Object
    Dim nextRow As Long
    Dim gapRows As Long
    Dim ownerDividerGapRows As Long
    Dim resultTables As Collection
    Dim stateResultTable As obj_ResultTable
    Dim eventsResultTable As obj_ResultTable
    Dim stateSourceAlias As String
    Dim stateTableAlias As String
    Dim eventsSourceAlias As String
    Dim eventsTableAlias As String
    Dim stateFields As Collection
    Dim eventsFields As Collection
    Dim outputStyle As ex_SheetStylesXmlProvider.t_OutputSheetStyle
    Dim hasOutputStyle As Boolean
    Dim warningRangeAddress As String
    Dim warningRange As Range
    Dim modeResult As Object
    Dim batchKeyResultsTable As obj_ResultTable

    On Error GoTo EH

    Set keys = mp_ResolveKeysFromModeInput(modeInput)
    If keys Is Nothing Or keys.Count = 0 Then
        Err.Raise vbObjectError + 6000, "ex_ModeReportCreation", "ReportCreation mode input has no keys. Provide KeysCollection in Output/Input."
    End If

    mp_ResolveOutputTables cfg, stateSourceAlias, stateTableAlias, eventsSourceAlias, eventsTableAlias
    Set stateFields = mp_GetFieldsAliases(stateSourceAlias, stateTableAlias)
    Set eventsFields = mp_GetFieldsAliases(eventsSourceAlias, eventsTableAlias)

    Set wsOut = mp_CreateOrClearSheet(RESULT_SHEET_NAME)
    ex_Messaging.m_ClearResultTableAnchors wsOut
    ex_Messaging.m_ClearResultRowAnchors wsOut
    hasOutputStyle = ex_SheetStylesXmlProvider.m_GetOutputSheetStyle(outputStyle, ThisWorkbook)
    If hasOutputStyle Then
        ex_OutputPanel.m_RenderForSheet wsOut, outputStyle
    Else
        warningRangeAddress = ex_SheetStylesXmlProvider.m_GetOutputWarningBannerRangeAddress(ThisWorkbook)
        ex_Messaging.m_RenderWarningBanner wsOut, UI_WARNING_TEXT, UI_WARNING_TITLE, warningRangeAddress
    End If
    mp_SetSheetTextFormat wsOut
    Set headerRows = New Collection
    Set sectionRows = New Collection
    Set contentRows = New Collection
    Set ownerDividerRows = New Collection
    Set resultTables = New Collection

    Set stateResultTable = mp_CreateResultTableFromFields(stateSourceAlias, stateTableAlias, stateFields)
    resultTables.Add stateResultTable
    Set eventsResultTable = mp_CreateResultTableFromFields(eventsSourceAlias, eventsTableAlias, eventsFields)
    resultTables.Add eventsResultTable

    mp_AddBatchTables resultTables, keys, stateSourceAlias, stateTableAlias, mp_GetRequiredConfig(stateSourceAlias & ".Sheet[" & stateTableAlias & "].Key"), eventsSourceAlias, eventsTableAlias, mp_GetRequiredConfig(eventsSourceAlias & ".Sheet[" & eventsTableAlias & "].Key")

    gapRows = mp_GetOutputGapRows()
    ownerDividerGapRows = mp_GetOwnerDividerGapRows()
    If hasOutputStyle Then
        nextRow = ex_SheetStylesXmlProvider.m_GetOutputViewStartRow(ThisWorkbook)
    Else
        nextRow = 1
        If Len(warningRangeAddress) > 0 Then
            On Error Resume Next
            Set warningRange = wsOut.Range(warningRangeAddress)
            On Error GoTo EH
            If Not warningRange Is Nothing Then
                nextRow = warningRange.Row + warningRange.Rows.Count + 1
            End If
        End If
    End If

    nextRow = mp_RenderTableRows(wsOut, keys, nextRow, headerRows, sectionRows, contentRows, ownerDividerRows, stateResultTable, stateSourceAlias, stateTableAlias, stateFields, ownerDividerGapRows)
    nextRow = mp_AppendGapRows(nextRow, gapRows)
    nextRow = mp_RenderTableRows(wsOut, keys, nextRow, headerRows, sectionRows, contentRows, ownerDividerRows, eventsResultTable, eventsSourceAlias, eventsTableAlias, eventsFields, ownerDividerGapRows)

    Set rowKindRanges = CreateObject("Scripting.Dictionary")
    rowKindRanges.CompareMode = 1
    Set rowKindRanges(KIND_HEADER) = headerRows
    Set rowKindRanges(KIND_SECTION) = sectionRows
    Set rowKindRanges(KIND_CONTENT) = contentRows
    Set rowKindRanges(KIND_OWNER_DIVIDER) = ownerDividerRows

    ex_OutputFormattingPipeline.m_ApplySheetPipeline wsOut, Nothing, Nothing, rowKindRanges, "ReportCreation"

    ' Передаем __Batch как саму таблицу __Batch.Sheet[KeyResults].
    If mp_TryGetBatchKeyResultsTable(resultTables, batchKeyResultsTable) Then
        ex_ScriptIO.m_SetObject modeInput, BATCH_RUNTIME_VAR, batchKeyResultsTable
    End If

    wsOut.Activate
    Set modeResult = CreateObject("Scripting.Dictionary")
    modeResult.CompareMode = 1
    Set modeResult("Output") = modeInput
    Set modeResult("Worksheet") = wsOut
    Set modeResult("ResultTables") = resultTables
    Set m_RunMode = modeResult
    Exit Function

EH:
    MsgBox "ReportCreation failed: [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description, vbExclamation
    Set m_RunMode = Nothing
End Function

Private Function mp_TryGetBatchKeyResultsTable( _
    ByVal resultTables As Collection, _
    ByRef outBatchKeyResultsTable As obj_ResultTable _
) As Boolean
    Dim i As Long
    Dim tableObj As obj_ResultTable

    If resultTables Is Nothing Then Exit Function

    ' Ищем служебную таблицу __Batch.Sheet[KeyResults] и передаем ее как объект __Batch.
    For i = 1 To resultTables.Count
        Set tableObj = resultTables(i)
        If Not tableObj Is Nothing Then
            If StrComp(tableObj.TableRef, "__Batch.Sheet[KeyResults]", vbTextCompare) = 0 Then
                Set outBatchKeyResultsTable = tableObj
                mp_TryGetBatchKeyResultsTable = Not (outBatchKeyResultsTable Is Nothing)
                Exit Function
            End If
        End If
    Next i
End Function

Private Function mp_ResolveKeysFromModeInput(ByVal modeInput As Object) As Collection
    Dim keys As Collection
    Dim keysObject As Object

    Set keys = New Collection

    If ex_ScriptIO.m_TryGetObject(modeInput, CONFIG_KEYS_COLLECTION_KEY, keysObject) Then
        If TypeName(keysObject) = "Collection" Then
            mp_AddUniqueKeysFromCollection keys, keysObject
        ElseIf TypeName(keysObject) = "Dictionary" Or TypeName(keysObject) = "Scripting.Dictionary" Then
            mp_AddUniqueKeysFromDictionary keys, keysObject
        End If
    End If

    Set mp_ResolveKeysFromModeInput = keys
End Function

Private Sub mp_AddUniqueKeysFromCollection(ByVal outKeys As Collection, ByVal sourceKeys As Collection)
    Dim i As Long

    If sourceKeys Is Nothing Then Exit Sub
    For i = 1 To sourceKeys.Count
        mp_AddUniqueKey outKeys, CStr(sourceKeys(i))
    Next i
End Sub

Private Sub mp_AddUniqueKeysFromDictionary(ByVal outKeys As Collection, ByVal sourceDict As Object)
    Dim dictKey As Variant

    If sourceDict Is Nothing Then Exit Sub
    For Each dictKey In sourceDict.Keys
        mp_AddUniqueKey outKeys, CStr(dictKey)
    Next dictKey
End Sub

Private Sub mp_AddUniqueKeysFromDelimited(ByVal outKeys As Collection, ByVal keysDelimited As String)
    Dim parts() As String
    Dim i As Long

    keysDelimited = Trim$(keysDelimited)
    If Len(keysDelimited) = 0 Then Exit Sub

    parts = Split(keysDelimited, ";")
    For i = LBound(parts) To UBound(parts)
        mp_AddUniqueKey outKeys, parts(i)
    Next i
End Sub

Private Sub mp_AddUniqueKey(ByVal outKeys As Collection, ByVal keyText As String)
    Dim i As Long
    Dim normalized As String

    If outKeys Is Nothing Then Exit Sub
    normalized = Trim$(CStr(keyText))
    If Len(normalized) = 0 Then Exit Sub

    For i = 1 To outKeys.Count
        If StrComp(CStr(outKeys(i)), normalized, vbTextCompare) = 0 Then Exit Sub
    Next i

    outKeys.Add normalized
End Sub

Private Function mp_RenderTableRows( _
    ByVal ws As Worksheet, _
    ByVal keys As Collection, _
    ByVal startRow As Long, _
    ByVal headerRows As Collection, _
    ByVal sectionRows As Collection, _
    ByVal contentRows As Collection, _
    ByVal ownerDividerRows As Collection, _
    ByVal resultTable As obj_ResultTable, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAliases As Collection, _
    ByVal ownerDividerGapRows As Long _
) As Long
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim keyFieldAlias As String
    Dim tableRef As String
    Dim nextRow As Long
    Dim hasRows As Boolean
    Dim fieldSelectSql As String
    Dim sortFieldAlias As String
    Dim i As Long
    Dim fieldAlias As String

    nextRow = startRow

    keyFieldAlias = mp_GetRequiredConfig(sourceAlias & ".Sheet[" & tableAlias & "].Key")
    If fieldAliases Is Nothing Or fieldAliases.Count = 0 Then
        Err.Raise vbObjectError + 6010, "ex_ModeReportCreation", "Missing fields aliases for '" & sourceAlias & ".Sheet[" & tableAlias & "]'."
    End If

    Set conn = mp_OpenSourceConnection(sourceAlias)
    tableRef = mp_BuildTableRef(sourceAlias, tableAlias)
    fieldSelectSql = mp_BuildSelectFieldsSql(sourceAlias, tableAlias, fieldAliases)
    sortFieldAlias = mp_GetOptionalConfig(sourceAlias & ".Sheet[" & tableAlias & "].Sort")

    sql = "SELECT " & _
            fieldSelectSql & " " & _
          "FROM " & tableRef & " WHERE " & mp_QuoteIdentifier(mp_MappedHeader(sourceAlias, tableAlias, keyFieldAlias)) & " IN (" & mp_BuildInList(keys) & ") " & _
          "ORDER BY " & mp_QuoteIdentifier(mp_MappedHeader(sourceAlias, tableAlias, keyFieldAlias)) & mp_BuildSortSqlTail(sourceAlias, tableAlias, sortFieldAlias, keyFieldAlias)

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 0, 1

    ws.Cells(nextRow, 1).Value = tableAlias
    sectionRows.Add nextRow
    nextRow = nextRow + 1

    For i = 1 To fieldAliases.Count
        fieldAlias = CStr(fieldAliases(i))
        ws.Cells(nextRow, i).Value = mp_MappedHeader(sourceAlias, tableAlias, fieldAlias)
    Next i
    headerRows.Add nextRow
    nextRow = nextRow + 1

    hasRows = mp_AppendRowsGeneric(ws, rs, nextRow, contentRows, ownerDividerRows, resultTable, sourceAlias, tableAlias, keyFieldAlias, fieldAliases, ownerDividerGapRows)
    If hasRows Then
        nextRow = mp_GetNextOutputRow(ws)
    Else
        ws.Cells(nextRow, 1).Value = "(no rows)"
        ws.Cells(nextRow, 2).Value = "-"
        contentRows.Add nextRow
        nextRow = nextRow + 1
    End If

    rs.Close
    conn.Close

    mp_RenderTableRows = nextRow
End Function

Private Function mp_AppendRowsGeneric( _
    ByVal ws As Worksheet, _
    ByVal rs As Object, _
    ByVal startRow As Long, _
    ByVal contentRows As Collection, _
    ByVal ownerDividerRows As Collection, _
    ByVal resultTable As obj_ResultTable, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal keyFieldAlias As String, _
    ByVal fieldAliases As Collection, _
    ByVal ownerDividerGapRows As Long _
) As Boolean
    Dim nextRow As Long
    Dim tableRowIndex As Long
    Dim colIndex As Long
    Dim fieldAlias As String
    Dim cellValue As String
    Dim fieldType As Long
    Dim prevOwnerKey As String
    Dim currentOwnerKey As String
    Dim rowObj As obj_ResultRow
    Dim rowAnchorName As String
    Dim rowOrdinal As Long

    If rs Is Nothing Then Exit Function
    If rs.EOF Then Exit Function
    If fieldAliases Is Nothing Or fieldAliases.Count = 0 Then Exit Function

    nextRow = startRow
    Do While Not rs.EOF
        tableRowIndex = -1
        If Not resultTable Is Nothing Then tableRowIndex = resultTable.Count
        currentOwnerKey = mp_AsText(rs.Fields(keyFieldAlias).Value)

        If Len(prevOwnerKey) > 0 Then
            If StrComp(prevOwnerKey, currentOwnerKey, vbTextCompare) <> 0 Then
                nextRow = mp_AppendOwnerDividerRows(ws, nextRow, ownerDividerRows, ownerDividerGapRows)
            End If
        End If

        For colIndex = 1 To fieldAliases.Count
            fieldAlias = CStr(fieldAliases(colIndex))
            fieldType = rs.Fields(fieldAlias).Type
            cellValue = mp_AsText(rs.Fields(fieldAlias).Value, fieldType)
            ws.Cells(nextRow, colIndex).Value = cellValue
            If Not resultTable Is Nothing Then
                mp_AddResultCell resultTable, tableRowIndex, sourceAlias, tableAlias, fieldAlias, cellValue
            End If
        Next colIndex

        If Not resultTable Is Nothing Then
            ' Для каждой строки сохраняем ссылку на таблицу-источник.
            ' Это позволяет в DSL писать: for (let r2 in row.SrcTableRows) { ... }
            mp_AddResultCell resultTable, tableRowIndex, sourceAlias, tableAlias, SRC_TABLE_ROWS_ALIAS, sourceAlias & ".Sheet[" & tableAlias & "]"
        End If

        If Not resultTable Is Nothing Then
            Set rowObj = resultTable.EnsureRow(tableRowIndex)
            rowOrdinal = tableRowIndex + 1
            rowAnchorName = ex_Messaging.m_BuildResultRowAnchorName(resultTable.TableRef, rowOrdinal)
            If Len(rowAnchorName) > 0 Then
                rowObj.RowAnchorName = rowAnchorName
                ex_Messaging.m_RegisterResultRowAnchor ws, rowAnchorName, nextRow
            End If
        End If

        prevOwnerKey = currentOwnerKey

        contentRows.Add nextRow
        nextRow = nextRow + 1
        rs.MoveNext
    Loop

    mp_AppendRowsGeneric = True
End Function

Private Function mp_AppendOwnerDividerRows( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal ownerDividerRows As Collection, _
    ByVal dividerRowsCount As Long _
) As Long
    Dim i As Long
    Dim nextRow As Long

    nextRow = startRow
    If dividerRowsCount <= 0 Then
        mp_AppendOwnerDividerRows = nextRow
        Exit Function
    End If

    For i = 1 To dividerRowsCount
        ws.Rows(nextRow).ClearContents
        If Not ownerDividerRows Is Nothing Then ownerDividerRows.Add nextRow
        nextRow = nextRow + 1
    Next i

    mp_AppendOwnerDividerRows = nextRow
End Function

Private Function mp_AppendGapRows(ByVal startRow As Long, ByVal gapRows As Long) As Long
    If gapRows <= 0 Then
        mp_AppendGapRows = startRow
    Else
        mp_AppendGapRows = startRow + gapRows
    End If
End Function

Private Function mp_GetOutputGapRows() As Long
    Dim rawValue As String
    Dim parsedValue As Long

    rawValue = Trim$(CStr(ex_ConfigProvider.m_GetConfigValue("Output.Layout.Gap.Default", "0")))
    If Len(rawValue) = 0 Then
        mp_GetOutputGapRows = 0
        Exit Function
    End If
    If Not IsNumeric(rawValue) Then
        mp_GetOutputGapRows = 0
        Exit Function
    End If

    parsedValue = CLng(rawValue)
    If parsedValue < 0 Then parsedValue = 0
    mp_GetOutputGapRows = parsedValue
End Function

Private Function mp_GetOwnerDividerGapRows() As Long
    Dim rawValue As String
    Dim parsedValue As Long

    rawValue = Trim$(CStr(ex_ConfigProvider.m_GetConfigValue("Output.Layout.Gap.OwnerDivider", "0")))
    If Len(rawValue) = 0 Then
        mp_GetOwnerDividerGapRows = 0
        Exit Function
    End If
    If Not IsNumeric(rawValue) Then
        mp_GetOwnerDividerGapRows = 0
        Exit Function
    End If

    parsedValue = CLng(rawValue)
    If parsedValue < 0 Then parsedValue = 0
    mp_GetOwnerDividerGapRows = parsedValue
End Function

Private Function mp_AsText(ByVal valueIn As Variant, Optional ByVal adoFieldType As Long = -1) As String
    mp_AsText = ex_SqlAdoHelpers.m_ToNormalizedText(valueIn, adoFieldType)
End Function

Private Sub mp_SetSheetTextFormat(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    ws.Cells.NumberFormat = "@"
End Sub

Private Function mp_NormalizePhoneText(ByVal valueIn As Variant) As String
    Dim rawText As String
    Dim numericText As String

    rawText = Trim$(mp_AsText(valueIn))
    If Len(rawText) = 0 Then Exit Function

    If InStr(1, rawText, "E+", vbTextCompare) > 0 Then
        On Error Resume Next
        numericText = Format$(CDbl(rawText), "0")
        On Error GoTo 0
        If Len(numericText) > 0 Then rawText = numericText
    End If

    If Left$(rawText, 1) <> "+" Then
        If Left$(rawText, 2) = "38" And Len(rawText) >= 11 Then
            rawText = "+" & rawText
        End If
    End If

    mp_NormalizePhoneText = rawText
End Function

Private Function mp_GetNextOutputRow(ByVal ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        mp_GetNextOutputRow = 2
    Else
        mp_GetNextOutputRow = lastRow + 1
    End If
End Function

Private Function mp_BuildInList(ByVal keys As Collection) As String
    Dim i As Long
    Dim token As String

    For i = 1 To keys.Count
        token = Replace$(CStr(keys(i)), "'", "''")
        If i > 1 Then mp_BuildInList = mp_BuildInList & ","
        mp_BuildInList = mp_BuildInList & "'" & token & "'"
    Next i
End Function

Private Function mp_BuildSelectFieldsSql( _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAliases As Collection _
) As String
    Dim i As Long
    Dim fieldAlias As String
    Dim mappedHeader As String

    If fieldAliases Is Nothing Or fieldAliases.Count = 0 Then
        Err.Raise vbObjectError + 6011, "ex_ModeReportCreation", "Cannot build SQL SELECT fields for empty fields list at '" & sourceAlias & ".Sheet[" & tableAlias & "]'."
    End If

    For i = 1 To fieldAliases.Count
        fieldAlias = CStr(fieldAliases(i))
        mappedHeader = mp_MappedHeader(sourceAlias, tableAlias, fieldAlias)
        If i > 1 Then mp_BuildSelectFieldsSql = mp_BuildSelectFieldsSql & ", "
        mp_BuildSelectFieldsSql = mp_BuildSelectFieldsSql & mp_QuoteIdentifier(mappedHeader) & " AS " & mp_QuoteIdentifier(fieldAlias)
    Next i
End Function

Private Function mp_BuildSortSqlTail( _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal sortFieldAlias As String, _
    ByVal keyFieldAlias As String _
) As String
    sortFieldAlias = Trim$(sortFieldAlias)
    If Len(sortFieldAlias) = 0 Then Exit Function
    If StrComp(sortFieldAlias, keyFieldAlias, vbTextCompare) = 0 Then Exit Function
    mp_BuildSortSqlTail = ", " & mp_QuoteIdentifier(mp_MappedHeader(sourceAlias, tableAlias, sortFieldAlias))
End Function

Private Function mp_GetFieldsAliases(ByVal sourceAlias As String, ByVal tableAlias As String) As Collection
    Dim rawList As String
    Dim listItems As Collection
    Dim keyAlias As String

    rawList = mp_GetRequiredConfig(sourceAlias & ".Sheet[" & tableAlias & "].FieldsAliases")
    Set listItems = mp_ParseSemicolonList(rawList)
    If listItems Is Nothing Or listItems.Count = 0 Then
        Err.Raise vbObjectError + 6012, "ex_ModeReportCreation", "Config key '" & sourceAlias & ".Sheet[" & tableAlias & "].FieldsAliases' is empty."
    End If

    keyAlias = mp_GetRequiredConfig(sourceAlias & ".Sheet[" & tableAlias & "].Key")
    If Not mp_CollectionContainsText(listItems, keyAlias) Then
        listItems.Add keyAlias, Before:=1
    End If

    Set mp_GetFieldsAliases = listItems
End Function

Private Sub mp_ResolveOutputTables( _
    ByVal cfg As Object, _
    ByRef outStateSourceAlias As String, _
    ByRef outStateTableAlias As String, _
    ByRef outEventsSourceAlias As String, _
    ByRef outEventsTableAlias As String _
)
    Dim outputTables As Collection
    Dim i As Long
    Dim tableAlias As String
    Dim sourceAlias As String
    Dim tableType As String

    Set outputTables = mp_ParseSemicolonList(mp_GetRequiredConfigFromDict(cfg, "Output.Sheets"))
    If outputTables Is Nothing Or outputTables.Count = 0 Then
        Err.Raise vbObjectError + 6013, "ex_ModeReportCreation", "Config key 'Output.Sheets' is empty."
    End If

    For i = 1 To outputTables.Count
        tableAlias = CStr(outputTables(i))
        sourceAlias = mp_GetSourceAliasForTable(cfg, tableAlias)
        tableType = LCase$(mp_GetRequiredConfigFromDict(cfg, sourceAlias & ".Sheet[" & tableAlias & "].Type"))

        Select Case tableType
            Case "state"
                outStateSourceAlias = sourceAlias
                outStateTableAlias = tableAlias
            Case "events"
                outEventsSourceAlias = sourceAlias
                outEventsTableAlias = tableAlias
        End Select
    Next i

    If Len(outStateSourceAlias) = 0 Or Len(outStateTableAlias) = 0 Then
        Err.Raise vbObjectError + 6014, "ex_ModeReportCreation", "Output table with Type='State' was not found in Output.Sheets."
    End If
    If Len(outEventsSourceAlias) = 0 Or Len(outEventsTableAlias) = 0 Then
        Err.Raise vbObjectError + 6015, "ex_ModeReportCreation", "Output table with Type='Events' was not found in Output.Sheets."
    End If
End Sub

Private Function mp_GetSourceAliasForTable(ByVal cfg As Object, ByVal tableAlias As String) As String
    Dim explicitSource As String
    Dim keyName As Variant
    Dim sourceAlias As String
    Dim sheetAliases As Collection

    explicitSource = mp_GetOptionalConfig("Output.Sheet[" & tableAlias & "].SourceAlias")
    If Len(explicitSource) > 0 Then
        mp_GetSourceAliasForTable = explicitSource
        Exit Function
    End If

    If cfg Is Nothing Then
        Err.Raise vbObjectError + 6016, "ex_ModeReportCreation", "Config dictionary is not available to resolve source alias for table '" & tableAlias & "'."
    End If

    For Each keyName In cfg.Keys
        If LCase$(Right$(CStr(keyName), Len(".SheetAliases"))) = LCase$(".SheetAliases") Then
            sourceAlias = Mid$(CStr(keyName), Len("Source.") + 1, Len(CStr(keyName)) - Len("Source.") - Len(".SheetAliases"))
            If Len(sourceAlias) > 0 Then
                Set sheetAliases = mp_ParseSemicolonList(CStr(cfg(CStr(keyName))))
                If mp_CollectionContainsText(sheetAliases, tableAlias) Then
                    mp_GetSourceAliasForTable = sourceAlias
                    Exit Function
                End If
            End If
        End If
    Next keyName

    Err.Raise vbObjectError + 6017, "ex_ModeReportCreation", "Cannot resolve source alias for table '" & tableAlias & "'. Add Output.Sheet[" & tableAlias & "].SourceAlias or include table alias in Source.<Alias>.SheetAliases."
End Function

Private Function mp_ParseSemicolonList(ByVal rawText As String) As Collection
    Dim parts() As String
    Dim i As Long
    Dim itemText As String
    Dim result As Collection
    Dim seen As Object

    rawText = Trim$(rawText)
    If Len(rawText) = 0 Then Exit Function

    parts = Split(rawText, ";")
    Set result = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1

    For i = LBound(parts) To UBound(parts)
        itemText = Trim$(CStr(parts(i)))
        If Len(itemText) > 0 Then
            If Not seen.Exists(itemText) Then
                seen.Add itemText, True
                result.Add itemText
            End If
        End If
    Next i

    If result.Count > 0 Then Set mp_ParseSemicolonList = result
End Function

Private Function mp_CollectionContainsText(ByVal listItems As Collection, ByVal valueText As String) As Boolean
    Dim i As Long

    If listItems Is Nothing Then Exit Function
    valueText = Trim$(valueText)
    If Len(valueText) = 0 Then Exit Function

    For i = 1 To listItems.Count
        If StrComp(CStr(listItems(i)), valueText, vbTextCompare) = 0 Then
            mp_CollectionContainsText = True
            Exit Function
        End If
    Next i
End Function

Private Function mp_GetRequiredConfig(ByVal keyName As String) As String
    mp_GetRequiredConfig = Trim$(CStr(ex_ConfigProvider.m_GetConfigValue(keyName, vbNullString)))
    If Len(mp_GetRequiredConfig) = 0 Then
        Err.Raise vbObjectError + 6018, "ex_ModeReportCreation", "Missing required config key '" & keyName & "'."
    End If
End Function

Private Function mp_GetOptionalConfig(ByVal keyName As String) As String
    mp_GetOptionalConfig = Trim$(CStr(ex_ConfigProvider.m_GetConfigValue(keyName, vbNullString)))
End Function

Private Function mp_GetOptionalConfigFromDict( _
    ByVal cfg As Object, _
    ByVal keyName As String, _
    ByVal defaultValue As String _
) As String
    keyName = Trim$(CStr(keyName))
    If Len(keyName) = 0 Then
        mp_GetOptionalConfigFromDict = CStr(defaultValue)
        Exit Function
    End If

    If Not cfg Is Nothing Then
        If cfg.Exists(keyName) Then
            mp_GetOptionalConfigFromDict = Trim$(CStr(cfg(keyName)))
            Exit Function
        End If
    End If

    mp_GetOptionalConfigFromDict = Trim$(CStr(ex_ConfigProvider.m_GetConfigValue(keyName, defaultValue)))
End Function

Private Function mp_GetRequiredConfigFromDict(ByVal cfg As Object, ByVal keyName As String) As String
    If cfg Is Nothing Then
        Err.Raise vbObjectError + 6019, "ex_ModeReportCreation", "Config dictionary is not initialized when reading key '" & keyName & "'."
    End If
    If Not cfg.Exists(keyName) Then
        Err.Raise vbObjectError + 6020, "ex_ModeReportCreation", "Missing required config key '" & keyName & "'."
    End If
    mp_GetRequiredConfigFromDict = Trim$(CStr(cfg(keyName)))
    If Len(mp_GetRequiredConfigFromDict) = 0 Then
        Err.Raise vbObjectError + 6021, "ex_ModeReportCreation", "Config key '" & keyName & "' is empty."
    End If
End Function

Private Function mp_CreateResultTable( _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAliases As Variant _
) As obj_ResultTable
    Dim tableObj As obj_ResultTable
    Dim i As Long
    Dim fieldAlias As String

    Set tableObj = New obj_ResultTable
    tableObj.Initialize sourceAlias & ".Sheet[" & tableAlias & "]"

    If IsArray(fieldAliases) Then
        For i = LBound(fieldAliases) To UBound(fieldAliases)
            fieldAlias = Trim$(CStr(fieldAliases(i)))
            If Len(fieldAlias) > 0 Then
                tableObj.AddFieldMap fieldAlias, ex_ResultRuntimeAdapter.m_BuildMapKey(sourceAlias, tableAlias, fieldAlias)
            End If
        Next i
    End If

    Set mp_CreateResultTable = tableObj
End Function

Private Function mp_CreateResultTableFromFields( _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAliases As Collection _
) As obj_ResultTable
    Dim tableObj As obj_ResultTable
    Dim i As Long
    Dim fieldAlias As String

    Set tableObj = New obj_ResultTable
    tableObj.Initialize sourceAlias & ".Sheet[" & tableAlias & "]"

    If fieldAliases Is Nothing Or fieldAliases.Count = 0 Then
        Err.Raise vbObjectError + 6022, "ex_ModeReportCreation", "Cannot create result table for empty fields list at '" & sourceAlias & ".Sheet[" & tableAlias & "]'."
    End If

    For i = 1 To fieldAliases.Count
        fieldAlias = Trim$(CStr(fieldAliases(i)))
        If Len(fieldAlias) > 0 Then
            tableObj.AddFieldMap fieldAlias, ex_ResultRuntimeAdapter.m_BuildMapKey(sourceAlias, tableAlias, fieldAlias)
        End If
    Next i

    Set mp_CreateResultTableFromFields = tableObj
End Function

Private Sub mp_AddResultCell( _
    ByVal resultTable As obj_ResultTable, _
    ByVal rowIndex As Long, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAlias As String, _
    ByVal valueText As String _
)
    If resultTable Is Nothing Then Exit Sub
    resultTable.SetRowValue rowIndex, fieldAlias, ex_ResultRuntimeAdapter.m_BuildMapKey(sourceAlias, tableAlias, fieldAlias), valueText
End Sub

Private Sub mp_AddBatchTables( _
    ByVal resultTables As Collection, _
    ByVal keys As Collection, _
    ByVal stateSourceAlias As String, _
    ByVal stateTableAlias As String, _
    ByVal stateKeyFieldAlias As String, _
    ByVal eventsSourceAlias As String, _
    ByVal eventsTableAlias As String, _
    ByVal eventsKeyFieldAlias As String _
)
    Dim keysResultsTable As obj_ResultTable
    Dim i As Long
    Dim stateTableRowsAlias As String
    Dim eventsTableRowsAlias As String

    If resultTables Is Nothing Then Exit Sub
    If keys Is Nothing Then Exit Sub

    ' __Batch состоит из одной таблицы KeyResults:
    ' Key + ссылки на строки таблиц источников для каждого ключа.
    stateTableRowsAlias = stateSourceAlias & "TableRows"
    eventsTableRowsAlias = eventsSourceAlias & "TableRows"

    Set keysResultsTable = mp_CreateResultTable("__Batch", "KeyResults", Array("Key", stateTableRowsAlias, eventsTableRowsAlias))
    For i = 1 To keys.Count
        mp_AddResultCell keysResultsTable, keysResultsTable.Count, "__Batch", "KeyResults", "Key", CStr(keys(i))
        mp_AddResultCell keysResultsTable, keysResultsTable.Count - 1, "__Batch", "KeyResults", stateTableRowsAlias, stateSourceAlias & ".Sheet[" & stateTableAlias & "]"
        mp_AddResultCell keysResultsTable, keysResultsTable.Count - 1, "__Batch", "KeyResults", eventsTableRowsAlias, eventsSourceAlias & ".Sheet[" & eventsTableAlias & "]"
    Next i
    resultTables.Add keysResultsTable
End Sub

Private Sub mp_MergeInjectedRuntimeContext( _
    ByVal targetScope As Object, _
    ByVal sourceContext As Object, _
    ByVal sourceKey As String _
)
    Dim sourceScope As Object
    Dim scopeKey As Variant

    If targetScope Is Nothing Then Exit Sub
    If sourceContext Is Nothing Then Exit Sub
    If Not sourceContext.Exists(sourceKey) Then Exit Sub
    If Not IsObject(sourceContext(sourceKey)) Then Exit Sub

    Set sourceScope = sourceContext(sourceKey)
    If sourceScope Is Nothing Then Exit Sub

    For Each scopeKey In sourceScope.Keys
        targetScope(CStr(scopeKey)) = sourceScope(CStr(scopeKey))
    Next scopeKey
End Sub

Private Function mp_OpenSourceConnection(ByVal sourceAlias As String) As Object
    Dim sourcePath As String
    Dim snapshotPath As String
    Dim conn As Object

    sourcePath = mp_ResolvePath(CStr(ex_ConfigProvider.m_GetConfigValue("Source." & sourceAlias & ".FilePath", vbNullString)))
    If Len(sourcePath) = 0 Then
        Err.Raise vbObjectError + 6001, "ex_ModeReportCreation", "Missing config key 'Source." & sourceAlias & ".FilePath'."
    End If
    If Dir(sourcePath) = vbNullString Then
        Err.Raise vbObjectError + 6002, "ex_ModeReportCreation", "Source file not found: " & sourcePath
    End If

    snapshotPath = ex_SourceSnapshot.m_GetSnapshotPath(sourcePath, "Source." & sourceAlias)

    Set conn = CreateObject("ADODB.Connection")
    conn.Open mp_BuildAdoConnectionString(snapshotPath)
    Set mp_OpenSourceConnection = conn
End Function

Private Function mp_BuildTableRef(ByVal sourceAlias As String, ByVal tableAlias As String) As String
    Dim sheetName As String
    sheetName = ex_ConfigProvider.m_GetResolvedSheetName(sourceAlias, tableAlias, Nothing, True, "ex_ModeReportCreation", 6003, 6003, 6022, 6023, 6024)
    mp_BuildTableRef = "[" & Replace$(sheetName, "]", "]]" ) & "$]"
End Function

Private Function mp_MappedHeader(ByVal sourceAlias As String, ByVal tableAlias As String, ByVal fieldAlias As String) As String
    Dim rawValue As String
    Dim splitPos As Long

    rawValue = Trim$(CStr(ex_ConfigProvider.m_GetConfigValue(sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]", vbNullString)))
    If Len(rawValue) = 0 Then
        Err.Raise vbObjectError + 6004, "ex_ModeReportCreation", "Missing mapping config for " & sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]"
    End If

    splitPos = InStr(1, rawValue, "|", vbBinaryCompare)
    If splitPos > 0 Then
        mp_MappedHeader = Trim$(Left$(rawValue, splitPos - 1))
    Else
        mp_MappedHeader = rawValue
    End If
End Function

Private Function mp_ResolvePath(ByVal pathText As String) As String
    Dim trimmedPath As String
    Dim fso As Object

    trimmedPath = Trim$(pathText)
    If Len(trimmedPath) = 0 Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.DriveExists(Left$(trimmedPath, 1)) Or InStr(1, trimmedPath, ":\", vbTextCompare) > 0 Then
        mp_ResolvePath = trimmedPath
    Else
        mp_ResolvePath = fso.GetAbsolutePathName(ThisWorkbook.Path & "\" & trimmedPath)
    End If
End Function

Private Function mp_BuildAdoConnectionString(ByVal sourcePath As String) As String
    Dim ext As String
    Dim props As String

    ext = LCase$(Mid$(sourcePath, InStrRev(sourcePath, ".") + 1))

    Select Case ext
        Case "xls"
            props = "Excel 8.0;HDR=YES;IMEX=1;ReadOnly=True"
        Case "xlsx"
            props = "Excel 12.0 Xml;HDR=YES;IMEX=1;ReadOnly=True"
        Case "xlsm"
            props = "Excel 12.0 Macro;HDR=YES;IMEX=1;ReadOnly=True"
        Case "xlsb"
            props = "Excel 12.0;HDR=YES;IMEX=1;ReadOnly=True"
        Case Else
            Err.Raise vbObjectError + 6005, "ex_ModeReportCreation", "Unsupported source file extension for ADO: ." & ext
    End Select

    mp_BuildAdoConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sourcePath & ";Extended Properties=""" & props & """;"
End Function

Private Function mp_QuoteIdentifier(ByVal valueText As String) As String
    mp_QuoteIdentifier = "[" & Replace$(Trim$(valueText), "]", "]]" ) & "]"
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

    Set mp_CreateOrClearSheet = ws
End Function
