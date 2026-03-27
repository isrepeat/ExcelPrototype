Attribute VB_Name = "ex_ModeMultiSources"
Option Explicit

Private Const SUMMARY_SHEET_NAME As String = "g_MultiSources"
Private Const RESULT_SHEET_PREFIX As String = "g_MS_"
Private Const SCRIPT_INPUT_RESULT_TABLES_KEY As String = "__ResultTables"
Private Const INPUT_KEY_USE_RESULT_LAYOUT As String = "__UseResultLayoutScript"
Private Const INPUT_KEY_QUERY_TABLE_REFS As String = "Query.TableRefs"
Private Const PREPROCESS_CONTEXT_HAS_SCRIPT As String = "HasScript"
Private Const LIKE_DIALECT_UNKNOWN As String = "unknown"
Private Const LIKE_DIALECT_STAR As String = "star"
Private Const LIKE_DIALECT_PERCENT As String = "percent"

Private g_LikeDialectByConnection As Object

Public Sub m_RunMultiSources()
    Dim cfg As Object
    Dim fio As String
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo EH

    fio = Trim$(ex_ConfigProvider.m_GetConfigValue("CommonKey", vbNullString))

    Set cfg = ex_ConfigProvider.m_LoadConfigDictionary("ex_ModeMultiSources", 6501, 6502)
    ex_ModePipeline.m_RunModePipeline cfg, "ex_ModeMultiSources.m_RunMode", mp_CreateScriptInputContext(fio), False
    Exit Sub

EH:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    mp_RestoreExcelUiStateSafe
    If errNumber = 0 Then errNumber = vbObjectError + 6500
    If Len(errSource) = 0 Then errSource = "ex_ModeMultiSources"
    If Len(errDescription) = 0 Then errDescription = "Unknown mode execution error."
    MsgBox "MultiSources failed: [" & errSource & " #" & CStr(errNumber) & "] " & errDescription, vbExclamation
End Sub

Public Function m_RunMode(ByVal cfg As Object, ByVal modeInput As Object, ByVal preProcessContext As Object) As Object
    Dim modeResult As Object
    Dim resultTables As Collection
    Dim resultTableRefs As Collection
    Dim outputEntries As Collection
    Dim summarySheet As Worksheet
    Dim commonKey As String
    Dim commonKeyType As String
    Dim useLikeMatch As Boolean
    Dim rowKindRanges As Object
    Dim headerRows As Collection
    Dim sectionRows As Collection
    Dim contentRows As Collection
    Dim fields As Collection
    Dim resultTable As obj_ResultTable
    Dim conn As Object
    Dim rs As Object
    Dim schemaRs As Object
    Dim i As Long
    Dim entry As Object
    Dim sourceAlias As String
    Dim tableAlias As String
    Dim rawToken As String
    Dim sourceFilePath As String
    Dim sheetName As String
    Dim keyFieldAlias As String
    Dim tableRef As String
    Dim hasRows As Boolean
    Dim fieldAlias As String
    Dim colIndex As Long
    Dim rowIndex As Long
    Dim tableVisualStartRow As Long
    Dim tableVisualEndRow As Long
    Dim fieldOrdinals As Object
    Dim expectedHeaders As Variant
    Dim keySourceHeader As String
    Dim detectedTableRef As String
    Dim isExplicitSheetRange As Boolean
    Dim selectSql As String
    Dim keySourceFieldName As String
    Dim likeDialect As String
    Dim hasOutputStyle As Boolean
    Dim outputStyle As ex_SheetStylesXmlProvider.t_OutputSheetStyle
    Dim summaryTopRow As Long
    Dim useResultLayoutScript As Boolean
    Dim wbCache As Object
    Dim longTextRuntimeCache As Object
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo EH

    If cfg Is Nothing Then
        Err.Raise vbObjectError + 6503, "ex_ModeMultiSources", "Config dictionary is not initialized."
    End If

    ex_ConfigVirtualSources.m_ExpandVirtualSourcesAndOutput cfg, "ex_ModeMultiSources"
    Set outputEntries = mp_ResolveQueryOutputEntries(cfg, modeInput, preProcessContext)
    If outputEntries Is Nothing Or outputEntries.Count = 0 Then
        Err.Raise vbObjectError + 6504, "ex_ModeMultiSources", "List is empty for input key: Query.TableRefs"
    End If

    commonKey = Trim$(ex_ScriptIO.m_GetStringOrDefault(modeInput, "CommonKey", vbNullString))
    If Len(commonKey) = 0 Then
        Err.Raise vbObjectError + 6505, "ex_ModeMultiSources", "Config key 'CommonKey' is empty."
    End If
    commonKeyType = LCase$(Trim$(ex_ConfigProvider.m_GetConfigEntryType("CommonKey", vbNullString)))
    useLikeMatch = (StrComp(commonKeyType, "rx", vbTextCompare) = 0)
    useResultLayoutScript = (StrComp( _
        ex_ScriptIO.m_GetStringOrDefault(modeInput, INPUT_KEY_USE_RESULT_LAYOUT, "0"), _
        "1", _
        vbTextCompare) = 0)

    mp_DeleteGeneratedResultSheets
    Set summarySheet = mp_CreateOrClearSheet(SUMMARY_SHEET_NAME)
    ex_Messaging.m_ClearResultTableAnchors summarySheet
    ex_Messaging.m_ClearResultRowAnchors summarySheet
    summarySheet.Cells.NumberFormat = "@"

    hasOutputStyle = ex_SheetStylesXmlProvider.m_GetOutputSheetStyle(outputStyle, ThisWorkbook)
    summaryTopRow = 1
    If hasOutputStyle Then
        summaryTopRow = ex_SheetStylesXmlProvider.m_GetOutputViewStartRow(ThisWorkbook)
    End If

    Set headerRows = New Collection
    Set sectionRows = New Collection
    Set contentRows = New Collection
    Set resultTables = New Collection
    Set resultTableRefs = New Collection
    Set wbCache = CreateObject("Scripting.Dictionary")
    wbCache.CompareMode = 1
    Set longTextRuntimeCache = ex_ResultSqlEngine.m_GetLongTextRuntimeCache("multisources")

    rowIndex = summaryTopRow
    For i = 1 To outputEntries.Count
        tableVisualStartRow = rowIndex
        Set entry = outputEntries(i)
        sourceAlias = Trim$(CStr(entry("SourceAlias")))
        tableAlias = Trim$(CStr(entry("TableAlias")))
        If Len(sourceAlias) = 0 Then
            Err.Raise vbObjectError + 6540, "ex_ModeMultiSources", "Output entry #" & CStr(i) & " has empty SourceAlias."
        End If
        If Len(tableAlias) = 0 Then
            Err.Raise vbObjectError + 6541, "ex_ModeMultiSources", "Output entry #" & CStr(i) & " has empty TableAlias."
        End If
        rawToken = Trim$(CStr(entry("RawToken")))

        sourceFilePath = mp_GetResolvedSourcePath(cfg, sourceAlias)
        If Len(sourceFilePath) = 0 Or Len(Dir$(sourceFilePath)) = 0 Then
            Err.Raise vbObjectError + 6506, "ex_ModeMultiSources", "Source file was not found: " & sourceFilePath
        End If
        sheetName = ex_ConfigProvider.m_GetResolvedSheetName(sourceAlias, tableAlias, cfg, False, "ex_ModeMultiSources", 6510, 6511, 6512, 6513, 6514)
        isExplicitSheetRange = mp_IsExplicitAdoRangeReference(sheetName)
        keyFieldAlias = mp_GetCfgRequired(cfg, sourceAlias & ".Sheet[" & tableAlias & "].Key")
        Set fields = mp_GetFieldsAliases(cfg, sourceAlias, tableAlias)
        If Not mp_CollectionContainsText(fields, keyFieldAlias) Then
            fields.Add keyFieldAlias, Before:=1
        End If

        summarySheet.Cells(rowIndex, 1).Value = mp_BuildSectionTitle(i, rawToken, sourceFilePath, sheetName)
        sectionRows.Add rowIndex
        rowIndex = rowIndex + 1

        For colIndex = 1 To fields.Count
            fieldAlias = CStr(fields(colIndex))
            summarySheet.Cells(rowIndex, colIndex).Value = mp_GetFieldDisplayHeader(cfg, sourceAlias, tableAlias, fieldAlias)
        Next colIndex
        headerRows.Add rowIndex
        rowIndex = rowIndex + 1

        Set resultTable = mp_CreateResultTableFromFields(sourceAlias, tableAlias, fields)
        resultTables.Add resultTable
        resultTableRefs.Add resultTable.TableRef

        Set conn = mp_OpenSourceConnection(cfg, sourceAlias)
        likeDialect = LIKE_DIALECT_UNKNOWN
        If useLikeMatch Then
            likeDialect = mp_GetLikeDialectForConnection(conn)
        End If
        tableRef = mp_BuildTableRef(sourceAlias, tableAlias, cfg)

        Set schemaRs = CreateObject("ADODB.Recordset")
        schemaRs.Open "SELECT * FROM " & tableRef & " WHERE 1=0", conn, 0, 1

        If Not isExplicitSheetRange Then
            expectedHeaders = mp_GetExpectedMappedHeaders(cfg, sourceAlias, tableAlias, fields)
            keySourceHeader = mp_MappedHeader(cfg, sourceAlias, tableAlias, keyFieldAlias)
            If mp_RecordsetLooksLikeGenericFields(schemaRs) Or (mp_RecordsetGetFieldOrdinal(schemaRs, keySourceHeader) < 0) Then
                If mp_TryDetectHeaderRangeFromTopRows(conn, tableRef, expectedHeaders, keySourceHeader, detectedTableRef) Then
                    schemaRs.Close
                    tableRef = detectedTableRef
                    schemaRs.Open "SELECT * FROM " & tableRef & " WHERE 1=0", conn, 0, 1
                    sheetName = mp_UnquoteSqlIdentifier(tableRef)
                    summarySheet.Cells(rowIndex - 2, 1).Value = mp_BuildSectionTitle(i, rawToken, sourceFilePath, sheetName)
                End If
            End If
        End If

        Set fieldOrdinals = mp_BuildFieldOrdinals(cfg, schemaRs, sourceAlias, tableAlias, fields)
        keySourceFieldName = mp_RecordsetFieldNameByOrdinal(schemaRs, CLng(fieldOrdinals(keyFieldAlias)), sourceAlias, tableAlias, keyFieldAlias)
        schemaRs.Close
        Set schemaRs = Nothing

        selectSql = mp_BuildFilteredSelectSql(tableRef, keySourceFieldName, commonKey, useLikeMatch, likeDialect)
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open selectSql, conn, 0, 1

        hasRows = mp_AppendFilteredRows( _
            summarySheet, _
            rs, _
            rowIndex, _
            contentRows, _
            resultTable, _
            sourceAlias, _
            tableAlias, _
            fields, _
            fieldOrdinals, _
            keyFieldAlias, _
            commonKey, _
            useLikeMatch, _
            cfg, _
            sheetName, _
            wbCache, _
            longTextRuntimeCache)
        If hasRows Then
            rowIndex = mp_GetNextOutputRow(summarySheet)
        Else
            summarySheet.Cells(rowIndex, 1).Value = "(no rows)"
            contentRows.Add rowIndex
            rowIndex = rowIndex + 1
        End If
        tableVisualEndRow = rowIndex - 1
        If tableVisualEndRow < tableVisualStartRow Then tableVisualEndRow = tableVisualStartRow
        ex_Messaging.m_RegisterResultTableAnchor summarySheet, resultTable.TableRef, tableVisualStartRow, tableVisualEndRow

        rs.Close
        conn.Close
        Set rs = Nothing
        Set conn = Nothing

        ' Inter-table spacing is controlled by ResultLayout script.
    Next i

    Set rowKindRanges = CreateObject("Scripting.Dictionary")
    rowKindRanges.CompareMode = 1
    Set rowKindRanges("header") = headerRows
    Set rowKindRanges("section") = sectionRows
    Set rowKindRanges("content") = contentRows

    If Not useResultLayoutScript Then
        mp_ApplySheetPipelineForPage summarySheet, "MultiSources", SUMMARY_SHEET_NAME, rowKindRanges
        If hasOutputStyle Then
            ex_OutputPanel.m_RenderForSheet summarySheet, outputStyle
        End If
    End If
    summarySheet.Activate

    Set modeResult = CreateObject("Scripting.Dictionary")
    modeResult.CompareMode = 1
    Set modeResult("Output") = modeInput
    Set modeResult("Worksheet") = summarySheet
    Set modeResult("ResultTables") = resultTables
    ex_ScriptIO.m_SetObject modeInput, SCRIPT_INPUT_RESULT_TABLES_KEY, resultTables
    ex_ScriptIO.m_SetObject modeInput, "__ResultTableRefs", resultTableRefs

    mp_CloseWorkbooks wbCache

    Set m_RunMode = modeResult
    Exit Function

EH:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    If errNumber = 0 Then errNumber = vbObjectError + 6520
    If Len(errSource) = 0 Then errSource = "ex_ModeMultiSources"
    If Len(errDescription) = 0 Then errDescription = "Mode run failed."

    On Error Resume Next
    If Not schemaRs Is Nothing Then
        If schemaRs.State <> 0 Then schemaRs.Close
    End If
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    If Not conn Is Nothing Then
        If conn.State <> 0 Then conn.Close
    End If
    mp_CloseWorkbooks wbCache
    On Error GoTo 0

    On Error Resume Next
    Set m_RunMode = mp_BuildFailureModeResult(modeInput, errNumber, errSource, errDescription)
    On Error GoTo 0
End Function

Private Function mp_ResolveQueryOutputEntries( _
    ByVal cfg As Object, _
    ByVal modeInput As Object, _
    ByVal preProcessContext As Object _
) As Collection
    Dim tableRefsText As String
    Dim hasPreProcessScript As Boolean

    hasPreProcessScript = mp_PreProcessHasScript(preProcessContext)
    tableRefsText = Trim$(ex_ScriptIO.m_GetStringOrDefault(modeInput, INPUT_KEY_QUERY_TABLE_REFS, vbNullString))

    If Len(tableRefsText) = 0 Then
        If hasPreProcessScript Then
            Err.Raise vbObjectError + 6542, "ex_ModeMultiSources", _
                "PreProcess script must explicitly pass or set input key 'Query.TableRefs'."
        End If

        tableRefsText = ex_ConfigVirtualSources.m_BuildAllTableRefsText(cfg, "ex_ModeMultiSources")
        tableRefsText = Trim$(tableRefsText)
        If Len(tableRefsText) = 0 Then
            Err.Raise vbObjectError + 6543, "ex_ModeMultiSources", _
                "Failed to build default Query.TableRefs from Source.*.SheetAliases."
        End If
        ex_ScriptIO.m_SetString modeInput, INPUT_KEY_QUERY_TABLE_REFS, tableRefsText
    End If

    Set mp_ResolveQueryOutputEntries = ex_ConfigVirtualSources.m_BuildOutputEntriesFromTableRefs( _
        cfg, tableRefsText, "ex_ModeMultiSources")
End Function

Private Function mp_PreProcessHasScript(ByVal preProcessContext As Object) As Boolean
    If preProcessContext Is Nothing Then Exit Function
    mp_PreProcessHasScript = (StrComp( _
        ex_ScriptIO.m_GetStringOrDefault(preProcessContext, PREPROCESS_CONTEXT_HAS_SCRIPT, "false"), _
        "true", _
        vbTextCompare) = 0)
End Function

Private Sub mp_RestoreExcelUiStateSafe()
    On Error Resume Next
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

Private Function mp_BuildFailureModeResult( _
    ByVal modeInput As Object, _
    ByVal errNumber As Long, _
    ByVal errSource As String, _
    ByVal errDescription As String _
) As Object
    Dim modeResult As Object
    Dim resultTables As Collection
    Dim resultTableRefs As Collection
    Dim summarySheet As Worksheet
    Dim outputObj As Object
    Dim hasOutputStyle As Boolean
    Dim outputStyle As ex_SheetStylesXmlProvider.t_OutputSheetStyle
    Dim summaryTopRow As Long

    Set summarySheet = mp_CreateOrClearSheet(SUMMARY_SHEET_NAME)
    ex_Messaging.m_ClearResultTableAnchors summarySheet
    ex_Messaging.m_ClearResultRowAnchors summarySheet

    hasOutputStyle = ex_SheetStylesXmlProvider.m_GetOutputSheetStyle(outputStyle, ThisWorkbook)
    summaryTopRow = 1
    If hasOutputStyle Then
        summaryTopRow = ex_SheetStylesXmlProvider.m_GetOutputViewStartRow(ThisWorkbook)
    End If

    summarySheet.Cells(summaryTopRow, 1).Value = "MultiSources"
    summarySheet.Cells(summaryTopRow, 2).Value = "Failed"
    summarySheet.Cells(summaryTopRow + 2, 1).Value = "ErrorSource"
    summarySheet.Cells(summaryTopRow + 2, 2).Value = errSource
    summarySheet.Cells(summaryTopRow + 3, 1).Value = "ErrorCode"
    summarySheet.Cells(summaryTopRow + 3, 2).Value = errNumber
    summarySheet.Cells(summaryTopRow + 4, 1).Value = "ErrorDescription"
    summarySheet.Cells(summaryTopRow + 4, 2).Value = errDescription

    mp_ApplySheetPipelineForPage summarySheet, "MultiSources", SUMMARY_SHEET_NAME
    If hasOutputStyle Then
        ex_OutputPanel.m_RenderForSheet summarySheet, outputStyle
    End If
    summarySheet.Activate

    Set modeResult = CreateObject("Scripting.Dictionary")
    modeResult.CompareMode = 1
    Set resultTables = New Collection
    Set resultTableRefs = New Collection

    If modeInput Is Nothing Then
        Set outputObj = New obj_ScriptIOPayload
    Else
        Set outputObj = modeInput
    End If

    Set modeResult("Output") = outputObj
    Set modeResult("Worksheet") = summarySheet
    Set modeResult("ResultTables") = resultTables
    ex_ScriptIO.m_SetObject outputObj, SCRIPT_INPUT_RESULT_TABLES_KEY, resultTables
    ex_ScriptIO.m_SetObject outputObj, "__ResultTableRefs", resultTableRefs

    Set mp_BuildFailureModeResult = modeResult
End Function

Private Sub mp_ApplySheetPipelineForPage( _
    ByVal ws As Worksheet, _
    ByVal modeKey As String, _
    ByVal pipelinePageName As String, _
    Optional ByVal rowKindRanges As Object = Nothing _
)
    Dim pipeline As Collection

    If ws Is Nothing Then Exit Sub

    On Error GoTo SoftFail
    Set pipeline = ex_StylePipelineEngine.m_BuildColumnStylesPipeline( _
        Nothing, _
        Nothing, _
        Trim$(modeKey), _
        ThisWorkbook, _
        Trim$(pipelinePageName) _
    )
    ex_StylePipelineEngine.m_ApplyColumnStylesPipeline ws, Nothing, pipeline, Trim$(modeKey), rowKindRanges
    On Error GoTo 0
    Exit Sub

SoftFail:
    On Error GoTo 0
End Sub

Private Function mp_BuildSectionTitle( _
    ByVal ordinal As Long, _
    ByVal rawToken As String, _
    ByVal sourceFilePath As String, _
    ByVal sheetName As String _
) As String
    rawToken = Trim$(rawToken)
    If Len(rawToken) > 0 Then
        mp_BuildSectionTitle = rawToken
    Else
        mp_BuildSectionTitle = "Entry #" & CStr(ordinal)
    End If
End Function

Private Function mp_GetFieldsAliases(ByVal cfg As Object, ByVal sourceAlias As String, ByVal tableAlias As String) As Collection
    Dim rawList As String

    rawList = mp_GetCfgRequired(cfg, sourceAlias & ".Sheet[" & tableAlias & "].FieldsAliases")
    Set mp_GetFieldsAliases = mp_ParseSemicolonList(rawList)
    If mp_GetFieldsAliases Is Nothing Or mp_GetFieldsAliases.Count = 0 Then
        Err.Raise vbObjectError + 6521, "ex_ModeMultiSources", _
            "Config key '" & sourceAlias & ".Sheet[" & tableAlias & "].FieldsAliases' is empty."
    End If
End Function

Private Function mp_ParseSemicolonList(ByVal rawText As String) As Collection
    Dim result As Collection
    Dim seen As Object
    Dim parts As Variant
    Dim i As Long
    Dim token As String

    rawText = Trim$(rawText)
    If Len(rawText) = 0 Then Exit Function

    Set result = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1

    parts = Split(rawText, ";")
    For i = LBound(parts) To UBound(parts)
        token = Trim$(CStr(parts(i)))
        If Len(token) > 0 Then
            If Not seen.Exists(token) Then
                seen(token) = True
                result.Add token
            End If
        End If
    Next i

    If result.Count > 0 Then Set mp_ParseSemicolonList = result
End Function

Private Function mp_CollectionContainsText(ByVal values As Collection, ByVal needle As String) As Boolean
    Dim i As Long

    If values Is Nothing Then Exit Function
    needle = Trim$(needle)
    If Len(needle) = 0 Then Exit Function

    For i = 1 To values.Count
        If StrComp(CStr(values(i)), needle, vbTextCompare) = 0 Then
            mp_CollectionContainsText = True
            Exit Function
        End If
    Next i
End Function

Private Function mp_GetFieldDisplayHeader( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAlias As String _
) As String
    Dim rawValue As String
    Dim splitPos As Long
    Dim labelText As String

    rawValue = mp_GetCfgRequired(cfg, sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]")
    splitPos = InStr(1, rawValue, "|", vbBinaryCompare)
    If splitPos > 0 Then
        labelText = Trim$(Mid$(rawValue, splitPos + 1))
        If Len(labelText) > 0 Then
            mp_GetFieldDisplayHeader = labelText
            Exit Function
        End If
    End If

    mp_GetFieldDisplayHeader = mp_MappedHeader(cfg, sourceAlias, tableAlias, fieldAlias)
End Function

Private Function mp_BuildSelectFieldsSql( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fields As Collection _
) As String
    Dim i As Long
    Dim fieldAlias As String
    Dim mappedHeader As String

    If fields Is Nothing Or fields.Count = 0 Then
        Err.Raise vbObjectError + 6522, "ex_ModeMultiSources", _
            "Cannot build SQL SELECT for empty fields list at '" & sourceAlias & ".Sheet[" & tableAlias & "]'."
    End If

    For i = 1 To fields.Count
        fieldAlias = CStr(fields(i))
        mappedHeader = mp_MappedHeader(cfg, sourceAlias, tableAlias, fieldAlias)
        If i > 1 Then mp_BuildSelectFieldsSql = mp_BuildSelectFieldsSql & ", "
        mp_BuildSelectFieldsSql = mp_BuildSelectFieldsSql & mp_QuoteIdentifier(mappedHeader) & " AS " & mp_QuoteIdentifier(fieldAlias)
    Next i
End Function

Private Function mp_BuildSortSqlTail( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal sortFieldAlias As String, _
    ByVal keyFieldAlias As String _
) As String
    sortFieldAlias = Trim$(sortFieldAlias)
    If Len(sortFieldAlias) = 0 Then Exit Function
    If StrComp(sortFieldAlias, keyFieldAlias, vbTextCompare) = 0 Then Exit Function
    mp_BuildSortSqlTail = " ORDER BY " & mp_QuoteIdentifier(mp_MappedHeader(cfg, sourceAlias, tableAlias, sortFieldAlias))
End Function

Private Function mp_BuildFilteredSelectSql( _
    ByVal tableRef As String, _
    ByVal keySourceFieldName As String, _
    ByVal keyValue As String, _
    Optional ByVal useLike As Boolean = False, _
    Optional ByVal likeDialect As String = LIKE_DIALECT_UNKNOWN _
) As String
    mp_BuildFilteredSelectSql = ex_ResultSqlEngine.m_BuildFilteredSelectSql( _
        tableRef, keySourceFieldName, keyValue, useLike, likeDialect)
End Function

Private Function mp_BuildAdoWhereLikePattern( _
    ByVal columnName As String, _
    ByVal patternText As String, _
    Optional ByVal likeDialect As String = LIKE_DIALECT_UNKNOWN _
) As String
    Dim colExpr As String
    Dim primaryPattern As String
    Dim altPattern As String
    Dim normalizedPattern As String

    colExpr = mp_QuoteIdentifier(columnName)
    primaryPattern = Trim$(patternText)
    normalizedPattern = mp_ConvertPatternForLikeDialect(primaryPattern, likeDialect)

    If StrComp(LCase$(Trim$(likeDialect)), LIKE_DIALECT_STAR, vbBinaryCompare) = 0 Or _
       StrComp(LCase$(Trim$(likeDialect)), LIKE_DIALECT_PERCENT, vbBinaryCompare) = 0 Then
        mp_BuildAdoWhereLikePattern = colExpr & " LIKE " & mp_QuoteSqlStringLiteral(normalizedPattern)
        Exit Function
    End If

    altPattern = mp_BuildAlternativeLikePattern(primaryPattern)

    If StrComp(primaryPattern, altPattern, vbBinaryCompare) = 0 Then
        mp_BuildAdoWhereLikePattern = colExpr & " LIKE " & mp_QuoteSqlStringLiteral(primaryPattern)
    Else
        mp_BuildAdoWhereLikePattern = "(" & _
            colExpr & " LIKE " & mp_QuoteSqlStringLiteral(primaryPattern) & _
            " OR " & colExpr & " LIKE " & mp_QuoteSqlStringLiteral(altPattern) & _
            ")"
    End If
End Function

Private Function mp_BuildAlternativeLikePattern(ByVal patternText As String) As String
    Dim hasStarSyntax As Boolean
    Dim hasPercentSyntax As Boolean

    hasStarSyntax = (InStr(1, patternText, "*", vbBinaryCompare) > 0) Or (InStr(1, patternText, "?", vbBinaryCompare) > 0)
    hasPercentSyntax = (InStr(1, patternText, "%", vbBinaryCompare) > 0) Or (InStr(1, patternText, "_", vbBinaryCompare) > 0)

    If hasStarSyntax And Not hasPercentSyntax Then
        mp_BuildAlternativeLikePattern = Replace$(Replace$(patternText, "*", "%"), "?", "_")
        Exit Function
    End If

    If hasPercentSyntax And Not hasStarSyntax Then
        mp_BuildAlternativeLikePattern = Replace$(Replace$(patternText, "%", "*"), "_", "?")
        Exit Function
    End If

    mp_BuildAlternativeLikePattern = patternText
End Function

Private Function mp_ConvertPatternForLikeDialect(ByVal patternText As String, ByVal likeDialect As String) As String
    likeDialect = LCase$(Trim$(likeDialect))

    If StrComp(likeDialect, LIKE_DIALECT_STAR, vbBinaryCompare) = 0 Then
        mp_ConvertPatternForLikeDialect = Replace$(Replace$(patternText, "%", "*"), "_", "?")
        Exit Function
    End If

    If StrComp(likeDialect, LIKE_DIALECT_PERCENT, vbBinaryCompare) = 0 Then
        mp_ConvertPatternForLikeDialect = Replace$(Replace$(patternText, "*", "%"), "?", "_")
        Exit Function
    End If

    mp_ConvertPatternForLikeDialect = patternText
End Function

Private Sub mp_EnsureLikeDialectCache()
    If g_LikeDialectByConnection Is Nothing Then
        Set g_LikeDialectByConnection = CreateObject("Scripting.Dictionary")
        g_LikeDialectByConnection.CompareMode = 1
    End If
End Sub

Private Function mp_GetLikeDialectForConnection(ByVal conn As Object) As String
    mp_GetLikeDialectForConnection = ex_ResultSqlEngine.m_GetLikeDialectForConnection(conn)
End Function

Private Function mp_GetConnectionCacheKey(ByVal conn As Object) As String
    On Error Resume Next
    If conn Is Nothing Then Exit Function
    mp_GetConnectionCacheKey = Trim$(CStr(conn.ConnectionString))
    On Error GoTo 0
End Function

Private Function mp_DetectLikeDialect(ByVal conn As Object) As String
    Dim starOk As Boolean
    Dim starMatches As Boolean
    Dim percentOk As Boolean
    Dim percentMatches As Boolean

    starOk = mp_TryLikeLiteralProbe(conn, "a*c", starMatches)
    percentOk = mp_TryLikeLiteralProbe(conn, "a%c", percentMatches)

    If starOk And percentOk Then
        If starMatches Xor percentMatches Then
            If starMatches Then
                mp_DetectLikeDialect = LIKE_DIALECT_STAR
            Else
                mp_DetectLikeDialect = LIKE_DIALECT_PERCENT
            End If
            Exit Function
        End If
    End If

    If starOk And starMatches Then
        mp_DetectLikeDialect = LIKE_DIALECT_STAR
        Exit Function
    End If

    If percentOk And percentMatches Then
        mp_DetectLikeDialect = LIKE_DIALECT_PERCENT
        Exit Function
    End If

    mp_DetectLikeDialect = LIKE_DIALECT_UNKNOWN
End Function

Private Function mp_TryLikeLiteralProbe(ByVal conn As Object, ByVal patternText As String, ByRef outMatches As Boolean) As Boolean
    Dim rs As Object
    Dim sqlText As String
    Dim hitValue As Variant

    On Error GoTo EH
    sqlText = "SELECT IIF('abc' LIKE " & mp_QuoteSqlStringLiteral(patternText) & ", 1, 0) AS Hit"
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sqlText, conn, 0, 1
    If rs.EOF Then GoTo ProbeFail

    hitValue = rs.Fields(0).Value
    outMatches = (Val(CStr(hitValue)) <> 0)
    mp_TryLikeLiteralProbe = True

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    On Error GoTo 0
    Exit Function

ProbeFail:
    outMatches = False
    mp_TryLikeLiteralProbe = False
    GoTo Cleanup

EH:
    outMatches = False
    mp_TryLikeLiteralProbe = False
    Resume Cleanup
End Function

Private Function mp_MappedHeader( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAlias As String _
) As String
    mp_MappedHeader = ex_ResultSqlEngine.m_MappedHeader( _
        cfg, _
        sourceAlias, _
        tableAlias, _
        fieldAlias, _
        "ex_ModeMultiSources", _
        vbObjectError + 6523, _
        vbObjectError + 6518, _
        vbObjectError + 6519)
End Function

Private Function mp_GetExpectedMappedHeaders( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fields As Collection _
) As Variant
    mp_GetExpectedMappedHeaders = ex_ResultSqlEngine.m_GetExpectedMappedHeaders( _
        cfg, _
        sourceAlias, _
        tableAlias, _
        fields, _
        "ex_ModeMultiSources", _
        vbObjectError + 6523, _
        vbObjectError + 6518, _
        vbObjectError + 6519)
End Function

Private Function mp_IsExplicitAdoRangeReference(ByVal value As String) As Boolean
    mp_IsExplicitAdoRangeReference = ex_ResultSqlEngine.m_IsExplicitAdoRangeReference(value)
End Function

Private Function mp_UnquoteSqlIdentifier(ByVal value As String) As String
    mp_UnquoteSqlIdentifier = ex_ResultSqlEngine.m_UnquoteSqlIdentifier(value)
End Function

Private Function mp_ExtractAdoSheetPrefix(ByVal tableRef As String) As String
    mp_ExtractAdoSheetPrefix = ex_ResultSqlEngine.m_ExtractAdoSheetPrefix(tableRef)
End Function

Private Function mp_BuildNormalizedHeaderTokenSet(ByVal expectedHeaders As Variant, ByVal keyHeader As String) As Object
    Set mp_BuildNormalizedHeaderTokenSet = ex_ResultSqlEngine.m_BuildNormalizedHeaderTokenSet(expectedHeaders, keyHeader)
End Function

Private Function mp_TryDetectHeaderRangeFromTopRows( _
    ByVal adoConn As Object, _
    ByVal tableRef As String, _
    ByVal expectedHeaders As Variant, _
    ByVal keyHeader As String, _
    ByRef outDetectedRef As String _
) As Boolean
    mp_TryDetectHeaderRangeFromTopRows = ex_ResultSqlEngine.m_TryDetectHeaderRangeFromTopRows( _
        adoConn, tableRef, expectedHeaders, keyHeader, outDetectedRef)
End Function

Private Function mp_TryBuildValidatedHeaderRangeRef( _
    ByVal adoConn As Object, _
    ByVal sheetPrefix As String, _
    ByVal headerRowAbs As Long, _
    ByVal colLetter As String, _
    ByVal keyHeader As String, _
    ByRef outRangeRef As String _
) As Boolean
    mp_TryBuildValidatedHeaderRangeRef = ex_ResultSqlEngine.m_TryBuildValidatedHeaderRangeRef( _
        adoConn, sheetPrefix, headerRowAbs, colLetter, keyHeader, outRangeRef)
End Function

Private Function mp_OpenSourceConnection(ByVal cfg As Object, ByVal sourceAlias As String) As Object
    Set mp_OpenSourceConnection = ex_ResultSqlEngine.m_OpenSourceConnection( _
        cfg, _
        sourceAlias, _
        "ex_ModeMultiSources", _
        vbObjectError + 6524, _
        vbObjectError + 6525, _
        vbObjectError + 6526, _
        vbObjectError + 6515, _
        vbObjectError + 6516, _
        vbObjectError + 6517, _
        vbObjectError + 6518, _
        vbObjectError + 6519)
End Function

Private Function mp_BuildTableRef(ByVal sourceAlias As String, ByVal tableAlias As String, ByVal cfg As Object) As String
    Dim sheetName As String

    sheetName = ex_ConfigProvider.m_GetResolvedSheetName(sourceAlias, tableAlias, cfg, True, "ex_ModeMultiSources", 6530, 6531, 6532, 6533, 6534)
    mp_BuildTableRef = ex_ResultSqlEngine.m_BuildTableRefFromSheetName(sheetName, "ex_ModeMultiSources", vbObjectError + 6535)
End Function

Private Function mp_QuoteIdentifier(ByVal valueText As String) As String
    mp_QuoteIdentifier = ex_ResultSqlEngine.m_QuoteIdentifier(valueText)
End Function

Private Function mp_QuoteSqlStringLiteral(ByVal valueText As String) As String
    mp_QuoteSqlStringLiteral = ex_ResultSqlEngine.m_QuoteSqlStringLiteral(valueText)
End Function

Private Function mp_BuildAdoConnectionString(ByVal sourcePath As String) As String
    mp_BuildAdoConnectionString = ex_ResultSqlEngine.m_BuildAdoConnectionString(sourcePath, "ex_ModeMultiSources", vbObjectError + 6526)
End Function

Private Function mp_BuildFieldOrdinals( _
    ByVal cfg As Object, _
    ByVal rs As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fields As Collection _
) As Object
    Set mp_BuildFieldOrdinals = ex_ResultSqlEngine.m_BuildFieldOrdinals( _
        cfg, _
        rs, _
        sourceAlias, _
        tableAlias, _
        fields, _
        "ex_ModeMultiSources", _
        vbObjectError + 6536, _
        vbObjectError + 6523, _
        vbObjectError + 6518, _
        vbObjectError + 6519)
End Function

Private Function mp_RecordsetFieldNameByOrdinal( _
    ByVal rs As Object, _
    ByVal fieldOrdinal As Long, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAlias As String _
) As String
    mp_RecordsetFieldNameByOrdinal = ex_ResultSqlEngine.m_RecordsetFieldNameByOrdinal( _
        rs, _
        fieldOrdinal, _
        sourceAlias, _
        tableAlias, _
        fieldAlias, _
        "ex_ModeMultiSources", _
        vbObjectError + 6538, _
        vbObjectError + 6539)
End Function

Private Function mp_AppendFilteredRows( _
    ByVal ws As Worksheet, _
    ByVal rs As Object, _
    ByVal startRow As Long, _
    ByVal contentRows As Collection, _
    ByVal resultTable As obj_ResultTable, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fields As Collection, _
    ByVal fieldOrdinals As Object, _
    ByVal keyFieldAlias As String, _
    ByVal keyValue As String, _
    Optional ByVal useLike As Boolean = False, _
    Optional ByVal cfg As Object = Nothing, _
    Optional ByVal configuredSheetName As String = vbNullString, _
    Optional ByVal wbCache As Object = Nothing, _
    Optional ByVal runtimeCache As Object = Nothing _
) As Boolean
    mp_AppendFilteredRows = ex_ResultSqlEngine.m_AppendFilteredRows( _
        ws, _
        rs, _
        startRow, _
        contentRows, _
        resultTable, _
        sourceAlias, _
        tableAlias, _
        fields, _
        fieldOrdinals, _
        keyFieldAlias, _
        keyValue, _
        useLike, _
        cfg, _
        configuredSheetName, _
        wbCache, _
        runtimeCache)
End Function

Private Sub mp_CloseWorkbooks(ByVal wbCache As Object)
    Dim key As Variant
    Dim wb As Workbook

    If wbCache Is Nothing Then Exit Sub

    On Error Resume Next
    For Each key In wbCache.Keys
        Set wb = wbCache(key)
        If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Next key
    wbCache.RemoveAll
    On Error GoTo 0
End Sub

Private Function mp_AsText(ByVal valueIn As Variant, Optional ByVal adoFieldType As Long = -1) As String
    mp_AsText = ex_SqlAdoHelpers.m_ToNormalizedText(valueIn, adoFieldType)
End Function

Private Function mp_ListRecordsetFields(ByVal rs As Object, Optional ByVal maxCount As Long = 25) As String
    mp_ListRecordsetFields = ex_ResultSqlEngine.m_ListRecordsetFields(rs, maxCount)
End Function

Private Function mp_RecordsetLooksLikeGenericFields(ByVal rs As Object) As Boolean
    mp_RecordsetLooksLikeGenericFields = ex_ResultSqlEngine.m_RecordsetLooksLikeGenericFields(rs)
End Function

Private Function mp_RecordsetGetFieldOrdinal(ByVal rs As Object, ByVal fieldName As String) As Long
    mp_RecordsetGetFieldOrdinal = ex_ResultSqlEngine.m_RecordsetGetFieldOrdinal(rs, fieldName)
End Function

Private Function mp_ToSafeText(ByVal valueIn As Variant) As String
    mp_ToSafeText = ex_ResultSqlEngine.m_ToSafeText(valueIn)
End Function

Private Function mp_IsEmptyVariantArray(ByVal valueRef As Variant) As Boolean
    mp_IsEmptyVariantArray = ex_ResultSqlEngine.m_IsEmptyVariantArray(valueRef)
End Function

Private Function mp_ToColumnLetter(ByVal columnIndex As Long) As String
    mp_ToColumnLetter = ex_ResultSqlEngine.m_ToColumnLetter(columnIndex)
End Function

Private Function mp_NormalizeHeader(ByVal valueText As String) As String
    mp_NormalizeHeader = ex_ResultSqlEngine.m_NormalizeHeader(valueText)
End Function

Private Function mp_NormalizeHeaderLoose(ByVal valueText As String) As String
    mp_NormalizeHeaderLoose = ex_ResultSqlEngine.m_NormalizeHeaderLoose(valueText)
End Function

Private Function mp_GetNextOutputRow(ByVal ws As Worksheet) As Long
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then
        mp_GetNextOutputRow = 1
    Else
        mp_GetNextOutputRow = lastRow + 1
    End If
End Function

Private Function mp_CreateResultTableFromFields( _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fields As Collection _
) As obj_ResultTable
    Set mp_CreateResultTableFromFields = ex_ResultSqlEngine.m_CreateResultTableFromFields(sourceAlias, tableAlias, fields)
End Function

Private Sub mp_AddResultCell( _
    ByVal resultTable As obj_ResultTable, _
    ByVal rowIndex As Long, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAlias As String, _
    ByVal valueText As String _
)
    ex_ResultSqlEngine.m_AddResultCell resultTable, rowIndex, sourceAlias, tableAlias, fieldAlias, valueText
End Sub

Private Function mp_CreateScriptInputContext(ByVal fio As String) As Object
    Dim payload As obj_ScriptIOPayload

    Set payload = New obj_ScriptIOPayload
    payload.m_SetString "CommonKey", Trim$(fio)

    Set mp_CreateScriptInputContext = payload
End Function

Private Sub mp_RenderEntrySheet( _
    ByVal ws As Worksheet, _
    ByVal ordinal As Long, _
    ByVal commonKey As String, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal rawToken As String, _
    ByVal sourceFilePath As String, _
    ByVal sheetName As String _
)
    ws.Cells(1, 1).Value = "MultiSources Result"
    ws.Cells(1, 2).Value = "Entry #" & CStr(ordinal)

    ws.Cells(3, 1).Value = "CommonKey"
    ws.Cells(3, 2).Value = commonKey

    ws.Cells(5, 1).Value = "SourceAlias"
    ws.Cells(5, 2).Value = sourceAlias

    ws.Cells(6, 1).Value = "TableAlias"
    ws.Cells(6, 2).Value = tableAlias

    ws.Cells(7, 1).Value = "OutputToken"
    ws.Cells(7, 2).Value = rawToken

    ws.Cells(8, 1).Value = "SourceFile"
    ws.Cells(8, 2).Value = sourceFilePath

    ws.Cells(9, 1).Value = "SheetName"
    ws.Cells(9, 2).Value = sheetName

    If Len(sourceFilePath) > 0 Then
        ws.Cells(11, 1).Value = "FileExists"
        ws.Cells(11, 2).Value = IIf(Len(Dir$(sourceFilePath)) > 0, "true", "false")
    End If
End Sub

Private Sub mp_DeleteGeneratedResultSheets()
    Dim namesToDelete As Collection
    Dim ws As Worksheet
    Dim sheetName As Variant

    Set namesToDelete = New Collection

    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, SUMMARY_SHEET_NAME, vbTextCompare) = 0 Then
            namesToDelete.Add ws.Name
        ElseIf StrComp(Left$(ws.Name, Len(RESULT_SHEET_PREFIX)), RESULT_SHEET_PREFIX, vbTextCompare) = 0 Then
            namesToDelete.Add ws.Name
        End If
    Next ws

    Application.DisplayAlerts = False
    On Error Resume Next
    For Each sheetName In namesToDelete
        ThisWorkbook.Worksheets(CStr(sheetName)).Delete
    Next sheetName
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

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

Private Function mp_BuildEntrySheetName(ByVal ordinal As Long, ByVal sourceAlias As String, ByVal tableAlias As String) As String
    Dim rawName As String

    rawName = RESULT_SHEET_PREFIX & Format$(ordinal, "000") & "_" & sourceAlias & "_" & tableAlias
    mp_BuildEntrySheetName = mp_NormalizeSheetName(rawName)
End Function

Private Function mp_NormalizeSheetName(ByVal rawName As String) As String
    Dim normalized As String
    Dim i As Long
    Dim ch As String

    normalized = Trim$(rawName)
    If Len(normalized) = 0 Then normalized = RESULT_SHEET_PREFIX & "Result"

    For i = 1 To Len(normalized)
        ch = Mid$(normalized, i, 1)
        Select Case ch
            Case "[", "]", ":", "*", "?", "/", "\"
                Mid$(normalized, i, 1) = "_"
        End Select
    Next i

    If Len(normalized) > 31 Then
        normalized = Left$(normalized, 31)
    End If

    If Len(Trim$(normalized)) = 0 Then
        normalized = RESULT_SHEET_PREFIX & "Result"
    End If

    mp_NormalizeSheetName = normalized
End Function

Private Function mp_GetResolvedSourcePath(ByVal cfg As Object, ByVal sourceAlias As String) As String
    mp_GetResolvedSourcePath = ex_ResultSqlEngine.m_GetResolvedSourcePath( _
        cfg, _
        sourceAlias, _
        "ex_ModeMultiSources", _
        vbObjectError + 6515, _
        vbObjectError + 6516, _
        vbObjectError + 6517, _
        vbObjectError + 6518, _
        vbObjectError + 6519)
End Function

Private Function mp_GetCfgRequired(ByVal cfg As Object, ByVal keyName As String) As String
    mp_GetCfgRequired = ex_ResultSqlEngine.m_GetCfgRequired( _
        cfg, _
        keyName, _
        "ex_ModeMultiSources", _
        vbObjectError + 6518, _
        vbObjectError + 6519)
End Function

Private Function mp_GetCfgOptional(ByVal cfg As Object, ByVal keyName As String, Optional ByVal defaultValue As String = vbNullString) As String
    If cfg Is Nothing Then
        mp_GetCfgOptional = defaultValue
        Exit Function
    End If

    If Not cfg.Exists(keyName) Then
        mp_GetCfgOptional = defaultValue
        Exit Function
    End If

    mp_GetCfgOptional = Trim$(CStr(cfg(keyName)))
End Function

Private Function mp_HasPlaceholderTokens(ByVal valueText As String) As Boolean
    mp_HasPlaceholderTokens = ex_ResultSqlEngine.m_HasPlaceholderTokens(valueText)
End Function

Private Function mp_ResolvePathLocal(ByVal inputPath As String) As String
    mp_ResolvePathLocal = ex_ResultSqlEngine.m_ResolvePathLocal(inputPath)
End Function
