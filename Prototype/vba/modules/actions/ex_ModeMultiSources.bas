Attribute VB_Name = "ex_ModeMultiSources"
Option Explicit

Private Const SUMMARY_SHEET_NAME As String = "g_MultiSources"
Private Const RESULT_SHEET_PREFIX As String = "g_MS_"
Private Const SCRIPT_INPUT_RESULT_TABLES_KEY As String = "__ResultTables"
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
    Dim gapRows As Long
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
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo EH

    If cfg Is Nothing Then
        Err.Raise vbObjectError + 6503, "ex_ModeMultiSources", "Config dictionary is not initialized."
    End If

    ex_ConfigVirtualSources.m_ExpandVirtualSourcesAndOutput cfg, "ex_ModeMultiSources"
    Set outputEntries = ex_ConfigVirtualSources.m_BuildOutputEntries(cfg, "ex_ModeMultiSources")
    If outputEntries Is Nothing Or outputEntries.Count = 0 Then
        Err.Raise vbObjectError + 6504, "ex_ModeMultiSources", "List is empty for config key: Output.Sheets"
    End If

    commonKey = Trim$(ex_ScriptIO.m_GetStringOrDefault(modeInput, "CommonKey", vbNullString))
    If Len(commonKey) = 0 Then
        Err.Raise vbObjectError + 6505, "ex_ModeMultiSources", "Config key 'CommonKey' is empty."
    End If
    commonKeyType = LCase$(Trim$(ex_ConfigProvider.m_GetConfigEntryType("CommonKey", vbNullString)))
    useLikeMatch = (StrComp(commonKeyType, "rx", vbTextCompare) = 0)

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

    gapRows = mp_GetGapRows(cfg)
    rowIndex = summaryTopRow
    For i = 1 To outputEntries.Count
        tableVisualStartRow = rowIndex
        Set entry = outputEntries(i)
        sourceAlias = Trim$(CStr(entry("SourceAlias")))
        tableAlias = Trim$(CStr(entry("TableAlias")))
        rawToken = Trim$(CStr(entry("RawToken")))

        sourceFilePath = mp_ResolvePathLocal(mp_GetCfgRequired(cfg, "Source." & sourceAlias & ".FilePath"))
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

        hasRows = mp_AppendFilteredRows(summarySheet, rs, rowIndex, contentRows, resultTable, sourceAlias, tableAlias, fields, fieldOrdinals, keyFieldAlias, commonKey, useLikeMatch)
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

        If i < outputEntries.Count Then
            rowIndex = rowIndex + gapRows
        End If
    Next i

    Set rowKindRanges = CreateObject("Scripting.Dictionary")
    rowKindRanges.CompareMode = 1
    Set rowKindRanges("header") = headerRows
    Set rowKindRanges("section") = sectionRows
    Set rowKindRanges("content") = contentRows

    mp_ApplySheetPipelineForPage summarySheet, "MultiSources", SUMMARY_SHEET_NAME, rowKindRanges
    If hasOutputStyle Then
        ex_OutputPanel.m_RenderForSheet summarySheet, outputStyle
    End If
    summarySheet.Activate

    Set modeResult = CreateObject("Scripting.Dictionary")
    modeResult.CompareMode = 1
    Set modeResult("Output") = modeInput
    Set modeResult("Worksheet") = summarySheet
    Set modeResult("ResultTables") = resultTables
    ex_ScriptIO.m_SetObject modeInput, SCRIPT_INPUT_RESULT_TABLES_KEY, resultTables
    ex_ScriptIO.m_SetObject modeInput, "__ResultTableRefs", resultTableRefs

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
    On Error GoTo 0

    On Error Resume Next
    Set m_RunMode = mp_BuildFailureModeResult(modeInput, errNumber, errSource, errDescription)
    On Error GoTo 0
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

Private Function mp_GetGapRows(ByVal cfg As Object) As Long
    Dim rawValue As String

    rawValue = mp_GetCfgOptional(cfg, "Output.Layout.Gap.Default", "1")
    If Len(rawValue) = 0 Then
        mp_GetGapRows = 1
        Exit Function
    End If
    If Not IsNumeric(rawValue) Then
        mp_GetGapRows = 1
        Exit Function
    End If

    mp_GetGapRows = CLng(rawValue)
    If mp_GetGapRows < 0 Then mp_GetGapRows = 0
End Function

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
    If useLike Then
        mp_BuildFilteredSelectSql = _
            "SELECT * FROM " & tableRef & _
            " WHERE " & mp_BuildAdoWhereLikePattern(keySourceFieldName, Trim$(keyValue), likeDialect)
        Exit Function
    End If

    mp_BuildFilteredSelectSql = _
        "SELECT * FROM " & tableRef & _
        " WHERE " & mp_QuoteIdentifier(keySourceFieldName) & " = " & mp_QuoteSqlStringLiteral(Trim$(keyValue))
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
    Dim cacheKey As String
    Dim detected As String

    mp_EnsureLikeDialectCache
    cacheKey = mp_GetConnectionCacheKey(conn)
    If Len(cacheKey) > 0 Then
        If g_LikeDialectByConnection.Exists(cacheKey) Then
            mp_GetLikeDialectForConnection = CStr(g_LikeDialectByConnection(cacheKey))
            Exit Function
        End If
    End If

    detected = mp_DetectLikeDialect(conn)
    If Len(cacheKey) > 0 Then g_LikeDialectByConnection(cacheKey) = detected
    mp_GetLikeDialectForConnection = detected
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
    Dim rawValue As String
    Dim splitPos As Long

    rawValue = mp_GetCfgRequired(cfg, sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]")
    splitPos = InStr(1, rawValue, "|", vbBinaryCompare)
    If splitPos > 0 Then
        mp_MappedHeader = Trim$(Left$(rawValue, splitPos - 1))
    Else
        mp_MappedHeader = Trim$(rawValue)
    End If

    If Len(mp_MappedHeader) = 0 Then
        Err.Raise vbObjectError + 6523, "ex_ModeMultiSources", _
            "Mapped source header is empty for " & sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]"
    End If
End Function

Private Function mp_GetExpectedMappedHeaders( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fields As Collection _
) As Variant
    Dim arr() As String
    Dim i As Long
    Dim headerText As String

    If fields Is Nothing Then
        mp_GetExpectedMappedHeaders = Array()
        Exit Function
    End If
    If fields.Count = 0 Then
        mp_GetExpectedMappedHeaders = Array()
        Exit Function
    End If

    ReDim arr(0 To fields.Count - 1)
    For i = 1 To fields.Count
        headerText = mp_MappedHeader(cfg, sourceAlias, tableAlias, CStr(fields(i)))
        arr(i - 1) = headerText
    Next i
    mp_GetExpectedMappedHeaders = arr
End Function

Private Function mp_IsExplicitAdoRangeReference(ByVal value As String) As Boolean
    value = Trim$(value)
    If InStr(1, value, "$", vbBinaryCompare) <= 0 Then Exit Function
    If InStr(1, value, ":", vbBinaryCompare) <= 0 Then Exit Function
    mp_IsExplicitAdoRangeReference = True
End Function

Private Function mp_UnquoteSqlIdentifier(ByVal value As String) As String
    value = Trim$(value)
    If Len(value) >= 2 Then
        If Left$(value, 1) = "[" And Right$(value, 1) = "]" Then
            value = Mid$(value, 2, Len(value) - 2)
        End If
    End If
    mp_UnquoteSqlIdentifier = Replace$(value, "]]", "]")
End Function

Private Function mp_ExtractAdoSheetPrefix(ByVal tableRef As String) As String
    Dim objectName As String
    Dim dollarPos As Long

    objectName = mp_UnquoteSqlIdentifier(tableRef)
    If Len(objectName) = 0 Then Exit Function

    dollarPos = InStr(1, objectName, "$", vbBinaryCompare)
    If dollarPos <= 0 Then Exit Function

    mp_ExtractAdoSheetPrefix = Left$(objectName, dollarPos)
End Function

Private Function mp_BuildNormalizedHeaderTokenSet(ByVal expectedHeaders As Variant, ByVal keyHeader As String) As Object
    Dim d As Object
    Dim i As Long
    Dim token As String

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    token = mp_NormalizeHeader(keyHeader)
    If Len(token) > 0 Then d(token) = True

    If Not mp_IsEmptyVariantArray(expectedHeaders) Then
        For i = LBound(expectedHeaders) To UBound(expectedHeaders)
            token = mp_NormalizeHeader(CStr(expectedHeaders(i)))
            If Len(token) > 0 Then d(token) = True
        Next i
    End If

    Set mp_BuildNormalizedHeaderTokenSet = d
End Function

Private Function mp_TryDetectHeaderRangeFromTopRows( _
    ByVal adoConn As Object, _
    ByVal tableRef As String, _
    ByVal expectedHeaders As Variant, _
    ByVal keyHeader As String, _
    ByRef outDetectedRef As String _
) As Boolean
    Const MAX_HEADER_ALIGNMENT_SHIFT As Long = 20
    Dim sheetPrefix As String
    Dim probeRef As String
    Dim rs As Object
    Dim rowsData As Variant
    Dim rowLower As Long
    Dim rowUpper As Long
    Dim fieldLower As Long
    Dim fieldUpper As Long
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim bestRowIndex As Long
    Dim bestScore As Long
    Dim bestLastCol As Long
    Dim rowTokens As Object
    Dim expectedSet As Object
    Dim keyToken As String
    Dim cellText As String
    Dim normalized As String
    Dim lastNonEmptyCol As Long
    Dim currentScore As Long
    Dim token As Variant
    Dim headerRowAbs As Long
    Dim colLetter As String
    Dim fallbackRowAbs As Long
    Dim alignmentShift As Long

    sheetPrefix = mp_ExtractAdoSheetPrefix(tableRef)
    If Len(sheetPrefix) = 0 Then Exit Function

    probeRef = "[" & sheetPrefix & "A1:ZZ200]"

    Set expectedSet = mp_BuildNormalizedHeaderTokenSet(expectedHeaders, keyHeader)
    If expectedSet Is Nothing Then Exit Function
    If expectedSet.Count = 0 Then Exit Function

    keyToken = mp_NormalizeHeader(keyHeader)
    If Len(keyToken) = 0 Then Exit Function

    On Error GoTo EH
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM " & probeRef, adoConn, 0, 1
    If rs.EOF Then
        rs.Close
        Exit Function
    End If

    rowsData = rs.GetRows
    rs.Close

    rowLower = LBound(rowsData, 2)
    rowUpper = UBound(rowsData, 2)
    fieldLower = LBound(rowsData, 1)
    fieldUpper = UBound(rowsData, 1)
    bestRowIndex = -1
    bestScore = 0

    For rowIndex = rowLower To rowUpper
        Set rowTokens = CreateObject("Scripting.Dictionary")
        rowTokens.CompareMode = 1
        lastNonEmptyCol = 0

        For colIndex = fieldLower To fieldUpper
            cellText = mp_ToSafeText(rowsData(colIndex, rowIndex))
            normalized = mp_NormalizeHeader(cellText)
            If Len(normalized) > 0 Then
                rowTokens(normalized) = True
                lastNonEmptyCol = (colIndex - fieldLower + 1)
            End If
        Next colIndex

        If rowTokens.Exists(keyToken) Then
            currentScore = 0
            For Each token In expectedSet.Keys
                If rowTokens.Exists(CStr(token)) Then currentScore = currentScore + 1
            Next token

            If currentScore > bestScore Then
                bestScore = currentScore
                bestRowIndex = rowIndex
                bestLastCol = lastNonEmptyCol
            End If
        End If
    Next rowIndex

    If bestRowIndex < 0 Then Exit Function
    If bestLastCol <= 0 Then bestLastCol = (fieldUpper - fieldLower + 1)
    If bestLastCol <= 0 Then Exit Function

    colLetter = mp_ToColumnLetter(bestLastCol)
    If Len(colLetter) = 0 Then Exit Function

    For alignmentShift = 1 To MAX_HEADER_ALIGNMENT_SHIFT
        headerRowAbs = (bestRowIndex - rowLower) + alignmentShift
        If headerRowAbs > 0 Then
            If mp_TryBuildValidatedHeaderRangeRef(adoConn, sheetPrefix, headerRowAbs, colLetter, keyHeader, outDetectedRef) Then
                mp_TryDetectHeaderRangeFromTopRows = True
                Exit Function
            End If
        End If
    Next alignmentShift

    fallbackRowAbs = (bestRowIndex - rowLower) + 1
    If fallbackRowAbs <= 0 Then Exit Function
    outDetectedRef = "[" & sheetPrefix & "A" & CStr(fallbackRowAbs) & ":" & colLetter & "1048576]"
    mp_TryDetectHeaderRangeFromTopRows = True
    Exit Function

EH:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    On Error GoTo 0
End Function

Private Function mp_TryBuildValidatedHeaderRangeRef( _
    ByVal adoConn As Object, _
    ByVal sheetPrefix As String, _
    ByVal headerRowAbs As Long, _
    ByVal colLetter As String, _
    ByVal keyHeader As String, _
    ByRef outRangeRef As String _
) As Boolean
    Dim rs As Object
    Dim candidateRef As String

    If adoConn Is Nothing Then Exit Function
    If headerRowAbs <= 0 Then Exit Function
    If Len(Trim$(sheetPrefix)) = 0 Then Exit Function
    If Len(Trim$(colLetter)) = 0 Then Exit Function
    If Len(Trim$(keyHeader)) = 0 Then Exit Function

    candidateRef = "[" & sheetPrefix & "A" & CStr(headerRowAbs) & ":" & colLetter & "1048576]"

    On Error GoTo EH
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM " & candidateRef & " WHERE 1=0", adoConn, 0, 1
    If mp_RecordsetGetFieldOrdinal(rs, keyHeader) >= 0 Then
        outRangeRef = candidateRef
        mp_TryBuildValidatedHeaderRangeRef = True
    End If
    rs.Close
    Exit Function

EH:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    On Error GoTo 0
End Function

Private Function mp_OpenSourceConnection(ByVal cfg As Object, ByVal sourceAlias As String) As Object
    Dim sourcePath As String
    Dim snapshotPath As String
    Dim conn As Object

    sourcePath = mp_ResolvePathLocal(mp_GetCfgRequired(cfg, "Source." & sourceAlias & ".FilePath"))
    If Len(sourcePath) = 0 Then
        Err.Raise vbObjectError + 6524, "ex_ModeMultiSources", "Source file path is empty for alias '" & sourceAlias & "'."
    End If
    If Dir$(sourcePath) = vbNullString Then
        Err.Raise vbObjectError + 6525, "ex_ModeMultiSources", "Source file not found: " & sourcePath
    End If

    snapshotPath = ex_SourceSnapshot.m_GetSnapshotPath(sourcePath, "Source." & sourceAlias)
    Set conn = CreateObject("ADODB.Connection")
    conn.Open mp_BuildAdoConnectionString(snapshotPath)

    Set mp_OpenSourceConnection = conn
End Function

Private Function mp_BuildTableRef(ByVal sourceAlias As String, ByVal tableAlias As String, ByVal cfg As Object) As String
    Dim sheetName As String

    sheetName = ex_ConfigProvider.m_GetResolvedSheetName(sourceAlias, tableAlias, cfg, True, "ex_ModeMultiSources", 6530, 6531, 6532, 6533, 6534)
    sheetName = Trim$(sheetName)
    If Len(sheetName) = 0 Then
        Err.Raise vbObjectError + 6535, "ex_ModeMultiSources", _
            "Resolved SheetName is empty for " & sourceAlias & ".Sheet[" & tableAlias & "]."
    End If

    If Left$(sheetName, 1) = "[" And Right$(sheetName, 1) = "]" Then
        mp_BuildTableRef = sheetName
        Exit Function
    End If

    If InStr(1, sheetName, "$", vbBinaryCompare) > 0 Then
        mp_BuildTableRef = "[" & Replace$(sheetName, "]", "]]") & "]"
    Else
        mp_BuildTableRef = "[" & Replace$(sheetName, "]", "]]") & "$]"
    End If
End Function

Private Function mp_QuoteIdentifier(ByVal valueText As String) As String
    mp_QuoteIdentifier = "[" & Replace$(Trim$(valueText), "]", "]]") & "]"
End Function

Private Function mp_QuoteSqlStringLiteral(ByVal valueText As String) As String
    mp_QuoteSqlStringLiteral = "'" & Replace$(CStr(valueText), "'", "''") & "'"
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
            Err.Raise vbObjectError + 6526, "ex_ModeMultiSources", _
                "Unsupported source file extension for ADO: ." & ext
    End Select

    mp_BuildAdoConnectionString = _
        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sourcePath & ";Extended Properties=""" & props & """;"
End Function

Private Function mp_BuildFieldOrdinals( _
    ByVal cfg As Object, _
    ByVal rs As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fields As Collection _
) As Object
    Dim byExact As Object
    Dim byLoose As Object
    Dim result As Object
    Dim i As Long
    Dim fieldName As String
    Dim fieldAlias As String
    Dim desiredHeader As String
    Dim exactToken As String
    Dim looseToken As String
    Dim availableFields As String
    Dim hintText As String

    If rs Is Nothing Then Exit Function
    If fields Is Nothing Then Exit Function

    Set byExact = CreateObject("Scripting.Dictionary")
    byExact.CompareMode = 1
    Set byLoose = CreateObject("Scripting.Dictionary")
    byLoose.CompareMode = 1

    For i = 0 To rs.Fields.Count - 1
        fieldName = CStr(rs.Fields(i).Name)
        exactToken = mp_NormalizeHeader(fieldName)
        looseToken = mp_NormalizeHeaderLoose(fieldName)
        If Len(exactToken) > 0 Then
            If Not byExact.Exists(exactToken) Then byExact(exactToken) = i
        End If
        If Len(looseToken) > 0 Then
            If Not byLoose.Exists(looseToken) Then byLoose(looseToken) = i
        End If
    Next i

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1

    For i = 1 To fields.Count
        fieldAlias = CStr(fields(i))
        desiredHeader = mp_MappedHeader(cfg, sourceAlias, tableAlias, fieldAlias)
        exactToken = mp_NormalizeHeader(desiredHeader)
        looseToken = mp_NormalizeHeaderLoose(desiredHeader)

        If Len(exactToken) > 0 And byExact.Exists(exactToken) Then
            result(fieldAlias) = CLng(byExact(exactToken))
        ElseIf Len(looseToken) > 0 And byLoose.Exists(looseToken) Then
            result(fieldAlias) = CLng(byLoose(looseToken))
        Else
            availableFields = mp_ListRecordsetFields(rs, 40)
            If mp_RecordsetLooksLikeGenericFields(rs) Then
                hintText = " Hint: ADO returned generic fields (F1..Fn). Set SheetName as explicit range with header row, e.g. 'Аркуш1$A3:I1048576'."
            End If
            Err.Raise vbObjectError + 6536, "ex_ModeMultiSources", _
                "Configured source header '" & desiredHeader & "' is not found for " & sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]. " & _
                "Available fields: " & availableFields & "." & hintText
        End If
    Next i

    Set mp_BuildFieldOrdinals = result
End Function

Private Function mp_RecordsetFieldNameByOrdinal( _
    ByVal rs As Object, _
    ByVal fieldOrdinal As Long, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAlias As String _
) As String
    If rs Is Nothing Then
        Err.Raise vbObjectError + 6538, "ex_ModeMultiSources", "Recordset is not initialized while resolving source field name."
    End If
    If fieldOrdinal < 0 Or fieldOrdinal >= rs.Fields.Count Then
        Err.Raise vbObjectError + 6539, "ex_ModeMultiSources", _
            "Field ordinal is out of range for " & sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]."
    End If
    mp_RecordsetFieldNameByOrdinal = CStr(rs.Fields(fieldOrdinal).Name)
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
    Optional ByVal useLike As Boolean = False _
) As Boolean
    Dim nextRow As Long
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim fieldAlias As String
    Dim fieldOrdinal As Long
    Dim valueText As String
    Dim keyOrdinal As Long
    Dim currentKey As String
    Dim tableRef As String
    Dim rowObj As obj_ResultRow
    Dim rowAnchorName As String
    Dim isMatch As Boolean

    If rs Is Nothing Then Exit Function
    If fields Is Nothing Then Exit Function
    If fieldOrdinals Is Nothing Then Exit Function
    If fields.Count = 0 Then Exit Function
    If Not fieldOrdinals.Exists(keyFieldAlias) Then Exit Function
    If rs.EOF Then Exit Function

    keyValue = Trim$(keyValue)
    keyOrdinal = CLng(fieldOrdinals(keyFieldAlias))
    nextRow = startRow
    tableRef = sourceAlias & ".Sheet[" & tableAlias & "]"

    Do While Not rs.EOF
        currentKey = mp_AsText(rs.Fields(keyOrdinal).Value, rs.Fields(keyOrdinal).Type)
        isMatch = useLike Or (StrComp(Trim$(currentKey), keyValue, vbTextCompare) = 0)
        If isMatch Then
            rowIndex = resultTable.Count
            For colIndex = 1 To fields.Count
                fieldAlias = CStr(fields(colIndex))
                fieldOrdinal = CLng(fieldOrdinals(fieldAlias))
                valueText = mp_AsText(rs.Fields(fieldOrdinal).Value, rs.Fields(fieldOrdinal).Type)
                ws.Cells(nextRow, colIndex).Value = valueText
                mp_AddResultCell resultTable, rowIndex, sourceAlias, tableAlias, fieldAlias, valueText
            Next colIndex
            Set rowObj = resultTable.EnsureRow(rowIndex)
            rowAnchorName = ex_Messaging.m_BuildResultRowAnchorName(tableRef, rowIndex + 1)
            If Len(rowAnchorName) > 0 Then
                rowObj.RowAnchorName = rowAnchorName
                ex_Messaging.m_RegisterResultRowAnchor ws, rowAnchorName, nextRow
            End If
            contentRows.Add nextRow
            nextRow = nextRow + 1
        End If
        rs.MoveNext
    Loop

    mp_AppendFilteredRows = (nextRow > startRow)
End Function

Private Function mp_AsText(ByVal valueIn As Variant, Optional ByVal adoFieldType As Long = -1) As String
    mp_AsText = ex_SqlAdoHelpers.m_ToNormalizedText(valueIn, adoFieldType)
End Function

Private Function mp_ListRecordsetFields(ByVal rs As Object, Optional ByVal maxCount As Long = 25) As String
    Dim i As Long
    Dim count As Long
    Dim fieldName As String

    If rs Is Nothing Then Exit Function
    If maxCount <= 0 Then maxCount = 25

    For i = 0 To rs.Fields.Count - 1
        fieldName = Trim$(CStr(rs.Fields(i).Name))
        If Len(fieldName) = 0 Then fieldName = "(empty)"
        If count > 0 Then mp_ListRecordsetFields = mp_ListRecordsetFields & ", "
        mp_ListRecordsetFields = mp_ListRecordsetFields & "[" & fieldName & "]"
        count = count + 1
        If count >= maxCount Then Exit For
    Next i

    If rs.Fields.Count > maxCount Then
        mp_ListRecordsetFields = mp_ListRecordsetFields & ", ..."
    End If
End Function

Private Function mp_RecordsetLooksLikeGenericFields(ByVal rs As Object) As Boolean
    Dim i As Long
    Dim probeCount As Long
    Dim fieldName As String

    If rs Is Nothing Then Exit Function
    If rs.Fields.Count = 0 Then Exit Function

    probeCount = rs.Fields.Count
    If probeCount > 10 Then probeCount = 10

    For i = 0 To probeCount - 1
        fieldName = UCase$(Trim$(CStr(rs.Fields(i).Name)))
        If Len(fieldName) < 2 Then Exit Function
        If Left$(fieldName, 1) <> "F" Then Exit Function
        If Not IsNumeric(Mid$(fieldName, 2)) Then Exit Function
    Next i

    mp_RecordsetLooksLikeGenericFields = True
End Function

Private Function mp_RecordsetGetFieldOrdinal(ByVal rs As Object, ByVal fieldName As String) As Long
    Dim i As Long
    Dim targetToken As String

    mp_RecordsetGetFieldOrdinal = -1
    If rs Is Nothing Then Exit Function

    targetToken = mp_NormalizeHeader(fieldName)
    If Len(targetToken) = 0 Then Exit Function

    For i = 0 To rs.Fields.Count - 1
        If StrComp(mp_NormalizeHeader(CStr(rs.Fields(i).Name)), targetToken, vbTextCompare) = 0 Then
            mp_RecordsetGetFieldOrdinal = i
            Exit Function
        End If
    Next i
End Function

Private Function mp_ToSafeText(ByVal valueIn As Variant) As String
    If IsError(valueIn) Then Exit Function
    If IsNull(valueIn) Then Exit Function
    If IsEmpty(valueIn) Then Exit Function

    mp_ToSafeText = Trim$(CStr(valueIn))
End Function

Private Function mp_IsEmptyVariantArray(ByVal valueRef As Variant) As Boolean
    Dim lb As Long
    Dim ub As Long

    If IsEmpty(valueRef) Then
        mp_IsEmptyVariantArray = True
        Exit Function
    End If
    If Not IsArray(valueRef) Then
        mp_IsEmptyVariantArray = True
        Exit Function
    End If

    On Error GoTo ErrHandler
    lb = LBound(valueRef)
    ub = UBound(valueRef)
    mp_IsEmptyVariantArray = (ub < lb)
    Exit Function

ErrHandler:
    mp_IsEmptyVariantArray = True
End Function

Private Function mp_ToColumnLetter(ByVal columnIndex As Long) As String
    Dim n As Long
    Dim remainder As Long

    If columnIndex < 1 Then columnIndex = 1
    n = columnIndex

    Do While n > 0
        remainder = (n - 1) Mod 26
        mp_ToColumnLetter = Chr$(65 + remainder) & mp_ToColumnLetter
        n = (n - remainder - 1) \ 26
    Loop
End Function

Private Function mp_NormalizeHeader(ByVal valueText As String) As String
    Dim normalized As String

    normalized = CStr(valueText)
    normalized = Replace$(normalized, vbCr, " ")
    normalized = Replace$(normalized, vbLf, " ")
    normalized = Replace$(normalized, vbTab, " ")
    normalized = Replace$(normalized, ChrW$(160), " ")
    normalized = Replace$(normalized, "#", ".")
    normalized = Replace$(normalized, ChrW$(&H2019), "'")
    normalized = Replace$(normalized, ChrW$(&H2BC), "'")
    normalized = Replace$(normalized, ChrW$(&H60), "'")
    normalized = Replace$(normalized, ChrW$(&HB4), "'")
    normalized = Replace$(normalized, "  ", " ")
    normalized = Replace$(normalized, "  ", " ")
    normalized = Trim$(normalized)
    normalized = LCase$(normalized)

    mp_NormalizeHeader = normalized
End Function

Private Function mp_NormalizeHeaderLoose(ByVal valueText As String) As String
    Dim normalized As String
    Dim i As Long
    Dim ch As String
    Dim codePoint As Long
    Dim resultText As String

    normalized = mp_NormalizeHeader(valueText)
    If Len(normalized) = 0 Then Exit Function

    For i = 1 To Len(normalized)
        ch = Mid$(normalized, i, 1)
        codePoint = AscW(ch)
        If (codePoint >= 48 And codePoint <= 57) _
           Or (codePoint >= 65 And codePoint <= 90) _
           Or (codePoint >= 97 And codePoint <= 122) _
           Or (codePoint >= &H410 And codePoint <= &H44F) _
           Or codePoint = &H401 _
           Or codePoint = &H451 _
           Or codePoint = &H404 _
           Or codePoint = &H454 _
           Or codePoint = &H406 _
           Or codePoint = &H456 _
           Or codePoint = &H407 _
           Or codePoint = &H457 _
           Or codePoint = &H490 _
           Or codePoint = &H491 Then
            resultText = resultText & ch
        End If
    Next i

    mp_NormalizeHeaderLoose = resultText
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
    Dim tableObj As obj_ResultTable
    Dim i As Long
    Dim fieldAlias As String

    Set tableObj = New obj_ResultTable
    tableObj.Initialize sourceAlias & ".Sheet[" & tableAlias & "]"

    For i = 1 To fields.Count
        fieldAlias = Trim$(CStr(fields(i)))
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
    Dim sourcePrefix As String
    Dim fileKey As String
    Dim resolverKey As String
    Dim resolverArgsKey As String
    Dim rawPath As String
    Dim resolverName As String
    Dim resolverCallName As String
    Dim resolverArgs As String
    Dim resolvedValue As Variant
    Dim resolvedPath As String

    sourcePrefix = "Source." & Trim$(sourceAlias)
    fileKey = sourcePrefix & ".FilePath"
    resolverKey = sourcePrefix & ".FileResolver"
    resolverArgsKey = sourcePrefix & ".FileResolverArgs"

    rawPath = mp_GetCfgRequired(cfg, fileKey)
    resolverName = mp_GetCfgOptional(cfg, resolverKey, vbNullString)
    resolverArgs = mp_GetCfgOptional(cfg, resolverArgsKey, vbNullString)

    If Len(resolverName) = 0 Then
        If mp_HasPlaceholderTokens(rawPath) Then
            Err.Raise vbObjectError + 6515, "ex_ModeMultiSources", _
                "Source path contains placeholders but no resolver is configured for key '" & fileKey & "'."
        End If

        mp_GetResolvedSourcePath = mp_ResolvePathLocal(rawPath)
        Exit Function
    End If

    If InStr(1, resolverName, "!", vbBinaryCompare) > 0 Then
        resolverCallName = resolverName
    Else
        resolverCallName = "'" & ThisWorkbook.Name & "'!" & resolverName
    End If

    On Error GoTo ResolverEH
    resolvedValue = Application.Run(resolverCallName, rawPath, resolverArgs)
    On Error GoTo 0

    resolvedPath = Trim$(CStr(resolvedValue))
    If Len(resolvedPath) = 0 Then
        Err.Raise vbObjectError + 6516, "ex_ModeMultiSources", _
            "Source file resolver '" & resolverName & "' returned an empty path for key '" & fileKey & "'."
    End If

    mp_GetResolvedSourcePath = mp_ResolvePathLocal(resolvedPath)
    Exit Function

ResolverEH:
    Err.Raise vbObjectError + 6517, "ex_ModeMultiSources", _
        "Source file resolver failed for key '" & fileKey & "' (resolver='" & resolverName & "'): " & Err.Description
End Function

Private Function mp_GetCfgRequired(ByVal cfg As Object, ByVal keyName As String) As String
    Dim valueText As String

    If cfg Is Nothing Or Not cfg.Exists(keyName) Then
        Err.Raise vbObjectError + 6518, "ex_ModeMultiSources", "Missing config key: " & keyName
    End If

    valueText = Trim$(CStr(cfg(keyName)))
    If Len(valueText) = 0 Then
        Err.Raise vbObjectError + 6519, "ex_ModeMultiSources", "Empty config value: " & keyName
    End If

    mp_GetCfgRequired = valueText
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
    Dim normalized As String

    normalized = Trim$(valueText)
    If Len(normalized) = 0 Then Exit Function

    mp_HasPlaceholderTokens = (InStr(1, normalized, "{", vbBinaryCompare) > 0) _
                              And (InStr(1, normalized, "}", vbBinaryCompare) > 0)
End Function

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
