Attribute VB_Name = "ex_ResultSqlEngine"
Option Explicit

Private Const LIKE_DIALECT_UNKNOWN As String = "unknown"
Private Const LIKE_DIALECT_STAR As String = "star"
Private Const LIKE_DIALECT_PERCENT As String = "percent"
Private Const LONG_TEXT_CACHE_AMBIGUOUS As String = "#AMBIGUOUS#"
Private Const HYDRATION_DEBUG_LOG_PATH As String = "Logs\personalcard_pipeline.log"

Private g_LikeDialectByConnection As Object
Private g_LongTextRuntimeCacheBySignature As Object

Public Function m_GetCfgRequired( _
    ByVal cfg As Object, _
    ByVal keyName As String, _
    ByVal errSource As String, _
    ByVal errMissingCode As Long, _
    ByVal errEmptyCode As Long _
) As String
    Dim valueText As String

    If cfg Is Nothing Or Not cfg.Exists(keyName) Then
        Err.Raise errMissingCode, errSource, "Missing config key: " & keyName
    End If

    valueText = Trim$(CStr(cfg(keyName)))
    If Len(valueText) = 0 Then
        Err.Raise errEmptyCode, errSource, "Empty config value: " & keyName
    End If

    m_GetCfgRequired = valueText
End Function

Public Function m_GetCfgOptional( _
    ByVal cfg As Object, _
    ByVal keyName As String, _
    Optional ByVal defaultValue As String = vbNullString _
) As String
    If cfg Is Nothing Then
        m_GetCfgOptional = defaultValue
        Exit Function
    End If

    If Not cfg.Exists(keyName) Then
        m_GetCfgOptional = defaultValue
        Exit Function
    End If

    m_GetCfgOptional = Trim$(CStr(cfg(keyName)))
End Function

Public Function m_HasPlaceholderTokens(ByVal valueText As String) As Boolean
    Dim normalized As String

    normalized = Trim$(valueText)
    If Len(normalized) = 0 Then Exit Function

    m_HasPlaceholderTokens = (InStr(1, normalized, "{", vbBinaryCompare) > 0) _
                             And (InStr(1, normalized, "}", vbBinaryCompare) > 0)
End Function

Public Function m_ResolvePathLocal(ByVal inputPath As String) As String
    Dim basePath As String

    inputPath = Trim$(inputPath)
    If Len(inputPath) = 0 Then Exit Function

    If Left$(inputPath, 2) = "\\" Or InStr(1, inputPath, ":\", vbTextCompare) > 0 Then
        m_ResolvePathLocal = inputPath
        Exit Function
    End If

    basePath = ThisWorkbook.Path
    If Len(basePath) = 0 Then
        m_ResolvePathLocal = inputPath
        Exit Function
    End If

    If Right$(basePath, 1) <> "\" Then
        basePath = basePath & "\"
    End If

    m_ResolvePathLocal = basePath & inputPath
End Function

Public Function m_GetResolvedSourcePath( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal errSource As String, _
    ByVal errNoResolverForTemplateCode As Long, _
    ByVal errResolverReturnedEmptyCode As Long, _
    ByVal errResolverFailedCode As Long, _
    Optional ByVal errMissingCfgCode As Long = vbObjectError + 1, _
    Optional ByVal errEmptyCfgCode As Long = vbObjectError + 2 _
) As String
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

    rawPath = m_GetCfgRequired(cfg, fileKey, errSource, errMissingCfgCode, errEmptyCfgCode)
    resolverName = m_GetCfgOptional(cfg, resolverKey, vbNullString)
    resolverArgs = m_GetCfgOptional(cfg, resolverArgsKey, vbNullString)

    If Len(resolverName) = 0 Then
        If m_HasPlaceholderTokens(rawPath) Then
            Err.Raise errNoResolverForTemplateCode, errSource, _
                "Source path contains placeholders but no resolver is configured for key '" & fileKey & "'."
        End If

        m_GetResolvedSourcePath = m_ResolvePathLocal(rawPath)
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
        Err.Raise errResolverReturnedEmptyCode, errSource, _
            "Source file resolver '" & resolverName & "' returned an empty path for key '" & fileKey & "'."
    End If

    m_GetResolvedSourcePath = m_ResolvePathLocal(resolvedPath)
    Exit Function

ResolverEH:
    Err.Raise errResolverFailedCode, errSource, _
        "Source file resolver failed for key '" & fileKey & "' (resolver='" & resolverName & "'): " & Err.Description
End Function

Public Function m_BuildAdoConnectionString( _
    ByVal sourcePath As String, _
    ByVal errSource As String, _
    ByVal errUnsupportedExtensionCode As Long _
) As String
    Dim ext As String
    Dim props As String

    ext = LCase$(Mid$(sourcePath, InStrRev(sourcePath, ".") + 1))
    Select Case ext
        Case "xls"
            props = "Excel 8.0;HDR=YES;IMEX=1;ReadOnly=True;TypeGuessRows=0;ImportMixedTypes=Text;MAXSCANROWS=0"
        Case "xlsx"
            props = "Excel 12.0 Xml;HDR=YES;IMEX=1;ReadOnly=True;TypeGuessRows=0;ImportMixedTypes=Text;MAXSCANROWS=0"
        Case "xlsm"
            props = "Excel 12.0 Macro;HDR=YES;IMEX=1;ReadOnly=True;TypeGuessRows=0;ImportMixedTypes=Text;MAXSCANROWS=0"
        Case "xlsb"
            props = "Excel 12.0;HDR=YES;IMEX=1;ReadOnly=True;TypeGuessRows=0;ImportMixedTypes=Text;MAXSCANROWS=0"
        Case Else
            Err.Raise errUnsupportedExtensionCode, errSource, _
                "Unsupported source file extension for ADO: ." & ext
    End Select

    m_BuildAdoConnectionString = _
        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sourcePath & ";Extended Properties=""" & props & """;"
End Function

Public Function m_OpenSourceConnection( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal errSource As String, _
    ByVal errPathEmptyCode As Long, _
    ByVal errNotFoundCode As Long, _
    ByVal errUnsupportedExtensionCode As Long, _
    ByVal errNoResolverForTemplateCode As Long, _
    ByVal errResolverReturnedEmptyCode As Long, _
    ByVal errResolverFailedCode As Long, _
    Optional ByVal errMissingCfgCode As Long = vbObjectError + 1, _
    Optional ByVal errEmptyCfgCode As Long = vbObjectError + 2 _
) As Object
    Dim sourcePath As String
    Dim snapshotPath As String
    Dim conn As Object

    sourcePath = m_GetResolvedSourcePath( _
        cfg, _
        sourceAlias, _
        errSource, _
        errNoResolverForTemplateCode, _
        errResolverReturnedEmptyCode, _
        errResolverFailedCode, _
        errMissingCfgCode, _
        errEmptyCfgCode)

    If Len(sourcePath) = 0 Then
        Err.Raise errPathEmptyCode, errSource, "Source file path is empty for alias '" & sourceAlias & "'."
    End If
    If Dir$(sourcePath) = vbNullString Then
        Err.Raise errNotFoundCode, errSource, "Source file not found: " & sourcePath
    End If

    snapshotPath = ex_SourceSnapshot.m_GetSnapshotPath(sourcePath, "Source." & sourceAlias)
    Set conn = CreateObject("ADODB.Connection")
    conn.Open m_BuildAdoConnectionString(snapshotPath, errSource, errUnsupportedExtensionCode)

    Set m_OpenSourceConnection = conn
End Function

Public Function m_BuildTableRefFromSheetName( _
    ByVal sheetName As String, _
    ByVal errSource As String, _
    ByVal errEmptySheetNameCode As Long _
) As String
    sheetName = Trim$(sheetName)
    If Len(sheetName) = 0 Then
        Err.Raise errEmptySheetNameCode, errSource, "Resolved SheetName is empty."
    End If

    If Left$(sheetName, 1) = "[" And Right$(sheetName, 1) = "]" Then
        m_BuildTableRefFromSheetName = sheetName
        Exit Function
    End If

    If InStr(1, sheetName, "$", vbBinaryCompare) > 0 Then
        m_BuildTableRefFromSheetName = "[" & Replace$(sheetName, "]", "]]" ) & "]"
    Else
        m_BuildTableRefFromSheetName = "[" & Replace$(sheetName, "]", "]]" ) & "$]"
    End If
End Function

Public Function m_QuoteIdentifier(ByVal valueText As String) As String
    m_QuoteIdentifier = "[" & Replace$(Trim$(valueText), "]", "]]" ) & "]"
End Function

Public Function m_QuoteSqlStringLiteral(ByVal valueText As String) As String
    m_QuoteSqlStringLiteral = "'" & Replace$(CStr(valueText), "'", "''") & "'"
End Function

Public Function m_BuildFilteredSelectSql( _
    ByVal tableRef As String, _
    ByVal keySourceFieldName As String, _
    ByVal keyValue As String, _
    Optional ByVal useLike As Boolean = False, _
    Optional ByVal likeDialect As String = LIKE_DIALECT_UNKNOWN _
) As String
    If useLike Then
        m_BuildFilteredSelectSql = _
            "SELECT * FROM " & tableRef & _
            " WHERE " & m_BuildAdoWhereLikePattern(keySourceFieldName, Trim$(keyValue), likeDialect)
        Exit Function
    End If

    m_BuildFilteredSelectSql = _
        "SELECT * FROM " & tableRef & _
        " WHERE " & m_QuoteIdentifier(keySourceFieldName) & " = " & m_QuoteSqlStringLiteral(Trim$(keyValue))
End Function

Public Function m_BuildAdoWhereLikePattern( _
    ByVal columnName As String, _
    ByVal patternText As String, _
    Optional ByVal likeDialect As String = LIKE_DIALECT_UNKNOWN _
) As String
    Dim colExpr As String
    Dim primaryPattern As String
    Dim altPattern As String
    Dim normalizedPattern As String

    colExpr = m_QuoteIdentifier(columnName)
    primaryPattern = Trim$(patternText)
    normalizedPattern = mp_ConvertPatternForLikeDialect(primaryPattern, likeDialect)

    If StrComp(LCase$(Trim$(likeDialect)), LIKE_DIALECT_STAR, vbBinaryCompare) = 0 Or _
       StrComp(LCase$(Trim$(likeDialect)), LIKE_DIALECT_PERCENT, vbBinaryCompare) = 0 Then
        m_BuildAdoWhereLikePattern = colExpr & " LIKE " & m_QuoteSqlStringLiteral(normalizedPattern)
        Exit Function
    End If

    altPattern = mp_BuildAlternativeLikePattern(primaryPattern)

    If StrComp(primaryPattern, altPattern, vbBinaryCompare) = 0 Then
        m_BuildAdoWhereLikePattern = colExpr & " LIKE " & m_QuoteSqlStringLiteral(primaryPattern)
    Else
        m_BuildAdoWhereLikePattern = "(" & _
            colExpr & " LIKE " & m_QuoteSqlStringLiteral(primaryPattern) & _
            " OR " & colExpr & " LIKE " & m_QuoteSqlStringLiteral(altPattern) & _
            ")"
    End If
End Function

Public Function m_GetLikeDialectForConnection(ByVal conn As Object) As String
    Dim cacheKey As String
    Dim detected As String

    mp_EnsureLikeDialectCache
    cacheKey = mp_GetConnectionCacheKey(conn)
    If Len(cacheKey) > 0 Then
        If g_LikeDialectByConnection.Exists(cacheKey) Then
            m_GetLikeDialectForConnection = CStr(g_LikeDialectByConnection(cacheKey))
            Exit Function
        End If
    End If

    detected = mp_DetectLikeDialect(conn)
    If Len(cacheKey) > 0 Then g_LikeDialectByConnection(cacheKey) = detected
    m_GetLikeDialectForConnection = detected
End Function

Public Function m_MappedHeader( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAlias As String, _
    ByVal errSource As String, _
    ByVal errMappedHeaderEmptyCode As Long, _
    Optional ByVal errMissingCfgCode As Long = vbObjectError + 1, _
    Optional ByVal errEmptyCfgCode As Long = vbObjectError + 2 _
) As String
    Dim rawValue As String
    Dim splitPos As Long

    rawValue = m_GetCfgRequired( _
        cfg, _
        sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]", _
        errSource, _
        errMissingCfgCode, _
        errEmptyCfgCode)

    splitPos = InStr(1, rawValue, "|", vbBinaryCompare)
    If splitPos > 0 Then
        m_MappedHeader = Trim$(Left$(rawValue, splitPos - 1))
    Else
        m_MappedHeader = Trim$(rawValue)
    End If

    If Len(m_MappedHeader) = 0 Then
        Err.Raise errMappedHeaderEmptyCode, errSource, _
            "Mapped source header is empty for " & sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]"
    End If
End Function

Public Function m_GetFieldDisplayHeader( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAlias As String, _
    ByVal errSource As String, _
    ByVal errMappedHeaderEmptyCode As Long, _
    Optional ByVal errMissingCfgCode As Long = vbObjectError + 1, _
    Optional ByVal errEmptyCfgCode As Long = vbObjectError + 2 _
) As String
    Dim rawValue As String
    Dim splitPos As Long
    Dim labelText As String

    rawValue = m_GetCfgRequired( _
        cfg, _
        sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]", _
        errSource, _
        errMissingCfgCode, _
        errEmptyCfgCode)

    splitPos = InStr(1, rawValue, "|", vbBinaryCompare)
    If splitPos > 0 Then
        labelText = Trim$(Mid$(rawValue, splitPos + 1))
        If Len(labelText) > 0 Then
            m_GetFieldDisplayHeader = labelText
            Exit Function
        End If
    End If

    m_GetFieldDisplayHeader = m_MappedHeader( _
        cfg, sourceAlias, tableAlias, fieldAlias, errSource, errMappedHeaderEmptyCode, errMissingCfgCode, errEmptyCfgCode)
End Function

Public Function m_GetExpectedMappedHeaders( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fields As Collection, _
    ByVal errSource As String, _
    ByVal errMappedHeaderEmptyCode As Long, _
    Optional ByVal errMissingCfgCode As Long = vbObjectError + 1, _
    Optional ByVal errEmptyCfgCode As Long = vbObjectError + 2 _
) As Variant
    Dim arr() As String
    Dim i As Long
    Dim headerText As String

    If fields Is Nothing Then
        m_GetExpectedMappedHeaders = Array()
        Exit Function
    End If
    If fields.Count = 0 Then
        m_GetExpectedMappedHeaders = Array()
        Exit Function
    End If

    ReDim arr(0 To fields.Count - 1)
    For i = 1 To fields.Count
        headerText = m_MappedHeader( _
            cfg, sourceAlias, tableAlias, CStr(fields(i)), errSource, errMappedHeaderEmptyCode, errMissingCfgCode, errEmptyCfgCode)
        arr(i - 1) = headerText
    Next i

    m_GetExpectedMappedHeaders = arr
End Function

Public Function m_IsExplicitAdoRangeReference(ByVal value As String) As Boolean
    value = Trim$(value)
    If InStr(1, value, "$", vbBinaryCompare) <= 0 Then Exit Function
    If InStr(1, value, ":", vbBinaryCompare) <= 0 Then Exit Function
    m_IsExplicitAdoRangeReference = True
End Function

Public Function m_UnquoteSqlIdentifier(ByVal value As String) As String
    value = Trim$(value)
    If Len(value) >= 2 Then
        If Left$(value, 1) = "[" And Right$(value, 1) = "]" Then
            value = Mid$(value, 2, Len(value) - 2)
        End If
    End If

    m_UnquoteSqlIdentifier = Replace$(value, "]]", "]")
End Function

Public Function m_ExtractAdoSheetPrefix(ByVal tableRef As String) As String
    Dim objectName As String
    Dim dollarPos As Long

    objectName = m_UnquoteSqlIdentifier(tableRef)
    If Len(objectName) = 0 Then Exit Function

    dollarPos = InStr(1, objectName, "$", vbBinaryCompare)
    If dollarPos <= 0 Then Exit Function

    m_ExtractAdoSheetPrefix = Left$(objectName, dollarPos)
End Function

Public Sub m_ResetLongTextRuntimeCache(Optional ByVal cacheSignature As String = vbNullString)
    Dim normalizedSignature As String

    mp_EnsureLongTextRuntimeCacheStorage
    normalizedSignature = mp_NormalizeLongTextRuntimeCacheSignature(cacheSignature)

    If Len(normalizedSignature) = 0 Then
        g_LongTextRuntimeCacheBySignature.RemoveAll
        Exit Sub
    End If

    If g_LongTextRuntimeCacheBySignature.Exists(normalizedSignature) Then
        g_LongTextRuntimeCacheBySignature.Remove normalizedSignature
    End If
End Sub

Public Function m_GetLongTextRuntimeCache(ByVal cacheSignature As String) As Object
    Dim normalizedSignature As String
    Dim runtimeCache As Object

    mp_EnsureLongTextRuntimeCacheStorage
    normalizedSignature = mp_NormalizeLongTextRuntimeCacheSignature(cacheSignature)
    If Len(normalizedSignature) = 0 Then normalizedSignature = "__default__"

    If g_LongTextRuntimeCacheBySignature.Exists(normalizedSignature) Then
        Set m_GetLongTextRuntimeCache = g_LongTextRuntimeCacheBySignature(normalizedSignature)
        Exit Function
    End If

    Set runtimeCache = CreateObject("Scripting.Dictionary")
    runtimeCache.CompareMode = 1
    g_LongTextRuntimeCacheBySignature.Add normalizedSignature, runtimeCache

    Set m_GetLongTextRuntimeCache = runtimeCache
End Function

Public Function m_FindHeaderColumnInWorksheetRow( _
    ByVal ws As Worksheet, _
    ByVal headerRow As Long, _
    ByVal maxCol As Long, _
    ByVal headerName As String _
) As Long
    Dim needle As String
    Dim needleAlt As String
    Dim c As Long
    Dim currentHeader As String
    Dim currentNorm As String

    If ws Is Nothing Then Exit Function
    If headerRow <= 0 Then Exit Function
    If maxCol <= 0 Then Exit Function

    needle = m_NormalizeHeader(headerName)
    needleAlt = m_NormalizeHeader(Replace$(headerName, "#", "."))

    For c = 1 To maxCol
        currentHeader = CStr(ws.Cells(headerRow, c).Value2)
        currentNorm = m_NormalizeHeader(currentHeader)
        If currentNorm = needle Or currentNorm = needleAlt Then
            m_FindHeaderColumnInWorksheetRow = c
            Exit Function
        End If
    Next c

    m_FindHeaderColumnInWorksheetRow = -1
End Function

Public Function m_FindWorksheetByConfiguredAdoName(ByVal wb As Workbook, ByVal configuredSheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim needle As String
    Dim needleAlt As String

    needle = mp_ExtractSheetNameToken(configuredSheetName)
    If Len(needle) = 0 Then Exit Function
    needleAlt = Replace$(needle, "#", ".")

    For Each ws In wb.Worksheets
        If StrComp(Trim$(ws.Name), needle, vbTextCompare) = 0 Then
            Set m_FindWorksheetByConfiguredAdoName = ws
            Exit Function
        End If
        If StrComp(Replace$(Trim$(ws.Name), ".", "#"), needle, vbTextCompare) = 0 Then
            Set m_FindWorksheetByConfiguredAdoName = ws
            Exit Function
        End If
        If StrComp(Trim$(ws.Name), needleAlt, vbTextCompare) = 0 Then
            Set m_FindWorksheetByConfiguredAdoName = ws
            Exit Function
        End If
    Next ws
End Function

Public Function m_FindFirstMarkerCell(ByVal ws As Worksheet, ByVal markerText As String) As Range
    Dim searchRange As Range

    If ws Is Nothing Then Exit Function
    markerText = Trim$(markerText)
    If Len(markerText) = 0 Then Exit Function

    Set searchRange = ws.UsedRange
    If searchRange Is Nothing Then Exit Function

    Set m_FindFirstMarkerCell = searchRange.Find( _
        What:=markerText, _
        After:=searchRange.Cells(searchRange.Cells.Count), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False)
End Function

Public Function m_FindMarkerCellInColumnAfterRow( _
    ByVal ws As Worksheet, _
    ByVal markerColumn As Long, _
    ByVal markerText As String, _
    ByVal minExclusiveRow As Long _
) As Range
    Dim searchRange As Range
    Dim firstFound As Range
    Dim currentFound As Range
    Dim firstAddress As String
    Dim bestRow As Long

    If ws Is Nothing Then Exit Function
    If markerColumn <= 0 Then Exit Function
    markerText = Trim$(markerText)
    If Len(markerText) = 0 Then Exit Function

    On Error Resume Next
    Set searchRange = Intersect(ws.Columns(markerColumn), ws.UsedRange)
    On Error GoTo 0
    If searchRange Is Nothing Then
        Set searchRange = ws.Columns(markerColumn)
    End If

    Set firstFound = searchRange.Find( _
        What:=markerText, _
        After:=searchRange.Cells(searchRange.Cells.Count), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False)
    If firstFound Is Nothing Then Exit Function

    bestRow = 0
    firstAddress = firstFound.Address
    Set currentFound = firstFound

    Do
        If currentFound.Row > minExclusiveRow Then
            If bestRow = 0 Or currentFound.Row < bestRow Then
                bestRow = currentFound.Row
                Set m_FindMarkerCellInColumnAfterRow = currentFound
            End If
        End If
        Set currentFound = searchRange.FindNext(currentFound)
        If currentFound Is Nothing Then Exit Do
    Loop While currentFound.Address <> firstAddress
End Function

Public Sub m_HydrateAdoLongTextFromWorksheetIfNeeded( _
    ByRef outValues As Variant, _
    ByVal rowCount As Long, _
    ByVal fields As Variant, _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal configuredSheetName As String, _
    ByVal keyValue As String, _
    ByVal wbCache As Object, _
    ByVal runtimeCache As Object _
)
    Dim hasUncachedLongText As Boolean
    Dim rangeStartMarker As String
    Dim rangeEndMarker As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim startCell As Range
    Dim endCell As Range
    Dim markerCol As Long
    Dim headerRow As Long
    Dim dataStartRow As Long
    Dim dataEndRow As Long
    Dim keyAlias As String
    Dim keyHeaderName As String
    Dim keyCol As Long
    Dim fieldColsByIdx As Object

    On Error GoTo SafeExit

    If rowCount <= 0 Then
        mp_DebugHydrationLog "hydrate", "skip reason=rowCount<=0 src='" & sourceAlias & "' tbl='" & tableAlias & "'"
        Exit Sub
    End If
    If m_IsEmptyVariantArray(fields) Then
        mp_DebugHydrationLog "hydrate", "skip reason=fields-empty src='" & sourceAlias & "' tbl='" & tableAlias & "'"
        Exit Sub
    End If
    If Len(Trim$(keyValue)) = 0 Then
        mp_DebugHydrationLog "hydrate", "skip reason=empty-key src='" & sourceAlias & "' tbl='" & tableAlias & "'"
        Exit Sub
    End If
    If m_IsExplicitAdoRangeReference(configuredSheetName) Then
        mp_DebugHydrationLog "hydrate", "skip reason=explicit-range src='" & sourceAlias & "' tbl='" & tableAlias & "' sheet='" & configuredSheetName & "'"
        Exit Sub
    End If
    If wbCache Is Nothing Then
        mp_DebugHydrationLog "hydrate", "skip reason=wbCache-nothing src='" & sourceAlias & "' tbl='" & tableAlias & "'"
        Exit Sub
    End If

    mp_DebugHydrationLog "hydrate", "start src='" & sourceAlias & "' tbl='" & tableAlias & "' key='" & mp_DebugToken(keyValue, 80) & "' rows=" & CStr(rowCount)

    hasUncachedLongText = m_ApplyLongTextRuntimeCache( _
        outValues, _
        rowCount, _
        fields, _
        sourceAlias, _
        tableAlias, _
        configuredSheetName, _
        keyValue, _
        runtimeCache)
    If Not hasUncachedLongText Then Exit Sub

    rangeStartMarker = m_GetCfgOptional(cfg, sourceAlias & ".Sheet[" & tableAlias & "].RangeStartMarker", vbNullString)
    rangeEndMarker = m_GetCfgOptional(cfg, sourceAlias & ".Sheet[" & tableAlias & "].RangeEndMarker", vbNullString)
    If Len(rangeStartMarker) = 0 Or Len(rangeEndMarker) = 0 Then
        mp_DebugHydrationLog "hydrate", "skip reason=markers-missing src='" & sourceAlias & "' tbl='" & tableAlias & "'"
        Exit Sub
    End If

    Set wb = mp_GetWorkbookForSource(wbCache, cfg, sourceAlias)
    If wb Is Nothing Then
        mp_DebugHydrationLog "hydrate", "skip reason=workbook-not-opened src='" & sourceAlias & "' tbl='" & tableAlias & "'"
        Exit Sub
    End If

    Set ws = m_FindWorksheetByConfiguredAdoName(wb, configuredSheetName)
    If ws Is Nothing Then
        mp_DebugHydrationLog "hydrate", "skip reason=worksheet-not-found src='" & sourceAlias & "' tbl='" & tableAlias & "' sheet='" & configuredSheetName & "'"
        Exit Sub
    End If

    Set startCell = m_FindFirstMarkerCell(ws, rangeStartMarker)
    If startCell Is Nothing Then
        mp_DebugHydrationLog "hydrate", "skip reason=start-marker-not-found src='" & sourceAlias & "' tbl='" & tableAlias & "' marker='" & mp_DebugToken(rangeStartMarker, 60) & "'"
        Exit Sub
    End If

    markerCol = startCell.Column
    Set endCell = m_FindMarkerCellInColumnAfterRow(ws, markerCol, rangeEndMarker, startCell.Row)
    If endCell Is Nothing Then
        mp_DebugHydrationLog "hydrate", "skip reason=end-marker-not-found src='" & sourceAlias & "' tbl='" & tableAlias & "' marker='" & mp_DebugToken(rangeEndMarker, 60) & "'"
        Exit Sub
    End If

    headerRow = startCell.Row - 1
    dataStartRow = startCell.Row
    dataEndRow = endCell.Row - 1
    If headerRow < 1 Or dataEndRow < dataStartRow Then
        mp_DebugHydrationLog "hydrate", "skip reason=invalid-data-range src='" & sourceAlias & "' tbl='" & tableAlias & "'"
        Exit Sub
    End If

    keyAlias = m_GetCfgRequired(cfg, sourceAlias & ".Sheet[" & tableAlias & "].Key", "ex_ResultSqlEngine", vbObjectError + 1370, vbObjectError + 1371)
    keyHeaderName = mp_GetMappedSourceHeaderForHydration(cfg, sourceAlias, tableAlias, keyAlias)
    keyCol = m_FindHeaderColumnInWorksheetRow(ws, headerRow, markerCol, keyHeaderName)
    If keyCol <= 0 Then
        mp_DebugHydrationLog "hydrate", "skip reason=key-column-not-found src='" & sourceAlias & "' tbl='" & tableAlias & "' keyAlias='" & keyAlias & "'"
        Exit Sub
    End If

    Set fieldColsByIdx = mp_FindFieldColumnsForHydration(fields, cfg, sourceAlias, tableAlias, ws, headerRow, markerCol)
    If fieldColsByIdx Is Nothing Then
        mp_DebugHydrationLog "hydrate", "skip reason=field-map-nothing src='" & sourceAlias & "' tbl='" & tableAlias & "'"
        Exit Sub
    End If
    If fieldColsByIdx.Count = 0 Then
        mp_DebugHydrationLog "hydrate", "skip reason=field-map-empty src='" & sourceAlias & "' tbl='" & tableAlias & "'"
        Exit Sub
    End If

    mp_DebugHydrationLog "hydrate", "worksheet-ready src='" & sourceAlias & "' tbl='" & tableAlias & "' dataRows=" & CStr(dataEndRow - dataStartRow + 1) & " mappedFields=" & CStr(fieldColsByIdx.Count)

    m_TryHydrateLongAdoValuesFromWorksheet _
        outValues, _
        rowCount, _
        fields, _
        fieldColsByIdx, _
        ws, _
        dataStartRow, _
        dataEndRow, _
        keyCol, _
        keyValue, _
        sourceAlias, _
        tableAlias, _
        configuredSheetName, _
        runtimeCache

    mp_DebugHydrationLog "hydrate", "done src='" & sourceAlias & "' tbl='" & tableAlias & "'"

SafeExit:
    If Err.Number <> 0 Then
        mp_DebugHydrationLog "hydrate", "error src='" & sourceAlias & "' tbl='" & tableAlias & "' err=['" & Err.Source & "' #" & CStr(Err.Number) & "] " & Err.Description
    End If
    On Error GoTo 0
End Sub

Public Function m_ApplyLongTextRuntimeCache( _
    ByRef outValues As Variant, _
    ByVal rowCount As Long, _
    ByVal fields As Variant, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal configuredSheetName As String, _
    ByVal keyValue As String, _
    ByVal runtimeCache As Object _
) As Boolean
    Dim hasRuntimeCache As Boolean
    Dim outIndex As Long
    Dim i As Long
    Dim outCol As Long
    Dim outValueText As String
    Dim prefixToken As String
    Dim cacheKey As String
    Dim looseCacheKey As String
    Dim cachedText As String
    Dim candidates As Long
    Dim strictHits As Long
    Dim looseHits As Long
    Dim misses As Long

    If rowCount <= 0 Then Exit Function
    If m_IsEmptyVariantArray(fields) Then Exit Function
    If Len(Trim$(keyValue)) = 0 Then Exit Function

    hasRuntimeCache = Not runtimeCache Is Nothing

    For outIndex = 1 To rowCount
        For i = LBound(fields) To UBound(fields)
            outCol = 1 + (i - LBound(fields))
            outValueText = CStr(outValues(outIndex, outCol))

            If Len(outValueText) > 250 Then
                candidates = candidates + 1
                prefixToken = mp_BuildLongTextPrefixToken(outValueText)
                cacheKey = mp_BuildLongTextRuntimeCacheKey(sourceAlias, tableAlias, configuredSheetName, keyValue, outIndex, i, prefixToken)
                looseCacheKey = mp_BuildLongTextLooseCacheKey(sourceAlias, tableAlias, configuredSheetName, keyValue, i, prefixToken)

                If hasRuntimeCache Then
                    If runtimeCache.Exists(cacheKey) Then
                        strictHits = strictHits + 1
                        outValues(outIndex, outCol) = CStr(runtimeCache(cacheKey))
                        GoTo ContinueCacheCandidate
                    End If

                    If runtimeCache.Exists(looseCacheKey) Then
                        cachedText = CStr(runtimeCache(looseCacheKey))
                        If StrComp(cachedText, LONG_TEXT_CACHE_AMBIGUOUS, vbBinaryCompare) <> 0 Then
                            looseHits = looseHits + 1
                            outValues(outIndex, outCol) = cachedText
                            GoTo ContinueCacheCandidate
                        End If
                    End If
                End If

                misses = misses + 1
                m_ApplyLongTextRuntimeCache = True
            End If
ContinueCacheCandidate:
        Next i
    Next outIndex

    If candidates > 0 Then
        mp_DebugHydrationLog "cache-pass", _
            "src='" & sourceAlias & "' tbl='" & tableAlias & "' key='" & mp_DebugToken(keyValue, 80) & _
            "' candidates=" & CStr(candidates) & _
            " strictHit=" & CStr(strictHits) & _
            " looseHit=" & CStr(looseHits) & _
            " miss=" & CStr(misses)
    End If
End Function

Public Sub m_TryHydrateLongAdoValuesFromWorksheet( _
    ByRef outValues As Variant, _
    ByVal rowCount As Long, _
    ByVal fields As Variant, _
    ByVal fieldColsByIdx As Object, _
    ByVal ws As Worksheet, _
    ByVal dataStartRow As Long, _
    ByVal dataEndRow As Long, _
    ByVal keyCol As Long, _
    ByVal keyValue As String, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal configuredSheetName As String, _
    ByVal runtimeCache As Object _
)
    Dim unresolvedCellKeys As Object
    Dim unresolvedPrefixesByCellKey As Object
    Dim unresolvedFieldIdx As Object
    Dim neededPrefixesByFieldIdx As Object
    Dim fieldArraysByIdx As Object
    Dim sourceRowsByFieldPrefix As Object
    Dim matches As Collection
    Dim hasRuntimeCache As Boolean
    Dim outIndex As Long
    Dim i As Long
    Dim outCol As Long
    Dim outValueText As String
    Dim cacheKey As String
    Dim looseCacheKey As String
    Dim cachedText As String
    Dim cellKey As String
    Dim fieldKey As Variant
    Dim prefixKey As Variant
    Dim keyArray As Variant
    Dim keyArrayIsSingleCell As Boolean
    Dim r As Long
    Dim fieldCol As Long
    Dim fieldArray As Variant
    Dim fieldArrayIsSingleCell As Boolean
    Dim neededPrefixes As Object
    Dim rowsBucket As Collection
    Dim preferredRow As Long
    Dim lookupKey As String
    Dim prefixToken As String
    Dim parsedRowIdx As Long
    Dim parsedFieldIdx As Long
    Dim sourceRow As Long
    Dim sourceOffset As Long
    Dim normalizedValue As Variant
    Dim unresolvedCount As Long
    Dim keyMatchCount As Long
    Dim prefixBucketCount As Long
    Dim hydratedCount As Long
    Dim prefixUniqueCount As Long
    Dim prefixAmbiguousCount As Long
    Dim fallbackCount As Long
    Dim unresolvedWithoutBucketCount As Long
    Dim strictHitCount As Long
    Dim looseHitCount As Long
    Dim usedFallback As Boolean

    On Error GoTo SafeExit

    If rowCount <= 0 Then Exit Sub
    If ws Is Nothing Then Exit Sub
    If fieldColsByIdx Is Nothing Then Exit Sub
    If m_IsEmptyVariantArray(fields) Then Exit Sub
    If Len(Trim$(keyValue)) = 0 Then Exit Sub
    If dataStartRow <= 0 Or dataEndRow < dataStartRow Then Exit Sub
    If keyCol <= 0 Then Exit Sub

    hasRuntimeCache = Not runtimeCache Is Nothing

    Set unresolvedCellKeys = CreateObject("Scripting.Dictionary")
    unresolvedCellKeys.CompareMode = 1
    Set unresolvedPrefixesByCellKey = CreateObject("Scripting.Dictionary")
    unresolvedPrefixesByCellKey.CompareMode = 1
    Set unresolvedFieldIdx = CreateObject("Scripting.Dictionary")
    unresolvedFieldIdx.CompareMode = 1
    Set neededPrefixesByFieldIdx = CreateObject("Scripting.Dictionary")
    neededPrefixesByFieldIdx.CompareMode = 1

    For outIndex = 1 To rowCount
        For i = LBound(fields) To UBound(fields)
            outCol = 1 + (i - LBound(fields))
            outValueText = CStr(outValues(outIndex, outCol))

            If Len(outValueText) > 250 Then
                If Not fieldColsByIdx.Exists(CStr(i)) Then GoTo ContinueFieldCandidate
                fieldCol = CLng(fieldColsByIdx(CStr(i)))
                If fieldCol <= 0 Then GoTo ContinueFieldCandidate

                cellKey = mp_BuildLongTextCellKey(outIndex, i)
                prefixToken = mp_BuildLongTextPrefixToken(outValueText)
                cacheKey = mp_BuildLongTextRuntimeCacheKey(sourceAlias, tableAlias, configuredSheetName, keyValue, outIndex, i, prefixToken)
                looseCacheKey = mp_BuildLongTextLooseCacheKey(sourceAlias, tableAlias, configuredSheetName, keyValue, i, prefixToken)

                If hasRuntimeCache Then
                    If runtimeCache.Exists(cacheKey) Then
                        strictHitCount = strictHitCount + 1
                        outValues(outIndex, outCol) = CStr(runtimeCache(cacheKey))
                        GoTo ContinueFieldCandidate
                    End If

                    If runtimeCache.Exists(looseCacheKey) Then
                        cachedText = CStr(runtimeCache(looseCacheKey))
                        If StrComp(cachedText, LONG_TEXT_CACHE_AMBIGUOUS, vbBinaryCompare) <> 0 Then
                            looseHitCount = looseHitCount + 1
                            outValues(outIndex, outCol) = cachedText
                            GoTo ContinueFieldCandidate
                        End If
                    End If
                End If

                unresolvedCellKeys(cellKey) = cacheKey
                unresolvedPrefixesByCellKey(cellKey) = prefixToken
                unresolvedFieldIdx(CStr(i)) = True

                If Not neededPrefixesByFieldIdx.Exists(CStr(i)) Then
                    Set neededPrefixes = CreateObject("Scripting.Dictionary")
                    neededPrefixes.CompareMode = 1
                    Set neededPrefixesByFieldIdx(CStr(i)) = neededPrefixes
                Else
                    Set neededPrefixes = neededPrefixesByFieldIdx(CStr(i))
                End If
                neededPrefixes(prefixToken) = True
            End If
ContinueFieldCandidate:
        Next i
    Next outIndex

    If unresolvedCellKeys.Count = 0 Then Exit Sub
    If unresolvedFieldIdx.Count = 0 Then Exit Sub

    unresolvedCount = unresolvedCellKeys.Count

    keyArray = ws.Range(ws.Cells(dataStartRow, keyCol), ws.Cells(dataEndRow, keyCol)).Value2
    Set matches = New Collection
    keyArrayIsSingleCell = Not IsArray(keyArray)

    If keyArrayIsSingleCell Then
        If mp_IsSameHydrationKeyValue(keyArray, keyValue) Then
            matches.Add dataStartRow
        End If
    Else
        For r = 1 To UBound(keyArray, 1)
            If mp_IsSameHydrationKeyValue(keyArray(r, 1), keyValue) Then
                matches.Add (dataStartRow + r - 1)
            End If
        Next r
    End If
    keyMatchCount = matches.Count
    If matches.Count = 0 Then
        mp_DebugHydrationLog "hydrate-core", _
            "src='" & sourceAlias & "' tbl='" & tableAlias & "' key='" & mp_DebugToken(keyValue, 80) & _
            "' unresolved=" & CStr(unresolvedCount) & _
            " cacheStrictHit=" & CStr(strictHitCount) & _
            " cacheLooseHit=" & CStr(looseHitCount) & _
            " keyMatches=0"
        Exit Sub
    End If

    Set fieldArraysByIdx = CreateObject("Scripting.Dictionary")
    fieldArraysByIdx.CompareMode = 1

    For Each fieldKey In unresolvedFieldIdx.Keys
        i = CLng(fieldKey)
        If Not fieldColsByIdx.Exists(CStr(i)) Then GoTo ContinueFieldArray

        fieldCol = CLng(fieldColsByIdx(CStr(i)))
        If fieldCol <= 0 Then GoTo ContinueFieldArray

        fieldArray = ws.Range(ws.Cells(dataStartRow, fieldCol), ws.Cells(dataEndRow, fieldCol)).Value2
        fieldArraysByIdx(CStr(i)) = fieldArray
ContinueFieldArray:
    Next fieldKey

    If fieldArraysByIdx.Count = 0 Then Exit Sub

    Set sourceRowsByFieldPrefix = CreateObject("Scripting.Dictionary")
    sourceRowsByFieldPrefix.CompareMode = 1

    For Each fieldKey In unresolvedFieldIdx.Keys
        i = CLng(fieldKey)
        If Not fieldArraysByIdx.Exists(CStr(i)) Then GoTo ContinueFieldPrefixIndex
        If Not neededPrefixesByFieldIdx.Exists(CStr(i)) Then GoTo ContinueFieldPrefixIndex

        fieldArray = fieldArraysByIdx(CStr(i))
        fieldArrayIsSingleCell = Not IsArray(fieldArray)
        Set neededPrefixes = neededPrefixesByFieldIdx(CStr(i))

        For r = 1 To matches.Count
            sourceRow = CLng(matches(r))
            sourceOffset = sourceRow - dataStartRow + 1
            If sourceOffset < 1 Then GoTo ContinueMatchRow

            If fieldArrayIsSingleCell Then
                normalizedValue = ex_SqlAdoHelpers.m_ToNormalizedCellValue(fieldArray)
            Else
                If sourceOffset > UBound(fieldArray, 1) Then GoTo ContinueMatchRow
                normalizedValue = ex_SqlAdoHelpers.m_ToNormalizedCellValue(fieldArray(sourceOffset, 1))
            End If

            prefixToken = mp_BuildLongTextPrefixToken(CStr(normalizedValue))
            If neededPrefixes.Exists(prefixToken) Then
                lookupKey = mp_BuildFieldPrefixLookupKey(i, prefixToken)
                mp_AddRowToMatchBucket sourceRowsByFieldPrefix, lookupKey, sourceRow
            End If
ContinueMatchRow:
        Next r
ContinueFieldPrefixIndex:
    Next fieldKey

    prefixBucketCount = sourceRowsByFieldPrefix.Count

    For Each fieldKey In unresolvedCellKeys.Keys
        cellKey = CStr(fieldKey)
        If Not mp_TryParseLongTextCellKey(cellKey, parsedRowIdx, parsedFieldIdx) Then GoTo ContinueHydrateCell
        If parsedRowIdx < 1 Then GoTo ContinueHydrateCell
        If parsedFieldIdx < LBound(fields) Or parsedFieldIdx > UBound(fields) Then GoTo ContinueHydrateCell
        If Not fieldArraysByIdx.Exists(CStr(parsedFieldIdx)) Then GoTo ContinueHydrateCell
        If Not unresolvedPrefixesByCellKey.Exists(cellKey) Then GoTo ContinueHydrateCell

        sourceRow = 0
        prefixToken = CStr(unresolvedPrefixesByCellKey(cellKey))
        lookupKey = mp_BuildFieldPrefixLookupKey(parsedFieldIdx, prefixToken)
        If sourceRowsByFieldPrefix.Exists(lookupKey) Then
            Set rowsBucket = sourceRowsByFieldPrefix(lookupKey)
            If rowsBucket.Count = 1 Then
                prefixUniqueCount = prefixUniqueCount + 1
                sourceRow = CLng(rowsBucket(1))
            ElseIf rowsBucket.Count > 1 Then
                prefixAmbiguousCount = prefixAmbiguousCount + 1
                If parsedRowIdx <= matches.Count Then
                    preferredRow = CLng(matches(parsedRowIdx))
                    If mp_CollectionContainsLong(rowsBucket, preferredRow) Then
                        sourceRow = preferredRow
                    End If
                End If
                If sourceRow = 0 Then
                    sourceRow = CLng(rowsBucket(1))
                End If
            End If
        Else
            unresolvedWithoutBucketCount = unresolvedWithoutBucketCount + 1
        End If
        If sourceRow = 0 Then
            ' Safety fallback: keep previous behavior if prefix lookup is ambiguous/missing.
            If parsedRowIdx > matches.Count Then GoTo ContinueHydrateCell
            sourceRow = CLng(matches(parsedRowIdx))
            fallbackCount = fallbackCount + 1
            usedFallback = True
        End If
        sourceOffset = sourceRow - dataStartRow + 1
        If sourceOffset < 1 Then GoTo ContinueHydrateCell

        fieldArray = fieldArraysByIdx(CStr(parsedFieldIdx))
        fieldArrayIsSingleCell = Not IsArray(fieldArray)
        If fieldArrayIsSingleCell Then
            normalizedValue = ex_SqlAdoHelpers.m_ToNormalizedCellValue(fieldArray)
        Else
            If sourceOffset > UBound(fieldArray, 1) Then GoTo ContinueHydrateCell
            normalizedValue = ex_SqlAdoHelpers.m_ToNormalizedCellValue(fieldArray(sourceOffset, 1))
        End If

        outCol = 1 + (parsedFieldIdx - LBound(fields))
        outValues(parsedRowIdx, outCol) = normalizedValue
    hydratedCount = hydratedCount + 1

        If hasRuntimeCache Then
            cacheKey = CStr(unresolvedCellKeys(cellKey))
            runtimeCache(cacheKey) = CStr(normalizedValue)

            looseCacheKey = mp_BuildLongTextLooseCacheKey( _
                sourceAlias, _
                tableAlias, _
                configuredSheetName, _
                keyValue, _
                parsedFieldIdx, _
                CStr(unresolvedPrefixesByCellKey(cellKey)))

            If runtimeCache.Exists(looseCacheKey) Then
                cachedText = CStr(runtimeCache(looseCacheKey))
                If StrComp(cachedText, CStr(normalizedValue), vbBinaryCompare) <> 0 Then
                    runtimeCache(looseCacheKey) = LONG_TEXT_CACHE_AMBIGUOUS
                End If
            Else
                runtimeCache(looseCacheKey) = CStr(normalizedValue)
            End If
        End If
ContinueHydrateCell:
    Next fieldKey

    mp_DebugHydrationLog "hydrate-core", _
        "src='" & sourceAlias & "' tbl='" & tableAlias & "' key='" & mp_DebugToken(keyValue, 80) & _
        "' unresolved=" & CStr(unresolvedCount) & _
        " cacheStrictHit=" & CStr(strictHitCount) & _
        " cacheLooseHit=" & CStr(looseHitCount) & _
        " keyMatches=" & CStr(keyMatchCount) & _
        " prefixBuckets=" & CStr(prefixBucketCount) & _
        " prefixUnique=" & CStr(prefixUniqueCount) & _
        " prefixAmbiguous=" & CStr(prefixAmbiguousCount) & _
        " noPrefixBucket=" & CStr(unresolvedWithoutBucketCount) & _
        " fallback=" & CStr(fallbackCount) & _
        " hydrated=" & CStr(hydratedCount) & _
        " usedFallback=" & LCase$(CStr(usedFallback))

SafeExit:
    If Err.Number <> 0 Then
        mp_DebugHydrationLog "hydrate-core", "error src='" & sourceAlias & "' tbl='" & tableAlias & "' err=['" & Err.Source & "' #" & CStr(Err.Number) & "] " & Err.Description
    End If
    On Error GoTo 0
End Sub

Public Function m_BuildNormalizedHeaderTokenSet(ByVal expectedHeaders As Variant, ByVal keyHeader As String) As Object
    Dim d As Object
    Dim i As Long
    Dim token As String

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    token = m_NormalizeHeader(keyHeader)
    If Len(token) > 0 Then d(token) = True

    If Not m_IsEmptyVariantArray(expectedHeaders) Then
        For i = LBound(expectedHeaders) To UBound(expectedHeaders)
            token = m_NormalizeHeader(CStr(expectedHeaders(i)))
            If Len(token) > 0 Then d(token) = True
        Next i
    End If

    Set m_BuildNormalizedHeaderTokenSet = d
End Function

Public Function m_TryDetectHeaderRangeFromTopRows( _
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

    sheetPrefix = m_ExtractAdoSheetPrefix(tableRef)
    If Len(sheetPrefix) = 0 Then Exit Function

    probeRef = "[" & sheetPrefix & "A1:ZZ200]"

    Set expectedSet = m_BuildNormalizedHeaderTokenSet(expectedHeaders, keyHeader)
    If expectedSet Is Nothing Then Exit Function
    If expectedSet.Count = 0 Then Exit Function

    keyToken = m_NormalizeHeader(keyHeader)
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
            cellText = m_ToSafeText(rowsData(colIndex, rowIndex))
            normalized = m_NormalizeHeader(cellText)
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

    colLetter = m_ToColumnLetter(bestLastCol)
    If Len(colLetter) = 0 Then Exit Function

    For alignmentShift = 1 To MAX_HEADER_ALIGNMENT_SHIFT
        headerRowAbs = (bestRowIndex - rowLower) + alignmentShift
        If headerRowAbs > 0 Then
            If m_TryBuildValidatedHeaderRangeRef(adoConn, sheetPrefix, headerRowAbs, colLetter, keyHeader, outDetectedRef) Then
                m_TryDetectHeaderRangeFromTopRows = True
                Exit Function
            End If
        End If
    Next alignmentShift

    fallbackRowAbs = (bestRowIndex - rowLower) + 1
    If fallbackRowAbs <= 0 Then Exit Function

    outDetectedRef = "[" & sheetPrefix & "A" & CStr(fallbackRowAbs) & ":" & colLetter & "1048576]"
    m_TryDetectHeaderRangeFromTopRows = True
    Exit Function

EH:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    On Error GoTo 0
End Function

Public Function m_TryBuildValidatedHeaderRangeRef( _
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
    If m_RecordsetGetFieldOrdinal(rs, keyHeader) >= 0 Then
        outRangeRef = candidateRef
        m_TryBuildValidatedHeaderRangeRef = True
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

Public Function m_BuildFieldOrdinals( _
    ByVal cfg As Object, _
    ByVal rs As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fields As Collection, _
    ByVal errSource As String, _
    ByVal errHeaderNotFoundCode As Long, _
    ByVal errMappedHeaderEmptyCode As Long, _
    Optional ByVal errMissingCfgCode As Long = vbObjectError + 1, _
    Optional ByVal errEmptyCfgCode As Long = vbObjectError + 2 _
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
        exactToken = m_NormalizeHeader(fieldName)
        looseToken = m_NormalizeHeaderLoose(fieldName)
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
        desiredHeader = m_MappedHeader( _
            cfg, sourceAlias, tableAlias, fieldAlias, errSource, errMappedHeaderEmptyCode, errMissingCfgCode, errEmptyCfgCode)

        exactToken = m_NormalizeHeader(desiredHeader)
        looseToken = m_NormalizeHeaderLoose(desiredHeader)

        If Len(exactToken) > 0 And byExact.Exists(exactToken) Then
            result(fieldAlias) = CLng(byExact(exactToken))
        ElseIf Len(looseToken) > 0 And byLoose.Exists(looseToken) Then
            result(fieldAlias) = CLng(byLoose(looseToken))
        Else
            availableFields = m_ListRecordsetFields(rs, 40)
            If m_RecordsetLooksLikeGenericFields(rs) Then
                hintText = " Hint: ADO returned generic fields (F1..Fn). Set SheetName as explicit range with header row, e.g. 'Аркуш1$A3:I1048576'."
            End If
            Err.Raise errHeaderNotFoundCode, errSource, _
                "Configured source header '" & desiredHeader & "' is not found for " & sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]. " & _
                "Available fields: " & availableFields & "." & hintText
        End If
    Next i

    Set m_BuildFieldOrdinals = result
End Function

Public Function m_RecordsetFieldNameByOrdinal( _
    ByVal rs As Object, _
    ByVal fieldOrdinal As Long, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAlias As String, _
    ByVal errSource As String, _
    ByVal errRecordsetNotInitializedCode As Long, _
    ByVal errOrdinalOutOfRangeCode As Long _
) As String
    If rs Is Nothing Then
        Err.Raise errRecordsetNotInitializedCode, errSource, _
            "Recordset is not initialized while resolving source field name."
    End If

    If fieldOrdinal < 0 Or fieldOrdinal >= rs.Fields.Count Then
        Err.Raise errOrdinalOutOfRangeCode, errSource, _
            "Field ordinal is out of range for " & sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]."
    End If

    m_RecordsetFieldNameByOrdinal = CStr(rs.Fields(fieldOrdinal).Name)
End Function

Public Function m_AppendFilteredRows( _
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
                m_AddResultCell resultTable, rowIndex, sourceAlias, tableAlias, fieldAlias, valueText
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

    m_AppendFilteredRows = (nextRow > startRow)
End Function

Public Function m_ToSafeText(ByVal valueIn As Variant) As String
    If IsError(valueIn) Then Exit Function
    If IsNull(valueIn) Then Exit Function
    If IsEmpty(valueIn) Then Exit Function

    m_ToSafeText = Trim$(CStr(valueIn))
End Function

Public Function m_IsEmptyVariantArray(ByVal valueRef As Variant) As Boolean
    Dim lb As Long
    Dim ub As Long

    If IsEmpty(valueRef) Then
        m_IsEmptyVariantArray = True
        Exit Function
    End If
    If Not IsArray(valueRef) Then
        m_IsEmptyVariantArray = True
        Exit Function
    End If

    On Error GoTo ErrHandler
    lb = LBound(valueRef)
    ub = UBound(valueRef)
    m_IsEmptyVariantArray = (ub < lb)
    Exit Function

ErrHandler:
    m_IsEmptyVariantArray = True
End Function

Public Function m_ToColumnLetter(ByVal columnIndex As Long) As String
    Dim n As Long
    Dim remainder As Long

    If columnIndex < 1 Then columnIndex = 1
    n = columnIndex

    Do While n > 0
        remainder = (n - 1) Mod 26
        m_ToColumnLetter = Chr$(65 + remainder) & m_ToColumnLetter
        n = (n - remainder - 1) \ 26
    Loop
End Function

Public Function m_NormalizeHeader(ByVal valueText As String) As String
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

    m_NormalizeHeader = normalized
End Function

Public Function m_NormalizeHeaderLoose(ByVal valueText As String) As String
    Dim normalized As String
    Dim i As Long
    Dim ch As String
    Dim codePoint As Long
    Dim resultText As String

    normalized = m_NormalizeHeader(valueText)
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

    m_NormalizeHeaderLoose = resultText
End Function

Public Function m_ListRecordsetFields(ByVal rs As Object, Optional ByVal maxCount As Long = 25) As String
    Dim i As Long
    Dim count As Long
    Dim fieldName As String

    If rs Is Nothing Then Exit Function
    If maxCount <= 0 Then maxCount = 25

    For i = 0 To rs.Fields.Count - 1
        fieldName = Trim$(CStr(rs.Fields(i).Name))
        If Len(fieldName) = 0 Then fieldName = "(empty)"
        If count > 0 Then m_ListRecordsetFields = m_ListRecordsetFields & ", "
        m_ListRecordsetFields = m_ListRecordsetFields & "[" & fieldName & "]"
        count = count + 1
        If count >= maxCount Then Exit For
    Next i

    If rs.Fields.Count > maxCount Then
        m_ListRecordsetFields = m_ListRecordsetFields & ", ..."
    End If
End Function

Public Function m_RecordsetLooksLikeGenericFields(ByVal rs As Object) As Boolean
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

    m_RecordsetLooksLikeGenericFields = True
End Function

Public Function m_RecordsetGetFieldOrdinal(ByVal rs As Object, ByVal fieldName As String) As Long
    Dim i As Long
    Dim targetToken As String

    m_RecordsetGetFieldOrdinal = -1
    If rs Is Nothing Then Exit Function

    targetToken = m_NormalizeHeader(fieldName)
    If Len(targetToken) = 0 Then Exit Function

    For i = 0 To rs.Fields.Count - 1
        If StrComp(m_NormalizeHeader(CStr(rs.Fields(i).Name)), targetToken, vbTextCompare) = 0 Then
            m_RecordsetGetFieldOrdinal = i
            Exit Function
        End If
    Next i
End Function

Public Function m_CreateResultTableFromFields( _
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

    Set m_CreateResultTableFromFields = tableObj
End Function

Public Sub m_AddResultCell( _
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

Private Sub mp_EnsureLikeDialectCache()
    If g_LikeDialectByConnection Is Nothing Then
        Set g_LikeDialectByConnection = CreateObject("Scripting.Dictionary")
        g_LikeDialectByConnection.CompareMode = 1
    End If
End Sub

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
    sqlText = "SELECT IIF('abc' LIKE " & m_QuoteSqlStringLiteral(patternText) & ", 1, 0) AS Hit"
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

Private Sub mp_EnsureLongTextRuntimeCacheStorage()
    If g_LongTextRuntimeCacheBySignature Is Nothing Then
        Set g_LongTextRuntimeCacheBySignature = CreateObject("Scripting.Dictionary")
        g_LongTextRuntimeCacheBySignature.CompareMode = 1
    End If
End Sub

Private Function mp_NormalizeLongTextRuntimeCacheSignature(ByVal cacheSignature As String) As String
    mp_NormalizeLongTextRuntimeCacheSignature = LCase$(Trim$(cacheSignature))
End Function

Private Function mp_BuildLongTextCellKey(ByVal outRowIndex As Long, ByVal fieldIndex As Long) As String
    mp_BuildLongTextCellKey = CStr(outRowIndex) & "|" & CStr(fieldIndex)
End Function

Private Function mp_TryParseLongTextCellKey(ByVal cellKey As String, ByRef outRowIndex As Long, ByRef outFieldIndex As Long) As Boolean
    Dim sepPos As Long
    Dim leftPart As String
    Dim rightPart As String

    sepPos = InStr(1, cellKey, "|", vbBinaryCompare)
    If sepPos <= 1 Then Exit Function
    If sepPos >= Len(cellKey) Then Exit Function

    leftPart = Mid$(cellKey, 1, sepPos - 1)
    rightPart = Mid$(cellKey, sepPos + 1)
    If Not IsNumeric(leftPart) Then Exit Function
    If Not IsNumeric(rightPart) Then Exit Function

    outRowIndex = CLng(leftPart)
    outFieldIndex = CLng(rightPart)
    mp_TryParseLongTextCellKey = True
End Function

Private Function mp_BuildLongTextRuntimeCacheKey( _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal configuredSheetName As String, _
    ByVal keyValue As String, _
    ByVal outRowIndex As Long, _
    ByVal fieldIndex As Long, _
    ByVal prefixToken As String _
) As String
    mp_BuildLongTextRuntimeCacheKey = LCase$(Trim$(sourceAlias)) & "|" & _
                                     LCase$(Trim$(tableAlias)) & "|" & _
                                     LCase$(Trim$(configuredSheetName)) & "|" & _
                                     Trim$(keyValue) & "|" & _
                                     "idx:" & CStr(outRowIndex) & "|" & _
                                     CStr(fieldIndex) & "|" & _
                                     prefixToken
End Function

Private Function mp_BuildLongTextLooseCacheKey( _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal configuredSheetName As String, _
    ByVal keyValue As String, _
    ByVal fieldIndex As Long, _
    ByVal prefixToken As String _
) As String
    mp_BuildLongTextLooseCacheKey = LCase$(Trim$(sourceAlias)) & "|" & _
                                    LCase$(Trim$(tableAlias)) & "|" & _
                                    LCase$(Trim$(configuredSheetName)) & "|" & _
                                    Trim$(keyValue) & "|" & _
                                    "fld:" & CStr(fieldIndex) & "|" & _
                                    prefixToken
End Function

Private Function mp_BuildLongTextPrefixToken(ByVal valueText As String) As String
    Dim token As String

    token = CStr(valueText)
    token = Replace$(token, vbCrLf, vbLf)
    token = Replace$(token, vbCr, vbLf)
    token = Replace$(token, vbLf, " ")
    token = Replace$(token, vbTab, " ")
    token = Replace$(token, ChrW$(160), " ")

    Do While InStr(1, token, "  ", vbBinaryCompare) > 0
        token = Replace$(token, "  ", " ")
    Loop

    token = Trim$(token)
    If Len(token) > 255 Then token = Left$(token, 255)
    mp_BuildLongTextPrefixToken = LCase$(token)
End Function

Private Function mp_IsSameHydrationKeyValue(ByVal leftValue As Variant, ByVal rightText As String) As Boolean
    Dim leftTrimmed As String
    Dim rightTrimmed As String
    Dim leftLoose As String
    Dim rightLoose As String

    leftTrimmed = Trim$(CStr(leftValue))
    rightTrimmed = Trim$(rightText)

    If StrComp(leftTrimmed, rightTrimmed, vbTextCompare) = 0 Then
        mp_IsSameHydrationKeyValue = True
        Exit Function
    End If

    leftLoose = m_NormalizeHeaderLoose(leftTrimmed)
    rightLoose = m_NormalizeHeaderLoose(rightTrimmed)
    If Len(leftLoose) = 0 Or Len(rightLoose) = 0 Then Exit Function

    mp_IsSameHydrationKeyValue = (StrComp(leftLoose, rightLoose, vbTextCompare) = 0)
End Function

Private Function mp_BuildFieldPrefixLookupKey(ByVal fieldIndex As Long, ByVal prefixToken As String) As String
    mp_BuildFieldPrefixLookupKey = CStr(fieldIndex) & "|" & prefixToken
End Function

Private Sub mp_AddRowToMatchBucket(ByVal buckets As Object, ByVal lookupKey As String, ByVal sourceRow As Long)
    Dim rowsBucket As Collection

    If buckets Is Nothing Then Exit Sub
    If Len(lookupKey) = 0 Then Exit Sub

    If buckets.Exists(lookupKey) Then
        Set rowsBucket = buckets(lookupKey)
    Else
        Set rowsBucket = New Collection
        Set buckets(lookupKey) = rowsBucket
    End If

    rowsBucket.Add sourceRow
End Sub

Private Function mp_CollectionContainsLong(ByVal values As Collection, ByVal needle As Long) As Boolean
    Dim i As Long

    If values Is Nothing Then Exit Function
    If values.Count = 0 Then Exit Function

    For i = 1 To values.Count
        If CLng(values(i)) = needle Then
            mp_CollectionContainsLong = True
            Exit Function
        End If
    Next i
End Function

Private Function mp_GetWorkbookForSource(ByVal wbCache As Object, ByVal cfg As Object, ByVal sourceAlias As String) As Workbook
    Dim sourcePath As String
    Dim snapshotPath As String
    Dim wb As Workbook

    If wbCache Is Nothing Then Exit Function

    If wbCache.Exists(sourceAlias) Then
        Set mp_GetWorkbookForSource = wbCache(sourceAlias)
        Exit Function
    End If

    sourcePath = m_GetResolvedSourcePath( _
        cfg, _
        sourceAlias, _
        "ex_ResultSqlEngine", _
        vbObjectError + 1762, _
        vbObjectError + 1760, _
        vbObjectError + 1761, _
        vbObjectError + 1370, _
        vbObjectError + 1371)

    If Dir(sourcePath) = vbNullString Then
        Err.Raise vbObjectError + 1360, "ex_ResultSqlEngine", "Source file not found: " & sourcePath
    End If

    snapshotPath = ex_SourceSnapshot.m_GetSnapshotPath(sourcePath, "Source." & sourceAlias)
    Set wb = Workbooks.Open(Filename:=snapshotPath, ReadOnly:=True, UpdateLinks:=0)

    On Error Resume Next
    wb.Windows(1).Visible = False
    On Error GoTo 0

    wbCache.Add sourceAlias, wb
    Set mp_GetWorkbookForSource = wb
End Function

Private Function mp_GetMappedSourceHeaderForHydration( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAlias As String _
) As String
    Dim mappedHeader As String

    mappedHeader = m_MappedHeader( _
        cfg, _
        sourceAlias, _
        tableAlias, _
        fieldAlias, _
        "ex_ResultSqlEngine", _
        vbObjectError + 1390, _
        vbObjectError + 1370, _
        vbObjectError + 1371)

    mappedHeader = Trim$(mappedHeader)
    If Len(mappedHeader) >= 2 Then
        If Left$(mappedHeader, 1) = "[" And Right$(mappedHeader, 1) = "]" Then
            mappedHeader = Trim$(Mid$(mappedHeader, 2, Len(mappedHeader) - 2))
        End If
    End If

    If Len(mappedHeader) = 0 Then
        Err.Raise vbObjectError + 1390, "ex_ResultSqlEngine", _
            "Mapped source header is empty for " & sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]"
    End If

    mp_GetMappedSourceHeaderForHydration = mappedHeader
End Function

Private Function mp_FindFieldColumnsForHydration( _
    ByVal fields As Variant, _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal ws As Worksheet, _
    ByVal headerRow As Long, _
    ByVal markerCol As Long _
) As Object
    Dim fieldColsByIdx As Object
    Dim i As Long
    Dim fieldAlias As String
    Dim fieldHeaderName As String
    Dim fieldCol As Long

    Set fieldColsByIdx = CreateObject("Scripting.Dictionary")
    fieldColsByIdx.CompareMode = 1

    For i = LBound(fields) To UBound(fields)
        fieldAlias = Trim$(CStr(fields(i)))
        If Len(fieldAlias) = 0 Then GoTo ContinueFieldAlias
        If ex_FetchDslEngine.m_IsVirtualFieldAlias(cfg, sourceAlias, tableAlias, fieldAlias) Then GoTo ContinueFieldAlias

        fieldHeaderName = mp_GetMappedSourceHeaderForHydration(cfg, sourceAlias, tableAlias, fieldAlias)
        fieldCol = m_FindHeaderColumnInWorksheetRow(ws, headerRow, markerCol, fieldHeaderName)
        If fieldCol > 0 Then
            fieldColsByIdx(CStr(i)) = fieldCol
        End If
ContinueFieldAlias:
    Next i

    Set mp_FindFieldColumnsForHydration = fieldColsByIdx
End Function

Private Function mp_ExtractSheetNameToken(ByVal configuredSheetName As String) As String
    Dim token As String
    Dim dollarPos As Long

    token = Trim$(configuredSheetName)
    If Len(token) = 0 Then Exit Function

    If Left$(token, 1) = "[" And Right$(token, 1) = "]" Then
        token = Mid$(token, 2, Len(token) - 2)
    End If
    token = mp_CleanAdoSchemaObjectName(token)

    dollarPos = InStr(1, token, "$", vbBinaryCompare)
    If dollarPos > 0 Then
        token = Left$(token, dollarPos - 1)
    End If

    mp_ExtractSheetNameToken = Trim$(token)
End Function

Private Function mp_NormalizeAdoObjectNameExact(ByVal value As String) As String
    mp_NormalizeAdoObjectNameExact = LCase$(Trim$(mp_CleanAdoSchemaObjectName(value)))
End Function

Private Function mp_CleanAdoSchemaObjectName(ByVal value As String) As String
    Dim cleaned As String

    cleaned = Trim$(value)
    If Len(cleaned) = 0 Then Exit Function

    If Left$(cleaned, 1) = "[" And Right$(cleaned, 1) = "]" Then
        cleaned = Mid$(cleaned, 2, Len(cleaned) - 2)
    End If

    cleaned = Replace$(cleaned, "]]", "]")
    cleaned = Replace$(cleaned, "'", vbNullString)
    cleaned = Replace$(cleaned, "`", vbNullString)

    mp_CleanAdoSchemaObjectName = Trim$(cleaned)
End Function

Private Sub mp_DebugHydrationLog(ByVal stage As String, ByVal messageText As String)
    On Error Resume Next
    ex_Messaging.m_LogToFile "[ex_ResultSqlEngine][" & CStr(stage) & "] " & CStr(messageText), HYDRATION_DEBUG_LOG_PATH
    On Error GoTo 0
End Sub

Private Function mp_DebugToken(ByVal valueText As String, Optional ByVal maxLen As Long = 120) As String
    Dim token As String

    token = CStr(valueText)
    token = Replace$(token, vbCrLf, " ")
    token = Replace$(token, vbCr, " ")
    token = Replace$(token, vbLf, " ")
    token = Trim$(token)

    If maxLen > 0 Then
        If Len(token) > maxLen Then
            token = Left$(token, maxLen) & "..."
        End If
    End If

    mp_DebugToken = token
End Function

Private Function mp_AsText(ByVal valueIn As Variant, Optional ByVal adoFieldType As Long = -1) As String
    mp_AsText = ex_SqlAdoHelpers.m_ToNormalizedText(valueIn, adoFieldType)
End Function
