Attribute VB_Name = "ex_ResultSqlEngine"
Option Explicit

Private Const LIKE_DIALECT_UNKNOWN As String = "unknown"
Private Const LIKE_DIALECT_STAR As String = "star"
Private Const LIKE_DIALECT_PERCENT As String = "percent"

Private g_LikeDialectByConnection As Object

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
            props = "Excel 8.0;HDR=YES;IMEX=1;ReadOnly=True"
        Case "xlsx"
            props = "Excel 12.0 Xml;HDR=YES;IMEX=1;ReadOnly=True"
        Case "xlsm"
            props = "Excel 12.0 Macro;HDR=YES;IMEX=1;ReadOnly=True"
        Case "xlsb"
            props = "Excel 12.0;HDR=YES;IMEX=1;ReadOnly=True"
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

Private Function mp_AsText(ByVal valueIn As Variant, Optional ByVal adoFieldType As Long = -1) As String
    mp_AsText = ex_SqlAdoHelpers.m_ToNormalizedText(valueIn, adoFieldType)
End Function
