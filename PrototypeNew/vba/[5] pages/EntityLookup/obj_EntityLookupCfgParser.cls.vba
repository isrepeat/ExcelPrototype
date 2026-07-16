VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_EntityLookupCfgParser"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const LOOKUP_ALL_TOKEN As String = "*"
Private Const LOOKUP_MAX_ROWS As Long = 30

Private Const RUNTIME_ERROR_TITLE As String = "PrototypeNew / EntityLookup runtime"
Private Const MODE_PREFIX As String = "EntityLookup.Column["
Private Const TABLE_PREFIX As String = "EntityLookup.Table.Column["
Private Const LOOKUP_PREFIX As String = "EntityLookup.Lookup["

Private m_CfgTableParser As obj_CfgTableParser
Private m_IsDisposed As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & TypeName(Me) & ".Class_Initialize"
#End If
    Set m_CfgTableParser = New obj_CfgTableParser
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Dispose
    On Error GoTo 0
End Sub

' //
' // API
' //
Public Function Initialize(ByVal configTable As obj_ConfigTable) As Boolean
    m_IsDisposed = False
    If m_CfgTableParser Is Nothing Then Set m_CfgTableParser = New obj_CfgTableParser
    If m_CfgTableParser Is Nothing Then Exit Function
    If Not m_CfgTableParser.Initialize(configTable, Me) Then Exit Function

    Initialize = True
End Function

Public Sub Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    If Not m_CfgTableParser Is Nothing Then m_CfgTableParser.Dispose
    Set m_CfgTableParser = Nothing
    On Error GoTo 0
End Sub

Public Function ResolveLatestByDmyPattern(ByVal rawValue As String) As String
    ResolveLatestByDmyPattern = ex_SourceResolver.fn_ResolveLatestByDmyPattern(rawValue)
End Function

Public Function TryGetLookupColumnKeys(ByRef outColumnKeys As Collection) As Boolean
    Dim cfgMap As Object
    Dim rawColumns As String

    Set outColumnKeys = Nothing
    If Not private_TryBuildConfigMap(cfgMap) Then Exit Function

    rawColumns = m_CfgTableParser.CfgParserBase.GetOptionalConfigValue(cfgMap, "EntityLookup.Table.Columns")
    If VBA.Len(VBA.Trim$(rawColumns)) = 0 Then
        rawColumns = m_CfgTableParser.CfgParserBase.GetOptionalConfigValue(cfgMap, "EntityLookup.Columns", "FIO")
    End If
    Set outColumnKeys = m_CfgTableParser.CfgParserBase.SplitListToCollection(rawColumns)
    If outColumnKeys Is Nothing Then Exit Function
    If outColumnKeys.Count <= 0 Then
        private_ShowConfigError "EntityLookup.Columns is empty."
        Exit Function
    End If

    TryGetLookupColumnKeys = True
End Function

Public Function TryGetColumnCaption( _
    ByVal columnKey As String, _
    ByRef outCaption As String _
) As Boolean
    Dim cfgMap As Object

    outCaption = VBA.Trim$(columnKey)
    If VBA.Len(outCaption) = 0 Then Exit Function
    If Not private_TryBuildConfigMap(cfgMap) Then Exit Function

    outCaption = m_CfgTableParser.CfgParserBase.GetOptionalConfigValue( _
        cfgMap, _
        private_BuildTableColumnPrefix(columnKey) & "Caption", _
        outCaption)

    outCaption = m_CfgTableParser.CfgParserBase.GetOptionalConfigValue( _
        cfgMap, _
        private_BuildColumnPrefix(columnKey) & "Caption", _
        outCaption)
    If VBA.Len(VBA.Trim$(outCaption)) = 0 Then outCaption = VBA.Trim$(columnKey)

    TryGetColumnCaption = True
End Function

Public Function TryGetLookupKeys(ByRef outLookupKeys As Collection) As Boolean
    Dim cfgMap As Object
    Dim rawLookups As String

    Set outLookupKeys = Nothing
    If Not private_TryBuildConfigMap(cfgMap) Then Exit Function

    rawLookups = m_CfgTableParser.CfgParserBase.GetOptionalConfigValue(cfgMap, "EntityLookup.Lookups")
    If VBA.Len(VBA.Trim$(rawLookups)) = 0 Then
        private_ShowConfigError "EntityLookup.Lookups is empty."
        Exit Function
    End If

    Set outLookupKeys = m_CfgTableParser.CfgParserBase.SplitListToCollection(rawLookups)
    If outLookupKeys Is Nothing Then Exit Function
    If outLookupKeys.Count <= 0 Then
        private_ShowConfigError "EntityLookup.Lookups is empty."
        Exit Function
    End If

    TryGetLookupKeys = True
End Function

Public Function TryGetLookupCandidatesConfig( _
    ByVal lookupKey As String, _
    ByRef outMinCount As Long, _
    ByRef outSectionCaption As String _
) As Boolean
    Dim cfgMap As Object
    Dim lookupPrefix As String

    outMinCount = 2
    outSectionCaption = "Candidates"

    lookupKey = VBA.Trim$(lookupKey)
    If VBA.Len(lookupKey) = 0 Then Exit Function
    If Not private_TryBuildConfigMap(cfgMap) Then Exit Function

    lookupPrefix = private_BuildLookupPrefix(lookupKey)
    If VBA.Len(lookupPrefix) = 0 Then Exit Function

    outMinCount = private_GetOptionalConfigLong(cfgMap, lookupPrefix & "CandidatesMinCount", outMinCount)
    outSectionCaption = m_CfgTableParser.CfgParserBase.GetOptionalConfigValue(cfgMap, lookupPrefix & "CandidatesSectionCaption", outSectionCaption)

    If outMinCount <= 0 Then outMinCount = 2
    If VBA.Len(VBA.Trim$(outSectionCaption)) = 0 Then outSectionCaption = "Candidates"

    TryGetLookupCandidatesConfig = True
End Function

Public Function TryGetLookupInputGridColumn( _
    ByVal lookupKey As String, _
    ByRef outInputGridColumn As Long _
) As Boolean
    Dim cfgMap As Object
    Dim targetColumnKey As String
    Dim tableColumnKeys As Collection
    Dim columnKeyObj As Variant
    Dim columnKey As String
    Dim currentGridCol As Long

    outInputGridColumn = 0

    lookupKey = VBA.Trim$(lookupKey)
    If VBA.Len(lookupKey) = 0 Then Exit Function
    If Not private_TryBuildConfigMap(cfgMap) Then Exit Function

    If Not private_TryGetLookupTargetColumn(cfgMap, lookupKey, targetColumnKey) Then Exit Function
    If Not private_TryGetTableColumnKeysFromMap(cfgMap, tableColumnKeys) Then Exit Function

    currentGridCol = 1
    For Each columnKeyObj In tableColumnKeys
        columnKey = VBA.Trim$(VBA.CStr(columnKeyObj))
        If VBA.Len(columnKey) = 0 Then GoTo ContinueColumn

        If VBA.StrComp(columnKey, targetColumnKey, VBA.vbTextCompare) = 0 Then
            outInputGridColumn = currentGridCol
            TryGetLookupInputGridColumn = True
            Exit Function
        End If

        currentGridCol = currentGridCol + 1
ContinueColumn:
    Next columnKeyObj

    private_ShowConfigError "Lookup target column '" & targetColumnKey & "' for lookup '" & lookupKey & "' is not listed in EntityLookup.Table.Columns."
End Function

Public Function TryGetLookupTargetColumn( _
    ByVal lookupKey As String, _
    ByRef outTargetColumnKey As String _
) As Boolean
    Dim cfgMap As Object

    outTargetColumnKey = VBA.vbNullString
    If Not private_TryBuildConfigMap(cfgMap) Then Exit Function
    TryGetLookupTargetColumn = private_TryGetLookupTargetColumn(cfgMap, lookupKey, outTargetColumnKey)
End Function

Public Function TryBuildLookupSqlParams( _
    ByVal columnKey As String, _
    ByVal queryText As String, _
    ByRef outSqlParams As obj_SqlParams, _
    ByRef outSearchColumnAlias As String, _
    ByRef outResultColumnAliases As Collection _
) As Boolean
    Dim cfgMap As Object
    Dim columnPrefix As String
    Dim sourceAlias As String
    Dim tableAlias As String
    Dim searchColumnRef As String
    Dim searchColumnAlias As String
    Dim tablePathPrefix As String
    Dim sourcePath As String
    Dim sheetName As String
    Dim rangeStartMarker As String
    Dim rangeEndMarker As String
    Dim searchSourceHeader As String
    Dim searchMappedHeader As String
    Dim resultAliasObj As Variant
    Dim resultAliasToken As String
    Dim sourceResultAlias As String
    Dim targetResultAlias As String
    Dim sourceColumnHeader As String
    Dim mappedColumnHeader As String
    Dim outputMappedHeader As String
    Dim sqlParams As obj_SqlParams
    Dim notBlankCondition As String
    Dim containsCondition As String
    Dim normalizedAliases As Collection
    Dim columnAliasSet As Object
    Dim selectedSourceAliases As Object
    Dim allSourceAliases As Collection
    Dim auxiliaryAliasObj As Variant
    Dim auxiliaryAlias As String

    Set outSqlParams = Nothing
    Set outResultColumnAliases = Nothing
    outSearchColumnAlias = VBA.vbNullString

    columnKey = VBA.Trim$(columnKey)
    queryText = VBA.Trim$(queryText)
    If VBA.Len(columnKey) = 0 Then
        private_ShowConfigError "Lookup column key is empty."
        Exit Function
    End If
    If VBA.Len(queryText) = 0 Then
        private_ShowConfigError "Enter search text for column '" & columnKey & "'."
        Exit Function
    End If

    If Not private_TryBuildConfigMap(cfgMap) Then Exit Function
    columnPrefix = private_BuildLookupPrefix(columnKey)
    If VBA.Len(columnPrefix) = 0 Then Exit Function

    If Not private_TryGetRequiredConfigValue(cfgMap, columnPrefix & "SearchColumnAlias", searchColumnRef) Then Exit Function
    If Not m_CfgTableParser.TryParseColumnRefToken(searchColumnRef, sourceAlias, tableAlias, searchColumnAlias) Then
        private_ShowConfigError "Invalid SearchColumnAlias for lookup '" & columnKey & "'. Expected format: SourceAlias.Sheet[TableAlias].Column[ColumnAlias]."
        Exit Function
    End If
    outSearchColumnAlias = searchColumnAlias

    tablePathPrefix = m_CfgTableParser.BuildTablePathPrefix(sourceAlias, tableAlias)
    If VBA.Len(tablePathPrefix) = 0 Then
        private_ShowConfigError "Failed to build table path prefix for column '" & columnKey & "'."
        Exit Function
    End If
    If Not private_TryBuildColumnAliasSet(cfgMap, tablePathPrefix, columnAliasSet) Then Exit Function
    If Not private_TryValidateDeclaredColumnAlias(columnAliasSet, tablePathPrefix, searchColumnAlias, "SearchColumnAlias for lookup '" & columnKey & "'") Then Exit Function

    If Not m_CfgTableParser.TryResolveSourcePath(cfgMap, sourceAlias, sourcePath) Then
        private_ShowConfigError "Failed to resolve source path for alias '" & sourceAlias & "'."
        Exit Function
    End If
    If Not m_CfgTableParser.TryResolveSheetName(cfgMap, tablePathPrefix, sheetName) Then
        private_ShowConfigError "Missing sheet name for table prefix '" & tablePathPrefix & "'."
        Exit Function
    End If
    If Not m_CfgTableParser.TryResolveRangeMarkers(cfgMap, tablePathPrefix, rangeStartMarker, rangeEndMarker) Then
        private_ShowConfigError "Invalid range markers for table prefix '" & tablePathPrefix & "'."
        Exit Function
    End If

    If Not m_CfgTableParser.TryResolveMapByColumnAlias(cfgMap, tablePathPrefix, searchColumnAlias, searchSourceHeader, searchMappedHeader) Then
        private_ShowConfigError "Failed to resolve search column alias '" & searchColumnAlias & "' for table prefix '" & tablePathPrefix & "'."
        Exit Function
    End If

    Set normalizedAliases = private_BuildQualifiedResultColumnMappings( _
        cfgMap, columnPrefix, sourceAlias, tableAlias, columnKey)
    If normalizedAliases Is Nothing Then Exit Function
    If normalizedAliases.Count <= 0 Then
        private_ShowConfigError "No qualified ResultColumnsAliases mappings are configured for lookup '" & columnKey & "'."
        Exit Function
    End If
    Set selectedSourceAliases = ex_Helpers.fn_CreateDictionaryTextCompare()

    ' Каждый маппинг задаётся в конфиге полностью квалифицированными ссылками:
    ' ResultColumnsAliases[Source.Sheet[Table].Column[SourceAlias]] =
    '     EntityLookup.Table.Column[TargetAlias].
    ' Внутри запись нормализуется в SourceAlias=>TargetAlias для SQL builder.
    For Each resultAliasObj In normalizedAliases
        resultAliasToken = VBA.Trim$(VBA.CStr(resultAliasObj))
        If VBA.Len(resultAliasToken) = 0 Then GoTo ContinueValidateResultAlias
        If Not private_TryParseResultAliasMapping(resultAliasToken, sourceResultAlias, targetResultAlias) Then
            private_ShowConfigError "Invalid normalized ResultColumnsAliases mapping '" & resultAliasToken & "' for lookup '" & columnKey & "'."
            Exit Function
        End If
        If Not private_TryValidateDeclaredColumnAlias(columnAliasSet, tablePathPrefix, sourceResultAlias, "ResultColumnsAliases for lookup '" & columnKey & "'") Then Exit Function
        selectedSourceAliases(sourceResultAlias) = True
        If VBA.StrComp(sourceResultAlias, searchColumnAlias, VBA.vbTextCompare) = 0 Then outSearchColumnAlias = targetResultAlias
ContinueValidateResultAlias:
    Next resultAliasObj

    Set sqlParams = New obj_SqlParams
    sqlParams.SourcePath = sourcePath
    sqlParams.SheetName = sheetName
    sqlParams.RangeStartMarker = rangeStartMarker
    sqlParams.RangeEndMarker = rangeEndMarker
    ' Лимит применяется самим движком запроса: SQL использует TOP 30,
    ' а универсальный движок прекращает чтение после тридцатой строки.
    sqlParams.MaxRows = LOOKUP_MAX_ROWS
    notBlankCondition = ex_HelpersSql.fn_BuildWhereNotBlankSql(searchSourceHeader)
    If VBA.Len(notBlankCondition) = 0 Then
        private_ShowConfigError "Failed to build non-empty search condition for column '" & columnKey & "'."
        Exit Function
    End If

    If VBA.StrComp(VBA.Trim$(queryText), LOOKUP_ALL_TOKEN, VBA.vbBinaryCompare) = 0 Then
        ' "*" отключает только текстовую фильтрацию LIKE. Общий фильтр
        ' пустых строк сохраняется, поэтому TOP N считается по заполненным
        ' записям, а хвост пустого диапазона не попадает в candidates.
        sqlParams.WhereConditions = notBlankCondition
    Else
        containsCondition = ex_HelpersSql.fn_BuildWhereContainsSql(searchSourceHeader, queryText)
        If VBA.Len(containsCondition) = 0 Then
            private_ShowConfigError "Failed to build search condition for column '" & columnKey & "'."
            Exit Function
        End If
        ' Фильтр непустых значений применяется ко всем lookup-запросам,
        ' независимо от конкретного внешнего справочника.
        sqlParams.WhereConditions = "(" & notBlankCondition & ") AND (" & containsCondition & ")"
    End If

    For Each resultAliasObj In normalizedAliases
        resultAliasToken = VBA.Trim$(VBA.CStr(resultAliasObj))
        If VBA.Len(resultAliasToken) = 0 Then GoTo ContinueResultAlias
        If Not private_TryParseResultAliasMapping(resultAliasToken, sourceResultAlias, targetResultAlias) Then Exit Function

        If Not m_CfgTableParser.TryResolveMapByColumnAlias(cfgMap, tablePathPrefix, sourceResultAlias, sourceColumnHeader, mappedColumnHeader) Then
            private_ShowConfigError "Failed to resolve result alias '" & sourceResultAlias & "' for table prefix '" & tablePathPrefix & "'."
            Exit Function
        End If
        ' В SQL читаем физический source header, но наружу отдаём target alias.
        ' Так candidate table может сразу заполнять колонку формы, не создавая
        ' дубликаты source-полей с теми же подписями.
        outputMappedHeader = private_GetOutputColumnCaption(cfgMap, targetResultAlias, mappedColumnHeader)
        If Not sqlParams.AddColumnMapping(sourceColumnHeader, outputMappedHeader, targetResultAlias) Then
            private_ShowConfigError "Failed to add SQL column mapping for alias '" & targetResultAlias & "'."
            Exit Function
        End If
ContinueResultAlias:
    Next resultAliasObj

    ' ColumnHeadersAliases задаёт полный контракт выборки источника. Колонки без
    ' цели в ResultColumnsAliases остаются служебными: проекция выводит их после
    ' активной формы, а код может использовать их для отбора и фильтрации.
    If Not m_CfgTableParser.TryGetRequiredColumnHeadersAliases(cfgMap, tablePathPrefix, allSourceAliases) Then Exit Function
    For Each auxiliaryAliasObj In allSourceAliases
        auxiliaryAlias = VBA.Trim$(VBA.CStr(auxiliaryAliasObj))
        If VBA.Len(auxiliaryAlias) = 0 Then GoTo ContinueAuxiliaryAlias
        If selectedSourceAliases.Exists(auxiliaryAlias) Then GoTo ContinueAuxiliaryAlias
        If Not m_CfgTableParser.TryResolveMapByColumnAlias( _
            cfgMap, tablePathPrefix, auxiliaryAlias, sourceColumnHeader, mappedColumnHeader) Then Exit Function
        If Not sqlParams.AddColumnMapping(sourceColumnHeader, mappedColumnHeader, auxiliaryAlias) Then
            private_ShowConfigError "Failed to add auxiliary SQL column mapping for source alias '" & auxiliaryAlias & "'."
            Exit Function
        End If
ContinueAuxiliaryAlias:
    Next auxiliaryAliasObj

    Set outSqlParams = sqlParams
    Set outResultColumnAliases = normalizedAliases
    TryBuildLookupSqlParams = True
End Function

Private Function private_TryParseResultAliasMapping( _
    ByVal aliasToken As String, _
    ByRef outSourceAlias As String, _
    ByRef outTargetAlias As String _
) As Boolean
    Dim arrowPos As Long

    outSourceAlias = VBA.vbNullString
    outTargetAlias = VBA.vbNullString
    aliasToken = VBA.Trim$(aliasToken)
    If VBA.Len(aliasToken) = 0 Then Exit Function

    ' Формат без стрелки оставляем обратносуместимым: FIO значит FIO=>FIO.
    ' Стрелка задает явное перенаправление результата lookup-а в другую
    ' колонку формы.
    arrowPos = VBA.InStr(1, aliasToken, "=>", VBA.vbBinaryCompare)
    If arrowPos > 0 Then
        outSourceAlias = VBA.Trim$(VBA.Left$(aliasToken, arrowPos - 1))
        outTargetAlias = VBA.Trim$(VBA.Mid$(aliasToken, arrowPos + 2))
    Else
        outSourceAlias = aliasToken
        outTargetAlias = aliasToken
    End If

    private_TryParseResultAliasMapping = (VBA.Len(outSourceAlias) > 0 And VBA.Len(outTargetAlias) > 0)
End Function

Private Function private_GetOutputColumnCaption( _
    ByVal cfgMap As Object, _
    ByVal targetAlias As String, _
    ByVal defaultCaption As String _
) As String
    targetAlias = VBA.Trim$(targetAlias)
    private_GetOutputColumnCaption = VBA.Trim$(defaultCaption)
    If VBA.Len(targetAlias) = 0 Then Exit Function
    If m_CfgTableParser Is Nothing Then Exit Function
    If m_CfgTableParser.CfgParserBase Is Nothing Then Exit Function

    private_GetOutputColumnCaption = m_CfgTableParser.CfgParserBase.GetOptionalConfigValue( _
        cfgMap, _
        private_BuildTableColumnPrefix(targetAlias) & "Caption", _
        private_GetOutputColumnCaption)
End Function

' //
' // Internal
' //
Private Function private_TryBuildConfigMap(ByRef outCfgMap As Object) As Boolean
    Dim configEntries As Collection

    Set outCfgMap = Nothing
    If m_CfgTableParser Is Nothing Then
        private_ShowConfigError "EntityLookup config parser is not initialized."
        Exit Function
    End If
    If m_CfgTableParser.CfgParserBase Is Nothing Then
        private_ShowConfigError "EntityLookup base config parser is not initialized."
        Exit Function
    End If

    If Not m_CfgTableParser.CfgParserBase.TryGetConfigEntries(configEntries) Then
        private_ShowConfigError "Failed to read EntityLookup config entries."
        Exit Function
    End If
    If Not m_CfgTableParser.CfgParserBase.BuildConfigDictionary(configEntries, outCfgMap) Then
        private_ShowConfigError "Failed to build EntityLookup config dictionary."
        Exit Function
    End If
    If outCfgMap Is Nothing Then
        private_ShowConfigError "EntityLookup config dictionary is empty."
        Exit Function
    End If

    private_TryBuildConfigMap = True
End Function

Private Function private_TryGetRequiredConfigValue( _
    ByVal cfgMap As Object, _
    ByVal keyName As String, _
    ByRef outValue As String _
) As Boolean
    outValue = VBA.vbNullString
    If m_CfgTableParser Is Nothing Then Exit Function
    If m_CfgTableParser.CfgParserBase Is Nothing Then Exit Function

    If Not m_CfgTableParser.CfgParserBase.TryGetRequiredConfigValue(cfgMap, keyName, outValue) Then
        private_ShowConfigError "Missing or empty config key: " & keyName
        Exit Function
    End If

    private_TryGetRequiredConfigValue = True
End Function

Private Function private_TryBuildColumnAliasSet( _
    ByVal cfgMap As Object, _
    ByVal tablePathPrefix As String, _
    ByRef outAliasSet As Object _
) As Boolean
    Dim columnAliases As Collection
    Dim aliasObj As Variant
    Dim aliasText As String

    Set outAliasSet = Nothing
    If m_CfgTableParser Is Nothing Then Exit Function

    If Not m_CfgTableParser.TryGetRequiredColumnHeadersAliases(cfgMap, tablePathPrefix, columnAliases) Then
        private_ShowConfigError "Missing or empty config key: " & tablePathPrefix & "ColumnHeadersAliases"
        Exit Function
    End If

    Set outAliasSet = ex_Helpers.fn_CreateDictionaryTextCompare()
    For Each aliasObj In columnAliases
        aliasText = VBA.Trim$(VBA.CStr(aliasObj))
        If VBA.Len(aliasText) = 0 Then GoTo ContinueAlias
        If outAliasSet.Exists(aliasText) Then
            private_ShowConfigError "Duplicate column alias '" & aliasText & "' in " & tablePathPrefix & "ColumnHeadersAliases."
            Exit Function
        End If
        outAliasSet(aliasText) = True
ContinueAlias:
    Next aliasObj

    If outAliasSet.Count <= 0 Then
        private_ShowConfigError "Missing or empty config key: " & tablePathPrefix & "ColumnHeadersAliases"
        Exit Function
    End If

    private_TryBuildColumnAliasSet = True
End Function

Private Function private_TryValidateDeclaredColumnAlias( _
    ByVal columnAliasSet As Object, _
    ByVal tablePathPrefix As String, _
    ByVal columnAlias As String, _
    ByVal usageText As String _
) As Boolean
    columnAlias = VBA.Trim$(columnAlias)
    If VBA.Len(columnAlias) = 0 Then
        private_ShowConfigError "Empty column alias used by " & usageText & "."
        Exit Function
    End If
    If columnAliasSet Is Nothing Then
        private_ShowConfigError "Column aliases are not initialized for table prefix '" & tablePathPrefix & "'."
        Exit Function
    End If
    If Not columnAliasSet.Exists(columnAlias) Then
        private_ShowConfigError "Alias '" & columnAlias & "' used by " & usageText & " is not declared in " & tablePathPrefix & "ColumnHeadersAliases."
        Exit Function
    End If

    private_TryValidateDeclaredColumnAlias = True
End Function

Private Function private_BuildColumnPrefix(ByVal columnKey As String) As String
    columnKey = VBA.Trim$(columnKey)
    If VBA.Len(columnKey) = 0 Then Exit Function
    private_BuildColumnPrefix = MODE_PREFIX & columnKey & "]."
End Function

Private Function private_BuildTableColumnPrefix(ByVal columnKey As String) As String
    columnKey = VBA.Trim$(columnKey)
    If VBA.Len(columnKey) = 0 Then Exit Function
    private_BuildTableColumnPrefix = TABLE_PREFIX & columnKey & "]."
End Function

Private Function private_BuildLookupPrefix(ByVal lookupKey As String) As String
    lookupKey = VBA.Trim$(lookupKey)
    If VBA.Len(lookupKey) = 0 Then Exit Function
    private_BuildLookupPrefix = LOOKUP_PREFIX & lookupKey & "]."
End Function

Private Function private_GetOptionalConfigLong( _
    ByVal cfgMap As Object, _
    ByVal keyName As String, _
    ByVal defaultValue As Long _
) As Long
    Dim rawValue As String

    private_GetOptionalConfigLong = defaultValue
    If m_CfgTableParser Is Nothing Then Exit Function
    If m_CfgTableParser.CfgParserBase Is Nothing Then Exit Function

    rawValue = m_CfgTableParser.CfgParserBase.GetOptionalConfigValue(cfgMap, keyName, VBA.CStr(defaultValue))
    rawValue = VBA.Trim$(rawValue)
    If VBA.Len(rawValue) = 0 Then Exit Function
    If Not VBA.IsNumeric(rawValue) Then Exit Function

    private_GetOptionalConfigLong = VBA.CLng(rawValue)
End Function

Private Function private_TryGetTableColumnKeysFromMap( _
    ByVal cfgMap As Object, _
    ByRef outColumnKeys As Collection _
) As Boolean
    Dim rawColumns As String

    Set outColumnKeys = Nothing
    If m_CfgTableParser Is Nothing Then Exit Function
    If m_CfgTableParser.CfgParserBase Is Nothing Then Exit Function

    rawColumns = m_CfgTableParser.CfgParserBase.GetOptionalConfigValue(cfgMap, "EntityLookup.Table.Columns")
    If VBA.Len(VBA.Trim$(rawColumns)) = 0 Then
        private_ShowConfigError "EntityLookup.Table.Columns is empty."
        Exit Function
    End If

    Set outColumnKeys = m_CfgTableParser.CfgParserBase.SplitListToCollection(rawColumns)
    If outColumnKeys Is Nothing Then Exit Function
    If outColumnKeys.Count <= 0 Then
        private_ShowConfigError "EntityLookup.Table.Columns is empty."
        Exit Function
    End If

    private_TryGetTableColumnKeysFromMap = True
End Function

Private Function private_TryBuildLookupTargetColumnSet( _
    ByVal cfgMap As Object, _
    ByRef outTargetColumns As Object _
) As Boolean
    Dim rawLookups As String
    Dim lookupKeys As Collection
    Dim lookupKeyObj As Variant
    Dim lookupKey As String
    Dim targetColumnKey As String

    Set outTargetColumns = ex_Helpers.fn_CreateDictionaryTextCompare()
    If m_CfgTableParser Is Nothing Then Exit Function
    If m_CfgTableParser.CfgParserBase Is Nothing Then Exit Function

    rawLookups = m_CfgTableParser.CfgParserBase.GetOptionalConfigValue(cfgMap, "EntityLookup.Lookups")
    If VBA.Len(VBA.Trim$(rawLookups)) = 0 Then
        private_ShowConfigError "EntityLookup.Lookups is empty."
        Exit Function
    End If

    Set lookupKeys = m_CfgTableParser.CfgParserBase.SplitListToCollection(rawLookups)
    If lookupKeys Is Nothing Then Exit Function
    If lookupKeys.Count <= 0 Then
        private_ShowConfigError "EntityLookup.Lookups is empty."
        Exit Function
    End If

    For Each lookupKeyObj In lookupKeys
        lookupKey = VBA.Trim$(VBA.CStr(lookupKeyObj))
        If VBA.Len(lookupKey) = 0 Then GoTo ContinueLookup
        If Not private_TryGetLookupTargetColumn(cfgMap, lookupKey, targetColumnKey) Then Exit Function
        If VBA.Len(targetColumnKey) > 0 Then outTargetColumns(VBA.LCase$(targetColumnKey)) = True
ContinueLookup:
    Next lookupKeyObj

    private_TryBuildLookupTargetColumnSet = True
End Function

Private Function private_TryGetLookupTargetColumn( _
    ByVal cfgMap As Object, _
    ByVal lookupKey As String, _
    ByRef outTargetColumnKey As String _
) As Boolean
    Dim lookupPrefix As String
    Dim targetColumnRef As String

    outTargetColumnKey = VBA.vbNullString
    If m_CfgTableParser Is Nothing Then Exit Function
    If m_CfgTableParser.CfgParserBase Is Nothing Then Exit Function

    lookupKey = VBA.Trim$(lookupKey)
    If VBA.Len(lookupKey) = 0 Then Exit Function

    lookupPrefix = private_BuildLookupPrefix(lookupKey)
    If VBA.Len(lookupPrefix) = 0 Then Exit Function

    targetColumnRef = m_CfgTableParser.CfgParserBase.GetOptionalConfigValue(cfgMap, lookupPrefix & "TargetColumn")
    targetColumnRef = VBA.Trim$(targetColumnRef)
    If VBA.Len(targetColumnRef) = 0 Then
        private_ShowConfigError "Missing or invalid config key: " & lookupPrefix & "TargetColumn"
        Exit Function
    End If
    If Not private_TryParseEntityLookupTableColumnRef(targetColumnRef, outTargetColumnKey) Then
        private_ShowConfigError "Invalid TargetColumn reference '" & targetColumnRef & "' for lookup '" & lookupKey & _
            "'. Expected EntityLookup.Table.Column[ColumnAlias]."
        Exit Function
    End If
    If Not private_IsDeclaredEntityLookupTableColumn(cfgMap, outTargetColumnKey) Then
        private_ShowConfigError "TargetColumn '" & targetColumnRef & "' references a column not declared in EntityLookup.Table.Columns."
        Exit Function
    End If

    private_TryGetLookupTargetColumn = True
End Function

Private Function private_BuildQualifiedResultColumnMappings( _
    ByVal cfgMap As Object, _
    ByVal lookupPrefix As String, _
    ByVal expectedSourceAlias As String, _
    ByVal expectedTableAlias As String, _
    ByVal lookupKey As String _
) As Collection
    Dim result As Collection
    Dim seen As Object
    Dim cfgKeyObj As Variant
    Dim cfgKey As String
    Dim mappingsPrefix As String
    Dim sourceRef As String
    Dim targetRef As String
    Dim sourceAlias As String
    Dim tableAlias As String
    Dim sourceColumnAlias As String
    Dim targetColumnAlias As String
    Dim mappingToken As String

    Set result = New Collection
    Set seen = ex_Helpers.fn_CreateDictionaryTextCompare()
    If cfgMap Is Nothing Then Exit Function

    mappingsPrefix = VBA.LCase$(lookupPrefix & "ResultColumnsAliases[")
    For Each cfgKeyObj In cfgMap.Keys
        cfgKey = VBA.CStr(cfgKeyObj)
        If VBA.Left$(VBA.LCase$(cfgKey), VBA.Len(mappingsPrefix)) <> mappingsPrefix Then GoTo ContinueMapping
        If VBA.Right$(cfgKey, 1) <> "]" Then GoTo ContinueMapping

        sourceRef = VBA.Mid$(cfgKey, VBA.Len(mappingsPrefix) + 1, VBA.Len(cfgKey) - VBA.Len(mappingsPrefix) - 1)
        targetRef = VBA.Trim$(VBA.CStr(cfgMap(cfgKeyObj)))
        If Not m_CfgTableParser.TryParseColumnRefToken(sourceRef, sourceAlias, tableAlias, sourceColumnAlias) Then
            private_ShowConfigError "Invalid source column reference in ResultColumnsAliases for lookup '" & lookupKey & "': " & sourceRef
            Exit Function
        End If
        If VBA.StrComp(sourceAlias, expectedSourceAlias, VBA.vbTextCompare) <> 0 Or _
           VBA.StrComp(tableAlias, expectedTableAlias, VBA.vbTextCompare) <> 0 Then
            private_ShowConfigError "ResultColumnsAliases source '" & sourceRef & "' does not belong to lookup source " & _
                expectedSourceAlias & ".Sheet[" & expectedTableAlias & "]."
            Exit Function
        End If
        If Not private_TryParseEntityLookupTableColumnRef(targetRef, targetColumnAlias) Then
            private_ShowConfigError "Invalid target column reference in ResultColumnsAliases for lookup '" & lookupKey & _
                "': " & targetRef & ". Expected EntityLookup.Table.Column[ColumnAlias]."
            Exit Function
        End If
        If Not private_IsDeclaredEntityLookupTableColumn(cfgMap, targetColumnAlias) Then
            private_ShowConfigError "ResultColumnsAliases target '" & targetRef & _
                "' is not declared in EntityLookup.Table.Columns."
            Exit Function
        End If

        mappingToken = sourceColumnAlias & "=>" & targetColumnAlias
        If seen.Exists(VBA.LCase$(sourceColumnAlias)) Then
            private_ShowConfigError "Duplicate source column mapping '" & sourceColumnAlias & "' for lookup '" & lookupKey & "'."
            Exit Function
        End If
        seen(VBA.LCase$(sourceColumnAlias)) = True
        result.Add mappingToken
ContinueMapping:
    Next cfgKeyObj

    Set private_BuildQualifiedResultColumnMappings = result
End Function

Private Function private_TryParseEntityLookupTableColumnRef( _
    ByVal referenceText As String, _
    ByRef outColumnAlias As String _
) As Boolean
    Const REF_PREFIX As String = "EntityLookup.Table.Column["
    Dim suffix As String

    outColumnAlias = VBA.vbNullString
    referenceText = VBA.Trim$(referenceText)
    If VBA.Len(referenceText) <= VBA.Len(REF_PREFIX) + 1 Then Exit Function
    If VBA.StrComp(VBA.Left$(referenceText, VBA.Len(REF_PREFIX)), REF_PREFIX, VBA.vbTextCompare) <> 0 Then Exit Function
    If VBA.Right$(referenceText, 1) <> "]" Then Exit Function

    suffix = VBA.Mid$(referenceText, VBA.Len(REF_PREFIX) + 1)
    outColumnAlias = VBA.Trim$(VBA.Left$(suffix, VBA.Len(suffix) - 1))
    If VBA.Len(outColumnAlias) = 0 Then Exit Function
    If VBA.InStr(1, outColumnAlias, "[", VBA.vbBinaryCompare) > 0 Then Exit Function
    If VBA.InStr(1, outColumnAlias, "]", VBA.vbBinaryCompare) > 0 Then Exit Function

    private_TryParseEntityLookupTableColumnRef = True
End Function

Private Function private_IsDeclaredEntityLookupTableColumn( _
    ByVal cfgMap As Object, _
    ByVal columnAlias As String _
) As Boolean
    Dim columnKeys As Collection
    Dim columnKeyObj As Variant

    columnAlias = VBA.Trim$(columnAlias)
    If VBA.Len(columnAlias) = 0 Then Exit Function
    If Not private_TryGetTableColumnKeysFromMap(cfgMap, columnKeys) Then Exit Function

    For Each columnKeyObj In columnKeys
        If VBA.StrComp(VBA.Trim$(VBA.CStr(columnKeyObj)), columnAlias, VBA.vbTextCompare) = 0 Then
            private_IsDeclaredEntityLookupTableColumn = True
            Exit Function
        End If
    Next columnKeyObj
End Function

Private Sub private_ShowConfigError(ByVal messageText As String)
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "EntityLookupCfgParser: " & VBA.CStr(messageText)
#End If
    VBA.MsgBox "PrototypeNew: " & VBA.CStr(messageText), vbExclamation, RUNTIME_ERROR_TITLE
End Sub
