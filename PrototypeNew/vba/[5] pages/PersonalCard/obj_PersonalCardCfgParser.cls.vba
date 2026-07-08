VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PersonalCardCfgParser"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private m_CfgTableParser As obj_CfgTableParser
Private m_IsDisposed As Boolean

Private Const ADDITIONAL_SQL_PARAMS_KEY As String = "AdditionalSqlParams"
Private Const ADDITIONAL_SQL_PARAM_ROW_PROCESSOR As String = "rowprocessor"

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

    If Not m_CfgTableParser.Initialize(configTable, Me) Then Exit Function

    Initialize = True
End Function

Public Sub Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next

    If Not m_CfgTableParser Is Nothing Then
        m_CfgTableParser.Dispose
    End If

    Set m_CfgTableParser = Nothing
    On Error GoTo 0
End Sub

Public Function ResolveLatestByDmyPattern(ByVal rawValue As String) As String
    ResolveLatestByDmyPattern = ex_SourceResolver.fn_ResolveLatestByDmyPattern(rawValue)
End Function

Public Function TryBuildSqlParams( _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByRef outSqlParams As obj_SqlParams _
) As Boolean
    Dim configEntries As Collection
    Dim cfgMap As Object
    Dim tablePathPrefix As String
    Dim sourcePath As String
    Dim sheetName As String
    Dim rangeStartMarker As String
    Dim rangeEndMarker As String
    Dim columnHeadersAliases As Collection
    Dim columnAliasObj As Variant
    Dim sourceColumnHeader As String
    Dim mappedColumnHeader As String
    Dim keyColumnAlias As String
    Dim keySourceColumnHeader As String
    Dim keyMappedColumnHeader As String
    Dim commonKeyValue As String
    Dim sqlParams As obj_SqlParams

    ' Сбрасываем выходные значения.
    Set outSqlParams = Nothing

    sourceAlias = VBA.Trim$(sourceAlias)
    tableAlias = VBA.Trim$(tableAlias)
    If VBA.Len(sourceAlias) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildSqlParams: sourceAlias is empty."
#End If
        Exit Function
    End If
    If VBA.Len(tableAlias) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildSqlParams: tableAlias is empty."
#End If
        Exit Function
    End If

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "PersonalCardCfgParser.TryBuildSqlParams: sourceAlias='" & sourceAlias & "' tableAlias='" & tableAlias & "'"
#End If

    ' 1) Читаем конфиг-таблицу и нормализуем ее в словарь ключ/значение.
    If Not m_CfgTableParser.CfgParserBase.TryGetConfigEntries(configEntries) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildSqlParams: failed to get config entries."
#End If
        Exit Function
    End If
    If Not m_CfgTableParser.CfgParserBase.BuildConfigDictionary(configEntries, cfgMap) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildSqlParams: failed to build config dictionary."
#End If
        Exit Function
    End If
    If cfgMap Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildSqlParams: config dictionary is Nothing."
#End If
        Exit Function
    End If

    ' 2) Table-specific значения берем через obj_CfgTableParser.
    If m_CfgTableParser Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildSqlParams: CfgTableParser is Nothing."
#End If
        Exit Function
    End If
    tablePathPrefix = m_CfgTableParser.BuildTablePathPrefix(sourceAlias, tableAlias)
    If VBA.Len(tablePathPrefix) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildSqlParams: table path prefix is empty."
#End If
        Exit Function
    End If

    If Not m_CfgTableParser.TryResolveSourcePath(cfgMap, sourceAlias, sourcePath) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildSqlParams: failed to resolve source path for sourceAlias='" & sourceAlias & "'."
#End If
        Exit Function
    End If
    If Not m_CfgTableParser.TryResolveSheetName(cfgMap, tablePathPrefix, sheetName) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildSqlParams: failed to resolve sheet name for tablePathPrefix='" & tablePathPrefix & "'."
#End If
        Exit Function
    End If
    If Not m_CfgTableParser.TryResolveRangeMarkers(cfgMap, tablePathPrefix, rangeStartMarker, rangeEndMarker) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildSqlParams: failed to resolve range markers for tablePathPrefix='" & tablePathPrefix & "'."
#End If
        Exit Function
    End If
    If Not m_CfgTableParser.TryGetRequiredColumnHeadersAliases(cfgMap, tablePathPrefix, columnHeadersAliases) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildSqlParams: failed to read ColumnHeadersAliases for tablePathPrefix='" & tablePathPrefix & "'."
#End If
        Exit Function
    End If

    ' 3) Здесь именно PersonalCard-парсер собирает obj_SqlParams.
    Set sqlParams = New obj_SqlParams
    sqlParams.SourcePath = sourcePath
    sqlParams.SheetName = sheetName
    sqlParams.RangeStartMarker = rangeStartMarker
    sqlParams.RangeEndMarker = rangeEndMarker

    ' 4) Формируем WHERE из конфига:
    ' <TablePathPrefix>Key -> алиас поля (например FIO)
    ' CommonKey            -> значение для фильтра
    If Not m_CfgTableParser.CfgParserBase.TryGetRequiredConfigValue(cfgMap, tablePathPrefix & "Key", keyColumnAlias) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildSqlParams: required key is missing '" & tablePathPrefix & "Key'."
#End If
        Exit Function
    End If
    If Not m_CfgTableParser.CfgParserBase.TryGetRequiredConfigValue(cfgMap, "CommonKey", commonKeyValue) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildSqlParams: CommonKey is missing."
#End If
        Exit Function
    End If
    If Not m_CfgTableParser.TryResolveMapByColumnAlias(cfgMap, tablePathPrefix, keyColumnAlias, keySourceColumnHeader, keyMappedColumnHeader) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildSqlParams: failed to resolve key column alias='" & keyColumnAlias & "' for tablePathPrefix='" & tablePathPrefix & "'."
#End If
        Exit Function
    End If
    sqlParams.WhereConditions = ex_HelpersSql.fn_BuildWhereEqualsSql(keySourceColumnHeader, commonKeyValue)
    If VBA.Len(sqlParams.WhereConditions) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildSqlParams: failed to build WHERE condition keySourceColumnHeader='" & keySourceColumnHeader & "' commonKey='" & commonKeyValue & "'."
#End If
        Exit Function
    End If

    For Each columnAliasObj In columnHeadersAliases
        If Not m_CfgTableParser.TryResolveMapByColumnAlias( _
            cfgMap, _
            tablePathPrefix, _
            VBA.CStr(columnAliasObj), _
            sourceColumnHeader, _
            mappedColumnHeader _
        ) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildSqlParams: failed to resolve column alias='" & VBA.CStr(columnAliasObj) & "' for tablePathPrefix='" & tablePathPrefix & "'."
#End If
            Exit Function
        End If
        If Not sqlParams.AddColumnMapping(sourceColumnHeader, mappedColumnHeader, VBA.CStr(columnAliasObj)) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildSqlParams: failed to add column mapping alias='" & VBA.CStr(columnAliasObj) & "' source='" & sourceColumnHeader & "' mapped='" & mappedColumnHeader & "'."
#End If
            Exit Function
        End If
    Next columnAliasObj

    If Not private_TryAttachAdditionalSqlParams(cfgMap, tablePathPrefix, sqlParams) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildSqlParams: failed to attach AdditionalSqlParams for tablePathPrefix='" & tablePathPrefix & "'."
#End If
        Exit Function
    End If

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "PersonalCardCfgParser.TryBuildSqlParams: done " & sqlParams.fn_ToString()
#End If
    Set outSqlParams = sqlParams
    TryBuildSqlParams = True
End Function

Public Function TryBuildAllSqlParams(ByRef outSqlParamsList As Collection) As Boolean
    Dim configEntries As Collection
    Dim cfgMap As Object
    Dim cfgKeyObj As Variant
    Dim cfgKey As String
    Dim sourceAlias As String
    Dim sheetAliases As Collection
    Dim tableAliasObj As Variant
    Dim tableAlias As String
    Dim tableRefKey As String
    Dim seenRefs As Object
    Dim sqlParams As obj_SqlParams

    Set outSqlParamsList = Nothing

    If m_CfgTableParser Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildAllSqlParams: CfgTableParser is Nothing."
#End If
        Exit Function
    End If
    If m_CfgTableParser.CfgParserBase Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildAllSqlParams: CfgParserBase is Nothing."
#End If
        Exit Function
    End If

    ' Читаем конфиг один раз и автоматически собираем все таблицы,
    ' объявленные через Source.*.SheetsAliases.
    If Not m_CfgTableParser.CfgParserBase.TryGetConfigEntries(configEntries) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildAllSqlParams: failed to get config entries."
#End If
        Exit Function
    End If
    If Not m_CfgTableParser.CfgParserBase.BuildConfigDictionary(configEntries, cfgMap) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildAllSqlParams: failed to build config dictionary."
#End If
        Exit Function
    End If
    If cfgMap Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildAllSqlParams: config dictionary is Nothing."
#End If
        Exit Function
    End If

    Set seenRefs = VBA.CreateObject("Scripting.Dictionary")
    seenRefs.CompareMode = 1

    Set outSqlParamsList = New Collection

    For Each cfgKeyObj In cfgMap.Keys
        cfgKey = VBA.LCase$(VBA.Trim$(VBA.CStr(cfgKeyObj)))
        sourceAlias = VBA.vbNullString
        If Not private_TryParseSourceSheetAliasesKey(cfgKey, sourceAlias) Then GoTo ContinueCfgKey

        Set sheetAliases = m_CfgTableParser.CfgParserBase.SplitListToCollection(VBA.CStr(cfgMap(cfgKeyObj)))
        If sheetAliases Is Nothing Then GoTo ContinueCfgKey

        For Each tableAliasObj In sheetAliases
            tableAlias = VBA.Trim$(VBA.CStr(tableAliasObj))
            If VBA.Len(tableAlias) = 0 Then GoTo ContinueTableAlias

            tableRefKey = sourceAlias & ".sheet[" & VBA.LCase$(tableAlias) & "]"
            If seenRefs.Exists(tableRefKey) Then GoTo ContinueTableAlias
            seenRefs(tableRefKey) = True

            Set sqlParams = Nothing
            If Not TryBuildSqlParams(sourceAlias, tableAlias, sqlParams) Then
#If LOGGING_DEBUG_ENABLED Then
                ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildAllSqlParams: TryBuildSqlParams failed sourceAlias='" & sourceAlias & "' tableAlias='" & tableAlias & "'."
#End If
                Exit Function
            End If
            If Not sqlParams Is Nothing Then outSqlParamsList.Add sqlParams
ContinueTableAlias:
        Next tableAliasObj
ContinueCfgKey:
    Next cfgKeyObj

    If outSqlParamsList.Count <= 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser.TryBuildAllSqlParams: no SQL params were built."
#End If
        Exit Function
    End If
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "PersonalCardCfgParser.TryBuildAllSqlParams: built count=" & VBA.CStr(outSqlParamsList.Count)
#End If
    TryBuildAllSqlParams = True
End Function

' //
' // Internal
' //
Private Function private_TryParseSourceSheetAliasesKey( _
    ByVal normalizedCfgKey As String, _
    ByRef outSourceAlias As String _
) As Boolean
    Const PREFIX As String = "source."
    Const SUFFIX As String = ".sheetsaliases"
    Dim keyLen As Long
    Dim sourceLen As Long

    outSourceAlias = VBA.vbNullString

    normalizedCfgKey = VBA.LCase$(VBA.Trim$(normalizedCfgKey))
    keyLen = VBA.Len(normalizedCfgKey)
    If keyLen <= VBA.Len(PREFIX) + VBA.Len(SUFFIX) Then Exit Function

    If VBA.Left$(normalizedCfgKey, VBA.Len(PREFIX)) <> PREFIX Then Exit Function
    If VBA.Right$(normalizedCfgKey, VBA.Len(SUFFIX)) <> SUFFIX Then Exit Function

    sourceLen = keyLen - VBA.Len(PREFIX) - VBA.Len(SUFFIX)
    outSourceAlias = VBA.Mid$(normalizedCfgKey, VBA.Len(PREFIX) + 1, sourceLen)
    outSourceAlias = VBA.Trim$(outSourceAlias)
    If VBA.Len(outSourceAlias) = 0 Then Exit Function

    private_TryParseSourceSheetAliasesKey = True
End Function

Private Function private_TryAttachAdditionalSqlParams( _
    ByVal cfgMap As Object, _
    ByVal tablePathPrefix As String, _
    ByVal sqlParams As obj_SqlParams _
) As Boolean
    Dim cfgBase As obj_CfgParserBase
    Dim rawAdditionalParams As String
    Dim additionalParamsByKey As Object
    Dim paramKeyObj As Variant
    Dim rowProcessorClassName As String
    Dim rowProcessor As obj_ISqlRowProcessor

    If sqlParams Is Nothing Then Exit Function
    If m_CfgTableParser Is Nothing Then Exit Function
    Set cfgBase = m_CfgTableParser.CfgParserBase
    If cfgBase Is Nothing Then Exit Function

    rawAdditionalParams = cfgBase.GetOptionalConfigValue(cfgMap, tablePathPrefix & ADDITIONAL_SQL_PARAMS_KEY)

    If VBA.Len(VBA.Trim$(rawAdditionalParams)) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "PersonalCardCfgParser: no AdditionalSqlParams for '" & tablePathPrefix & "'."
#End If
        Set sqlParams.RowProcessor = Nothing
        private_TryAttachAdditionalSqlParams = True
        Exit Function
    End If

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "PersonalCardCfgParser: AdditionalSqlParams for '" & tablePathPrefix & "' raw='" & rawAdditionalParams & "'"
#End If
    If Not private_TryParseAdditionalSqlParams(rawAdditionalParams, additionalParamsByKey) Then Exit Function
    If additionalParamsByKey Is Nothing Then Exit Function

    For Each paramKeyObj In additionalParamsByKey.Keys
        Select Case VBA.LCase$(VBA.CStr(paramKeyObj))
            Case ADDITIONAL_SQL_PARAM_ROW_PROCESSOR
                ' supported
            Case Else
#If LOGGING_DEBUG_ENABLED Then
                ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser: unsupported additional SQL param '" & VBA.CStr(paramKeyObj) & "' for '" & tablePathPrefix & "'."
#End If
                Exit Function
        End Select
    Next paramKeyObj

    If additionalParamsByKey.Exists(ADDITIONAL_SQL_PARAM_ROW_PROCESSOR) Then
        rowProcessorClassName = VBA.Trim$(VBA.CStr(additionalParamsByKey(ADDITIONAL_SQL_PARAM_ROW_PROCESSOR)))
        If VBA.Len(rowProcessorClassName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser: RowProcessor value is empty for '" & tablePathPrefix & "'."
#End If
            Exit Function
        End If

        If Not ex_SqlRowProcessorFactory.fn_TryCreateByClassName(rowProcessorClassName, rowProcessor) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser: failed to create RowProcessor class='" & rowProcessorClassName & "' for '" & tablePathPrefix & "'."
#End If
            Exit Function
        End If
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "PersonalCardCfgParser: RowProcessor attached class='" & rowProcessorClassName & "' type='" & TypeName(rowProcessor) & "'"
#End If
        Set sqlParams.RowProcessor = rowProcessor
    Else
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "PersonalCardCfgParser: AdditionalSqlParams has no RowProcessor for '" & tablePathPrefix & "'."
#End If
        Set sqlParams.RowProcessor = Nothing
    End If

    private_TryAttachAdditionalSqlParams = True
End Function

Private Function private_TryParseAdditionalSqlParams( _
    ByVal rawAdditionalParams As String, _
    ByRef outParamsByKey As Object _
) As Boolean
    Dim rawParts As Variant
    Dim rawPart As Variant
    Dim tokenText As String
    Dim eqPos As Long
    Dim paramKey As String
    Dim paramValue As String

    Set outParamsByKey = ex_Helpers.fn_CreateDictionaryTextCompare()
    If outParamsByKey Is Nothing Then Exit Function

    rawAdditionalParams = VBA.Trim$(rawAdditionalParams)
    If VBA.Len(rawAdditionalParams) = 0 Then
        private_TryParseAdditionalSqlParams = True
        Exit Function
    End If

    rawParts = VBA.Split(rawAdditionalParams, ";")
    For Each rawPart In rawParts
        tokenText = VBA.Trim$(VBA.CStr(rawPart))
        If VBA.Len(tokenText) = 0 Then GoTo ContinueToken

        eqPos = VBA.InStr(1, tokenText, "=", VBA.vbBinaryCompare)
        If eqPos <= 1 Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser: invalid additional SQL param token '" & tokenText & "'. Expected format Key=Value."
#End If
            Exit Function
        End If

        paramKey = VBA.LCase$(VBA.Trim$(VBA.Left$(tokenText, eqPos - 1)))
        paramValue = VBA.Trim$(VBA.Mid$(tokenText, eqPos + 1))
        If VBA.Len(paramKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PersonalCardCfgParser: additional SQL param key is empty in token '" & tokenText & "'."
#End If
            Exit Function
        End If

        outParamsByKey(paramKey) = paramValue
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "PersonalCardCfgParser: parsed AdditionalSqlParam key='" & paramKey & "' value='" & paramValue & "'"
#End If
ContinueToken:
    Next rawPart

    private_TryParseAdditionalSqlParams = True
End Function
