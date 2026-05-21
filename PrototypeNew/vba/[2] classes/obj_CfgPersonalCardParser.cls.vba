VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_CfgPersonalCardParser"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private m_CfgTableParser As obj_CfgTableParser
Private m_IsDisposed As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
    Set m_CfgTableParser = New obj_CfgTableParser
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
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
    If VBA.Len(sourceAlias) = 0 Then Exit Function
    If VBA.Len(tableAlias) = 0 Then Exit Function

    ' 1) Читаем конфиг-таблицу и нормализуем ее в словарь ключ/значение.
    If Not m_CfgTableParser.CfgParserBase.TryGetConfigEntries(configEntries) Then Exit Function
    If Not m_CfgTableParser.CfgParserBase.BuildConfigDictionary(configEntries, cfgMap) Then Exit Function
    If cfgMap Is Nothing Then Exit Function

    ' 2) Table-specific значения берем через obj_CfgTableParser.
    If m_CfgTableParser Is Nothing Then Exit Function
    tablePathPrefix = m_CfgTableParser.BuildTablePathPrefix(sourceAlias, tableAlias)
    If VBA.Len(tablePathPrefix) = 0 Then Exit Function

    If Not m_CfgTableParser.TryResolveSourcePath(cfgMap, sourceAlias, sourcePath) Then Exit Function
    If Not m_CfgTableParser.TryResolveSheetName(cfgMap, tablePathPrefix, sheetName) Then Exit Function
    If Not m_CfgTableParser.TryResolveRangeMarkers(cfgMap, tablePathPrefix, rangeStartMarker, rangeEndMarker) Then Exit Function
    If Not m_CfgTableParser.TryGetRequiredColumnHeadersAliases(cfgMap, tablePathPrefix, columnHeadersAliases) Then Exit Function

    ' 3) Здесь именно PersonalCard-парсер собирает obj_SqlParams.
    Set sqlParams = New obj_SqlParams
    sqlParams.SourcePath = sourcePath
    sqlParams.SheetName = sheetName
    sqlParams.RangeStartMarker = rangeStartMarker
    sqlParams.RangeEndMarker = rangeEndMarker

    ' 4) Формируем WHERE из конфига:
    ' <TablePathPrefix>Key -> алиас поля (например FIO)
    ' CommonKey            -> значение для фильтра
    If Not m_CfgTableParser.CfgParserBase.TryGetRequiredConfigValue(cfgMap, tablePathPrefix & "Key", keyColumnAlias) Then Exit Function
    If Not m_CfgTableParser.CfgParserBase.TryGetRequiredConfigValue(cfgMap, "CommonKey", commonKeyValue) Then Exit Function
    If Not m_CfgTableParser.TryResolveMapByColumnAlias(cfgMap, tablePathPrefix, keyColumnAlias, keySourceColumnHeader, keyMappedColumnHeader) Then Exit Function
    sqlParams.WhereConditions = ex_HelpersSql.fn_BuildWhereEqualsSql(keySourceColumnHeader, commonKeyValue)
    If VBA.Len(sqlParams.WhereConditions) = 0 Then Exit Function

    For Each columnAliasObj In columnHeadersAliases
        If Not m_CfgTableParser.TryResolveMapByColumnAlias( _
            cfgMap, _
            tablePathPrefix, _
            VBA.CStr(columnAliasObj), _
            sourceColumnHeader, _
            mappedColumnHeader _
        ) Then Exit Function
        If Not sqlParams.AddColumnMapping(sourceColumnHeader, mappedColumnHeader) Then Exit Function
    Next columnAliasObj

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

    If m_CfgTableParser Is Nothing Then Exit Function
    If m_CfgTableParser.CfgParserBase Is Nothing Then Exit Function

    ' Читаем конфиг один раз и автоматически собираем все таблицы,
    ' объявленные через Source.*.SheetAliases.
    If Not m_CfgTableParser.CfgParserBase.TryGetConfigEntries(configEntries) Then Exit Function
    If Not m_CfgTableParser.CfgParserBase.BuildConfigDictionary(configEntries, cfgMap) Then Exit Function
    If cfgMap Is Nothing Then Exit Function

    Set seenRefs = CreateObject("Scripting.Dictionary")
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
            If Not TryBuildSqlParams(sourceAlias, tableAlias, sqlParams) Then Exit Function
            If Not sqlParams Is Nothing Then outSqlParamsList.Add sqlParams
ContinueTableAlias:
        Next tableAliasObj
ContinueCfgKey:
    Next cfgKeyObj

    If outSqlParamsList.Count <= 0 Then Exit Function
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
    Const SUFFIX As String = ".sheetaliases"
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
