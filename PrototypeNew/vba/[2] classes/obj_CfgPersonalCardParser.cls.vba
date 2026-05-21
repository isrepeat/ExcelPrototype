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
