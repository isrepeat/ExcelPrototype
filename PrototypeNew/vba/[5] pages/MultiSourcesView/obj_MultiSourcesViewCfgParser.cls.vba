VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_MultiSourcesViewCfgParser"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const RUNTIME_ERROR_TITLE As String = "PrototypeNew / MultiSourcesView runtime"

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

Public Function ResolveAllByDmyPattern( _
    ByVal rawValue As String, _
    Optional ByVal resolverArgs As String = "" _
) As Collection
    Set ResolveAllByDmyPattern = ex_SourceResolver.fn_ResolveAllByDmyPattern(rawValue, resolverArgs)
End Function

Public Function TryGetScenarioClassName( _
    ByRef outClassName As String, _
    ByRef outHasScenario As Boolean _
) As Boolean
    Dim cfgMap As Object

    outClassName = VBA.vbNullString
    outHasScenario = False
    If Not private_TryBuildConfigMap(cfgMap) Then Exit Function
    If cfgMap.Exists("MultiSourcesView.ScenarioClass") Then
        outClassName = VBA.Trim$(VBA.CStr( _
            cfgMap("MultiSourcesView.ScenarioClass")))
        outHasScenario = True
        If VBA.Len(outClassName) = 0 Then
            private_ShowConfigError _
                "MultiSourcesView.ScenarioClass is specified but empty."
            Exit Function
        End If
    End If
    TryGetScenarioClassName = True
End Function

Public Function TryGetViewSettings( _
    ByRef outTableRefs As Collection, _
    ByRef outColumns As Collection, _
    ByRef outMaxRows As Long _
) As Boolean
    Dim cfgMap As Object
    Dim rawTableRefs As String
    Dim rawColumns As String
    Dim maxRowsText As String
    Dim itemValue As Variant
    Dim sourceAlias As String
    Dim tableAlias As String

    Set outTableRefs = Nothing
    Set outColumns = Nothing
    outMaxRows = 500

    If Not private_TryBuildConfigMap(cfgMap) Then Exit Function
    If Not private_TryGetRequiredConfigValue(cfgMap, "MultiSourcesView.Tables", rawTableRefs) Then Exit Function
    If Not private_TryGetRequiredConfigValue(cfgMap, "MultiSourcesView.Columns", rawColumns) Then Exit Function

    Set outTableRefs = m_CfgTableParser.CfgParserBase.SplitListToCollection(rawTableRefs)
    Set outColumns = m_CfgTableParser.CfgParserBase.SplitListToCollection(rawColumns)
    If outTableRefs Is Nothing Or outTableRefs.Count = 0 Then
        private_ShowConfigError "MultiSourcesView.Tables is empty."
        Exit Function
    End If
    If outColumns Is Nothing Or outColumns.Count = 0 Then
        private_ShowConfigError "MultiSourcesView.Columns is empty."
        Exit Function
    End If

    For Each itemValue In outTableRefs
        If Not m_CfgTableParser.TryParseTableRefToken(VBA.Trim$(VBA.CStr(itemValue)), sourceAlias, tableAlias) Then
            private_ShowConfigError "Invalid table reference '" & VBA.CStr(itemValue) & "'. Expected <SourceAlias>.Sheet[<TableAlias>]."
            Exit Function
        End If
    Next itemValue

    maxRowsText = m_CfgTableParser.CfgParserBase.GetOptionalConfigValue(cfgMap, "MultiSourcesView.MaxRows", "500")
    If VBA.IsNumeric(maxRowsText) Then outMaxRows = VBA.CLng(maxRowsText)
    If outMaxRows < 1 Then outMaxRows = 500
    TryGetViewSettings = True
End Function

Public Function TryValidateSkeleton(ByRef outStatusText As String) As Boolean
    Dim cfgMap As Object
    Dim leftRef As String
    Dim rightRef As String
    Dim keyColumns As String
    Dim compareColumns As String

    outStatusText = VBA.vbNullString
    If Not private_TryBuildConfigMap(cfgMap) Then Exit Function

    If Not private_TryGetRequiredConfigValue(cfgMap, "Comparing.LeftTable", leftRef) Then Exit Function
    If Not private_TryGetRequiredConfigValue(cfgMap, "Comparing.RightTable", rightRef) Then Exit Function
    If Not private_TryGetRequiredConfigValue(cfgMap, "Comparing.KeyColumns", keyColumns) Then Exit Function
    If Not private_TryGetRequiredConfigValue(cfgMap, "Comparing.CompareColumns", compareColumns) Then Exit Function

    If Not private_ValidateTableRef(leftRef, "Comparing.LeftTable") Then Exit Function
    If Not private_ValidateTableRef(rightRef, "Comparing.RightTable") Then Exit Function
    If Not private_ValidateColumnList(keyColumns, "Comparing.KeyColumns") Then Exit Function
    If Not private_ValidateColumnList(compareColumns, "Comparing.CompareColumns") Then Exit Function

    outStatusText = _
        "Config ready: " & leftRef & " -> " & rightRef & _
        "; keys: " & keyColumns & _
        "; compare: " & compareColumns
    TryValidateSkeleton = True
End Function

Public Function TryGetCompareSettings( _
    ByRef outLeftTableRef As String, _
    ByRef outRightTableRef As String, _
    ByRef outKeyColumns As Collection, _
    ByRef outCompareColumns As Collection, _
    ByRef outCompareColumnFormats As Object, _
    ByRef outIgnoreCase As Boolean, _
    ByRef outTrimText As Boolean _
) As Boolean
    Dim cfgMap As Object
    Dim rawKeyColumns As String
    Dim rawCompareColumns As String

    outLeftTableRef = VBA.vbNullString
    outRightTableRef = VBA.vbNullString
    Set outKeyColumns = Nothing
    Set outCompareColumns = Nothing
    Set outCompareColumnFormats = Nothing
    outIgnoreCase = True
    outTrimText = True

    If Not private_TryBuildConfigMap(cfgMap) Then Exit Function

    If Not private_TryGetRequiredConfigValue(cfgMap, "Comparing.LeftTable", outLeftTableRef) Then Exit Function
    If Not private_TryGetRequiredConfigValue(cfgMap, "Comparing.RightTable", outRightTableRef) Then Exit Function
    If Not private_TryGetRequiredConfigValue(cfgMap, "Comparing.KeyColumns", rawKeyColumns) Then Exit Function
    If Not private_TryGetRequiredConfigValue(cfgMap, "Comparing.CompareColumns", rawCompareColumns) Then Exit Function

    Set outKeyColumns = m_CfgTableParser.CfgParserBase.SplitListToCollection(rawKeyColumns)
    If Not private_TryParseCompareColumnSpecs(rawCompareColumns, outCompareColumns, outCompareColumnFormats) Then Exit Function
    If outKeyColumns Is Nothing Then
        private_ShowConfigError "Comparing.KeyColumns is empty."
        Exit Function
    End If
    If outKeyColumns.Count <= 0 Then
        private_ShowConfigError "Comparing.KeyColumns is empty."
        Exit Function
    End If
    If outCompareColumns Is Nothing Then
        private_ShowConfigError "Comparing.CompareColumns is empty."
        Exit Function
    End If
    If outCompareColumns.Count <= 0 Then
        private_ShowConfigError "Comparing.CompareColumns is empty."
        Exit Function
    End If

    outIgnoreCase = private_ParseBoolean(m_CfgTableParser.CfgParserBase.GetOptionalConfigValue(cfgMap, "Comparing.IgnoreCase", "true"), True)
    outTrimText = private_ParseBoolean(m_CfgTableParser.CfgParserBase.GetOptionalConfigValue(cfgMap, "Comparing.TrimText", "true"), True)
    TryGetCompareSettings = True
End Function

Public Function TryBuildTableSqlParams( _
    ByVal tableRef As String, _
    ByVal outputColumnAliases As Collection, _
    ByRef outSqlParams As obj_SqlParams _
) As Boolean
    Dim cfgMap As Object
    Dim sourceAlias As String
    Dim tableAlias As String
    Dim tablePathPrefix As String
    Dim sourcePath As String
    Dim sheetName As String
    Dim rangeStartMarker As String
    Dim rangeEndMarker As String
    Dim aliasItem As Variant
    Dim columnAlias As String
    Dim sourceColumnHeader As String
    Dim mappedColumnHeader As String
    Dim sqlParams As obj_SqlParams

    Set outSqlParams = Nothing
    If outputColumnAliases Is Nothing Then
        private_ShowConfigError "Comparing output columns are not specified."
        Exit Function
    End If
    If outputColumnAliases.Count <= 0 Then
        private_ShowConfigError "Comparing output columns are not specified."
        Exit Function
    End If
    If Not private_TryBuildConfigMap(cfgMap) Then Exit Function

    If Not m_CfgTableParser.TryParseTableRefToken(tableRef, sourceAlias, tableAlias) Then
        private_ShowConfigError "Invalid table reference '" & tableRef & "'. Expected <SourceAlias>.Sheet[<TableAlias>]."
        Exit Function
    End If

    tablePathPrefix = m_CfgTableParser.BuildTablePathPrefix(sourceAlias, tableAlias)
    If VBA.Len(tablePathPrefix) = 0 Then
        private_ShowConfigError "Failed to build table path prefix for '" & tableRef & "'."
        Exit Function
    End If
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

    Set sqlParams = New obj_SqlParams
    sqlParams.SourcePath = sourcePath
    sqlParams.SheetName = sheetName
    sqlParams.RangeStartMarker = rangeStartMarker
    sqlParams.RangeEndMarker = rangeEndMarker

    For Each aliasItem In outputColumnAliases
        columnAlias = VBA.Trim$(VBA.CStr(aliasItem))
        If VBA.Len(columnAlias) = 0 Then GoTo ContinueAlias
        If Not m_CfgTableParser.TryResolveMapByColumnAlias(cfgMap, tablePathPrefix, columnAlias, sourceColumnHeader, mappedColumnHeader) Then
            private_ShowConfigError "Failed to resolve column alias '" & columnAlias & "' for table prefix '" & tablePathPrefix & "'."
            Exit Function
        End If
        If Not sqlParams.AddColumnMapping(sourceColumnHeader, mappedColumnHeader, columnAlias) Then
            private_ShowConfigError "Failed to add SQL column mapping for alias '" & columnAlias & "'."
            Exit Function
        End If
ContinueAlias:
    Next aliasItem

    Set outSqlParams = sqlParams
    TryBuildTableSqlParams = True
End Function

Public Function TryBuildTableSqlParamsList( _
    ByVal tableRef As String, _
    ByVal outputColumnAliases As Collection, _
    ByRef outSqlParamsList As Collection _
) As Boolean
    Dim cfgMap As Object
    Dim sourceAlias As String
    Dim tableAlias As String
    Dim tablePathPrefix As String
    Dim sourceKeyPrefix As String
    Dim sourcePathKey As String
    Dim resolverKey As String
    Dim resolverArgsKey As String
    Dim runtimeAliasPatternKey As String
    Dim rawSourcePath As String
    Dim sourcePathPattern As String
    Dim resolverName As String
    Dim resolverArgs As String
    Dim runtimeAliasPattern As String
    Dim sourcePaths As Collection
    Dim sourcePath As Variant
    Dim sqlParams As obj_SqlParams

    Set outSqlParamsList = Nothing
    If outputColumnAliases Is Nothing Then
        private_ShowConfigError "MultiSourcesView output columns are not specified."
        Exit Function
    End If
    If outputColumnAliases.Count = 0 Then
        private_ShowConfigError "MultiSourcesView output columns are not specified."
        Exit Function
    End If
    If Not private_TryBuildConfigMap(cfgMap) Then Exit Function
    If Not m_CfgTableParser.TryParseTableRefToken(tableRef, sourceAlias, tableAlias) Then
        private_ShowConfigError "Invalid table reference '" & tableRef & "'. Expected <SourceAlias>.Sheet[<TableAlias>]."
        Exit Function
    End If

    tablePathPrefix = m_CfgTableParser.BuildTablePathPrefix(sourceAlias, tableAlias)
    sourceKeyPrefix = "Source." & sourceAlias & "."
    sourcePathKey = sourceKeyPrefix & "FilePath"
    resolverKey = sourceKeyPrefix & "FilePathResolver"
    resolverArgsKey = sourceKeyPrefix & "FilePathResolverArgs"
    runtimeAliasPatternKey = sourceKeyPrefix & "RuntimeAliasPattern"
    If cfgMap.Exists(sourcePathKey) Then rawSourcePath = VBA.Trim$(VBA.CStr(cfgMap(sourcePathKey)))
    sourcePathPattern = rawSourcePath
    resolverName = m_CfgTableParser.CfgParserBase.GetOptionalConfigValue(cfgMap, resolverKey)
    resolverArgs = m_CfgTableParser.CfgParserBase.GetOptionalConfigValue(cfgMap, resolverArgsKey)
    runtimeAliasPattern = m_CfgTableParser.CfgParserBase.GetOptionalConfigValue(cfgMap, runtimeAliasPatternKey)
    If VBA.Len(VBA.Trim$(rawSourcePath)) = 0 Then
        private_ShowConfigError "Missing source path for alias '" & sourceAlias & "'."
        Exit Function
    End If

    If VBA.InStr(1, resolverName, "ResolveAllByDmyPattern", _
        VBA.vbTextCompare) > 0 Then
        ' Коллекция файлов требует уникальных runtime-алиасов. Стабильный
        ' sourceAlias остается ключом конфигурации, а отдельный шаблон
        ' определяет имя каждого физического экземпляра источника.
        If VBA.Len(VBA.Trim$(runtimeAliasPattern)) = 0 Then
            private_ShowConfigError "Missing " & runtimeAliasPatternKey & _
                " for multi-source alias '" & sourceAlias & "'."
            Exit Function
        End If
        On Error GoTo EH_RESOLVE_ALL
        Set sourcePaths = ResolveAllByDmyPattern(rawSourcePath, resolverArgs)
        On Error GoTo 0
    ElseIf VBA.InStr(1, resolverName, "ResolveLatestByDmyPattern", _
        VBA.vbTextCompare) > 0 Then
        Set sourcePaths = New Collection
        On Error GoTo EH_RESOLVE_LATEST
        sourcePaths.Add ResolveLatestByDmyPattern(rawSourcePath)
        On Error GoTo 0
    Else
        Set sourcePaths = New Collection
        If Not m_CfgTableParser.TryResolveSourcePath(cfgMap, sourceAlias, rawSourcePath) Then Exit Function
        sourcePaths.Add rawSourcePath
    End If
    If sourcePaths Is Nothing Then
        private_ShowConfigError "Source resolver returned no files for alias '" & sourceAlias & "'."
        Exit Function
    End If
    If sourcePaths.Count = 0 Then
        private_ShowConfigError "Source resolver returned no files for alias '" & sourceAlias & "'."
        Exit Function
    End If

    Set outSqlParamsList = New Collection
    For Each sourcePath In sourcePaths
        If Not private_TryBuildSqlParamsForPath( _
            cfgMap, tablePathPrefix, VBA.CStr(sourcePath), outputColumnAliases, sqlParams) Then Exit Function
        If VBA.Len(VBA.Trim$(runtimeAliasPattern)) > 0 Then
            sqlParams.SourceAliasTemplate = sourceAlias
            On Error GoTo EH_EXPAND_ALIAS
            sqlParams.SourceAlias = ex_SourceResolver.fn_ExpandDmyRuntimeAliasByResolvedPath( _
                runtimeAliasPattern, sourcePathPattern, VBA.CStr(sourcePath))
            On Error GoTo 0
        Else
            sqlParams.SourceAlias = sourceAlias
            sqlParams.SourceAliasTemplate = VBA.vbNullString
        End If
        outSqlParamsList.Add sqlParams
    Next sourcePath
    TryBuildTableSqlParamsList = True
    Exit Function

EH_RESOLVE_ALL:
    private_ShowConfigError "Failed to resolve source files for alias '" & sourceAlias & "': " & Err.Description
    Err.Clear
    On Error GoTo 0
    Exit Function

EH_RESOLVE_LATEST:
    private_ShowConfigError "Failed to resolve latest source file for alias '" & _
        sourceAlias & "': " & Err.Description
    Err.Clear
    On Error GoTo 0
    Exit Function

EH_EXPAND_ALIAS:
    private_ShowConfigError "Failed to expand runtime alias pattern '" & runtimeAliasPattern & _
        "' for file '" & VBA.CStr(sourcePath) & "': " & Err.Description
    Err.Clear
    On Error GoTo 0
End Function

Private Function private_TryBuildSqlParamsForPath( _
    ByVal cfgMap As Object, _
    ByVal tablePathPrefix As String, _
    ByVal sourcePath As String, _
    ByVal outputColumnAliases As Collection, _
    ByRef outSqlParams As obj_SqlParams _
) As Boolean
    Dim sheetName As String
    Dim rangeStartMarker As String
    Dim rangeEndMarker As String
    Dim aliasItem As Variant
    Dim columnAlias As String
    Dim sourceColumnHeader As String
    Dim mappedColumnHeader As String
    Dim sqlParams As obj_SqlParams

    Set outSqlParams = Nothing
    If Not m_CfgTableParser.TryResolveSheetName(cfgMap, tablePathPrefix, sheetName) Then Exit Function
    If Not m_CfgTableParser.TryResolveRangeMarkers( _
        cfgMap, tablePathPrefix, rangeStartMarker, rangeEndMarker) Then Exit Function

    Set sqlParams = New obj_SqlParams
    sqlParams.SourcePath = sourcePath
    sqlParams.SheetName = sheetName
    sqlParams.RangeStartMarker = rangeStartMarker
    sqlParams.RangeEndMarker = rangeEndMarker
    For Each aliasItem In outputColumnAliases
        columnAlias = VBA.Trim$(VBA.CStr(aliasItem))
        If VBA.Len(columnAlias) > 0 Then
            If Not m_CfgTableParser.TryResolveMapByColumnAlias( _
                cfgMap, tablePathPrefix, columnAlias, sourceColumnHeader, mappedColumnHeader) Then Exit Function
            If Not sqlParams.AddColumnMapping( _
                sourceColumnHeader, mappedColumnHeader, columnAlias) Then Exit Function
        End If
    Next aliasItem

    Set outSqlParams = sqlParams
    private_TryBuildSqlParamsForPath = True
End Function

Private Function private_TryBuildConfigMap(ByRef outCfgMap As Object) As Boolean
    Dim configEntries As Collection

    Set outCfgMap = Nothing
    If m_CfgTableParser Is Nothing Then Exit Function
    If m_CfgTableParser.CfgParserBase Is Nothing Then Exit Function
    If Not m_CfgTableParser.CfgParserBase.TryGetConfigEntries(configEntries) Then Exit Function
    If Not m_CfgTableParser.CfgParserBase.BuildConfigDictionary(configEntries, outCfgMap) Then Exit Function
    If outCfgMap Is Nothing Then Exit Function
    private_TryBuildConfigMap = True
End Function

Private Function private_ValidateTableRef(ByVal tableRef As String, ByVal keyName As String) As Boolean
    Dim sourceAlias As String
    Dim tableAlias As String

    If m_CfgTableParser.TryParseTableRefToken(tableRef, sourceAlias, tableAlias) Then
        private_ValidateTableRef = True
        Exit Function
    End If

    private_ShowConfigError keyName & " has invalid format. Expected <SourceAlias>.Sheet[<TableAlias>]."
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
        private_ShowConfigError "Required config key is missing or empty: " & keyName
        Exit Function
    End If

    private_TryGetRequiredConfigValue = True
End Function

Private Function private_ValidateColumnList(ByVal rawColumns As String, ByVal keyName As String) As Boolean
    Dim columns As Collection
    Dim columnFormats As Object

    If VBA.StrComp(keyName, "Comparing.CompareColumns", VBA.vbTextCompare) = 0 Then
        If Not private_TryParseCompareColumnSpecs(rawColumns, columns, columnFormats) Then Exit Function
    Else
        Set columns = m_CfgTableParser.CfgParserBase.SplitListToCollection(rawColumns)
    End If
    If columns Is Nothing Then Exit Function
    If columns.Count <= 0 Then
        private_ShowConfigError keyName & " is empty."
        Exit Function
    End If

    private_ValidateColumnList = True
End Function

Private Function private_TryParseCompareColumnSpecs( _
    ByVal rawColumns As String, _
    ByRef outColumns As Collection, _
    ByRef outColumnFormats As Object _
) As Boolean
    Dim tokens As Collection
    Dim token As Variant
    Dim columnAlias As String
    Dim formatKind As String

    Set outColumns = New Collection
    Set outColumnFormats = Nothing
    Set outColumnFormats = VBA.CreateObject("Scripting.Dictionary")
    outColumnFormats.CompareMode = 1

    Set tokens = private_SplitColumnSpecList(rawColumns)
    If tokens Is Nothing Then Exit Function

    For Each token In tokens
        If Not private_TryParseColumnSpec(VBA.CStr(token), columnAlias, formatKind) Then Exit Function
        If VBA.Len(columnAlias) = 0 Then GoTo ContinueToken

        outColumns.Add columnAlias
        If VBA.Len(formatKind) > 0 Then outColumnFormats(columnAlias) = formatKind

ContinueToken:
    Next token

    private_TryParseCompareColumnSpecs = True
End Function

Private Function private_SplitColumnSpecList(ByVal rawColumns As String) As Collection
    Dim result As Collection
    Dim token As String
    Dim i As Long
    Dim ch As String
    Dim braceDepth As Long

    Set result = New Collection
    rawColumns = VBA.Trim$(VBA.CStr(rawColumns))
    If VBA.Len(rawColumns) = 0 Then
        Set private_SplitColumnSpecList = result
        Exit Function
    End If

    ' CompareColumns поддерживает метаданные в postfix-скобках:
    ' PlannedReturnDate{fmt:Date}. Поэтому список режем по ;/, только вне {...}.
    For i = 1 To VBA.Len(rawColumns)
        ch = VBA.Mid$(rawColumns, i, 1)
        If ch = "{" Then
            braceDepth = braceDepth + 1
            token = token & ch
        ElseIf ch = "}" Then
            If braceDepth > 0 Then braceDepth = braceDepth - 1
            token = token & ch
        ElseIf (ch = ";" Or ch = ",") And braceDepth = 0 Then
            token = VBA.Trim$(token)
            If VBA.Len(token) > 0 Then result.Add token
            token = VBA.vbNullString
        Else
            token = token & ch
        End If
    Next i

    If braceDepth <> 0 Then
        private_ShowConfigError "Comparing.CompareColumns has unbalanced metadata braces."
        Exit Function
    End If

    token = VBA.Trim$(token)
    If VBA.Len(token) > 0 Then result.Add token
    Set private_SplitColumnSpecList = result
End Function

Private Function private_TryParseColumnSpec( _
    ByVal rawSpec As String, _
    ByRef outColumnAlias As String, _
    ByRef outFormatKind As String _
) As Boolean
    Dim openPos As Long
    Dim closePos As Long
    Dim metaText As String

    outColumnAlias = VBA.vbNullString
    outFormatKind = VBA.vbNullString
    rawSpec = VBA.Trim$(VBA.CStr(rawSpec))
    If VBA.Len(rawSpec) = 0 Then
        private_TryParseColumnSpec = True
        Exit Function
    End If

    openPos = VBA.InStr(1, rawSpec, "{", VBA.vbBinaryCompare)
    If openPos <= 0 Then
        outColumnAlias = rawSpec
        private_TryParseColumnSpec = True
        Exit Function
    End If

    closePos = VBA.InStrRev(rawSpec, "}")
    If closePos <= openPos Or closePos <> VBA.Len(rawSpec) Then
        private_ShowConfigError "Invalid Comparing.CompareColumns metadata: " & rawSpec
        Exit Function
    End If

    outColumnAlias = VBA.Trim$(VBA.Left$(rawSpec, openPos - 1))
    metaText = VBA.Mid$(rawSpec, openPos + 1, closePos - openPos - 1)
    If VBA.Len(outColumnAlias) = 0 Then
        private_ShowConfigError "Invalid Comparing.CompareColumns metadata: column alias is empty."
        Exit Function
    End If
    If Not private_TryParseColumnMeta(metaText, outFormatKind) Then Exit Function

    private_TryParseColumnSpec = True
End Function

Private Function private_TryParseColumnMeta( _
    ByVal metaText As String, _
    ByRef outFormatKind As String _
) As Boolean
    Dim parts As Variant
    Dim part As Variant
    Dim pairParts As Variant
    Dim metaKey As String
    Dim metaValue As String

    outFormatKind = VBA.vbNullString
    metaText = VBA.Trim$(VBA.CStr(metaText))
    If VBA.Len(metaText) = 0 Then
        private_TryParseColumnMeta = True
        Exit Function
    End If

    parts = VBA.Split(metaText, ",")
    For Each part In parts
        pairParts = VBA.Split(VBA.CStr(part), ":")
        If UBound(pairParts) <> 1 Then
            private_ShowConfigError "Invalid Comparing.CompareColumns metadata item: " & VBA.CStr(part)
            Exit Function
        End If

        metaKey = VBA.LCase$(VBA.Trim$(VBA.CStr(pairParts(0))))
        metaValue = VBA.Trim$(VBA.CStr(pairParts(1)))
        Select Case metaKey
            Case "fmt", "format"
                Select Case VBA.LCase$(metaValue)
                    Case "date"
                        outFormatKind = "date"
                    Case Else
                        private_ShowConfigError "Unsupported Comparing.CompareColumns format: " & metaValue
                        Exit Function
                End Select
            Case Else
                private_ShowConfigError "Unsupported Comparing.CompareColumns metadata key: " & metaKey
                Exit Function
        End Select
    Next part

    private_TryParseColumnMeta = True
End Function

Private Function private_ParseBoolean(ByVal rawValue As String, ByVal defaultValue As Boolean) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(rawValue))
        Case "true", "1", "yes", "y", "да"
            private_ParseBoolean = True
        Case "false", "0", "no", "n", "нет"
            private_ParseBoolean = False
        Case Else
            private_ParseBoolean = defaultValue
    End Select
End Function

Private Sub private_ShowConfigError(ByVal messageText As String)
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "ComparingCfgParser: " & VBA.CStr(messageText)
#End If
    VBA.MsgBox "PrototypeNew: " & VBA.CStr(messageText), vbExclamation, RUNTIME_ERROR_TITLE
End Sub
