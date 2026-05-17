VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_CfgTableParser"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private m_CfgParserBase As obj_CfgParserBase
Private m_IsDisposed As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
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
Public Function Initialize(ByVal cfgParserBase As obj_CfgParserBase) As Boolean
    m_IsDisposed = False
    Set m_CfgParserBase = cfgParserBase
    Initialize = Not m_CfgParserBase Is Nothing
End Function

Public Sub Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Set m_CfgParserBase = Nothing
    On Error GoTo 0
End Sub

Public Function BuildTablePathPrefix( _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String _
) As String
    sourceAlias = VBA.Trim$(sourceAlias)
    tableAlias = VBA.Trim$(tableAlias)
    If VBA.Len(sourceAlias) = 0 Then Exit Function
    If VBA.Len(tableAlias) = 0 Then Exit Function
    BuildTablePathPrefix = sourceAlias & ".Sheet[" & tableAlias & "]."
End Function

Public Function TryResolveSourcePath( _
    ByVal cfgMap As Object, _
    ByVal sourceAlias As String, _
    ByRef outSourcePath As String _
) As Boolean
    TryResolveSourcePath = private_TryResolveSourcePath(cfgMap, sourceAlias, outSourcePath)
End Function

Public Function TryResolveSheetName( _
    ByVal cfgMap As Object, _
    ByVal tablePathPrefix As String, _
    ByRef outSheetName As String _
) As Boolean
    TryResolveSheetName = private_TryResolveSheetName(cfgMap, tablePathPrefix, outSheetName)
End Function

Public Function TryResolveRangeMarkers( _
    ByVal cfgMap As Object, _
    ByVal tablePathPrefix As String, _
    ByRef outRangeStartMarker As String, _
    ByRef outRangeEndMarker As String _
) As Boolean
    outRangeStartMarker = VBA.vbNullString
    outRangeEndMarker = VBA.vbNullString
    If m_CfgParserBase Is Nothing Then Exit Function
    If cfgMap Is Nothing Then Exit Function

    outRangeStartMarker = m_CfgParserBase.GetOptionalConfigValue(cfgMap, tablePathPrefix & "RangeStartMarker")
    outRangeEndMarker = m_CfgParserBase.GetOptionalConfigValue(cfgMap, tablePathPrefix & "RangeEndMarker")
    If (VBA.Len(outRangeStartMarker) > 0 Xor VBA.Len(outRangeEndMarker) > 0) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "CfgTableParser: both markers are required together. Prefix='" & tablePathPrefix & "'."
#End If
        Exit Function
    End If

    TryResolveRangeMarkers = True
End Function

Public Function TryGetRequiredColumnHeadersAliases( _
    ByVal cfgMap As Object, _
    ByVal tablePathPrefix As String, _
    ByRef outColumnHeadersAliases As Collection _
) As Boolean
    Set outColumnHeadersAliases = Nothing
    If m_CfgParserBase Is Nothing Then Exit Function
    If cfgMap Is Nothing Then Exit Function
    TryGetRequiredColumnHeadersAliases = m_CfgParserBase.TryGetRequiredConfigList(cfgMap, tablePathPrefix & "ColumnHeadersAliases", outColumnHeadersAliases)
End Function

Public Function TryResolveMapByColumnAlias( _
    ByVal cfgMap As Object, _
    ByVal tablePathPrefix As String, _
    ByVal columnAlias As String, _
    ByRef outSourceColumnHeader As String, _
    ByRef outMappedColumnHeader As String _
) As Boolean
    Dim mapKey As String
    Dim mapRaw As String

    If m_CfgParserBase Is Nothing Then Exit Function
    If cfgMap Is Nothing Then Exit Function

    columnAlias = VBA.Trim$(columnAlias)
    If VBA.Len(columnAlias) = 0 Then Exit Function

    mapKey = tablePathPrefix & "Map[" & columnAlias & "]"
    If Not m_CfgParserBase.TryGetRequiredConfigValue(cfgMap, mapKey, mapRaw) Then Exit Function
    If Not TryParseMapValue(mapRaw, outSourceColumnHeader, outMappedColumnHeader) Then Exit Function

    TryResolveMapByColumnAlias = True
End Function

Public Function TryParseTableRefToken( _
    ByVal tokenText As String, _
    ByRef outSourceAlias As String, _
    ByRef outTableAlias As String _
) As Boolean
    Dim normalized As String
    Dim pos As Long
    Dim suffix As String
    Dim closePos As Long

    outSourceAlias = VBA.vbNullString
    outTableAlias = VBA.vbNullString

    ' Формат токена: <SourceAlias>.Sheet[<TableAlias>]
    tokenText = VBA.Trim$(tokenText)
    normalized = VBA.LCase$(tokenText)
    If VBA.Len(tokenText) = 0 Then Exit Function

    pos = VBA.InStr(1, normalized, ".sheet[", VBA.vbTextCompare)
    If pos <= 1 Then Exit Function

    outSourceAlias = VBA.Trim$(VBA.Left$(tokenText, pos - 1))
    suffix = VBA.Mid$(tokenText, pos + 7)
    closePos = VBA.InStr(1, suffix, "]", VBA.vbBinaryCompare)
    If closePos <= 1 Then Exit Function

    outTableAlias = VBA.Trim$(VBA.Left$(suffix, closePos - 1))
    If VBA.Len(outSourceAlias) = 0 Or VBA.Len(outTableAlias) = 0 Then Exit Function
    If VBA.Len(VBA.Trim$(VBA.Mid$(suffix, closePos + 1))) > 0 Then Exit Function

    TryParseTableRefToken = True
End Function

Public Function TryParseMapValue( _
    ByVal rawMapValue As String, _
    ByRef outSourceHeader As String, _
    ByRef outLabel As String _
) As Boolean
    Dim p As Long

    outSourceHeader = VBA.vbNullString
    outLabel = VBA.vbNullString

    ' Поддерживаем форматы Map:
    '   HeaderName
    '   HeaderName|Display Label
    ' Если Display Label не задан, используем HeaderName.
    rawMapValue = VBA.Trim$(rawMapValue)
    If VBA.Len(rawMapValue) = 0 Then Exit Function

    p = VBA.InStr(1, rawMapValue, "|", VBA.vbBinaryCompare)
    If p > 0 Then
        outSourceHeader = VBA.Trim$(VBA.Left$(rawMapValue, p - 1))
        outLabel = VBA.Trim$(VBA.Mid$(rawMapValue, p + 1))
    Else
        outSourceHeader = rawMapValue
    End If

    If VBA.Len(outSourceHeader) >= 2 Then
        If VBA.Left$(outSourceHeader, 1) = "[" And VBA.Right$(outSourceHeader, 1) = "]" Then
            outSourceHeader = VBA.Trim$(VBA.Mid$(outSourceHeader, 2, VBA.Len(outSourceHeader) - 2))
        End If
    End If

    If VBA.Len(outSourceHeader) = 0 Then Exit Function
    If VBA.Len(outLabel) = 0 Then outLabel = outSourceHeader
    TryParseMapValue = True
End Function

' //
' // Internal
' //
Private Function private_TryResolveSourcePath( _
    ByVal cfgMap As Object, _
    ByVal sourceAlias As String, _
    ByRef outSourcePath As String _
) As Boolean
    Dim pathKey As String
    Dim resolverKey As String
    Dim resolverArgsKey As String
    Dim sourcePathRaw As String
    Dim sourcePathResolved As String
    Dim resolverName As String
    Dim resolverArgs As String

    outSourcePath = VBA.vbNullString
    pathKey = "Source." & sourceAlias & ".FilePath"

    If Not m_CfgParserBase.TryGetRequiredConfigValue(cfgMap, pathKey, sourcePathRaw) Then Exit Function

    resolverKey = "Source." & sourceAlias & ".FileResolver"
    resolverArgsKey = "Source." & sourceAlias & ".FileResolverArgs"
    resolverName = m_CfgParserBase.GetOptionalConfigValue(cfgMap, resolverKey)
    resolverArgs = m_CfgParserBase.GetOptionalConfigValue(cfgMap, resolverArgsKey)

    sourcePathResolved = sourcePathRaw
    If VBA.Len(resolverName) > 0 Then
        If Not m_CfgParserBase.TryResolveWithOptionalResolver(sourcePathRaw, resolverName, resolverArgs, sourcePathResolved) Then Exit Function
    End If

    outSourcePath = m_CfgParserBase.ResolvePathLocal(sourcePathResolved)
    private_TryResolveSourcePath = (VBA.Len(outSourcePath) > 0)
End Function

Private Function private_TryResolveSheetName( _
    ByVal cfgMap As Object, _
    ByVal tablePathPrefix As String, _
    ByRef outSheetName As String _
) As Boolean
    Dim sheetNameRaw As String
    Dim sheetNameResolved As String
    Dim resolverName As String
    Dim resolverArgs As String

    outSheetName = VBA.vbNullString

    If Not m_CfgParserBase.TryGetRequiredConfigValue(cfgMap, tablePathPrefix & "SheetName", sheetNameRaw) Then Exit Function

    resolverName = m_CfgParserBase.GetOptionalConfigValue(cfgMap, tablePathPrefix & "SheetNameResolver")
    resolverArgs = m_CfgParserBase.GetOptionalConfigValue(cfgMap, tablePathPrefix & "SheetNameResolverArgs")

    sheetNameResolved = sheetNameRaw
    If VBA.Len(resolverName) > 0 Then
        If Not m_CfgParserBase.TryResolveWithOptionalResolver(sheetNameRaw, resolverName, resolverArgs, sheetNameResolved) Then Exit Function
    End If

    outSheetName = sheetNameResolved
    private_TryResolveSheetName = (VBA.Len(outSheetName) > 0)
End Function
