VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PrsnlEvntBuilderCfgParser"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const PROFILES_PROVIDER_CLASS_KEY As String = "EntityLookup.ProfilesProviderClass"
Private Const ENTITY_LOOKUP_TABLE_COLUMNS_KEY As String = "EntityLookup.Table.Columns"
Private Const ENTITY_LOOKUP_TABLE_COLUMN_PREFIX As String = "EntityLookup.Table.Column["
Private Const ENTITY_LOOKUP_TABLE_COLUMN_CAPTION_SUFFIX As String = "].Caption"
Private Const EXPORT_CONFIG_PREFIX As String = "Export."
Private Const EXPORT_FILE_PATH_SUFFIX As String = ".FilePath"
Private Const EXPORT_CLASS_SUFFIX As String = ".ExporterClass"
Private Const EXPORT_SHEET_NAME_SUFFIX As String = ".SheetName"
Private Const EXPORT_RANGE_START_MARKER_SUFFIX As String = ".RangeStartMarker"
Private Const EXPORT_RANGE_END_MARKER_SUFFIX As String = ".RangeEndMarker"
Private Const DEFAULT_EXPORTER_CLASS As String = "obj_PEB_ExptrDailyScope"

Private m_ConfigTable As obj_ConfigTable
Private m_CfgParserBase As obj_CfgParserBase
Private m_ConfigEntries As Collection
Private m_CfgMap As Object
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
Public Function Initialize(ByVal configTable As obj_ConfigTable) As Boolean
    m_IsDisposed = False
    Set m_ConfigTable = configTable
    Set m_CfgParserBase = Nothing
    Set m_ConfigEntries = Nothing
    Set m_CfgMap = Nothing

    If m_ConfigTable Is Nothing Then Exit Function

    Set m_CfgParserBase = New obj_CfgParserBase
    If Not m_CfgParserBase.Initialize(m_ConfigTable) Then Exit Function
    If Not m_CfgParserBase.TryGetConfigEntries(m_ConfigEntries) Then Exit Function
    If Not m_CfgParserBase.BuildConfigDictionary(m_ConfigEntries, m_CfgMap) Then Exit Function

    Initialize = True
End Function

Public Sub Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Set m_ConfigTable = Nothing
    Set m_CfgParserBase = Nothing
    Set m_ConfigEntries = Nothing
    Set m_CfgMap = Nothing
    On Error GoTo 0
End Sub

Public Function TryGetProfilesProviderClass(ByRef outProviderClassName As String) As Boolean
    outProviderClassName = VBA.vbNullString
    If m_CfgParserBase Is Nothing Then Exit Function
    If m_CfgMap Is Nothing Then Exit Function

    If Not m_CfgParserBase.TryGetRequiredConfigValue(m_CfgMap, PROFILES_PROVIDER_CLASS_KEY, outProviderClassName) Then
        VBA.MsgBox "PrototypeNew: required config key '" & PROFILES_PROVIDER_CLASS_KEY & "' is missing.", VBA.vbExclamation, "PrototypeNew / PrsnlEvntBuilder"
        Exit Function
    End If

    TryGetProfilesProviderClass = True
End Function

Public Function TryGetEntityLookupColumnAliasByCaption(ByRef outAliasByCaption As Object) As Boolean
    Dim columnAliasesText As String
    Dim columnAliases As Collection
    Dim aliasObj As Variant
    Dim columnAlias As String
    Dim columnCaption As String

    Set outAliasByCaption = ex_Helpers.fn_CreateDictionaryTextCompare()
    If m_CfgParserBase Is Nothing Then Exit Function
    If m_CfgMap Is Nothing Then Exit Function

    columnAliasesText = m_CfgParserBase.GetOptionalConfigValue(m_CfgMap, ENTITY_LOOKUP_TABLE_COLUMNS_KEY, VBA.vbNullString)
    Set columnAliases = m_CfgParserBase.SplitListToCollection(columnAliasesText)
    If columnAliases Is Nothing Then
        TryGetEntityLookupColumnAliasByCaption = True
        Exit Function
    End If

    For Each aliasObj In columnAliases
        columnAlias = VBA.Trim$(VBA.CStr(aliasObj))
        If VBA.Len(columnAlias) = 0 Then GoTo ContinueAlias

        columnCaption = m_CfgParserBase.GetOptionalConfigValue( _
            m_CfgMap, _
            ENTITY_LOOKUP_TABLE_COLUMN_PREFIX & columnAlias & ENTITY_LOOKUP_TABLE_COLUMN_CAPTION_SUFFIX, _
            VBA.vbNullString)
        columnCaption = VBA.Trim$(columnCaption)

        ' Видимый caption на листе остается именем колонки, а стабильный ключ
        ' из конфига сохраняем как alias DynamicTable.
        If VBA.Len(columnCaption) > 0 Then outAliasByCaption(columnCaption) = columnAlias

ContinueAlias:
    Next aliasObj

    TryGetEntityLookupColumnAliasByCaption = True
End Function

Public Function TryGetExportSettings( _
    ByRef outExportAliases As Collection, _
    ByRef outExporterClassByAlias As Object, _
    ByRef outExportConfigTableByAlias As Object _
) As Boolean
    Dim entryObj As Variant
    Dim configEntry As obj_ConfigEntry
    Dim keyText As String
    Dim exportAlias As String
    Dim keySuffix As String
    Dim aliasObj As Variant
    Dim exporterClassName As String
    Dim targetWorkbookPath As String
    Dim targetSheetName As String
    Dim rangeStartMarker As String
    Dim rangeEndMarker As String
    Dim exportConfigTable As obj_ConfigTable

    Set outExportAliases = New Collection
    Set outExporterClassByAlias = ex_Helpers.fn_CreateDictionaryTextCompare()
    Set outExportConfigTableByAlias = ex_Helpers.fn_CreateDictionaryTextCompare()

    If m_CfgParserBase Is Nothing Then Exit Function
    If m_CfgMap Is Nothing Then Exit Function
    If m_ConfigEntries Is Nothing Then
        TryGetExportSettings = True
        Exit Function
    End If

    For Each entryObj In m_ConfigEntries
        If Not VBA.IsObject(entryObj) Then GoTo ContinueEntry
        Set configEntry = Nothing
        On Error Resume Next
        Set configEntry = entryObj
        On Error GoTo 0
        If configEntry Is Nothing Then GoTo ContinueEntry

        keyText = VBA.Trim$(configEntry.Key)
        If Not private_TryParseExportConfigKey(keyText, exportAlias, keySuffix) Then GoTo ContinueEntry
        If Not private_ExportAliasExists(outExportAliases, exportAlias) Then outExportAliases.Add exportAlias

ContinueEntry:
    Next entryObj

    For Each aliasObj In outExportAliases
        exportAlias = VBA.Trim$(VBA.CStr(aliasObj))
        If VBA.Len(exportAlias) = 0 Then GoTo ContinueAlias

        exporterClassName = m_CfgParserBase.GetOptionalConfigValue( _
            m_CfgMap, _
            EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_CLASS_SUFFIX, _
            DEFAULT_EXPORTER_CLASS)
        exporterClassName = VBA.Trim$(exporterClassName)
        If VBA.Len(exporterClassName) = 0 Then exporterClassName = DEFAULT_EXPORTER_CLASS

        targetWorkbookPath = m_CfgParserBase.GetOptionalConfigValue( _
            m_CfgMap, _
            EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_FILE_PATH_SUFFIX, _
            VBA.vbNullString)
        targetSheetName = m_CfgParserBase.GetOptionalConfigValue( _
            m_CfgMap, _
            EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_SHEET_NAME_SUFFIX, _
            VBA.vbNullString)
        rangeStartMarker = m_CfgParserBase.GetOptionalConfigValue( _
            m_CfgMap, _
            EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_RANGE_START_MARKER_SUFFIX, _
            VBA.vbNullString)
        rangeEndMarker = m_CfgParserBase.GetOptionalConfigValue( _
            m_CfgMap, _
            EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_RANGE_END_MARKER_SUFFIX, _
            VBA.vbNullString)

        Set exportConfigTable = private_BuildExportConfigTable( _
            exportAlias, _
            exporterClassName, _
            targetWorkbookPath, _
            targetSheetName, _
            rangeStartMarker, _
            rangeEndMarker)
        If exportConfigTable Is Nothing Then Exit Function

        outExporterClassByAlias(exportAlias) = exporterClassName
        Set outExportConfigTableByAlias(exportAlias) = exportConfigTable

ContinueAlias:
    Next aliasObj

    TryGetExportSettings = True
End Function

' //
' // Internal
' //
Private Function private_TryParseExportConfigKey( _
    ByVal keyText As String, _
    ByRef outExportAlias As String, _
    ByRef outKeySuffix As String _
) As Boolean
    Dim keyLower As String
    Dim suffixPos As Long

    outExportAlias = VBA.vbNullString
    outKeySuffix = VBA.vbNullString

    keyText = VBA.Trim$(keyText)
    keyLower = VBA.LCase$(keyText)
    If VBA.Left$(keyLower, VBA.Len(VBA.LCase$(EXPORT_CONFIG_PREFIX))) <> VBA.LCase$(EXPORT_CONFIG_PREFIX) Then Exit Function

    suffixPos = VBA.InStr(VBA.Len(EXPORT_CONFIG_PREFIX) + 1, keyText, ".", VBA.vbTextCompare)
    If suffixPos <= VBA.Len(EXPORT_CONFIG_PREFIX) + 1 Then Exit Function

    outKeySuffix = VBA.Mid$(keyText, suffixPos)
    If VBA.StrComp(outKeySuffix, EXPORT_FILE_PATH_SUFFIX, VBA.vbTextCompare) <> 0 _
        And VBA.StrComp(outKeySuffix, EXPORT_CLASS_SUFFIX, VBA.vbTextCompare) <> 0 _
        And VBA.StrComp(outKeySuffix, EXPORT_SHEET_NAME_SUFFIX, VBA.vbTextCompare) <> 0 _
        And VBA.StrComp(outKeySuffix, EXPORT_RANGE_START_MARKER_SUFFIX, VBA.vbTextCompare) <> 0 _
        And VBA.StrComp(outKeySuffix, EXPORT_RANGE_END_MARKER_SUFFIX, VBA.vbTextCompare) <> 0 Then Exit Function

    outExportAlias = VBA.Trim$(VBA.Mid$(keyText, VBA.Len(EXPORT_CONFIG_PREFIX) + 1, suffixPos - VBA.Len(EXPORT_CONFIG_PREFIX) - 1))
    If VBA.Len(outExportAlias) = 0 Then Exit Function

    private_TryParseExportConfigKey = True
End Function

Private Function private_ExportAliasExists(ByVal exportAliases As Collection, ByVal exportAlias As String) As Boolean
    Dim aliasObj As Variant

    If exportAliases Is Nothing Then Exit Function
    exportAlias = VBA.Trim$(exportAlias)
    If VBA.Len(exportAlias) = 0 Then Exit Function

    For Each aliasObj In exportAliases
        If VBA.StrComp(VBA.Trim$(VBA.CStr(aliasObj)), exportAlias, VBA.vbTextCompare) = 0 Then
            private_ExportAliasExists = True
            Exit Function
        End If
    Next aliasObj
End Function

Private Function private_BuildExportConfigTable( _
    ByVal exportAlias As String, _
    ByVal exporterClassName As String, _
    ByVal targetWorkbookPath As String, _
    ByVal targetSheetName As String, _
    ByVal rangeStartMarker As String, _
    ByVal rangeEndMarker As String _
) As obj_ConfigTable
    Dim configTable As obj_ConfigTable

    Set configTable = New obj_ConfigTable
    If Not configTable.Initialize() Then Exit Function

    If Not configTable.AddRow(VBA.vbNullString, EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_CLASS_SUFFIX, exporterClassName) Then Exit Function
    If Not configTable.AddRow(VBA.vbNullString, EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_FILE_PATH_SUFFIX, targetWorkbookPath) Then Exit Function
    If Not configTable.AddRow(VBA.vbNullString, EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_SHEET_NAME_SUFFIX, targetSheetName) Then Exit Function
    If Not configTable.AddRow(VBA.vbNullString, EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_RANGE_START_MARKER_SUFFIX, rangeStartMarker) Then Exit Function
    If Not configTable.AddRow(VBA.vbNullString, EXPORT_CONFIG_PREFIX & exportAlias & EXPORT_RANGE_END_MARKER_SUFFIX, rangeEndMarker) Then Exit Function

    Set private_BuildExportConfigTable = configTable
End Function
