VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PageSDBCtrl"
Option Explicit

Private Const CONTROLLER_RUNTIME_OBJECT_KEY As String = "RuntimeObjects.SupportingDocumentBuilder.Controller"
Private Const DRAFT_VALUES_CONTAINER_NAME As String = "EventDraftValues"
Private Const WORD_EXPORT_PREVIEW_CONTEXT_KEY As String = "WordExportPreviewText"
Private Const WORD_EXPORT_PREVIEW_EDITABLE_KEY As String = "WordExportPreviewEditable"
Private Const SECTION_TYPE_CONTEXT_KEY As String = "SectionType"
Private Const WRITE_TO_WORD_CONTEXT_KEY As String = "WriteToWord"
Private m_Page As obj_IPage
Private m_Data As obj_SDB_Data
Private m_ProfileConfigTable As obj_ConfigTable
Private m_WordExporterClassName As String
Private m_SelectedSection As String
Private m_WordExportPreviewText As String
Private m_IsDisposed As Boolean

Private Sub Class_Terminate()
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Me.Dispose
    On Error GoTo 0
End Sub

Public Property Get RuntimeObjectSourceKey() As String
    RuntimeObjectSourceKey = CONTROLLER_RUNTIME_OBJECT_KEY
End Property

Public Property Get WordExportPreviewText() As String
    WordExportPreviewText = m_WordExportPreviewText
End Property

Public Property Get IsWordPreviewVisible() As Boolean
    IsWordPreviewVisible = _
        (VBA.Len(VBA.Trim$(m_WordExportPreviewText)) > 0)
End Property

Public Function Initialize(ByVal page As Object) As Boolean
    Dim pageInterface As obj_IPage
    Dim pageBase As obj_PageBase

    If page Is Nothing Then
        VBA.MsgBox "SupportingDocumentBuilder controller requires a page.", _
            VBA.vbExclamation, "Supporting Document Builder"
        Exit Function
    End If
    On Error Resume Next
    Set pageInterface = page
    On Error GoTo 0
    If pageInterface Is Nothing Then
        VBA.MsgBox "SupportingDocumentBuilder page must implement obj_IPage.", _
            VBA.vbExclamation, "Supporting Document Builder"
        Exit Function
    End If
    Set m_Page = pageInterface
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Not pageBase.RuntimeSources.SetObjectSource(CONTROLLER_RUNTIME_OBJECT_KEY, Me) Then Exit Function
    m_SelectedSection = VBA.vbNullString
    m_IsDisposed = False
    Initialize = True
End Function

Public Sub Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    Set m_Page = Nothing
    Set m_Data = Nothing
    Set m_ProfileConfigTable = Nothing
    m_WordExporterClassName = VBA.vbNullString
    m_WordExportPreviewText = VBA.vbNullString
End Sub

Public Function UpdateDataFromConfigTable(ByVal configTable As obj_ConfigTable) As Boolean
    Dim cfgParser As obj_CfgParserBase
    Dim configEntries As Collection
    Dim configMap As Object
    Dim providerClassName As String
    Dim sdbFactory As obj_SDB_Factory

    If configTable Is Nothing Then
        VBA.MsgBox "SupportingDocumentBuilder configuration table is missing.", _
            VBA.vbExclamation, "Supporting Document Builder"
        Exit Function
    End If
    Set cfgParser = New obj_CfgParserBase
    If Not cfgParser.Initialize(configTable) Then Exit Function
    If Not cfgParser.TryGetConfigEntries(configEntries) Then Exit Function
    If Not cfgParser.BuildConfigDictionary(configEntries, configMap) Then Exit Function
    If Not cfgParser.TryGetRequiredConfigValue( _
        configMap, _
        "SupportingDocumentBuilder.ProfilesProviderClass", _
        providerClassName) Then
        VBA.MsgBox "SupportingDocumentBuilder.ProfilesProviderClass is required.", _
            VBA.vbExclamation, "Supporting Document Builder"
        Exit Function
    End If
    Set sdbFactory = New obj_SDB_Factory
    If Not sdbFactory.TryCreateProfilesProvider(providerClassName, m_Data) Then Exit Function

    Set m_ProfileConfigTable = configTable
    If Not cfgParser.TryGetRequiredConfigValue( _
        configMap, "Export.Word.ExporterClass", m_WordExporterClassName) Then
        VBA.MsgBox "SupportingDocumentBuilder requires Export.Word.ExporterClass.", _
            VBA.vbExclamation, "Supporting Document Builder"
        Exit Function
    End If
    UpdateDataFromConfigTable = True
End Function

Public Function PrepareRuntime(Optional ByVal notifyChange As Boolean = False) As Boolean
    If m_Data Is Nothing Then
        VBA.MsgBox "SupportingDocumentBuilder profiles provider is not initialized.", _
            VBA.vbExclamation, "Supporting Document Builder"
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(m_SelectedSection)) = 0 Then
        m_SelectedSection = VBA.Trim$(m_Data.DefaultSectionName)
    End If
    If VBA.Len(m_SelectedSection) = 0 Then
        VBA.MsgBox "SupportingDocumentBuilder does not declare a default section.", _
            VBA.vbExclamation, "Supporting Document Builder"
        Exit Function
    End If
    PrepareRuntime = True
End Function

Public Function ResolveProfileVisibilityState(ByVal tagsText As String) As String
    ResolveProfileVisibilityState = "collapsed"
    If m_Data Is Nothing Then Exit Function
    ResolveProfileVisibilityState = _
        m_Data.ResolveProfileVisibilityState(m_SelectedSection, tagsText)
End Function

Public Function OnSupportingDocumentSectionClick( _
    Optional ByVal sectionName As Variant _
) As Boolean
    Dim requestedSection As String

    If m_Data Is Nothing Then Exit Function
    If VBA.IsMissing(sectionName) Then
        requestedSection = m_Data.DefaultSectionName
    Else
        requestedSection = VBA.Trim$(VBA.CStr(sectionName))
        If VBA.Len(requestedSection) = 0 Then _
            requestedSection = m_Data.DefaultSectionName
    End If
    If Not m_Data.IsSectionName(requestedSection) Then
        VBA.MsgBox "SupportingDocumentBuilder: unsupported section '" & _
            requestedSection & "'.", VBA.vbExclamation, _
            "Supporting Document Builder"
        Exit Function
    End If
    m_SelectedSection = requestedSection
    m_WordExportPreviewText = VBA.vbNullString
    OnSupportingDocumentSectionClick = private_RerenderPage("sdb:section-selected")
End Function

Public Function OnGenerateWordPreviewClick( _
    Optional ByVal ignored As Variant _
) As Boolean
    OnGenerateWordPreviewClick = private_RunWordExporter(False)
End Function

Public Function OnExportToWordClick(Optional ByVal ignored As Variant) As Boolean
    OnExportToWordClick = private_RunWordExporter(True)
End Function

Private Function private_RunWordExporter(ByVal writeToWord As Boolean) As Boolean
    Dim sourceTables As Collection
    Dim exportContext As Object
    Dim exporter As obj_IDataExporter

    If Not private_EnsureCurrentConfig() Then Exit Function
    ' sourceTables передаёт табличные данные, а один и тот же exportContext
    ' служит двусторонним контрактом: Controller пишет параметры операции,
    ' exporter возвращает в этот словарь текст и свойства preview.
    If Not private_BuildSourceTables(sourceTables, exportContext) Then Exit Function
    exportContext(WRITE_TO_WORD_CONTEXT_KEY) = writeToWord
    If Not private_CreateWordExporter(exporter) Then Exit Function
    If Not exporter.Export(sourceTables, exportContext) Then Exit Function

    If writeToWord Then
        rt_Messaging.fn_ShowStatusBarSuccess "Export to WORD: done", 3
    Else
        If Not private_CapturePreview(exportContext) Then Exit Function
        rt_Messaging.fn_ShowStatusBarSuccess "WORD preview: done", 3
    End If
    private_RunWordExporter = True
End Function

Private Function private_BuildSourceTables( _
    ByRef outTables As Collection, _
    ByRef outContext As Object _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim sourceTable As obj_TableDynamic
    Dim sourceRow As obj_Row
    Dim fieldAliases As Variant
    Dim aliasObj As Variant
    Dim valueRange As Range
    Dim headerText As String

    Set outTables = Nothing
    Set outContext = Nothing
    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function

    fieldAliases = Array( _
        "_Rank", "_FIO", "_PositionCode", "_PositionName", "_DateFrom", _
        "_DocDate", "_DocNo", "_ReportPositionCode", "_Destination", _
        "_IncomingNo", "_IncomingDate")
    Set sourceTable = New obj_TableDynamic
    If VBA.Len(VBA.Trim$(m_SelectedSection)) = 0 Then
        VBA.MsgBox "SupportingDocumentBuilder section is not selected.", _
            VBA.vbExclamation, "Supporting Document Builder"
        Exit Function
    End If
    sourceTable.SectionTitle = m_SelectedSection
    Set sourceRow = New obj_Row

    For Each aliasObj In fieldAliases
        Set valueRange = Nothing
        If Not private_TryGetDraftValueRange( _
            pageBase, VBA.CStr(aliasObj), valueRange) Then
            VBA.MsgBox "SupportingDocumentBuilder: field '" & VBA.CStr(aliasObj) & _
                "' is not rendered.", VBA.vbExclamation, "Supporting Document Builder"
            Exit Function
        End If
        If valueRange Is Nothing Then Exit Function
        headerText = private_ReadHeaderText( _
            pageBase.Worksheet.Cells(valueRange.Row - 1, valueRange.Column))
        If VBA.Len(headerText) = 0 Then headerText = VBA.CStr(aliasObj)
        If Not private_AddSourceColumn( _
            sourceTable, headerText, VBA.CStr(aliasObj)) Then Exit Function
        sourceRow.PushCellRaw valueRange.Cells(1, 1).Text
    Next aliasObj
    If Not sourceTable.PushRow(sourceRow) Then Exit Function

    Set outTables = New Collection
    outTables.Add sourceTable
    Set outContext = VBA.CreateObject("Scripting.Dictionary")
    outContext.CompareMode = 1
    outContext(SECTION_TYPE_CONTEXT_KEY) = m_SelectedSection
    private_BuildSourceTables = True
End Function

Private Function private_TryGetDraftValueRange( _
    ByVal pageBase As obj_PageBase, _
    ByVal fieldAlias As String, _
    ByRef outRange As Range _
) As Boolean
    Dim draftValuesRange As Range
    Dim tagEntries As Collection
    Dim tagEntryObj As Variant
    Dim tagEntry As Object
    Dim candidateRange As Range
    Dim ws As Worksheet

    Set outRange = Nothing
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function
    If Not pageBase.TryGetLayoutContainerRange( _
        DRAFT_VALUES_CONTAINER_NAME, draftValuesRange) Then Exit Function
    If draftValuesRange Is Nothing Then Exit Function
    If Not pageBase.TryGetLayoutTagEntriesInRange( _
        draftValuesRange, tagEntries, "visible") Then Exit Function
    If tagEntries Is Nothing Then Exit Function

    For Each tagEntryObj In tagEntries
        If Not VBA.IsObject(tagEntryObj) Then GoTo ContinueEntry
        Set tagEntry = tagEntryObj
        If tagEntry Is Nothing Then GoTo ContinueEntry
        If Not tagEntry.Exists("Tag") Then GoTo ContinueEntry
        If VBA.StrComp( _
            VBA.Trim$(VBA.CStr(tagEntry("Tag"))), _
            VBA.Trim$(fieldAlias), _
            VBA.vbTextCompare) <> 0 Then GoTo ContinueEntry

        Set candidateRange = ws.Range( _
            ws.Cells(VBA.CLng(tagEntry("RowStart")), VBA.CLng(tagEntry("ColStart"))), _
            ws.Cells(VBA.CLng(tagEntry("RowEnd")), VBA.CLng(tagEntry("ColEnd"))))
        Set outRange = Application.Intersect(candidateRange, draftValuesRange)
        If Not outRange Is Nothing Then
            private_TryGetDraftValueRange = True
            Exit Function
        End If
ContinueEntry:
    Next tagEntryObj
End Function

Private Function private_CreateWordExporter( _
    ByRef outExporter As obj_IDataExporter _
) As Boolean
    Dim sdbFactory As obj_SDB_Factory

    Set outExporter = Nothing
    If VBA.Len(VBA.Trim$(m_WordExporterClassName)) = 0 Then
        VBA.MsgBox "SupportingDocumentBuilder WORD exporter is not configured.", _
            VBA.vbExclamation, "Supporting Document Builder"
        Exit Function
    End If
    If m_ProfileConfigTable Is Nothing Then Exit Function
    Set sdbFactory = New obj_SDB_Factory
    private_CreateWordExporter = sdbFactory.TryCreateDataExporter( _
        m_WordExporterClassName, m_ProfileConfigTable, _
        m_ProfileConfigTable, outExporter)
End Function

Private Function private_CapturePreview(ByVal exportContext As Object) As Boolean
    If exportContext Is Nothing Then Exit Function
    If Not exportContext.Exists(WORD_EXPORT_PREVIEW_CONTEXT_KEY) Then
        VBA.MsgBox "WORD exporter did not return preview text.", _
            VBA.vbExclamation, "Supporting Document Builder"
        Exit Function
    End If
    m_WordExportPreviewText = VBA.CStr( _
        exportContext(WORD_EXPORT_PREVIEW_CONTEXT_KEY))
    If VBA.Len(VBA.Trim$(m_WordExportPreviewText)) = 0 Then
        VBA.MsgBox "WORD exporter returned an empty preview.", _
            VBA.vbExclamation, "Supporting Document Builder"
        Exit Function
    End If
    If exportContext.Exists(WORD_EXPORT_PREVIEW_EDITABLE_KEY) Then
        If VBA.CBool(exportContext(WORD_EXPORT_PREVIEW_EDITABLE_KEY)) Then
            VBA.MsgBox "Table preview must be read-only.", _
                VBA.vbExclamation, "Supporting Document Builder"
            Exit Function
        End If
    End If
    If m_Page Is Nothing Then Exit Function
    ' До генерации панель Collapsed и отсутствует в активной layout-геометрии,
    ' поэтому частичный reflow самого WordExportPanel не может её раскрыть.
    ' Полный render повторно вычисляет visibility, а Page сохраняет draft-строку.
    private_CapturePreview = rt_PageManager.fn_RenderPage( _
        m_Page, "sdb:word-preview-ready")
End Function

Private Function private_EnsureCurrentConfig() As Boolean
    Dim sdbPage As obj_PageSDB
    If m_Page Is Nothing Then Exit Function
    If Not TypeOf m_Page Is obj_PageSDB Then Exit Function
    Set sdbPage = m_Page
    private_EnsureCurrentConfig = sdbPage.EnsureModeConfigCurrent()
End Function

Private Function private_RerenderPage(ByVal reasonText As String) As Boolean
    If m_Page Is Nothing Then Exit Function
    private_RerenderPage = rt_PageManager.fn_RenderPage(m_Page, reasonText)
End Function

Private Function private_ReadHeaderText(ByVal headerCell As Range) As String
    Dim mergedRange As Range
    If headerCell Is Nothing Then Exit Function
    If headerCell.MergeCells Then
        Set mergedRange = headerCell.MergeArea
        private_ReadHeaderText = VBA.Trim$(VBA.CStr(mergedRange.Cells(1, 1).Value2))
    Else
        private_ReadHeaderText = VBA.Trim$(VBA.CStr(headerCell.Value2))
    End If
End Function

Private Function private_AddSourceColumn( _
    ByVal tableObj As obj_TableDynamic, _
    ByVal columnName As String, _
    ByVal columnAlias As String _
) As Boolean
    Dim columnObj As obj_Column
    Dim normalizedAlias As String
    If tableObj Is Nothing Then Exit Function
    Set columnObj = New obj_Column
    columnObj.Name = VBA.Trim$(columnName)
    columnObj.Position = tableObj.ColumnCount + 1
    normalizedAlias = VBA.Trim$(columnAlias)
    If Not columnObj.AddAlias(normalizedAlias) Then Exit Function
    If VBA.Left$(normalizedAlias, 1) = "_" Then
        If Not columnObj.AddAlias(VBA.Mid$(normalizedAlias, 2)) Then Exit Function
    End If
    private_AddSourceColumn = tableObj.PushColumn(columnObj)
End Function
