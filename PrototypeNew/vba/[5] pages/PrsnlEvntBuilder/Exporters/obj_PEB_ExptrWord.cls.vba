VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PEB_ExptrWord"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
#Const TVO_POSITION_FALLBACK_ENABLED = True
#Const TVO_POSITION_FALLBACK_HIGHLIGHT_ENABLED = True

Implements obj_IDataExporter

' Runtime path is relative to ThisWorkbook.Path, same as page UI paths.
Private Const WORD_RESULT_TEMPLATES_REL_PATH As String = "modes\PrsnlEvntBuilder\PrsnlEvntBuilderWordResultTemplates.xml"
Private Const PREVIEW_FALLBACK_COLOR As String = "#FF0000"
Private Const CONTEXT_SECTION_TYPE As String = "SectionType"
Private Const CONTEXT_VALIDATION_ENABLED As String = "ValidateWord"
Private Const CONTEXT_WORD_PREVIEW_TEXT As String = "WordExportPreviewText"
Private Const CONTEXT_MANUAL_ORDER_NO As String = "ManualOrderNo"
Private Const SENTINEL_SHORT_DATE As Date = #1/1/1900#
Private Const SOURCE_ALIAS_IPN As String = "IPN"
Private Const SOURCE_ALIAS_RANK As String = "Rank"
Private Const SOURCE_ALIAS_FIO As String = "FIO"
Private Const SOURCE_ALIAS_POSITION_CODE As String = "PositionCode"
Private Const SOURCE_ALIAS_POSITION_NAME As String = "PositionName"
Private Const SOURCE_ALIAS_HOSPITAL_SHORT As String = "HospitalShort"
Private Const SOURCE_ALIAS_VACATION As String = "Vacation"
Private Const SOURCE_ALIAS_REPORT_RANK As String = "ReportRank"
Private Const SOURCE_ALIAS_REPORT_PERSON As String = "ReportPerson"
Private Const SOURCE_ALIAS_REPORT_POSITION_CODE As String = "ReportPositionCode"
Private Const SOURCE_ALIAS_INCOMING_NO As String = "IncomingNo"
Private Const SOURCE_ALIAS_INCOMING_DATE As String = "IncomingDate"
Private Const SOURCE_ALIAS_DOCUMENT_NOTE As String = "DocumentNote"
Private Const SOURCE_ALIAS_DOC_DATE As String = "DocDate"
Private Const SOURCE_ALIAS_DATE_FROM As String = "DateFrom"
Private Const SOURCE_ALIAS_DURATION_DAYS As String = "DurationDays"
Private Const SOURCE_ALIAS_VH_DATE As String = "VhDate"
Private Const SOURCE_ALIAS_VLK_DATE As String = "VlkDate"
Private Const WORD_ALIAS_RANK_GENITIVE As String = "RankGenitive"
Private Const WORD_ALIAS_FIO_GENITIVE As String = "FIOGenitive"
Private Const WORD_ALIAS_FIO_ACCUSATIVE As String = "FIOAccusative"
Private Const WORD_ALIAS_FIO_INITIALS_GENITIVE As String = "FIOInitialsGenitive"
Private Const WORD_ALIAS_POSITION_GENITIVE As String = "PositionGenitive"
Private Const WORD_ALIAS_POSITION_DEFAULT As String = "Position"
Private Const WORD_ALIAS_RANK_DATIVE As String = "RankDative"
Private Const WORD_ALIAS_FIO_DATIVE As String = "FIODative"
Private Const WORD_ALIAS_POSITION_DATIVE As String = "PositionDative"
Private Const WORD_ALIAS_HOSPITAL_GENITIVE As String = "HospitalGenitive"
Private Const WORD_ALIAS_HOSPITAL_ACCUSATIVE As String = "HospitalAccusative"
Private Const WORD_ALIAS_HOSPITAL_DATIVE As String = "HospitalDative"
Private Const WORD_ALIAS_REPORT_RANK_GENITIVE As String = "ReportRankGenitive"
Private Const WORD_ALIAS_REPORT_PERSON_GENITIVE As String = "ReportPersonGenitive"
Private Const WORD_ALIAS_REPORT_PERSON_INITIALS_GENITIVE As String = "ReportPersonInitialsGenitive"
Private Const WORD_ALIAS_REPORT_POSITION_GENITIVE As String = "ReportPositionGenitive"
Private Const WORD_ALIAS_REPORTER_GENITIVE As String = "ReporterGenitive"
Private Const WORD_ALIAS_ORDER_NO As String = "OrderNo"
Private Const WORD_ALIAS_INCOMING_DATE_SHORT As String = "IncomingDateShort"
Private Const WORD_ALIAS_DOC_DATE_SHORT As String = "DocDateShort"
Private Const WORD_ALIAS_DATE_FROM_SHORT As String = "DateFromShort"
Private Const WORD_ALIAS_VH_DATE_SHORT As String = "VhDateShort"
Private Const WORD_ALIAS_VLK_DATE_SHORT As String = "VlkDateShort"

' DateTo отсутствует в draft-форме: для частичной ежегодной отпуска экспортёр
' вычисляет его из DateFrom + DurationDays - 1 и добавляет в контекст под этим
' alias. Суффикс Short означает единое компактное представление полной даты с
' определённым годом, но ещё не готовый длинный текст для WORD. Шаблон должен
' явно выбрать представление через formatter, например:
'   {[DateToShort]|dateformat:"\dd \month \yyyy року"}
Private Const WORD_ALIAS_DATE_TO_SHORT As String = "DateToShort"
Private Const WORD_ALIAS_VACATION_DAYS As String = "VacationDays"
Private Const WORD_ALIAS_ADDITIONAL_WAY_DAYS As String = "AdditionalWayDays"
Private Const WORD_ALIAS_VACATION_TOTAL_DAYS As String = "VacationTotalDays"
Private Const WORD_ALIAS_VACATION_DATES_SAME_MONTH As String = "VacationDatesSameMonth"
Private Const WORD_ALIAS_ENROLL_TO_FOOD_SUPPORT_DATE As String = "EnrollToFoodSupportDateShort"
Private Const WORD_ALIAS_REMOVE_FROM_FOOD_SUPPORT_DATE As String = "RemoveFromFoodSupportDateShort"
Private Const WORD_ALIAS_PREV_VK_NUM As String = "PrevVkNum"
Private Const WORD_ALIAS_PREV_VK_DATE As String = "PrevVkDateShort"
Private Const LATEST_MOVEMENT_DEPARTURE_ORDER_KEY As String = "Наказ вибуття"
Private Const LATEST_MOVEMENT_ESCORT_DOCUMENT_KEY As String = "Супровідний документ"

' Канонический компактный формат date aliases внутри контекста шаблона.
' Он нужен не для окончательного отображения в WORD, а чтобы formatter pipeline
' мог снова однозначно распознать дату независимо от региональных настроек Excel.
' Поэтому здесь намеренно нет названий месяцев, слова "року" и типографических
' пробелов: всё это контролируется непосредственно маской dateformat в XML.
' Этим форматом заполняются DateToShort, DateFromShort, PrevVkDateShort
' и другие вычисленные/нормализованные aliases с суффиксом Short.
Private Const WORD_SHORT_DATE_STORAGE_FORMAT As String = "dd.mm.yyyy"
Private Const WORD_ANCHOR_PREFIX As String = "{\export:"
Private Const WORD_ANCHOR_BEGIN_SUFFIX As String = "_Begin}"
Private Const WORD_ANCHOR_END_SUFFIX As String = "_End}"
Private Const WD_FIND_STOP As Long = 0
' PEB_* охватывает весь экспортированный пункт человека. Вложенная PEM_*
' охватывает первый символ пункта и хранит ключ сортировки в своём имени.
Private Const WORD_RECORD_BOOKMARK_PREFIX As String = "PEB_"
Private Const WORD_METADATA_BOOKMARK_PREFIX As String = "PEM_"
Private Const WORD_BOOKMARK_MAX_LENGTH As Long = 40
Private Const WORD_TEMPLATE_TO_HOSPITAL As String = "ToHospital"
Private Const WORD_TEMPLATE_FROM_HOSPITAL As String = "FromHospital"
Private Const REPORT_TVO_TEXT As String = "тимчасово виконуючого обов'язки"
Private Const META_SECTION_TYPE_DOCUMENT As String = "Мета: документ"
Private Const META_SECTION_TYPE_TVO As String = "Мета: ТВО"
Private Const LOOP_COLLECTION_META_DOCUMENT_TABLES As String = "MetaDocumentTables"
Private Const LOOP_COLLECTION_META_TVO_TABLES As String = "MetaTvoTables"


Private m_IsDisposed As Boolean
Private m_Base As obj_DataExporterBase
Private m_TemplateParser As obj_PEB_WordResultTplParser
Private m_ExporterDataProvider As obj_PEB_ExptrDataPrvdr

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Me.Dispose
    On Error GoTo 0
End Sub

' //
' // Interface
' //
Private Function obj_IDataExporter_Export( _
    ByVal sourceTables As Collection, _
    Optional ByVal context As Object = Nothing _
) As Boolean
    obj_IDataExporter_Export = Me.Export(sourceTables, context)
End Function

' //
' // API
' //
Public Function Initialize( _
    ByVal configTable As obj_ConfigTable, _
    Optional ByVal profileConfigTable As obj_ConfigTable = Nothing _
) As Boolean
    Dim dataProviderConfigTable As obj_ConfigTable

    m_IsDisposed = False
    Set m_Base = New obj_DataExporterBase
    Set m_TemplateParser = New obj_PEB_WordResultTplParser
    Set m_ExporterDataProvider = New obj_PEB_ExptrDataPrvdr
    Set dataProviderConfigTable = configTable
    If Not profileConfigTable Is Nothing Then Set dataProviderConfigTable = profileConfigTable

    If Not m_Base.Initialize(configTable, "WORD", "PrototypeNew / WORD export") Then Exit Function
    If Not m_TemplateParser.Initialize(WORD_RESULT_TEMPLATES_REL_PATH) Then Exit Function
    If Not m_ExporterDataProvider.Initialize(dataProviderConfigTable) Then Exit Function

    Initialize = True
End Function

Public Sub Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    If Not m_Base Is Nothing Then m_Base.Dispose
    If Not m_TemplateParser Is Nothing Then m_TemplateParser.Dispose
    If Not m_ExporterDataProvider Is Nothing Then m_ExporterDataProvider.Dispose
    Set m_Base = Nothing
    Set m_TemplateParser = Nothing
    Set m_ExporterDataProvider = Nothing
    On Error GoTo 0
End Sub

Public Function Export( _
    ByVal sourceTables As Collection, _
    Optional ByVal context As Object = Nothing _
) As Boolean
    Dim sourceTable As obj_TableDynamic
    Dim namedCollections As Object
    Dim sectionTypeText As String
    Dim previewText As String
    Dim templateId As String
    Dim recordIpn As String
    Dim exportValidationError As String
    Dim latestMovementTvoChain As Collection
    Dim latestMovementRecord As Object
    Dim builderData As obj_PrsnlEvntBuilderData
    Dim validationEnabled As Boolean
    Dim groupingHospitalShort As String
    Dim groupingDateShort As String

    If m_IsDisposed Then
        VBA.MsgBox "PrototypeNew: WORD exporter is disposed.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    If Not m_Base.TryGetMainSourceTable(sourceTables, sourceTable) Then Exit Function

    sectionTypeText = private_GetContextText(context, CONTEXT_SECTION_TYPE)
    If VBA.Len(sectionTypeText) = 0 Then sectionTypeText = VBA.Trim$(sourceTable.SectionTitle)
    If VBA.Len(sectionTypeText) = 0 Then
        VBA.MsgBox "PrototypeNew: WORD export requires SectionType in export context or source table SectionTitle.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    ' WORD validation defaults to disabled when the context key is absent.
    validationEnabled = private_GetContextBoolean(context, CONTEXT_VALIDATION_ENABLED)
    If Not m_ExporterDataProvider.IsExportAllowed(sourceTable, sectionTypeText, exportValidationError, latestMovementTvoChain, latestMovementRecord, validationEnabled) Then
        VBA.MsgBox exportValidationError, VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    If Not private_TryEnrichMainSourceTableForWord(sourceTable, context) Then Exit Function
    If Not private_TryEnrichPreviousVacationTicketForWord(sourceTable, sectionTypeText, latestMovementRecord) Then Exit Function
    If Not private_TryNormalizeDocumentNotesForWord(sourceTables, sectionTypeText) Then Exit Function
    If Not private_TryEnrichMetaDocumentDatesForWord(sourceTables) Then Exit Function
    If Not private_TryAppendMovementTvoTablesForReturn(sourceTables, sourceTable, sectionTypeText, latestMovementTvoChain) Then Exit Function
    If Not private_TryEnrichMetaTvoTablesForWord(sourceTables) Then Exit Function
    Set namedCollections = private_BuildNamedLoopCollections(sourceTables)
    If namedCollections Is Nothing Then Exit Function
    Set builderData = New obj_PrsnlEvntBuilderData
    If Not builderData.TryResolveWordTemplateId(sectionTypeText, templateId) Then
        VBA.MsgBox "PrototypeNew: WORD result template is not mapped for section: " & sectionTypeText, VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    If Not m_TemplateParser.TryRenderForTemplateId(templateId, sectionTypeText, sourceTables, namedCollections, previewText) Then Exit Function
    If Not private_TrySetContextText(context, CONTEXT_WORD_PREVIEW_TEXT, previewText) Then Exit Function

    ' CTRL+3 is the preview action. The dedicated button/CTRL+4 passes
    ' WriteToWord=True and persists the same rendered text in the document.
    If private_GetContextBoolean(context, "WriteToWord") Then
        If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_IPN, recordIpn) Then
            VBA.MsgBox "PrototypeNew: WORD export requires IPN to create a record bookmark.", VBA.vbExclamation, "PrototypeNew / WORD export"
            Exit Function
        End If
        recordIpn = VBA.Trim$(recordIpn)
        If VBA.Len(recordIpn) = 0 Then
            VBA.MsgBox "PrototypeNew: WORD export requires a non-empty IPN to create a record bookmark.", VBA.vbExclamation, "PrototypeNew / WORD export"
            Exit Function
        End If
        If private_IsHospitalGroupingTemplate(templateId) Then
            ' Метаданные нужны только двум больничным секциям. Значения берём
            ' из уже обогащённого WORD-контекста. Как и WORD-шаблон FromHospital,
            ' при пустом поле "З" используем дату входящего документа.
            If Not private_TryGetMainTableValue( _
                sourceTable, SOURCE_ALIAS_HOSPITAL_SHORT, groupingHospitalShort) Then Exit Function
            If Not private_TryGetMainTableValue( _
                sourceTable, WORD_ALIAS_DATE_FROM_SHORT, groupingDateShort) Then Exit Function
            If VBA.Len(VBA.Trim$(groupingDateShort)) = 0 Then
                If Not private_TryGetMainTableValue( _
                    sourceTable, WORD_ALIAS_INCOMING_DATE_SHORT, groupingDateShort) Then Exit Function
            End If
        End If
        If Not private_TryAppendBeforeWordEndAnchor( _
            templateId, recordIpn, previewText, groupingHospitalShort, groupingDateShort) Then Exit Function
    End If

    Export = True
End Function

Private Function private_TryAppendBeforeWordEndAnchor( _
    ByVal templateId As String, _
    ByVal recordIpn As String, _
    ByVal renderedText As String, _
    Optional ByVal groupingHospitalShort As String = "", _
    Optional ByVal groupingDateShort As String = "" _
) As Boolean
    Dim targetPath As String
    Dim templatePath As String
    Dim beginMarker As String
    Dim endMarker As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim beginRange As Object
    Dim endSearchRange As Object
    Dim endRange As Object
    Dim insertRange As Object
    Dim insertedStart As Long
    Dim insertedEnd As Long
    Dim plainRenderedText As String
    Dim documentOpened As Boolean
    Dim errorDescription As String
    Dim bookmarkName As String
    Dim metadataBookmarkName As String
    Dim metadataRange As Object

    On Error GoTo EH

    templatePath = VBA.Trim$(m_Base.TargetWorkbookPath)
    If VBA.Len(templatePath) = 0 Then
        VBA.MsgBox "PrototypeNew: required profile key 'Export.Word.FilePath' is empty.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    If Not private_IsAbsolutePath(templatePath) Then templatePath = ThisWorkbook.Path & Application.PathSeparator & templatePath
    If VBA.Len(VBA.Dir$(templatePath, VBA.vbNormal Or VBA.vbReadOnly Or VBA.vbHidden Or VBA.vbSystem)) = 0 Then
        VBA.MsgBox "PrototypeNew: WORD template file was not found: " & templatePath, VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    targetPath = private_BuildResultDocumentPath(templatePath)
    If VBA.Len(targetPath) = 0 Then
        VBA.MsgBox "PrototypeNew: failed to build the WORD result path from template: " & templatePath, VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    ' Export.Word.FilePath always points to an immutable template. The first
    ' export creates a sibling *_result document; subsequent exports append to it.
    If VBA.Len(VBA.Dir$(targetPath, VBA.vbNormal Or VBA.vbReadOnly Or VBA.vbHidden Or VBA.vbSystem)) = 0 Then
        VBA.FileCopy templatePath, targetPath
    End If

    ' Never write through a document which is open in Word (ours or another
    ' process). This also avoids silently editing a user's unsaved document.
    If private_IsDocumentOpenInRunningWord(targetPath) Or private_IsFileLocked(targetPath) Then
        VBA.MsgBox "Невозможно выполнить экспорт, пока документ открыт." & VBA.vbCrLf & _
            "Закройте документ и повторите попытку:" & VBA.vbCrLf & targetPath, _
            VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    beginMarker = WORD_ANCHOR_PREFIX & VBA.Trim$(templateId) & WORD_ANCHOR_BEGIN_SUFFIX
    endMarker = WORD_ANCHOR_PREFIX & VBA.Trim$(templateId) & WORD_ANCHOR_END_SUFFIX
    plainRenderedText = private_StripPreviewColorMarkers(renderedText)
    bookmarkName = private_BuildRecordBookmarkName(templateId, recordIpn)
    If VBA.Len(bookmarkName) = 0 Then
        VBA.MsgBox "PrototypeNew: failed to build a WORD bookmark for IPN '" & recordIpn & "'.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    If private_IsHospitalGroupingTemplate(templateId) Then
        metadataBookmarkName = private_BuildGroupingMetadataBookmarkName( _
            groupingHospitalShort, groupingDateShort)
    End If

    If Not rt_PEB_WordExportRuntime.fn_GetOrCreateWordApp(wordApp) Then Exit Function
    Set wordDoc = wordApp.Documents.Open(targetPath, False, False, False)
    documentOpened = True

    If Not private_TryFindWordText(wordDoc.Content, beginMarker, beginRange) Then
        VBA.MsgBox "PrototypeNew: WORD begin anchor was not found: " & beginMarker, VBA.vbExclamation, "PrototypeNew / WORD export"
        GoTo CleanFail
    End If
    Set insertRange = wordDoc.Range(beginRange.End, wordDoc.Content.End)
    If Not private_TryFindWordText(insertRange, endMarker, endRange) Then
        VBA.MsgBox "PrototypeNew: WORD end anchor was not found after begin anchor: " & endMarker, VBA.vbExclamation, "PrototypeNew / WORD export"
        GoTo CleanFail
    End If

    bookmarkName = private_BuildUniqueBookmarkName(wordDoc, bookmarkName)
    If VBA.Len(bookmarkName) = 0 Then
        VBA.MsgBox "PrototypeNew: failed to create a unique WORD bookmark for IPN '" & recordIpn & "'.", VBA.vbExclamation, "PrototypeNew / WORD export"
        GoTo CleanFail
    End If
    If VBA.Len(metadataBookmarkName) > 0 Then
        metadataBookmarkName = private_BuildUniqueBookmarkName(wordDoc, metadataBookmarkName)
        If VBA.Len(metadataBookmarkName) = 0 Then GoTo CleanFail
    End If

    ' Append the rendered text exactly as produced by the export template.
    ' Separators between records must be defined explicitly in that template.
    ' Existing records and their bookmarks are left untouched.
    insertedStart = endRange.Start
    Set insertRange = wordDoc.Range(insertedStart, insertedStart)
    insertRange.Text = plainRenderedText
    insertedEnd = insertedStart + VBA.Len(plainRenderedText)
    Set insertRange = wordDoc.Range(insertedStart, insertedEnd)
    ' Do not inherit highlight from a neighbouring anchor or an older export.
    insertRange.HighlightColorIndex = 0
    wordDoc.Bookmarks.Add bookmarkName, insertRange
    If VBA.Len(metadataBookmarkName) > 0 Then
        ' PEM-закладка не добавляет текст: она охватывает первый уже существующий
        ' символ пункта. Ненулевая длина стабильно сохраняется самим Word.
        Set metadataRange = wordDoc.Range(insertedStart, insertedStart + 1)
        wordDoc.Bookmarks.Add metadataBookmarkName, metadataRange
    End If

    ' Word расширяет предыдущую PEB-закладку, когда новый текст вставляется
    ' непосредственно в её End. После каждой вставки восстанавливаем непересекающиеся
    ' границы всех пунктов секции. End-якорь ищем повторно: ранее полученный
    ' Word Range не обязан автоматически сдвинуться после вставки.
    Set endSearchRange = wordDoc.Range(beginRange.End, wordDoc.Content.End)
    Set endRange = Nothing
    If Not private_TryFindWordText(endSearchRange, endMarker, endRange) Then GoTo CleanFail
    If Not private_TryNormalizeSectionRecordBookmarks( _
        wordDoc, templateId, beginRange.End, endRange.Start) Then GoTo CleanFail

    wordDoc.Save
    wordDoc.Close False
    documentOpened = False
    private_TryAppendBeforeWordEndAnchor = True
    Exit Function

CleanFail:
    On Error Resume Next
    If documentOpened Then wordDoc.Close False
    On Error GoTo 0
    Exit Function
EH:
    errorDescription = Err.Description
    On Error Resume Next
    If documentOpened Then wordDoc.Close False
    On Error GoTo 0
    VBA.MsgBox "PrototypeNew: WORD export failed: " & errorDescription, VBA.vbExclamation, "PrototypeNew / WORD export"
End Function

' Перегруппировывает только уже записанный *_result документ. Секционные
' текстовые якоря обязательны: они не позволяют смешать пункты разных приказов.
Public Function RegroupResultDocumentHospitalPoints(ByRef regroupedPointCount As Long) As Boolean
    Dim templatePath As String
    Dim targetPath As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim sectionPointCount As Long
    Dim documentOpened As Boolean
    Dim errorDescription As String

    On Error GoTo EH
    regroupedPointCount = 0

    templatePath = VBA.Trim$(m_Base.TargetWorkbookPath)
    If VBA.Len(templatePath) = 0 Then
        VBA.MsgBox "PrototypeNew: required profile key 'Export.Word.FilePath' is empty.", _
            VBA.vbExclamation, "PrototypeNew / WORD document"
        Exit Function
    End If
    If Not private_IsAbsolutePath(templatePath) Then
        templatePath = ThisWorkbook.Path & Application.PathSeparator & templatePath
    End If
    targetPath = private_BuildResultDocumentPath(templatePath)
    If VBA.Len(targetPath) = 0 Or _
        VBA.Len(VBA.Dir$(targetPath, VBA.vbNormal Or VBA.vbReadOnly Or VBA.vbHidden Or VBA.vbSystem)) = 0 Then
        VBA.MsgBox "PrototypeNew: WORD result document was not found:" & VBA.vbCrLf & targetPath, _
            VBA.vbExclamation, "PrototypeNew / WORD document"
        Exit Function
    End If
    If private_IsDocumentOpenInRunningWord(targetPath) Or private_IsFileLocked(targetPath) Then
        VBA.MsgBox "Невозможно перегруппировать пункты, пока документ открыт." & VBA.vbCrLf & _
            "Закройте документ и повторите попытку:" & VBA.vbCrLf & targetPath, _
            VBA.vbExclamation, "PrototypeNew / WORD document"
        Exit Function
    End If

    If Not rt_PEB_WordExportRuntime.fn_GetOrCreateWordApp(wordApp) Then Exit Function
    Set wordDoc = wordApp.Documents.Open(targetPath, False, False, False)
    documentOpened = True

    If Not private_TryRegroupHospitalSection( _
        wordApp, wordDoc, WORD_TEMPLATE_TO_HOSPITAL, sectionPointCount) Then GoTo CleanFail
    regroupedPointCount = regroupedPointCount + sectionPointCount
    If Not private_TryRegroupHospitalSection( _
        wordApp, wordDoc, WORD_TEMPLATE_FROM_HOSPITAL, sectionPointCount) Then GoTo CleanFail
    regroupedPointCount = regroupedPointCount + sectionPointCount

    wordDoc.Save
    wordDoc.Close False
    documentOpened = False
    RegroupResultDocumentHospitalPoints = True
    Exit Function

CleanFail:
    On Error Resume Next
    If documentOpened Then wordDoc.Close False
    On Error GoTo 0
    Exit Function
EH:
    errorDescription = Err.Description
    On Error Resume Next
    If documentOpened Then wordDoc.Close False
    On Error GoTo 0
    VBA.MsgBox "PrototypeNew: WORD grouping failed: " & errorDescription, _
        VBA.vbExclamation, "PrototypeNew / WORD document"
End Function

Private Function private_TryRegroupHospitalSection( _
    ByVal wordApp As Object, _
    ByVal wordDoc As Object, _
    ByVal templateId As String, _
    ByRef outPointCount As Long _
) As Boolean
    Dim beginMarker As String
    Dim endMarker As String
    Dim beginRange As Object
    Dim endSearchRange As Object
    Dim endRange As Object
    Dim bookmarkObj As Object
    Dim bookmarkName As String
    Dim metadataBookmarkName As String
    Dim recordPrefix As String
    Dim hospitalShortText As String
    Dim dateShortText As String
    Dim entries() As Object
    Dim entry As Object
    Dim swapEntry As Object
    Dim entryCount As Long
    Dim i As Long
    Dim j As Long
    Dim blockStart As Long
    Dim blockEnd As Long
    Dim recordRange As Object
    Dim metadataRange As Object
    Dim scratchDoc As Object
    Dim scratchRange As Object
    Dim scratchContent As Object
    Dim targetRange As Object
    Dim targetStart As Long
    Dim currentStart As Long
    Dim errorDescription As String

    On Error GoTo EH
    outPointCount = 0
    If wordApp Is Nothing Or wordDoc Is Nothing Then Exit Function

    beginMarker = WORD_ANCHOR_PREFIX & templateId & WORD_ANCHOR_BEGIN_SUFFIX
    endMarker = WORD_ANCHOR_PREFIX & templateId & WORD_ANCHOR_END_SUFFIX
    If Not private_TryFindWordText(wordDoc.Content, beginMarker, beginRange) Then
        VBA.MsgBox "PrototypeNew: WORD begin anchor was not found: " & beginMarker, _
            VBA.vbExclamation, "PrototypeNew / WORD document"
        Exit Function
    End If
    Set endSearchRange = wordDoc.Range(beginRange.End, wordDoc.Content.End)
    If Not private_TryFindWordText(endSearchRange, endMarker, endRange) Then
        VBA.MsgBox "PrototypeNew: WORD end anchor was not found after begin anchor: " & endMarker, _
            VBA.vbExclamation, "PrototypeNew / WORD document"
        Exit Function
    End If

    ' Исправляет в том числе документы, созданные до появления нормализации:
    ' старые PEB-закладки могли растянуться до конца всей секции.
    If Not private_TryNormalizeSectionRecordBookmarks( _
        wordDoc, templateId, beginRange.End, endRange.Start) Then Exit Function

    ' Собираем только PEB-закладки нужного templateId, расположенные строго
    ' между Begin/End текущей секции и имеющие связанную PEM-закладку.
    recordPrefix = VBA.LCase$(WORD_RECORD_BOOKMARK_PREFIX & _
        private_NormalizeBookmarkPart(templateId) & "_")
    For Each bookmarkObj In wordDoc.Bookmarks
        bookmarkName = VBA.CStr(bookmarkObj.Name)
        If VBA.Left$(VBA.LCase$(bookmarkName), VBA.Len(recordPrefix)) <> recordPrefix Then GoTo ContinueBookmark
        If bookmarkObj.Range.Start < beginRange.End Or bookmarkObj.Range.End > endRange.Start Then GoTo ContinueBookmark

        Set recordRange = bookmarkObj.Range
        If Not private_TryFindGroupingMetadataBookmark( _
            wordDoc, recordRange, metadataBookmarkName, hospitalShortText, dateShortText) Then _
            GoTo ContinueBookmark
        Set metadataRange = wordDoc.Bookmarks(metadataBookmarkName).Range

        entryCount = entryCount + 1
        ReDim Preserve entries(1 To entryCount)
        Set entry = VBA.CreateObject("Scripting.Dictionary")
        entry.CompareMode = 1
        entry("BookmarkName") = bookmarkName
        entry("MetadataBookmarkName") = metadataBookmarkName
        entry("SortKey") = private_BuildGroupingSortKey(dateShortText, hospitalShortText)
        entry("OriginalIndex") = entryCount
        entry("Length") = bookmarkObj.Range.End - bookmarkObj.Range.Start
        entry("Start") = bookmarkObj.Range.Start
        entry("End") = bookmarkObj.Range.End
        Set entries(entryCount) = entry

        If blockStart = 0 Or bookmarkObj.Range.Start < blockStart Then blockStart = bookmarkObj.Range.Start
        If bookmarkObj.Range.End > blockEnd Then blockEnd = bookmarkObj.Range.End
ContinueBookmark:
    Next bookmarkObj

    If entryCount <= 1 Then
        outPointCount = entryCount
        private_TryRegroupHospitalSection = True
        Exit Function
    End If

    ' Стабильная сортировка: одинаковые ключи не меняются местами. Ключ строится
    ' как дата YYYYMMDD, затем короткое название учреждения.
    For i = 1 To entryCount - 1
        For j = i + 1 To entryCount
            If VBA.StrComp(VBA.CStr(entries(j)("SortKey")), _
                VBA.CStr(entries(i)("SortKey")), VBA.vbTextCompare) < 0 Then
                Set swapEntry = entries(i)
                Set entries(i) = entries(j)
                Set entries(j) = swapEntry
            End If
        Next j
    Next i

    ' FormattedText нельзя надёжно хранить как обычную строку. Собираем пункты
    ' в нужном порядке во временном документе, сохраняя стили и разметку Word.
    Set scratchDoc = wordApp.Documents.Add
    For i = 1 To entryCount
        Set recordRange = wordDoc.Bookmarks(VBA.CStr(entries(i)("BookmarkName"))).Range.Duplicate
        Set scratchRange = scratchDoc.Range(scratchDoc.Content.End - 1, scratchDoc.Content.End - 1)
        scratchRange.FormattedText = recordRange.FormattedText
    Next i
    Set scratchContent = scratchDoc.Range(0, scratchDoc.Content.End - 1)

    ' Старые закладки удаляем до замены диапазона, иначе Word может оставить
    ' их границы на прежних позициях или автоматически удалить непредсказуемо.
    For i = 1 To entryCount
        bookmarkName = VBA.CStr(entries(i)("BookmarkName"))
        metadataBookmarkName = VBA.CStr(entries(i)("MetadataBookmarkName"))
        If wordDoc.Bookmarks.Exists(metadataBookmarkName) Then wordDoc.Bookmarks(metadataBookmarkName).Delete
        If wordDoc.Bookmarks.Exists(bookmarkName) Then wordDoc.Bookmarks(bookmarkName).Delete
    Next i

    targetStart = blockStart
    Set targetRange = wordDoc.Range(blockStart, blockEnd)
    targetRange.FormattedText = scratchContent.FormattedText
    scratchDoc.Close False
    Set scratchDoc = Nothing

    ' После вставки FormattedText восстанавливаем обе закладки по сохранённым
    ' длинам каждого пункта и снова скрываем диапазон метаданных.
    currentStart = targetStart
    For i = 1 To entryCount
        Set recordRange = wordDoc.Range( _
            currentStart, currentStart + VBA.CLng(entries(i)("Length")))
        wordDoc.Bookmarks.Add VBA.CStr(entries(i)("BookmarkName")), recordRange
        Set metadataRange = wordDoc.Range(currentStart, currentStart + 1)
        wordDoc.Bookmarks.Add VBA.CStr(entries(i)("MetadataBookmarkName")), metadataRange
        currentStart = currentStart + VBA.CLng(entries(i)("Length"))
    Next i

    outPointCount = entryCount
    private_TryRegroupHospitalSection = True
    Exit Function

EH:
    errorDescription = Err.Description
    On Error Resume Next
    If Not scratchDoc Is Nothing Then scratchDoc.Close False
    On Error GoTo 0
    VBA.MsgBox "PrototypeNew: failed to regroup WORD section '" & templateId & _
        "': " & errorDescription, VBA.vbExclamation, "PrototypeNew / WORD document"
End Function

Private Function private_TryNormalizeSectionRecordBookmarks( _
    ByVal wordDoc As Object, _
    ByVal templateId As String, _
    ByVal sectionStart As Long, _
    ByVal sectionEnd As Long _
) As Boolean
    Dim recordPrefix As String
    Dim bookmarkObj As Object
    Dim bookmarkName As String
    Dim bookmarkNames() As String
    Dim bookmarkStarts() As Long
    Dim bookmarkCount As Long
    Dim i As Long
    Dim j As Long
    Dim swapName As String
    Dim swapStart As Long
    Dim normalizedEnd As Long
    Dim normalizedRange As Object

    If wordDoc Is Nothing Then Exit Function
    If sectionEnd <= sectionStart Then Exit Function
    recordPrefix = VBA.LCase$(WORD_RECORD_BOOKMARK_PREFIX & _
        private_NormalizeBookmarkPart(templateId) & "_")

    ' Для восстановления достаточно надёжных левых границ. Word растягивает End,
    ' но Start каждой ранее созданной записи сохраняет корректно.
    For Each bookmarkObj In wordDoc.Bookmarks
        bookmarkName = VBA.CStr(bookmarkObj.Name)
        If VBA.Left$(VBA.LCase$(bookmarkName), VBA.Len(recordPrefix)) <> recordPrefix Then _
            GoTo ContinueBookmark
        If bookmarkObj.Range.Start < sectionStart Or bookmarkObj.Range.Start >= sectionEnd Then _
            GoTo ContinueBookmark
        bookmarkCount = bookmarkCount + 1
        ReDim Preserve bookmarkNames(1 To bookmarkCount)
        ReDim Preserve bookmarkStarts(1 To bookmarkCount)
        bookmarkNames(bookmarkCount) = bookmarkName
        bookmarkStarts(bookmarkCount) = bookmarkObj.Range.Start
ContinueBookmark:
    Next bookmarkObj

    For i = 1 To bookmarkCount - 1
        For j = i + 1 To bookmarkCount
            If bookmarkStarts(j) < bookmarkStarts(i) Then
                swapStart = bookmarkStarts(i)
                bookmarkStarts(i) = bookmarkStarts(j)
                bookmarkStarts(j) = swapStart
                swapName = bookmarkNames(i)
                bookmarkNames(i) = bookmarkNames(j)
                bookmarkNames(j) = swapName
            End If
        Next j
    Next i

    ' Сначала удаляем только внешние PEB-закладки. Вложенные PEM-закладки
    ' сохраняются, поскольку их диапазоны и имена не изменяются.
    For i = 1 To bookmarkCount
        If wordDoc.Bookmarks.Exists(bookmarkNames(i)) Then
            wordDoc.Bookmarks(bookmarkNames(i)).Delete
        End If
    Next i

    For i = 1 To bookmarkCount
        If i < bookmarkCount Then
            normalizedEnd = bookmarkStarts(i + 1)
        Else
            normalizedEnd = sectionEnd
        End If
        If normalizedEnd <= bookmarkStarts(i) Then Exit Function
        Set normalizedRange = wordDoc.Range(bookmarkStarts(i), normalizedEnd)
        wordDoc.Bookmarks.Add bookmarkNames(i), normalizedRange
    Next i

    private_TryNormalizeSectionRecordBookmarks = True
End Function

' Завершает подготовку всех сформированных блоков результирующего документа:
' сохраняет экспортированный текст, но удаляет маркеры
' {\export:<id>_Begin}/{\export:<id>_End}, а также примыкающие к содержимому
' переводы строк CR/LF и ручные разрывы строк. Идентификаторы читаются из самого
' документа, поэтому новые секции не требуют изменений VBA. Операция необратима
' для текущего result-файла: для дальнейшего экспорта потребуется новый результат.
Public Function RemoveResultDocumentAnchors(ByRef clearedBlockCount As Long) As Boolean
    Dim templatePath As String
    Dim targetPath As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim documentText As String
    Dim beginMatches As Object
    Dim beginMatch As Object
    Dim rx As Object
    Dim beginRange As Object
    Dim endSearchRange As Object
    Dim endRange As Object
    Dim boundaryRange As Object
    Dim blockText As String
    Dim leadingBreakCount As Long
    Dim trailingBreakCount As Long
    Dim anchorId As String
    Dim beginMarker As String
    Dim endMarker As String
    Dim matchIndex As Long
    Dim documentOpened As Boolean
    Dim errorDescription As String
    Dim bookmarkIndex As Long
    Dim bookmarkObj As Object

    On Error GoTo EH
    clearedBlockCount = 0

    templatePath = VBA.Trim$(m_Base.TargetWorkbookPath)
    If VBA.Len(templatePath) = 0 Then
        VBA.MsgBox "PrototypeNew: required profile key 'Export.Word.FilePath' is empty.", VBA.vbExclamation, "PrototypeNew / WORD document"
        Exit Function
    End If
    If Not private_IsAbsolutePath(templatePath) Then templatePath = ThisWorkbook.Path & Application.PathSeparator & templatePath

    targetPath = private_BuildResultDocumentPath(templatePath)
    If VBA.Len(targetPath) = 0 Or VBA.Len(VBA.Dir$(targetPath, VBA.vbNormal Or VBA.vbReadOnly Or VBA.vbHidden Or VBA.vbSystem)) = 0 Then
        VBA.MsgBox "PrototypeNew: WORD result document was not found:" & VBA.vbCrLf & targetPath, VBA.vbExclamation, "PrototypeNew / WORD document"
        Exit Function
    End If

    If private_IsDocumentOpenInRunningWord(targetPath) Or private_IsFileLocked(targetPath) Then
        VBA.MsgBox "Невозможно очистить документ, пока он открыт." & VBA.vbCrLf & _
            "Закройте документ и повторите попытку:" & VBA.vbCrLf & targetPath, _
            VBA.vbExclamation, "PrototypeNew / WORD document"
        Exit Function
    End If

    If Not rt_PEB_WordExportRuntime.fn_GetOrCreateWordApp(wordApp) Then Exit Function
    Set wordDoc = wordApp.Documents.Open(targetPath, False, False, False)
    documentOpened = True

    documentText = VBA.CStr(wordDoc.Content.Text)
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = True
    rx.Pattern = "\{\\export:([^}\r\n]+)_Begin\}"
    Set beginMatches = rx.Execute(documentText)

    ' Идём с конца документа, чтобы удаление блока не смещало сохранённые
    ' позиции начальных якорей предыдущих блоков.
    For matchIndex = beginMatches.Count - 1 To 0 Step -1
        Set beginMatch = beginMatches(matchIndex)
        anchorId = VBA.CStr(beginMatch.SubMatches(0))
        beginMarker = VBA.CStr(beginMatch.Value)
        endMarker = WORD_ANCHOR_PREFIX & anchorId & WORD_ANCHOR_END_SUFFIX

        ' FirstIndex относится к строке Content.Text и может расходиться с
        ' координатами Word Range из-за полей и служебных символов документа.
        ' Поэтому regex определяет только текст якоря, а фактический диапазон
        ' повторно находим средствами Word — так завершающая "}" не остаётся.
        If Not private_TryFindWordText(wordDoc.Content, beginMarker, beginRange) Then
            VBA.MsgBox "PrototypeNew: WORD begin anchor was not found: " & beginMarker, VBA.vbExclamation, "PrototypeNew / WORD document"
            GoTo CleanFail
        End If
        Set endSearchRange = wordDoc.Range(beginRange.End, wordDoc.Content.End)
        If Not private_TryFindWordText(endSearchRange, endMarker, endRange) Then
            VBA.MsgBox "PrototypeNew: WORD end anchor was not found after begin anchor: " & endMarker, VBA.vbExclamation, "PrototypeNew / WORD document"
            GoTo CleanFail
        End If

        blockText = VBA.CStr(wordDoc.Range(beginRange.End, endRange.Start).Text)
        leadingBreakCount = private_CountLeadingWordLineBreaks(blockText)
        trailingBreakCount = private_CountTrailingWordLineBreaks(blockText)
        ' В пустом блоке одни и те же переводы строк одновременно являются
        ' начальными и хвостовыми. Относим пересечение к началу блока, чтобы
        ' не удалить текст за его границами.
        If leadingBreakCount + trailingBreakCount > VBA.Len(blockText) Then
            trailingBreakCount = VBA.Len(blockText) - leadingBreakCount
        End If

        ' Внутри блока тоже идём справа налево. Так сохраняются текст,
        ' форматирование и закладки записей между двумя якорями.
        endRange.Text = VBA.vbNullString
        If trailingBreakCount > 0 Then
            Set boundaryRange = wordDoc.Range(endRange.Start - trailingBreakCount, endRange.Start)
            boundaryRange.Text = VBA.vbNullString
        End If
        If leadingBreakCount > 0 Then
            Set boundaryRange = wordDoc.Range(beginRange.End, beginRange.End + leadingBreakCount)
            boundaryRange.Text = VBA.vbNullString
        End If
        beginRange.Text = VBA.vbNullString
        clearedBlockCount = clearedBlockCount + 1
    Next matchIndex

    ' Финализация удаляет и скрытые данные сортировки. После этого документ
    ' остаётся визуально чистым, а повторная перегруппировка намеренно невозможна.
    For bookmarkIndex = wordDoc.Bookmarks.Count To 1 Step -1
        Set bookmarkObj = wordDoc.Bookmarks.Item(bookmarkIndex)
        If VBA.Left$(VBA.CStr(bookmarkObj.Name), VBA.Len(WORD_METADATA_BOOKMARK_PREFIX)) = _
            WORD_METADATA_BOOKMARK_PREFIX Then
            bookmarkObj.Delete
        End If
    Next bookmarkIndex

    wordDoc.Save
    wordDoc.Close False
    documentOpened = False
    RemoveResultDocumentAnchors = True
    Exit Function

CleanFail:
    On Error Resume Next
    If documentOpened Then wordDoc.Close False
    On Error GoTo 0
    Exit Function
EH:
    errorDescription = Err.Description
    On Error Resume Next
    If documentOpened Then wordDoc.Close False
    On Error GoTo 0
    VBA.MsgBox "PrototypeNew: WORD document cleanup failed: " & errorDescription, VBA.vbExclamation, "PrototypeNew / WORD document"
End Function

Private Function private_CountLeadingWordLineBreaks(ByVal valueText As String) As Long
    Dim charIndex As Long
    Dim currentChar As String

    For charIndex = 1 To VBA.Len(valueText)
        currentChar = VBA.Mid$(valueText, charIndex, 1)
        If currentChar <> VBA.vbCr And currentChar <> VBA.vbLf And currentChar <> VBA.Chr$(11) Then Exit For
        private_CountLeadingWordLineBreaks = private_CountLeadingWordLineBreaks + 1
    Next charIndex
End Function

Private Function private_CountTrailingWordLineBreaks(ByVal valueText As String) As Long
    Dim charIndex As Long
    Dim currentChar As String

    For charIndex = VBA.Len(valueText) To 1 Step -1
        currentChar = VBA.Mid$(valueText, charIndex, 1)
        If currentChar <> VBA.vbCr And currentChar <> VBA.vbLf And currentChar <> VBA.Chr$(11) Then Exit For
        private_CountTrailingWordLineBreaks = private_CountTrailingWordLineBreaks + 1
    Next charIndex
End Function

Private Function private_BuildResultDocumentPath(ByVal templatePath As String) As String
    Dim extensionPos As Long
    Dim slashPos As Long
    Dim backslashPos As Long

    templatePath = VBA.Trim$(templatePath)
    If VBA.Len(templatePath) = 0 Then Exit Function

    slashPos = VBA.InStrRev(templatePath, "/")
    backslashPos = VBA.InStrRev(templatePath, "\")
    extensionPos = VBA.InStrRev(templatePath, ".")

    If extensionPos > slashPos And extensionPos > backslashPos Then
        private_BuildResultDocumentPath = VBA.Left$(templatePath, extensionPos - 1) & _
            "_result" & VBA.Mid$(templatePath, extensionPos)
    Else
        private_BuildResultDocumentPath = templatePath & "_result.docx"
    End If
End Function

Private Function private_BuildRecordBookmarkName(ByVal templateId As String, ByVal recordIpn As String) As String
    Dim safeTemplateId As String
    Dim safeIpn As String
    Dim availableTemplateLength As Long

    safeTemplateId = private_NormalizeBookmarkPart(templateId)
    safeIpn = private_NormalizeBookmarkPart(recordIpn)
    If VBA.Len(safeTemplateId) = 0 Or VBA.Len(safeIpn) = 0 Then Exit Function

    availableTemplateLength = WORD_BOOKMARK_MAX_LENGTH - VBA.Len(WORD_RECORD_BOOKMARK_PREFIX) - VBA.Len(safeIpn) - 1
    If availableTemplateLength <= 0 Then Exit Function
    If VBA.Len(safeTemplateId) > availableTemplateLength Then safeTemplateId = VBA.Left$(safeTemplateId, availableTemplateLength)

    private_BuildRecordBookmarkName = WORD_RECORD_BOOKMARK_PREFIX & safeTemplateId & "_" & safeIpn
End Function

Private Function private_BuildGroupingMetadataBookmarkName( _
    ByVal hospitalShortText As String, _
    ByVal dateShortText As String _
) As String
    Dim sortKey As String
    Dim separatorPos As Long
    Dim dateKey As String
    Dim hospitalKey As String
    Dim availableHospitalLength As Long

    hospitalShortText = VBA.Trim$(hospitalShortText)
    dateShortText = VBA.Trim$(dateShortText)
    If VBA.Len(hospitalShortText) = 0 Or VBA.Len(dateShortText) = 0 Then Exit Function

    sortKey = private_BuildGroupingSortKey(dateShortText, hospitalShortText)
    separatorPos = VBA.InStr(1, sortKey, "|", VBA.vbBinaryCompare)
    If separatorPos <= 1 Then Exit Function
    dateKey = VBA.Left$(sortKey, separatorPos - 1)
    hospitalKey = private_NormalizeMetadataBookmarkPart(hospitalShortText)
    availableHospitalLength = WORD_BOOKMARK_MAX_LENGTH - _
        VBA.Len(WORD_METADATA_BOOKMARK_PREFIX) - VBA.Len(dateKey) - 1
    If availableHospitalLength <= 0 Or VBA.Len(hospitalKey) = 0 Then Exit Function
    If VBA.Len(hospitalKey) > availableHospitalLength Then
        hospitalKey = VBA.Left$(hospitalKey, availableHospitalLength)
    End If
    private_BuildGroupingMetadataBookmarkName = WORD_METADATA_BOOKMARK_PREFIX & _
        dateKey & "_" & hospitalKey
End Function

Private Function private_TryParseGroupingMetadataBookmarkName( _
    ByVal metadataBookmarkName As String, _
    ByRef outHospitalShortText As String, _
    ByRef outDateShortText As String _
) As Boolean
    Dim payloadText As String
    Dim separatorPos As Long
    Dim dateKey As String

    outHospitalShortText = VBA.vbNullString
    outDateShortText = VBA.vbNullString
    If VBA.Left$(metadataBookmarkName, VBA.Len(WORD_METADATA_BOOKMARK_PREFIX)) <> _
        WORD_METADATA_BOOKMARK_PREFIX Then Exit Function

    payloadText = VBA.Mid$(metadataBookmarkName, VBA.Len(WORD_METADATA_BOOKMARK_PREFIX) + 1)
    separatorPos = VBA.InStr(1, payloadText, "_", VBA.vbBinaryCompare)
    If separatorPos <= 1 Then Exit Function
    dateKey = VBA.Left$(payloadText, separatorPos - 1)
    If VBA.Len(dateKey) <> 8 Or Not VBA.IsNumeric(dateKey) Then Exit Function
    outDateShortText = VBA.Mid$(dateKey, 7, 2) & "." & _
        VBA.Mid$(dateKey, 5, 2) & "." & VBA.Left$(dateKey, 4)
    outHospitalShortText = VBA.Trim$(VBA.Mid$(payloadText, separatorPos + 1))
    private_TryParseGroupingMetadataBookmarkName = ( _
        VBA.Len(outDateShortText) > 0 And VBA.Len(outHospitalShortText) > 0)
End Function

Private Function private_TryFindGroupingMetadataBookmark( _
    ByVal wordDoc As Object, _
    ByVal recordRange As Object, _
    ByRef outMetadataBookmarkName As String, _
    ByRef outHospitalShortText As String, _
    ByRef outDateShortText As String _
) As Boolean
    Dim bookmarkObj As Object
    Dim bookmarkName As String

    outMetadataBookmarkName = VBA.vbNullString
    outHospitalShortText = VBA.vbNullString
    outDateShortText = VBA.vbNullString
    If wordDoc Is Nothing Or recordRange Is Nothing Then Exit Function

    ' После сохранения и повторного открытия Word может сдвинуть нулевую
    ' закладку на один служебный символ относительно начала PEB-диапазона.
    ' Поэтому проверяем принадлежность всему пункту. Правую границу исключаем,
    ' чтобы не принять PEM-закладку следующего соседнего пункта.
    For Each bookmarkObj In wordDoc.Bookmarks
        bookmarkName = VBA.CStr(bookmarkObj.Name)
        If VBA.Left$(bookmarkName, VBA.Len(WORD_METADATA_BOOKMARK_PREFIX)) <> _
            WORD_METADATA_BOOKMARK_PREFIX Then GoTo ContinueBookmark
        If bookmarkObj.Range.Start < recordRange.Start Then GoTo ContinueBookmark
        If bookmarkObj.Range.Start >= recordRange.End Then GoTo ContinueBookmark
        If Not private_TryParseGroupingMetadataBookmarkName( _
            bookmarkName, outHospitalShortText, outDateShortText) Then GoTo ContinueBookmark
        outMetadataBookmarkName = bookmarkName
        private_TryFindGroupingMetadataBookmark = True
        Exit Function
ContinueBookmark:
    Next bookmarkObj
End Function

Private Function private_NormalizeMetadataBookmarkPart(ByVal valueText As String) As String
    Dim resultText As String
    Dim charIndex As Long
    Dim charText As String
    Const ALLOWED_CHARS As String = _
        "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_" & _
        "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ" & _
        "абвгдеёжзийклмнопрстуфхцчшщъыьэюяІіЇїЄєҐґ"

    valueText = VBA.Trim$(valueText)
    For charIndex = 1 To VBA.Len(valueText)
        charText = VBA.Mid$(valueText, charIndex, 1)
        If VBA.InStr(1, ALLOWED_CHARS, charText, VBA.vbBinaryCompare) > 0 Then
            resultText = resultText & charText
        ElseIf VBA.Right$(resultText, 1) <> "_" Then
            resultText = resultText & "_"
        End If
    Next charIndex
    Do While VBA.Right$(resultText, 1) = "_"
        resultText = VBA.Left$(resultText, VBA.Len(resultText) - 1)
    Loop
    private_NormalizeMetadataBookmarkPart = resultText
End Function

Private Function private_BuildGroupingSortKey( _
    ByVal dateShortText As String, _
    ByVal hospitalShortText As String _
) As String
    Dim dateParts As Variant
    Dim dateKey As String

    dateShortText = VBA.Trim$(dateShortText)
    hospitalShortText = VBA.LCase$(VBA.Trim$(hospitalShortText))
    dateParts = VBA.Split(dateShortText, ".")
    If UBound(dateParts) = 2 Then
        dateKey = VBA.CStr(dateParts(2)) & VBA.Right$("0" & VBA.CStr(dateParts(1)), 2) & _
            VBA.Right$("0" & VBA.CStr(dateParts(0)), 2)
    Else
        dateKey = dateShortText
    End If
    private_BuildGroupingSortKey = dateKey & "|" & hospitalShortText
End Function

Private Function private_IsHospitalGroupingTemplate(ByVal templateId As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(templateId))
        Case VBA.LCase$(WORD_TEMPLATE_TO_HOSPITAL), _
             VBA.LCase$(WORD_TEMPLATE_FROM_HOSPITAL)
            private_IsHospitalGroupingTemplate = True
    End Select
End Function

Private Function private_BuildUniqueBookmarkName(ByVal wordDoc As Object, ByVal baseName As String) As String
    Dim candidateName As String
    Dim suffixText As String
    Dim sequenceNo As Long
    Dim baseMaxLength As Long

    baseName = VBA.Trim$(baseName)
    If wordDoc Is Nothing Or VBA.Len(baseName) = 0 Then Exit Function
    If Not wordDoc.Bookmarks.Exists(baseName) Then
        private_BuildUniqueBookmarkName = baseName
        Exit Function
    End If

    sequenceNo = 2
    Do
        suffixText = "_" & VBA.CStr(sequenceNo)
        baseMaxLength = WORD_BOOKMARK_MAX_LENGTH - VBA.Len(suffixText)
        candidateName = VBA.Left$(baseName, baseMaxLength) & suffixText
        If Not wordDoc.Bookmarks.Exists(candidateName) Then
            private_BuildUniqueBookmarkName = candidateName
            Exit Function
        End If
        sequenceNo = sequenceNo + 1
    Loop While sequenceNo < 100000
End Function

Private Function private_NormalizeBookmarkPart(ByVal valueText As String) As String
    Dim rx As Object
    valueText = VBA.Trim$(valueText)
    If VBA.Len(valueText) = 0 Then Exit Function
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.Pattern = "[^A-Za-z0-9_]"
    private_NormalizeBookmarkPart = rx.Replace(valueText, "_")
End Function

Private Function private_StripPreviewColorMarkers(ByVal renderedText As String) As String
    Dim rx As Object
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = True
    rx.Pattern = "\[\[/?color(?:=[^\]]+)?\]\]"
    private_StripPreviewColorMarkers = rx.Replace(renderedText, VBA.vbNullString)
End Function

Private Function private_TryFindWordText(ByVal sourceRange As Object, ByVal targetText As String, ByRef outRange As Object) As Boolean
    Dim findRange As Object
    Set outRange = Nothing
    If sourceRange Is Nothing Then Exit Function
    Set findRange = sourceRange.Duplicate
    With findRange.Find
        .ClearFormatting
        .Text = targetText
        .Forward = True
        .Wrap = WD_FIND_STOP
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
    End With
    If findRange.Find.Execute Then
        Set outRange = findRange.Duplicate
        private_TryFindWordText = True
    End If
End Function

Private Function private_IsDocumentOpenInRunningWord(ByVal targetPath As String) As Boolean
    Dim wordApp As Object
    Dim doc As Object

    On Error Resume Next
    Set wordApp = VBA.GetObject(, "Word.Application")
    On Error GoTo 0
    If wordApp Is Nothing Then Exit Function

    For Each doc In wordApp.Documents
        If VBA.StrComp(VBA.CStr(doc.FullName), targetPath, VBA.vbTextCompare) = 0 Then
            private_IsDocumentOpenInRunningWord = True
            Exit Function
        End If
    Next doc
End Function

Private Function private_IsFileLocked(ByVal targetPath As String) As Boolean
    Dim fileHandle As Integer

    On Error GoTo Locked
    fileHandle = VBA.FreeFile
    Open targetPath For Binary Access Read Write Lock Read Write As #fileHandle
    Close #fileHandle
    private_IsFileLocked = False
    Exit Function

Locked:
    private_IsFileLocked = True
    On Error Resume Next
    If fileHandle > 0 Then Close #fileHandle
    On Error GoTo 0
End Function

Private Function private_IsAbsolutePath(ByVal pathText As String) As Boolean
    pathText = VBA.Trim$(pathText)
    private_IsAbsolutePath = (VBA.Len(pathText) >= 3 And VBA.Mid$(pathText, 2, 2) = ":\") _
        Or (VBA.Left$(pathText, 2) = "\\")
End Function

Private Function private_GetContextBoolean(ByVal context As Object, ByVal keyText As String) As Boolean
    Dim rawValue As Variant
    If context Is Nothing Then Exit Function
    On Error Resume Next
    If context.Exists(keyText) Then rawValue = context(keyText)
    If Err.Number <> 0 Then
        Err.Clear
        rawValue = VBA.CallByName(context, keyText, VbGet)
    End If
    On Error GoTo 0
    If VBA.VarType(rawValue) = VBA.vbBoolean Then
        private_GetContextBoolean = VBA.CBool(rawValue)
    Else
        private_GetContextBoolean = (VBA.StrComp(VBA.Trim$(VBA.CStr(rawValue)), "True", VBA.vbTextCompare) = 0)
    End If
End Function

' //
' // Internal
' //
Private Function private_TryEnrichMainSourceTableForWord( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal context As Object _
) As Boolean
    Dim ipnText As String
    Dim rankText As String
    Dim fioText As String
    Dim positionCodeText As String
    Dim positionText As String
    Dim hospitalShortText As String
    Dim vacationText As String
    Dim reportRankText As String
    Dim reportPersonText As String
    Dim reportPositionCodeText As String
    Dim incomingNoText As String
    Dim incomingDateText As String
    Dim docDateText As String
    Dim dateFromText As String
    Dim vhDateText As String
    Dim vlkDateText As String
    Dim orderNoText As String
    Dim sectionTypeText As String
    Dim durationDaysText As String
    Dim rankGenitive As String
    Dim fioGenitive As String
    Dim fioAccusative As String
    Dim fioInitialsGenitive As String
    Dim positionGenitive As String
    Dim hospitalGenitive As String
    Dim hospitalAccusative As String
    Dim hospitalDative As String
    Dim reportRankGenitive As String
    Dim reportPersonGenitive As String
    Dim reportPersonInitialsGenitive As String
    Dim reportPositionGenitive As String
    Dim reportTvoPositionGenitive As String
    Dim reporterGenitive As String
    Dim isReporterTvo As Boolean
    Dim dateFromDate As Date
    Dim hasDateFrom As Boolean
    Dim removeFromFoodSupportDateText As String
    Dim enrollToFoodSupportDateText As String
    Dim foodSupportDateValue As Date
    Dim vacationTotalDays As Long
    Dim vacationDays As Long
    Dim additionalWayDays As Long
    Dim vacationDateTo As Date
    Dim builderData As obj_PrsnlEvntBuilderData
    If sourceTable Is Nothing Then Exit Function
    If m_ExporterDataProvider Is Nothing Then Exit Function
    If m_ExporterDataProvider.CommonData Is Nothing Then Exit Function

    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_IPN, ipnText) Then ipnText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_RANK, rankText) Then rankText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_FIO, fioText) Then fioText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_POSITION_CODE, positionCodeText) Then positionCodeText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_POSITION_NAME, positionText) Then positionText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_HOSPITAL_SHORT, hospitalShortText) Then hospitalShortText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_VACATION, vacationText) Then vacationText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_REPORT_RANK, reportRankText) Then reportRankText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_REPORT_PERSON, reportPersonText) Then reportPersonText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_REPORT_POSITION_CODE, reportPositionCodeText) Then reportPositionCodeText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_INCOMING_NO, incomingNoText) Then incomingNoText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_INCOMING_DATE, incomingDateText) Then incomingDateText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_DOC_DATE, docDateText) Then docDateText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_DATE_FROM, dateFromText) Then dateFromText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_DURATION_DAYS, durationDaysText) Then durationDaysText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_VH_DATE, vhDateText) Then vhDateText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_VLK_DATE, vlkDateText) Then vlkDateText = VBA.vbNullString

    ' WORD preview работает не только с исходными колонками формы. Перед
    ' render-ом мы дописываем в main DynamicTable вычисленные поля:
    ' склонения ФИО/звания/посады/лечебного учреждения, строку рапортующего
    ' и date aliases. Все даты передаём в едином компактном виде
    ' ...Short = 01.02.2025; полный вид формирует XML через dateformat.
    orderNoText = private_GetContextText(context, CONTEXT_MANUAL_ORDER_NO)
    If Not m_ExporterDataProvider.CommonData.SetOrderNo(orderNoText) Then Exit Function
    If VBA.Len(VBA.Trim$(orderNoText)) > 0 Then
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_ORDER_NO, orderNoText) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(orderNoText)) > 0 And Not m_ExporterDataProvider.CommonData.HasOrderDate Then
        rt_Messaging.fn_ShowStatusBarWarning _
            "Order date was not found for order number '" & orderNoText & "'. Short dates use 01.01.1900.", _
            5
    End If

    ' Все склонения берутся из общего provider-а. Если справочник пустой или
    ' ключ не найден, provider сам показывает MsgBox с конкретной причиной.
    If Not m_ExporterDataProvider.CommonData.TryResolveRankGenitive(rankText, rankGenitive) Then Exit Function
    If Not m_ExporterDataProvider.CommonData.TryResolveFioGenitive(ipnText, fioGenitive) Then Exit Function
    If Not m_ExporterDataProvider.CommonData.TryResolveFioAccusative(ipnText, fioAccusative) Then Exit Function
    If Not m_ExporterDataProvider.CommonData.TryResolveFioInitialsGenitive(ipnText, fioInitialsGenitive) Then Exit Function
    If Not m_ExporterDataProvider.CommonData.TryResolvePositionGenitive(positionCodeText, positionGenitive) Then Exit Function
    If Not m_ExporterDataProvider.CommonData.TryResolveHospitalGenitive(hospitalShortText, hospitalGenitive) Then Exit Function
    If Not m_ExporterDataProvider.CommonData.TryResolveHospitalAccusative(hospitalShortText, hospitalAccusative) Then Exit Function
    If Not m_ExporterDataProvider.CommonData.TryResolveHospitalDative(hospitalShortText, hospitalDative) Then Exit Function
    If Not m_ExporterDataProvider.CommonData.TryResolveRankGenitive(reportRankText, reportRankGenitive) Then Exit Function
    If Not m_ExporterDataProvider.CommonData.TryResolveFioGenitiveByName(reportPersonText, reportPersonGenitive) Then Exit Function
    If Not m_ExporterDataProvider.CommonData.TryResolveFioInitialsGenitiveByName(reportPersonText, reportPersonInitialsGenitive) Then Exit Function

    If Not m_ExporterDataProvider.TryResolveReporterTvoPositionGenitive(reportPersonText, reportTvoPositionGenitive, isReporterTvo) Then Exit Function
    If isReporterTvo Then
        reportPositionGenitive = reportTvoPositionGenitive
    Else
        If Not m_ExporterDataProvider.CommonData.TryResolvePositionGenitive(reportPositionCodeText, reportPositionGenitive) Then Exit Function
    End If

    If isReporterTvo Then
        reporterGenitive = private_JoinNonEmptyParts( _
            private_JoinNonEmptyParts(REPORT_TVO_TEXT & " " & private_LowerFirstLetter(reportPositionGenitive), reportRankGenitive), _
            reportPersonInitialsGenitive)
    Else
        reporterGenitive = private_JoinNonEmptyParts( _
            private_JoinNonEmptyParts(private_LowerFirstLetter(reportPositionGenitive), reportRankGenitive), _
            reportPersonInitialsGenitive)
    End If

    ' Upsert не влияет на видимую форму экспорта: это служебное обогащение
    ' DynamicTable перед шаблонизацией. XML-шаблон может читать новые поля,
    ' но UI не обязан их отрисовывать отдельными колонками.
    If VBA.Len(VBA.Trim$(rankGenitive)) > 0 Then
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_RANK_GENITIVE, rankGenitive) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(fioGenitive)) > 0 Then
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_FIO_GENITIVE, fioGenitive) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(fioAccusative)) > 0 Then
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_FIO_ACCUSATIVE, fioAccusative) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(fioInitialsGenitive)) > 0 Then
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_FIO_INITIALS_GENITIVE, fioInitialsGenitive) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(positionGenitive)) > 0 Then
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_POSITION_GENITIVE, positionGenitive) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(positionText)) > 0 Then
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_POSITION_DEFAULT, positionText) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(hospitalGenitive)) > 0 Then
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_HOSPITAL_GENITIVE, hospitalGenitive) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(hospitalAccusative)) > 0 Then
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_HOSPITAL_ACCUSATIVE, hospitalAccusative) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(hospitalDative)) > 0 Then
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_HOSPITAL_DATIVE, hospitalDative) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(vacationText)) > 0 Then
        If Not private_TryUpsertMainTableValue(sourceTable, SOURCE_ALIAS_VACATION, vacationText) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(reportRankGenitive)) > 0 Then
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_REPORT_RANK_GENITIVE, reportRankGenitive) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(reportPersonGenitive)) > 0 Then
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_REPORT_PERSON_GENITIVE, reportPersonGenitive) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(reportPersonInitialsGenitive)) > 0 Then
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_REPORT_PERSON_INITIALS_GENITIVE, reportPersonInitialsGenitive) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(reportPositionGenitive)) > 0 Then
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_REPORT_POSITION_GENITIVE, reportPositionGenitive) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(reporterGenitive)) > 0 Then
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_REPORTER_GENITIVE, reporterGenitive) Then Exit Function
    End If
    If VBA.Len(incomingNoText) > 0 Then
        If Not private_TryUpsertMainTableValue( _
            sourceTable, SOURCE_ALIAS_INCOMING_NO, _
            m_ExporterDataProvider.NormalizeIncomingNoForExport(incomingNoText)) Then Exit Function
    End If
    If Not private_TryUpsertShortDateValue(sourceTable, SOURCE_ALIAS_INCOMING_DATE, WORD_ALIAS_INCOMING_DATE_SHORT, incomingDateText) Then Exit Function
    If Not private_TryUpsertShortDateValue(sourceTable, SOURCE_ALIAS_DOC_DATE, WORD_ALIAS_DOC_DATE_SHORT, docDateText) Then Exit Function
    If Not private_TryUpsertShortDateValue(sourceTable, SOURCE_ALIAS_DATE_FROM, WORD_ALIAS_DATE_FROM_SHORT, dateFromText) Then Exit Function
    If Not private_TryUpsertShortDateValue(sourceTable, SOURCE_ALIAS_VH_DATE, WORD_ALIAS_VH_DATE_SHORT, vhDateText) Then Exit Function
    If Not private_TryUpsertShortDateValue(sourceTable, SOURCE_ALIAS_VLK_DATE, WORD_ALIAS_VLK_DATE_SHORT, vlkDateText) Then Exit Function

    If Not private_TryResolveDateByRawText(dateFromText, hasDateFrom, dateFromDate) Then Exit Function

    sectionTypeText = private_GetContextText(context, CONTEXT_SECTION_TYPE)
    If VBA.Len(sectionTypeText) = 0 Then sectionTypeText = VBA.Trim$(sourceTable.SectionTitle)

    ' Продолжительность отпуска храним в контексте тремя отдельными числами.
    ' Например, "15+4 (дорога)" превращается в 15 дней отпуска, 4 дня дороги
    ' и 19 календарных дней общего периода. Исходную пользовательскую строку
    ' шаблон больше не выводит.
    If VBA.Len(VBA.Trim$(durationDaysText)) > 0 Then
        If Not private_TryParseVacationDuration( _
            durationDaysText, vacationDays, additionalWayDays, vacationTotalDays) Then Exit Function
        If Not private_TryUpsertMainTableValue( _
            sourceTable, WORD_ALIAS_VACATION_DAYS, VBA.CStr(vacationDays)) Then Exit Function
        If Not private_TryUpsertMainTableValue( _
            sourceTable, WORD_ALIAS_ADDITIONAL_WAY_DAYS, VBA.CStr(additionalWayDays)) Then Exit Function
        If Not private_TryUpsertMainTableValue( _
            sourceTable, WORD_ALIAS_VACATION_TOTAL_DAYS, VBA.CStr(vacationTotalDays)) Then Exit Function
    End If

    ' Для всех секций выбытия или перевода в отпуск вычисляем дату окончания
    ' включительно: DateFrom + VacationTotalDays - 1. Окончательное отображение
    ' даты (месяц словами, "року", NBSP) по-прежнему задаёт XML-шаблон.
    Set builderData = New obj_PrsnlEvntBuilderData
    If builderData.UsesMovementVacationDestination(sectionTypeText) Then
        If hasDateFrom Then
            If vacationTotalDays > 0 Then
                vacationDateTo = VBA.DateAdd("d", vacationTotalDays - 1, dateFromDate)
                If Not private_TryUpsertMainTableValue( _
                    sourceTable, WORD_ALIAS_DATE_TO_SHORT, VBA.Format$(vacationDateTo, WORD_SHORT_DATE_STORAGE_FORMAT)) Then Exit Function
                If Not private_TryUpsertMainTableValue( _
                    sourceTable, WORD_ALIAS_VACATION_DATES_SAME_MONTH, _
                    VBA.CStr(VBA.Year(dateFromDate) = VBA.Year(vacationDateTo) And _
                             VBA.Month(dateFromDate) = VBA.Month(vacationDateTo))) Then Exit Function
            End If
        End If
    End If

    ' WORD и Movement используют один helper: max(OrderDate + 1, DateFrom).
    If Not m_ExporterDataProvider.CommonData.TryCalculateFoodSupportDate( _
        hasDateFrom, dateFromDate, foodSupportDateValue) Then
        VBA.MsgBox "PrototypeNew: cannot calculate food-support date because both order date and DateFrom are unavailable.", _
            VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    removeFromFoodSupportDateText = VBA.Format$(foodSupportDateValue, WORD_SHORT_DATE_STORAGE_FORMAT)
    enrollToFoodSupportDateText = removeFromFoodSupportDateText

    If VBA.Len(VBA.Trim$(removeFromFoodSupportDateText)) > 0 Then
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_REMOVE_FROM_FOOD_SUPPORT_DATE, removeFromFoodSupportDateText) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(enrollToFoodSupportDateText)) > 0 Then
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_ENROLL_TO_FOOD_SUPPORT_DATE, enrollToFoodSupportDateText) Then Exit Function
    End If

    private_TryEnrichMainSourceTableForWord = True
End Function

Private Function private_TryEnrichPreviousVacationTicketForWord( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal sectionTypeText As String, _
    ByVal latestMovementRecord As Object _
) As Boolean
    Dim data As obj_PrsnlEvntBuilderData
    Dim escortDocumentText As String
    Dim departureOrderText As String
    Dim departureOrderDate As Date
    Dim previousTicketDateText As String
    Dim isSupportedSection As Boolean

    If sourceTable Is Nothing Then Exit Function

    Set data = New obj_PrsnlEvntBuilderData
    isSupportedSection = _
        (VBA.StrComp(sectionTypeText, data.SectionTypeTransferTreatmentVacationToTreatment, VBA.vbTextCompare) = 0) _
        Or (VBA.StrComp(sectionTypeText, data.SectionTypeTransferAnnualVacationToTreatment, VBA.vbTextCompare) = 0) _
        Or (VBA.StrComp(sectionTypeText, data.SectionTypeTransferTreatmentVacationToVlk, VBA.vbTextCompare) = 0)
    If Not isSupportedSection Then
        private_TryEnrichPreviousVacationTicketForWord = True
        Exit Function
    End If

    If latestMovementRecord Is Nothing Then
        VBA.MsgBox "PrototypeNew: latest Movement record is unavailable for the previous vacation ticket.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    If latestMovementRecord.Exists(LATEST_MOVEMENT_ESCORT_DOCUMENT_KEY) Then
        escortDocumentText = private_NormalizeTemplateScalar( _
            VBA.CStr(latestMovementRecord(LATEST_MOVEMENT_ESCORT_DOCUMENT_KEY)))
    End If
    If latestMovementRecord.Exists(LATEST_MOVEMENT_DEPARTURE_ORDER_KEY) Then
        departureOrderText = private_NormalizeTemplateScalar( _
            VBA.CStr(latestMovementRecord(LATEST_MOVEMENT_DEPARTURE_ORDER_KEY)))
    End If

    If VBA.Len(VBA.Trim$(escortDocumentText)) = 0 Then
        VBA.MsgBox "PrototypeNew: latest Movement record has no 'Супровідний документ' value for PrevVkNum.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(departureOrderText)) = 0 Then
        VBA.MsgBox "PrototypeNew: latest Movement record has no 'Наказ вибуття' value for PrevVkDate.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    If Not m_ExporterDataProvider.CommonData.TryResolveOrderDateByNumber(departureOrderText, departureOrderDate) Then
        VBA.MsgBox "PrototypeNew: failed to resolve PrevVkDate by departure order '" & departureOrderText & "'.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    previousTicketDateText = VBA.Format$(departureOrderDate, WORD_SHORT_DATE_STORAGE_FORMAT)
    If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_PREV_VK_NUM, escortDocumentText) Then Exit Function
    If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_PREV_VK_DATE, previousTicketDateText) Then Exit Function

    private_TryEnrichPreviousVacationTicketForWord = True
End Function

Private Function private_TryNormalizeDocumentNotesForWord( _
    ByVal sourceTables As Collection, _
    ByVal sectionTypeText As String _
) As Boolean
    Dim data As obj_PrsnlEvntBuilderData
    Dim sourceTable As obj_TableDynamic
    Dim sourceRow As obj_Row
    Dim tableValue As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim documentNoteText As String
    Dim normalizedDocumentNote As String
    Dim isArrival As Boolean

    If sourceTables Is Nothing Then Exit Function

    Set data = New obj_PrsnlEvntBuilderData
    isArrival = data.IsMovementClosingSectionType(sectionTypeText)

    For Each tableValue In sourceTables
        Set sourceTable = tableValue
        If sourceTable Is Nothing Then GoTo ContinueTable

        If Not sourceTable.TryGetColumnIndexByAlias(SOURCE_ALIAS_DOCUMENT_NOTE, columnIndex) Then
            If Not sourceTable.TryGetColumnIndexByName(SOURCE_ALIAS_DOCUMENT_NOTE, columnIndex) Then GoTo ContinueTable
        End If

        For rowIndex = 1 To sourceTable.RowCount
            Set sourceRow = sourceTable.Rows.Item(rowIndex)
            If sourceRow Is Nothing Then GoTo ContinueRow

            documentNoteText = private_NormalizeTemplateScalar(VBA.CStr(sourceRow.GetCellValue(columnIndex)))
            normalizedDocumentNote = private_NormalizeDocumentNoteForWord(documentNoteText, isArrival)
            If VBA.StrComp(documentNoteText, normalizedDocumentNote, VBA.vbBinaryCompare) <> 0 Then
                If Not sourceRow.SetCellRaw(columnIndex, normalizedDocumentNote) Then Exit Function
            End If
ContinueRow:
        Next rowIndex
ContinueTable:
    Next tableValue

    private_TryNormalizeDocumentNotesForWord = True
End Function

Private Function private_TryEnrichMetaDocumentDatesForWord(ByVal sourceTables As Collection) As Boolean
    Dim sourceTable As obj_TableDynamic
    Dim tableIndex As Long
    Dim docDateText As String

    If sourceTables Is Nothing Then Exit Function

    For tableIndex = 2 To sourceTables.Count
        Set sourceTable = sourceTables.Item(tableIndex)
        If sourceTable Is Nothing Then GoTo ContinueTable
        If VBA.StrComp( _
            private_NormalizeCollectionKey(sourceTable.SectionTitle), _
            private_NormalizeCollectionKey(META_SECTION_TYPE_DOCUMENT), _
            VBA.vbTextCompare) <> 0 Then GoTo ContinueTable

        If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_DOC_DATE, docDateText) Then
            VBA.MsgBox "PrototypeNew: meta document table has no DocDate column.", VBA.vbExclamation, "PrototypeNew / WORD export"
            Exit Function
        End If
        If Not private_TryUpsertShortDateValue( _
            sourceTable, SOURCE_ALIAS_DOC_DATE, WORD_ALIAS_DOC_DATE_SHORT, docDateText) Then Exit Function
ContinueTable:
    Next tableIndex

    private_TryEnrichMetaDocumentDatesForWord = True
End Function

Private Function private_NormalizeDocumentNoteForWord( _
    ByVal documentNoteText As String, _
    ByVal isArrival As Boolean _
) As String
    Dim normalizedKey As String

    documentNoteText = private_NormalizeTemplateScalar(documentNoteText)
    normalizedKey = VBA.LCase$(VBA.Trim$(documentNoteText))

    Select Case normalizedKey
        Case "м/к"
            If isArrival Then
                private_NormalizeDocumentNoteForWord = "виписка із медичної карти стаціонарного хворого"
            Else
                private_NormalizeDocumentNoteForWord = "медична карта стаціонарного хворого"
            End If
        Case Else
            private_NormalizeDocumentNoteForWord = documentNoteText
    End Select
End Function

Private Function private_TryParseVacationDuration( _
    ByVal durationText As String, _
    ByRef outVacationDays As Long, _
    ByRef outAdditionalWayDays As Long, _
    ByRef outTotalDays As Long _
) As Boolean
    Dim rx As Object
    Dim matches As Object

    outVacationDays = 0
    outAdditionalWayDays = 0
    outTotalDays = 0
    durationText = VBA.Trim$(durationText)
    If VBA.Len(durationText) = 0 Then
        private_TryParseVacationDuration = True
        Exit Function
    End If

    On Error GoTo EH
    ' Первое число — количество дней самого отпуска. Первое число после плюса
    ' — дополнительные дни на дорогу. Текст после второго числа не учитывается.
    ' Для совместимости также принимается прежняя форма "20 (2)".
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True
    rx.Pattern = "^\s*(\d+)(?:\s*(?:\+|\()\s*(\d+))?"
    Set matches = rx.Execute(durationText)
    If matches.Count = 0 Then
        VBA.MsgBox "PrototypeNew: unsupported vacation duration: '" & durationText & "'. Expected, for example, '15' or '15+4 (дорога)'.", _
            VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    outVacationDays = VBA.CLng(matches(0).SubMatches(0))
    If VBA.Len(VBA.CStr(matches(0).SubMatches(1))) > 0 Then
        outAdditionalWayDays = VBA.CLng(matches(0).SubMatches(1))
    End If
    outTotalDays = outVacationDays + outAdditionalWayDays
    private_TryParseVacationDuration = True
    Exit Function

EH:
    VBA.MsgBox "PrototypeNew: failed to calculate vacation duration from '" & durationText & "': " & Err.Description, VBA.vbExclamation, "PrototypeNew / WORD export"
End Function

Private Function private_TryAppendMovementTvoTablesForReturn( _
    ByVal sourceTables As Collection, _
    ByVal mainSourceTable As obj_TableDynamic, _
    ByVal sectionTypeText As String, _
    ByVal latestMovementTvoChain As Collection _
) As Boolean
    Dim data As obj_PrsnlEvntBuilderData
    Dim existingTable As obj_TableDynamic
    Dim tableIndex As Long
    Dim chainValue As Variant
    Dim chainItem As Object
    Dim tvoTable As obj_TableDynamic
    Dim rankText As String
    Dim fioText As String
    Dim positionCode As String
    Dim positionText As String
    If sourceTables Is Nothing Then Exit Function
    If mainSourceTable Is Nothing Then Exit Function
    Set data = New obj_PrsnlEvntBuilderData
    If Not data.IsMovementClosingSectionType(sectionTypeText) Then
        private_TryAppendMovementTvoTablesForReturn = True
        Exit Function
    End If

    ' Явно добавленная пользователем meta-ТВО строка имеет приоритет над
    ' восстановлением из Movement и предотвращает дублирование пунктов.
    For tableIndex = 2 To sourceTables.Count
        Set existingTable = sourceTables.Item(tableIndex)
        If Not existingTable Is Nothing Then
            If VBA.StrComp( _
                private_NormalizeCollectionKey(existingTable.SectionTitle), _
                private_NormalizeCollectionKey(META_SECTION_TYPE_TVO), _
                VBA.vbTextCompare) = 0 Then
                private_TryAppendMovementTvoTablesForReturn = True
                Exit Function
            End If
        End If
    Next tableIndex

    ' Цепочка является частью snapshot, уже полученного IsExportAllowed.
    ' Это исключает повторный запрос последней строки Movement при preview.
    If latestMovementTvoChain Is Nothing Then
        private_TryAppendMovementTvoTablesForReturn = True
        Exit Function
    End If

    For Each chainValue In latestMovementTvoChain
        Set chainItem = chainValue
        fioText = VBA.CStr(chainItem("FIO"))
        positionCode = VBA.CStr(chainItem("PositionCode"))
        If Not m_ExporterDataProvider.CommonData.TryResolveRankByIpn(VBA.CStr(chainItem("IPN")), rankText) Then Exit Function
        ' Все три формы должности будут прочитаны одним запросом на этапе
        ' meta enrichment; здесь сохраняем только код из Movement.
        positionText = VBA.vbNullString

        Set tvoTable = private_BuildTvoSourceTable( _
            rankText, fioText, VBA.CStr(chainItem("IPN")), positionCode, positionText)
        If tvoTable Is Nothing Then Exit Function
        sourceTables.Add tvoTable
    Next chainValue

    private_TryAppendMovementTvoTablesForReturn = True
End Function

Private Function private_BuildTvoSourceTable( _
    ByVal rankText As String, _
    ByVal fioText As String, _
    ByVal ipnText As String, _
    ByVal positionCodeText As String, _
    ByVal positionText As String _
) As obj_TableDynamic
    Dim tableObj As obj_TableDynamic
    Dim rowObj As obj_Row

    Set tableObj = New obj_TableDynamic
    tableObj.SectionTitle = META_SECTION_TYPE_TVO
    If Not private_TryAddWordSourceColumn(tableObj, SOURCE_ALIAS_RANK) Then Exit Function
    If Not private_TryAddWordSourceColumn(tableObj, SOURCE_ALIAS_FIO) Then Exit Function
    If Not private_TryAddWordSourceColumn(tableObj, SOURCE_ALIAS_IPN) Then Exit Function
    If Not private_TryAddWordSourceColumn(tableObj, SOURCE_ALIAS_POSITION_CODE) Then Exit Function
    If Not private_TryAddWordSourceColumn(tableObj, WORD_ALIAS_POSITION_DEFAULT) Then Exit Function

    Set rowObj = New obj_Row
    rowObj.PushCellRaw rankText
    rowObj.PushCellRaw fioText
    rowObj.PushCellRaw ipnText
    rowObj.PushCellRaw positionCodeText
    rowObj.PushCellRaw positionText
    If Not tableObj.PushRow(rowObj) Then Exit Function

    Set private_BuildTvoSourceTable = tableObj
End Function

Private Function private_TryAddWordSourceColumn( _
    ByVal tableObj As obj_TableDynamic, _
    ByVal aliasText As String _
) As Boolean
    Dim columnObj As obj_Column

    If tableObj Is Nothing Then Exit Function
    aliasText = VBA.Trim$(aliasText)
    If VBA.Len(aliasText) = 0 Then Exit Function
    Set columnObj = New obj_Column
    columnObj.Name = aliasText
    columnObj.Position = tableObj.ColumnCount + 1
    If Not columnObj.AddAlias(aliasText) Then Exit Function
    private_TryAddWordSourceColumn = tableObj.PushColumn(columnObj)
End Function

Private Function private_TryEnrichMetaTvoTablesForWord(ByVal sourceTables As Collection) As Boolean
    Dim tableIndex As Long
    Dim sourceTable As obj_TableDynamic
    Dim rankText As String
    Dim ipnText As String
    Dim positionCodeText As String
    Dim positionText As String
    Dim rankGenitive As String
    Dim rankDative As String
    Dim fioGenitive As String
    Dim fioDative As String
    Dim positionGenitive As String
    Dim positionDative As String
    Dim positionGenitiveFound As Boolean
    Dim positionDativeFound As Boolean
    Dim fallbackPositionText As String
    Dim enrichedCount As Long
    If sourceTables Is Nothing Then Exit Function

    ' Служебные падежные колонки добавляются только в export-source копии
    ' meta-ТВО таблиц. Визуальная форма при этом остается из пяти полей.
    For tableIndex = 2 To sourceTables.Count
        Set sourceTable = Nothing
        Set sourceTable = sourceTables.Item(tableIndex)
        If sourceTable Is Nothing Then GoTo ContinueTable
        If VBA.StrComp( _
            private_NormalizeCollectionKey(sourceTable.SectionTitle), _
            private_NormalizeCollectionKey(META_SECTION_TYPE_TVO), _
            VBA.vbTextCompare) <> 0 Then GoTo ContinueTable

        If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_RANK, rankText) Then rankText = VBA.vbNullString
        If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_IPN, ipnText) Then ipnText = VBA.vbNullString
        If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_POSITION_CODE, positionCodeText) Then positionCodeText = VBA.vbNullString
        If Not private_TryGetMainTableValue(sourceTable, WORD_ALIAS_POSITION_DEFAULT, positionText) Then positionText = VBA.vbNullString

        ' Для каждого базового Rank/FIO/Position шаблон получает единый набор:
        ' default без суффикса, Genitive и Dative.
        If Not m_ExporterDataProvider.CommonData.TryResolveRankGenitive(rankText, rankGenitive) Then Exit Function
        If Not m_ExporterDataProvider.CommonData.TryResolveRankDative(rankText, rankDative) Then Exit Function
        If Not m_ExporterDataProvider.CommonData.TryResolveFioGenitive(ipnText, fioGenitive) Then Exit Function
        If Not m_ExporterDataProvider.CommonData.TryResolveFioDative(ipnText, fioDative) Then Exit Function
#If TVO_POSITION_FALLBACK_ENABLED Then
        If Not m_ExporterDataProvider.CommonData.TryResolvePositionFormsOptional( _
            positionCodeText, fallbackPositionText, positionGenitive, positionDative, positionGenitiveFound) Then Exit Function
        positionDativeFound = positionGenitiveFound
        If positionGenitiveFound Then positionText = fallbackPositionText

        ' Missing dictionary rows are allowed only by this compile-time feature.
        ' The source row supplies the uninflected name and red marks it in preview.
        If Not positionGenitiveFound Or Not positionDativeFound Then
            If VBA.Len(VBA.Trim$(positionText)) = 0 Then
                VBA.MsgBox "PrototypeNew: position code '" & positionCodeText & _
                    "' was not found in ШПО / Посади, and the source row has no position name for the TVO person.", _
                    VBA.vbExclamation, "PrototypeNew / WORD export"
                Exit Function
            End If
#If TVO_POSITION_FALLBACK_HIGHLIGHT_ENABLED Then
            fallbackPositionText = private_WrapPreviewColor(positionText, PREVIEW_FALLBACK_COLOR)
#Else
            fallbackPositionText = positionText
#End If
            positionText = fallbackPositionText
            If Not positionGenitiveFound Then positionGenitive = fallbackPositionText
            If Not positionDativeFound Then positionDative = fallbackPositionText
        End If
#Else
        ' Strict mode preserves the original behavior: a missing position row
        ' is reported by CommonData and stops preview/export immediately.
        If Not m_ExporterDataProvider.CommonData.TryResolvePositionFormsOptional( _
            positionCodeText, positionText, positionGenitive, positionDative, positionGenitiveFound) Then Exit Function
        If Not positionGenitiveFound Then
            VBA.MsgBox "PrototypeNew: declension row was not found in ШПО / Посади for key: " & positionCodeText, _
                VBA.vbExclamation, "PrototypeNew / WORD export"
            Exit Function
        End If
#End If

        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_RANK_GENITIVE, rankGenitive) Then Exit Function
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_RANK_DATIVE, rankDative) Then Exit Function
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_FIO_GENITIVE, fioGenitive) Then Exit Function
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_FIO_DATIVE, fioDative) Then Exit Function
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_POSITION_DEFAULT, positionText) Then Exit Function
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_POSITION_GENITIVE, positionGenitive) Then Exit Function
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_POSITION_DATIVE, positionDative) Then Exit Function
        enrichedCount = enrichedCount + 1

ContinueTable:
    Next tableIndex

#If LOGGING_DEBUG_ENABLED Then
    If enrichedCount > 0 Then ex_Core.fn_Diagnostic_LogInfo "peb-word:meta-tvo-enriched count=" & VBA.CStr(enrichedCount)
#End If
    private_TryEnrichMetaTvoTablesForWord = True
End Function

Private Function private_WrapPreviewColor(ByVal valueText As String, ByVal colorText As String) As String
    If VBA.Len(valueText) = 0 Then Exit Function
    private_WrapPreviewColor = "[[color=" & colorText & "]]" & valueText & "[[/color]]"
End Function

Private Function private_TryResolveDateByRawText( _
    ByVal rawDateText As String, _
    ByRef outHasDate As Boolean, _
    ByRef outDateValue As Date _
) As Boolean
    Dim trimmedDateText As String

    outHasDate = False
    trimmedDateText = VBA.Trim$(rawDateText)
    If VBA.Len(trimmedDateText) = 0 Then
        private_TryResolveDateByRawText = True
        Exit Function
    End If

    If m_ExporterDataProvider Is Nothing Then Exit Function
    If m_ExporterDataProvider.CommonData Is Nothing Then Exit Function

    If m_ExporterDataProvider.CommonData.HasOrderDate Then
        If Not ex_Helpers.fn_TryResolveDateWithContext(trimmedDateText, m_ExporterDataProvider.CommonData.OrderDate, outDateValue) Then
            VBA.MsgBox "PrototypeNew: failed to resolve date from value '" & trimmedDateText & "'.", VBA.vbExclamation, "PrototypeNew / WORD export"
            Exit Function
        End If
    ElseIf ex_Helpers.fn_IsShortDateValue(trimmedDateText) Then
        outDateValue = SENTINEL_SHORT_DATE
    ElseIf Not ex_Helpers.fn_TryResolveDateWithContext(trimmedDateText, SENTINEL_SHORT_DATE, outDateValue) Then
        VBA.MsgBox "PrototypeNew: failed to resolve date from value '" & trimmedDateText & "'.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    outHasDate = True
    private_TryResolveDateByRawText = True
End Function

Private Function private_TryUpsertShortDateValue( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal sourceAlias As String, _
    ByVal shortAlias As String, _
    ByVal rawDateText As String _
) As Boolean
    Dim resolvedDate As Date
    Dim trimmedDateText As String
    Dim shortDateText As String

    trimmedDateText = VBA.Trim$(rawDateText)
    If VBA.Len(trimmedDateText) = 0 Then
        ' Алиас должен присутствовать в контексте даже при пустой дате.
        ' Тогда шаблонная проверка [*DateShort] вернёт False и просто не
        ' выведет дату, не останавливая экспорт из-за отсутствующего поля.
        If Not private_TryUpsertMainTableValue( _
            sourceTable, shortAlias, VBA.vbNullString) Then Exit Function
        private_TryUpsertShortDateValue = True
        Exit Function
    End If

    ' Базовая дата для сокращенных дат берется из common data:
    ' номер приказа -> дата приказа. Если номера/даты приказа нет, короткие
    ' даты намеренно превращаются в 01.01.1900, чтобы проблема была видна
    ' глазами в WORD preview.
    If m_ExporterDataProvider Is Nothing Then Exit Function
    If m_ExporterDataProvider.CommonData Is Nothing Then Exit Function
    If m_ExporterDataProvider.CommonData.HasOrderDate Then
        If Not ex_Helpers.fn_TryResolveDateWithContext(trimmedDateText, m_ExporterDataProvider.CommonData.OrderDate, resolvedDate) Then
            VBA.MsgBox "PrototypeNew: failed to resolve full date for '" & sourceAlias & _
                "' from value '" & trimmedDateText & "'.", VBA.vbExclamation, "PrototypeNew / WORD export"
            Exit Function
        End If
    ElseIf ex_Helpers.fn_IsShortDateValue(trimmedDateText) Then
        resolvedDate = SENTINEL_SHORT_DATE
    ElseIf Not ex_Helpers.fn_TryResolveDateWithContext(trimmedDateText, SENTINEL_SHORT_DATE, resolvedDate) Then
        VBA.MsgBox "PrototypeNew: failed to resolve full date for '" & sourceAlias & _
            "' from value '" & trimmedDateText & "'.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    shortDateText = VBA.Format$(resolvedDate, WORD_SHORT_DATE_STORAGE_FORMAT)

    ' Short — единое представление даты в контексте. Его можно вывести напрямую
    ' либо преобразовать в длинный вид через dateformat непосредственно в XML.
    If Not private_TryUpsertMainTableValue(sourceTable, shortAlias, shortDateText) Then Exit Function

    private_TryUpsertShortDateValue = True
End Function

Private Function private_JoinNonEmptyParts(ByVal leftText As String, ByVal rightText As String) As String
    leftText = VBA.Trim$(leftText)
    rightText = VBA.Trim$(rightText)
    If VBA.Len(leftText) = 0 Then
        private_JoinNonEmptyParts = rightText
    ElseIf VBA.Len(rightText) = 0 Then
        private_JoinNonEmptyParts = leftText
    Else
        private_JoinNonEmptyParts = leftText & " " & rightText
    End If
End Function

Private Function private_KeepSurnameWithInitialsTogether(ByVal valueText As String) As String
    Static initialsRx As Object

    valueText = VBA.Trim$(valueText)
    If VBA.Len(valueText) = 0 Then Exit Function

    If initialsRx Is Nothing Then
        Set initialsRx = VBA.CreateObject("VBScript.RegExp")
        initialsRx.Global = True
        initialsRx.IgnoreCase = False
        ' Два инициала могут быть записаны как "І. І." или "І.І.". Меняем
        ' только пробел перед всей группой: "Прізвище І. І.".
        initialsRx.Pattern = "(\S+)[ \t]+([А-ЯІЇЄҐA-Z]\.[ \t]*[А-ЯІЇЄҐA-Z]\.)"
    End If

    private_KeepSurnameWithInitialsTogether = initialsRx.Replace( _
        valueText, "$1" & VBA.ChrW$(160) & "$2")
End Function

Private Function private_NormalizeHospitalWordTypography(ByVal valueText As String) As String
    Static hospitalNumberRx As Object

    valueText = private_NormalizeLocationWordTypography(valueText)
    valueText = private_KeepSurnameWithInitialsTogether(valueText)
    If VBA.Len(valueText) = 0 Then Exit Function

    If hospitalNumberRx Is Nothing Then
        Set hospitalNumberRx = VBA.CreateObject("VBScript.RegExp")
        hospitalNumberRx.Global = True
        hospitalNumberRx.IgnoreCase = False
        ' Не даём WORD разорвать обозначение и значение: "№ 18".
        hospitalNumberRx.Pattern = "№[ \t]+(\S)"
    End If

    private_NormalizeHospitalWordTypography = hospitalNumberRx.Replace( _
        valueText, "№" & VBA.ChrW$(160) & "$1")
End Function

Private Function private_NormalizeLocationWordTypography(ByVal valueText As String) As String
    Static settlementRx As Object
    Static regionRx As Object

    valueText = VBA.Trim$(valueText)
    If VBA.Len(valueText) = 0 Then Exit Function

    If regionRx Is Nothing Then
        Set regionRx = VBA.CreateObject("VBScript.RegExp")
        regionRx.Global = True
        regionRx.IgnoreCase = True
        ' В адресных значениях сокращение области всегда раскрываем полностью.
        regionRx.Pattern = "обл\."
    End If
    valueText = regionRx.Replace(valueText, "області")

    If settlementRx Is Nothing Then
        Set settlementRx = VBA.CreateObject("VBScript.RegExp")
        settlementRx.Global = True
        settlementRx.IgnoreCase = True
        ' Сначала проверяем длинное "сел.", затем "м."/"с.". После точки
        ' всегда оставляем ровно один неразрывный пробел.
        settlementRx.Pattern = "(сел|м|с)\.[ \t]*"
    End If
    private_NormalizeLocationWordTypography = settlementRx.Replace( _
        valueText, "$1." & VBA.ChrW$(160))
End Function

Private Function private_ApplyGeneratedAliasWordTypography( _
    ByVal columnAlias As String, _
    ByVal valueText As String _
) As String
    ' Вызывается после private_NormalizeTemplateScalar: входные NBSP сначала
    ' удаляются как потенциальный мусор, а затем заново добавляются только в
    ' известных генерируемых WORD aliases. Исходные поля формы не изменяются.
    Select Case VBA.LCase$(VBA.Trim$(columnAlias))
        ' Все поддерживаемые склонённые формы персоны проходят через один
        ' formatter. Для полного ФИО это безопасный no-op, но если справочник
        ' вернёт форму с инициалами, граница фамилия/инициалы также будет NBSP.
        Case VBA.LCase$(WORD_ALIAS_FIO_GENITIVE), _
             VBA.LCase$(WORD_ALIAS_FIO_ACCUSATIVE), _
             VBA.LCase$(WORD_ALIAS_FIO_DATIVE), _
             VBA.LCase$(WORD_ALIAS_FIO_INITIALS_GENITIVE), _
             VBA.LCase$(WORD_ALIAS_REPORT_PERSON_GENITIVE), _
             VBA.LCase$(WORD_ALIAS_REPORT_PERSON_INITIALS_GENITIVE), _
             VBA.LCase$(WORD_ALIAS_REPORTER_GENITIVE)
            valueText = private_KeepSurnameWithInitialsTogether(valueText)

        Case VBA.LCase$(WORD_ALIAS_HOSPITAL_GENITIVE), _
             VBA.LCase$(WORD_ALIAS_HOSPITAL_ACCUSATIVE)
            valueText = private_NormalizeHospitalWordTypography(valueText)

        Case VBA.LCase$(SOURCE_ALIAS_VACATION)
            valueText = private_NormalizeLocationWordTypography(valueText)
    End Select

    private_ApplyGeneratedAliasWordTypography = valueText
End Function

Private Function private_LowerFirstLetter(ByVal valueText As String) As String
    valueText = VBA.Trim$(valueText)
    If VBA.Len(valueText) = 0 Then Exit Function
    private_LowerFirstLetter = VBA.LCase$(VBA.Left$(valueText, 1)) & VBA.Mid$(valueText, 2)
End Function

Private Function private_BuildNamedLoopCollections(ByVal sourceTables As Collection) As Object
    Dim result As Object
    Dim mainTables As Collection
    Dim metaTables As Collection
    Dim metaDocumentTables As Collection
    Dim metaTvoTables As Collection
    Dim sourceTable As obj_TableDynamic
    Dim tableIndex As Long
    Dim sectionTitle As String
    Dim normalizedSectionTitle As String

    If sourceTables Is Nothing Then Exit Function
    If sourceTables.Count <= 0 Then Exit Function

    Set result = VBA.CreateObject("Scripting.Dictionary")
    result.CompareMode = 1
    Set mainTables = New Collection
    Set metaTables = New Collection
    Set metaDocumentTables = New Collection
    Set metaTvoTables = New Collection

    For tableIndex = 1 To sourceTables.Count
        Set sourceTable = Nothing
        Set sourceTable = sourceTables.Item(tableIndex)
        If sourceTable Is Nothing Then GoTo ContinueTable

        If tableIndex = 1 Then
            mainTables.Add sourceTable
        Else
            metaTables.Add sourceTable

            sectionTitle = VBA.Trim$(sourceTable.SectionTitle)
            normalizedSectionTitle = private_NormalizeCollectionKey(sectionTitle)
            If VBA.StrComp(normalizedSectionTitle, private_NormalizeCollectionKey(META_SECTION_TYPE_DOCUMENT), VBA.vbTextCompare) = 0 Then
                metaDocumentTables.Add sourceTable
            End If
            If VBA.StrComp(normalizedSectionTitle, private_NormalizeCollectionKey(META_SECTION_TYPE_TVO), VBA.vbTextCompare) = 0 Then
                metaTvoTables.Add sourceTable
            End If
        End If

        sectionTitle = VBA.Trim$(sourceTable.SectionTitle)
        If VBA.Len(sectionTitle) > 0 Then
            If Not private_TryAddNamedTableCollection(result, sectionTitle, sourceTable) Then Exit Function
            If Not private_TryAddNamedTableCollection(result, private_NormalizeCollectionKey(sectionTitle), sourceTable) Then Exit Function
        End If

ContinueTable:
    Next tableIndex

    Set result("MainTable") = mainTables
    Set result("MetaTables") = metaTables
    Set result(LOOP_COLLECTION_META_DOCUMENT_TABLES) = metaDocumentTables
    Set result(LOOP_COLLECTION_META_TVO_TABLES) = metaTvoTables

    Set private_BuildNamedLoopCollections = result
End Function

Private Function private_TryAddNamedTableCollection( _
    ByVal collectionsMap As Object, _
    ByVal collectionKey As String, _
    ByVal sourceTable As obj_TableDynamic _
) As Boolean
    Dim tableCollection As Collection

    private_TryAddNamedTableCollection = True
    If collectionsMap Is Nothing Then Exit Function
    If sourceTable Is Nothing Then Exit Function

    collectionKey = VBA.Trim$(collectionKey)
    If VBA.Len(collectionKey) = 0 Then Exit Function

    If collectionsMap.Exists(collectionKey) Then
        Set tableCollection = collectionsMap(collectionKey)
    Else
        Set tableCollection = New Collection
        Set collectionsMap(collectionKey) = tableCollection
    End If

    tableCollection.Add sourceTable
End Function

Private Function private_NormalizeCollectionKey(ByVal valueText As String) As String
    valueText = VBA.LCase$(VBA.Trim$(valueText))
    valueText = VBA.Replace(valueText, " ", VBA.vbNullString)
    valueText = VBA.Replace(valueText, ":", VBA.vbNullString)
    valueText = VBA.Replace(valueText, "-", VBA.vbNullString)
    valueText = VBA.Replace(valueText, "_", VBA.vbNullString)
    private_NormalizeCollectionKey = valueText
End Function

Private Function private_TryGetMainTableValue( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal columnAlias As String, _
    ByRef outValue As String _
) As Boolean
    Dim columnIndex As Long
    Dim sourceRow As obj_Row

    outValue = VBA.vbNullString
    If sourceTable Is Nothing Then Exit Function
    If sourceTable.RowCount <= 0 Then Exit Function
    columnAlias = VBA.Trim$(columnAlias)
    If VBA.Len(columnAlias) = 0 Then Exit Function

    If Not sourceTable.TryGetColumnIndexByAlias(columnAlias, columnIndex) Then
        If Not sourceTable.TryGetColumnIndexByName(columnAlias, columnIndex) Then Exit Function
    End If

    Set sourceRow = sourceTable.Rows.Item(1)
    If sourceRow Is Nothing Then Exit Function

    outValue = private_NormalizeTemplateScalar(VBA.CStr(sourceRow.GetCellValue(columnIndex)))
    private_TryGetMainTableValue = True
End Function

Private Function private_TryUpsertMainTableValue( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal columnAlias As String, _
    ByVal valueText As String _
) As Boolean
    Dim columnIndex As Long
    Dim sourceRow As obj_Row
    Dim columnObj As obj_Column

    If sourceTable Is Nothing Then Exit Function
    If sourceTable.RowCount <= 0 Then Exit Function
    columnAlias = VBA.Trim$(columnAlias)
    If VBA.Len(columnAlias) = 0 Then Exit Function

    If Not sourceTable.TryGetColumnIndexByAlias(columnAlias, columnIndex) Then
        If Not sourceTable.TryGetColumnIndexByName(columnAlias, columnIndex) Then
            Set columnObj = New obj_Column
            columnObj.Name = columnAlias
            columnObj.Position = sourceTable.ColumnCount + 1
            If Not columnObj.AddAlias(columnAlias) Then Exit Function
            If Not sourceTable.PushColumn(columnObj) Then Exit Function
            columnIndex = sourceTable.ColumnCount
        End If
    End If

    Set sourceRow = sourceTable.Rows.Item(1)
    If sourceRow Is Nothing Then Exit Function
    valueText = private_NormalizeTemplateScalar(valueText)
    valueText = private_ApplyGeneratedAliasWordTypography(columnAlias, valueText)
    private_TryUpsertMainTableValue = sourceRow.SetCellRaw(columnIndex, valueText)
End Function

Private Function private_GetContextText(ByVal context As Object, ByVal keyText As String) As String
    If context Is Nothing Then Exit Function
    keyText = VBA.Trim$(keyText)
    If VBA.Len(keyText) = 0 Then Exit Function

    On Error Resume Next
    If context.Exists(keyText) Then private_GetContextText = private_NormalizeTemplateScalar(VBA.CStr(context(keyText)))
    If Err.Number <> 0 Then
        Err.Clear
        private_GetContextText = private_NormalizeTemplateScalar(VBA.CStr(VBA.CallByName(context, keyText, VbGet)))
    End If
    On Error GoTo 0
End Function

Private Function private_TrySetContextText( _
    ByVal context As Object, _
    ByVal keyText As String, _
    ByVal valueText As String _
) As Boolean
    Dim normalizedKey As String
    Dim valueToStore As String

    If context Is Nothing Then
        VBA.MsgBox "PrototypeNew: WORD export context is not specified.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    keyText = VBA.Trim$(keyText)
    If VBA.Len(keyText) = 0 Then Exit Function
    normalizedKey = VBA.LCase$(keyText)
    valueToStore = valueText
    If VBA.StrComp(normalizedKey, VBA.LCase$(CONTEXT_WORD_PREVIEW_TEXT), VBA.vbTextCompare) <> 0 Then
        valueToStore = private_NormalizeTemplateScalar(valueText)
    End If

    On Error Resume Next
    context(keyText) = valueToStore
    If Err.Number <> 0 Then
        Err.Clear
        VBA.CallByName context, keyText, VbLet, valueToStore
    End If
    private_TrySetContextText = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function private_NormalizeTemplateScalar(ByVal valueText As String) As String
    Dim rx As Object

    valueText = VBA.CStr(valueText)
    valueText = VBA.Replace(valueText, VBA.vbCrLf, " ")
    valueText = VBA.Replace(valueText, VBA.vbCr, " ")
    valueText = VBA.Replace(valueText, VBA.vbLf, " ")
    valueText = VBA.Replace(valueText, VBA.vbTab, " ")

    ' Remove invisible separators that break #if truthiness/formatting in templates.
    valueText = VBA.Replace(valueText, VBA.ChrW$(160), " ")
    valueText = VBA.Replace(valueText, VBA.ChrW$(8239), " ")
    valueText = VBA.Replace(valueText, VBA.ChrW$(8203), VBA.vbNullString)
    valueText = VBA.Replace(valueText, VBA.ChrW$(8204), VBA.vbNullString)
    valueText = VBA.Replace(valueText, VBA.ChrW$(8205), VBA.vbNullString)
    valueText = VBA.Replace(valueText, VBA.ChrW$(65279), VBA.vbNullString)

    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = False
    rx.Pattern = "\s+"
    valueText = rx.Replace(valueText, " ")

    private_NormalizeTemplateScalar = VBA.Trim$(valueText)
End Function
