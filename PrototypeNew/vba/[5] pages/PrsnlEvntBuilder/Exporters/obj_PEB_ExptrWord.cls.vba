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
Private Const CONTEXT_WORD_PREVIEW_TEXT As String = "WordExportPreviewText"
Private Const CONTEXT_MANUAL_ORDER_NO As String = "ManualOrderNo"
Private Const SENTINEL_SHORT_DATE As Date = #1/1/1900#
Private Const SOURCE_ALIAS_IPN As String = "IPN"
Private Const SOURCE_ALIAS_RANK As String = "Rank"
Private Const SOURCE_ALIAS_FIO As String = "FIO"
Private Const SOURCE_ALIAS_POSITION_CODE As String = "PositionCode"
Private Const SOURCE_ALIAS_POSITION_NAME As String = "PositionName"
Private Const SOURCE_ALIAS_HOSPITAL_SHORT As String = "HospitalShort"
Private Const SOURCE_ALIAS_REPORT_RANK As String = "ReportRank"
Private Const SOURCE_ALIAS_REPORT_PERSON As String = "ReportPerson"
Private Const SOURCE_ALIAS_REPORT_POSITION_CODE As String = "ReportPositionCode"
Private Const SOURCE_ALIAS_INCOMING_DATE As String = "IncomingDate"
Private Const SOURCE_ALIAS_DOC_DATE As String = "DocDate"
Private Const SOURCE_ALIAS_DATE_FROM As String = "DateFrom"
Private Const SOURCE_ALIAS_DURATION_DAYS As String = "DurationDays"
Private Const SOURCE_ALIAS_VH_DATE As String = "VhDate"
Private Const SOURCE_ALIAS_VLK_DATE As String = "VlkDate"
Private Const WORD_ALIAS_RANK_GENITIVE As String = "RankGenitive"
Private Const WORD_ALIAS_FIO_GENITIVE As String = "FIOGenitive"
Private Const WORD_ALIAS_FIO_INITIALS_GENITIVE As String = "FIOInitialsGenitive"
Private Const WORD_ALIAS_POSITION_GENITIVE As String = "PositionGenitive"
Private Const WORD_ALIAS_POSITION_DEFAULT As String = "Position"
Private Const WORD_ALIAS_RANK_DATIVE As String = "RankDative"
Private Const WORD_ALIAS_FIO_DATIVE As String = "FIODative"
Private Const WORD_ALIAS_POSITION_DATIVE As String = "PositionDative"
Private Const WORD_ALIAS_HOSPITAL_GENITIVE As String = "HospitalGenitive"
Private Const WORD_ALIAS_HOSPITAL_ACCUSATIVE As String = "HospitalAccusative"
Private Const WORD_ALIAS_REPORT_RANK_GENITIVE As String = "ReportRankGenitive"
Private Const WORD_ALIAS_REPORT_PERSON_GENITIVE As String = "ReportPersonGenitive"
Private Const WORD_ALIAS_REPORT_PERSON_INITIALS_GENITIVE As String = "ReportPersonInitialsGenitive"
Private Const WORD_ALIAS_REPORT_POSITION_GENITIVE As String = "ReportPositionGenitive"
Private Const WORD_ALIAS_REPORTER_GENITIVE As String = "ReporterGenitive"
Private Const WORD_ALIAS_INCOMING_DATE_RESOLVED As String = "IncomingDateResolved"
Private Const WORD_ALIAS_DOC_DATE_RESOLVED As String = "DocDateResolved"
Private Const WORD_ALIAS_DATE_FROM_RESOLVED As String = "DateFromResolved"
Private Const WORD_ALIAS_VH_DATE_RESOLVED As String = "VhDateResolved"
Private Const WORD_ALIAS_VLK_DATE_RESOLVED As String = "VlkDateResolved"
Private Const WORD_ALIAS_INCOMING_DATE_SHORT As String = "IncomingDateShort"
Private Const WORD_ALIAS_DOC_DATE_SHORT As String = "DocDateShort"
Private Const WORD_ALIAS_DATE_FROM_SHORT As String = "DateFromShort"
Private Const WORD_ALIAS_VH_DATE_SHORT As String = "VhDateShort"
Private Const WORD_ALIAS_VLK_DATE_SHORT As String = "VlkDateShort"
Private Const WORD_ALIAS_INCOMING_DATE_FULL As String = "IncomingDateFull"
Private Const WORD_ALIAS_DOC_DATE_FULL As String = "DocDateFull"
Private Const WORD_ALIAS_DATE_FROM_FULL As String = "DateFromFull"
Private Const WORD_ALIAS_DATE_TO_FULL As String = "DateToFull"
Private Const WORD_ALIAS_VH_DATE_FULL As String = "VhDateFull"
Private Const WORD_ALIAS_VLK_DATE_FULL As String = "VlkDateFull"
Private Const WORD_ALIAS_INCOMING_DATE_FULL_PLUS_ONE As String = "IncomingDateFullPlusOne"
Private Const WORD_ALIAS_DOC_DATE_FULL_PLUS_ONE As String = "DocDateFullPlusOne"
Private Const WORD_ALIAS_DATE_FROM_FULL_PLUS_ONE As String = "DateFromFullPlusOne"
Private Const WORD_ALIAS_VH_DATE_FULL_PLUS_ONE As String = "VhDateFullPlusOne"
Private Const WORD_ALIAS_VLK_DATE_FULL_PLUS_ONE As String = "VlkDateFullPlusOne"
Private Const WORD_ALIAS_ENROLL_TO_FOOD_SUPPORT_DATE As String = "EnrollToFoodSupportDate"
Private Const WORD_ALIAS_REMOVE_FROM_FOOD_SUPPORT_DATE As String = "RemoveFromFoodSupportDate"
Private Const WORD_RESOLVED_DATE_STORAGE_FORMAT As String = "dd.mm.yyyy"
Private Const WORD_ANCHOR_PREFIX As String = "{\export:"
Private Const WORD_ANCHOR_BEGIN_SUFFIX As String = "_Begin}"
Private Const WORD_ANCHOR_END_SUFFIX As String = "_End}"
Private Const WD_FIND_STOP As Long = 0
Private Const WORD_RECORD_BOOKMARK_PREFIX As String = "PEB_"
Private Const WORD_BOOKMARK_MAX_LENGTH As Long = 40
Private Const REPORT_TVO_TEXT As String = "тимчасово виконуючого обов'язки"
Private Const META_SECTION_TYPE_DOCUMENT As String = "Мета: документ"
Private Const META_SECTION_TYPE_TVO As String = "Мета: ТВО"
Private Const SECTION_TYPE_TO_ANNUAL_VACATION_PART As String = "У частину щорічної основної відпустки"
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

    If VBA.StrComp(private_GetContextText(context, "ExportMode"), "Rewrite Last", VBA.vbTextCompare) = 0 Then
        VBA.MsgBox "PrototypeNew: WORD exporter does not support Rewrite Last because it generates a preview and does not persist person records.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

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
    If Not m_ExporterDataProvider.IsExportAllowed(sourceTable, sectionTypeText, exportValidationError, latestMovementTvoChain) Then
        VBA.MsgBox exportValidationError, VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    If Not private_TryEnrichMainSourceTableForWord(sourceTable, context) Then Exit Function
    If Not private_TryAppendMovementTvoTablesForReturn(sourceTables, sourceTable, sectionTypeText, latestMovementTvoChain) Then Exit Function
    If Not private_TryEnrichMetaTvoTablesForWord(sourceTables) Then Exit Function
    Set namedCollections = private_BuildNamedLoopCollections(sourceTables)
    If namedCollections Is Nothing Then Exit Function
    If Not m_TemplateParser.TryRenderForSectionType(sectionTypeText, sourceTables, namedCollections, previewText, templateId) Then Exit Function
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
        If Not private_TryAppendBeforeWordEndAnchor(templateId, recordIpn, previewText) Then Exit Function
    End If

    Export = True
End Function

Private Function private_TryAppendBeforeWordEndAnchor( _
    ByVal templateId As String, _
    ByVal recordIpn As String, _
    ByVal renderedText As String _
) As Boolean
    Dim targetPath As String
    Dim templatePath As String
    Dim beginMarker As String
    Dim endMarker As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim beginRange As Object
    Dim endRange As Object
    Dim insertRange As Object
    Dim insertedStart As Long
    Dim insertedEnd As Long
    Dim plainRenderedText As String
    Dim documentOpened As Boolean
    Dim errorDescription As String
    Dim bookmarkName As String

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

    bookmarkName = private_BuildUniqueRecordBookmarkName(wordDoc, bookmarkName)
    If VBA.Len(bookmarkName) = 0 Then
        VBA.MsgBox "PrototypeNew: failed to create a unique WORD bookmark for IPN '" & recordIpn & "'.", VBA.vbExclamation, "PrototypeNew / WORD export"
        GoTo CleanFail
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

Private Function private_BuildUniqueRecordBookmarkName(ByVal wordDoc As Object, ByVal baseName As String) As String
    Dim candidateName As String
    Dim suffixText As String
    Dim sequenceNo As Long
    Dim baseMaxLength As Long

    baseName = VBA.Trim$(baseName)
    If wordDoc Is Nothing Or VBA.Len(baseName) = 0 Then Exit Function
    If Not wordDoc.Bookmarks.Exists(baseName) Then
        private_BuildUniqueRecordBookmarkName = baseName
        Exit Function
    End If

    sequenceNo = 2
    Do
        suffixText = "_" & VBA.CStr(sequenceNo)
        baseMaxLength = WORD_BOOKMARK_MAX_LENGTH - VBA.Len(suffixText)
        candidateName = VBA.Left$(baseName, baseMaxLength) & suffixText
        If Not wordDoc.Bookmarks.Exists(candidateName) Then
            private_BuildUniqueRecordBookmarkName = candidateName
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
    Dim reportRankText As String
    Dim reportPersonText As String
    Dim reportPositionCodeText As String
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
    Dim fioInitialsGenitive As String
    Dim positionGenitive As String
    Dim hospitalGenitive As String
    Dim hospitalAccusative As String
    Dim reportRankGenitive As String
    Dim reportPersonGenitive As String
    Dim reportPersonInitialsGenitive As String
    Dim reportPositionGenitive As String
    Dim reportTvoPositionGenitive As String
    Dim reporterGenitive As String
    Dim isReporterTvo As Boolean
    Dim dateFromDate As Date
    Dim incomingDate As Date
    Dim hasDateFrom As Boolean
    Dim hasIncomingDate As Boolean
    Dim removeFromFoodSupportDateText As String
    Dim enrollToFoodSupportDateText As String
    Dim enrollToFoodSupportDateValue As Date
    Dim vacationTotalDays As Long
    Dim vacationDateTo As Date
    If sourceTable Is Nothing Then Exit Function
    If m_ExporterDataProvider Is Nothing Then Exit Function
    If m_ExporterDataProvider.CommonData Is Nothing Then Exit Function

    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_IPN, ipnText) Then ipnText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_RANK, rankText) Then rankText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_FIO, fioText) Then fioText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_POSITION_CODE, positionCodeText) Then positionCodeText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_POSITION_NAME, positionText) Then positionText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_HOSPITAL_SHORT, hospitalShortText) Then hospitalShortText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_REPORT_RANK, reportRankText) Then reportRankText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_REPORT_PERSON, reportPersonText) Then reportPersonText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_REPORT_POSITION_CODE, reportPositionCodeText) Then reportPositionCodeText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_INCOMING_DATE, incomingDateText) Then incomingDateText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_DOC_DATE, docDateText) Then docDateText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_DATE_FROM, dateFromText) Then dateFromText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_DURATION_DAYS, durationDaysText) Then durationDaysText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_VH_DATE, vhDateText) Then vhDateText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_VLK_DATE, vlkDateText) Then vlkDateText = VBA.vbNullString

    ' WORD preview работает не только с исходными колонками формы. Перед
    ' render-ом мы дописываем в main DynamicTable вычисленные поля:
    ' склонения ФИО/звания/посады/лечебного учреждения, строку рапортующего
    ' и resolved date aliases. Для дат пишем сразу несколько представлений:
    ' ...Resolved/...Short = 01.02.2025, ...Full = 01 лютого 2025 року,
    ' ...FullPlusOne = полный формат даты + 1 день.
    orderNoText = private_GetContextText(context, CONTEXT_MANUAL_ORDER_NO)
    If Not m_ExporterDataProvider.CommonData.SetOrderNo(orderNoText) Then Exit Function
    If VBA.Len(VBA.Trim$(orderNoText)) > 0 And Not m_ExporterDataProvider.CommonData.HasOrderDate Then
        rt_Messaging.fn_ShowStatusBarWarning _
            "Order date was not found for order number '" & orderNoText & "'. Short dates use 01.01.1900.", _
            5
    End If

    ' Все склонения берутся из общего provider-а. Если справочник пустой или
    ' ключ не найден, provider сам показывает MsgBox с конкретной причиной.
    If Not m_ExporterDataProvider.CommonData.TryResolveRankGenitive(rankText, rankGenitive) Then Exit Function
    If Not m_ExporterDataProvider.CommonData.TryResolveFioGenitive(ipnText, fioGenitive) Then Exit Function
    If Not m_ExporterDataProvider.CommonData.TryResolveFioInitialsGenitive(ipnText, fioInitialsGenitive) Then Exit Function
    If Not m_ExporterDataProvider.CommonData.TryResolvePositionGenitive(positionCodeText, positionGenitive) Then Exit Function
    If Not m_ExporterDataProvider.CommonData.TryResolveHospitalGenitive(hospitalShortText, hospitalGenitive) Then Exit Function
    If Not m_ExporterDataProvider.CommonData.TryResolveHospitalAccusative(hospitalShortText, hospitalAccusative) Then Exit Function
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
    If Not private_TryUpsertResolvedDateValues( _
        sourceTable, SOURCE_ALIAS_INCOMING_DATE, WORD_ALIAS_INCOMING_DATE_RESOLVED, _
        WORD_ALIAS_INCOMING_DATE_SHORT, WORD_ALIAS_INCOMING_DATE_FULL, _
        WORD_ALIAS_INCOMING_DATE_FULL_PLUS_ONE, incomingDateText) Then Exit Function
    If Not private_TryUpsertResolvedDateValues( _
        sourceTable, SOURCE_ALIAS_DOC_DATE, WORD_ALIAS_DOC_DATE_RESOLVED, _
        WORD_ALIAS_DOC_DATE_SHORT, WORD_ALIAS_DOC_DATE_FULL, _
        WORD_ALIAS_DOC_DATE_FULL_PLUS_ONE, docDateText) Then Exit Function
    If Not private_TryUpsertResolvedDateValues( _
        sourceTable, SOURCE_ALIAS_DATE_FROM, WORD_ALIAS_DATE_FROM_RESOLVED, _
        WORD_ALIAS_DATE_FROM_SHORT, WORD_ALIAS_DATE_FROM_FULL, _
        WORD_ALIAS_DATE_FROM_FULL_PLUS_ONE, dateFromText) Then Exit Function
    If Not private_TryUpsertResolvedDateValues( _
        sourceTable, SOURCE_ALIAS_VH_DATE, WORD_ALIAS_VH_DATE_RESOLVED, _
        WORD_ALIAS_VH_DATE_SHORT, WORD_ALIAS_VH_DATE_FULL, _
        WORD_ALIAS_VH_DATE_FULL_PLUS_ONE, vhDateText) Then Exit Function
    If Not private_TryUpsertResolvedDateValues( _
        sourceTable, SOURCE_ALIAS_VLK_DATE, WORD_ALIAS_VLK_DATE_RESOLVED, _
        WORD_ALIAS_VLK_DATE_SHORT, WORD_ALIAS_VLK_DATE_FULL, _
        WORD_ALIAS_VLK_DATE_FULL_PLUS_ONE, vlkDateText) Then Exit Function

    If Not private_TryResolveDateByRawText(dateFromText, hasDateFrom, dateFromDate) Then Exit Function
    If Not private_TryResolveDateByRawText(incomingDateText, hasIncomingDate, incomingDate) Then Exit Function

    sectionTypeText = private_GetContextText(context, CONTEXT_SECTION_TYPE)
    If VBA.Len(sectionTypeText) = 0 Then sectionTypeText = VBA.Trim$(sourceTable.SectionTitle)

    ' В форме частичной ежегодной отпуска пользователь вводит только дату
    ' начала и продолжительность. Дату окончания не храним в UI: вычисляем
    ' включительно (start + total days - 1) и отдаем шаблону как DateToFull.
    If VBA.StrComp(sectionTypeText, SECTION_TYPE_TO_ANNUAL_VACATION_PART, VBA.vbTextCompare) = 0 Then
        If hasDateFrom Then
            If Not private_TrySumDurationDays(durationDaysText, vacationTotalDays) Then Exit Function
            If vacationTotalDays > 0 Then
                vacationDateTo = VBA.DateAdd("d", vacationTotalDays - 1, dateFromDate)
                If Not private_TryUpsertMainTableValue( _
                    sourceTable, WORD_ALIAS_DATE_TO_FULL, ex_Helpers.fn_FormatUaDateLong(vacationDateTo)) Then Exit Function
            End If
        End If
    End If

    If hasDateFrom Then
        removeFromFoodSupportDateText = ex_Helpers.fn_FormatUaDateLong(dateFromDate)
    ElseIf hasIncomingDate Then
        removeFromFoodSupportDateText = ex_Helpers.fn_FormatUaDateLong(incomingDate)
    End If

    If hasIncomingDate Then
        enrollToFoodSupportDateValue = VBA.DateAdd("d", 1, incomingDate)
        If hasDateFrom Then
            If dateFromDate > enrollToFoodSupportDateValue Then enrollToFoodSupportDateValue = dateFromDate
        End If
        enrollToFoodSupportDateText = ex_Helpers.fn_FormatUaDateLong(enrollToFoodSupportDateValue)
    ElseIf hasDateFrom Then
        enrollToFoodSupportDateText = ex_Helpers.fn_FormatUaDateLong(dateFromDate)
    End If

    If VBA.Len(VBA.Trim$(removeFromFoodSupportDateText)) > 0 Then
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_REMOVE_FROM_FOOD_SUPPORT_DATE, removeFromFoodSupportDateText) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(enrollToFoodSupportDateText)) > 0 Then
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_ENROLL_TO_FOOD_SUPPORT_DATE, enrollToFoodSupportDateText) Then Exit Function
    End If

    private_TryEnrichMainSourceTableForWord = True
End Function

Private Function private_TrySumDurationDays( _
    ByVal durationText As String, _
    ByRef outTotalDays As Long _
) As Boolean
    Dim rx As Object
    Dim matches As Object
    Dim matchObj As Object

    outTotalDays = 0
    durationText = VBA.Trim$(durationText)
    If VBA.Len(durationText) = 0 Then
        private_TrySumDurationDays = True
        Exit Function
    End If

    On Error GoTo EH
    ' DurationDays допускает пользовательские записи вроде "20+2" или
    ' "20 (2)". Суммируем все целые числа, повторяя смысл Excel-формулы формы.
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.Pattern = "\d+"
    Set matches = rx.Execute(durationText)
    For Each matchObj In matches
        outTotalDays = outTotalDays + VBA.CLng(matchObj.Value)
    Next matchObj
    private_TrySumDurationDays = True
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

Private Function private_TryUpsertResolvedDateValues( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal sourceAlias As String, _
    ByVal resolvedAlias As String, _
    ByVal shortAlias As String, _
    ByVal fullAlias As String, _
    ByVal fullPlusOneAlias As String, _
    ByVal rawDateText As String _
) As Boolean
    Dim resolvedDate As Date
    Dim trimmedDateText As String
    Dim shortDateText As String
    Dim fullDateText As String
    Dim fullPlusOneDateText As String

    trimmedDateText = VBA.Trim$(rawDateText)
    If VBA.Len(trimmedDateText) = 0 Then
        private_TryUpsertResolvedDateValues = True
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

    shortDateText = VBA.Format$(resolvedDate, WORD_RESOLVED_DATE_STORAGE_FORMAT)
    fullDateText = ex_Helpers.fn_FormatUaDateLong(resolvedDate)
    fullPlusOneDateText = ex_Helpers.fn_FormatUaDateLong(VBA.DateAdd("d", 1, resolvedDate))

    ' Resolved оставляем как технический alias для formatter pipeline, а Short/Full
    ' даем шаблону как готовые текстовые представления без лишнего форматирования.
    If Not private_TryUpsertMainTableValue(sourceTable, resolvedAlias, shortDateText) Then Exit Function
    If Not private_TryUpsertMainTableValue(sourceTable, shortAlias, shortDateText) Then Exit Function
    If Not private_TryUpsertMainTableValue(sourceTable, fullAlias, fullDateText) Then Exit Function
    If Not private_TryUpsertMainTableValue(sourceTable, fullPlusOneAlias, fullPlusOneDateText) Then Exit Function

    private_TryUpsertResolvedDateValues = True
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
