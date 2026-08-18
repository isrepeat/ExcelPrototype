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
Private Const PREVIEW_FALLBACK_COLOR As String = "#FFFF00"
Private Const PREVIEW_TRUNCATION_WARNING_LENGTH As Long = 254
Private Const PREVIEW_TRUNCATED_VALUE_TAG As String = "preview-truncated-value"
Private Const PREVIEW_LOOKUP_WARNING_VALUE_TAG As String = "preview-lookup-warning-value"
Private Const CONTEXT_SECTION_TYPE As String = "SectionType"
Private Const CONTEXT_VALIDATION_ENABLED As String = "ValidateWord"
Private Const CONTEXT_WORD_PREVIEW_TEXT As String = "WordExportPreviewText"
Private Const CONTEXT_MANUAL_ORDER_NO As String = "ManualOrderNo"
Private Const CONTEXT_MANUAL_ORDER_DATE_SERIAL As String = "ManualOrderDateSerial"
Private Const CONTEXT_REPORT_IS_TVO As String = "ReportIsTvo"
Private Const CONTEXT_MOVEMENT_PREVALIDATED As String = "MovementPrevalidated"
Private Const SENTINEL_SHORT_DATE As Date = #1/1/1900#
Private Const SOURCE_ALIAS_IPN As String = "IPN"
Private Const TEMPORARY_IPN_MARKER As String = "ТП-ІПН"
Private Const SOURCE_ALIAS_RANK As String = "Rank"
Private Const SOURCE_ALIAS_FIO As String = "FIO"
Private Const SOURCE_ALIAS_POSITION_CODE As String = "PositionCode"
Private Const SOURCE_ALIAS_POSITION_NAME As String = "PositionName"
Private Const SOURCE_ALIAS_HOSPITAL_SHORT As String = "HospitalShort"
Private Const SOURCE_ALIAS_TO_HOSPITAL_SHORT As String = "ToHospitalShort"
Private Const SOURCE_ALIAS_DESTINATION As String = "Destination"
Private Const SOURCE_ALIAS_REPORT_RANK As String = "ReportRank"
Private Const SOURCE_ALIAS_REPORT_PERSON As String = "ReportPerson"
Private Const SOURCE_ALIAS_REPORT_POSITION_CODE As String = "ReportPositionCode"
Private Const SOURCE_ALIAS_INCOMING_NO As String = "IncomingNo"
Private Const SOURCE_ALIAS_INCOMING_DATE As String = "IncomingDate"
Private Const SOURCE_ALIAS_DOCUMENT_NOTE As String = "DocumentNote"
Private Const SOURCE_ALIAS_DOC_DATE As String = "DocDate"
Private Const SOURCE_ALIAS_DATE_FROM As String = "DateFrom"
Private Const SOURCE_ALIAS_DURATION_DAYS As String = "DurationDays"
Private Const SOURCE_ALIAS_VACATION_TICKET_NO As String = "VacationTicketNo"
Private Const SOURCE_ALIAS_VACATION_TICKET_DATE As String = "VacationTicketDate"
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
Private Const WORD_ALIAS_TO_HOSPITAL_ACCUSATIVE As String = "ToHospitalAccusative"
Private Const WORD_ALIAS_REPORT_RANK_GENITIVE As String = "ReportRankGenitive"
Private Const WORD_ALIAS_REPORT_PERSON_GENITIVE As String = "ReportPersonGenitive"
Private Const WORD_ALIAS_REPORT_PERSON_INITIALS_GENITIVE As String = "ReportPersonInitialsGenitive"
Private Const WORD_ALIAS_REPORT_POSITION_GENITIVE As String = "ReportPositionGenitive"
Private Const WORD_ALIAS_REPORTER_GENITIVE As String = "ReporterGenitive"
Private Const WORD_ALIAS_ORDER_NO As String = "OrderNo"
Private Const WORD_ALIAS_INCOMING_DATE_SHORT As String = "IncomingDateShort"
Private Const WORD_ALIAS_DOC_DATE_SHORT As String = "DocDateShort"
Private Const WORD_ALIAS_DATE_FROM_SHORT As String = "DateFromShort"
Private Const WORD_ALIAS_EFFECTIVE_RETURN_DATE_SHORT As String = "EffectiveReturnDateShort"
Private Const WORD_ALIAS_VACATION_TICKET_DATE_SHORT As String = "VacationTicketDateShort"
Private Const WORD_ALIAS_VLK_DATE_SHORT As String = "VlkDateShort"
Private Const WORD_GROUP_BOOKMARK_PREFIX As String = "PEG_"
Private Const WORD_NESTED_GROUP_BOOKMARK_PREFIX As String = "PEN_"

' DateTo отсутствует в draft-форме: для частичной ежегодной отпуска экспортёр
' вычисляет его из DateFrom + DurationDays - 1 и добавляет в контекст под этим
' alias. Суффикс Short означает единое компактное представление полной даты с
' определённым годом, но ещё не готовый длинный текст для WORD. Шаблон должен
' явно выбрать представление через formatter, например:
'   {[DateToShort]|dateformat:"\dd \month \yyyy року"}
Private Const WORD_ALIAS_DATE_TO_SHORT As String = "DateToShort"
Private Const WORD_ALIAS_VACATION_DAYS As String = "VacationDays"
Private Const WORD_ALIAS_VACATION_DURATION_PHRASE As String = "VacationDurationPhrase"
Private Const WORD_ALIAS_ADDITIONAL_DURATION_PHRASES As String = "AdditionalDurationPhrases"
Private Const WORD_ALIAS_VACATION_TOTAL_DAYS As String = "VacationTotalDays"
Private Const WORD_ALIAS_VACATION_DATES_SAME_MONTH As String = "VacationDatesSameMonth"
Private Const WORD_ALIAS_VACATION_DATES_SAME_YEAR As String = "VacationDatesSameYear"
Private Const WORD_ALIAS_ENROLL_TO_FOOD_SUPPORT_DATE As String = "EnrollToFoodSupportDateShort"
Private Const WORD_ALIAS_REMOVE_FROM_FOOD_SUPPORT_DATE As String = "RemoveFromFoodSupportDateShort"
Private Const WORD_ALIAS_REQUIRES_FOOD_SUPPORT_CHANGE As String = "RequiresFoodSupportChange"
Private Const WORD_ALIAS_REQUIRES_FOOD_SUPPORT_ENROLL As String = "RequiresFoodSupportEnroll"
Private Const WORD_ALIAS_REQUIRES_FOOD_SUPPORT_REMOVE As String = "RequiresFoodSupportRemove"
Private Const WORD_ALIAS_PREV_VACATION_TICKET_NO As String = "PrevVacationTicketNo"
Private Const WORD_ALIAS_PREV_VACATION_TICKET_DATE_SHORT As String = "PrevVacationTicketDateShort"

' Канонический компактный формат date aliases внутри контекста шаблона.
' Он нужен не для окончательного отображения в WORD, а чтобы formatter pipeline
' мог снова однозначно распознать дату независимо от региональных настроек Excel.
' Поэтому здесь намеренно нет названий месяцев, слова "року" и типографических
' пробелов: всё это контролируется непосредственно маской dateformat в XML.
' Этим форматом заполняются DateToShort, DateFromShort, PrevVacationTicketDateShort
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
Private Const WORD_BOOKMARK_MARKER_BEGIN As String = "[[PEB_BOOKMARK_BEGIN:"
Private Const WORD_BOOKMARK_MARKER_END As String = "[[PEB_BOOKMARK_END:"
Private Const WORD_BOOKMARK_MARKER_ROOT As String = "[[PEB_BOOKMARK_"
Private Const WORD_BOOKMARK_MARKER_SUFFIX As String = "]]"
Private Const REPORT_TVO_TEXT As String = "тимчасово виконуючого обов'язки"
Private Const META_SECTION_TYPE_DOCUMENT As String = "Мета: документ"
Private Const META_SECTION_TYPE_TVO As String = "Мета: ТВО"
Private Const LOOP_COLLECTION_META_DOCUMENT_TABLES As String = "MetaDocumentTables"
Private Const LOOP_COLLECTION_META_TVO_TABLES As String = "MetaTvoTables"


Private m_IsDisposed As Boolean
Private m_Base As obj_DataExporterBase
Private m_TemplateParser As obj_WordResultTplParser
Private m_ExporterCfgDataProvider As obj_PEB_ExptrCfgDataPrvdr
Private m_LastBuildHasWarnings As Boolean
Private m_OwnsExporterCfgDataProvider As Boolean

Public Function ToggleSupportedBookmarks( _
    ByRef outMarkersAreVisible As Boolean, _
    ByRef convertedCount As Long, _
    Optional ByVal orderNo As String = "" _
) As Boolean
    Dim templatePath As String
    Dim targetPath As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim documentOpened As Boolean
    Dim errorDescription As String
    Dim previousScreenUpdating As Boolean
    Dim screenUpdatingChanged As Boolean

    On Error GoTo EH
    outMarkersAreVisible = False
    convertedCount = 0

    templatePath = VBA.Trim$(m_Base.TargetWorkbookPath)
    If VBA.Len(templatePath) = 0 Then
        VBA.MsgBox "PrototypeNew: required profile key 'Export.Word.FilePath' is empty.", _
            VBA.vbExclamation, "PrototypeNew / WORD bookmarks"
        Exit Function
    End If
    If Not private_IsAbsolutePath(templatePath) Then _
        templatePath = ThisWorkbook.Path & Application.PathSeparator & templatePath
    targetPath = private_BuildResultDocumentPath(templatePath, orderNo)
    If VBA.Len(targetPath) = 0 Or VBA.Len(VBA.Dir$(targetPath, _
        VBA.vbNormal Or VBA.vbReadOnly Or VBA.vbHidden Or VBA.vbSystem)) = 0 Then
        VBA.MsgBox "PrototypeNew: WORD result document was not found:" & _
            VBA.vbCrLf & targetPath, VBA.vbExclamation, _
            "PrototypeNew / WORD bookmarks"
        Exit Function
    End If
    If Not rt_WordExportRuntime.fn_TryAcquireWordDocument( _
        targetPath, wordApp, wordDoc, documentOpened) Then Exit Function
    On Error Resume Next
    previousScreenUpdating = wordApp.ScreenUpdating
    wordApp.ScreenUpdating = False
    screenUpdatingChanged = (Err.Number = 0)
    Err.Clear
    On Error GoTo EH

    If VBA.InStr(1, VBA.CStr(wordDoc.Content.Text), _
        WORD_BOOKMARK_MARKER_ROOT, VBA.vbTextCompare) > 0 Then
        If Not private_RestoreSupportedBookmarks(wordDoc, convertedCount) Then _
            GoTo CleanFail
        outMarkersAreVisible = False
    Else
        If Not private_ShowSupportedBookmarksAsMarkers(wordDoc, convertedCount) Then _
            GoTo CleanFail
        outMarkersAreVisible = True
    End If

    ' Переключение закладок является ручной операцией просмотра/редактирования.
    ' Документ намеренно оставляем открытым с несохранёнными изменениями:
    ' пользователь сам решает, сохранить преобразование или отменить его.
    If screenUpdatingChanged Then wordApp.ScreenUpdating = previousScreenUpdating
    wordApp.Visible = True
    wordDoc.Activate
    ToggleSupportedBookmarks = True
    Exit Function

CleanFail:
    On Error Resume Next
    If screenUpdatingChanged Then wordApp.ScreenUpdating = previousScreenUpdating
    If documentOpened Then wordDoc.Close False
    On Error GoTo 0
    Exit Function
EH:
    errorDescription = Err.Description
    On Error Resume Next
    If screenUpdatingChanged Then wordApp.ScreenUpdating = previousScreenUpdating
    If documentOpened Then wordDoc.Close False
    On Error GoTo 0
    VBA.MsgBox "Не удалось переключить WORD-закладки: " & errorDescription, _
        VBA.vbExclamation, "PrsnlEventBuilder / WORD bookmarks"
End Function

Public Function GetRecordBookmarks( _
    ByVal orderNo As String, _
    ByRef outEvents As Collection _
) As Boolean
    Dim targetPath As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim documentOpened As Boolean
    Dim bookmarkObj As Object
    Dim bookmarkName As String
    Dim bookmarkIpn As String
    Dim captionText As String
    Dim optionObj As obj_SelectOption

    On Error GoTo EH
    Set outEvents = New Collection
    If Not private_TryBuildExistingResultPath(orderNo, targetPath, False) Then
        If VBA.Len(targetPath) = 0 Then Exit Function
        GetRecordBookmarks = True
        Exit Function
    End If
    If Not rt_WordExportRuntime.fn_TryAcquireWordDocument( _
        targetPath, wordApp, wordDoc, documentOpened) Then Exit Function
    If VBA.InStr(1, VBA.CStr(wordDoc.Content.Text), _
        WORD_BOOKMARK_MARKER_ROOT, VBA.vbTextCompare) > 0 Then
        VBA.MsgBox "Сначала восстановите WORD-закладки из видимых маркеров.", _
            VBA.vbExclamation, "PrsnlEventBuilder / WORD events"
        GoTo CleanFail
    End If

    For Each bookmarkObj In wordDoc.Bookmarks
        bookmarkName = VBA.CStr(bookmarkObj.Name)
        If VBA.Left$(VBA.UCase$(bookmarkName), _
            VBA.Len(WORD_RECORD_BOOKMARK_PREFIX)) = _
            WORD_RECORD_BOOKMARK_PREFIX Then
            captionText = private_BuildWordEventCaption( _
                VBA.CStr(bookmarkObj.Range.Text))
            bookmarkIpn = private_ExtractRecordBookmarkIpn(bookmarkName)
            Set optionObj = New obj_SelectOption
            optionObj.Id = bookmarkName
            optionObj.Caption = "WORD" & VBA.vbTab & bookmarkName & _
                VBA.vbTab & captionText & VBA.vbTab & bookmarkIpn
            outEvents.Add optionObj
        End If
    Next bookmarkObj
    If documentOpened Then wordDoc.Close False
    GetRecordBookmarks = True
    Exit Function
CleanFail:
    If documentOpened Then wordDoc.Close False
    Exit Function
EH:
    On Error Resume Next
    If documentOpened Then wordDoc.Close False
    On Error GoTo 0
    VBA.MsgBox "Не удалось прочитать события WORD: " & Err.Description, _
        VBA.vbExclamation, "PrsnlEventBuilder / WORD events"
End Function

Public Function TryGetTemplateHashByName( _
    ByVal templateName As String, _
    ByRef outTemplateHash As String _
) As Boolean
    outTemplateHash = VBA.vbNullString
    If m_IsDisposed Or m_TemplateParser Is Nothing Then Exit Function
    TryGetTemplateHashByName = m_TemplateParser.TryGetTemplateHashByName( _
        templateName, outTemplateHash)
End Function

Public Function TryGetTemplateNameByHash( _
    ByVal templateHash As String, _
    ByRef outTemplateName As String _
) As Boolean
    outTemplateName = VBA.vbNullString
    If m_IsDisposed Or m_TemplateParser Is Nothing Then Exit Function
    TryGetTemplateNameByHash = m_TemplateParser.TryGetTemplateNameByHash( _
        templateHash, outTemplateName)
End Function

Private Function private_ExtractRecordBookmarkIpn( _
    ByVal bookmarkName As String _
) As String
    Dim separatorIndex As Long
    Dim bookmarkParts As Variant
    Dim lastPartIndex As Long

    bookmarkParts = VBA.Split(bookmarkName, "_")
    lastPartIndex = UBound(bookmarkParts)
    If lastPartIndex >= 2 Then
        ' Новый формат: PEB_<HH>_<IPN>[_<Occurrence>].
        If VBA.Len(VBA.CStr(bookmarkParts(1))) = 2 Then
            private_ExtractRecordBookmarkIpn = VBA.CStr(bookmarkParts(2))
            Exit Function
        End If
        ' Legacy duplicate: PEB_<TruncatedName>_<IPN>_<Occurrence>.
        If lastPartIndex >= 3 And VBA.IsNumeric(bookmarkParts(lastPartIndex)) And _
            VBA.IsNumeric(bookmarkParts(lastPartIndex - 1)) Then
            private_ExtractRecordBookmarkIpn = _
                VBA.CStr(bookmarkParts(lastPartIndex - 1))
            Exit Function
        End If
    End If
    separatorIndex = VBA.InStrRev(bookmarkName, "_", -1, VBA.vbBinaryCompare)
    If separatorIndex <= VBA.Len(WORD_RECORD_BOOKMARK_PREFIX) Then Exit Function
    If separatorIndex >= VBA.Len(bookmarkName) Then Exit Function
    private_ExtractRecordBookmarkIpn = VBA.Mid$(bookmarkName, separatorIndex + 1)
End Function

Public Function DeleteRecordBookmark( _
    ByVal bookmarkName As String, _
    ByVal orderNo As String, _
    ByRef outSnapshot As Object _
) As Boolean
    Dim targetPath As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim documentOpened As Boolean
    Dim targetRange As Object
    Dim undoRange As Object
    Dim snapshot As Object
    Dim perfTotalStartedAt As Double
    Dim perfStageStartedAt As Double
    Dim acquireMs As Double
    Dim snapshotMs As Double
    Dim deleteMs As Double
    Dim saveMs As Double

    Set outSnapshot = Nothing
    perfTotalStartedAt = VBA.Timer

    On Error GoTo EH
    bookmarkName = VBA.Trim$(bookmarkName)
    If VBA.Left$(VBA.UCase$(bookmarkName), _
        VBA.Len(WORD_RECORD_BOOKMARK_PREFIX)) <> _
        WORD_RECORD_BOOKMARK_PREFIX Then
        VBA.MsgBox "Некорректная WORD-закладка события: " & bookmarkName, _
            VBA.vbExclamation, "PrsnlEventBuilder / WORD events"
        Exit Function
    End If
    If Not private_TryBuildExistingResultPath(orderNo, targetPath, True) Then Exit Function
    perfStageStartedAt = VBA.Timer
    If Not rt_WordExportRuntime.fn_TryAcquireWordDocument( _
        targetPath, wordApp, wordDoc, documentOpened) Then Exit Function
    acquireMs = private_PerfElapsedMs(perfStageStartedAt)
    If Not wordDoc.Bookmarks.Exists(bookmarkName) Then
        VBA.MsgBox "WORD-событие изменилось после обновления списка. " & _
            "Обновите список и повторите удаление.", _
            VBA.vbExclamation, "PrsnlEventBuilder / WORD events"
        GoTo CleanFail
    End If

    perfStageStartedAt = VBA.Timer
    Set targetRange = wordDoc.Bookmarks(bookmarkName).Range
    Set undoRange = private_GetDeleteUndoRange(wordDoc, targetRange)
    If undoRange Is Nothing Then Set undoRange = targetRange.Duplicate
    Set snapshot = VBA.CreateObject("Scripting.Dictionary")
    snapshot("Kind") = "Word"
    snapshot("EventId") = bookmarkName
    snapshot("OrderNo") = orderNo
    snapshot("TargetPath") = targetPath
    snapshot("RangeStart") = VBA.CLng(undoRange.Start)
    snapshot("RangeLength") = VBA.CLng(undoRange.End) - VBA.CLng(undoRange.Start)
    snapshot("WordOpenXml") = VBA.CStr(undoRange.WordOpenXML)
    snapshotMs = private_PerfElapsedMs(perfStageStartedAt)
    perfStageStartedAt = VBA.Timer
    If Not private_TryDeleteExportRange(wordDoc, undoRange) Then GoTo CleanFail
    deleteMs = private_PerfElapsedMs(perfStageStartedAt)
    perfStageStartedAt = VBA.Timer
    wordDoc.Save
    saveMs = private_PerfElapsedMs(perfStageStartedAt)
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "perf:word-delete totalMs='" & _
        VBA.Format$(private_PerfElapsedMs(perfTotalStartedAt), "0") & _
        "' acquireMs='" & VBA.Format$(acquireMs, "0") & _
        "' snapshotMs='" & VBA.Format$(snapshotMs, "0") & _
        "' deleteMs='" & VBA.Format$(deleteMs, "0") & _
        "' saveMs='" & VBA.Format$(saveMs, "0") & "'"
#End If
    If documentOpened Then wordDoc.Close False
    Set outSnapshot = snapshot
    DeleteRecordBookmark = True
    Exit Function
CleanFail:
    If documentOpened Then wordDoc.Close False
    Exit Function
EH:
    On Error Resume Next
    If documentOpened Then wordDoc.Close False
    On Error GoTo 0
    VBA.MsgBox "Не удалось удалить событие WORD: " & Err.Description, _
        VBA.vbExclamation, "PrsnlEventBuilder / WORD events"
End Function

Private Function private_TryDeleteExportRange( _
    ByVal wordDoc As Object, _
    ByVal deleteRange As Object _
) As Boolean
    Dim bookmarkObj As Object
    Dim bookmarkRange As Object
    Dim bookmarkNames As Collection
    Dim bookmarkNameObj As Variant
    Dim bookmarkName As String

    If wordDoc Is Nothing Or deleteRange Is Nothing Then Exit Function
    Set bookmarkNames = New Collection

    ' Undo scope уже выбран как запись или самая внешняя группа, которая
    ' опустеет после её удаления. Поэтому глобальный O(n²) поиск пустых групп
    ' не нужен: одним проходом собираем только экспортные bookmarks внутри scope.
    For Each bookmarkObj In wordDoc.Bookmarks
        bookmarkName = VBA.UCase$(VBA.CStr(bookmarkObj.Name))
        If VBA.Left$(bookmarkName, VBA.Len(WORD_RECORD_BOOKMARK_PREFIX)) <> _
            WORD_RECORD_BOOKMARK_PREFIX And _
            VBA.Left$(bookmarkName, VBA.Len(WORD_GROUP_BOOKMARK_PREFIX)) <> _
            WORD_GROUP_BOOKMARK_PREFIX And _
            VBA.Left$(bookmarkName, VBA.Len(WORD_NESTED_GROUP_BOOKMARK_PREFIX)) <> _
            WORD_NESTED_GROUP_BOOKMARK_PREFIX And _
            VBA.Left$(bookmarkName, VBA.Len(WORD_METADATA_BOOKMARK_PREFIX)) <> _
            WORD_METADATA_BOOKMARK_PREFIX Then GoTo ContinueBookmark
        Set bookmarkRange = bookmarkObj.Range
        If bookmarkRange.Start >= deleteRange.Start And _
            bookmarkRange.End <= deleteRange.End Then _
            bookmarkNames.Add VBA.CStr(bookmarkObj.Name)
ContinueBookmark:
    Next bookmarkObj

    For Each bookmarkNameObj In bookmarkNames
        bookmarkName = VBA.CStr(bookmarkNameObj)
        If wordDoc.Bookmarks.Exists(bookmarkName) Then _
            wordDoc.Bookmarks(bookmarkName).Delete
    Next bookmarkNameObj
    deleteRange.Delete
    private_TryDeleteExportRange = True
End Function

Private Function private_PerfElapsedMs(ByVal startedAt As Double) As Double
    Dim elapsedSeconds As Double

    elapsedSeconds = VBA.Timer - startedAt
    If elapsedSeconds < 0 Then elapsedSeconds = elapsedSeconds + 86400#
    private_PerfElapsedMs = elapsedSeconds * 1000#
End Function

Public Function RestoreDeletedRecord(ByVal snapshot As Object) As Boolean
    Dim targetPath As String
    Dim bookmarkName As String
    Dim wordOpenXml As String
    Dim insertStart As Long
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim insertRange As Object
    Dim documentOpened As Boolean
    Dim contentInserted As Boolean

    On Error GoTo EH
    If snapshot Is Nothing Then Exit Function
    If Not snapshot.Exists("TargetPath") Or _
        Not snapshot.Exists("RangeStart") Or _
        Not snapshot.Exists("WordOpenXml") Or _
        Not snapshot.Exists("EventId") Then Exit Function
    targetPath = VBA.CStr(snapshot("TargetPath"))
    bookmarkName = VBA.CStr(snapshot("EventId"))
    wordOpenXml = VBA.CStr(snapshot("WordOpenXml"))
    insertStart = VBA.CLng(snapshot("RangeStart"))
    If Not rt_WordExportRuntime.fn_TryAcquireWordDocument( _
        targetPath, wordApp, wordDoc, documentOpened) Then Exit Function
    If wordDoc.Bookmarks.Exists(bookmarkName) Then GoTo StateChanged
    If insertStart < 0 Or insertStart > wordDoc.Content.End Then GoTo StateChanged
    Set insertRange = wordDoc.Range(insertStart, insertStart)
    insertRange.InsertXML wordOpenXml
    contentInserted = True
    If Not wordDoc.Bookmarks.Exists(bookmarkName) Then GoTo StateChanged
    wordDoc.Save
    If documentOpened Then wordDoc.Close False
    RestoreDeletedRecord = True
    Exit Function

StateChanged:
    If contentInserted Then
        On Error Resume Next
        insertRange.Delete
        On Error GoTo 0
    End If
    VBA.MsgBox "WORD-документ изменился после удаления. " & _
        "Автоматическое восстановление в прежнюю позицию отменено.", _
        VBA.vbExclamation, "PrsnlEventBuilder / WORD undo"
    If documentOpened Then wordDoc.Close False
    Exit Function
EH:
    On Error Resume Next
    If documentOpened Then wordDoc.Close False
    On Error GoTo 0
    VBA.MsgBox "Не удалось восстановить WORD-событие: " & Err.Description, _
        VBA.vbExclamation, "PrsnlEventBuilder / WORD undo"
End Function

Private Function private_GetDeleteUndoRange( _
    ByVal wordDoc As Object, _
    ByVal recordRange As Object _
) As Object
    Dim bookmarkObj As Object
    Dim candidateRange As Object
    Dim recordBookmark As Object
    Dim recordRangeObj As Object
    Dim candidateEntries As Collection
    Dim candidateEntry As Object
    Dim candidateItem As Variant
    Dim bestLength As Long

    If wordDoc Is Nothing Or recordRange Is Nothing Then Exit Function
    Set private_GetDeleteUndoRange = recordRange.Duplicate
    bestLength = VBA.CLng(recordRange.End) - VBA.CLng(recordRange.Start)
    Set candidateEntries = New Collection

    ' Сначала оставляем только группы, содержащие удаляемую запись. Обычно это
    ' одна-две группы; Range остальных bookmarks больше не читается повторно.
    For Each bookmarkObj In wordDoc.Bookmarks
        If VBA.Left$(VBA.UCase$(VBA.CStr(bookmarkObj.Name)), _
            VBA.Len(WORD_GROUP_BOOKMARK_PREFIX)) <> WORD_GROUP_BOOKMARK_PREFIX And _
            VBA.Left$(VBA.UCase$(VBA.CStr(bookmarkObj.Name)), _
            VBA.Len(WORD_NESTED_GROUP_BOOKMARK_PREFIX)) <> WORD_NESTED_GROUP_BOOKMARK_PREFIX Then _
            GoTo ContinueBookmark
        Set candidateRange = bookmarkObj.Range
        If recordRange.Start < candidateRange.Start Or _
            recordRange.End > candidateRange.End Then GoTo ContinueBookmark
        Set candidateEntry = VBA.CreateObject("Scripting.Dictionary")
        candidateEntry.Add "Range", candidateRange.Duplicate
        candidateEntry("RecordCount") = 0&
        candidateEntries.Add candidateEntry
ContinueBookmark:
    Next bookmarkObj

    ' Каждый record Range читается один раз и сравнивается только с найденными
    ' родителями. Прежний вложенный обход повторно обращался к Word COM для
    ' полного набора bookmarks по каждой группе.
    For Each recordBookmark In wordDoc.Bookmarks
        If VBA.Left$(VBA.UCase$(VBA.CStr(recordBookmark.Name)), _
            VBA.Len(WORD_RECORD_BOOKMARK_PREFIX)) = WORD_RECORD_BOOKMARK_PREFIX Then
            Set recordRangeObj = recordBookmark.Range
            For Each candidateItem In candidateEntries
                Set candidateEntry = candidateItem
                Set candidateRange = candidateEntry("Range")
                If recordRangeObj.Start >= candidateRange.Start And _
                    recordRangeObj.End <= candidateRange.End Then _
                    candidateEntry("RecordCount") = _
                        VBA.CLng(candidateEntry("RecordCount")) + 1
            Next candidateItem
        End If
    Next recordBookmark

    For Each candidateItem In candidateEntries
        Set candidateEntry = candidateItem
        If VBA.CLng(candidateEntry("RecordCount")) = 1 Then
            Set candidateRange = candidateEntry("Range")
            If candidateRange.End - candidateRange.Start > bestLength Then
                Set private_GetDeleteUndoRange = candidateRange.Duplicate
                bestLength = candidateRange.End - candidateRange.Start
            End If
        End If
    Next candidateItem
End Function

Private Function private_TryBuildExistingResultPath( _
    ByVal orderNo As String, ByRef outTargetPath As String, _
    ByVal showMissingMessage As Boolean _
) As Boolean
    Dim templatePath As String
    templatePath = VBA.Trim$(m_Base.TargetWorkbookPath)
    If VBA.Len(templatePath) = 0 Then
        VBA.MsgBox "PrototypeNew: required profile key 'Export.Word.FilePath' is empty.", _
            VBA.vbExclamation, "PrototypeNew / WORD events"
        Exit Function
    End If
    If Not private_IsAbsolutePath(templatePath) Then _
        templatePath = ThisWorkbook.Path & Application.PathSeparator & templatePath
    outTargetPath = private_BuildResultDocumentPath(templatePath, orderNo)
    If VBA.Len(outTargetPath) = 0 Then Exit Function
    If VBA.Len(VBA.Dir$(outTargetPath, VBA.vbNormal Or VBA.vbReadOnly Or _
        VBA.vbHidden Or VBA.vbSystem)) = 0 Then
        If showMissingMessage Then
            VBA.MsgBox "WORD result document was not found:" & _
                VBA.vbCrLf & outTargetPath, VBA.vbExclamation, _
                "PrsnlEventBuilder / WORD events"
        End If
        Exit Function
    End If
    private_TryBuildExistingResultPath = True
End Function

Private Function private_BuildWordEventCaption(ByVal valueText As String) As String
    valueText = VBA.Replace(valueText, VBA.vbCr, " ")
    valueText = VBA.Replace(valueText, VBA.vbLf, " ")
    valueText = VBA.Replace(valueText, VBA.Chr$(11), " ")
    valueText = VBA.Replace(valueText, VBA.vbTab, " ")
    Do While VBA.InStr(1, valueText, "  ", VBA.vbBinaryCompare) > 0
        valueText = VBA.Replace(valueText, "  ", " ")
    Loop
    valueText = VBA.Trim$(valueText)
    If VBA.Len(valueText) > 180 Then valueText = VBA.Left$(valueText, 177) & "..."
    private_BuildWordEventCaption = valueText
End Function

Private Function private_ShowSupportedBookmarksAsMarkers( _
    ByVal wordDoc As Object, _
    ByRef convertedCount As Long _
) As Boolean
    Dim bookmarkIndex As Long
    Dim bookmarkObj As Object
    Dim bookmarkName As String
    Dim bookmarkNames() As String
    Dim bookmarkStarts() As Long
    Dim bookmarkEnds() As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim positionCount As Long
    Dim positions() As Long
    Dim used() As Boolean
    Dim bestIndex As Long
    Dim insertRange As Object
    Dim markerText As String

    For bookmarkIndex = 1 To wordDoc.Bookmarks.Count
        Set bookmarkObj = wordDoc.Bookmarks.Item(bookmarkIndex)
        bookmarkName = VBA.CStr(bookmarkObj.Name)
        If private_IsSupportedBookmarkName(bookmarkName) Then
            convertedCount = convertedCount + 1
            If convertedCount = 1 Then
                ReDim bookmarkNames(1 To 1)
                ReDim bookmarkStarts(1 To 1)
                ReDim bookmarkEnds(1 To 1)
            Else
                ReDim Preserve bookmarkNames(1 To convertedCount)
                ReDim Preserve bookmarkStarts(1 To convertedCount)
                ReDim Preserve bookmarkEnds(1 To convertedCount)
            End If
            bookmarkNames(convertedCount) = bookmarkName
            bookmarkStarts(convertedCount) = bookmarkObj.Range.Start
            bookmarkEnds(convertedCount) = bookmarkObj.Range.End
        End If
    Next bookmarkIndex
    If convertedCount = 0 Then
        VBA.MsgBox "В документе нет поддерживаемых WORD-закладок PEB/PEM/PEG/PEN.", _
            VBA.vbInformation, "PrsnlEventBuilder / WORD bookmarks"
        Exit Function
    End If

    ' Собираем уникальные границы. Вставка справа налево сохраняет исходные
    ' координаты, а единая строка для общей границы не создаёт пересечений:
    ' сначала закрываются внутренние диапазоны, затем открываются внешние.
    For i = 1 To convertedCount
        For k = 1 To 2
            If k = 1 Then j = bookmarkStarts(i) Else j = bookmarkEnds(i)
            bestIndex = 0
            For bestIndex = 1 To positionCount
                If positions(bestIndex) = j Then Exit For
            Next bestIndex
            If bestIndex > positionCount Then bestIndex = 0
            If bestIndex = 0 Then
                positionCount = positionCount + 1
                If positionCount = 1 Then
                    ReDim positions(1 To 1)
                Else
                    ReDim Preserve positions(1 To positionCount)
                End If
                positions(positionCount) = j
            End If
        Next k
    Next i
    For i = 1 To positionCount - 1
        For j = i + 1 To positionCount
            If positions(j) > positions(i) Then
                k = positions(i): positions(i) = positions(j): positions(j) = k
            End If
        Next j
    Next i

    For i = 1 To positionCount
        markerText = VBA.vbNullString
        ReDim used(1 To convertedCount)
        For k = 1 To convertedCount
            bestIndex = 0
            For j = 1 To convertedCount
                If Not used(j) And bookmarkEnds(j) = positions(i) Then
                    If bestIndex = 0 Then
                        bestIndex = j
                    ElseIf bookmarkStarts(j) > bookmarkStarts(bestIndex) Then
                        bestIndex = j
                    ElseIf bookmarkStarts(j) = bookmarkStarts(bestIndex) Then
                        If private_GetBookmarkNestingPriority(bookmarkNames(j)) > _
                            private_GetBookmarkNestingPriority( _
                                bookmarkNames(bestIndex)) Then bestIndex = j
                    End If
                End If
            Next j
            If bestIndex = 0 Then Exit For
            used(bestIndex) = True
            markerText = markerText & WORD_BOOKMARK_MARKER_END & _
                bookmarkNames(bestIndex) & WORD_BOOKMARK_MARKER_SUFFIX
        Next k
        ReDim used(1 To convertedCount)
        For k = 1 To convertedCount
            bestIndex = 0
            For j = 1 To convertedCount
                If Not used(j) And bookmarkStarts(j) = positions(i) Then
                    If bestIndex = 0 Then
                        bestIndex = j
                    ElseIf bookmarkEnds(j) > bookmarkEnds(bestIndex) Then
                        bestIndex = j
                    ElseIf bookmarkEnds(j) = bookmarkEnds(bestIndex) Then
                        If private_GetBookmarkNestingPriority(bookmarkNames(j)) < _
                            private_GetBookmarkNestingPriority( _
                                bookmarkNames(bestIndex)) Then bestIndex = j
                    End If
                End If
            Next j
            If bestIndex = 0 Then Exit For
            used(bestIndex) = True
            markerText = markerText & WORD_BOOKMARK_MARKER_BEGIN & _
                bookmarkNames(bestIndex) & WORD_BOOKMARK_MARKER_SUFFIX
        Next k
        Set insertRange = wordDoc.Range(positions(i), positions(i))
        insertRange.Text = markerText
    Next i
    For i = 1 To convertedCount
        If wordDoc.Bookmarks.Exists(bookmarkNames(i)) Then _
            wordDoc.Bookmarks(bookmarkNames(i)).Delete
    Next i
    private_ShowSupportedBookmarksAsMarkers = True
End Function

' Чем больше значение, тем глубже закладка в иерархии WORD:
' PEG (внешняя группа) -> PEN (вложенная группа) -> PEB (пункт) -> PEM (метаданные).
' При общей границе внешние диапазоны открываются первыми, а закрываются последними.
Private Function private_GetBookmarkNestingPriority( _
    ByVal bookmarkName As String _
) As Long
    Select Case VBA.UCase$(VBA.Left$(bookmarkName, 4))
        Case WORD_GROUP_BOOKMARK_PREFIX
            private_GetBookmarkNestingPriority = 1
        Case WORD_NESTED_GROUP_BOOKMARK_PREFIX
            private_GetBookmarkNestingPriority = 2
        Case WORD_RECORD_BOOKMARK_PREFIX
            private_GetBookmarkNestingPriority = 3
        Case WORD_METADATA_BOOKMARK_PREFIX
            private_GetBookmarkNestingPriority = 4
    End Select
End Function

Private Function private_RestoreSupportedBookmarks( _
    ByVal wordDoc As Object, _
    ByRef convertedCount As Long _
) As Boolean
    Dim bookmarkName As String
    Dim contentRanges As Collection
    Dim bookmarkNames As Collection
    Dim beginStarts As Object
    Dim beginEnds As Object
    Dim endStarts As Object
    Dim endEnds As Object
    Dim markerRanges As Collection
    Dim beginStart As Long
    Dim beginEnd As Long
    Dim endStart As Long
    Dim endEnd As Long
    Dim contentRange As Object
    Dim i As Long

    Set contentRanges = New Collection
    Set bookmarkNames = New Collection
    Set markerRanges = New Collection
    Set beginStarts = VBA.CreateObject("Scripting.Dictionary")
    Set beginEnds = VBA.CreateObject("Scripting.Dictionary")
    Set endStarts = VBA.CreateObject("Scripting.Dictionary")
    Set endEnds = VBA.CreateObject("Scripting.Dictionary")
    beginStarts.CompareMode = VBA.vbTextCompare
    beginEnds.CompareMode = VBA.vbTextCompare
    endStarts.CompareMode = VBA.vbTextCompare
    endEnds.CompareMode = VBA.vbTextCompare
    If Not private_TryScanWordBookmarkMarkers( _
        wordDoc, beginStarts, beginEnds, endStarts, endEnds, _
        bookmarkNames, markerRanges) Then GoTo InvalidMarkers
    If beginStarts.Count <> endStarts.Count Then GoTo InvalidMarkers

    convertedCount = bookmarkNames.Count
    For i = 1 To bookmarkNames.Count
        bookmarkName = VBA.CStr(bookmarkNames(i))
        If Not endStarts.Exists(bookmarkName) Then GoTo InvalidMarkers
        beginStart = VBA.CLng(beginStarts(bookmarkName))
        beginEnd = VBA.CLng(beginEnds(bookmarkName))
        endStart = VBA.CLng(endStarts(bookmarkName))
        endEnd = VBA.CLng(endEnds(bookmarkName))
        If beginEnd <= endStart Then
            Set contentRange = wordDoc.Range(beginEnd, endStart)
        ElseIf endEnd <= beginStart Then
            ' Схлопнутая закладка записывается как END -> BEGIN в одной позиции.
            ' После удаления обоих маркеров этот anchor снова станет Range(Start, Start).
            Set contentRange = wordDoc.Range(endStart, endStart)
        Else
            GoTo InvalidMarkers
        End If
        contentRanges.Add contentRange
    Next i

    If Not private_TryDeleteWordMarkerRangesRightToLeft( _
        markerRanges) Then GoTo InvalidMarkers
    If VBA.InStr(1, VBA.CStr(wordDoc.Content.Text), _
        WORD_BOOKMARK_MARKER_ROOT, VBA.vbTextCompare) > 0 Then _
        GoTo InvalidMarkers
    For i = 1 To convertedCount
        wordDoc.Bookmarks.Add VBA.CStr(bookmarkNames(i)), contentRanges(i)
    Next i
    private_RestoreSupportedBookmarks = True
    Exit Function

InvalidMarkers:
    VBA.MsgBox _
        "Набор WORD-маркеров повреждён: для каждой начальной метки должна " & _
        "существовать парная конечная метка. Восстановление отменено.", _
        VBA.vbExclamation, "PrsnlEventBuilder / WORD bookmarks"
End Function

Private Function private_TryScanWordBookmarkMarkers( _
    ByVal wordDoc As Object, _
    ByVal beginStarts As Object, _
    ByVal beginEnds As Object, _
    ByVal endStarts As Object, _
    ByVal endEnds As Object, _
    ByVal bookmarkNames As Collection, _
    ByVal markerRanges As Collection _
) As Boolean
    Dim scanStart As Long
    Dim markerStartRange As Object
    Dim markerRange As Object
    Dim markerProbeRange As Object
    Dim markerProbeText As String
    Dim suffixOffset As Long
    Dim markerEnd As Long
    Dim markerText As String
    Dim bookmarkName As String
    Dim isBeginMarker As Boolean
    Dim nameStart As Long
    Dim nameLength As Long

    If wordDoc Is Nothing Then Exit Function
    scanStart = wordDoc.Content.Start
    Do While scanStart < wordDoc.Content.End
        If Not private_TryFindWordText(wordDoc.Range( _
            scanStart, wordDoc.Content.End), WORD_BOOKMARK_MARKER_ROOT, _
            markerStartRange) Then Exit Do
        ' Имя Word bookmark ограничено 40 символами. Читаем небольшой фрагмент
        ' возле найденного префикса и определяем конец обычным InStr, не делая
        ' второй COM Find для каждого из 122 маркеров.
        markerEnd = markerStartRange.Start + 96
        If markerEnd > wordDoc.Content.End Then markerEnd = wordDoc.Content.End
        Set markerProbeRange = wordDoc.Range(markerStartRange.Start, markerEnd)
        markerProbeText = VBA.CStr(markerProbeRange.Text)
        suffixOffset = VBA.InStr(1, markerProbeText, _
            WORD_BOOKMARK_MARKER_SUFFIX, VBA.vbBinaryCompare)
        If suffixOffset <= 0 Then Exit Function
        markerEnd = markerStartRange.Start + suffixOffset - 1 + _
            VBA.Len(WORD_BOOKMARK_MARKER_SUFFIX)
        Set markerRange = wordDoc.Range(markerStartRange.Start, markerEnd)
        markerText = VBA.CStr(markerRange.Text)
        If VBA.Left$(markerText, VBA.Len(WORD_BOOKMARK_MARKER_BEGIN)) = _
            WORD_BOOKMARK_MARKER_BEGIN Then
            isBeginMarker = True
            nameStart = VBA.Len(WORD_BOOKMARK_MARKER_BEGIN) + 1
        ElseIf VBA.Left$(markerText, VBA.Len(WORD_BOOKMARK_MARKER_END)) = _
            WORD_BOOKMARK_MARKER_END Then
            isBeginMarker = False
            nameStart = VBA.Len(WORD_BOOKMARK_MARKER_END) + 1
        Else
            Exit Function
        End If
        nameLength = VBA.Len(markerText) - nameStart - _
            VBA.Len(WORD_BOOKMARK_MARKER_SUFFIX) + 1
        If nameLength <= 0 Then Exit Function
        bookmarkName = VBA.Mid$(markerText, nameStart, nameLength)
        If Not private_IsSupportedBookmarkName(bookmarkName) Then Exit Function
        If isBeginMarker Then
            If beginStarts.Exists(bookmarkName) Then Exit Function
            beginStarts.Add bookmarkName, markerRange.Start
            beginEnds.Add bookmarkName, markerRange.End
            bookmarkNames.Add bookmarkName
        Else
            If endStarts.Exists(bookmarkName) Then Exit Function
            endStarts.Add bookmarkName, markerRange.Start
            endEnds.Add bookmarkName, markerRange.End
        End If
        markerRanges.Add markerRange.Duplicate
        scanStart = markerRange.End
    Loop
    private_TryScanWordBookmarkMarkers = (markerRanges.Count > 0)
End Function

Private Function private_TryDeleteWordMarkerRangesRightToLeft( _
    ByVal markerRanges As Collection _
) As Boolean
    Dim markerIndex As Long
    Dim markerRange As Object

    If markerRanges Is Nothing Then Exit Function
    For markerIndex = markerRanges.Count To 1 Step -1
        Set markerRange = markerRanges.Item(markerIndex)
        If markerRange Is Nothing Then Exit Function
        markerRange.Delete
    Next markerIndex
    private_TryDeleteWordMarkerRangesRightToLeft = True
End Function

Private Function private_IsSupportedBookmarkName(ByVal bookmarkName As String) As Boolean
    Dim prefixText As String
    prefixText = VBA.UCase$(VBA.Left$(bookmarkName, 4))
    private_IsSupportedBookmarkName = _
        (prefixText = WORD_RECORD_BOOKMARK_PREFIX Or _
         prefixText = WORD_METADATA_BOOKMARK_PREFIX Or _
         prefixText = WORD_GROUP_BOOKMARK_PREFIX Or _
         prefixText = WORD_NESTED_GROUP_BOOKMARK_PREFIX)
End Function


Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & TypeName(Me) & ".Class_Terminate"
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
    Optional ByVal profileConfigTable As obj_ConfigTable = Nothing, _
    Optional ByVal exporterCfgDataProvider As obj_PEB_ExptrCfgDataPrvdr = Nothing _
) As Boolean
    Dim exporterCfgDataProviderConfigTable As obj_ConfigTable

    m_IsDisposed = False
    m_OwnsExporterCfgDataProvider = False
    Set m_Base = New obj_DataExporterBase
    Set m_TemplateParser = New obj_WordResultTplParser
    ' Target-настройки берём из exporter config, а профильные источники
    ' Personnel/Movement — из полной таблицы профиля, если controller её передал.
    Set exporterCfgDataProviderConfigTable = configTable
    If Not profileConfigTable Is Nothing Then Set exporterCfgDataProviderConfigTable = profileConfigTable

    If Not m_Base.Initialize(configTable, "WORD", "PrototypeNew / WORD export") Then Exit Function
    If Not m_TemplateParser.Initialize(WORD_RESULT_TEMPLATES_REL_PATH) Then Exit Function
    If exporterCfgDataProvider Is Nothing Then
        Set m_ExporterCfgDataProvider = New obj_PEB_ExptrCfgDataPrvdr
        m_OwnsExporterCfgDataProvider = True
        If Not m_ExporterCfgDataProvider.Initialize(exporterCfgDataProviderConfigTable) Then Exit Function
    Else
        ' Страница уже использует этот provider для «Історія руху». Совместное
        ' владение исключает второй ADO handle и попытку пересоздать занятый
        ' Movement snapshot при формировании WORD preview.
        Set m_ExporterCfgDataProvider = exporterCfgDataProvider
    End If

    Initialize = True
End Function

Public Sub Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    If Not m_Base Is Nothing Then m_Base.Dispose
    If Not m_TemplateParser Is Nothing Then m_TemplateParser.Dispose
    If m_OwnsExporterCfgDataProvider Then
        If Not m_ExporterCfgDataProvider Is Nothing Then m_ExporterCfgDataProvider.Dispose
    End If
    Set m_Base = Nothing
    Set m_TemplateParser = Nothing
    Set m_ExporterCfgDataProvider = Nothing
    m_OwnsExporterCfgDataProvider = False
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
    Dim recordText As String
    Dim templateId As String
    Dim templateHash As String
    Dim recordIpn As String
    Dim exportValidationError As String
    Dim latestMovementTvoChain As Collection
    Dim ignoredLatestMovementRecord As Object
    Dim builderData As obj_PrsnlEvntBuilderData
    Dim validationEnabled As Boolean
    Dim movementWasPrevalidated As Boolean
    Dim hasGrouping As Boolean
    Dim groupByText As String
    Dim groupOrderText As String
    Dim groupKeyText As String
    Dim groupHeaderText As String
    Dim nestedGroupByText As String
    Dim nestedGroupOrderText As String
    Dim nestedGroupKeyText As String
    Dim nestedGroupHeaderText As String
    Dim groupByParts As Variant
    Dim groupOrderParts As Variant
    Dim usePreparedPreview As Boolean
    Dim writeToWord As Boolean
    Dim documentFilepath As String

    If m_IsDisposed Then
        VBA.MsgBox "PrototypeNew: WORD exporter is disposed.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    m_LastBuildHasWarnings = False
    If Not m_Base.TryGetMainSourceTable(sourceTables, sourceTable) Then Exit Function

    sectionTypeText = private_GetContextText(context, CONTEXT_SECTION_TYPE)
    If VBA.Len(sectionTypeText) = 0 Then sectionTypeText = VBA.Trim$(sourceTable.SectionTitle)
    If VBA.Len(sectionTypeText) = 0 Then
        VBA.MsgBox "PrototypeNew: WORD export requires SectionType in export context or source table SectionTitle.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    ' WORD validation defaults to disabled when the context key is absent.
    validationEnabled = private_GetContextBoolean(context, CONTEXT_VALIDATION_ENABLED)
    movementWasPrevalidated = private_GetContextBoolean( _
        context, CONTEXT_MOVEMENT_PREVALIDATED)
    ' Если Movement для тех же ІПН и секции уже успешно экспортирован,
    ' controller передаёт receipt выполненной проверки. Иначе WORD выполняет
    ' обычную самостоятельную валидацию текущего snapshot.
    If Not m_ExporterCfgDataProvider.IsExportAllowed( _
        sourceTable, sectionTypeText, exportValidationError, _
        latestMovementTvoChain, ignoredLatestMovementRecord, _
        validationEnabled, False, movementWasPrevalidated) Then
        VBA.MsgBox exportValidationError, VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    If Not private_TryEnrichMainSourceTableForWord(sourceTable, context) Then Exit Function
    If Not private_TryEnrichPreviousVacationTicketForWord( _
        sourceTable, sectionTypeText) Then Exit Function
    If Not private_TryNormalizeDocumentNotesForWord(sourceTables, sectionTypeText) Then Exit Function
    If Not private_TryEnrichMetaDocumentDatesForWord(sourceTables) Then Exit Function
    If Not private_TryAppendMovementTvoTablesForReturn(sourceTables, sourceTable, sectionTypeText, latestMovementTvoChain) Then Exit Function
    If Not private_TryEnrichMetaTvoTablesForWord(sourceTables) Then Exit Function
    Set namedCollections = private_BuildNamedLoopCollections(sourceTables)
    If namedCollections Is Nothing Then Exit Function
    If VBA.Len(private_GetContextText( _
        context, CONTEXT_MANUAL_ORDER_NO)) = 0 Then
        VBA.MsgBox "PrototypeNew: documentFilepath requires a non-empty " & _
            "ManualOrderNo in export context.", VBA.vbExclamation, _
            "PrototypeNew / WORD export"
        Exit Function
    End If
    If Not m_TemplateParser.TryRenderDocumentFilepath( _
        sectionTypeText, sourceTables, namedCollections, context, _
        documentFilepath) Then Exit Function
    Set builderData = New obj_PrsnlEvntBuilderData
    If Not builderData.TryResolveWordTemplateId(sectionTypeText, templateId) Then
        VBA.MsgBox "PrototypeNew: WORD result template is not mapped for section: " & sectionTypeText, VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    If Not m_TemplateParser.TryGetTemplateHashByName( _
        templateId, templateHash) Then Exit Function
    previewText = private_GetContextText(context, CONTEXT_WORD_PREVIEW_TEXT)
    writeToWord = private_GetContextBoolean(context, "WriteToWord")
    usePreparedPreview = (writeToWord And _
        VBA.Len(VBA.Trim$(previewText)) > 0)
    If Not usePreparedPreview Then
        If Not m_TemplateParser.TryRenderForTemplateId( _
            templateId, _
            sectionTypeText, _
            sourceTables, _
            namedCollections, _
            recordText) Then Exit Function
    End If
    If Not m_TemplateParser.TryGetGroupingDefinition( _
        templateId, hasGrouping, groupByText, groupOrderText) Then Exit Function

    ' Заголовки групп рендерятся и для preview, и для фактической вставки.
    ' Preview показывает их линейно в порядке DSL, а Word позже решает по
    ' закладкам, какие из них действительно требуется добавить в документ.
    If hasGrouping Then
        groupByParts = VBA.Split(groupByText, ";")
        groupOrderParts = VBA.Split(groupOrderText, ";")
        groupByText = VBA.Trim$(VBA.CStr(groupByParts(0)))
        groupOrderText = VBA.Trim$(VBA.CStr(groupOrderParts(0)))
        If Not private_TryGetMainTableValue( _
            sourceTable, groupByText, groupKeyText) Then Exit Function
        groupKeyText = VBA.Trim$(groupKeyText)
        If VBA.Len(groupKeyText) = 0 Then
            VBA.MsgBox "PrototypeNew: grouped WORD template '" & templateId & _
                "' requires a non-empty value [" & groupByText & "].", _
                VBA.vbExclamation, "PrototypeNew / WORD export"
            Exit Function
        End If
        If Not m_TemplateParser.TryRenderGroupHeaderForTemplateId( _
            templateId, groupByText, sectionTypeText, sourceTables, namedCollections, _
            groupHeaderText) Then Exit Function
        If UBound(groupByParts) > 1 Then
            VBA.MsgBox "PrototypeNew: WORD export currently supports no more " & _
                "than two grouping levels: " & templateId, _
                VBA.vbExclamation, "PrototypeNew / WORD export"
            Exit Function
        End If
        If UBound(groupByParts) = 1 Then
            nestedGroupByText = VBA.Trim$(VBA.CStr(groupByParts(1)))
            nestedGroupOrderText = VBA.Trim$(VBA.CStr(groupOrderParts(1)))
            If Not private_TryGetMainTableValue( _
                sourceTable, nestedGroupByText, nestedGroupKeyText) Then Exit Function
            nestedGroupKeyText = VBA.Trim$(nestedGroupKeyText)
            If VBA.Len(nestedGroupKeyText) = 0 Then
                VBA.MsgBox "PrototypeNew: grouped WORD template '" & templateId & _
                    "' requires a non-empty value [" & nestedGroupByText & "].", _
                    VBA.vbExclamation, "PrototypeNew / WORD export"
                Exit Function
            End If
            If Not m_TemplateParser.TryRenderGroupHeaderForTemplateId( _
                templateId, nestedGroupByText, sectionTypeText, sourceTables, _
                namedCollections, nestedGroupHeaderText) Then Exit Function
        End If
    End If

    If usePreparedPreview Then
        ' Banner содержит линейное preview вместе с groupHeader. Перед поздней
        ' Word-группировкой отделяем от него только редактируемое тело записи.
        If Not private_TryExtractRecordTextFromPreview( _
            templateId, previewText, groupHeaderText, _
            nestedGroupHeaderText, recordText) Then Exit Function
    Else
        previewText = groupHeaderText & nestedGroupHeaderText & recordText
    End If
    If Not private_TrySetContextText( _
        context, CONTEXT_WORD_PREVIEW_TEXT, previewText) Then Exit Function

    ' CTRL+3 только возвращает линейное preview. CTRL+4 передаёт WriteToWord=True,
    ' после чего groupHeader и recordText вставляются по отдельным правилам.
    If writeToWord Then
        If m_LastBuildHasWarnings Or _
           VBA.InStr(1, previewText, "[[preview-warning]]", _
               VBA.vbTextCompare) > 0 Or _
           VBA.InStr(1, previewText, "{[", VBA.vbBinaryCompare) > 0 Then
            If VBA.MsgBox( _
                "Не всі значення вдалося знайти або відмінити." & _
                VBA.vbCrLf & "Проблемні фрагменти залишені у результаті " & _
                "та позначені жовтим у preview." & VBA.vbCrLf & _
                VBA.vbCrLf & "Продовжити експорт WORD?", _
                VBA.vbQuestion Or VBA.vbYesNo Or VBA.vbDefaultButton2, _
                "PrsnlEventBuilder / Неповний результат") <> VBA.vbYes Then _
                Exit Function
        End If
        If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_IPN, recordIpn) Then
            VBA.MsgBox "PrototypeNew: WORD export requires IPN to create a record bookmark.", VBA.vbExclamation, "PrototypeNew / WORD export"
            Exit Function
        End If
        recordIpn = VBA.Trim$(recordIpn)
        If VBA.Len(recordIpn) = 0 Then
            VBA.MsgBox "PrototypeNew: WORD export requires a non-empty IPN to create a record bookmark.", VBA.vbExclamation, "PrototypeNew / WORD export"
            Exit Function
        End If
        If Not private_TryAppendBeforeWordEndAnchor( _
            templateId, templateHash, recordIpn, recordText, _
            groupKeyText, groupOrderText, groupHeaderText, _
            nestedGroupKeyText, nestedGroupOrderText, nestedGroupHeaderText, _
            documentFilepath) Then Exit Function
    End If

    Export = True
End Function

Private Function private_TryExtractRecordTextFromPreview( _
    ByVal templateId As String, _
    ByVal previewText As String, _
    ByVal groupHeaderText As String, _
    ByVal nestedGroupHeaderText As String, _
    ByRef outRecordText As String _
) As Boolean
    Dim headerPrefix As String
    Dim normalizedHeaderPrefix As String
    Dim normalizedPreviewText As String

    outRecordText = previewText
    ' Текст, прочитанный обратно из Excel Banner, уже не содержит служебных
    ' inline-маркеров, поэтому сравниваем его с plain-вариантом заголовков,
    ' а не с исходной строкой template renderer.
    If Not private_TryStripPreviewInlineMarkers( _
        groupHeaderText & nestedGroupHeaderText, headerPrefix) Then Exit Function
    If VBA.Len(headerPrefix) = 0 Then
        private_TryExtractRecordTextFromPreview = True
        Exit Function
    End If

    ' Template renderer формирует переносы как CRLF. После записи такого текста
    ' в ячейку Excel и обратного чтения через Range.Value2 Excel возвращает
    ' внутренние переносы как LF. Визуально заголовки одинаковы, но бинарное
    ' сравнение считает CRLF и LF разными; кроме того, CRLF занимает два символа,
    ' поэтому Len(headerPrefix) нельзя применять к исходному previewText.
    '
    ' Нормализуем обе стороны до LF и отделяем тело записи из этой же
    ' нормализованной строки. Так длина проверенного prefix точно совпадает с
    ' позицией, с которой начинается редактируемый record text.
    normalizedHeaderPrefix = private_NormalizePreviewLineBreaks(headerPrefix)
    normalizedPreviewText = private_NormalizePreviewLineBreaks(previewText)

    If VBA.Len(normalizedPreviewText) < VBA.Len(normalizedHeaderPrefix) Or _
        VBA.StrComp( _
            VBA.Left$(normalizedPreviewText, VBA.Len(normalizedHeaderPrefix)), _
            normalizedHeaderPrefix, _
            VBA.vbBinaryCompare) <> 0 Then
        VBA.MsgBox _
            "PrototypeNew: grouped preview headers were changed for template '" & _
            templateId & "'. Edit only the record text below the group headers.", _
            VBA.vbExclamation, _
            "PrototypeNew / WORD preview"
        Exit Function
    End If

    outRecordText = VBA.Mid$( _
        normalizedPreviewText, VBA.Len(normalizedHeaderPrefix) + 1)
    private_TryExtractRecordTextFromPreview = True
End Function

Private Function private_NormalizePreviewLineBreaks( _
    ByVal valueText As String _
) As String
    ' Сначала схлопываем CRLF как единый перенос. Если сперва заменить только
    ' CR, исходный CRLF превратится в два LF и снова изменит длину prefix.
    valueText = VBA.Replace(valueText, VBA.vbCrLf, VBA.vbLf)
    valueText = VBA.Replace(valueText, VBA.vbCr, VBA.vbLf)
    private_NormalizePreviewLineBreaks = valueText
End Function

Private Function private_TryAppendBeforeWordEndAnchor( _
    ByVal templateId As String, _
    ByVal templateHash As String, _
    ByVal recordIpn As String, _
    ByVal renderedText As String, _
    Optional ByVal groupKeyText As String = "", _
    Optional ByVal groupOrderText As String = "", _
    Optional ByVal groupHeaderText As String = "", _
    Optional ByVal nestedGroupKeyText As String = "", _
    Optional ByVal nestedGroupOrderText As String = "", _
    Optional ByVal nestedGroupHeaderText As String = "", _
    Optional ByVal documentFilepath As String = "" _
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
    Dim groupBookmarkName As String
    Dim groupRange As Object
    Dim groupStart As Long
    Dim groupEnd As Long
    Dim nestedGroupBookmarkName As String
    Dim nestedGroupRange As Object
    Dim nestedGroupStart As Long
    Dim nestedGroupContentStart As Long
    Dim plainGroupHeaderText As String
    Dim plainNestedGroupHeaderText As String
    Dim undoAction As obj_PEB_ExportUndoAction
    Dim undoActionReady As Boolean

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

    targetPath = VBA.Trim$(documentFilepath)
    If VBA.Len(targetPath) = 0 Then
        VBA.MsgBox "PrototypeNew: rendered documentFilepath is empty.", _
            VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    If Not private_IsAbsolutePath(targetPath) Then _
        targetPath = ThisWorkbook.Path & Application.PathSeparator & targetPath

    ' Export.Word.FilePath указывает на неизменяемый входной DOCX. Полный путь
    ' результата уже вычислен DSL-тегом documentFilepath; первый экспорт
    ' копирует в него шаблон, последующие дописывают пункты по якорям.
    If VBA.Len(VBA.Dir$(targetPath, VBA.vbNormal Or VBA.vbReadOnly Or VBA.vbHidden Or VBA.vbSystem)) = 0 Then
        VBA.FileCopy templatePath, targetPath
    End If

    beginMarker = WORD_ANCHOR_PREFIX & VBA.Trim$(templateId) & WORD_ANCHOR_BEGIN_SUFFIX
    endMarker = WORD_ANCHOR_PREFIX & VBA.Trim$(templateId) & WORD_ANCHOR_END_SUFFIX
    If Not private_TryStripPreviewInlineMarkers( _
        renderedText, plainRenderedText) Then Exit Function
    plainRenderedText = private_NormalizeWordParagraphBreaks(plainRenderedText)
    bookmarkName = private_BuildRecordBookmarkName(templateHash, recordIpn)
    If VBA.Len(bookmarkName) = 0 Then
        VBA.MsgBox "PrototypeNew: failed to build a WORD bookmark for IPN '" & recordIpn & "'.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(groupKeyText)) > 0 Then
        groupBookmarkName = private_BuildGroupBookmarkName( _
            templateHash, groupKeyText, groupOrderText)
        If VBA.Len(groupBookmarkName) = 0 Then Exit Function
        If Not private_TryStripPreviewInlineMarkers( _
            groupHeaderText, plainGroupHeaderText) Then Exit Function
        groupHeaderText = private_NormalizeWordParagraphBreaks( _
            plainGroupHeaderText)
        If VBA.Len(VBA.Trim$(nestedGroupKeyText)) > 0 Then
            nestedGroupBookmarkName = private_BuildNestedGroupBookmarkName( _
                groupBookmarkName, nestedGroupKeyText)
            If VBA.Len(nestedGroupBookmarkName) = 0 Then Exit Function
            If Not private_TryStripPreviewInlineMarkers( _
                nestedGroupHeaderText, plainNestedGroupHeaderText) Then Exit Function
            nestedGroupHeaderText = private_NormalizeWordParagraphBreaks( _
                plainNestedGroupHeaderText)
        End If
    End If

    If Not rt_WordExportRuntime.fn_TryAcquireWordDocument( _
        targetPath, wordApp, wordDoc, documentOpened) Then Exit Function

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
    ' Текст существующих заголовков не анализируется. Иерархию хранят закладки:
    ' PEG — внешняя группа (дата), PEN — вложенная группа (больница),
    ' PEB — отдельная запись человека. Begin/End ограничивают всю секцию.
    insertedStart = endRange.Start
    If VBA.Len(groupBookmarkName) > 0 Then
        ' Undo старых версий мог оставить пустые PEG/PEN. Такая закладка не
        ' подтверждает существование заголовка и должна считаться отсутствующей.
        private_DeleteCollapsedWordBookmarkIfExists _
            wordDoc, nestedGroupBookmarkName
        private_DeleteCollapsedWordBookmarkIfExists _
            wordDoc, groupBookmarkName
        If wordDoc.Bookmarks.Exists(groupBookmarkName) Then
            Set groupRange = wordDoc.Bookmarks(groupBookmarkName).Range
            If groupRange.Start < beginRange.End Or groupRange.End > endRange.Start Then
                VBA.MsgBox "PrototypeNew: WORD group bookmark is outside its section: " & _
                    groupBookmarkName, VBA.vbExclamation, "PrototypeNew / WORD export"
                GoTo CleanFail
            End If
            groupStart = groupRange.Start
            groupEnd = groupRange.End
            ' Вложенные группы начинаются после заголовка родительской даты.
            ' Старые версии включали этот заголовок в первую PEN-закладку,
            ' поэтому вычисляем безопасную нижнюю границу и для таких документов.
            nestedGroupContentStart = groupStart + VBA.Len(groupHeaderText)
            If nestedGroupContentStart > groupEnd Then
                nestedGroupContentStart = groupStart
            End If
            ' Существующая PEN-группа означает, что оба заголовка уже выведены.
            ' Новую запись вставляем в её End и затем расширяем PEN и PEG.
            If VBA.Len(nestedGroupBookmarkName) > 0 And _
                wordDoc.Bookmarks.Exists(nestedGroupBookmarkName) Then
                Set nestedGroupRange = wordDoc.Bookmarks( _
                    nestedGroupBookmarkName).Range
                If nestedGroupRange.Start < groupRange.Start Or _
                    nestedGroupRange.End > groupRange.End Then
                    VBA.MsgBox "PrototypeNew: nested WORD group bookmark is " & _
                        "outside its parent: " & nestedGroupBookmarkName, _
                        VBA.vbExclamation, "PrototypeNew / WORD export"
                    GoTo CleanFail
                End If
                nestedGroupStart = nestedGroupRange.Start
                If nestedGroupStart < nestedGroupContentStart Then
                    nestedGroupStart = nestedGroupContentStart
                End If
                insertedStart = nestedGroupRange.End
                wordDoc.Bookmarks(nestedGroupBookmarkName).Delete
            Else
                ' Внешняя дата уже существует, но больницы внутри неё ещё нет:
                ' вставляем только вложенный groupHeader и первую запись.
                insertedStart = groupRange.End
                If VBA.Len(nestedGroupBookmarkName) > 0 Then
                    If Not private_TryFindNestedGroupInsertPosition( _
                        wordDoc, groupBookmarkName, nestedGroupBookmarkName, _
                        nestedGroupOrderText, groupRange.Start, _
                        nestedGroupContentStart, groupRange.End, insertedStart) Then _
                        GoTo CleanFail
                    nestedGroupStart = insertedStart
                    plainRenderedText = nestedGroupHeaderText & plainRenderedText
                End If
            End If
            wordDoc.Bookmarks(groupBookmarkName).Delete
        Else
            If Not private_TryFindGroupInsertPosition( _
                wordDoc, templateHash, groupBookmarkName, groupOrderText, _
                beginRange.End, endRange.Start, insertedStart) Then GoTo CleanFail
            ' Нет даже внешней группы: единым блоком вставляются заголовок даты,
            ' заголовок больницы и первая запись человека.
            If VBA.Len(nestedGroupBookmarkName) > 0 Then
                plainRenderedText = groupHeaderText & nestedGroupHeaderText & _
                    plainRenderedText
                nestedGroupStart = insertedStart + VBA.Len(groupHeaderText)
            Else
                plainRenderedText = groupHeaderText & plainRenderedText
            End If
            groupStart = insertedStart
        End If
    End If

    ' В плоском режиме текст по-прежнему вставляется непосредственно перед End.
    ' В grouped-режиме заголовок создаётся один раз, а записи дописываются внутрь
    ' закладки своей группы.
    Set insertRange = wordDoc.Range(insertedStart, insertedStart)
    insertRange.Text = plainRenderedText
    insertedEnd = insertedStart + VBA.Len(plainRenderedText)

    Set insertRange = wordDoc.Range(insertedStart, insertedEnd)
    ' Do not inherit highlight from a neighbouring anchor or an older export.
    insertRange.HighlightColorIndex = 0
    If Not private_TryTrimExpandedRecordBookmarkAtInsertion( _
        wordDoc, insertedStart) Then GoTo CleanFail
    wordDoc.Bookmarks.Add bookmarkName, insertRange
    If VBA.Len(groupBookmarkName) > 0 Then
        If groupEnd > 0 Then
            groupEnd = groupEnd + VBA.Len(plainRenderedText)
        Else
            groupEnd = insertedEnd
        End If
        Set groupRange = wordDoc.Range(groupStart, groupEnd)
        wordDoc.Bookmarks.Add groupBookmarkName, groupRange
        If VBA.Len(nestedGroupBookmarkName) > 0 Then
            Set nestedGroupRange = wordDoc.Range(nestedGroupStart, insertedEnd)
            wordDoc.Bookmarks.Add nestedGroupBookmarkName, nestedGroupRange
        End If
    End If

    ' Bookmark остаётся точным локатором конкретной вставки. Undo action
    ' хранит также позицию и текст для симметричного redo,
    ' но не удерживает Word Document/Range после закрытия файла.
    Set undoAction = New obj_PEB_ExportUndoAction
    undoActionReady = undoAction.InitializeWord( _
        targetPath, bookmarkName, insertedStart, plainRenderedText, _
        VBA.vbNullString, groupBookmarkName, groupStart, _
        nestedGroupBookmarkName, nestedGroupStart)
    If Not undoActionReady Then GoTo CleanFail

    wordDoc.Save
    If documentOpened Then
        wordDoc.Close False
        documentOpened = False
    End If
    If Not rt_UndoManager.fn_PushExecutedAction(undoAction) Then
        rt_Messaging.fn_ShowStatusBarWarning "WORD export completed, but its undo action was not registered.", 5
    End If
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

Private Sub private_DeleteCollapsedWordBookmarkIfExists( _
    ByVal wordDoc As Object, _
    ByVal bookmarkName As String _
)
    Dim bookmarkRange As Object

    If wordDoc Is Nothing Then Exit Sub
    bookmarkName = VBA.Trim$(bookmarkName)
    If VBA.Len(bookmarkName) = 0 Then Exit Sub
    If Not wordDoc.Bookmarks.Exists(bookmarkName) Then Exit Sub

    Set bookmarkRange = wordDoc.Bookmarks(bookmarkName).Range
    If bookmarkRange.Start >= bookmarkRange.End Then
        wordDoc.Bookmarks(bookmarkName).Delete
    End If
End Sub

Private Function private_TryTrimExpandedRecordBookmarkAtInsertion( _
    ByVal wordDoc As Object, _
    ByVal insertedStart As Long _
) As Boolean
    Dim bookmarkObj As Object
    Dim bookmarkName As String
    Dim bookmarkStart As Long
    Dim correctedRange As Object

    If wordDoc Is Nothing Then Exit Function
    For Each bookmarkObj In wordDoc.Bookmarks
        bookmarkName = VBA.CStr(bookmarkObj.Name)
        If VBA.Left$(VBA.UCase$(bookmarkName), _
            VBA.Len(WORD_RECORD_BOOKMARK_PREFIX)) = _
            WORD_RECORD_BOOKMARK_PREFIX Then
            bookmarkStart = bookmarkObj.Range.Start
            If bookmarkStart < insertedStart And _
                bookmarkObj.Range.End > insertedStart Then
                wordDoc.Bookmarks(bookmarkName).Delete
                Set correctedRange = wordDoc.Range(bookmarkStart, insertedStart)
                wordDoc.Bookmarks.Add bookmarkName, correctedRange
                Exit For
            End If
        End If
    Next bookmarkObj
    private_TryTrimExpandedRecordBookmarkAtInsertion = True
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
Public Function RemoveResultDocumentAnchors( _
    ByRef clearedBlockCount As Long, _
    Optional ByVal orderNo As String = "" _
) As Boolean
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

    targetPath = private_BuildResultDocumentPath(templatePath, orderNo)
    If VBA.Len(targetPath) = 0 Or VBA.Len(VBA.Dir$(targetPath, VBA.vbNormal Or VBA.vbReadOnly Or VBA.vbHidden Or VBA.vbSystem)) = 0 Then
        VBA.MsgBox "PrototypeNew: WORD result document was not found:" & VBA.vbCrLf & targetPath, VBA.vbExclamation, "PrototypeNew / WORD document"
        Exit Function
    End If

    If Not rt_WordExportRuntime.fn_TryAcquireWordDocument( _
        targetPath, wordApp, wordDoc, documentOpened) Then Exit Function

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
    If documentOpened Then
        wordDoc.Close False
        documentOpened = False
    End If
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

Private Function private_BuildResultDocumentPath( _
    ByVal templatePath As String, _
    ByVal orderNo As String _
) As String
    Dim extensionPos As Long
    Dim slashPos As Long
    Dim backslashPos As Long
    Dim safeOrderNo As String

    templatePath = VBA.Trim$(templatePath)
    If VBA.Len(templatePath) = 0 Then Exit Function
    safeOrderNo = private_NormalizeResultFileNamePart(orderNo)
    If VBA.Len(safeOrderNo) = 0 Then
        VBA.MsgBox "PrototypeNew: WORD result document requires a non-empty order number.", _
            VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    slashPos = VBA.InStrRev(templatePath, "/")
    backslashPos = VBA.InStrRev(templatePath, "\")
    extensionPos = VBA.InStrRev(templatePath, ".")

    If extensionPos > slashPos And extensionPos > backslashPos Then
        private_BuildResultDocumentPath = VBA.Left$(templatePath, extensionPos - 1) & _
            "_" & safeOrderNo & VBA.Mid$(templatePath, extensionPos)
    Else
        private_BuildResultDocumentPath = templatePath & "_" & safeOrderNo & ".docx"
    End If
End Function

Private Function private_NormalizeResultFileNamePart(ByVal valueText As String) As String
    Dim invalidChars As Variant
    Dim invalidChar As Variant

    valueText = VBA.Trim$(valueText)
    invalidChars = VBA.Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each invalidChar In invalidChars
        valueText = VBA.Replace(valueText, VBA.CStr(invalidChar), "_")
    Next invalidChar
    private_NormalizeResultFileNamePart = valueText
End Function

Private Function private_BuildRecordBookmarkName(ByVal templateHash As String, ByVal recordIpn As String) As String
    Dim safeTemplateHash As String
    Dim safeIpn As String

    safeTemplateHash = VBA.UCase$(private_NormalizeBookmarkPart(templateHash))
    safeIpn = private_NormalizeBookmarkPart(recordIpn)
    If Not safeTemplateHash Like "[0-9A-F][0-9A-F]" Then Exit Function
    If VBA.Len(safeIpn) = 0 Then Exit Function

    private_BuildRecordBookmarkName = WORD_RECORD_BOOKMARK_PREFIX & _
        safeTemplateHash & "_" & safeIpn
End Function

Private Function private_BuildGroupBookmarkName( _
    ByVal templateHash As String, _
    ByVal groupKeyText As String, _
    ByVal groupOrderText As String _
) As String
    Dim templatePart As String
    Dim keyPart As String
    Dim keyHash As String
    Dim bookmarkPrefix As String
    Dim availableKeyLength As Long
    Dim availableVisibleKeyLength As Long
    Dim dateParts As Variant

    templatePart = VBA.UCase$(private_NormalizeBookmarkPart(templateHash))
    If Not templatePart Like "[0-9A-F][0-9A-F]" Then Exit Function
    groupKeyText = VBA.Trim$(groupKeyText)
    Select Case VBA.LCase$(VBA.Trim$(groupOrderText))
        Case "date"
            dateParts = VBA.Split(groupKeyText, ".")
            If UBound(dateParts) <> 2 Then
                VBA.MsgBox "PrototypeNew: groupOrder='date' expects DD.MM.YYYY, got: " & _
                    groupKeyText, VBA.vbExclamation, "PrototypeNew / WORD export"
                Exit Function
            End If
            keyPart = VBA.CStr(dateParts(2)) & _
                VBA.Right$("0" & VBA.CStr(dateParts(1)), 2) & _
                VBA.Right$("0" & VBA.CStr(dateParts(0)), 2)
        Case "text", "none"
            keyPart = private_NormalizeBookmarkPart( _
                VBA.LCase$(groupKeyText))
        Case Else
            Exit Function
    End Select

    bookmarkPrefix = WORD_GROUP_BOOKMARK_PREFIX & templatePart & "_"
    availableKeyLength = WORD_BOOKMARK_MAX_LENGTH - VBA.Len(bookmarkPrefix)
    If availableKeyLength <= 0 Then
        VBA.MsgBox "PrototypeNew: WORD group bookmark prefix leaves no room " & _
            "for a group key: " & bookmarkPrefix, VBA.vbExclamation, _
            "PrototypeNew / WORD export"
        Exit Function
    End If

    If VBA.Len(keyPart) > availableKeyLength Then
        keyHash = private_BuildStableBookmarkHash( _
            VBA.LCase$(VBA.Trim$(groupKeyText)))
        availableVisibleKeyLength = availableKeyLength - _
            VBA.Len(keyHash) - 1
        If availableVisibleKeyLength <= 0 Then
            VBA.MsgBox "PrototypeNew: WORD group bookmark prefix is too long " & _
                "to append a stable group hash: " & bookmarkPrefix, _
                VBA.vbExclamation, "PrototypeNew / WORD export"
            Exit Function
        End If
        keyPart = VBA.Left$(keyPart, availableVisibleKeyLength) & _
            "_" & keyHash
    End If

    private_BuildGroupBookmarkName = bookmarkPrefix & keyPart
End Function

Private Function private_BuildNestedGroupBookmarkName( _
    ByVal parentBookmarkName As String, _
    ByVal groupKeyText As String _
) As String
    Dim parentHash As String
    Dim keyHash As String
    Dim keyPart As String
    Dim availableKeyLength As Long

    ' Родительский hash делает одинаковую больницу независимой для разных дат.
    ' Видимая часть ключа сохраняет сортировку, полный hash защищает усечённые
    ' длинные названия от совпадений и помогает уложиться в лимит Word.
    parentHash = private_BuildStableBookmarkHash(parentBookmarkName)
    keyHash = private_BuildStableBookmarkHash(VBA.LCase$(VBA.Trim$(groupKeyText)))
    keyPart = private_NormalizeMetadataBookmarkPart( _
        VBA.LCase$(groupKeyText))
    availableKeyLength = WORD_BOOKMARK_MAX_LENGTH - _
        VBA.Len(WORD_NESTED_GROUP_BOOKMARK_PREFIX) - _
        VBA.Len(parentHash) - VBA.Len(keyHash) - 2
    If availableKeyLength <= 0 Or VBA.Len(keyPart) = 0 Then Exit Function
    If VBA.Len(keyPart) > availableKeyLength Then
        keyPart = VBA.Left$(keyPart, availableKeyLength)
    End If
    private_BuildNestedGroupBookmarkName = _
        WORD_NESTED_GROUP_BOOKMARK_PREFIX & parentHash & "_" & _
        keyPart & "_" & keyHash
End Function

Private Function private_BuildStableBookmarkHash(ByVal valueText As String) As String
    Dim hashValue As Double
    Dim charCode As Long
    Dim charIndex As Long
    Dim hexText As String

    hashValue = 5381#
    For charIndex = 1 To VBA.Len(valueText)
        charCode = VBA.AscW(VBA.Mid$(valueText, charIndex, 1))
        If charCode < 0 Then charCode = charCode + 65536
        hashValue = hashValue * 33# + charCode
        hashValue = hashValue - _
            VBA.Int(hashValue / 2147483647#) * 2147483647#
    Next charIndex
    hexText = VBA.Hex$(VBA.CLng(hashValue))
    private_BuildStableBookmarkHash = VBA.Right$("00000000" & hexText, 8)
End Function

Private Function private_TryFindNestedGroupInsertPosition( _
    ByVal wordDoc As Object, _
    ByVal parentBookmarkName As String, _
    ByVal newGroupBookmarkName As String, _
    ByVal groupOrderText As String, _
    ByVal parentStart As Long, _
    ByVal parentContentStart As Long, _
    ByVal parentEnd As Long, _
    ByRef outInsertPosition As Long _
) As Boolean
    Dim bookmarkObj As Object
    Dim bookmarkName As String
    Dim prefixText As String
    Dim candidateStart As Long

    outInsertPosition = parentEnd
    If wordDoc Is Nothing Then Exit Function
    If VBA.LCase$(VBA.Trim$(groupOrderText)) = "none" Then
        private_TryFindNestedGroupInsertPosition = True
        Exit Function
    End If
    ' Сравниваются имена служебных закладок, а не текст заголовков документа.
    ' Поэтому ручное форматирование заголовка не влияет на поиск группы.
    prefixText = VBA.LCase$(WORD_NESTED_GROUP_BOOKMARK_PREFIX & _
        private_BuildStableBookmarkHash(parentBookmarkName) & "_")
    For Each bookmarkObj In wordDoc.Bookmarks
        bookmarkName = VBA.CStr(bookmarkObj.Name)
        If VBA.Left$(VBA.LCase$(bookmarkName), VBA.Len(prefixText)) <> prefixText Then _
            GoTo ContinueBookmark
        If bookmarkObj.Range.Start < parentStart Or bookmarkObj.Range.End > parentEnd Then _
            GoTo ContinueBookmark
        If VBA.StrComp(bookmarkName, newGroupBookmarkName, VBA.vbTextCompare) > 0 Then
            candidateStart = bookmarkObj.Range.Start
            ' Legacy PEN могла начинаться вместе с PEG и включать дату.
            ' Вставка новой больницы никогда не должна подниматься выше даты.
            If candidateStart < parentContentStart Then
                candidateStart = parentContentStart
            End If
            If candidateStart < outInsertPosition Then _
                outInsertPosition = candidateStart
        End If
ContinueBookmark:
    Next bookmarkObj
    private_TryFindNestedGroupInsertPosition = True
End Function

Private Function private_TryFindGroupInsertPosition( _
    ByVal wordDoc As Object, _
    ByVal templateHash As String, _
    ByVal newGroupBookmarkName As String, _
    ByVal groupOrderText As String, _
    ByVal sectionStart As Long, _
    ByVal sectionEnd As Long, _
    ByRef outInsertPosition As Long _
) As Boolean
    Dim bookmarkObj As Object
    Dim bookmarkName As String
    Dim prefixText As String

    outInsertPosition = sectionEnd
    If wordDoc Is Nothing Then Exit Function
    If VBA.LCase$(VBA.Trim$(groupOrderText)) = "none" Then
        private_TryFindGroupInsertPosition = True
        Exit Function
    End If
    prefixText = VBA.LCase$(WORD_GROUP_BOOKMARK_PREFIX & _
        private_NormalizeBookmarkPart(templateHash) & "_")
    For Each bookmarkObj In wordDoc.Bookmarks
        bookmarkName = VBA.CStr(bookmarkObj.Name)
        If VBA.Left$(VBA.LCase$(bookmarkName), VBA.Len(prefixText)) <> prefixText Then _
            GoTo ContinueBookmark
        If bookmarkObj.Range.Start < sectionStart Or bookmarkObj.Range.End > sectionEnd Then _
            GoTo ContinueBookmark
        If VBA.StrComp(bookmarkName, newGroupBookmarkName, VBA.vbTextCompare) > 0 Then
            If bookmarkObj.Range.Start < outInsertPosition Then _
                outInsertPosition = bookmarkObj.Range.Start
        End If
ContinueBookmark:
    Next bookmarkObj
    private_TryFindGroupInsertPosition = True
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

Private Function private_BuildUniqueBookmarkName(ByVal wordDoc As Object, ByVal baseName As String) As String
    Dim candidateName As String
    Dim suffixText As String
    Dim sequenceNo As Long

    baseName = VBA.Trim$(baseName)
    If wordDoc Is Nothing Or VBA.Len(baseName) = 0 Then Exit Function
    If Not wordDoc.Bookmarks.Exists(baseName) Then
        private_BuildUniqueBookmarkName = baseName
        Exit Function
    End If

    sequenceNo = 2
    Do
        suffixText = "_" & VBA.CStr(sequenceNo)
        candidateName = baseName & suffixText
        If VBA.Len(candidateName) > WORD_BOOKMARK_MAX_LENGTH Then
            VBA.MsgBox "PrototypeNew: WORD bookmark occurrence suffix exceeds " & _
                "the name limit: " & candidateName, VBA.vbExclamation, _
                "PrototypeNew / WORD export"
            Exit Function
        End If
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

Private Function private_TryStripPreviewInlineMarkers( _
    ByVal renderedText As String, _
    ByRef outPlainText As String _
) As Boolean
    Dim inlineTextProfile As obj_InlineTextProfile
    Dim ignoredRuns As Collection

    outPlainText = VBA.vbNullString
    Set inlineTextProfile = New obj_InlineTextProfile
    inlineTextProfile.InlineMarkersEnabled = True
    ' Word получает тот же plain-текст, который общий inline pipeline передает
    ' Excel Banner перед посимвольным оформлением.
    If Not inlineTextProfile.TryResolveInlineText( _
        renderedText, outPlainText, ignoredRuns) Then Exit Function

    private_TryStripPreviewInlineMarkers = True
End Function

' Excel хранит перенос строки внутри ячейки как LF, а Word использует CR
' как границу абзаца. Неразрывные и типографические пробелы не изменяем.
Private Function private_NormalizeWordParagraphBreaks( _
    ByVal sourceText As String _
) As String
    sourceText = VBA.Replace(sourceText, VBA.vbCrLf, VBA.vbCr)
    sourceText = VBA.Replace(sourceText, VBA.vbLf, VBA.vbCr)
    private_NormalizeWordParagraphBreaks = sourceText
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
    Dim toHospitalShortText As String
    Dim destinationText As String
    Dim movementDestinationText As String
    Dim reportRankText As String
    Dim reportPersonText As String
    Dim reportPositionCodeText As String
    Dim incomingNoText As String
    Dim incomingDateText As String
    Dim docDateText As String
    Dim dateFromText As String
    Dim vacationTicketDateText As String
    Dim vlkDateText As String
    Dim orderNoText As String
    Dim resolvedOrderDate As Date
    Dim sectionTypeText As String
    Dim durationDaysText As String
    Dim vacationTicketNoText As String
    Dim normalizedVacationTicketNoText As String
    Dim rankGenitive As String
    Dim rankGenitiveFound As Boolean
    Dim rankGenitiveCellTag As String
    Dim fioGenitive As String
    Dim fioAccusative As String
    Dim fioInitialsGenitive As String
    Dim fioDeclensionFound As Boolean
    Dim fioDeclensionCellTag As String
    Dim positionGenitive As String
    Dim positionGenitiveFound As Boolean
    Dim positionGenitiveCellTag As String
    Dim hospitalGenitive As String
    Dim hospitalAccusative As String
    Dim hospitalDative As String
    Dim toHospitalAccusative As String
    Dim usesDestinationHospital As Boolean
    Dim reportRankGenitive As String
    Dim reportRankGenitiveFound As Boolean
    Dim reportRankGenitiveCellTag As String
    Dim reportPersonGenitive As String
    Dim reportPersonInitialsGenitive As String
    Dim reportPositionGenitive As String
    Dim reportTvoPositionGenitive As String
    Dim reporterGenitive As String
    Dim isReporterTvo As Boolean
    Dim isReporterTvoOverride As Boolean
    Dim dateFromDate As Date
    Dim hasDateFrom As Boolean
    Dim removeFromFoodSupportDateText As String
    Dim enrollToFoodSupportDateText As String
    Dim foodSupportDateValue As Date
    Dim requiresFoodSupportChange As Boolean
    Dim requiresFoodSupportEnroll As Boolean
    Dim requiresFoodSupportRemove As Boolean
    Dim vacationTotalDays As Long
    Dim vacationDays As Long
    Dim additionalDurationPhrases As String
    Dim vacationDateTo As Date
    Dim builderData As obj_PrsnlEvntBuilderData
    Dim uaLocationInflector As obj_IUaInflector
    Dim inflectedDestinationText As String
    If sourceTable Is Nothing Then Exit Function
    If m_ExporterCfgDataProvider Is Nothing Then Exit Function
    If m_ExporterCfgDataProvider.CommonData Is Nothing Then Exit Function

    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_IPN, ipnText) Then ipnText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_RANK, rankText) Then rankText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_FIO, fioText) Then fioText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_POSITION_CODE, positionCodeText) Then positionCodeText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_POSITION_NAME, positionText) Then positionText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_HOSPITAL_SHORT, hospitalShortText) Then hospitalShortText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_TO_HOSPITAL_SHORT, toHospitalShortText) Then toHospitalShortText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_DESTINATION, destinationText) Then destinationText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_REPORT_RANK, reportRankText) Then reportRankText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_REPORT_PERSON, reportPersonText) Then reportPersonText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_REPORT_POSITION_CODE, reportPositionCodeText) Then reportPositionCodeText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_INCOMING_NO, incomingNoText) Then incomingNoText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_INCOMING_DATE, incomingDateText) Then incomingDateText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_DOC_DATE, docDateText) Then docDateText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_DATE_FROM, dateFromText) Then dateFromText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_DURATION_DAYS, durationDaysText) Then durationDaysText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_VACATION_TICKET_NO, vacationTicketNoText) Then vacationTicketNoText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_VACATION_TICKET_DATE, vacationTicketDateText) Then vacationTicketDateText = VBA.vbNullString
    If Not private_TryGetMainTableValue(sourceTable, SOURCE_ALIAS_VLK_DATE, vlkDateText) Then vlkDateText = VBA.vbNullString

    ' WORD preview работает не только с исходными колонками формы. Перед
    ' render-ом мы дописываем в main DynamicTable вычисленные поля:
    ' склонения ФИО/звания/посады/лечебного учреждения, строку рапортующего
    ' и date aliases. Все даты передаём в едином компактном виде
    ' ...Short = 01.02.2025; полный вид формирует XML через dateformat.
    orderNoText = private_GetContextText(context, CONTEXT_MANUAL_ORDER_NO)
    If Not VBA.IsNumeric(private_GetContextText( _
        context, CONTEXT_MANUAL_ORDER_DATE_SERIAL)) Then Exit Function
    resolvedOrderDate = VBA.CDate(VBA.CDbl(private_GetContextText( _
        context, CONTEXT_MANUAL_ORDER_DATE_SERIAL)))
    If Not m_ExporterCfgDataProvider.CommonData.SetResolvedOrderPair( _
        orderNoText, resolvedOrderDate) Then Exit Function
    If Not m_ExporterCfgDataProvider.CommonData.TryFormatVacationTicketNoForExport( _
        vacationTicketNoText, orderNoText, normalizedVacationTicketNoText) Then Exit Function
    If VBA.Len(normalizedVacationTicketNoText) > 0 Then
        If Not private_TryUpsertMainTableValue(sourceTable, SOURCE_ALIAS_VACATION_TICKET_NO, normalizedVacationTicketNoText) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(orderNoText)) > 0 Then
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_ORDER_NO, orderNoText) Then Exit Function
    End If
    ' Все склонения берутся из общего provider-а. Отсутствие ФИО в АЛФ не
    ' блокирует preview/export: ниже используем исходную форму и помечаем её
    ' семантическим warning-тегом для style pipeline.
    If Not m_ExporterCfgDataProvider.CommonData.TryResolveRankGenitiveOptional( _
        rankText, rankGenitive, rankGenitiveFound) Then Exit Function
    If Not rankGenitiveFound And VBA.Len(VBA.Trim$(rankText)) > 0 Then
        m_LastBuildHasWarnings = True
        rankGenitive = rankText
        rankGenitiveCellTag = PREVIEW_LOOKUP_WARNING_VALUE_TAG
        ex_Core.fn_Diagnostic_LogError _
            "peb-word:rank-declension-missing rank='" & rankText & "'"
    End If
    If VBA.StrComp(VBA.Trim$(ipnText), TEMPORARY_IPN_MARKER, _
        VBA.vbTextCompare) = 0 Then
        If Not m_ExporterCfgDataProvider.CommonData.TryResolveFioFormsByNameOptional( _
            fioText, fioGenitive, fioAccusative, fioInitialsGenitive, _
            fioDeclensionFound) Then Exit Function
    Else
        If Not m_ExporterCfgDataProvider.CommonData.TryResolveFioFormsOptional( _
            ipnText, fioGenitive, fioAccusative, fioInitialsGenitive, _
            fioDeclensionFound) Then Exit Function
    End If
    If Not fioDeclensionFound Then
        If VBA.Len(VBA.Trim$(fioText)) = 0 Then
            VBA.MsgBox "PrototypeNew: FIO was not found in ШПО / АЛФ for IPN '" & _
                ipnText & "', and the source form has no FIO.", _
                VBA.vbExclamation, "PrototypeNew / WORD export"
            Exit Function
        End If
        fioGenitive = fioText
        fioAccusative = fioText
        fioInitialsGenitive = fioText
        fioDeclensionCellTag = PREVIEW_LOOKUP_WARNING_VALUE_TAG
        m_LastBuildHasWarnings = True
        ex_Core.fn_Diagnostic_LogError _
            "peb-word:fio-declension-missing ipn='" & ipnText & _
            "' fio='" & fioText & "'"
    End If
    If Not m_ExporterCfgDataProvider.CommonData.TryResolvePositionGenitiveOptional( _
        positionCodeText, positionGenitive, positionGenitiveFound, _
        rankText) Then Exit Function
    If Not positionGenitiveFound And _
       VBA.Len(VBA.Trim$(positionCodeText)) > 0 Then
        If VBA.Len(VBA.Trim$(positionText)) = 0 Then
            VBA.MsgBox "PrototypeNew: position was not found in ШПО / Посади " & _
                "for code '" & positionCodeText & _
                "', and the source form has no position name.", _
                VBA.vbExclamation, "PrototypeNew / WORD export"
            Exit Function
        End If
        positionGenitive = positionText
        positionGenitiveCellTag = PREVIEW_LOOKUP_WARNING_VALUE_TAG
        m_LastBuildHasWarnings = True
        ex_Core.fn_Diagnostic_LogError _
            "peb-word:position-declension-missing code='" & _
            positionCodeText & "' position='" & positionText & "'"
    End If
    ' Значение остается обычным текстом. Признак вероятного обрезания ADO
    ' передаем отдельно через семантический тег ячейки DynamicTable.
    If VBA.Len(positionGenitiveCellTag) = 0 And _
       VBA.Len(positionGenitive) >= PREVIEW_TRUNCATION_WARNING_LENGTH Then
        positionGenitiveCellTag = PREVIEW_TRUNCATED_VALUE_TAG
    End If
    If Not m_ExporterCfgDataProvider.CommonData.TryResolveHospitalGenitive(hospitalShortText, hospitalGenitive) Then Exit Function
    If Not m_ExporterCfgDataProvider.CommonData.TryResolveHospitalAccusative(hospitalShortText, hospitalAccusative) Then Exit Function
    If Not m_ExporterCfgDataProvider.CommonData.TryResolveHospitalDative(hospitalShortText, hospitalDative) Then Exit Function
    sectionTypeText = private_GetContextText(context, CONTEXT_SECTION_TYPE)
    If VBA.Len(sectionTypeText) = 0 Then sectionTypeText = VBA.Trim$(sourceTable.SectionTitle)
    Set builderData = New obj_PrsnlEvntBuilderData
    usesDestinationHospital = (VBA.StrComp( _
        VBA.Trim$(sectionTypeText), _
        VBA.Trim$(builderData.SectionTypeTransferTreatmentToStationaryVlk), _
        VBA.vbTextCompare) = 0)
    If usesDestinationHospital Then
        If Not m_ExporterCfgDataProvider.CommonData.TryResolveHospitalAccusative( _
            toHospitalShortText, toHospitalAccusative) Then Exit Function
        movementDestinationText = toHospitalShortText
    Else
        movementDestinationText = destinationText
    End If
    ' Временно сохраняем место выбытия в исходном виде. Раскрытие и склонение
    ' сокращений (например, «р-н» -> «району», «обл.» -> «області»)
    ' отключено, чтобы отпускной пункт получал исходную сокращённую фразу.
    ' If Not usesDestinationHospital And VBA.Len(VBA.Trim$(destinationText)) > 0 Then
    '     Set uaLocationInflector = New obj_UaLocationInflector
    '     If Not uaLocationInflector.TryInflect( _
    '         destinationText, "genitive", inflectedDestinationText) Then
    '         VBA.MsgBox "Не вдалося відмінити адресу призначення: " & destinationText, _
    '             VBA.vbExclamation, "PrsnlEventBuilder / WORD export"
    '         Exit Function
    '     End If
    '     destinationText = inflectedDestinationText
    ' End If
    If Not m_ExporterCfgDataProvider.CommonData.TryResolveRankGenitiveOptional( _
        reportRankText, reportRankGenitive, _
        reportRankGenitiveFound) Then Exit Function
    If Not reportRankGenitiveFound And _
       VBA.Len(VBA.Trim$(reportRankText)) > 0 Then
        m_LastBuildHasWarnings = True
        reportRankGenitive = reportRankText
        reportRankGenitiveCellTag = PREVIEW_LOOKUP_WARNING_VALUE_TAG
        ex_Core.fn_Diagnostic_LogError _
            "peb-word:report-rank-declension-missing rank='" & _
            reportRankText & "'"
    End If
    If Not m_ExporterCfgDataProvider.CommonData.TryResolveFioGenitiveByName(reportPersonText, reportPersonGenitive) Then Exit Function
    If Not m_ExporterCfgDataProvider.CommonData.TryResolveFioInitialsGenitiveByName(reportPersonText, reportPersonInitialsGenitive) Then Exit Function

    isReporterTvoOverride = private_GetContextBoolean( _
        context, CONTEXT_REPORT_IS_TVO)
    If isReporterTvoOverride Then
        isReporterTvo = True
        If Not m_ExporterCfgDataProvider.CommonData.TryResolvePositionGenitive( _
            reportPositionCodeText, reportPositionGenitive) Then Exit Function
    Else
        If Not m_ExporterCfgDataProvider.TryResolveReporterTvoPositionGenitive( _
            reportPersonText, reportTvoPositionGenitive, isReporterTvo) Then Exit Function
        If isReporterTvo Then
            reportPositionGenitive = reportTvoPositionGenitive
        Else
            If Not m_ExporterCfgDataProvider.CommonData.TryResolvePositionGenitive( _
                reportPositionCodeText, reportPositionGenitive) Then Exit Function
        End If
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
        If Not private_TryUpsertMainTableValue( _
            sourceTable, WORD_ALIAS_RANK_GENITIVE, rankGenitive, _
            rankGenitiveCellTag) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(fioGenitive)) > 0 Then
        If Not private_TryUpsertMainTableValue( _
            sourceTable, WORD_ALIAS_FIO_GENITIVE, fioGenitive, _
            fioDeclensionCellTag) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(fioAccusative)) > 0 Then
        If Not private_TryUpsertMainTableValue( _
            sourceTable, WORD_ALIAS_FIO_ACCUSATIVE, fioAccusative, _
            fioDeclensionCellTag) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(fioInitialsGenitive)) > 0 Then
        If Not private_TryUpsertMainTableValue( _
            sourceTable, WORD_ALIAS_FIO_INITIALS_GENITIVE, _
            fioInitialsGenitive, fioDeclensionCellTag) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(positionGenitive)) > 0 Then
        If Not private_TryUpsertMainTableValue( _
            sourceTable, WORD_ALIAS_POSITION_GENITIVE, positionGenitive, _
            positionGenitiveCellTag) Then Exit Function
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
    If VBA.Len(VBA.Trim$(toHospitalAccusative)) > 0 Then
        If Not private_TryUpsertMainTableValue( _
            sourceTable, WORD_ALIAS_TO_HOSPITAL_ACCUSATIVE, _
            toHospitalAccusative) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(destinationText)) > 0 Then
        If Not private_TryUpsertMainTableValue(sourceTable, SOURCE_ALIAS_DESTINATION, destinationText) Then Exit Function
    End If
    If VBA.Len(VBA.Trim$(reportRankGenitive)) > 0 Then
        If Not private_TryUpsertMainTableValue( _
            sourceTable, WORD_ALIAS_REPORT_RANK_GENITIVE, _
            reportRankGenitive, reportRankGenitiveCellTag) Then Exit Function
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
            m_ExporterCfgDataProvider.NormalizeIncomingNoForExport(incomingNoText)) Then Exit Function
    End If
    If Not private_TryUpsertShortDateValue(sourceTable, SOURCE_ALIAS_INCOMING_DATE, WORD_ALIAS_INCOMING_DATE_SHORT, incomingDateText) Then Exit Function
    If Not private_TryUpsertShortDateValue(sourceTable, SOURCE_ALIAS_DOC_DATE, WORD_ALIAS_DOC_DATE_SHORT, docDateText) Then Exit Function
    If Not private_TryUpsertShortDateValue(sourceTable, SOURCE_ALIAS_DATE_FROM, WORD_ALIAS_DATE_FROM_SHORT, dateFromText) Then Exit Function
    If Not private_TryUpsertShortDateValue(sourceTable, SOURCE_ALIAS_VACATION_TICKET_DATE, WORD_ALIAS_VACATION_TICKET_DATE_SHORT, vacationTicketDateText) Then Exit Function
    If Not private_TryUpsertShortDateValue(sourceTable, SOURCE_ALIAS_VLK_DATE, WORD_ALIAS_VLK_DATE_SHORT, vlkDateText) Then Exit Function
    If VBA.Len(VBA.Trim$(dateFromText)) > 0 Then
        If Not private_TryUpsertShortDateValue( _
            sourceTable, SOURCE_ALIAS_DATE_FROM, _
            WORD_ALIAS_EFFECTIVE_RETURN_DATE_SHORT, dateFromText) Then Exit Function
    Else
        If Not private_TryUpsertShortDateValue( _
            sourceTable, SOURCE_ALIAS_INCOMING_DATE, _
            WORD_ALIAS_EFFECTIVE_RETURN_DATE_SHORT, incomingDateText) Then Exit Function
    End If

    If Not private_TryResolveDateByRawText(dateFromText, hasDateFrom, dateFromDate) Then Exit Function

    sectionTypeText = private_GetContextText(context, CONTEXT_SECTION_TYPE)
    If VBA.Len(sectionTypeText) = 0 Then sectionTypeText = VBA.Trim$(sourceTable.SectionTitle)

    ' Первый операнд задаёт отпуск, каждый следующий +N — отдельный
    ' дополнительный период. Причина в скобках определяет самостоятельный
    ' текстовый оборот, а общий срок включает все операнды.
    If VBA.Len(VBA.Trim$(durationDaysText)) > 0 Then
        If Not private_TryParseVacationDuration( _
            durationDaysText, vacationDays, additionalDurationPhrases, _
            vacationTotalDays) Then Exit Function
        If Not private_TryUpsertMainTableValue( _
            sourceTable, WORD_ALIAS_VACATION_DAYS, VBA.CStr(vacationDays)) Then Exit Function
        If Not private_TryUpsertMainTableValue( _
            sourceTable, WORD_ALIAS_VACATION_DURATION_PHRASE, _
            private_FormatVacationDurationPhrase(vacationDays)) Then Exit Function
        If Not private_TryUpsertMainTableValue( _
            sourceTable, WORD_ALIAS_ADDITIONAL_DURATION_PHRASES, _
            additionalDurationPhrases) Then Exit Function
        If Not private_TryUpsertMainTableValue( _
            sourceTable, WORD_ALIAS_VACATION_TOTAL_DAYS, VBA.CStr(vacationTotalDays)) Then Exit Function
    End If

    ' Для всех секций выбытия или перевода в отпуск вычисляем дату окончания
    ' включительно: DateFrom + VacationTotalDays - 1. Окончательное отображение
    ' даты (месяц словами, "року", NBSP) по-прежнему задаёт XML-шаблон.
    private_ResolveFoodSupportChange _
        sectionTypeText, hospitalShortText, movementDestinationText, builderData, _
        requiresFoodSupportEnroll, requiresFoodSupportRemove
    requiresFoodSupportChange = requiresFoodSupportEnroll Or requiresFoodSupportRemove
    If Not private_TryUpsertMainTableValue( _
        sourceTable, WORD_ALIAS_REQUIRES_FOOD_SUPPORT_CHANGE, _
        VBA.CStr(requiresFoodSupportChange)) Then Exit Function
    If Not private_TryUpsertMainTableValue( _
        sourceTable, WORD_ALIAS_REQUIRES_FOOD_SUPPORT_ENROLL, _
        VBA.CStr(requiresFoodSupportEnroll)) Then Exit Function
    If Not private_TryUpsertMainTableValue( _
        sourceTable, WORD_ALIAS_REQUIRES_FOOD_SUPPORT_REMOVE, _
        VBA.CStr(requiresFoodSupportRemove)) Then Exit Function

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
                If Not private_TryUpsertMainTableValue( _
                    sourceTable, WORD_ALIAS_VACATION_DATES_SAME_YEAR, _
                    VBA.CStr(VBA.Year(dateFromDate) = VBA.Year(vacationDateTo))) Then Exit Function
            End If
        End If
    End If

    If requiresFoodSupportChange Then
        ' WORD и Movement используют один helper: max(OrderDate + 1, DateFrom).
        If Not m_ExporterCfgDataProvider.CommonData.TryCalculateFoodSupportDate( _
            hasDateFrom, dateFromDate, foodSupportDateValue) Then
            VBA.MsgBox "PrototypeNew: cannot calculate food-support date because both order date and DateFrom are unavailable.", _
                VBA.vbExclamation, "PrototypeNew / WORD export"
            Exit Function
        End If

        removeFromFoodSupportDateText = VBA.Format$( _
            foodSupportDateValue, WORD_SHORT_DATE_STORAGE_FORMAT)
        enrollToFoodSupportDateText = removeFromFoodSupportDateText

        If VBA.Len(VBA.Trim$(removeFromFoodSupportDateText)) > 0 Then
            If Not private_TryUpsertMainTableValue( _
                sourceTable, WORD_ALIAS_REMOVE_FROM_FOOD_SUPPORT_DATE, _
                removeFromFoodSupportDateText) Then Exit Function
        End If
        If VBA.Len(VBA.Trim$(enrollToFoodSupportDateText)) > 0 Then
            If Not private_TryUpsertMainTableValue( _
                sourceTable, WORD_ALIAS_ENROLL_TO_FOOD_SUPPORT_DATE, _
                enrollToFoodSupportDateText) Then Exit Function
        End If
    End If

    private_TryEnrichMainSourceTableForWord = True
End Function

Private Sub private_ResolveFoodSupportChange( _
    ByVal sectionTypeText As String, _
    ByVal hospitalShortText As String, _
    ByVal destinationText As String, _
    ByVal builderData As obj_PrsnlEvntBuilderData, _
    ByRef requiresEnroll As Boolean, _
    ByRef requiresRemove As Boolean _
)
    Dim normalizedSectionType As String
    Dim transferParts() As String
    Dim sourceIsInternal As Boolean
    Dim destinationIsInternal As Boolean
    Dim isArrival As Boolean

    requiresEnroll = False
    requiresRemove = False
    If builderData Is Nothing Then Exit Sub

    normalizedSectionType = VBA.LCase$(VBA.Trim$(sectionTypeText))
    If VBA.InStr(1, normalizedSectionType, "=>", VBA.vbBinaryCompare) > 0 Then
        transferParts = VBA.Split(normalizedSectionType, "=>")
        If UBound(transferParts) <> 1 Then Exit Sub

        sourceIsInternal = private_IsInternalTransferSource( _
            transferParts(0), hospitalShortText)
        destinationIsInternal = private_IsInternalTransferDestination( _
            transferParts(1), hospitalShortText, destinationText)

        requiresRemove = sourceIsInternal And Not destinationIsInternal
        requiresEnroll = Not sourceIsInternal And destinationIsInternal
        Exit Sub
    End If

    isArrival = builderData.IsMovementClosingSectionType(sectionTypeText)
    If private_IsInternalTreatmentSide( _
        normalizedSectionType, hospitalShortText, destinationText) Then Exit Sub

    requiresEnroll = isArrival
    requiresRemove = Not isArrival
End Sub

Private Function private_IsInternalTransferSource( _
    ByVal sectionSideText As String, _
    ByVal hospitalShortText As String _
) As Boolean
    Dim normalizedSide As String

    normalizedSide = VBA.LCase$(VBA.Trim$(sectionSideText))
    If private_IsMedicalCompanyText(normalizedSide) Or _
        private_IsAmbulatoryTreatmentText(normalizedSide) Then
        private_IsInternalTransferSource = True
    ElseIf normalizedSide = "лікування" Then
        private_IsInternalTransferSource = _
            private_IsMedicalCompanyText(VBA.LCase$(hospitalShortText)) Or _
            private_IsAmbulatoryTreatmentText(VBA.LCase$(hospitalShortText))
    End If
End Function

Private Function private_IsInternalTransferDestination( _
    ByVal sectionSideText As String, _
    ByVal hospitalShortText As String, _
    ByVal destinationText As String _
) As Boolean
    Dim normalizedSide As String
    Dim normalizedPlace As String

    normalizedSide = VBA.LCase$(VBA.Trim$(sectionSideText))
    If private_IsMedicalCompanyText(normalizedSide) Or _
        private_IsAmbulatoryTreatmentText(normalizedSide) Then
        private_IsInternalTransferDestination = True
        Exit Function
    End If

    If normalizedSide = "лікування" Then
        normalizedPlace = VBA.LCase$(VBA.Trim$(hospitalShortText))
    ElseIf VBA.InStr(1, normalizedSide, "відпуст", VBA.vbTextCompare) > 0 And _
        VBA.InStr(1, normalizedSide, "лікуван", VBA.vbTextCompare) > 0 Then
        normalizedPlace = VBA.LCase$(VBA.Trim$(destinationText))
    End If

    private_IsInternalTransferDestination = _
        private_IsMedicalCompanyText(normalizedPlace) Or _
        private_IsAmbulatoryTreatmentText(normalizedPlace)
End Function

Private Function private_IsInternalTreatmentSide( _
    ByVal sectionSideText As String, _
    ByVal hospitalShortText As String, _
    ByVal destinationText As String _
) As Boolean
    Dim normalizedSide As String
    Dim normalizedPlace As String

    normalizedSide = VBA.LCase$(VBA.Trim$(sectionSideText))
    normalizedPlace = VBA.LCase$(VBA.Trim$(hospitalShortText & " " & destinationText))

    If private_IsMedicalCompanyText(normalizedSide) Or _
        private_IsAmbulatoryTreatmentText(normalizedSide) Then
        private_IsInternalTreatmentSide = True
        Exit Function
    End If

    ' Для общих секций «лікування» и «відпустка для лікування» место
    ' определяется выбранной больницей/направлением. Медрота без номера
    ' считается медротой А7383; явно указанная другая часть — внешнее место.
    If VBA.InStr(1, normalizedSide, "лікуван", VBA.vbTextCompare) > 0 Then
        private_IsInternalTreatmentSide = _
            private_IsMedicalCompanyText(normalizedPlace) Or _
            private_IsAmbulatoryTreatmentText(normalizedPlace)
    End If
End Function

Private Function private_IsMedicalCompanyText(ByVal normalizedText As String) As Boolean
    Dim unitRx As Object
    Dim unitMatches As Object
    Dim normalizedUnitNo As String

    If VBA.InStr(1, normalizedText, "медичн", VBA.vbTextCompare) = 0 Or _
        VBA.InStr(1, normalizedText, "рот", VBA.vbTextCompare) = 0 Then Exit Function

    ' Медрота без номера части означает медроту А7383. Если номер указан
    ' явно, внутренней считается только медрота нашей воинской части.
    Set unitRx = VBA.CreateObject("VBScript.RegExp")
    unitRx.Global = False
    unitRx.IgnoreCase = True
    unitRx.Pattern = "[АA]\s*\d{4}"
    Set unitMatches = unitRx.Execute(normalizedText)
    If unitMatches.Count = 0 Then
        private_IsMedicalCompanyText = True
        Exit Function
    End If

    normalizedUnitNo = VBA.Replace$( _
        VBA.LCase$(VBA.CStr(unitMatches(0).Value)), " ", VBA.vbNullString)
    private_IsMedicalCompanyText = _
        normalizedUnitNo = "а7383" Or normalizedUnitNo = "a7383"
End Function

Private Function private_IsAmbulatoryTreatmentText(ByVal normalizedText As String) As Boolean
    private_IsAmbulatoryTreatmentText = _
        VBA.InStr(1, normalizedText, "амбулатор", VBA.vbTextCompare) > 0 And _
        VBA.InStr(1, normalizedText, "лікуван", VBA.vbTextCompare) > 0
End Function

Private Function private_TryEnrichPreviousVacationTicketForWord( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal sectionTypeText As String _
) As Boolean
    Dim data As obj_PrsnlEvntBuilderData
    Dim ipnText As String
    Dim fioText As String
    Dim previousTicketFound As Boolean
    Dim escortDocumentText As String
    Dim departureOrderText As String
    Dim departureOrderDate As Date
    Dim normalizedPreviousTicketNo As String
    Dim previousTicketDateText As String
    Dim isSupportedSection As Boolean

    If sourceTable Is Nothing Then Exit Function

    Set data = New obj_PrsnlEvntBuilderData
    isSupportedSection = _
        (VBA.StrComp(sectionTypeText, data.SectionTypeTransferTreatmentVacationToTreatment, VBA.vbTextCompare) = 0) _
        Or (VBA.StrComp(sectionTypeText, data.SectionTypeTransferTreatmentVacationToTreatmentVacation, VBA.vbTextCompare) = 0) _
        Or (VBA.StrComp(sectionTypeText, data.SectionTypeTransferAnnualVacationToTreatment, VBA.vbTextCompare) = 0) _
        Or (VBA.StrComp(sectionTypeText, data.SectionTypeTransferAnnualVacationToFamilyVacation, VBA.vbTextCompare) = 0) _
        Or (VBA.StrComp(sectionTypeText, data.SectionTypeTransferFamilyVacationToAnnualVacation, VBA.vbTextCompare) = 0) _
        Or (VBA.StrComp(sectionTypeText, data.SectionTypeTransferTreatmentVacationToVlk, VBA.vbTextCompare) = 0) _
        Or (VBA.StrComp(sectionTypeText, data.SectionTypeTransferAnnualVacationToVlk, VBA.vbTextCompare) = 0) _
        Or (VBA.StrComp(sectionTypeText, data.SectionTypeTransferFamilyVacationToVlk, VBA.vbTextCompare) = 0) _
        Or (VBA.StrComp(sectionTypeText, data.SectionTypeTransferMedicalCompanyToTreatmentVacation, VBA.vbTextCompare) = 0) _
        Or (VBA.StrComp(sectionTypeText, data.SectionTypeTransferMedicalCompanyTreatmentVacationToTreatment, VBA.vbTextCompare) = 0)
    If Not isSupportedSection Then
        private_TryEnrichPreviousVacationTicketForWord = True
        Exit Function
    End If

    If Not private_TryGetMainTableValue( _
        sourceTable, SOURCE_ALIAS_IPN, ipnText) Then Exit Function
    If Not private_TryGetMainTableValue( _
        sourceTable, SOURCE_ALIAS_FIO, fioText) Then fioText = VBA.vbNullString
    If Not m_ExporterCfgDataProvider.TryGetLatestMovementVacationTicket( _
        ipnText, _
        previousTicketFound, _
        escortDocumentText, _
        departureOrderText, _
        fioText) Then Exit Function
    If Not previousTicketFound Then
        VBA.MsgBox "PrototypeNew: Movement has no previous record containing both " & _
            "'Супровідний документ' and 'Вибуття.Наказ' for IPN '" & _
            ipnText & "'.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    escortDocumentText = private_NormalizeTemplateScalar(escortDocumentText)
    departureOrderText = private_NormalizeTemplateScalar(departureOrderText)

    If Not m_ExporterCfgDataProvider.CommonData.TryResolveOrderDateByNumber(departureOrderText, departureOrderDate) Then
        VBA.MsgBox "PrototypeNew: failed to resolve PrevVacationTicketDate by departure order '" & departureOrderText & "'.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    If Not m_ExporterCfgDataProvider.CommonData.TryFormatVacationTicketNoForExport( _
        escortDocumentText, departureOrderText, normalizedPreviousTicketNo) Then Exit Function

    previousTicketDateText = VBA.Format$(departureOrderDate, WORD_SHORT_DATE_STORAGE_FORMAT)
    If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_PREV_VACATION_TICKET_NO, normalizedPreviousTicketNo) Then Exit Function
    If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_PREV_VACATION_TICKET_DATE_SHORT, previousTicketDateText) Then Exit Function

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

Private Function private_FormatVacationDurationPhrase( _
    ByVal vacationDays As Long _
) As String
    Dim lastDigit As Long
    Dim lastTwoDigits As Long
    Dim daysUnit As String

    lastDigit = vacationDays Mod 10
    lastTwoDigits = vacationDays Mod 100

    ' После «терміном на» числительное требует винительного падежа:
    ' 1/21 — «календарну добу», 2–4/22–24 — «календарні доби».
    ' Числа 11–14 являются исключением и используют «календарних діб».
    If lastTwoDigits >= 11 And lastTwoDigits <= 14 Then
        daysUnit = "календарних діб"
    Else
        Select Case lastDigit
            Case 1
                daysUnit = "календарну добу"
            Case 2 To 4
                daysUnit = "календарні доби"
            Case Else
                daysUnit = "календарних діб"
        End Select
    End If

    private_FormatVacationDurationPhrase = "терміном на " & _
        VBA.CStr(vacationDays) & " " & daysUnit
End Function

Private Function private_TryParseVacationDuration( _
    ByVal durationText As String, _
    ByRef outVacationDays As Long, _
    ByRef outAdditionalPhrases As String, _
    ByRef outTotalDays As Long _
) As Boolean
    Dim fullRx As Object
    Dim operandRx As Object
    Dim matches As Object
    Dim operandMatch As Object
    Dim additionalDays As Long
    Dim reasonText As String
    Dim phraseText As String

    outVacationDays = 0
    outAdditionalPhrases = VBA.vbNullString
    outTotalDays = 0
    durationText = VBA.Trim$(durationText)
    If VBA.Len(durationText) = 0 Then
        private_TryParseVacationDuration = True
        Exit Function
    End If

    On Error GoTo EH
    ' Поддерживается основной срок и произвольное число операндов +N.
    ' Причина каждого дополнительного периода опциональна, но если указана,
    ' должна принадлежать известному набору.
    Set fullRx = VBA.CreateObject("VBScript.RegExp")
    fullRx.Global = False
    fullRx.IgnoreCase = True
    fullRx.Pattern = "^\s*(\d+)((?:\s*\+\s*\d+(?:\s*\([^)]*\))?)*)\s*$"
    Set matches = fullRx.Execute(durationText)
    If matches.Count = 0 Then
        VBA.MsgBox "PrototypeNew: unsupported vacation duration: '" & _
            durationText & "'. Expected, for example, " & _
            "'15+2 (дорога) +1 (донація крові)'.", _
            VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    outVacationDays = VBA.CLng(matches(0).SubMatches(0))
    outTotalDays = outVacationDays

    Set operandRx = VBA.CreateObject("VBScript.RegExp")
    operandRx.Global = True
    operandRx.IgnoreCase = True
    operandRx.Pattern = "\+\s*(\d+)(?:\s*\(([^)]*)\))?"
    Set matches = operandRx.Execute(durationText)
    For Each operandMatch In matches
        additionalDays = VBA.CLng(operandMatch.SubMatches(0))
        reasonText = VBA.LCase$(VBA.Trim$( _
            VBA.CStr(operandMatch.SubMatches(1))))
        If Not private_TryFormatAdditionalDurationPhrase( _
            additionalDays, reasonText, phraseText) Then Exit Function
        outAdditionalPhrases = outAdditionalPhrases & phraseText
        outTotalDays = outTotalDays + additionalDays
    Next operandMatch
    If VBA.Len(outAdditionalPhrases) > 0 Then _
        outAdditionalPhrases = outAdditionalPhrases & ","
    private_TryParseVacationDuration = True
    Exit Function

EH:
    VBA.MsgBox "PrototypeNew: failed to calculate vacation duration from '" & durationText & "': " & Err.Description, VBA.vbExclamation, "PrototypeNew / WORD export"
End Function

Private Function private_TryFormatAdditionalDurationPhrase( _
    ByVal additionalDays As Long, _
    ByVal reasonText As String, _
    ByRef outPhrase As String _
) As Boolean
    Dim daysText As String
    Dim agreementText As String

    outPhrase = VBA.vbNullString
    If additionalDays <= 0 Then
        VBA.MsgBox "PrototypeNew: additional vacation days must be greater " & _
            "than zero.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    If additionalDays = 1 Then
        daysText = "1 доби"
    Else
        daysText = VBA.CStr(additionalDays) & " діб"
    End If

    Select Case reasonText
        Case "дорога"
            If additionalDays = 1 Then
                agreementText = "необхідної"
            Else
                agreementText = "необхідних"
            End If
            outPhrase = ", без урахування " & daysText & ", " & _
                agreementText & " для проїзду до місця проведення " & _
                "відпустки та у зворотному напрямку"
        Case "донація", "донація крові"
            If additionalDays = 1 Then
                agreementText = "наданої"
            Else
                agreementText = "наданих"
            End If
            outPhrase = ", без урахування " & daysText & ", " & _
                agreementText & " за донацію крові"
        Case VBA.vbNullString
            outPhrase = ", без урахування " & daysText
        Case Else
            VBA.MsgBox "PrototypeNew: unsupported additional vacation " & _
                "reason: '" & reasonText & "'. Supported reasons: " & _
                "'дорога', 'донація крові', 'донація'.", _
                VBA.vbExclamation, "PrototypeNew / WORD export"
            Exit Function
    End Select

    private_TryFormatAdditionalDurationPhrase = True
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
        If Not m_ExporterCfgDataProvider.CommonData.TryResolveRankByIpn(VBA.CStr(chainItem("IPN")), rankText) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError _
                "word-export:tvo-rank-lookup failed fio='" & VBA.Replace$(fioText, "'", "''") & _
                "' ipn='" & VBA.Replace$(VBA.CStr(chainItem("IPN")), "'", "''") & "'"
#End If
            Exit Function
        End If
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
    Dim fioAccusative As String
    Dim fioDative As String
    Dim positionGenitive As String
    Dim positionDative As String
    Dim positionGenitiveFound As Boolean
    Dim positionDativeFound As Boolean
    Dim fallbackPositionText As String
    Dim positionGenitiveCellTag As String
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
        If Not m_ExporterCfgDataProvider.CommonData.TryResolveRankGenitive(rankText, rankGenitive) Then Exit Function
        If Not m_ExporterCfgDataProvider.CommonData.TryResolveRankDative(rankText, rankDative) Then Exit Function
        If Not m_ExporterCfgDataProvider.CommonData.TryResolveFioGenitive(ipnText, fioGenitive) Then Exit Function
        If Not m_ExporterCfgDataProvider.CommonData.TryResolveFioAccusative(ipnText, fioAccusative) Then Exit Function
        If Not m_ExporterCfgDataProvider.CommonData.TryResolveFioDative(ipnText, fioDative) Then Exit Function
#If TVO_POSITION_FALLBACK_ENABLED Then
        If Not m_ExporterCfgDataProvider.CommonData.TryResolvePositionFormsOptional( _
            positionCodeText, fallbackPositionText, positionGenitive, positionDative, _
            positionGenitiveFound, rankText) Then Exit Function
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
        If Not m_ExporterCfgDataProvider.CommonData.TryResolvePositionFormsOptional( _
            positionCodeText, positionText, positionGenitive, positionDative, _
            positionGenitiveFound, rankText) Then Exit Function
        If Not positionGenitiveFound Then
            VBA.MsgBox "PrototypeNew: declension row was not found in ШПО / Посади for key: " & positionCodeText, _
                VBA.vbExclamation, "PrototypeNew / WORD export"
            Exit Function
        End If
#End If
        If VBA.Len(positionGenitive) >= PREVIEW_TRUNCATION_WARNING_LENGTH Then
            positionGenitiveCellTag = PREVIEW_TRUNCATED_VALUE_TAG
        Else
            positionGenitiveCellTag = VBA.vbNullString
        End If
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_RANK_GENITIVE, rankGenitive) Then Exit Function
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_RANK_DATIVE, rankDative) Then Exit Function
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_FIO_GENITIVE, fioGenitive) Then Exit Function
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_FIO_ACCUSATIVE, fioAccusative) Then Exit Function
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_FIO_DATIVE, fioDative) Then Exit Function
        If Not private_TryUpsertMainTableValue(sourceTable, WORD_ALIAS_POSITION_DEFAULT, positionText) Then Exit Function
        If Not private_TryUpsertMainTableValue( _
            sourceTable, WORD_ALIAS_POSITION_GENITIVE, positionGenitive, _
            positionGenitiveCellTag) Then Exit Function
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

    If m_ExporterCfgDataProvider Is Nothing Then Exit Function
    If m_ExporterCfgDataProvider.CommonData Is Nothing Then Exit Function

    If m_ExporterCfgDataProvider.CommonData.HasOrderDate Then
        If Not ex_Helpers.fn_TryResolveDateWithContext(trimmedDateText, m_ExporterCfgDataProvider.CommonData.OrderDate, outDateValue) Then
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
    If m_ExporterCfgDataProvider Is Nothing Then Exit Function
    If m_ExporterCfgDataProvider.CommonData Is Nothing Then Exit Function
    If m_ExporterCfgDataProvider.CommonData.HasOrderDate Then
        If Not ex_Helpers.fn_TryResolveDateWithContext(trimmedDateText, m_ExporterCfgDataProvider.CommonData.OrderDate, resolvedDate) Then
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
    Static leadingInitialsRx As Object

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

    If leadingInitialsRx Is Nothing Then
        Set leadingInitialsRx = VBA.CreateObject("VBScript.RegExp")
        leadingInitialsRx.Global = True
        leadingInitialsRx.IgnoreCase = False
        ' Обратная форма также встречается в названиях учреждений и должностях:
        ' "І. І. Прізвище". Не даём фамилии оторваться от группы инициалов.
        leadingInitialsRx.Pattern = "([А-ЯІЇЄҐA-Z]\.[ \t]*[А-ЯІЇЄҐA-Z]\.)[ \t]+(\S+)"
    End If

    valueText = initialsRx.Replace(valueText, "$1" & VBA.ChrW$(160) & "$2")
    private_KeepSurnameWithInitialsTogether = leadingInitialsRx.Replace( _
        valueText, "$1" & VBA.ChrW$(160) & "$2")
End Function

Private Function private_KeepSettlementPrefixTogether(ByVal valueText As String) As String
    Static settlementRx As Object

    valueText = VBA.Trim$(valueText)
    If VBA.Len(valueText) = 0 Then Exit Function

    If settlementRx Is Nothing Then
        Set settlementRx = VBA.CreateObject("VBScript.RegExp")
        settlementRx.Global = True
        settlementRx.IgnoreCase = True
        ' Это исключительно WORD-типографика: не даём обозначению населённого
        ' пункта оторваться от его названия при переносе строки.
        settlementRx.Pattern = "(сел|м|с)\.[ \t]*"
    End If

    private_KeepSettlementPrefixTogether = settlementRx.Replace( _
        valueText, "$1." & VBA.ChrW$(160))
End Function

Private Function private_NormalizeHospitalWordTypography(ByVal valueText As String) As String
    Static hospitalNumberRx As Object

    valueText = private_KeepSettlementPrefixTogether(valueText)
    valueText = private_KeepSurnameWithInitialsTogether(valueText)
    If VBA.Len(valueText) = 0 Then Exit Function

    If hospitalNumberRx Is Nothing Then
        Set hospitalNumberRx = VBA.CreateObject("VBScript.RegExp")
        hospitalNumberRx.Global = True
        hospitalNumberRx.IgnoreCase = False
        ' Нормализуем как «№ 18», так и уже слитное «№18» и не даём WORD
        ' разорвать обозначение и число: итоговая форма всегда «№ 18».
        hospitalNumberRx.Pattern = "№[ \t]*(\d)"
    End If

    private_NormalizeHospitalWordTypography = hospitalNumberRx.Replace( _
        valueText, "№" & VBA.ChrW$(160) & "$1")
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
             VBA.LCase$(WORD_ALIAS_HOSPITAL_ACCUSATIVE), _
             VBA.LCase$(WORD_ALIAS_HOSPITAL_DATIVE), _
             VBA.LCase$(WORD_ALIAS_TO_HOSPITAL_ACCUSATIVE)
            valueText = private_NormalizeHospitalWordTypography(valueText)

        Case VBA.LCase$(SOURCE_ALIAS_DESTINATION)
            valueText = private_KeepSettlementPrefixTogether(valueText)
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
    ByVal valueText As String, _
    Optional ByVal cellTag As String = VBA.vbNullString _
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
    If Not sourceRow.SetCellRaw(columnIndex, valueText) Then Exit Function

    cellTag = VBA.Trim$(cellTag)
    If VBA.Len(cellTag) > 0 Then
        If Not sourceRow.AddCellTag(columnIndex, cellTag) Then Exit Function
    End If

    private_TryUpsertMainTableValue = True
End Function

Private Function private_GetContextText(ByVal context As Object, ByVal keyText As String) As String
    Dim rawValueText As String

    If context Is Nothing Then Exit Function
    keyText = VBA.Trim$(keyText)
    If VBA.Len(keyText) = 0 Then Exit Function

    On Error Resume Next
    If context.Exists(keyText) Then rawValueText = VBA.CStr(context(keyText))
    If Err.Number <> 0 Then
        Err.Clear
        rawValueText = VBA.CStr(VBA.CallByName(context, keyText, VbGet))
    End If
    On Error GoTo 0

    ' Готовый preview является многострочным документным текстом, а не
    ' скалярным значением шаблона. Его CR/LF и неразрывные пробелы должны
    ' пройти в Word без private_NormalizeTemplateScalar.
    If VBA.StrComp( _
        keyText, CONTEXT_WORD_PREVIEW_TEXT, VBA.vbTextCompare) = 0 Then
        private_GetContextText = rawValueText
    Else
        private_GetContextText = private_NormalizeTemplateScalar(rawValueText)
    End If
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
