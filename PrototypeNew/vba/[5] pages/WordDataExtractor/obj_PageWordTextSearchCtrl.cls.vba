VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PageWordTextSearchCtrl"
Option Explicit

Implements obj_IPageCtrl

Private Const OBJECT_KEY As String = "RuntimeObjects.WordDataExtractor.Controller"
Private Const TABLES_KEY As String = "RuntimeItems.WordDataExtractor.Tables"
Private Const SOURCES_KEY As String = "RuntimeItems.WordDataExtractor.Sources"
Private Const DEFAULT_CONTEXT_LINES_BEFORE As Long = 10
Private Const DEFAULT_CONTEXT_LINES_AFTER As Long = 5
Private Const MAX_CONTEXT_LINE_COUNT As Long = 100
Private Const RESULTS_CONTAINER_NAME As String = "SearchResultsPanel"
Private Const PREVIEW_CONTAINER_NAME As String = "SearchPreviewText"

Private m_Page As obj_IPage
Private m_DocumentPath As String
Private m_DocumentPathPattern As String
Private m_DocumentFilename As String
Private m_DocumentPaths As Collection
Private m_DocumentPathResolver As String
Private m_DocumentPathResolverArgs As String
Private m_SourceOptions As Collection
Private m_SourceIds As Collection
Private m_SourcePatterns As Object
Private m_SourceEnabled As Object
Private m_SourceCaptions As Object
Private m_DateFromDay As String
Private m_DateFromMonth As String
Private m_DateFromYear As String
Private m_DateToDay As String
Private m_DateToMonth As String
Private m_DateToYear As String
Private m_SearchText As String
Private m_IsRegexMode As Boolean
Private m_AllTables As Collection
Private m_ShowEmptyTables As Boolean
Private m_IsReady As Boolean
Private m_IsSearchRunning As Boolean
Private m_IsCancelRequested As Boolean
Private m_SelectedPreviewText As String
Private m_ContextLinesBefore As Long
Private m_ContextLinesAfter As Long

Private Function obj_IPageCtrl_Initialize( _
    ByVal page As obj_IPage _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim items As Collection

    If page Is Nothing Then Exit Function
    Set m_Page = page
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Not pageBase.RuntimeSources.SetObjectSource( _
        OBJECT_KEY, Me) Then Exit Function
    If Not pageBase.RegisterSelectionHandler( _
        Me, "OnResultSelectionChanged") Then Exit Function
    Set items = New Collection
    Set m_AllTables = New Collection
    m_ShowEmptyTables = False
    m_SelectedPreviewText = "Выберите строку результата"
    m_ContextLinesBefore = DEFAULT_CONTEXT_LINES_BEFORE
    m_ContextLinesAfter = DEFAULT_CONTEXT_LINES_AFTER
    If Not pageBase.RuntimeSources.SetItemsSource( _
        TABLES_KEY, items, False) Then Exit Function
    Set m_SourceOptions = New Collection
    Set m_SourceIds = New Collection
    Set m_SourcePatterns = VBA.CreateObject("Scripting.Dictionary")
    m_SourcePatterns.CompareMode = VBA.vbTextCompare
    Set m_SourceEnabled = VBA.CreateObject("Scripting.Dictionary")
    m_SourceEnabled.CompareMode = VBA.vbTextCompare
    Set m_SourceCaptions = VBA.CreateObject("Scripting.Dictionary")
    m_SourceCaptions.CompareMode = VBA.vbTextCompare
    If Not pageBase.RuntimeSources.SetItemsSource( _
        SOURCES_KEY, m_SourceOptions, False) Then Exit Function
    obj_IPageCtrl_Initialize = True
End Function

Private Function obj_IPageCtrl_UpdateData( _
    ByVal configControl As obj_ConfigControlVM _
) As Boolean
    Dim configTable As obj_ConfigTable
    Dim wordDataExtrCfgParser As obj_WordDataExtrCfgParser
    Dim documentDir As String
    Dim documentFilename As String
    Dim documentSources As String

    m_IsReady = False
    If configControl Is Nothing Then Exit Function
    If Not configControl.TryBuildConfigTableFromRendered( _
        configTable) Then Exit Function
    Set wordDataExtrCfgParser = New obj_WordDataExtrCfgParser
    If Not wordDataExtrCfgParser.Initialize(configTable) Then Exit Function

    m_DocumentPath = VBA.Trim$( _
        wordDataExtrCfgParser.GetOptionalValue( _
            "WordDataExtractor.DocumentPath"))
    If VBA.Len(m_DocumentPath) = 0 Then
        documentDir = VBA.Trim$( _
            wordDataExtrCfgParser.GetOptionalValue( _
                "WordDataExtractor.DocumentDir"))
        documentFilename = VBA.Trim$( _
            wordDataExtrCfgParser.GetOptionalValue( _
                "WordDataExtractor.DocumentFilename"))
        m_DocumentFilename = documentFilename
        documentSources = VBA.Trim$( _
            wordDataExtrCfgParser.GetOptionalValue( _
                "WordDataExtractor.DocumentSources"))
        If (VBA.Len(documentDir) = 0 Xor _
            VBA.Len(documentFilename) = 0) And _
            VBA.Len(documentSources) = 0 Then
            private_Error "Для поиска должны быть заполнены оба ключа: " & _
                "WordDataExtractor.DocumentDir и " & _
                "WordDataExtractor.DocumentFilename."
            Exit Function
        End If
        If VBA.Len(documentDir) > 0 Then
            m_DocumentPath = private_CombineDirectoryAndFilename( _
                documentDir, documentFilename)
            If VBA.Len(m_DocumentPath) = 0 Then Exit Function
        End If
    End If
    m_DocumentPath = private_EnsureDefaultDocumentExtension( _
        m_DocumentPath)
    m_DocumentPathPattern = m_DocumentPath

    m_DocumentPathResolver = VBA.Trim$( _
        wordDataExtrCfgParser.GetOptionalValue( _
            "WordDataExtractor.DocumentPathResolver"))
    m_DocumentPathResolverArgs = VBA.Trim$( _
        wordDataExtrCfgParser.GetOptionalValue( _
            "WordDataExtractor.DocumentPathResolverArgs"))
    If Not private_TryLoadDocumentSources( _
        wordDataExtrCfgParser) Then Exit Function
    If Not private_TryLoadDateParts( _
        private_GetResolverArg(m_DocumentPathResolverArgs, "dateFrom"), _
        "dateFrom", m_DateFromDay, m_DateFromMonth, _
        m_DateFromYear) Then Exit Function
    If Not private_TryLoadDateParts( _
        private_GetResolverArg(m_DocumentPathResolverArgs, "dateTo"), _
        "dateTo", m_DateToDay, m_DateToMonth, _
        m_DateToYear) Then Exit Function
    If Not private_ResolveDocuments() Then Exit Function

    m_IsReady = True
    obj_IPageCtrl_UpdateData = True
End Function

Private Sub obj_IPageCtrl_Dispose()
    Dim pageBase As obj_PageBase

    If Not m_Page Is Nothing Then
        Set pageBase = m_Page.GetPageBase()
        If Not pageBase Is Nothing Then _
            pageBase.ClearSelectionHandler Me
    End If
    Set m_AllTables = Nothing
    Set m_DocumentPaths = Nothing
    Set m_SourceOptions = Nothing
    Set m_SourceIds = Nothing
    Set m_SourcePatterns = Nothing
    Set m_SourceEnabled = Nothing
    Set m_SourceCaptions = Nothing
    Set m_Page = Nothing
End Sub

Public Property Get SearchText() As String
    SearchText = m_SearchText
End Property

Public Property Get ShowEmptyTables() As Boolean
    ShowEmptyTables = m_ShowEmptyTables
End Property

Public Property Get IsRegexMode() As Boolean
    IsRegexMode = m_IsRegexMode
End Property

Public Property Get SelectedPreviewText() As String
    SelectedPreviewText = m_SelectedPreviewText
End Property

Public Property Get ContextLinesBefore() As Long
    ContextLinesBefore = m_ContextLinesBefore
End Property

Public Property Get ContextLinesAfter() As Long
    ContextLinesAfter = m_ContextLinesAfter
End Property

Public Property Get DateFromDay() As String
    DateFromDay = m_DateFromDay
End Property

Public Property Get DateFromMonth() As String
    DateFromMonth = m_DateFromMonth
End Property

Public Property Get DateFromYear() As String
    DateFromYear = m_DateFromYear
End Property

Public Property Get DateToDay() As String
    DateToDay = m_DateToDay
End Property

Public Property Get DateToMonth() As String
    DateToMonth = m_DateToMonth
End Property

Public Property Get DateToYear() As String
    DateToYear = m_DateToYear
End Property

Public Function OnSearchTextChanged(Optional ByVal arg As Variant) As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim changedCell As Range
    Dim changedCellAddress As String

    If VBA.IsMissing(arg) Then Exit Function
    changedCellAddress = VBA.Trim$(VBA.CStr(arg))
    If VBA.Len(changedCellAddress) = 0 Then Exit Function
    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set changedCell = ws.Range(changedCellAddress)
    On Error GoTo 0
    If changedCell Is Nothing Then Exit Function

    m_SearchText = VBA.Trim$(VBA.CStr(changedCell.Value2))
    OnSearchTextChanged = True
End Function

Public Function OnContextLinesBeforeChanged( _
    Optional ByVal arg As Variant _
) As Boolean
    If VBA.IsMissing(arg) Then Exit Function
    OnContextLinesBeforeChanged = private_TryUpdateContextLineCount( _
        arg, "выше совпадения", m_ContextLinesBefore)
End Function

Public Function OnContextLinesAfterChanged( _
    Optional ByVal arg As Variant _
) As Boolean
    If VBA.IsMissing(arg) Then Exit Function
    OnContextLinesAfterChanged = private_TryUpdateContextLineCount( _
        arg, "ниже совпадения", m_ContextLinesAfter)
End Function

Private Function private_TryUpdateContextLineCount( _
    ByVal arg As Variant, _
    ByVal fieldCaption As String, _
    ByRef outLineCount As Long _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim changedCell As Range
    Dim rawValue As String
    Dim numericValue As Double

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Set changedCell = ws.Range(VBA.Trim$(VBA.CStr(arg)))
    On Error GoTo 0
    If changedCell Is Nothing Then Exit Function

    rawValue = VBA.Trim$(VBA.CStr(changedCell.Value2))
    If VBA.Len(rawValue) = 0 Or Not VBA.IsNumeric(rawValue) Then
        private_Error "Количество строк " & fieldCaption & _
            " должно быть целым числом от 0 до " & _
            VBA.CStr(MAX_CONTEXT_LINE_COUNT) & "."
        changedCell.Value2 = outLineCount
        Exit Function
    End If
    numericValue = VBA.CDbl(rawValue)
    If numericValue <> VBA.Fix(numericValue) Or numericValue < 0# Or _
        numericValue > MAX_CONTEXT_LINE_COUNT Then
        private_Error "Количество строк " & fieldCaption & _
            " должно быть целым числом от 0 до " & _
            VBA.CStr(MAX_CONTEXT_LINE_COUNT) & "."
        changedCell.Value2 = outLineCount
        Exit Function
    End If

    outLineCount = VBA.CLng(numericValue)
    private_TryUpdateContextLineCount = True
End Function
Public Function OnDateFromDayChanged(Optional ByVal arg As Variant) As Boolean
    If VBA.IsMissing(arg) Then Exit Function
    OnDateFromDayChanged = private_TryReadChangedCellText( _
        arg, m_DateFromDay)
End Function

Public Function OnDateFromMonthChanged(Optional ByVal arg As Variant) As Boolean
    If VBA.IsMissing(arg) Then Exit Function
    OnDateFromMonthChanged = private_TryReadChangedCellText( _
        arg, m_DateFromMonth)
End Function

Public Function OnDateFromYearChanged(Optional ByVal arg As Variant) As Boolean
    If VBA.IsMissing(arg) Then Exit Function
    OnDateFromYearChanged = private_TryReadChangedCellText( _
        arg, m_DateFromYear)
End Function

Public Function OnDateToDayChanged(Optional ByVal arg As Variant) As Boolean
    If VBA.IsMissing(arg) Then Exit Function
    OnDateToDayChanged = private_TryReadChangedCellText( _
        arg, m_DateToDay)
End Function

Public Function OnDateToMonthChanged(Optional ByVal arg As Variant) As Boolean
    If VBA.IsMissing(arg) Then Exit Function
    OnDateToMonthChanged = private_TryReadChangedCellText( _
        arg, m_DateToMonth)
End Function

Public Function OnDateToYearChanged(Optional ByVal arg As Variant) As Boolean
    If VBA.IsMissing(arg) Then Exit Function
    OnDateToYearChanged = private_TryReadChangedCellText( _
        arg, m_DateToYear)
End Function

Public Function SearchAndRender(Optional ByVal arg As Variant) As Boolean
    Dim resultTables As Collection
    Dim documentPathItem As Variant
    Dim documentText As String
    Dim documentTable As obj_TableDynamic
    Dim searchRegex As Object
    Dim documentIndex As Long
    Dim documentError As String
    Dim failedDocumentCount As Long

    If m_IsSearchRunning Then
        private_Error "Поиск уже выполняется."
        Exit Function
    End If

    If Not m_IsReady Then
        private_Error "Конфигурация поиска не готова. " & _
            "Повторно откройте режим с главной страницы."
        Exit Function
    End If
    ' Shape-click не всегда сопровождается SheetChange для input-ячейки:
    ' например после частичного render или если Excel завершил edit mode прямо
    ' кликом по Shape. Перед запуском считаем фактическое значение из UI, чтобы
    ' controller не расходился с текстом, который видит пользователь.
    If Not private_TrySyncSearchTextFromUi() Then Exit Function
    If VBA.Len(VBA.Trim$(m_SearchText)) = 0 Then
        private_Error "Введите часть текста для поиска в WORD-документах."
        Exit Function
    End If
    If Not private_ResolveDocuments() Then Exit Function
    If m_DocumentPaths Is Nothing Then
        private_Error "Список WORD-документов для поиска не подготовлен."
        Exit Function
    End If
    If m_DocumentPaths.Count = 0 Then
        private_Error "Список WORD-документов для поиска пуст."
        Exit Function
    End If
    If Not private_TryCreateSearchRegex(searchRegex) Then Exit Function

    m_IsSearchRunning = True
    m_IsCancelRequested = False
    m_SelectedPreviewText = "Выберите строку результата"
    rt_Messaging.fn_ShowStatusBarProgress _
        "WORD search", 0, m_DocumentPaths.Count
    Set resultTables = New Collection
    Set m_AllTables = resultTables
    For Each documentPathItem In m_DocumentPaths
        ' Поиск остаётся кооперативным: Excel обрабатывает нажатие кнопки
        ' отмены только между синхронными COM-вызовами Word.
        VBA.DoEvents
        If m_IsCancelRequested Then Exit For
        documentIndex = documentIndex + 1
        Set documentTable = Nothing
        If Not private_ReadWordDocument( _
            VBA.CStr(documentPathItem), documentText, _
            documentError) Then
            ' Ошибка одного файла не должна прерывать пакетный поиск.
            ' Диагностическая строка сохраняет имя проблемного документа
            ' в результатах и позволяет обработать остальные файлы.
            failedDocumentCount = failedDocumentCount + 1
            If Not private_BuildDocumentErrorTable( _
                VBA.CStr(documentPathItem), documentError, _
                documentIndex, documentTable) Then GoTo SearchFailed
        Else
            If Not private_BuildSearchResultTable( _
                VBA.CStr(documentPathItem), documentText, _
                documentIndex, searchRegex, _
                documentTable) Then GoTo SearchFailed
        End If
        resultTables.Add documentTable
        ' Публикуем накопленный результат сразу после появления новой
        ' видимой таблицы, не дожидаясь завершения всего списка файлов.
        If documentTable.RowCount > 0 Or m_ShowEmptyTables Then
            If Not private_PublishVisibleTables() Then GoTo SearchFailed
            If Not rt_PageManager.fn_RenderPage( _
                m_Page, "word-text-search:progress") Then GoTo SearchFailed
        End If
        rt_Messaging.fn_ShowStatusBarProgress _
            "WORD search", documentIndex, m_DocumentPaths.Count
    Next documentPathItem

    If Not private_PublishVisibleTables() Then GoTo SearchFailed
    If Not rt_PageManager.fn_RenderPage( _
        m_Page, "word-text-search:search") Then GoTo SearchFailed
    If m_IsCancelRequested Then
        rt_Messaging.fn_ShowStatusBarNotice _
            "WORD search отменён. Обработано документов: " & _
            VBA.CStr(documentIndex) & " из " & _
            VBA.CStr(m_DocumentPaths.Count) & ".", 4
    Else
        rt_Messaging.fn_ShowStatusBarSuccess _
            "WORD search: '" & m_SearchText & _
            "', обработано документов: " & _
            VBA.CStr(documentIndex) & ", строк результата: " & _
            VBA.CStr(private_CountTableRows(resultTables)) & _
            ", ошибок чтения: " & VBA.CStr(failedDocumentCount) & ".", 4
    End If
    m_IsSearchRunning = False
    SearchAndRender = True
    Exit Function

SearchFailed:
    m_IsSearchRunning = False
End Function

Private Function private_TrySyncSearchTextFromUi() As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim inputRange As Range
    Dim inputColumns As Range

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function

    If Not ex_ControlPartsRuntime.fn_TryResolveControlPartScope( _
        ws, "input", "SearchText", "cell", _
        inputRange, inputColumns) Then Exit Function
    If inputRange Is Nothing Then
        private_Error _
            "Не удалось определить ячейку поля текста для поиска. " & _
            "Повторно откройте режим с главной страницы."
        Exit Function
    End If

    m_SearchText = VBA.Trim$(VBA.CStr(inputRange.Cells(1, 1).Value2))
    private_TrySyncSearchTextFromUi = True
End Function

Public Function CancelSearch(Optional ByVal arg As Variant) As Boolean
    If Not m_IsSearchRunning Then
        rt_Messaging.fn_ShowStatusBarNotice _
            "WORD search сейчас не выполняется.", 3
        CancelSearch = True
        Exit Function
    End If
    m_IsCancelRequested = True
    rt_Messaging.fn_ShowStatusBarNotice _
        "Запрошена отмена WORD search...", 3
    CancelSearch = True
End Function

Public Function OnResultSelectionChanged( _
    Optional ByVal arg As Variant _
) As Boolean
    Dim selectedCell As Range
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim resultsRange As Range
    Dim markerCell As Range
    Dim selectedContextCell As Range
    Dim previewRange As Range
    Dim previewColumns As Range

    OnResultSelectionChanged = True
    If VBA.IsMissing(arg) Then Exit Function
    If Not VBA.IsObject(arg) Then Exit Function
    If Not TypeOf arg Is Range Then Exit Function
    Set selectedCell = arg.Cells(1, 1)
    ex_Core.fn_Diagnostic_LogInfo _
        "word-text-search:preview-selection target='" & _
        selectedCell.Address(False, False) & "'"
    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function
    If Not selectedCell.Worksheet Is ws Then Exit Function

    If Not pageBase.TryGetLayoutContainerRange( _
        RESULTS_CONTAINER_NAME, resultsRange) Then
        ex_Core.fn_Diagnostic_LogError _
            "word-text-search:preview-skip reason='results-range-missing'"
        Exit Function
    End If
    If resultsRange Is Nothing Then Exit Function
    If Application.Intersect(selectedCell, resultsRange) Is Nothing Then
        ex_Core.fn_Diagnostic_LogInfo _
            "word-text-search:preview-skip reason='outside-results'"
        Exit Function
    End If

    ' Схема результата принадлежит этому mode-specific контроллеру:
    ' первая колонка содержит номер/Ошибка, вторая — полный context.
    If selectedCell.Column = resultsRange.Column Then
        Set markerCell = selectedCell
        Set selectedContextCell = selectedCell.Offset(0, 1)
    ElseIf selectedCell.Column = resultsRange.Column + 1 Then
        Set markerCell = selectedCell.Offset(0, -1)
        Set selectedContextCell = selectedCell
    Else
        ex_Core.fn_Diagnostic_LogInfo _
            "word-text-search:preview-skip reason='outside-result-columns'"
        Exit Function
    End If
    If Not private_IsSearchResultMarker(markerCell.Value2) Then
        ex_Core.fn_Diagnostic_LogInfo _
            "word-text-search:preview-skip reason='not-data-row'"
        Exit Function
    End If
    If Not ex_ControlPartsRuntime.fn_TryResolveControlPartScope( _
        ws, "label", PREVIEW_CONTAINER_NAME, "cell", _
        previewRange, previewColumns) Then Exit Function
    If previewRange Is Nothing Then
        ex_Core.fn_Diagnostic_LogError _
            "word-text-search:preview-skip reason='preview-range-empty'"
        Exit Function
    End If

    ' Preview обновляется адресно, без полного render страницы: выделение
    ' остаётся на строке результата, а длинный текст не раздувает таблицу.
    m_SelectedPreviewText = _
        VBA.CStr(selectedContextCell.Cells(1, 1).Value2)
    previewRange.Cells(1, 1).Value2 = m_SelectedPreviewText
    ex_Core.fn_Diagnostic_LogInfo _
        "word-text-search:preview-updated target='" & _
        selectedCell.Address(False, False) & "' textLength=" & _
        VBA.CStr(VBA.Len(m_SelectedPreviewText))
End Function

Private Function private_IsSearchResultMarker( _
    ByVal markerValue As Variant _
) As Boolean
    Dim markerText As String

    markerText = VBA.Trim$(VBA.CStr(markerValue))
    If VBA.Len(markerText) = 0 Then Exit Function
    If VBA.IsNumeric(markerValue) Then
        private_IsSearchResultMarker = True
        Exit Function
    End If
    private_IsSearchResultMarker = ( _
        VBA.StrComp(markerText, "Ошибка", VBA.vbTextCompare) = 0)
End Function

Public Function ToggleRegexMode(Optional ByVal arg As Variant) As Boolean
    m_IsRegexMode = Not m_IsRegexMode
    If Not rt_PageManager.fn_RenderPage( _
        m_Page, "word-text-search:toggle-regex") Then Exit Function
    ToggleRegexMode = True
End Function

Public Function ToggleDocumentSource(Optional ByVal arg As Variant) As Boolean
    Dim sourceId As String

    If VBA.IsMissing(arg) Then Exit Function
    sourceId = VBA.Trim$(VBA.CStr(arg))
    If m_SourceEnabled Is Nothing Then Exit Function
    If Not m_SourceEnabled.Exists(sourceId) Then
        private_Error "Источник документов не найден в конфигурации: " & sourceId
        Exit Function
    End If

    m_SourceEnabled(sourceId) = Not VBA.CBool(m_SourceEnabled(sourceId))
    If Not private_RebuildSourceOptions() Then Exit Function
    If Not rt_PageManager.fn_RenderPage( _
        m_Page, "word-text-search:toggle-source") Then Exit Function
    ToggleDocumentSource = True
End Function

Public Function Rerender(Optional ByVal arg As Variant) As Boolean
    Rerender = rt_PageManager.fn_RenderPage( _
        m_Page, "word-text-search:rerender")
End Function

Public Function ToggleEmptyTables(Optional ByVal arg As Variant) As Boolean
    m_ShowEmptyTables = Not m_ShowEmptyTables
    If Not private_PublishVisibleTables() Then Exit Function
    If Not rt_PageManager.fn_RenderPage( _
        m_Page, "word-text-search:toggle-empty-tables") Then Exit Function
    ToggleEmptyTables = True
End Function

Private Function private_TryReadChangedCellText( _
    ByVal arg As Variant, _
    ByRef outText As String _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim changedCell As Range
    Dim changedCellAddress As String

    changedCellAddress = VBA.Trim$(VBA.CStr(arg))
    If VBA.Len(changedCellAddress) = 0 Then Exit Function
    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set changedCell = ws.Range(changedCellAddress)
    On Error GoTo 0
    If changedCell Is Nothing Then Exit Function
    outText = VBA.Trim$(VBA.CStr(changedCell.Value2))
    private_TryReadChangedCellText = True
End Function

Private Function private_TryLoadDateParts( _
    ByVal rawDateText As String, _
    ByVal fieldName As String, _
    ByRef outDay As String, _
    ByRef outMonth As String, _
    ByRef outYear As String _
) As Boolean
    Dim dateParts As Variant
    Dim normalizedDateText As String

    outDay = VBA.vbNullString
    outMonth = VBA.vbNullString
    outYear = VBA.vbNullString
    rawDateText = VBA.Trim$(rawDateText)
    If VBA.Len(rawDateText) = 0 Then
        private_TryLoadDateParts = True
        Exit Function
    End If

    rawDateText = VBA.Replace$(rawDateText, "/", ".")
    rawDateText = VBA.Replace$(rawDateText, "-", ".")
    dateParts = VBA.Split(rawDateText, ".")
    If UBound(dateParts) <> 2 Then
        private_Error "Дата " & fieldName & " в конфигурации должна " & _
            "иметь формат ДД.ММ.ГГГГ: '" & rawDateText & "'."
        Exit Function
    End If
    outDay = VBA.Trim$(VBA.CStr(dateParts(0)))
    outMonth = VBA.Trim$(VBA.CStr(dateParts(1)))
    outYear = VBA.Trim$(VBA.CStr(dateParts(2)))
    If Not private_TryBuildDateText( _
        fieldName, outDay, outMonth, outYear, _
        normalizedDateText) Then Exit Function
    dateParts = VBA.Split(normalizedDateText, ".")
    outDay = VBA.CStr(dateParts(0))
    outMonth = VBA.CStr(dateParts(1))
    outYear = VBA.CStr(dateParts(2))
    private_TryLoadDateParts = True
End Function

Private Function private_TryBuildDateText( _
    ByVal fieldName As String, _
    ByVal dayText As String, _
    ByVal monthText As String, _
    ByVal yearText As String, _
    ByRef outDateText As String _
) As Boolean
    Dim dayValue As Long
    Dim monthValue As Long
    Dim yearValue As Long
    Dim resolvedDate As Date

    outDateText = VBA.vbNullString
    dayText = VBA.Trim$(dayText)
    monthText = VBA.Trim$(monthText)
    yearText = VBA.Trim$(yearText)
    If VBA.Len(dayText) = 0 And VBA.Len(monthText) = 0 And _
        VBA.Len(yearText) = 0 Then
        private_TryBuildDateText = True
        Exit Function
    End If
    If VBA.Len(dayText) = 0 Or VBA.Len(monthText) = 0 Or _
        VBA.Len(yearText) = 0 Then
        private_Error "Для " & fieldName & _
            " заполните все три поля: день, месяц и год."
        Exit Function
    End If
    If Not VBA.IsNumeric(dayText) Or Not VBA.IsNumeric(monthText) Or _
        Not VBA.IsNumeric(yearText) Then
        private_Error "Для " & fieldName & _
            " день, месяц и год должны быть числами."
        Exit Function
    End If

    On Error GoTo InvalidDate
    dayValue = VBA.CLng(dayText)
    monthValue = VBA.CLng(monthText)
    yearValue = VBA.CLng(yearText)
    If yearValue < 1000 Or yearValue > 9999 Then GoTo InvalidDate
    resolvedDate = VBA.DateSerial(yearValue, monthValue, dayValue)
    If VBA.Day(resolvedDate) <> dayValue Or _
        VBA.Month(resolvedDate) <> monthValue Or _
        VBA.Year(resolvedDate) <> yearValue Then GoTo InvalidDate
    On Error GoTo 0

    outDateText = VBA.Format$(resolvedDate, "dd.mm.yyyy")
    private_TryBuildDateText = True
    Exit Function

InvalidDate:
    Err.Clear
    On Error GoTo 0
    private_Error "Для " & fieldName & " указана некорректная дата: " & _
        dayText & "." & monthText & "." & yearText & "."
End Function

Private Function private_TryLoadDocumentSources( _
    ByVal wordDataExtrCfgParser As obj_WordDataExtrCfgParser _
) As Boolean
    Dim sourceIdsText As String
    Dim sourceIds As Variant
    Dim sourceIdItem As Variant
    Dim sourceId As String
    Dim sourceDir As String
    Dim sourcePath As String
    Dim sourceFilename As String
    Dim sourceCaption As String
    Dim enabledText As String
    Dim isEnabled As Boolean

    If wordDataExtrCfgParser Is Nothing Then Exit Function
    Set m_SourcePatterns = VBA.CreateObject("Scripting.Dictionary")
    m_SourcePatterns.CompareMode = VBA.vbTextCompare
    Set m_SourceEnabled = VBA.CreateObject("Scripting.Dictionary")
    m_SourceEnabled.CompareMode = VBA.vbTextCompare
    Set m_SourceCaptions = VBA.CreateObject("Scripting.Dictionary")
    m_SourceCaptions.CompareMode = VBA.vbTextCompare
    Set m_SourceIds = New Collection

    sourceIdsText = VBA.Trim$(wordDataExtrCfgParser.GetOptionalValue( _
        "WordDataExtractor.DocumentSources"))
    If VBA.Len(sourceIdsText) = 0 Then
        private_TryLoadDocumentSources = private_RebuildSourceOptions()
        Exit Function
    End If

    sourceIds = VBA.Split(sourceIdsText, ";")
    For Each sourceIdItem In sourceIds
        sourceId = VBA.Trim$(VBA.CStr(sourceIdItem))
        If VBA.Len(sourceId) = 0 Then GoTo ContinueSource
        If m_SourcePatterns.Exists(sourceId) Then
            private_Error "Идентификатор источника документов указан дважды: " & sourceId
            Exit Function
        End If

        sourcePath = VBA.Trim$(wordDataExtrCfgParser.GetOptionalValue( _
            "WordDataExtractor.DocumentSource[" & sourceId & "].DocumentPath"))
        sourceDir = VBA.Trim$(wordDataExtrCfgParser.GetOptionalValue( _
            "WordDataExtractor.DocumentSource[" & sourceId & "].DocumentDir"))
        sourceFilename = VBA.Trim$(wordDataExtrCfgParser.GetOptionalValue( _
            "WordDataExtractor.DocumentSource[" & sourceId & "].DocumentFilename"))
        If VBA.Len(sourcePath) > 0 And VBA.Len(sourceDir) > 0 Then
            private_Error "Для источника '" & sourceId & _
                "' нельзя одновременно задавать DocumentPath и DocumentDir."
            Exit Function
        End If
        If VBA.Len(sourcePath) = 0 Then
            If VBA.Len(sourceDir) = 0 Then
                private_Error "Для источника '" & sourceId & _
                    "' не указан DocumentDir или DocumentPath."
                Exit Function
            End If
            If VBA.Len(sourceFilename) = 0 Then
                private_Error "Для источника '" & sourceId & _
                    "' не указан обязательный DocumentFilename."
                Exit Function
            End If
            sourcePath = private_CombineDirectoryAndFilename( _
                sourceDir, sourceFilename)
        End If
        sourcePath = private_EnsureDefaultDocumentExtension(sourcePath)

        sourceCaption = VBA.Trim$(wordDataExtrCfgParser.GetOptionalValue( _
            "WordDataExtractor.DocumentSource[" & sourceId & "].Caption", _
            sourcePath))
        enabledText = VBA.LCase$(VBA.Trim$(wordDataExtrCfgParser.GetOptionalValue( _
            "WordDataExtractor.DocumentSource[" & sourceId & "].Enabled", "true")))
        Select Case enabledText
            Case "true", "1", "yes", "on"
                isEnabled = True
            Case "false", "0", "no", "off"
                isEnabled = False
            Case Else
                private_Error "Некорректное значение Enabled у источника '" & _
                    sourceId & "': " & enabledText
                Exit Function
        End Select
        m_SourcePatterns.Add sourceId, sourcePath
        m_SourceCaptions.Add sourceId, sourceCaption
        m_SourceEnabled.Add sourceId, isEnabled
        m_SourceIds.Add sourceId
ContinueSource:
    Next sourceIdItem
    If m_SourcePatterns.Count = 0 Then
        private_Error "WordDataExtractor.DocumentSources не содержит источников."
        Exit Function
    End If
    private_TryLoadDocumentSources = private_RebuildSourceOptions()
End Function

Private Function private_RebuildSourceOptions() As Boolean
    Dim pageBase As obj_PageBase
    Dim sourceId As Variant
    Dim sourceOption As obj_SelectOption
    Dim marker As String

    Set m_SourceOptions = New Collection
    If Not m_SourceIds Is Nothing Then
        For Each sourceId In m_SourceIds
            Set sourceOption = New obj_SelectOption
            If Not sourceOption.Initialize() Then Exit Function
            sourceOption.Id = VBA.CStr(sourceId)
            If VBA.CBool(m_SourceEnabled(sourceId)) Then
                marker = VBA.ChrW$(&H2611) & " "
                If Not sourceOption.SetState("selected", True) Then Exit Function
            Else
                marker = VBA.ChrW$(&H2610) & " "
            End If
            sourceOption.Caption = marker & VBA.CStr(m_SourceCaptions(sourceId))
            m_SourceOptions.Add sourceOption
        Next sourceId
    End If
    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    private_RebuildSourceOptions = pageBase.RuntimeSources.SetItemsSource( _
        SOURCES_KEY, m_SourceOptions, False)
End Function

Private Sub private_SortDocumentRecords( _
    ByVal records As Collection, _
    ByVal sortDescending As Boolean _
)
    Dim i As Long
    Dim j As Long
    Dim leftRecord As Object
    Dim rightRecord As Object

    If records Is Nothing Then Exit Sub
    For i = 1 To records.Count - 1
        For j = i + 1 To records.Count
            Set leftRecord = records.Item(i)
            Set rightRecord = records.Item(j)
            If private_ShouldSwapDocumentRecords( _
                leftRecord, rightRecord, sortDescending) Then
                records.Remove j
                records.Add rightRecord, Before:=i
                Set leftRecord = records.Item(i)
            End If
        Next j
    Next i
End Sub

Private Function private_ShouldSwapDocumentRecords( _
    ByVal leftRecord As Object, _
    ByVal rightRecord As Object, _
    ByVal sortDescending As Boolean _
) As Boolean
    Dim dateCompare As Long
    Dim textCompare As Long

    If VBA.CDate(leftRecord("Date")) < VBA.CDate(rightRecord("Date")) Then
        dateCompare = -1
    ElseIf VBA.CDate(leftRecord("Date")) > VBA.CDate(rightRecord("Date")) Then
        dateCompare = 1
    End If
    If sortDescending Then dateCompare = -dateCompare
    If dateCompare <> 0 Then
        private_ShouldSwapDocumentRecords = (dateCompare > 0)
        Exit Function
    End If

    textCompare = VBA.StrComp(VBA.CStr(leftRecord("Name")), _
        VBA.CStr(rightRecord("Name")), VBA.vbTextCompare)
    If textCompare = 0 Then
        textCompare = VBA.StrComp(VBA.CStr(leftRecord("Path")), _
            VBA.CStr(rightRecord("Path")), VBA.vbTextCompare)
    End If
    If textCompare = 0 Then
        If VBA.CLng(leftRecord("Sequence")) < _
            VBA.CLng(rightRecord("Sequence")) Then
            textCompare = -1
        ElseIf VBA.CLng(leftRecord("Sequence")) > _
            VBA.CLng(rightRecord("Sequence")) Then
            textCompare = 1
        End If
    End If
    private_ShouldSwapDocumentRecords = (textCompare > 0)
End Function

Private Function private_GetFileName(ByVal filePath As String) As String
    Dim fso As Object

    If VBA.Len(VBA.Trim$(filePath)) = 0 Then Exit Function
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    private_GetFileName = fso.GetFileName(filePath)
End Function

Private Function private_ResolveDocuments() As Boolean
    Dim effectiveResolverArgs As String
    Dim dateFromText As String
    Dim dateToText As String
    Dim sourceId As Variant
    Dim sourcePattern As String
    Dim resolvedPaths As Collection
    Dim resolvedPath As Variant
    Dim resolvedDate As Date
    Dim records As Collection
    Dim record As Object
    Dim selectedCount As Long
    Dim sequence As Long
    Dim i As Long
    Dim seenPaths As Object
    Dim normalizedPath As String
    Dim sourceResolverArgs As String

    If Not private_TryBuildDateText( _
        "dateFrom", m_DateFromDay, m_DateFromMonth, _
        m_DateFromYear, dateFromText) Then Exit Function
    If Not private_TryBuildDateText( _
        "dateTo", m_DateToDay, m_DateToMonth, _
        m_DateToYear, dateToText) Then Exit Function
    effectiveResolverArgs = private_SetResolverArg( _
        m_DocumentPathResolverArgs, "dateFrom", dateFromText)
    effectiveResolverArgs = private_SetResolverArg( _
        effectiveResolverArgs, "dateTo", dateToText)
    Set m_DocumentPaths = Nothing

    If Not m_SourcePatterns Is Nothing Then
        If m_SourcePatterns.Count > 0 Then
            If VBA.StrComp(m_DocumentPathResolver, _
                "ResolveAllByDmyPattern", VBA.vbTextCompare) <> 0 Then
                private_Error "Для нескольких источников требуется " & _
                    "WordDataExtractor.DocumentPathResolver=" & _
                    "ResolveAllByDmyPattern."
                Exit Function
            End If
            Set records = New Collection
            Set seenPaths = VBA.CreateObject("Scripting.Dictionary")
            seenPaths.CompareMode = VBA.vbTextCompare
            sourceResolverArgs = private_SetResolverArg( _
                effectiveResolverArgs, "allowEmpty", "true")
            For Each sourceId In m_SourceIds
                If VBA.CBool(m_SourceEnabled(sourceId)) Then
                    selectedCount = selectedCount + 1
                    sourcePattern = VBA.CStr(m_SourcePatterns(sourceId))
                    On Error GoTo EH
                    Set resolvedPaths = ex_SourceResolver.fn_ResolveAllByDmyPattern( _
                        sourcePattern, sourceResolverArgs)
                    On Error GoTo 0
                    For Each resolvedPath In resolvedPaths
                        normalizedPath = VBA.Trim$(VBA.CStr(resolvedPath))
                        If seenPaths.Exists(normalizedPath) Then
                            GoTo ContinueResolvedPath
                        End If
                        If Not ex_SourceResolver.fn_TryGetDmyDateByResolvedPath( _
                            sourcePattern, normalizedPath, resolvedDate) Then
                            private_Error "Не удалось прочитать дату из имени документа: " & _
                                normalizedPath
                            Exit Function
                        End If
                        seenPaths.Add normalizedPath, True
                        sequence = sequence + 1
                        Set record = VBA.CreateObject("Scripting.Dictionary")
                        record("Path") = normalizedPath
                        record("Date") = resolvedDate
                        record("Name") = private_GetFileName(normalizedPath)
                        record("Sequence") = sequence
                        records.Add record
ContinueResolvedPath:
                    Next resolvedPath
            End If
            Next sourceId
            If selectedCount = 0 Then
                private_Error "Выберите хотя бы один источник WORD-документов."
                Exit Function
            End If
            If records.Count = 0 Then
                private_Error "В выбранных источниках не найдено WORD-документов."
                Exit Function
            End If
            private_SortDocumentRecords records, _
                (VBA.InStr(1, effectiveResolverArgs, "order=desc", _
                    VBA.vbTextCompare) > 0)
            Set m_DocumentPaths = New Collection
            For i = 1 To records.Count
                Set record = records.Item(i)
                m_DocumentPaths.Add VBA.CStr(record("Path"))
            Next i
            private_ResolveDocuments = True
            Exit Function
        End If
    End If

    If VBA.Len(m_DocumentPathResolver) = 0 Then
        If VBA.Len(dateFromText) > 0 Or _
            VBA.Len(dateToText) > 0 Then
            private_Error "Поля dateFrom/dateTo требуют " & _
                "WordDataExtractor.DocumentPathResolver."
            Exit Function
        End If
        Set m_DocumentPaths = New Collection
        If VBA.Len(m_DocumentPath) > 0 Then
            m_DocumentPaths.Add m_DocumentPath
        End If
    ElseIf VBA.StrComp(m_DocumentPathResolver, _
        "ResolveAllByDmyPattern", VBA.vbTextCompare) = 0 Then
        On Error GoTo EH
        Set m_DocumentPaths = _
            ex_SourceResolver.fn_ResolveAllByDmyPattern( _
                m_DocumentPathPattern, effectiveResolverArgs)
        On Error GoTo 0
    Else
        private_Error "Поисковый controller не поддерживает " & _
            "WordDataExtractor.DocumentPathResolver: " & _
            m_DocumentPathResolver
        Exit Function
    End If
    If m_DocumentPaths Is Nothing Then
        private_Error "Resolver документов не вернул коллекцию файлов."
        Exit Function
    End If
    If m_DocumentPaths.Count = 0 Then
        private_Error "Resolver документов не нашёл ни одного файла."
        Exit Function
    End If
    private_ResolveDocuments = True
    Exit Function

EH:
    private_Error "Не удалось разрешить список Word-документов для поиска: " & _
        Err.Description
    Err.Clear
    On Error GoTo 0
End Function

Private Function private_GetResolverArg( _
    ByVal resolverArgs As String, _
    ByVal argName As String _
) As String
    Dim tokens As Variant
    Dim token As Variant
    Dim separatorPosition As Long
    Dim keyText As String

    tokens = VBA.Split(resolverArgs, ";")
    For Each token In tokens
        separatorPosition = VBA.InStr(1, VBA.CStr(token), _
            "=", VBA.vbBinaryCompare)
        If separatorPosition <= 1 Then GoTo ContinueToken
        keyText = VBA.Trim$(VBA.Left$( _
            VBA.CStr(token), separatorPosition - 1))
        If VBA.StrComp(keyText, argName, VBA.vbTextCompare) = 0 Then
            private_GetResolverArg = VBA.Trim$(VBA.Mid$( _
                VBA.CStr(token), separatorPosition + 1))
            Exit Function
        End If
ContinueToken:
    Next token
End Function

Private Function private_SetResolverArg( _
    ByVal resolverArgs As String, _
    ByVal argName As String, _
    ByVal argValue As String _
) As String
    Dim tokens As Variant
    Dim token As Variant
    Dim tokenText As String
    Dim keyText As String
    Dim separatorPosition As Long
    Dim resultText As String

    tokens = VBA.Split(resolverArgs, ";")
    For Each token In tokens
        tokenText = VBA.Trim$(VBA.CStr(token))
        If VBA.Len(tokenText) = 0 Then GoTo ContinueToken
        separatorPosition = VBA.InStr(1, tokenText, _
            "=", VBA.vbBinaryCompare)
        keyText = tokenText
        If separatorPosition > 1 Then
            keyText = VBA.Trim$(VBA.Left$( _
                tokenText, separatorPosition - 1))
        End If
        If VBA.StrComp(keyText, argName, _
            VBA.vbTextCompare) <> 0 Then
            resultText = private_AppendResolverToken( _
                resultText, tokenText)
        End If
ContinueToken:
    Next token

    argValue = VBA.Trim$(argValue)
    If VBA.Len(argValue) > 0 Then
        resultText = private_AppendResolverToken( _
            resultText, argName & "=" & argValue)
    End If
    private_SetResolverArg = resultText
End Function

Private Function private_AppendResolverToken( _
    ByVal resolverArgs As String, _
    ByVal tokenText As String _
) As String
    If VBA.Len(resolverArgs) = 0 Then
        private_AppendResolverToken = tokenText
    Else
        private_AppendResolverToken = _
            resolverArgs & ";" & tokenText
    End If
End Function

Private Function private_PublishVisibleTables() As Boolean
    Dim pageBase As obj_PageBase
    Dim visibleTables As Collection
    Dim tableItem As Variant
    Dim tableObj As obj_TableDynamic

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set visibleTables = New Collection
    If Not m_AllTables Is Nothing Then
        For Each tableItem In m_AllTables
            Set tableObj = tableItem
            If tableObj Is Nothing Then Exit Function
            If m_ShowEmptyTables Or tableObj.RowCount > 0 Then
                visibleTables.Add tableObj
            End If
        Next tableItem
    End If
    private_PublishVisibleTables = _
        pageBase.RuntimeSources.SetItemsSource( _
            TABLES_KEY, visibleTables, False)
End Function

Private Function private_BuildSearchResultTable( _
    ByVal documentPath As String, _
    ByVal documentText As String, _
    ByVal documentIndex As Long, _
    ByVal searchRegex As Object, _
    ByRef outTable As obj_TableDynamic _
) As Boolean
    Dim textLines As Collection
    Dim lineText As Variant
    Dim lineIndex As Long
    Dim matchIndex As Long
    Dim contextText As String
    Dim rowObj As obj_Row
    Dim fso As Object

    Set outTable = New obj_TableDynamic
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    outTable.SectionTitle = fso.GetFileName(documentPath)
    outTable.SourceAlias = "text-search-" & VBA.CStr(documentIndex)
    outTable.SourceAliasTemplate = "WordDataExtractorTextSearch"
    If Not private_AddSearchColumn(outTable, _
        "№ запису", "matchIndex") Then Exit Function
    If Not private_AddSearchColumn(outTable, _
        "Контекст", "context") Then Exit Function
    Set textLines = private_CollectLines(documentText)

    For Each lineText In textLines
        lineIndex = lineIndex + 1
        If private_IsSearchMatch( _
            VBA.CStr(lineText), searchRegex) Then
            matchIndex = matchIndex + 1
            contextText = private_BuildLineContext( _
                textLines, lineIndex)
            Set rowObj = New obj_Row
            rowObj.PushCellRaw VBA.CStr(matchIndex)
            rowObj.PushCellRaw contextText
            If Not outTable.PushRow(rowObj) Then Exit Function
        End If
    Next lineText
    private_BuildSearchResultTable = True
End Function

Private Function private_BuildDocumentErrorTable( _
    ByVal documentPath As String, _
    ByVal errorText As String, _
    ByVal documentIndex As Long, _
    ByRef outTable As obj_TableDynamic _
) As Boolean
    Dim rowObj As obj_Row
    Dim fso As Object

    Set outTable = New obj_TableDynamic
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    outTable.SectionTitle = fso.GetFileName(documentPath)
    outTable.SourceAlias = "text-search-" & VBA.CStr(documentIndex)
    outTable.SourceAliasTemplate = "WordDataExtractorTextSearch"
    If Not private_AddSearchColumn(outTable, _
        "№ запису", "matchIndex") Then Exit Function
    If Not private_AddSearchColumn(outTable, _
        "Контекст", "context") Then Exit Function

    Set rowObj = New obj_Row
    rowObj.PushCellRaw "Ошибка"
    rowObj.PushCellRaw "Документ пропущен: " & errorText
    If Not outTable.PushRow(rowObj) Then Exit Function
    private_BuildDocumentErrorTable = True
End Function

Private Function private_TryCreateSearchRegex( _
    ByRef outRegex As Object _
) As Boolean
    Set outRegex = Nothing
    If Not m_IsRegexMode Then
        private_TryCreateSearchRegex = True
        Exit Function
    End If

    On Error GoTo EH
    Set outRegex = VBA.CreateObject("VBScript.RegExp")
    outRegex.Pattern = m_SearchText
    outRegex.IgnoreCase = True
    outRegex.Global = False
    outRegex.MultiLine = False
    Call outRegex.Test(VBA.vbNullString)
    private_TryCreateSearchRegex = True
    Exit Function

EH:
    Set outRegex = Nothing
    private_Error "Некорректное регулярное выражение '" & _
        m_SearchText & "': " & Err.Description
    Err.Clear
    On Error GoTo 0
End Function

Private Function private_IsSearchMatch( _
    ByVal lineText As String, _
    ByVal searchRegex As Object _
) As Boolean
    If m_IsRegexMode Then
        If searchRegex Is Nothing Then Exit Function
        private_IsSearchMatch = searchRegex.Test(lineText)
    Else
        private_IsSearchMatch = _
            (VBA.InStr(1, lineText, m_SearchText, _
                VBA.vbTextCompare) > 0)
    End If
End Function

Private Function private_AddSearchColumn( _
    ByVal tableObj As obj_TableDynamic, _
    ByVal captionText As String, _
    ByVal aliasText As String _
) As Boolean
    Dim columnObj As obj_Column

    If tableObj Is Nothing Then Exit Function
    Set columnObj = New obj_Column
    columnObj.Name = captionText
    If Not columnObj.AddAlias(aliasText) Then Exit Function
    private_AddSearchColumn = tableObj.PushColumn(columnObj)
End Function

Private Function private_CollectLines( _
    ByVal documentText As String _
) As Collection
    Dim result As Collection
    Dim rawLines As Variant
    Dim rawLine As Variant
    Dim lineText As String

    Set result = New Collection
    documentText = VBA.Replace( _
        documentText, VBA.Chr$(11), VBA.vbCr)
    rawLines = VBA.Split(documentText, VBA.vbCr)
    For Each rawLine In rawLines
        lineText = VBA.Trim$(VBA.CStr(rawLine))
        ' Пустая строка является частью исходного контекста Word. Она также
        ' учитывается в заданных границах контекста и сохраняет разрывы между блоками.
        result.Add lineText
    Next rawLine
    Set private_CollectLines = result
End Function

Private Function private_BuildLineContext( _
    ByVal textLines As Collection, _
    ByVal lineIndex As Long _
) As String
    Dim contextText As String
    Dim firstLineIndex As Long
    Dim lastLineIndex As Long
    Dim currentLineIndex As Long

    If textLines Is Nothing Then Exit Function
    firstLineIndex = lineIndex - m_ContextLinesBefore
    If firstLineIndex < 1 Then firstLineIndex = 1
    lastLineIndex = lineIndex + m_ContextLinesAfter
    If lastLineIndex > textLines.Count Then
        lastLineIndex = textLines.Count
    End If

    For currentLineIndex = firstLineIndex To lastLineIndex
        If VBA.Len(contextText) > 0 Then
            contextText = contextText & VBA.vbLf
        End If
        contextText = contextText & _
            VBA.CStr(textLines.Item(currentLineIndex))
    Next currentLineIndex
    private_BuildLineContext = contextText
End Function

Private Function private_CountTableRows(ByVal tables As Collection) As Long
    Dim tableItem As Variant
    Dim tableObj As obj_TableDynamic

    If tables Is Nothing Then Exit Function
    For Each tableItem In tables
        Set tableObj = tableItem
        If Not tableObj Is Nothing Then
            private_CountTableRows = _
                private_CountTableRows + tableObj.RowCount
        End If
    Next tableItem
End Function

Private Function private_ReadWordDocument( _
    ByVal filePath As String, _
    ByRef outText As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim fso As Object
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim errDescription As String

    outText = VBA.vbNullString
    outErrorText = VBA.vbNullString
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        outErrorText = "WORD-документ не найден: " & filePath
        ex_Core.fn_Diagnostic_LogError _
            "WordTextSearch: " & outErrorText
        Exit Function
    End If

    On Error GoTo EH
    ex_Core.fn_Diagnostic_LogInfo _
        "word-text-search:document-read-start path='" & _
        VBA.Replace$(filePath, "'", "''") & "'"
    If Not rt_WordExportRuntime.fn_GetOrCreateWordApp( _
        wordApp) Then
        outErrorText = "Не удалось запустить Microsoft Word."
        ex_Core.fn_Diagnostic_LogError _
            "WordTextSearch: " & outErrorText
        Exit Function
    End If

    wordApp.DisplayAlerts = 0
    Set wordDoc = wordApp.Documents.Open(filePath, False, True, False)
    If Not private_TryBuildDocumentText(wordDoc, outText) Then
        VBA.Err.Raise VBA.vbObjectError + 2102, _
            "WordTextSearch", _
            "Не удалось собрать текст из WORD-документа."
    End If
    wordDoc.Close False
    Set wordDoc = Nothing
    ex_Core.fn_Diagnostic_LogInfo _
        "word-text-search:document-read-done path='" & _
        VBA.Replace$(filePath, "'", "''") & "'"
    private_ReadWordDocument = True
    Exit Function

EH:
    errDescription = Err.Description
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close False
    On Error GoTo 0
    outErrorText = "Не удалось прочитать WORD-документ: " & _
        errDescription
    ex_Core.fn_Diagnostic_LogError _
        "WordTextSearch: " & outErrorText & "; путь: " & filePath
End Function

Private Function private_TryBuildDocumentText( _
    ByVal wordDoc As Object, _
    ByRef outText As String _
) As Boolean
    Dim resultText As String

    outText = VBA.vbNullString
    If wordDoc Is Nothing Then Exit Function
    On Error GoTo EH
    resultText = VBA.CStr(wordDoc.Content.Text)
    resultText = private_NormalizeWordWhitespace(resultText)
    resultText = VBA.Replace(resultText, VBA.Chr$(7), VBA.vbCr)
    outText = resultText
    private_TryBuildDocumentText = (VBA.Len(outText) > 0)
    Exit Function

EH:
    ex_Core.fn_Diagnostic_LogError _
        "WordTextSearch: ошибка сборки текста документа: " & _
        Err.Description
End Function
Private Function private_NormalizeWordWhitespace( _
    ByVal sourceText As String _
) As String
    Dim codePoint As Long

    sourceText = VBA.Replace(sourceText, VBA.ChrW$(160), " ")
    sourceText = VBA.Replace(sourceText, VBA.ChrW$(5760), " ")
    For codePoint = 8192 To 8202
        sourceText = VBA.Replace( _
            sourceText, VBA.ChrW$(codePoint), " ")
    Next codePoint
    sourceText = VBA.Replace(sourceText, VBA.ChrW$(8239), " ")
    sourceText = VBA.Replace(sourceText, VBA.ChrW$(8287), " ")
    sourceText = VBA.Replace(sourceText, VBA.ChrW$(12288), " ")
    sourceText = VBA.Replace( _
        sourceText, VBA.ChrW$(8203), VBA.vbNullString)
    sourceText = VBA.Replace( _
        sourceText, VBA.ChrW$(8288), VBA.vbNullString)
    sourceText = VBA.Replace( _
        sourceText, VBA.ChrW$(65279), VBA.vbNullString)
    private_NormalizeWordWhitespace = sourceText
End Function

Private Function private_EnsureDefaultDocumentExtension( _
    ByVal documentPath As String _
) As String
    Dim lastSeparatorPosition As Long
    Dim lastDotPosition As Long

    documentPath = VBA.Trim$(documentPath)
    If VBA.Len(documentPath) = 0 Then Exit Function
    lastSeparatorPosition = VBA.InStrRev(documentPath, "\")
    If VBA.InStrRev(documentPath, "/") > _
        lastSeparatorPosition Then
        lastSeparatorPosition = VBA.InStrRev(documentPath, "/")
    End If
    lastDotPosition = VBA.InStrRev(documentPath, ".")
    If lastDotPosition <= lastSeparatorPosition Then
        documentPath = documentPath & ".docx"
    ElseIf lastDotPosition = VBA.Len(documentPath) Then
        documentPath = documentPath & "docx"
    End If
    private_EnsureDefaultDocumentExtension = documentPath
End Function

Private Function private_CombineDirectoryAndFilename( _
    ByVal directoryPath As String, _
    ByVal filename As String _
) As String
    Dim trailingChar As String

    directoryPath = VBA.Trim$(directoryPath)
    filename = VBA.Trim$(filename)
    If VBA.Len(directoryPath) = 0 Or _
        VBA.Len(filename) = 0 Then Exit Function
    trailingChar = VBA.Right$(directoryPath, 1)
    If trailingChar = "\" Or trailingChar = "/" Then
        private_CombineDirectoryAndFilename = _
            directoryPath & filename
    Else
        private_CombineDirectoryAndFilename = _
            directoryPath & "\" & filename
    End If
End Function

Private Sub private_Error(ByVal messageText As String)
    ex_Core.fn_Diagnostic_LogError _
        "WordTextSearch: " & messageText
    VBA.MsgBox messageText, VBA.vbExclamation, _
        "PrototypeNew / Word text search"
End Sub
