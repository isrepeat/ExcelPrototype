VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PageWordDataExtractorCtrl"
Option Explicit

Implements obj_IPageCtrl

Private Const OBJECT_KEY As String = "RuntimeObjects.WordDataExtractor.Controller"
Private Const TABLES_KEY As String = "RuntimeItems.WordDataExtractor.Tables"
Private m_Page As obj_IPage
Private m_RulesPath As String
Private m_PipelineId As String
Private m_DocumentPath As String
Private m_DocumentPathPattern As String
Private m_DocumentPaths As Collection
Private m_DocumentDateColumnCaption As String
Private m_TransformerClassName As String
Private m_ConfigTable As obj_ConfigTable
Private m_AllTables As Collection
Private m_ShowEmptyTables As Boolean
Private m_IsReady As Boolean

Public Property Get RuntimeObjectSourceKey() As String
    RuntimeObjectSourceKey = OBJECT_KEY
End Property

Public Property Get DocumentPath() As String
    DocumentPath = m_DocumentPath
End Property

Public Function Initialize(ByVal page As obj_IPage) As Boolean
    Dim pageBase As obj_PageBase, items As Collection
    If page Is Nothing Then Exit Function
    Set m_Page = page
    Set pageBase = page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Not pageBase.RuntimeSources.SetObjectSource(OBJECT_KEY, Me) Then Exit Function
    Set items = New Collection
    Set m_AllTables = New Collection
    m_ShowEmptyTables = True
    If Not pageBase.RuntimeSources.SetItemsSource(TABLES_KEY, items, False) Then Exit Function
    Initialize = True
End Function

Private Function obj_IPageCtrl_Initialize( _
    ByVal page As obj_IPage _
) As Boolean
    obj_IPageCtrl_Initialize = Initialize(page)
End Function

Public Sub Dispose()
    Set m_ConfigTable = Nothing
    Set m_AllTables = Nothing
    Set m_DocumentPaths = Nothing
    Set m_Page = Nothing
End Sub

Private Sub obj_IPageCtrl_Dispose()
    Dispose
End Sub

Public Property Get ShowEmptyTables() As Boolean
    ShowEmptyTables = m_ShowEmptyTables
End Property

Public Function UpdateData(ByVal configControl As obj_ConfigControlVM) As Boolean
    Dim configTable As obj_ConfigTable
    Dim wordDataExtrCfgParser As obj_WordDataExtrCfgParser
    Dim documentDir As String
    Dim documentFilename As String
    Dim documentPathResolver As String
    Dim documentPathResolverArgs As String
    m_IsReady = False
    If configControl Is Nothing Then Exit Function
    If Not configControl.TryBuildConfigTableFromRendered(configTable) Then Exit Function
    Set wordDataExtrCfgParser = New obj_WordDataExtrCfgParser
    If Not wordDataExtrCfgParser.Initialize(configTable) Then Exit Function
    If Not wordDataExtrCfgParser.TryGetRequiredValue( _
        "WordDataExtractor.RulesFile", m_RulesPath) Then
        private_Error "В профиле отсутствует обязательный ключ WordDataExtractor.RulesFile."
        Exit Function
    End If
    If Not wordDataExtrCfgParser.TryGetRequiredValue( _
        "WordDataExtractor.Pipeline", m_PipelineId) Then
        private_Error "В профиле отсутствует обязательный ключ WordDataExtractor.Pipeline."
        Exit Function
    End If
    m_DocumentPath = VBA.Trim$( _
        wordDataExtrCfgParser.GetOptionalValue( _
            "WordDataExtractor.DocumentPath"))
    If VBA.Len(m_DocumentPath) = 0 Then
        ' Раздельная запись является альтернативой, а не дополнением:
        ' непустой DocumentPath всегда имеет приоритет над Dir/Filename.
        documentDir = VBA.Trim$( _
            wordDataExtrCfgParser.GetOptionalValue( _
                "WordDataExtractor.DocumentDir"))
        documentFilename = VBA.Trim$( _
            wordDataExtrCfgParser.GetOptionalValue( _
                "WordDataExtractor.DocumentFilename"))
        If VBA.Len(documentDir) = 0 Xor VBA.Len(documentFilename) = 0 Then
            private_Error "Для альтернативного пути должны быть заполнены оба ключа: " & _
                "WordDataExtractor.DocumentDir и WordDataExtractor.DocumentFilename."
            Exit Function
        End If
        If VBA.Len(documentDir) > 0 Then
            m_DocumentPath = private_CombineDirectoryAndFilename( _
                documentDir, documentFilename)
            If VBA.Len(m_DocumentPath) = 0 Then Exit Function
        End If
    End If
    m_DocumentPath = private_EnsureDefaultDocumentExtension(m_DocumentPath)
    m_DocumentPathPattern = m_DocumentPath
    documentPathResolver = VBA.Trim$( _
        wordDataExtrCfgParser.GetOptionalValue( _
            "WordDataExtractor.DocumentPathResolver"))
    documentPathResolverArgs = VBA.Trim$( _
        wordDataExtrCfgParser.GetOptionalValue( _
            "WordDataExtractor.DocumentPathResolverArgs"))
    Set m_DocumentPaths = Nothing
    If VBA.Len(documentPathResolver) = 0 Then
        Set m_DocumentPaths = New Collection
        If VBA.Len(m_DocumentPath) > 0 Then _
            m_DocumentPaths.Add m_DocumentPath
    ElseIf VBA.StrComp(documentPathResolver, _
        "ResolveAllByDmyPattern", VBA.vbTextCompare) = 0 Then
        On Error GoTo EH_RESOLVE_DOCUMENTS
        Set m_DocumentPaths = _
            ex_SourceResolver.fn_ResolveAllByDmyPattern( _
                m_DocumentPathPattern, documentPathResolverArgs)
        On Error GoTo 0
    Else
        private_Error "Не поддерживается WordDataExtractor.DocumentPathResolver: " & _
            documentPathResolver
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
    m_DocumentDateColumnCaption = VBA.Trim$( _
        wordDataExtrCfgParser.GetOptionalValue( _
            "WordDataExtractor.DocumentDateColumnCaption"))
    m_TransformerClassName = _
        wordDataExtrCfgParser.GetOptionalValue( _
            "WordDataExtractor.TransformerClass")
    Set m_ConfigTable = configTable
    m_IsReady = True
    UpdateData = True
    Exit Function
EH_RESOLVE_DOCUMENTS:
    private_Error "Не удалось разрешить список Word-документов: " & _
        Err.Description
    Err.Clear
    On Error GoTo 0
End Function

Private Function obj_IPageCtrl_UpdateData( _
    ByVal configControl As obj_ConfigControlVM _
) As Boolean
    obj_IPageCtrl_UpdateData = UpdateData(configControl)
End Function

Private Function private_EnsureDefaultDocumentExtension( _
    ByVal documentPath As String _
) As String
    Dim lastSeparatorPosition As Long
    Dim lastDotPosition As Long

    documentPath = VBA.Trim$(documentPath)
    If VBA.Len(documentPath) = 0 Then Exit Function

    lastSeparatorPosition = VBA.InStrRev(documentPath, "\")
    If VBA.InStrRev(documentPath, "/") > lastSeparatorPosition Then
        lastSeparatorPosition = VBA.InStrRev(documentPath, "/")
    End If
    lastDotPosition = VBA.InStrRev(documentPath, ".")

    ' Точка в имени директории не является расширением файла.
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
    If VBA.Len(directoryPath) = 0 Or VBA.Len(filename) = 0 Then Exit Function

    trailingChar = VBA.Right$(directoryPath, 1)
    If trailingChar = "\" Or trailingChar = "/" Then
        private_CombineDirectoryAndFilename = directoryPath & filename
    Else
        private_CombineDirectoryAndFilename = directoryPath & "\" & filename
    End If
End Function

Public Function ExtractAndRender(Optional ByVal arg As Variant) As Boolean
    Dim engine As obj_WDE_RulesEngine
    Dim tables As Collection
    Dim transformedTables As Collection
    Dim resultTables As Collection
    Dim documentText As String
    Dim documentPathItem As Variant
    Dim documentDateText As String
    Dim pageBase As obj_PageBase
    If Not m_IsReady Then
        private_Error "Конфигурация не готова. Повторно откройте режим с главной страницы."
        Exit Function
    End If
    If m_DocumentPaths Is Nothing Then
        private_Error "Список Word-документов не подготовлен."
        Exit Function
    End If
    If m_DocumentPaths.Count = 0 Then
        private_Error "В текущем профиле не заполнен WordDataExtractor.DocumentPath " & _
            "или пара WordDataExtractor.DocumentDir/DocumentFilename."
        Exit Function
    End If
    Set engine = New obj_WDE_RulesEngine
    If Not engine.Initialize(m_RulesPath) Then Exit Function
    Set resultTables = New Collection
    For Each documentPathItem In m_DocumentPaths
        If Not private_ReadWordDocument( _
            VBA.CStr(documentPathItem), documentText) Then Exit Function
        If Not engine.ExtractTables( _
            m_PipelineId, documentText, tables) Then Exit Function
        If Not private_TryTransformTables( _
            tables, transformedTables) Then Exit Function
        If Not private_TryAddPersonnelCountToDatesTable( _
            transformedTables) Then Exit Function
        documentDateText = VBA.vbNullString
        If VBA.Len(m_DocumentDateColumnCaption) > 0 Then
            If Not private_TryResolveDocumentDate( _
                VBA.CStr(documentPathItem), documentDateText) Then _
                Exit Function
        End If
        If Not private_MergeDocumentTables(resultTables, _
            transformedTables, documentDateText) Then Exit Function
    Next documentPathItem
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set m_AllTables = resultTables
    If Not private_PublishVisibleTables() Then Exit Function
    If Not rt_PageManager.fn_RenderPage(m_Page, "worddataextractor:extract") Then Exit Function
    rt_Messaging.fn_ShowStatusBarSuccess _
        "WordDataExtractor: обработано документов: " & _
        VBA.CStr(m_DocumentPaths.Count) & ", извлечено таблиц: " & _
        VBA.CStr(resultTables.Count) & ". Режим: " & _
        private_CurrentTablesModeCaption() & ".", 4
    ExtractAndRender = True
End Function

Private Function private_TryAddPersonnelCountToDatesTable( _
    ByVal tables As Collection _
) As Boolean
    Const DATES_TABLE_ALIAS As String = "order-dates"
    Const FIO_COLUMN_ALIAS As String = "fio"
    Const PERSON_COUNT_COLUMN_ALIAS As String = "personCount"

    Dim tableItem As Variant
    Dim tableObj As obj_TableDynamic
    Dim datesTable As obj_TableDynamic
    Dim personCountColumn As obj_Column
    Dim rowObj As obj_Row
    Dim ignoredFioColumnIndex As Long
    Dim rowIndex As Long
    Dim personCount As Long

    If tables Is Nothing Then Exit Function

    For Each tableItem In tables
        Set tableObj = tableItem
        If tableObj Is Nothing Then Exit Function
        If VBA.StrComp(VBA.Trim$(tableObj.SourceAlias), DATES_TABLE_ALIAS, _
            VBA.vbTextCompare) = 0 Then
            Set datesTable = tableObj
        ElseIf tableObj.TryGetColumnIndexByAlias( _
            FIO_COLUMN_ALIAS, ignoredFioColumnIndex) Then
            personCount = personCount + tableObj.RowCount
        End If
    Next tableItem

    If datesTable Is Nothing Then
        private_Error "В результатах WordDataExtractor отсутствует обязательная таблица 'Дати'."
        Exit Function
    End If
    If datesTable.RowCount = 0 Then
        private_Error "Обязательная таблица 'Дати' не содержит строку приказа."
        Exit Function
    End If

    Set personCountColumn = New obj_Column
    personCountColumn.Name = "Кількість осіб"
    If Not personCountColumn.AddAlias(PERSON_COUNT_COLUMN_ALIAS) Then Exit Function
    If Not datesTable.InsertColumnAt( _
        personCountColumn, datesTable.ColumnCount + 1) Then Exit Function

    For rowIndex = 1 To datesTable.RowCount
        Set rowObj = datesTable.Rows.Item(rowIndex)
        If rowObj Is Nothing Then Exit Function
        If Not rowObj.SetCellRaw( _
            datesTable.ColumnCount, VBA.CStr(personCount)) Then Exit Function
    Next rowIndex

    private_TryAddPersonnelCountToDatesTable = True
End Function

Public Function ToggleEmptyTables(Optional ByVal arg As Variant) As Boolean
    m_ShowEmptyTables = Not m_ShowEmptyTables
    If Not private_PublishVisibleTables() Then Exit Function
    If Not rt_PageManager.fn_RenderPage( _
        m_Page, "worddataextractor:toggle-empty-tables") Then Exit Function
    rt_Messaging.fn_ShowStatusBarSuccess _
        "WordDataExtractor: " & private_CurrentTablesModeCaption() & ".", 3
    ToggleEmptyTables = True
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
            Select Case VBA.LCase$(VBA.Trim$(tableObj.SourceAlias))
                Case "returned-from-annual-leave"
                    If Not private_AddVisibleTableByAlias(visibleTables, _
                        "returned-from-annual-leave") Then Exit Function
                    If Not private_AddVisibleTableByAlias(visibleTables, _
                        "returned-from-family-leave") Then Exit Function
                    If Not private_AddVisibleTableByAlias(visibleTables, _
                        "returned-from-inpatient-vlk") Then Exit Function
                Case "returned-from-family-leave", _
                     "returned-from-inpatient-vlk"
                    ' Эти таблицы добавляются рядом с ежегодным отпуском.
                Case Else
                    If m_ShowEmptyTables Or tableObj.RowCount > 0 Then
                        visibleTables.Add tableObj
                    End If
            End Select
        Next tableItem
    End If

    private_PublishVisibleTables = pageBase.RuntimeSources.SetItemsSource( _
        TABLES_KEY, visibleTables, False)
End Function

Private Function private_AddVisibleTableByAlias( _
    ByVal targetTables As Collection, _
    ByVal sourceAlias As String _
) As Boolean
    Dim tableItem As Variant
    Dim tableObj As obj_TableDynamic

    If targetTables Is Nothing Or m_AllTables Is Nothing Then Exit Function
    For Each tableItem In m_AllTables
        Set tableObj = tableItem
        If tableObj Is Nothing Then Exit Function
        If VBA.StrComp(VBA.Trim$(tableObj.SourceAlias), sourceAlias, _
            vbTextCompare) = 0 Then
            If m_ShowEmptyTables Or tableObj.RowCount > 0 Then
                targetTables.Add tableObj
            End If
            private_AddVisibleTableByAlias = True
            Exit Function
        End If
    Next tableItem
End Function

Private Function private_CurrentTablesModeCaption() As String
    If m_ShowEmptyTables Then
        private_CurrentTablesModeCaption = "показаны все таблицы"
    Else
        private_CurrentTablesModeCaption = "показаны только непустые таблицы"
    End If
End Function

Private Function private_TryResolveDocumentDate( _
    ByVal resolvedPath As String, _
    ByRef outDateText As String _
) As Boolean
    outDateText = VBA.vbNullString
    If VBA.Len(VBA.Trim$(m_DocumentPathPattern)) = 0 Then
        private_Error "Не задан шаблон пути для извлечения даты документа."
        Exit Function
    End If
    On Error GoTo EH
    outDateText = ex_SourceResolver.fn_ExpandDmyRuntimeAliasByResolvedPath( _
        "{dd}.{mm}.{yyyy}", m_DocumentPathPattern, resolvedPath)
    private_TryResolveDocumentDate = _
        (VBA.Len(VBA.Trim$(outDateText)) > 0)
    Exit Function
EH:
    private_Error "Не удалось извлечь дату из имени Word-документа '" & _
        resolvedPath & "': " & Err.Description
    Err.Clear
    On Error GoTo 0
End Function

Private Function private_MergeDocumentTables( _
    ByVal targetTables As Collection, _
    ByVal sourceTables As Collection, _
    ByVal documentDateText As String _
) As Boolean
    Dim sourceTableItem As Variant
    Dim sourceTable As obj_TableDynamic
    Dim targetTable As obj_TableDynamic
    Dim rowObj As obj_Row
    Dim clonedRow As obj_Row
    Dim rowIndex As Long

    If targetTables Is Nothing Or sourceTables Is Nothing Then Exit Function
    For Each sourceTableItem In sourceTables
        Set sourceTable = sourceTableItem
        If sourceTable Is Nothing Then Exit Function
        If VBA.Len(m_DocumentDateColumnCaption) > 0 Then
            If Not private_AddDocumentDateColumn( _
                sourceTable, documentDateText) Then Exit Function
        End If
        Set targetTable = private_FindTableByAlias( _
            targetTables, sourceTable.SourceAlias)
        If targetTable Is Nothing Then
            targetTables.Add sourceTable
        Else
            If VBA.StrComp(targetTable.HeaderText, sourceTable.HeaderText, _
                VBA.vbBinaryCompare) <> 0 Then
                private_Error "Нельзя объединить таблицу '" & _
                    sourceTable.SourceAlias & _
                    "': схемы колонок в документах отличаются."
                Exit Function
            End If
            For rowIndex = 1 To sourceTable.Rows.Count
                Set rowObj = sourceTable.Rows.Item(rowIndex)
                If rowObj Is Nothing Then Exit Function
                Set clonedRow = rowObj.Clone(targetTable.ColumnCount)
                If clonedRow Is Nothing Then Exit Function
                If Not targetTable.PushRow(clonedRow) Then Exit Function
            Next rowIndex
        End If
    Next sourceTableItem
    private_MergeDocumentTables = True
End Function

Private Function private_AddDocumentDateColumn( _
    ByVal tableObj As obj_TableDynamic, _
    ByVal documentDateText As String _
) As Boolean
    Dim dateColumn As obj_Column
    Dim rowObj As obj_Row
    Dim rowIndex As Long

    If tableObj Is Nothing Then Exit Function
    Set dateColumn = New obj_Column
    dateColumn.Name = m_DocumentDateColumnCaption
    dateColumn.FormatKind = "date"
    If Not dateColumn.AddAlias("documentDate") Then Exit Function
    If Not tableObj.InsertColumnAt(dateColumn, 1) Then Exit Function
    For rowIndex = 1 To tableObj.Rows.Count
        Set rowObj = tableObj.Rows.Item(rowIndex)
        If rowObj Is Nothing Then Exit Function
        If Not rowObj.SetCellRaw(1, documentDateText) Then Exit Function
    Next rowIndex
    private_AddDocumentDateColumn = True
End Function

Private Function private_FindTableByAlias( _
    ByVal tables As Collection, _
    ByVal sourceAlias As String _
) As obj_TableDynamic
    Dim tableItem As Variant
    Dim tableObj As obj_TableDynamic

    If tables Is Nothing Then Exit Function
    For Each tableItem In tables
        Set tableObj = tableItem
        If Not tableObj Is Nothing Then
            If VBA.StrComp(tableObj.SourceAlias, sourceAlias, _
                VBA.vbTextCompare) = 0 Then
                Set private_FindTableByAlias = tableObj
                Exit Function
            End If
        End If
    Next tableItem
End Function

Private Function private_TryTransformTables( _
    ByVal sourceTables As Collection, _
    ByRef outTables As Collection _
) As Boolean
    Dim transformer As obj_ITableTransformer
    Dim sourceTableItem As Variant
    Dim sourceTable As obj_TableDynamic
    Dim transformedTable As obj_TableDynamic

    Set outTables = Nothing
    If sourceTables Is Nothing Then Exit Function

    ' Пустой class сохраняет исходный pipeline без копирования таблиц.
    If VBA.Len(VBA.Trim$(m_TransformerClassName)) = 0 Then
        Set outTables = sourceTables
        private_TryTransformTables = True
        Exit Function
    End If
    If m_ConfigTable Is Nothing Then
        private_Error "Конфигурация transformer недоступна."
        Exit Function
    End If
    If Not private_TryCreateTableTransformer( _
        m_TransformerClassName, transformer) Then Exit Function

    Set outTables = New Collection
    For Each sourceTableItem In sourceTables
        Set sourceTable = sourceTableItem
        If sourceTable Is Nothing Then
            transformer.Dispose
            Exit Function
        End If
        Set transformedTable = Nothing
        If Not transformer.Transform( _
            sourceTable, m_ConfigTable, transformedTable) Then
            transformer.Dispose
            Exit Function
        End If
        If transformedTable Is Nothing Then
            transformer.Dispose
            private_Error "Transformer '" & m_TransformerClassName & _
                "' не вернул результирующую таблицу."
            Exit Function
        End If
        outTables.Add transformedTable
    Next sourceTableItem
    transformer.Dispose

    private_TryTransformTables = True
End Function

Private Function private_TryCreateTableTransformer( _
    ByVal transformerClassName As String, _
    ByRef outTransformer As obj_ITableTransformer _
) As Boolean
    Dim orderTransformer As obj_WDE_OrderTransformer

    Set outTransformer = Nothing
    transformerClassName = VBA.Trim$(transformerClassName)
    If VBA.Len(transformerClassName) = 0 Then Exit Function

    ' Тот же runtime-подход, что и для exporter-ов PEB: имя приходит из
    ' профиля, а конкретный VBA project class создаётся локальным Select Case.
    Select Case VBA.LCase$(transformerClassName)
        Case VBA.LCase$("obj_WDE_OrderTransformer")
            Set orderTransformer = New obj_WDE_OrderTransformer
            Set outTransformer = orderTransformer

        Case Else
            private_Error "Класс transformer не поддерживается: " & _
                transformerClassName
            Exit Function
    End Select

    private_TryCreateTableTransformer = Not outTransformer Is Nothing
End Function

Public Function Rerender(Optional ByVal arg As Variant) As Boolean
    Rerender = rt_PageManager.fn_RenderPage(m_Page, "worddataextractor:rerender")
End Function

Public Function OpenRulesFile(Optional ByVal arg As Variant) As Boolean
    Dim resolvedPath As String
    Dim shellRunner As Object
    Dim commandText As String
    Dim errorText As String

    If Not m_IsReady Then
        private_Error "Конфигурация не готова. Повторно откройте режим с главной страницы."
        Exit Function
    End If

    resolvedPath = private_ResolveWorkbookRelativePath(m_RulesPath)
    If VBA.Len(resolvedPath) = 0 Or _
        VBA.Len(VBA.Dir$(resolvedPath, VBA.vbNormal)) = 0 Then
        private_Error "Файл правил не найден: " & resolvedPath
        Exit Function
    End If

    On Error GoTo EH
    commandText = "notepad.exe """ & resolvedPath & """"
    Set shellRunner = VBA.CreateObject("WScript.Shell")
    shellRunner.Run commandText, VBA.vbNormalFocus, False
    Set shellRunner = Nothing
    OpenRulesFile = True
    Exit Function
EH:
    errorText = Err.Description
    Set shellRunner = Nothing
    private_Error "Не удалось открыть файл правил: " & errorText
End Function

Private Function private_ResolveWorkbookRelativePath(ByVal filePath As String) As String
    filePath = VBA.Trim$(filePath)
    If VBA.Len(filePath) = 0 Then Exit Function

    ' RulesFile допускает как абсолютный, так и относительный путь профиля.
    If VBA.InStr(1, filePath, ":", VBA.vbBinaryCompare) > 0 Or _
        VBA.Left$(filePath, 2) = "\\" Then
        private_ResolveWorkbookRelativePath = filePath
    Else
        private_ResolveWorkbookRelativePath = _
            ex_XmlCore.fn_CombineBasePath(ThisWorkbook, filePath)
    End If
End Function

Private Function private_ReadWordDocument(ByVal filePath As String, ByRef outText As String) As Boolean
    Dim fso As Object, wordApp As Object, wordDoc As Object
    Dim errDescription As String
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        private_Error "WORD-документ не найден: " & filePath
        Exit Function
    End If
    On Error GoTo EH
    If Not rt_PEB_WordExportRuntime.fn_GetOrCreateWordApp(wordApp) Then Exit Function
    Set wordDoc = wordApp.Documents.Open(filePath, False, True, False)
    ' Rules определяют секции по текстовым заголовкам и не зависят от
    ' отображаемой нумерации Word, поэтому документ читается одним COM-вызовом.
    If Not private_TryBuildDocumentText(wordDoc, outText) Then
        VBA.Err.Raise VBA.vbObjectError + 2101, "WordDataExtractor", _
            "Не удалось собрать текст из абзацев WORD-документа."
    End If
    wordDoc.Close False
    Set wordDoc = Nothing
    private_ReadWordDocument = True
    Exit Function
EH:
    errDescription = Err.Description
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close False
    On Error GoTo 0
    private_Error "Не удалось прочитать WORD-документ: " & errDescription
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

    ' Пробелы Word нормализуются централизованно до выполнения любого scope,
    ' context или field regex. Поэтому начальные фразы секций не обязаны
    ' отдельно перечислять NBSP и прочие визуально неотличимые разделители.
    resultText = private_NormalizeWordWhitespace(resultText)
    resultText = VBA.Replace(resultText, VBA.Chr$(7), VBA.vbCr)

    outText = resultText
    private_TryBuildDocumentText = (VBA.Len(outText) > 0)
    Exit Function
EH:
    ex_Core.fn_Diagnostic_LogError "WordDataExtractor: ошибка сборки текста документа: " & Err.Description
End Function

Private Function private_NormalizeWordWhitespace( _
    ByVal sourceText As String _
) As String
    Dim codePoint As Long

    sourceText = VBA.Replace(sourceText, VBA.ChrW$(160), " ")
    sourceText = VBA.Replace(sourceText, VBA.ChrW$(5760), " ")
    ' U+2000..U+200A: en/em/figure/thin/hair spaces и их варианты.
    For codePoint = 8192 To 8202
        sourceText = VBA.Replace(sourceText, VBA.ChrW$(codePoint), " ")
    Next codePoint
    sourceText = VBA.Replace(sourceText, VBA.ChrW$(8239), " ")
    sourceText = VBA.Replace(sourceText, VBA.ChrW$(8287), " ")
    sourceText = VBA.Replace(sourceText, VBA.ChrW$(12288), " ")
    ' Невидимые разделители не должны склеивать или разрывать ключевые фразы.
    sourceText = VBA.Replace(sourceText, VBA.ChrW$(8203), VBA.vbNullString)
    sourceText = VBA.Replace(sourceText, VBA.ChrW$(8288), VBA.vbNullString)
    sourceText = VBA.Replace(sourceText, VBA.ChrW$(65279), VBA.vbNullString)
    private_NormalizeWordWhitespace = sourceText
End Function

Private Sub private_Error(ByVal messageText As String)
    ex_Core.fn_Diagnostic_LogError "WordDataExtractor: " & messageText
    VBA.MsgBox messageText, VBA.vbExclamation, "PrototypeNew / WordDataExtractor"
End Sub
