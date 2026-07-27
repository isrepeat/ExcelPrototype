VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PageWordTextSearchCtrl"
Option Explicit

Implements obj_IPageCtrl

Private Const OBJECT_KEY As String = "RuntimeObjects.WordDataExtractor.Controller"
Private Const TABLES_KEY As String = "RuntimeItems.WordDataExtractor.Tables"
Private Const CONTEXT_LINE_RADIUS As Long = 5

Private m_Page As obj_IPage
Private m_DocumentPath As String
Private m_DocumentPathPattern As String
Private m_DocumentPaths As Collection
Private m_SearchText As String
Private m_IsRegexMode As Boolean
Private m_AllTables As Collection
Private m_ShowEmptyTables As Boolean
Private m_IsReady As Boolean

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
    Set items = New Collection
    Set m_AllTables = New Collection
    m_ShowEmptyTables = True
    If Not pageBase.RuntimeSources.SetItemsSource( _
        TABLES_KEY, items, False) Then Exit Function
    obj_IPageCtrl_Initialize = True
End Function

Private Function obj_IPageCtrl_UpdateData( _
    ByVal configControl As obj_ConfigControlVM _
) As Boolean
    Dim configTable As obj_ConfigTable
    Dim wordDataExtrCfgParser As obj_WordDataExtrCfgParser
    Dim documentDir As String
    Dim documentFilename As String
    Dim documentPathResolver As String
    Dim documentPathResolverArgs As String

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
        If VBA.Len(documentDir) = 0 Xor _
            VBA.Len(documentFilename) = 0 Then
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

    documentPathResolver = VBA.Trim$( _
        wordDataExtrCfgParser.GetOptionalValue( _
            "WordDataExtractor.DocumentPathResolver"))
    documentPathResolverArgs = VBA.Trim$( _
        wordDataExtrCfgParser.GetOptionalValue( _
            "WordDataExtractor.DocumentPathResolverArgs"))
    Set m_DocumentPaths = Nothing
    If VBA.Len(documentPathResolver) = 0 Then
        Set m_DocumentPaths = New Collection
        If VBA.Len(m_DocumentPath) > 0 Then
            m_DocumentPaths.Add m_DocumentPath
        End If
    ElseIf VBA.StrComp(documentPathResolver, _
        "ResolveAllByDmyPattern", VBA.vbTextCompare) = 0 Then
        On Error GoTo EH_RESOLVE_DOCUMENTS
        Set m_DocumentPaths = _
            ex_SourceResolver.fn_ResolveAllByDmyPattern( _
                m_DocumentPathPattern, documentPathResolverArgs)
        On Error GoTo 0
    Else
        private_Error "Поисковый controller не поддерживает " & _
            "WordDataExtractor.DocumentPathResolver: " & _
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

    m_IsReady = True
    obj_IPageCtrl_UpdateData = True
    Exit Function

EH_RESOLVE_DOCUMENTS:
    private_Error "Не удалось разрешить список Word-документов для поиска: " & _
        Err.Description
    Err.Clear
    On Error GoTo 0
End Function

Private Sub obj_IPageCtrl_Dispose()
    Set m_AllTables = Nothing
    Set m_DocumentPaths = Nothing
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

Public Function SearchAndRender(Optional ByVal arg As Variant) As Boolean
    Dim resultTables As Collection
    Dim documentPathItem As Variant
    Dim documentText As String
    Dim documentTable As obj_TableDynamic
    Dim searchRegex As Object
    Dim documentIndex As Long

    If Not m_IsReady Then
        private_Error "Конфигурация поиска не готова. " & _
            "Повторно откройте режим с главной страницы."
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(m_SearchText)) = 0 Then
        private_Error "Введите часть текста для поиска в WORD-документах."
        Exit Function
    End If
    If m_DocumentPaths Is Nothing Then
        private_Error "Список WORD-документов для поиска не подготовлен."
        Exit Function
    End If
    If m_DocumentPaths.Count = 0 Then
        private_Error "Список WORD-документов для поиска пуст."
        Exit Function
    End If
    If Not private_TryCreateSearchRegex(searchRegex) Then Exit Function

    Set resultTables = New Collection
    For Each documentPathItem In m_DocumentPaths
        documentIndex = documentIndex + 1
        If Not private_ReadWordDocument( _
            VBA.CStr(documentPathItem), documentText) Then Exit Function
        If Not private_BuildSearchResultTable( _
            VBA.CStr(documentPathItem), documentText, _
            documentIndex, searchRegex, _
            documentTable) Then Exit Function
        resultTables.Add documentTable
    Next documentPathItem

    Set m_AllTables = resultTables
    If Not private_PublishVisibleTables() Then Exit Function
    If Not rt_PageManager.fn_RenderPage( _
        m_Page, "word-text-search:search") Then Exit Function
    rt_Messaging.fn_ShowStatusBarSuccess _
        "WORD search: '" & m_SearchText & _
        "', обработано документов: " & _
        VBA.CStr(m_DocumentPaths.Count) & ", найдено записей: " & _
        VBA.CStr(private_CountTableRows(resultTables)) & ".", 4
    SearchAndRender = True
End Function

Public Function ToggleRegexMode(Optional ByVal arg As Variant) As Boolean
    m_IsRegexMode = Not m_IsRegexMode
    If Not rt_PageManager.fn_RenderPage( _
        m_Page, "word-text-search:toggle-regex") Then Exit Function
    ToggleRegexMode = True
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
    Set textLines = private_CollectNonEmptyLines(documentText)

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

Private Function private_CollectNonEmptyLines( _
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
        If VBA.Len(lineText) > 0 Then result.Add lineText
    Next rawLine
    Set private_CollectNonEmptyLines = result
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
    firstLineIndex = lineIndex - CONTEXT_LINE_RADIUS
    If firstLineIndex < 1 Then firstLineIndex = 1
    lastLineIndex = lineIndex + CONTEXT_LINE_RADIUS
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
    ByRef outText As String _
) As Boolean
    Dim fso As Object
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim errDescription As String

    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        private_Error "WORD-документ не найден: " & filePath
        Exit Function
    End If
    On Error GoTo EH
    If Not rt_PEB_WordExportRuntime.fn_GetOrCreateWordApp( _
        wordApp) Then Exit Function
    Set wordDoc = wordApp.Documents.Open(filePath, False, True, False)
    If Not private_TryBuildDocumentText(wordDoc, outText) Then
        VBA.Err.Raise VBA.vbObjectError + 2102, _
            "WordTextSearch", _
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
    private_Error "Не удалось прочитать WORD-документ: " & _
        errDescription
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
