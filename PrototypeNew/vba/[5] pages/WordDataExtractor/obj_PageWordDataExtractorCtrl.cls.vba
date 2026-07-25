VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PageWordDataExtractorCtrl"
Option Explicit

Private Const OBJECT_KEY As String = "RuntimeObjects.WordDataExtractor.Controller"
Private Const TABLES_KEY As String = "RuntimeItems.WordDataExtractor.Tables"
Private m_Page As obj_IPage
Private m_RulesPath As String
Private m_PipelineId As String
Private m_DocumentPath As String
Private m_TransformerClassName As String
Private m_ConfigTable As obj_ConfigTable
Private m_AllTables As Collection
Private m_ShowEmptyTables As Boolean
Private m_DataReadOptions As obj_ExternalDataReadOptions
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
    Set m_DataReadOptions = New obj_ExternalDataReadOptions
    If Not pageBase.RuntimeSources.SetItemsSource(TABLES_KEY, items, False) Then Exit Function
    Initialize = True
End Function

Public Sub Dispose()
    Set m_ConfigTable = Nothing
    Set m_AllTables = Nothing
    Set m_DataReadOptions = Nothing
    Set m_Page = Nothing
End Sub

Public Property Get ShowEmptyTables() As Boolean
    ShowEmptyTables = m_ShowEmptyTables
End Property

Public Property Get AdoSupportLongValues() As Boolean
    If m_DataReadOptions Is Nothing Then Exit Property
    AdoSupportLongValues = m_DataReadOptions.AdoSupportLongValues
End Property

Public Function UpdateData(ByVal configControl As obj_ConfigControlVM) As Boolean
    Dim configTable As obj_ConfigTable
    Dim wordDataExtractorCfgParser As obj_WordDataExtractorCfgParser
    Dim documentDir As String
    Dim documentFilename As String
    m_IsReady = False
    If configControl Is Nothing Then Exit Function
    If Not configControl.TryBuildConfigTableFromRendered(configTable) Then Exit Function
    Set wordDataExtractorCfgParser = New obj_WordDataExtractorCfgParser
    If Not wordDataExtractorCfgParser.Initialize(configTable) Then Exit Function
    If Not wordDataExtractorCfgParser.TryGetRequiredValue( _
        "WordDataExtractor.RulesFile", m_RulesPath) Then
        private_Error "В профиле отсутствует обязательный ключ WordDataExtractor.RulesFile."
        Exit Function
    End If
    If Not wordDataExtractorCfgParser.TryGetRequiredValue( _
        "WordDataExtractor.Pipeline", m_PipelineId) Then
        private_Error "В профиле отсутствует обязательный ключ WordDataExtractor.Pipeline."
        Exit Function
    End If
    m_DocumentPath = VBA.Trim$( _
        wordDataExtractorCfgParser.GetOptionalValue( _
            "WordDataExtractor.DocumentPath"))
    If VBA.Len(m_DocumentPath) = 0 Then
        ' Раздельная запись является альтернативой, а не дополнением:
        ' непустой DocumentPath всегда имеет приоритет над Dir/Filename.
        documentDir = VBA.Trim$( _
            wordDataExtractorCfgParser.GetOptionalValue( _
                "WordDataExtractor.DocumentDir"))
        documentFilename = VBA.Trim$( _
            wordDataExtractorCfgParser.GetOptionalValue( _
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
    m_TransformerClassName = _
        wordDataExtractorCfgParser.GetOptionalValue( _
            "WordDataExtractor.TransformerClass")
    Set m_ConfigTable = configTable
    m_IsReady = True
    UpdateData = True
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
    Dim engine As obj_WDE_RulesEngine, tables As Collection, documentText As String
    Dim resultTables As Collection
    Dim pageBase As obj_PageBase
    If Not m_IsReady Then
        private_Error "Конфигурация не готова. Повторно откройте режим с главной страницы."
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(m_DocumentPath)) = 0 Then
        private_Error "В текущем профиле не заполнен WordDataExtractor.DocumentPath " & _
            "или пара WordDataExtractor.DocumentDir/DocumentFilename."
        Exit Function
    End If
    If Not private_ReadWordDocument(m_DocumentPath, documentText) Then Exit Function
    Set engine = New obj_WDE_RulesEngine
    If Not engine.Initialize(m_RulesPath) Then Exit Function
    If Not engine.ExtractTables(m_PipelineId, documentText, tables) Then Exit Function
    If Not private_TryTransformTables(tables, resultTables) Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set m_AllTables = resultTables
    If Not private_PublishVisibleTables() Then Exit Function
    If Not rt_PageManager.fn_RenderPage(m_Page, "worddataextractor:extract") Then Exit Function
    rt_Messaging.fn_ShowStatusBarSuccess _
        "WordDataExtractor: извлечено таблиц: " & _
        VBA.CStr(resultTables.Count) & ". Режим: " & _
        private_CurrentTablesModeCaption() & ".", 4
    ExtractAndRender = True
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

Public Function ToggleAdoSupportLongValues(Optional ByVal arg As Variant) As Boolean
    If m_DataReadOptions Is Nothing Then
        private_Error "Настройки чтения внешних данных не инициализированы."
        Exit Function
    End If
    If Not m_DataReadOptions.ToggleLongValuesMode() Then Exit Function
    If Not ex_ControlRefreshRuntime.fn_TryRefreshStaticControl("AdoLongValues") Then
        If Not m_DataReadOptions.ToggleLongValuesMode() Then Exit Function
        private_Error "Не удалось обновить кнопку поддержки длинных ADO-значений."
        Exit Function
    End If

    ToggleAdoSupportLongValues = True
    If m_DataReadOptions.AdoSupportLongValues Then
        rt_Messaging.fn_ShowStatusBarSuccess "ADO support for long values enabled", 3
    Else
        rt_Messaging.fn_ShowStatusBarWarning "ADO support for long values disabled", 3
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
            Set orderTransformer.DataReadOptions = m_DataReadOptions
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
