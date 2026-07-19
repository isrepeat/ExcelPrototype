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
    If Not pageBase.RuntimeSources.SetItemsSource(TABLES_KEY, items, False) Then Exit Function
    Initialize = True
End Function

Public Sub Dispose()
    Set m_ConfigTable = Nothing
    Set m_Page = Nothing
End Sub

Public Function UpdateData(ByVal configControl As obj_ConfigControlVM) As Boolean
    Dim configTable As obj_ConfigTable, parser As obj_CfgParserBase
    Dim entries As Collection, cfgMap As Object
    m_IsReady = False
    If configControl Is Nothing Then Exit Function
    If Not configControl.TryBuildConfigTableFromRendered(configTable) Then Exit Function
    Set parser = New obj_CfgParserBase
    If Not parser.Initialize(configTable) Then Exit Function
    If Not parser.TryGetConfigEntries(entries) Then Exit Function
    If Not parser.BuildConfigDictionary(entries, cfgMap) Then Exit Function
    If Not parser.TryGetRequiredConfigValue(cfgMap, "WordDataExtractor.RulesFile", m_RulesPath) Then
        private_Error "В профиле отсутствует обязательный ключ WordDataExtractor.RulesFile."
        Exit Function
    End If
    If Not parser.TryGetRequiredConfigValue(cfgMap, "WordDataExtractor.Pipeline", m_PipelineId) Then
        private_Error "В профиле отсутствует обязательный ключ WordDataExtractor.Pipeline."
        Exit Function
    End If
    m_DocumentPath = parser.GetOptionalConfigValue(cfgMap, "WordDataExtractor.DocumentPath")
    m_TransformerClassName = parser.GetOptionalConfigValue( _
        cfgMap, "WordDataExtractor.TransformerClass")
    Set m_ConfigTable = configTable
    m_IsReady = True
    UpdateData = True
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
        private_Error "В текущем профиле не заполнен WordDataExtractor.DocumentPath."
        Exit Function
    End If
    If Not private_ReadWordDocument(m_DocumentPath, documentText) Then Exit Function
    Set engine = New obj_WDE_RulesEngine
    If Not engine.Initialize(m_RulesPath) Then Exit Function
    If Not engine.ExtractTables(m_PipelineId, documentText, tables) Then Exit Function
    If Not private_TryTransformTables(tables, resultTables) Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Not pageBase.RuntimeSources.SetItemsSource( _
        TABLES_KEY, resultTables, False) Then Exit Function
    If Not rt_PageManager.fn_RenderPage(m_Page, "worddataextractor:extract") Then Exit Function
    rt_Messaging.fn_ShowStatusBarSuccess _
        "WordDataExtractor: извлечено таблиц: " & _
        VBA.CStr(resultTables.Count) & ".", 4
    ExtractAndRender = True
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
    Set outTransformer = Nothing
    transformerClassName = VBA.Trim$(transformerClassName)
    If VBA.Len(transformerClassName) = 0 Then Exit Function

    ' Тот же runtime-подход, что и для exporter-ов PEB: имя приходит из
    ' профиля, а конкретный VBA project class создаётся локальным Select Case.
    Select Case VBA.LCase$(transformerClassName)
        Case VBA.LCase$("obj_WDE_OrderTransformer")
            Set outTransformer = New obj_WDE_OrderTransformer

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
    ' Document.Content.Text не содержит отображаемые метки автоматических
    ' списков Word. Собираем текст по абзацам и сохраняем ListString, чтобы
    ' извлекаемый текст соответствовал видимому документу. Текущие scope
    ' определяются заголовками и от нумерации секций не зависят.
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
    Dim paragraph As Object
    Dim paragraphText As String
    Dim listLabel As String
    Dim resultText As String

    outText = VBA.vbNullString
    If wordDoc Is Nothing Then Exit Function

    On Error GoTo EH
    For Each paragraph In wordDoc.Paragraphs
        paragraphText = VBA.CStr(paragraph.Range.Text)
        listLabel = VBA.vbNullString

        ' У обычного абзаца ListString пуст. Ошибку отдельного повреждённого
        ' списка не превращаем в ошибку чтения всего документа.
        On Error Resume Next
        listLabel = VBA.Trim$(VBA.CStr(paragraph.Range.ListFormat.ListString))
        Err.Clear
        On Error GoTo EH

        If VBA.Len(listLabel) > 0 Then
            resultText = resultText & listLabel
            If VBA.Len(paragraphText) > 0 Then
                If VBA.Left$(paragraphText, 1) <> " " And _
                    VBA.Left$(paragraphText, 1) <> VBA.vbTab Then
                    resultText = resultText & " "
                End If
            End If
        End If
        resultText = resultText & paragraphText
    Next paragraph

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
