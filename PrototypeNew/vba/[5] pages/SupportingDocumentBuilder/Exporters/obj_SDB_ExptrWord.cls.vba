VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_SDB_ExptrWord"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Implements obj_IDataExporter

' Runtime path is relative to ThisWorkbook.Path, same as page UI paths.
Private Const WORD_RESULT_TEMPLATES_REL_PATH As String = "modes\SupportingDocumentBuilder\SupportingDocumentBuilderWordResultTemplates.xml"
Private Const CONTEXT_SECTION_TYPE As String = "SectionType"
Private Const CONTEXT_WORD_PREVIEW_TEXT As String = "WordExportPreviewText"
Private Const CONTEXT_WORD_PREVIEW_EDITABLE As String = "WordExportPreviewEditable"
Private Const WORD_ANCHOR_PREFIX As String = "{\export:"
Private Const WORD_ANCHOR_BEGIN_SUFFIX As String = "_Begin}"
Private Const WORD_ANCHOR_END_SUFFIX As String = "_End}"
Private Const WD_FIND_STOP As Long = 0
Private Const WD_STYLE_TABLE_GRID As Long = -155
Private Const WD_ALIGN_PARAGRAPH_LEFT As Long = 0
Private Const WD_ALIGN_PARAGRAPH_CENTER As Long = 1
Private Const WD_BORDER_TOP As Long = -1
Private Const WD_BORDER_LEFT As Long = -2
Private Const WD_BORDER_BOTTOM As Long = -3
Private Const WD_BORDER_RIGHT As Long = -4
Private Const WD_LINE_STYLE_NONE As Long = 0
Private Const WD_LINE_STYLE_SINGLE As Long = 1
Private Const WD_AUTO_FIT_FIXED As Long = 0
Private Const WD_PREFERRED_WIDTH_PERCENT As Long = 2
Private Const WD_STYLE_NORMAL As Long = -1
Private Const WD_ORIENT_PORTRAIT As Long = 0
Private Const WD_ORIENT_LANDSCAPE As Long = 1
Private Const WD_PAPER_A4 As Long = 7
Private Const WD_LINE_SPACE_SINGLE As Long = 0
Private Const WD_FORMAT_DOCUMENT_DEFAULT As Long = 16


Private m_IsDisposed As Boolean
Private m_Base As obj_DataExporterBase
Private m_TemplateParser As obj_WordResultTplParser
Private m_WordResultStylesParser As obj_WordResultStylesParser

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
Public Function Initialize(ByVal configTable As obj_ConfigTable) As Boolean
    m_IsDisposed = False
    Set m_Base = New obj_DataExporterBase
    Set m_TemplateParser = New obj_WordResultTplParser
    Set m_WordResultStylesParser = New obj_WordResultStylesParser
    If Not m_Base.Initialize(configTable, "WORD", "PrototypeNew / WORD export") Then Exit Function
    If Not m_TemplateParser.Initialize(WORD_RESULT_TEMPLATES_REL_PATH) Then Exit Function
    Initialize = True
End Function

Public Sub Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    If Not m_Base Is Nothing Then m_Base.Dispose
    If Not m_TemplateParser Is Nothing Then m_TemplateParser.Dispose
    Set m_Base = Nothing
    Set m_TemplateParser = Nothing
    Set m_WordResultStylesParser = Nothing
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
    Dim usePreparedPreview As Boolean
    Dim writeToWord As Boolean
    Dim generationMode As String
    Dim documentStylesMarkup As String
    Dim sectionStylesMarkup As String
    Dim documentFilepath As String

    On Error GoTo EH

    If m_IsDisposed Then
        VBA.MsgBox "PrototypeNew: WORD exporter is disposed.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    If Not m_Base.TryGetMainSourceTable(sourceTables, sourceTable) Then Exit Function
    If Not m_TemplateParser.TryGetDocumentSettings( _
        generationMode, documentStylesMarkup, sectionStylesMarkup) Then Exit Function

    sectionTypeText = private_GetContextText(context, CONTEXT_SECTION_TYPE)
    If VBA.Len(sectionTypeText) = 0 Then sectionTypeText = VBA.Trim$(sourceTable.SectionTitle)
    If VBA.Len(sectionTypeText) = 0 Then
        VBA.MsgBox "PrototypeNew: WORD export requires SectionType in export context or source table SectionTitle.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    If Not private_ValidateCertificateFields(sourceTable) Then Exit Function
    Set namedCollections = ex_Helpers.fn_CreateDictionaryTextCompare()
    If Not m_TemplateParser.TryRenderDocumentFilepath( _
        sectionTypeText, sourceTables, namedCollections, context, _
        documentFilepath) Then Exit Function
    ' В SDB идентификатор шаблона совпадает с секцией. Каталог секций остаётся
    ' в obj_SDB_Data, а экспортёр не содержит перечня документов.
    templateId = sectionTypeText
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
        If Not private_TrySetContextBoolean( _
            context, _
            CONTEXT_WORD_PREVIEW_EDITABLE, _
            Not m_TemplateParser.LastRenderHasTables) Then Exit Function
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo _
            "sdb-word-export:record-rendered length=" & _
            VBA.CStr(VBA.Len(recordText)) & " editable=" & _
            VBA.LCase$(VBA.CStr(Not m_TemplateParser.LastRenderHasTables))
#End If
    End If
    If usePreparedPreview Then
        recordText = previewText
    Else
        previewText = recordText
    End If
    If Not private_TrySetContextText( _
        context, CONTEXT_WORD_PREVIEW_TEXT, previewText) Then Exit Function

    If writeToWord Then
        Select Case generationMode
            Case "anchored"
                If Not private_TryAppendBeforeWordEndAnchor( _
                    templateId, recordText, documentFilepath) Then Exit Function
            Case "newdocument"
                If Not private_TryGenerateNewWordDocument( _
                    recordText, documentStylesMarkup, sectionStylesMarkup, _
                    documentFilepath) Then Exit Function
            Case Else
                VBA.MsgBox "SupportingDocumentBuilder: unsupported WORD " & _
                    "generation mode '" & generationMode & "'.", _
                    VBA.vbExclamation, "Supporting Document Builder"
                Exit Function
        End Select
    End If

    Export = True
    Exit Function

EH:
    ex_Core.fn_Diagnostic_LogError _
        "sdb-word-export:error errNo=" & VBA.CStr(Err.Number) & _
        " err='" & VBA.Replace(Err.Description, "'", "''") & "'"
    VBA.MsgBox "SupportingDocumentBuilder WORD export failed: [" & _
        VBA.CStr(Err.Number) & "] " & Err.Description, _
        VBA.vbExclamation, "Supporting Document Builder"
End Function

Private Function private_TrySetContextBoolean( _
    ByVal context As Object, _
    ByVal keyText As String, _
    ByVal valueToStore As Boolean _
) As Boolean
    If context Is Nothing Then
        VBA.MsgBox "SupportingDocumentBuilder export context is missing.", _
            VBA.vbExclamation, "Supporting Document Builder"
        Exit Function
    End If
    keyText = VBA.Trim$(keyText)
    If VBA.Len(keyText) = 0 Then Exit Function

    On Error Resume Next
    context(keyText) = valueToStore
    If Err.Number <> 0 Then
        Err.Clear
        VBA.CallByName context, keyText, VbLet, valueToStore
    End If
    If Err.Number <> 0 Then
        VBA.MsgBox "SupportingDocumentBuilder could not write export context key '" & _
            keyText & "': [" & VBA.CStr(Err.Number) & "] " & _
            Err.Description, VBA.vbExclamation, _
            "Supporting Document Builder"
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    private_TrySetContextBoolean = True
End Function

Private Function private_ValidateCertificateFields( _
    ByVal sourceTable As obj_TableDynamic _
) As Boolean
    Dim requiredAliases As Variant
    Dim requiredCaptions As Variant
    Dim fieldIndex As Long
    Dim fieldValue As String
    Dim positionCodeText As String

    requiredAliases = Array( _
        "Rank", "FIO", "PositionCode", "DateFrom", "DocDate", "DocNo", _
        "ReportPositionCode", "Destination", "IncomingNo", "IncomingDate")
    requiredCaptions = Array( _
        "Звання", "ПІБ", "Код посади", "Дата початку відрядження", _
        "Дата наказу", "Номер наказу", "Посада автора повідомлення", _
        "Військова частина", "Вихідний номер", "Вихідна дата")

    For fieldIndex = LBound(requiredAliases) To UBound(requiredAliases)
        fieldValue = VBA.vbNullString
        If Not private_TryGetMainTableValue( _
            sourceTable, VBA.CStr(requiredAliases(fieldIndex)), fieldValue) Then
            fieldValue = VBA.vbNullString
        End If
        If VBA.Len(VBA.Trim$(fieldValue)) = 0 Then
            VBA.MsgBox "SupportingDocumentBuilder: заповніть обов'язкове поле '" & _
                VBA.CStr(requiredCaptions(fieldIndex)) & "'.", _
                VBA.vbExclamation, "Supporting Document Builder"
            Exit Function
        End If
    Next fieldIndex

    If Not private_TryGetMainTableValue( _
        sourceTable, "PositionCode", positionCodeText) Then Exit Function
    If VBA.InStr(1, positionCodeText, "РОЗП", VBA.vbTextCompare) = 0 Then
        fieldValue = VBA.vbNullString
        If Not private_TryGetMainTableValue( _
            sourceTable, "PositionName", fieldValue) Then fieldValue = VBA.vbNullString
        If VBA.Len(VBA.Trim$(fieldValue)) = 0 Then
            VBA.MsgBox "SupportingDocumentBuilder: для коду посади, який не " & _
                "містить РОЗП, заповніть поле 'Посада'.", _
                VBA.vbExclamation, "Supporting Document Builder"
            Exit Function
        End If
    End If
    private_ValidateCertificateFields = True
End Function

Private Function private_TryGenerateNewWordDocument( _
    ByVal renderedText As String, _
    ByVal documentStylesMarkup As String, _
    ByVal sectionStylesMarkup As String, _
    ByVal documentFilepath As String _
) As Boolean
    Dim outputPath As String
    Dim plainRenderedText As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim insertRange As Object
    Dim insertedEnd As Long
    Dim errorDescription As String
    Dim errorNumber As Long
    Dim ownsWordApp As Boolean
    Dim previousScreenUpdating As Boolean
    Dim previousDisplayAlerts As Long
    Dim wordAppStateCaptured As Boolean

    On Error GoTo EH
    outputPath = VBA.Trim$(documentFilepath)
    If VBA.Len(outputPath) = 0 Then
        VBA.MsgBox "SupportingDocumentBuilder: rendered documentFilepath is " & _
            "empty.", VBA.vbExclamation, "Supporting Document Builder"
        Exit Function
    End If
    If Not private_IsAbsolutePath(outputPath) Then _
        outputPath = ThisWorkbook.Path & Application.PathSeparator & outputPath

    If Not private_TryStripPreviewInlineMarkers( _
        renderedText, plainRenderedText) Then Exit Function
    plainRenderedText = private_NormalizeWordParagraphBreaks(plainRenderedText)
    If Not private_TryNormalizeNewDocumentPages(plainRenderedText) Then Exit Function

    ' Холодный CreateObject Word занимает основную часть времени экспорта.
    ' Повторно используем уже запущенный экземпляр, в том числе оставшийся
    ' открытым после предыдущего сгенерированного документа.
    On Error Resume Next
    Set wordApp = VBA.GetObject(, "Word.Application")
    On Error GoTo EH
    If wordApp Is Nothing Then
        Set wordApp = VBA.CreateObject("Word.Application")
        ownsWordApp = True
    End If
    previousScreenUpdating = VBA.CBool(wordApp.ScreenUpdating)
    previousDisplayAlerts = VBA.CLng(wordApp.DisplayAlerts)
    wordAppStateCaptured = True
    wordApp.ScreenUpdating = False
    wordApp.DisplayAlerts = 0
    Set wordDoc = wordApp.Documents.Add
    If Not m_WordResultStylesParser.Initialize( _
        documentStylesMarkup, sectionStylesMarkup) Then GoTo CleanFail
    If Not private_TryApplyNewDocumentStyles( _
        wordDoc, m_WordResultStylesParser) Then GoTo CleanFail

    Set insertRange = wordDoc.Range(0, 0)
    If Not private_TryInsertRenderedHtmlBlocks( _
        wordDoc, insertRange, plainRenderedText, insertedEnd) Then GoTo CleanFail

    ' newDocument всегда является полной повторной генерацией результата.
    ' SaveAs2 заменяет только путь, вычисленный documentFilepath; входной DOCX
    ' из профиля в этом режиме не открывается и не изменяется.
    wordDoc.SaveAs2 outputPath, WD_FORMAT_DOCUMENT_DEFAULT
    wordApp.DisplayAlerts = previousDisplayAlerts
    wordApp.ScreenUpdating = previousScreenUpdating
    wordApp.Visible = True
    wordDoc.Activate
    ' Это настройка отображения окна, но её значение входит в глобальную
    ' конфигурацию documentStyles вместе с остальным видом результата.
    wordApp.ActiveWindow.View.TableGridlines = _
        m_WordResultStylesParser.ShowTableGridlines
    private_TryGenerateNewWordDocument = True
    Exit Function

CleanFail:
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close False
    If Not wordApp Is Nothing And wordAppStateCaptured Then
        wordApp.DisplayAlerts = previousDisplayAlerts
        wordApp.ScreenUpdating = previousScreenUpdating
    End If
    If Not wordApp Is Nothing And ownsWordApp Then wordApp.Quit
    On Error GoTo 0
    Exit Function

EH:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close False
    If Not wordApp Is Nothing And wordAppStateCaptured Then
        wordApp.DisplayAlerts = previousDisplayAlerts
        wordApp.ScreenUpdating = previousScreenUpdating
    End If
    If Not wordApp Is Nothing And ownsWordApp Then wordApp.Quit
    On Error GoTo 0
    VBA.MsgBox "SupportingDocumentBuilder: failed to generate a new WORD " & _
        "document: [" & VBA.CStr(errorNumber) & "] " & errorDescription, _
        VBA.vbExclamation, "Supporting Document Builder"
End Function

Private Function private_TryNormalizeNewDocumentPages( _
    ByRef renderedText As String _
) As Boolean
    Dim pageStart As Long
    Dim pageEnd As Long
    Dim pageIndex As Long

    pageStart = VBA.InStr(1, renderedText, "<page>", VBA.vbTextCompare)
    If pageStart = 0 Then
        VBA.MsgBox "SupportingDocumentBuilder: newDocument template requires " & _
            "at least one <page> block.", VBA.vbExclamation, _
            "Supporting Document Builder"
        Exit Function
    End If
    Do While pageStart > 0
        pageEnd = VBA.InStr(pageStart, renderedText, "</page>", _
            VBA.vbTextCompare)
        If pageEnd = 0 Then
            VBA.MsgBox "SupportingDocumentBuilder: unclosed <page> block.", _
                VBA.vbExclamation, "Supporting Document Builder"
            Exit Function
        End If
        pageIndex = pageIndex + 1
        renderedText = VBA.Left$(renderedText, pageStart - 1) & _
            IIf(pageIndex = 1, VBA.vbNullString, VBA.Chr$(12)) & _
            VBA.Mid$(renderedText, pageStart + VBA.Len("<page>"))
        pageEnd = VBA.InStr(pageStart, renderedText, "</page>", _
            VBA.vbTextCompare)
        renderedText = VBA.Left$(renderedText, pageEnd - 1) & _
            VBA.Mid$(renderedText, pageEnd + VBA.Len("</page>"))
        pageStart = VBA.InStr(pageStart, renderedText, "<page>", _
            VBA.vbTextCompare)
    Loop
    private_TryNormalizeNewDocumentPages = True
End Function

Private Function private_TryApplyNewDocumentStyles( _
    ByVal wordDoc As Object, _
    ByVal wordResultStylesParser As obj_WordResultStylesParser _
) As Boolean
    Dim normalStyle As Object
    Dim pageSetup As Object

    If wordDoc Is Nothing Or wordResultStylesParser Is Nothing Then Exit Function
    On Error GoTo EH
    Set normalStyle = wordDoc.Styles(WD_STYLE_NORMAL)
    normalStyle.Font.Name = wordResultStylesParser.FontName
    normalStyle.Font.Size = wordResultStylesParser.FontSize
    normalStyle.ParagraphFormat.SpaceBefore = _
        wordResultStylesParser.ParagraphSpacingBefore
    normalStyle.ParagraphFormat.SpaceAfter = _
        wordResultStylesParser.ParagraphSpacingAfter
    normalStyle.ParagraphFormat.LineSpacingRule = WD_LINE_SPACE_SINGLE
    ' Пустой стартовый абзац получает обновлённый Normal до вставки контента;
    ' поэтому обычный текст и новые таблицы наследуют глобальные параметры, а
    ' локальные атрибуты ячеек применяются поверх них позднее.
    wordDoc.Content.Style = normalStyle

    Set pageSetup = wordDoc.Sections(1).PageSetup
    pageSetup.PaperSize = WD_PAPER_A4
    Select Case wordResultStylesParser.Orientation
        Case "portrait": pageSetup.Orientation = WD_ORIENT_PORTRAIT
        Case "landscape": pageSetup.Orientation = WD_ORIENT_LANDSCAPE
    End Select
    pageSetup.TopMargin = wordResultStylesParser.MarginTopPoints
    pageSetup.RightMargin = wordResultStylesParser.MarginRightPoints
    pageSetup.BottomMargin = wordResultStylesParser.MarginBottomPoints
    pageSetup.LeftMargin = wordResultStylesParser.MarginLeftPoints
    private_TryApplyNewDocumentStyles = True
    Exit Function
EH:
    VBA.MsgBox "SupportingDocumentBuilder: failed to apply global WORD styles: [" & _
        VBA.CStr(Err.Number) & "] " & Err.Description, VBA.vbExclamation, _
        "Supporting Document Builder"
End Function

Private Function private_TryAppendBeforeWordEndAnchor( _
    ByVal templateId As String, _
    ByVal renderedText As String, _
    ByVal documentFilepath As String _
) As Boolean
    Dim templatePath As String
    Dim targetPath As String
    Dim beginMarker As String
    Dim endMarker As String
    Dim plainRenderedText As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim beginRange As Object
    Dim searchRange As Object
    Dim endRange As Object
    Dim insertRange As Object
    Dim insertedEnd As Long
    Dim ownsWordApp As Boolean

    On Error GoTo EH
    templatePath = VBA.Trim$(m_Base.TargetWorkbookPath)
    If VBA.Len(templatePath) = 0 Then
        VBA.MsgBox "SupportingDocumentBuilder: Export.Word.FilePath is empty.", _
            VBA.vbExclamation, "Supporting Document Builder"
        Exit Function
    End If
    If Not private_IsAbsolutePath(templatePath) Then _
        templatePath = ThisWorkbook.Path & Application.PathSeparator & templatePath
    If VBA.Len(VBA.Dir$(templatePath, VBA.vbNormal Or VBA.vbReadOnly Or _
        VBA.vbHidden Or VBA.vbSystem)) = 0 Then
        VBA.MsgBox "SupportingDocumentBuilder: WORD template was not found: " & _
            templatePath, VBA.vbExclamation, "Supporting Document Builder"
        Exit Function
    End If

    targetPath = VBA.Trim$(documentFilepath)
    If VBA.Len(targetPath) = 0 Then
        VBA.MsgBox "SupportingDocumentBuilder: rendered documentFilepath is empty.", _
            VBA.vbExclamation, "Supporting Document Builder"
        Exit Function
    End If
    If Not private_IsAbsolutePath(targetPath) Then _
        targetPath = ThisWorkbook.Path & Application.PathSeparator & targetPath
    If VBA.Len(VBA.Dir$(targetPath, VBA.vbNormal Or VBA.vbReadOnly Or _
        VBA.vbHidden Or VBA.vbSystem)) = 0 Then VBA.FileCopy templatePath, targetPath

    On Error Resume Next
    Set wordApp = VBA.GetObject(, "Word.Application")
    On Error GoTo EH
    If wordApp Is Nothing Then
        Set wordApp = VBA.CreateObject("Word.Application")
        ownsWordApp = True
    End If
    Set wordDoc = wordApp.Documents.Open(targetPath)
    beginMarker = WORD_ANCHOR_PREFIX & VBA.Trim$(templateId) & _
        WORD_ANCHOR_BEGIN_SUFFIX
    endMarker = WORD_ANCHOR_PREFIX & VBA.Trim$(templateId) & _
        WORD_ANCHOR_END_SUFFIX
    If Not private_TryFindWordText(wordDoc.Content, beginMarker, beginRange) Then
        VBA.MsgBox "SupportingDocumentBuilder: WORD begin anchor was not found: " & _
            beginMarker, VBA.vbExclamation, "Supporting Document Builder"
        GoTo CleanFail
    End If
    Set searchRange = wordDoc.Range(beginRange.End, wordDoc.Content.End)
    If Not private_TryFindWordText(searchRange, endMarker, endRange) Then
        VBA.MsgBox "SupportingDocumentBuilder: WORD end anchor was not found: " & _
            endMarker, VBA.vbExclamation, "Supporting Document Builder"
        GoTo CleanFail
    End If

    If Not private_TryStripPreviewInlineMarkers( _
        renderedText, plainRenderedText) Then GoTo CleanFail
    plainRenderedText = private_NormalizeWordParagraphBreaks(plainRenderedText)
    Set insertRange = wordDoc.Range(endRange.Start, endRange.Start)
    If VBA.InStr(1, plainRenderedText, "<table", VBA.vbTextCompare) > 0 Or _
        VBA.InStr(1, plainRenderedText, "<image", VBA.vbTextCompare) > 0 Then
        If Not private_TryInsertRenderedHtmlBlocks( _
            wordDoc, insertRange, plainRenderedText, insertedEnd) Then GoTo CleanFail
    Else
        insertRange.Text = plainRenderedText
    End If
    wordDoc.Save
    wordDoc.Close False
    If ownsWordApp Then wordApp.Quit
    private_TryAppendBeforeWordEndAnchor = True
    Exit Function

CleanFail:
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close False
    If ownsWordApp And Not wordApp Is Nothing Then wordApp.Quit
    On Error GoTo 0
    Exit Function
EH:
    VBA.MsgBox "SupportingDocumentBuilder: anchored WORD export failed: [" & _
        VBA.CStr(Err.Number) & "] " & Err.Description, VBA.vbExclamation, _
        "Supporting Document Builder"
    Resume CleanFail
End Function

Private Function private_TryInsertRenderedHtmlBlocks( _
    ByVal wordDoc As Object, _
    ByVal targetRange As Object, _
    ByVal renderedText As String, _
    ByRef outInsertedEnd As Long _
) As Boolean
    Dim cursorPos As Long
    Dim searchPos As Long
    Dim tableStart As Long
    Dim tableEnd As Long
    Dim imageStart As Long
    Dim imageEnd As Long
    Dim beforeText As String
    Dim tableMarkup As String
    Dim htmlDoc As Object
    Dim htmlTable As Object
    Dim htmlRow As Object
    Dim htmlCell As Object
    Dim htmlImages As Object
    Dim htmlImage As Object
    Dim wordRange As Object
    Dim wordTable As Object
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim rowCount As Long
    Dim columnCount As Long
    Dim tableStyle As String
    Dim declaredColumnCount As Long
    Dim rowColumnCount As Long
    Dim cellCount As Long
    Dim cellIndex As Long
    Dim cellSpan As Long
    Dim startColumns() As Long
    Dim cellSpans() As Long
    Dim cellTexts() As String
    Dim cellIsHeader() As Boolean
    Dim cellBorders() As String
    Dim cellAlignments() As String
    Dim cellBoldValues() As String
    Dim cellFontSizes() As String
    Dim cellImageSources() As String
    Dim cellImageWidths() As String
    Dim cellImageHeights() As String
    Dim cellFontColors() As String
    Dim cellBackgrounds() As String
    Dim tableBorders As String
    Dim tableBordersDisabled As Boolean
    Dim columnWidthsText As String
    Dim previousBlockWasTable As Boolean

    outInsertedEnd = 0
    If wordDoc Is Nothing Or targetRange Is Nothing Then Exit Function
    cursorPos = targetRange.Start
    searchPos = 1

    Do
        tableStart = VBA.InStr(searchPos, renderedText, "<table", VBA.vbTextCompare)
        imageStart = VBA.InStr(searchPos, renderedText, "<image", VBA.vbTextCompare)
        If tableStart = 0 And imageStart = 0 Then Exit Do
        If imageStart > 0 And (tableStart = 0 Or imageStart < tableStart) Then
            imageEnd = VBA.InStr(imageStart, renderedText, "</image>", VBA.vbTextCompare)
            If imageEnd = 0 Then
                VBA.MsgBox "SupportingDocumentBuilder: preview contains an " & _
                    "unclosed <image> element.", VBA.vbExclamation, _
                    "Supporting Document Builder"
                Exit Function
            End If
            imageEnd = imageEnd + VBA.Len("</image>") - 1
            beforeText = VBA.Mid$(renderedText, searchPos, imageStart - searchPos)
            If VBA.Len(beforeText) > 0 Then
                Set wordRange = wordDoc.Range(cursorPos, cursorPos)
                wordRange.Text = beforeText
                cursorPos = wordRange.End
            End If
            Set wordRange = wordDoc.Range(cursorPos, cursorPos)
            If Not private_TryInsertRenderedImage( _
                wordDoc, wordRange, _
                VBA.Mid$(renderedText, imageStart, imageEnd - imageStart + 1), _
                cursorPos) Then Exit Function
            previousBlockWasTable = False
            searchPos = imageEnd + 1
            GoTo ContinueBlock
        End If
        tableEnd = VBA.InStr(tableStart, renderedText, "</table>", VBA.vbTextCompare)
        If tableEnd = 0 Then
            VBA.MsgBox "SupportingDocumentBuilder: preview contains an unclosed " & _
                "<table> element.", VBA.vbExclamation, _
                "Supporting Document Builder"
            Exit Function
        End If
        tableEnd = tableEnd + VBA.Len("</table>") - 1
        beforeText = VBA.Mid$(renderedText, searchPos, tableStart - searchPos)
        If previousBlockWasTable And _
            VBA.Len(VBA.Trim$(beforeText)) = 0 Then
            ' XML-отступы между элементами не создают абзац в Word, поэтому
            ' считаются тем же случаем, что и полностью соседние таблицы.
            beforeText = VBA.vbNullString
            Set wordRange = wordDoc.Range(cursorPos, cursorPos)
            wordRange.Text = VBA.vbCr
            cursorPos = wordRange.End
            previousBlockWasTable = False
        End If
        If VBA.Len(beforeText) > 0 Then
            Set wordRange = wordDoc.Range(cursorPos, cursorPos)
            wordRange.Text = beforeText
            cursorPos = wordRange.End
            previousBlockWasTable = False
        End If

        tableMarkup = VBA.Mid$(renderedText, tableStart, tableEnd - tableStart + 1)
        Set htmlDoc = VBA.CreateObject("htmlfile")
        htmlDoc.Open
        htmlDoc.Write "<html><body>" & tableMarkup & "</body></html>"
        htmlDoc.Close
        If htmlDoc.getElementsByTagName("table").Length <> 1 Then
            VBA.MsgBox "SupportingDocumentBuilder: failed to parse rendered table.", _
                VBA.vbExclamation, "Supporting Document Builder"
            Exit Function
        End If
        Set htmlTable = htmlDoc.getElementsByTagName("table").Item(0)
        rowCount = htmlTable.Rows.Length
        columnCount = 0
        For Each htmlRow In htmlTable.Rows
            rowColumnCount = 0
            For Each htmlCell In htmlRow.Cells
                cellSpan = VBA.CLng(htmlCell.colSpan)
                If cellSpan <= 0 Then cellSpan = 1
                rowColumnCount = rowColumnCount + cellSpan
            Next htmlCell
            If rowColumnCount > columnCount Then columnCount = rowColumnCount
        Next htmlRow
        If Not m_WordResultStylesParser.TryParseOptionalPositiveLong( _
            m_WordResultStylesParser.HtmlAttributeText(htmlTable, "columns"), _
            declaredColumnCount) Then Exit Function
        If declaredColumnCount > columnCount Then columnCount = declaredColumnCount
        If rowCount <= 0 Or columnCount <= 0 Then
            VBA.MsgBox "SupportingDocumentBuilder: rendered table must contain " & _
                "at least one row and one cell.", VBA.vbExclamation, _
                "Supporting Document Builder"
            Exit Function
        End If

        Set wordRange = wordDoc.Range(cursorPos, cursorPos)
        Set wordTable = wordDoc.Tables.Add(wordRange, rowCount, columnCount)
        columnWidthsText = VBA.Trim$( _
            m_WordResultStylesParser.HtmlAttributeText( _
                htmlTable, "columnWidths"))
        If VBA.Len(columnWidthsText) > 0 Then
            If Not private_TryApplyWordTableColumnWidths( _
                wordTable, columnWidthsText, columnCount) Then Exit Function
        End If
        tableBorders = m_WordResultStylesParser.HtmlAttributeText( _
            htmlTable, "borders")
        If Not m_WordResultStylesParser.TryParseTableBorders( _
            tableBorders, tableBordersDisabled) Then Exit Function
        If tableBordersDisabled Then wordTable.Borders.Enable = False
        rowIndex = 1
        For Each htmlRow In htmlTable.Rows
            cellCount = htmlRow.Cells.Length
            If cellCount <= 0 Then
                VBA.MsgBox "SupportingDocumentBuilder: every table row must " & _
                    "contain at least one cell.", VBA.vbExclamation, _
                    "Supporting Document Builder"
                Exit Function
            End If
            ReDim startColumns(0 To cellCount - 1)
            ReDim cellSpans(0 To cellCount - 1)
            ReDim cellTexts(0 To cellCount - 1)
            ReDim cellIsHeader(0 To cellCount - 1)
            ReDim cellBorders(0 To cellCount - 1)
            ReDim cellAlignments(0 To cellCount - 1)
            ReDim cellBoldValues(0 To cellCount - 1)
            ReDim cellFontSizes(0 To cellCount - 1)
            ReDim cellImageSources(0 To cellCount - 1)
            ReDim cellImageWidths(0 To cellCount - 1)
            ReDim cellImageHeights(0 To cellCount - 1)
            ReDim cellFontColors(0 To cellCount - 1)
            ReDim cellBackgrounds(0 To cellCount - 1)
            columnIndex = 1
            For cellIndex = 0 To cellCount - 1
                Set htmlCell = htmlRow.Cells.Item(cellIndex)
                cellSpan = VBA.CLng(htmlCell.colSpan)
                If cellSpan <= 0 Then cellSpan = 1
                startColumns(cellIndex) = columnIndex
                cellSpans(cellIndex) = cellSpan
                cellTexts(cellIndex) = private_NormalizeWordParagraphBreaks( _
                    m_WordResultStylesParser.HtmlValueText(htmlCell.innerText))
                Set htmlImages = htmlCell.getElementsByTagName("img")
                If htmlImages.Length = 0 Then _
                    Set htmlImages = htmlCell.getElementsByTagName("image")
                If htmlImages.Length > 1 Then
                    VBA.MsgBox "SupportingDocumentBuilder: a table cell supports " & _
                        "no more than one <image>.", VBA.vbExclamation, _
                        "Supporting Document Builder"
                    Exit Function
                End If
                If htmlImages.Length = 1 Then
                    Set htmlImage = htmlImages.Item(0)
                    cellImageSources(cellIndex) = _
                        m_WordResultStylesParser.HtmlAttributeText( _
                            htmlImage, "src")
                    cellImageWidths(cellIndex) = _
                        m_WordResultStylesParser.HtmlAttributeText( _
                            htmlImage, "widthPt")
                    cellImageHeights(cellIndex) = _
                        m_WordResultStylesParser.HtmlAttributeText( _
                            htmlImage, "heightPt")
                End If
                cellIsHeader(cellIndex) = (VBA.LCase$( _
                    VBA.CStr(htmlCell.tagName)) = "th")
                cellBorders(cellIndex) = _
                    m_WordResultStylesParser.HtmlAttributeText( _
                        htmlCell, "borders")
                cellAlignments(cellIndex) = _
                    m_WordResultStylesParser.HtmlAttributeText(htmlCell, "align")
                cellBoldValues(cellIndex) = _
                    m_WordResultStylesParser.HtmlAttributeText(htmlCell, "bold")
                cellFontSizes(cellIndex) = _
                    m_WordResultStylesParser.HtmlAttributeText( _
                        htmlCell, "fontSize")
                cellFontColors(cellIndex) = _
                    m_WordResultStylesParser.HtmlAttributeText( _
                        htmlCell, "fontColor")
                cellBackgrounds(cellIndex) = _
                    m_WordResultStylesParser.HtmlAttributeText( _
                        htmlCell, "background")
                columnIndex = columnIndex + cellSpan
            Next cellIndex
            ' Объединяем справа налево, чтобы более ранние grid-координаты
            ' не сдвигались после Merge ячеек с colspan.
            For cellIndex = cellCount - 1 To 0 Step -1
                columnIndex = startColumns(cellIndex)
                cellSpan = cellSpans(cellIndex)
                If cellSpan > 1 Then
                    wordTable.Cell(rowIndex, columnIndex).Merge _
                        wordTable.Cell(rowIndex, columnIndex + cellSpan - 1)
                End If
                If VBA.Len(cellImageSources(cellIndex)) > 0 Then
                    wordTable.Cell(rowIndex, columnIndex).Range.Text = _
                        cellTexts(cellIndex)
                    Set wordRange = wordDoc.Range( _
                        wordTable.Cell(rowIndex, columnIndex).Range.Start, _
                        wordTable.Cell(rowIndex, columnIndex).Range.Start)
                    If Not private_TryInsertWordImage( _
                        wordDoc, wordRange, cellImageSources(cellIndex), _
                        cellImageWidths(cellIndex), cellImageHeights(cellIndex), _
                        cursorPos) Then Exit Function
                Else
                    wordTable.Cell(rowIndex, columnIndex).Range.Text = _
                        cellTexts(cellIndex)
                End If
                If Not private_TryApplyWordCellFormatting( _
                    wordTable.Cell(rowIndex, columnIndex), _
                    cellBorders(cellIndex), _
                    cellAlignments(cellIndex), _
                    cellBoldValues(cellIndex), _
                    cellFontSizes(cellIndex), _
                    cellFontColors(cellIndex), _
                    cellBackgrounds(cellIndex), _
                    cellIsHeader(cellIndex)) Then Exit Function
            Next cellIndex
            rowIndex = rowIndex + 1
        Next htmlRow
        tableStyle = VBA.Trim$( _
            m_WordResultStylesParser.HtmlAttributeText(htmlTable, "wordStyle"))
        If VBA.Len(tableStyle) > 0 Then
            If Not private_TryApplyWordTableStyle( _
                wordTable, tableStyle) Then Exit Function
        End If
        cursorPos = wordTable.Range.End
        previousBlockWasTable = True
        searchPos = tableEnd + 1
ContinueBlock:
    Loop

    beforeText = VBA.Mid$(renderedText, searchPos)
    If VBA.Len(beforeText) > 0 Then
        Set wordRange = wordDoc.Range(cursorPos, cursorPos)
        wordRange.Text = beforeText
        cursorPos = wordRange.End
    End If
    outInsertedEnd = cursorPos
    private_TryInsertRenderedHtmlBlocks = True
    Exit Function
End Function

Private Function private_TryApplyWordTableColumnWidths( _
    ByVal wordTable As Object, _
    ByVal columnWidthsText As String, _
    ByVal expectedColumnCount As Long _
) As Boolean
    Dim columnWidths As Collection
    Dim columnIndex As Long
    Dim rowIndex As Long
    Dim gridCellIndex As Long
    Dim gridCells As Collection
    Dim wordCell As Object

    If wordTable Is Nothing Then Exit Function
    columnWidthsText = VBA.Trim$(columnWidthsText)
    If VBA.Len(columnWidthsText) = 0 Then
        private_TryApplyWordTableColumnWidths = True
        Exit Function
    End If
    If Not m_WordResultStylesParser.TryParseColumnWidths( _
        columnWidthsText, expectedColumnCount, columnWidths) Then Exit Function

    On Error GoTo EH
    wordTable.AllowAutoFit = False
    wordTable.AutoFitBehavior WD_AUTO_FIT_FIXED
    wordTable.PreferredWidthType = WD_PREFERRED_WIDTH_PERCENT
    wordTable.PreferredWidth = 100
    ' Сохраняем ссылки до первого изменения ширины. После него Word может уже
    ' считать сетку неоднородной и возвращать error 5941 даже для Cell(r, c).
    Set gridCells = New Collection
    For rowIndex = 1 To wordTable.Rows.Count
        For columnIndex = 1 To expectedColumnCount
            Set wordCell = wordTable.Cell(rowIndex, columnIndex)
            gridCells.Add wordCell
        Next columnIndex
    Next rowIndex
    gridCellIndex = 1
    For rowIndex = 1 To wordTable.Rows.Count
        For columnIndex = 1 To expectedColumnCount
            Set wordCell = gridCells.Item(gridCellIndex)
            wordCell.PreferredWidthType = WD_PREFERRED_WIDTH_PERCENT
            wordCell.PreferredWidth = VBA.CSng(columnWidths.Item(columnIndex))
            gridCellIndex = gridCellIndex + 1
        Next columnIndex
    Next rowIndex
    private_TryApplyWordTableColumnWidths = True
    Exit Function

EH:
    VBA.MsgBox "SupportingDocumentBuilder: failed to apply Word table column " & _
        "widths '" & columnWidthsText & "': [" & VBA.CStr(Err.Number) & "] " & _
        Err.Description, VBA.vbExclamation, _
        "Supporting Document Builder"
End Function

Private Function private_TryApplyWordCellFormatting( _
    ByVal wordCell As Object, _
    ByVal bordersText As String, _
    ByVal alignmentText As String, _
    ByVal boldText As String, _
    ByVal fontSizeText As String, _
    ByVal fontColorText As String, _
    ByVal backgroundText As String, _
    ByVal isHeader As Boolean _
) As Boolean
    Dim borders As Collection
    Dim borderObj As Variant
    Dim borderName As String
    Dim colorValue As Long
    Dim parsedAlignment As String
    Dim hasBold As Boolean
    Dim parsedBold As Boolean
    Dim hasFontSize As Boolean
    Dim parsedFontSize As Single

    If wordCell Is Nothing Then Exit Function
    On Error GoTo EH

    If Not m_WordResultStylesParser.TryParseAlignment( _
        alignmentText, parsedAlignment) Then Exit Function
    Select Case parsedAlignment
        Case "left"
            wordCell.Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_LEFT
        Case "center"
            wordCell.Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER
    End Select

    wordCell.Range.Font.Bold = isHeader
    If Not m_WordResultStylesParser.TryParseOptionalBoolean( _
        boldText, hasBold, parsedBold) Then Exit Function
    If hasBold Then wordCell.Range.Font.Bold = parsedBold
    If Not m_WordResultStylesParser.TryParseOptionalPositiveSingle( _
        fontSizeText, hasFontSize, parsedFontSize) Then Exit Function
    If hasFontSize Then wordCell.Range.Font.Size = parsedFontSize
    If VBA.Len(VBA.Trim$(fontColorText)) > 0 Then
        If Not m_WordResultStylesParser.TryParseHtmlColor( _
            fontColorText, colorValue) Then Exit Function
        wordCell.Range.Font.Color = colorValue
    End If
    If VBA.Len(VBA.Trim$(backgroundText)) > 0 Then
        If Not m_WordResultStylesParser.TryParseHtmlColor( _
            backgroundText, colorValue) Then Exit Function
        wordCell.Shading.BackgroundPatternColor = colorValue
    End If

    If Not m_WordResultStylesParser.TryParseBorders( _
        bordersText, borders) Then Exit Function
    If borders.Count > 0 Then
        For Each borderObj In borders
            borderName = VBA.CStr(borderObj)
            Select Case borderName
                Case "bottom"
                    wordCell.Borders(WD_BORDER_BOTTOM).LineStyle = _
                        WD_LINE_STYLE_SINGLE
                Case "top"
                    wordCell.Borders(WD_BORDER_TOP).LineStyle = _
                        WD_LINE_STYLE_SINGLE
                Case "left"
                    wordCell.Borders(WD_BORDER_LEFT).LineStyle = _
                        WD_LINE_STYLE_SINGLE
                Case "right"
                    wordCell.Borders(WD_BORDER_RIGHT).LineStyle = _
                        WD_LINE_STYLE_SINGLE
                Case "none"
                    wordCell.Borders(WD_BORDER_TOP).LineStyle = WD_LINE_STYLE_NONE
                    wordCell.Borders(WD_BORDER_LEFT).LineStyle = WD_LINE_STYLE_NONE
                    wordCell.Borders(WD_BORDER_BOTTOM).LineStyle = WD_LINE_STYLE_NONE
                    wordCell.Borders(WD_BORDER_RIGHT).LineStyle = WD_LINE_STYLE_NONE
            End Select
        Next borderObj
    End If

    private_TryApplyWordCellFormatting = True
    Exit Function

EH:
    VBA.MsgBox "SupportingDocumentBuilder: failed to format Word table cell: [" & _
        VBA.CStr(Err.Number) & "] " & Err.Description, VBA.vbExclamation, _
        "Supporting Document Builder"
End Function

Private Function private_TryInsertRenderedImage( _
    ByVal wordDoc As Object, _
    ByVal targetRange As Object, _
    ByVal imageMarkup As String, _
    ByRef outCursorPos As Long _
) As Boolean
    Dim imageDoc As Object
    Dim imageNode As Object
    Dim imagePath As String
    Dim widthText As String
    Dim heightText As String

    If wordDoc Is Nothing Or targetRange Is Nothing Then Exit Function
    Set imageDoc = VBA.CreateObject("MSXML2.DOMDocument.6.0")
    imageDoc.async = False
    If Not imageDoc.LoadXML(imageMarkup) Then
        VBA.MsgBox "SupportingDocumentBuilder: invalid <image> markup: " & _
            imageDoc.parseError.reason, VBA.vbExclamation, _
            "Supporting Document Builder"
        Exit Function
    End If
    Set imageNode = imageDoc.DocumentElement
    If imageNode Is Nothing Then Exit Function
    imagePath = VBA.Trim$( _
        m_WordResultStylesParser.HtmlAttributeText(imageNode, "src"))
    If VBA.Len(imagePath) = 0 Then
        VBA.MsgBox "SupportingDocumentBuilder: <image> requires src.", _
            VBA.vbExclamation, "Supporting Document Builder"
        Exit Function
    End If
    widthText = VBA.Trim$( _
        m_WordResultStylesParser.HtmlAttributeText(imageNode, "widthPt"))
    heightText = VBA.Trim$( _
        m_WordResultStylesParser.HtmlAttributeText(imageNode, "heightPt"))
    private_TryInsertRenderedImage = private_TryInsertWordImage( _
        wordDoc, targetRange, imagePath, widthText, heightText, outCursorPos)
End Function

Private Function private_TryInsertWordImage( _
    ByVal wordDoc As Object, _
    ByVal targetRange As Object, _
    ByVal imagePath As String, _
    ByVal widthText As String, _
    ByVal heightText As String, _
    ByRef outCursorPos As Long _
) As Boolean
    Dim inlineShape As Object
    Dim hasWidth As Boolean
    Dim hasHeight As Boolean
    Dim parsedWidth As Single
    Dim parsedHeight As Single

    If wordDoc Is Nothing Or targetRange Is Nothing Then Exit Function
    imagePath = VBA.Trim$(imagePath)
    If VBA.Len(imagePath) = 0 Then Exit Function
    If Not private_IsAbsolutePath(imagePath) Then _
        imagePath = ThisWorkbook.Path & Application.PathSeparator & imagePath
    If VBA.Len(VBA.Dir$(imagePath, _
        VBA.vbNormal Or VBA.vbReadOnly Or VBA.vbHidden Or VBA.vbSystem)) = 0 Then
        VBA.MsgBox "SupportingDocumentBuilder: image file was not found: " & _
            imagePath, VBA.vbExclamation, "Supporting Document Builder"
        Exit Function
    End If

    If Not m_WordResultStylesParser.TryParseOptionalPositiveSingle( _
        widthText, hasWidth, parsedWidth) Then Exit Function
    If Not m_WordResultStylesParser.TryParseOptionalPositiveSingle( _
        heightText, hasHeight, parsedHeight) Then Exit Function
    On Error GoTo EH
    Set inlineShape = wordDoc.InlineShapes.AddPicture( _
        imagePath, False, True, targetRange)
    If hasWidth Then inlineShape.Width = parsedWidth
    If hasHeight Then inlineShape.Height = parsedHeight
    outCursorPos = inlineShape.Range.End
    private_TryInsertWordImage = True
    Exit Function

EH:
    VBA.MsgBox "SupportingDocumentBuilder: failed to insert image '" & _
        imagePath & "': [" & VBA.CStr(Err.Number) & "] " & Err.Description, _
        VBA.vbExclamation, "Supporting Document Builder"
End Function

Private Function private_TryApplyWordTableStyle( _
    ByVal wordTable As Object, _
    ByVal tableStyle As String _
) As Boolean
    Dim parsedTableStyle As String

    If wordTable Is Nothing Then Exit Function
    If Not m_WordResultStylesParser.TryParseTableStyle( _
        tableStyle, parsedTableStyle) Then Exit Function
    If VBA.Len(parsedTableStyle) = 0 Then
        private_TryApplyWordTableStyle = True
        Exit Function
    End If

    On Error GoTo EH
    Select Case parsedTableStyle
        Case "builtin:tablegrid"
            ' Built-in style ID одинаков во всех локализациях Word.
            wordTable.Style = WD_STYLE_TABLE_GRID
    End Select
    private_TryApplyWordTableStyle = True
    Exit Function

EH:
    VBA.MsgBox "SupportingDocumentBuilder: failed to apply Word table style '" & _
        tableStyle & "': [" & VBA.CStr(Err.Number) & "] " & _
        Err.Description, VBA.vbExclamation, _
        "Supporting Document Builder"
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
