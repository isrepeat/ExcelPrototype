VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_WordResultTplParser"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const PROFILES_NS As String = "urn:excelprototype:profiles"
Private Const TOKEN_NEWLINE As String = "{#newline}"
Private Const TOKEN_SPACE As String = "{#space}"
Private Const TOKEN_SPACE_SENTINEL As String = "<<__PEB_SPACE__>>"
Private Const TOKEN_JOIN_LINE As String = "#^"
Private Const TOKEN_JOIN_LINE_BRACED As String = "{#^}"
Private Const TOKEN_TRIM_INDENT As String = "#_"
Private Const TOKEN_TRIM_INDENT_BRACED As String = "{#_}"
Private Const TOKEN_IF_OPEN As String = "{#if"
Private Const TOKEN_IF_ELSE As String = "{#else}"
Private Const TOKEN_IF_CLOSE As String = "{#endif}"
Private Const TOKEN_FOR_OPEN As String = "{#for"
Private Const TOKEN_FOR_CLOSE As String = "{#endfor}"
Private Const TOKEN_INCLUDE_OPEN As String = "{#include"
Private Const TOKEN_LET As String = "#let"
Private Const FORMATTER_UPPER_FIRST_LETTER As String = "upperFirstLetter"
Private Const FORMATTER_LOWER_FIRST_LETTER As String = "lowerFirstLetter"
Private Const FORMATTER_REGEX_REPLACE As String = "regexreplace"
Private Const FORMATTER_DATE_OFFSET As String = "dateoffset"
Private Const FORMATTER_DATE_FORMAT As String = "dateformat"
Private Const FORMATTER_DATE_STORAGE_FORMAT As String = "dd.mm.yyyy"
Private Const PREVIEW_VALUE_MARKER As String = "preview-value"
Private Const PREVIEW_TRUNCATED_MARKER As String = "preview-truncated"
Private Const PREVIEW_TRUNCATED_VALUE_TAG As String = "preview-truncated-value"
Private Const PREVIEW_WARNING_MARKER As String = "preview-warning"
Private Const PREVIEW_LOOKUP_WARNING_VALUE_TAG As String = "preview-lookup-warning-value"

Private m_TemplateRelPath As String
Private m_TemplateDoc As Object
Private m_TemplateLastModified As Date
Private m_HasTemplateLastModified As Boolean
Private m_IsDisposed As Boolean
Private m_NamedCollections As Object
Private m_LastRenderHasTables As Boolean

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
' // API
' //
Public Function Initialize(ByVal templateRelPath As String) As Boolean
    m_IsDisposed = False
    m_TemplateRelPath = VBA.Trim$(templateRelPath)
    If VBA.Len(m_TemplateRelPath) = 0 Then Exit Function
    If Not private_TryReloadTemplateIfChanged(True) Then Exit Function
    Initialize = True
End Function

Public Sub Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    m_TemplateRelPath = VBA.vbNullString
    Set m_TemplateDoc = Nothing
    m_TemplateLastModified = 0
    m_HasTemplateLastModified = False
    Set m_NamedCollections = Nothing
End Sub

Public Function TryRenderForTemplateId( _
    ByVal templateId As String, _
    ByVal sectionTypeText As String, _
    ByVal sourceTables As Collection, _
    ByVal namedCollections As Object, _
    ByRef outResultText As String _
) As Boolean
    Dim templateText As String
    Dim renderVars As Object
    Dim loopRows As Object

    outResultText = VBA.vbNullString
    m_LastRenderHasTables = False
    If m_IsDisposed Then Exit Function
    If sourceTables Is Nothing Then Exit Function
    If sourceTables.Count <= 0 Then Exit Function
    If Not private_TryReloadTemplateIfChanged() Then Exit Function

    If Not private_TryGetTemplateTextById(templateId, templateText) Then Exit Function
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "peb-word-parser:render template='" & templateId & _
        "' section='" & sectionTypeText & "'"
#End If
    ' Keep named loop collections in parser state so nested #for blocks can
    ' resolve collections by explicit DSL key (no implicit backward fallback).
    Set m_NamedCollections = namedCollections

    Set renderVars = VBA.CreateObject("Scripting.Dictionary")
    renderVars.CompareMode = 1
    Set loopRows = VBA.CreateObject("Scripting.Dictionary")
    loopRows.CompareMode = 1

    outResultText = private_RenderTemplate(templateText, sectionTypeText, sourceTables, renderVars, loopRows)
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "word-parser:render-complete template='" & _
        templateId & "' resultLength=" & VBA.CStr(VBA.Len(outResultText)) & _
        " hasTables=" & VBA.LCase$(VBA.CStr(m_LastRenderHasTables))
#End If
    TryRenderForTemplateId = True
End Function

Public Property Get LastRenderHasTables() As Boolean
    LastRenderHasTables = m_LastRenderHasTables
End Property

Public Function TryGetGroupingDefinition( _
    ByVal templateId As String, _
    ByRef outHasGrouping As Boolean, _
    ByRef outGroupBy As String, _
    ByRef outGroupOrder As String _
) As Boolean
    Dim templateNode As Object
    Dim groupByParts As Variant
    Dim groupOrderParts As Variant
    Dim groupIndex As Long
    Dim groupHeaderNode As Object

    outHasGrouping = False
    outGroupBy = VBA.vbNullString
    outGroupOrder = VBA.vbNullString
    If m_IsDisposed Then Exit Function
    If Not private_TryReloadTemplateIfChanged() Then Exit Function
    Set templateNode = m_TemplateDoc.selectSingleNode( _
        "/p:wordResultTemplates/p:template[@id=" & _
        ex_XmlCore.fn_XPathLiteral(VBA.Trim$(templateId)) & "]")
    If templateNode Is Nothing Then
        VBA.MsgBox "PrototypeNew: WORD result template was not found by id: " & _
            templateId, VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    outGroupBy = VBA.Trim$(ex_XmlCore.fn_NodeAttrText( _
        templateNode, "groupBy"))
    outGroupOrder = VBA.LCase$(VBA.Trim$(ex_XmlCore.fn_NodeAttrText( _
        templateNode, "groupOrder")))
    outHasGrouping = (VBA.Len(outGroupBy) > 0)
    If Not outHasGrouping Then
        If Not templateNode.selectSingleNode("p:groupHeader") Is Nothing Then
            VBA.MsgBox "PrototypeNew: groupHeader requires template@groupBy: " & _
                templateId, VBA.vbExclamation, "PrototypeNew / WORD export"
            Exit Function
        End If
        TryGetGroupingDefinition = True
        Exit Function
    End If
    ' Позиция элемента в groupBy соответствует уровню вложенности:
    ' первый alias создаёт внешнюю группу, второй — группу внутри неё.
    ' groupOrder и groupHeader@by валидируются здесь, чтобы Word exporter
    ' получал уже однозначное описание структуры.
    groupByParts = VBA.Split(outGroupBy, ";")
    groupOrderParts = VBA.Split(outGroupOrder, ";")
    If UBound(groupByParts) <> UBound(groupOrderParts) Then
        VBA.MsgBox "PrototypeNew: groupBy and groupOrder must contain the same " & _
            "number of levels: " & templateId, VBA.vbExclamation, _
            "PrototypeNew / WORD export"
        Exit Function
    End If
    For groupIndex = LBound(groupByParts) To UBound(groupByParts)
        groupByParts(groupIndex) = VBA.Trim$(VBA.CStr(groupByParts(groupIndex)))
        groupOrderParts(groupIndex) = VBA.LCase$(VBA.Trim$( _
            VBA.CStr(groupOrderParts(groupIndex))))
        If VBA.Len(groupByParts(groupIndex)) = 0 Then Exit Function
        Select Case groupOrderParts(groupIndex)
            Case "date", "text", "none"
            Case Else
                VBA.MsgBox "PrototypeNew: unsupported template groupOrder '" & _
                    groupOrderParts(groupIndex) & "': " & templateId, _
                    VBA.vbExclamation, "PrototypeNew / WORD export"
                Exit Function
        End Select
        Set groupHeaderNode = templateNode.selectSingleNode( _
            "p:groupHeader[@by=" & _
            ex_XmlCore.fn_XPathLiteral(groupByParts(groupIndex)) & "]")
        If groupHeaderNode Is Nothing And UBound(groupByParts) = 0 Then
            Set groupHeaderNode = templateNode.selectSingleNode("p:groupHeader")
        End If
        If groupHeaderNode Is Nothing Then
            VBA.MsgBox "PrototypeNew: grouped WORD template requires groupHeader " & _
                "for [" & groupByParts(groupIndex) & "]: " & templateId, _
                VBA.vbExclamation, "PrototypeNew / WORD export"
            Exit Function
        End If
    Next groupIndex
    TryGetGroupingDefinition = True
End Function

Public Function TryRenderGroupHeaderForTemplateId( _
    ByVal templateId As String, _
    ByVal groupByText As String, _
    ByVal sectionTypeText As String, _
    ByVal sourceTables As Collection, _
    ByVal namedCollections As Object, _
    ByRef outResultText As String _
) As Boolean
    Dim templateText As String
    Dim renderVars As Object
    Dim loopRows As Object

    outResultText = VBA.vbNullString
    If m_IsDisposed Then Exit Function
    If sourceTables Is Nothing Or sourceTables.Count <= 0 Then Exit Function
    If Not private_TryReloadTemplateIfChanged() Then Exit Function
    If Not private_TryGetGroupHeaderTextById( _
        templateId, groupByText, templateText) Then Exit Function
    Set m_NamedCollections = namedCollections
    Set renderVars = VBA.CreateObject("Scripting.Dictionary")
    renderVars.CompareMode = 1
    Set loopRows = VBA.CreateObject("Scripting.Dictionary")
    loopRows.CompareMode = 1
    outResultText = private_RenderTemplate( _
        templateText, sectionTypeText, sourceTables, renderVars, loopRows)
    TryRenderGroupHeaderForTemplateId = True
End Function

' //
' // Internal
' //
Private Function private_TryReloadTemplateIfChanged( _
    Optional ByVal forceReload As Boolean = False _
) As Boolean
    Dim templatePath As String
    Dim currentLastModified As Date
    Dim refreshedDoc As Object

    templatePath = ex_XmlCore.fn_CombineBasePath(ThisWorkbook, m_TemplateRelPath)
    If VBA.Len(VBA.Trim$(templatePath)) = 0 Then
        VBA.MsgBox "PrototypeNew: failed to resolve WORD result templates path: " & m_TemplateRelPath, VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    On Error GoTo EH
    currentLastModified = VBA.FileDateTime(templatePath)
    On Error GoTo 0

    If Not forceReload And m_HasTemplateLastModified Then
        ' Быстрый путь: DOM остается валиден, пока timestamp файла не изменился.
        If currentLastModified = m_TemplateLastModified Then
            private_TryReloadTemplateIfChanged = True
            Exit Function
        End If
    End If

    ' Загружаем в отдельный DOM и только после успешного parse заменяем кеш.
    Set refreshedDoc = ex_XmlCore.fn_LoadDomByFilePath( _
        templatePath, _
        "Missing WORD result templates file: ", _
        "Failed to parse WORD result templates file: ", PROFILES_NS)
    If refreshedDoc Is Nothing Then
        VBA.MsgBox "PrototypeNew: failed to reload WORD result templates file: " & templatePath, VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    Set m_TemplateDoc = refreshedDoc
    m_TemplateLastModified = currentLastModified
    m_HasTemplateLastModified = True
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "peb-word-parser:templates-reloaded path='" & templatePath & "'"
#End If
    private_TryReloadTemplateIfChanged = True
    Exit Function

EH:
    VBA.MsgBox "PrototypeNew: failed to read WORD result templates modification date: " & templatePath & VBA.vbCrLf & Err.Description, VBA.vbExclamation, "PrototypeNew / WORD export"
End Function

Private Function private_TryGetTemplateTextById( _
    ByVal templateId As String, _
    ByRef outTemplateText As String _
) As Boolean
    Dim doc As Object
    Dim node As Object
    Dim xpath As String
    Dim includeChain As Collection
    Dim templateNode As Object
    Dim childNode As Object
    Dim contentText As String

    outTemplateText = VBA.vbNullString
    templateId = VBA.Trim$(templateId)
    If VBA.Len(templateId) = 0 Then
        VBA.MsgBox "PrototypeNew: WORD result template id is empty.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    Set doc = m_TemplateDoc
    If doc Is Nothing Then
        VBA.MsgBox "PrototypeNew: failed to load WORD result templates: " & m_TemplateRelPath, VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    xpath = "/p:wordResultTemplates/p:template[@id=" & ex_XmlCore.fn_XPathLiteral(templateId) & "]"
    Set templateNode = doc.selectSingleNode(xpath)
    If templateNode Is Nothing Then
        VBA.MsgBox "PrototypeNew: WORD result template was not found by id: " & templateId, VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    Set includeChain = New Collection
    For Each childNode In templateNode.childNodes
        Select Case VBA.LCase$(VBA.CStr(childNode.baseName))
            Case "text"
                contentText = contentText & private_ExpandSharedTemplateIncludes( _
                    VBA.CStr(childNode.Text), doc, includeChain)
            Case "table"
                m_LastRenderHasTables = True
                contentText = contentText & private_SerializeHtmlNode(childNode)
        End Select
    Next childNode
    If VBA.Len(contentText) = 0 Then
        VBA.MsgBox "SupportingDocumentBuilder: template has no text or table blocks: " & _
            templateId, VBA.vbExclamation, "Supporting Document Builder"
        Exit Function
    End If
    outTemplateText = contentText
    private_TryGetTemplateTextById = True
End Function

Private Function private_SerializeHtmlNode(ByVal sourceNode As Object) As String
    Dim nodeName As String
    Dim resultText As String
    Dim childNode As Object
    Dim attrNode As Object
    Dim loopExpression As String

    If sourceNode Is Nothing Then Exit Function
    If sourceNode.nodeType = 3 Or sourceNode.nodeType = 4 Then
        private_SerializeHtmlNode = VBA.CStr(sourceNode.nodeValue)
        Exit Function
    End If
    nodeName = VBA.LCase$(VBA.CStr(sourceNode.baseName))
    Select Case nodeName
        Case "table", "caption", "thead", "tbody", "tfoot", "tr", "th", "td"
        Case Else
            VBA.MsgBox "SupportingDocumentBuilder: unsupported table tag <" & _
                nodeName & ">.", VBA.vbExclamation, "Supporting Document Builder"
            Exit Function
    End Select
    On Error Resume Next
    loopExpression = VBA.Trim$(VBA.CStr(sourceNode.getAttribute("for")))
    On Error GoTo 0
    If nodeName = "tr" And VBA.Len(loopExpression) > 0 Then
        resultText = "{#for " & loopExpression & "}"
    End If
    resultText = resultText & "<" & nodeName
    For Each attrNode In sourceNode.Attributes
        If VBA.StrComp(VBA.CStr(attrNode.nodeName), "for", VBA.vbTextCompare) <> 0 Then
            resultText = resultText & " " & VBA.CStr(attrNode.nodeName) & _
                "=""" & private_HtmlEncode(VBA.CStr(attrNode.nodeValue)) & """"
        End If
    Next attrNode
    resultText = resultText & ">"
    For Each childNode In sourceNode.childNodes
        resultText = resultText & private_SerializeHtmlNode(childNode)
    Next childNode
    resultText = resultText & "</" & nodeName & ">"
    If nodeName = "tr" And VBA.Len(loopExpression) > 0 Then
        resultText = resultText & "{#endfor}"
    End If
    private_SerializeHtmlNode = resultText
End Function

Private Function private_HtmlEncode(ByVal valueText As String) As String
    valueText = VBA.Replace(valueText, "&", "&amp;")
    valueText = VBA.Replace(valueText, """", "&quot;")
    valueText = VBA.Replace(valueText, "<", "&lt;")
    private_HtmlEncode = VBA.Replace(valueText, ">", "&gt;")
End Function

Private Function private_TryGetTemplateChildTextById( _
    ByVal templateId As String, _
    ByVal childName As String, _
    ByRef outTemplateText As String _
) As Boolean
    Dim node As Object
    Dim xpath As String
    Dim includeChain As Collection

    outTemplateText = VBA.vbNullString
    templateId = VBA.Trim$(templateId)
    childName = VBA.Trim$(childName)
    If VBA.Len(templateId) = 0 Or VBA.Len(childName) = 0 Then Exit Function
    xpath = "/p:wordResultTemplates/p:template[@id=" & _
        ex_XmlCore.fn_XPathLiteral(templateId) & "]/p:" & childName
    Set node = m_TemplateDoc.selectSingleNode(xpath)
    If node Is Nothing Then
        VBA.MsgBox "PrototypeNew: WORD template element '" & childName & _
            "' was not found: " & templateId, _
            VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If
    Set includeChain = New Collection
    outTemplateText = private_ExpandSharedTemplateIncludes( _
        VBA.CStr(node.Text), m_TemplateDoc, includeChain)
    private_TryGetTemplateChildTextById = True
End Function

Private Function private_TryGetGroupHeaderTextById( _
    ByVal templateId As String, _
    ByVal groupByText As String, _
    ByRef outTemplateText As String _
) As Boolean
    Dim templateNode As Object
    Dim node As Object
    Dim includeChain As Collection

    outTemplateText = VBA.vbNullString
    Set templateNode = m_TemplateDoc.selectSingleNode( _
        "/p:wordResultTemplates/p:template[@id=" & _
        ex_XmlCore.fn_XPathLiteral(VBA.Trim$(templateId)) & "]")
    If templateNode Is Nothing Then Exit Function
    ' Заголовок выбирается по alias, а не по его порядку среди XML-узлов.
    ' Благодаря этому перестановка groupHeader в файле не меняет семантику DSL.
    Set node = templateNode.selectSingleNode( _
        "p:groupHeader[@by=" & _
        ex_XmlCore.fn_XPathLiteral(VBA.Trim$(groupByText)) & "]")
    If node Is Nothing And templateNode.selectNodes("p:groupHeader").Length = 1 Then
        Set node = templateNode.selectSingleNode("p:groupHeader")
    End If
    If node Is Nothing Then
        VBA.MsgBox "PrototypeNew: WORD groupHeader was not found for [" & _
            groupByText & "]: " & templateId, VBA.vbExclamation, _
            "PrototypeNew / WORD export"
        Exit Function
    End If
    Set includeChain = New Collection
    outTemplateText = private_ExpandSharedTemplateIncludes( _
        VBA.CStr(node.Text), m_TemplateDoc, includeChain)
    private_TryGetGroupHeaderTextById = True
End Function

Private Function private_RenderTemplate( _
    ByVal templateText As String, _
    ByVal sectionTypeText As String, _
    ByVal sourceTables As Collection, _
    ByVal renderVars As Object, _
    ByVal loopRows As Object, _
    Optional ByVal normalizeOutput As Boolean = True _
) As String
    Dim resultText As String
    Dim rx As Object
    Dim matches As Object
    Dim matchObj As Object
    Dim matchIndex As Long
    Dim placeholderName As String
    Dim placeholderValue As String

    ' Поддерживаем PEB WORD-preview syntax:
    '   {#include id}    -> вставить sharedTemplate из XML
    '   #let A = expr;   -> сохранить локальное значение для последующих blocks/placeholders
    '   {#if expr}...{#else}...{#endif}
    '   {#for item in Collection}...{#endfor}
    '   #^ / {#^}        -> удалить token и следующий перенос строки
    '   #_ / {#_}        -> удалить token и пробелы/табуляцию до первого символа
    '   {#newline}       -> перенос строки
    '   {#space}         -> один пробел
    '   {SectionType}    -> тип секции из export context/source table
    '   {[Column]|formatter} -> значение колонки DynamicTable по alias/заголовку + formatter pipeline
    '      dateoffset:"+1"             -> сдвинуть дату на N дней
    '      dateformat:"\dd \month"     -> отформатировать дату в текст
    '   {item.[Column]} -> значение колонки строки внутри #for
    '
    ' Это не полный универсальный parser: внешние VBA-вызовы, падежи,
    ' даты, colors и морфология намеренно не перенесены в этот PEB preview.
    resultText = VBA.CStr(templateText)
    ' Render order is important: first evaluate control structures, then apply
    ' layout tokens (#^/#_), then resolve placeholders and normalize output.
    resultText = private_ResolveLetBindings(resultText, sectionTypeText, sourceTables, renderVars, loopRows)
    resultText = private_RenderForBlocks(resultText, sectionTypeText, sourceTables, renderVars, loopRows)
    resultText = private_RenderIfBlocks(resultText, sectionTypeText, sourceTables, renderVars, loopRows)
    resultText = private_ApplyLayoutTokens(resultText)
    resultText = VBA.Replace(resultText, TOKEN_NEWLINE, VBA.vbCrLf)
    ' Keep explicit template spaces protected until final output normalization.
    resultText = VBA.Replace(resultText, TOKEN_SPACE, TOKEN_SPACE_SENTINEL)

    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = False
    ' Берем только обычные placeholders. Служебные токены {#...} сюда не попадают.
    rx.Pattern = "\{([^#][^{}]*)\}"

    Set matches = rx.Execute(resultText)
    If matches Is Nothing Then
        If normalizeOutput Then
            private_RenderTemplate = private_NormalizeRenderedText(resultText)
        Else
            private_RenderTemplate = resultText
        End If
        Exit Function
    End If

    For matchIndex = matches.Count - 1 To 0 Step -1
        Set matchObj = matches.Item(matchIndex)
        placeholderName = VBA.Trim$(VBA.CStr(matchObj.SubMatches(0)))
        placeholderValue = private_GetPlaceholderValue(placeholderName, sectionTypeText, sourceTables, renderVars, loopRows)
        ' Parser размечает только семантику участка. Конкретный цвет выбирает
        ' общий style pipeline по inlinePart rule страницы.
        If VBA.Len(placeholderValue) > 0 Then
            If private_MainPlaceholderHasCellTag( _
                placeholderName, sourceTables, _
                PREVIEW_LOOKUP_WARNING_VALUE_TAG) Then
                placeholderValue = "[[" & PREVIEW_WARNING_MARKER & "]]" & _
                    placeholderValue & "[[/" & PREVIEW_WARNING_MARKER & "]]"
            ElseIf private_MainPlaceholderHasCellTag( _
                placeholderName, sourceTables, PREVIEW_TRUNCATED_VALUE_TAG) Then
                placeholderValue = "[[" & PREVIEW_TRUNCATED_MARKER & "]]" & _
                    placeholderValue & "[[/" & PREVIEW_TRUNCATED_MARKER & "]]"
            ElseIf VBA.InStr(1, placeholderValue, "[[", VBA.vbBinaryCompare) = 0 Then
                placeholderValue = "[[" & PREVIEW_VALUE_MARKER & "]]" & _
                    placeholderValue & "[[/" & PREVIEW_VALUE_MARKER & "]]"
            End If
        End If
        resultText = VBA.Left$(resultText, matchObj.FirstIndex) & _
            placeholderValue & _
            VBA.Mid$(resultText, matchObj.FirstIndex + matchObj.Length + 1)
    Next matchIndex

    If normalizeOutput Then
        private_RenderTemplate = private_ResolveSpaceSentinels(private_NormalizeRenderedText(resultText))
    Else
        ' Keep sentinel unresolved in nested renders, otherwise outer #_ can
        ' trim the restored leading space before conditional fragments.
        private_RenderTemplate = resultText
    End If
End Function

Private Function private_ResolveSpaceSentinels(ByVal valueText As String) As String
    ' Final pass: convert protected {#space} sentinels back to regular spaces.
    private_ResolveSpaceSentinels = VBA.Replace(valueText, TOKEN_SPACE_SENTINEL, " ")
End Function

Private Function private_GetPlaceholderValue( _
    ByVal placeholderName As String, _
    ByVal sectionTypeText As String, _
    ByVal sourceTables As Collection, _
    ByVal renderVars As Object, _
    ByVal loopRows As Object _
) As String
    Dim placeholderParts As Collection
    Dim valueName As String
    Dim resultValue As String
    Dim formatterIndex As Long

    Set placeholderParts = private_SplitByDelimiterOutsideQuotes(VBA.Trim$(placeholderName), "|")
    If placeholderParts Is Nothing Then Exit Function
    If placeholderParts.Count <= 0 Then Exit Function

    valueName = VBA.Trim$(VBA.CStr(placeholderParts.Item(1)))
    resultValue = private_GetRawPlaceholderValue(valueName, sectionTypeText, sourceTables, renderVars, loopRows)

    ' Formatter pipeline пишется прямо в placeholder:
    '   {[FIO]|lowerFirstLetter|regexreplace:"\s+"," "}
    ' Разбиваем по | только вне кавычек, чтобы regex с alternation не ломал
    ' список formatter-ов.
    For formatterIndex = 2 To placeholderParts.Count
        resultValue = private_ApplyFormatter(resultValue, VBA.Trim$(VBA.CStr(placeholderParts.Item(formatterIndex))))
    Next formatterIndex

    private_GetPlaceholderValue = resultValue
End Function

Private Function private_GetRawPlaceholderValue( _
    ByVal placeholderName As String, _
    ByVal sectionTypeText As String, _
    ByVal sourceTables As Collection, _
    ByVal renderVars As Object, _
    ByVal loopRows As Object _
) As String
    Dim loopVarName As String
    Dim loopFieldName As String
    Dim dotPos As Long

    placeholderName = VBA.Trim$(placeholderName)
    If VBA.Len(placeholderName) = 0 Then Exit Function

    ' Pseudo-placeholder, которого нет в таблице формы.
    If VBA.StrComp(placeholderName, "SectionType", VBA.vbTextCompare) = 0 Then
        private_GetRawPlaceholderValue = VBA.Trim$(sectionTypeText)
        Exit Function
    End If

    If Not renderVars Is Nothing Then
        If renderVars.Exists(placeholderName) Then
            private_GetRawPlaceholderValue = VBA.Trim$(VBA.CStr(renderVars(placeholderName)))
            Exit Function
        End If
    End If

    dotPos = VBA.InStr(1, placeholderName, ".", VBA.vbBinaryCompare)
    If dotPos > 1 Then
        loopVarName = VBA.Trim$(VBA.Left$(placeholderName, dotPos - 1))
        loopFieldName = VBA.Trim$(VBA.Mid$(placeholderName, dotPos + 1))
        If private_IsBracketFieldToken(loopFieldName) Then
            loopFieldName = private_UnwrapBracketFieldToken(loopFieldName)
            If private_TryGetLoopFieldValue(loopRows, loopVarName, loopFieldName, private_GetRawPlaceholderValue) Then Exit Function
        ElseIf VBA.Left$(loopFieldName, 2) = "__" Then
            If private_TryGetLoopFieldValue(loopRows, loopVarName, loopFieldName, private_GetRawPlaceholderValue) Then Exit Function
        End If
        Exit Function
    End If

    If private_IsBracketFieldToken(placeholderName) Then
        ' Поля DynamicTable в новом DSL читаются только через {[AliasOrHeader]}.
        ' Так они визуально отличаются от DSL-переменных вроде {SectionType}
        ' или {item.__first}.
        private_GetRawPlaceholderValue = private_GetMainTableFieldValue(private_UnwrapBracketFieldToken(placeholderName), sourceTables)
    End If
End Function

Private Function private_ExpandSharedTemplateIncludes( _
    ByVal sourceText As String, _
    ByVal doc As Object, _
    ByVal includeChain As Collection _
) As String
    Dim resultText As String
    Dim rx As Object
    Dim matches As Object
    Dim matchObj As Object
    Dim includeId As String
    Dim includeText As String
    Dim expandedText As String

    resultText = VBA.CStr(sourceText)
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True
    rx.MultiLine = True
    rx.Pattern = "\{#include\s+([A-Za-z_][A-Za-z0-9_.-]*)\s*\}"

    Do
        Set matches = rx.Execute(resultText)
        If matches Is Nothing Then Exit Do
        If matches.Count <= 0 Then Exit Do

        Set matchObj = matches.Item(0)
        includeId = VBA.Trim$(VBA.CStr(matchObj.SubMatches(0)))
        If private_IncludeChainContains(includeChain, includeId) Then
            VBA.MsgBox "PrototypeNew: circular WORD shared template include: " & includeId, VBA.vbExclamation, "PrototypeNew / WORD export"
            private_ExpandSharedTemplateIncludes = resultText
            Exit Function
        End If

        includeText = private_GetSharedTemplateText(doc, includeId)
        includeChain.Add includeId
        expandedText = private_ExpandSharedTemplateIncludes(includeText, doc, includeChain)
        includeChain.Remove includeChain.Count

        resultText = VBA.Left$(resultText, matchObj.FirstIndex) & _
            expandedText & _
            VBA.Mid$(resultText, matchObj.FirstIndex + matchObj.Length + 1)
    Loop

    If VBA.InStr(1, resultText, TOKEN_INCLUDE_OPEN, VBA.vbTextCompare) > 0 Then
        VBA.MsgBox "PrototypeNew: invalid WORD #include directive. Use {#include SharedTemplateId}.", VBA.vbExclamation, "PrototypeNew / WORD export"
    End If

    private_ExpandSharedTemplateIncludes = resultText
End Function

Private Function private_GetSharedTemplateText(ByVal doc As Object, ByVal sharedTemplateId As String) As String
    Dim node As Object
    Dim xpath As String

    sharedTemplateId = VBA.Trim$(sharedTemplateId)
    If doc Is Nothing Then Exit Function
    If VBA.Len(sharedTemplateId) = 0 Then Exit Function

    xpath = "/p:wordResultTemplates/p:sharedTemplates/p:sharedTemplate[@id=" & ex_XmlCore.fn_XPathLiteral(sharedTemplateId) & "]/p:text"
    Set node = doc.selectSingleNode(xpath)
    If node Is Nothing Then
        VBA.MsgBox "PrototypeNew: WORD shared template was not found: " & sharedTemplateId, VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    private_GetSharedTemplateText = VBA.CStr(node.Text)
End Function

Private Function private_IncludeChainContains(ByVal includeChain As Collection, ByVal includeId As String) As Boolean
    Dim itemIndex As Long

    If includeChain Is Nothing Then Exit Function
    For itemIndex = 1 To includeChain.Count
        If VBA.StrComp(VBA.CStr(includeChain.Item(itemIndex)), includeId, VBA.vbTextCompare) = 0 Then
            private_IncludeChainContains = True
            Exit Function
        End If
    Next itemIndex
End Function

Private Function private_ResolveLetBindings( _
    ByVal sourceText As String, _
    ByVal sectionTypeText As String, _
    ByVal sourceTables As Collection, _
    ByVal renderVars As Object, _
    ByVal loopRows As Object _
) As String
    Dim resultText As String
    Dim rx As Object
    Dim matches As Object
    Dim matchObj As Object
    Dim varName As String
    Dim expressionText As String
    Dim expressionValue As String

    resultText = VBA.CStr(sourceText)
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = False
    rx.MultiLine = True
    rx.Pattern = "#let\s+([A-Za-z_][A-Za-z0-9_]*)\s*=\s*([^;]+?)\s*;"

    Do
        Set matches = rx.Execute(resultText)
        If matches Is Nothing Then Exit Do
        If matches.Count <= 0 Then Exit Do

        Set matchObj = matches.Item(0)
        varName = VBA.Trim$(VBA.CStr(matchObj.SubMatches(0)))
        expressionText = VBA.Trim$(VBA.CStr(matchObj.SubMatches(1)))
        expressionValue = private_EvaluateExpressionText(expressionText, sectionTypeText, sourceTables, renderVars, loopRows)
        If Not renderVars Is Nothing Then renderVars(varName) = expressionValue

        resultText = VBA.Left$(resultText, matchObj.FirstIndex) & _
            VBA.Mid$(resultText, matchObj.FirstIndex + matchObj.Length + 1)
    Loop

    If VBA.InStr(1, resultText, TOKEN_LET, VBA.vbTextCompare) > 0 Then
        VBA.MsgBox "PrototypeNew: invalid WORD #let syntax. Use #let VarName = expression;.", VBA.vbExclamation, "PrototypeNew / WORD export"
    End If

    private_ResolveLetBindings = resultText
End Function

Private Function private_RenderForBlocks( _
    ByVal sourceText As String, _
    ByVal sectionTypeText As String, _
    ByVal sourceTables As Collection, _
    ByVal renderVars As Object, _
    ByVal loopRows As Object _
) As String
    Dim resultText As String
    Dim closePos As Long
    Dim openPos As Long
    Dim openEndPos As Long
    Dim headerText As String
    Dim bodyText As String
    Dim renderedText As String
    Dim loopVarName As String
    Dim collectionName As String

    resultText = VBA.CStr(sourceText)

    Do
        closePos = VBA.InStr(1, resultText, TOKEN_FOR_CLOSE, VBA.vbTextCompare)
        If closePos <= 0 Then Exit Do
        openPos = VBA.InStrRev(resultText, TOKEN_FOR_OPEN, closePos, VBA.vbTextCompare)
        If openPos <= 0 Then Exit Do

        openEndPos = VBA.InStr(openPos, resultText, "}", VBA.vbBinaryCompare)
        If openEndPos <= openPos Or openEndPos > closePos Then Exit Do

        headerText = VBA.Mid$(resultText, openPos, openEndPos - openPos + 1)
        If Not private_TryParseForHeader(headerText, loopVarName, collectionName) Then Exit Do

        bodyText = VBA.Mid$(resultText, openEndPos + 1, closePos - openEndPos - 1)
        renderedText = private_RenderForCollection(collectionName, loopVarName, bodyText, sectionTypeText, sourceTables, renderVars, loopRows)

        resultText = VBA.Left$(resultText, openPos - 1) & _
            renderedText & _
            VBA.Mid$(resultText, closePos + VBA.Len(TOKEN_FOR_CLOSE))
    Loop

    If VBA.InStr(1, resultText, TOKEN_FOR_OPEN, VBA.vbTextCompare) > 0 _
        Or VBA.InStr(1, resultText, TOKEN_FOR_CLOSE, VBA.vbTextCompare) > 0 Then
        VBA.MsgBox "PrototypeNew: invalid WORD #for block. Use {#for item in Collection}...{#endfor}.", VBA.vbExclamation, "PrototypeNew / WORD export"
    End If

    private_RenderForBlocks = resultText
End Function

Private Function private_TryParseForHeader( _
    ByVal headerText As String, _
    ByRef outLoopVarName As String, _
    ByRef outCollectionName As String _
) As Boolean
    Dim rx As Object
    Dim matches As Object

    outLoopVarName = VBA.vbNullString
    outCollectionName = VBA.vbNullString

    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True
    rx.Pattern = "^\{#for\s+([A-Za-z_][A-Za-z0-9_]*)\s+in\s+([^}]+)\}$"

    Set matches = rx.Execute(VBA.Trim$(headerText))
    If matches Is Nothing Then Exit Function
    If matches.Count <= 0 Then Exit Function

    outLoopVarName = VBA.Trim$(VBA.CStr(matches.Item(0).SubMatches(0)))
    outCollectionName = VBA.Trim$(VBA.CStr(matches.Item(0).SubMatches(1)))
    private_TryParseForHeader = (VBA.Len(outLoopVarName) > 0 And VBA.Len(outCollectionName) > 0)
End Function

Private Function private_RenderForCollection( _
    ByVal collectionName As String, _
    ByVal loopVarName As String, _
    ByVal bodyText As String, _
    ByVal sectionTypeText As String, _
    ByVal sourceTables As Collection, _
    ByVal renderVars As Object, _
    ByVal loopRows As Object _
) As String
    Dim loopTableRows As Collection
    Dim loopItem As Variant
    Dim loopRowCtx As Object
    Dim rowIndex As Long
    Dim renderedPart As String
    Dim oldLoopRow As Object
    Dim hadOldLoopRow As Boolean

    Set loopTableRows = private_BuildLoopRows(collectionName, sourceTables)
    If loopTableRows Is Nothing Then Exit Function
    If loopTableRows.Count <= 0 Then Exit Function

    If Not loopRows Is Nothing Then
        hadOldLoopRow = loopRows.Exists(loopVarName)
        If hadOldLoopRow Then
            Set oldLoopRow = loopRows(loopVarName)
        End If
    End If

    rowIndex = 0
    For Each loopItem In loopTableRows
        rowIndex = rowIndex + 1
        Set loopRowCtx = loopItem
        loopRowCtx("__first") = (rowIndex = 1)
        loopRowCtx("__last") = (rowIndex = loopTableRows.Count)
        loopRowCtx("__index") = rowIndex

        If loopRows.Exists(loopVarName) Then loopRows.Remove loopVarName
        loopRows.Add loopVarName, loopRowCtx
        renderedPart = renderedPart & private_RenderTemplate(bodyText, sectionTypeText, sourceTables, renderVars, loopRows, False)
    Next loopItem

    If Not loopRows Is Nothing Then
        If hadOldLoopRow Then
            If loopRows.Exists(loopVarName) Then loopRows.Remove loopVarName
            loopRows.Add loopVarName, oldLoopRow
        ElseIf loopRows.Exists(loopVarName) Then
            loopRows.Remove loopVarName
        End If
    End If

    private_RenderForCollection = renderedPart
End Function

Private Function private_RenderIfBlocks( _
    ByVal sourceText As String, _
    ByVal sectionTypeText As String, _
    ByVal sourceTables As Collection, _
    ByVal renderVars As Object, _
    ByVal loopRows As Object _
) As String
    Dim resultText As String
    Dim closePos As Long
    Dim openPos As Long
    Dim openEndPos As Long
    Dim elsePos As Long
    Dim secondElsePos As Long
    Dim headerText As String
    Dim conditionText As String
    Dim trueBodyText As String
    Dim falseBodyText As String
    Dim selectedBodyText As String
    Dim renderedText As String
    Dim conditionResult As Boolean

    resultText = VBA.CStr(sourceText)

    Do
        closePos = VBA.InStr(1, resultText, TOKEN_IF_CLOSE, VBA.vbTextCompare)
        If closePos <= 0 Then Exit Do
        openPos = VBA.InStrRev(resultText, TOKEN_IF_OPEN, closePos, VBA.vbTextCompare)
        If openPos <= 0 Then Exit Do

        openEndPos = VBA.InStr(openPos, resultText, "}", VBA.vbBinaryCompare)
        If openEndPos <= openPos Or openEndPos > closePos Then Exit Do

        headerText = VBA.Mid$(resultText, openPos, openEndPos - openPos + 1)
        conditionText = private_ParseIfCondition(headerText)

        ' К этому моменту найден самый внутренний #if: вложенные блоки уже
        ' обработаны предыдущими итерациями. Поэтому первый #else между
        ' заголовком и #endif принадлежит именно текущему условию.
        elsePos = VBA.InStr(openEndPos + 1, resultText, TOKEN_IF_ELSE, VBA.vbTextCompare)
        If elsePos > 0 And elsePos < closePos Then
            secondElsePos = VBA.InStr(elsePos + VBA.Len(TOKEN_IF_ELSE), resultText, TOKEN_IF_ELSE, VBA.vbTextCompare)
            If secondElsePos > 0 And secondElsePos < closePos Then Exit Do
            trueBodyText = VBA.Mid$(resultText, openEndPos + 1, elsePos - openEndPos - 1)
            falseBodyText = VBA.Mid$(resultText, elsePos + VBA.Len(TOKEN_IF_ELSE), closePos - elsePos - VBA.Len(TOKEN_IF_ELSE))
        Else
            trueBodyText = VBA.Mid$(resultText, openEndPos + 1, closePos - openEndPos - 1)
            falseBodyText = VBA.vbNullString
        End If

        conditionResult = private_EvaluateCondition(conditionText, sectionTypeText, sourceTables, renderVars, loopRows)
        If conditionResult Then
            selectedBodyText = trueBodyText
        Else
            selectedBodyText = falseBodyText
        End If
        If VBA.Len(selectedBodyText) > 0 Then
            renderedText = private_RenderTemplate(selectedBodyText, sectionTypeText, sourceTables, renderVars, loopRows, False)
        Else
            renderedText = VBA.vbNullString
        End If

        resultText = VBA.Left$(resultText, openPos - 1) & _
            renderedText & _
            VBA.Mid$(resultText, closePos + VBA.Len(TOKEN_IF_CLOSE))
    Loop

    If VBA.InStr(1, resultText, TOKEN_IF_OPEN, VBA.vbTextCompare) > 0 _
        Or VBA.InStr(1, resultText, TOKEN_IF_ELSE, VBA.vbTextCompare) > 0 _
        Or VBA.InStr(1, resultText, TOKEN_IF_CLOSE, VBA.vbTextCompare) > 0 Then
        VBA.MsgBox "PrototypeNew: invalid WORD #if block. Use {#if expression}...{#else}...{#endif}.", VBA.vbExclamation, "PrototypeNew / WORD export"
    End If

    private_RenderIfBlocks = resultText
End Function

Private Function private_ParseIfCondition(ByVal headerText As String) As String
    headerText = VBA.Trim$(headerText)
    If VBA.Left$(headerText, VBA.Len(TOKEN_IF_OPEN)) <> TOKEN_IF_OPEN Then Exit Function
    If VBA.Right$(headerText, 1) <> "}" Then Exit Function
    private_ParseIfCondition = VBA.Trim$(VBA.Mid$(headerText, VBA.Len(TOKEN_IF_OPEN) + 1, VBA.Len(headerText) - VBA.Len(TOKEN_IF_OPEN) - 1))
End Function

Private Function private_EvaluateCondition( _
    ByVal conditionText As String, _
    ByVal sectionTypeText As String, _
    ByVal sourceTables As Collection, _
    ByVal renderVars As Object, _
    ByVal loopRows As Object _
) As Boolean
    Dim orParts() As String
    Dim andParts() As String
    Dim orIndex As Long
    Dim andIndex As Long
    Dim allAndTrue As Boolean

    conditionText = VBA.Trim$(conditionText)
    If VBA.Len(conditionText) = 0 Then Exit Function

    orParts = VBA.Split(conditionText, " #or ")
    For orIndex = LBound(orParts) To UBound(orParts)
        andParts = VBA.Split(VBA.Trim$(VBA.CStr(orParts(orIndex))), " #and ")
        allAndTrue = True
        For andIndex = LBound(andParts) To UBound(andParts)
            If Not private_EvaluateSimpleCondition(VBA.Trim$(VBA.CStr(andParts(andIndex))), sectionTypeText, sourceTables, renderVars, loopRows) Then
                allAndTrue = False
                Exit For
            End If
        Next andIndex

        If allAndTrue Then
            private_EvaluateCondition = True
            Exit Function
        End If
    Next orIndex
End Function

Private Function private_EvaluateSimpleCondition( _
    ByVal conditionText As String, _
    ByVal sectionTypeText As String, _
    ByVal sourceTables As Collection, _
    ByVal renderVars As Object, _
    ByVal loopRows As Object _
) As Boolean
    Dim operators As Variant
    Dim opIndex As Long
    Dim opText As String
    Dim opPos As Long
    Dim leftValue As String
    Dim rightValue As String

    conditionText = VBA.Trim$(conditionText)
    If VBA.Left$(conditionText, VBA.Len("#not ")) = "#not " Then
        private_EvaluateSimpleCondition = Not private_EvaluateSimpleCondition( _
            VBA.Trim$(VBA.Mid$(conditionText, VBA.Len("#not ") + 1)), _
            sectionTypeText, sourceTables, renderVars, loopRows)
        Exit Function
    End If

    operators = Array(">=", "<=", "==", "!=", ">", "<")
    For opIndex = LBound(operators) To UBound(operators)
        opText = VBA.CStr(operators(opIndex))
        opPos = VBA.InStr(1, conditionText, opText, VBA.vbBinaryCompare)
        If opPos > 0 Then
            leftValue = private_EvaluateExpressionText(VBA.Left$(conditionText, opPos - 1), sectionTypeText, sourceTables, renderVars, loopRows)
            rightValue = private_EvaluateExpressionText(VBA.Mid$(conditionText, opPos + VBA.Len(opText)), sectionTypeText, sourceTables, renderVars, loopRows)
            private_EvaluateSimpleCondition = private_CompareConditionValues(leftValue, rightValue, opText)
            Exit Function
        End If
    Next opIndex

    private_EvaluateSimpleCondition = private_IsTruthy(private_EvaluateExpressionText(conditionText, sectionTypeText, sourceTables, renderVars, loopRows))
End Function

Private Function private_CompareConditionValues( _
    ByVal leftValue As String, _
    ByVal rightValue As String, _
    ByVal opText As String _
) As Boolean
    If VBA.IsNumeric(leftValue) And VBA.IsNumeric(rightValue) Then
        Select Case opText
            Case "==": private_CompareConditionValues = (CDbl(leftValue) = CDbl(rightValue))
            Case "!=": private_CompareConditionValues = (CDbl(leftValue) <> CDbl(rightValue))
            Case ">": private_CompareConditionValues = (CDbl(leftValue) > CDbl(rightValue))
            Case "<": private_CompareConditionValues = (CDbl(leftValue) < CDbl(rightValue))
            Case ">=": private_CompareConditionValues = (CDbl(leftValue) >= CDbl(rightValue))
            Case "<=": private_CompareConditionValues = (CDbl(leftValue) <= CDbl(rightValue))
        End Select
        Exit Function
    End If

    Select Case opText
        Case "==": private_CompareConditionValues = (VBA.StrComp(leftValue, rightValue, VBA.vbTextCompare) = 0)
        Case "!=": private_CompareConditionValues = (VBA.StrComp(leftValue, rightValue, VBA.vbTextCompare) <> 0)
        Case ">": private_CompareConditionValues = (VBA.StrComp(leftValue, rightValue, VBA.vbTextCompare) > 0)
        Case "<": private_CompareConditionValues = (VBA.StrComp(leftValue, rightValue, VBA.vbTextCompare) < 0)
        Case ">=": private_CompareConditionValues = (VBA.StrComp(leftValue, rightValue, VBA.vbTextCompare) >= 0)
        Case "<=": private_CompareConditionValues = (VBA.StrComp(leftValue, rightValue, VBA.vbTextCompare) <= 0)
    End Select
End Function

Private Function private_EvaluateExpressionText( _
    ByVal expressionText As String, _
    ByVal sectionTypeText As String, _
    ByVal sourceTables As Collection, _
    ByVal renderVars As Object, _
    ByVal loopRows As Object _
) As String
    Dim trimmedText As String
    Dim rx As Object
    Dim matches As Object
    Dim matchObj As Object
    Dim resultText As String
    Dim placeholderName As String
    Dim placeholderValue As String
    Dim functionValue As String

    trimmedText = VBA.Trim$(expressionText)
    If VBA.Len(trimmedText) = 0 Then Exit Function
    If private_TryEvaluateHelperFunction(trimmedText, sectionTypeText, sourceTables, renderVars, loopRows, functionValue) Then
        private_EvaluateExpressionText = functionValue
        Exit Function
    End If
    If private_IsQuoted(trimmedText) Then
        private_EvaluateExpressionText = VBA.Mid$(trimmedText, 2, VBA.Len(trimmedText) - 2)
        Exit Function
    End If

    ' Bracket tokens always represent DynamicTable fields. A missing column is
    ' an empty value, not a literal expression such as "[DateFromShort]".
    ' Returning the token text made missing optional fields truthy in #if and
    ' prevented their fallback branches from rendering.
    If private_IsBracketFieldToken(trimmedText) _
        Or (VBA.InStr(1, trimmedText, ".[", VBA.vbBinaryCompare) > 1 And VBA.Right$(trimmedText, 1) = "]") Then
        private_EvaluateExpressionText = private_GetPlaceholderValue( _
            trimmedText, sectionTypeText, sourceTables, renderVars, loopRows)
        Exit Function
    End If

    If VBA.Left$(trimmedText, 1) <> "{" And VBA.InStr(1, trimmedText, "{", VBA.vbBinaryCompare) <= 0 Then
        placeholderValue = private_GetPlaceholderValue(trimmedText, sectionTypeText, sourceTables, renderVars, loopRows)
        If VBA.Len(placeholderValue) > 0 _
            Or private_TokenExists(trimmedText, sectionTypeText, sourceTables, renderVars, loopRows) Then
            private_EvaluateExpressionText = placeholderValue
        Else
            private_EvaluateExpressionText = trimmedText
        End If
        Exit Function
    End If

    resultText = trimmedText
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = False
    rx.Pattern = "\{([^#][^{}]*)\}"
    Set matches = rx.Execute(resultText)
    If matches Is Nothing Then
        private_EvaluateExpressionText = resultText
        Exit Function
    End If

    For Each matchObj In matches
        placeholderName = VBA.Trim$(VBA.CStr(matchObj.SubMatches(0)))
        placeholderValue = private_GetPlaceholderValue(placeholderName, sectionTypeText, sourceTables, renderVars, loopRows)
        resultText = VBA.Replace(resultText, matchObj.Value, placeholderValue)
    Next matchObj

    private_EvaluateExpressionText = resultText
End Function

Private Function private_TokenExists( _
    ByVal tokenText As String, _
    ByVal sectionTypeText As String, _
    ByVal sourceTables As Collection, _
    ByVal renderVars As Object, _
    ByVal loopRows As Object _
) As Boolean
    Dim sourceTable As obj_TableDynamic
    Dim dotPos As Long
    Dim loopVarName As String
    Dim loopFieldName As String
    Dim dummyValue As String
    Dim columnIndex As Long

    tokenText = VBA.Trim$(tokenText)
    If VBA.Len(tokenText) = 0 Then Exit Function
    If VBA.StrComp(tokenText, "SectionType", VBA.vbTextCompare) = 0 Then
        private_TokenExists = True
        Exit Function
    End If
    If Not renderVars Is Nothing Then
        If renderVars.Exists(tokenText) Then
            private_TokenExists = True
            Exit Function
        End If
    End If

    dotPos = VBA.InStr(1, tokenText, ".", VBA.vbBinaryCompare)
    If dotPos > 1 Then
        loopVarName = VBA.Trim$(VBA.Left$(tokenText, dotPos - 1))
        loopFieldName = VBA.Trim$(VBA.Mid$(tokenText, dotPos + 1))
        If private_IsBracketFieldToken(loopFieldName) Then
            private_TokenExists = private_TryGetLoopFieldValue( _
                loopRows, _
                loopVarName, _
                private_UnwrapBracketFieldToken(loopFieldName), _
                dummyValue)
        ElseIf VBA.Left$(loopFieldName, 2) = "__" Then
            private_TokenExists = private_TryGetLoopFieldValue(loopRows, loopVarName, loopFieldName, dummyValue)
        End If
        Exit Function
    End If

    If Not private_IsBracketFieldToken(tokenText) Then Exit Function

    Set sourceTable = private_TryGetSourceTable(sourceTables, 1)
    If sourceTable Is Nothing Then Exit Function
    tokenText = private_UnwrapBracketFieldToken(tokenText)
    private_TokenExists = sourceTable.TryGetColumnIndexByAlias(tokenText, columnIndex)
    If Not private_TokenExists Then private_TokenExists = sourceTable.TryGetColumnIndexByName(tokenText, columnIndex)
End Function

Private Function private_GetMainTableFieldValue(ByVal fieldName As String, ByVal sourceTables As Collection) As String
    Dim sourceTable As obj_TableDynamic
    Dim sourceRow As obj_Row
    Dim columnIndex As Long

    fieldName = VBA.Trim$(fieldName)
    If VBA.Len(fieldName) = 0 Then Exit Function
    If sourceTables Is Nothing Then Exit Function
    If sourceTables.Count <= 0 Then Exit Function

    Set sourceTable = private_TryGetSourceTable(sourceTables, 1)
    If sourceTable Is Nothing Then Exit Function
    If sourceTable.RowCount <= 0 Then Exit Function

    columnIndex = private_FindSourceColumnIndex(sourceTable, fieldName)
    If columnIndex <= 0 Then Exit Function

    Set sourceRow = sourceTable.Rows.Item(1)
    If sourceRow Is Nothing Then Exit Function

    private_GetMainTableFieldValue = VBA.Trim$(sourceRow.GetCellValue(columnIndex))
End Function

Private Function private_MainPlaceholderHasCellTag( _
    ByVal placeholderName As String, _
    ByVal sourceTables As Collection, _
    ByVal tagName As String _
) As Boolean
    Dim placeholderParts As Collection
    Dim fieldToken As String
    Dim fieldName As String
    Dim sourceTable As obj_TableDynamic
    Dim sourceRow As obj_Row
    Dim cellObj As obj_Cell
    Dim columnIndex As Long

    Set placeholderParts = private_SplitByDelimiterOutsideQuotes( _
        VBA.Trim$(placeholderName), "|")
    If placeholderParts Is Nothing Then Exit Function
    If placeholderParts.Count <= 0 Then Exit Function

    fieldToken = VBA.Trim$(VBA.CStr(placeholderParts.Item(1)))
    If Not private_IsBracketFieldToken(fieldToken) Then Exit Function
    fieldName = private_UnwrapBracketFieldToken(fieldToken)

    Set sourceTable = private_TryGetSourceTable(sourceTables, 1)
    If sourceTable Is Nothing Then Exit Function
    If sourceTable.RowCount <= 0 Then Exit Function

    columnIndex = private_FindSourceColumnIndex(sourceTable, fieldName)
    If columnIndex <= 0 Then Exit Function

    Set sourceRow = sourceTable.Rows.Item(1)
    If sourceRow Is Nothing Then Exit Function
    If Not sourceRow.TryGetCellAt(columnIndex, cellObj) Then Exit Function
    If cellObj Is Nothing Then Exit Function

    private_MainPlaceholderHasCellTag = cellObj.HasTag(tagName)
End Function

Private Function private_IsBracketFieldToken(ByVal tokenText As String) As Boolean
    tokenText = VBA.Trim$(tokenText)
    If VBA.Len(tokenText) < 3 Then Exit Function
    private_IsBracketFieldToken = (VBA.Left$(tokenText, 1) = "[" And VBA.Right$(tokenText, 1) = "]")
End Function

Private Function private_UnwrapBracketFieldToken(ByVal tokenText As String) As String
    tokenText = VBA.Trim$(tokenText)
    If Not private_IsBracketFieldToken(tokenText) Then Exit Function
    private_UnwrapBracketFieldToken = VBA.Trim$(VBA.Mid$(tokenText, 2, VBA.Len(tokenText) - 2))
End Function

Private Function private_IsTruthy(ByVal valueText As String) As Boolean
    valueText = VBA.LCase$(VBA.Trim$(valueText))
    If VBA.Len(valueText) = 0 Then Exit Function
    If valueText = "false" Or valueText = "0" Or valueText = "no" Then Exit Function
    private_IsTruthy = True
End Function

Private Function private_IsQuoted(ByVal valueText As String) As Boolean
    valueText = VBA.Trim$(valueText)
    If VBA.Len(valueText) < 2 Then Exit Function
    private_IsQuoted = (VBA.Left$(valueText, 1) = """" And VBA.Right$(valueText, 1) = """") _
        Or (VBA.Left$(valueText, 1) = "'" And VBA.Right$(valueText, 1) = "'")
End Function

Private Function private_TryEvaluateHelperFunction( _
    ByVal expressionText As String, _
    ByVal sectionTypeText As String, _
    ByVal sourceTables As Collection, _
    ByVal renderVars As Object, _
    ByVal loopRows As Object, _
    ByRef outValue As String _
) As Boolean
    Dim functionName As String
    Dim argsText As String
    Dim args As Collection
    Dim textArg As String
    Dim patternArg As String
    Dim replacementArg As String
    Dim openPos As Long

    outValue = VBA.vbNullString
    expressionText = VBA.Trim$(expressionText)
    If VBA.Left$(expressionText, VBA.Len("$ex_Helpers.")) <> "$ex_Helpers." Then Exit Function
    If VBA.Right$(expressionText, 1) <> ")" Then Exit Function

    ' Минимальная поддержка внешних helper-вызовов для #let/#if.
    ' Сейчас переносим только regex-методы, которые нужны шаблонам:
    ' $ex_Helpers.m_RegexIsMatch("{[ReportPerson]}", "...")
    ' Аргументы сначала проходят через placeholder evaluation, поэтому внутри
    ' строк можно ссылаться на поля DynamicTable.
    openPos = VBA.InStr(1, expressionText, "(", VBA.vbBinaryCompare)
    If openPos <= VBA.Len("$ex_Helpers.") + 1 Then Exit Function

    functionName = VBA.Mid$(expressionText, VBA.Len("$ex_Helpers.") + 1, openPos - VBA.Len("$ex_Helpers.") - 1)
    argsText = VBA.Mid$(expressionText, openPos + 1, VBA.Len(expressionText) - openPos - 1)
    Set args = private_SplitArguments(argsText)

    Select Case VBA.LCase$(VBA.Trim$(functionName))
        Case "m_regexismatch"
            If args.Count <> 2 Then Exit Function
            textArg = private_EvaluateFunctionArgument(VBA.CStr(args.Item(1)), sectionTypeText, sourceTables, renderVars, loopRows)
            patternArg = private_EvaluateFunctionArgument(VBA.CStr(args.Item(2)), sectionTypeText, sourceTables, renderVars, loopRows)
            outValue = VBA.CStr(ex_Helpers.m_RegexIsMatch(textArg, patternArg))
            private_TryEvaluateHelperFunction = True
        Case "m_regexfirstmatch"
            If args.Count <> 2 Then Exit Function
            textArg = private_EvaluateFunctionArgument(VBA.CStr(args.Item(1)), sectionTypeText, sourceTables, renderVars, loopRows)
            patternArg = private_EvaluateFunctionArgument(VBA.CStr(args.Item(2)), sectionTypeText, sourceTables, renderVars, loopRows)
            outValue = ex_Helpers.m_RegexFirstMatch(textArg, patternArg)
            private_TryEvaluateHelperFunction = True
        Case "m_regexgetgroup"
            If args.Count < 2 Or args.Count > 3 Then Exit Function
            textArg = private_EvaluateFunctionArgument(VBA.CStr(args.Item(1)), sectionTypeText, sourceTables, renderVars, loopRows)
            patternArg = private_EvaluateFunctionArgument(VBA.CStr(args.Item(2)), sectionTypeText, sourceTables, renderVars, loopRows)
            If args.Count = 3 Then
                outValue = ex_Helpers.m_RegexGetGroup(textArg, patternArg, VBA.CLng(VBA.Val(VBA.CStr(args.Item(3)))))
            Else
                outValue = ex_Helpers.m_RegexGetGroup(textArg, patternArg)
            End If
            private_TryEvaluateHelperFunction = True
        Case "m_regexreplace"
            If args.Count <> 3 Then Exit Function
            textArg = private_EvaluateFunctionArgument(VBA.CStr(args.Item(1)), sectionTypeText, sourceTables, renderVars, loopRows)
            patternArg = private_EvaluateFunctionArgument(VBA.CStr(args.Item(2)), sectionTypeText, sourceTables, renderVars, loopRows)
            replacementArg = private_EvaluateFunctionArgument(VBA.CStr(args.Item(3)), sectionTypeText, sourceTables, renderVars, loopRows)
            outValue = ex_Helpers.m_RegexReplace(textArg, patternArg, replacementArg)
            private_TryEvaluateHelperFunction = True
    End Select
End Function

Private Function private_EvaluateFunctionArgument( _
    ByVal argumentText As String, _
    ByVal sectionTypeText As String, _
    ByVal sourceTables As Collection, _
    ByVal renderVars As Object, _
    ByVal loopRows As Object _
) As String
    argumentText = VBA.Trim$(argumentText)
    If private_IsQuoted(argumentText) Then argumentText = VBA.Mid$(argumentText, 2, VBA.Len(argumentText) - 2)
    If VBA.InStr(1, argumentText, "{", VBA.vbBinaryCompare) > 0 Then
        private_EvaluateFunctionArgument = private_EvaluateExpressionText(argumentText, sectionTypeText, sourceTables, renderVars, loopRows)
    Else
        private_EvaluateFunctionArgument = argumentText
    End If
End Function

Private Function private_BuildLoopRows(ByVal collectionName As String, ByVal sourceTables As Collection) As Collection
    Dim resultRows As Collection
    Dim tableCollection As Collection
    Dim tableItem As Variant
    Dim tableObj As obj_TableDynamic

    Set resultRows = New Collection
    collectionName = VBA.Trim$(collectionName)
    If VBA.Len(collectionName) = 0 Then Exit Function

    If Not private_TryGetNamedTableCollection(collectionName, tableCollection) Then Exit Function

    For Each tableItem In tableCollection
        Set tableObj = Nothing
        On Error Resume Next
        Set tableObj = tableItem
        On Error GoTo 0
        If Not tableObj Is Nothing Then private_AddTableRowsToLoopCollection tableObj, resultRows
    Next tableItem

    Set private_BuildLoopRows = resultRows
End Function

Private Function private_TryGetNamedTableCollection( _
    ByVal collectionName As String, _
    ByRef outCollection As Collection _
) As Boolean
    Dim normalizedKey As String
    Dim collectionObj As Object

    Set outCollection = Nothing
    If m_NamedCollections Is Nothing Then Exit Function

    collectionName = VBA.Trim$(collectionName)
    If VBA.Len(collectionName) = 0 Then Exit Function

    If m_NamedCollections.Exists(collectionName) Then
        Set collectionObj = m_NamedCollections(collectionName)
    Else
        normalizedKey = private_NormalizeCollectionKey(collectionName)
        If VBA.Len(normalizedKey) > 0 Then
            If m_NamedCollections.Exists(normalizedKey) Then Set collectionObj = m_NamedCollections(normalizedKey)
        End If
    End If

    If collectionObj Is Nothing Then Exit Function
    Set outCollection = collectionObj
    private_TryGetNamedTableCollection = True
End Function

Private Sub private_AddTableRowsToLoopCollection(ByVal tableObj As obj_TableDynamic, ByVal resultRows As Collection)
    Dim rowIndex As Long
    Dim rowCtx As Object

    If tableObj Is Nothing Then Exit Sub
    If resultRows Is Nothing Then Exit Sub

    For rowIndex = 1 To tableObj.RowCount
        Set rowCtx = private_BuildLoopRowContext(tableObj, rowIndex)
        If Not rowCtx Is Nothing Then resultRows.Add rowCtx
    Next rowIndex
End Sub

Private Function private_BuildLoopRowContext(ByVal tableObj As obj_TableDynamic, ByVal rowIndex As Long) As Object
    Dim rowObj As obj_Row
    Dim result As Object
    Dim colIndex As Long
    Dim colObj As obj_Column
    Dim aliasItem As Variant
    Dim valueText As String

    If tableObj Is Nothing Then Exit Function
    If rowIndex <= 0 Or rowIndex > tableObj.RowCount Then Exit Function

    Set rowObj = tableObj.Rows.Item(rowIndex)
    If rowObj Is Nothing Then Exit Function

    Set result = VBA.CreateObject("Scripting.Dictionary")
    result.CompareMode = 1
    result("__sectionTitle") = tableObj.SectionTitle

    For colIndex = 1 To tableObj.ColumnCount
        Set colObj = tableObj.Columns.Item(colIndex)
        If colObj Is Nothing Then GoTo ContinueColumn
        valueText = VBA.Trim$(rowObj.GetCellValue(colIndex))
        If VBA.Len(VBA.Trim$(colObj.Name)) > 0 Then result(VBA.Trim$(colObj.Name)) = valueText
        For Each aliasItem In colObj.Aliases
            If VBA.Len(VBA.Trim$(VBA.CStr(aliasItem))) > 0 Then result(VBA.Trim$(VBA.CStr(aliasItem))) = valueText
        Next aliasItem

ContinueColumn:
    Next colIndex

    Set private_BuildLoopRowContext = result
End Function

Private Function private_TryGetLoopFieldValue( _
    ByVal loopRows As Object, _
    ByVal loopVarName As String, _
    ByVal loopFieldName As String, _
    ByRef outValue As String _
) As Boolean
    Dim rowCtx As Object

    outValue = VBA.vbNullString
    If loopRows Is Nothing Then Exit Function
    loopVarName = VBA.Trim$(loopVarName)
    loopFieldName = VBA.Trim$(loopFieldName)
    If VBA.Len(loopVarName) = 0 Or VBA.Len(loopFieldName) = 0 Then Exit Function
    If Not loopRows.Exists(loopVarName) Then Exit Function

    Set rowCtx = loopRows(loopVarName)
    If rowCtx Is Nothing Then Exit Function
    If Not rowCtx.Exists(loopFieldName) Then Exit Function

    outValue = VBA.Trim$(VBA.CStr(rowCtx(loopFieldName)))
    private_TryGetLoopFieldValue = True
End Function

Private Function private_TryGetSourceTable(ByVal sourceTables As Collection, ByVal tableIndex As Long) As obj_TableDynamic
    Dim tableObj As Object

    If sourceTables Is Nothing Then Exit Function
    If tableIndex <= 0 Or tableIndex > sourceTables.Count Then Exit Function

    On Error Resume Next
    Set tableObj = sourceTables.Item(tableIndex)
    Set private_TryGetSourceTable = tableObj
    On Error GoTo 0
End Function

Private Function private_NormalizeCollectionKey(ByVal valueText As String) As String
    valueText = VBA.LCase$(VBA.Trim$(valueText))
    valueText = VBA.Replace(valueText, " ", VBA.vbNullString)
    valueText = VBA.Replace(valueText, ":", VBA.vbNullString)
    valueText = VBA.Replace(valueText, "-", VBA.vbNullString)
    valueText = VBA.Replace(valueText, "_", VBA.vbNullString)
    private_NormalizeCollectionKey = valueText
End Function

Private Function private_ApplyFormatter(ByVal valueText As String, ByVal formatterText As String) As String
    Dim formatterName As String
    Dim argsText As String
    Dim colonPos As Long
    Dim args As Collection
    Dim dateValue As Date
    Dim offsetDays As Long
    Dim formatPattern As String

    formatterText = VBA.Trim$(formatterText)
    formatterName = formatterText
    colonPos = VBA.InStr(1, formatterText, ":", VBA.vbBinaryCompare)
    If colonPos > 0 Then
        formatterName = VBA.Trim$(VBA.Left$(formatterText, colonPos - 1))
        argsText = VBA.Trim$(VBA.Mid$(formatterText, colonPos + 1))
    End If

    Select Case VBA.LCase$(formatterName)
        Case VBA.LCase$(FORMATTER_UPPER_FIRST_LETTER)
            private_ApplyFormatter = private_UpperFirstLetter(valueText)
        Case VBA.LCase$(FORMATTER_LOWER_FIRST_LETTER)
            private_ApplyFormatter = private_LowerFirstLetter(valueText)
        Case VBA.LCase$(FORMATTER_REGEX_REPLACE)
            ' Формат: |regexreplace:"pattern","replacement".
            ' Аргументы разбираются с учетом кавычек, поэтому запятые внутри
            ' regex/replacement не режут список параметров.
            Set args = private_SplitArguments(argsText)
            If Not args Is Nothing Then
                If args.Count = 2 Then
                    private_ApplyFormatter = ex_Helpers.m_RegexReplace( _
                        valueText, _
                        private_UnquoteFormatterArgument(VBA.CStr(args.Item(1))), _
                        private_UnquoteFormatterArgument(VBA.CStr(args.Item(2))))
                    Exit Function
                End If
            End If
            private_ApplyFormatter = valueText
        Case VBA.LCase$(FORMATTER_DATE_OFFSET)
            ' Формат: |dateoffset:"+1".
            ' Возвращаем дату в стабильном виде dd.mm.yyyy, чтобы следующий
            ' formatter dateformat мог независимо выбрать отображение.
            If Not private_TryParseFormatterDate(valueText, dateValue) Then
                private_ApplyFormatter = valueText
                Exit Function
            End If
            argsText = private_UnquoteFormatterArgument(argsText)
            If VBA.Len(argsText) = 0 Or Not VBA.IsNumeric(argsText) Then
                private_ApplyFormatter = valueText
                Exit Function
            End If
            offsetDays = VBA.CLng(argsText)
            private_ApplyFormatter = VBA.Format$(VBA.DateAdd("d", offsetDays, dateValue), FORMATTER_DATE_STORAGE_FORMAT)
        Case VBA.LCase$(FORMATTER_DATE_FORMAT)
            ' Формат: |dateformat:"\dd \month \yyyy року".
            ' Отображение даты остается ответственностью шаблона, а не WORD exporter-а.
            If Not private_TryParseFormatterDate(valueText, dateValue) Then
                private_ApplyFormatter = valueText
                Exit Function
            End If
            formatPattern = private_UnquoteFormatterArgument(argsText)
            private_ApplyFormatter = ex_Helpers.fn_FormatUaDatePattern(dateValue, formatPattern)
        Case Else
            private_ApplyFormatter = valueText
    End Select
End Function

Private Function private_TryParseFormatterDate(ByVal valueText As String, ByRef outDate As Date) As Boolean
    valueText = VBA.Trim$(valueText)
    If VBA.Len(valueText) = 0 Then Exit Function

    ' Formatter stage получает уже resolved date из WORD exporter-а. Контекст
    ' 01.01.1900 здесь нужен только как безопасный base для полностью заданных
    ' дат или sentinel-дат, а не для бизнес-резолва короткой даты.
    private_TryParseFormatterDate = ex_Helpers.fn_TryResolveDateWithContext( _
        valueText, _
        VBA.DateSerial(1900, 1, 1), _
        outDate)
End Function

Private Function private_SplitArguments(ByVal argsText As String) As Collection
    Set private_SplitArguments = private_SplitByDelimiterOutsideQuotes(argsText, ",")
End Function

Private Function private_SplitByDelimiterOutsideQuotes( _
    ByVal sourceText As String, _
    ByVal delimiterText As String _
) As Collection
    Dim result As Collection
    Dim i As Long
    Dim ch As String
    Dim currentPart As String
    Dim quoteChar As String
    Dim inQuote As Boolean

    Set result = New Collection
    ' Универсальный splitter для DSL-строк: нужен и для formatter pipeline,
    ' и для аргументов helper-вызовов. Главное правило - delimiter внутри
    ' кавычек является частью значения, а не разделителем.
    For i = 1 To VBA.Len(sourceText)
        ch = VBA.Mid$(sourceText, i, 1)
        If inQuote Then
            currentPart = currentPart & ch
            If ch = quoteChar Then inQuote = False
        Else
            If ch = """" Or ch = "'" Then
                inQuote = True
                quoteChar = ch
                currentPart = currentPart & ch
            ElseIf ch = delimiterText Then
                result.Add VBA.Trim$(currentPart)
                currentPart = VBA.vbNullString
            Else
                currentPart = currentPart & ch
            End If
        End If
    Next i

    If VBA.Len(VBA.Trim$(currentPart)) > 0 Or VBA.Len(sourceText) > 0 Then result.Add VBA.Trim$(currentPart)
    Set private_SplitByDelimiterOutsideQuotes = result
End Function

Private Function private_UnquoteFormatterArgument(ByVal valueText As String) As String
    valueText = VBA.Trim$(valueText)
    If private_IsQuoted(valueText) Then
        private_UnquoteFormatterArgument = VBA.Mid$(valueText, 2, VBA.Len(valueText) - 2)
    Else
        private_UnquoteFormatterArgument = valueText
    End If
End Function

Private Function private_UpperFirstLetter(ByVal valueText As String) As String
    private_UpperFirstLetter = private_ChangeFirstVisibleLetterCase( _
        valueText, True)
End Function

Private Function private_LowerFirstLetter(ByVal valueText As String) As String
    private_LowerFirstLetter = private_ChangeFirstVisibleLetterCase( _
        valueText, False)
End Function

Private Function private_ChangeFirstVisibleLetterCase( _
    ByVal valueText As String, _
    ByVal makeUpper As Boolean _
) As String
    Dim visibleStart As Long
    Dim markerEnd As Long
    Dim firstCharacter As String

    private_ChangeFirstVisibleLetterCase = valueText
    If VBA.Len(valueText) = 0 Then Exit Function

    visibleStart = 1
    ' Preview color является служебной обёрткой значения. Formatter должен
    ' менять регистр первого символа текста, а не первой скобки маркера.
    If VBA.StrComp( _
        VBA.Left$(valueText, VBA.Len("[[color=")), _
        "[[color=", VBA.vbTextCompare) = 0 Then
        markerEnd = VBA.InStr(1, valueText, "]]", VBA.vbBinaryCompare)
        If markerEnd > 0 Then visibleStart = markerEnd + VBA.Len("]]")
    End If
    If visibleStart > VBA.Len(valueText) Then Exit Function

    firstCharacter = VBA.Mid$(valueText, visibleStart, 1)
    If makeUpper Then
        firstCharacter = VBA.UCase$(firstCharacter)
    Else
        firstCharacter = VBA.LCase$(firstCharacter)
    End If
    private_ChangeFirstVisibleLetterCase = _
        VBA.Left$(valueText, visibleStart - 1) & _
        firstCharacter & VBA.Mid$(valueText, visibleStart + 1)
End Function

Private Function private_FindSourceColumnIndex(ByVal sourceTable As obj_TableDynamic, ByVal columnName As String) As Long
    If sourceTable Is Nothing Then Exit Function
    columnName = VBA.Trim$(VBA.CStr(columnName))
    If VBA.Len(columnName) = 0 Then Exit Function

    If sourceTable.TryGetColumnIndexByAlias(columnName, private_FindSourceColumnIndex) Then Exit Function
    If sourceTable.TryGetColumnIndexByName(columnName, private_FindSourceColumnIndex) Then Exit Function
End Function

Private Function private_ApplyLayoutTokens(ByVal valueText As String) As String
    ' #^ склеивает строки: удаляем marker, ближайший перенос строки после него
    ' и отступ в начале следующей строки. Это удобно для шаблонов, где строка
    ' в XML разбита ради читаемости, но в итоговом тексте должна продолжаться.
    valueText = private_RegexReplace(valueText, "(\{#\^\}|#\^)[ \t]*(\r\n|\r|\n)[ \t]*", VBA.vbNullString)
    valueText = VBA.Replace(valueText, TOKEN_JOIN_LINE_BRACED, VBA.vbNullString)
    valueText = VBA.Replace(valueText, TOKEN_JOIN_LINE, VBA.vbNullString)

    ' #_ убирает отступ после marker до первого видимого символа. Перенос строки
    ' marker не трогает: для склейки строк используем #^.
    valueText = private_RegexReplace(valueText, "(\{#_\}|#_)[ \t]*", VBA.vbNullString)
    valueText = VBA.Replace(valueText, TOKEN_TRIM_INDENT_BRACED, VBA.vbNullString)
    valueText = VBA.Replace(valueText, TOKEN_TRIM_INDENT, VBA.vbNullString)

    private_ApplyLayoutTokens = valueText
End Function

Private Function private_NormalizeRenderedText(ByVal valueText As String) As String
    valueText = private_CollapseSpacesPreservingLineBreaks(valueText)
    private_NormalizeRenderedText = VBA.Trim$(valueText)
End Function

Private Function private_CollapseSpacesPreservingLineBreaks(ByVal valueText As String) As String
    Dim lines() As String
    Dim i As Long

    valueText = VBA.Replace(valueText, VBA.vbCrLf, VBA.vbLf)
    valueText = VBA.Replace(valueText, VBA.vbCr, VBA.vbLf)
    lines = VBA.Split(valueText, VBA.vbLf)

    For i = LBound(lines) To UBound(lines)
        lines(i) = private_CollapseSpaces(VBA.CStr(lines(i)))
    Next i

    private_CollapseSpacesPreservingLineBreaks = VBA.Join(lines, VBA.vbCrLf)
End Function

Private Function private_CollapseSpaces(ByVal valueText As String) As String
    valueText = VBA.Trim$(valueText)
    Do While VBA.InStr(1, valueText, "  ", VBA.vbBinaryCompare) > 0
        valueText = VBA.Replace(valueText, "  ", " ")
    Loop
    private_CollapseSpaces = valueText
End Function

Private Function private_RegexReplace( _
    ByVal sourceText As String, _
    ByVal patternText As String, _
    ByVal replacementText As String _
) As String
    Dim rx As Object

    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = False
    rx.MultiLine = True
    rx.Pattern = patternText

    private_RegexReplace = rx.Replace(sourceText, replacementText)
End Function
