VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PEB_WordResultTplParser"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = False
#Const LOGGING_VERBOSE_ENABLED = False

Private Const PROFILES_NS As String = "urn:excelprototype:profiles"
Private Const TOKEN_NEWLINE As String = "{#newline}"
Private Const TOKEN_JOIN_LINE As String = "#^"
Private Const TOKEN_JOIN_LINE_BRACED As String = "{#^}"
Private Const TOKEN_TRIM_INDENT As String = "#_"
Private Const TOKEN_TRIM_INDENT_BRACED As String = "{#_}"
Private Const TOKEN_IF_OPEN As String = "{#if"
Private Const TOKEN_IF_CLOSE As String = "{#endif}"
Private Const TOKEN_FOR_OPEN As String = "{#for"
Private Const TOKEN_FOR_CLOSE As String = "{#endfor}"
Private Const TOKEN_INCLUDE_OPEN As String = "{#include"
Private Const TOKEN_LET As String = "#let"
Private Const FORMATTER_UPPER_FIRST_LETTER As String = "upperFirstLetter"
Private Const FORMATTER_LOWER_FIRST_LETTER As String = "lowerFirstLetter"

Private m_TemplateRelPath As String
Private m_IsDisposed As Boolean

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
    Initialize = (VBA.Len(m_TemplateRelPath) > 0)
End Function

Public Sub Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    m_TemplateRelPath = VBA.vbNullString
End Sub

Public Function TryRenderForSectionType( _
    ByVal sectionTypeText As String, _
    ByVal sourceTables As Collection, _
    ByRef outResultText As String _
) As Boolean
    Dim templateText As String
    Dim renderVars As Object
    Dim loopRows As Object

    outResultText = VBA.vbNullString
    If m_IsDisposed Then Exit Function
    If sourceTables Is Nothing Then Exit Function
    If sourceTables.Count <= 0 Then Exit Function

    If Not private_TryGetTemplateTextBySectionType(sectionTypeText, templateText) Then Exit Function

    Set renderVars = VBA.CreateObject("Scripting.Dictionary")
    renderVars.CompareMode = 1
    Set loopRows = VBA.CreateObject("Scripting.Dictionary")
    loopRows.CompareMode = 1

    outResultText = private_RenderTemplate(templateText, sectionTypeText, sourceTables, renderVars, loopRows)

    TryRenderForSectionType = True
End Function

' //
' // Internal
' //
Private Function private_TryGetTemplateTextBySectionType( _
    ByVal sectionTypeText As String, _
    ByRef outTemplateText As String _
) As Boolean
    Dim doc As Object
    Dim node As Object
    Dim xpath As String
    Dim includeChain As Collection

    outTemplateText = VBA.vbNullString
    sectionTypeText = VBA.Trim$(sectionTypeText)
    If VBA.Len(sectionTypeText) = 0 Then
        VBA.MsgBox "PrototypeNew: WORD template section type is empty.", VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    Set doc = ex_XmlCore.fn_LoadDomByRelativePath( _
        ThisWorkbook, _
        m_TemplateRelPath, _
        "Missing WORD result templates file: ", _
        "Failed to parse WORD result templates file: ", _
        PROFILES_NS)
    If doc Is Nothing Then
        VBA.MsgBox "PrototypeNew: failed to load WORD result templates: " & m_TemplateRelPath, VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    xpath = "/p:wordResultTemplates/p:template[@sectionType=" & ex_XmlCore.fn_XPathLiteral(sectionTypeText) & "]/p:text"
    Set node = doc.selectSingleNode(xpath)
    If node Is Nothing Then
        VBA.MsgBox "PrototypeNew: WORD result template was not found for section type: " & sectionTypeText, VBA.vbExclamation, "PrototypeNew / WORD export"
        Exit Function
    End If

    Set includeChain = New Collection
    outTemplateText = private_ExpandSharedTemplateIncludes(VBA.CStr(node.Text), doc, includeChain)
    private_TryGetTemplateTextBySectionType = True
End Function

Private Function private_RenderTemplate( _
    ByVal templateText As String, _
    ByVal sectionTypeText As String, _
    ByVal sourceTables As Collection, _
    ByVal renderVars As Object, _
    ByVal loopRows As Object _
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
    '   {#if expr}...{#endif}
    '   {#for item in Collection}...{#endfor}
    '   #^ / {#^}        -> удалить token и следующий перенос строки
    '   #_ / {#_}        -> удалить token и пробелы/табуляцию до первого символа
    '   {#newline}       -> перенос строки
    '   {SectionType}    -> тип секции из export context/source table
    '   {Column|formatter} -> значение по alias/заголовку + formatter pipeline
    '
    ' Это все еще не полный PersonalCard parser: внешние VBA-вызовы, падежи,
    ' даты, colors и морфология намеренно не перенесены в этот PEB preview.
    resultText = VBA.CStr(templateText)
    resultText = private_ResolveLetBindings(resultText, sectionTypeText, sourceTables, renderVars, loopRows)
    resultText = private_RenderForBlocks(resultText, sectionTypeText, sourceTables, renderVars, loopRows)
    resultText = private_RenderIfBlocks(resultText, sectionTypeText, sourceTables, renderVars, loopRows)
    resultText = private_ApplyLayoutTokens(resultText)
    resultText = VBA.Replace(resultText, TOKEN_NEWLINE, VBA.vbCrLf)

    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = False
    ' Берем только обычные placeholders. Служебные токены {#...} сюда не попадают.
    rx.Pattern = "\{([^#][^{}]*)\}"

    Set matches = rx.Execute(resultText)
    If matches Is Nothing Then
        private_RenderTemplate = resultText
        Exit Function
    End If

    For matchIndex = matches.Count - 1 To 0 Step -1
        Set matchObj = matches.Item(matchIndex)
        placeholderName = VBA.Trim$(VBA.CStr(matchObj.SubMatches(0)))
        placeholderValue = private_GetPlaceholderValue(placeholderName, sectionTypeText, sourceTables, renderVars, loopRows)
        resultText = VBA.Left$(resultText, matchObj.FirstIndex) & _
            placeholderValue & _
            VBA.Mid$(resultText, matchObj.FirstIndex + matchObj.Length + 1)
    Next matchIndex

    private_RenderTemplate = private_NormalizeRenderedText(resultText)
End Function

Private Function private_GetPlaceholderValue( _
    ByVal placeholderName As String, _
    ByVal sectionTypeText As String, _
    ByVal sourceTables As Collection, _
    ByVal renderVars As Object, _
    ByVal loopRows As Object _
) As String
    Dim placeholderParts() As String
    Dim valueName As String
    Dim resultValue As String
    Dim formatterIndex As Long

    placeholderParts = VBA.Split(VBA.Trim$(placeholderName), "|")
    valueName = VBA.Trim$(VBA.CStr(placeholderParts(LBound(placeholderParts))))
    resultValue = private_GetRawPlaceholderValue(valueName, sectionTypeText, sourceTables, renderVars, loopRows)

    For formatterIndex = LBound(placeholderParts) + 1 To UBound(placeholderParts)
        resultValue = private_ApplyFormatter(resultValue, VBA.Trim$(VBA.CStr(placeholderParts(formatterIndex))))
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
    Dim sourceTable As obj_TableDynamic
    Dim sourceRow As obj_Row
    Dim columnIndex As Long
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
        If private_TryGetLoopFieldValue(loopRows, loopVarName, loopFieldName, private_GetRawPlaceholderValue) Then Exit Function
    End If

    If sourceTables Is Nothing Then Exit Function
    If sourceTables.Count <= 0 Then Exit Function
    Set sourceTable = private_TryGetSourceTable(sourceTables, 1)
    If sourceTable Is Nothing Then Exit Function

    If sourceTable.RowCount <= 0 Then Exit Function
    columnIndex = private_FindSourceColumnIndex(sourceTable, placeholderName)
    If columnIndex <= 0 Then Exit Function

    Set sourceRow = sourceTable.Rows.Item(1)
    If sourceRow Is Nothing Then Exit Function
    private_GetRawPlaceholderValue = VBA.Trim$(sourceRow.GetCellValue(columnIndex))
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
        renderedPart = renderedPart & private_RenderTemplate(bodyText, sectionTypeText, sourceTables, renderVars, loopRows)
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
    Dim headerText As String
    Dim conditionText As String
    Dim bodyText As String
    Dim renderedText As String

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
        bodyText = VBA.Mid$(resultText, openEndPos + 1, closePos - openEndPos - 1)

        If private_EvaluateCondition(conditionText, sectionTypeText, sourceTables, renderVars, loopRows) Then
            renderedText = private_RenderTemplate(bodyText, sectionTypeText, sourceTables, renderVars, loopRows)
        Else
            renderedText = VBA.vbNullString
        End If

        resultText = VBA.Left$(resultText, openPos - 1) & _
            renderedText & _
            VBA.Mid$(resultText, closePos + VBA.Len(TOKEN_IF_CLOSE))
    Loop

    If VBA.InStr(1, resultText, TOKEN_IF_OPEN, VBA.vbTextCompare) > 0 _
        Or VBA.InStr(1, resultText, TOKEN_IF_CLOSE, VBA.vbTextCompare) > 0 Then
        VBA.MsgBox "PrototypeNew: invalid WORD #if block. Use {#if expression}...{#endif}.", VBA.vbExclamation, "PrototypeNew / WORD export"
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

    trimmedText = VBA.Trim$(expressionText)
    If VBA.Len(trimmedText) = 0 Then Exit Function
    If private_IsQuoted(trimmedText) Then
        private_EvaluateExpressionText = VBA.Mid$(trimmedText, 2, VBA.Len(trimmedText) - 2)
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
        private_TokenExists = private_TryGetLoopFieldValue(loopRows, loopVarName, loopFieldName, dummyValue)
        Exit Function
    End If

    Set sourceTable = private_TryGetSourceTable(sourceTables, 1)
    If sourceTable Is Nothing Then Exit Function
    private_TokenExists = sourceTable.TryGetColumnIndexByAlias(tokenText, columnIndex)
    If Not private_TokenExists Then private_TokenExists = sourceTable.TryGetColumnIndexByName(tokenText, columnIndex)
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

Private Function private_BuildLoopRows(ByVal collectionName As String, ByVal sourceTables As Collection) As Collection
    Dim resultRows As Collection
    Dim tableIndex As Long
    Dim tableObj As obj_TableDynamic

    Set resultRows = New Collection
    collectionName = VBA.Trim$(collectionName)
    If VBA.Len(collectionName) = 0 Then Exit Function
    If sourceTables Is Nothing Then Exit Function

    If VBA.StrComp(collectionName, "MainTable", VBA.vbTextCompare) = 0 Then
        Set tableObj = private_TryGetSourceTable(sourceTables, 1)
        If Not tableObj Is Nothing Then private_AddTableRowsToLoopCollection tableObj, resultRows
        Set private_BuildLoopRows = resultRows
        Exit Function
    End If

    If VBA.StrComp(collectionName, "MetaTables", VBA.vbTextCompare) = 0 Then
        For tableIndex = 2 To sourceTables.Count
            Set tableObj = private_TryGetSourceTable(sourceTables, tableIndex)
            If Not tableObj Is Nothing Then private_AddTableRowsToLoopCollection tableObj, resultRows
        Next tableIndex
        Set private_BuildLoopRows = resultRows
        Exit Function
    End If

    For tableIndex = 1 To sourceTables.Count
        Set tableObj = private_TryGetSourceTable(sourceTables, tableIndex)
        If tableObj Is Nothing Then GoTo ContinueTable
        If VBA.StrComp(VBA.Trim$(tableObj.SectionTitle), collectionName, VBA.vbTextCompare) = 0 _
            Or VBA.StrComp(private_NormalizeCollectionKey(tableObj.SectionTitle), private_NormalizeCollectionKey(collectionName), VBA.vbTextCompare) = 0 Then
            private_AddTableRowsToLoopCollection tableObj, resultRows
        End If

ContinueTable:
    Next tableIndex

    Set private_BuildLoopRows = resultRows
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
    formatterText = VBA.Trim$(formatterText)
    Select Case VBA.LCase$(formatterText)
        Case VBA.LCase$(FORMATTER_UPPER_FIRST_LETTER)
            private_ApplyFormatter = private_UpperFirstLetter(valueText)
        Case VBA.LCase$(FORMATTER_LOWER_FIRST_LETTER)
            private_ApplyFormatter = private_LowerFirstLetter(valueText)
        Case Else
            private_ApplyFormatter = valueText
    End Select
End Function

Private Function private_UpperFirstLetter(ByVal valueText As String) As String
    If VBA.Len(valueText) = 0 Then Exit Function
    private_UpperFirstLetter = VBA.UCase$(VBA.Left$(valueText, 1)) & VBA.Mid$(valueText, 2)
End Function

Private Function private_LowerFirstLetter(ByVal valueText As String) As String
    If VBA.Len(valueText) = 0 Then Exit Function
    private_LowerFirstLetter = VBA.LCase$(VBA.Left$(valueText, 1)) & VBA.Mid$(valueText, 2)
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
