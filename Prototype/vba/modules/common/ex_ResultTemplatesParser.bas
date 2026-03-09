Attribute VB_Name = "ex_ResultTemplatesParser"
Option Explicit

Private Const PROFILES_NS As String = "urn:excelprototype:profiles"
Private Const RESULT_TEMPLATES_REL_PATH As String = "config\modes\PersonalCard\PersonalCardResultTemplates.xml"

' Reserved keywords/tokens supported by this parser.
' Reserved date tokens:
' - {#dd}, {#dd+N}, {#dd-N}
' Reserved line-join token:
' - {#_}  (ignores a line break around token)
' - #_    (same behavior, shorthand)
Private Const DATE_TOKEN_PATTERN As String = "\{#dd(?:[+-]\d+)?\}"
Private Const RESERVED_TOKEN_PREFIX As String = "#"
Private Const RESERVED_DATE_DAY_KEYWORD As String = "dd"
Private Const RESERVED_DATE_OUTPUT_FORMAT As String = "dd"
Private Const RESERVED_JOINLINE_TOKEN As String = "{#_}"
Private Const RESERVED_JOINLINE_TOKEN_SHORT As String = "#_"
Private Const IF_BLOCK_OPEN As String = "{#if"
Private Const IF_BLOCK_CLOSE As String = "{#endif}"
Private Const BOOLEAN_TRUE As String = "true"
Private Const BOOLEAN_FALSE As String = "false"

Private Const FORMATTER_UPPER As String = "upper"
Private Const FORMATTER_LOWER As String = "lower"
Private Const FORMATTER_CAPITALIZE As String = "capitalize"
Private Const FORMATTER_FIRSTCHAR As String = "firstchar"
Private Const FORMATTER_UPPERFIRSTWORD As String = "upperfirstword"
Private Const FORMATTER_UPPERFIRSTLETTER As String = "upperfirstletter"
Private Const FORMATTER_LOWERFIRSTWORD As String = "lowerfirstword"
Private Const FORMATTER_LOWERFIRSTLETTER As String = "lowerfirstletter"
Private Const FORMATTER_GENITIVE As String = "genitive"
Private Const FORMATTER_ACCUSATIVE As String = "accusative"
Private Const FORMATTER_DATIVE As String = "dative"
Private Const FORMATTER_TRUNCATE As String = "truncate"
Private Const FORMATTER_REPLACE As String = "replace"

Private Const CASE_GENITIVE As String = "genitive"
Private Const CASE_ACCUSATIVE As String = "accusative"
Private Const CASE_DATIVE As String = "dative"

Private Const NBSP_CODE_POINT As Long = 160
Private Const NARROW_NBSP_CODE_POINT As Long = 8239
Private Const TEMPLATE_ERROR_PREFIX As String = "[TEMPLATE ERROR]"

Public Function m_GetTemplateText(ByVal templateId As String) As String
    Dim doc As Object
    Dim node As Object
    Dim xpath As String
    Dim templateText As String

    On Error GoTo EH

    templateId = mp_TrimWhitespace(templateId)
    If Len(templateId) = 0 Then
        Err.Raise vbObjectError + 1760, "ex_ResultTemplatesParser", "Template id is empty."
    End If

    Set doc = ex_XmlCore.m_LoadDomByRelativePath( _
        ThisWorkbook, _
        RESULT_TEMPLATES_REL_PATH, _
        PROFILES_NS, _
        "Missing result templates file: ", _
        "Failed to parse result templates file: " _
    )
    If doc Is Nothing Then
        Err.Raise vbObjectError + 1761, "ex_ResultTemplatesParser", "Unable to load result templates xml."
    End If

    xpath = "/p:resultTemplates/p:template[@id=" & ex_XmlCore.m_XPathLiteral(templateId) & "]/p:text"
    Set node = doc.selectSingleNode(xpath)
    If node Is Nothing Then
        Err.Raise vbObjectError + 1762, "ex_ResultTemplatesParser", "Template not found: '" & templateId & "'."
    End If

    templateText = CStr(node.Text)
    m_GetTemplateText = mp_NormalizeTemplateText(templateText)
    Exit Function

EH:
    m_GetTemplateText = mp_PrependTemplateError(vbNullString, "m_GetTemplateText('" & templateId & "')")
End Function

Public Function m_ReplaceToken( _
    ByVal sourceText As String, _
    ByVal tokenText As String, _
    ByVal replacementText As String _
) As String
    On Error GoTo EH
    m_ReplaceToken = Replace(CStr(sourceText), CStr(tokenText), CStr(replacementText))
    Exit Function

EH:
    m_ReplaceToken = mp_PrependTemplateError(CStr(sourceText), "m_ReplaceToken('" & CStr(tokenText) & "')")
End Function

Public Function m_ReplacePlaceholder( _
    ByVal sourceText As String, _
    ByVal placeholderName As String, _
    ByVal replacementText As String _
) As String
    Dim normalizedName As String
    Dim placeholderToken As String
    Dim resultText As String
    Dim rx As Object
    Dim matches As Object
    Dim i As Long
    Dim formatter As String
    Dim formattedValue As String
    Dim matchStart As Long
    Dim matchLen As Long

    On Error GoTo EH

    normalizedName = mp_TrimWhitespace(placeholderName)
    If Len(normalizedName) = 0 Then
        m_ReplacePlaceholder = CStr(sourceText)
        Exit Function
    End If

    placeholderToken = "{" & normalizedName & "}"
    resultText = m_ReplaceToken(CStr(sourceText), placeholderToken, CStr(replacementText))

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = False
    rx.Pattern = "\{" & mp_EscapeRegex(normalizedName) & "\|([^{}]+)\}"

    Set matches = rx.Execute(resultText)
    If Not matches Is Nothing Then
        For i = matches.Count - 1 To 0 Step -1
            formatter = CStr(matches(i).SubMatches(0))
            formattedValue = mp_ApplyFormatterPipeline(CStr(replacementText), formatter)
            matchStart = CLng(matches(i).FirstIndex)
            matchLen = CLng(matches(i).Length)

            resultText = Left$(resultText, matchStart) & formattedValue & Mid$(resultText, matchStart + matchLen + 1)
        Next i
    End If

    resultText = mp_ReplaceIfConditionForPlaceholder(resultText, normalizedName, CStr(replacementText))
    m_ReplacePlaceholder = resultText
    Exit Function

EH:
    If Len(resultText) = 0 Then resultText = CStr(sourceText)
    m_ReplacePlaceholder = mp_PrependTemplateError(resultText, "m_ReplacePlaceholder('" & CStr(placeholderName) & "')")
End Function

Public Function m_ResolveTemplate( _
    ByVal sourceText As String, _
    Optional ByVal baseDateText As String = vbNullString _
) As String
    On Error GoTo EH
    ' Final pass for template text.
    m_ResolveTemplate = mp_ResolveConditionalBlocks(CStr(sourceText))
    m_ResolveTemplate = mp_ResolveJoinLineTokens(m_ResolveTemplate)
    m_ResolveTemplate = mp_ResolveDateExpressions(m_ResolveTemplate, baseDateText)
    Exit Function

EH:
    m_ResolveTemplate = mp_PrependTemplateError(CStr(sourceText), "m_ResolveTemplate")
End Function

Public Function m_FormatValue( _
    ByVal sourceValue As String, _
    ByVal formatterName As String _
) As String
    Dim normalizedFormatter As String

    On Error GoTo EH

    normalizedFormatter = mp_TrimWhitespace(formatterName)
    If Len(normalizedFormatter) = 0 Then
        m_FormatValue = CStr(sourceValue)
        Exit Function
    End If

    m_FormatValue = mp_ApplyFormatterPipeline(CStr(sourceValue), normalizedFormatter)
    Exit Function

EH:
    m_FormatValue = mp_PrependTemplateError(CStr(sourceValue), "m_FormatValue('" & CStr(formatterName) & "')")
End Function

Private Function mp_CapitalizeText(ByVal textValue As String) As String
    textValue = CStr(textValue)
    If Len(textValue) = 0 Then Exit Function
    If Len(textValue) = 1 Then
        mp_CapitalizeText = UCase$(textValue)
        Exit Function
    End If

    mp_CapitalizeText = UCase$(Left$(textValue, 1)) & LCase$(Mid$(textValue, 2))
End Function

Private Function mp_ApplyFormatter(ByVal sourceValue As String, ByVal formatterName As String) As String
    Dim normalizedFormatter As String

    normalizedFormatter = LCase$(mp_TrimWhitespace(formatterName))

    Select Case normalizedFormatter
        Case FORMATTER_UPPER
            mp_ApplyFormatter = UCase$(CStr(sourceValue))
        Case FORMATTER_LOWER
            mp_ApplyFormatter = LCase$(CStr(sourceValue))
        Case FORMATTER_CAPITALIZE
            mp_ApplyFormatter = mp_CapitalizeText(CStr(sourceValue))
        Case FORMATTER_FIRSTCHAR
            mp_ApplyFormatter = mp_FirstNonSpaceChar(CStr(sourceValue))
        Case FORMATTER_UPPERFIRSTLETTER
            mp_ApplyFormatter = mp_UppercaseFirstLetter(CStr(sourceValue))
        Case FORMATTER_UPPERFIRSTWORD
            mp_ApplyFormatter = mp_UppercaseFirstWord(CStr(sourceValue))
        Case FORMATTER_LOWERFIRSTLETTER
            mp_ApplyFormatter = ex_MorphUaLite.m_LowercaseFirstLetter(CStr(sourceValue))
        Case FORMATTER_LOWERFIRSTWORD
            mp_ApplyFormatter = mp_LowercaseFirstWord(CStr(sourceValue))
        Case FORMATTER_GENITIVE
            mp_ApplyFormatter = mp_InflectPhraseToCase(CStr(sourceValue), CASE_GENITIVE)
        Case FORMATTER_ACCUSATIVE
            mp_ApplyFormatter = mp_InflectPhraseToCase(CStr(sourceValue), CASE_ACCUSATIVE)
        Case FORMATTER_DATIVE
            mp_ApplyFormatter = mp_InflectPhraseToCase(CStr(sourceValue), CASE_DATIVE)
        Case Else
            Err.Raise vbObjectError + 1766, "ex_ResultTemplatesParser", _
                "Unsupported formatter '" & formatterName & "'."
    End Select
End Function

Private Function mp_UppercaseFirstLetter(ByVal sourceValue As String) As String
    Dim textValue As String

    textValue = CStr(sourceValue)
    If Len(textValue) = 0 Then
        mp_UppercaseFirstLetter = textValue
        Exit Function
    End If

    mp_UppercaseFirstLetter = UCase$(Left$(textValue, 1)) & Mid$(textValue, 2)
End Function

Private Function mp_UppercaseFirstWord(ByVal sourceValue As String) As String
    Dim textValue As String
    Dim wordStart As Long
    Dim wordEnd As Long
    Dim ch As String

    textValue = CStr(sourceValue)
    If Len(textValue) = 0 Then
        mp_UppercaseFirstWord = textValue
        Exit Function
    End If

    wordStart = 1
    Do While wordStart <= Len(textValue)
        ch = Mid$(textValue, wordStart, 1)
        If Not mp_IsWhitespaceChar(ch) Then Exit Do
        wordStart = wordStart + 1
    Loop

    If wordStart > Len(textValue) Then
        mp_UppercaseFirstWord = textValue
        Exit Function
    End If

    wordEnd = wordStart
    Do While wordEnd <= Len(textValue)
        ch = Mid$(textValue, wordEnd, 1)
        If mp_IsWhitespaceChar(ch) Then Exit Do
        wordEnd = wordEnd + 1
    Loop

    mp_UppercaseFirstWord = _
        Left$(textValue, wordStart - 1) & _
        UCase$(Mid$(textValue, wordStart, wordEnd - wordStart)) & _
        Mid$(textValue, wordEnd)
End Function

Private Function mp_LowercaseFirstWord(ByVal sourceValue As String) As String
    Dim textValue As String
    Dim wordStart As Long
    Dim wordEnd As Long
    Dim ch As String

    textValue = CStr(sourceValue)
    If Len(textValue) = 0 Then
        mp_LowercaseFirstWord = textValue
        Exit Function
    End If

    wordStart = 1
    Do While wordStart <= Len(textValue)
        ch = Mid$(textValue, wordStart, 1)
        If Not mp_IsWhitespaceChar(ch) Then Exit Do
        wordStart = wordStart + 1
    Loop

    If wordStart > Len(textValue) Then
        mp_LowercaseFirstWord = textValue
        Exit Function
    End If

    wordEnd = wordStart
    Do While wordEnd <= Len(textValue)
        ch = Mid$(textValue, wordEnd, 1)
        If mp_IsWhitespaceChar(ch) Then Exit Do
        wordEnd = wordEnd + 1
    Loop

    mp_LowercaseFirstWord = _
        Left$(textValue, wordStart - 1) & _
        LCase$(Mid$(textValue, wordStart, wordEnd - wordStart)) & _
        Mid$(textValue, wordEnd)
End Function

Private Function mp_ApplyFormatterPipeline(ByVal sourceValue As String, ByVal formatterPipeline As String) As String
    Dim actions() As String
    Dim i As Long
    Dim actionSpec As String
    Dim formattedValue As String

    formatterPipeline = CStr(formatterPipeline)
    actionSpec = mp_TrimWhitespace(formatterPipeline)
    If Len(actionSpec) = 0 Then
        mp_ApplyFormatterPipeline = CStr(sourceValue)
        Exit Function
    End If

    formattedValue = CStr(sourceValue)
    actions = Split(actionSpec, "|")

    For i = LBound(actions) To UBound(actions)
        actionSpec = mp_TrimWhitespace(CStr(actions(i)))
        If Len(actionSpec) = 0 Then
            Err.Raise vbObjectError + 1770, "ex_ResultTemplatesParser", "Empty formatter action in '" & formatterPipeline & "'."
        End If
        formattedValue = mp_ApplyFormatterAction(formattedValue, actionSpec)
    Next i

    mp_ApplyFormatterPipeline = formattedValue
End Function

Private Function mp_ApplyFormatterAction(ByVal sourceValue As String, ByVal actionSpec As String) As String
    Dim colonPos As Long
    Dim actionName As String
    Dim actionArgs As String
    Dim commaPos As Long
    Dim replaceFrom As String
    Dim replaceTo As String
    Dim maxLen As Long

    actionSpec = mp_TrimWhitespace(CStr(actionSpec))
    colonPos = InStr(1, actionSpec, ":", vbBinaryCompare)
    If colonPos > 0 Then
        actionName = LCase$(mp_TrimWhitespace(Left$(actionSpec, colonPos - 1)))
        actionArgs = Mid$(actionSpec, colonPos + 1)
    Else
        actionName = LCase$(actionSpec)
        actionArgs = vbNullString
    End If

    Select Case actionName
        Case FORMATTER_TRUNCATE
            actionArgs = mp_TrimWhitespace(actionArgs)
            If Not mp_TryParseNonNegativeLong(actionArgs, maxLen) Then
                Err.Raise vbObjectError + 1771, "ex_ResultTemplatesParser", "truncate requires non-negative integer argument: '" & actionSpec & "'."
            End If
            If maxLen <= 0 Then
                mp_ApplyFormatterAction = vbNullString
            ElseIf Len(sourceValue) <= maxLen Then
                mp_ApplyFormatterAction = CStr(sourceValue)
            Else
                mp_ApplyFormatterAction = Left$(CStr(sourceValue), maxLen)
            End If
            Exit Function

        Case FORMATTER_REPLACE
            commaPos = InStr(1, actionArgs, ",", vbBinaryCompare)
            If commaPos <= 0 Then
                Err.Raise vbObjectError + 1772, "ex_ResultTemplatesParser", "replace requires two args 'from,to': '" & actionSpec & "'."
            End If
            replaceFrom = Left$(actionArgs, commaPos - 1)
            replaceTo = Mid$(actionArgs, commaPos + 1)
            If Len(replaceFrom) = 0 Then
                Err.Raise vbObjectError + 1773, "ex_ResultTemplatesParser", "replace 'from' argument cannot be empty: '" & actionSpec & "'."
            End If
            mp_ApplyFormatterAction = Replace(CStr(sourceValue), replaceFrom, replaceTo)
            Exit Function

        Case Else
            If Len(actionArgs) > 0 Then
                Err.Raise vbObjectError + 1774, "ex_ResultTemplatesParser", "Formatter '" & actionName & "' does not support arguments."
            End If
            mp_ApplyFormatterAction = mp_ApplyFormatter(CStr(sourceValue), actionName)
            Exit Function
    End Select
End Function

Private Function mp_TryParseNonNegativeLong(ByVal textValue As String, ByRef outValue As Long) As Boolean
    Dim i As Long
    Dim ch As String
    Dim parsed As Double

    textValue = mp_TrimWhitespace(CStr(textValue))
    If Len(textValue) = 0 Then Exit Function

    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i

    parsed = CDbl(textValue)
    If parsed < 0# Or parsed > 2147483647# Then Exit Function

    outValue = CLng(parsed)
    mp_TryParseNonNegativeLong = True
End Function

Private Function mp_FirstNonSpaceChar(ByVal textValue As String) As String
    Dim i As Long
    Dim ch As String

    textValue = CStr(textValue)
    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        If Not mp_IsWhitespaceChar(ch) Then
            mp_FirstNonSpaceChar = ch
            Exit Function
        End If
    Next i
End Function

Private Function mp_IsWhitespaceChar(ByVal ch As String) As Boolean
    Dim codePoint As Long

    If Len(ch) = 0 Then Exit Function

    codePoint = AscW(Left$(ch, 1))
    mp_IsWhitespaceChar = _
        (codePoint = 32) Or _
        (codePoint = 9) Or _
        (codePoint = 10) Or _
        (codePoint = 13) Or _
        (codePoint = NBSP_CODE_POINT) Or _
        (codePoint = NARROW_NBSP_CODE_POINT)
End Function

Private Function mp_TrimWhitespace(ByVal textValue As String) As String
    Dim startPos As Long
    Dim endPos As Long

    textValue = CStr(textValue)
    startPos = 1
    endPos = Len(textValue)

    Do While startPos <= endPos
        If Not mp_IsWhitespaceChar(Mid$(textValue, startPos, 1)) Then Exit Do
        startPos = startPos + 1
    Loop

    Do While endPos >= startPos
        If Not mp_IsWhitespaceChar(Mid$(textValue, endPos, 1)) Then Exit Do
        endPos = endPos - 1
    Loop

    If startPos > endPos Then
        mp_TrimWhitespace = vbNullString
    Else
        mp_TrimWhitespace = Mid$(textValue, startPos, endPos - startPos + 1)
    End If
End Function

Private Function mp_InflectPhraseToCase(ByVal sourceValue As String, ByVal caseName As String) As String
    Dim convertedText As String

    sourceValue = CStr(sourceValue)
    convertedText = ex_MorphUaLite.m_InflectPhraseToCase(sourceValue, caseName)
    If Len(convertedText) = 0 Then
        mp_InflectPhraseToCase = sourceValue
    Else
        mp_InflectPhraseToCase = convertedText
    End If
End Function

Private Function mp_EscapeRegex(ByVal textValue As String) As String
    Dim i As Long
    Dim ch As String
    Dim escaped As String

    escaped = vbNullString
    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        Select Case ch
            Case "\", ".", "^", "$", "|", "(", ")", "[", "]", "{", "}", "*", "+", "?"
                escaped = escaped & "\" & ch
            Case Else
                escaped = escaped & ch
        End Select
    Next i

    mp_EscapeRegex = escaped
End Function

Private Function mp_ReplaceIfConditionForPlaceholder( _
    ByVal sourceText As String, _
    ByVal placeholderName As String, _
    ByVal replacementText As String _
) As String
    Dim rx As Object
    Dim replacementCondition As String

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = True
    rx.Pattern = "\{#if\s+" & mp_EscapeRegex(placeholderName) & "\s*\}"

    replacementCondition = IF_BLOCK_OPEN & " " & mp_BooleanTextFromValue(replacementText) & "}"
    mp_ReplaceIfConditionForPlaceholder = rx.Replace(CStr(sourceText), replacementCondition)
End Function

Private Function mp_BooleanTextFromValue(ByVal textValue As String) As String
    If mp_IsTruthyConditionValue(textValue) Then
        mp_BooleanTextFromValue = BOOLEAN_TRUE
    Else
        mp_BooleanTextFromValue = BOOLEAN_FALSE
    End If
End Function

Private Function mp_IsTruthyConditionValue(ByVal textValue As String) As Boolean
    Dim normalized As String
    normalized = LCase$(mp_TrimWhitespace(CStr(textValue)))

    If Len(normalized) = 0 Then Exit Function
    If normalized = BOOLEAN_FALSE Then Exit Function
    mp_IsTruthyConditionValue = True
End Function

Private Function mp_ResolveConditionalBlocks(ByVal sourceText As String) As String
    Dim resultText As String
    Dim openPos As Long
    Dim headerEndPos As Long
    Dim closeStartPos As Long
    Dim closeEndPos As Long
    Dim conditionText As String
    Dim innerText As String
    Dim replacementText As String

    resultText = CStr(sourceText)

    Do
        openPos = InStr(1, resultText, IF_BLOCK_OPEN, vbTextCompare)
        If openPos = 0 Then Exit Do

        If Not mp_TryParseIfHeader(resultText, openPos, headerEndPos, conditionText) Then
            Err.Raise vbObjectError + 1767, "ex_ResultTemplatesParser", "Invalid if-block syntax near position " & CStr(openPos) & "."
        End If

        If Not mp_TryFindMatchingIfClose(resultText, headerEndPos + 1, closeStartPos, closeEndPos) Then
            Err.Raise vbObjectError + 1768, "ex_ResultTemplatesParser", "Missing closing {#endif} for if-block near position " & CStr(openPos) & "."
        End If

        innerText = Mid$(resultText, headerEndPos + 1, closeStartPos - headerEndPos - 1)
        innerText = mp_ResolveConditionalBlocks(innerText)

        If mp_IsTruthyConditionValue(conditionText) Then
            replacementText = innerText
        Else
            replacementText = vbNullString
        End If

        resultText = Left$(resultText, openPos - 1) & replacementText & Mid$(resultText, closeEndPos + 1)
    Loop

    If InStr(1, resultText, IF_BLOCK_CLOSE, vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 1769, "ex_ResultTemplatesParser", "Unexpected {#endif} without matching {#if ...}."
    End If

    mp_ResolveConditionalBlocks = resultText
End Function

Private Function mp_ResolveJoinLineTokens(ByVal sourceText As String) As String
    Dim resultText As String

    resultText = CStr(sourceText)
    resultText = Replace(resultText, vbCrLf, vbLf)
    resultText = Replace(resultText, vbCr, vbLf)

    resultText = mp_ResolveJoinLineToken(resultText, RESERVED_JOINLINE_TOKEN)
    resultText = mp_ResolveJoinLineToken(resultText, RESERVED_JOINLINE_TOKEN_SHORT)
    mp_ResolveJoinLineTokens = resultText
End Function

Private Function mp_ResolveJoinLineToken(ByVal sourceText As String, ByVal tokenText As String) As String
    Dim resultText As String
    Dim rx As Object
    Dim tokenPattern As String

    resultText = CStr(sourceText)
    tokenPattern = mp_EscapeRegex(CStr(tokenText))

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = False

    ' Join with next line: "...TOKEN\n    ..."
    rx.Pattern = tokenPattern & "[ \t]*" & vbLf & "[ \t]*"
    resultText = rx.Replace(resultText, vbNullString)

    ' Join with previous line: "...\n    TOKEN..."
    rx.Pattern = "[ \t]*" & vbLf & "[ \t]*" & tokenPattern
    resultText = rx.Replace(resultText, vbNullString)

    ' Fallback: strip standalone token if no line-break pattern was matched.
    resultText = Replace(resultText, CStr(tokenText), vbNullString)
    mp_ResolveJoinLineToken = resultText
End Function

Private Function mp_TryParseIfHeader( _
    ByVal sourceText As String, _
    ByVal ifOpenPos As Long, _
    ByRef outHeaderEndPos As Long, _
    ByRef outConditionText As String _
) As Boolean
    Dim closePos As Long
    Dim rawCondition As String

    If StrComp(Mid$(sourceText, ifOpenPos, Len(IF_BLOCK_OPEN)), IF_BLOCK_OPEN, vbTextCompare) <> 0 Then Exit Function

    closePos = InStr(ifOpenPos + Len(IF_BLOCK_OPEN), sourceText, "}", vbBinaryCompare)
    If closePos = 0 Then Exit Function

    rawCondition = Mid$(sourceText, ifOpenPos + Len(IF_BLOCK_OPEN), closePos - ifOpenPos - Len(IF_BLOCK_OPEN))
    outConditionText = mp_TrimWhitespace(rawCondition)
    If Len(outConditionText) = 0 Then Exit Function

    outHeaderEndPos = closePos
    mp_TryParseIfHeader = True
End Function

Private Function mp_TryFindMatchingIfClose( _
    ByVal sourceText As String, _
    ByVal searchFromPos As Long, _
    ByRef outCloseStartPos As Long, _
    ByRef outCloseEndPos As Long _
) As Boolean
    Dim depth As Long
    Dim nextOpenPos As Long
    Dim nextClosePos As Long

    depth = 1
    Do While searchFromPos <= Len(sourceText)
        nextOpenPos = InStr(searchFromPos, sourceText, IF_BLOCK_OPEN, vbTextCompare)
        nextClosePos = InStr(searchFromPos, sourceText, IF_BLOCK_CLOSE, vbTextCompare)

        If nextClosePos = 0 Then Exit Function

        If nextOpenPos > 0 And nextOpenPos < nextClosePos Then
            depth = depth + 1
            searchFromPos = nextOpenPos + Len(IF_BLOCK_OPEN)
        Else
            depth = depth - 1
            If depth = 0 Then
                outCloseStartPos = nextClosePos
                outCloseEndPos = nextClosePos + Len(IF_BLOCK_CLOSE) - 1
                mp_TryFindMatchingIfClose = True
                Exit Function
            End If
            searchFromPos = nextClosePos + Len(IF_BLOCK_CLOSE)
        End If
    Loop
End Function

Private Function mp_ResolveDateExpressions( _
    ByVal sourceText As String, _
    Optional ByVal baseDateText As String = vbNullString _
) As String
    Dim baseDate As Date
    Dim rx As Object
    Dim matches As Object
    Dim i As Long
    Dim tokenText As String
    Dim resolvedDay As String
    Dim resolvedByToken As Object
    Dim key As Variant

    baseDate = mp_ParseBaseDate(baseDateText)
    Set resolvedByToken = CreateObject("Scripting.Dictionary")
    resolvedByToken.CompareMode = 1

    Set rx = mp_CreateDateTokenRegex()
    Set matches = rx.Execute(CStr(sourceText))
    mp_ResolveDateExpressions = CStr(sourceText)

    If matches Is Nothing Then Exit Function
    If matches.Count = 0 Then Exit Function

    For i = 0 To matches.Count - 1
        tokenText = CStr(matches(i).Value)
        If Not resolvedByToken.Exists(tokenText) Then
            resolvedDay = mp_ResolveDateToken(tokenText, baseDate)
            resolvedByToken.Add tokenText, resolvedDay
        End If
    Next i

    For Each key In resolvedByToken.Keys
        mp_ResolveDateExpressions = Replace(mp_ResolveDateExpressions, CStr(key), CStr(resolvedByToken(CStr(key))))
    Next key
End Function

Private Function mp_NormalizeTemplateText(ByVal templateText As String) As String
    templateText = Replace(templateText, vbCrLf, vbLf)
    templateText = Replace(templateText, vbCr, vbLf)
    If Left$(templateText, 1) = vbLf Then templateText = Mid$(templateText, 2)
    If Right$(templateText, 1) = vbLf Then templateText = Left$(templateText, Len(templateText) - 1)
    mp_NormalizeTemplateText = templateText
End Function

Private Function mp_PrependTemplateError(ByVal sourceText As String, ByVal operationName As String) As String
    Dim errorLine As String
    Dim fullText As String

    errorLine = TEMPLATE_ERROR_PREFIX & " " & operationName & ": [" & CStr(Err.Source) & " #" & CStr(Err.Number) & "] " & CStr(Err.Description)
    fullText = CStr(sourceText)

    If Len(fullText) = 0 Then
        mp_PrependTemplateError = errorLine
        Exit Function
    End If

    If StrComp(Left$(fullText, Len(errorLine)), errorLine, vbBinaryCompare) = 0 Then
        mp_PrependTemplateError = fullText
        Exit Function
    End If

    mp_PrependTemplateError = errorLine & vbLf & fullText
End Function

Private Function mp_CreateDateTokenRegex() As Object
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = False
    regex.Pattern = DATE_TOKEN_PATTERN
    Set mp_CreateDateTokenRegex = regex
End Function

Private Function mp_ParseBaseDate(ByVal baseDateText As String) As Date
    Dim parsedDate As Date

    baseDateText = mp_TrimWhitespace(baseDateText)
    If Len(baseDateText) = 0 Then
        mp_ParseBaseDate = Date
        Exit Function
    End If

    On Error GoTo ParseError
    parsedDate = CDate(baseDateText)
    mp_ParseBaseDate = parsedDate
    Exit Function

ParseError:
    Err.Raise vbObjectError + 1763, "ex_ResultTemplatesParser", _
        "Invalid base date '" & baseDateText & "' for date expressions."
End Function

Private Function mp_ResolveDateToken(ByVal tokenText As String, ByVal baseDate As Date) As String
    Dim innerText As String
    Dim offsetText As String
    Dim offsetValue As Long
    Dim resolvedDate As Date

    innerText = Mid$(tokenText, 2, Len(tokenText) - 2) ' without braces
    If Left$(innerText, 1) <> RESERVED_TOKEN_PREFIX Then
        Err.Raise vbObjectError + 1765, "ex_ResultTemplatesParser", _
            "Unsupported reserved token '" & tokenText & "'. Reserved tokens must start with '#'."
    End If
    innerText = Mid$(innerText, 2)

    If LCase$(Left$(innerText, 2)) <> RESERVED_DATE_DAY_KEYWORD Then
        Err.Raise vbObjectError + 1765, "ex_ResultTemplatesParser", _
            "Unsupported reserved token '" & tokenText & "'."
    End If

    offsetText = Mid$(innerText, 3) ' suffix after dd
    If Len(offsetText) = 0 Then
        offsetValue = 0
    Else
        On Error GoTo ParseError
        offsetValue = CLng(offsetText)
    End If

    resolvedDate = DateAdd("d", offsetValue, baseDate)
    mp_ResolveDateToken = Format$(resolvedDate, RESERVED_DATE_OUTPUT_FORMAT)
    Exit Function

ParseError:
    Err.Raise vbObjectError + 1764, "ex_ResultTemplatesParser", _
        "Invalid date token '" & tokenText & "'. Use {#dd}, {#dd+N}, {#dd-N}."
End Function
