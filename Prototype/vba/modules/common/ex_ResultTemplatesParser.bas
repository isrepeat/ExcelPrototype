Attribute VB_Name = "ex_ResultTemplatesParser"
Option Explicit

Private Const PROFILES_NS As String = "urn:excelprototype:profiles"

' Reserved keywords/tokens supported by this parser.
' Reserved numeric offset token:
' - <NUMERIC>{+N}
' - <NUMERIC>{-N}
' Reserved line-join token:
' - {#^}  (ignores a line break around token)
' - #^    (same behavior, shorthand)
' Legacy line-join token (backward compatibility):
' - {#_}
' Reserved trim-indentation token:
' - #_    (removes token and horizontal whitespace after it)
' Reserved computed-value token:
' - #let varName = $ModuleName.MethodName(arg1, arg2);
' Reserved if-condition unary operator:
' - #not
' Reserved if-condition numeric comparison operators:
' - ==, !=, >, <, >=, <=
Private Const NUMERIC_OFFSET_TOKEN_PATTERN As String = "\{([+-]\d+)\}"
Private Const LEGACY_DAY_TOKEN_PATTERN As String = "\{#dd(?:[+-]\d+)?\}"
Private Const RESERVED_JOINLINE_TOKEN As String = "{#^}"
Private Const RESERVED_JOINLINE_TOKEN_SHORT As String = "#^"
Private Const RESERVED_JOINLINE_TOKEN_LEGACY As String = "{#_}"
Private Const RESERVED_TRIMINDENT_TOKEN_SHORT As String = "#_"
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
Private Const FORMATTER_DATEFORMAT As String = "dateformat"
Private Const FORMATTER_TO_DATE_DAY As String = "todate_day"
Private Const FORMATTER_TO_DATE_DAY_WITH_MONTH As String = "todate_daywithmonth"
Private Const FORMATTER_CALENDAR_DAYS_UA As String = "calendardaysua"
Private Const FORMATTER_SURNAME_INITIALS As String = "surnameinitials"
Private Const FORMATTER_FIO_SURNAME As String = "fiosurname"
Private Const FORMATTER_FIO_INITIALS As String = "fioinitials"

Private Const CASE_GENITIVE As String = "genitive"
Private Const CASE_ACCUSATIVE As String = "accusative"
Private Const CASE_DATIVE As String = "dative"

Private Const NBSP_CODE_POINT As Long = 160
Private Const NARROW_NBSP_CODE_POINT As Long = 8239
Private Const TEMPLATE_ERROR_PREFIX As String = "[TEMPLATE ERROR]"
Private Const HIGHLIGHT_MARKER_START As String = "/Start"
Private Const HIGHLIGHT_MARKER_END As String = "/End"
Private Const DEFAULT_HIGHLIGHT_COLOR_HEX As String = "#66CCFF"

Public Function m_GetTemplateText( _
    ByVal templateId As String, _
    ByVal resultTemplatesRelPath As String _
) As String
    Dim doc As Object
    Dim node As Object
    Dim xpath As String
    Dim templateText As String
    Dim templatesPath As String

    templateId = mp_TrimWhitespace(templateId)
    If Len(templateId) = 0 Then
        Err.Raise vbObjectError + 1760, "ex_ResultTemplatesParser", "Template id is empty."
    End If
    templatesPath = mp_TrimWhitespace(resultTemplatesRelPath)
    If Len(templatesPath) = 0 Then
        Err.Raise vbObjectError + 1819, "ex_ResultTemplatesParser", "Result templates path is empty."
    End If

    On Error GoTo EH

    Set doc = ex_XmlCore.m_LoadDomByRelativePath( _
        ThisWorkbook, _
        templatesPath, _
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
    m_GetTemplateText = mp_PrependTemplateError(vbNullString, "m_GetTemplateText('" & templateId & "', '" & templatesPath & "')")
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
    ByVal replacementText As String, _
    Optional ByVal highlightColorHex As String = vbNullString _
) As String
    Dim normalizedName As String
    Dim resultText As String
    Dim rx As Object
    Dim matches As Object
    Dim i As Long
    Dim formatter As String
    Dim formattedValue As String
    Dim matchStart As Long
    Dim matchLen As Long
    Dim offsetText As String
    Dim resolvedValue As String

    On Error GoTo EH

    normalizedName = mp_TrimWhitespace(placeholderName)
    If Len(normalizedName) = 0 Then
        m_ReplacePlaceholder = CStr(sourceText)
        Exit Function
    End If

    If mp_ShouldCollapsePlaceholderValue(replacementText) Then
        resultText = mp_CollapseNamedPlaceholderTokens(CStr(sourceText), normalizedName)
        resultText = mp_ReplaceIfConditionForPlaceholder(resultText, normalizedName, vbNullString)
        m_ReplacePlaceholder = resultText
        Exit Function
    End If

    resultText = CStr(sourceText)

    ' Resolve formatted placeholders first: {Name|formatter}{+N}
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = False
    rx.Pattern = "\{" & mp_EscapeRegex(normalizedName) & "\|([^{}]+)\}(\{([+-]\d+)\})?"

    Set matches = rx.Execute(resultText)
    If Not matches Is Nothing Then
        For i = matches.Count - 1 To 0 Step -1
            formatter = CStr(matches(i).SubMatches(0))
            If matches(i).SubMatches.Count >= 3 Then
                offsetText = CStr(matches(i).SubMatches(2))
            Else
                offsetText = vbNullString
            End If
            formattedValue = mp_ApplyFormatterPipeline(CStr(replacementText), formatter)
            resolvedValue = mp_ResolveNumericOffsetForPlaceholderValue(formattedValue, offsetText, normalizedName)
            resolvedValue = mp_WrapHighlightMarkers(resolvedValue, highlightColorHex)
            matchStart = CLng(matches(i).FirstIndex)
            matchLen = CLng(matches(i).Length)

            resultText = Left$(resultText, matchStart) & resolvedValue & Mid$(resultText, matchStart + matchLen + 1)
        Next i
    End If

    ' Resolve plain placeholders: {Name}{+N}
    rx.Pattern = "\{" & mp_EscapeRegex(normalizedName) & "\}(\{([+-]\d+)\})?"
    Set matches = rx.Execute(resultText)
    If Not matches Is Nothing Then
        For i = matches.Count - 1 To 0 Step -1
            If matches(i).SubMatches.Count >= 2 Then
                offsetText = CStr(matches(i).SubMatches(1))
            Else
                offsetText = vbNullString
            End If

            resolvedValue = mp_ResolveNumericOffsetForPlaceholderValue(CStr(replacementText), offsetText, normalizedName)
            resolvedValue = mp_WrapHighlightMarkers(resolvedValue, highlightColorHex)
            matchStart = CLng(matches(i).FirstIndex)
            matchLen = CLng(matches(i).Length)

            resultText = Left$(resultText, matchStart) & resolvedValue & Mid$(resultText, matchStart + matchLen + 1)
        Next i
    End If

    resultText = mp_ReplaceIfConditionForPlaceholder(resultText, normalizedName, CStr(replacementText))
    m_ReplacePlaceholder = resultText
    Exit Function

EH:
    If Len(resultText) = 0 Then resultText = CStr(sourceText)
    m_ReplacePlaceholder = mp_PrependTemplateError(resultText, "m_ReplacePlaceholder('" & CStr(placeholderName) & "')")
End Function

Private Function mp_ResolveNumericOffsetForPlaceholderValue( _
    ByVal sourceValue As String, _
    ByVal offsetText As String, _
    ByVal placeholderName As String _
) As String
    Dim baseNumber As Double
    Dim baseIsInteger As Boolean
    Dim offsetValue As Double
    Dim resultNumber As Double

    sourceValue = CStr(sourceValue)
    offsetText = mp_TrimWhitespace(CStr(offsetText))
    If Len(offsetText) = 0 Then
        mp_ResolveNumericOffsetForPlaceholderValue = sourceValue
        Exit Function
    End If

    If Not mp_TryParseTemplateNumeric(sourceValue, baseNumber, baseIsInteger) Then
        Err.Raise vbObjectError + 1824, "ex_ResultTemplatesParser", _
            "Placeholder '" & CStr(placeholderName) & "' value '" & CStr(sourceValue) & "' is not numeric for offset '{" & CStr(offsetText) & "}'."
    End If
    If Not mp_TryParseSignedInteger(offsetText, offsetValue) Then
        Err.Raise vbObjectError + 1825, "ex_ResultTemplatesParser", _
            "Invalid numeric offset '{" & CStr(offsetText) & "}' after placeholder '" & CStr(placeholderName) & "'. Use '{+N}' or '{-N}'."
    End If

    resultNumber = baseNumber + offsetValue
    If baseIsInteger And Fix(resultNumber) = resultNumber Then
        mp_ResolveNumericOffsetForPlaceholderValue = mp_FormatIntegerWithBasePadding(CLng(resultNumber), sourceValue)
    Else
        mp_ResolveNumericOffsetForPlaceholderValue = CStr(resultNumber)
    End If
End Function

Public Function m_ResolveTemplate( _
    ByVal sourceText As String, _
    Optional ByVal baseDateText As String = vbNullString, _
    Optional ByVal highlightColorHex As String = vbNullString _
) As String
    On Error GoTo EH
    ' Final pass for template text.
    m_ResolveTemplate = mp_ResolveTemplateLetBindings(CStr(sourceText), highlightColorHex)
    m_ResolveTemplate = mp_ResolveConditionalBlocks(m_ResolveTemplate)
    m_ResolveTemplate = mp_ResolveJoinLineTokens(m_ResolveTemplate)
    m_ResolveTemplate = mp_ResolveTrimIndentTokens(m_ResolveTemplate)
    m_ResolveTemplate = mp_ResolveDateExpressions(m_ResolveTemplate, baseDateText)
    m_ResolveTemplate = mp_CollapseUnresolvedPlaceholders(m_ResolveTemplate)
    m_ResolveTemplate = mp_CollapseHorizontalWhitespaceRuns(m_ResolveTemplate)
    Exit Function

EH:
    m_ResolveTemplate = mp_PrependTemplateError(CStr(sourceText), "m_ResolveTemplate")
End Function

Public Function m_ExtractHighlightSegments( _
    ByVal sourceText As String, _
    ByRef outSegments As Collection, _
    Optional ByVal fallbackColorHex As String = vbNullString _
) As String
    Dim resultText As String
    Dim scanPos As Long
    Dim startPos As Long
    Dim valueStartPos As Long
    Dim endPos As Long
    Dim prefixText As String
    Dim valueText As String
    Dim colorHex As String
    Dim colorStartPos As Long
    Dim colorEndPos As Long
    Dim segmentStart As Long
    Dim segment As Object

    On Error GoTo EH

    Set outSegments = New Collection
    resultText = vbNullString
    sourceText = CStr(sourceText)
    fallbackColorHex = mp_TrimWhitespace(CStr(fallbackColorHex))
    If Len(fallbackColorHex) = 0 Then fallbackColorHex = DEFAULT_HIGHLIGHT_COLOR_HEX

    scanPos = 1
    Do
        startPos = InStr(scanPos, sourceText, HIGHLIGHT_MARKER_START, vbBinaryCompare)
        If startPos = 0 Then
            resultText = resultText & Mid$(sourceText, scanPos)
            Exit Do
        End If

        prefixText = Mid$(sourceText, scanPos, startPos - scanPos)
        resultText = resultText & prefixText

        valueStartPos = startPos + Len(HIGHLIGHT_MARKER_START)
        colorHex = fallbackColorHex

        If valueStartPos <= Len(sourceText) Then
            If Mid$(sourceText, valueStartPos, 1) = "(" Then
                colorStartPos = valueStartPos + 1
                colorEndPos = InStr(colorStartPos, sourceText, ")", vbBinaryCompare)
                If colorEndPos > 0 Then
                    colorHex = mp_TrimWhitespace(Mid$(sourceText, colorStartPos, colorEndPos - colorStartPos))
                    If Len(colorHex) = 0 Then colorHex = fallbackColorHex
                    valueStartPos = colorEndPos + 1
                End If
            End If
        End If

        endPos = InStr(valueStartPos, sourceText, HIGHLIGHT_MARKER_END, vbBinaryCompare)
        If endPos = 0 Then
            resultText = resultText & Mid$(sourceText, startPos)
            Exit Do
        End If

        valueText = Mid$(sourceText, valueStartPos, endPos - valueStartPos)
        segmentStart = Len(resultText) + 1
        resultText = resultText & valueText

        If Len(valueText) > 0 Then
            Set segment = CreateObject("Scripting.Dictionary")
            segment.CompareMode = 1 ' vbTextCompare
            segment("Start") = CLng(segmentStart)
            segment("Length") = CLng(Len(valueText))
            segment("ColorHex") = CStr(colorHex)
            outSegments.Add segment
        End If

        scanPos = endPos + Len(HIGHLIGHT_MARKER_END)
    Loop

    m_ExtractHighlightSegments = resultText
    Exit Function

EH:
    Set outSegments = New Collection
    m_ExtractHighlightSegments = CStr(sourceText)
End Function

Public Function m_RemoveHighlightMarkers(ByVal sourceText As String) As String
    Dim segments As Collection

    m_RemoveHighlightMarkers = m_ExtractHighlightSegments(CStr(sourceText), segments, DEFAULT_HIGHLIGHT_COLOR_HEX)
End Function

Private Function mp_ShouldCollapsePlaceholderValue(ByVal replacementText As String) As Boolean
    Dim normalized As String

    normalized = mp_TrimWhitespace(CStr(replacementText))
    If Len(normalized) = 0 Then
        mp_ShouldCollapsePlaceholderValue = True
        Exit Function
    End If

    If normalized = "-" Or normalized = "?" Then
        mp_ShouldCollapsePlaceholderValue = True
    End If
End Function

Private Function mp_WrapHighlightMarkers( _
    ByVal valueText As String, _
    ByVal highlightColorHex As String _
) As String
    highlightColorHex = mp_TrimWhitespace(CStr(highlightColorHex))
    valueText = CStr(valueText)

    If Len(valueText) = 0 Or Len(highlightColorHex) = 0 Then
        mp_WrapHighlightMarkers = valueText
        Exit Function
    End If

    mp_WrapHighlightMarkers = HIGHLIGHT_MARKER_START & "(" & highlightColorHex & ")" & valueText & HIGHLIGHT_MARKER_END
End Function

Private Function mp_CollapseNamedPlaceholderTokens( _
    ByVal sourceText As String, _
    ByVal placeholderName As String _
) As String
    Dim rx As Object
    Dim tokenPattern As String
    Dim resultText As String

    resultText = CStr(sourceText)
    tokenPattern = "\{" & mp_EscapeRegex(CStr(placeholderName)) & "(\|[^{}]+)?\}"

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = False
    rx.Pattern = " ?" & tokenPattern & " ?"

    mp_CollapseNamedPlaceholderTokens = rx.Replace(resultText, vbNullString)
End Function

Private Function mp_CollapseUnresolvedPlaceholders(ByVal sourceText As String) As String
    Dim rx As Object
    Dim resultText As String

    resultText = CStr(sourceText)

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = False
    rx.Pattern = "\{[A-Za-z_][A-Za-z0-9_]*(\|[^{}]+)?\}"

    mp_CollapseUnresolvedPlaceholders = rx.Replace(resultText, vbNullString)
End Function

Private Function mp_CollapseHorizontalWhitespaceRuns(ByVal sourceText As String) As String
    Dim rx As Object

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = False
    rx.Pattern = "[ \t]{2,}"

    mp_CollapseHorizontalWhitespaceRuns = rx.Replace(CStr(sourceText), " ")
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

Private Function mp_FormatUaDateDay(ByVal sourceDateText As String) As String
    mp_FormatUaDateDay = ex_DateHelpers.m_FormatDateDay(CStr(sourceDateText))
End Function

Private Function mp_FormatUaDateDayWithMonth(ByVal sourceDateText As String) As String
    mp_FormatUaDateDayWithMonth = ex_DateHelpers.m_FormatDateDayWithMonth(CStr(sourceDateText))
End Function

Private Function mp_FormatSurnameInitials(ByVal sourceValue As String) As String
    mp_FormatSurnameInitials = ex_MorphUaLite.m_ToShortFioNormalized(CStr(sourceValue))
End Function

Private Function mp_FormatFioSurname(ByVal sourceValue As String) As String
    mp_FormatFioSurname = ex_MorphUaLite.m_ToFioSurnameNormalized(CStr(sourceValue))
End Function

Private Function mp_FormatFioInitials(ByVal sourceValue As String) As String
    mp_FormatFioInitials = ex_MorphUaLite.m_ToFioInitials(CStr(sourceValue))
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
        Case FORMATTER_TO_DATE_DAY
            mp_ApplyFormatter = mp_FormatUaDateDay(CStr(sourceValue))
        Case FORMATTER_TO_DATE_DAY_WITH_MONTH
            mp_ApplyFormatter = mp_FormatUaDateDayWithMonth(CStr(sourceValue))
        Case FORMATTER_CALENDAR_DAYS_UA
            mp_ApplyFormatter = ex_DateHelpers.m_FormatCalendarDaysUa(CStr(sourceValue))
        Case FORMATTER_SURNAME_INITIALS
            mp_ApplyFormatter = mp_FormatSurnameInitials(CStr(sourceValue))
        Case FORMATTER_FIO_SURNAME
            mp_ApplyFormatter = mp_FormatFioSurname(CStr(sourceValue))
        Case FORMATTER_FIO_INITIALS
            mp_ApplyFormatter = mp_FormatFioInitials(CStr(sourceValue))
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

        Case FORMATTER_DATEFORMAT
            actionArgs = mp_UnquoteFormatterArgument(actionArgs)
            If Len(actionArgs) = 0 Then
                Err.Raise vbObjectError + 1801, "ex_ResultTemplatesParser", "dateformat requires non-empty format argument: '" & actionSpec & "'."
            End If
            mp_ApplyFormatterAction = ex_DateHelpers.m_FormatDateByPattern(CStr(sourceValue), actionArgs)
            Exit Function

        Case Else
            If Len(actionArgs) > 0 Then
                Err.Raise vbObjectError + 1774, "ex_ResultTemplatesParser", "Formatter '" & actionName & "' does not support arguments."
            End If
            mp_ApplyFormatterAction = mp_ApplyFormatter(CStr(sourceValue), actionName)
            Exit Function
    End Select
End Function

Private Function mp_UnquoteFormatterArgument(ByVal argText As String) As String
    Dim normalized As String

    normalized = mp_TrimWhitespace(CStr(argText))
    If Len(normalized) >= 2 Then
        If (Left$(normalized, 1) = """" And Right$(normalized, 1) = """") Or _
           (Left$(normalized, 1) = "'" And Right$(normalized, 1) = "'") Then
            normalized = Mid$(normalized, 2, Len(normalized) - 2)
        End If
    End If

    mp_UnquoteFormatterArgument = normalized
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
    Dim rxExpr As Object
    Dim replacementCondition As String
    Dim resultText As String
    Dim updatedText As String

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = True
    rx.Pattern = "\{#if\s+" & mp_EscapeRegex(placeholderName) & "\s*\}"

    replacementCondition = IF_BLOCK_OPEN & " " & mp_BooleanTextFromValue(replacementText) & "}"
    resultText = rx.Replace(CStr(sourceText), replacementCondition)

    Set rxExpr = CreateObject("VBScript.RegExp")
    rxExpr.Global = True
    rxExpr.IgnoreCase = False
    rxExpr.Pattern = "(\{#if\s+[^}]*)\b" & mp_EscapeRegex(placeholderName) & "\b"

    Do
        updatedText = rxExpr.Replace(resultText, "$1" & CStr(replacementText))
        If StrComp(updatedText, resultText, vbBinaryCompare) = 0 Then Exit Do
        resultText = updatedText
    Loop

    mp_ReplaceIfConditionForPlaceholder = resultText
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
    Dim tokens As Variant
    Dim tokenCount As Long
    Dim tokenLBound As Long
    Dim tokenUBound As Long
    Dim tokenIndex As Long

    normalized = mp_NormalizeConditionWhitespace(CStr(textValue))
    If Len(normalized) = 0 Then Exit Function

    tokens = Split(normalized, " ")
    tokenCount = mp_GetArrayItemCount(tokens)
    If tokenCount <= 0 Then Exit Function

    tokenLBound = LBound(tokens)
    tokenUBound = UBound(tokens)
    tokenIndex = tokenLBound

    mp_IsTruthyConditionValue = mp_EvaluateConditionOr(tokens, tokenIndex, tokenLBound, tokenUBound, CStr(textValue))
    If tokenIndex <= tokenUBound Then
        Err.Raise vbObjectError + 1812, "ex_ResultTemplatesParser", _
            "Unsupported if-condition '" & CStr(textValue) & "' near token '" & CStr(tokens(tokenIndex)) & "'."
    End If
End Function

Private Function mp_EvaluateConditionOr( _
    ByVal tokens As Variant, _
    ByRef tokenIndex As Long, _
    ByVal tokenLBound As Long, _
    ByVal tokenUBound As Long, _
    ByVal sourceCondition As String _
) As Boolean
    Dim resultValue As Boolean
    Dim nextValue As Boolean
    Dim tokenText As String

    resultValue = mp_EvaluateConditionAnd(tokens, tokenIndex, tokenLBound, tokenUBound, sourceCondition)
    Do While tokenIndex <= tokenUBound
        tokenText = LCase$(CStr(tokens(tokenIndex)))
        If tokenText <> "#or" Then Exit Do

        tokenIndex = tokenIndex + 1
        nextValue = mp_EvaluateConditionAnd(tokens, tokenIndex, tokenLBound, tokenUBound, sourceCondition)
        resultValue = (resultValue Or nextValue)
    Loop

    mp_EvaluateConditionOr = resultValue
End Function

Private Function mp_EvaluateConditionAnd( _
    ByVal tokens As Variant, _
    ByRef tokenIndex As Long, _
    ByVal tokenLBound As Long, _
    ByVal tokenUBound As Long, _
    ByVal sourceCondition As String _
) As Boolean
    Dim resultValue As Boolean
    Dim nextValue As Boolean
    Dim tokenText As String

    resultValue = mp_EvaluateConditionUnary(tokens, tokenIndex, tokenLBound, tokenUBound, sourceCondition)
    Do While tokenIndex <= tokenUBound
        tokenText = LCase$(CStr(tokens(tokenIndex)))
        If tokenText <> "#and" Then Exit Do

        tokenIndex = tokenIndex + 1
        nextValue = mp_EvaluateConditionUnary(tokens, tokenIndex, tokenLBound, tokenUBound, sourceCondition)
        resultValue = (resultValue And nextValue)
    Loop

    mp_EvaluateConditionAnd = resultValue
End Function

Private Function mp_EvaluateConditionUnary( _
    ByVal tokens As Variant, _
    ByRef tokenIndex As Long, _
    ByVal tokenLBound As Long, _
    ByVal tokenUBound As Long, _
    ByVal sourceCondition As String _
) As Boolean
    Dim isNegated As Boolean
    Dim tokenText As String
    Dim atomStart As Long
    Dim atomText As String
    Dim atomValue As Boolean

    Do While tokenIndex <= tokenUBound
        tokenText = LCase$(CStr(tokens(tokenIndex)))
        If tokenText <> "#not" Then Exit Do
        isNegated = Not isNegated
        tokenIndex = tokenIndex + 1
    Loop

    If tokenIndex > tokenUBound Then
        atomValue = False
    Else
        tokenText = LCase$(CStr(tokens(tokenIndex)))
        If tokenText = "#and" Or tokenText = "#or" Then
            ' Missing operand is treated as False to handle empty placeholders in expressions.
            atomValue = False
        Else
            atomStart = tokenIndex
            Do While tokenIndex <= tokenUBound
                tokenText = LCase$(CStr(tokens(tokenIndex)))
                If tokenText = "#and" Or tokenText = "#or" Then Exit Do
                tokenIndex = tokenIndex + 1
            Loop

            atomText = mp_JoinConditionTokens(tokens, atomStart, tokenIndex - 1)
            atomValue = mp_EvaluateAtomicCondition(atomText, sourceCondition)
        End If
    End If

    If isNegated Then atomValue = Not atomValue
    mp_EvaluateConditionUnary = atomValue
End Function

Private Function mp_EvaluateAtomicCondition(ByVal atomicText As String, ByVal sourceCondition As String) As Boolean
    Dim normalized As String
    Dim hasNumericComparison As Boolean

    normalized = LCase$(mp_TrimWhitespace(CStr(atomicText)))
    If Len(normalized) = 0 Then Exit Function

    hasNumericComparison = mp_HasNumericComparisonOperator(normalized)
    If hasNumericComparison Then
        If Not mp_TryEvaluateNumericComparisonCondition(normalized, mp_EvaluateAtomicCondition) Then
            Err.Raise vbObjectError + 1810, "ex_ResultTemplatesParser", _
                "Unsupported numeric if-condition '" & CStr(sourceCondition) & "'. Use '<NUMBER> <OP> <NUMBER>' where OP is ==, !=, >, <, >=, <=."
        End If
        Exit Function
    End If

    If normalized = BOOLEAN_FALSE Then
        mp_EvaluateAtomicCondition = False
    Else
        mp_EvaluateAtomicCondition = True
    End If
End Function

Private Function mp_JoinConditionTokens(ByVal tokens As Variant, ByVal startIndex As Long, ByVal endIndex As Long) As String
    Dim i As Long
    Dim resultText As String

    If endIndex < startIndex Then Exit Function

    For i = startIndex To endIndex
        If Len(resultText) > 0 Then resultText = resultText & " "
        resultText = resultText & CStr(tokens(i))
    Next i

    mp_JoinConditionTokens = resultText
End Function

Private Function mp_NormalizeConditionWhitespace(ByVal conditionText As String) As String
    Dim i As Long
    Dim ch As String
    Dim resultText As String
    Dim needSeparator As Boolean

    conditionText = CStr(conditionText)
    For i = 1 To Len(conditionText)
        ch = Mid$(conditionText, i, 1)
        If mp_IsWhitespaceChar(ch) Then
            If Len(resultText) > 0 Then needSeparator = True
        Else
            If needSeparator Then
                resultText = resultText & " "
                needSeparator = False
            End If
            resultText = resultText & ch
        End If
    Next i

    mp_NormalizeConditionWhitespace = mp_TrimWhitespace(resultText)
End Function

Private Function mp_TryStripNotPrefix(ByRef conditionText As String) As Boolean
    Dim nextCh As String

    conditionText = LCase$(mp_TrimWhitespace(CStr(conditionText)))
    If Len(conditionText) < 4 Then Exit Function
    If Left$(conditionText, 4) <> "#not" Then Exit Function

    If Len(conditionText) = 4 Then
        conditionText = vbNullString
        mp_TryStripNotPrefix = True
        Exit Function
    End If

    nextCh = Mid$(conditionText, 5, 1)
    If Not mp_IsWhitespaceChar(nextCh) Then Exit Function

    conditionText = LCase$(mp_TrimWhitespace(Mid$(conditionText, 5)))
    mp_TryStripNotPrefix = True
End Function

Private Function mp_HasNumericComparisonOperator(ByVal conditionText As String) As Boolean
    conditionText = CStr(conditionText)
    If InStr(1, conditionText, "==", vbBinaryCompare) > 0 Then
        mp_HasNumericComparisonOperator = True
        Exit Function
    End If
    If InStr(1, conditionText, "!=", vbBinaryCompare) > 0 Then
        mp_HasNumericComparisonOperator = True
        Exit Function
    End If
    If InStr(1, conditionText, ">=", vbBinaryCompare) > 0 Then
        mp_HasNumericComparisonOperator = True
        Exit Function
    End If
    If InStr(1, conditionText, "<=", vbBinaryCompare) > 0 Then
        mp_HasNumericComparisonOperator = True
        Exit Function
    End If
    If InStr(1, conditionText, ">", vbBinaryCompare) > 0 Then
        mp_HasNumericComparisonOperator = True
        Exit Function
    End If
    If InStr(1, conditionText, "<", vbBinaryCompare) > 0 Then
        mp_HasNumericComparisonOperator = True
        Exit Function
    End If
End Function

Private Function mp_TryEvaluateNumericComparisonCondition( _
    ByVal conditionText As String, _
    ByRef outResult As Boolean _
) As Boolean
    Dim rx As Object
    Dim matches As Object
    Dim leftText As String
    Dim rightText As String
    Dim operatorText As String
    Dim leftValue As Double
    Dim rightValue As Double
    Dim leftIsInteger As Boolean
    Dim rightIsInteger As Boolean

    conditionText = mp_TrimWhitespace(CStr(conditionText))
    If Len(conditionText) = 0 Then Exit Function

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = False
    rx.Pattern = "^\s*(.+?)\s*(==|!=|>=|<=|>|<)\s*(.+?)\s*$"

    Set matches = rx.Execute(conditionText)
    If matches Is Nothing Then Exit Function
    If matches.Count = 0 Then Exit Function

    leftText = mp_TrimWhitespace(CStr(matches(0).SubMatches(0)))
    operatorText = CStr(matches(0).SubMatches(1))
    rightText = mp_TrimWhitespace(CStr(matches(0).SubMatches(2)))

    If Not mp_TryParseTemplateNumeric(leftText, leftValue, leftIsInteger) Then Exit Function
    If Not mp_TryParseTemplateNumeric(rightText, rightValue, rightIsInteger) Then Exit Function

    Select Case operatorText
        Case "=="
            outResult = (leftValue = rightValue)
        Case "!="
            outResult = (leftValue <> rightValue)
        Case ">"
            outResult = (leftValue > rightValue)
        Case "<"
            outResult = (leftValue < rightValue)
        Case ">="
            outResult = (leftValue >= rightValue)
        Case "<="
            outResult = (leftValue <= rightValue)
        Case Else
            Exit Function
    End Select

    mp_TryEvaluateNumericComparisonCondition = True
End Function

Private Function mp_ResolveTemplateLetBindings( _
    ByVal sourceText As String, _
    Optional ByVal highlightColorHex As String = vbNullString _
) As String
    Dim resultText As String
    Dim rx As Object
    Dim matches As Object
    Dim i As Long
    Dim letVarName As String
    Dim letExpression As String
    Dim letValue As String
    Dim matchStart As Long
    Dim matchLen As Long
    Dim valuesByVar As Object
    Dim key As Variant

    resultText = CStr(sourceText)

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = False
    rx.Pattern = "#let\s+([A-Za-z_][A-Za-z0-9_]*)\s*=\s*([^;]+?)\s*;"

    Set matches = rx.Execute(resultText)
    If matches Is Nothing Then
        mp_ResolveTemplateLetBindings = resultText
        Exit Function
    End If
    If matches.Count = 0 Then
        mp_ResolveTemplateLetBindings = resultText
        Exit Function
    End If

    Set valuesByVar = CreateObject("Scripting.Dictionary")
    valuesByVar.CompareMode = 1 ' vbTextCompare

    For i = matches.Count - 1 To 0 Step -1
        letVarName = mp_TrimWhitespace(CStr(matches(i).SubMatches(0)))
        letExpression = mp_TrimWhitespace(CStr(matches(i).SubMatches(1)))
        letValue = mp_EvaluateTemplateLetExpression(letExpression)

        If valuesByVar.Exists(letVarName) Then
            Err.Raise vbObjectError + 1778, "ex_ResultTemplatesParser", "Template let variable '" & letVarName & "' is already declared."
        End If
        valuesByVar.Add letVarName, letValue

        matchStart = CLng(matches(i).FirstIndex)
        matchLen = CLng(matches(i).Length)
        resultText = Left$(resultText, matchStart) & Mid$(resultText, matchStart + matchLen + 1)
    Next i

    If InStr(1, resultText, "#let", vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 1809, "ex_ResultTemplatesParser", "Invalid #let syntax. Use '#let <VAR> = <EXPR>;' format."
    End If

    resultText = mp_RemoveEmptyLetContainers(resultText)

    For Each key In valuesByVar.Keys
        resultText = m_ReplacePlaceholder(resultText, CStr(key), CStr(valuesByVar(CStr(key))), highlightColorHex)
    Next key

    mp_ResolveTemplateLetBindings = resultText
End Function

Private Function mp_RemoveEmptyLetContainers(ByVal sourceText As String) As String
    Dim resultText As String
    Dim updatedText As String
    Dim rx As Object

    resultText = CStr(sourceText)

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = False
    rx.Pattern = "\{[ \t\r\n]*\}"

    Do
        updatedText = rx.Replace(resultText, vbNullString)
        If StrComp(updatedText, resultText, vbBinaryCompare) = 0 Then Exit Do
        resultText = updatedText
    Loop

    mp_RemoveEmptyLetContainers = resultText
End Function

Private Function mp_EvaluateTemplateLetExpression(ByVal expressionText As String) As String
    Dim normalizedExpression As String
    Dim openPos As Long
    Dim closePos As Long
    Dim helperRef As String
    Dim argsText As String
    Dim args As Variant

    normalizedExpression = mp_TrimWhitespace(CStr(expressionText))
    If Right$(normalizedExpression, 1) = ";" Then
        normalizedExpression = mp_TrimWhitespace(Left$(normalizedExpression, Len(normalizedExpression) - 1))
    End If
    If Len(normalizedExpression) = 0 Then
        Err.Raise vbObjectError + 1779, "ex_ResultTemplatesParser", "Template let expression is empty."
    End If

    openPos = InStr(1, normalizedExpression, "(", vbBinaryCompare)
    closePos = InStrRev(normalizedExpression, ")", -1, vbBinaryCompare)
    If openPos <= 1 Or closePos <= openPos Then
        Err.Raise vbObjectError + 1780, "ex_ResultTemplatesParser", "Unsupported template let expression: '" & normalizedExpression & "'."
    End If

    helperRef = mp_TrimWhitespace(Left$(normalizedExpression, openPos - 1))
    argsText = Mid$(normalizedExpression, openPos + 1, closePos - openPos - 1)

    args = mp_SplitLetExpressionArgs(argsText)

    If Left$(helperRef, 1) <> "$" Then
        Err.Raise vbObjectError + 1783, "ex_ResultTemplatesParser", _
            "Template let helper must use '$<MODULE>.<METHOD>(...)' syntax: '" & normalizedExpression & "'."
    End If

    mp_EvaluateTemplateLetExpression = mp_RunExternalTemplateHelper(helperRef, args)
End Function

Private Function mp_SplitLetExpressionArgs(ByVal argsText As String) As Variant
    Dim parts() As String
    Dim currentPart As String
    Dim ch As String
    Dim i As Long
    Dim normalizedPart As String
    Dim count As Long
    Dim inSingleQuote As Boolean
    Dim inDoubleQuote As Boolean

    argsText = mp_TrimWhitespace(CStr(argsText))
    If Len(argsText) = 0 Then
        mp_SplitLetExpressionArgs = Array()
        Exit Function
    End If

    ReDim parts(0 To Len(argsText))

    For i = 1 To Len(argsText)
        ch = Mid$(argsText, i, 1)
        If ch = "'" And Not inDoubleQuote Then
            inSingleQuote = Not inSingleQuote
            currentPart = currentPart & ch
            GoTo ContinueChar
        End If
        If ch = """" And Not inSingleQuote Then
            inDoubleQuote = Not inDoubleQuote
            currentPart = currentPart & ch
            GoTo ContinueChar
        End If

        If ch = "," And Not inSingleQuote And Not inDoubleQuote Then
            normalizedPart = mp_TrimWhitespace(currentPart)
            If Len(normalizedPart) > 0 Then
                parts(count) = normalizedPart
                count = count + 1
            End If
            currentPart = vbNullString
            GoTo ContinueChar
        End If

        currentPart = currentPart & ch
ContinueChar:
    Next i

    normalizedPart = mp_TrimWhitespace(currentPart)
    If Len(normalizedPart) > 0 Then
        parts(count) = normalizedPart
        count = count + 1
    End If

    If count = 0 Then
        mp_SplitLetExpressionArgs = Array()
        Exit Function
    End If

    ReDim Preserve parts(0 To count - 1)
    mp_SplitLetExpressionArgs = parts
End Function

Private Function mp_GetArrayItemCount(ByVal values As Variant) As Long
    On Error GoTo EmptyArray
    If IsArray(values) Then
        mp_GetArrayItemCount = UBound(values) - LBound(values) + 1
    End If
    Exit Function
EmptyArray:
    mp_GetArrayItemCount = 0
End Function

Private Function mp_RunExternalTemplateHelper(ByVal helperRef As String, ByVal args As Variant) As String
    Dim methodRef As String
    Dim argCount As Long
    Dim parsedArgs() As Variant
    Dim i As Long
    Dim invokeResult As Variant

    methodRef = mp_TrimWhitespace(CStr(helperRef))
    If Left$(methodRef, 1) = "$" Then methodRef = Mid$(methodRef, 2)
    methodRef = mp_TrimWhitespace(methodRef)

    If Len(methodRef) = 0 Then
        Err.Raise vbObjectError + 1784, "ex_ResultTemplatesParser", "Template helper reference is empty."
    End If
    If InStr(1, methodRef, ".", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 1785, "ex_ResultTemplatesParser", "Template helper must use '<MODULE>.<METHOD>' syntax: '" & helperRef & "'."
    End If

    argCount = mp_GetArrayItemCount(args)
    If argCount > 5 Then
        Err.Raise vbObjectError + 1786, "ex_ResultTemplatesParser", "Template helper supports at most 5 arguments: '" & helperRef & "'."
    End If

    If argCount > 0 Then
        ReDim parsedArgs(0 To argCount - 1)
        For i = 0 To argCount - 1
            mp_ValidateTemplateHelperArgumentRaw CStr(args(i)), helperRef, i + 1
            parsedArgs(i) = mp_ParseTemplateHelperArgument(CStr(args(i)))
        Next i
    End If

    Select Case argCount
        Case 0
            invokeResult = Application.Run(methodRef)
        Case 1
            invokeResult = Application.Run(methodRef, parsedArgs(0))
        Case 2
            invokeResult = Application.Run(methodRef, parsedArgs(0), parsedArgs(1))
        Case 3
            invokeResult = Application.Run(methodRef, parsedArgs(0), parsedArgs(1), parsedArgs(2))
        Case 4
            invokeResult = Application.Run(methodRef, parsedArgs(0), parsedArgs(1), parsedArgs(2), parsedArgs(3))
        Case 5
            invokeResult = Application.Run(methodRef, parsedArgs(0), parsedArgs(1), parsedArgs(2), parsedArgs(3), parsedArgs(4))
    End Select

    mp_RunExternalTemplateHelper = mp_NormalizeTemplateHelperResult(invokeResult, helperRef)
End Function

Private Function mp_ParseTemplateHelperArgument(ByVal argText As String) As Variant
    Dim normalized As String
    Dim numberValue As Double
    Dim unquoted As String

    normalized = mp_TrimWhitespace(CStr(argText))
    normalized = mp_RemoveInlineHighlightMarkers(normalized)
    normalized = mp_TrimWhitespace(normalized)
    If Len(normalized) = 0 Then
        mp_ParseTemplateHelperArgument = vbNullString
        Exit Function
    End If

    If (Left$(normalized, 1) = """" And Right$(normalized, 1) = """") Or _
       (Left$(normalized, 1) = "'" And Right$(normalized, 1) = "'") Then
        unquoted = Mid$(normalized, 2, Len(normalized) - 2)
        unquoted = mp_RemoveInlineHighlightMarkers(unquoted)
        mp_ParseTemplateHelperArgument = unquoted
        Exit Function
    End If

    Select Case LCase$(normalized)
        Case "true"
            mp_ParseTemplateHelperArgument = True
            Exit Function
        Case "false"
            mp_ParseTemplateHelperArgument = False
            Exit Function
    End Select

    If IsNumeric(normalized) Then
        numberValue = CDbl(normalized)
        If Fix(numberValue) = numberValue Then
            mp_ParseTemplateHelperArgument = CLng(numberValue)
        Else
            mp_ParseTemplateHelperArgument = numberValue
        End If
        Exit Function
    End If

    mp_ParseTemplateHelperArgument = normalized
End Function

Private Sub mp_ValidateTemplateHelperArgumentRaw( _
    ByVal rawArgText As String, _
    ByVal helperRef As String, _
    ByVal argIndex As Long _
)
    Dim normalized As String
    Dim checkText As String
    Dim placeholderToken As String

    normalized = mp_TrimWhitespace(CStr(rawArgText))
    If Len(normalized) = 0 Then Exit Sub

    ' Escaped braces are allowed: \{ and \}
    checkText = Replace(normalized, "\{", vbNullString)
    checkText = Replace(checkText, "\}", vbNullString)

    placeholderToken = mp_FindFirstPlaceholderLikeToken(checkText)
    If Len(placeholderToken) > 0 Then
        Err.Raise vbObjectError + 1810, "ex_ResultTemplatesParser", _
            "Template helper '" & helperRef & "' argument #" & CStr(argIndex) & _
            " contains unresolved placeholder '" & placeholderToken & "'. " & _
            "Ensure placeholders are resolved before '#let $<MODULE>.<METHOD>(...)'."
    End If

    If InStr(1, checkText, "{", vbBinaryCompare) > 0 Or InStr(1, checkText, "}", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 1811, "ex_ResultTemplatesParser", _
            "Template helper '" & helperRef & "' argument #" & CStr(argIndex) & _
            " contains unescaped '{' or '}'. Use '\{' and '\}' for literal braces."
    End If
End Sub

Private Function mp_FindFirstPlaceholderLikeToken(ByVal sourceText As String) As String
    Dim rx As Object
    Dim matches As Object

    sourceText = CStr(sourceText)
    If Len(sourceText) = 0 Then Exit Function

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = False
    rx.Pattern = "\{[A-Za-z_#][^{}]*\}"

    Set matches = rx.Execute(sourceText)
    If matches Is Nothing Then Exit Function
    If matches.Count = 0 Then Exit Function

    mp_FindFirstPlaceholderLikeToken = CStr(matches(0).Value)
End Function

Private Function mp_NormalizeTemplateHelperResult(ByVal resultValue As Variant, ByVal helperRef As String) As String
    If IsObject(resultValue) Then
        Err.Raise vbObjectError + 1787, "ex_ResultTemplatesParser", "Template helper '" & helperRef & "' returned object result, string/number/boolean expected."
    End If
    If IsError(resultValue) Then
        Err.Raise vbObjectError + 1788, "ex_ResultTemplatesParser", "Template helper '" & helperRef & "' returned error value."
    End If

    If VarType(resultValue) = vbBoolean Then
        mp_NormalizeTemplateHelperResult = mp_BooleanTextFromValue(CStr(resultValue))
        Exit Function
    End If

    If IsNull(resultValue) Then
        mp_NormalizeTemplateHelperResult = vbNullString
    Else
        mp_NormalizeTemplateHelperResult = CStr(resultValue)
    End If
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
    resultText = mp_ResolveJoinLineToken(resultText, RESERVED_JOINLINE_TOKEN_LEGACY)
    mp_ResolveJoinLineTokens = resultText
End Function

Private Function mp_ResolveTrimIndentTokens(ByVal sourceText As String) As String
    Dim resultText As String

    resultText = CStr(sourceText)
    resultText = mp_ResolveTrimIndentToken(resultText, RESERVED_TRIMINDENT_TOKEN_SHORT)
    mp_ResolveTrimIndentTokens = resultText
End Function

Private Function mp_ResolveTrimIndentToken(ByVal sourceText As String, ByVal tokenText As String) As String
    Dim resultText As String
    Dim rx As Object
    Dim tokenPattern As String

    resultText = CStr(sourceText)
    tokenPattern = mp_EscapeRegex(CStr(tokenText))

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = False

    ' Remove token and all horizontal whitespace after it.
    rx.Pattern = tokenPattern & "[ \t]*"
    resultText = rx.Replace(resultText, vbNullString)

    mp_ResolveTrimIndentToken = resultText
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
    Dim rx As Object
    Dim matches As Object
    Dim tokenStart As Long
    Dim tokenLen As Long
    Dim tokenOffsetText As String
    Dim prefixText As String
    Dim valueStart As Long
    Dim valueLen As Long
    Dim valueText As String
    Dim baseNumber As Double
    Dim baseIsInteger As Boolean
    Dim offsetValue As Double
    Dim resultNumber As Double
    Dim resolvedValue As String
    Dim leftPart As String
    Dim rightPart As String
    Dim resultText As String
    Dim hasHighlight As Boolean
    Dim highlightColorHex As String

    resultText = CStr(sourceText)
    baseDateText = CStr(baseDateText) ' reserved for signature compatibility.

    mp_EnsureLegacyDayTokensNotUsed resultText

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = False
    rx.Pattern = NUMERIC_OFFSET_TOKEN_PATTERN

    Do
        Set matches = rx.Execute(resultText)
        If matches Is Nothing Then Exit Do
        If matches.Count = 0 Then Exit Do

        tokenStart = CLng(matches(0).FirstIndex)
        tokenLen = CLng(matches(0).Length)
        tokenOffsetText = CStr(matches(0).SubMatches(0))
        prefixText = Left$(resultText, tokenStart)

        If Not mp_TryFindImmediateLeftValueSpan(prefixText, valueStart, valueLen, valueText) Then
            Err.Raise vbObjectError + 1764, "ex_ResultTemplatesParser", _
                "Numeric offset token '{" & tokenOffsetText & "}' has no left value."
        End If
        hasHighlight = mp_TryExtractInlineHighlightColor(valueText, highlightColorHex)

        If Not mp_TryParseTemplateNumeric(valueText, baseNumber, baseIsInteger) Then
            Err.Raise vbObjectError + 1765, "ex_ResultTemplatesParser", _
                "Left value '" & valueText & "' before token '{" & tokenOffsetText & "}' is not numeric."
        End If
        If Not mp_TryParseSignedInteger(tokenOffsetText, offsetValue) Then
            Err.Raise vbObjectError + 1766, "ex_ResultTemplatesParser", _
                "Invalid numeric offset '{" & tokenOffsetText & "}'. Use '{+N}' or '{-N}'."
        End If

        resultNumber = baseNumber + offsetValue
        If baseIsInteger And Fix(resultNumber) = resultNumber Then
            resolvedValue = mp_FormatIntegerWithBasePadding(CLng(resultNumber), valueText)
        Else
            resolvedValue = CStr(resultNumber)
        End If
        If hasHighlight Then
            resolvedValue = mp_WrapHighlightMarkers(resolvedValue, highlightColorHex)
        End If

        leftPart = Left$(resultText, valueStart - 1)
        rightPart = Mid$(resultText, tokenStart + tokenLen + 1)
        resultText = leftPart & resolvedValue & rightPart
    Loop

    mp_ResolveDateExpressions = resultText
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

Private Sub mp_EnsureLegacyDayTokensNotUsed(ByVal sourceText As String)
    Dim rx As Object
    Dim matches As Object

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = False
    rx.Pattern = LEGACY_DAY_TOKEN_PATTERN

    Set matches = rx.Execute(CStr(sourceText))
    If matches Is Nothing Then Exit Sub
    If matches.Count = 0 Then Exit Sub

    Err.Raise vbObjectError + 1763, "ex_ResultTemplatesParser", _
        "Legacy token '" & CStr(matches(0).Value) & "' is not supported. Use '<NUMERIC>{+N}' or '<NUMERIC>{-N}'."
End Sub

Private Function mp_TryFindImmediateLeftValueSpan( _
    ByVal textValue As String, _
    ByRef outStartPos As Long, _
    ByRef outLength As Long, _
    ByRef outValueText As String _
) As Boolean
    Dim endPos As Long
    Dim startPos As Long

    textValue = CStr(textValue)
    endPos = Len(textValue)

    Do While endPos > 0
        If Not mp_IsWhitespaceChar(Mid$(textValue, endPos, 1)) Then Exit Do
        endPos = endPos - 1
    Loop
    If endPos <= 0 Then Exit Function

    startPos = endPos
    Do While startPos > 1
        If mp_IsWhitespaceChar(Mid$(textValue, startPos - 1, 1)) Then Exit Do
        startPos = startPos - 1
    Loop

    outStartPos = startPos
    outLength = endPos - startPos + 1
    outValueText = Mid$(textValue, outStartPos, outLength)
    mp_TryFindImmediateLeftValueSpan = (outLength > 0)
End Function

Private Function mp_TryParseTemplateNumeric(ByVal numberText As String, ByRef outValue As Double, ByRef outIsInteger As Boolean) As Boolean
    Dim rx As Object
    Dim normalized As String

    numberText = mp_RemoveInlineHighlightMarkers(CStr(numberText))
    numberText = mp_TrimWhitespace(CStr(numberText))
    If Len(numberText) = 0 Then Exit Function

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = False
    rx.Pattern = "^[+-]?\d+(?:[.,]\d+)?$"

    If Not rx.Test(numberText) Then Exit Function

    normalized = Replace(numberText, ",", ".")
    outValue = Val(normalized)
    outIsInteger = (InStr(1, normalized, ".", vbBinaryCompare) = 0)
    mp_TryParseTemplateNumeric = True
End Function

Private Function mp_RemoveInlineHighlightMarkers(ByVal sourceText As String) As String
    Dim rxStart As Object
    Dim resultText As String

    resultText = CStr(sourceText)
    If Len(resultText) = 0 Then
        mp_RemoveInlineHighlightMarkers = vbNullString
        Exit Function
    End If

    Set rxStart = CreateObject("VBScript.RegExp")
    rxStart.Global = True
    rxStart.IgnoreCase = False
    rxStart.Pattern = mp_EscapeRegex(HIGHLIGHT_MARKER_START) & "(\([^)]*\))?"
    resultText = rxStart.Replace(resultText, vbNullString)
    resultText = Replace(resultText, HIGHLIGHT_MARKER_END, vbNullString)

    mp_RemoveInlineHighlightMarkers = resultText
End Function

Private Function mp_TryExtractInlineHighlightColor(ByVal sourceText As String, ByRef outColorHex As String) As Boolean
    Dim rx As Object
    Dim matches As Object

    outColorHex = vbNullString
    sourceText = CStr(sourceText)
    If Len(sourceText) = 0 Then Exit Function
    If InStr(1, sourceText, HIGHLIGHT_MARKER_START, vbBinaryCompare) = 0 Then Exit Function
    If InStr(1, sourceText, HIGHLIGHT_MARKER_END, vbBinaryCompare) = 0 Then Exit Function

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = False
    rx.Pattern = mp_EscapeRegex(HIGHLIGHT_MARKER_START) & "(?:\(([^)]*)\))?"

    Set matches = rx.Execute(sourceText)
    If matches Is Nothing Then Exit Function
    If matches.Count = 0 Then Exit Function

    outColorHex = mp_TrimWhitespace(CStr(matches(0).SubMatches(0)))
    If Len(outColorHex) = 0 Then outColorHex = DEFAULT_HIGHLIGHT_COLOR_HEX
    mp_TryExtractInlineHighlightColor = True
End Function

Private Function mp_TryParseSignedInteger(ByVal numberText As String, ByRef outValue As Double) As Boolean
    Dim rx As Object

    numberText = mp_TrimWhitespace(CStr(numberText))
    If Len(numberText) = 0 Then Exit Function

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = False
    rx.Pattern = "^[+-]\d+$"

    If Not rx.Test(numberText) Then Exit Function

    outValue = CDbl(numberText)
    mp_TryParseSignedInteger = True
End Function

Private Function mp_FormatIntegerWithBasePadding(ByVal resultValue As Long, ByVal baseNumberText As String) As String
    Dim signText As String
    Dim absResult As Long
    Dim baseDigits As String
    Dim width As Long
    Dim paddedText As String
    Dim shouldPad As Boolean

    signText = vbNullString
    If resultValue < 0 Then signText = "-"

    absResult = resultValue
    If absResult < 0 Then absResult = -absResult

    baseDigits = mp_TrimWhitespace(CStr(baseNumberText))
    If Left$(baseDigits, 1) = "+" Or Left$(baseDigits, 1) = "-" Then
        baseDigits = Mid$(baseDigits, 2)
    End If

    width = Len(baseDigits)
    shouldPad = (width > 1 And Left$(baseDigits, 1) = "0")
    If shouldPad Then
        paddedText = Format$(absResult, String$(width, "0"))
    Else
        paddedText = CStr(absResult)
    End If

    mp_FormatIntegerWithBasePadding = signText & paddedText
End Function
