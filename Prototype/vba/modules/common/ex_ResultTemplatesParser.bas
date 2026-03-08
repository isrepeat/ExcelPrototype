Attribute VB_Name = "ex_ResultTemplatesParser"
Option Explicit

Private Const PROFILES_NS As String = "urn:excelprototype:profiles"
Private Const RESULT_TEMPLATES_REL_PATH As String = "config\modes\PersonalCard\PersonalCardResultTemplates.xml"
' Reserved date tokens:
' - {#dd}, {#dd+N}, {#dd-N}
Private Const DATE_TOKEN_PATTERN As String = "\{#dd(?:[+-]\d+)?\}"

Public Function m_GetTemplateText(ByVal templateId As String) As String
    Dim doc As Object
    Dim node As Object
    Dim xpath As String
    Dim templateText As String

    templateId = Trim$(templateId)
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
End Function

Public Function m_ReplaceToken( _
    ByVal sourceText As String, _
    ByVal tokenText As String, _
    ByVal replacementText As String _
) As String
    m_ReplaceToken = Replace(CStr(sourceText), CStr(tokenText), CStr(replacementText))
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

    normalizedName = Trim$(placeholderName)
    If Len(normalizedName) = 0 Then
        m_ReplacePlaceholder = CStr(sourceText)
        Exit Function
    End If

    placeholderToken = "{" & normalizedName & "}"
    resultText = m_ReplaceToken(CStr(sourceText), placeholderToken, CStr(replacementText))

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = False
    rx.Pattern = "\{" & mp_EscapeRegex(normalizedName) & "\|([^{}|]+)\}"

    Set matches = rx.Execute(resultText)
    If Not matches Is Nothing Then
        For i = matches.Count - 1 To 0 Step -1
            formatter = CStr(matches(i).SubMatches(0))
            formattedValue = mp_ApplyFormatter(CStr(replacementText), formatter)
            matchStart = CLng(matches(i).FirstIndex)
            matchLen = CLng(matches(i).Length)

            resultText = Left$(resultText, matchStart) & formattedValue & Mid$(resultText, matchStart + matchLen + 1)
        Next i
    End If

    m_ReplacePlaceholder = resultText
End Function

Public Function m_ResolveTemplate( _
    ByVal sourceText As String, _
    Optional ByVal baseDateText As String = vbNullString _
) As String
    ' Final pass for template text (reserved date tokens and future finalizers).
    m_ResolveTemplate = mp_ResolveDateExpressions(CStr(sourceText), baseDateText)
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

    normalizedFormatter = LCase$(Trim$(formatterName))

    Select Case normalizedFormatter
        Case "upper"
            mp_ApplyFormatter = UCase$(CStr(sourceValue))
        Case "lower"
            mp_ApplyFormatter = LCase$(CStr(sourceValue))
        Case "capitalize"
            mp_ApplyFormatter = mp_CapitalizeText(CStr(sourceValue))
        Case "firstchar"
            mp_ApplyFormatter = mp_FirstNonSpaceChar(CStr(sourceValue))
        Case "genitive"
            mp_ApplyFormatter = mp_InflectPhraseToCase(CStr(sourceValue), "genitive")
        Case "accusative"
            mp_ApplyFormatter = mp_InflectPhraseToCase(CStr(sourceValue), "accusative")
        Case Else
            Err.Raise vbObjectError + 1766, "ex_ResultTemplatesParser", _
                "Unsupported formatter '" & formatterName & "'."
    End Select
End Function

Private Function mp_FirstNonSpaceChar(ByVal textValue As String) As String
    Dim i As Long
    Dim ch As String

    textValue = CStr(textValue)
    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then
            mp_FirstNonSpaceChar = ch
            Exit Function
        End If
    Next i
End Function

Private Function mp_InflectPhraseToCase(ByVal sourceValue As String, ByVal caseName As String) As String
    Dim i As Long
    Dim token As String
    Dim ch As String
    Dim resultText As String

    sourceValue = CStr(sourceValue)
    token = vbNullString
    resultText = vbNullString

    For i = 1 To Len(sourceValue)
        ch = Mid$(sourceValue, i, 1)
        If mp_IsWordChar(ch) Then
            token = token & ch
        Else
            If Len(token) > 0 Then
                resultText = resultText & mp_InflectWordByCase(token, caseName)
                token = vbNullString
            End If
            resultText = resultText & ch
        End If
    Next i

    If Len(token) > 0 Then
        resultText = resultText & mp_InflectWordByCase(token, caseName)
    End If

    mp_InflectPhraseToCase = resultText
End Function

Private Function mp_InflectWordByCase(ByVal wordText As String, ByVal caseName As String) As String
    Dim hyphenPos As Long
    Dim leftPart As String
    Dim rightPart As String
    Dim lowWord As String
    Dim inflectedLow As String

    wordText = CStr(wordText)
    If Len(wordText) = 0 Then Exit Function

    hyphenPos = InStr(1, wordText, "-", vbBinaryCompare)
    If hyphenPos > 1 And hyphenPos < Len(wordText) Then
        leftPart = Left$(wordText, hyphenPos - 1)
        rightPart = Mid$(wordText, hyphenPos + 1)
        mp_InflectWordByCase = mp_InflectWordByCase(leftPart, caseName) & "-" & mp_InflectWordByCase(rightPart, caseName)
        Exit Function
    End If

    lowWord = LCase$(wordText)
    If caseName = "genitive" Then
        inflectedLow = mp_InflectWordToGenitiveLow(lowWord)
    ElseIf caseName = "accusative" Then
        inflectedLow = mp_InflectWordToAccusativeLow(lowWord)
    Else
        inflectedLow = lowWord
    End If

    mp_InflectWordByCase = mp_ApplyWordCasePattern(wordText, inflectedLow)
End Function

Private Function mp_InflectWordToGenitiveLow(ByVal lowWord As String) As String
    Select Case lowWord
        Case "старший": mp_InflectWordToGenitiveLow = "старшого": Exit Function
        Case "молодший": mp_InflectWordToGenitiveLow = "молодшого": Exit Function
        Case "головний": mp_InflectWordToGenitiveLow = "головного": Exit Function
        Case "черговий": mp_InflectWordToGenitiveLow = "чергового": Exit Function
        Case "лейтенант": mp_InflectWordToGenitiveLow = "лейтенанта": Exit Function
        Case "капітан": mp_InflectWordToGenitiveLow = "капітана": Exit Function
        Case "майор": mp_InflectWordToGenitiveLow = "майора": Exit Function
        Case "полковник": mp_InflectWordToGenitiveLow = "полковника": Exit Function
        Case "сержант": mp_InflectWordToGenitiveLow = "сержанта": Exit Function
        Case "солдат": mp_InflectWordToGenitiveLow = "солдата": Exit Function
        Case "командир": mp_InflectWordToGenitiveLow = "командира": Exit Function
    End Select

    mp_InflectWordToGenitiveLow = mp_InflectWordByHeuristics(lowWord, "genitive")
End Function

Private Function mp_InflectWordToAccusativeLow(ByVal lowWord As String) As String
    Select Case lowWord
        Case "старший": mp_InflectWordToAccusativeLow = "старшого": Exit Function
        Case "молодший": mp_InflectWordToAccusativeLow = "молодшого": Exit Function
        Case "головний": mp_InflectWordToAccusativeLow = "головного": Exit Function
        Case "черговий": mp_InflectWordToAccusativeLow = "чергового": Exit Function
        Case "лейтенант": mp_InflectWordToAccusativeLow = "лейтенанта": Exit Function
        Case "капітан": mp_InflectWordToAccusativeLow = "капітана": Exit Function
        Case "майор": mp_InflectWordToAccusativeLow = "майора": Exit Function
        Case "полковник": mp_InflectWordToAccusativeLow = "полковника": Exit Function
        Case "сержант": mp_InflectWordToAccusativeLow = "сержанта": Exit Function
        Case "солдат": mp_InflectWordToAccusativeLow = "солдата": Exit Function
        Case "командир": mp_InflectWordToAccusativeLow = "командира": Exit Function
    End Select

    mp_InflectWordToAccusativeLow = mp_InflectWordByHeuristics(lowWord, "accusative")
End Function

Private Function mp_InflectWordByHeuristics(ByVal lowWord As String, ByVal caseName As String) As String
    Dim last1 As String
    Dim last2 As String
    Dim last3 As String
    Dim last4 As String

    If Len(lowWord) = 0 Then Exit Function
    If Len(lowWord) <= 2 Then
        mp_InflectWordByHeuristics = lowWord
        Exit Function
    End If

    last1 = Right$(lowWord, 1)
    last2 = Right$(lowWord, 2)
    last3 = Right$(lowWord, 3)
    last4 = Right$(lowWord, 4)

    If caseName = "genitive" Then
        If last2 = "ко" Then
            mp_InflectWordByHeuristics = lowWord
            Exit Function
        End If

        If last2 = "ий" Then
            mp_InflectWordByHeuristics = Left$(lowWord, Len(lowWord) - 2) & "ого"
            Exit Function
        End If
        If last2 = "ій" Then
            mp_InflectWordByHeuristics = Left$(lowWord, Len(lowWord) - 2) & "ього"
            Exit Function
        End If
        If last1 = "а" Then
            mp_InflectWordByHeuristics = Left$(lowWord, Len(lowWord) - 1) & "и"
            Exit Function
        End If
        If last1 = "я" Then
            mp_InflectWordByHeuristics = Left$(lowWord, Len(lowWord) - 1) & "і"
            Exit Function
        End If
        If last1 = "ь" Then
            mp_InflectWordByHeuristics = Left$(lowWord, Len(lowWord) - 1) & "я"
            Exit Function
        End If
        If last1 = "о" Then
            mp_InflectWordByHeuristics = lowWord
            Exit Function
        End If

        If last3 = "ич" Or last4 = "ович" Or last4 = "евич" Then
            mp_InflectWordByHeuristics = lowWord & "а"
            Exit Function
        End If

        mp_InflectWordByHeuristics = lowWord & "а"
        Exit Function
    End If

    If caseName = "accusative" Then
        If last2 = "ко" Then
            mp_InflectWordByHeuristics = lowWord
            Exit Function
        End If

        If last2 = "ий" Then
            mp_InflectWordByHeuristics = Left$(lowWord, Len(lowWord) - 2) & "ого"
            Exit Function
        End If
        If last2 = "ій" Then
            mp_InflectWordByHeuristics = Left$(lowWord, Len(lowWord) - 2) & "ього"
            Exit Function
        End If
        If last1 = "а" Then
            mp_InflectWordByHeuristics = Left$(lowWord, Len(lowWord) - 1) & "у"
            Exit Function
        End If
        If last1 = "я" Then
            mp_InflectWordByHeuristics = Left$(lowWord, Len(lowWord) - 1) & "ю"
            Exit Function
        End If
        If last1 = "о" Then
            mp_InflectWordByHeuristics = lowWord
            Exit Function
        End If
        If last1 = "ь" Then
            mp_InflectWordByHeuristics = Left$(lowWord, Len(lowWord) - 1) & "я"
            Exit Function
        End If

        If last3 = "ич" Or last4 = "ович" Or last4 = "евич" Then
            mp_InflectWordByHeuristics = lowWord & "а"
            Exit Function
        End If

        mp_InflectWordByHeuristics = lowWord & "а"
        Exit Function
    End If

    mp_InflectWordByHeuristics = lowWord
End Function

Private Function mp_ApplyWordCasePattern(ByVal sourceWord As String, ByVal inflectedLower As String) As String
    If sourceWord = UCase$(sourceWord) Then
        mp_ApplyWordCasePattern = UCase$(inflectedLower)
        Exit Function
    End If

    If sourceWord = LCase$(sourceWord) Then
        mp_ApplyWordCasePattern = LCase$(inflectedLower)
        Exit Function
    End If

    mp_ApplyWordCasePattern = mp_CapitalizeText(inflectedLower)
End Function

Private Function mp_IsWordChar(ByVal ch As String) As Boolean
    Dim codePoint As Long

    If ch = "'" Then
        mp_IsWordChar = True
        Exit Function
    End If

    codePoint = AscW(ch)
    mp_IsWordChar = _
        (codePoint >= 48 And codePoint <= 57) Or _
        (codePoint >= 65 And codePoint <= 90) Or _
        (codePoint >= 97 And codePoint <= 122) Or _
        (codePoint >= 1024 And codePoint <= 1279)
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

    baseDateText = Trim$(baseDateText)
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
    If Left$(innerText, 1) <> "#" Then
        Err.Raise vbObjectError + 1765, "ex_ResultTemplatesParser", _
            "Unsupported reserved token '" & tokenText & "'. Reserved tokens must start with '#'."
    End If
    innerText = Mid$(innerText, 2)

    If LCase$(Left$(innerText, 2)) <> "dd" Then
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
    mp_ResolveDateToken = Format$(resolvedDate, "dd")
    Exit Function

ParseError:
    Err.Raise vbObjectError + 1764, "ex_ResultTemplatesParser", _
        "Invalid date token '" & tokenText & "'. Use {#dd}, {#dd+N}, {#dd-N}."
End Function
