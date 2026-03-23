Attribute VB_Name = "ex_MorphUaLite"
Option Explicit

' ============================================
' Ukrainian morphology (lite, VBA-only)
' ============================================

Private Const CASE_GENITIVE As String = "genitive"
Private Const CASE_ACCUSATIVE As String = "accusative"
Private Const CASE_DATIVE As String = "dative"
Private Const USE_PURE_FIO_CONVERTERS As Boolean = True
Private Const USE_SELECTION_TEXT_CONVERTERS As Boolean = False
Private Const USE_ENHANCED_PHRASE_INFLECTION As Boolean = True
Private Const USE_LEGACY_PHRASE_FALLBACK As Boolean = False
Private Const ENHANCED_MAX_WORDS_PER_SEGMENT As Long = 4
Private Const NBSP_CODE_POINT As Long = 160
Private Const NARROW_NBSP_CODE_POINT As Long = 8239

Public Function m_TryConvertSelectionTextToGenitive(ByVal normalizedText As String, ByRef convertedText As String) As Boolean
    m_TryConvertSelectionTextToGenitive = mp_TryConvertSelectionTextToGenitive(normalizedText, convertedText)
End Function

Public Function m_TryConvertSelectionTextToAccusative(ByVal normalizedText As String, ByRef convertedText As String) As Boolean
    m_TryConvertSelectionTextToAccusative = mp_TryConvertSelectionTextToAccusative(normalizedText, convertedText)
End Function

Public Function m_TryConvertSelectionTextToDative(ByVal normalizedText As String, ByRef convertedText As String) As Boolean
    m_TryConvertSelectionTextToDative = mp_TryConvertSelectionTextToDative(normalizedText, convertedText)
End Function

Public Function m_TryParseSentenceWithFio(ByVal normalizedText As String, ByRef leadPhrase As String, ByRef surname As String, ByRef firstName As String, ByRef patronymic As String, ByRef tailPhrase As String) As Boolean
    m_TryParseSentenceWithFio = mp_TryParseSentenceWithFio(normalizedText, leadPhrase, surname, firstName, patronymic, tailPhrase)
End Function

Public Function m_TryParseFio(ByVal normalizedText As String, ByRef surname As String, ByRef firstName As String, ByRef patronymic As String) As Boolean
    m_TryParseFio = mp_TryParseFio(normalizedText, surname, firstName, patronymic)
End Function

Public Function m_NormalizeFioInput(ByVal inputText As String) As String
    m_NormalizeFioInput = mp_NormalizeFioInput(inputText)
End Function

Public Function m_IsValidFioToken(ByVal token As String) As Boolean
    m_IsValidFioToken = mp_IsValidFioToken(token)
End Function

Public Function m_DetectFioGender(ByVal firstName As String, ByVal patronymic As String, ByRef gender As String) As Boolean
    m_DetectFioGender = mp_DetectFioGender(firstName, patronymic, gender)
End Function

Public Function m_TrimTokenPunctuation(ByVal token As String) As String
    m_TrimTokenPunctuation = mp_TrimTokenPunctuation(token)
End Function

Public Function m_JoinArraySlice(ByRef parts() As String, ByVal startIndex As Long, Optional ByVal endIndex As Long = -1) As String
    m_JoinArraySlice = mp_JoinArraySlice(parts, startIndex, endIndex)
End Function

Public Function m_ToTitleCaseWord(ByVal textValue As String) As String
    m_ToTitleCaseWord = mp_ToTitleCaseWord(textValue)
End Function

Public Function m_ToShortFioNormalized(ByVal sourceText As String) As String
    Dim surname As String
    Dim initials As String

    surname = m_ToFioSurnameNormalized(sourceText)
    initials = m_ToFioInitials(sourceText)

    m_ToShortFioNormalized = mp_ToTitleCaseWord(surname)
    If Len(initials) > 0 Then
        If Len(m_ToShortFioNormalized) > 0 Then
            m_ToShortFioNormalized = m_ToShortFioNormalized & " " & initials
        Else
            m_ToShortFioNormalized = initials
        End If
    End If
End Function

Public Function m_ToFioSurnameNormalized(ByVal sourceText As String) As String
    Dim normalized As String
    Dim surname As String
    Dim firstName As String
    Dim patronymic As String

    normalized = mp_NormalizeFioInput(CStr(sourceText))
    If Len(normalized) = 0 Then Exit Function

    If Not mp_TryParseFio(normalized, surname, firstName, patronymic) Then
        m_ToFioSurnameNormalized = normalized
        Exit Function
    End If

    m_ToFioSurnameNormalized = surname
End Function

Public Function m_ToFioInitials(ByVal sourceText As String) As String
    Dim normalized As String
    Dim surname As String
    Dim firstName As String
    Dim patronymic As String
    Dim firstInitial As String
    Dim patronymicInitial As String

    normalized = mp_NormalizeFioInput(CStr(sourceText))
    If Len(normalized) = 0 Then Exit Function

    If Not mp_TryParseFio(normalized, surname, firstName, patronymic) Then
        Exit Function
    End If

    If Len(firstName) > 0 Then firstInitial = UCase$(Left$(firstName, 1))
    If Len(patronymic) > 0 Then patronymicInitial = UCase$(Left$(patronymic, 1))

    If Len(firstInitial) > 0 Then
        m_ToFioInitials = firstInitial & "."
    End If
    If Len(patronymicInitial) > 0 Then
        m_ToFioInitials = m_ToFioInitials & patronymicInitial & "."
    End If
End Function

Public Function m_LowercaseFirstLetter(ByVal sourceText As String) As String
    m_LowercaseFirstLetter = mp_LowercaseFirstLetter(sourceText)
End Function

Public Function m_IsWhitespaceLikeChar(ByVal ch As String) As Boolean
    m_IsWhitespaceLikeChar = mp_IsWhitespaceLikeChar(ch)
End Function

Public Sub m_SplitTrailingLineBreaks(ByVal sourceText As String, ByRef bodyText As String, ByRef trailingBreaks As String)
    mp_SplitTrailingLineBreaks sourceText, bodyText, trailingBreaks
End Sub

Public Function m_InflectPhraseToCase(ByVal sourceText As String, ByVal targetCase As String) As String
    Dim caseName As String
    Dim normalizedText As String
    Dim convertedText As String
    Dim fallbackText As String

    sourceText = CStr(sourceText)
    caseName = LCase$(Trim$(targetCase))
    If caseName <> CASE_GENITIVE And caseName <> CASE_ACCUSATIVE And caseName <> CASE_DATIVE Then
        m_InflectPhraseToCase = sourceText
        Exit Function
    End If

    normalizedText = m_NormalizeFioInput(sourceText)
    If Len(normalizedText) = 0 Then
        m_InflectPhraseToCase = sourceText
        Exit Function
    End If

    If mp_ShouldSkipPhraseInflection(normalizedText) Then
        m_InflectPhraseToCase = normalizedText
        Exit Function
    End If

    If USE_PURE_FIO_CONVERTERS Then
        If caseName = CASE_GENITIVE Then
            If mp_TryConvertPureFioToGenitive(normalizedText, convertedText) Then
                m_InflectPhraseToCase = convertedText
                Exit Function
            End If
        ElseIf caseName = CASE_ACCUSATIVE Then
            If mp_TryConvertPureFioToAccusative(normalizedText, convertedText) Then
                m_InflectPhraseToCase = convertedText
                Exit Function
            End If
        ElseIf caseName = CASE_DATIVE Then
            If mp_TryConvertPureFioToDative(normalizedText, convertedText) Then
                m_InflectPhraseToCase = convertedText
                Exit Function
            End If
        End If
    End If

    If USE_SELECTION_TEXT_CONVERTERS Then
        If caseName = CASE_GENITIVE Then
            If m_TryConvertSelectionTextToGenitive(normalizedText, convertedText) Then
                m_InflectPhraseToCase = convertedText
                Exit Function
            End If
        ElseIf caseName = CASE_ACCUSATIVE Then
            If m_TryConvertSelectionTextToAccusative(normalizedText, convertedText) Then
                m_InflectPhraseToCase = convertedText
                Exit Function
            End If
        ElseIf caseName = CASE_DATIVE Then
            If m_TryConvertSelectionTextToDative(normalizedText, convertedText) Then
                m_InflectPhraseToCase = convertedText
                Exit Function
            End If
        End If
    End If

    If USE_ENHANCED_PHRASE_INFLECTION Then
        fallbackText = mp_InflectPhraseByDashSegmentsToCaseEnhanced(normalizedText, caseName, ENHANCED_MAX_WORDS_PER_SEGMENT)
        If Len(fallbackText) > 0 Then
            m_InflectPhraseToCase = fallbackText
            Exit Function
        End If
    End If

    If USE_LEGACY_PHRASE_FALLBACK Then
        fallbackText = m_InflectPhraseByDashSegmentsToCase(normalizedText, caseName, 8)
        If Len(fallbackText) > 0 Then
            m_InflectPhraseToCase = fallbackText
            Exit Function
        End If
    End If

    m_InflectPhraseToCase = normalizedText
End Function

Public Function m_InflectPhraseByDashSegmentsToCase(ByVal phraseText As String, ByVal targetCase As String, ByVal maxWordsPerSegment As Long) As String
    Dim normalized As String
    normalized = Trim$(phraseText)
    If Len(normalized) = 0 Then Exit Function

    Dim caseName As String
    caseName = LCase$(Trim$(targetCase))
    If caseName <> CASE_GENITIVE And caseName <> CASE_ACCUSATIVE And caseName <> CASE_DATIVE Then Exit Function

    If maxWordsPerSegment <= 0 Then maxWordsPerSegment = 4

    normalized = mp_NormalizeSpacedDashSeparators(normalized)

    Dim segments() As String
    segments = Split(normalized, " - ")

    Dim i As Long
    For i = LBound(segments) To UBound(segments)
        segments(i) = mp_InflectLeadingWordsInSegmentToCase(Trim$(segments(i)), maxWordsPerSegment, caseName)
    Next i

    ' Word Ctrl+NumpadMinus -> en dash (U+2013).
    m_InflectPhraseByDashSegmentsToCase = Join(segments, " – ")
End Function

Private Function mp_NormalizeSpacedDashSeparators(ByVal textValue As String) As String
    Static regex As Object

    If regex Is Nothing Then
        Set regex = CreateObject("VBScript.RegExp")
        regex.Global = True
        regex.IgnoreCase = True
        regex.Pattern = "\s+[—–-]\s+"
    End If

    mp_NormalizeSpacedDashSeparators = regex.Replace(textValue, " - ")
End Function

Private Function mp_InflectLeadingWordsInSegmentToCase(ByVal segmentText As String, ByVal maxWordsToInflect As Long, ByVal targetCase As String) As String
    If Len(segmentText) = 0 Then Exit Function

    Dim parts() As String
    parts = Split(segmentText, " ")

    Dim i As Long
    Dim changedCount As Long
    Dim core As String
    Dim suffix As String
    Dim inflected As String
    Dim lowCore As String
    Dim lowPrevCore As String
    Dim candidateCount As Long

    For i = LBound(parts) To UBound(parts)
        If changedCount >= maxWordsToInflect Then Exit For
        If i > 10 Then Exit For

        core = mp_TrimTokenPunctuation(parts(i))
        If Len(core) = 0 Then GoTo ContinueLoop
        If Not mp_IsValidGeneralWord(core) Then GoTo ContinueLoop
        If IsNumeric(core) Then GoTo ContinueLoop
        If mp_IsStopWord(core) Then GoTo ContinueLoop

        If i > LBound(parts) Then
            lowPrevCore = LCase$(mp_TrimTokenPunctuation(parts(i - 1)))
        Else
            lowPrevCore = vbNullString
        End If

        candidateCount = candidateCount + 1
        If candidateCount > maxWordsToInflect Then Exit For

        lowCore = LCase$(core)

        If mp_ShouldKeepWordUnchangedByContext(lowPrevCore, lowCore, targetCase) Then GoTo ContinueLoop

        suffix = mp_TokenTailPunctuation(parts(i))
        If mp_InflectCommonWordToCase(core, targetCase, inflected) Then
            parts(i) = inflected & suffix
            If LCase$(inflected) <> LCase$(core) Then
                changedCount = changedCount + 1
            End If
        End If

ContinueLoop:
    Next i

    mp_InflectLeadingWordsInSegmentToCase = Join(parts, " ")
End Function

Private Function mp_InflectPhraseByDashSegmentsToCaseEnhanced(ByVal phraseText As String, ByVal targetCase As String, ByVal maxWordsPerSegment As Long) As String
    Dim normalized As String
    Dim caseName As String
    Dim segments() As String
    Dim i As Long

    normalized = Trim$(phraseText)
    If Len(normalized) = 0 Then Exit Function

    caseName = LCase$(Trim$(targetCase))
    If caseName <> CASE_GENITIVE And caseName <> CASE_ACCUSATIVE And caseName <> CASE_DATIVE Then Exit Function

    If maxWordsPerSegment <= 0 Then maxWordsPerSegment = ENHANCED_MAX_WORDS_PER_SEGMENT

    normalized = mp_NormalizeSpacedDashSeparators(normalized)
    segments = Split(normalized, " - ")

    For i = LBound(segments) To UBound(segments)
        segments(i) = mp_InflectWordsInSegmentToCaseEnhanced(Trim$(segments(i)), maxWordsPerSegment, caseName)
    Next i

    mp_InflectPhraseByDashSegmentsToCaseEnhanced = Join(segments, " – ")
End Function

Private Function mp_InflectWordsInSegmentToCaseEnhanced(ByVal segmentText As String, ByVal maxWordsToInflect As Long, ByVal targetCase As String) As String
    Dim parts() As String
    Dim i As Long
    Dim changedCount As Long
    Dim candidateCount As Long
    Dim tokenText As String
    Dim prefix As String
    Dim core As String
    Dim suffix As String
    Dim inflected As String
    Dim lowCore As String
    Dim lowPrevCore As String

    If Len(segmentText) = 0 Then Exit Function

    parts = Split(segmentText, " ")
    For i = LBound(parts) To UBound(parts)
        If changedCount >= maxWordsToInflect Then Exit For
        If i > 10 Then Exit For

        tokenText = CStr(parts(i))
        If Len(tokenText) = 0 Then GoTo ContinueLoop

        prefix = vbNullString
        core = vbNullString
        suffix = vbNullString
        mp_SplitTokenDecorations tokenText, prefix, core, suffix
        If Len(core) = 0 Then GoTo ContinueLoop
        If Not mp_IsValidGeneralWord(core) Then GoTo ContinueLoop
        If IsNumeric(core) Then GoTo ContinueLoop
        If mp_IsStopWord(core) Then GoTo ContinueLoop

        If i > LBound(parts) Then
            lowPrevCore = LCase$(mp_TrimTokenPunctuation(parts(i - 1)))
        Else
            lowPrevCore = vbNullString
        End If

        candidateCount = candidateCount + 1
        If candidateCount > maxWordsToInflect Then Exit For

        lowCore = LCase$(core)
        If mp_ShouldKeepWordUnchangedByContext(lowPrevCore, lowCore, targetCase) Then GoTo ContinueLoop

        If mp_InflectCommonWordToCase(core, targetCase, inflected) Then
            parts(i) = prefix & inflected & suffix
            If StrComp(inflected, core, vbTextCompare) <> 0 Then
                changedCount = changedCount + 1
            End If
        End If

ContinueLoop:
    Next i

    mp_InflectWordsInSegmentToCaseEnhanced = Join(parts, " ")
End Function

Private Function mp_ShouldSkipPhraseInflection(ByVal phraseText As String) As Boolean
    Dim normalizedPhrase As String
    normalizedPhrase = Trim$(phraseText)
    If Len(normalizedPhrase) = 0 Then Exit Function

    Dim parts() As String
    parts = Split(normalizedPhrase, " ")
    If UBound(parts) < 0 Then Exit Function

    Dim firstWord As String
    firstWord = LCase$(mp_TrimTokenPunctuation(parts(0)))

    Select Case firstWord
        Case "який", "яка", "яке", "які", "якого", "якої", "якому", "яким", "якій", "якими", "яких"
            mp_ShouldSkipPhraseInflection = True
    End Select
End Function

Private Sub mp_SplitTokenDecorations( _
    ByVal tokenText As String, _
    ByRef outPrefix As String, _
    ByRef outCore As String, _
    ByRef outSuffix As String _
)
    Dim startPos As Long
    Dim endPos As Long
    Dim ch As String

    outPrefix = vbNullString
    outCore = vbNullString
    outSuffix = vbNullString

    tokenText = CStr(tokenText)
    startPos = 1
    endPos = Len(tokenText)

    Do While startPos <= endPos
        ch = Mid$(tokenText, startPos, 1)
        If Not mp_IsTokenPrefixDecoration(ch) Then Exit Do
        outPrefix = outPrefix & ch
        startPos = startPos + 1
    Loop

    Do While endPos >= startPos
        ch = Mid$(tokenText, endPos, 1)
        If Not mp_IsTokenSuffixDecoration(ch) Then Exit Do
        outSuffix = ch & outSuffix
        endPos = endPos - 1
    Loop

    If startPos <= endPos Then
        outCore = Mid$(tokenText, startPos, endPos - startPos + 1)
    Else
        outCore = vbNullString
    End If
End Sub

Private Function mp_IsTokenPrefixDecoration(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    mp_IsTokenPrefixDecoration = (InStr(".,;:!?()[]{}""«»", ch) > 0)
End Function

Private Function mp_IsTokenSuffixDecoration(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    mp_IsTokenSuffixDecoration = (InStr(".,;:!?()[]{}""«»", ch) > 0)
End Function

Private Function mp_InflectCommonWordToCase(ByVal sourceWord As String, ByVal targetCase As String, ByRef resultWord As String) As Boolean
    If Not mp_IsValidGeneralWord(sourceWord) Then Exit Function

    If InStr(sourceWord, "-") > 0 Then
        mp_InflectCommonWordToCase = mp_TryInflectHyphenCommonWordToCase(sourceWord, targetCase, resultWord)
        Exit Function
    End If

    Select Case LCase$(Trim$(targetCase))
        Case CASE_ACCUSATIVE
            mp_InflectCommonWordToCase = mp_TryInflectSimpleCommonWordToAccusative(sourceWord, resultWord)
        Case CASE_DATIVE
            mp_InflectCommonWordToCase = mp_TryInflectSimpleCommonWordToDative(sourceWord, resultWord)
        Case Else
            mp_InflectCommonWordToCase = mp_TryInflectSimpleCommonWordToGenitive(sourceWord, resultWord)
    End Select
End Function

Private Function mp_TryInflectHyphenCommonWordToCase(ByVal sourceWord As String, ByVal targetCase As String, ByRef resultWord As String) As Boolean
    Dim parts() As String
    parts = Split(sourceWord, "-")
    If UBound(parts) < 1 Then Exit Function

    Dim i As Long
    Dim partInflected As String
    Dim currentPart As String
    Dim isFixedFirstPart As Boolean

    isFixedFirstPart = mp_IsFixedHyphenFirstPart(parts(0))

    For i = LBound(parts) To UBound(parts)
        currentPart = parts(i)
        partInflected = currentPart

        If Len(currentPart) > 0 Then
            If Not (i = LBound(parts) And isFixedFirstPart) Then
                Select Case LCase$(Trim$(targetCase))
                    Case CASE_ACCUSATIVE
                        Call mp_TryInflectSimpleCommonWordToAccusative(currentPart, partInflected)
                    Case CASE_DATIVE
                        Call mp_TryInflectSimpleCommonWordToDative(currentPart, partInflected)
                    Case Else
                        Call mp_TryInflectSimpleCommonWordToGenitive(currentPart, partInflected)
                End Select
            End If
        End If

        If i = LBound(parts) Then
            resultWord = partInflected
        Else
            resultWord = resultWord & "-" & partInflected
        End If
    Next i

    mp_TryInflectHyphenCommonWordToCase = True
End Function

Private Function mp_TryInflectSimpleCommonWordToGenitive(ByVal sourceWord As String, ByRef resultWord As String) As Boolean
    If Not mp_IsValidGeneralWord(sourceWord) Then Exit Function

    Dim low As String
    low = LCase$(sourceWord)
    low = mp_NormalizeTrailingYiSoftSign(low)

    Dim exceptions As Object
    Set exceptions = mp_GetCommonWordExceptionsGenitiveDict()
    If exceptions.Exists(low) Then
        resultWord = mp_ApplyWordCase(sourceWord, CStr(exceptions(low)))
        mp_TryInflectSimpleCommonWordToGenitive = True
        Exit Function
    End If

    Dim outLow As String
    outLow = low

    If mp_EndsWith(low, "ий") Then
        outLow = Left$(low, Len(low) - 2) & "ого"
    ElseIf mp_EndsWith(low, "ій") Then
        outLow = Left$(low, Len(low) - 2) & "ія"
    ElseIf mp_EndsWith(low, "а") Then
        outLow = Left$(low, Len(low) - 1) & "и"
    ElseIf mp_EndsWith(low, "я") Then
        If mp_IsLikelyNeuterNounOnYa(low) Then
            outLow = low
        Else
            outLow = Left$(low, Len(low) - 1) & "і"
        End If
    ElseIf mp_EndsWith(low, "ець") Then
        outLow = mp_InflectLowWordEndingEtsToGenitive(low)
    ElseIf mp_EndsWith(low, "ь") Then
        outLow = Left$(low, Len(low) - 1) & "я"
    ElseIf mp_EndsWith(low, "й") Then
        outLow = Left$(low, Len(low) - 1) & "я"
    ElseIf mp_EndsWithConsonant(low) Then
        outLow = low & "а"
    End If

    resultWord = mp_ApplyWordCase(sourceWord, outLow)
    mp_TryInflectSimpleCommonWordToGenitive = True
End Function

Private Function mp_TryInflectSimpleCommonWordToAccusative(ByVal sourceWord As String, ByRef resultWord As String) As Boolean
    If Not mp_IsValidGeneralWord(sourceWord) Then Exit Function

    Dim low As String
    low = LCase$(sourceWord)
    low = mp_NormalizeTrailingYiSoftSign(low)

    ' Keep likely dative-plural forms unchanged (e.g., "екіпажам")
    ' to avoid over-inflection like "екіпажама" in accusative pass.
    If mp_EndsWith(low, "ам") Or mp_EndsWith(low, "ям") Then
        resultWord = sourceWord
        mp_TryInflectSimpleCommonWordToAccusative = True
        Exit Function
    End If

    Dim exceptions As Object
    Set exceptions = mp_GetCommonWordExceptionsAccusativeDict()
    If exceptions.Exists(low) Then
        resultWord = mp_ApplyWordCase(sourceWord, CStr(exceptions(low)))
        mp_TryInflectSimpleCommonWordToAccusative = True
        Exit Function
    End If

    Dim outLow As String
    outLow = low

    If mp_EndsWith(low, "ий") Then
        outLow = Left$(low, Len(low) - 2) & "ого"
    ElseIf mp_EndsWith(low, "ій") Then
        outLow = Left$(low, Len(low) - 2) & "ія"
    ElseIf mp_EndsWith(low, "а") Then
        outLow = Left$(low, Len(low) - 1) & "у"
    ElseIf mp_EndsWith(low, "я") Then
        If mp_IsLikelyNeuterNounOnYa(low) Then
            outLow = low
        Else
            outLow = Left$(low, Len(low) - 1) & "ю"
        End If
    ElseIf mp_EndsWith(low, "ець") Then
        outLow = mp_InflectLowWordEndingEtsToGenitive(low)
    ElseIf mp_EndsWith(low, "ь") Then
        outLow = Left$(low, Len(low) - 1) & "я"
    ElseIf mp_EndsWith(low, "й") Then
        outLow = Left$(low, Len(low) - 1) & "я"
    ElseIf mp_EndsWithConsonant(low) Then
        outLow = low & "а"
    End If

    resultWord = mp_ApplyWordCase(sourceWord, outLow)
    mp_TryInflectSimpleCommonWordToAccusative = True
End Function

Private Function mp_TryInflectSimpleCommonWordToDative(ByVal sourceWord As String, ByRef resultWord As String) As Boolean
    If Not mp_IsValidGeneralWord(sourceWord) Then Exit Function

    Dim low As String
    low = LCase$(sourceWord)
    low = mp_NormalizeTrailingYiSoftSign(low)

    Dim exceptions As Object
    Set exceptions = mp_GetCommonWordExceptionsDativeDict()
    If exceptions.Exists(low) Then
        resultWord = mp_ApplyWordCase(sourceWord, CStr(exceptions(low)))
        mp_TryInflectSimpleCommonWordToDative = True
        Exit Function
    End If

    Dim outLow As String
    outLow = low

    If mp_EndsWith(low, "ий") Then
        outLow = Left$(low, Len(low) - 2) & "ому"
    ElseIf mp_EndsWith(low, "ій") Then
        outLow = Left$(low, Len(low) - 2) & "ію"
    ElseIf mp_EndsWith(low, "а") Then
        outLow = Left$(low, Len(low) - 1) & "і"
    ElseIf mp_EndsWith(low, "я") Then
        If mp_IsLikelyNeuterNounOnYa(low) Then
            outLow = low
        Else
            outLow = Left$(low, Len(low) - 1) & "і"
        End If
    ElseIf mp_EndsWith(low, "ець") Then
        outLow = mp_InflectLowWordEndingEtsToDative(low)
    ElseIf mp_EndsWith(low, "ь") Then
        outLow = Left$(low, Len(low) - 1) & "ю"
    ElseIf mp_EndsWith(low, "й") Then
        outLow = Left$(low, Len(low) - 1) & "ю"
    ElseIf mp_EndsWithConsonant(low) Then
        outLow = low & "у"
    End If

    resultWord = mp_ApplyWordCase(sourceWord, outLow)
    mp_TryInflectSimpleCommonWordToDative = True
End Function

Private Function mp_IsLikelyNeuterNounOnYa(ByVal low As String) As Boolean
    If mp_EndsWith(low, "ення") Or mp_EndsWith(low, "єння") Or mp_EndsWith(low, "іння") Or _
       mp_EndsWith(low, "ання") Or mp_EndsWith(low, "ття") Or mp_EndsWith(low, "лля") Then
        mp_IsLikelyNeuterNounOnYa = True
    End If
End Function

Private Function mp_InflectLowWordEndingEtsToGenitive(ByVal lowWord As String) As String
    If mp_EndsWith(lowWord, "лець") Then
        mp_InflectLowWordEndingEtsToGenitive = Left$(lowWord, Len(lowWord) - 4) & "льця"
    ElseIf mp_EndsWith(lowWord, "єць") Then
        mp_InflectLowWordEndingEtsToGenitive = Left$(lowWord, Len(lowWord) - 3) & "йця"
    ElseIf mp_EndsWith(lowWord, "ець") Then
        mp_InflectLowWordEndingEtsToGenitive = Left$(lowWord, Len(lowWord) - 3) & "ця"
    Else
        mp_InflectLowWordEndingEtsToGenitive = lowWord
    End If
End Function

Private Function mp_InflectLowWordEndingEtsToDative(ByVal lowWord As String) As String
    If mp_EndsWith(lowWord, "лець") Then
        mp_InflectLowWordEndingEtsToDative = Left$(lowWord, Len(lowWord) - 4) & "льцю"
    ElseIf mp_EndsWith(lowWord, "єць") Then
        mp_InflectLowWordEndingEtsToDative = Left$(lowWord, Len(lowWord) - 3) & "йцю"
    ElseIf mp_EndsWith(lowWord, "ець") Then
        mp_InflectLowWordEndingEtsToDative = Left$(lowWord, Len(lowWord) - 3) & "цю"
    Else
        mp_InflectLowWordEndingEtsToDative = lowWord
    End If
End Function

Private Function mp_ShouldKeepWordUnchangedByContext(ByVal lowPrevCore As String, ByVal lowCore As String, ByVal targetCase As String) As Boolean
    If Len(lowCore) = 0 Then Exit Function

    ' Do not re-inflect likely instrumental forms when context already governs instrumental.
    ' Example: "забезпечення продовольством" must keep "продовольством" unchanged.
    If mp_IsLikelyInstrumentalGovernanceContext(lowPrevCore) And mp_IsLikelyInstrumentalForm(lowCore) Then
        mp_ShouldKeepWordUnchangedByContext = True
        Exit Function
    End If

    If targetCase = CASE_ACCUSATIVE Then
        If mp_IsLikelyAlreadyAccusative(lowCore) Or mp_IsLikelyAlreadyGenitive(lowCore) Then
            mp_ShouldKeepWordUnchangedByContext = True
            Exit Function
        End If
    ElseIf targetCase = CASE_DATIVE Then
        If mp_IsLikelyAlreadyDative(lowCore) Then
            mp_ShouldKeepWordUnchangedByContext = True
            Exit Function
        End If
    Else
        If mp_IsLikelyAlreadyGenitive(lowCore) Then
            mp_ShouldKeepWordUnchangedByContext = True
            Exit Function
        End If
    End If

    If mp_IsGenitiveGovernedHeadWord(lowPrevCore) Then
        If mp_HasConsonantClusterEnding(lowCore) Or mp_IsLikelyMasculineGenitiveOnA(lowCore) Then
            mp_ShouldKeepWordUnchangedByContext = True
        End If
    End If
End Function

Private Function mp_IsLikelyInstrumentalGovernanceContext(ByVal lowWord As String) As Boolean
    If Len(lowWord) = 0 Then Exit Function

    Select Case lowWord
        Case "з", "із", "зі", "перед", "над", "під", "між", "за"
            mp_IsLikelyInstrumentalGovernanceContext = True
            Exit Function
    End Select

    If mp_EndsWith(lowWord, "ення") Or mp_EndsWith(lowWord, "єння") Or mp_EndsWith(lowWord, "іння") Or _
       mp_EndsWith(lowWord, "ання") Or mp_EndsWith(lowWord, "ття") Or mp_EndsWith(lowWord, "лля") Then
        mp_IsLikelyInstrumentalGovernanceContext = True
        Exit Function
    End If

    If mp_EndsWith(lowWord, "им") Or mp_EndsWith(lowWord, "ім") Then
        mp_IsLikelyInstrumentalGovernanceContext = True
    End If
End Function

Private Function mp_IsLikelyInstrumentalForm(ByVal lowWord As String) As Boolean
    If Len(lowWord) < 4 Then Exit Function

    If mp_EndsWith(lowWord, "ом") Or mp_EndsWith(lowWord, "ем") Or _
       mp_EndsWith(lowWord, "ою") Or mp_EndsWith(lowWord, "ею") Or mp_EndsWith(lowWord, "єю") Or _
       mp_EndsWith(lowWord, "ами") Or mp_EndsWith(lowWord, "ями") Or mp_EndsWith(lowWord, "ьми") Or _
       mp_EndsWith(lowWord, "им") Or mp_EndsWith(lowWord, "ім") Then
        mp_IsLikelyInstrumentalForm = True
    End If
End Function

Private Function mp_IsLikelyAlreadyAccusative(ByVal low As String) As Boolean
    If mp_EndsWith(low, "у") Or mp_EndsWith(low, "ю") Then
        mp_IsLikelyAlreadyAccusative = True
        Exit Function
    End If

    If mp_EndsWith(low, "ого") Or mp_EndsWith(low, "ього") Then
        mp_IsLikelyAlreadyAccusative = True
    End If
End Function

Private Function mp_IsLikelyAlreadyGenitive(ByVal low As String) As Boolean
    If mp_IsLikelyAdjectiveGenitive(low) Then
        mp_IsLikelyAlreadyGenitive = True
        Exit Function
    End If

    If mp_EndsWith(low, "у") Or mp_EndsWith(low, "ю") Or mp_EndsWith(low, "ів") Or _
       mp_EndsWith(low, "їв") Or mp_EndsWith(low, "ей") Then
        mp_IsLikelyAlreadyGenitive = True
        Exit Function
    End If

    If Len(low) > 3 Then
        If mp_EndsWith(low, "и") Or mp_EndsWith(low, "і") Then
            mp_IsLikelyAlreadyGenitive = True
        End If
    End If
End Function

Private Function mp_IsLikelyAlreadyDative(ByVal low As String) As Boolean
    If mp_EndsWith(low, "ому") Or mp_EndsWith(low, "ьому") Or mp_EndsWith(low, "еві") Or _
       mp_EndsWith(low, "єві") Or mp_EndsWith(low, "ові") Then
        mp_IsLikelyAlreadyDative = True
        Exit Function
    End If

    If Len(low) > 3 Then
        If mp_EndsWith(low, "і") Or mp_EndsWith(low, "у") Or mp_EndsWith(low, "ю") Then
            mp_IsLikelyAlreadyDative = True
        End If
    End If
End Function

Private Function mp_IsLikelyAdjectiveGenitive(ByVal low As String) As Boolean
    If mp_EndsWith(low, "ого") Or mp_EndsWith(low, "ього") Or mp_EndsWith(low, "ої") Or _
       mp_EndsWith(low, "ьої") Or mp_EndsWith(low, "єї") Or mp_EndsWith(low, "их") Or _
       mp_EndsWith(low, "іх") Then
        mp_IsLikelyAdjectiveGenitive = True
    End If
End Function

Private Function mp_IsGenitiveGovernedHeadWord(ByVal lowWord As String) As Boolean
    If Len(lowWord) = 0 Then Exit Function

    If mp_IsLikelyMasculineGenitiveOnA(lowWord) Then
        mp_IsGenitiveGovernedHeadWord = True
        Exit Function
    End If

    If mp_EndsWith(lowWord, "ення") Or mp_EndsWith(lowWord, "єння") Or mp_EndsWith(lowWord, "іння") Or _
       mp_EndsWith(lowWord, "ання") Or mp_EndsWith(lowWord, "ття") Or mp_EndsWith(lowWord, "лля") Then
        mp_IsGenitiveGovernedHeadWord = True
    End If
End Function

Private Function mp_IsLikelyMasculineGenitiveOnA(ByVal lowWord As String) As Boolean
    If Len(lowWord) < 4 Then Exit Function
    If Not (mp_EndsWith(lowWord, "а") Or mp_EndsWith(lowWord, "я")) Then Exit Function

    If mp_IsConsonantChar(Mid$(lowWord, Len(lowWord) - 1, 1)) Then
        mp_IsLikelyMasculineGenitiveOnA = True
    End If
End Function

Private Function mp_HasConsonantClusterEnding(ByVal lowWord As String) As Boolean
    Dim wordLen As Long
    wordLen = Len(lowWord)
    If wordLen < 2 Then Exit Function

    Dim lastCh As String
    Dim prevCh As String
    lastCh = Mid$(lowWord, wordLen, 1)
    prevCh = Mid$(lowWord, wordLen - 1, 1)

    If Not mp_IsConsonantChar(lastCh) Then Exit Function

    If mp_IsConsonantChar(prevCh) Then
        mp_HasConsonantClusterEnding = True
        Exit Function
    End If

    If prevCh = "ь" And wordLen >= 3 Then
        prevCh = Mid$(lowWord, wordLen - 2, 1)
        If mp_IsConsonantChar(prevCh) Then
            mp_HasConsonantClusterEnding = True
        End If
    End If
End Function

Private Function mp_IsConsonantChar(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    mp_IsConsonantChar = (InStr("бвгґджзйклмнпрстфхцчшщ", ch) > 0)
End Function

Private Function mp_IsValidGeneralWord(ByVal token As String) As Boolean
    mp_IsValidGeneralWord = mp_IsValidWordToken(token)
End Function

Private Function mp_IsValidWordToken(ByVal token As String) As Boolean
    Dim i As Long
    Dim code As Long

    token = Trim$(token)
    If Len(token) = 0 Then Exit Function

    If Not mp_IsLetterCodePoint(mp_CodePointAt(token, 1)) Then Exit Function

    For i = 2 To Len(token)
        code = mp_CodePointAt(token, i)

        If mp_IsLetterCodePoint(code) Then
            GoTo ContinueLoop
        End If

        Select Case code
            Case &H2D, &H27, &H60, &H2019, &H2BC, &H42C, &H44C ' -, ', `, ’, ʼ, Ь, ь
                GoTo ContinueLoop
        End Select

        Exit Function
ContinueLoop:
    Next i

    mp_IsValidWordToken = True
End Function

Private Function mp_CodePointAt(ByVal sourceText As String, ByVal index As Long) As Long
    mp_CodePointAt = AscW(Mid$(sourceText, index, 1))
    If mp_CodePointAt < 0 Then mp_CodePointAt = mp_CodePointAt + 65536
End Function

Private Function mp_IsLetterCodePoint(ByVal code As Long) As Boolean
    Select Case code
        Case 65 To 90, 97 To 122, _
             &H410 To &H44F, _
             &H401, &H451, _
             &H404, &H454, _
             &H406, &H456, _
             &H407, &H457, _
             &H490, &H491
            mp_IsLetterCodePoint = True
    End Select
End Function

Private Function mp_IsStopWord(ByVal token As String) As Boolean
    Dim low As String
    low = LCase$(token)

    Select Case low
        Case "і", "й", "та", "або", "в", "у", "на", "до", "з", "із", "зі", "по", "за", "від", "при", "для", "про"
            mp_IsStopWord = True
    End Select
End Function

Private Function mp_TrimTokenPunctuation(ByVal token As String) As String
    Dim s As String
    s = Trim$(token)

    Do While Len(s) > 0 And InStr(".,;:!?()[]{}""«»", Left$(s, 1)) > 0
        s = Mid$(s, 2)
    Loop

    Do While Len(s) > 0 And InStr(".,;:!?()[]{}""«»", Right$(s, 1)) > 0
        s = Left$(s, Len(s) - 1)
    Loop

    mp_TrimTokenPunctuation = s
End Function

Private Function mp_TokenTailPunctuation(ByVal token As String) As String
    Dim i As Long
    For i = Len(token) To 1 Step -1
        If InStr(".,;:!?)""»", Mid$(token, i, 1)) = 0 Then Exit For
    Next i

    If i < Len(token) Then
        mp_TokenTailPunctuation = Mid$(token, i + 1)
    End If
End Function

Private Function mp_IsFixedHyphenFirstPart(ByVal partText As String) As Boolean
    Dim low As String
    low = LCase$(partText)

    Select Case low
        Case "штаб", "обер", "унтер", "віце", "екс", "лейб", "контр", "псевдо"
            mp_IsFixedHyphenFirstPart = True
    End Select
End Function

Private Function mp_ApplyWordCase(ByVal sourceWord As String, ByVal inflectedLower As String) As String
    If sourceWord = UCase$(sourceWord) Then
        mp_ApplyWordCase = UCase$(inflectedLower)
    ElseIf sourceWord = LCase$(sourceWord) Then
        mp_ApplyWordCase = LCase$(inflectedLower)
    Else
        mp_ApplyWordCase = mp_ToTitleCaseWord(inflectedLower)
    End If
End Function

Private Function mp_ToTitleCaseWord(ByVal textValue As String) As String
    Dim parts() As String
    parts = Split(textValue, "-")

    Dim i As Long
    Dim part As String
    For i = LBound(parts) To UBound(parts)
        part = LCase$(parts(i))
        If Len(part) > 0 Then
            parts(i) = UCase$(Left$(part, 1)) & Mid$(part, 2)
        End If
    Next i

    mp_ToTitleCaseWord = Join(parts, "-")
End Function

Private Function mp_EndsWith(ByVal textValue As String, ByVal suffix As String) As Boolean
    If Len(textValue) < Len(suffix) Then Exit Function
    mp_EndsWith = (Right$(textValue, Len(suffix)) = suffix)
End Function

Private Function mp_NormalizeTrailingYiSoftSign(ByVal lowWord As String) As String
    mp_NormalizeTrailingYiSoftSign = lowWord
    If Len(lowWord) < 2 Then Exit Function

    If Right$(lowWord, 2) = "йь" Then
        mp_NormalizeTrailingYiSoftSign = Left$(lowWord, Len(lowWord) - 1)
    End If
End Function

Private Function mp_EndsWithConsonant(ByVal textValue As String) As Boolean
    If Len(textValue) = 0 Then Exit Function
    mp_EndsWithConsonant = (InStr("бвгґджзйклмнпрстфхцчшщ", Right$(textValue, 1)) > 0)
End Function

Private Function mp_GetCommonWordExceptionsGenitiveDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    d("штаб") = "штабу"
    d("майстер") = "майстра"

    Set mp_GetCommonWordExceptionsGenitiveDict = d
End Function

Private Function mp_GetCommonWordExceptionsAccusativeDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    d("майстер") = "майстра"

    Set mp_GetCommonWordExceptionsAccusativeDict = d
End Function

Private Function mp_GetCommonWordExceptionsDativeDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    d("штаб") = "штабу"
    d("майстер") = "майстру"

    Set mp_GetCommonWordExceptionsDativeDict = d
End Function

Private Function mp_TryConvertSelectionTextToGenitive(ByVal normalizedText As String, ByRef convertedText As String) As Boolean
    If mp_TryConvertPureFioToGenitive(normalizedText, convertedText) Then
        mp_TryConvertSelectionTextToGenitive = True
        Exit Function
    End If

    If Not mp_IsSingleSentenceText(normalizedText) Then Exit Function

    If mp_TryConvertSentenceWithFioToGenitive(normalizedText, convertedText) Then
        mp_TryConvertSelectionTextToGenitive = True
    End If
End Function

Private Function mp_TryConvertSelectionTextToAccusative(ByVal normalizedText As String, ByRef convertedText As String) As Boolean
    If mp_TryConvertPureFioToAccusative(normalizedText, convertedText) Then
        mp_TryConvertSelectionTextToAccusative = True
        Exit Function
    End If

    If Not mp_IsSingleSentenceText(normalizedText) Then Exit Function

    If mp_TryConvertSentenceWithFioToAccusative(normalizedText, convertedText) Then
        mp_TryConvertSelectionTextToAccusative = True
    End If
End Function

Private Function mp_TryConvertSelectionTextToDative(ByVal normalizedText As String, ByRef convertedText As String) As Boolean
    If mp_TryConvertPureFioToDative(normalizedText, convertedText) Then
        mp_TryConvertSelectionTextToDative = True
        Exit Function
    End If

    If Not mp_IsSingleSentenceText(normalizedText) Then Exit Function

    If mp_TryConvertSentenceWithFioToDative(normalizedText, convertedText) Then
        mp_TryConvertSelectionTextToDative = True
    End If
End Function

Private Function mp_IsSingleSentenceText(ByVal normalizedText As String) As Boolean
    Static regex As Object

    If regex Is Nothing Then
        Set regex = CreateObject("VBScript.RegExp")
        regex.Global = False
        regex.IgnoreCase = True
        regex.Pattern = "^[^.!?]+([.!?]+)?$"
    End If

    mp_IsSingleSentenceText = regex.Test(Trim$(normalizedText))
End Function

Private Function mp_TryConvertPureFioToGenitive(ByVal normalizedText As String, ByRef convertedText As String) As Boolean
    Dim surname As String, firstName As String, patronymic As String
    If Not mp_TryParseFio(normalizedText, surname, firstName, patronymic) Then Exit Function

    Dim gender As String
    If Not mp_DetectFioGender(firstName, patronymic, gender) Then Exit Function

    Dim genSurname As String, genFirstName As String, genPatronymic As String
    If Not mp_InflectSurnameToGenitive(surname, gender, genSurname) Then Exit Function
    If Not mp_InflectNameToGenitive(firstName, gender, genFirstName) Then Exit Function
    If Not mp_InflectPatronymicToGenitive(patronymic, gender, genPatronymic) Then Exit Function

    convertedText = genSurname & " " & genFirstName & " " & genPatronymic
    mp_TryConvertPureFioToGenitive = True
End Function

Private Function mp_TryConvertPureFioToAccusative(ByVal normalizedText As String, ByRef convertedText As String) As Boolean
    Dim surname As String, firstName As String, patronymic As String
    If Not mp_TryParseFio(normalizedText, surname, firstName, patronymic) Then Exit Function

    Dim gender As String
    If Not mp_DetectFioGender(firstName, patronymic, gender) Then Exit Function

    Dim accSurname As String, accFirstName As String, accPatronymic As String
    If Not mp_InflectSurnameToAccusative(surname, gender, accSurname) Then Exit Function
    If Not mp_InflectNameToAccusative(firstName, gender, accFirstName) Then Exit Function
    If Not mp_InflectPatronymicToAccusative(patronymic, gender, accPatronymic) Then Exit Function

    convertedText = accSurname & " " & accFirstName & " " & accPatronymic
    mp_TryConvertPureFioToAccusative = True
End Function

Private Function mp_TryConvertPureFioToDative(ByVal normalizedText As String, ByRef convertedText As String) As Boolean
    Dim surname As String, firstName As String, patronymic As String
    If Not mp_TryParseFio(normalizedText, surname, firstName, patronymic) Then Exit Function

    Dim gender As String
    If Not mp_DetectFioGender(firstName, patronymic, gender) Then Exit Function

    Dim datSurname As String, datFirstName As String, datPatronymic As String
    If Not mp_InflectSurnameToDative(surname, gender, datSurname) Then Exit Function
    If Not mp_InflectNameToDative(firstName, gender, datFirstName) Then Exit Function
    If Not mp_InflectPatronymicToDative(patronymic, gender, datPatronymic) Then Exit Function

    convertedText = datSurname & " " & datFirstName & " " & datPatronymic
    mp_TryConvertPureFioToDative = True
End Function

Private Function mp_TryConvertSentenceWithFioToGenitive(ByVal normalizedText As String, ByRef convertedText As String) As Boolean
    Dim leadPhrase As String
    Dim surname As String
    Dim firstName As String
    Dim patronymic As String
    Dim tailPhrase As String

    If Not mp_TryParseSentenceWithFio(normalizedText, leadPhrase, surname, firstName, patronymic, tailPhrase) Then Exit Function

    Dim fioGenitive As String
    Dim gender As String
    If Not mp_TryInflectFioToGenitive(surname, firstName, patronymic, fioGenitive, gender) Then Exit Function

    Dim leadPhraseGen As String
    leadPhraseGen = m_InflectPhraseByDashSegmentsToCase(leadPhrase, "genitive", 4)
    If Len(leadPhraseGen) = 0 Then Exit Function

    If Len(Trim$(tailPhrase)) > 0 Then
        Dim tailPhraseGen As String
        If mp_ShouldSkipTailInflection(tailPhrase) Then
            tailPhraseGen = tailPhrase
        Else
            tailPhraseGen = m_InflectPhraseByDashSegmentsToCase(tailPhrase, "genitive", 4)
            If Len(tailPhraseGen) = 0 Then Exit Function
        End If
        convertedText = leadPhraseGen & " " & fioGenitive & ", " & mp_LowercaseFirstLetter(tailPhraseGen)
    Else
        convertedText = leadPhraseGen & " " & fioGenitive
    End If

    mp_TryConvertSentenceWithFioToGenitive = True
End Function

Private Function mp_TryConvertSentenceWithFioToAccusative(ByVal normalizedText As String, ByRef convertedText As String) As Boolean
    Dim leadPhrase As String
    Dim surname As String
    Dim firstName As String
    Dim patronymic As String
    Dim tailPhrase As String

    If Not mp_TryParseSentenceWithFio(normalizedText, leadPhrase, surname, firstName, patronymic, tailPhrase) Then Exit Function

    Dim fioAccusative As String
    Dim gender As String
    If Not mp_TryInflectFioToAccusative(surname, firstName, patronymic, fioAccusative, gender) Then Exit Function

    Dim leadPhraseAcc As String
    leadPhraseAcc = m_InflectPhraseByDashSegmentsToCase(leadPhrase, "accusative", 4)
    If Len(leadPhraseAcc) = 0 Then Exit Function

    If Len(Trim$(tailPhrase)) > 0 Then
        Dim tailPhraseAcc As String
        If mp_ShouldSkipTailInflection(tailPhrase) Then
            tailPhraseAcc = tailPhrase
        Else
            tailPhraseAcc = m_InflectPhraseByDashSegmentsToCase(tailPhrase, "accusative", 4)
            If Len(tailPhraseAcc) = 0 Then Exit Function
        End If
        convertedText = leadPhraseAcc & " " & fioAccusative & ", " & mp_LowercaseFirstLetter(tailPhraseAcc)
    Else
        convertedText = leadPhraseAcc & " " & fioAccusative
    End If

    mp_TryConvertSentenceWithFioToAccusative = True
End Function

Private Function mp_TryConvertSentenceWithFioToDative(ByVal normalizedText As String, ByRef convertedText As String) As Boolean
    Dim leadPhrase As String
    Dim surname As String
    Dim firstName As String
    Dim patronymic As String
    Dim tailPhrase As String

    If Not mp_TryParseSentenceWithFio(normalizedText, leadPhrase, surname, firstName, patronymic, tailPhrase) Then Exit Function

    Dim fioDative As String
    Dim gender As String
    If Not mp_TryInflectFioToDative(surname, firstName, patronymic, fioDative, gender) Then Exit Function

    Dim leadPhraseDat As String
    leadPhraseDat = m_InflectPhraseByDashSegmentsToCase(leadPhrase, CASE_DATIVE, 4)
    If Len(leadPhraseDat) = 0 Then Exit Function

    If Len(Trim$(tailPhrase)) > 0 Then
        Dim tailPhraseDat As String
        If mp_ShouldSkipTailInflection(tailPhrase) Then
            tailPhraseDat = tailPhrase
        Else
            tailPhraseDat = m_InflectPhraseByDashSegmentsToCase(tailPhrase, CASE_DATIVE, 4)
            If Len(tailPhraseDat) = 0 Then Exit Function
        End If
        convertedText = leadPhraseDat & " " & fioDative & ", " & mp_LowercaseFirstLetter(tailPhraseDat)
    Else
        convertedText = leadPhraseDat & " " & fioDative
    End If

    mp_TryConvertSentenceWithFioToDative = True
End Function

Private Function mp_ShouldSkipTailInflection(ByVal tailPhrase As String) As Boolean
    mp_ShouldSkipTailInflection = mp_ShouldSkipPhraseInflection(tailPhrase)
End Function

Private Function mp_TryParseSentenceWithFio(ByVal normalizedText As String, ByRef leadPhrase As String, ByRef surname As String, ByRef firstName As String, ByRef patronymic As String, ByRef tailPhrase As String) As Boolean
    Dim parts() As String
    parts = Split(normalizedText, " ")
    If UBound(parts) < 3 Then Exit Function

    Dim fioStart As Long
    If Not mp_FindFioStartIndex(parts, fioStart, surname, firstName, patronymic) Then Exit Function

    If fioStart < 1 Then Exit Function

    leadPhrase = mp_JoinArraySlice(parts, 0, fioStart - 1)
    If fioStart + 3 <= UBound(parts) Then
        tailPhrase = mp_JoinArraySlice(parts, fioStart + 3, UBound(parts))
    Else
        tailPhrase = vbNullString
    End If

    If Len(leadPhrase) = 0 Then Exit Function

    mp_TryParseSentenceWithFio = True
End Function

Private Function mp_FindFioStartIndex(ByRef parts() As String, ByRef fioStart As Long, ByRef surname As String, ByRef firstName As String, ByRef patronymic As String) As Boolean
    Dim i As Long
    Dim c1 As String
    Dim c2 As String
    Dim c3 As String
    Dim gender As String

    For i = LBound(parts) To UBound(parts) - 2
        c1 = mp_TrimTokenPunctuation(parts(i))
        c2 = mp_TrimTokenPunctuation(parts(i + 1))
        c3 = mp_TrimTokenPunctuation(parts(i + 2))

        If mp_IsValidFioToken(c1) And mp_IsValidFioToken(c2) And mp_IsValidFioToken(c3) Then
            If mp_DetectFioGender(c2, c3, gender) Then
                fioStart = i
                surname = c1
                firstName = c2
                patronymic = c3
                mp_FindFioStartIndex = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function mp_TryInflectFioToGenitive(ByVal surname As String, ByVal firstName As String, ByVal patronymic As String, ByRef fioGenitive As String, ByRef gender As String) As Boolean
    If Not mp_DetectFioGender(firstName, patronymic, gender) Then Exit Function

    Dim genSurname As String
    Dim genFirstName As String
    Dim genPatronymic As String

    If Not mp_InflectSurnameToGenitive(surname, gender, genSurname) Then Exit Function
    If Not mp_InflectNameToGenitive(firstName, gender, genFirstName) Then Exit Function
    If Not mp_InflectPatronymicToGenitive(patronymic, gender, genPatronymic) Then Exit Function

    fioGenitive = genSurname & " " & genFirstName & " " & genPatronymic
    mp_TryInflectFioToGenitive = True
End Function

Private Function mp_TryInflectFioToAccusative(ByVal surname As String, ByVal firstName As String, ByVal patronymic As String, ByRef fioAccusative As String, ByRef gender As String) As Boolean
    If Not mp_DetectFioGender(firstName, patronymic, gender) Then Exit Function

    Dim accSurname As String
    Dim accFirstName As String
    Dim accPatronymic As String

    If Not mp_InflectSurnameToAccusative(surname, gender, accSurname) Then Exit Function
    If Not mp_InflectNameToAccusative(firstName, gender, accFirstName) Then Exit Function
    If Not mp_InflectPatronymicToAccusative(patronymic, gender, accPatronymic) Then Exit Function

    fioAccusative = accSurname & " " & accFirstName & " " & accPatronymic
    mp_TryInflectFioToAccusative = True
End Function

Private Function mp_TryInflectFioToDative(ByVal surname As String, ByVal firstName As String, ByVal patronymic As String, ByRef fioDative As String, ByRef gender As String) As Boolean
    If Not mp_DetectFioGender(firstName, patronymic, gender) Then Exit Function

    Dim datSurname As String
    Dim datFirstName As String
    Dim datPatronymic As String

    If Not mp_InflectSurnameToDative(surname, gender, datSurname) Then Exit Function
    If Not mp_InflectNameToDative(firstName, gender, datFirstName) Then Exit Function
    If Not mp_InflectPatronymicToDative(patronymic, gender, datPatronymic) Then Exit Function

    fioDative = datSurname & " " & datFirstName & " " & datPatronymic
    mp_TryInflectFioToDative = True
End Function

Private Function mp_JoinArraySlice(ByRef parts() As String, ByVal startIndex As Long, Optional ByVal endIndex As Long = -1) As String
    If endIndex < 0 Then endIndex = UBound(parts)
    If startIndex > endIndex Then Exit Function

    Dim i As Long
    For i = startIndex To endIndex
        If Len(mp_JoinArraySlice) = 0 Then
            mp_JoinArraySlice = parts(i)
        Else
            mp_JoinArraySlice = mp_JoinArraySlice & " " & parts(i)
        End If
    Next i
End Function

Private Function mp_LowercaseFirstLetter(ByVal sourceText As String) As String
    If Len(sourceText) = 0 Then Exit Function
    mp_LowercaseFirstLetter = LCase$(Left$(sourceText, 1)) & Mid$(sourceText, 2)
End Function

Private Function mp_TryParseFio(ByVal normalizedText As String, ByRef surname As String, ByRef firstName As String, ByRef patronymic As String) As Boolean
    Dim parts() As String
    parts = Split(normalizedText, " ")
    If UBound(parts) <> 2 Then Exit Function

    Dim p0 As String
    Dim p1 As String
    Dim p2 As String

    p0 = mp_TrimTokenPunctuation(parts(0))
    p1 = mp_TrimTokenPunctuation(parts(1))
    p2 = mp_TrimTokenPunctuation(parts(2))

    If Not mp_IsValidFioToken(p0) Then Exit Function
    If Not mp_IsValidFioToken(p1) Then Exit Function
    If Not mp_IsValidFioToken(p2) Then Exit Function

    surname = p0
    firstName = p1
    patronymic = p2
    mp_TryParseFio = True
End Function

Private Function mp_NormalizeFioInput(ByVal inputText As String) As String
    Dim s As String
    s = Trim$(inputText)
    s = Replace$(s, vbCr, " ")
    s = Replace$(s, vbLf, " ")
    s = Replace$(s, vbTab, " ")
    s = Replace$(s, ChrW$(NBSP_CODE_POINT), " ")
    s = Replace$(s, ChrW$(NARROW_NBSP_CODE_POINT), " ")
    s = Replace$(s, Chr$(7), " ")

    s = Replace$(s, "’", "'")
    s = Replace$(s, "ʼ", "'")
    s = Replace$(s, "`", "'")

    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop

    mp_NormalizeFioInput = Trim$(s)
End Function

Private Function mp_IsValidFioToken(ByVal token As String) As Boolean
    mp_IsValidFioToken = mp_IsValidWordToken(token)
End Function

Private Function mp_DetectFioGender(ByVal firstName As String, ByVal patronymic As String, ByRef gender As String) As Boolean
    Dim p As String
    p = LCase$(patronymic)

    If mp_EndsWith(p, "ович") Or mp_EndsWith(p, "евич") Or mp_EndsWith(p, "йович") Or _
       mp_EndsWith(p, "овича") Or mp_EndsWith(p, "евича") Or mp_EndsWith(p, "йовича") Then
        gender = "male"
        mp_DetectFioGender = True
        Exit Function
    End If

    If mp_EndsWith(p, "івна") Or mp_EndsWith(p, "ївна") Or mp_EndsWith(p, "овна") Or mp_EndsWith(p, "евна") Or _
       mp_EndsWith(p, "івни") Or mp_EndsWith(p, "ївни") Or mp_EndsWith(p, "овни") Or mp_EndsWith(p, "евни") Then
        gender = "female"
        mp_DetectFioGender = True
        Exit Function
    End If

    Dim maleNames As Object
    Set maleNames = mp_GetMaleNamesDict()
    Dim femaleNames As Object
    Set femaleNames = mp_GetFemaleNamesDict()

    Dim n As String
    n = LCase$(firstName)

    If maleNames.Exists(n) And Not femaleNames.Exists(n) Then
        gender = "male"
        mp_DetectFioGender = True
        Exit Function
    End If

    If femaleNames.Exists(n) And Not maleNames.Exists(n) Then
        gender = "female"
        mp_DetectFioGender = True
    End If
End Function

Private Function mp_InflectSurnameToGenitive(ByVal surname As String, ByVal gender As String, ByRef resultText As String) As Boolean
    mp_InflectSurnameToGenitive = mp_InflectTokenByHyphenParts(surname, gender, "surname", resultText)
End Function

Private Function mp_InflectNameToGenitive(ByVal firstName As String, ByVal gender As String, ByRef resultText As String) As Boolean
    mp_InflectNameToGenitive = mp_InflectTokenByHyphenParts(firstName, gender, "name", resultText)
End Function

Private Function mp_InflectPatronymicToGenitive(ByVal patronymic As String, ByVal gender As String, ByRef resultText As String) As Boolean
    mp_InflectPatronymicToGenitive = mp_InflectTokenByHyphenParts(patronymic, gender, "patronymic", resultText)
End Function

Private Function mp_InflectSurnameToAccusative(ByVal surname As String, ByVal gender As String, ByRef resultText As String) As Boolean
    mp_InflectSurnameToAccusative = mp_InflectTokenByHyphenPartsAccusative(surname, gender, "surname", resultText)
End Function

Private Function mp_InflectNameToAccusative(ByVal firstName As String, ByVal gender As String, ByRef resultText As String) As Boolean
    mp_InflectNameToAccusative = mp_InflectTokenByHyphenPartsAccusative(firstName, gender, "name", resultText)
End Function

Private Function mp_InflectPatronymicToAccusative(ByVal patronymic As String, ByVal gender As String, ByRef resultText As String) As Boolean
    mp_InflectPatronymicToAccusative = mp_InflectTokenByHyphenPartsAccusative(patronymic, gender, "patronymic", resultText)
End Function

Private Function mp_InflectSurnameToDative(ByVal surname As String, ByVal gender As String, ByRef resultText As String) As Boolean
    mp_InflectSurnameToDative = mp_InflectTokenByHyphenPartsDative(surname, gender, "surname", resultText)
End Function

Private Function mp_InflectNameToDative(ByVal firstName As String, ByVal gender As String, ByRef resultText As String) As Boolean
    mp_InflectNameToDative = mp_InflectTokenByHyphenPartsDative(firstName, gender, "name", resultText)
End Function

Private Function mp_InflectPatronymicToDative(ByVal patronymic As String, ByVal gender As String, ByRef resultText As String) As Boolean
    mp_InflectPatronymicToDative = mp_InflectTokenByHyphenPartsDative(patronymic, gender, "patronymic", resultText)
End Function

Private Function mp_InflectTokenByHyphenParts(ByVal token As String, ByVal gender As String, ByVal tokenType As String, ByRef resultText As String) As Boolean
    Dim parts() As String
    parts = Split(token, "-")

    Dim i As Long
    Dim partResult As String

    For i = LBound(parts) To UBound(parts)
        partResult = ""

        Select Case tokenType
            Case "surname"
                If Not mp_InflectSurnamePart(parts(i), gender, partResult) Then Exit Function
            Case "name"
                If Not mp_InflectNamePart(parts(i), gender, partResult) Then Exit Function
            Case "patronymic"
                If Not mp_InflectPatronymicPart(parts(i), gender, partResult) Then Exit Function
            Case Else
                Exit Function
        End Select

        If i = LBound(parts) Then
            resultText = partResult
        Else
            resultText = resultText & "-" & partResult
        End If
    Next i

    mp_InflectTokenByHyphenParts = True
End Function

Private Function mp_InflectTokenByHyphenPartsAccusative(ByVal token As String, ByVal gender As String, ByVal tokenType As String, ByRef resultText As String) As Boolean
    Dim parts() As String
    parts = Split(token, "-")

    Dim i As Long
    Dim partResult As String

    For i = LBound(parts) To UBound(parts)
        partResult = ""

        Select Case tokenType
            Case "surname"
                If Not mp_InflectSurnamePartToAccusative(parts(i), gender, partResult) Then Exit Function
            Case "name"
                If Not mp_InflectNamePartToAccusative(parts(i), gender, partResult) Then Exit Function
            Case "patronymic"
                If Not mp_InflectPatronymicPartToAccusative(parts(i), gender, partResult) Then Exit Function
            Case Else
                Exit Function
        End Select

        If i = LBound(parts) Then
            resultText = partResult
        Else
            resultText = resultText & "-" & partResult
        End If
    Next i

    mp_InflectTokenByHyphenPartsAccusative = True
End Function

Private Function mp_InflectTokenByHyphenPartsDative(ByVal token As String, ByVal gender As String, ByVal tokenType As String, ByRef resultText As String) As Boolean
    Dim parts() As String
    parts = Split(token, "-")

    Dim i As Long
    Dim partResult As String

    For i = LBound(parts) To UBound(parts)
        partResult = ""

        Select Case tokenType
            Case "surname"
                If Not mp_InflectSurnamePartToDative(parts(i), gender, partResult) Then Exit Function
            Case "name"
                If Not mp_InflectNamePartToDative(parts(i), gender, partResult) Then Exit Function
            Case "patronymic"
                If Not mp_InflectPatronymicPartToDative(parts(i), gender, partResult) Then Exit Function
            Case Else
                Exit Function
        End Select

        If i = LBound(parts) Then
            resultText = partResult
        Else
            resultText = resultText & "-" & partResult
        End If
    Next i

    mp_InflectTokenByHyphenPartsDative = True
End Function

Private Function mp_InflectSurnamePart(ByVal originalPart As String, ByVal gender As String, ByRef partResult As String) As Boolean
    Dim low As String
    low = LCase$(originalPart)
    low = mp_NormalizeTrailingYiSoftSign(low)

    Dim exceptions As Object
    Set exceptions = mp_GetSurnameExceptionsDict()
    If exceptions.Exists(low) Then
        partResult = mp_ApplyWordCase(originalPart, CStr(exceptions(low)))
        mp_InflectSurnamePart = True
        Exit Function
    End If

    If mp_IsIndeclinableSurname(low, gender) Then
        partResult = originalPart
        mp_InflectSurnamePart = True
        Exit Function
    End If

    Dim outLow As String
    outLow = low

    If gender = "male" Then
        If mp_EndsWith(low, "ський") Then
            outLow = Left$(low, Len(low) - 5) & "ського"
        ElseIf mp_EndsWith(low, "цький") Then
            outLow = Left$(low, Len(low) - 5) & "цького"
        ElseIf mp_EndsWith(low, "зький") Then
            outLow = Left$(low, Len(low) - 5) & "зького"
        ElseIf mp_EndsWith(low, "ой") Or mp_EndsWith(low, "ый") Then
            outLow = Left$(low, Len(low) - 2) & "ого"
        ElseIf mp_EndsWith(low, "ець") Then
            outLow = mp_InflectLowWordEndingEtsToGenitive(low)
        ElseIf mp_EndsWith(low, "ий") Then
            outLow = Left$(low, Len(low) - 2) & "ого"
        ElseIf mp_EndsWith(low, "ній") Then
            outLow = Left$(low, Len(low) - 3) & "нього"
        ElseIf mp_EndsWith(low, "ій") Then
            outLow = Left$(low, Len(low) - 2) & "ія"
        ElseIf mp_EndsWith(low, "ко") Then
            outLow = Left$(low, Len(low) - 1) & "а"
        ElseIf mp_EndsWith(low, "а") Then
            outLow = Left$(low, Len(low) - 1) & "и"
        ElseIf mp_EndsWith(low, "я") Then
            outLow = Left$(low, Len(low) - 1) & "і"
        ElseIf mp_EndsWith(low, "ь") Then
            outLow = Left$(low, Len(low) - 1) & "я"
        ElseIf mp_EndsWith(low, "й") Then
            outLow = Left$(low, Len(low) - 1) & "я"
        ElseIf mp_EndsWithConsonant(low) Then
            outLow = low & "а"
        End If
    ElseIf gender = "female" Then
        If mp_EndsWith(low, "ська") Then
            outLow = Left$(low, Len(low) - 4) & "ської"
        ElseIf mp_EndsWith(low, "цька") Then
            outLow = Left$(low, Len(low) - 4) & "цької"
        ElseIf mp_EndsWith(low, "зька") Then
            outLow = Left$(low, Len(low) - 4) & "зької"
        ElseIf mp_EndsWith(low, "ова") Or mp_EndsWith(low, "ева") Or mp_EndsWith(low, "єва") Or _
               mp_EndsWith(low, "іна") Or mp_EndsWith(low, "їна") Or mp_EndsWith(low, "ина") Then
            outLow = Left$(low, Len(low) - 1) & "ої"
        ElseIf mp_EndsWith(low, "а") Then
            outLow = Left$(low, Len(low) - 1) & "и"
        ElseIf mp_EndsWith(low, "я") Then
            outLow = Left$(low, Len(low) - 1) & "і"
        End If
    Else
        Exit Function
    End If

    partResult = mp_ApplyWordCase(originalPart, outLow)
    mp_InflectSurnamePart = True
End Function

Private Function mp_InflectNamePart(ByVal originalPart As String, ByVal gender As String, ByRef partResult As String) As Boolean
    Dim low As String
    low = LCase$(originalPart)
    low = mp_NormalizeTrailingYiSoftSign(low)

    Dim exceptions As Object
    Set exceptions = mp_GetNameExceptionsDict()
    If exceptions.Exists(low) Then
        partResult = mp_ApplyWordCase(originalPart, CStr(exceptions(low)))
        mp_InflectNamePart = True
        Exit Function
    End If

    Dim outLow As String
    outLow = low

    If gender = "male" Then
        If mp_EndsWith(low, "ій") Then
            outLow = Left$(low, Len(low) - 2) & "ія"
        ElseIf mp_EndsWith(low, "ець") Then
            outLow = mp_InflectLowWordEndingEtsToGenitive(low)
        ElseIf mp_EndsWith(low, "й") Then
            outLow = Left$(low, Len(low) - 1) & "я"
        ElseIf mp_EndsWith(low, "ь") Then
            outLow = Left$(low, Len(low) - 1) & "я"
        ElseIf mp_EndsWith(low, "о") Then
            outLow = Left$(low, Len(low) - 1) & "а"
        ElseIf mp_EndsWith(low, "а") Then
            outLow = Left$(low, Len(low) - 1) & "и"
        ElseIf mp_EndsWith(low, "я") Then
            outLow = Left$(low, Len(low) - 1) & "і"
        ElseIf mp_EndsWithConsonant(low) Then
            outLow = low & "а"
        End If
    ElseIf gender = "female" Then
        If mp_EndsWith(low, "ія") Then
            outLow = Left$(low, Len(low) - 2) & "ії"
        ElseIf mp_EndsWith(low, "я") Then
            outLow = Left$(low, Len(low) - 1) & "і"
        ElseIf mp_EndsWith(low, "а") Then
            outLow = Left$(low, Len(low) - 1) & "и"
        End If
    Else
        Exit Function
    End If

    partResult = mp_ApplyWordCase(originalPart, outLow)
    mp_InflectNamePart = True
End Function

Private Function mp_InflectPatronymicPart(ByVal originalPart As String, ByVal gender As String, ByRef partResult As String) As Boolean
    Dim low As String
    low = LCase$(originalPart)

    Dim exceptions As Object
    Set exceptions = mp_GetPatronymicExceptionsDict()
    If exceptions.Exists(low) Then
        partResult = mp_ApplyWordCase(originalPart, CStr(exceptions(low)))
        mp_InflectPatronymicPart = True
        Exit Function
    End If

    Dim outLow As String
    outLow = low

    If gender = "male" Then
        If mp_EndsWith(low, "ович") Or mp_EndsWith(low, "евич") Or mp_EndsWith(low, "йович") Then
            outLow = low & "а"
        End If
    ElseIf gender = "female" Then
        If mp_EndsWith(low, "івна") Or mp_EndsWith(low, "ївна") Or mp_EndsWith(low, "овна") Or mp_EndsWith(low, "евна") Then
            outLow = Left$(low, Len(low) - 1) & "и"
        End If
    Else
        Exit Function
    End If

    partResult = mp_ApplyWordCase(originalPart, outLow)
    mp_InflectPatronymicPart = True
End Function

Private Function mp_InflectSurnamePartToAccusative(ByVal originalPart As String, ByVal gender As String, ByRef partResult As String) As Boolean
    Dim low As String
    low = LCase$(originalPart)
    low = mp_NormalizeTrailingYiSoftSign(low)

    If mp_IsIndeclinableSurname(low, gender) Then
        partResult = originalPart
        mp_InflectSurnamePartToAccusative = True
        Exit Function
    End If

    Dim outLow As String
    outLow = low

    If gender = "male" Then
        If mp_EndsWith(low, "ський") Then
            outLow = Left$(low, Len(low) - 5) & "ського"
        ElseIf mp_EndsWith(low, "цький") Then
            outLow = Left$(low, Len(low) - 5) & "цького"
        ElseIf mp_EndsWith(low, "зький") Then
            outLow = Left$(low, Len(low) - 5) & "зького"
        ElseIf mp_EndsWith(low, "ой") Or mp_EndsWith(low, "ый") Then
            outLow = Left$(low, Len(low) - 2) & "ого"
        ElseIf mp_EndsWith(low, "ець") Then
            outLow = mp_InflectLowWordEndingEtsToGenitive(low)
        ElseIf mp_EndsWith(low, "ий") Then
            outLow = Left$(low, Len(low) - 2) & "ого"
        ElseIf mp_EndsWith(low, "ній") Then
            outLow = Left$(low, Len(low) - 3) & "нього"
        ElseIf mp_EndsWith(low, "ій") Then
            outLow = Left$(low, Len(low) - 2) & "ія"
        ElseIf mp_EndsWith(low, "ко") Then
            outLow = Left$(low, Len(low) - 1) & "а"
        ElseIf mp_EndsWith(low, "а") Then
            outLow = Left$(low, Len(low) - 1) & "у"
        ElseIf mp_EndsWith(low, "я") Then
            outLow = Left$(low, Len(low) - 1) & "ю"
        ElseIf mp_EndsWith(low, "ь") Then
            outLow = Left$(low, Len(low) - 1) & "я"
        ElseIf mp_EndsWith(low, "й") Then
            outLow = Left$(low, Len(low) - 1) & "я"
        ElseIf mp_EndsWithConsonant(low) Then
            outLow = low & "а"
        End If
    ElseIf gender = "female" Then
        If mp_EndsWith(low, "ська") Then
            outLow = Left$(low, Len(low) - 4) & "ську"
        ElseIf mp_EndsWith(low, "цька") Then
            outLow = Left$(low, Len(low) - 4) & "цьку"
        ElseIf mp_EndsWith(low, "зька") Then
            outLow = Left$(low, Len(low) - 4) & "зьку"
        ElseIf mp_EndsWith(low, "а") Then
            outLow = Left$(low, Len(low) - 1) & "у"
        ElseIf mp_EndsWith(low, "я") Then
            outLow = Left$(low, Len(low) - 1) & "ю"
        End If
    Else
        Exit Function
    End If

    partResult = mp_ApplyWordCase(originalPart, outLow)
    mp_InflectSurnamePartToAccusative = True
End Function

Private Function mp_InflectNamePartToAccusative(ByVal originalPart As String, ByVal gender As String, ByRef partResult As String) As Boolean
    Dim low As String
    low = LCase$(originalPart)
    low = mp_NormalizeTrailingYiSoftSign(low)

    Dim outLow As String
    outLow = low

    If gender = "male" Then
        If mp_EndsWith(low, "ій") Then
            outLow = Left$(low, Len(low) - 2) & "ія"
        ElseIf mp_EndsWith(low, "ець") Then
            outLow = mp_InflectLowWordEndingEtsToGenitive(low)
        ElseIf mp_EndsWith(low, "й") Then
            outLow = Left$(low, Len(low) - 1) & "я"
        ElseIf mp_EndsWith(low, "ь") Then
            outLow = Left$(low, Len(low) - 1) & "я"
        ElseIf mp_EndsWith(low, "о") Then
            outLow = Left$(low, Len(low) - 1) & "а"
        ElseIf mp_EndsWith(low, "а") Then
            outLow = Left$(low, Len(low) - 1) & "у"
        ElseIf mp_EndsWith(low, "я") Then
            outLow = Left$(low, Len(low) - 1) & "ю"
        ElseIf mp_EndsWithConsonant(low) Then
            outLow = low & "а"
        End If
    ElseIf gender = "female" Then
        If mp_EndsWith(low, "ія") Then
            outLow = Left$(low, Len(low) - 2) & "ію"
        ElseIf mp_EndsWith(low, "я") Then
            outLow = Left$(low, Len(low) - 1) & "ю"
        ElseIf mp_EndsWith(low, "а") Then
            outLow = Left$(low, Len(low) - 1) & "у"
        End If
    Else
        Exit Function
    End If

    partResult = mp_ApplyWordCase(originalPart, outLow)
    mp_InflectNamePartToAccusative = True
End Function

Private Function mp_InflectPatronymicPartToAccusative(ByVal originalPart As String, ByVal gender As String, ByRef partResult As String) As Boolean
    Dim low As String
    low = LCase$(originalPart)

    Dim outLow As String
    outLow = low

    If gender = "male" Then
        If mp_EndsWith(low, "ович") Or mp_EndsWith(low, "евич") Or mp_EndsWith(low, "йович") Then
            outLow = low & "а"
        ElseIf mp_EndsWith(low, "овича") Or mp_EndsWith(low, "евича") Or mp_EndsWith(low, "йовича") Then
            outLow = low
        End If
    ElseIf gender = "female" Then
        If mp_EndsWith(low, "івна") Or mp_EndsWith(low, "ївна") Or mp_EndsWith(low, "овна") Or mp_EndsWith(low, "евна") Then
            outLow = Left$(low, Len(low) - 1) & "у"
        ElseIf mp_EndsWith(low, "івни") Or mp_EndsWith(low, "ївни") Or mp_EndsWith(low, "овни") Or mp_EndsWith(low, "евни") Then
            outLow = Left$(low, Len(low) - 1) & "у"
        End If
    Else
        Exit Function
    End If

    partResult = mp_ApplyWordCase(originalPart, outLow)
    mp_InflectPatronymicPartToAccusative = True
End Function

Private Function mp_InflectSurnamePartToDative(ByVal originalPart As String, ByVal gender As String, ByRef partResult As String) As Boolean
    Dim low As String
    low = LCase$(originalPart)
    low = mp_NormalizeTrailingYiSoftSign(low)

    If mp_IsIndeclinableSurname(low, gender) Then
        partResult = originalPart
        mp_InflectSurnamePartToDative = True
        Exit Function
    End If

    Dim outLow As String
    outLow = low

    If gender = "male" Then
        If mp_EndsWith(low, "ський") Then
            outLow = Left$(low, Len(low) - 5) & "ському"
        ElseIf mp_EndsWith(low, "цький") Then
            outLow = Left$(low, Len(low) - 5) & "цькому"
        ElseIf mp_EndsWith(low, "зький") Then
            outLow = Left$(low, Len(low) - 5) & "зькому"
        ElseIf mp_EndsWith(low, "ой") Or mp_EndsWith(low, "ый") Or mp_EndsWith(low, "ий") Then
            outLow = Left$(low, Len(low) - 2) & "ому"
        ElseIf mp_EndsWith(low, "ець") Then
            outLow = mp_InflectLowWordEndingEtsToDative(low)
        ElseIf mp_EndsWith(low, "ній") Then
            outLow = Left$(low, Len(low) - 3) & "ньому"
        ElseIf mp_EndsWith(low, "ій") Then
            outLow = Left$(low, Len(low) - 2) & "ію"
        ElseIf mp_EndsWith(low, "ко") Then
            outLow = Left$(low, Len(low) - 1) & "у"
        ElseIf mp_EndsWith(low, "а") Then
            outLow = Left$(low, Len(low) - 1) & "і"
        ElseIf mp_EndsWith(low, "я") Then
            outLow = Left$(low, Len(low) - 1) & "і"
        ElseIf mp_EndsWith(low, "ь") Then
            outLow = Left$(low, Len(low) - 1) & "ю"
        ElseIf mp_EndsWith(low, "й") Then
            outLow = Left$(low, Len(low) - 1) & "ю"
        ElseIf mp_EndsWithConsonant(low) Then
            outLow = low & "у"
        End If
    ElseIf gender = "female" Then
        If mp_EndsWith(low, "ська") Then
            outLow = Left$(low, Len(low) - 4) & "ській"
        ElseIf mp_EndsWith(low, "цька") Then
            outLow = Left$(low, Len(low) - 4) & "цькій"
        ElseIf mp_EndsWith(low, "зька") Then
            outLow = Left$(low, Len(low) - 4) & "зькій"
        ElseIf mp_EndsWith(low, "ова") Or mp_EndsWith(low, "ева") Or mp_EndsWith(low, "єва") Or _
               mp_EndsWith(low, "іна") Or mp_EndsWith(low, "їна") Or mp_EndsWith(low, "ина") Then
            outLow = Left$(low, Len(low) - 1) & "ій"
        ElseIf mp_EndsWith(low, "а") Then
            outLow = Left$(low, Len(low) - 1) & "і"
        ElseIf mp_EndsWith(low, "я") Then
            outLow = Left$(low, Len(low) - 1) & "і"
        End If
    Else
        Exit Function
    End If

    partResult = mp_ApplyWordCase(originalPart, outLow)
    mp_InflectSurnamePartToDative = True
End Function

Private Function mp_InflectNamePartToDative(ByVal originalPart As String, ByVal gender As String, ByRef partResult As String) As Boolean
    Dim low As String
    low = LCase$(originalPart)
    low = mp_NormalizeTrailingYiSoftSign(low)

    Dim exceptions As Object
    Set exceptions = mp_GetNameExceptionsDativeDict()
    If exceptions.Exists(low) Then
        partResult = mp_ApplyWordCase(originalPart, CStr(exceptions(low)))
        mp_InflectNamePartToDative = True
        Exit Function
    End If

    Dim outLow As String
    outLow = low

    If gender = "male" Then
        If mp_EndsWith(low, "ій") Then
            outLow = Left$(low, Len(low) - 2) & "ію"
        ElseIf mp_EndsWith(low, "ець") Then
            outLow = mp_InflectLowWordEndingEtsToDative(low)
        ElseIf mp_EndsWith(low, "й") Then
            outLow = Left$(low, Len(low) - 1) & "ю"
        ElseIf mp_EndsWith(low, "ь") Then
            outLow = Left$(low, Len(low) - 1) & "ю"
        ElseIf mp_EndsWith(low, "о") Then
            outLow = Left$(low, Len(low) - 1) & "ові"
        ElseIf mp_EndsWith(low, "а") Then
            outLow = Left$(low, Len(low) - 1) & "і"
        ElseIf mp_EndsWith(low, "я") Then
            outLow = Left$(low, Len(low) - 1) & "і"
        ElseIf mp_EndsWithConsonant(low) Then
            outLow = low & "у"
        End If
    ElseIf gender = "female" Then
        If mp_EndsWith(low, "ія") Then
            outLow = Left$(low, Len(low) - 2) & "ії"
        ElseIf mp_EndsWith(low, "я") Then
            outLow = Left$(low, Len(low) - 1) & "і"
        ElseIf mp_EndsWith(low, "а") Then
            outLow = Left$(low, Len(low) - 1) & "і"
        End If
    Else
        Exit Function
    End If

    partResult = mp_ApplyWordCase(originalPart, outLow)
    mp_InflectNamePartToDative = True
End Function

Private Function mp_InflectPatronymicPartToDative(ByVal originalPart As String, ByVal gender As String, ByRef partResult As String) As Boolean
    Dim low As String
    low = LCase$(originalPart)

    Dim outLow As String
    outLow = low

    If gender = "male" Then
        If mp_EndsWith(low, "ович") Or mp_EndsWith(low, "евич") Or mp_EndsWith(low, "йович") Then
            outLow = low & "у"
        ElseIf mp_EndsWith(low, "овича") Or mp_EndsWith(low, "евича") Or mp_EndsWith(low, "йовича") Then
            outLow = Left$(low, Len(low) - 1) & "у"
        End If
    ElseIf gender = "female" Then
        If mp_EndsWith(low, "івна") Or mp_EndsWith(low, "ївна") Or mp_EndsWith(low, "овна") Or mp_EndsWith(low, "евна") Then
            outLow = Left$(low, Len(low) - 1) & "і"
        ElseIf mp_EndsWith(low, "івни") Or mp_EndsWith(low, "ївни") Or mp_EndsWith(low, "овни") Or mp_EndsWith(low, "евни") Then
            outLow = low
        End If
    Else
        Exit Function
    End If

    partResult = mp_ApplyWordCase(originalPart, outLow)
    mp_InflectPatronymicPartToDative = True
End Function

Private Function mp_IsIndeclinableSurname(ByVal lowSurname As String, ByVal gender As String) As Boolean
    ' Surnames like "КАНІВСЬКИХ" are indeclinable in Ukrainian for both genders.
    If mp_EndsWith(lowSurname, "их") Or mp_EndsWith(lowSurname, "їх") Or mp_EndsWith(lowSurname, "ых") Then
        mp_IsIndeclinableSurname = True
        Exit Function
    End If

    If gender = "female" Then
        If mp_EndsWith(lowSurname, "енко") Or mp_EndsWith(lowSurname, "ко") Then
            mp_IsIndeclinableSurname = True
            Exit Function
        End If

        If mp_EndsWithConsonant(lowSurname) Or mp_EndsWith(lowSurname, "о") Then
            mp_IsIndeclinableSurname = True
        End If
    End If
End Function

Private Sub mp_SplitTrailingLineBreaks(ByVal sourceText As String, ByRef bodyText As String, ByRef trailingBreaks As String)
    bodyText = sourceText

    Do While Len(bodyText) > 0
        Dim tailChar As String
        tailChar = Right$(bodyText, 1)

        If tailChar = vbCr Or tailChar = vbLf Or AscW(tailChar) = 11 Or AscW(tailChar) = 7 Then
            trailingBreaks = tailChar & trailingBreaks
            bodyText = Left$(bodyText, Len(bodyText) - 1)
        Else
            Exit Do
        End If
    Loop
End Sub

Private Function mp_IsWhitespaceLikeChar(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    mp_IsWhitespaceLikeChar = (ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Or AscW(ch) = 160)
End Function

Private Function mp_GetNameExceptionsDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    d("ілля") = "іллі"
    d("лев") = "лева"
    d("любов") = "любові"
    d("матвій") = "матвія"
    d("лука") = "луки"

    Set mp_GetNameExceptionsDict = d
End Function

Private Function mp_GetNameExceptionsDativeDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    d("ілля") = "іллі"
    d("лев") = "леву"
    d("любов") = "любові"
    d("матвій") = "матвію"
    d("лука") = "луці"

    Set mp_GetNameExceptionsDativeDict = d
End Function

Private Function mp_GetSurnameExceptionsDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    d("середа") = "середи"

    Set mp_GetSurnameExceptionsDict = d
End Function

Private Function mp_GetPatronymicExceptionsDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set mp_GetPatronymicExceptionsDict = d
End Function

Private Function mp_GetMaleNamesDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    d("іван") = True
    d("петро") = True
    d("андрій") = True
    d("олександр") = True
    d("микола") = True
    d("богдан") = True
    d("тарас") = True
    d("дмитро") = True
    d("максим") = True
    d("василь") = True
    d("володимир") = True
    d("юрій") = True
    d("сергій") = True
    d("степан") = True
    d("роман") = True
    d("павло") = True

    Set mp_GetMaleNamesDict = d
End Function

Private Function mp_GetFemaleNamesDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    d("марія") = True
    d("олена") = True
    d("наталія") = True
    d("тетяна") = True
    d("оксана") = True
    d("ірина") = True
    d("анна") = True
    d("катерина") = True
    d("людмила") = True
    d("світлана") = True
    d("юлія") = True
    d("ольга") = True
    d("вікторія") = True

    Set mp_GetFemaleNamesDict = d
End Function
