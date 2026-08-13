Sub RegexReplaceInMatchedRanges()

    Dim rangePattern As String
    Dim replacePattern As String
    Dim replacement As String

    Dim reRanges As Object
    Dim reReplace As Object

    Dim matches As Object
    Dim m As Object

    Dim docText As String
    Dim rangeText As String
    Dim newRangeText As String
    Dim wordFindText As String

    Dim rng As Range
    Dim findRange As Range
    Dim searchEnd As Long
    Dim processedCount As Long
    Dim errorDescription As String

    On Error GoTo EH

    rangePattern = InputBox( _
        "Enter regex to identify target ranges:", _
        "Range Pattern")

    If rangePattern = "" Then Exit Sub

    replacePattern = InputBox( _
        "Enter regex to replace inside those ranges:", _
        "Replace Pattern")

    If replacePattern = "" Then Exit Sub

    replacement = InputBox( _
        "Enter replacement text ($1, $2, ... supported):", _
        "Replacement")

    docText = ActiveDocument.Content.Text

    Set reRanges = CreateObject("VBScript.RegExp")
    reRanges.Global = True
    reRanges.Multiline = True
    reRanges.IgnoreCase = False
    reRanges.pattern = rangePattern

    Set reReplace = CreateObject("VBScript.RegExp")
    reReplace.Global = True
    reReplace.Multiline = True
    reReplace.IgnoreCase = False
    reReplace.pattern = replacePattern

    Set matches = reRanges.Execute(docText)

    If matches.count = 0 Then
        ' Для буквальных шаблонов повторяем путь WORD Exporter напрямую:
        ' Word Find -> Range.Text. Это также покрывает текст, который Word
        ' отображает в основном story, но не возвращает идентично в снимке
        ' Content.Text для VBScript.RegExp.
        processedCount = private_ReplaceLiteralWordRanges( _
            ActiveDocument, rangePattern, reReplace, replacement)
        If processedCount > 0 Then
            MsgBox processedCount & " range(s) processed.", _
                   vbInformation, _
                   "Regex Replace"
            Exit Sub
        End If

        MsgBox "В актуальном тексте документа не найдено совпадений для:" & _
               VBA.vbCrLf & rangePattern, _
               vbExclamation, _
               "Regex Replace"
        Exit Sub
    End If

    Dim i As Long
    searchEnd = ActiveDocument.Content.End

    For i = matches.count - 1 To 0 Step -1

        Set m = matches(i)

        ' FirstIndex относится к строке Content.Text и может расходиться с
        ' координатами Word Range из-за полей и служебных символов. Как и WORD
        ' Exporter, повторно находим фактический диапазон средствами Word.
        wordFindText = VBA.CStr(m.Value)
        wordFindText = VBA.Replace(wordFindText, "^", "^^")
        wordFindText = VBA.Replace(wordFindText, VBA.vbCrLf, "^p")
        wordFindText = VBA.Replace(wordFindText, VBA.vbCr, "^p")
        wordFindText = VBA.Replace(wordFindText, VBA.vbLf, "^l")
        wordFindText = VBA.Replace(wordFindText, VBA.Chr$(11), "^l")
        wordFindText = VBA.Replace(wordFindText, VBA.vbTab, "^t")

        If VBA.Len(wordFindText) = 0 Then
            Err.Raise VBA.vbObjectError + 1, _
                      "RegexReplaceInMatchedRanges", _
                      "Range pattern returned an empty match."
        End If

        Set findRange = ActiveDocument.Range( _
            Start:=ActiveDocument.Content.Start, _
            End:=searchEnd)
        With findRange.Find
            .ClearFormatting
            .Text = wordFindText
            .Forward = False
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = True
            .MatchWildcards = False
        End With
        If Not findRange.Find.Execute Then
            Err.Raise VBA.vbObjectError + 2, _
                      "RegexReplaceInMatchedRanges", _
                      "Word could not resolve regex match to a document Range: " & _
                      VBA.CStr(m.Value)
        End If
        Set rng = findRange.Duplicate
        searchEnd = rng.Start

        rangeText = rng.Text

        newRangeText = reReplace.Replace(rangeText, replacement)

        ' Сначала VBScript.RegExp обрабатывает ссылки $1, $2, ...,
        ' затем текстовые escape-последовательности превращаются
        ' в реальные управляющие символы Word.
        newRangeText = private_DecodeReplacementEscapes(newRangeText)

        ' Как и WORD Exporter, записываем весь заранее нормализованный блок
        ' одним присваиванием Range.Text.
        rng.Text = newRangeText

    Next i

    MsgBox matches.count & " range(s) processed.", _
           vbInformation, _
           "Regex Replace"

    Exit Sub

EH:
    errorDescription = Err.Description
    MsgBox "Regex replace failed: " & errorDescription, _
           vbExclamation, _
           "Regex Replace"

End Sub

Private Function private_ReplaceLiteralWordRanges( _
    ByVal wordDoc As Document, _
    ByVal targetText As String, _
    ByVal reReplace As Object, _
    ByVal replacement As String _
) As Long
    Dim rootStoryRange As Range
    Dim storyRange As Range
    Dim nextStoryRange As Range
    Dim findRange As Range
    Dim matchedRange As Range
    Dim searchEnd As Long
    Dim newRangeText As String
    Dim found As Boolean

    For Each rootStoryRange In wordDoc.StoryRanges
        Set storyRange = rootStoryRange

        Do While Not storyRange Is Nothing
            Set nextStoryRange = storyRange.NextStoryRange
            searchEnd = storyRange.End

            Do While searchEnd > storyRange.Start
                Set findRange = storyRange.Duplicate
                findRange.SetRange _
                    Start:=storyRange.Start, _
                    End:=searchEnd

                With findRange.Find
                    .ClearFormatting
                    .Text = targetText
                    .Forward = False
                    .Wrap = wdFindStop
                    .Format = False
                    .MatchCase = True
                    .MatchWildcards = False
                    found = .Execute
                End With

                If Not found Then Exit Do

                Set matchedRange = findRange.Duplicate
                searchEnd = matchedRange.Start
                newRangeText = reReplace.Replace( _
                    matchedRange.Text, replacement)
                newRangeText = private_DecodeReplacementEscapes( _
                    newRangeText)

                ' Тот же контракт, что у WORD Exporter: нормализованный блок
                ' записывается в фактический Word Range одним присваиванием.
                matchedRange.Text = newRangeText
                private_ReplaceLiteralWordRanges = _
                    private_ReplaceLiteralWordRanges + 1
            Loop

            Set storyRange = nextStoryRange
        Loop
    Next rootStoryRange
End Function

Private Function private_DecodeReplacementEscapes( _
    ByVal valueText As String _
) As String
    valueText = VBA.Replace(valueText, "\r", VBA.vbCr)
    valueText = VBA.Replace(valueText, "\n", VBA.vbLf)
    valueText = VBA.Replace(valueText, "\t", VBA.vbTab)
    private_DecodeReplacementEscapes = valueText
End Function
