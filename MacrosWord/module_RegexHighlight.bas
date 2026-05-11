Option Explicit

Private Const MP_REGEX_FILE_NAME As String = "Word_RegexHighlight.txt"
Private Const MP_REGEX_DEFAULT_FILL_COLOR As Long = 8388736 ' RGB(128, 0, 128)
Private Const MP_REGEX_IGNORE_CASE As Boolean = False
Private Const MP_REGEX_MULTILINE As Boolean = True
Private Const MP_REGEX_GLOBAL As Boolean = True
Private Const MP_STATUSBAR_CLEAR_DELAY As String = "00:00:02"
Private Const MP_REGEX_BOOKMARK_PREFIX As String = "rxh_"

Public Sub m_RegexHighlightByPattern()
    If Documents.Count = 0 Then
        MsgBox "Нет открытого Word-документа.", vbExclamation, "Regex Highlight"
        Exit Sub
    End If

    Dim regexFilePath As String
    Dim errorText As String

    If Not mp_TryResolveRegexFilePath(ActiveDocument, regexFilePath, errorText) Then
        MsgBox errorText, vbExclamation, "Regex Highlight"
        Exit Sub
    End If

    Dim pattern As String
    Dim preparedPattern As String
    Dim highlightGroupIndex As Long
    Dim fillColor As Long

    If Not mp_TryReadRegexConfigFromFile(regexFilePath, pattern, fillColor, errorText) Then
        MsgBox errorText, vbExclamation, "Regex Highlight"
        Exit Sub
    End If

    If Not mp_TryPrepareRegexPattern(pattern, preparedPattern, highlightGroupIndex, errorText) Then
        MsgBox errorText, vbExclamation, "Regex Highlight"
        Exit Sub
    End If

    Dim doc As Document
    Set doc = ActiveDocument

    Dim regex As Object
    On Error GoTo RegexCreateFailed
    Set regex = mp_CreateRegex(preparedPattern)
    On Error GoTo FailHighlight

    Dim targetRange As Range
    Dim highlightedCount As Long
    Dim bookmarkCount As Long
    Dim staleClearedCount As Long
    Dim staleBookmarkCount As Long
    Dim undoStarted As Boolean

    mp_BeginUndoGroup "Regex Highlight", undoStarted

    Set targetRange = doc.Content
    mp_ClearRegexHighlightsByBookmarks doc, staleClearedCount, staleBookmarkCount
    highlightedCount = mp_HighlightMatchesInRange(targetRange, regex, fillColor, highlightGroupIndex, bookmarkCount)

    mp_EndUndoGroup undoStarted
    mp_SetStatusBarMessage "Regex Highlight: подсвечено " & highlightedCount & "; закладок " & bookmarkCount
    Exit Sub

RegexCreateFailed:
    MsgBox "Некорректный regex паттерн: " & Err.Description, vbExclamation, "Regex Highlight"
    Exit Sub

FailHighlight:
    mp_EndUndoGroup undoStarted
    MsgBox "Ошибка подсветки: " & Err.Description, vbExclamation, "Regex Highlight"
End Sub

Public Sub m_RegexClearHighlightInDocument()
    If Documents.Count = 0 Then
        MsgBox "Нет открытого Word-документа.", vbExclamation, "Regex Highlight"
        Exit Sub
    End If

    On Error GoTo FailClear

    Dim doc As Document
    Set doc = ActiveDocument

    Dim clearedCount As Long
    Dim removedBookmarkCount As Long

    mp_ClearRegexHighlightsByBookmarks doc, clearedCount, removedBookmarkCount
    mp_SetStatusBarMessage "Regex Highlight: снято " & clearedCount & "; удалено закладок " & removedBookmarkCount
    Exit Sub

FailClear:
    MsgBox "Ошибка при снятии подсветки: " & Err.Description, vbExclamation, "Regex Highlight"
End Sub

Private Function mp_CreateRegex(ByVal pattern As String) As Object
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    regex.Pattern = pattern
    regex.Global = MP_REGEX_GLOBAL
    regex.IgnoreCase = MP_REGEX_IGNORE_CASE
    regex.MultiLine = MP_REGEX_MULTILINE

    Set mp_CreateRegex = regex
End Function

Private Function mp_TryPrepareRegexPattern(ByVal rawPattern As String, ByRef preparedPattern As String, ByRef highlightGroupIndex As Long, ByRef errorText As String) As Boolean
    preparedPattern = rawPattern
    highlightGroupIndex = 0

    If InStr(1, rawPattern, "(?<", vbTextCompare) = 0 Then
        mp_TryPrepareRegexPattern = True
        Exit Function
    End If

    If Not mp_TryRewriteNamedHighlightGroup(rawPattern, preparedPattern, highlightGroupIndex, errorText) Then Exit Function
    mp_TryPrepareRegexPattern = True
End Function

Private Function mp_TryRewriteNamedHighlightGroup(ByVal sourcePattern As String, ByRef rewrittenPattern As String, ByRef highlightGroupIndex As Long, ByRef errorText As String) As Boolean
    Dim i As Long
    Dim patternLength As Long
    Dim inCharClass As Boolean
    Dim captureIndex As Long
    Dim ch As String

    patternLength = Len(sourcePattern)
    i = 1
    rewrittenPattern = vbNullString

    Do While i <= patternLength
        ch = Mid$(sourcePattern, i, 1)

        If ch = "\" Then
            rewrittenPattern = rewrittenPattern & ch
            If i < patternLength Then
                rewrittenPattern = rewrittenPattern & Mid$(sourcePattern, i + 1, 1)
                i = i + 2
            Else
                i = i + 1
            End If
            GoTo ContinueLoop
        End If

        If inCharClass Then
            rewrittenPattern = rewrittenPattern & ch
            If ch = "]" Then inCharClass = False
            i = i + 1
            GoTo ContinueLoop
        End If

        If ch = "[" Then
            inCharClass = True
            rewrittenPattern = rewrittenPattern & ch
            i = i + 1
            GoTo ContinueLoop
        End If

        If ch = "(" Then
            Dim token3 As String
            token3 = Mid$(sourcePattern, i, 3)
            If token3 = "(?:" Or token3 = "(?=" Or token3 = "(?!" Then
                rewrittenPattern = rewrittenPattern & token3
                i = i + 3
                GoTo ContinueLoop
            End If

            Dim token4 As String
            token4 = Mid$(sourcePattern, i, 4)
            If token4 = "(?<=" Or token4 = "(?<!" Then
                rewrittenPattern = rewrittenPattern & token4
                i = i + 4
                GoTo ContinueLoop
            End If

            If LCase$(token3) = "(?<" Then
                Dim closingPos As Long
                closingPos = InStr(i + 3, sourcePattern, ">", vbBinaryCompare)
                If closingPos = 0 Then
                    errorText = "Незавершенная именованная группа в regex."
                    Exit Function
                End If

                captureIndex = captureIndex + 1

                Dim groupName As String
                groupName = LCase$(Mid$(sourcePattern, i + 3, closingPos - (i + 3)))
                If groupName = "rxhighlight" Then
                    If highlightGroupIndex <> 0 Then
                        errorText = "В regex найдено больше одной группы (?<rxHighlight>...)."
                        Exit Function
                    End If
                    highlightGroupIndex = captureIndex
                End If

                rewrittenPattern = rewrittenPattern & "("
                i = closingPos + 1
                GoTo ContinueLoop
            End If

            captureIndex = captureIndex + 1
            rewrittenPattern = rewrittenPattern & "("
            i = i + 1
            GoTo ContinueLoop
        End If

        rewrittenPattern = rewrittenPattern & ch
        i = i + 1

ContinueLoop:
    Loop

    mp_TryRewriteNamedHighlightGroup = True
End Function

Private Function mp_TryResolveHighlightSegment(ByVal matchText As Object, ByVal highlightGroupIndex As Long, ByRef segmentOffset As Long, ByRef segmentLength As Long) As Boolean
    Dim fullValue As String
    fullValue = CStr(matchText.Value)

    If Len(fullValue) = 0 Then Exit Function

    If highlightGroupIndex <= 0 Then
        segmentOffset = 0
        segmentLength = Len(fullValue)
        mp_TryResolveHighlightSegment = True
        Exit Function
    End If

    If highlightGroupIndex > matchText.SubMatches.Count Then
        segmentOffset = 0
        segmentLength = Len(fullValue)
        mp_TryResolveHighlightSegment = True
        Exit Function
    End If

    Dim groupValue As String
    groupValue = CStr(matchText.SubMatches(highlightGroupIndex - 1))
    If Len(groupValue) = 0 Then Exit Function

    Dim foundPos As Long
    foundPos = InStr(1, fullValue, groupValue, vbBinaryCompare)

    If foundPos <= 0 Then
        segmentOffset = 0
        segmentLength = Len(fullValue)
    Else
        segmentOffset = foundPos - 1
        segmentLength = Len(groupValue)
    End If

    mp_TryResolveHighlightSegment = True
End Function

Private Function mp_HighlightMatchesInRange(ByVal sourceRange As Range, ByVal regex As Object, ByVal fillColor As Long, ByVal highlightGroupIndex As Long, ByRef bookmarkCount As Long) As Long
    On Error GoTo HighlightRangeFailed

    Dim matches As Object
    Set matches = regex.Execute(sourceRange.Text)

    Dim i As Long
    Dim matchText As Object
    Dim hitRange As Range
    Dim hitStart As Long
    Dim segmentOffset As Long
    Dim segmentLength As Long

    For i = 0 To matches.Count - 1
        On Error GoTo MatchFailed
        Set matchText = matches(i)

        If Not mp_TryResolveHighlightSegment(matchText, highlightGroupIndex, segmentOffset, segmentLength) Then GoTo ContinueMatchLoop

        If segmentLength > 0 Then
            hitStart = sourceRange.Start + CLng(matchText.FirstIndex) + segmentOffset
            Set hitRange = mp_GetVerifiedMatchRange(sourceRange.Document, hitStart, segmentLength, Mid$(CStr(matchText.Value), segmentOffset + 1, segmentLength))
            If hitRange Is Nothing Then GoTo ContinueMatchLoop

            hitRange.HighlightColorIndex = wdNoHighlight
            With hitRange.Shading
                .Texture = wdTextureNone
                .ForegroundPatternColor = wdColorAutomatic
                .BackgroundPatternColor = fillColor
            End With
            mp_AddRegexBookmark sourceRange.Document, hitRange, bookmarkCount
            mp_HighlightMatchesInRange = mp_HighlightMatchesInRange + 1
        End If

ContinueMatchLoop:
        On Error GoTo HighlightRangeFailed
    Next i

    Exit Function

MatchFailed:
    Err.Clear
    Resume ContinueMatchLoop

HighlightRangeFailed:
    Err.Clear
End Function

Private Sub mp_ClearRegexHighlightsByBookmarks(ByVal doc As Document, ByRef clearedCount As Long, ByRef removedBookmarkCount As Long)
    On Error GoTo ClearBookmarksFailed

    Dim i As Long
    Dim currentBookmark As Bookmark
    Dim bookmarkRange As Range

    For i = doc.Bookmarks.Count To 1 Step -1
        Set currentBookmark = doc.Bookmarks(i)

        If StrComp(Left$(currentBookmark.Name, Len(MP_REGEX_BOOKMARK_PREFIX)), MP_REGEX_BOOKMARK_PREFIX, vbTextCompare) = 0 Then
            Set bookmarkRange = currentBookmark.Range
            bookmarkRange.HighlightColorIndex = wdNoHighlight
            With bookmarkRange.Shading
                .Texture = wdTextureNone
                .ForegroundPatternColor = wdColorAutomatic
                .BackgroundPatternColor = wdColorAutomatic
            End With
            clearedCount = clearedCount + 1

            currentBookmark.Delete
            removedBookmarkCount = removedBookmarkCount + 1
        End If
    Next i

    Exit Sub

ClearBookmarksFailed:
    Err.Clear
End Sub

Private Sub mp_AddRegexBookmark(ByVal doc As Document, ByVal targetRange As Range, ByRef bookmarkCount As Long)
    On Error GoTo AddBookmarkFailed

    Dim bookmarkName As String
    bookmarkName = mp_GetNextRegexBookmarkName(doc, bookmarkCount)

    doc.Bookmarks.Add Name:=bookmarkName, Range:=targetRange.Duplicate
    Exit Sub

AddBookmarkFailed:
    Err.Clear
End Sub

Private Function mp_GetNextRegexBookmarkName(ByVal doc As Document, ByRef bookmarkCount As Long) As String
    Dim candidateName As String

    Do
        bookmarkCount = bookmarkCount + 1
        candidateName = MP_REGEX_BOOKMARK_PREFIX & Format$(bookmarkCount, "00000000")
    Loop While doc.Bookmarks.Exists(candidateName)

    mp_GetNextRegexBookmarkName = candidateName
End Function

Private Function mp_GetVerifiedMatchRange(ByVal doc As Document, ByVal baseStart As Long, ByVal matchLength As Long, ByVal expectedText As String) As Range
    If matchLength <= 0 Then Exit Function

    Dim offsets(0 To 2) As Long
    offsets(0) = 0
    offsets(1) = 1
    offsets(2) = -1

    Dim i As Long
    Dim candidateStart As Long
    Dim candidateEnd As Long
    Dim contentEnd As Long
    Dim candidateRange As Range

    contentEnd = doc.Content.End

    For i = LBound(offsets) To UBound(offsets)
        candidateStart = baseStart + offsets(i)
        candidateEnd = candidateStart + matchLength

        If candidateStart >= 0 And candidateEnd <= contentEnd Then
            Set candidateRange = doc.Range(Start:=candidateStart, End:=candidateEnd)
            If StrComp(candidateRange.Text, expectedText, vbBinaryCompare) = 0 Then
                Set mp_GetVerifiedMatchRange = candidateRange
                Exit Function
            End If
        End If
    Next i
End Function

Private Function mp_TryResolveRegexFilePath(ByVal doc As Document, ByRef regexFilePath As String, ByRef errorText As String) As Boolean
    If Len(Trim$(doc.Path)) = 0 Then
        errorText = "Сначала сохраните документ, чтобы использовать regex-файл из его папки."
        Exit Function
    End If

    regexFilePath = doc.Path & "\" & MP_REGEX_FILE_NAME

    If Len(Dir$(regexFilePath, vbNormal)) = 0 Then
        errorText = "Файл с regex не найден: " & regexFilePath
        Exit Function
    End If

    mp_TryResolveRegexFilePath = True
End Function

Private Function mp_TryReadRegexConfigFromFile(ByVal regexFilePath As String, ByRef pattern As String, ByRef fillColor As Long, ByRef errorText As String) As Boolean
    Dim fullText As String
    fillColor = MP_REGEX_DEFAULT_FILL_COLOR

    If mp_TryReadFileTextUtf8(regexFilePath, fullText) Then
        mp_TryReadRegexConfigFromFile = mp_TryParseRegexConfigText(fullText, pattern, fillColor, errorText)
        Exit Function
    End If

    If mp_TryReadFileTextAnsi(regexFilePath, fullText, errorText) Then
        mp_TryReadRegexConfigFromFile = mp_TryParseRegexConfigText(fullText, pattern, fillColor, errorText)
    End If
End Function

Private Function mp_TryReadFileTextUtf8(ByVal filePath As String, ByRef textValue As String) As Boolean
    On Error GoTo Utf8ReadFailed

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    textValue = stream.ReadText(-1)
    stream.Close

    If Len(textValue) > 0 Then
        If AscW(Left$(textValue, 1)) = 65279 Then
            textValue = Mid$(textValue, 2)
        End If
    End If

    mp_TryReadFileTextUtf8 = True
    Exit Function

Utf8ReadFailed:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
End Function

Private Function mp_TryReadFileTextAnsi(ByVal filePath As String, ByRef textValue As String, ByRef errorText As String) As Boolean
    Dim fileNumber As Integer
    Dim fileIsOpen As Boolean

    On Error GoTo AnsiReadFailed

    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
    fileIsOpen = True
    textValue = Input$(LOF(fileNumber), #fileNumber)
    Close #fileNumber
    fileIsOpen = False

    mp_TryReadFileTextAnsi = True
    Exit Function

AnsiReadFailed:
    If fileIsOpen Then Close #fileNumber
    errorText = "Не удалось прочитать файл regex """ & filePath & """: " & Err.Description
End Function

Private Function mp_TryParseRegexConfigText(ByVal textValue As String, ByRef pattern As String, ByRef fillColor As Long, ByRef errorText As String) As Boolean
    pattern = vbNullString

    Dim normalizedText As String
    normalizedText = Replace(textValue, vbCrLf, vbLf)
    normalizedText = Replace(normalizedText, vbCr, vbLf)

    Dim lines() As String
    lines = Split(normalizedText, vbLf)

    Dim i As Long
    Dim lineText As String
    Dim separatorPos As Long
    Dim keyText As String
    Dim valueText As String

    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(lines(i))
        If Len(lineText) = 0 Then GoTo ContinueLoop
        If Left$(lineText, 2) = "//" Then GoTo ContinueLoop

        separatorPos = InStr(1, lineText, "=", vbBinaryCompare)
        If separatorPos <= 1 Then
            errorText = "Неверный формат строки " & (i + 1) & ": " & lineText
            Exit Function
        End If

        keyText = LCase$(Trim$(Left$(lineText, separatorPos - 1)))
        valueText = Trim$(Mid$(lineText, separatorPos + 1))
        valueText = mp_UnquoteValue(valueText)

        Select Case keyText
            Case "regex", "pattern"
                pattern = valueText

            Case "color", "highlight_color", "highlight"
                If Not mp_TryParseColorHex(valueText, fillColor, errorText) Then
                    errorText = "Ошибка в строке " & (i + 1) & ": " & errorText
                    Exit Function
                End If

            Case Else
                ' Неизвестные ключи игнорируются для удобного расширения конфига.
        End Select

ContinueLoop:
    Next i

    If Len(pattern) = 0 Then
        errorText = "В файле не найден ключ regex."
        Exit Function
    End If

    mp_TryParseRegexConfigText = True
End Function

Private Function mp_UnquoteValue(ByVal valueText As String) As String
    Dim outValue As String
    outValue = Trim$(valueText)

    If Len(outValue) >= 2 Then
        If (Left$(outValue, 1) = """" And Right$(outValue, 1) = """") Or _
           (Left$(outValue, 1) = "'" And Right$(outValue, 1) = "'") Then
            outValue = Mid$(outValue, 2, Len(outValue) - 2)
        End If
    End If

    mp_UnquoteValue = outValue
End Function

Private Function mp_TryParseColorHex(ByVal rawValue As String, ByRef fillColor As Long, ByRef errorText As String) As Boolean
    Dim valueText As String
    valueText = Trim$(rawValue)

    If Len(valueText) = 0 Then
        errorText = "Значение color пустое."
        Exit Function
    End If

    If Left$(valueText, 2) = "#<" And Right$(valueText, 1) = ">" Then
        valueText = Mid$(valueText, 3, Len(valueText) - 3)
    End If

    Dim redValue As Long
    Dim greenValue As Long
    Dim blueValue As Long

    If Not mp_TryParseHexColor(valueText, redValue, greenValue, blueValue) Then
        errorText = "color должен быть только в hex-формате #RRGGBB (или #<RRGGBB>)."
        Exit Function
    End If

    fillColor = RGB(redValue, greenValue, blueValue)
    mp_TryParseColorHex = True
End Function

Private Function mp_TryParseHexColor(ByVal rawHex As String, ByRef redValue As Long, ByRef greenValue As Long, ByRef blueValue As Long) As Boolean
    Dim hexText As String
    hexText = Trim$(rawHex)

    If Left$(hexText, 1) = "#" Then
        hexText = Mid$(hexText, 2)
    End If

    If Len(hexText) <> 6 Then Exit Function
    If Not hexText Like "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]" Then Exit Function

    redValue = CLng("&H" & Mid$(hexText, 1, 2))
    greenValue = CLng("&H" & Mid$(hexText, 3, 2))
    blueValue = CLng("&H" & Mid$(hexText, 5, 2))
    mp_TryParseHexColor = True
End Function

Private Sub mp_BeginUndoGroup(ByVal recordName As String, ByRef started As Boolean)
    On Error GoTo UndoNotAvailable

    Dim undoRecord As Object
    Set undoRecord = CallByName(Application, "UndoRecord", VbGet)
    CallByName undoRecord, "StartCustomRecord", VbMethod, recordName
    started = True
    Exit Sub

UndoNotAvailable:
    started = False
    Err.Clear
End Sub

Private Sub mp_EndUndoGroup(ByRef started As Boolean)
    On Error Resume Next

    If started Then
        Dim undoRecord As Object
        Set undoRecord = CallByName(Application, "UndoRecord", VbGet)
        CallByName undoRecord, "EndCustomRecord", VbMethod
        started = False
    End If
End Sub

Private Sub mp_SetStatusBarMessage(ByVal messageText As String)
    On Error Resume Next
    Application.StatusBar = messageText
    Application.OnTime When:=Now + TimeValue(MP_STATUSBAR_CLEAR_DELAY), Name:="m_RegexHighlight_ClearStatusBar"
End Sub

Public Sub m_RegexHighlight_ClearStatusBar()
    On Error Resume Next
    Application.StatusBar = vbNullString
End Sub
