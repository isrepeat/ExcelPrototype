Option Explicit
#Const ENABLE_LOGGING = False

Private Const cnst_REGEX_FILE_NAME As String = "Word_RegexHighlight.txt"
Private Const cnst_REGEX_DEFAULT_FILL_COLOR As Long = 8388736 ' RGB(128, 0, 128)
Private Const cnst_REGEX_IGNORE_CASE As Boolean = False
Private Const cnst_REGEX_MULTILINE As Boolean = True
Private Const cnst_REGEX_GLOBAL As Boolean = True
Private Const cnst_STATUSBAR_CLEAR_DELAY As String = "00:00:02"
Private Const cnst_REGEX_BOOKMARK_PREFIX As String = "rxh_"
Private Const cnst_DIAG_PROGRESS_INTERVAL As Long = 250

Public Sub fn_RegexHighlightByPattern()
    Dim regexFilePath As String
    Dim errorText As String
    Dim pattern As String
    Dim preparedPattern As String
    Dim highlightGroupIndex As Long
    Dim fillColor As Long
    Dim groupStyleByName As Object
    Dim namedGroupIndexes As Object
    Dim groupStyleByIndex As Object
    Dim doc As Document
    Dim regex As Object
    Dim targetRange As Range
    Dim highlightedCount As Long
    Dim bookmarkCount As Long
    Dim staleClearedCount As Long
    Dim staleBookmarkCount As Long
    Dim totalMatches As Long
    Dim skippedMatches As Long
    Dim undoStarted As Boolean

    If Documents.Count = 0 Then
#If ENABLE_LOGGING Then
        ex_Diagnostincs.fn_Diagnostic_LogWarning "RegexHighlight: no open document."
#End If
        VBA.MsgBox "No open Word document.", VBA.vbExclamation, "Regex Highlight"
        Exit Sub
    End If

#If ENABLE_LOGGING Then
    ex_Diagnostincs.fn_Diagnostic_LogInfo "RegexHighlight: start doc='" & ActiveDocument.Name & "'"
#End If

    If Not private_TryResolveRegexFilePath( _
        ActiveDocument, _
        regexFilePath, _
        errorText) Then
#If ENABLE_LOGGING Then
        ex_Diagnostincs.fn_Diagnostic_LogError "RegexHighlight: resolve regex file failed: " & errorText
#End If
        VBA.MsgBox errorText, VBA.vbExclamation, "Regex Highlight"
        Exit Sub
    End If

#If ENABLE_LOGGING Then
    ex_Diagnostincs.fn_Diagnostic_LogInfo "RegexHighlight: regex file='" & regexFilePath & "'"
#End If

    If Not private_TryReadRegexConfigFromFile( _
        regexFilePath, _
        pattern, _
        fillColor, _
        groupStyleByName, _
        errorText) Then
#If ENABLE_LOGGING Then
        ex_Diagnostincs.fn_Diagnostic_LogError "RegexHighlight: read config failed: " & errorText
#End If
        VBA.MsgBox errorText, VBA.vbExclamation, "Regex Highlight"
        Exit Sub
    End If

    If Not private_TryPrepareRegexPattern( _
        pattern, _
        preparedPattern, _
        highlightGroupIndex, _
        namedGroupIndexes, _
        errorText) Then
#If ENABLE_LOGGING Then
        ex_Diagnostincs.fn_Diagnostic_LogError "RegexHighlight: prepare regex pattern failed: " & errorText
#End If
        VBA.MsgBox errorText, VBA.vbExclamation, "Regex Highlight"
        Exit Sub
    End If

    If Not private_TryBuildGroupColorIndexMap( _
        groupStyleByName, _
        namedGroupIndexes, _
        groupStyleByIndex, _
        errorText) Then
#If ENABLE_LOGGING Then
        ex_Diagnostincs.fn_Diagnostic_LogError "RegexHighlight: resolve group colors failed: " & errorText
#End If
        VBA.MsgBox errorText, VBA.vbExclamation, "Regex Highlight"
        Exit Sub
    End If

#If ENABLE_LOGGING Then
    ex_Diagnostincs.fn_Diagnostic_LogInfo _
        "RegexHighlight: config loaded patternLength=" & Len(preparedPattern) & _
        " highlightGroupIndex=" & highlightGroupIndex & _
        " namedGroups=" & namedGroupIndexes.Count & _
        " groupColorRules=" & groupStyleByIndex.Count
#End If
    Set doc = ActiveDocument

    On Error GoTo RegexCreateFailed
    Set regex = private_CreateRegex(preparedPattern)
    On Error GoTo FailHighlight

    private_BeginUndoGroup "Regex Highlight", undoStarted

    Set targetRange = doc.Content
    private_ClearRegexHighlightsByBookmarks _
        doc, staleClearedCount, staleBookmarkCount
#If ENABLE_LOGGING Then
    ex_Diagnostincs.fn_Diagnostic_LogInfo _
        "RegexHighlight: previous highlights cleared count=" & staleClearedCount & _
        " removedBookmarks=" & staleBookmarkCount
#End If

    highlightedCount = private_HighlightMatchesInRange( _
        targetRange, _
        regex, _
        fillColor, _
        highlightGroupIndex, _
        groupStyleByIndex, _
        bookmarkCount, totalMatches, skippedMatches)

    private_EndUndoGroup undoStarted
#If ENABLE_LOGGING Then
    ex_Diagnostincs.fn_Diagnostic_LogInfo _
        "RegexHighlight: done highlighted=" & highlightedCount & _
        " matches=" & totalMatches & _
        " skipped=" & skippedMatches & _
        " bookmarks=" & bookmarkCount
#End If

    private_SetStatusBarMessage _
        "Regex Highlight: highlighted " & highlightedCount & "/" & totalMatches & _
        "; skipped " & skippedMatches & _
        "; bookmarks " & bookmarkCount
    Exit Sub

RegexCreateFailed:
#If ENABLE_LOGGING Then
    ex_Diagnostincs.fn_Diagnostic_LogException "RegexHighlight: regex-create-failed", Err.Number, Err.Description
#End If
    VBA.MsgBox "Invalid regex pattern: " & Err.Description, VBA.vbExclamation, "Regex Highlight"
    Exit Sub

FailHighlight:
    private_EndUndoGroup undoStarted
#If ENABLE_LOGGING Then
    ex_Diagnostincs.fn_Diagnostic_LogException "RegexHighlight: highlight-failed", Err.Number, Err.Description
#End If
    VBA.MsgBox "Highlight failed: " & Err.Description, VBA.vbExclamation, "Regex Highlight"
End Sub

Public Sub fn_RegexClearHighlightInDocument()
    Dim doc As Document
    Dim clearedCount As Long
    Dim removedBookmarkCount As Long

    If Documents.Count = 0 Then
#If ENABLE_LOGGING Then
        ex_Diagnostincs.fn_Diagnostic_LogWarning "RegexHighlightClear: no open document."
#End If
        VBA.MsgBox "No open Word document.", VBA.vbExclamation, "Regex Highlight"
        Exit Sub
    End If

#If ENABLE_LOGGING Then
    ex_Diagnostincs.fn_Diagnostic_LogInfo "RegexHighlightClear: start doc='" & ActiveDocument.Name & "'"
#End If

    On Error GoTo FailClear
    Set doc = ActiveDocument

    private_ClearRegexHighlightsByBookmarks _
        doc, clearedCount, removedBookmarkCount
#If ENABLE_LOGGING Then
    ex_Diagnostincs.fn_Diagnostic_LogInfo _
        "RegexHighlightClear: done cleared=" & clearedCount & _
        " removedBookmarks=" & removedBookmarkCount
#End If
    private_SetStatusBarMessage "Regex Highlight: cleared " & clearedCount & "; removed bookmarks " & removedBookmarkCount
    Exit Sub

FailClear:
#If ENABLE_LOGGING Then
    ex_Diagnostincs.fn_Diagnostic_LogException "RegexHighlightClear: failed", Err.Number, Err.Description
#End If
    VBA.MsgBox "Clear highlight failed: " & Err.Description, VBA.vbExclamation, "Regex Highlight"
End Sub

Public Sub fn_RegexHighlight_ClearStatusBar()
    On Error Resume Next
    Application.StatusBar = VBA.vbNullString
#If ENABLE_LOGGING Then
    ex_Diagnostincs.fn_Diagnostic_LogStatusBarMessage "clear", VBA.vbNullString, 0
#End If
End Sub

Private Function private_CreateRegex(ByVal pattern As String) As Object
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    regex.Pattern = pattern
    regex.Global = cnst_REGEX_GLOBAL
    regex.IgnoreCase = cnst_REGEX_IGNORE_CASE
    regex.MultiLine = cnst_REGEX_MULTILINE

    Set private_CreateRegex = regex
End Function

Private Function private_CreateTextCompareDictionary() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Set private_CreateTextCompareDictionary = dict
End Function

Private Function private_TryPrepareRegexPattern( _
    ByVal rawPattern As String, _
    ByRef preparedPattern As String, _
    ByRef highlightGroupIndex As Long, _
    ByRef namedGroupIndexes As Object, _
    ByRef errorText As String _
) As Boolean
    preparedPattern = rawPattern
    highlightGroupIndex = 0
    Set namedGroupIndexes = private_CreateTextCompareDictionary()

    If InStr(1, rawPattern, "(?<", vbTextCompare) = 0 Then
        private_TryPrepareRegexPattern = True
        Exit Function
    End If

    If Not private_TryRewriteNamedHighlightGroup( _
        rawPattern, _
        preparedPattern, _
        highlightGroupIndex, _
        namedGroupIndexes, _
        errorText) Then Exit Function
    private_TryPrepareRegexPattern = True
End Function

Private Function private_TryRewriteNamedHighlightGroup( _
    ByVal sourcePattern As String, _
    ByRef rewrittenPattern As String, _
    ByRef highlightGroupIndex As Long, _
    ByRef namedGroupIndexes As Object, _
    ByRef errorText As String _
) As Boolean
    Dim i As Long
    Dim patternLength As Long
    Dim inCharClass As Boolean
    Dim captureIndex As Long
    Dim ch As String
    Dim token3 As String
    Dim token4 As String
    Dim closingPos As Long
    Dim groupName As String

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
            token3 = Mid$(sourcePattern, i, 3)
            If token3 = "(?:" Or token3 = "(?=" Or token3 = "(?!" Then
                rewrittenPattern = rewrittenPattern & token3
                i = i + 3
                GoTo ContinueLoop
            End If

            token4 = Mid$(sourcePattern, i, 4)
            If token4 = "(?<=" Or token4 = "(?<!" Then
                rewrittenPattern = rewrittenPattern & token4
                i = i + 4
                GoTo ContinueLoop
            End If

            If LCase$(token3) = "(?<" Then
                closingPos = InStr(i + 3, sourcePattern, ">", vbBinaryCompare)
                If closingPos = 0 Then
                    errorText = "Unterminated named regex group."
                    Exit Function
                End If

                captureIndex = captureIndex + 1

                groupName = LCase$(Mid$(sourcePattern, i + 3, closingPos - (i + 3)))
                If Len(groupName) = 0 Then
                    errorText = "Named regex group cannot be empty."
                    Exit Function
                End If

                If namedGroupIndexes.Exists(groupName) Then
                    errorText = "Duplicate named regex group found (?<" & groupName & ">...)."
                    Exit Function
                End If
                namedGroupIndexes.Add groupName, captureIndex

                If groupName = "rxhighlight" Then
                    If highlightGroupIndex <> 0 Then
                        errorText = "More than one (?<rxHighlight>...) group found in regex."
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

    private_TryRewriteNamedHighlightGroup = True
End Function

Private Function private_TryResolveHighlightSegment( _
    ByVal matchText As Object, _
    ByVal highlightGroupIndex As Long, _
    ByRef segmentOffset As Long, _
    ByRef segmentLength As Long _
) As Boolean
    Dim fullValue As String
    fullValue = CStr(matchText.Value)

    If Len(fullValue) = 0 Then Exit Function

    If highlightGroupIndex <= 0 Then
        segmentOffset = 0
        segmentLength = Len(fullValue)
        private_TryResolveHighlightSegment = True
        Exit Function
    End If

    If highlightGroupIndex > matchText.SubMatches.Count Then
        segmentOffset = 0
        segmentLength = Len(fullValue)
        private_TryResolveHighlightSegment = True
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

    private_TryResolveHighlightSegment = True
End Function

Private Function private_HighlightMatchesInRange( _
    ByVal sourceRange As Range, _
    ByVal regex As Object, _
    ByVal fillColor As Long, _
    ByVal highlightGroupIndex As Long, _
    ByVal groupStyleByIndex As Object, _
    ByRef bookmarkCount As Long, _
    Optional ByRef totalMatches As Long = 0, _
    Optional ByRef skippedMatches As Long = 0 _
) As Long
    Dim matches As Object
    Dim i As Long
    Dim matchText As Object
    Dim hitRange As Range
    Dim hitStart As Long
    Dim segmentOffset As Long
    Dim segmentLength As Long
    Dim configuredGroupIndex As Variant
    Dim configuredStyle As Object
    Dim hasBackColor As Boolean
    Dim backColor As Long
    Dim hasFontColor As Boolean
    Dim fontColor As Long
    Dim hasConfiguredGroupColors As Boolean
    Dim highlightedByConfiguredGroups As Boolean

    On Error GoTo HighlightRangeFailed

    Set matches = regex.Execute(sourceRange.Text)
    totalMatches = matches.Count

#If ENABLE_LOGGING Then
    ex_Diagnostincs.fn_Diagnostic_LogInfo _
        "RegexHighlight: regex.execute matches=" & totalMatches & _
        " rangeStart=" & sourceRange.Start & _
        " rangeEnd=" & sourceRange.End & _
        " textLength=" & Len(sourceRange.Text)
#End If

    hasConfiguredGroupColors = Not groupStyleByIndex Is Nothing
    If hasConfiguredGroupColors Then hasConfiguredGroupColors = (groupStyleByIndex.Count > 0)

    For i = 0 To matches.Count - 1
        On Error GoTo MatchFailed
        Set matchText = matches(i)

        highlightedByConfiguredGroups = False
        If hasConfiguredGroupColors Then
            For Each configuredGroupIndex In groupStyleByIndex.Keys
                Set configuredStyle = Nothing
                Set configuredStyle = groupStyleByIndex(configuredGroupIndex)

                If Not private_TryResolveHighlightSegment( _
                    matchText, _
                    CLng(configuredGroupIndex), _
                    segmentOffset, _
                    segmentLength) Then GoTo ContinueConfiguredGroupLoop
                If configuredStyle Is Nothing Then GoTo ContinueConfiguredGroupLoop

                If segmentLength > 0 Then
                    hitStart = sourceRange.Start + CLng(matchText.FirstIndex) + segmentOffset
                    Set hitRange = private_GetVerifiedMatchRange( _
                        sourceRange.Document, _
                        hitStart, _
                        segmentLength, _
                        Mid$(CStr(matchText.Value), segmentOffset + 1, segmentLength))
                    If hitRange Is Nothing Then GoTo ContinueConfiguredGroupLoop

                    hasBackColor = False
                    backColor = 0
                    hasFontColor = False
                    fontColor = 0

                    If configuredStyle.Exists("has_backcolor") Then hasBackColor = CBool(configuredStyle("has_backcolor"))
                    If configuredStyle.Exists("backcolor") Then backColor = CLng(configuredStyle("backcolor"))
                    If configuredStyle.Exists("has_fontcolor") Then hasFontColor = CBool(configuredStyle("has_fontcolor"))
                    If configuredStyle.Exists("fontcolor") Then fontColor = CLng(configuredStyle("fontcolor"))

                    private_ApplyMatchStyle _
                        hitRange, _
                        hasBackColor, _
                        backColor, _
                        hasFontColor, _
                        fontColor
                    private_AddRegexBookmark sourceRange.Document, hitRange, bookmarkCount
                    private_HighlightMatchesInRange = private_HighlightMatchesInRange + 1
                    highlightedByConfiguredGroups = True
                End If

ContinueConfiguredGroupLoop:
            Next configuredGroupIndex

            If highlightedByConfiguredGroups Then GoTo ContinueMatchLoop
        End If

        If Not private_TryResolveHighlightSegment( _
            matchText, _
            highlightGroupIndex, _
            segmentOffset, _
            segmentLength) Then GoTo ContinueMatchLoop

        If segmentLength > 0 Then
            hitStart = sourceRange.Start + CLng(matchText.FirstIndex) + segmentOffset
            Set hitRange = private_GetVerifiedMatchRange( _
                sourceRange.Document, _
                hitStart, _
                segmentLength, _
                Mid$(CStr(matchText.Value), segmentOffset + 1, segmentLength))
            If hitRange Is Nothing Then GoTo ContinueMatchLoop

            private_ApplyMatchStyle _
                hitRange, _
                True, _
                fillColor, _
                False, _
                wdColorAutomatic
            private_AddRegexBookmark sourceRange.Document, hitRange, bookmarkCount
            private_HighlightMatchesInRange = private_HighlightMatchesInRange + 1
        End If

ContinueMatchLoop:
        If cnst_DIAG_PROGRESS_INTERVAL > 0 Then
            If (i + 1) Mod cnst_DIAG_PROGRESS_INTERVAL = 0 Then
#If ENABLE_LOGGING Then
                ex_Diagnostincs.fn_Diagnostic_LogVerbose _
                    "RegexHighlight: processed matches " & (i + 1) & "/" & totalMatches
#End If
            End If
        End If
        On Error GoTo HighlightRangeFailed
    Next i

    Exit Function

MatchFailed:
    skippedMatches = skippedMatches + 1
#If ENABLE_LOGGING Then
    ex_Diagnostincs.fn_Diagnostic_LogWarning _
        "RegexHighlight: match skipped at index=" & i & " err='" & Err.Description & "'"
#End If
    Err.Clear
    Resume ContinueMatchLoop

HighlightRangeFailed:
#If ENABLE_LOGGING Then
    ex_Diagnostincs.fn_Diagnostic_LogException "RegexHighlight: range failed", Err.Number, Err.Description
#End If
    Err.Clear
End Function

Private Sub private_ApplyMatchStyle( _
    ByVal hitRange As Range, _
    ByVal hasBackColor As Boolean, _
    ByVal backColor As Long, _
    ByVal hasFontColor As Boolean, _
    ByVal fontColor As Long _
)
    hitRange.HighlightColorIndex = wdNoHighlight
    With hitRange.Shading
        .Texture = wdTextureNone
        .ForegroundPatternColor = wdColorAutomatic
        .BackgroundPatternColor = wdColorAutomatic
    End With
    hitRange.Font.Color = wdColorAutomatic

    If hasBackColor Then
        hitRange.Shading.BackgroundPatternColor = backColor
    End If

    If hasFontColor Then
        hitRange.Font.Color = fontColor
    End If
End Sub

Private Sub private_ClearRegexHighlightsByBookmarks( _
    ByVal doc As Document, _
    ByRef clearedCount As Long, _
    ByRef removedBookmarkCount As Long _
)
    Dim i As Long
    Dim currentBookmark As Bookmark
    Dim bookmarkRange As Range

    On Error GoTo ClearBookmarksFailed

    For i = doc.Bookmarks.Count To 1 Step -1
        Set currentBookmark = doc.Bookmarks(i)

        If StrComp(Left$(currentBookmark.Name, Len(cnst_REGEX_BOOKMARK_PREFIX)), cnst_REGEX_BOOKMARK_PREFIX, vbTextCompare) = 0 Then
            Set bookmarkRange = currentBookmark.Range
            bookmarkRange.HighlightColorIndex = wdNoHighlight
            With bookmarkRange.Shading
                .Texture = wdTextureNone
                .ForegroundPatternColor = wdColorAutomatic
                .BackgroundPatternColor = wdColorAutomatic
            End With
            bookmarkRange.Font.Color = wdColorAutomatic
            clearedCount = clearedCount + 1

            currentBookmark.Delete
            removedBookmarkCount = removedBookmarkCount + 1
        End If
    Next i

    Exit Sub

ClearBookmarksFailed:
#If ENABLE_LOGGING Then
    ex_Diagnostincs.fn_Diagnostic_LogException "RegexHighlight: clear bookmarks failed", Err.Number, Err.Description
#End If
    Err.Clear
End Sub

Private Sub private_AddRegexBookmark( _
    ByVal doc As Document, _
    ByVal targetRange As Range, _
    ByRef bookmarkCount As Long _
)
    Dim bookmarkName As String

    On Error GoTo AddBookmarkFailed

    bookmarkName = private_GetNextRegexBookmarkName(doc, bookmarkCount)

    doc.Bookmarks.Add Name:=bookmarkName, Range:=targetRange.Duplicate
    Exit Sub

AddBookmarkFailed:
    Err.Clear
End Sub

Private Function private_GetNextRegexBookmarkName(ByVal doc As Document, ByRef bookmarkCount As Long) As String
    Dim candidateName As String

    Do
        bookmarkCount = bookmarkCount + 1
        candidateName = cnst_REGEX_BOOKMARK_PREFIX & Format$(bookmarkCount, "00000000")
    Loop While doc.Bookmarks.Exists(candidateName)

    private_GetNextRegexBookmarkName = candidateName
End Function

Private Function private_GetVerifiedMatchRange( _
    ByVal doc As Document, _
    ByVal baseStart As Long, _
    ByVal matchLength As Long, _
    ByVal expectedText As String _
) As Range
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
                Set private_GetVerifiedMatchRange = candidateRange
                Exit Function
            End If
        End If
    Next i
End Function

Private Function private_TryResolveRegexFilePath( _
    ByVal doc As Document, _
    ByRef regexFilePath As String, _
    ByRef errorText As String _
) As Boolean
    If Len(Trim$(doc.Path)) = 0 Then
        errorText = "Save the document first to use the regex file from its folder."
        Exit Function
    End If

    regexFilePath = doc.Path & "\" & cnst_REGEX_FILE_NAME

    If Len(Dir$(regexFilePath, vbNormal)) = 0 Then
        errorText = "Regex file not found: " & regexFilePath
        Exit Function
    End If

    private_TryResolveRegexFilePath = True
End Function

Private Function private_TryBuildGroupColorIndexMap( _
    ByVal groupStyleByName As Object, _
    ByVal namedGroupIndexes As Object, _
    ByRef groupStyleByIndex As Object, _
    ByRef errorText As String _
) As Boolean
    Dim groupName As Variant
    Dim captureIndex As Long

    Set groupStyleByIndex = private_CreateTextCompareDictionary()

    If groupStyleByName Is Nothing Then
        private_TryBuildGroupColorIndexMap = True
        Exit Function
    End If

    If groupStyleByName.Count = 0 Then
        private_TryBuildGroupColorIndexMap = True
        Exit Function
    End If

    If namedGroupIndexes Is Nothing Or namedGroupIndexes.Count = 0 Then
        errorText = "Config has group_color.<name>, but regex has no named groups."
        Exit Function
    End If

    For Each groupName In groupStyleByName.Keys
        If Not namedGroupIndexes.Exists(CStr(groupName)) Then
            errorText = "Named regex group for group_color." & CStr(groupName) & " was not found."
            Exit Function
        End If

        captureIndex = CLng(namedGroupIndexes(CStr(groupName)))
        Set groupStyleByIndex(CStr(captureIndex)) = groupStyleByName(groupName)
    Next groupName

    private_TryBuildGroupColorIndexMap = True
End Function

Private Function private_TryReadRegexConfigFromFile( _
    ByVal regexFilePath As String, _
    ByRef pattern As String, _
    ByRef fillColor As Long, _
    ByRef groupStyleByName As Object, _
    ByRef errorText As String _
) As Boolean
    Dim fullText As String

    fillColor = cnst_REGEX_DEFAULT_FILL_COLOR
    Set groupStyleByName = private_CreateTextCompareDictionary()

    If private_TryReadFileTextUtf8(regexFilePath, fullText) Then
        private_TryReadRegexConfigFromFile = private_TryParseRegexConfigText( _
            fullText, _
            pattern, _
            fillColor, _
            groupStyleByName, _
            errorText)
        Exit Function
    End If

    If private_TryReadFileTextAnsi( _
        regexFilePath, _
        fullText, _
        errorText) Then
        private_TryReadRegexConfigFromFile = private_TryParseRegexConfigText( _
            fullText, _
            pattern, _
            fillColor, _
            groupStyleByName, _
            errorText)
    End If
End Function

Private Function private_TryReadFileTextUtf8(ByVal filePath As String, ByRef textValue As String) As Boolean
    Dim stream As Object

    On Error GoTo Utf8ReadFailed

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

    private_TryReadFileTextUtf8 = True
    Exit Function

Utf8ReadFailed:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
End Function

Private Function private_TryReadFileTextAnsi( _
    ByVal filePath As String, _
    ByRef textValue As String, _
    ByRef errorText As String _
) As Boolean
    Dim fileNumber As Integer
    Dim fileIsOpen As Boolean

    On Error GoTo AnsiReadFailed

    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
    fileIsOpen = True
    textValue = Input$(LOF(fileNumber), #fileNumber)
    Close #fileNumber
    fileIsOpen = False

    private_TryReadFileTextAnsi = True
    Exit Function

AnsiReadFailed:
    If fileIsOpen Then Close #fileNumber
    errorText = "Failed to read regex file \"" & filePath & "\": " & Err.Description
End Function

Private Function private_TryParseRegexConfigText( _
    ByVal textValue As String, _
    ByRef pattern As String, _
    ByRef fillColor As Long, _
    ByRef groupStyleByName As Object, _
    ByRef errorText As String _
) As Boolean
    Dim normalizedText As String
    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim separatorPos As Long
    Dim keyText As String
    Dim valueText As String
    Dim groupName As String
    Dim styleRule As Object

    pattern = VBA.vbNullString
    Set groupStyleByName = private_CreateTextCompareDictionary()

    normalizedText = Replace(textValue, vbCrLf, vbLf)
    normalizedText = Replace(normalizedText, vbCr, vbLf)
    lines = Split(normalizedText, vbLf)

    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(lines(i))
        If Len(lineText) = 0 Then GoTo ContinueLoop
        If Left$(lineText, 2) = "//" Then GoTo ContinueLoop

        separatorPos = InStr(1, lineText, "=", vbBinaryCompare)
        If separatorPos <= 1 Then
            errorText = "Invalid config format at line " & (i + 1) & ": " & lineText
            Exit Function
        End If

        keyText = LCase$(Trim$(Left$(lineText, separatorPos - 1)))
        valueText = Trim$(Mid$(lineText, separatorPos + 1))
        valueText = private_UnquoteValue(valueText)

        Select Case keyText
            Case "regex", "pattern"
                pattern = valueText

            Case "color", "highlight_color", "highlight"
                If Not private_TryParseColorHex( _
                    valueText, _
                    fillColor, _
                    errorText) Then
                    errorText = "Error at line " & (i + 1) & ": " & errorText
                    Exit Function
                End If

            Case Else
                If Left$(keyText, 12) = "group_color." Then
                    groupName = Trim$(Mid$(keyText, 13))
                    If Len(groupName) = 0 Then
                        errorText = "Error at line " & (i + 1) & ": empty group name in group_color.<name>."
                        Exit Function
                    End If

                    Set styleRule = Nothing
                    If Not private_TryParseGroupStyleValue( _
                        valueText, _
                        styleRule, _
                        errorText) Then
                        errorText = "Error at line " & (i + 1) & ": " & errorText
                        Exit Function
                    End If

                    Set groupStyleByName(groupName) = styleRule
                End If

                ' Неизвестные ключи игнорируются для удобного расширения конфига.
        End Select

ContinueLoop:
    Next i

    If Len(pattern) = 0 Then
        errorText = "Key 'regex' was not found in config file."
        Exit Function
    End If

    private_TryParseRegexConfigText = True
End Function

Private Function private_TryParseGroupStyleValue( _
    ByVal rawValue As String, _
    ByRef styleRule As Object, _
    ByRef errorText As String _
) As Boolean
    Dim valueText As String
    Dim bodyText As String
    Dim sections() As String
    Dim section As Variant
    Dim sectionText As String
    Dim separatorPos As Long
    Dim styleKey As String
    Dim styleValue As String
    Dim hasBackColor As Boolean
    Dim backColor As Long
    Dim hasFontColor As Boolean
    Dim fontColor As Long

    valueText = Trim$(rawValue)
    If Len(valueText) = 0 Then
        errorText = "group_color value cannot be empty."
        Exit Function
    End If

    If Left$(valueText, 1) <> "{" Or Right$(valueText, 1) <> "}" Then
        errorText = "group_color must use object format: {backcolor:#RRGGBB; fontcolor:#RRGGBB}."
        Exit Function
    End If

    bodyText = Mid$(valueText, 2, Len(valueText) - 2)
    sections = Split(bodyText, ";")

    For Each section In sections
        sectionText = Trim$(CStr(section))
        If Len(sectionText) = 0 Then GoTo ContinueSectionLoop

        separatorPos = InStr(1, sectionText, ":", vbBinaryCompare)
        If separatorPos <= 1 Then
            errorText = "Invalid group_color object segment: " & sectionText
            Exit Function
        End If

        styleKey = LCase$(Trim$(Left$(sectionText, separatorPos - 1)))
        styleValue = Trim$(Mid$(sectionText, separatorPos + 1))
        styleValue = private_UnquoteValue(styleValue)

        Select Case styleKey
            Case "backcolor", "background", "bgcolor", "fillcolor", "highlight", "color"
                If Not private_TryParseColorHex( _
                    styleValue, _
                    backColor, _
                    errorText) Then Exit Function
                hasBackColor = True

            Case "fontcolor", "textcolor", "forecolor", "font"
                If Not private_TryParseColorHex( _
                    styleValue, _
                    fontColor, _
                    errorText) Then Exit Function
                hasFontColor = True

            Case Else
                errorText = "Unsupported group_color style key: " & styleKey
                Exit Function
        End Select

ContinueSectionLoop:
    Next section

    If Not hasBackColor And Not hasFontColor Then
        errorText = "group_color object must define backcolor and/or fontcolor."
        Exit Function
    End If

    Set styleRule = private_CreateGroupStyleRule( _
        hasBackColor, _
        backColor, _
        hasFontColor, _
        fontColor)
    private_TryParseGroupStyleValue = True
End Function

Private Function private_CreateGroupStyleRule( _
    ByVal hasBackColor As Boolean, _
    ByVal backColor As Long, _
    ByVal hasFontColor As Boolean, _
    ByVal fontColor As Long _
) As Object
    Dim styleRule As Object

    Set styleRule = private_CreateTextCompareDictionary()
    styleRule("has_backcolor") = hasBackColor
    styleRule("backcolor") = backColor
    styleRule("has_fontcolor") = hasFontColor
    styleRule("fontcolor") = fontColor
    Set private_CreateGroupStyleRule = styleRule
End Function

Private Function private_UnquoteValue(ByVal valueText As String) As String
    Dim outValue As String
    outValue = Trim$(valueText)

    If Len(outValue) >= 2 Then
        If (Left$(outValue, 1) = """" And Right$(outValue, 1) = """") Or _
           (Left$(outValue, 1) = "'" And Right$(outValue, 1) = "'") Then
            outValue = Mid$(outValue, 2, Len(outValue) - 2)
        End If
    End If

    private_UnquoteValue = outValue
End Function

Private Function private_TryParseColorHex( _
    ByVal rawValue As String, _
    ByRef fillColor As Long, _
    ByRef errorText As String _
) As Boolean
    Dim valueText As String
    Dim redValue As Long
    Dim greenValue As Long
    Dim blueValue As Long

    valueText = Trim$(rawValue)

    If Len(valueText) = 0 Then
        errorText = "Color value is empty."
        Exit Function
    End If

    If Left$(valueText, 2) = "#<" And Right$(valueText, 1) = ">" Then
        valueText = Mid$(valueText, 3, Len(valueText) - 3)
    End If

    If Not private_TryParseHexColor( _
        valueText, _
        redValue, _
        greenValue, _
        blueValue) Then
        errorText = "Color must be in hex format #RRGGBB (or #<RRGGBB>)."
        Exit Function
    End If

    fillColor = RGB(redValue, greenValue, blueValue)
    private_TryParseColorHex = True
End Function

Private Function private_TryParseHexColor( _
    ByVal rawHex As String, _
    ByRef redValue As Long, _
    ByRef greenValue As Long, _
    ByRef blueValue As Long _
) As Boolean
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
    private_TryParseHexColor = True
End Function

Private Sub private_BeginUndoGroup(ByVal recordName As String, ByRef started As Boolean)
    Dim undoRecord As Object

    On Error GoTo UndoNotAvailable

    Set undoRecord = CallByName(Application, "UndoRecord", VbGet)
    CallByName undoRecord, "StartCustomRecord", VbMethod, recordName
    started = True
    Exit Sub

UndoNotAvailable:
    started = False
    Err.Clear
End Sub

Private Sub private_EndUndoGroup(ByRef started As Boolean)
    Dim undoRecord As Object

    On Error Resume Next

    If started Then
        Set undoRecord = CallByName(Application, "UndoRecord", VbGet)
        CallByName undoRecord, "EndCustomRecord", VbMethod
        started = False
    End If
End Sub

Private Sub private_SetStatusBarMessage(ByVal messageText As String)
    On Error Resume Next
    Application.StatusBar = messageText
    Application.OnTime When:=VBA.Now + VBA.TimeValue(cnst_STATUSBAR_CLEAR_DELAY), Name:="fn_RegexHighlight_ClearStatusBar"
#If ENABLE_LOGGING Then
    ex_Diagnostincs.fn_Diagnostic_LogStatusBarMessage "show", messageText, 2
#End If
End Sub
