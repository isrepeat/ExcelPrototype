Option Explicit

Private Const MP_DATE_TOKEN_PATTERN As String = "\{#(?:[+-]\d+)?\}"
Private Const MP_TEMPLATE_START_LINE As String = "===================TEMPLATE_START====================="
Private Const MP_TEMPLATE_END_LINE As String = "====================TEMPLATE_END======================"
Private Const MP_STATUSBAR_CLEAR_DELAY As String = "00:00:03"
Private m_LastCollectedText As String
Private m_LastShortFormCache As String
Private m_ViewStateCaptured As Boolean
Private m_ViewDoc As Document
Private m_ViewWindow As Window
Private m_ViewSelectionStart As Long
Private m_ViewSelectionEnd As Long
Private m_ViewSelectionStoryType As Long
Private m_ViewVerticalScroll As Long
Private m_ViewHorizontalScroll As Long

Public Sub m_DrillOrder_InitializeDate()
    If Documents.Count = 0 Then
        MsgBox "No Word document is open.", vbExclamation, "Order Initialization"
        Exit Sub
    End If

    Dim dateInput As String
    dateInput = Trim$(InputBox$("Enter day number (1-31):", "Order Initialization"))

    If Len(dateInput) = 0 Then
        mp_SetStatusBarMessage "Warning: day number was not entered."
        Exit Sub
    End If

    If Not IsNumeric(dateInput) Then
        MsgBox "Enter a number from 1 to 31.", vbExclamation, "Order Initialization"
        Exit Sub
    End If

    Dim dateNumber As Long
    dateNumber = CLng(dateInput)

    If dateNumber < 1 Or dateNumber > 31 Then
        MsgBox "Day number must be in range 1-31.", vbExclamation, "Order Initialization"
        Exit Sub
    End If

    Dim tokenMap As Object
    Dim foundTokenCount As Long
    Dim parseError As String

    If Not mp_BuildDateTokenMap(ActiveDocument, dateNumber, tokenMap, foundTokenCount, parseError) Then
        MsgBox parseError, vbExclamation, "Order Initialization"
        Exit Sub
    End If

    If foundTokenCount = 0 Then
        mp_SetStatusBarMessage "Warning: tokens {#}, {#+N}, {#-N} were not found in the document."
        Exit Sub
    End If

    Dim undoStarted As Boolean
    mp_BeginUndoGroup "Order Initialization", undoStarted

    On Error GoTo FailApply

    Dim replacementsCount As Long
    replacementsCount = mp_ApplyTokenMapInDocument(ActiveDocument, tokenMap)

    If replacementsCount = 0 Then
        MsgBox "Tokens were found, but replacement failed.", vbExclamation, "Order Initialization"
        GoTo Finalize
    End If

    mp_SetStatusBarMessage "Done. Tokens replaced: " & replacementsCount

Finalize:
    mp_EndUndoGroup undoStarted
    Exit Sub

FailApply:
    mp_EndUndoGroup undoStarted
    MsgBox "Execution error: " & Err.Description, vbExclamation, "Order Initialization"
End Sub

Public Sub m_DrillOrder_InsertTemplateBlockBelowSelection()
    If Documents.Count = 0 Then
        MsgBox "No Word document is open.", vbExclamation, "Insert Template Block"
        Exit Sub
    End If

    If Selection Is Nothing Then
        MsgBox "Unable to determine current cursor position.", vbExclamation, "Insert Template Block"
        Exit Sub
    End If

    If Selection.Range.StoryType <> wdMainTextStory Then
        MsgBox "Cursor must be in the main document text.", vbExclamation, "Insert Template Block"
        Exit Sub
    End If

    Dim doc As Document
    Set doc = ActiveDocument

    Dim sourceStart As Long
    Dim sourceEnd As Long
    Dim insertAt As Long
    Dim errorText As String

    If Not mp_TryGetTemplateBlockBoundsAtPosition(doc, MP_TEMPLATE_START_LINE, MP_TEMPLATE_END_LINE, Selection.Range.Start, sourceStart, sourceEnd, insertAt, errorText) Then
        MsgBox errorText, vbExclamation, "Insert Template Block"
        Exit Sub
    End If

    Dim sourceRange As Range
    Set sourceRange = doc.Range(Start:=sourceStart, End:=sourceEnd)
    If Not mp_RangeHasMeaningfulText(sourceRange) Then
        MsgBox "No template text found between TEMPLATE_START and TEMPLATE_END.", vbExclamation, "Insert Template Block"
        Exit Sub
    End If

    Dim undoStarted As Boolean
    mp_BeginUndoGroup "Insert Template Block", undoStarted
    On Error GoTo FailInsert

    Dim insertAnchor As Range
    Dim pasteRange As Range
    Dim insertPos As Long
    Dim insertedLength As Long

    Set insertAnchor = mp_GetCollapsedRangeAt(doc, insertAt)
    If insertAnchor Is Nothing Then
        MsgBox "Unable to determine insertion position below template.", vbExclamation, "Insert Template Block"
        GoTo Finalize
    End If

    ' Добавляем пустую строку перед вставляемым шаблонным блоком.
    insertAnchor.Text = vbCr
    insertAnchor.Collapse wdCollapseEnd

    insertPos = insertAnchor.Start
    insertAnchor.FormattedText = sourceRange.FormattedText

    insertedLength = insertAnchor.End - insertPos
    If insertedLength <= 0 Then
        MsgBox "Unable to insert a copy of the template block.", vbExclamation, "Insert Template Block"
        GoTo Finalize
    End If

    Set pasteRange = doc.Range(Start:=insertPos, End:=insertPos + insertedLength)
    pasteRange.Font.Color = wdColorAutomatic

    mp_SetStatusBarMessage "Template block inserted below delimiter."

Finalize:
    mp_EndUndoGroup undoStarted
    Exit Sub

FailInsert:
    mp_EndUndoGroup undoStarted
    MsgBox "Template block insertion error: " & Err.Description, vbExclamation, "Insert Template Block"
End Sub

Private Function mp_BuildDateTokenMap(ByVal doc As Document, ByVal baseDate As Long, ByRef tokenMap As Object, ByRef foundTokenCount As Long, ByRef errorText As String) As Boolean
    Set tokenMap = CreateObject("Scripting.Dictionary")

    Dim regex As Object
    Set regex = mp_CreateDateTokenRegex()

    Dim story As Range
    Dim currentRange As Range
    Dim matches As Object
    Dim i As Long
    Dim token As String
    Dim replacementText As String

    For Each story In doc.StoryRanges
        Set currentRange = story

        Do While Not currentRange Is Nothing
            Set matches = regex.Execute(currentRange.Text)
            foundTokenCount = foundTokenCount + matches.Count

            For i = 0 To matches.Count - 1
                token = matches(i).Value

                If Not tokenMap.Exists(token) Then
                    If Not mp_TryResolveDateToken(token, baseDate, replacementText, errorText) Then
                        Exit Function
                    End If
                    tokenMap.Add token, replacementText
                End If
            Next i

            Set currentRange = currentRange.NextStoryRange
        Loop
    Next story

    mp_BuildDateTokenMap = True
End Function

Private Function mp_ApplyTokenMapInDocument(ByVal doc As Document, ByVal tokenMap As Object) As Long
    Dim story As Range
    Dim currentRange As Range
    Dim token As Variant

    For Each story In doc.StoryRanges
        Set currentRange = story

        Do While Not currentRange Is Nothing
            For Each token In tokenMap.Keys
                mp_ApplyTokenMapInDocument = mp_ApplyTokenMapInDocument + _
                    mp_ReplaceTokenInRange(currentRange, CStr(token), CStr(tokenMap(token)))
            Next token

            Set currentRange = currentRange.NextStoryRange
        Loop
    Next story
End Function

Private Function mp_ReplaceTokenInRange(ByVal sourceRange As Range, ByVal token As String, ByVal replacementText As String) As Long
    Dim findRange As Range
    Set findRange = sourceRange.Duplicate

    With findRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = token
        .Replacement.Text = replacementText
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With

    Do While findRange.Find.Execute(Replace:=wdReplaceOne)
        mp_ReplaceTokenInRange = mp_ReplaceTokenInRange + 1
        findRange.Collapse wdCollapseEnd
    Loop
End Function

Private Function mp_CreateDateTokenRegex() As Object
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = False
    regex.Pattern = MP_DATE_TOKEN_PATTERN
    Set mp_CreateDateTokenRegex = regex
End Function

Private Function mp_TryResolveDateToken(ByVal token As String, ByVal baseDate As Long, ByRef replacementText As String, ByRef errorText As String) As Boolean
    Dim offsetPart As String
    offsetPart = Mid$(token, 3, Len(token) - 3)

    Dim offsetValue As Long
    If Len(offsetPart) = 0 Then
        offsetValue = 0
    Else
        On Error GoTo ParseError
        offsetValue = CLng(offsetPart)
    End If

    Dim resolvedDate As Long
    resolvedDate = baseDate + offsetValue
    If resolvedDate < 1 Or resolvedDate > 31 Then
        errorText = "Token """ & token & """ with base day " & Format$(baseDate, "00") & _
                    " gives invalid day " & resolvedDate & ". Allowed range is 1..31."
        Exit Function
    End If

    replacementText = Format$(resolvedDate, "00")
    mp_TryResolveDateToken = True
    Exit Function

ParseError:
    errorText = "Unable to parse token """ & token & """. Allowed tokens: {#}, {#+N}, {#-N}."
End Function

Private Function mp_DuplicateTemplateBlocksBelowByDelimiter(ByVal doc As Document, ByVal startLine As String, ByVal endLine As String) As Long
    Dim blockBounds As Collection
    Set blockBounds = mp_CollectTemplateBlockBounds(doc, startLine, endLine)
    If blockBounds Is Nothing Then Exit Function
    If blockBounds.Count = 0 Then Exit Function

    Dim i As Long
    Dim bounds As Variant
    Dim sourceRange As Range
    Dim insertAnchor As Range
    Dim pasteRange As Range
    Dim sourceText As String
    Dim sourceStart As Long
    Dim sourceEnd As Long
    Dim insertAt As Long
    Dim insertedLength As Long
    Dim insertPos As Long

    ' Идем снизу вверх, чтобы вставки не сдвигали позиции еще не обработанных блоков.
    For i = blockBounds.Count To 1 Step -1
        bounds = blockBounds(i)
        sourceStart = CLng(bounds(0))
        sourceEnd = CLng(bounds(1))
        insertAt = CLng(bounds(2))

        If sourceEnd <= sourceStart Then GoTo ContinueLoop
        If sourceStart < 0 Then sourceStart = 0
        If sourceEnd > doc.Content.End Then sourceEnd = doc.Content.End

        Set sourceRange = doc.Range(Start:=sourceStart, End:=sourceEnd)
        sourceText = sourceRange.Text
        If Len(Trim$(Replace$(sourceText, vbCr, ""))) = 0 Then GoTo ContinueLoop

        ' Вставляем копию блока под собой.
        Set insertAnchor = mp_GetCollapsedRangeAt(doc, insertAt)
        If insertAnchor Is Nothing Then GoTo ContinueLoop
        insertPos = insertAnchor.Start
        insertAnchor.FormattedText = sourceRange.FormattedText

        insertedLength = insertAnchor.End - insertPos
        If insertedLength <= 0 Then GoTo ContinueLoop

        Set pasteRange = doc.Range(Start:=insertPos, End:=insertPos + insertedLength)
        pasteRange.Font.Color = wdColorAutomatic

        mp_DuplicateTemplateBlocksBelowByDelimiter = mp_DuplicateTemplateBlocksBelowByDelimiter + 1

ContinueLoop:
    Next i
End Function

Private Function mp_CollectTemplateBlockBounds(ByVal doc As Document, ByVal startLine As String, ByVal endLine As String) As Collection
    Dim mainStory As Range
    Set mainStory = doc.StoryRanges(wdMainTextStory)
    If mainStory Is Nothing Then Exit Function

    Dim bounds As Collection
    Set bounds = New Collection

    Dim p As Paragraph
    Dim normalizedText As String
    Dim openDelimiterEnd As Long
    Dim blockStart As Long
    Dim blockEnd As Long

    For Each p In mainStory.Paragraphs
        normalizedText = mp_NormalizeParagraphText(p.Range.Text)

        If mp_IsTemplateDelimiterParagraph(normalizedText, startLine) Then
            openDelimiterEnd = p.Range.End
        ElseIf mp_IsTemplateDelimiterParagraph(normalizedText, endLine) Then
            If openDelimiterEnd <> 0 Then
                blockStart = openDelimiterEnd
                blockEnd = p.Range.Start
                If blockEnd > blockStart Then
                    bounds.Add Array(blockStart, blockEnd, p.Range.End)
                End If
                openDelimiterEnd = 0
            End If
        End If
    Next p

    Set mp_CollectTemplateBlockBounds = bounds
End Function

Private Function mp_TryGetTemplateBlockBoundsAtPosition(ByVal doc As Document, ByVal startLine As String, ByVal endLine As String, ByVal position As Long, ByRef blockStart As Long, ByRef blockEnd As Long, ByRef insertAt As Long, ByRef errorText As String) As Boolean
    Dim mainStory As Range
    Set mainStory = doc.StoryRanges(wdMainTextStory)
    If mainStory Is Nothing Then
        errorText = "Main document text was not found."
        Exit Function
    End If

    If position < mainStory.Start Or position > mainStory.End Then
        errorText = "Cursor must be in the main document text."
        Exit Function
    End If

    Dim p As Paragraph
    Dim normalizedText As String
    Dim openDelimiterEnd As Long
    Dim closeDelimiterStart As Long
    Dim closeDelimiterEnd As Long

    openDelimiterEnd = -1

    For Each p In mainStory.Paragraphs
        normalizedText = mp_NormalizeParagraphText(p.Range.Text)
        If mp_IsTemplateDelimiterParagraph(normalizedText, startLine) Then
            openDelimiterEnd = p.Range.End
        ElseIf mp_IsTemplateDelimiterParagraph(normalizedText, endLine) Then
            If openDelimiterEnd = -1 Then GoTo ContinueLoop

            closeDelimiterStart = p.Range.Start
            closeDelimiterEnd = p.Range.End

            If position >= openDelimiterEnd And position <= closeDelimiterStart Then
                blockStart = openDelimiterEnd
                blockEnd = closeDelimiterStart
                insertAt = closeDelimiterEnd
                mp_TryGetTemplateBlockBoundsAtPosition = True
                Exit Function
            End If

            openDelimiterEnd = -1
        End If

ContinueLoop:
    Next p

    errorText = "Cursor must be between TEMPLATE_START and TEMPLATE_END."
End Function

Private Function mp_RangeHasMeaningfulText(ByVal sourceRange As Range) As Boolean
    If sourceRange Is Nothing Then Exit Function

    Dim textValue As String
    textValue = sourceRange.Text
    textValue = Replace$(textValue, vbCr, "")
    textValue = Replace$(textValue, vbLf, "")
    textValue = Replace$(textValue, vbTab, "")
    textValue = Replace$(textValue, Chr$(7), "")
    textValue = Replace$(textValue, ChrW$(160), "")

    mp_RangeHasMeaningfulText = (Len(Trim$(textValue)) > 0)
End Function

Private Function mp_NormalizeParagraphText(ByVal textValue As String) As String
    Dim s As String
    s = Replace$(textValue, vbCr, "")
    s = Replace$(s, Chr$(7), "")
    mp_NormalizeParagraphText = Trim$(s)
End Function

Private Function mp_IsTemplateDelimiterParagraph(ByVal normalizedParagraphText As String, ByVal delimiterLine As String) As Boolean
    If Len(normalizedParagraphText) = 0 Then Exit Function
    mp_IsTemplateDelimiterParagraph = (StrComp(normalizedParagraphText, delimiterLine, vbBinaryCompare) = 0)
End Function

Private Function mp_GetCollapsedRangeAt(ByVal doc As Document, ByVal position As Long) As Range
    On Error GoTo FailPoint

    Dim safePos As Long
    safePos = position

    If safePos < 0 Then safePos = 0
    If safePos > doc.Content.End - 1 Then safePos = doc.Content.End - 1

    Set mp_GetCollapsedRangeAt = doc.Range(Start:=safePos, End:=safePos)
    Exit Function

FailPoint:
    Set mp_GetCollapsedRangeAt = Nothing
End Function

' ============================================
' Formatting API (Word)
' ============================================

Public Function m_FindRangesByFontColor(ByVal targetColor As Long) As Collection
    Dim targetDoc As Document
    Set targetDoc = mp_ResolveActiveDocument()
    If targetDoc Is Nothing Then Exit Function

    Dim hits As Collection
    Set hits = New Collection

    Dim story As Range
    Dim currentRange As Range
    Dim findRange As Range

    For Each story In targetDoc.StoryRanges
        Set currentRange = story

        Do While Not currentRange Is Nothing
            Set findRange = currentRange.Duplicate

            With findRange.Find
                .ClearFormatting
                .Text = ""
                .Format = True
                .Font.Color = targetColor
                .Forward = True
                .Wrap = wdFindStop
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With

            Do While findRange.Find.Execute
                hits.Add findRange.Duplicate
                findRange.Collapse wdCollapseEnd
            Loop

            Set currentRange = currentRange.NextStoryRange
        Loop
    Next story

    Set m_FindRangesByFontColor = hits
End Function

Public Function m_FindRangesByHighlightColor(ByVal targetHighlight As WdColorIndex) As Collection
    If targetHighlight = wdNoHighlight Then
        MsgBox "Search by wdNoHighlight is not supported.", vbExclamation, "Highlight Search"
        Exit Function
    End If

    Dim targetDoc As Document
    Set targetDoc = mp_ResolveActiveDocument()
    If targetDoc Is Nothing Then Exit Function

    Dim hits As Collection
    Set hits = New Collection

    Dim story As Range
    Dim currentRange As Range
    Dim findRange As Range

    For Each story In targetDoc.StoryRanges
        Set currentRange = story

        Do While Not currentRange Is Nothing
            Set findRange = currentRange.Duplicate

            With findRange.Find
                .ClearFormatting
                .Text = ""
                .Format = True
                .Highlight = True
                .Forward = True
                .Wrap = wdFindStop
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With

            Do While findRange.Find.Execute
                If findRange.HighlightColorIndex = targetHighlight Then
                    hits.Add findRange.Duplicate
                End If
                findRange.Collapse wdCollapseEnd
            Loop

            Set currentRange = currentRange.NextStoryRange
        Loop
    Next story

    Set m_FindRangesByHighlightColor = hits
End Function

Public Function m_CollectTextFromRanges(ByVal ranges As Collection, Optional ByVal delimiter As String = vbCrLf) As String
    If ranges Is Nothing Then Exit Function

    Dim i As Long
    Dim chunk As Range
    Dim buffer As String

    For i = 1 To ranges.Count
        Set chunk = ranges(i)
        buffer = buffer & chunk.Text
        If i < ranges.Count Then buffer = buffer & delimiter
    Next i

    m_CollectTextFromRanges = buffer
End Function

Public Function m_CopyTextByFontColor(ByVal targetColor As Long, Optional ByVal delimiter As String = vbCrLf) As Long
    Dim hits As Collection
    Set hits = m_FindRangesByFontColor(targetColor)
    If hits Is Nothing Then Exit Function

    m_LastCollectedText = m_CollectTextFromRanges(hits, delimiter)
    m_CopyTextByFontColor = hits.Count

    If m_CopyTextByFontColor > 0 Then
        mp_CopyTextToClipboard m_LastCollectedText
    End If
End Function

Public Function m_CopyTextByHighlightColor(ByVal targetHighlight As WdColorIndex, Optional ByVal delimiter As String = vbCrLf) As Long
    Dim hits As Collection
    Set hits = m_FindRangesByHighlightColor(targetHighlight)
    If hits Is Nothing Then Exit Function

    m_LastCollectedText = m_CollectTextFromRanges(hits, delimiter)
    m_CopyTextByHighlightColor = hits.Count

    If m_CopyTextByHighlightColor > 0 Then
        mp_CopyTextToClipboard m_LastCollectedText
    End If
End Function

Public Sub m_PasteClipboardAtSelection()
    If Documents.Count = 0 Then
        MsgBox "No Word document is open.", vbExclamation, "Paste"
        Exit Sub
    End If

    Selection.Paste
End Sub

Public Sub m_PasteLastCollectedTextAtSelection()
    If Documents.Count = 0 Then
        MsgBox "No Word document is open.", vbExclamation, "Paste"
        Exit Sub
    End If

    If Len(m_LastCollectedText) = 0 Then
        mp_SetStatusBarMessage "Warning: there is no saved text to paste."
        Exit Sub
    End If

    Selection.Range.Text = m_LastCollectedText
End Sub

Public Sub m_PasteTextAtSelection(ByVal textValue As String)
    If Documents.Count = 0 Then
        MsgBox "No Word document is open.", vbExclamation, "Paste"
        Exit Sub
    End If

    Selection.Range.Text = textValue
End Sub

Public Function m_GetLastCollectedText() As String
    m_GetLastCollectedText = m_LastCollectedText
End Function

Public Function m_ReplaceFontColor(ByVal fromColor As Long, ByVal toColor As Long) As Long
    Dim hits As Collection
    Set hits = m_FindRangesByFontColor(fromColor)
    If hits Is Nothing Then Exit Function

    Dim i As Long
    Dim chunk As Range

    For i = 1 To hits.Count
        Set chunk = hits(i)
        chunk.Font.Color = toColor
    Next i

    m_ReplaceFontColor = hits.Count
End Function

Public Function m_ReplaceHighlightColor(ByVal fromHighlight As WdColorIndex, ByVal toHighlight As WdColorIndex) As Long
    Dim hits As Collection
    Set hits = m_FindRangesByHighlightColor(fromHighlight)
    If hits Is Nothing Then Exit Function

    Dim i As Long
    Dim chunk As Range

    For i = 1 To hits.Count
        Set chunk = hits(i)
        chunk.HighlightColorIndex = toHighlight
    Next i

    m_ReplaceHighlightColor = hits.Count
End Function

' ============================================
' Dialog wrappers
' ============================================

Public Sub m_API_CopyTextByFontColor_Dialog()
    Dim colorValue As Long
    If Not mp_PromptFontColor("Enter font color for search and copy." & vbCrLf & _
                           "Formats: #RRGGBB, RRGGBB, Long, names (brown/red/blue/...)", _
                           "Copy by Font Color", colorValue) Then Exit Sub

    Dim count As Long
    count = m_CopyTextByFontColor(colorValue)

    If count = 0 Then
        mp_SetStatusBarMessage "Warning: no fragments with the specified font color were found."
    Else
        mp_SetStatusBarMessage "Copy by font color: found and copied fragments: " & count
    End If
End Sub

Public Sub m_API_CopyTextByHighlightColor_Dialog()
    Dim hiColor As WdColorIndex
    If Not mp_PromptHighlightColor("Enter highlight color for search and copy." & vbCrLf & _
                                "Examples: yellow, red, brightgreen, 7, 6", _
                                "Copy by Highlight", hiColor, False) Then Exit Sub

    Dim count As Long
    count = m_CopyTextByHighlightColor(hiColor)

    If count = 0 Then
        mp_SetStatusBarMessage "Warning: no fragments with the specified highlight color were found."
    Else
        mp_SetStatusBarMessage "Copy by highlight: found and copied fragments: " & count
    End If
End Sub

Public Sub m_API_ReplaceFontColor_Dialog()
    Dim fromColor As Long
    If Not mp_PromptFontColor("Enter SOURCE font color." & vbCrLf & _
                           "Formats: #RRGGBB, RRGGBB, Long, names (brown/red/blue/...)", _
                           "Replace Font Color", fromColor) Then Exit Sub

    Dim toColor As Long
    If Not mp_PromptFontColor("Enter NEW font color.", "Replace Font Color", toColor) Then Exit Sub

    Dim changed As Long
    changed = m_ReplaceFontColor(fromColor, toColor)
    mp_SetStatusBarMessage "Replace font color: fragments changed: " & changed
End Sub

Public Sub m_API_ReplaceHighlightColor_Dialog()
    Dim fromHighlight As WdColorIndex
    If Not mp_PromptHighlightColor("Enter SOURCE highlight color." & vbCrLf & _
                                "Examples: yellow, red, brightgreen, 7, 6", _
                                "Replace Highlight Color", fromHighlight, False) Then Exit Sub

    Dim toHighlight As WdColorIndex
    If Not mp_PromptHighlightColor("Enter NEW highlight color." & vbCrLf & _
                                "Use nohighlight to clear highlight.", _
                                "Replace Highlight Color", toHighlight, True) Then Exit Sub

    Dim changed As Long
    changed = m_ReplaceHighlightColor(fromHighlight, toHighlight)
    mp_SetStatusBarMessage "Replace highlight color: fragments changed: " & changed
End Sub

' ============================================
' Internal helpers
' ============================================

Private Function mp_ResolveActiveDocument() As Document
    If Documents.Count = 0 Then
        MsgBox "No Word document is open.", vbExclamation, "Formatting"
        Exit Function
    End If
    Set mp_ResolveActiveDocument = ActiveDocument
End Function

Private Sub mp_CaptureViewState()
    On Error GoTo FailCapture

    mp_ClearViewState
    If Documents.Count = 0 Then Exit Sub
    If Selection Is Nothing Then Exit Sub

    Set m_ViewDoc = ActiveDocument
    Set m_ViewWindow = ActiveWindow
    m_ViewSelectionStoryType = Selection.Range.StoryType
    m_ViewSelectionStart = Selection.Range.Start
    m_ViewSelectionEnd = Selection.Range.End

    On Error Resume Next
    If Not m_ViewWindow Is Nothing Then
        m_ViewVerticalScroll = m_ViewWindow.VerticalPercentScrolled
        m_ViewHorizontalScroll = m_ViewWindow.HorizontalPercentScrolled
    End If
    On Error GoTo 0

    m_ViewStateCaptured = True
    Exit Sub

FailCapture:
    mp_ClearViewState
    Err.Clear
End Sub

Private Sub mp_RestoreViewState()
    If Not m_ViewStateCaptured Then Exit Sub

    On Error Resume Next

    If m_ViewDoc Is Nothing Then Exit Sub
    If Documents.Count = 0 Then Exit Sub

    m_ViewDoc.Activate
    If Not m_ViewWindow Is Nothing Then m_ViewWindow.Activate

    If m_ViewSelectionStoryType = wdMainTextStory Then
        Dim mainStory As Range
        Dim maxPos As Long
        Dim startPos As Long
        Dim endPos As Long

        Set mainStory = m_ViewDoc.StoryRanges(wdMainTextStory)
        If Not mainStory Is Nothing Then
            maxPos = mainStory.End
            startPos = m_ViewSelectionStart
            endPos = m_ViewSelectionEnd

            If startPos < 0 Then startPos = 0
            If startPos > maxPos Then startPos = maxPos
            If endPos < startPos Then endPos = startPos
            If endPos > maxPos Then endPos = maxPos

            Selection.SetRange Start:=startPos, End:=endPos
        End If
    End If

    If Not m_ViewWindow Is Nothing Then
        m_ViewWindow.VerticalPercentScrolled = m_ViewVerticalScroll
        m_ViewWindow.HorizontalPercentScrolled = m_ViewHorizontalScroll
    End If
End Sub

Private Sub mp_ClearViewState()
    m_ViewStateCaptured = False
    Set m_ViewDoc = Nothing
    Set m_ViewWindow = Nothing
    m_ViewSelectionStart = 0
    m_ViewSelectionEnd = 0
    m_ViewSelectionStoryType = 0
    m_ViewVerticalScroll = 0
    m_ViewHorizontalScroll = 0
End Sub

Private Sub mp_AddUndoAnchorAtCapturedPosition()
    If Not m_ViewStateCaptured Then Exit Sub
    If m_ViewDoc Is Nothing Then Exit Sub
    If m_ViewSelectionStoryType <> wdMainTextStory Then Exit Sub

    On Error GoTo FailAnchor

    Dim mainStory As Range
    Dim anchorPos As Long
    Dim anchorRange As Range
    Const ANCHOR_NAME As String = "MP_UNDO_ANCHOR"

    Set mainStory = m_ViewDoc.StoryRanges(wdMainTextStory)
    If mainStory Is Nothing Then Exit Sub

    anchorPos = m_ViewSelectionStart
    If anchorPos < mainStory.Start Then anchorPos = mainStory.Start
    If anchorPos > mainStory.End Then anchorPos = mainStory.End

    Set anchorRange = m_ViewDoc.Range(Start:=anchorPos, End:=anchorPos)

    On Error Resume Next
    If m_ViewDoc.Bookmarks.Exists(ANCHOR_NAME) Then
        m_ViewDoc.Bookmarks(ANCHOR_NAME).Delete
    End If
    On Error GoTo FailAnchor

    m_ViewDoc.Bookmarks.Add Name:=ANCHOR_NAME, Range:=anchorRange
    If m_ViewDoc.Bookmarks.Exists(ANCHOR_NAME) Then
        m_ViewDoc.Bookmarks(ANCHOR_NAME).Delete
    End If
    Exit Sub

FailAnchor:
    Err.Clear
End Sub

Private Sub mp_BeginUndoGroup(ByVal recordName As String, ByRef started As Boolean)
    On Error GoTo UndoNotAvailable

    mp_CaptureViewState

    Dim undoRecord As Object
    Set undoRecord = CallByName(Application, "UndoRecord", VbGet)
    CallByName undoRecord, "StartCustomRecord", VbMethod, recordName
    started = True
    Exit Sub

UndoNotAvailable:
    started = False
    mp_RestoreViewState
    mp_ClearViewState
    Err.Clear
End Sub

Private Sub mp_EndUndoGroup(ByRef started As Boolean)
    On Error Resume Next

    If started Then
        mp_AddUndoAnchorAtCapturedPosition

        Dim undoRecord As Object
        Set undoRecord = CallByName(Application, "UndoRecord", VbGet)
        CallByName undoRecord, "EndCustomRecord", VbMethod
        started = False
    End If

    mp_RestoreViewState
    mp_ClearViewState
End Sub

Private Sub mp_SetStatusBarMessage(ByVal messageText As String)
    Application.StatusBar = messageText
    Application.OnTime When:=Now + TimeValue(MP_STATUSBAR_CLEAR_DELAY), Name:="m_ClearStatusBar"
End Sub

Private Function mp_CopyTextToClipboard(ByVal textValue As String) As Boolean
    If Len(textValue) = 0 Then Exit Function

    ' Forms.DataObject в части окружений "съедает" неразрывный пробел.
    ' Для таких строк сразу идем через Word-буфер.
    If InStr(1, textValue, ChrW$(160), vbBinaryCompare) > 0 Then
        GoTo FallbackWordClipboard
    End If

    ' Быстрый путь без открытия временного документа (обычно без мигания Word).
    On Error GoTo FallbackWordClipboard
    Dim dataObject As Object
    Set dataObject = CreateObject("Forms.DataObject")
    dataObject.SetText textValue
    dataObject.PutInClipboard
    mp_CopyTextToClipboard = True
    Exit Function

FallbackWordClipboard:
    Err.Clear

    Dim tempDoc As Document
    Dim tempPayload As String
    Dim sourceDoc As Document
    Dim currentScreenUpdating As Boolean
    Dim nbMarker As String
    currentScreenUpdating = Application.ScreenUpdating

    On Error GoTo CleanFail
    Application.ScreenUpdating = False

    If Documents.Count > 0 Then
        Set sourceDoc = ActiveDocument
    End If

    ' Создаем временный документ невидимо, чтобы не было мигания окна Word.
    Set tempDoc = Documents.Add(, , , False)
    nbMarker = "<<MP_NBSP_MARKER>>"
    tempPayload = Replace$(textValue, ChrW$(160), nbMarker)
    tempDoc.Range.Text = tempPayload

    ' Превращаем маркер в "настоящий" неразрывный пробел Word (^s).
    If InStr(1, tempPayload, nbMarker, vbBinaryCompare) > 0 Then
        With tempDoc.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = nbMarker
            .Replacement.Text = "^s"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchWildcards = False
            .Execute Replace:=wdReplaceAll
        End With
    End If

    tempDoc.Range.Copy
    tempDoc.Close SaveChanges:=wdDoNotSaveChanges

    If Not sourceDoc Is Nothing Then
        sourceDoc.Activate
    End If

    Application.ScreenUpdating = currentScreenUpdating
    mp_CopyTextToClipboard = True
    Exit Function

CleanFail:
    On Error Resume Next
    If Not tempDoc Is Nothing Then tempDoc.Close SaveChanges:=wdDoNotSaveChanges
    Application.ScreenUpdating = currentScreenUpdating
    MsgBox "Failed to copy text to clipboard: " & Err.Description, vbExclamation, "Copy"
End Function

Private Function mp_PromptFontColor(ByVal prompt As String, ByVal title As String, ByRef colorValue As Long) As Boolean
    Dim rawValue As String
    rawValue = Trim$(InputBox$(prompt, title))
    If Len(rawValue) = 0 Then Exit Function

    If Not mp_TryParseFontColor(rawValue, colorValue) Then
        MsgBox "Unable to parse font color. Use #RRGGBB, RRGGBB, Long, or a color name.", vbExclamation, title
        Exit Function
    End If

    mp_PromptFontColor = True
End Function

Private Function mp_PromptHighlightColor(ByVal prompt As String, ByVal title As String, ByRef colorValue As WdColorIndex, ByVal allowNoHighlight As Boolean) As Boolean
    Dim rawValue As String
    rawValue = Trim$(InputBox$(prompt, title))
    If Len(rawValue) = 0 Then Exit Function

    If Not mp_TryParseHighlightColor(rawValue, colorValue, allowNoHighlight) Then
        MsgBox "Unable to parse highlight color. Example: yellow/red/brightgreen/7/6.", vbExclamation, title
        Exit Function
    End If

    mp_PromptHighlightColor = True
End Function

Private Function mp_TryParseFontColor(ByVal rawValue As String, ByRef colorValue As Long) As Boolean
    Dim s As String
    s = LCase$(Trim$(rawValue))

    Select Case s
        Case "black": colorValue = wdColorBlack: mp_TryParseFontColor = True: Exit Function
        Case "white": colorValue = wdColorWhite: mp_TryParseFontColor = True: Exit Function
        Case "red": colorValue = wdColorRed: mp_TryParseFontColor = True: Exit Function
        Case "blue": colorValue = wdColorBlue: mp_TryParseFontColor = True: Exit Function
        Case "green": colorValue = wdColorGreen: mp_TryParseFontColor = True: Exit Function
        Case "yellow": colorValue = wdColorYellow: mp_TryParseFontColor = True: Exit Function
        Case "brown": colorValue = wdColorBrown: mp_TryParseFontColor = True: Exit Function
        Case "orange": colorValue = wdColorOrange: mp_TryParseFontColor = True: Exit Function
        Case "gray", "grey": colorValue = wdColorGray50: mp_TryParseFontColor = True: Exit Function
    End Select

    If mp_IsLikelyHexColor(s) Then
        colorValue = mp_HexToLongColor(s)
        mp_TryParseFontColor = True
        Exit Function
    End If

    If IsNumeric(s) Then
        On Error GoTo ParseFail
        colorValue = CLng(s)
        mp_TryParseFontColor = True
        Exit Function
    End If

ParseFail:
    mp_TryParseFontColor = False
End Function

Private Function mp_TryParseHighlightColor(ByVal rawValue As String, ByRef colorValue As WdColorIndex, ByVal allowNoHighlight As Boolean) As Boolean
    Dim s As String
    s = LCase$(Trim$(rawValue))

    Select Case s
        Case "black": colorValue = wdBlack
        Case "blue": colorValue = wdBlue
        Case "turquoise": colorValue = wdTurquoise
        Case "brightgreen", "lime": colorValue = wdBrightGreen
        Case "pink": colorValue = wdPink
        Case "red": colorValue = wdRed
        Case "yellow": colorValue = wdYellow
        Case "white": colorValue = wdWhite
        Case "darkblue", "navy": colorValue = wdDarkBlue
        Case "teal": colorValue = wdTeal
        Case "green": colorValue = wdGreen
        Case "violet", "purple": colorValue = wdViolet
        Case "darkred": colorValue = wdDarkRed
        Case "darkyellow", "olive": colorValue = wdDarkYellow
        Case "gray50", "grey50": colorValue = wdGray50
        Case "gray25", "grey25": colorValue = wdGray25
        Case "nohighlight", "none", "clear"
            If Not allowNoHighlight Then Exit Function
            colorValue = wdNoHighlight
        Case Else
            If Not IsNumeric(s) Then Exit Function
            On Error GoTo ParseFail
            colorValue = CInt(s)
    End Select

    If Not mp_IsValidHighlightColor(colorValue, allowNoHighlight) Then Exit Function
    mp_TryParseHighlightColor = True
    Exit Function

ParseFail:
    mp_TryParseHighlightColor = False
End Function

Private Function mp_IsValidHighlightColor(ByVal value As WdColorIndex, ByVal allowNoHighlight As Boolean) As Boolean
    Select Case value
        Case wdBlack, wdBlue, wdTurquoise, wdBrightGreen, wdPink, wdRed, wdYellow, wdWhite, _
             wdDarkBlue, wdTeal, wdGreen, wdViolet, wdDarkRed, wdDarkYellow, wdGray50, wdGray25
            mp_IsValidHighlightColor = True
        Case wdNoHighlight
            mp_IsValidHighlightColor = allowNoHighlight
    End Select
End Function

Private Function mp_IsLikelyHexColor(ByVal s As String) As Boolean
    Dim cleaned As String
    cleaned = Replace$(s, "#", "")

    If Len(cleaned) <> 6 Then Exit Function
    If cleaned Like "[0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f][0-9a-f]" Then
        mp_IsLikelyHexColor = True
    End If
End Function

Private Function mp_HexToLongColor(ByVal hexColor As String) As Long
    Dim s As String
    s = Replace$(Trim$(hexColor), "#", "")

    Dim r As Long, g As Long, b As Long
    r = CLng("&H" & Mid$(s, 1, 2))
    g = CLng("&H" & Mid$(s, 3, 2))
    b = CLng("&H" & Mid$(s, 5, 2))

    mp_HexToLongColor = RGB(r, g, b)
End Function

' ============================================
' Finalization
' ============================================

Public Sub m_Finalization()
    If Documents.Count = 0 Then
        MsgBox "No Word document is open.", vbExclamation, "Finalization"
        Exit Sub
    End If

    Dim undoStarted As Boolean
    mp_BeginUndoGroup "Order Finalization", undoStarted

    On Error GoTo FailFinalize

    m_ReplaceSequences True

    Dim removedProtectionMarkers As Long
    removedProtectionMarkers = mp_RemoveProtectedNumberMarkersOutsideTemplateRegions(ActiveDocument, MP_TEMPLATE_START_LINE, MP_TEMPLATE_END_LINE)

    Dim removedTemplateBlocks As Long
    removedTemplateBlocks = mp_RemoveTemplateBlocksAndDelimiters(ActiveDocument, MP_TEMPLATE_START_LINE, MP_TEMPLATE_END_LINE)

    Dim highlightedHeadings As Long
    highlightedHeadings = mp_HighlightSectionHeadings(ActiveDocument, wdYellow)

    mp_SetStatusBarMessage "Finalization: protection markers removed: " & removedProtectionMarkers & _
                           "; template blocks removed: " & removedTemplateBlocks & _
                           "; headings highlighted: " & highlightedHeadings

Finalize:
    mp_EndUndoGroup undoStarted
    Exit Sub

FailFinalize:
    mp_EndUndoGroup undoStarted
    MsgBox "Finalization error: " & Err.Description, vbExclamation, "Finalization"
End Sub

Private Function mp_RemoveTemplateBlocksAndDelimiters(ByVal doc As Document, ByVal startLine As String, ByVal endLine As String) As Long
    Dim deleteBounds As Collection
    Set deleteBounds = mp_CollectTemplateDeletionBounds(doc, startLine, endLine)
    If deleteBounds Is Nothing Then Exit Function
    If deleteBounds.Count = 0 Then Exit Function

    Dim i As Long
    Dim bounds As Variant
    Dim deleteRange As Range
    Dim startPos As Long
    Dim endPos As Long

    For i = deleteBounds.Count To 1 Step -1
        bounds = deleteBounds(i)
        startPos = CLng(bounds(0))
        endPos = CLng(bounds(1))

        If endPos <= startPos Then GoTo ContinueLoop

        Set deleteRange = doc.Range(Start:=startPos, End:=endPos)
        deleteRange.Delete
        mp_RemoveTemplateBlocksAndDelimiters = mp_RemoveTemplateBlocksAndDelimiters + 1

ContinueLoop:
    Next i
End Function

Private Function mp_CollectTemplateDeletionBounds(ByVal doc As Document, ByVal startLine As String, ByVal endLine As String) As Collection
    Set mp_CollectTemplateDeletionBounds = mp_CollectTemplateRegionBounds(doc, startLine, endLine)
End Function

Private Function mp_CollectTemplateRegionBounds(ByVal doc As Document, ByVal startLine As String, ByVal endLine As String) As Collection
    Dim mainStory As Range
    Set mainStory = doc.StoryRanges(wdMainTextStory)
    If mainStory Is Nothing Then Exit Function

    Dim bounds As Collection
    Set bounds = New Collection

    Dim p As Paragraph
    Dim normalizedText As String
    Dim openDelimiterStart As Long

    openDelimiterStart = -1

    For Each p In mainStory.Paragraphs
        normalizedText = mp_NormalizeParagraphText(p.Range.Text)

        If mp_IsTemplateDelimiterParagraph(normalizedText, startLine) Then
            openDelimiterStart = p.Range.Start
        ElseIf mp_IsTemplateDelimiterParagraph(normalizedText, endLine) Then
            If openDelimiterStart <> -1 Then
                bounds.Add Array(openDelimiterStart, p.Range.End)
                openDelimiterStart = -1
            End If
        End If
    Next p

    Set mp_CollectTemplateRegionBounds = bounds
End Function

Private Function mp_RemoveProtectedNumberMarkersOutsideTemplateRegions(ByVal doc As Document, ByVal startLine As String, ByVal endLine As String) As Long
    Dim mainStory As Range
    Set mainStory = doc.StoryRanges(wdMainTextStory)
    If mainStory Is Nothing Then Exit Function

    Dim protectedBounds As Collection
    Set protectedBounds = mp_CollectTemplateRegionBounds(doc, startLine, endLine)

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = False
    regex.Pattern = "^(\s*)\^\s*([0-9Nn]+(?:\.[0-9Nn]+)*)(\.?)(?=$|[\s\)\]\}\.;,:/\\\-–—])(\s*.*)$"

    Dim i As Long
    Dim pRange As Range
    Dim bodyText As String
    Dim trailingText As String
    Dim matches As Object
    Dim newBody As String

    For i = mainStory.Paragraphs.Count To 1 Step -1
        Set pRange = mainStory.Paragraphs(i).Range
        If mp_IsPositionInProtectedBounds(pRange.Start, protectedBounds) Then GoTo ContinueLoop

        bodyText = vbNullString
        trailingText = vbNullString
        mp_SplitTrailingLineBreaks pRange.Text, bodyText, trailingText
        If Len(bodyText) = 0 Then GoTo ContinueLoop

        Set matches = regex.Execute(bodyText)
        If matches.Count = 0 Then GoTo ContinueLoop

        newBody = CStr(matches(0).SubMatches(0)) & _
                  CStr(matches(0).SubMatches(1)) & _
                  CStr(matches(0).SubMatches(2)) & _
                  CStr(matches(0).SubMatches(3))

        If StrComp(newBody, bodyText, vbBinaryCompare) <> 0 Then
            pRange.Text = newBody & trailingText
            mp_RemoveProtectedNumberMarkersOutsideTemplateRegions = mp_RemoveProtectedNumberMarkersOutsideTemplateRegions + 1
        End If

ContinueLoop:
    Next i
End Function

Private Function mp_HighlightSectionHeadings(ByVal doc As Document, ByVal highlightColor As WdColorIndex) As Long
    Dim mainStory As Range
    Set mainStory = doc.StoryRanges(wdMainTextStory)
    If mainStory Is Nothing Then Exit Function

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = False
    regex.Pattern = "^\s*\d+(\.\d+)*\."

    Dim p As Paragraph
    Dim paragraphText As String
    Dim paintRange As Range
    Dim firstCharRange As Range
    Dim scanPos As Long

    For Each p In mainStory.Paragraphs
        paragraphText = mp_NormalizeParagraphText(p.Range.Text)
        If regex.Test(paragraphText) Then
            Set paintRange = p.Range.Duplicate
            mp_TrimParagraphEnding paintRange

            If paintRange.End > paintRange.Start Then
                ' Ищем первый не-пробельный символ пункта.
                scanPos = paintRange.Start
                Do While scanPos < paintRange.End
                    If Not mp_IsWhitespaceLikeChar(doc.Range(scanPos, scanPos + 1).Text) Then Exit Do
                    scanPos = scanPos + 1
                Loop

                If scanPos < paintRange.End Then
                    Set firstCharRange = doc.Range(scanPos, scanPos + 1)

                    ' Не окрашен -> не трогаем.
                    If firstCharRange.HighlightColorIndex <> wdNoHighlight Then
                        ' Уже желтый -> не трогаем.
                        If firstCharRange.HighlightColorIndex <> highlightColor Then
                            paintRange.HighlightColorIndex = highlightColor
                            mp_HighlightSectionHeadings = mp_HighlightSectionHeadings + 1
                        End If
                    End If
                End If
            End If
        End If
    Next p
End Function

Private Sub mp_TrimParagraphEnding(ByRef targetRange As Range)
    Do While targetRange.End > targetRange.Start
        Dim tailRange As Range
        Set tailRange = targetRange.Duplicate
        tailRange.SetRange Start:=targetRange.End - 1, End:=targetRange.End

        If tailRange.Text = vbCr Or AscW(tailRange.Text) = 7 Then
            targetRange.End = targetRange.End - 1
        Else
            Exit Do
        End If
    Loop
End Sub

' ============================================
' Manual numbering repair
' ============================================
Public Sub m_UpdateManualNumberingWholeDocument()
    If Documents.Count = 0 Then
        MsgBox "No Word document is open.", vbExclamation, "Numbering Update"
        Exit Sub
    End If

    Dim undoStarted As Boolean
    mp_BeginUndoGroup "Numbering Update", undoStarted
    On Error GoTo FailRenumber

    Dim changedCount As Long
    changedCount = mp_UpdateManualNumberingWholeDocumentInternal(ActiveDocument)

    If changedCount = 0 Then
        mp_SetStatusBarMessage "Numbering is already up to date, or no manual items were found."
    Else
        mp_SetStatusBarMessage "Numbering update: items fixed: " & changedCount
    End If

Finalize:
    mp_EndUndoGroup undoStarted
    Exit Sub

FailRenumber:
    mp_EndUndoGroup undoStarted
    MsgBox "Numbering update error: " & Err.Description, vbExclamation, "Numbering Update"
End Sub

Private Function mp_UpdateManualNumberingWholeDocumentInternal(ByVal doc As Document) As Long
    Dim mainStory As Range
    Set mainStory = doc.StoryRanges(wdMainTextStory)
    If mainStory Is Nothing Then Exit Function

    Dim templateBounds As Collection
    Set templateBounds = mp_CollectTemplateRegionBounds(doc, MP_TEMPLATE_START_LINE, MP_TEMPLATE_END_LINE)

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = False
    ' Supports numeric and N-based section tokens, with or without trailing dot:
    ' 1., 1.2., N., NN., N.N, NN.n, 1N.2, etc.
    ' Important: allows punctuation/comments right after the number (for example "6.4;//...").
    regex.Pattern = "^(\s*)(\^\s*)?([0-9Nn]+(?:\.[0-9Nn]+)*)(?:\.)?(?=$|[\s\)\]\}\.;,:/\\\-–—])(\s*.*)$"

    Dim candidates As Collection
    Set candidates = New Collection

    Dim p As Paragraph
    Dim bodyText As String
    Dim trailingText As String
    Dim matches As Object
    Dim m As Object
    Dim item As Object
    Dim markerText As String
    Dim numPath As String
    Dim numLevel As Long
    Dim tailText As String

    For Each p In mainStory.Paragraphs
        If mp_IsPositionInProtectedBounds(p.Range.Start, templateBounds) Then GoTo ContinueCollectLoop

        bodyText = vbNullString
        trailingText = vbNullString
        mp_SplitTrailingLineBreaks p.Range.Text, bodyText, trailingText
        If Len(bodyText) = 0 Then GoTo ContinueCollectLoop

        Set matches = regex.Execute(bodyText)
        If matches.Count = 0 Then GoTo ContinueCollectLoop

        Set m = matches(0)
        markerText = CStr(m.SubMatches(1))
        numPath = CStr(m.SubMatches(2))
        tailText = CStr(m.SubMatches(3))

        ' Не считаем даты (dd.mm.yyyy) ручной нумерацией.
        If mp_LooksLikeCalendarDate(numPath, tailText) Then GoTo ContinueCollectLoop

        numLevel = mp_CountNumberPathLevels(numPath)
        If numLevel <= 0 Then GoTo ContinueCollectLoop

        Set item = CreateObject("Scripting.Dictionary")
        item.Add "para_start", p.Range.Start
        item.Add "indent_len", Len(CStr(m.SubMatches(0)))
        item.Add "marker_len", Len(markerText)
        item.Add "protected", (Len(markerText) > 0)
        item.Add "path", numPath
        item.Add "path_len", Len(numPath)
        item.Add "level", numLevel
        item.Add "new_path", numPath
        item.Add "replace", False
        candidates.Add item

ContinueCollectLoop:
    Next p

    If candidates.Count = 0 Then Exit Function

    Dim state() As Long
    Dim prevLevel As Long
    Dim hasState As Boolean
    Dim i As Long
    Dim j As Long
    Dim currentItem As Object
    Dim currentLevel As Long
    Dim isProtected As Boolean
    Dim newPath As String

    For i = 1 To candidates.Count
        Set currentItem = candidates(i)
        currentLevel = CLng(currentItem("level"))
        isProtected = CBool(currentItem("protected"))

        If isProtected Then
            If Not mp_TrySetNumberingStateFromPath(state, CStr(currentItem("path")), currentLevel) Then GoTo ContinueStateLoop
            currentItem("new_path") = CStr(currentItem("path"))
            currentItem("replace") = False
            hasState = True
        ElseIf Not hasState Then
            mp_EnsureNumberStateCapacity state, currentLevel
            For j = 1 To currentLevel
                state(j) = 1
            Next j

            newPath = mp_BuildNumberPath(state, currentLevel)
            currentItem("new_path") = newPath
            currentItem("replace") = (StrComp(newPath, CStr(currentItem("path")), vbBinaryCompare) <> 0)
            hasState = True
        Else
            mp_AdvanceNumberingState state, prevLevel, currentLevel

            newPath = mp_BuildNumberPath(state, currentLevel)
            currentItem("new_path") = newPath
            currentItem("replace") = (StrComp(newPath, CStr(currentItem("path")), vbBinaryCompare) <> 0)
        End If

ContinueStateLoop:
        prevLevel = currentLevel
    Next i

    Dim numberRange As Range
    Dim numberStart As Long
    Dim numberEnd As Long

    For i = candidates.Count To 1 Step -1
        Set currentItem = candidates(i)
        If Not CBool(currentItem("replace")) Then GoTo ContinueApplyLoop

        numberStart = CLng(currentItem("para_start")) + CLng(currentItem("indent_len")) + CLng(currentItem("marker_len"))
        numberEnd = numberStart + CLng(currentItem("path_len"))
        Set numberRange = doc.Range(Start:=numberStart, End:=numberEnd)
        numberRange.Text = CStr(currentItem("new_path"))
        mp_UpdateManualNumberingWholeDocumentInternal = mp_UpdateManualNumberingWholeDocumentInternal + 1

ContinueApplyLoop:
    Next i
End Function

Private Function mp_TrySetNumberingStateFromPath(ByRef values() As Long, ByVal numberPath As String, ByVal level As Long) As Boolean
    If level <= 0 Then Exit Function
    If Len(numberPath) = 0 Then Exit Function

    Dim parts() As String
    parts = Split(numberPath, ".")
    If UBound(parts) + 1 <> level Then Exit Function

    mp_EnsureNumberStateCapacity values, level

    Dim i As Long
    For i = 1 To level
        If Not IsNumeric(parts(i - 1)) Then Exit Function
        values(i) = CLng(parts(i - 1))
    Next i

    Dim upperBound As Long
    upperBound = mp_GetLongArrayUpperBound(values)
    For i = level + 1 To upperBound
        values(i) = 0
    Next i

    mp_TrySetNumberingStateFromPath = True
End Function

Private Function mp_CountNumberPathLevels(ByVal numberPath As String) As Long
    If Len(numberPath) = 0 Then Exit Function
    mp_CountNumberPathLevels = UBound(Split(numberPath, ".")) + 1
End Function

Private Function mp_LooksLikeCalendarDate(ByVal numberPath As String, ByVal tailText As String) As Boolean
    Static dateRegex As Object

    If dateRegex Is Nothing Then
        Set dateRegex = CreateObject("VBScript.RegExp")
        dateRegex.Global = False
        dateRegex.IgnoreCase = False
        ' Примеры: 04.03.2026, 04.03.2026., 04.03.2026 ->
        dateRegex.Pattern = "^\s*(0?[1-9]|[12]\d|3[01])\.(0?[1-9]|1[0-2])\.(19|20)\d{2}(\D.*)?$"
    End If

    mp_LooksLikeCalendarDate = dateRegex.Test(numberPath & "." & tailText)
End Function

Private Sub mp_AdvanceNumberingState(ByRef values() As Long, ByVal previousLevel As Long, ByVal newLevel As Long)
    If newLevel <= 0 Then Exit Sub

    mp_EnsureNumberStateCapacity values, newLevel

    Dim i As Long
    Dim upperBound As Long
    upperBound = mp_GetLongArrayUpperBound(values)

    If previousLevel <= 0 Then
        If values(newLevel) <= 0 Then values(newLevel) = 1
    ElseIf newLevel = previousLevel Then
        values(newLevel) = values(newLevel) + 1
    ElseIf newLevel > previousLevel Then
        For i = previousLevel + 1 To newLevel
            values(i) = 1
        Next i
    Else
        values(newLevel) = values(newLevel) + 1
    End If

    For i = newLevel + 1 To upperBound
        values(i) = 0
    Next i
End Sub

Private Function mp_BuildNumberPath(ByRef values() As Long, ByVal level As Long) As String
    If level <= 0 Then Exit Function

    Dim i As Long
    For i = 1 To level
        If Len(mp_BuildNumberPath) = 0 Then
            mp_BuildNumberPath = CStr(values(i))
        Else
            mp_BuildNumberPath = mp_BuildNumberPath & "." & CStr(values(i))
        End If
    Next i
End Function

Private Sub mp_EnsureNumberStateCapacity(ByRef values() As Long, ByVal requiredLevel As Long)
    Dim currentUpper As Long
    currentUpper = mp_GetLongArrayUpperBound(values)

    If currentUpper = 0 Then
        ReDim values(1 To requiredLevel)
    ElseIf currentUpper < requiredLevel Then
        ReDim Preserve values(1 To requiredLevel)
    End If
End Sub

Private Function mp_GetLongArrayUpperBound(ByRef values() As Long) As Long
    On Error GoTo NotInitialized
    mp_GetLongArrayUpperBound = UBound(values)
    Exit Function

NotInitialized:
    mp_GetLongArrayUpperBound = 0
    Err.Clear
End Function

' ============================================
' FIO -> Genitive (Selection only)
' ============================================

Public Sub m_FioToGenitive_Selection()
    If Documents.Count = 0 Then
        MsgBox "No Word document is open.", vbExclamation, "FIO to Genitive"
        Exit Sub
    End If

    Dim sourceRange As Range
    Set sourceRange = Selection.Range.Duplicate
    If sourceRange Is Nothing Then Exit Sub

    Dim parseRange As Range
    Set parseRange = sourceRange.Duplicate
    mp_TrimTrailingDecorations parseRange
    If parseRange.End <= parseRange.Start Then
        mp_MarkInvalidFioSelection sourceRange
        Exit Sub
    End If

    Dim normalizedText As String
    normalizedText = mp_NormalizeFioInput(parseRange.Text)
    If Len(normalizedText) = 0 Then
        mp_MarkInvalidFioSelection sourceRange
        Exit Sub
    End If

    Dim convertedText As String
    If Not mp_TryConvertSelectionTextToGenitive(normalizedText, convertedText) Then
        mp_MarkInvalidFioSelection sourceRange
        Exit Sub
    End If

    Dim undoStarted As Boolean
    mp_BeginUndoGroup "FIO to Genitive", undoStarted
    On Error GoTo FailInflect

    parseRange.Text = convertedText

    If mp_UpdateShortFormCacheFromText(convertedText) Then
        mp_SetStatusBarMessage "FIO converted to genitive case. Short form saved to cache."
    Else
        m_LastShortFormCache = vbNullString
        mp_SetStatusBarMessage "FIO converted to genitive case. Short form was not saved."
    End If

Finalize:
    mp_EndUndoGroup undoStarted
    Exit Sub

FailInflect:
    mp_EndUndoGroup undoStarted
    MsgBox "FIO conversion error: " & Err.Description, vbExclamation, "FIO to Genitive"
End Sub

Public Sub m_FioToAccusative_Selection()
    If Documents.Count = 0 Then
        MsgBox "No Word document is open.", vbExclamation, "FIO to Accusative"
        Exit Sub
    End If

    Dim sourceRange As Range
    Set sourceRange = Selection.Range.Duplicate
    If sourceRange Is Nothing Then Exit Sub

    Dim parseRange As Range
    Set parseRange = sourceRange.Duplicate
    mp_TrimTrailingDecorations parseRange
    If parseRange.End <= parseRange.Start Then
        mp_MarkInvalidFioSelection sourceRange
        Exit Sub
    End If

    Dim normalizedText As String
    normalizedText = mp_NormalizeFioInput(parseRange.Text)
    If Len(normalizedText) = 0 Then
        mp_MarkInvalidFioSelection sourceRange
        Exit Sub
    End If

    Dim convertedText As String
    If Not mp_TryConvertSelectionTextToAccusative(normalizedText, convertedText) Then
        mp_MarkInvalidFioSelection sourceRange
        Exit Sub
    End If

    Dim undoStarted As Boolean
    mp_BeginUndoGroup "FIO to Accusative", undoStarted
    On Error GoTo FailAccusative

    parseRange.Text = convertedText

    If mp_UpdateShortFormCacheFromSourceAsGenitive(normalizedText) Then
        mp_SetStatusBarMessage "FIO converted to accusative case. Short form (genitive) saved to cache."
    Else
        m_LastShortFormCache = vbNullString
        mp_SetStatusBarMessage "FIO converted to accusative case. Short form was not saved."
    End If

Finalize:
    mp_EndUndoGroup undoStarted
    Exit Sub

FailAccusative:
    mp_EndUndoGroup undoStarted
    MsgBox "FIO conversion error: " & Err.Description, vbExclamation, "FIO to Accusative"
End Sub

Public Sub m_FioToShortForm_Selection()
    If Documents.Count = 0 Then
        MsgBox "No Word document is open.", vbExclamation, "FIO Short Form"
        Exit Sub
    End If

    If Len(m_LastShortFormCache) = 0 Then
        Dim sourceRange As Range
        Dim parseRange As Range
        Dim normalizedText As String

        Set sourceRange = Selection.Range.Duplicate
        If Not sourceRange Is Nothing Then
            Set parseRange = sourceRange.Duplicate
            mp_TrimTrailingDecorations parseRange
            If parseRange.End > parseRange.Start Then
                normalizedText = mp_NormalizeFioInput(parseRange.Text)
            End If

            If Len(normalizedText) > 0 Then
                Call mp_UpdateShortFormCacheFromText(normalizedText)
            End If
        End If
    End If

    If Len(m_LastShortFormCache) = 0 Then
        mp_SetStatusBarMessage "Warning: short-form cache is empty. Run m_FioToGenitive_Selection first or select a rank+FIO phrase."
        Exit Sub
    End If

    Dim undoStarted As Boolean
    mp_BeginUndoGroup "Report Placeholder Insert", undoStarted
    On Error GoTo FailShortForm

    Dim pivotPos As Long
    pivotPos = Selection.Range.End

    If mp_ReplaceReportPlaceholderBelow(ActiveDocument, pivotPos, m_LastShortFormCache) Then
        mp_SetStatusBarMessage "Inserted below template: report " & m_LastShortFormCache
        m_LastShortFormCache = vbNullString
    Else
        MsgBox "No template like ""report ***"" was found below cursor/selection.", vbExclamation, "FIO Short Form"
    End If

Finalize:
    mp_EndUndoGroup undoStarted
    Exit Sub

FailShortForm:
    mp_EndUndoGroup undoStarted
    MsgBox "Report placeholder insertion error: " & Err.Description, vbExclamation, "FIO Short Form"
End Sub

Private Function mp_UpdateShortFormCacheFromText(ByVal normalizedText As String) As Boolean
    Dim leadPhrase As String
    Dim surname As String
    Dim firstName As String
    Dim patronymic As String
    Dim tailPhrase As String

    If mp_TryParseSentenceWithFio(normalizedText, leadPhrase, surname, firstName, patronymic, tailPhrase) Then
        mp_UpdateShortFormCacheFromText = mp_TryComposeShortFormText(leadPhrase, surname, firstName, patronymic, m_LastShortFormCache)
        Exit Function
    End If

    If mp_TryParseFio(normalizedText, surname, firstName, patronymic) Then
        mp_UpdateShortFormCacheFromText = mp_TryComposeShortFormText(vbNullString, surname, firstName, patronymic, m_LastShortFormCache)
    End If
End Function

Private Function mp_UpdateShortFormCacheFromSourceAsGenitive(ByVal normalizedSourceText As String) As Boolean
    Dim genitiveText As String
    If Not mp_TryConvertSelectionTextToGenitive(normalizedSourceText, genitiveText) Then Exit Function

    mp_UpdateShortFormCacheFromSourceAsGenitive = mp_UpdateShortFormCacheFromText(genitiveText)
End Function

Private Function mp_TryComposeShortFormText(ByVal leadPhrase As String, ByVal surname As String, ByVal firstName As String, ByVal patronymic As String, ByRef shortText As String) As Boolean
    If Not mp_IsValidFioToken(surname) Then Exit Function
    If Not mp_IsValidFioToken(firstName) Then Exit Function
    If Not mp_IsValidFioToken(patronymic) Then Exit Function

    Dim surnameTitle As String
    surnameTitle = mp_ToTitleCaseWord(LCase$(surname))

    Dim firstInitial As String
    Dim patrInitial As String
    firstInitial = mp_InitialWithDot(firstName)
    patrInitial = mp_InitialWithDot(patronymic)
    If Len(firstInitial) = 0 Or Len(patrInitial) = 0 Then Exit Function

    leadPhrase = mp_NormalizeWhitespace(leadPhrase)

    If Len(leadPhrase) > 0 Then
        shortText = mp_LowercaseFirstLetter(leadPhrase) & " " & surnameTitle & ChrW$(160) & firstInitial & patrInitial
    Else
        shortText = surnameTitle & ChrW$(160) & firstInitial & patrInitial
    End If

    mp_TryComposeShortFormText = True
End Function

Private Function mp_ReplaceReportPlaceholderBelow(ByVal doc As Document, ByVal pivotPosition As Long, ByVal shortPhrase As String) As Boolean
    Dim mainStory As Range
    Set mainStory = doc.StoryRanges(wdMainTextStory)
    If mainStory Is Nothing Then Exit Function

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    ' Простой шаблон: "рапорт" + пробелы + 3+ звездочек.
    regex.Pattern = "(рапорт\s+)(\*{3,})"

    Dim matches As Object
    Set matches = regex.Execute(mainStory.Text)
    If matches.Count = 0 Then Exit Function

    Dim i As Long
    Dim m As Object
    Dim selectedMatch As Object
    Dim fallbackMatch As Object
    Dim absoluteStart As Long

    For i = 0 To matches.Count - 1
        Set m = matches(i)
        If fallbackMatch Is Nothing Then Set fallbackMatch = m
        absoluteStart = mainStory.Start + CLng(m.FirstIndex)

        If absoluteStart >= pivotPosition Then
            Set selectedMatch = m
            Exit For
        End If
    Next i

    If selectedMatch Is Nothing Then Set selectedMatch = fallbackMatch
    If selectedMatch Is Nothing Then Exit Function

    absoluteStart = mainStory.Start + CLng(selectedMatch.FirstIndex)

    Dim targetRange As Range
    Dim matchValue As String
    Dim prefixText As String
    Dim placeholderStart As Long
    Dim placeholderEnd As Long
    Dim placeholderRange As Range
    Dim windowStart As Long
    Dim windowEnd As Long
    Dim scanRange As Range
    Dim localPos As Long

    matchValue = CStr(selectedMatch.Value)
    Set targetRange = doc.Range(Start:=absoluteStart, End:=absoluteStart + Len(matchValue))

    ' Защита от редкого смещения индекса на 1 символ влево/вправо.
    If CStr(targetRange.Text) <> matchValue Then
        windowStart = absoluteStart - 2
        If windowStart < mainStory.Start Then windowStart = mainStory.Start

        windowEnd = absoluteStart + Len(matchValue) + 2
        If windowEnd > mainStory.End Then windowEnd = mainStory.End

        Set scanRange = doc.Range(Start:=windowStart, End:=windowEnd)
        localPos = InStr(1, CStr(scanRange.Text), matchValue, vbBinaryCompare)
        If localPos = 0 Then Exit Function

        Set targetRange = doc.Range( _
            Start:=windowStart + localPos - 1, _
            End:=windowStart + localPos - 1 + Len(matchValue))
    End If

    prefixText = CStr(selectedMatch.SubMatches(0))
    If Len(prefixText) = 0 Then Exit Function

    placeholderStart = targetRange.Start + Len(prefixText)
    placeholderEnd = targetRange.End
    If placeholderEnd <= placeholderStart Then Exit Function
    Set placeholderRange = doc.Range(Start:=placeholderStart, End:=placeholderEnd)

    placeholderRange.Text = shortPhrase
    mp_ReplaceReportPlaceholderBelow = True
End Function

Private Sub mp_MarkInvalidFioSelection(ByVal targetRange As Range)
    On Error Resume Next
    Dim markerRange As Range
    Set markerRange = targetRange.Duplicate
    markerRange.HighlightColorIndex = wdYellow
End Sub

Private Function mp_TryBuildShortFioPhrase(ByVal normalizedText As String, ByRef resultText As String) As Boolean
    Dim leadPhrase As String
    Dim surname As String
    Dim firstName As String
    Dim patronymic As String

    If Not mp_TryExtractRankAndFioFromSentence(normalizedText, leadPhrase, surname, firstName, patronymic) Then
        If mp_TryParseFio(normalizedText, surname, firstName, patronymic) Then
            leadPhrase = vbNullString
        ElseIf mp_TryBuildShortFioPhraseLegacy(normalizedText, resultText) Then
            mp_TryBuildShortFioPhrase = True
            Exit Function
        Else
            Exit Function
        End If
    End If

    If Not mp_IsValidFioToken(surname) Then Exit Function
    If Not mp_IsValidFioToken(firstName) Then Exit Function
    If Not mp_IsValidFioToken(patronymic) Then Exit Function

    Dim surnameTitle As String
    surnameTitle = mp_ToTitleCaseWord(LCase$(surname))

    Dim firstInitial As String
    Dim patrInitial As String
    firstInitial = mp_InitialWithDot(firstName)
    patrInitial = mp_InitialWithDot(patronymic)
    If Len(firstInitial) = 0 Or Len(patrInitial) = 0 Then Exit Function

    If Len(leadPhrase) > 0 Then
        resultText = mp_LowercaseFirstLetter(leadPhrase) & " " & surnameTitle & ChrW$(160) & firstInitial & patrInitial
    Else
        resultText = surnameTitle & ChrW$(160) & firstInitial & patrInitial
    End If
    mp_TryBuildShortFioPhrase = True
End Function

Private Function mp_TryBuildShortFioPhraseLegacy(ByVal normalizedText As String, ByRef resultText As String) As Boolean
    Dim parts() As String
    parts = Split(normalizedText, " ")
    If UBound(parts) < 3 Then Exit Function

    Dim surnameIndex As Long
    surnameIndex = UBound(parts) - 2
    If surnameIndex < 0 Then Exit Function

    Dim leadPhrase As String
    If surnameIndex > 0 Then
        leadPhrase = mp_JoinArraySlice(parts, 0, surnameIndex - 1)
    Else
        leadPhrase = vbNullString
    End If

    Dim surname As String
    Dim firstName As String
    Dim patronymic As String
    surname = mp_TrimTokenPunctuation(parts(surnameIndex))
    firstName = mp_TrimTokenPunctuation(parts(surnameIndex + 1))
    patronymic = mp_TrimTokenPunctuation(parts(surnameIndex + 2))

    If Not mp_IsValidFioToken(surname) Then Exit Function
    If Not mp_IsValidFioToken(firstName) Then Exit Function
    If Not mp_IsValidFioToken(patronymic) Then Exit Function

    Dim surnameTitle As String
    surnameTitle = mp_ToTitleCaseWord(LCase$(surname))

    Dim firstInitial As String
    Dim patrInitial As String
    firstInitial = mp_InitialWithDot(firstName)
    patrInitial = mp_InitialWithDot(patronymic)
    If Len(firstInitial) = 0 Or Len(patrInitial) = 0 Then Exit Function

    If Len(leadPhrase) > 0 Then
        resultText = mp_LowercaseFirstLetter(leadPhrase) & " " & surnameTitle & ChrW$(160) & firstInitial & patrInitial
    Else
        resultText = surnameTitle & ChrW$(160) & firstInitial & patrInitial
    End If
    mp_TryBuildShortFioPhraseLegacy = True
End Function

Private Function mp_InitialWithDot(ByVal token As String) As String
    Dim normalizedToken As String
    normalizedToken = mp_TrimTokenPunctuation(token)
    If Len(normalizedToken) = 0 Then Exit Function

    mp_InitialWithDot = UCase$(Left$(normalizedToken, 1)) & "."
End Function

Private Function mp_TryExtractRankAndFioFromSentence(ByVal normalizedText As String, ByRef leadPhrase As String, ByRef surname As String, ByRef firstName As String, ByRef patronymic As String) As Boolean
    Static fioRegex As Object

    If fioRegex Is Nothing Then
        Set fioRegex = CreateObject("VBScript.RegExp")
        fioRegex.Global = True
        fioRegex.IgnoreCase = True
        fioRegex.Pattern = "([А-ЯІЇЄҐа-яіїєґA-Za-z][А-ЯІЇЄҐа-яіїєґA-Za-zЬь'’`ʼ\-]+)\s+([А-ЯІЇЄҐа-яіїєґA-Za-z][А-ЯІЇЄҐа-яіїєґA-Za-zЬь'’`ʼ\-]+)\s+([А-ЯІЇЄҐа-яіїєґA-Za-z][А-ЯІЇЄҐа-яіїєґA-Za-zЬь'’`ʼ\-]+)"
    End If

    Dim matches As Object
    Set matches = fioRegex.Execute(normalizedText)
    If matches.Count = 0 Then Exit Function

    Dim i As Long
    Dim m As Object
    Dim candidateLead As String
    Dim textBefore As String
    Dim gender As String

    For i = 0 To matches.Count - 1
        Set m = matches(i)

        surname = CStr(m.SubMatches(0))
        firstName = CStr(m.SubMatches(1))
        patronymic = CStr(m.SubMatches(2))

        If Not mp_DetectFioGender(firstName, patronymic, gender) Then GoTo ContinueLoop

        textBefore = Left$(normalizedText, CLng(m.FirstIndex))
        candidateLead = mp_TrimTokenPunctuation(mp_GetPhraseTailAfterDelimiters(textBefore))
        candidateLead = mp_NormalizeWhitespace(candidateLead)
        If Len(candidateLead) = 0 Then GoTo ContinueLoop

        leadPhrase = candidateLead
        mp_TryExtractRankAndFioFromSentence = True
        Exit Function

ContinueLoop:
    Next i
End Function

Private Function mp_GetPhraseTailAfterDelimiters(ByVal sourceText As String) As String
    Dim i As Long
    Dim ch As String

    For i = Len(sourceText) To 1 Step -1
        ch = Mid$(sourceText, i, 1)
        If InStr(".,;:!?)(" & ChrW$(8212) & ChrW$(8211), ch) > 0 Then
            mp_GetPhraseTailAfterDelimiters = Mid$(sourceText, i + 1)
            Exit Function
        End If
    Next i

    mp_GetPhraseTailAfterDelimiters = sourceText
End Function

Private Function mp_NormalizeWhitespace(ByVal textValue As String) As String
    Dim s As String
    s = Trim$(textValue)
    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop
    mp_NormalizeWhitespace = s
End Function

Private Function mp_ReplaceNearestReportPlaceholder(ByVal doc As Document, ByVal pivotPosition As Long, ByVal shortPhrase As String) As Boolean
    Dim mainStory As Range
    Set mainStory = doc.StoryRanges(wdMainTextStory)
    If mainStory Is Nothing Then Exit Function

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "рапорт\s+\*{3,}"

    Dim matches As Object
    Set matches = regex.Execute(mainStory.Text)
    If matches.Count = 0 Then Exit Function

    Dim i As Long
    Dim m As Object
    Dim bestIndex As Long
    Dim bestDistance As Long
    Dim currentDistance As Long
    Dim absoluteStart As Long

    bestIndex = -1
    bestDistance = -1

    For i = 0 To matches.Count - 1
        Set m = matches(i)
        absoluteStart = mainStory.Start + CLng(m.FirstIndex)
        currentDistance = Abs(absoluteStart - pivotPosition)

        If bestIndex = -1 Or currentDistance < bestDistance Then
            bestIndex = i
            bestDistance = currentDistance
        End If
    Next i

    If bestIndex < 0 Then Exit Function

    Set m = matches(bestIndex)

    Dim targetStart As Long
    Dim targetEnd As Long
    Dim targetRange As Range
    Dim reportWord As String

    targetStart = mainStory.Start + CLng(m.FirstIndex)
    targetEnd = targetStart + Len(CStr(m.Value))
    Set targetRange = doc.Range(Start:=targetStart, End:=targetEnd)

    reportWord = mp_GetLeadingWord(targetRange.Text)
    If Len(reportWord) = 0 Then reportWord = "рапорт"

    targetRange.Text = reportWord & " " & shortPhrase
    mp_ReplaceNearestReportPlaceholder = True
End Function

Private Function mp_GetLeadingWord(ByVal textValue As String) As String
    Dim i As Long
    Dim ch As String
    textValue = Trim$(textValue)
    If Len(textValue) = 0 Then Exit Function

    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        If ch = " " Or ch = vbTab Then Exit For
    Next i

    If i <= Len(textValue) Then
        mp_GetLeadingWord = Left$(textValue, i - 1)
    Else
        mp_GetLeadingWord = textValue
    End If
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

Private Sub mp_TrimTrailingDecorations(ByRef targetRange As Range)
    Do While targetRange.End > targetRange.Start
        Dim tailRange As Range
        Set tailRange = targetRange.Duplicate
        tailRange.SetRange Start:=targetRange.End - 1, End:=targetRange.End

        If mp_IsTailDecorationChar(tailRange.Text) Then
            targetRange.End = targetRange.End - 1
        Else
            Exit Do
        End If
    Loop
End Sub

Private Function mp_IsTailDecorationChar(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function

    If ch = vbCr Or ch = vbLf Or ch = vbTab Or ch = " " Then
        mp_IsTailDecorationChar = True
        Exit Function
    End If

    Dim codePoint As Long
    codePoint = AscW(ch)
    If codePoint = 160 Or codePoint = 11 Or codePoint = 7 Then
        mp_IsTailDecorationChar = True
        Exit Function
    End If

    If InStr(".,;:!?)]}""»", ch) > 0 Then
        mp_IsTailDecorationChar = True
    End If
End Function

Private Function mp_IsWhitespaceLikeChar(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    mp_IsWhitespaceLikeChar = (ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Or AscW(ch) = 160)
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
    leadPhraseGen = mp_InflectPhraseByDashSegments(leadPhrase, 4)
    If Len(leadPhraseGen) = 0 Then Exit Function

    If Len(Trim$(tailPhrase)) > 0 Then
        Dim tailPhraseGen As String
        tailPhraseGen = mp_InflectPhraseByDashSegments(tailPhrase, 4)
        If Len(tailPhraseGen) = 0 Then Exit Function
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

    ' Для фразы-звания сохраняем прежнюю, стабильную логику склонения.
    ' Для большинства воинских званий (одушевленные, муж. род) форма совпадает.
    Dim leadPhraseAcc As String
    leadPhraseAcc = mp_InflectPhraseByDashSegments(leadPhrase, 4)
    If Len(leadPhraseAcc) = 0 Then Exit Function

    If Len(Trim$(tailPhrase)) > 0 Then
        Dim tailPhraseAcc As String
        tailPhraseAcc = mp_InflectPhraseByDashSegments(tailPhrase, 4)
        If Len(tailPhraseAcc) = 0 Then Exit Function
        convertedText = leadPhraseAcc & " " & fioAccusative & ", " & mp_LowercaseFirstLetter(tailPhraseAcc)
    Else
        convertedText = leadPhraseAcc & " " & fioAccusative
    End If

    mp_TryConvertSentenceWithFioToAccusative = True
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

Private Function mp_SplitLeadingWord(ByVal phraseText As String, ByRef firstWord As String, ByRef restText As String) As Boolean
    Static regex As Object
    If regex Is Nothing Then
        Set regex = CreateObject("VBScript.RegExp")
        regex.Global = False
        regex.IgnoreCase = True
        regex.Pattern = "^\s*([А-ЯІЇЄҐа-яіїєґA-Za-z'’`ʼ\-]+)(.*)$"
    End If

    Dim matches As Object
    Set matches = regex.Execute(phraseText)
    If matches.Count = 0 Then Exit Function

    firstWord = matches(0).SubMatches(0)
    restText = matches(0).SubMatches(1)
    mp_SplitLeadingWord = True
End Function

Private Function mp_InflectCommonWordToGenitive(ByVal sourceWord As String, ByRef resultWord As String) As Boolean
    If Not mp_IsValidGeneralWord(sourceWord) Then Exit Function

    If InStr(sourceWord, "-") > 0 Then
        If mp_TryInflectHyphenCommonWordToGenitive(sourceWord, resultWord) Then
            mp_InflectCommonWordToGenitive = True
        End If
        Exit Function
    End If

    mp_InflectCommonWordToGenitive = mp_TryInflectSimpleCommonWordToGenitive(sourceWord, resultWord)
End Function

Private Function mp_TryInflectSimpleCommonWordToGenitive(ByVal sourceWord As String, ByRef resultWord As String) As Boolean
    If Not mp_IsValidGeneralWord(sourceWord) Then Exit Function

    Dim low As String
    low = LCase$(sourceWord)

    Dim exceptions As Object
    Set exceptions = mp_GetCommonWordExceptionsDict()
    If exceptions.Exists(low) Then
        resultWord = mp_ApplyWordCase(sourceWord, exceptions(low))
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

Private Function mp_TryInflectHyphenCommonWordToGenitive(ByVal sourceWord As String, ByRef resultWord As String) As Boolean
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
                If mp_TryInflectSimpleCommonWordToGenitive(currentPart, partInflected) Then
                    ' partInflected already set.
                End If
            End If
        End If

        If i = LBound(parts) Then
            resultWord = partInflected
        Else
            resultWord = resultWord & "-" & partInflected
        End If
    Next i

    mp_TryInflectHyphenCommonWordToGenitive = True
End Function

Private Function mp_IsFixedHyphenFirstPart(ByVal partText As String) As Boolean
    Dim low As String
    low = LCase$(partText)

    Select Case low
        Case "штаб", "обер", "унтер", "віце", "екс", "лейб", "контр", "псевдо"
            mp_IsFixedHyphenFirstPart = True
    End Select
End Function

Private Function mp_IsLikelyNeuterNounOnYa(ByVal low As String) As Boolean
    If mp_EndsWith(low, "ення") Or mp_EndsWith(low, "єння") Or mp_EndsWith(low, "іння") Or _
       mp_EndsWith(low, "ання") Or mp_EndsWith(low, "ття") Or mp_EndsWith(low, "лля") Then
        mp_IsLikelyNeuterNounOnYa = True
    End If
End Function

Private Function mp_InflectLowWordEndingEtsToGenitive(ByVal lowWord As String) As String
    ' Типовые чередования для -ець:
    ' стрілець -> стрільця, боєць -> бійця, молодець -> молодця.
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

Private Function mp_IsValidGeneralWord(ByVal token As String) As Boolean
    Static regex As Object

    If regex Is Nothing Then
        Set regex = CreateObject("VBScript.RegExp")
        regex.Global = False
        regex.IgnoreCase = True
        regex.Pattern = "^[А-ЯІЇЄҐа-яіїєґA-Za-z][А-ЯІЇЄҐа-яіїєґA-Za-zЬь'’`ʼ\-]*$"
    End If

    mp_IsValidGeneralWord = regex.Test(token)
End Function

Private Function mp_LowercaseFirstLetter(ByVal sourceText As String) As String
    If Len(sourceText) = 0 Then Exit Function
    mp_LowercaseFirstLetter = LCase$(Left$(sourceText, 1)) & Mid$(sourceText, 2)
End Function

Private Function mp_InflectPhraseByDashSegments(ByVal phraseText As String, ByVal maxWordsPerSegment As Long) As String
    Dim normalized As String
    normalized = Trim$(phraseText)
    If Len(normalized) = 0 Then Exit Function

    normalized = mp_NormalizeSpacedDashSeparators(normalized)

    Dim segments() As String
    segments = Split(normalized, " - ")

    Dim i As Long
    For i = LBound(segments) To UBound(segments)
        segments(i) = mp_InflectLeadingWordsInSegment(Trim$(segments(i)), maxWordsPerSegment)
    Next i

    ' Word Ctrl+NumpadMinus -> en dash (U+2013).
    mp_InflectPhraseByDashSegments = Join(segments, " – ")
End Function

Private Function mp_NormalizeSpacedDashSeparators(ByVal textValue As String) As String
    Static regex As Object

    If regex Is Nothing Then
        Set regex = CreateObject("VBScript.RegExp")
        regex.Global = True
        regex.IgnoreCase = True
        ' Нормализуем только тире/дефисы, окруженные пробелами.
        ' Дефисы внутри слов (оператор-електрик) остаются нетронутыми.
        regex.Pattern = "\s+[—–-]\s+"
    End If

    mp_NormalizeSpacedDashSeparators = regex.Replace(textValue, " - ")
End Function

Private Function mp_InflectLeadingWordsInSegment(ByVal segmentText As String, ByVal maxWordsToInflect As Long) As String
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
        If i > 10 Then Exit For ' Защита: склоняем только начало сегмента.

        core = mp_TrimTokenPunctuation(parts(i))
        If Len(core) = 0 Then GoTo ContinueLoop
        If Not mp_IsValidGeneralWord(core) Then GoTo ContinueLoop
        If mp_IsNumericToken(core) Then GoTo ContinueLoop
        If mp_IsStopWord(core) Then GoTo ContinueLoop

        If i > LBound(parts) Then
            lowPrevCore = LCase$(mp_TrimTokenPunctuation(parts(i - 1)))
        Else
            lowPrevCore = vbNullString
        End If

        candidateCount = candidateCount + 1
        If candidateCount > maxWordsToInflect Then Exit For

        lowCore = LCase$(core)
        If mp_ShouldKeepWordUnchangedByContext(lowPrevCore, lowCore) Then GoTo ContinueLoop

        If changedCount > 0 And mp_IsLikelyAlreadyGenitive(lowCore) Then Exit For

        suffix = mp_TokenTailPunctuation(parts(i))
        If mp_InflectCommonWordToGenitive(core, inflected) Then
            parts(i) = inflected & suffix
            If LCase$(inflected) <> LCase$(core) Then
                changedCount = changedCount + 1
            End If
        End If

ContinueLoop:
    Next i

    mp_InflectLeadingWordsInSegment = Join(parts, " ")
End Function

Private Function mp_ShouldKeepWordUnchangedByContext(ByVal lowPrevCore As String, ByVal lowCore As String) As Boolean
    If Len(lowCore) = 0 Then Exit Function

    ' Если слово уже похоже на форму родительного, не трогаем.
    If mp_IsLikelyAlreadyGenitive(lowCore) Then
        mp_ShouldKeepWordUnchangedByContext = True
        Exit Function
    End If

    ' После существительных на -ення/-іння/... часто идет зависимое слово
    ' в родительном множественного с нулевым окончанием: "забезпечення стрільб".
    If mp_IsGenitiveGovernedHeadWord(lowPrevCore) Then
        If mp_HasConsonantClusterEnding(lowCore) Then
            mp_ShouldKeepWordUnchangedByContext = True
        End If
    End If
End Function

Private Function mp_IsGenitiveGovernedHeadWord(ByVal lowWord As String) As Boolean
    If Len(lowWord) = 0 Then Exit Function

    If mp_EndsWith(lowWord, "ення") Or mp_EndsWith(lowWord, "єння") Or mp_EndsWith(lowWord, "іння") Or _
       mp_EndsWith(lowWord, "ання") Or mp_EndsWith(lowWord, "ття") Or mp_EndsWith(lowWord, "лля") Then
        mp_IsGenitiveGovernedHeadWord = True
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

Private Function mp_IsLikelyAdjectiveGenitive(ByVal low As String) As Boolean
    If mp_EndsWith(low, "ого") Or mp_EndsWith(low, "ього") Or mp_EndsWith(low, "ої") Or _
       mp_EndsWith(low, "ьої") Or mp_EndsWith(low, "єї") Or mp_EndsWith(low, "их") Or _
       mp_EndsWith(low, "іх") Then
        mp_IsLikelyAdjectiveGenitive = True
    End If
End Function

Private Function mp_IsStopWord(ByVal token As String) As Boolean
    Dim low As String
    low = LCase$(token)

    Select Case low
        Case "і", "й", "та", "або", "в", "у", "на", "до", "з", "із", "зі", "по", "за", "від", "при", "для", "про"
            mp_IsStopWord = True
    End Select
End Function

Private Function mp_IsNumericToken(ByVal token As String) As Boolean
    mp_IsNumericToken = IsNumeric(token)
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
    s = Replace$(s, ChrW$(160), " ")
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
    Static regex As Object

    If regex Is Nothing Then
        Set regex = CreateObject("VBScript.RegExp")
        regex.Global = False
        regex.IgnoreCase = True
        regex.Pattern = "^[А-ЯІЇЄҐа-яіїєґA-Za-z][А-ЯІЇЄҐа-яіїєґA-Za-zЬь'’`ʼ\-]*$"
    End If

    mp_IsValidFioToken = regex.Test(token)
End Function

Private Function mp_DetectFioGender(ByVal firstName As String, ByVal patronymic As String, ByRef gender As String) As Boolean
    Dim p As String
    p = LCase$(patronymic)

    ' Підтримка як називного, так і родового відмінків по батькові.
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

Private Function mp_InflectSurnamePart(ByVal originalPart As String, ByVal gender As String, ByRef partResult As String) As Boolean
    Dim low As String
    low = LCase$(originalPart)

    Dim exceptions As Object
    Set exceptions = mp_GetSurnameExceptionsDict()
    If exceptions.Exists(low) Then
        partResult = mp_ApplyWordCase(originalPart, exceptions(low))
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
        ElseIf mp_EndsWith(low, "ець") Then
            outLow = mp_InflectLowWordEndingEtsToGenitive(low)
        ElseIf mp_EndsWith(low, "ий") Then
            outLow = Left$(low, Len(low) - 2) & "ого"
        ElseIf mp_EndsWith(low, "ій") Then
            outLow = Left$(low, Len(low) - 2) & "ія"
        ElseIf mp_EndsWith(low, "ко") Then
            outLow = Left$(low, Len(low) - 1) & "а"
        ElseIf mp_EndsWith(low, "а") Then
            outLow = Left$(low, Len(low) - 1) & "и"
        ElseIf mp_EndsWith(low, "я") Then
            outLow = Left$(low, Len(low) - 1) & "і"
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

    Dim exceptions As Object
    Set exceptions = mp_GetNameExceptionsDict()
    If exceptions.Exists(low) Then
        partResult = mp_ApplyWordCase(originalPart, exceptions(low))
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
        partResult = mp_ApplyWordCase(originalPart, exceptions(low))
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

    If mp_IsIndeclinableSurname(low, gender) Then
        partResult = originalPart
        mp_InflectSurnamePartToAccusative = True
        Exit Function
    End If

    Dim outLow As String
    outLow = low

    If gender = "male" Then
        ' Для одушевленных мужских ФІО знахідний часто совпадает с родовим,
        ' но для форм на -а/-я используем -у/-ю.
        If mp_EndsWith(low, "ський") Then
            outLow = Left$(low, Len(low) - 5) & "ського"
        ElseIf mp_EndsWith(low, "цький") Then
            outLow = Left$(low, Len(low) - 5) & "цького"
        ElseIf mp_EndsWith(low, "зький") Then
            outLow = Left$(low, Len(low) - 5) & "зького"
        ElseIf mp_EndsWith(low, "ець") Then
            outLow = mp_InflectLowWordEndingEtsToGenitive(low)
        ElseIf mp_EndsWith(low, "ий") Then
            outLow = Left$(low, Len(low) - 2) & "ого"
        ElseIf mp_EndsWith(low, "ій") Then
            outLow = Left$(low, Len(low) - 2) & "ія"
        ElseIf mp_EndsWith(low, "ко") Then
            outLow = Left$(low, Len(low) - 1) & "а"
        ElseIf mp_EndsWith(low, "а") Then
            outLow = Left$(low, Len(low) - 1) & "у"
        ElseIf mp_EndsWith(low, "я") Then
            outLow = Left$(low, Len(low) - 1) & "ю"
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

    Dim outLow As String
    outLow = low

    If gender = "male" Then
        ' Для одушевленных мужских имен знахідний часто совпадает с родовим,
        ' но для форм на -а/-я используем -у/-ю.
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

Private Function mp_IsIndeclinableSurname(ByVal lowSurname As String, ByVal gender As String) As Boolean
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

Private Function mp_EndsWithConsonant(ByVal textValue As String) As Boolean
    If Len(textValue) = 0 Then Exit Function
    Dim ch As String
    ch = Right$(textValue, 1)
    mp_EndsWithConsonant = (InStr("бвгґджзйклмнпрстфхцчшщ", ch) > 0)
End Function

Private Function mp_GetNameExceptionsDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1 ' TextCompare

    d("ілля") = "іллі"
    d("лев") = "лева"
    d("любов") = "любові"
    d("матвій") = "матвія"
    d("лука") = "луки"

    Set mp_GetNameExceptionsDict = d
End Function

Private Function mp_GetCommonWordExceptionsDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1 ' TextCompare

    d("капітан") = "капітана"
    d("лейтенант") = "лейтенанта"
    d("старший") = "старшого"
    d("майор") = "майора"
    d("підполковник") = "підполковника"
    d("полковник") = "полковника"
    d("офіцер") = "офіцера"
    d("начальник") = "начальника"
    d("штаб") = "штабу"

    Set mp_GetCommonWordExceptionsDict = d
End Function

Private Function mp_GetSurnameExceptionsDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1 ' TextCompare

    ' Дополняется по мере накопления кейсов.
    d("середа") = "середи"

    Set mp_GetSurnameExceptionsDict = d
End Function

Private Function mp_GetPatronymicExceptionsDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1 ' TextCompare
    Set mp_GetPatronymicExceptionsDict = d
End Function

Private Function mp_GetMaleNamesDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1 ' TextCompare

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
    d.CompareMode = 1 ' TextCompare

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

' ============================================
' Sequence cleanup
' ============================================

Public Sub m_ReplaceSequences(Optional ByVal useExternalUndoGroup As Boolean = False)
    Dim undoStarted As Boolean

    If Documents.Count = 0 Then
        MsgBox "No Word document is open.", vbExclamation, "Sequence Cleanup"
        Exit Sub
    End If

    If Not useExternalUndoGroup Then
        mp_BeginUndoGroup "Sequence Cleanup", undoStarted
    End If

    On Error GoTo FailReplace

    ' 0) Заменить "<неразрывный пробел>;;" на ";;"
    mp_ReplaceAllOutsideTemplateRegions ActiveDocument, "/^s;;", ";;", MP_TEMPLATE_START_LINE, MP_TEMPLATE_END_LINE

    ' 1) Удалить ";;"
    mp_ReplaceAllOutsideTemplateRegions ActiveDocument, ";;", "", MP_TEMPLATE_START_LINE, MP_TEMPLATE_END_LINE

    ' 2) Заменить "/<неразрывный пробел>" на "/"
    mp_ReplaceAllOutsideTemplateRegions ActiveDocument, "/^s", "/", MP_TEMPLATE_START_LINE, MP_TEMPLATE_END_LINE

    mp_SetStatusBarMessage "Replacement completed."

Finalize:
    If Not useExternalUndoGroup Then
        mp_EndUndoGroup undoStarted
    End If
    Exit Sub

FailReplace:
    If Not useExternalUndoGroup Then
        mp_EndUndoGroup undoStarted
    End If
    MsgBox "Sequence cleanup error: " & Err.Description, vbExclamation, "Sequence Cleanup"
End Sub

Public Sub m_ClearStatusBar()
    Application.StatusBar = vbNullString
End Sub

Private Sub mp_ReplaceAllInActiveDocument(ByVal findText As String, ByVal replaceText As String)
    mp_ReplaceAllInRange ActiveDocument.Content, findText, replaceText
End Sub

Private Sub mp_ReplaceAllOutsideTemplateRegions(ByVal doc As Document, ByVal findText As String, ByVal replaceText As String, ByVal startLine As String, ByVal endLine As String)
    Dim mainStory As Range
    Set mainStory = doc.StoryRanges(wdMainTextStory)
    If mainStory Is Nothing Then Exit Sub

    Dim protectedBounds As Collection
    Set protectedBounds = mp_CollectTemplateRegionBounds(doc, startLine, endLine)

    Dim matchRanges As Collection
    Set matchRanges = mp_CollectMatchRangesOutsideBounds(mainStory, findText, protectedBounds)
    If matchRanges Is Nothing Then Exit Sub
    If matchRanges.Count = 0 Then Exit Sub

    Dim anchorPos As Long
    anchorPos = mp_GetUndoAnchorPosition(doc, mainStory)

    mp_ApplyReplaceTextToRanges matchRanges, replaceText, anchorPos
End Sub

Private Function mp_CollectMatchRangesOutsideBounds(ByVal sourceRange As Range, ByVal findText As String, ByVal protectedBounds As Collection) As Collection
    If sourceRange Is Nothing Then Exit Function
    If sourceRange.End <= sourceRange.Start Then Exit Function

    Dim results As Collection
    Set results = New Collection

    Dim findRange As Range
    Dim safetyCounter As Long
    Set findRange = sourceRange.Duplicate

    With findRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = vbNullString
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = False
    End With

    Do While findRange.Find.Execute
        safetyCounter = safetyCounter + 1
        If safetyCounter > 1000000 Then
            Err.Raise vbObjectError + 514, "mp_CollectMatchRangesOutsideBounds", "Replacement search limit exceeded."
        End If

        If Not mp_IsPositionInProtectedBounds(findRange.Start, protectedBounds) Then
            results.Add findRange.Duplicate
        End If

        findRange.Collapse wdCollapseEnd
    Loop

    Set mp_CollectMatchRangesOutsideBounds = results
End Function

Private Function mp_IsPositionInProtectedBounds(ByVal position As Long, ByVal protectedBounds As Collection) As Boolean
    If protectedBounds Is Nothing Then Exit Function
    If protectedBounds.Count = 0 Then Exit Function

    Dim i As Long
    Dim bounds As Variant
    Dim startPos As Long
    Dim endPos As Long

    For i = 1 To protectedBounds.Count
        bounds = protectedBounds(i)
        startPos = CLng(bounds(0))
        endPos = CLng(bounds(1))

        If position >= startPos And position < endPos Then
            mp_IsPositionInProtectedBounds = True
            Exit Function
        End If
    Next i
End Function

Private Function mp_GetUndoAnchorPosition(ByVal doc As Document, ByVal mainStory As Range) As Long
    mp_GetUndoAnchorPosition = -1

    If Not m_ViewStateCaptured Then Exit Function
    If m_ViewDoc Is Nothing Then Exit Function
    If m_ViewSelectionStoryType <> wdMainTextStory Then Exit Function
    If Not (m_ViewDoc Is doc) Then Exit Function
    If mainStory Is Nothing Then Exit Function

    mp_GetUndoAnchorPosition = m_ViewSelectionStart
    If mp_GetUndoAnchorPosition < mainStory.Start Then mp_GetUndoAnchorPosition = mainStory.Start
    If mp_GetUndoAnchorPosition > mainStory.End Then mp_GetUndoAnchorPosition = mainStory.End
End Function

Private Sub mp_ApplyReplaceTextToRanges(ByVal matchRanges As Collection, ByVal replaceText As String, ByVal anchorPos As Long)
    If matchRanges Is Nothing Then Exit Sub
    If matchRanges.Count = 0 Then Exit Sub

    Dim i As Long
    Dim pivotIndex As Long
    Dim chunk As Range

    pivotIndex = mp_FindClosestRangeIndex(matchRanges, anchorPos)

    If pivotIndex <= 0 Then
        For i = 1 To matchRanges.Count
            Set chunk = matchRanges(i)
            If chunk.End > chunk.Start Then chunk.Text = replaceText
        Next i
        Exit Sub
    End If

    For i = pivotIndex To matchRanges.Count
        Set chunk = matchRanges(i)
        If chunk.End > chunk.Start Then chunk.Text = replaceText
    Next i

    For i = pivotIndex - 1 To 1 Step -1
        Set chunk = matchRanges(i)
        If chunk.End > chunk.Start Then chunk.Text = replaceText
    Next i
End Sub

Private Function mp_FindClosestRangeIndex(ByVal matchRanges As Collection, ByVal anchorPos As Long) As Long
    If matchRanges Is Nothing Then Exit Function
    If matchRanges.Count = 0 Then Exit Function
    If anchorPos < 0 Then Exit Function

    Dim i As Long
    Dim chunk As Range
    Dim distance As Long
    Dim bestDistance As Long

    bestDistance = -1

    For i = 1 To matchRanges.Count
        Set chunk = matchRanges(i)
        distance = Abs(chunk.Start - anchorPos)
        If bestDistance = -1 Or distance < bestDistance Then
            bestDistance = distance
            mp_FindClosestRangeIndex = i
        End If
    Next i
End Function

Private Sub mp_ReplaceAllInRange(ByVal targetRange As Range, ByVal findText As String, ByVal replaceText As String)
    If targetRange Is Nothing Then Exit Sub
    If targetRange.End <= targetRange.Start Then Exit Sub

    Dim findRange As Range
    Dim safetyCounter As Long

    Set findRange = targetRange.Duplicate

    With findRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = False
    End With

    Do While findRange.Find.Execute(Replace:=wdReplaceOne)
        safetyCounter = safetyCounter + 1
        If safetyCounter > 1000000 Then
            Err.Raise vbObjectError + 513, "mp_ReplaceAllInRange", "Replacement limit exceeded in range."
        End If
        findRange.Collapse wdCollapseEnd
    Loop
End Sub
