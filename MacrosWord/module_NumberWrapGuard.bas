Option Explicit

Private Const MP_STATUSBAR_CLEAR_DELAY As String = "00:00:03"
Private Const MP_BREAK_CHAR_CODE As Long = 11 ' Manual line break (Shift+Enter)

Public Sub m_NumberWrapGuard_InsertBreakBeforeSplitNumberWord()
    If Documents.Count = 0 Then
        MsgBox "No Word document is open.", vbExclamation, "Number Wrap Guard"
        Exit Sub
    End If

    Dim doc As Document
    Dim processingRange As Range
    Dim hasSelectionScope As Boolean

    Set doc = ActiveDocument

    If Not Selection Is Nothing Then
        If Selection.Range.End > Selection.Range.Start Then
            If Selection.Range.StoryType <> wdMainTextStory Then
                MsgBox "Selected text is not in the main document body." & vbCrLf & _
                       "Select a range in the main text or clear selection to process the whole document.", _
                       vbExclamation, "Number Wrap Guard"
                Exit Sub
            End If

            Set processingRange = Selection.Range.Duplicate
            hasSelectionScope = True
        End If
    End If

    If processingRange Is Nothing Then
        Set processingRange = doc.Content
    End If

    Dim undoStarted As Boolean
    Dim previousCursor As WdCursorType
    Dim screenUpdatingWasEnabled As Boolean
    Dim insertedCount As Long

    previousCursor = Application.System.Cursor
    screenUpdatingWasEnabled = Application.ScreenUpdating
    Application.System.Cursor = wdCursorWait
    DoEvents
    Application.ScreenUpdating = False

    mp_BeginUndoGroup "Number Wrap Guard", undoStarted
    On Error GoTo FailApply

    insertedCount = mp_InsertBreaksForSplitNumberWordPairs(processingRange)

    mp_EndUndoGroup undoStarted
    Application.System.Cursor = previousCursor
    Application.ScreenUpdating = screenUpdatingWasEnabled

    If hasSelectionScope Then
        mp_SetStatusBarMessage "Selection processed. Breaks inserted: " & insertedCount & "."
    Else
        mp_SetStatusBarMessage "Document processed. Breaks inserted: " & insertedCount & "."
    End If
    Exit Sub

FailApply:
    mp_EndUndoGroup undoStarted
    On Error Resume Next
    Application.System.Cursor = previousCursor
    Application.ScreenUpdating = screenUpdatingWasEnabled
    MsgBox "Execution error: " & Err.Description, vbExclamation, "Number Wrap Guard"
End Sub

Private Function mp_InsertBreaksForSplitNumberWordPairs(ByVal sourceRange As Range) As Long
    If sourceRange Is Nothing Then Exit Function

    Dim doc As Document
    Set doc = sourceRange.Document
    If doc Is Nothing Then Exit Function

    Dim insertStarts() As Long
    Dim insertCount As Long
    ReDim insertStarts(1 To 64)

    Dim searchRange As Range
    Dim insertionPoint As Range
    Dim i As Long
    Dim rawStart As Long
    Dim rawEnd As Long
    Dim nextSearchStart As Long
    Dim numberStart As Long
    Dim numberEnd As Long
    Dim wordStart As Long
    Dim numberFirstLine As Long
    Dim numberLastLine As Long
    Dim wordLine As Long
    Dim charBefore As String
    Dim nextWordFound As Boolean

    Set searchRange = doc.Range(Start:=sourceRange.Start, End:=sourceRange.End)

    Do
        mp_ConfigureNumberFind searchRange.Find
        If Not searchRange.Find.Execute Then Exit Do

        rawStart = searchRange.Start
        rawEnd = searchRange.End
        numberStart = rawStart
        numberEnd = rawEnd

        mp_ExpandNumberBounds doc, sourceRange.Start, sourceRange.End, numberStart, numberEnd
        If numberEnd <= numberStart Then GoTo ContinueSearchLoop

        If numberStart > doc.Content.Start Then
            charBefore = doc.Range(Start:=numberStart - 1, End:=numberStart).Text
            If charBefore = vbCr Or charBefore = mp_BreakChar() Then GoTo ContinueSearchLoop
        End If

        numberFirstLine = doc.Range(Start:=numberStart, End:=numberStart + 1).Information(wdFirstCharacterLineNumber)
        numberLastLine = doc.Range(Start:=numberEnd - 1, End:=numberEnd).Information(wdFirstCharacterLineNumber)
        If numberFirstLine = numberLastLine Then
            ' Проверяем следующее слово в контексте всего документа,
            ' иначе при выделении "впритык" к числу перенос может не сработать.
            nextWordFound = mp_TryFindNextWordStart(doc, numberEnd, doc.Content.End, wordStart)
            If Not nextWordFound Then GoTo ContinueSearchLoop

            wordLine = doc.Range(Start:=wordStart, End:=wordStart + 1).Information(wdFirstCharacterLineNumber)
            If wordLine = numberLastLine Then GoTo ContinueSearchLoop
        End If

        insertCount = insertCount + 1
        If insertCount > UBound(insertStarts) Then
            ReDim Preserve insertStarts(1 To UBound(insertStarts) * 2)
        End If
        insertStarts(insertCount) = numberStart

ContinueSearchLoop:
        nextSearchStart = numberEnd
        If nextSearchStart <= rawStart Then nextSearchStart = rawEnd
        If nextSearchStart >= sourceRange.End Then Exit Do
        Set searchRange = doc.Range(Start:=nextSearchStart, End:=sourceRange.End)
    Loop

    If insertCount = 0 Then Exit Function

    ReDim Preserve insertStarts(1 To insertCount)

    For i = insertCount To 1 Step -1
        If i < insertCount Then
            If insertStarts(i) = insertStarts(i + 1) Then GoTo ContinueInsertLoop
        End If

        Set insertionPoint = doc.Range(Start:=insertStarts(i), End:=insertStarts(i))
        If insertionPoint.Start > doc.Content.Start Then
            charBefore = doc.Range(Start:=insertionPoint.Start - 1, End:=insertionPoint.Start).Text
            If charBefore = vbCr Or charBefore = mp_BreakChar() Then GoTo ContinueInsertLoop
        End If

        insertionPoint.InsertBefore mp_BreakChar()
        mp_InsertBreaksForSplitNumberWordPairs = mp_InsertBreaksForSplitNumberWordPairs + 1

ContinueInsertLoop:
    Next i
End Function

Private Function mp_BreakChar() As String
    mp_BreakChar = ChrW$(MP_BREAK_CHAR_CODE)
End Function

Private Sub mp_ConfigureNumberFind(ByVal finder As Find)
    With finder
        .ClearFormatting
        .Text = "[0-9]@"
        .Replacement.ClearFormatting
        .Replacement.Text = vbNullString
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWildcards = True
    End With
End Sub

Private Sub mp_ExpandNumberBounds(ByVal doc As Document, ByVal scopeStart As Long, ByVal scopeEnd As Long, ByRef numberStart As Long, ByRef numberEnd As Long)
    Dim ch As String

    Do While numberStart > scopeStart
        ch = doc.Range(Start:=numberStart - 1, End:=numberStart).Text
        If Not mp_IsDigitOrHiddenJoiner(ch) Then Exit Do
        numberStart = numberStart - 1
    Loop

    Do While numberEnd < scopeEnd
        ch = doc.Range(Start:=numberEnd, End:=numberEnd + 1).Text
        If Not mp_IsDigitOrHiddenJoiner(ch) Then Exit Do
        numberEnd = numberEnd + 1
    Loop
End Sub

Private Function mp_TryFindNextWordStart(ByVal doc As Document, ByVal startPos As Long, ByVal scopeEnd As Long, ByRef wordStart As Long) As Boolean
    If startPos >= scopeEnd Then Exit Function

    Dim probeEnd As Long
    probeEnd = startPos + 64
    If probeEnd > scopeEnd Then probeEnd = scopeEnd
    If probeEnd <= startPos Then Exit Function

    Dim probeText As String
    probeText = doc.Range(Start:=startPos, End:=probeEnd).Text
    If Len(probeText) = 0 Then Exit Function

    Dim probeIndex As Long
    probeIndex = 1
    Do While probeIndex <= Len(probeText)
        If Not mp_IsJoinSeparatorChar(Mid$(probeText, probeIndex, 1)) Then Exit Do
        probeIndex = probeIndex + 1
    Loop

    If probeIndex <= 1 Then Exit Function
    If probeIndex > Len(probeText) Then Exit Function
    If Not mp_IsLetterChar(Mid$(probeText, probeIndex, 1)) Then Exit Function

    wordStart = startPos + probeIndex - 1
    If wordStart >= scopeEnd Then Exit Function
    mp_TryFindNextWordStart = True
End Function

Private Function mp_IsJoinSeparatorChar(ByVal ch As String) As Boolean
    Select Case ch
        Case " ", vbTab, ChrW$(160), "-", "–", "—"
            mp_IsJoinSeparatorChar = True
    End Select
End Function

Private Function mp_IsLetterChar(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    mp_IsLetterChar = (ch Like "[A-Za-zА-Яа-яІіЇїЄєҐґ]")
End Function

Private Function mp_IsDigitOrHiddenJoiner(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    If ch Like "#" Then
        mp_IsDigitOrHiddenJoiner = True
        Exit Function
    End If

    Select Case AscW(ch)
        Case 173, 8203, 8204, 8205
            mp_IsDigitOrHiddenJoiner = True
    End Select
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
    Application.OnTime When:=Now + TimeValue(MP_STATUSBAR_CLEAR_DELAY), Name:="m_NumberWrapGuard_ClearStatusBar"
End Sub

Public Sub m_NumberWrapGuard_ClearStatusBar()
    On Error Resume Next
    Application.StatusBar = vbNullString
End Sub
