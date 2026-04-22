Option Explicit

Private Const MP_NEWLINE_STYLE_NAME As String = "__NewLine"
Private Const MP_STATUSBAR_CLEAR_DELAY As String = "00:00:03"

Public Sub m_NewLine_ApplyStyleAfterFirstEmptyParagraphInDocument()
    mp_ApplyStyleAfterEmptyParagraphsInDocument
End Sub

Private Sub mp_ApplyStyleAfterEmptyParagraphsInDocument()
    If Documents.Count = 0 Then
        MsgBox "No Word document is open.", vbExclamation, "NewLine Paragraph Style"
        Exit Sub
    End If

    Dim removedEmptyCount As Long
    Dim styledParagraphCount As Long
    Dim targetParagraphStart As Long
    Dim errorText As String
    Dim undoStarted As Boolean
    Dim screenUpdatingWasEnabled As Boolean
    Dim doc As Document
    Dim previousCursor As WdCursorType
    Dim processingRange As Range
    Dim hasSelectionScope As Boolean

    Set doc = ActiveDocument

    If Not Selection Is Nothing Then
        If Selection.Range.End > Selection.Range.Start Then
            If Selection.Range.StoryType <> wdMainTextStory Then
                MsgBox "Selected text is not in the main document body." & vbCrLf & _
                       "Select a range in the main text or clear selection to process the whole document.", _
                       vbExclamation, "NewLine Paragraph Style"
                Exit Sub
            End If

            Set processingRange = Selection.Range.Duplicate
            hasSelectionScope = True
        End If
    End If

    If processingRange Is Nothing Then
        Set processingRange = doc.Content
    End If

    screenUpdatingWasEnabled = Application.ScreenUpdating
    previousCursor = Application.System.Cursor
    Application.System.Cursor = wdCursorWait
    DoEvents
    Application.ScreenUpdating = False

    mp_BeginUndoGroup "NewLine Paragraph Style", undoStarted
    On Error GoTo FailApply

    If Not m_TryApplyStyleAfterEmptyParagraphs(processingRange, removedEmptyCount, styledParagraphCount, targetParagraphStart, errorText, MP_NEWLINE_STYLE_NAME) Then
        mp_EndUndoGroup undoStarted
        Application.ScreenUpdating = screenUpdatingWasEnabled
        Application.System.Cursor = previousCursor
        MsgBox errorText, vbExclamation, "NewLine Paragraph Style"
        Exit Sub
    End If

    mp_EndUndoGroup undoStarted
    Application.ScreenUpdating = screenUpdatingWasEnabled
    Application.System.Cursor = previousCursor

    If hasSelectionScope Then
        mp_SetStatusBarMessage "Selection processed. Empty lines removed: " & removedEmptyCount & "; paragraphs styled: " & styledParagraphCount & "."
    Else
        mp_SetStatusBarMessage "Document processed. Empty lines removed: " & removedEmptyCount & "; paragraphs styled: " & styledParagraphCount & "."
    End If
    Exit Sub

FailApply:
    mp_EndUndoGroup undoStarted
    On Error Resume Next
    Application.ScreenUpdating = screenUpdatingWasEnabled
    Application.System.Cursor = previousCursor
    MsgBox "Execution error: " & Err.Description, vbExclamation, "NewLine Paragraph Style"
End Sub

Public Function m_TryApplyStyleAfterFirstEmptyParagraph(ByVal sourceRange As Range, ByRef removedEmptyCount As Long, ByRef targetParagraphStart As Long, ByRef errorText As String, Optional ByVal styleName As String = "__NewLine") As Boolean
    Dim styledParagraphCount As Long
    m_TryApplyStyleAfterFirstEmptyParagraph = m_TryApplyStyleAfterEmptyParagraphs(sourceRange, removedEmptyCount, styledParagraphCount, targetParagraphStart, errorText, styleName)
End Function

Public Function m_TryApplyStyleAfterEmptyParagraphs(ByVal sourceRange As Range, ByRef removedEmptyCount As Long, ByRef styledParagraphCount As Long, ByRef targetParagraphStart As Long, ByRef errorText As String, Optional ByVal styleName As String = "__NewLine") As Boolean
    removedEmptyCount = 0
    styledParagraphCount = 0
    targetParagraphStart = 0
    errorText = vbNullString

    If sourceRange Is Nothing Then
        errorText = "Source range is not specified."
        Exit Function
    End If

    Dim doc As Document
    Set doc = sourceRange.Document
    If doc Is Nothing Then
        errorText = "Unable to resolve target document."
        Exit Function
    End If

    styleName = Trim$(styleName)
    If Len(styleName) = 0 Then
        errorText = "Style name is empty."
        Exit Function
    End If

    Dim targetStyle As Style
    If Not mp_TryResolveStyle(doc, styleName, targetStyle, errorText) Then Exit Function

    Dim paragraphList As Paragraphs
    Set paragraphList = sourceRange.Paragraphs
    If paragraphList Is Nothing Then
        errorText = "No paragraphs found in the specified range."
        Exit Function
    End If

    If paragraphList.Count < 2 Then
        errorText = "No paragraphs with leading empty line were found."
        Exit Function
    End If

    On Error GoTo ApplyFailed

    Dim i As Long
    Dim targetParagraph As Paragraph
    Dim currentParagraph As Paragraph
    Dim deleteStarts() As Long
    Dim deleteEnds() As Long
    Dim targetStarts() As Long
    Dim deleteCount As Long
    Dim previousEmpty As Boolean
    Dim currentEmpty As Boolean
    Dim previousStart As Long
    Dim previousEnd As Long

    ReDim deleteStarts(1 To paragraphList.Count)
    ReDim deleteEnds(1 To paragraphList.Count)
    ReDim targetStarts(1 To paragraphList.Count)

    ' Pass 1: только собираем позиции целевых абзацев и пустых строк.
    Set currentParagraph = paragraphList(1)
    previousEmpty = mp_IsParagraphEmptyLine(currentParagraph)
    previousStart = currentParagraph.Range.Start
    previousEnd = currentParagraph.Range.End

    For i = 2 To paragraphList.Count
        Set currentParagraph = paragraphList(i)
        currentEmpty = mp_IsParagraphEmptyLine(currentParagraph)

        If (Not currentEmpty) And previousEmpty Then
            deleteCount = deleteCount + 1
            deleteStarts(deleteCount) = previousStart
            deleteEnds(deleteCount) = previousEnd
            targetStarts(deleteCount) = currentParagraph.Range.Start
        End If

        previousEmpty = currentEmpty
        previousStart = currentParagraph.Range.Start
        previousEnd = currentParagraph.Range.End
    Next i

    If deleteCount = 0 Then
        errorText = "No paragraphs with leading empty line were found."
        Exit Function
    End If

    ReDim Preserve deleteStarts(1 To deleteCount)
    ReDim Preserve deleteEnds(1 To deleteCount)
    ReDim Preserve targetStarts(1 To deleteCount)

    Dim deleteRange As Range

    ' Pass 2: применяем стиль и удаляем отмеченные строки в обратном порядке.
    For i = deleteCount To 1 Step -1
        Set targetParagraph = doc.Range(Start:=targetStarts(i), End:=targetStarts(i)).Paragraphs(1)
        If targetParagraphStart = 0 Then targetParagraphStart = targetParagraph.Range.Start

        If CStr(targetParagraph.Range.Style) <> styleName Then
            targetParagraph.Range.Style = targetStyle
        End If

        Set deleteRange = doc.Range(Start:=deleteStarts(i), End:=deleteEnds(i))
        deleteRange.Delete
        removedEmptyCount = removedEmptyCount + 1
        styledParagraphCount = styledParagraphCount + 1
    Next i

    m_TryApplyStyleAfterEmptyParagraphs = True
    Exit Function

ApplyFailed:
    errorText = "Unable to apply style and delete empty lines: " & Err.Description
End Function

Private Function mp_IsParagraphEmptyLine(ByVal paragraphValue As Paragraph) As Boolean
    If paragraphValue Is Nothing Then Exit Function

    Dim paragraphText As String
    paragraphText = paragraphValue.Range.Text

    Dim i As Long
    Dim ch As String
    For i = 1 To Len(paragraphText)
        ch = Mid$(paragraphText, i, 1)
        Select Case ch
            Case vbCr
                mp_IsParagraphEmptyLine = True
                Exit Function
            Case " ", vbTab, ChrW$(160), ChrW$(11), ChrW$(7)
                ' Skip whitespace-like characters.
            Case Else
                Exit Function
        End Select
    Next i

    mp_IsParagraphEmptyLine = True
End Function

Private Function mp_TryResolveStyle(ByVal doc As Document, ByVal styleName As String, ByRef resolvedStyle As Style, ByRef errorText As String) As Boolean
    On Error GoTo StyleNotFound
    Set resolvedStyle = doc.Styles(styleName)
    mp_TryResolveStyle = True
    Exit Function

StyleNotFound:
    errorText = "Style '" & styleName & "' is not found in the current document."
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
    Application.OnTime When:=Now + TimeValue(MP_STATUSBAR_CLEAR_DELAY), Name:="m_NewLineParagraphStyle_ClearStatusBar"
End Sub

Public Sub m_NewLineParagraphStyle_ClearStatusBar()
    On Error Resume Next
    Application.StatusBar = vbNullString
End Sub
