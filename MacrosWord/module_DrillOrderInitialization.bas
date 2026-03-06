Option Explicit

Private Const MP_DATE_TOKEN_PATTERN As String = "\{#(?:[+-]\d+)?\}"
Private Const MP_TEMPLATE_DELIMITER_LINE As String = "========================================================"
Private Const MP_STATUSBAR_CLEAR_DELAY As String = "00:00:03"
Private m_LastCollectedText As String

Public Sub m_DrillOrder_InitializeDate()
    If Documents.Count = 0 Then
        MsgBox "Нет открытого документа Word.", vbExclamation, "Инициализация приказа"
        Exit Sub
    End If

    Dim dateInput As String
    dateInput = Trim$(InputBox$("Введите число даты (1-31):", "Инициализация приказа"))

    If Len(dateInput) = 0 Then
        mp_SetStatusBarMessage "Внимание: число даты не введено."
        Exit Sub
    End If

    If Not IsNumeric(dateInput) Then
        MsgBox "Введите число от 1 до 31.", vbExclamation, "Инициализация приказа"
        Exit Sub
    End If

    Dim dateNumber As Long
    dateNumber = CLng(dateInput)

    If dateNumber < 1 Or dateNumber > 31 Then
        MsgBox "Число даты должно быть в диапазоне 1-31.", vbExclamation, "Инициализация приказа"
        Exit Sub
    End If

    Dim tokenMap As Object
    Dim foundTokenCount As Long
    Dim parseError As String

    If Not mp_BuildDateTokenMap(ActiveDocument, dateNumber, tokenMap, foundTokenCount, parseError) Then
        MsgBox parseError, vbExclamation, "Инициализация приказа"
        Exit Sub
    End If

    If foundTokenCount = 0 Then
        mp_SetStatusBarMessage "Внимание: токены {#}, {#+N}, {#-N} не найдены в документе."
        Exit Sub
    End If

    Dim undoStarted As Boolean
    mp_BeginUndoGroup "Инициализация приказа", undoStarted

    On Error GoTo FailApply

    Dim replacementsCount As Long
    replacementsCount = mp_ApplyTokenMapInDocument(ActiveDocument, tokenMap)

    If replacementsCount = 0 Then
        MsgBox "Токены найдены, но заменить их не удалось.", vbExclamation, "Инициализация приказа"
        GoTo Finalize
    End If

    Dim duplicatedCount As Long
    duplicatedCount = mp_DuplicateTemplateBlocksBelowByDelimiter(ActiveDocument, MP_TEMPLATE_DELIMITER_LINE)

    mp_SetStatusBarMessage "Готово. Заменено токенов: " & replacementsCount & _
                           "; Продублировано шаблонных блоков: " & duplicatedCount

Finalize:
    mp_EndUndoGroup undoStarted
    Exit Sub

FailApply:
    mp_EndUndoGroup undoStarted
    MsgBox "Ошибка выполнения: " & Err.Description, vbExclamation, "Инициализация приказа"
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
        errorText = "Токен """ & token & """ при базовом числе " & Format$(baseDate, "00") & _
                    " дает недопустимое число " & resolvedDate & ". Допустим диапазон 1..31."
        Exit Function
    End If

    replacementText = Format$(resolvedDate, "00")
    mp_TryResolveDateToken = True
    Exit Function

ParseError:
    errorText = "Не удалось распознать токен """ & token & """. Допустимы только {#}, {#+N}, {#-N}."
End Function

Private Function mp_DuplicateTemplateBlocksBelowByDelimiter(ByVal doc As Document, ByVal delimiterLine As String) As Long
    Dim blockBounds As Collection
    Set blockBounds = mp_CollectTemplateBlockBounds(doc, delimiterLine)
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

Private Function mp_CollectTemplateBlockBounds(ByVal doc As Document, ByVal delimiterLine As String) As Collection
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

        If normalizedText = delimiterLine Then
            If openDelimiterEnd = 0 Then
                openDelimiterEnd = p.Range.End
            Else
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

Private Function mp_NormalizeParagraphText(ByVal textValue As String) As String
    Dim s As String
    s = Replace$(textValue, vbCr, "")
    s = Replace$(s, Chr$(7), "")
    mp_NormalizeParagraphText = Trim$(s)
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
        MsgBox "Поиск по wdNoHighlight не поддержан.", vbExclamation, "Поиск по выделению"
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
        MsgBox "Нет открытого документа Word.", vbExclamation, "Вставка"
        Exit Sub
    End If

    Selection.Paste
End Sub

Public Sub m_PasteLastCollectedTextAtSelection()
    If Documents.Count = 0 Then
        MsgBox "Нет открытого документа Word.", vbExclamation, "Вставка"
        Exit Sub
    End If

    If Len(m_LastCollectedText) = 0 Then
        mp_SetStatusBarMessage "Внимание: нет сохраненного текста для вставки."
        Exit Sub
    End If

    Selection.Range.Text = m_LastCollectedText
End Sub

Public Sub m_PasteTextAtSelection(ByVal textValue As String)
    If Documents.Count = 0 Then
        MsgBox "Нет открытого документа Word.", vbExclamation, "Вставка"
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
    If Not mp_PromptFontColor("Введите цвет шрифта для поиска и копирования." & vbCrLf & _
                           "Форматы: #RRGGBB, RRGGBB, Long, имена (brown/red/blue/...)", _
                           "Копирование по цвету шрифта", colorValue) Then Exit Sub

    Dim count As Long
    count = m_CopyTextByFontColor(colorValue)

    If count = 0 Then
        mp_SetStatusBarMessage "Внимание: фрагменты с заданным цветом шрифта не найдены."
    Else
        mp_SetStatusBarMessage "Копирование по цвету шрифта: найдено и скопировано " & count & " фрагм."
    End If
End Sub

Public Sub m_API_CopyTextByHighlightColor_Dialog()
    Dim hiColor As WdColorIndex
    If Not mp_PromptHighlightColor("Введите цвет выделения (highlight) для поиска и копирования." & vbCrLf & _
                                "Примеры: yellow, red, brightgreen, 7, 6", _
                                "Копирование по выделению", hiColor, False) Then Exit Sub

    Dim count As Long
    count = m_CopyTextByHighlightColor(hiColor)

    If count = 0 Then
        mp_SetStatusBarMessage "Внимание: фрагменты с заданным цветом выделения не найдены."
    Else
        mp_SetStatusBarMessage "Копирование по выделению: найдено и скопировано " & count & " фрагм."
    End If
End Sub

Public Sub m_API_ReplaceFontColor_Dialog()
    Dim fromColor As Long
    If Not mp_PromptFontColor("Введите ИСХОДНЫЙ цвет шрифта." & vbCrLf & _
                           "Форматы: #RRGGBB, RRGGBB, Long, имена (brown/red/blue/...)", _
                           "Замена цвета шрифта", fromColor) Then Exit Sub

    Dim toColor As Long
    If Not mp_PromptFontColor("Введите НОВЫЙ цвет шрифта.", "Замена цвета шрифта", toColor) Then Exit Sub

    Dim changed As Long
    changed = m_ReplaceFontColor(fromColor, toColor)
    mp_SetStatusBarMessage "Замена цвета шрифта: изменено фрагментов " & changed
End Sub

Public Sub m_API_ReplaceHighlightColor_Dialog()
    Dim fromHighlight As WdColorIndex
    If Not mp_PromptHighlightColor("Введите ИСХОДНЫЙ цвет выделения." & vbCrLf & _
                                "Примеры: yellow, red, brightgreen, 7, 6", _
                                "Замена цвета выделения", fromHighlight, False) Then Exit Sub

    Dim toHighlight As WdColorIndex
    If Not mp_PromptHighlightColor("Введите НОВЫЙ цвет выделения." & vbCrLf & _
                                "Можно nohighlight для снятия выделения.", _
                                "Замена цвета выделения", toHighlight, True) Then Exit Sub

    Dim changed As Long
    changed = m_ReplaceHighlightColor(fromHighlight, toHighlight)
    mp_SetStatusBarMessage "Замена цвета выделения: изменено фрагментов " & changed
End Sub

' ============================================
' Internal helpers
' ============================================

Private Function mp_ResolveActiveDocument() As Document
    If Documents.Count = 0 Then
        MsgBox "Нет открытого документа Word.", vbExclamation, "Работа с форматированием"
        Exit Function
    End If
    Set mp_ResolveActiveDocument = ActiveDocument
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
    If Not started Then Exit Sub

    On Error Resume Next
    Dim undoRecord As Object
    Set undoRecord = CallByName(Application, "UndoRecord", VbGet)
    CallByName undoRecord, "EndCustomRecord", VbMethod
    started = False
End Sub

Private Sub mp_SetStatusBarMessage(ByVal messageText As String)
    Application.StatusBar = messageText
    Application.OnTime When:=Now + TimeValue(MP_STATUSBAR_CLEAR_DELAY), Name:="m_ClearStatusBar"
End Sub

Private Function mp_CopyTextToClipboard(ByVal textValue As String) As Boolean
    If Len(textValue) = 0 Then Exit Function

    Dim tempDoc As Document
    Dim currentScreenUpdating As Boolean
    currentScreenUpdating = Application.ScreenUpdating

    On Error GoTo CleanFail
    Application.ScreenUpdating = False

    Set tempDoc = Documents.Add
    tempDoc.Range.Text = textValue
    tempDoc.Range.Copy
    tempDoc.Close SaveChanges:=wdDoNotSaveChanges

    Application.ScreenUpdating = currentScreenUpdating
    mp_CopyTextToClipboard = True
    Exit Function

CleanFail:
    On Error Resume Next
    If Not tempDoc Is Nothing Then tempDoc.Close SaveChanges:=wdDoNotSaveChanges
    Application.ScreenUpdating = currentScreenUpdating
    MsgBox "Не удалось скопировать текст в буфер обмена: " & Err.Description, vbExclamation, "Копирование"
End Function

Private Function mp_PromptFontColor(ByVal prompt As String, ByVal title As String, ByRef colorValue As Long) As Boolean
    Dim rawValue As String
    rawValue = Trim$(InputBox$(prompt, title))
    If Len(rawValue) = 0 Then Exit Function

    If Not mp_TryParseFontColor(rawValue, colorValue) Then
        MsgBox "Не удалось распознать цвет шрифта. Используйте #RRGGBB, RRGGBB, Long или имя цвета.", vbExclamation, title
        Exit Function
    End If

    mp_PromptFontColor = True
End Function

Private Function mp_PromptHighlightColor(ByVal prompt As String, ByVal title As String, ByRef colorValue As WdColorIndex, ByVal allowNoHighlight As Boolean) As Boolean
    Dim rawValue As String
    rawValue = Trim$(InputBox$(prompt, title))
    If Len(rawValue) = 0 Then Exit Function

    If Not mp_TryParseHighlightColor(rawValue, colorValue, allowNoHighlight) Then
        MsgBox "Не удалось распознать цвет выделения. Пример: yellow/red/brightgreen/7/6.", vbExclamation, title
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
        MsgBox "Нет открытого документа Word.", vbExclamation, "Финализация"
        Exit Sub
    End If

    Dim undoStarted As Boolean
    mp_BeginUndoGroup "Финализация приказа", undoStarted

    On Error GoTo FailFinalize

    m_ReplaceSequences True

    Dim removedTemplateBlocks As Long
    removedTemplateBlocks = mp_RemoveTemplateBlocksAndDelimiters(ActiveDocument, MP_TEMPLATE_DELIMITER_LINE)

    Dim highlightedHeadings As Long
    highlightedHeadings = mp_HighlightSectionHeadings(ActiveDocument, wdYellow)

    mp_SetStatusBarMessage "Финализация: удалено шаблонных блоков " & removedTemplateBlocks & _
                           "; подсвечено пунктов: " & highlightedHeadings

Finalize:
    mp_EndUndoGroup undoStarted
    Exit Sub

FailFinalize:
    mp_EndUndoGroup undoStarted
    MsgBox "Ошибка финализации: " & Err.Description, vbExclamation, "Финализация"
End Sub

Private Function mp_RemoveTemplateBlocksAndDelimiters(ByVal doc As Document, ByVal delimiterLine As String) As Long
    Dim deleteBounds As Collection
    Set deleteBounds = mp_CollectTemplateDeletionBounds(doc, delimiterLine)
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

Private Function mp_CollectTemplateDeletionBounds(ByVal doc As Document, ByVal delimiterLine As String) As Collection
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

        If normalizedText = delimiterLine Then
            If openDelimiterStart = -1 Then
                openDelimiterStart = p.Range.Start
            Else
                bounds.Add Array(openDelimiterStart, p.Range.End)
                openDelimiterStart = -1
            End If
        End If
    Next p

    Set mp_CollectTemplateDeletionBounds = bounds
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

    For Each p In mainStory.Paragraphs
        paragraphText = mp_NormalizeParagraphText(p.Range.Text)
        If regex.Test(paragraphText) Then
            Set paintRange = p.Range.Duplicate
            mp_TrimParagraphEnding paintRange

            If paintRange.End > paintRange.Start Then
                paintRange.HighlightColorIndex = highlightColor
                mp_HighlightSectionHeadings = mp_HighlightSectionHeadings + 1
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
' FIO -> Genitive (Selection only)
' ============================================

Public Sub m_FioToGenitive_Selection()
    If Documents.Count = 0 Then
        MsgBox "Нет открытого документа Word.", vbExclamation, "ФІО в родовий відмінок"
        Exit Sub
    End If

    Dim sourceRange As Range
    Set sourceRange = Selection.Range.Duplicate
    If sourceRange Is Nothing Then Exit Sub

    Dim sourceBodyText As String
    Dim trailingBreaks As String
    mp_SplitTrailingLineBreaks sourceRange.Text, sourceBodyText, trailingBreaks

    Dim normalizedText As String
    normalizedText = mp_NormalizeFioInput(sourceBodyText)
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
    mp_BeginUndoGroup "ФІО у родовий відмінок", undoStarted
    On Error GoTo FailInflect

    sourceRange.Text = convertedText & trailingBreaks
    mp_SetStatusBarMessage "ФІО змінено на родовий відмінок."

Finalize:
    mp_EndUndoGroup undoStarted
    Exit Sub

FailInflect:
    mp_EndUndoGroup undoStarted
    MsgBox "Ошибка преобразования ФІО: " & Err.Description, vbExclamation, "ФІО в родовий відмінок"
End Sub

Private Sub mp_MarkInvalidFioSelection(ByVal targetRange As Range)
    On Error Resume Next
    Dim markerRange As Range
    Set markerRange = targetRange.Duplicate
    markerRange.HighlightColorIndex = wdYellow
End Sub

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

    Dim tailPhraseGen As String
    tailPhraseGen = mp_InflectPhraseByDashSegments(tailPhrase, 4)
    If Len(tailPhraseGen) = 0 Then Exit Function

    convertedText = leadPhraseGen & " " & fioGenitive & ", " & mp_LowercaseFirstLetter(tailPhraseGen)
    mp_TryConvertSentenceWithFioToGenitive = True
End Function

Private Function mp_TryParseSentenceWithFio(ByVal normalizedText As String, ByRef leadPhrase As String, ByRef surname As String, ByRef firstName As String, ByRef patronymic As String, ByRef tailPhrase As String) As Boolean
    Dim parts() As String
    parts = Split(normalizedText, " ")
    If UBound(parts) < 4 Then Exit Function

    Dim fioStart As Long
    If Not mp_FindFioStartIndex(parts, fioStart, surname, firstName, patronymic) Then Exit Function

    If fioStart < 1 Then Exit Function
    If fioStart + 3 > UBound(parts) Then Exit Function

    leadPhrase = mp_JoinArraySlice(parts, 0, fioStart - 1)
    tailPhrase = mp_JoinArraySlice(parts, fioStart + 3, UBound(parts))

    If Len(leadPhrase) = 0 Then Exit Function
    If Len(tailPhrase) = 0 Then Exit Function

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

    Dim low As String
    low = LCase$(sourceWord)

    Dim exceptions As Object
    Set exceptions = mp_GetCommonWordExceptionsDict()
    If exceptions.Exists(low) Then
        resultWord = mp_ApplyWordCase(sourceWord, exceptions(low))
        mp_InflectCommonWordToGenitive = True
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
    ElseIf mp_EndsWith(low, "ь") Then
        outLow = Left$(low, Len(low) - 1) & "я"
    ElseIf mp_EndsWith(low, "й") Then
        outLow = Left$(low, Len(low) - 1) & "я"
    ElseIf mp_EndsWithConsonant(low) Then
        outLow = low & "а"
    End If

    resultWord = mp_ApplyWordCase(sourceWord, outLow)
    mp_InflectCommonWordToGenitive = True
End Function

Private Function mp_IsLikelyNeuterNounOnYa(ByVal low As String) As Boolean
    If mp_EndsWith(low, "ення") Or mp_EndsWith(low, "єння") Or mp_EndsWith(low, "іння") Or _
       mp_EndsWith(low, "ання") Or mp_EndsWith(low, "ття") Or mp_EndsWith(low, "лля") Then
        mp_IsLikelyNeuterNounOnYa = True
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

    normalized = Replace$(normalized, " – ", " - ")
    normalized = Replace$(normalized, " — ", " - ")

    Dim segments() As String
    segments = Split(normalized, " - ")

    Dim i As Long
    For i = LBound(segments) To UBound(segments)
        segments(i) = mp_InflectLeadingWordsInSegment(Trim$(segments(i)), maxWordsPerSegment)
    Next i

    mp_InflectPhraseByDashSegments = Join(segments, " - ")
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

    For i = LBound(parts) To UBound(parts)
        If changedCount >= maxWordsToInflect Then Exit For
        If i > 10 Then Exit For ' Защита: склоняем только начало сегмента.

        core = mp_TrimTokenPunctuation(parts(i))
        If Len(core) = 0 Then GoTo ContinueLoop
        If Not mp_IsValidGeneralWord(core) Then GoTo ContinueLoop
        If mp_IsNumericToken(core) Then GoTo ContinueLoop
        If mp_IsStopWord(core) Then GoTo ContinueLoop

        lowCore = LCase$(core)
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

    If Not mp_IsValidFioToken(parts(0)) Then Exit Function
    If Not mp_IsValidFioToken(parts(1)) Then Exit Function
    If Not mp_IsValidFioToken(parts(2)) Then Exit Function

    surname = parts(0)
    firstName = parts(1)
    patronymic = parts(2)
    mp_TryParseFio = True
End Function

Private Function mp_NormalizeFioInput(ByVal inputText As String) As String
    Dim s As String
    s = Trim$(inputText)
    s = Replace$(s, vbCr, " ")
    s = Replace$(s, vbLf, " ")
    s = Replace$(s, vbTab, " ")
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
        regex.Pattern = "^[А-ЯІЇЄҐ][А-ЯІЇЄҐЬ'’`ʼ\-]*$"
    End If

    mp_IsValidFioToken = regex.Test(token)
End Function

Private Function mp_DetectFioGender(ByVal firstName As String, ByVal patronymic As String, ByRef gender As String) As Boolean
    Dim p As String
    p = LCase$(patronymic)

    If mp_EndsWith(p, "ович") Or mp_EndsWith(p, "евич") Or mp_EndsWith(p, "йович") Then
        gender = "male"
        mp_DetectFioGender = True
        Exit Function
    End If

    If mp_EndsWith(p, "івна") Or mp_EndsWith(p, "ївна") Or mp_EndsWith(p, "овна") Or mp_EndsWith(p, "евна") Then
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
            outLow = Left$(low, Len(low) - 3) & "ця"
        ElseIf mp_EndsWith(low, "ий") Then
            outLow = Left$(low, Len(low) - 2) & "ого"
        ElseIf mp_EndsWith(low, "ій") Then
            outLow = Left$(low, Len(low) - 2) & "ія"
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

Private Function mp_IsIndeclinableSurname(ByVal lowSurname As String, ByVal gender As String) As Boolean
    If mp_EndsWith(lowSurname, "енко") Or mp_EndsWith(lowSurname, "ко") Then
        mp_IsIndeclinableSurname = True
        Exit Function
    End If

    If gender = "female" Then
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
        MsgBox "Нет открытого документа Word.", vbExclamation, "Очистка последовательностей"
        Exit Sub
    End If

    If Not useExternalUndoGroup Then
        mp_BeginUndoGroup "Очистка последовательностей", undoStarted
    End If

    On Error GoTo FailReplace

    ' 1) Удалить ";;"
    mp_ReplaceAllInActiveDocument ";;", ""

    ' 2) Заменить "/<неразрывный пробел>" на "/"
    mp_ReplaceAllInActiveDocument "/^s", "/"

    ' 3) Убрать неразрывный пробел перед "від"
    mp_ReplaceAllInActiveDocument "^s від", " від"

    mp_SetStatusBarMessage "Замена выполнена"

Finalize:
    If Not useExternalUndoGroup Then
        mp_EndUndoGroup undoStarted
    End If
    Exit Sub

FailReplace:
    If Not useExternalUndoGroup Then
        mp_EndUndoGroup undoStarted
    End If
    MsgBox "Ошибка очистки последовательностей: " & Err.Description, vbExclamation, "Очистка последовательностей"
End Sub

Public Sub m_ClearStatusBar()
    Application.StatusBar = vbNullString
End Sub

Private Sub mp_ReplaceAllInActiveDocument(ByVal findText As String, ByVal replaceText As String)
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub
