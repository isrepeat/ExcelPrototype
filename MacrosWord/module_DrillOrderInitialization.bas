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
