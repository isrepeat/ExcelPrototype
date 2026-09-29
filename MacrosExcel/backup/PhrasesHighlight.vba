Option Explicit

' =========================
' COLORS
' =========================
Private Const COLOR_QUOTED_PHRASE As Long = &HADADAD      ' #ADADAD, "..."
Private Const COLOR_NAME_PATTERN As Long = &H50D092       ' #92D050, ім. ... / імені ...
Private Const COLOR_INNER_QUOTES As Long = &H50D092       ' #92D050, «...»
Private Const COLOR_NUMBER_PATTERN As Long = &H50D092     ' #92D050, №123 / N123

Public Sub fn_HighlightOrganizationPatterns()
    Dim c As Range
    
    For Each c In Selection
        If Not IsError(c.value) Then
            If Len(CStr(c.value)) > 0 Then
                private_HighlightCellPatterns c
            End If
        End If
    Next c
End Sub

Private Sub private_HighlightCellPatterns(ByVal c As Range)
    private_HighlightBetweenChars c, """", """", COLOR_QUOTED_PHRASE
    private_HighlightNameMarkerPatterns c, COLOR_NAME_PATTERN
    private_HighlightInnerDoubleQuotedPhrases c, COLOR_INNER_QUOTES
    private_HighlightNumberPatterns c, COLOR_NUMBER_PATTERN
End Sub

Private Sub private_HighlightBetweenChars(ByVal c As Range, ByVal openChar As String, ByVal closeChar As String, ByVal colorValue As Long)
    Dim txt As String
    Dim p1 As Long, p2 As Long, startAt As Long
    
    txt = CStr(c.value)
    startAt = 1
    
    Do
        p1 = InStr(startAt, txt, openChar)
        If p1 = 0 Then Exit Do
        
        p2 = InStr(p1 + 1, txt, closeChar)
        If p2 = 0 Then Exit Do
        
        c.Characters(p1, p2 - p1 + 1).Font.Color = colorValue
        startAt = p2 + 1
    Loop
End Sub

Private Sub private_HighlightInnerDoubleQuotedPhrases(ByVal c As Range, ByVal colorValue As Long)
    Dim txt As String
    Dim quotePositions() As Long
    Dim quoteCount As Long
    Dim p As Long
    Dim i As Long

    txt = CStr(c.value)
    p = 1

    Do
        p = InStr(p, txt, """")
        If p = 0 Then Exit Do

        quoteCount = quoteCount + 1
        ReDim Preserve quotePositions(1 To quoteCount)
        quotePositions(quoteCount) = p
        p = p + 1
    Loop

    If quoteCount < 4 Then Exit Sub

    ' Для конструкции с одинаковыми внешними/внутренними кавычками:
    ' "... "внутренний фрагмент" ..."
    For i = 2 To quoteCount - 2 Step 2
        If quotePositions(i + 1) > quotePositions(i) Then
            c.Characters(quotePositions(i), quotePositions(i + 1) - quotePositions(i) + 1).Font.Color = colorValue
        End If
    Next i
End Sub

Private Sub private_HighlightNameMarkerPatterns(ByVal c As Range, ByVal colorValue As Long)
    Dim txt As String
    Dim pIm As Long, pImeni As Long
    Dim p As Long, markerLen As Long
    Dim endPos As Long, searchFrom As Long
    
    txt = CStr(c.value)
    searchFrom = 1
    
    Do
        pIm = InStr(searchFrom, txt, "³ì.", vbTextCompare)
        pImeni = InStr(searchFrom, txt, "³ìåí³", vbTextCompare)
        
        If pIm = 0 And pImeni = 0 Then Exit Do
        
        If pIm > 0 And (pImeni = 0 Or pIm < pImeni) Then
            p = pIm
            markerLen = Len("³ì.")
        Else
            p = pImeni
            markerLen = Len("³ìåí³")
        End If
        
        endPos = private_FindNameMarkerEnd(txt, p, markerLen)
        
        If endPos >= p Then
            c.Characters(p, endPos - p + 1).Font.Color = colorValue
            searchFrom = endPos + 1
        Else
            searchFrom = p + markerLen
        End If
    Loop
End Sub

Private Function private_FindNameMarkerEnd(ByVal txt As String, ByVal markerPos As Long, ByVal markerLen As Long) As Long
    Dim pos As Long
    Dim firstToken As String
    Dim nextToken As String
    Dim tokenEnd As Long
    
    pos = markerPos + markerLen
    pos = private_SkipSpaces(txt, pos)
    
    firstToken = private_GetNextToken(txt, pos, tokenEnd)
    If firstToken = "" Then Exit Function
    
    ' Вариант с инициалами в одном токене: ім. Л.Т. Малої / імені М.І. Ситенка
    If private_IsInitialsToken(firstToken) Then
        pos = private_SkipSpaces(txt, tokenEnd + 1)
        nextToken = private_GetNextToken(txt, pos, tokenEnd)
        
        If nextToken <> "" Then
            private_FindNameMarkerEnd = tokenEnd
        End If
        
        Exit Function
    End If
    
    ' Вариант с раздельными инициалами: ім. В. Т. Зайцева / імені О. С. Коломійченка
    If private_IsSingleInitial(firstToken) Then
        pos = private_SkipSpaces(txt, tokenEnd + 1)
        nextToken = private_GetNextToken(txt, pos, tokenEnd)
        
        Do While nextToken <> "" And private_IsSingleInitial(nextToken)
            pos = private_SkipSpaces(txt, tokenEnd + 1)
            nextToken = private_GetNextToken(txt, pos, tokenEnd)
        Loop
        
        If nextToken <> "" Then
            private_FindNameMarkerEnd = tokenEnd
        End If
        
        Exit Function
    End If
    
    ' Без инициалов: подсветить максимум 3 слова после маркера ім. / імені
    private_FindNameMarkerEnd = private_GetEndAfterNWords(txt, markerPos + markerLen, 3)
End Function

Private Function private_GetEndAfterNWords(ByVal txt As String, ByVal startPos As Long, ByVal maxWords As Long) As Long
    Dim pos As Long
    Dim token As String
    Dim tokenEnd As Long
    Dim count As Long
    
    pos = private_SkipSpaces(txt, startPos)
    
    Do While pos <= Len(txt) And count < maxWords
        token = private_GetNextToken(txt, pos, tokenEnd)
        If token = "" Then Exit Do
        
        count = count + 1
        private_GetEndAfterNWords = tokenEnd
        
        pos = private_SkipSpaces(txt, tokenEnd + 1)
    Loop
End Function

Private Function private_GetNextToken(ByVal txt As String, ByVal startPos As Long, ByRef tokenEnd As Long) As String
    Dim i As Long
    Dim ch As String
    
    startPos = private_SkipSpaces(txt, startPos)
    
    If startPos > Len(txt) Then Exit Function
    If private_IsStopChar(Mid(txt, startPos, 1)) Then Exit Function
    
    i = startPos
    
    Do While i <= Len(txt)
        ch = Mid(txt, i, 1)
        
        If ch = " " Or private_IsStopChar(ch) Then Exit Do
        
        i = i + 1
    Loop
    
    tokenEnd = i - 1
    private_GetNextToken = Mid(txt, startPos, tokenEnd - startPos + 1)
End Function

Private Sub private_HighlightNumberPatterns(ByVal c As Range, ByVal colorValue As Long)
    Dim txt As String
    Dim i As Long
    Dim startPos As Long
    Dim endPos As Long
    Dim ch As String
    
    txt = CStr(c.value)
    i = 1
    
    Do While i <= Len(txt)
        ch = Mid(txt, i, 1)
        
        If ch = "¹" Or ch = "N" Or ch = "n" Then
            startPos = i
            i = i + 1
            
            Do While i <= Len(txt) And Mid(txt, i, 1) = " "
                i = i + 1
            Loop
            
            If i <= Len(txt) And private_IsDigit(Mid(txt, i, 1)) Then
                Do While i <= Len(txt)
                    ch = Mid(txt, i, 1)
                    
                    If Not (private_IsDigit(ch) Or ch = "/" Or ch = "-" Or ch = "." Or ch = " ") Then
                        Exit Do
                    End If
                    
                    i = i + 1
                Loop
                
                endPos = i - 1
                c.Characters(startPos, endPos - startPos + 1).Font.Color = colorValue
            End If
        Else
            i = i + 1
        End If
    Loop
End Sub

Private Function private_SkipSpaces(ByVal txt As String, ByVal startPos As Long) As Long
    Do While startPos <= Len(txt) And Mid(txt, startPos, 1) = " "
        startPos = startPos + 1
    Loop
    
    private_SkipSpaces = startPos
End Function

Private Function private_IsInitialsToken(ByVal s As String) As Boolean
    private_IsInitialsToken = _
        s Like "[À-ß²¯ª¥].[À-ß²¯ª¥]." Or _
        s Like "[À-ß²¯ª¥].[À-ß²¯ª¥].[À-ß²¯ª¥]."
End Function

Private Function private_IsSingleInitial(ByVal s As String) As Boolean
    private_IsSingleInitial = s Like "[À-ß²¯ª¥]."
End Function

Private Function private_IsDigit(ByVal ch As String) As Boolean
    private_IsDigit = ch >= "0" And ch <= "9"
End Function

Private Function private_IsStopChar(ByVal ch As String) As Boolean
    private_IsStopChar = _
        ch = "," Or _
        ch = ";" Or _
        ch = ":" Or _
        ch = """" Or _
        ch = ChrW(&HAB) Or _
        ch = ChrW(&HBB) Or _
        ch = "(" Or _
        ch = ")" Or _
        ch = vbLf Or _
        ch = vbCr
End Function