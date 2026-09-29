Attribute VB_Name = "ex_ScriptDslParser"
Option Explicit

' Shared Script DSL parsing helpers used by Script runtimes.

Public Function m_ParseScript( _
    ByVal scriptText As String, _
    ByRef outBlocks As Collection, _
    ByRef outErrorText As String _
) As Boolean
    m_ParseScript = ex_ScriptDSL.m_ParseScriptToBlocks(scriptText, outBlocks, outErrorText)
End Function

Public Function m_NormalizeScript(ByVal scriptText As String) As String
    Dim lines As Variant
    Dim i As Long
    Dim rawLine As String
    Dim cleaned As String
    Dim normalized As String
    Dim inBacktickLiteral As Boolean
    Dim hasBacktickLiterals As Boolean

    scriptText = Replace(scriptText, vbCrLf, vbLf)
    scriptText = Replace(scriptText, vbCr, vbLf)
    hasBacktickLiterals = (InStr(1, scriptText, "`", vbBinaryCompare) > 0)
    scriptText = mp_StripMultiLineComments(scriptText, hasBacktickLiterals)
    lines = Split(scriptText, vbLf)

    For i = LBound(lines) To UBound(lines)
        rawLine = CStr(lines(i))
        rawLine = Replace(rawLine, vbTab, " ")
        rawLine = Replace(rawLine, ChrW$(160), " ")
        rawLine = mp_StripSingleLineComment(rawLine, inBacktickLiteral, hasBacktickLiterals)
        cleaned = Trim$(rawLine)
        If Len(normalized) > 0 Then normalized = normalized & vbLf
        normalized = normalized & cleaned
    Next i

    m_NormalizeScript = normalized
End Function

Public Function m_TrimTrailingSemicolon(ByVal lineText As String) As String
    lineText = Trim$(lineText)
    If Right$(lineText, 1) = ";" Then
        m_TrimTrailingSemicolon = Trim$(Left$(lineText, Len(lineText) - 1))
    Else
        m_TrimTrailingSemicolon = lineText
    End If
End Function

Public Function m_TryParseCallArgs( _
    ByVal statementText As String, _
    ByVal functionName As String, _
    ByRef outArgs As Collection _
) As Boolean
    Dim openPos As Long
    Dim closePos As Long
    Dim callee As String
    Dim argsText As String

    statementText = Trim$(statementText)
    If Len(statementText) = 0 Then Exit Function

    openPos = InStr(1, statementText, "(", vbBinaryCompare)
    If openPos <= 1 Then Exit Function
    If Right$(statementText, 1) <> ")" Then Exit Function

    callee = Trim$(Left$(statementText, openPos - 1))
    If StrComp(callee, functionName, vbTextCompare) <> 0 Then Exit Function

    closePos = Len(statementText)
    argsText = Trim$(Mid$(statementText, openPos + 1, closePos - openPos - 1))
    Set outArgs = m_ParseQuotedArgs(argsText)

    m_TryParseCallArgs = True
End Function

Public Function m_ParseQuotedArgs(ByVal argsText As String) As Collection
    Dim result As Collection
    Dim i As Long
    Dim ch As String
    Dim quoteChar As String
    Dim currentArg As String
    Dim inQuote As Boolean

    Set result = New Collection
    argsText = Trim$(argsText)
    If Len(argsText) = 0 Then
        Set m_ParseQuotedArgs = result
        Exit Function
    End If

    For i = 1 To Len(argsText)
        ch = Mid$(argsText, i, 1)
        If inQuote Then
            If ch = quoteChar Then
                inQuote = False
                result.Add currentArg
                currentArg = vbNullString
            Else
                currentArg = currentArg & ch
            End If
        Else
            If ch = """" Or ch = "`" Then
                inQuote = True
                quoteChar = ch
                currentArg = vbNullString
            ElseIf ch = "," Or ch = " " Or ch = vbTab Then
                ' separators between args
            Else
                Err.Raise vbObjectError + 6138, "ex_ScriptDslParser", _
                    "Arguments must be quoted strings. Unexpected token near '" & ch & "'."
            End If
        End If
    Next i

    If inQuote Then
        Err.Raise vbObjectError + 6139, "ex_ScriptDslParser", "Unclosed quoted argument in script call."
    End If

    Set m_ParseQuotedArgs = result
End Function

Private Function mp_StripMultiLineComments( _
    ByVal sourceText As String, _
    Optional ByVal hasBacktickLiterals As Boolean = True _
) As String
    Dim i As Long
    Dim ch As String
    Dim nextCh As String
    Dim inQuotes As Boolean
    Dim inCommentBlock As Boolean
    Dim inBacktickLiteral As Boolean
    Dim result As String

    i = 1
    Do While i <= Len(sourceText)
        If hasBacktickLiterals Then
            If mp_IsBacktickAt(sourceText, i) And Not inQuotes Then
                inBacktickLiteral = Not inBacktickLiteral
                result = result & "`"
                i = i + 1
                GoTo ContinueLoop
            End If
        End If

        ch = Mid$(sourceText, i, 1)

        If inCommentBlock Then
            If ch = "*" And i < Len(sourceText) Then
                nextCh = Mid$(sourceText, i + 1, 1)
                If nextCh = "/" Then
                    inCommentBlock = False
                    i = i + 2
                    GoTo ContinueLoop
                End If
            End If

            If ch = vbLf Then result = result & vbLf
            i = i + 1
            GoTo ContinueLoop
        End If

        If Not inBacktickLiteral Then
            If ch = """" Then
                inQuotes = Not inQuotes
                result = result & ch
                i = i + 1
                GoTo ContinueLoop
            End If

            If Not inQuotes Then
                If ch = "/" And i < Len(sourceText) Then
                    nextCh = Mid$(sourceText, i + 1, 1)
                    If nextCh = "*" Then
                        inCommentBlock = True
                        i = i + 2
                        GoTo ContinueLoop
                    End If
                End If
            End If
        End If

        result = result & ch
        i = i + 1

ContinueLoop:
    Loop

    mp_StripMultiLineComments = result
End Function

Private Function mp_StripSingleLineComment( _
    ByVal lineText As String, _
    ByRef inBacktickLiteral As Boolean, _
    Optional ByVal hasBacktickLiterals As Boolean = True _
) As String
    Dim i As Long
    Dim inQuotes As Boolean
    Dim ch As String
    Dim nextCh As String

    For i = 1 To Len(lineText)
        If hasBacktickLiterals Then
            If mp_IsBacktickAt(lineText, i) And Not inQuotes Then
                inBacktickLiteral = Not inBacktickLiteral
                GoTo ContinueLoop
            End If
        End If

        ch = Mid$(lineText, i, 1)
        If Not inBacktickLiteral Then
            If ch = """" Then
                inQuotes = Not inQuotes
            End If

            If Not inQuotes Then
                If ch = "/" And i < Len(lineText) Then
                    nextCh = Mid$(lineText, i + 1, 1)
                    If nextCh = "/" Then
                        mp_StripSingleLineComment = Left$(lineText, i - 1)
                        Exit Function
                    End If
                End If
            End If
        End If
ContinueLoop:
    Next i

    mp_StripSingleLineComment = lineText
End Function

Private Function mp_IsBacktickAt(ByVal textValue As String, ByVal pos As Long) As Boolean
    If pos < 1 Then Exit Function
    If pos > Len(textValue) Then Exit Function
    mp_IsBacktickAt = (Mid$(textValue, pos, 1) = "`")
End Function
