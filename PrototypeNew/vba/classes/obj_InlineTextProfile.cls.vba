VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_InlineTextProfile"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean

Private Const DELIMITER_DOUBLE As String = "double"
Private Const DELIMITER_SINGLE As String = "single"

Private m_PartName As String
Private m_InlineMarkersEnabled As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
    m_InlineMarkersEnabled = False
End Sub
Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Dispose
    On Error GoTo 0
End Sub

' //
' // API
' //
Public Function Initialize() As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    On Error GoTo 0
End Sub

Public Property Get PartName() As String
    PartName = m_PartName
End Property

Public Property Let PartName(ByVal value As String)
    m_PartName = VBA.CStr(value)
End Property

Public Property Get InlineMarkersEnabled() As Boolean
    InlineMarkersEnabled = m_InlineMarkersEnabled
End Property

Public Property Let InlineMarkersEnabled(ByVal value As Boolean)
    m_InlineMarkersEnabled = VBA.CBool(value)
End Property

Public Function TryResolveInlineText( _
    ByVal rawText As String, _
    ByRef outText As String, _
    ByRef outRuns As Collection _
) As Boolean
    ' Profile contains shared inline rules for a specific partName:
    ' 1) parse [[tag]]...[[/tag]] and [[color=...]]...[[/color]]
    ' 2) return runs for deferred apply step.
    If Not m_InlineMarkersEnabled Then
        outText = rawText
        Set outRuns = Nothing
        TryResolveInlineText = True
        Exit Function
    End If

    If Not private_TryParseInlineMarkers(rawText, outText, outRuns) Then Exit Function
    TryResolveInlineText = True
End Function

Public Function TryResolveInlineRunStyle( _
    ByVal markerName As String, _
    ByRef outFontColor As Long, _
    ByRef outFontBold As Boolean, _
    ByRef outFontItalic As Boolean, _
    ByRef outFontUnderline As Boolean _
) As Boolean
    If Not private_TryResolveInlineMarkerStyle( _
        markerName, outFontColor, outFontBold, outFontItalic, outFontUnderline) Then Exit Function

    TryResolveInlineRunStyle = True
End Function

Public Sub ApplyInlineRuns(ByVal targetRange As Range, ByVal runs As Collection)
    Dim runInfo As Object
    Dim startIndex As Long
    Dim runLength As Long
    Dim fontColor As Long
    Dim fontBold As Boolean
    Dim fontItalic As Boolean
    Dim fontUnderline As Boolean
    Dim applyFontFlags As Boolean

    If Not m_InlineMarkersEnabled Then Exit Sub
    If targetRange Is Nothing Then Exit Sub
    If runs Is Nothing Then Exit Sub

    For Each runInfo In runs
        startIndex = VBA.CLng(runInfo("Start"))
        runLength = VBA.CLng(runInfo("Length"))

        If startIndex <= 0 Or runLength <= 0 Then GoTo ContinueRun
        If Not private_TryResolveInlineRunStyleFromInfo( _
            runInfo, fontColor, fontBold, fontItalic, fontUnderline, applyFontFlags) Then GoTo ContinueRun

        On Error Resume Next
        With targetRange.Characters(startIndex, runLength).Font
            .Color = fontColor
            If applyFontFlags Then
                .Bold = fontBold
                .Italic = fontItalic
                If fontUnderline Then
                    .Underline = xlUnderlineStyleSingle
                Else
                    .Underline = xlUnderlineStyleNone
                End If
            End If
        End With
        On Error GoTo 0

ContinueRun:
    Next runInfo
End Sub

Public Sub ApplyInlineRunsToShape(ByVal targetShape As Shape, ByVal runs As Collection)
    Dim runInfo As Object
    Dim startIndex As Long
    Dim runLength As Long
    Dim fontColor As Long
    Dim fontBold As Boolean
    Dim fontItalic As Boolean
    Dim fontUnderline As Boolean
    Dim underlineValue As XlUnderlineStyle
    Dim underlineStyle2 As MsoTextUnderlineType
    Dim applyFontFlags As Boolean

    If Not m_InlineMarkersEnabled Then Exit Sub
    If targetShape Is Nothing Then Exit Sub
    If runs Is Nothing Then Exit Sub

    For Each runInfo In runs
        startIndex = VBA.CLng(runInfo("Start"))
        runLength = VBA.CLng(runInfo("Length"))

        If startIndex <= 0 Or runLength <= 0 Then GoTo ContinueRun
        If Not private_TryResolveInlineRunStyleFromInfo( _
            runInfo, fontColor, fontBold, fontItalic, fontUnderline, applyFontFlags) Then GoTo ContinueRun

        If applyFontFlags Then
            If fontUnderline Then
                underlineValue = xlUnderlineStyleSingle
                underlineStyle2 = msoUnderlineSingleLine
            Else
                underlineValue = xlUnderlineStyleNone
                underlineStyle2 = msoNoUnderline
            End If
        End If

        On Error Resume Next
        With targetShape.TextFrame.Characters(startIndex, runLength).Font
            .Color = fontColor
            If applyFontFlags Then
                .Bold = fontBold
                .Italic = fontItalic
                .Underline = underlineValue
            End If
        End With
        With targetShape.TextFrame2.TextRange.Characters(startIndex, runLength).Font
            .Fill.ForeColor.RGB = fontColor
            If applyFontFlags Then
                .Bold = fontBold
                .Italic = fontItalic
                .UnderlineStyle = underlineStyle2
            End If
        End With
        On Error GoTo 0

ContinueRun:
    Next runInfo
End Sub

Private Function private_TryParseInlineMarkers( _
    ByVal rawText As String, _
    ByRef outText As String, _
    ByRef outRuns As Collection _
) As Boolean
    Dim markerStack As Collection
    Dim markerInfo As Object
    Dim textLen As Long
    Dim i As Long
    Dim closePos As Long
    Dim tokenText As String
    Dim rawTextLower As String
    Dim tagName As String
    Dim isClosing As Boolean
    Dim colorSpec As String
    Dim escapedRightBracketCount As Long

    Set markerStack = New Collection
    Set outRuns = New Collection

    outText = VBA.vbNullString
    rawTextLower = VBA.LCase$(rawText)
    textLen = VBA.Len(rawText)
    i = 1

    Do While i <= textLen
        ' Escape helper: "[[[...]]]" renders outer brackets literally.
        ' Example: "[[[red]AAA[/red]]]" -> "[AAA]"
        If i + 2 <= textLen Then
            If VBA.Mid$(rawText, i, 3) = "[[[" Then
                outText = outText & "["
                escapedRightBracketCount = escapedRightBracketCount + 1
                ' Keep the remaining "[[" so inline marker can still be parsed as double-bracket token.
                i = i + 1
                GoTo ContinueLoop
            End If
        End If

        If i + 1 <= textLen Then
            If VBA.Mid$(rawText, i, 2) = "[[" Then
                closePos = VBA.InStr(i + 2, rawText, "]]", VBA.vbBinaryCompare)
                If closePos > 0 Then
                    tokenText = VBA.Trim$(VBA.Mid$(rawText, i + 2, closePos - i - 2))
                    If private_TryParseInlineToken(tokenText, isClosing, tagName, colorSpec) Then
                        If private_TryHandleInlineToken( _
                            isClosing, tagName, colorSpec, DELIMITER_DOUBLE, closePos + 2, rawTextLower, markerStack, outText, outRuns) Then
                            i = closePos + 2
                            GoTo ContinueLoop
                        End If
                    End If
                End If
            End If
        End If

        If VBA.Mid$(rawText, i, 1) = "[" Then
            closePos = VBA.InStr(i + 1, rawText, "]", VBA.vbBinaryCompare)
            If closePos > 0 Then
                tokenText = VBA.Trim$(VBA.Mid$(rawText, i + 1, closePos - i - 1))
                If private_TryParseInlineToken(tokenText, isClosing, tagName, colorSpec) Then
                    If private_TryHandleInlineToken( _
                        isClosing, tagName, colorSpec, DELIMITER_SINGLE, closePos + 1, rawTextLower, markerStack, outText, outRuns) Then
                        i = closePos + 1
                        GoTo ContinueLoop
                    End If
                End If
            End If
        End If

        If escapedRightBracketCount > 0 Then
            If i + 1 <= textLen Then
                If VBA.Mid$(rawText, i, 2) = "]]" Then
                    outText = outText & "]"
                    escapedRightBracketCount = escapedRightBracketCount - 1
                    i = i + 2
                    GoTo ContinueLoop
                End If
            End If
        End If

        outText = outText & VBA.Mid$(rawText, i, 1)
        i = i + 1

ContinueLoop:
    Loop

    ' Unclosed markers are treated as literals in source text and therefore
    ' have no run records.
    For Each markerInfo In markerStack
        ' no-op; explicit loop for readability and future diagnostics
    Next markerInfo

    private_TryParseInlineMarkers = True
End Function

Private Function private_TryParseInlineToken( _
    ByVal tokenText As String, _
    ByRef outIsClosing As Boolean, _
    ByRef outTagName As String, _
    ByRef outColorSpec As String _
) As Boolean
    Dim tokenBody As String
    Dim eqPos As Long
    Dim leftPart As String
    Dim rightPart As String

    tokenBody = VBA.Trim$(tokenText)
    If VBA.Len(tokenBody) = 0 Then Exit Function

    outIsClosing = False
    outTagName = VBA.vbNullString
    outColorSpec = VBA.vbNullString

    If VBA.Left$(tokenBody, 1) = "/" Then
        outIsClosing = True
        tokenBody = VBA.Trim$(VBA.Mid$(tokenBody, 2))
        tokenBody = VBA.LCase$(tokenBody)

        If tokenBody = "color" Then
            outTagName = "color"
            private_TryParseInlineToken = True
            Exit Function
        End If

        If private_IsSupportedInlineMarkerName(tokenBody) Then
            outTagName = tokenBody
            private_TryParseInlineToken = True
        End If
        Exit Function
    End If

    eqPos = VBA.InStr(1, tokenBody, "=", VBA.vbBinaryCompare)
    If eqPos > 0 Then
        leftPart = VBA.LCase$(VBA.Trim$(VBA.Left$(tokenBody, eqPos - 1)))
        rightPart = private_Unquote(VBA.Trim$(VBA.Mid$(tokenBody, eqPos + 1)))

        If leftPart = "color" And VBA.Len(rightPart) > 0 Then
            outTagName = "color"
            outColorSpec = rightPart
            private_TryParseInlineToken = True
            Exit Function
        End If
    End If

    tokenBody = VBA.LCase$(tokenBody)
    If private_IsSupportedInlineMarkerName(tokenBody) Then
        outTagName = tokenBody
        private_TryParseInlineToken = True
    End If
End Function

Private Function private_TryHandleInlineToken( _
    ByVal isClosing As Boolean, _
    ByVal tagName As String, _
    ByVal colorSpec As String, _
    ByVal delimiterType As String, _
    ByVal searchFrom As Long, _
    ByVal rawTextLower As String, _
    ByVal markerStack As Collection, _
    ByRef outText As String, _
    ByVal outRuns As Collection _
) As Boolean
    Dim markerInfo As Object
    Dim runStart As Long
    Dim runLength As Long
    Dim startTag As String
    Dim startDelimiter As String
    Dim startColorSpec As String

    If markerStack Is Nothing Then Exit Function
    If outRuns Is Nothing Then Exit Function

    If isClosing Then
        If markerStack.Count = 0 Then Exit Function

        Set markerInfo = markerStack(markerStack.Count)
        startTag = VBA.LCase$(VBA.CStr(markerInfo("Tag")))
        startDelimiter = VBA.LCase$(VBA.CStr(markerInfo("Delimiter")))
        If VBA.StrComp(startTag, tagName, VBA.vbBinaryCompare) <> 0 Then Exit Function
        If VBA.StrComp(startDelimiter, delimiterType, VBA.vbBinaryCompare) <> 0 Then Exit Function

        runStart = VBA.CLng(markerInfo("Start"))
        runLength = VBA.Len(outText) - runStart + 1

        If runLength > 0 Then
            startColorSpec = VBA.vbNullString
            On Error Resume Next
            startColorSpec = VBA.CStr(markerInfo("ColorSpec"))
            On Error GoTo 0

            private_AddInlineRun outRuns, tagName, runStart, runLength, startColorSpec
        End If

        markerStack.Remove markerStack.Count
        private_TryHandleInlineToken = True
        Exit Function
    End If

    If Not private_HasMatchingClose(rawTextLower, searchFrom, delimiterType, tagName) Then Exit Function

    Set markerInfo = CreateObject("Scripting.Dictionary")
    markerInfo.CompareMode = 1
    markerInfo("Tag") = tagName
    markerInfo("Start") = VBA.Len(outText) + 1
    markerInfo("Delimiter") = delimiterType
    If VBA.Len(VBA.Trim$(colorSpec)) > 0 Then markerInfo("ColorSpec") = VBA.Trim$(colorSpec)
    markerStack.Add markerInfo

    private_TryHandleInlineToken = True
End Function

Private Function private_HasMatchingClose( _
    ByVal rawTextLower As String, _
    ByVal searchFrom As Long, _
    ByVal delimiterType As String, _
    ByVal tagName As String _
) As Boolean
    Dim closeToken As String

    If VBA.StrComp(delimiterType, DELIMITER_DOUBLE, VBA.vbBinaryCompare) = 0 Then
        closeToken = "[[/" & tagName & "]]"
    Else
        closeToken = "[/" & tagName & "]"
    End If

    private_HasMatchingClose = (VBA.InStr(searchFrom, rawTextLower, closeToken, VBA.vbBinaryCompare) > 0)
End Function

Private Function private_TryResolveInlineRunStyleFromInfo( _
    ByVal runInfo As Object, _
    ByRef outFontColor As Long, _
    ByRef outFontBold As Boolean, _
    ByRef outFontItalic As Boolean, _
    ByRef outFontUnderline As Boolean, _
    ByRef outApplyFontFlags As Boolean _
) As Boolean
    Dim tagName As String
    Dim colorSpec As String

    If runInfo Is Nothing Then Exit Function

    tagName = VBA.vbNullString
    colorSpec = VBA.vbNullString

    On Error Resume Next
    tagName = VBA.LCase$(VBA.Trim$(VBA.CStr(runInfo("Tag"))))
    colorSpec = VBA.Trim$(VBA.CStr(runInfo("ColorSpec")))
    On Error GoTo 0

    If tagName = "color" Or VBA.Len(colorSpec) > 0 Then
        If Not private_TryResolveInlineColorSpec(colorSpec, outFontColor) Then Exit Function
        outFontBold = False
        outFontItalic = False
        outFontUnderline = False
        outApplyFontFlags = False
        private_TryResolveInlineRunStyleFromInfo = True
        Exit Function
    End If

    If Not private_TryResolveInlineMarkerStyle(tagName, outFontColor, outFontBold, outFontItalic, outFontUnderline) Then Exit Function
    outApplyFontFlags = True
    private_TryResolveInlineRunStyleFromInfo = True
End Function

Private Sub private_AddInlineRun( _
    ByVal runs As Collection, _
    ByVal tagName As String, _
    ByVal startIndex As Long, _
    ByVal runLength As Long, _
    Optional ByVal colorSpec As String = VBA.vbNullString _
)
    Dim runInfo As Object

    If runs Is Nothing Then Exit Sub
    If startIndex <= 0 Or runLength <= 0 Then Exit Sub

    Set runInfo = CreateObject("Scripting.Dictionary")
    runInfo.CompareMode = 1
    runInfo("Tag") = VBA.LCase$(VBA.Trim$(tagName))
    runInfo("Start") = VBA.CLng(startIndex)
    runInfo("Length") = VBA.CLng(runLength)
    If VBA.Len(VBA.Trim$(colorSpec)) > 0 Then
        runInfo("ColorSpec") = VBA.Trim$(colorSpec)
    End If

    runs.Add runInfo
End Sub

Private Function private_IsSupportedInlineMarkerName(ByVal markerName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(markerName))
        Case "ok", "warn", "error", "accent", "muted", "red"
            private_IsSupportedInlineMarkerName = True
    End Select
End Function

Private Function private_TryResolveInlineMarkerStyle( _
    ByVal markerName As String, _
    ByRef outFontColor As Long, _
    ByRef outFontBold As Boolean, _
    ByRef outFontItalic As Boolean, _
    ByRef outFontUnderline As Boolean _
) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(markerName))
        Case "ok"
            outFontColor = VBA.RGB(146, 208, 80)
            outFontBold = True
            outFontItalic = False
            outFontUnderline = False

        Case "warn"
            outFontColor = VBA.RGB(255, 192, 0)
            outFontBold = True
            outFontItalic = False
            outFontUnderline = False

        Case "error"
            outFontColor = VBA.RGB(255, 99, 71)
            outFontBold = True
            outFontItalic = False
            outFontUnderline = False

        Case "accent"
            outFontColor = VBA.RGB(91, 155, 213)
            outFontBold = True
            outFontItalic = False
            outFontUnderline = False

        Case "muted"
            outFontColor = VBA.RGB(180, 180, 180)
            outFontBold = False
            outFontItalic = True
            outFontUnderline = False

        Case "red"
            outFontColor = VBA.RGB(255, 0, 0)
            outFontBold = True
            outFontItalic = False
            outFontUnderline = False

        Case Else
            Exit Function
    End Select

    private_TryResolveInlineMarkerStyle = True
End Function

Private Function private_TryResolveInlineColorSpec(ByVal colorSpec As String, ByRef outFontColor As Long) As Boolean
    Dim normalized As String

    normalized = VBA.LCase$(VBA.Trim$(colorSpec))
    If VBA.Len(normalized) = 0 Then Exit Function

    Select Case normalized
        Case "red"
            outFontColor = VBA.RGB(255, 0, 0)
        Case "green"
            outFontColor = VBA.RGB(0, 176, 80)
        Case "blue"
            outFontColor = VBA.RGB(0, 112, 192)
        Case "orange"
            outFontColor = VBA.RGB(237, 125, 49)
        Case "yellow"
            outFontColor = VBA.RGB(255, 192, 0)
        Case "purple"
            outFontColor = VBA.RGB(112, 48, 160)
        Case "teal"
            outFontColor = VBA.RGB(0, 128, 128)
        Case "cyan"
            outFontColor = VBA.RGB(0, 176, 240)
        Case "magenta"
            outFontColor = VBA.RGB(192, 0, 192)
        Case "brown"
            outFontColor = VBA.RGB(128, 64, 0)
        Case "gray", "grey"
            outFontColor = VBA.RGB(127, 127, 127)
        Case "black"
            outFontColor = VBA.RGB(0, 0, 0)
        Case "white"
            outFontColor = VBA.RGB(255, 255, 255)
        Case Else
            If private_TryParseHexColor(normalized, outFontColor) Then
                private_TryResolveInlineColorSpec = True
                Exit Function
            End If
            Exit Function
    End Select

    private_TryResolveInlineColorSpec = True
End Function

Private Function private_TryParseHexColor(ByVal colorSpec As String, ByRef outColor As Long) As Boolean
    Dim hexText As String
    Dim redPart As Long
    Dim greenPart As Long
    Dim bluePart As Long

    hexText = VBA.Trim$(colorSpec)
    If VBA.Left$(hexText, 1) = "#" Then hexText = VBA.Mid$(hexText, 2)
    If VBA.Len(hexText) <> 6 Then Exit Function

    On Error GoTo EH_HEX
    redPart = VBA.CLng("&H" & VBA.Mid$(hexText, 1, 2))
    greenPart = VBA.CLng("&H" & VBA.Mid$(hexText, 3, 2))
    bluePart = VBA.CLng("&H" & VBA.Mid$(hexText, 5, 2))

    outColor = VBA.RGB(redPart, greenPart, bluePart)
    private_TryParseHexColor = True
    Exit Function

EH_HEX:
    Err.Clear
End Function

Private Function private_Unquote(ByVal textValue As String) As String
    textValue = VBA.Trim$(textValue)
    If VBA.Len(textValue) < 2 Then
        private_Unquote = textValue
        Exit Function
    End If

    If (VBA.Left$(textValue, 1) = VBA.Chr$(34) And VBA.Right$(textValue, 1) = VBA.Chr$(34)) Or _
       (VBA.Left$(textValue, 1) = "'" And VBA.Right$(textValue, 1) = "'") Then
        private_Unquote = VBA.Mid$(textValue, 2, VBA.Len(textValue) - 2)
    Else
        private_Unquote = textValue
    End If
End Function


