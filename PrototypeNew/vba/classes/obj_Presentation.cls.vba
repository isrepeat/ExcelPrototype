VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_Presentation"
Option Explicit

Private m_EffectiveVisible As Boolean
Private m_StyleName As String
Private m_PartName As String
Private m_SpanRows As Long
Private m_SpacerRowsAfter As Long
Private m_InlineMarkersEnabled As Boolean

Private Sub Class_Initialize()
    m_EffectiveVisible = True
    m_InlineMarkersEnabled = False
End Sub

' //
' // API
' //
Public Property Get EffectiveVisible() As Boolean
    EffectiveVisible = m_EffectiveVisible
End Property

Public Property Let EffectiveVisible(ByVal value As Boolean)
    m_EffectiveVisible = CBool(value)
End Property

Public Property Get StyleName() As String
    StyleName = m_StyleName
End Property

Public Property Let StyleName(ByVal value As String)
    m_StyleName = CStr(value)
End Property

Public Property Get PartName() As String
    PartName = m_PartName
End Property

Public Property Let PartName(ByVal value As String)
    m_PartName = CStr(value)
End Property

Public Property Get SpanRows() As Long
    SpanRows = m_SpanRows
End Property

Public Property Let SpanRows(ByVal value As Long)
    If value < 0 Then
        m_SpanRows = 0
    Else
        m_SpanRows = CLng(value)
    End If
End Property

Public Property Get SpacerRowsAfter() As Long
    SpacerRowsAfter = m_SpacerRowsAfter
End Property

Public Property Let SpacerRowsAfter(ByVal value As Long)
    If value < 0 Then
        m_SpacerRowsAfter = 0
    Else
        m_SpacerRowsAfter = CLng(value)
    End If
End Property

Public Property Get InlineMarkersEnabled() As Boolean
    InlineMarkersEnabled = m_InlineMarkersEnabled
End Property

Public Property Let InlineMarkersEnabled(ByVal value As Boolean)
    m_InlineMarkersEnabled = CBool(value)
End Property

Public Function m_TryResolveInlineText( _
    ByVal rawText As String, _
    ByRef outText As String, _
    ByRef outRuns As Collection _
) As Boolean
    If Not m_InlineMarkersEnabled Then
        outText = rawText
        Set outRuns = Nothing
        m_TryResolveInlineText = True
        Exit Function
    End If

    If Not mp_TryParseInlineMarkers(rawText, outText, outRuns) Then Exit Function
    m_TryResolveInlineText = True
End Function

Public Function m_TryResolveInlineRunStyle( _
    ByVal markerName As String, _
    ByRef outFontColor As Long, _
    ByRef outFontBold As Boolean, _
    ByRef outFontItalic As Boolean, _
    ByRef outFontUnderline As Boolean _
) As Boolean
    If Not mp_TryResolveInlineMarkerStyle( _
        markerName, outFontColor, outFontBold, outFontItalic, outFontUnderline) Then Exit Function

    m_TryResolveInlineRunStyle = True
End Function

Public Sub m_ApplyInlineRuns(ByVal targetRange As Range, ByVal runs As Collection)
    Dim runInfo As Object
    Dim tagName As String
    Dim startIndex As Long
    Dim runLength As Long
    Dim fontColor As Long
    Dim fontBold As Boolean
    Dim fontItalic As Boolean
    Dim fontUnderline As Boolean

    If Not m_InlineMarkersEnabled Then Exit Sub
    If targetRange Is Nothing Then Exit Sub
    If runs Is Nothing Then Exit Sub

    For Each runInfo In runs
        tagName = LCase$(CStr(runInfo("Tag")))
        startIndex = CLng(runInfo("Start"))
        runLength = CLng(runInfo("Length"))

        If startIndex <= 0 Or runLength <= 0 Then GoTo ContinueRun
        If Not mp_TryResolveInlineMarkerStyle(tagName, fontColor, fontBold, fontItalic, fontUnderline) Then GoTo ContinueRun

        On Error Resume Next
        With targetRange.Characters(startIndex, runLength).Font
            .Color = fontColor
            .Bold = fontBold
            .Italic = fontItalic
            If fontUnderline Then
                .Underline = xlUnderlineStyleSingle
            Else
                .Underline = xlUnderlineStyleNone
            End If
        End With
        On Error GoTo 0

ContinueRun:
    Next runInfo
End Sub

' //
' // Internal
' //
Private Function mp_TryParseInlineMarkers( _
    ByVal rawText As String, _
    ByRef outText As String, _
    ByRef outRuns As Collection _
) As Boolean
    Dim markerStack As Collection
    Dim textLen As Long
    Dim i As Long
    Dim closePos As Long
    Dim markerToken As String
    Dim markerName As String
    Dim isClosing As Boolean
    Dim literalToken As String
    Dim markerInfo As Object
    Dim rawTextLower As String
    Dim runStart As Long
    Dim runLength As Long

    Set markerStack = New Collection
    Set outRuns = New Collection
    outText = vbNullString
    rawTextLower = LCase$(rawText)
    textLen = Len(rawText)
    i = 1

    Do While i <= textLen
        If i < textLen Then
            If Mid$(rawText, i, 2) = "[[" Then
                closePos = InStr(i + 2, rawText, "]]", vbBinaryCompare)
                If closePos > 0 Then
                    markerToken = Trim$(Mid$(rawText, i + 2, closePos - i - 2))
                    isClosing = (Len(markerToken) > 0 And Left$(markerToken, 1) = "/")

                    If isClosing Then
                        markerName = LCase$(Trim$(Mid$(markerToken, 2)))
                    Else
                        markerName = LCase$(Trim$(markerToken))
                    End If

                    If mp_IsSupportedInlineMarkerName(markerName) Then
                        If isClosing Then
                            If markerStack.Count > 0 Then
                                Set markerInfo = markerStack(markerStack.Count)
                                If StrComp(CStr(markerInfo("Tag")), markerName, vbBinaryCompare) = 0 Then
                                    runStart = CLng(markerInfo("Start"))
                                    runLength = Len(outText) - runStart + 1
                                    If runLength > 0 Then
                                        mp_AddInlineRun outRuns, markerName, runStart, runLength
                                    End If

                                    markerStack.Remove markerStack.Count
                                    i = closePos + 2
                                    GoTo ContinueLoop
                                End If
                            End If
                        Else
                            If InStr(closePos + 2, rawTextLower, "[[/" & markerName & "]]", vbBinaryCompare) > 0 Then
                                Set markerInfo = CreateObject("Scripting.Dictionary")
                                markerInfo.CompareMode = 1
                                markerInfo("Tag") = markerName
                                markerInfo("Start") = Len(outText) + 1
                                markerStack.Add markerInfo

                                i = closePos + 2
                                GoTo ContinueLoop
                            End If
                        End If
                    End If

                    literalToken = Mid$(rawText, i, closePos - i + 2)
                    outText = outText & literalToken
                    i = closePos + 2
                    GoTo ContinueLoop
                End If
            End If
        End If

        outText = outText & Mid$(rawText, i, 1)
        i = i + 1

ContinueLoop:
    Loop

    mp_TryParseInlineMarkers = True
End Function

Private Sub mp_AddInlineRun( _
    ByVal runs As Collection, _
    ByVal tagName As String, _
    ByVal startIndex As Long, _
    ByVal runLength As Long _
)
    Dim runInfo As Object

    If runs Is Nothing Then Exit Sub
    If startIndex <= 0 Or runLength <= 0 Then Exit Sub

    Set runInfo = CreateObject("Scripting.Dictionary")
    runInfo.CompareMode = 1
    runInfo("Tag") = LCase$(Trim$(tagName))
    runInfo("Start") = CLng(startIndex)
    runInfo("Length") = CLng(runLength)

    runs.Add runInfo
End Sub

Private Function mp_IsSupportedInlineMarkerName(ByVal markerName As String) As Boolean
    Select Case LCase$(Trim$(markerName))
        Case "ok", "warn", "error", "accent", "muted"
            mp_IsSupportedInlineMarkerName = True
    End Select
End Function

Private Function mp_TryResolveInlineMarkerStyle( _
    ByVal markerName As String, _
    ByRef outFontColor As Long, _
    ByRef outFontBold As Boolean, _
    ByRef outFontItalic As Boolean, _
    ByRef outFontUnderline As Boolean _
) As Boolean
    Select Case LCase$(Trim$(markerName))
        Case "ok"
            outFontColor = RGB(146, 208, 80)
            outFontBold = True
            outFontItalic = False
            outFontUnderline = False

        Case "warn"
            outFontColor = RGB(255, 192, 0)
            outFontBold = True
            outFontItalic = False
            outFontUnderline = False

        Case "error"
            outFontColor = RGB(255, 99, 71)
            outFontBold = True
            outFontItalic = False
            outFontUnderline = False

        Case "accent"
            outFontColor = RGB(91, 155, 213)
            outFontBold = True
            outFontItalic = False
            outFontUnderline = False

        Case "muted"
            outFontColor = RGB(180, 180, 180)
            outFontBold = False
            outFontItalic = True
            outFontUnderline = False

        Case Else
            Exit Function
    End Select

    mp_TryResolveInlineMarkerStyle = True
End Function
