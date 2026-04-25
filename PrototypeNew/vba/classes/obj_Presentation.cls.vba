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
    m_EffectiveVisible = VBA.CBool(value)
End Property

Public Property Get StyleName() As String
    StyleName = m_StyleName
End Property

Public Property Let StyleName(ByVal value As String)
    m_StyleName = VBA.CStr(value)
End Property

Public Property Get PartName() As String
    PartName = m_PartName
End Property

Public Property Let PartName(ByVal value As String)
    m_PartName = VBA.CStr(value)
End Property

Public Property Get SpanRows() As Long
    SpanRows = m_SpanRows
End Property

Public Property Let SpanRows(ByVal value As Long)
    If value < 0 Then
        m_SpanRows = 0
    Else
        m_SpanRows = VBA.CLng(value)
    End If
End Property

Public Property Get SpacerRowsAfter() As Long
    SpacerRowsAfter = m_SpacerRowsAfter
End Property

Public Property Let SpacerRowsAfter(ByVal value As Long)
    If value < 0 Then
        m_SpacerRowsAfter = 0
    Else
        m_SpacerRowsAfter = VBA.CLng(value)
    End If
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
        tagName = VBA.LCase$(VBA.CStr(runInfo("Tag")))
        startIndex = VBA.CLng(runInfo("Start"))
        runLength = VBA.CLng(runInfo("Length"))

        If startIndex <= 0 Or runLength <= 0 Then GoTo ContinueRun
        If Not private_TryResolveInlineMarkerStyle(tagName, fontColor, fontBold, fontItalic, fontUnderline) Then GoTo ContinueRun

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
Private Function private_TryParseInlineMarkers( _
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
    outText = VBA.vbNullString
    rawTextLower = VBA.LCase$(rawText)
    textLen = VBA.Len(rawText)
    i = 1

    Do While i <= textLen
        If i < textLen Then
            If VBA.Mid$(rawText, i, 2) = "[[" Then
                closePos = VBA.InStr(i + 2, rawText, "]]", VBA.vbBinaryCompare)
                If closePos > 0 Then
                    markerToken = VBA.Trim$(VBA.Mid$(rawText, i + 2, closePos - i - 2))
                    isClosing = (VBA.Len(markerToken) > 0 And VBA.Left$(markerToken, 1) = "/")

                    If isClosing Then
                        markerName = VBA.LCase$(VBA.Trim$(VBA.Mid$(markerToken, 2)))
                    Else
                        markerName = VBA.LCase$(VBA.Trim$(markerToken))
                    End If

                    If private_IsSupportedInlineMarkerName(markerName) Then
                        If isClosing Then
                            If markerStack.Count > 0 Then
                                Set markerInfo = markerStack(markerStack.Count)
                                If VBA.StrComp(VBA.CStr(markerInfo("Tag")), markerName, VBA.vbBinaryCompare) = 0 Then
                                    runStart = VBA.CLng(markerInfo("Start"))
                                    runLength = VBA.Len(outText) - runStart + 1
                                    If runLength > 0 Then
                                        private_AddInlineRun outRuns, markerName, runStart, runLength
                                    End If

                                    markerStack.Remove markerStack.Count
                                    i = closePos + 2
                                    GoTo ContinueLoop
                                End If
                            End If
                        Else
                            If VBA.InStr(closePos + 2, rawTextLower, "[[/" & markerName & "]]", VBA.vbBinaryCompare) > 0 Then
                                Set markerInfo = CreateObject("Scripting.Dictionary")
                                markerInfo.CompareMode = 1
                                markerInfo("Tag") = markerName
                                markerInfo("Start") = VBA.Len(outText) + 1
                                markerStack.Add markerInfo

                                i = closePos + 2
                                GoTo ContinueLoop
                            End If
                        End If
                    End If

                    literalToken = VBA.Mid$(rawText, i, closePos - i + 2)
                    outText = outText & literalToken
                    i = closePos + 2
                    GoTo ContinueLoop
                End If
            End If
        End If

        outText = outText & VBA.Mid$(rawText, i, 1)
        i = i + 1

ContinueLoop:
    Loop

    private_TryParseInlineMarkers = True
End Function

Private Sub private_AddInlineRun( _
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
    runInfo("Tag") = VBA.LCase$(VBA.Trim$(tagName))
    runInfo("Start") = VBA.CLng(startIndex)
    runInfo("Length") = VBA.CLng(runLength)

    runs.Add runInfo
End Sub

Private Function private_IsSupportedInlineMarkerName(ByVal markerName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(markerName))
        Case "ok", "warn", "error", "accent", "muted"
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

        Case Else
            Exit Function
    End Select

    private_TryResolveInlineMarkerStyle = True
End Function
