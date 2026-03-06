Option Explicit

'========================
' Helpers (WORD)
'========================

Private Sub ApplyFontColor(ByVal r As Long, ByVal g As Long, ByVal b As Long)
    On Error Resume Next
    If Selection.Range Is Nothing Then Exit Sub
    Selection.Range.Font.Color = RGB(r, g, b)
End Sub

' ????????? ?????? (??????). ? Word ??? ?? RGB, ? ????? ???????????????? ??????.
Private Sub ApplyHighlight(ByVal idx As WdColorIndex)
    On Error Resume Next
    If Selection.Range Is Nothing Then Exit Sub
    Selection.Range.HighlightColorIndex = idx
End Sub

' ??????? (Shading) — ???????? ??? ???????, ????? ??????, ???????? ? Shading.
Private Sub ApplyShadingColor(ByVal r As Long, ByVal g As Long, ByVal b As Long)
    On Error Resume Next
    If Selection.Range Is Nothing Then Exit Sub
    Selection.Range.Shading.BackgroundPatternColor = RGB(r, g, b)
End Sub

'========================
' FONT COLORS (Text)
'========================
Public Sub Font_Red()
    ApplyFontColor 255, 0, 0
End Sub

Public Sub Font_Yellow()
    ApplyFontColor 255, 255, 0
End Sub

Public Sub Font_Gray()
    ApplyFontColor 173, 173, 173
End Sub

Public Sub Font_White()
    ApplyFontColor 255, 255, 255
End Sub

'========================
' HIGHLIGHT (Text marker)
'========================
Public Sub Fill_Red()
    ApplyHighlight wdRed
End Sub

Public Sub Fill_Yellow()
    ApplyHighlight wdYellow
End Sub

Public Sub Fill_None()
    On Error Resume Next
    If Selection.Range Is Nothing Then Exit Sub
    Selection.Range.HighlightColorIndex = wdNoHighlight
End Sub

'========================
' OPTIONAL: "True" shading fill (paragraph/cell)
'========================
Public Sub Shading_Red()
    ApplyShadingColor 255, 0, 0
End Sub

Public Sub Shading_Yellow()
    ApplyShadingColor 255, 255, 0
End Sub

Public Sub Shading_None()
    On Error Resume Next
    If Selection.Range Is Nothing Then Exit Sub
    Selection.Range.Shading.BackgroundPatternColor = wdColorAutomatic
End Sub