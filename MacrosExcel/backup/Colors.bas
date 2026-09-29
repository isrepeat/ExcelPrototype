Option Explicit

'========================
' Helpers
'========================
Private Function HexToLong(ByVal hexColor As String) As Long
    ' Accepts "#RRGGBB" or "RRGGBB"
    Dim s As String
    s = Replace$(Trim$(hexColor), "#", "")
    If Len(s) <> 6 Then
        Err.Raise vbObjectError + 1, "HexToLong", "HEX must be 6 chars (RRGGBB)."
    End If

    Dim r As Long, g As Long, b As Long
    r = CLng("&H" & Mid$(s, 1, 2))
    g = CLng("&H" & Mid$(s, 3, 2))
    b = CLng("&H" & Mid$(s, 5, 2))

    ' Excel/VBA uses BGR for .Color
    HexToLong = RGB(r, g, b)
End Function

Private Sub ApplyFontColor(ByVal hexColor As String)
    On Error Resume Next
    If TypeName(Selection) = "Range" Then
        Selection.Font.Color = HexToLong(hexColor)
    End If
End Sub

Private Sub ApplyFillColor(ByVal hexColor As String)
    On Error Resume Next
    If TypeName(Selection) = "Range" Then
        Selection.Interior.pattern = xlSolid
        Selection.Interior.Color = HexToLong(hexColor)
    End If
End Sub

'========================
' FONT COLORS (Text)
'========================
Public Sub Font_Red()
    ApplyFontColor "#FF0000"
End Sub

Public Sub Font_Yellow()
    ApplyFontColor "#FFFF00"
End Sub

Public Sub Font_Gray()
    ApplyFontColor "#ADADAD"
End Sub

Public Sub Font_White()
    ApplyFontColor "#FFFFFF"
End Sub

'========================
' FILL COLORS (Background)
'========================
Public Sub Fill_Red()
    ApplyFillColor "#FF0000"
End Sub

Public Sub Fill_Yellow()
    ApplyFillColor "#FFFF00"
End Sub

' Optional: remove fill
Public Sub Fill_None()
    On Error Resume Next
    If TypeName(Selection) = "Range" Then Selection.Interior.pattern = xlNone
End Sub