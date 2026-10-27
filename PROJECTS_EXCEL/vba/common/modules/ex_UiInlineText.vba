Option Explicit

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_WrapCaption( _
    ByVal shape As Object, _
    Optional ByVal preservePadding As Boolean = False _
)
    Dim textFrame As Object
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo EH
    Set textFrame = shape.TextFrame2
    textFrame.AutoSize = msoAutoSizeNone
    If Not preservePadding Then
        textFrame.MarginTop = 0
        textFrame.MarginBottom = 0
    End If
    textFrame.WordWrap = msoTrue
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    VBA.Err.Raise errorNumber, "Caption.Wrap", "Shape '" & shape.Name & "': " & errorDescription
End Sub

Public Sub fn_ApplyCaption( _
    ByVal shape As Object, _
    ByVal rawText As String, _
    ByVal baseColor As Long _
)
    Dim plain As String
    Dim remaining As String
    Dim opening As Long
    Dim closing As Long
    Dim ending As Long
    Dim colorText As String
    Dim segment As String
    Dim runs As Collection
    Dim run As Variant
    Dim color As Long

    Set runs = New Collection
    remaining = rawText
    Do
        opening = VBA.InStr(remaining, "[[color=")
        If opening = 0 Then
            plain = plain & remaining
            Exit Do
        End If
        plain = plain & VBA.Left$(remaining, opening - 1)
        remaining = VBA.Mid$(remaining, opening)
        closing = VBA.InStr(remaining, "]]")
        ending = VBA.InStr(remaining, "[[/color]]")
        If closing = 0 Or ending <= closing Then VBA.Err.Raise 5, , "Invalid inline color marker."
        colorText = VBA.Mid$(remaining, 9, closing - 9)
        If VBA.Len(colorText) <> 7 Or VBA.Left$(colorText, 1) <> "#" Then VBA.Err.Raise 5, , "Invalid inline color: " & colorText
        If VBA.Mid$(colorText, 2) Like "*[!0-9A-Fa-f]*" Then VBA.Err.Raise 5, , "Invalid inline color: " & colorText
        segment = VBA.Mid$(remaining, closing + 2, ending - closing - 2)
        If VBA.InStr(segment, "[[") > 0 Then VBA.Err.Raise 5, , "Nested inline markers are not supported."
        color = VBA.RGB(VBA.CLng("&H" & VBA.Mid$(colorText, 2, 2)), _
            VBA.CLng("&H" & VBA.Mid$(colorText, 4, 2)), VBA.CLng("&H" & VBA.Mid$(colorText, 6, 2)))
        If VBA.Len(segment) > 0 Then runs.Add VBA.Array(VBA.Len(plain) + 1, VBA.Len(segment), color)
        plain = plain & segment
        remaining = VBA.Mid$(remaining, ending + 10)
    Loop
    If VBA.InStr(plain, "[[/color]]") > 0 Then VBA.Err.Raise 5, , "Unmatched inline color marker."
    shape.TextFrame2.TextRange.Text = plain
    ' Replacing rich text can inherit the first character color; reset the entire caption first.
    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = baseColor
    For Each run In runs
        shape.TextFrame2.TextRange.Characters(VBA.CLng(run(0)), VBA.CLng(run(1))).Font.Fill.ForeColor.RGB = VBA.CLng(run(2))
    Next run
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------