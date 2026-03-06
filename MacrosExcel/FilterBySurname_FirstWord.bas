Option Explicit

Sub FilterBySurname_FirstWord()
    Dim rngData As Range, colIndex As Long
    Dim query As String
    Dim ws As Worksheet

    Set ws = ActiveSheet

    ' Ask user: surname OR surname + name OR full name
    ' Filter logic: starts with entered text
    query = NormalizeQuery(InputBox( _
        "Ââåäèòå: Ôàìèëèþ èëè Ôàìèëèÿ Èìÿ (èëè Ôàìèëèÿ Èìÿ Îò÷åñòâî)." & vbCrLf & _
        "Ôèëüòð: íà÷èíàåòñÿ ñ ââåäåííîãî òåêñòà.", _
        "Ôèëüòð ïî ÔÈÎ"))
    If Len(query) = 0 Then Exit Sub

    ' If active cell is inside Excel Table (ListObject) - filter table
    On Error Resume Next
    If Not ActiveCell.ListObject Is Nothing Then
        With ActiveCell.ListObject
            colIndex = ActiveCell.Column - .Range.Columns(1).Column + 1
            .Range.AutoFilter Field:=colIndex, Criteria1:=query & "*"
        End With
        Exit Sub
    End If
    On Error GoTo 0

    ' Regular range (CurrentRegion)
    Set rngData = ActiveCell.CurrentRegion
    If rngData.Rows.Count < 2 Then Exit Sub

    colIndex = ActiveCell.Column - rngData.Column + 1
    If colIndex < 1 Or colIndex > rngData.Columns.Count Then Exit Sub

    If Not ws.AutoFilterMode Then rngData.AutoFilter
    rngData.AutoFilter Field:=colIndex, Criteria1:=query & "*"
End Sub

Private Function NormalizeQuery(ByVal s As String) As String
    ' Normalize spaces and line breaks
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, ChrW(160), " ") ' non-breaking space
    s = Trim(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeQuery = s
End Function


Sub FilterSurname()
    FilterBySurname_FirstWord
End Sub