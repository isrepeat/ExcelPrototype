Option Explicit

Public Sub fn_FilterBySurnameFirstWord()
    Dim rngData As Range, colIndex As Long
    Dim query As String
    Dim ws As Worksheet

    Set ws = ActiveSheet

    ' Match records whose selected column starts with the supplied text.
    query = private_NormalizeQuery(InputBox( _
        "Enter a surname, a surname and first name, or a full name." & vbCrLf & _
        "The filter matches values that start with the entered text.", _
        "Filter by name"))
    If Len(query) = 0 Then Exit Sub

    ' Filter the Excel Table when the active cell belongs to one.
    On Error Resume Next
    If Not ActiveCell.ListObject Is Nothing Then
        With ActiveCell.ListObject
            colIndex = ActiveCell.Column - .Range.Columns(1).Column + 1
            .Range.AutoFilter Field:=colIndex, Criteria1:=query & "*"
        End With
        Exit Sub
    End If
    On Error GoTo 0

    ' Otherwise, filter the active cell's current region.
    Set rngData = ActiveCell.CurrentRegion
    If rngData.Rows.Count < 2 Then Exit Sub

    colIndex = ActiveCell.Column - rngData.Column + 1
    If colIndex < 1 Or colIndex > rngData.Columns.Count Then Exit Sub

    If Not ws.AutoFilterMode Then rngData.AutoFilter
    rngData.AutoFilter Field:=colIndex, Criteria1:=query & "*"
End Sub

Private Function private_NormalizeQuery(ByVal s As String) As String
    ' Normalize line breaks, non-breaking spaces, and repeated spaces.
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, ChrW(160), " ") ' non-breaking space
    s = Trim(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    private_NormalizeQuery = s
End Function

Public Sub fn_FilterSurname()
    fn_FilterBySurname_FirstWord
End Sub
