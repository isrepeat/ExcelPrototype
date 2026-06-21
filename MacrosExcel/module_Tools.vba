Option Explicit

Sub fn_ExtendTable()
   Dim tbl As ListObject
    Dim newRows As Long

    Set tbl = ActiveSheet.ListObjects(1)
    newRows = 10

    tbl.Resize tbl.Range.Resize(tbl.Range.Rows.count + newRows)
End Sub

Public Sub fn_MoveTableRowsUp()
    private_MoveSelectedTableRows -1
End Sub

Public Sub fn_MoveTableRowsDown()
    private_MoveSelectedTableRows 1
End Sub

Private Sub private_MoveSelectedTableRows(ByVal Direction As Long)
    Dim tbl As ListObject
    Dim sel As Range
    Dim tablePart As Range
    Dim firstRow As Long, lastRow As Long
    Dim firstColOffset As Long
    Dim blockData As Variant
    Dim swapData As Variant

    If TypeName(Selection) <> "Range" Then Exit Sub
    Set sel = Selection

    On Error Resume Next
    Set tbl = sel.Cells(1, 1).ListObject
    On Error GoTo 0

    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    Set tablePart = Intersect(sel, tbl.DataBodyRange)
    If tablePart Is Nothing Then Exit Sub

    ' Проверяем, что выделение непрерывное
    If sel.Areas.count > 1 Then Exit Sub

    firstRow = tablePart.Row - tbl.DataBodyRange.Row + 1
    lastRow = tablePart.Row + tablePart.Rows.count - tbl.DataBodyRange.Row

    firstColOffset = ActiveCell.Column - tbl.Range.Column + 1

    If Direction = -1 Then
        If firstRow <= 1 Then Exit Sub

        blockData = tbl.DataBodyRange.Rows(firstRow & ":" & lastRow).value
        swapData = tbl.DataBodyRange.Rows(firstRow - 1).value

        tbl.DataBodyRange.Rows(firstRow - 1).Resize(UBound(blockData, 1)).value = blockData
        tbl.DataBodyRange.Rows(lastRow).value = swapData
        tbl.DataBodyRange.Rows(firstRow - 1 & ":" & lastRow - 1).Select

    ElseIf Direction = 1 Then
        If lastRow >= tbl.ListRows.count Then Exit Sub

        blockData = tbl.DataBodyRange.Rows(firstRow & ":" & lastRow).value
        swapData = tbl.DataBodyRange.Rows(lastRow + 1).value

        tbl.DataBodyRange.Rows(firstRow + 1).Resize(UBound(blockData, 1)).value = blockData
        tbl.DataBodyRange.Rows(firstRow).value = swapData
        tbl.DataBodyRange.Rows(firstRow + 1 & ":" & lastRow + 1).Select
    End If

    tbl.DataBodyRange.Cells(IIf(Direction = -1, firstRow - 1, firstRow + 1), firstColOffset).Activate
End Sub
