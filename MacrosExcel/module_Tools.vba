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
    Dim rowsCount As Long
    Dim firstColOffset As Long
    Dim blockRange As Range
    Dim swapRange As Range
    Dim tmpRange As Range
    Dim tmpCol As Long

    If TypeName(Selection) <> "Range" Then Exit Sub
    Set sel = Selection

    If sel.Areas.count > 1 Then Exit Sub

    On Error Resume Next
    Set tbl = ActiveCell.ListObject
    On Error GoTo 0

    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    Set tablePart = Intersect(sel, tbl.DataBodyRange)
    If tablePart Is Nothing Then Exit Sub

    firstRow = tablePart.Row - tbl.DataBodyRange.Row + 1
    lastRow = tablePart.Row + tablePart.Rows.count - tbl.DataBodyRange.Row
    rowsCount = lastRow - firstRow + 1

    firstColOffset = ActiveCell.Column - tbl.DataBodyRange.Column + 1
    If firstColOffset < 1 Or firstColOffset > tbl.ListColumns.count Then firstColOffset = 1

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    tmpCol = tbl.Range.Column + tbl.Range.Columns.count + 2
    Set tmpRange = tbl.Parent.Cells(tbl.Range.Row, tmpCol).Resize(1, tbl.ListColumns.count)

    If Direction = -1 Then

        If firstRow <= 1 Then GoTo SafeExit

        Set blockRange = tbl.DataBodyRange.Rows(firstRow).Resize(rowsCount)
        Set swapRange = tbl.DataBodyRange.Rows(firstRow - 1)

        swapRange.Copy Destination:=tmpRange
        blockRange.Copy Destination:=tbl.DataBodyRange.Rows(firstRow - 1)
        tmpRange.Copy Destination:=tbl.DataBodyRange.Rows(lastRow)

        tmpRange.Clear

        tbl.DataBodyRange.Rows(firstRow - 1).Resize(rowsCount).Select
        tbl.DataBodyRange.Cells(firstRow - 1, firstColOffset).Activate

    ElseIf Direction = 1 Then

        If lastRow >= tbl.ListRows.count Then GoTo SafeExit

        Set blockRange = tbl.DataBodyRange.Rows(firstRow).Resize(rowsCount)
        Set swapRange = tbl.DataBodyRange.Rows(lastRow + 1)

        swapRange.Copy Destination:=tmpRange
        blockRange.Copy Destination:=tbl.DataBodyRange.Rows(firstRow + 1)
        tmpRange.Copy Destination:=tbl.DataBodyRange.Rows(firstRow)

        tmpRange.Clear

        tbl.DataBodyRange.Rows(firstRow + 1).Resize(rowsCount).Select
        tbl.DataBodyRange.Cells(firstRow + 1, firstColOffset).Activate

    End If

SafeExit:
    Application.CutCopyMode = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub