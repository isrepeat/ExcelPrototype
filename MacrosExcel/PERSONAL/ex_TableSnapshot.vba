Option Explicit

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_ExportVisibleTable()
    Dim sourceTable As ListObject
    Dim sourceBook As Workbook
    Dim resultBook As Workbook
    Dim resultSheet As Worksheet
    Dim resultTable As ListObject
    Dim sourceColumn As ListColumn
    Dim visibleColumns As Collection
    Dim visibleBlocks As Collection
    Dim sourceBlock As Range
    Dim sourceCell As Range
    Dim targetStart As Range
    Dim targetCell As Range
    Dim targetRow As Long
    Dim targetColumn As Long
    Dim bodyRows As Long
    Dim totalRows As Long
    Dim previousEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim previousAlerts As Boolean
    Dim stateChanged As Boolean
    Dim errorMessage As String

    On Error GoTo EH
    If Not TypeOf Application.ActiveSheet Is Worksheet Then _
        Err.Raise 5, "PERSONAL.Table.WorksheetRequired"
    Set sourceTable = Application.ActiveCell.ListObject
    If sourceTable Is Nothing Then _
        Err.Raise 5, "PERSONAL.Table.SelectCellInsideTable"
    Set sourceBook = sourceTable.Parent.Parent
    If Len(sourceBook.Path) = 0 Then _
        Err.Raise 5, "PERSONAL.Table.SaveSourceWorkbookFirst"

    ' Capture the visible rows and columns in worksheet order.
    Set visibleColumns = New Collection
    For Each sourceColumn In sourceTable.ListColumns
        If Not sourceColumn.Range.EntireColumn.Hidden Then
            visibleColumns.Add sourceColumn
        End If
    Next sourceColumn
    If visibleColumns.Count = 0 Then _
        Err.Raise 5, "PERSONAL.Table.NoVisibleColumns"
    Set visibleBlocks = private_VisibleRowBlocks(sourceTable.DataBodyRange)
    For Each sourceBlock In visibleBlocks
        bodyRows = bodyRows + sourceBlock.Rows.Count
    Next sourceBlock
    If sourceTable.ShowTotals Then
        If Not sourceTable.TotalsRowRange.EntireRow.Hidden Then
            totalRows = 1
        End If
    End If
    If bodyRows + totalRows + 4 > sourceTable.Parent.Rows.Count Or _
            visibleColumns.Count + 1 > sourceTable.Parent.Columns.Count Then _
        Err.Raise 6, "PERSONAL.Table.SnapshotDoesNotFit"

    ' Create the destination and transfer the workbook-local table style if needed.
    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    previousAlerts = Application.DisplayAlerts
    stateChanged = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set resultBook = Application.Workbooks.Add(xlWBATWorksheet)
    Set resultSheet = resultBook.Worksheets(1)
    resultSheet.Name = sourceTable.Parent.Name
    resultBook.Date1904 = sourceBook.Date1904
    private_ImportTableStyle sourceTable, resultBook
    resultBook.Styles("Normal").Interior.Color = RGB(40, 37, 30)
    resultBook.Styles("Normal").Font.Color = RGB(255, 255, 255)
    resultBook.Windows(1).DisplayGridlines = False
    Set targetStart = resultSheet.Range("B4")

    ' Write headers before creating the ListObject so text is kept verbatim.
    targetColumn = 0
    For Each sourceColumn In visibleColumns
        targetColumn = targetColumn + 1
        Set targetCell = targetStart.Cells(1, targetColumn)
        targetCell.NumberFormat = "@"
        targetCell.Value2 = sourceColumn.Name
        targetCell.EntireColumn.ColumnWidth = sourceColumn.Range.EntireColumn.ColumnWidth
    Next sourceColumn
    Set resultTable = resultSheet.ListObjects.Add( _
        xlSrcRange, targetStart.Resize(1 + bodyRows, visibleColumns.Count), , xlYes)
    resultTable.Name = sourceTable.Name
    resultTable.TableStyle = sourceTable.TableStyle
    resultTable.ShowTableStyleFirstColumn = sourceTable.ShowTableStyleFirstColumn
    resultTable.ShowTableStyleLastColumn = sourceTable.ShowTableStyleLastColumn
    resultTable.ShowTableStyleRowStripes = sourceTable.ShowTableStyleRowStripes
    resultTable.ShowTableStyleColumnStripes = sourceTable.ShowTableStyleColumnStripes
    If bodyRows = 0 Then
        If resultTable.ListRows.Count > 0 Then resultTable.ListRows(1).Delete
    End If
    resultTable.ShowTotals = (totalRows = 1)

    ' Copy cell formats and cached values in contiguous blocks.
    targetColumn = 0
    For Each sourceColumn In visibleColumns
        targetColumn = targetColumn + 1
        If sourceTable.ShowHeaders Then
            Set sourceCell = sourceTable.HeaderRowRange.Cells(1, sourceColumn.Index)
            private_CopySnapshotCells sourceCell, targetStart.Cells(1, targetColumn), False
        End If
        targetRow = 2
        For Each sourceBlock In visibleBlocks
            private_CopySnapshotCells sourceBlock.Columns(sourceColumn.Index), _
                targetStart.Cells(targetRow, targetColumn), True
            targetRow = targetRow + sourceBlock.Rows.Count
        Next sourceBlock
        If totalRows = 1 Then
            Set sourceCell = sourceTable.TotalsRowRange.Cells(1, sourceColumn.Index)
            private_CopySnapshotCells sourceCell, resultTable.TotalsRowRange.Cells(1, targetColumn), True
        End If
    Next sourceColumn
    targetRow = 2
    For Each sourceBlock In visibleBlocks
        For Each sourceCell In sourceBlock.Columns(1).Cells
            targetStart.Rows(targetRow).RowHeight = sourceCell.RowHeight
            targetRow = targetRow + 1
        Next sourceCell
    Next sourceBlock
    If sourceTable.ShowHeaders Then
        If Not sourceTable.HeaderRowRange.EntireRow.Hidden Then
            targetStart.RowHeight = sourceTable.HeaderRowRange.RowHeight
        End If
    End If
    If totalRows = 1 Then resultTable.TotalsRowRange.RowHeight = sourceTable.TotalsRowRange.RowHeight
    resultSheet.Activate
    targetStart.Select
    Application.DisplayAlerts = True
    resultBook.SaveAs Filename:=private_NextSnapshotPath(sourceBook.Path, sourceTable.Name), _
        FileFormat:=xlOpenXMLWorkbook

CleanExit:
    If stateChanged Then
        Application.CutCopyMode = False
        Application.DisplayAlerts = previousAlerts
        Application.ScreenUpdating = previousScreenUpdating
        Application.EnableEvents = previousEvents
    End If
    If Len(errorMessage) > 0 Then
        ex_Core.fn_Diagnostic_WriteLog "TABLE_SNAPSHOT_ERROR | " & errorMessage
        ex_WindowsUi.fn_ShowMessage errorMessage, VBA.vbExclamation
    End If
    Exit Sub
EH:
    errorMessage = Err.Source & VBA.vbCrLf & Err.Description
    On Error Resume Next
    If Not resultBook Is Nothing Then resultBook.Close SaveChanges:=False
    On Error GoTo 0
    GoTo CleanExit
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_NextSnapshotPath( _
    ByVal folderPath As String, _
    ByVal tableName As String _
) As String
    Dim fileSystem As Object
    Dim basePath As String
    Dim candidatePath As String
    Dim version As Long

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FolderExists(folderPath) Then _
        Err.Raise 76, "PERSONAL.Table.SnapshotFolderUnavailable", folderPath

    basePath = fileSystem.BuildPath(folderPath, "Slice_" & Replace$(tableName, "\", "_"))
    candidatePath = basePath & ".xlsx"
    version = 1
    Do While fileSystem.FileExists(candidatePath) Or fileSystem.FolderExists(candidatePath)
        version = version + 1
        candidatePath = basePath & "_v" & VBA.CStr(version) & ".xlsx"
    Loop
    private_NextSnapshotPath = candidatePath
End Function

Private Function private_VisibleRowBlocks(ByVal sourceBody As Range) As Collection
    Dim blocks As Collection
    Dim rowIndex As Long
    Dim firstRow As Long

    Set blocks = New Collection
    Set private_VisibleRowBlocks = blocks
    If sourceBody Is Nothing Then Exit Function
    For rowIndex = 1 To sourceBody.Rows.Count
        If sourceBody.Rows(rowIndex).EntireRow.Hidden Then
            If firstRow > 0 Then
                blocks.Add sourceBody.Rows(firstRow).Resize(rowIndex - firstRow)
                firstRow = 0
            End If
        ElseIf firstRow = 0 Then
            firstRow = rowIndex
        End If
    Next rowIndex
    If firstRow > 0 Then blocks.Add sourceBody.Rows(firstRow).Resize(rowIndex - firstRow)
End Function

Private Sub private_CopySnapshotCells( _
    ByVal sourceCells As Range, _
    ByVal targetCell As Range, _
    ByVal includeValues As Boolean _
)
    sourceCells.Copy
    targetCell.PasteSpecial Paste:=xlPasteFormats
    If includeValues Then targetCell.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
End Sub

Private Sub private_ImportTableStyle( _
    ByVal sourceTable As ListObject, _
    ByVal resultBook As Workbook _
)
    Dim sourceStyle As TableStyle
    Dim importedStyle As TableStyle
    Dim stagingSheet As Worksheet

    If Len(sourceTable.TableStyle) = 0 Then Exit Sub
    Set sourceStyle = sourceTable.Parent.Parent.TableStyles(sourceTable.TableStyle)
    If sourceStyle.BuiltIn Then Exit Sub

    ' Excel transfers a custom style when it copies the entire table across workbooks.
    Set stagingSheet = resultBook.Worksheets.Add
    sourceTable.Range.Copy Destination:=stagingSheet.Range("A1")
    Application.CutCopyMode = False
    Set importedStyle = resultBook.TableStyles(sourceStyle.Name)
    stagingSheet.Delete
End Sub