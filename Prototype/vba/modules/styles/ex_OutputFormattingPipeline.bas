Attribute VB_Name = "ex_OutputFormattingPipeline"
Option Explicit

Public Sub m_FormatAsTable(ByVal ws As Worksheet, ByVal startRow As Long, ByVal rowCount As Long, ByVal colCount As Long)
    Dim headerRange As Range
    Dim allRange As Range

    Set headerRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, colCount))
    Set allRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + rowCount - 1, colCount))

    allRange.Font.Name = "Segoe UI"
    allRange.Font.Size = 10
    headerRange.Font.Bold = True
    allRange.HorizontalAlignment = xlCenter
    allRange.VerticalAlignment = xlCenter
    allRange.EntireColumn.AutoFit
    allRange.AutoFilter

    ws.Activate
    ActiveWindow.FreezePanes = False
End Sub

Public Sub m_ApplyComparingStyleLayers( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal rowCount As Long, _
    ByVal colCount As Long, _
    ByRef baseStyle As t_BaseSheetStyle, _
    ByRef outputStyle As t_OutputSheetStyle, _
    ByVal hasOutputStyle As Boolean _
)
    Dim layerOrder As Variant
    Dim layerName As Variant
    Dim fullRowCount As Long

    If ws Is Nothing Then Exit Sub
    If Not ex_SheetStylesXmlProvider.m_GetLayerOrder(hasOutputStyle, layerOrder, ThisWorkbook) Then Exit Sub

    fullRowCount = startRow + rowCount - 1

    For Each layerName In layerOrder
        Select Case CStr(layerName)
            Case ex_SheetStylesXmlProvider.LAYER_BASE
                ex_SheetStylesXmlProvider.m_ApplyBaseLayer ws, fullRowCount, colCount, baseStyle
            Case ex_SheetStylesXmlProvider.LAYER_OUTPUT
                mp_ApplyOutputStyleToResult ws, startRow, rowCount, colCount, outputStyle
                ex_SheetStylesXmlProvider.m_ApplyStatusLayer ws, startRow, rowCount, colCount, outputStyle
        End Select
    Next layerName
End Sub

Public Sub m_ApplyTimelineStyleLayers( _
    ByVal ws As Worksheet, _
    ByVal headerRows As Collection, _
    ByVal sectionRows As Collection, _
    ByRef outputStyle As t_OutputSheetStyle, _
    ByRef baseStyle As t_BaseSheetStyle, _
    ByVal hasOutputStyle As Boolean _
)
    Dim layerOrder As Variant
    Dim layerName As Variant
    Dim rowCount As Long
    Dim colCount As Long

    If ws Is Nothing Then Exit Sub
    If Not ex_SheetStylesXmlProvider.m_GetLayerOrder(hasOutputStyle, layerOrder, ThisWorkbook) Then Exit Sub
    If Not ex_SheetStylesXmlProvider.m_GetUsedRangeSize(ws, rowCount, colCount) Then Exit Sub

    For Each layerName In layerOrder
        Select Case CStr(layerName)
            Case ex_SheetStylesXmlProvider.LAYER_BASE
                ex_SheetStylesXmlProvider.m_ApplyBaseLayer ws, rowCount, colCount, baseStyle
            Case ex_SheetStylesXmlProvider.LAYER_OUTPUT
                mp_ApplyTimelineOutputStyle ws, headerRows, sectionRows, outputStyle
        End Select
    Next layerName
End Sub

Public Sub m_ApplyOutputPanelLayers( _
    ByVal ws As Worksheet, _
    ByRef outputStyle As t_OutputSheetStyle, _
    ByVal hasOutputStyle As Boolean, _
    ByVal viewStartRow As Long, _
    ByVal viewEndRow As Long, _
    ByVal viewColCount As Long _
)
    If ws Is Nothing Then Exit Sub
    If Not hasOutputStyle Then Exit Sub

    ex_OutputPanel.m_RenderForSheet ws, outputStyle
    ex_OutputPanel.m_ApplyFixedWidthViewZoneLayer ws, outputStyle, viewStartRow, viewEndRow, viewColCount
End Sub

Public Sub m_ApplyViewZoneWrapText( _
    ByVal ws As Worksheet, _
    ByVal viewStartRow As Long, _
    ByVal viewEndRow As Long, _
    ByVal viewColCount As Long, _
    Optional ByVal wrapEnabled As Boolean = True, _
    Optional ByVal excludedRows As Collection = Nothing, _
    Optional ByVal excludedRows2 As Collection = Nothing _
)
    Dim targetRange As Range
    Dim rowIndex As Long
    Dim rowRange As Range
    Dim excludedMap As Object
    Dim excludedMap2 As Object

    If ws Is Nothing Then Exit Sub
    If viewStartRow < 1 Then Exit Sub
    If viewEndRow < viewStartRow Then Exit Sub
    If viewColCount < 1 Then Exit Sub

    Set targetRange = ws.Range(ws.Cells(viewStartRow, 1), ws.Cells(viewEndRow, viewColCount))
    targetRange.WrapText = wrapEnabled

    If excludedRows Is Nothing Then Exit Sub

    Set excludedMap = mp_BuildRowsMap(excludedRows)
    Set excludedMap2 = mp_BuildRowsMap(excludedRows2)

    For rowIndex = viewStartRow To viewEndRow
        If mp_RowIsInMap(excludedMap, rowIndex) Then
            Set rowRange = ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, viewColCount))
            rowRange.WrapText = False
        ElseIf mp_RowIsInMap(excludedMap2, rowIndex) Then
            Set rowRange = ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, viewColCount))
            rowRange.WrapText = False
        End If
    Next rowIndex
End Sub

Public Sub m_ApplyTimelineDataRowsHeight( _
    ByVal ws As Worksheet, _
    ByVal viewStartRow As Long, _
    ByVal viewEndRow As Long, _
    ByVal viewColCount As Long, _
    ByVal headerRows As Collection, _
    ByVal sectionRows As Collection, _
    Optional ByVal dataRowHeight As Double = 32 _
)
    Dim rowIndex As Long
    Dim headerRowsMap As Object
    Dim sectionRowsMap As Object
    Dim lastDataRow As Long

    If ws Is Nothing Then Exit Sub
    If viewStartRow < 1 Then Exit Sub
    If viewEndRow < viewStartRow Then Exit Sub
    If viewColCount < 1 Then Exit Sub
    If dataRowHeight <= 0 Then Exit Sub

    Set headerRowsMap = mp_BuildRowsMap(headerRows)
    Set sectionRowsMap = mp_BuildRowsMap(sectionRows)

    lastDataRow = mp_GetLastUsedRow(ws)
    If lastDataRow <= 0 Then Exit Sub
    If lastDataRow < viewStartRow Then Exit Sub
    If lastDataRow > viewEndRow Then lastDataRow = viewEndRow

    For rowIndex = viewStartRow To lastDataRow
        If mp_RowIsInMap(headerRowsMap, rowIndex) Then GoTo ContinueRow
        If mp_RowIsInMap(sectionRowsMap, rowIndex) Then GoTo ContinueRow
        ws.Rows(rowIndex).RowHeight = dataRowHeight
ContinueRow:
    Next rowIndex
End Sub

Private Sub mp_ApplyOutputStyleToResult( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal rowCount As Long, _
    ByVal colCount As Long, _
    ByRef style As t_OutputSheetStyle _
)
    Dim targetRange As Range
    Dim headerRange As Range

    Set targetRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + rowCount - 1, colCount))
    Set headerRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, colCount))

    targetRange.Interior.Pattern = xlSolid
    targetRange.Interior.Color = style.ContentBackColor
    targetRange.Font.Color = style.ContentColor
    targetRange.Font.Name = style.FontName
    targetRange.Font.Size = style.FontSize
    targetRange.RowHeight = style.RowHeight
    targetRange.HorizontalAlignment = style.HorizontalAlignment
    targetRange.VerticalAlignment = style.VerticalAlignment

    headerRange.Interior.Pattern = xlSolid
    headerRange.Interior.Color = style.HeaderBackColor
    headerRange.Font.Color = style.HeaderColor
    headerRange.Font.Bold = style.HeaderBold
End Sub

Private Sub mp_ApplyTimelineOutputStyle(ByVal ws As Worksheet, ByVal headerRows As Collection, ByVal sectionRows As Collection, ByRef style As t_OutputSheetStyle)
    Dim usedRows As Long
    Dim usedCols As Long
    Dim usedRange As Range
    Dim rowId As Variant
    Dim rowIndex As Long
    Dim lastCol As Long
    Dim titleRange As Range
    Dim sectionFillRange As Range
    Dim sectionTitleCols As Long

    If ws Is Nothing Then Exit Sub
    If ws.UsedRange Is Nothing Then Exit Sub

    usedRows = ws.UsedRange.Rows.Count
    usedCols = ws.UsedRange.Columns.Count
    Set usedRange = ws.Range(ws.Cells(1, 1), ws.Cells(usedRows, usedCols))

    usedRange.Interior.Pattern = xlSolid
    usedRange.Interior.Color = style.ContentBackColor
    usedRange.Font.Name = style.FontName
    usedRange.Font.Size = style.FontSize
    usedRange.Font.Color = style.ContentColor
    usedRange.HorizontalAlignment = style.HorizontalAlignment
    usedRange.VerticalAlignment = style.VerticalAlignment
    ws.Rows("1:" & CStr(usedRows)).RowHeight = style.RowHeight

    ' Header/section rows should not wrap before AutoFit, otherwise column widths can be over-expanded.
    For Each rowId In headerRows
        rowIndex = CLng(rowId)
        lastCol = mp_GetLastUsedColumnInRow(ws, rowIndex)
        If lastCol > 0 Then
            ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, lastCol)).WrapText = False
        End If
    Next rowId

    For Each rowId In sectionRows
        rowIndex = CLng(rowId)
        lastCol = mp_GetLastUsedColumnInRow(ws, rowIndex)
        If lastCol > 0 Then
            ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, lastCol)).WrapText = False
        End If
    Next rowId

    usedRange.EntireColumn.AutoFit

    For Each rowId In sectionRows
        rowIndex = CLng(rowId)
        Set sectionFillRange = ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, usedCols))
        sectionTitleCols = style.SectionMergeColumns
        If sectionTitleCols < 1 Then sectionTitleCols = 1
        If sectionTitleCols > usedCols Then sectionTitleCols = usedCols
        Set titleRange = ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, sectionTitleCols))
        titleRange.UnMerge
        titleRange.Merge
        titleRange.HorizontalAlignment = style.HorizontalAlignment
        titleRange.VerticalAlignment = style.VerticalAlignment
        sectionFillRange.Interior.Pattern = xlSolid
        sectionFillRange.Interior.Color = style.SectionBackColor
        titleRange.Font.Bold = style.SectionBold
        titleRange.Font.Color = style.SectionColor
    Next rowId

    For Each rowId In headerRows
        rowIndex = CLng(rowId)
        lastCol = mp_GetLastUsedColumnInRow(ws, rowIndex)
        If lastCol > 0 Then
            Set titleRange = ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, lastCol))
            titleRange.Interior.Pattern = xlSolid
            titleRange.Interior.Color = style.HeaderBackColor
            titleRange.Font.Bold = style.HeaderBold
            titleRange.Font.Color = style.HeaderColor
        End If
    Next rowId
End Sub

Private Function mp_GetLastUsedColumnInRow(ByVal ws As Worksheet, ByVal rowIndex As Long) As Long
    If ws Is Nothing Then Exit Function
    If rowIndex <= 0 Then Exit Function

    mp_GetLastUsedColumnInRow = ws.Cells(rowIndex, ws.Columns.Count).End(xlToLeft).Column
    If mp_GetLastUsedColumnInRow = 1 Then
        If Len(Trim$(CStr(ws.Cells(rowIndex, 1).Value))) = 0 Then
            mp_GetLastUsedColumnInRow = 0
        End If
    End If
End Function

Private Function mp_RowIsListed(ByVal rowsCollection As Collection, ByVal rowIndex As Long) As Boolean
    Dim itemValue As Variant

    If rowsCollection Is Nothing Then Exit Function

    For Each itemValue In rowsCollection
        If CLng(itemValue) = rowIndex Then
            mp_RowIsListed = True
            Exit Function
        End If
    Next itemValue
End Function

Private Function mp_BuildRowsMap(ByVal rowsCollection As Collection) As Object
    Dim rowsMap As Object
    Dim itemValue As Variant

    Set rowsMap = CreateObject("Scripting.Dictionary")
    rowsMap.CompareMode = 0

    If rowsCollection Is Nothing Then
        Set mp_BuildRowsMap = rowsMap
        Exit Function
    End If

    For Each itemValue In rowsCollection
        rowsMap(CStr(CLng(itemValue))) = True
    Next itemValue

    Set mp_BuildRowsMap = rowsMap
End Function

Private Function mp_RowIsInMap(ByVal rowsMap As Object, ByVal rowIndex As Long) As Boolean
    If rowsMap Is Nothing Then Exit Function
    mp_RowIsInMap = rowsMap.Exists(CStr(rowIndex))
End Function

Private Function mp_GetLastUsedRow(ByVal ws As Worksheet) As Long
    Dim lastCell As Range

    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If lastCell Is Nothing Then Exit Function
    mp_GetLastUsedRow = lastCell.Row
End Function
