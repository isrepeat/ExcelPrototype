Attribute VB_Name = "ex_ResultWriter"
Option Explicit

' =============================================================================
' ex_ResultWriter
' =============================================================================
' Purpose:
'   Render compare result into worksheet "Result" with a dark UI theme.
'
' Responsibilities:
'   - Create / get "Result" worksheet
'   - Clear old content and write a 2D Variant array to cells
'   - Apply base formatting (header row, filters, autofit, freeze panes)
'   - Apply dark background to the whole visible area (plus extra margin)
'   - Highlight ONLY 3 statuses by row color:
'         Added   -> green
'         Changed -> purple
'         Removed -> red
'     Any other status (OK / Error) stays with default dark background.
'   - Draw full grid borders (like Excel menu "All Borders") on dark background
'
' Notes about Excel limitations:
'   Excel has no "sheet background color" like a UI canvas.
'   We simulate it by filling a big cell range and restricting scroll area.
' =============================================================================

Public Sub WriteTableToResultSheet(ByVal tableData As Variant)
    Dim ws As Worksheet
    Dim rowCount As Long
    Dim colCount As Long
    Dim targetRange As Range
    
    ' Get or create Result sheet
    Set ws = GetOrCreateWorksheet("Result")
    
    ' Clear previous content and reset scroll area
    ws.Cells.Clear
    ws.ScrollArea = ""
    
    ' Determine table size from 2D array
    rowCount = UBound(tableData, 1)
    colCount = UBound(tableData, 2)
    
    ' Write data in one operation (fast)
    Set targetRange = ws.Range(ws.Cells(1, 1), ws.Cells(rowCount, colCount))
    targetRange.Value = tableData
    
    ' Base table formatting (fonts, header, filters, freeze panes)
    FormatAsTable _
        ws, _
        rowCount, _
        colCount
    
    ' Apply unified dark theme + status highlighting
    ex_SheetTheme.ApplyDarkThemeToSheet _
        ws, _
        True
End Sub


Private Function GetOrCreateWorksheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim fullName As String
    
    ' Add g_ prefix to sheet names
    fullName = "g_" & sheetName
    
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, fullName, vbTextCompare) = 0 Then
            Set GetOrCreateWorksheet = ws
            Exit Function
        End If
    Next ws
    
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = fullName
    
    Set GetOrCreateWorksheet = ws
End Function

Private Sub FormatAsTable( _
    ByVal ws As Worksheet, _
    ByVal rowCount As Long, _
    ByVal colCount As Long _
)
    Dim headerRange As Range
    Dim allRange As Range
    
    Set headerRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount))
    Set allRange = ws.Range(ws.Cells(1, 1), ws.Cells(rowCount, colCount))
    
    allRange.Font.Name = "Segoe UI"
    allRange.Font.Size = 10
    
    headerRange.Font.Bold = True
    
    allRange.HorizontalAlignment = xlCenter
    headerRange.HorizontalAlignment = xlCenter
    
    allRange.EntireColumn.AutoFit
    allRange.AutoFilter
    
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub ApplyDarkSheetBackground( _
    ByVal ws As Worksheet, _
    ByVal rowCount As Long, _
    ByVal colCount As Long _
)
    Dim visibleRange As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim bgRange As Range
    
    ws.Activate
    Set visibleRange = ActiveWindow.VisibleRange
    
    lastRow = visibleRange.Row + visibleRange.Rows.Count - 1 + 200
    lastCol = visibleRange.Column + visibleRange.Columns.Count - 1 + 30
    
    If lastRow < rowCount + 200 Then
        lastRow = rowCount + 200
    End If
    
    If lastCol < colCount + 10 Then
        lastCol = colCount + 10
    End If
    
    Set bgRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    bgRange.Interior.Pattern = xlSolid
    bgRange.Interior.Color = RGB(30, 30, 30)
    bgRange.Font.Color = RGB(235, 235, 235)
    
    ActiveWindow.DisplayGridlines = False
    ws.ScrollArea = bgRange.Address
End Sub

Private Sub ApplyStatusHighlight( _
    ByVal ws As Worksheet, _
    ByVal rowCount As Long, _
    ByVal colCount As Long _
)
    Dim statusCol As Long
    Dim r As Long
    Dim statusValue As String
    Dim rowRange As Range
    
    statusCol = FindColumnIndex(ws, colCount, "Status")
    If statusCol = 0 Then
        Exit Sub
    End If
    
    For r = 2 To rowCount
        statusValue = CStr(ws.Cells(r, statusCol).Value)
        Set rowRange = ws.Range(ws.Cells(r, 1), ws.Cells(r, colCount))
        
        Select Case LCase$(statusValue)
            Case "added"
                rowRange.Interior.Pattern = xlSolid
                rowRange.Interior.Color = RGB(46, 125, 50)
            Case "changed"
                rowRange.Interior.Pattern = xlSolid
                rowRange.Interior.Color = RGB(123, 31, 162)
            Case "removed"
                rowRange.Interior.Pattern = xlSolid
                rowRange.Interior.Color = RGB(183, 28, 28)
            Case Else
                rowRange.Interior.Pattern = xlSolid
                rowRange.Interior.Color = RGB(30, 30, 30)
        End Select
        
        rowRange.Font.Color = RGB(235, 235, 235)
    Next r
End Sub

Private Sub ApplyAllBordersToRange(ByVal targetRange As Range)
    With targetRange
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
        
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideHorizontal).Weight = xlThin
        
        .Borders.Color = RGB(80, 80, 80)
    End With
End Sub

Private Function FindColumnIndex( _
    ByVal ws As Worksheet, _
    ByVal colCount As Long, _
    ByVal headerName As String _
) As Long
    Dim c As Long
    Dim v As String
    
    For c = 1 To colCount
        v = CStr(ws.Cells(1, c).Value)
        If StrComp(v, headerName, vbTextCompare) = 0 Then
            FindColumnIndex = c
            Exit Function
        End If
    Next c
    
    FindColumnIndex = 0
End Function
