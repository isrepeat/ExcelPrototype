Attribute VB_Name = "ex_SheetTheme"
Option Explicit

' =============================================================================
' ex_SheetTheme
' =============================================================================
' Purpose:
'   Apply unified dark theme to any worksheet:
'   - dark background for visible area
'   - light font
'   - visible grid borders ("All Borders" style)
'   - optional status highlighting (Added / Changed / Removed)
'
' This module contains NO data logic.
' =============================================================================

' -----------------------------------------------------------------------------
' Public API
' -----------------------------------------------------------------------------

' Apply full dark theme to a worksheet.
' If hasStatusColumn = True, Status-based coloring will be applied.
Public Sub ApplyDarkThemeToSheet( _
    ByVal ws As Worksheet, _
    Optional ByVal hasStatusColumn As Boolean = False _
)
    Dim usedRange As Range
    Dim rowCount As Long
    Dim colCount As Long
    
    If ws.UsedRange Is Nothing Then
        Exit Sub
    End If
    
    Set usedRange = ws.UsedRange
    rowCount = usedRange.Rows.Count
    colCount = usedRange.Columns.Count
    
    ApplyDarkBackground ws, rowCount, colCount
    ApplyGridBorders ws, rowCount, colCount
    
    If hasStatusColumn Then
        ApplyStatusHighlight ws, rowCount, colCount
    End If
End Sub

' -----------------------------------------------------------------------------
' Dark background for visible area + margin
' -----------------------------------------------------------------------------
Private Sub ApplyDarkBackground( _
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

' -----------------------------------------------------------------------------
' Draw full grid borders (like "All Borders")
' -----------------------------------------------------------------------------
Private Sub ApplyGridBorders( _
    ByVal ws As Worksheet, _
    ByVal rowCount As Long, _
    ByVal colCount As Long _
)
    Dim targetRange As Range
    
    Set targetRange = ws.Range(ws.Cells(1, 1), ws.Cells(rowCount + 200, colCount + 50))
    
    With targetRange
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
        .Borders.Color = RGB(80, 80, 80)
        .Borders.Weight = xlThin
    End With
End Sub

' -----------------------------------------------------------------------------
' Status-based row coloring (ONLY for Result sheet)
' -----------------------------------------------------------------------------
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
                rowRange.Interior.Color = RGB(46, 125, 50)
            Case "changed"
                rowRange.Interior.Color = RGB(123, 31, 162)
            Case "removed"
                rowRange.Interior.Color = RGB(183, 28, 28)
            Case Else
                rowRange.Interior.Color = RGB(30, 30, 30)
        End Select
        
        rowRange.Font.Color = RGB(235, 235, 235)
    Next r
End Sub

' -----------------------------------------------------------------------------
' Utilities
' -----------------------------------------------------------------------------
Private Function FindColumnIndex( _
    ByVal ws As Worksheet, _
    ByVal colCount As Long, _
    ByVal headerName As String _
) As Long
    Dim c As Long
    
    For c = 1 To colCount
        If StrComp(CStr(ws.Cells(1, c).Value), headerName, vbTextCompare) = 0 Then
            FindColumnIndex = c
            Exit Function
        End If
    Next c
    
    FindColumnIndex = 0
End Function
