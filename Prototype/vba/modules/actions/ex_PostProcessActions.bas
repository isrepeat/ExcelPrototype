Attribute VB_Name = "ex_PostProcessActions"
Option Explicit

Public Sub m_HighlightRow( _
    ByVal rowRef As pp_ResultRow, _
    Optional ByVal colorHex As String = "#FFF2CC" _
)
    Dim colorValue As Long
    Dim rowRange As Range
    Dim ws As Worksheet
    Dim usedCols As Long

    If rowRef Is Nothing Then Exit Sub
    If Len(Trim$(colorHex)) = 0 Then colorHex = "#FFF2CC"
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    usedCols = mp_GetLastUsedColumn(ws)
    If usedCols <= 0 Then usedCols = 1

    If Not ex_XmlCore.m_TryParseColor(colorHex, colorValue) Then
        Err.Raise vbObjectError + 1650, "ex_PostProcessActions", "Invalid highlight color: " & colorHex
    End If

    Set rowRange = ws.Range(ws.Cells(rowRef.RowIndex, 1), ws.Cells(rowRef.RowIndex, usedCols))
    rowRange.Interior.Pattern = xlSolid
    rowRange.Interior.Color = colorValue
End Sub

Public Sub m_AddNote( _
    ByVal rowRef As pp_ResultRow, _
    ByVal noteText As String _
)
    Dim noteCell As Range
    Dim ws As Worksheet

    If rowRef Is Nothing Then Exit Sub
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    Set noteCell = ws.Cells(rowRef.RowIndex, 1)
    On Error Resume Next
    If Not noteCell.Comment Is Nothing Then noteCell.Comment.Delete
    On Error GoTo 0
    noteCell.AddComment noteText
End Sub

Public Sub m_AppendFooterText(ByVal footerText As String)
    Dim ws As Worksheet
    Dim startRow As Long

    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    startRow = mp_GetLastUsedRow(ws) + 2
    If startRow < 1 Then startRow = 1
    ws.Cells(startRow, 1).Value = footerText
End Sub

Private Function mp_GetLastUsedRow(ByVal ws As Worksheet) As Long
    On Error GoTo ExitFn
    mp_GetLastUsedRow = ws.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
ExitFn:
End Function

Private Function mp_GetLastUsedColumn(ByVal ws As Worksheet) As Long
    On Error GoTo ExitFn
    mp_GetLastUsedColumn = ws.Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
ExitFn:
End Function
