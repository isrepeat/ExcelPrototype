Attribute VB_Name = "ex_Messaging"
Option Explicit

Private Const DEFAULT_DARK_ROWS As Long = 200
Private Const DEFAULT_DARK_COLS As Long = 52

' =============================================================================
' Status bar notification
' =============================================================================

Public Sub m_ShowNotice(ByVal msg As String, Optional ByVal seconds As Double = 2)
    Application.StatusBar = msg
    Application.OnTime Now + TimeSerial(0, 0, seconds), "ex_Messaging.m_ClearStatusBar"
End Sub

Public Sub m_ClearStatusBar()
    ' Очищает статус бар
    Application.StatusBar = False
End Sub

Public Sub m_ApplyDarkSheetBase( _
    ByVal ws As Worksheet, _
    Optional ByVal rowCount As Long = DEFAULT_DARK_ROWS, _
    Optional ByVal colCount As Long = DEFAULT_DARK_COLS _
)
    Dim seedRange As Range

    If ws Is Nothing Then Exit Sub
    If rowCount < 1 Then rowCount = DEFAULT_DARK_ROWS
    If colCount < 1 Then colCount = DEFAULT_DARK_COLS

    ws.Activate

    Set seedRange = ws.Range(ws.Cells(1, 1), ws.Cells(rowCount, colCount))
    seedRange.Interior.Pattern = xlSolid
    seedRange.Interior.Color = RGB(32, 32, 32)
    seedRange.Font.Color = RGB(235, 235, 235)
    seedRange.Borders.LineStyle = xlContinuous
    seedRange.Borders.Color = RGB(62, 62, 62)
    seedRange.Borders.Weight = xlThin
    ActiveWindow.DisplayGridlines = False
End Sub

Public Sub m_RenderErrorBanner( _
    ByVal ws As Worksheet, _
    ByVal errDescription As String, _
    Optional ByVal errSource As String = vbNullString, _
    Optional ByVal errNumber As Long = 0, _
    Optional ByVal titleText As String = "ERROR: Operation failed", _
    Optional ByVal bannerRangeAddress As String = "A1:H4" _
)
    Dim bannerRange As Range
    Dim rowCount As Long
    Dim messageText As String

    If ws Is Nothing Then Exit Sub

    messageText = Trim$(errDescription)
    If Len(messageText) = 0 Then
        messageText = "Unknown error."
    End If

    Set bannerRange = ws.Range(bannerRangeAddress)
    rowCount = bannerRange.Rows.Count
    If rowCount < 4 Then rowCount = 4

    bannerRange.ClearContents
    bannerRange.UnMerge
    ws.Range(ws.Cells(bannerRange.Row, bannerRange.Column), ws.Cells(bannerRange.Row, bannerRange.Column + bannerRange.Columns.Count - 1)).Merge
    ws.Range(ws.Cells(bannerRange.Row + 1, bannerRange.Column), ws.Cells(bannerRange.Row + 1, bannerRange.Column + bannerRange.Columns.Count - 1)).Merge
    ws.Range(ws.Cells(bannerRange.Row + 2, bannerRange.Column), ws.Cells(bannerRange.Row + 2, bannerRange.Column + bannerRange.Columns.Count - 1)).Merge
    ws.Range(ws.Cells(bannerRange.Row + 3, bannerRange.Column), ws.Cells(bannerRange.Row + 3, bannerRange.Column + bannerRange.Columns.Count - 1)).Merge

    ws.Cells(bannerRange.Row, bannerRange.Column).Value = titleText
    ws.Cells(bannerRange.Row + 1, bannerRange.Column).Value = messageText
    ws.Cells(bannerRange.Row + 2, bannerRange.Column).Value = "Source: " & IIf(Len(Trim$(errSource)) > 0, errSource, "n/a")
    ws.Cells(bannerRange.Row + 3, bannerRange.Column).Value = "Code: " & CStr(errNumber)

    With ws.Range(ws.Cells(bannerRange.Row, bannerRange.Column), ws.Cells(bannerRange.Row + rowCount - 1, bannerRange.Column + bannerRange.Columns.Count - 1))
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(192, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = False
    End With

    ws.Range(ws.Cells(bannerRange.Row, bannerRange.Column), ws.Cells(bannerRange.Row, bannerRange.Column + bannerRange.Columns.Count - 1)).Font.Bold = True
    ws.Rows(CStr(bannerRange.Row) & ":" & CStr(bannerRange.Row + rowCount - 1)).RowHeight = 24
End Sub
