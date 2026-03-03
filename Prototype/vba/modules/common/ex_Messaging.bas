Attribute VB_Name = "ex_Messaging"
Option Explicit

Private Const DEFAULT_DARK_ROWS As Long = 200
Private Const DEFAULT_DARK_COLS As Long = 52

Private g_StatusClearTime As Date
Private g_StatusClearScheduled As Boolean
Private g_StatusClearProcedureName As String

' =============================================================================
' Status bar notification
' =============================================================================

Public Sub m_ShowNotice(ByVal msg As String, Optional ByVal seconds As Double = 2)
    If seconds <= 0 Then seconds = 2

    Application.StatusBar = msg

    On Error Resume Next
    If g_StatusClearScheduled Then
        Application.OnTime EarliestTime:=g_StatusClearTime, Procedure:=mp_GetStatusClearProcedureName(), Schedule:=False
        g_StatusClearScheduled = False
    End If

    g_StatusClearTime = Now + (seconds / 86400#)
    Application.OnTime EarliestTime:=g_StatusClearTime, Procedure:=mp_GetStatusClearProcedureName(), Schedule:=True
    g_StatusClearScheduled = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub m_ClearStatusBar()
    ' Очищает статус бар
    g_StatusClearScheduled = False
    Application.StatusBar = False
End Sub

Private Function mp_GetStatusClearProcedureName() As String
    If Len(g_StatusClearProcedureName) = 0 Then
        g_StatusClearProcedureName = "'" & ThisWorkbook.Name & "'!ex_Messaging.m_ClearStatusBar"
    End If
    mp_GetStatusClearProcedureName = g_StatusClearProcedureName
End Function

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
    Dim bannerStyle As ex_SheetStylesXmlProvider.t_ErrorBannerStyle
    Dim hasBannerStyle As Boolean
    Dim rowCount As Long
    Dim messageText As String
    Dim rowOffset As Long

    If ws Is Nothing Then Exit Sub

    messageText = Trim$(errDescription)
    If Len(messageText) = 0 Then
        messageText = "Unknown error."
    End If

    hasBannerStyle = ex_SheetStylesXmlProvider.m_GetErrorBannerStyle(bannerStyle, ThisWorkbook)
    If Len(Trim$(bannerRangeAddress)) = 0 Then
        bannerRangeAddress = ex_SheetStylesXmlProvider.m_GetOutputErrorBannerRangeAddress(ThisWorkbook)
    End If

    Set bannerRange = ws.Range(bannerRangeAddress)
    rowCount = bannerRange.Rows.Count
    If hasBannerStyle Then
        If bannerStyle.Rows > rowCount Then rowCount = bannerStyle.Rows
    End If
    If rowCount < 4 Then rowCount = 4

    bannerRange.ClearContents
    bannerRange.UnMerge
    For rowOffset = 0 To rowCount - 1
        ws.Range(ws.Cells(bannerRange.Row + rowOffset, bannerRange.Column), ws.Cells(bannerRange.Row + rowOffset, bannerRange.Column + bannerRange.Columns.Count - 1)).Merge
    Next rowOffset

    ws.Cells(bannerRange.Row, bannerRange.Column).Value = titleText
    ws.Cells(bannerRange.Row + 1, bannerRange.Column).Value = messageText
    ws.Cells(bannerRange.Row + 2, bannerRange.Column).Value = "Source: " & IIf(Len(Trim$(errSource)) > 0, errSource, "n/a")
    ws.Cells(bannerRange.Row + 3, bannerRange.Column).Value = "Code: " & CStr(errNumber)

    With ws.Range(ws.Cells(bannerRange.Row, bannerRange.Column), ws.Cells(bannerRange.Row + rowCount - 1, bannerRange.Column + bannerRange.Columns.Count - 1))
        .WrapText = IIf(hasBannerStyle, bannerStyle.WrapText, True)
        .VerticalAlignment = IIf(hasBannerStyle, bannerStyle.VerticalAlignment, xlCenter)
        .HorizontalAlignment = IIf(hasBannerStyle, bannerStyle.HorizontalAlignment, xlLeft)
        .Interior.Pattern = xlSolid
        .Interior.Color = IIf(hasBannerStyle, bannerStyle.BackColor, RGB(192, 0, 0))
        .Font.Color = IIf(hasBannerStyle, bannerStyle.FontColor, RGB(255, 255, 255))
        .Font.Bold = False

        If hasBannerStyle And bannerStyle.ShowGrid Then
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders.Color = bannerStyle.GridColor
            .Borders.Weight = xlThin
        ElseIf hasBannerStyle Then
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlNone
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
        End If
    End With

    ws.Range(ws.Cells(bannerRange.Row, bannerRange.Column), ws.Cells(bannerRange.Row, bannerRange.Column + bannerRange.Columns.Count - 1)).Font.Bold = IIf(hasBannerStyle, bannerStyle.TitleBold, True)
    mp_ApplyBannerRowHeights ws, bannerRange, rowCount, IIf(hasBannerStyle, bannerStyle.RowHeight, 24), bannerRange.Row + 1, messageText, IIf(hasBannerStyle, bannerStyle.WrapText, True)
End Sub

Public Sub m_RenderWarningBanner( _
    ByVal ws As Worksheet, _
    ByVal warningText As String, _
    Optional ByVal titleText As String = "WARNING", _
    Optional ByVal bannerRangeAddress As String = "A1:H3" _
)
    Dim bannerRange As Range
    Dim bannerStyle As ex_SheetStylesXmlProvider.t_ErrorBannerStyle
    Dim hasBannerStyle As Boolean
    Dim rowCount As Long
    Dim messageText As String
    Dim rowOffset As Long

    If ws Is Nothing Then Exit Sub

    messageText = Trim$(warningText)
    If Len(messageText) = 0 Then
        messageText = "Action required."
    End If

    hasBannerStyle = ex_SheetStylesXmlProvider.m_GetWarningBannerStyle(bannerStyle, ThisWorkbook)
    If Len(Trim$(bannerRangeAddress)) = 0 Then
        bannerRangeAddress = ex_SheetStylesXmlProvider.m_GetOutputWarningBannerRangeAddress(ThisWorkbook)
    End If

    Set bannerRange = ws.Range(bannerRangeAddress)
    rowCount = bannerRange.Rows.Count
    If hasBannerStyle Then
        If bannerStyle.Rows > rowCount Then rowCount = bannerStyle.Rows
    End If
    If rowCount < 2 Then rowCount = 2

    bannerRange.ClearContents
    bannerRange.UnMerge
    For rowOffset = 0 To rowCount - 1
        ws.Range(ws.Cells(bannerRange.Row + rowOffset, bannerRange.Column), ws.Cells(bannerRange.Row + rowOffset, bannerRange.Column + bannerRange.Columns.Count - 1)).Merge
    Next rowOffset

    ws.Cells(bannerRange.Row, bannerRange.Column).Value = titleText
    ws.Cells(bannerRange.Row + 1, bannerRange.Column).Value = messageText

    With ws.Range(ws.Cells(bannerRange.Row, bannerRange.Column), ws.Cells(bannerRange.Row + rowCount - 1, bannerRange.Column + bannerRange.Columns.Count - 1))
        .WrapText = IIf(hasBannerStyle, bannerStyle.WrapText, True)
        .VerticalAlignment = IIf(hasBannerStyle, bannerStyle.VerticalAlignment, xlCenter)
        .HorizontalAlignment = IIf(hasBannerStyle, bannerStyle.HorizontalAlignment, xlLeft)
        .Interior.Pattern = xlSolid
        .Interior.Color = IIf(hasBannerStyle, bannerStyle.BackColor, RGB(76, 63, 16))
        .Font.Color = IIf(hasBannerStyle, bannerStyle.FontColor, RGB(255, 229, 153))
        .Font.Bold = False

        If hasBannerStyle And bannerStyle.ShowGrid Then
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders.Color = bannerStyle.GridColor
            .Borders.Weight = xlThin
        ElseIf hasBannerStyle Then
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlNone
            .Borders(xlEdgeRight).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
        End If
    End With

    ws.Range(ws.Cells(bannerRange.Row, bannerRange.Column), ws.Cells(bannerRange.Row, bannerRange.Column + bannerRange.Columns.Count - 1)).Font.Bold = IIf(hasBannerStyle, bannerStyle.TitleBold, True)
    mp_ApplyBannerRowHeights ws, bannerRange, rowCount, IIf(hasBannerStyle, bannerStyle.RowHeight, 24), bannerRange.Row + 1, messageText, IIf(hasBannerStyle, bannerStyle.WrapText, True)
End Sub

Private Sub mp_ApplyBannerRowHeights( _
    ByVal ws As Worksheet, _
    ByVal bannerRange As Range, _
    ByVal rowCount As Long, _
    ByVal baseRowHeight As Double, _
    ByVal messageRowIndex As Long, _
    ByVal messageText As String, _
    ByVal wrapTextEnabled As Boolean _
)
    Dim measuredHeight As Double
    Dim messageRange As Range
    Dim rowStart As Long
    Dim rowEnd As Long

    If ws Is Nothing Then Exit Sub
    If bannerRange Is Nothing Then Exit Sub
    If rowCount < 1 Then Exit Sub
    If baseRowHeight <= 0 Then baseRowHeight = 24

    rowStart = bannerRange.Row
    rowEnd = bannerRange.Row + rowCount - 1
    ws.Rows(CStr(rowStart) & ":" & CStr(rowEnd)).RowHeight = baseRowHeight

    If Not wrapTextEnabled Then Exit Sub
    If Len(Trim$(messageText)) = 0 Then Exit Sub
    If messageRowIndex < rowStart Or messageRowIndex > rowEnd Then Exit Sub

    Set messageRange = ws.Range( _
        ws.Cells(messageRowIndex, bannerRange.Column), _
        ws.Cells(messageRowIndex, bannerRange.Column + bannerRange.Columns.Count - 1) _
    )
    measuredHeight = mp_MeasureBannerTextHeight(ws, messageRange, messageText)
    If measuredHeight > baseRowHeight Then
        ws.Rows(messageRowIndex).RowHeight = measuredHeight
    End If
End Sub

Private Function mp_MeasureBannerTextHeight( _
    ByVal ws As Worksheet, _
    ByVal targetRange As Range, _
    ByVal messageText As String _
) As Double
    Dim textBoxShape As Object

    On Error GoTo EH
    If ws Is Nothing Then Exit Function
    If targetRange Is Nothing Then Exit Function
    If Len(messageText) = 0 Then Exit Function

    Set textBoxShape = ws.Shapes.AddTextbox(1, targetRange.Left, targetRange.Top, targetRange.Width, 8)
    textBoxShape.Line.Visible = 0
    textBoxShape.Fill.Visible = 0
    textBoxShape.TextFrame2.MarginLeft = 0
    textBoxShape.TextFrame2.MarginRight = 0
    textBoxShape.TextFrame2.MarginTop = 0
    textBoxShape.TextFrame2.MarginBottom = 0
    textBoxShape.TextFrame2.WordWrap = -1
    textBoxShape.TextFrame2.AutoSize = 1
    textBoxShape.TextFrame2.TextRange.Text = messageText
    textBoxShape.TextFrame2.TextRange.Font.Size = targetRange.Font.Size
    textBoxShape.TextFrame2.TextRange.Font.Name = CStr(targetRange.Font.Name)

    mp_MeasureBannerTextHeight = textBoxShape.Height + 2

Cleanup:
    On Error Resume Next
    If Not textBoxShape Is Nothing Then textBoxShape.Delete
    On Error GoTo 0
    Exit Function

EH:
    mp_MeasureBannerTextHeight = 0
    Resume Cleanup
End Function
