Attribute VB_Name = "ex_SheetHelpers"
Option Explicit

Public Function m_GetStandardRowHeight( _
    ByVal ws As Worksheet, _
    Optional ByVal fallbackHeight As Double = 15# _
) As Double
    If ws Is Nothing Then
        m_GetStandardRowHeight = fallbackHeight
        Exit Function
    End If

    m_GetStandardRowHeight = ws.StandardHeight
    If m_GetStandardRowHeight <= 0 Then m_GetStandardRowHeight = fallbackHeight
End Function

Public Function m_RoundUpMeasuredHeight( _
    ByVal measuredHeight As Double, _
    Optional ByVal roundPad As Double = 0# _
) As Double
    Dim roundedHeight As Double

    If measuredHeight <= 0 Then Exit Function
    roundedHeight = Int(measuredHeight + 0.999)
    m_RoundUpMeasuredHeight = roundedHeight + roundPad
End Function

Public Function m_MeasureTextHeight( _
    ByVal ws As Worksheet, _
    ByVal targetRange As Range, _
    ByVal messageText As String, _
    Optional ByVal sideMargin As Double = 0#, _
    Optional ByVal verticalMargin As Double = 0#, _
    Optional ByVal forceFontSize As Double = 0#, _
    Optional ByVal forceFontName As String = vbNullString _
) As Double
    Dim textBoxShape As Object

    On Error GoTo EH
    If ws Is Nothing Then Exit Function
    If targetRange Is Nothing Then Exit Function
    If Len(messageText) = 0 Then Exit Function

    Set textBoxShape = ws.Shapes.AddTextbox(1, targetRange.Left, targetRange.Top, targetRange.Width, 8)
    textBoxShape.Line.Visible = 0
    textBoxShape.Fill.Visible = 0
    textBoxShape.TextFrame2.MarginLeft = sideMargin
    textBoxShape.TextFrame2.MarginRight = sideMargin
    textBoxShape.TextFrame2.MarginTop = verticalMargin
    textBoxShape.TextFrame2.MarginBottom = verticalMargin
    textBoxShape.TextFrame2.WordWrap = -1
    textBoxShape.TextFrame2.AutoSize = 1
    textBoxShape.TextFrame2.TextRange.Text = messageText

    If forceFontSize > 0 Then
        textBoxShape.TextFrame2.TextRange.Font.Size = forceFontSize
    Else
        textBoxShape.TextFrame2.TextRange.Font.Size = targetRange.Font.Size
    End If
    If Len(forceFontName) > 0 Then
        textBoxShape.TextFrame2.TextRange.Font.Name = forceFontName
    Else
        textBoxShape.TextFrame2.TextRange.Font.Name = CStr(targetRange.Font.Name)
    End If

    m_MeasureTextHeight = textBoxShape.Height

Cleanup:
    On Error Resume Next
    If Not textBoxShape Is Nothing Then textBoxShape.Delete
    On Error GoTo 0
    Exit Function

EH:
    m_MeasureTextHeight = 0
    Resume Cleanup
End Function

Public Sub m_ApplySingleRowTextAutoHeight( _
    ByVal ws As Worksheet, _
    ByVal targetRange As Range, _
    ByVal messageText As String, _
    Optional ByVal baseRowHeight As Double = 0#, _
    Optional ByVal minRowHeight As Double = 0#, _
    Optional ByVal autoHeightEnabled As Boolean = True, _
    Optional ByVal wrapTextEnabled As Boolean = True, _
    Optional ByVal autoHeightMarginTop As Double = 0#, _
    Optional ByVal autoHeightMarginBottom As Double = 0#, _
    Optional ByVal measureSideMargin As Double = 0#, _
    Optional ByVal measureVerticalMargin As Double = 0#, _
    Optional ByVal measureExtraHeight As Double = 0#, _
    Optional ByVal roundPad As Double = 0#, _
    Optional ByVal forceFontSize As Double = 0#, _
    Optional ByVal forceFontName As String = vbNullString _
)
    Dim targetRow As Long
    Dim standardRowHeight As Double
    Dim measuredHeight As Double

    If ws Is Nothing Then Exit Sub
    If targetRange Is Nothing Then Exit Sub
    targetRow = targetRange.Row
    If targetRow < 1 Or targetRow > ws.Rows.Count Then Exit Sub

    standardRowHeight = m_GetStandardRowHeight(ws, 15#)

    If baseRowHeight > 0 Then
        ws.Rows(targetRow).RowHeight = baseRowHeight
    Else
        ws.Rows(targetRow).RowHeight = standardRowHeight
    End If

    If Not autoHeightEnabled Or Not wrapTextEnabled Then
        If minRowHeight > 0 Then
            If ws.Rows(targetRow).RowHeight < minRowHeight Then
                ws.Rows(targetRow).RowHeight = minRowHeight
            End If
        End If
        Exit Sub
    End If

    If Len(Trim$(messageText)) = 0 Then
        If minRowHeight > 0 Then
            If ws.Rows(targetRow).RowHeight < minRowHeight Then
                ws.Rows(targetRow).RowHeight = minRowHeight
            End If
        End If
        Exit Sub
    End If

    On Error Resume Next
    targetRange.WrapText = True
    On Error GoTo 0

    measuredHeight = m_MeasureTextHeight( _
        ws, _
        targetRange, _
        messageText, _
        measureSideMargin, _
        measureVerticalMargin, _
        forceFontSize, _
        forceFontName _
    )
    measuredHeight = measuredHeight + autoHeightMarginTop + autoHeightMarginBottom + measureExtraHeight
    If measuredHeight < minRowHeight Then measuredHeight = minRowHeight

    If measuredHeight > 0 Then
        ws.Rows(targetRow).RowHeight = m_RoundUpMeasuredHeight(measuredHeight, roundPad)
    ElseIf minRowHeight > 0 Then
        ws.Rows(targetRow).RowHeight = minRowHeight
    End If
End Sub
