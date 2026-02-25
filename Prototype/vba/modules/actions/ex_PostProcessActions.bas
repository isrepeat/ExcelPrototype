Attribute VB_Name = "ex_PostProcessActions"
Option Explicit

Private Const PROFILES_NS As String = "urn:excelprototype:profiles"
Private Const SHEET_STYLES_REL_PATH As String = "config\SheetStyles.xml"
Private Const POST_PROCESS_FOOTER_STYLE_LABEL As String = "post process footer style"

Private Type t_PostProcessFooterStyle
    Columns As Long
    Overflow As String
    BackColor As Long
    FontColor As Long
    FontSize As Double
    RowHeight As Double
    AutoHeight As Boolean
End Type

Public Sub m_HighlightRow( _
    ByVal rowRef As obj_ResultRow, _
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
    ByVal rowRef As obj_ResultRow, _
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

Public Sub m_AppendPostProcessFooterText(ByVal postProcessFooterText As String)
    Dim ws As Worksheet
    Dim startRow As Long
    Dim endCol As Long
    Dim postProcessFooterStyle As t_PostProcessFooterStyle
    Dim postProcessFooterRange As Range

    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    If Not mp_TryLoadPostProcessFooterStyle(postProcessFooterStyle) Then
        Err.Raise vbObjectError + 1651, "ex_PostProcessActions", "Unable to apply postProcessFooter text: invalid '/SheetStyles/postProcessFooterStyle'."
    End If

    startRow = mp_GetLastUsedRow(ws) + 2
    If startRow < 1 Then startRow = 1

    endCol = postProcessFooterStyle.Columns
    If endCol < 1 Then endCol = 1
    If endCol > ws.Columns.Count Then endCol = ws.Columns.Count

    Set postProcessFooterRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, endCol))
    If postProcessFooterRange.MergeCells Then postProcessFooterRange.UnMerge
    If endCol > 1 Then postProcessFooterRange.Merge

    postProcessFooterRange.Value = postProcessFooterText
    postProcessFooterRange.Interior.Pattern = xlSolid
    postProcessFooterRange.Interior.Color = postProcessFooterStyle.BackColor
    postProcessFooterRange.Font.Color = postProcessFooterStyle.FontColor
    postProcessFooterRange.Font.Size = postProcessFooterStyle.FontSize
    postProcessFooterRange.HorizontalAlignment = xlLeft
    postProcessFooterRange.VerticalAlignment = xlCenter

    Select Case postProcessFooterStyle.Overflow
        Case "wrap"
            postProcessFooterRange.WrapText = True
            postProcessFooterRange.ShrinkToFit = False
        Case "shrink"
            postProcessFooterRange.WrapText = False
            postProcessFooterRange.ShrinkToFit = True
        Case Else
            postProcessFooterRange.WrapText = False
            postProcessFooterRange.ShrinkToFit = False
    End Select

    mp_ApplyPostProcessFooterRowHeight ws, postProcessFooterRange, postProcessFooterText, postProcessFooterStyle
End Sub

Private Function mp_TryLoadPostProcessFooterStyle(ByRef outStyle As t_PostProcessFooterStyle) As Boolean
    Dim doc As Object
    Dim node As Object
    Dim overflowText As String

    Set doc = ex_XmlCore.m_LoadDomByRelativePath( _
        ThisWorkbook, _
        SHEET_STYLES_REL_PATH, _
        PROFILES_NS, _
        "Missing SheetStyles file: ", _
        "Failed to parse SheetStyles file: " _
    )
    If doc Is Nothing Then Exit Function

    Set node = doc.selectSingleNode("/p:SheetStyles/p:postProcessFooterStyle")
    If node Is Nothing Then
        MsgBox "SheetStyles must contain '/SheetStyles/postProcessFooterStyle'.", vbExclamation
        Exit Function
    End If

    If Not ex_XmlCore.m_ReadRequiredAttrLong(node, "columns", outStyle.Columns, "postProcessFooterStyle@columns", POST_PROCESS_FOOTER_STYLE_LABEL) Then Exit Function
    overflowText = LCase$(Trim$(ex_XmlCore.m_ReadRequiredAttrText(node, "overflow", "postProcessFooterStyle@overflow", POST_PROCESS_FOOTER_STYLE_LABEL)))
    If Len(overflowText) = 0 Then Exit Function
    Select Case overflowText
        Case "wrap", "clip", "shrink"
            outStyle.Overflow = overflowText
        Case Else
            MsgBox "Invalid value for postProcessFooter style attribute 'postProcessFooterStyle@overflow': expected wrap, clip, or shrink.", vbExclamation
            Exit Function
    End Select
    If Not ex_XmlCore.m_ReadRequiredAttrHexColor(node, "backColor", outStyle.BackColor, "postProcessFooterStyle@backColor", POST_PROCESS_FOOTER_STYLE_LABEL) Then Exit Function
    If Not ex_XmlCore.m_ReadRequiredAttrHexColor(node, "fontColor", outStyle.FontColor, "postProcessFooterStyle@fontColor", POST_PROCESS_FOOTER_STYLE_LABEL) Then Exit Function
    If Not ex_XmlCore.m_ReadRequiredAttrDouble(node, "fontSize", outStyle.FontSize, "postProcessFooterStyle@fontSize", POST_PROCESS_FOOTER_STYLE_LABEL) Then Exit Function
    If Not ex_XmlCore.m_ReadRequiredAttrDouble(node, "rowHeight", outStyle.RowHeight, "postProcessFooterStyle@rowHeight", POST_PROCESS_FOOTER_STYLE_LABEL) Then Exit Function
    If Not ex_XmlCore.m_ReadRequiredAttrBoolean(node, "autoHeight", outStyle.AutoHeight, "postProcessFooterStyle@autoHeight", POST_PROCESS_FOOTER_STYLE_LABEL) Then Exit Function

    If outStyle.Columns < 1 Then
        MsgBox "Invalid value for postProcessFooter style attribute 'postProcessFooterStyle@columns': must be >= 1.", vbExclamation
        Exit Function
    End If
    If outStyle.FontSize <= 0 Then
        MsgBox "Invalid value for postProcessFooter style attribute 'postProcessFooterStyle@fontSize': must be > 0.", vbExclamation
        Exit Function
    End If
    If outStyle.RowHeight <= 0 Then
        MsgBox "Invalid value for postProcessFooter style attribute 'postProcessFooterStyle@rowHeight': must be > 0.", vbExclamation
        Exit Function
    End If

    mp_TryLoadPostProcessFooterStyle = True
End Function

Private Sub mp_ApplyPostProcessFooterRowHeight( _
    ByVal ws As Worksheet, _
    ByVal postProcessFooterRange As Range, _
    ByVal postProcessFooterText As String, _
    ByRef postProcessFooterStyle As t_PostProcessFooterStyle _
)
    Dim targetRow As Long
    Dim measuredHeight As Double

    If ws Is Nothing Then Exit Sub
    If postProcessFooterRange Is Nothing Then Exit Sub

    targetRow = postProcessFooterRange.Row
    If targetRow <= 0 Then Exit Sub

    If Not postProcessFooterStyle.AutoHeight Or StrComp(postProcessFooterStyle.Overflow, "wrap", vbTextCompare) <> 0 Then
        ws.Rows(targetRow).RowHeight = postProcessFooterStyle.RowHeight
        Exit Sub
    End If

    measuredHeight = mp_MeasurePostProcessFooterTextHeight(ws, postProcessFooterRange, postProcessFooterText, postProcessFooterStyle.FontSize)
    If measuredHeight <= 0 Then
        ws.Rows(targetRow).RowHeight = postProcessFooterStyle.RowHeight
        Exit Sub
    End If

    If measuredHeight < postProcessFooterStyle.RowHeight Then
        ws.Rows(targetRow).RowHeight = postProcessFooterStyle.RowHeight
    Else
        ws.Rows(targetRow).RowHeight = measuredHeight
    End If
End Sub

Private Function mp_MeasurePostProcessFooterTextHeight( _
    ByVal ws As Worksheet, _
    ByVal postProcessFooterRange As Range, _
    ByVal postProcessFooterText As String, _
    ByVal fontSize As Double _
) As Double
    Dim textBoxShape As Object

    On Error GoTo EH
    If ws Is Nothing Then Exit Function
    If postProcessFooterRange Is Nothing Then Exit Function
    If Len(postProcessFooterText) = 0 Then Exit Function

    Set textBoxShape = ws.Shapes.AddTextbox(1, postProcessFooterRange.Left, postProcessFooterRange.Top, postProcessFooterRange.Width, 8)
    textBoxShape.Line.Visible = 0
    textBoxShape.Fill.Visible = 0
    textBoxShape.TextFrame2.MarginLeft = 0
    textBoxShape.TextFrame2.MarginRight = 0
    textBoxShape.TextFrame2.MarginTop = 0
    textBoxShape.TextFrame2.MarginBottom = 0
    textBoxShape.TextFrame2.WordWrap = -1
    textBoxShape.TextFrame2.AutoSize = 1
    textBoxShape.TextFrame2.TextRange.Text = postProcessFooterText
    textBoxShape.TextFrame2.TextRange.Font.Size = fontSize
    textBoxShape.TextFrame2.TextRange.Font.Name = CStr(postProcessFooterRange.Font.Name)

    mp_MeasurePostProcessFooterTextHeight = textBoxShape.Height + 2

Cleanup:
    On Error Resume Next
    If Not textBoxShape Is Nothing Then textBoxShape.Delete
    On Error GoTo 0
    Exit Function

EH:
    mp_MeasurePostProcessFooterTextHeight = 0
    Resume Cleanup
End Function

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
