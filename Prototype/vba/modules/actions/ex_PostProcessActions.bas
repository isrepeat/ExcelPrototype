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

Public Sub m_HighlightRowCell( _
    ByVal rowRef As obj_ResultRow, _
    ByVal columnRef As String, _
    Optional ByVal colorHex As String = "#FFF2CC" _
)
    Dim colorValue As Long
    Dim targetCol As Long
    Dim targetCell As Range
    Dim ws As Worksheet
    Dim usedCols As Long

    If rowRef Is Nothing Then Exit Sub
    columnRef = Trim$(columnRef)
    If Len(columnRef) = 0 Then
        Err.Raise vbObjectError + 1652, "ex_PostProcessActions", "Column reference is empty for row cell highlight."
    End If
    If Len(Trim$(colorHex)) = 0 Then colorHex = "#FFF2CC"

    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    usedCols = mp_GetLastUsedColumn(ws)
    If usedCols <= 0 Then usedCols = 1

    If Not mp_TryResolveColumnIndexInRow(rowRef, columnRef, targetCol) Then
        Err.Raise vbObjectError + 1653, "ex_PostProcessActions", "Unknown row cell reference '" & columnRef & "'. Use 1-based column index or field alias."
    End If
    If targetCol < 1 Or targetCol > usedCols Then
        Err.Raise vbObjectError + 1654, "ex_PostProcessActions", "Column index '" & CStr(targetCol) & "' is out of used range 1.." & CStr(usedCols) & "."
    End If

    If Not ex_XmlCore.m_TryParseColor(colorHex, colorValue) Then
        Err.Raise vbObjectError + 1650, "ex_PostProcessActions", "Invalid highlight color: " & colorHex
    End If

    Set targetCell = ws.Cells(rowRef.RowIndex, targetCol)
    targetCell.Interior.Pattern = xlSolid
    targetCell.Interior.Color = colorValue
End Sub

Public Function m_RegexIsMatch( _
    ByVal textValue As String, _
    ByVal regexPattern As String _
) As Boolean
    Dim rx As Object

    Set rx = mp_CreateRegex(regexPattern)
    m_RegexIsMatch = rx.Test(CStr(textValue))
End Function

Public Function m_RegexFirstMatch( _
    ByVal textValue As String, _
    ByVal regexPattern As String _
) As String
    Dim rx As Object
    Dim matches As Object

    Set rx = mp_CreateRegex(regexPattern)
    Set matches = rx.Execute(CStr(textValue))
    If matches.Count > 0 Then
        m_RegexFirstMatch = CStr(matches(0).Value)
    End If
End Function

Public Function m_RowToText( _
    ByVal rowRef As obj_ResultRow, _
    ByVal separatorText As String _
) As String
    m_RowToText = mp_GetRowText(rowRef, separatorText)
End Function

Public Function m_RowCellRegexIsMatch( _
    ByVal rowRef As obj_ResultRow, _
    ByVal columnRef As String, _
    ByVal regexPattern As String _
) As Boolean
    m_RowCellRegexIsMatch = m_RegexIsMatch(mp_GetRowCellLiveText(rowRef, columnRef), regexPattern)
End Function

Public Function m_RowCellRegexFirstMatch( _
    ByVal rowRef As obj_ResultRow, _
    ByVal columnRef As String, _
    ByVal regexPattern As String _
) As String
    m_RowCellRegexFirstMatch = m_RegexFirstMatch(mp_GetRowCellLiveText(rowRef, columnRef), regexPattern)
End Function

Public Sub m_EmphasizeRowCellTextByRegex( _
    ByVal rowRef As obj_ResultRow, _
    ByVal columnRef As String, _
    ByVal regexPattern As String, _
    Optional ByVal fontColorHex As String = "#FF0000", _
    Optional ByVal uppercaseMatches As String = "false" _
)
    Dim targetCell As Range
    Dim targetCol As Long
    Dim originalText As String
    Dim transformedText As String
    Dim rx As Object
    Dim matches As Object
    Dim i As Long
    Dim matchObj As Object
    Dim colorValue As Long
    Dim makeUpper As Boolean

    If rowRef Is Nothing Then
        Err.Raise vbObjectError + 1664, "ex_PostProcessActions", "Row reference is required for regex text emphasis."
    End If

    If Not mp_TryResolveColumnIndexInRow(rowRef, columnRef, targetCol) Then
        Err.Raise vbObjectError + 1665, "ex_PostProcessActions", "Unknown row cell reference '" & columnRef & "' for regex text emphasis."
    End If
    Set targetCell = mp_GetRowCellRange(rowRef.RowIndex, targetCol)
    originalText = CStr(targetCell.Value)

    If Len(Trim$(fontColorHex)) = 0 Then fontColorHex = "#FF0000"
    If Not ex_XmlCore.m_TryParseColor(fontColorHex, colorValue) Then
        Err.Raise vbObjectError + 1666, "ex_PostProcessActions", "Invalid regex emphasis color: " & fontColorHex
    End If
    makeUpper = mp_ParseRequiredBoolean(uppercaseMatches, "uppercaseMatches")

    Set rx = mp_CreateRegex(regexPattern, True)
    Set matches = rx.Execute(originalText)
    If matches Is Nothing Or matches.Count = 0 Then Exit Sub

    If makeUpper Then
        transformedText = originalText
        For i = 0 To matches.Count - 1
            Set matchObj = matches(i)
            If matchObj.Length > 0 Then
                transformedText = Left$(transformedText, matchObj.FirstIndex) & UCase$(Mid$(transformedText, matchObj.FirstIndex + 1, matchObj.Length)) & Mid$(transformedText, matchObj.FirstIndex + matchObj.Length + 1)
            End If
        Next i
        targetCell.Value = transformedText
    End If

    For i = 0 To matches.Count - 1
        Set matchObj = matches(i)
        If matchObj.Length > 0 Then
            targetCell.Characters(matchObj.FirstIndex + 1, matchObj.Length).Font.Color = colorValue
            targetCell.Characters(matchObj.FirstIndex + 1, matchObj.Length).Font.Bold = True
        End If
    Next i
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

Private Function mp_TryResolveColumnIndexInRow( _
    ByVal rowRef As obj_ResultRow, _
    ByVal columnRef As String, _
    ByRef outColumnIndex As Long _
) As Boolean
    Dim numericIndex As Long
    Dim columns As Collection
    Dim i As Long
    Dim colObj As obj_ResultColumn

    If rowRef Is Nothing Then Exit Function
    columnRef = Trim$(columnRef)
    If Len(columnRef) = 0 Then Exit Function

    If ex_XmlCore.m_TryParseLong(columnRef, numericIndex) Then
        If numericIndex < 1 Then Exit Function
        Set columns = rowRef.Columns
        If numericIndex > columns.Count Then Exit Function
        outColumnIndex = numericIndex
        mp_TryResolveColumnIndexInRow = True
        Exit Function
    End If

    Set columns = rowRef.Columns
    For i = 1 To columns.Count
        Set colObj = columns(i)
        If StrComp(colObj.Alias, columnRef, vbTextCompare) = 0 Then
            outColumnIndex = i
            mp_TryResolveColumnIndexInRow = True
            Exit Function
        End If
    Next i
End Function

Private Function mp_GetRowText( _
    ByVal rowRef As obj_ResultRow, _
    ByVal separatorText As String _
) As String
    Dim columns As Collection
    Dim i As Long
    Dim colObj As obj_ResultColumn
    Dim result As String

    If rowRef Is Nothing Then
        Err.Raise vbObjectError + 1655, "ex_PostProcessActions", "Row reference is required for row text build."
    End If
    If Len(separatorText) = 0 Then
        Err.Raise vbObjectError + 1663, "ex_PostProcessActions", "Separator is required for row text build."
    End If

    Set columns = rowRef.Columns
    For i = 1 To columns.Count
        Set colObj = columns(i)
        If i > 1 Then result = result & separatorText
        result = result & CStr(colObj.Value)
    Next i

    mp_GetRowText = result
End Function

Private Function mp_GetRowCellText( _
    ByVal rowRef As obj_ResultRow, _
    ByVal columnRef As String _
) As String
    Dim numericIndex As Long
    Dim columns As Collection

    If rowRef Is Nothing Then
        Err.Raise vbObjectError + 1656, "ex_PostProcessActions", "Row reference is required for regex cell parsing."
    End If

    columnRef = Trim$(columnRef)
    If Len(columnRef) = 0 Then
        Err.Raise vbObjectError + 1657, "ex_PostProcessActions", "Column reference is empty for regex cell parsing."
    End If

    If ex_XmlCore.m_TryParseLong(columnRef, numericIndex) Then
        If numericIndex < 1 Then
            Err.Raise vbObjectError + 1658, "ex_PostProcessActions", "Column index must be >= 1 for regex cell parsing."
        End If

        Set columns = rowRef.Columns
        If numericIndex > columns.Count Then
            Err.Raise vbObjectError + 1659, "ex_PostProcessActions", "Column index '" & CStr(numericIndex) & "' is out of row bounds (max " & CStr(columns.Count) & ")."
        End If

        mp_GetRowCellText = CStr(columns(numericIndex).Value)
        Exit Function
    End If

    If Not rowRef.HasAlias(columnRef) Then
        Err.Raise vbObjectError + 1660, "ex_PostProcessActions", "Field alias '" & columnRef & "' is not available in row."
    End If
    mp_GetRowCellText = CStr(rowRef.Column(columnRef))
End Function

Private Function mp_GetRowCellLiveText( _
    ByVal rowRef As obj_ResultRow, _
    ByVal columnRef As String _
) As String
    Dim targetCol As Long
    Dim targetCell As Range

    If rowRef Is Nothing Then
        Err.Raise vbObjectError + 1671, "ex_PostProcessActions", "Row reference is required for live cell text parsing."
    End If
    If Not mp_TryResolveColumnIndexInRow(rowRef, columnRef, targetCol) Then
        Err.Raise vbObjectError + 1672, "ex_PostProcessActions", "Unknown row cell reference '" & columnRef & "' for live cell text parsing."
    End If

    Set targetCell = mp_GetRowCellRange(rowRef.RowIndex, targetCol)
    mp_GetRowCellLiveText = CStr(targetCell.Value)
End Function

Private Function mp_CreateRegex( _
    ByVal regexPattern As String, _
    Optional ByVal globalMatches As Boolean = False _
) As Object
    Dim rx As Object

    regexPattern = Trim$(regexPattern)
    If Len(regexPattern) = 0 Then
        Err.Raise vbObjectError + 1661, "ex_PostProcessActions", "Regex pattern is empty."
    End If

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = globalMatches
    rx.IgnoreCase = True
    rx.MultiLine = True

    On Error GoTo PatternErr
    rx.Pattern = regexPattern
    On Error GoTo 0

    Set mp_CreateRegex = rx
    Exit Function

PatternErr:
    Err.Raise vbObjectError + 1662, "ex_PostProcessActions", "Invalid regex pattern '" & regexPattern & "': " & Err.Description
End Function

Private Function mp_ParseRequiredBoolean(ByVal valueText As String, ByVal fieldName As String) As Boolean
    Dim parsedValue As Boolean

    valueText = Trim$(valueText)
    If Not ex_XmlCore.m_TryParseBoolean(valueText, parsedValue) Then
        Err.Raise vbObjectError + 1667, "ex_PostProcessActions", "Invalid boolean for '" & fieldName & "': '" & valueText & "'."
    End If

    mp_ParseRequiredBoolean = parsedValue
End Function

Private Function mp_GetRowCellRange( _
    ByVal rowIndex As Long, _
    ByVal columnIndex As Long _
) As Range
    Dim ws As Worksheet

    Set ws = ActiveSheet
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1668, "ex_PostProcessActions", "Active sheet is not available for regex text emphasis."
    End If
    If rowIndex < 1 Then
        Err.Raise vbObjectError + 1669, "ex_PostProcessActions", "Row index must be >= 1 for regex text emphasis."
    End If
    If columnIndex < 1 Then
        Err.Raise vbObjectError + 1670, "ex_PostProcessActions", "Column index must be >= 1 for regex text emphasis."
    End If

    Set mp_GetRowCellRange = ws.Cells(rowIndex, columnIndex)
End Function
