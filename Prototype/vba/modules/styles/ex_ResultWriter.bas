Attribute VB_Name = "ex_ResultWriter"
Option Explicit

Public Sub m_WriteTableToResultSheet(ByVal tableData As Variant)
    Dim ws As Worksheet
    Dim rowCount As Long
    Dim colCount As Long
    Dim targetRange As Range
    Dim outputStyle As t_OutputSheetStyle
    Dim baseStyle As t_BaseSheetStyle
    Dim hasOutputStyle As Boolean
    Dim layerOrder As Variant
    Dim layerName As Variant

    On Error GoTo EH

    If Not ex_SheetStylesXmlProvider.m_EnsureInitialized(ThisWorkbook) Then
        MsgBox "Не удалось инициализировать реестр стилей.", vbExclamation
        Exit Sub
    End If
    If Not ex_SheetStylesXmlProvider.m_GetBaseSheetStyle(baseStyle, ThisWorkbook) Then Exit Sub
    hasOutputStyle = ex_SheetStylesXmlProvider.m_GetOutputSheetStyle(outputStyle, ThisWorkbook)
    If Not ex_SheetStylesXmlProvider.m_GetLayerOrder(hasOutputStyle, layerOrder, ThisWorkbook) Then Exit Sub

    Set ws = mp_GetOrCreateWorksheet("Result")
    ws.Cells.Clear
    ws.ScrollArea = ""

    rowCount = UBound(tableData, 1)
    colCount = UBound(tableData, 2)

    Set targetRange = ws.Range(ws.Cells(1, 1), ws.Cells(rowCount, colCount))
    targetRange.Value = tableData

    mp_FormatAsTable ws, rowCount, colCount

    For Each layerName In layerOrder
        Select Case CStr(layerName)
            Case ex_SheetStylesXmlProvider.LAYER_BASE
                ex_SheetStylesXmlProvider.m_ApplyBaseLayer ws, rowCount, colCount, baseStyle
            Case ex_SheetStylesXmlProvider.LAYER_OUTPUT
                mp_ApplyOutputStyleToResult ws, rowCount, colCount, outputStyle
                ex_SheetStylesXmlProvider.m_ApplyStatusLayer ws, rowCount, colCount, outputStyle
        End Select
    Next layerName

    Exit Sub
EH:
    MsgBox "Result writer error: " & Err.Description, vbExclamation
End Sub

Private Function mp_GetOrCreateWorksheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim fullName As String

    fullName = "g_" & sheetName

    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, fullName, vbTextCompare) = 0 Then
            Set mp_GetOrCreateWorksheet = ws
            Exit Function
        End If
    Next ws

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = fullName
    ex_SheetStylesXmlProvider.m_ApplyDefaultSheetView ws
    Set mp_GetOrCreateWorksheet = ws
End Function

Private Sub mp_FormatAsTable(ByVal ws As Worksheet, ByVal rowCount As Long, ByVal colCount As Long)
    Dim headerRange As Range
    Dim allRange As Range

    Set headerRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount))
    Set allRange = ws.Range(ws.Cells(1, 1), ws.Cells(rowCount, colCount))

    allRange.Font.Name = "Segoe UI"
    allRange.Font.Size = 10
    headerRange.Font.Bold = True
    allRange.HorizontalAlignment = xlCenter
    allRange.VerticalAlignment = xlCenter
    allRange.EntireColumn.AutoFit
    allRange.AutoFilter

    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub mp_ApplyOutputStyleToResult( _
    ByVal ws As Worksheet, _
    ByVal rowCount As Long, _
    ByVal colCount As Long, _
    ByRef style As t_OutputSheetStyle _
)
    Dim targetRange As Range
    Dim headerRange As Range

    Set targetRange = ws.Range(ws.Cells(1, 1), ws.Cells(rowCount, colCount))
    Set headerRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount))

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
