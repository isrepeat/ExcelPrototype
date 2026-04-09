VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_TableListControlVM"
Option Explicit
Implements obj_IControl

#Const ENALBE_STYLES = True
#Const ENALBE_BORDERS = True

Private m_ControlName As String
Private m_ItemsSourceRaw As String
Private m_ItemVisibilityRaw As String
Private m_LayoutSheet As String
Private m_RowStart As Long
Private m_ColStart As Long
Private m_RowEnd As Long
Private m_ColEnd As Long
Private m_TableItems As Collection
Private m_IsConfigured As Boolean

Private Sub obj_IControl_Configure(ByVal controlNode As Object)
    m_IsConfigured = False
    Set m_TableItems = Nothing

    If controlNode Is Nothing Then
        MsgBox "Table: control node is not specified.", vbExclamation
        Exit Sub
    End If

    m_ControlName = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "name"))
    If Len(m_ControlName) = 0 Then m_ControlName = "tablelist"

    m_ItemsSourceRaw = Trim$(CStr(ex_XmlCore.m_NodeAttrText(controlNode, "itemsSource")))
    If Len(m_ItemsSourceRaw) = 0 Then 
        MsgBox "Table: itemsSource is not specified for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If
    m_ItemVisibilityRaw = Trim$(CStr(ex_XmlCore.m_NodeAttrText(controlNode, "itemVisibility")))

    m_LayoutSheet = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "__layoutSheet"))
    If Len(m_LayoutSheet) = 0 Then
        MsgBox "Table: runtime layout sheet is missing for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    If Not mp_TryReadLayoutLongAttr(controlNode, "__layoutRowStart", m_RowStart, True) Then Exit Sub
    If Not mp_TryReadLayoutLongAttr(controlNode, "__layoutColStart", m_ColStart, True) Then Exit Sub
    If Not mp_TryReadLayoutLongAttr(controlNode, "__layoutRowEnd", m_RowEnd, True) Then Exit Sub
    If Not mp_TryReadLayoutLongAttr(controlNode, "__layoutColEnd", m_ColEnd, True) Then Exit Sub

    If m_RowStart <= 0 Or m_ColStart <= 0 Then
        MsgBox "Table: invalid row/column start for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    If m_RowEnd < m_RowStart Then
        MsgBox "Table: control '" & m_ControlName & "' has invalid spanRows range.", vbExclamation
        Exit Sub
    End If

    If m_ColEnd < m_ColStart Then
        MsgBox "Table: control '" & m_ControlName & "' has invalid spanCells range.", vbExclamation
        Exit Sub
    End If

    If Not ex_ListItemsSourceRuntime.m_TryResolveItemsSource(m_ItemsSourceRaw, m_TableItems) Then Exit Sub
    If Not mp_TryApplyItemVisibilityFilter(m_TableItems) Then Exit Sub

    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim valueBlock As Variant
    Dim targetRange As Range
    Dim styleSegments As Collection

    If Not m_IsConfigured Then
        MsgBox "Table: control '" & m_ControlName & "' is not configured.", vbExclamation
        Exit Sub
    End If

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then
        MsgBox "Table: workbook is not specified for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    Set ws = mp_GetWorksheetByName(wb, m_LayoutSheet)
    If ws Is Nothing Then
        MsgBox "Table: sheet '" & m_LayoutSheet & "' was not found for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    If m_TableItems Is Nothing Then
        MsgBox "Table: itemsSource is not resolved for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    ' Build a full in-memory block first, then write once to the sheet.
    ' This avoids per-cell COM calls, which are the main render bottleneck.
    If Not mp_TryBuildRenderBuffer(valueBlock, styleSegments) Then Exit Sub
    If IsEmpty(valueBlock) Then Exit Sub

    Set targetRange = ws.Range( _
        ws.Cells(m_RowStart, m_ColStart), _
        ws.Cells(m_RowStart + UBound(valueBlock, 1) - 1, m_ColStart + UBound(valueBlock, 2) - 1))

    ' Single bulk assignment is significantly faster than incremental writes.
    targetRange.Value2 = valueBlock

    If Not mp_TryRegisterControlPartSegments(ws, styleSegments) Then Exit Sub

#If ENALBE_STYLES Then
    mp_ApplyStyleSegments ws, styleSegments
#End If
End Sub

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case LCase$(Trim$(attrName))
        Case "itemssource", "itemvisibility"
            obj_IControl_SupportsAttribute = True
    End Select
End Function

Private Function mp_TryApplyItemVisibilityFilter(ByRef tableItems As Collection) As Boolean
    Dim filteredItems As Collection
    Dim tableItem As Variant
    Dim isVisible As Boolean

    If Len(m_ItemVisibilityRaw) = 0 Then
        mp_TryApplyItemVisibilityFilter = True
        Exit Function
    End If

    If tableItems Is Nothing Then
        MsgBox "Table: itemsSource is not resolved for control '" & m_ControlName & "'.", vbExclamation
        Exit Function
    End If

    Set filteredItems = New Collection

    For Each tableItem In tableItems
        If Not IsObject(tableItem) Then
            MsgBox "Table: itemsSource entry must be an object for itemVisibility evaluation in control '" & m_ControlName & "'.", vbExclamation
            Exit Function
        End If

        If Not ex_BindingRuntime.m_TryResolveVisibilityBinding(m_ItemVisibilityRaw, tableItem, isVisible) Then Exit Function
        If isVisible Then filteredItems.Add tableItem
    Next tableItem

    Set tableItems = filteredItems
    mp_TryApplyItemVisibilityFilter = True
End Function

Private Function mp_TryReadLayoutLongAttr( _
    ByVal controlNode As Object, _
    ByVal attrName As String, _
    ByRef outValue As Long, _
    ByVal isRequired As Boolean _
) As Boolean
    Dim rawText As String

    rawText = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, attrName))
    If Len(rawText) = 0 Then
        If isRequired Then
            MsgBox "Table: runtime layout attribute '" & attrName & "' is missing for control '" & m_ControlName & "'.", vbExclamation
            Exit Function
        End If

        outValue = 0
        mp_TryReadLayoutLongAttr = True
        Exit Function
    End If

    If Not IsNumeric(rawText) Then
        MsgBox "Table: runtime layout attribute '" & attrName & "' must be numeric for control '" & m_ControlName & "'.", vbExclamation
        Exit Function
    End If

    outValue = CLng(rawText)
    mp_TryReadLayoutLongAttr = True
End Function

Private Function mp_TryBuildRenderBuffer(ByRef outValueBlock As Variant, ByRef outStyleSegments As Collection) As Boolean
    Dim tableItem As Variant
    Dim tableModel As obj_TableDynamic
    Dim tableRows As Collection
    Dim tableColumnCount As Long
    Dim availableCols As Long
    Dim maxRows As Long
    Dim plannedRows As Long
    Dim currentOutputRow As Long
    Dim dataRowsToWrite As Long
    Dim rowItem As Variant
    Dim rowModel As obj_Row
    Dim colOffset As Long
    Dim tokens As Variant
    Dim writeStart As Long
    Dim writeEnd As Long

    availableCols = mp_GetAvailableColumnCount()
    maxRows = m_RowEnd - m_RowStart + 1

    If availableCols <= 0 Or maxRows <= 0 Then
        MsgBox "Table: invalid render bounds for control '" & m_ControlName & "'.", vbExclamation
        Exit Function
    End If

    plannedRows = 0

    ' Pass 1: estimate final output size and validate table shape constraints.
    ' Pre-sizing the matrix avoids repeated ReDim/Preserve in the hot path.

    For Each tableItem In m_TableItems
        If plannedRows >= maxRows Then Exit For

        If Not mp_TryResolveTableModel(tableItem, tableModel) Then Exit Function

        tableColumnCount = tableModel.ColumnCount
        If tableColumnCount <= 0 Then
            MsgBox "Table: table item has no columns.", vbExclamation
            Exit Function
        End If

        If tableColumnCount > availableCols Then
            MsgBox "Table: control '" & m_ControlName & "' requires " & CStr(tableColumnCount) & _
                   " columns, but span provides only " & CStr(availableCols) & ".", vbExclamation
            Exit Function
        End If

        Set tableRows = tableModel.Rows
        dataRowsToWrite = 0
        If Not tableRows Is Nothing Then dataRowsToWrite = tableRows.Count

        ' section + header + data + spacer
        plannedRows = plannedRows + 2 + dataRowsToWrite + 1
        If plannedRows > maxRows Then
            plannedRows = maxRows
            Exit For
        End If
    Next tableItem

    If plannedRows = 0 Then
        outValueBlock = Empty
        mp_TryBuildRenderBuffer = True
        Exit Function
    End If

    ReDim outValueBlock(1 To plannedRows, 1 To availableCols)

#If ENALBE_STYLES Then
    Set outStyleSegments = New Collection
#End If

    currentOutputRow = 0

    ' Pass 2: fill matrix sequentially and collect compact style segments.
    ' Data rows use obj_Row.m_CopyToMatrixRow to keep per-cell overhead minimal.

    For Each tableItem In m_TableItems
        If currentOutputRow >= plannedRows Then Exit For

        If Not mp_TryResolveTableModel(tableItem, tableModel) Then Exit Function

        tableColumnCount = tableModel.ColumnCount
        Set tableRows = tableModel.Rows

        ' Section row
        currentOutputRow = currentOutputRow + 1
        outValueBlock(currentOutputRow, 1) = tableModel.SectionTitle
#If ENALBE_STYLES Then
        mp_AddStyleSegment outStyleSegments, "section", tableColumnCount, currentOutputRow, currentOutputRow
#End If
        If currentOutputRow >= plannedRows Then Exit For

        ' Header row
        currentOutputRow = currentOutputRow + 1
        tokens = Split(tableModel.HeaderText, "|")
        For colOffset = 1 To tableColumnCount
            If colOffset - 1 <= UBound(tokens) Then
                outValueBlock(currentOutputRow, colOffset) = Trim$(CStr(tokens(colOffset - 1)))
            End If
        Next colOffset
#If ENALBE_STYLES Then
        mp_AddStyleSegment outStyleSegments, "header", tableColumnCount, currentOutputRow, currentOutputRow
#End If
        If currentOutputRow >= plannedRows Then Exit For

        ' Data rows
        writeStart = currentOutputRow + 1
        If Not tableRows Is Nothing Then
            On Error GoTo EH_INVALID_ROW
            For Each rowItem In tableRows
                If currentOutputRow >= plannedRows Then Exit For

                currentOutputRow = currentOutputRow + 1
                Set rowModel = rowItem
                rowModel.m_CopyToMatrixRow outValueBlock, currentOutputRow, tableColumnCount
            Next rowItem
            On Error GoTo 0
        End If
        writeEnd = currentOutputRow
#If ENALBE_STYLES Then
        If writeEnd >= writeStart Then
            mp_AddStyleSegment outStyleSegments, "data", tableColumnCount, writeStart, writeEnd
        End If
#End If
        If currentOutputRow >= plannedRows Then Exit For

        ' Spacer row
        currentOutputRow = currentOutputRow + 1
#If ENALBE_STYLES Then
        mp_AddStyleSegment outStyleSegments, "spacer", tableColumnCount, currentOutputRow, currentOutputRow
#End If
    Next tableItem

    mp_TryBuildRenderBuffer = True
    Exit Function

EH_INVALID_ROW:
    On Error GoTo 0
    MsgBox "Table: unsupported row object in table rows. Expected obj_Row.", vbExclamation
End Function

Private Sub mp_AddStyleSegment( _
    ByVal styleSegments As Collection, _
    ByVal styleKind As String, _
    ByVal columnCount As Long, _
    ByVal rowStart As Long, _
    ByVal rowEnd As Long _
)
    Dim segment As Object
    Dim lastSegment As Object

    If styleSegments Is Nothing Then Exit Sub
    If rowEnd < rowStart Then Exit Sub

    ' Merge adjacent segments with same style+width to reduce style operations later.
    If styleSegments.Count > 0 Then
        Set lastSegment = styleSegments(styleSegments.Count)
        If StrComp(CStr(lastSegment("StyleKind")), styleKind, vbTextCompare) = 0 Then
            If CLng(lastSegment("ColumnCount")) = columnCount Then
                If CLng(lastSegment("RowEnd")) + 1 = rowStart Then
                    lastSegment("RowEnd") = rowEnd
                    Exit Sub
                End If
            End If
        End If
    End If

    Set segment = CreateObject("Scripting.Dictionary")
    segment.CompareMode = 1
    segment("StyleKind") = styleKind
    segment("ColumnCount") = columnCount
    segment("RowStart") = rowStart
    segment("RowEnd") = rowEnd

    styleSegments.Add segment
End Sub

Private Sub mp_ApplyStyleSegments(ByVal ws As Worksheet, ByVal styleSegments As Collection)
    Dim groupedRanges As Object
    Dim segment As Object
    Dim segmentRange As Range
    Dim groupedKey As String
    Dim groupedRange As Range
    Dim key As Variant
    Dim sepPos As Long
    Dim styleKind As String
    Dim columnCount As Long
    Dim backColor As Long
    Dim fontColor As Long
    Dim borderColor As Long
    Dim fontSize As Double
    Dim fontBold As Boolean

    If ws Is Nothing Then Exit Sub
    If styleSegments Is Nothing Then Exit Sub

    ' Group by (style kind + column count) and union ranges.
    ' This turns many small style calls into a few larger batch calls.
    Set groupedRanges = CreateObject("Scripting.Dictionary")
    groupedRanges.CompareMode = 1

    For Each segment In styleSegments
        styleKind = LCase$(CStr(segment("StyleKind")))
        columnCount = CLng(segment("ColumnCount"))

        Set segmentRange = mp_BuildSegmentRange( _
            ws, _
            CLng(segment("RowStart")), _
            CLng(segment("RowEnd")), _
            columnCount)
        If segmentRange Is Nothing Then GoTo ContinueSegment

        groupedKey = styleKind & "|" & CStr(columnCount)

        If groupedRanges.Exists(groupedKey) Then
            Set groupedRange = groupedRanges(groupedKey)
            Set groupedRanges(groupedKey) = Application.Union(groupedRange, segmentRange)
        Else
            groupedRanges.Add groupedKey, segmentRange
        End If

ContinueSegment:
    Next segment

    For Each key In groupedRanges.Keys
        sepPos = InStrRev(CStr(key), "|", -1, vbBinaryCompare)
        If sepPos <= 1 Then GoTo ContinueGroup

        styleKind = Left$(CStr(key), sepPos - 1)

        If Not mp_TryResolveStylePreset(styleKind, backColor, fontColor, borderColor, fontSize, fontBold) Then Exit Sub

        Set groupedRange = groupedRanges(CStr(key))
        If groupedRange Is Nothing Then GoTo ContinueGroup

        mp_ApplyRowStyle groupedRange, backColor, fontColor, borderColor, fontSize, fontBold

ContinueGroup:
    Next key
End Sub

Private Function mp_TryRegisterControlPartSegments(ByVal ws As Worksheet, ByVal styleSegments As Collection) As Boolean
    Dim segment As Object
    Dim partName As String
    Dim segmentRange As Range

    If ws Is Nothing Then Exit Function
    If styleSegments Is Nothing Then
        mp_TryRegisterControlPartSegments = True
        Exit Function
    End If

    For Each segment In styleSegments
        partName = mp_MapStyleKindToControlPart(CStr(segment("StyleKind")))
        If Len(partName) = 0 Then GoTo ContinueSegment

        Set segmentRange = mp_BuildSegmentRange( _
            ws, _
            CLng(segment("RowStart")), _
            CLng(segment("RowEnd")), _
            CLng(segment("ColumnCount")))
        If segmentRange Is Nothing Then GoTo ContinueSegment

        If Not ex_ControlPartsRuntime.m_RegisterControlPart( _
            ws, _
            "tablelist", _
            m_ControlName, _
            partName, _
            segmentRange) Then Exit Function

ContinueSegment:
    Next segment

    mp_TryRegisterControlPartSegments = True
End Function

Private Function mp_MapStyleKindToControlPart(ByVal styleKind As String) As String
    Select Case LCase$(Trim$(styleKind))
        Case "section"
            mp_MapStyleKindToControlPart = "section"
        Case "header"
            mp_MapStyleKindToControlPart = "header"
        Case "data"
            mp_MapStyleKindToControlPart = "rows"
        Case "spacer"
            mp_MapStyleKindToControlPart = "spacer"
    End Select
End Function

Private Function mp_BuildSegmentRange( _
    ByVal ws As Worksheet, _
    ByVal relativeRowStart As Long, _
    ByVal relativeRowEnd As Long, _
    ByVal columnCount As Long _
) As Range
    Dim absRowStart As Long
    Dim absRowEnd As Long

    If ws Is Nothing Then Exit Function
    If relativeRowStart <= 0 Or relativeRowEnd < relativeRowStart Then Exit Function
    If columnCount <= 0 Then Exit Function

    absRowStart = m_RowStart + relativeRowStart - 1
    absRowEnd = m_RowStart + relativeRowEnd - 1

    Set mp_BuildSegmentRange = ws.Range( _
        ws.Cells(absRowStart, m_ColStart), _
        ws.Cells(absRowEnd, m_ColStart + columnCount - 1))
End Function

Private Function mp_TryResolveStylePreset( _
    ByVal styleKind As String, _
    ByRef backColor As Long, _
    ByRef fontColor As Long, _
    ByRef borderColor As Long, _
    ByRef fontSize As Double, _
    ByRef fontBold As Boolean _
) As Boolean
    Select Case LCase$(Trim$(styleKind))
        Case "section"
            backColor = RGB(23, 58, 94)
            fontColor = RGB(234, 246, 255)
            borderColor = RGB(14, 34, 57)
            fontSize = 11
            fontBold = True

        Case "header"
            backColor = RGB(43, 74, 107)
            fontColor = RGB(221, 238, 255)
            borderColor = RGB(31, 54, 80)
            fontSize = 10
            fontBold = True

        Case "data"
            backColor = RGB(58, 58, 58)
            fontColor = RGB(240, 240, 240)
            borderColor = RGB(42, 42, 42)
            fontSize = 10
            fontBold = False

        Case "spacer"
            backColor = RGB(31, 31, 31)
            fontColor = RGB(31, 31, 31)
            borderColor = RGB(31, 31, 31)
            fontSize = 8
            fontBold = False

        Case Else
            MsgBox "Table: unsupported style segment kind '" & styleKind & "'.", vbExclamation
            Exit Function
    End Select

    mp_TryResolveStylePreset = True
End Function

Private Function mp_GetAvailableColumnCount() As Long
    If m_ColEnd <= 0 Or m_ColStart <= 0 Then Exit Function
    mp_GetAvailableColumnCount = m_ColEnd - m_ColStart + 1
End Function

Private Function mp_TryResolveTableModel(ByVal tableItem As Variant, ByRef outTable As obj_TableDynamic) As Boolean
    Dim fixedTable As obj_Table

    If Not IsObject(tableItem) Then
        MsgBox "Table: itemsSource entry must be an object of type obj_TableDynamic or obj_Table.", vbExclamation
        Exit Function
    End If

    Select Case LCase$(TypeName(tableItem))
        Case "obj_tabledynamic"
            Set outTable = tableItem
            mp_TryResolveTableModel = True

        Case "obj_table"
            Set fixedTable = tableItem
            Set outTable = mp_ConvertFixedTableToDynamic(fixedTable)
            If outTable Is Nothing Then Exit Function
            mp_TryResolveTableModel = True

        Case Else
            MsgBox "Table: unsupported table model type '" & TypeName(tableItem) & "'. Expected obj_TableDynamic or obj_Table.", vbExclamation
    End Select
End Function

Private Function mp_ConvertFixedTableToDynamic(ByVal fixedTable As obj_Table) As obj_TableDynamic
    Dim dynamicTable As obj_TableDynamic
    Dim sourceColumns As Collection
    Dim sourceRows As Collection
    Dim sourceColumn As obj_Column
    Dim sourceRow As obj_Row
    Dim targetColumn As obj_Column
    Dim targetRow As obj_Row
    Dim colIndex As Long

    If fixedTable Is Nothing Then
        MsgBox "Table: fixed table model is not specified.", vbExclamation
        Exit Function
    End If

    Set dynamicTable = New obj_TableDynamic
    dynamicTable.SectionTitle = fixedTable.SectionTitle

    Set sourceColumns = fixedTable.Columns
    For Each sourceColumn In sourceColumns
        Set targetColumn = New obj_Column
        targetColumn.Position = sourceColumn.Position
        targetColumn.Name = sourceColumn.Name
        If Not dynamicTable.m_AddColumn(targetColumn) Then Exit Function
    Next sourceColumn

    Set sourceRows = fixedTable.Rows
    For Each sourceRow In sourceRows
        Set targetRow = New obj_Row
        For colIndex = 1 To dynamicTable.ColumnCount
            targetRow.m_AddCell sourceRow.m_GetCell(colIndex)
        Next colIndex
        If Not dynamicTable.m_AddRow(targetRow) Then Exit Function
    Next sourceRow

    Set mp_ConvertFixedTableToDynamic = dynamicTable
End Function

Private Sub mp_ApplyRowStyle( _
    ByVal targetRange As Range, _
    ByVal backColor As Long, _
    ByVal fontColor As Long, _
    ByVal borderColor As Long, _
    ByVal fontSize As Double, _
    ByVal fontBold As Boolean _
)
    With targetRange
        .Interior.Color = backColor
        .Font.Color = fontColor
        .Font.Name = "Calibri"
        .Font.Size = fontSize
        .Font.Bold = fontBold
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignCenter
        .WrapText = False
#If ENALBE_BORDERS Then
        .Borders.LineStyle = xlContinuous
        .Borders.Color = borderColor
        .Borders.Weight = xlThin
#End If
    End With
End Sub

Private Function mp_CanWriteRow(ByVal rowIndex As Long) As Boolean
    If rowIndex <= 0 Then Exit Function

    If m_RowEnd > 0 Then
        mp_CanWriteRow = (rowIndex <= m_RowEnd)
    Else
        mp_CanWriteRow = True
    End If
End Function

Private Function mp_GetWorksheetByName(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set mp_GetWorksheetByName = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function
