VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_TableSingleControlVM"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean
Implements obj_IControl

#Const ENALBE_STYLES = True
#Const ENALBE_BORDERS = True

Private m_ControlBase As obj_ControlBase
Private m_ControlName As String
Private m_ItemsSourceRaw As String
Private m_LayoutSheetName As String
Private m_RowStart As Long
Private m_ColStart As Long
Private m_RowEnd As Long
Private m_ColEnd As Long
Private m_TableItems As Collection
Private m_IsConfigured As Boolean
Private m_Page As obj_IPage

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Dispose
    On Error GoTo 0
End Sub

' //
' // Interface
' //
Private Sub obj_IControl_Configure(ByVal controlNode As Object)
    Dim pageBase As obj_PageBase

    m_IsConfigured = False
    Set m_TableItems = Nothing
    Set m_ControlBase = Nothing

    Set pageBase = m_Page.GetPageBase()
    Set m_ControlBase = New obj_ControlBase
    If Not m_ControlBase.Initialize(m_Page) Then Exit Sub
    If Not m_ControlBase.Configure(pageBase, controlNode, "TableSingle", "tableSingle", m_ControlName) Then Exit Sub

    m_ItemsSourceRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "itemsSource")))
    If VBA.Len(m_ItemsSourceRaw) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableSingle: itemsSource is not specified for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    m_LayoutSheetName = VBA.Trim$(ex_XmlCore.fn_NodeAttrText(controlNode, "__layoutSheetName"))
    If VBA.Len(m_LayoutSheetName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableSingle: runtime layout sheet is missing for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    If Not private_TryReadLayoutLongAttr(controlNode, "__layoutRowStart", m_RowStart, True) Then Exit Sub
    If Not private_TryReadLayoutLongAttr(controlNode, "__layoutColStart", m_ColStart, True) Then Exit Sub
    If Not private_TryReadLayoutLongAttr(controlNode, "__layoutRowEnd", m_RowEnd, True) Then Exit Sub
    If Not private_TryReadLayoutLongAttr(controlNode, "__layoutColEnd", m_ColEnd, True) Then Exit Sub

    If m_RowStart <= 0 Or m_ColStart <= 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableSingle: invalid row/column start for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If
    If m_RowEnd < m_RowStart Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableSingle: control '" & m_ControlName & "' has invalid spanRows range."
#End If
        Exit Sub
    End If
    If m_ColEnd < m_ColStart Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableSingle: control '" & m_ControlName & "' has invalid spanColls range."
#End If
        Exit Sub
    End If

    Set pageBase = m_ControlBase.PageBase
    If pageBase Is Nothing Then Exit Sub
    If Not ex_RuntimeSourceResolver.fn_TryResolveItemsSource(pageBase.RuntimeSources, m_ItemsSourceRaw, m_TableItems) Then Exit Sub

    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render()
    Dim ws As Worksheet
    Dim valueBlock As Variant
    Dim targetRange As Range
    Dim styleSegments As Collection
    Dim page As obj_PageBase

    If Not m_IsConfigured Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableSingle: control '" & m_ControlName & "' is not configured."
#End If
        Exit Sub
    End If

    Set page = Nothing
    If Not m_ControlBase Is Nothing Then Set page = m_ControlBase.PageBase
    If page Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableSingle: page is not specified for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    Set ws = private_GetWorksheetByName(page, m_LayoutSheetName)
    If ws Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableSingle: sheet '" & m_LayoutSheetName & "' was not found for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    If m_TableItems Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableSingle: itemsSource is not resolved for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    If Not private_TryBuildRenderBufferSingle(valueBlock, styleSegments) Then Exit Sub
    If VBA.IsEmpty(valueBlock) Then Exit Sub

    Set targetRange = ws.Range( _
        ws.Cells(m_RowStart, m_ColStart), _
        ws.Cells(m_RowStart + UBound(valueBlock, 1) - 1, m_ColStart + UBound(valueBlock, 2) - 1))

    targetRange.Value2 = valueBlock

    If Not private_TryRegisterControlPartSegments(ws, styleSegments) Then Exit Sub

#If ENALBE_STYLES Then
    private_ApplyStyleSegments ws, styleSegments
#End If
End Sub

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(attrName))
        Case "itemssource"
            obj_IControl_SupportsAttribute = True
    End Select
End Function

' //
' // API
' //
Public Function Initialize(ByVal page As obj_IPage) As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    m_IsDisposed = False
    m_IsConfigured = False
    Set m_Page = page
    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Err.Clear
    Set m_ControlBase = Nothing
    Set m_TableItems = Nothing
    Set m_Page = Nothing
    On Error GoTo 0
End Sub

' (No public API yet.)
'
' //
' // Internal
' //
Private Function private_TryReadLayoutLongAttr( _
    ByVal controlNode As Object, _
    ByVal attrName As String, _
    ByRef outValue As Long, _
    ByVal isRequired As Boolean _
) As Boolean
    Dim rawText As String

    rawText = VBA.Trim$(ex_XmlCore.fn_NodeAttrText(controlNode, attrName))
    If VBA.Len(rawText) = 0 Then
        If isRequired Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "TableSingle: runtime layout attribute '" & attrName & "' is missing for control '" & m_ControlName & "'."
#End If
            Exit Function
        End If
        outValue = 0
        private_TryReadLayoutLongAttr = True
        Exit Function
    End If

    If Not VBA.IsNumeric(rawText) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableSingle: runtime layout attribute '" & attrName & "' must be numeric for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    outValue = VBA.CLng(rawText)
    private_TryReadLayoutLongAttr = True
End Function

Private Function private_TryBuildRenderBufferSingle(ByRef outValueBlock As Variant, ByRef outStyleSegments As Collection) As Boolean
    Dim tableDynamic As obj_TableDynamic
    Dim tableRows As list__obj_Row
    Dim tableColumnCount As Long
    Dim availableCols As Long
    Dim maxRows As Long
    Dim plannedRows As Long
    Dim currentOutputRow As Long
    Dim sourceRow As obj_Row
    Dim row As obj_Row
    Dim colOffset As Long
    Dim tokens As Variant
    Dim writeStart As Long
    Dim writeEnd As Long
    Dim rowIndex As Long

    availableCols = private_GetAvailableColumnCount()
    maxRows = m_RowEnd - m_RowStart + 1

    If availableCols <= 0 Or maxRows <= 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableSingle: invalid render bounds for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    If Not private_TryResolveSingleTableModel(tableDynamic) Then Exit Function
    If tableDynamic Is Nothing Then
        outValueBlock = Empty
        private_TryBuildRenderBufferSingle = True
        Exit Function
    End If

    tableColumnCount = tableDynamic.ColumnCount
    If tableColumnCount <= 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableSingle: table item has no columns."
#End If
        Exit Function
    End If
    If tableColumnCount > availableCols Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableSingle: control '" & m_ControlName & "' requires " & VBA.CStr(tableColumnCount) & _
               " columns, but span provides only " & VBA.CStr(availableCols) & "."
#End If
        Exit Function
    End If

    Set tableRows = tableDynamic.Rows
    plannedRows = 3
    If Not tableRows Is Nothing Then plannedRows = 2 + tableRows.Count + 1
    If plannedRows > maxRows Then plannedRows = maxRows

    If plannedRows <= 0 Then
        outValueBlock = Empty
        private_TryBuildRenderBufferSingle = True
        Exit Function
    End If

    ReDim outValueBlock(1 To plannedRows, 1 To availableCols)

#If ENALBE_STYLES Then
    Set outStyleSegments = New Collection
#End If

    currentOutputRow = 0

    currentOutputRow = currentOutputRow + 1
    outValueBlock(currentOutputRow, 1) = tableDynamic.SectionTitle
#If ENALBE_STYLES Then
    private_AddStyleSegment outStyleSegments, "section", tableColumnCount, currentOutputRow, currentOutputRow
#End If
    If currentOutputRow >= plannedRows Then
        private_TryBuildRenderBufferSingle = True
        Exit Function
    End If

    currentOutputRow = currentOutputRow + 1
    tokens = VBA.Split(tableDynamic.HeaderText, "|")
    For colOffset = 1 To tableColumnCount
        If colOffset - 1 <= UBound(tokens) Then
            outValueBlock(currentOutputRow, colOffset) = VBA.Trim$(VBA.CStr(tokens(colOffset - 1)))
        End If
    Next colOffset
#If ENALBE_STYLES Then
    private_AddStyleSegment outStyleSegments, "header", tableColumnCount, currentOutputRow, currentOutputRow
#End If
    If currentOutputRow >= plannedRows Then
        private_TryBuildRenderBufferSingle = True
        Exit Function
    End If

    writeStart = currentOutputRow + 1
    If Not tableRows Is Nothing Then
        For rowIndex = 1 To tableRows.Count
            If currentOutputRow >= plannedRows Then Exit For
            Set sourceRow = tableRows.Item(rowIndex)
            If sourceRow Is Nothing Then GoTo ContinueSourceRow
            currentOutputRow = currentOutputRow + 1
            Set row = sourceRow
            row.CopyToMatrixRow outValueBlock, currentOutputRow, tableColumnCount
ContinueSourceRow:
        Next rowIndex
    End If
    writeEnd = currentOutputRow
#If ENALBE_STYLES Then
    If writeEnd >= writeStart Then
        private_AddStyleSegment outStyleSegments, "data", tableColumnCount, writeStart, writeEnd
    End If
#End If
    If currentOutputRow >= plannedRows Then
        private_TryBuildRenderBufferSingle = True
        Exit Function
    End If

    currentOutputRow = currentOutputRow + 1
#If ENALBE_STYLES Then
    private_AddStyleSegment outStyleSegments, "spacer", tableColumnCount, currentOutputRow, currentOutputRow
#End If

    private_TryBuildRenderBufferSingle = True
End Function

Private Sub private_AddStyleSegment( _
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

    If styleSegments.Count > 0 Then
        Set lastSegment = styleSegments(styleSegments.Count)
        If VBA.StrComp(VBA.CStr(lastSegment("StyleKind")), styleKind, VBA.vbTextCompare) = 0 Then
            If VBA.CLng(lastSegment("ColumnCount")) = columnCount Then
                If VBA.CLng(lastSegment("RowEnd")) + 1 = rowStart Then
                    lastSegment("RowEnd") = rowEnd
                    Exit Sub
                End If
            End If
        End If
    End If

    Set segment = VBA.CreateObject("Scripting.Dictionary")
    segment.CompareMode = 1
    segment("StyleKind") = styleKind
    segment("ColumnCount") = columnCount
    segment("RowStart") = rowStart
    segment("RowEnd") = rowEnd
    styleSegments.Add segment
End Sub

Private Sub private_ApplyStyleSegments(ByVal ws As Worksheet, ByVal styleSegments As Collection)
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

    Set groupedRanges = VBA.CreateObject("Scripting.Dictionary")
    groupedRanges.CompareMode = 1

    For Each segment In styleSegments
        styleKind = VBA.LCase$(VBA.CStr(segment("StyleKind")))
        columnCount = VBA.CLng(segment("ColumnCount"))

        Set segmentRange = private_BuildSegmentRange( _
            ws, _
            VBA.CLng(segment("RowStart")), _
            VBA.CLng(segment("RowEnd")), _
            columnCount)
        If segmentRange Is Nothing Then GoTo ContinueSegment

        groupedKey = styleKind & "|" & VBA.CStr(columnCount)

        If groupedRanges.Exists(groupedKey) Then
            Set groupedRange = groupedRanges(groupedKey)
            Set groupedRanges(groupedKey) = Application.Union(groupedRange, segmentRange)
        Else
            groupedRanges.Add groupedKey, segmentRange
        End If

ContinueSegment:
    Next segment

    For Each key In groupedRanges.Keys
        sepPos = VBA.InStrRev(VBA.CStr(key), "|", -1, VBA.vbBinaryCompare)
        If sepPos <= 1 Then GoTo ContinueGroup

        styleKind = VBA.Left$(VBA.CStr(key), sepPos - 1)
        If Not private_TryResolveStylePreset(styleKind, backColor, fontColor, borderColor, fontSize, fontBold) Then Exit Sub

        Set groupedRange = groupedRanges(VBA.CStr(key))
        If groupedRange Is Nothing Then GoTo ContinueGroup

        private_ApplyRowStyle groupedRange, backColor, fontColor, borderColor, fontSize, fontBold

ContinueGroup:
    Next key
End Sub

Private Function private_TryRegisterControlPartSegments(ByVal ws As Worksheet, ByVal styleSegments As Collection) As Boolean
    Dim segment As Object
    Dim partName As String
    Dim segmentRange As Range

    If ws Is Nothing Then Exit Function
    If styleSegments Is Nothing Then
        private_TryRegisterControlPartSegments = True
        Exit Function
    End If

    For Each segment In styleSegments
        partName = private_MapStyleKindToControlPart(VBA.CStr(segment("StyleKind")))
        If VBA.Len(partName) = 0 Then GoTo ContinueSegment

        Set segmentRange = private_BuildSegmentRange( _
            ws, _
            VBA.CLng(segment("RowStart")), _
            VBA.CLng(segment("RowEnd")), _
            VBA.CLng(segment("ColumnCount")))
        If segmentRange Is Nothing Then GoTo ContinueSegment

        If Not ex_ControlPartsRuntime.fn_RegisterControlPart( _
            ws, _
            "table", _
            m_ControlName, _
            partName, _
            segmentRange) Then Exit Function

ContinueSegment:
    Next segment

    private_TryRegisterControlPartSegments = True
End Function

Private Function private_MapStyleKindToControlPart(ByVal styleKind As String) As String
    Select Case VBA.LCase$(VBA.Trim$(styleKind))
        Case "section"
            private_MapStyleKindToControlPart = "section"
        Case "header"
            private_MapStyleKindToControlPart = "header"
        Case "data"
            private_MapStyleKindToControlPart = "rows"
        Case "spacer"
            private_MapStyleKindToControlPart = "spacer"
    End Select
End Function

Private Function private_BuildSegmentRange( _
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

    Set private_BuildSegmentRange = ws.Range( _
        ws.Cells(absRowStart, m_ColStart), _
        ws.Cells(absRowEnd, m_ColStart + columnCount - 1))
End Function

Private Function private_TryResolveStylePreset( _
    ByVal styleKind As String, _
    ByRef backColor As Long, _
    ByRef fontColor As Long, _
    ByRef borderColor As Long, _
    ByRef fontSize As Double, _
    ByRef fontBold As Boolean _
) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(styleKind))
        Case "section"
            backColor = VBA.RGB(23, 58, 94)
            fontColor = VBA.RGB(234, 246, 255)
            borderColor = VBA.RGB(14, 34, 57)
            fontSize = 11
            fontBold = True

        Case "header"
            backColor = VBA.RGB(43, 74, 107)
            fontColor = VBA.RGB(221, 238, 255)
            borderColor = VBA.RGB(31, 54, 80)
            fontSize = 10
            fontBold = True

        Case "data"
            backColor = VBA.RGB(58, 58, 58)
            fontColor = VBA.RGB(240, 240, 240)
            borderColor = VBA.RGB(42, 42, 42)
            fontSize = 10
            fontBold = False

        Case "spacer"
            backColor = VBA.RGB(31, 31, 31)
            fontColor = VBA.RGB(31, 31, 31)
            borderColor = VBA.RGB(31, 31, 31)
            fontSize = 8
            fontBold = False

        Case Else
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "TableSingle: unsupported style segment kind '" & styleKind & "'."
#End If
            Exit Function
    End Select

    private_TryResolveStylePreset = True
End Function

Private Function private_GetAvailableColumnCount() As Long
    If m_ColEnd <= 0 Or m_ColStart <= 0 Then Exit Function
    private_GetAvailableColumnCount = m_ColEnd - m_ColStart + 1
End Function

Private Function private_TryResolveSingleTableModel(ByRef outTable As obj_TableDynamic) As Boolean
    Dim tableItem As Object

    If m_TableItems Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableSingle: itemsSource is not resolved for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If
    If m_TableItems.Count = 0 Then
        private_TryResolveSingleTableModel = True
        Exit Function
    End If

    If Not VBA.IsObject(m_TableItems(1)) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableSingle: itemsSource entry must be obj_TableDynamic or obj_Table."
#End If
        Exit Function
    End If

    Set tableItem = m_TableItems(1)
    If Not private_TryResolveTableModel(tableItem, outTable) Then Exit Function
    private_TryResolveSingleTableModel = True
End Function

Private Function private_TryResolveTableModel(ByVal tableItem As Variant, ByRef outTable As obj_TableDynamic) As Boolean
    Dim fixedTable As obj_Table

    If Not VBA.IsObject(tableItem) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableSingle: itemsSource entry must be obj_TableDynamic or obj_Table."
#End If
        Exit Function
    End If

    Select Case VBA.LCase$(VBA.TypeName(tableItem))
        Case "obj_tabledynamic"
            Set outTable = tableItem
            private_TryResolveTableModel = True

        Case "obj_table"
            Set fixedTable = tableItem
            Set outTable = private_ConvertFixedTableToDynamic(fixedTable)
            If outTable Is Nothing Then Exit Function
            private_TryResolveTableModel = True

        Case Else
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "TableSingle: unsupported table model type '" & VBA.TypeName(tableItem) & "'. Expected obj_TableDynamic or obj_Table."
#End If
    End Select
End Function

Private Function private_ConvertFixedTableToDynamic(ByVal fixedTable As obj_Table) As obj_TableDynamic
    Dim tableDynamic As obj_TableDynamic
    Dim sourceColumns As list__obj_Column
    Dim sourceRows As list__obj_Row
    Dim sourceColumn As obj_Column
    Dim sourceRow As obj_Row
    Dim targetColumn As obj_Column
    Dim targetRow As obj_Row
    Dim colIndex As Long
    Dim sourceColumnIndex As Long
    Dim sourceRowIndex As Long

    If fixedTable Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableSingle: fixed table model is not specified."
#End If
        Exit Function
    End If

    Set tableDynamic = New obj_TableDynamic
    tableDynamic.SectionTitle = fixedTable.SectionTitle

    Set sourceColumns = fixedTable.Columns
    For sourceColumnIndex = 1 To sourceColumns.Count
        Set sourceColumn = sourceColumns.Item(sourceColumnIndex)
        If sourceColumn Is Nothing Then GoTo ContinueSourceColumn
        Set targetColumn = New obj_Column
        targetColumn.Position = sourceColumn.Position
        targetColumn.Name = sourceColumn.Name
        If Not tableDynamic.AddColumn(targetColumn) Then Exit Function
ContinueSourceColumn:
    Next sourceColumnIndex

    Set sourceRows = fixedTable.Rows
    For sourceRowIndex = 1 To sourceRows.Count
        Set sourceRow = sourceRows.Item(sourceRowIndex)
        If sourceRow Is Nothing Then GoTo ContinueSourceRowInFixedTable
        Set targetRow = New obj_Row
        For colIndex = 1 To tableDynamic.ColumnCount
            targetRow.AddCell sourceRow.GetCell(colIndex)
        Next colIndex
        If Not tableDynamic.AddRow(targetRow) Then Exit Function
ContinueSourceRowInFixedTable:
    Next sourceRowIndex

    Set private_ConvertFixedTableToDynamic = tableDynamic
End Function

Private Sub private_ApplyRowStyle( _
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

Private Function private_GetWorksheetByName(ByVal page As obj_PageBase, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    If page Is Nothing Then Exit Function
    Set ws = page.Worksheet
    If ws Is Nothing Then Exit Function

    sheetName = VBA.LCase$(VBA.Trim$(sheetName))
    If VBA.Len(sheetName) > 0 Then
        If VBA.StrComp(VBA.LCase$(VBA.Trim$(ws.Name)), sheetName, VBA.vbTextCompare) <> 0 Then Exit Function
    End If

    Set private_GetWorksheetByName = ws
End Function
