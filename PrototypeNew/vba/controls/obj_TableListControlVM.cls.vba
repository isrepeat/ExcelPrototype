VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_TableListControlVM"
Option Explicit
Implements obj_IControl

#Const ENALBE_STYLES = True
#Const ENALBE_BORDERS = True

Private m_Base As obj_ControlBase
Private m_ControlName As String
Private m_ItemsSourceRaw As String
Private m_ItemVisibilityRaw As String
Private m_LayoutSheetName As String
Private m_RowStart As Long
Private m_ColStart As Long
Private m_RowEnd As Long
Private m_ColEnd As Long
Private m_TableItems As Collection
Private m_IsConfigured As Boolean

' //
' // Interface
' //
Private Sub obj_IControl_Configure(ByVal page As obj_PageBase, ByVal controlNode As Object)
    Dim currentPage As obj_PageBase

    m_IsConfigured = False
    Set m_TableItems = Nothing
    Set m_Base = Nothing

    Set m_Base = New obj_ControlBase
    If Not m_Base.Configure(page, controlNode, "TableList", "tablelist", m_ControlName) Then Exit Sub

    m_ItemsSourceRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "itemsSource")))
    If VBA.Len(m_ItemsSourceRaw) = 0 Then
        VBA.MsgBox "TableList: itemsSource is not specified for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    m_ItemVisibilityRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "itemVisibility")))

    m_LayoutSheetName = VBA.Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "__layoutSheetName"))
    If VBA.Len(m_LayoutSheetName) = 0 Then
        VBA.MsgBox "TableList: runtime layout sheet is missing for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    If Not private_TryReadLayoutLongAttr(controlNode, "__layoutRowStart", m_RowStart, True) Then Exit Sub
    If Not private_TryReadLayoutLongAttr(controlNode, "__layoutColStart", m_ColStart, True) Then Exit Sub
    If Not private_TryReadLayoutLongAttr(controlNode, "__layoutRowEnd", m_RowEnd, True) Then Exit Sub
    If Not private_TryReadLayoutLongAttr(controlNode, "__layoutColEnd", m_ColEnd, True) Then Exit Sub

    If m_RowStart <= 0 Or m_ColStart <= 0 Then
        VBA.MsgBox "TableList: invalid row/column start for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    If m_RowEnd < m_RowStart Then
        VBA.MsgBox "TableList: control '" & m_ControlName & "' has invalid spanRows range.", VBA.vbExclamation
        Exit Sub
    End If

    If m_ColEnd < m_ColStart Then
        VBA.MsgBox "TableList: control '" & m_ControlName & "' has invalid spanCells range.", VBA.vbExclamation
        Exit Sub
    End If

    Set currentPage = m_Base.PageBase
    If currentPage Is Nothing Then Exit Sub
    If Not ex_RuntimeSourceResolver.m_TryResolveItemsSource(currentPage.RuntimeSources, m_ItemsSourceRaw, m_TableItems) Then Exit Sub
    If Not private_TryApplyItemVisibilityFilter(m_TableItems) Then Exit Sub

    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render()
    Dim ws As Worksheet
    Dim valueBlock As Variant
    Dim targetRange As Range
    Dim styleSegments As Collection
    Dim page As obj_PageBase

    If Not m_IsConfigured Then
        VBA.MsgBox "TableList: control '" & m_ControlName & "' is not configured.", VBA.vbExclamation
        Exit Sub
    End If

    Set page = Nothing
    If Not m_Base Is Nothing Then Set page = m_Base.PageBase
    If page Is Nothing Then
        VBA.MsgBox "TableList: page is not specified for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    Set ws = private_GetWorksheetByName(page, m_LayoutSheetName)
    If ws Is Nothing Then
        VBA.MsgBox "TableList: sheet '" & m_LayoutSheetName & "' was not found for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    If m_TableItems Is Nothing Then
        VBA.MsgBox "TableList: itemsSource is not resolved for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    ' Build in-memory first, then write once to minimize COM overhead.
    If Not private_TryBuildRenderBuffer(valueBlock, styleSegments) Then Exit Sub
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
        Case "itemssource", "itemvisibility"
            obj_IControl_SupportsAttribute = True
    End Select
End Function

' //
' // API
' //
' (No public API yet.)
'
' //
' // Internal
' //
Private Function private_TryApplyItemVisibilityFilter(ByRef tableItems As Collection) As Boolean
    Dim filteredItems As Collection
    Dim tableItem As Variant
    Dim isVisible As Boolean

    If VBA.Len(m_ItemVisibilityRaw) = 0 Then
        private_TryApplyItemVisibilityFilter = True
        Exit Function
    End If

    If tableItems Is Nothing Then
        VBA.MsgBox "TableList: itemsSource is not resolved for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Function
    End If

    Set filteredItems = New Collection

    For Each tableItem In tableItems
        If Not VBA.IsObject(tableItem) Then
            VBA.MsgBox "TableList: itemsSource entry must be an object for itemVisibility evaluation in control '" & m_ControlName & "'.", VBA.vbExclamation
            Exit Function
        End If

        If Not ex_BindingRuntime.m_TryResolveVisibilityBinding(m_ItemVisibilityRaw, tableItem, isVisible) Then Exit Function
        If isVisible Then filteredItems.Add tableItem
    Next tableItem

    Set tableItems = filteredItems
    private_TryApplyItemVisibilityFilter = True
End Function

Private Function private_TryReadLayoutLongAttr( _
    ByVal controlNode As Object, _
    ByVal attrName As String, _
    ByRef outValue As Long, _
    ByVal isRequired As Boolean _
) As Boolean
    Dim rawText As String

    rawText = VBA.Trim$(ex_XmlCore.m_NodeAttrText(controlNode, attrName))
    If VBA.Len(rawText) = 0 Then
        If isRequired Then
            VBA.MsgBox "TableList: runtime layout attribute '" & attrName & "' is missing for control '" & m_ControlName & "'.", VBA.vbExclamation
            Exit Function
        End If

        outValue = 0
        private_TryReadLayoutLongAttr = True
        Exit Function
    End If

    If Not VBA.IsNumeric(rawText) Then
        VBA.MsgBox "TableList: runtime layout attribute '" & attrName & "' must be numeric for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Function
    End If

    outValue = VBA.CLng(rawText)
    private_TryReadLayoutLongAttr = True
End Function

Private Function private_TryBuildRenderBuffer(ByRef outValueBlock As Variant, ByRef outStyleSegments As Collection) As Boolean
    Dim tableItem As Variant
    Dim tableView As obj_TableViewItem
    Dim availableCols As Long
    Dim maxRows As Long
    Dim plannedRows As Long
    Dim rowsForItem As Long
    Dim currentOutputRow As Long

    availableCols = private_GetAvailableColumnCount()
    maxRows = m_RowEnd - m_RowStart + 1

    If availableCols <= 0 Or maxRows <= 0 Then
        VBA.MsgBox "TableList: invalid render bounds for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Function
    End If

    plannedRows = 0

    ' Pass 1: estimate output size up-front to allocate matrix once.
    For Each tableItem In m_TableItems
        If plannedRows >= maxRows Then Exit For

        Set tableView = Nothing
        If Not private_TryResolveTableViewItem(tableItem, tableView) Then Exit Function
        If tableView Is Nothing Then GoTo ContinueEstimate

        If Not private_TryEstimateTableOutputRows(tableView, availableCols, rowsForItem) Then Exit Function
        plannedRows = plannedRows + rowsForItem
        If plannedRows > maxRows Then
            plannedRows = maxRows
            Exit For
        End If

ContinueEstimate:
    Next tableItem

    If plannedRows = 0 Then
        outValueBlock = Empty
        private_TryBuildRenderBuffer = True
        Exit Function
    End If

    ReDim outValueBlock(1 To plannedRows, 1 To availableCols)

#If ENALBE_STYLES Then
    Set outStyleSegments = New Collection
#End If

    currentOutputRow = 0

    ' Pass 2: fill matrix sequentially.
    For Each tableItem In m_TableItems
        If currentOutputRow >= plannedRows Then Exit For

        Set tableView = Nothing
        If Not private_TryResolveTableViewItem(tableItem, tableView) Then Exit Function
        If tableView Is Nothing Then GoTo ContinueWrite

        If Not private_TryWriteTableItemToBuffer( _
            tableView, outValueBlock, outStyleSegments, availableCols, plannedRows, currentOutputRow) Then Exit Function

ContinueWrite:
    Next tableItem

    private_TryBuildRenderBuffer = True
End Function

Private Function private_TryEstimateTableOutputRows( _
    ByVal tableView As obj_TableViewItem, _
    ByVal availableCols As Long, _
    ByRef outRows As Long _
) As Boolean
    Dim tableModel As obj_TableDynamic
    Dim rowItems As Collection
    Dim rowItemRaw As Variant
    Dim rowView As obj_RowViewItem

    outRows = 0

    If tableView Is Nothing Then
        private_TryEstimateTableOutputRows = True
        Exit Function
    End If

    If Not tableView.IsVisible() Then
        private_TryEstimateTableOutputRows = True
        Exit Function
    End If

    Set tableModel = tableView.Model
    If tableModel Is Nothing Then
        VBA.MsgBox "TableList: table view has no model.", VBA.vbExclamation
        Exit Function
    End If

    If tableModel.ColumnCount <= 0 Then
        VBA.MsgBox "TableList: table item has no columns.", VBA.vbExclamation
        Exit Function
    End If

    If tableModel.ColumnCount > availableCols Then
        VBA.MsgBox "TableList: control '" & m_ControlName & "' requires " & VBA.CStr(tableModel.ColumnCount) & _
               " columns, but span provides only " & VBA.CStr(availableCols) & ".", VBA.vbExclamation
        Exit Function
    End If

    outRows = outRows + private_GetBannerRenderRows(tableView.Banner)

    ' section + header
    outRows = outRows + 2

    Set rowItems = tableView.RowItems
    If Not rowItems Is Nothing And rowItems.Count > 0 Then
        For Each rowItemRaw In rowItems
            Set rowView = Nothing
            If Not private_TryResolveRowViewItem(rowItemRaw, rowView) Then Exit Function
            If rowView Is Nothing Then GoTo ContinueRowEstimate
            If Not rowView.IsVisible() Then GoTo ContinueRowEstimate

            outRows = outRows + private_GetBannerRenderRows(rowView.Banner)
            outRows = outRows + 1
            outRows = outRows + rowView.SpacerRowsAfter

ContinueRowEstimate:
        Next rowItemRaw
    Else
        outRows = outRows + tableModel.RowCount
    End If

    ' Spacer after table
    outRows = outRows + 1

    private_TryEstimateTableOutputRows = True
End Function

Private Function private_TryWriteTableItemToBuffer( _
    ByVal tableView As obj_TableViewItem, _
    ByRef valueBlock As Variant, _
    ByVal styleSegments As Collection, _
    ByVal availableCols As Long, _
    ByVal plannedRows As Long, _
    ByRef ioCurrentOutputRow As Long _
) As Boolean
    Dim tableModel As obj_TableDynamic
    Dim rowItems As Collection
    Dim rowItemRaw As Variant
    Dim rowView As obj_RowViewItem
    Dim tableRows As Collection
    Dim rowRaw As Variant
    Dim rowModel As obj_Row
    Dim colOffset As Long
    Dim tokens As Variant
    Dim writeStart As Long
    Dim writeEnd As Long

    If tableView Is Nothing Then
        private_TryWriteTableItemToBuffer = True
        Exit Function
    End If

    If Not tableView.IsVisible() Then
        private_TryWriteTableItemToBuffer = True
        Exit Function
    End If

    Set tableModel = tableView.Model
    If tableModel Is Nothing Then
        VBA.MsgBox "TableList: table view has no model.", VBA.vbExclamation
        Exit Function
    End If

    If tableModel.ColumnCount <= 0 Then
        VBA.MsgBox "TableList: table item has no columns.", VBA.vbExclamation
        Exit Function
    End If

    If tableModel.ColumnCount > availableCols Then
        VBA.MsgBox "TableList: control '" & m_ControlName & "' requires " & VBA.CStr(tableModel.ColumnCount) & _
               " columns, but span provides only " & VBA.CStr(availableCols) & ".", VBA.vbExclamation
        Exit Function
    End If

    If Not private_TryAppendBannerBlock( _
        tableView.Banner, "tablebanner", tableModel.ColumnCount, valueBlock, styleSegments, plannedRows, ioCurrentOutputRow) Then Exit Function
    If ioCurrentOutputRow >= plannedRows Then
        private_TryWriteTableItemToBuffer = True
        Exit Function
    End If

    ' Section row
    ioCurrentOutputRow = ioCurrentOutputRow + 1
    valueBlock(ioCurrentOutputRow, 1) = tableModel.SectionTitle
#If ENALBE_STYLES Then
    private_AddStyleSegment styleSegments, "section", tableModel.ColumnCount, ioCurrentOutputRow, ioCurrentOutputRow
#End If
    If ioCurrentOutputRow >= plannedRows Then
        private_TryWriteTableItemToBuffer = True
        Exit Function
    End If

    ' Header row
    ioCurrentOutputRow = ioCurrentOutputRow + 1
    tokens = VBA.Split(tableModel.HeaderText, "|")
    For colOffset = 1 To tableModel.ColumnCount
        If colOffset - 1 <= UBound(tokens) Then
            valueBlock(ioCurrentOutputRow, colOffset) = VBA.Trim$(VBA.CStr(tokens(colOffset - 1)))
        End If
    Next colOffset
#If ENALBE_STYLES Then
    private_AddStyleSegment styleSegments, "header", tableModel.ColumnCount, ioCurrentOutputRow, ioCurrentOutputRow
#End If
    If ioCurrentOutputRow >= plannedRows Then
        private_TryWriteTableItemToBuffer = True
        Exit Function
    End If

    Set rowItems = tableView.RowItems

    If Not rowItems Is Nothing And rowItems.Count > 0 Then
        For Each rowItemRaw In rowItems
            If ioCurrentOutputRow >= plannedRows Then Exit For

            Set rowView = Nothing
            If Not private_TryResolveRowViewItem(rowItemRaw, rowView) Then Exit Function
            If rowView Is Nothing Then GoTo ContinueRowView

            If Not private_TryAppendRowViewData( _
                rowView, tableModel.ColumnCount, valueBlock, styleSegments, plannedRows, ioCurrentOutputRow) Then Exit Function

ContinueRowView:
        Next rowItemRaw
    Else
        Set tableRows = tableModel.Rows
        If Not tableRows Is Nothing Then
            writeStart = ioCurrentOutputRow + 1
            On Error GoTo EH_INVALID_ROW
            For Each rowRaw In tableRows
                If ioCurrentOutputRow >= plannedRows Then Exit For

                ioCurrentOutputRow = ioCurrentOutputRow + 1
                Set rowModel = rowRaw
                rowModel.CopyToMatrixRow valueBlock, ioCurrentOutputRow, tableModel.ColumnCount
            Next rowRaw
            On Error GoTo 0
            writeEnd = ioCurrentOutputRow
#If ENALBE_STYLES Then
            If writeEnd >= writeStart Then
                private_AddStyleSegment styleSegments, "data", tableModel.ColumnCount, writeStart, writeEnd
            End If
#End If
        End If
    End If

    If ioCurrentOutputRow >= plannedRows Then
        private_TryWriteTableItemToBuffer = True
        Exit Function
    End If

    ' Spacer row
    ioCurrentOutputRow = ioCurrentOutputRow + 1
#If ENALBE_STYLES Then
    private_AddStyleSegment styleSegments, "spacer", tableModel.ColumnCount, ioCurrentOutputRow, ioCurrentOutputRow
#End If

    private_TryWriteTableItemToBuffer = True
    Exit Function

EH_INVALID_ROW:
    On Error GoTo 0
    VBA.MsgBox "TableList: unsupported row object in table rows. Expected obj_Row.", VBA.vbExclamation
End Function

Private Function private_TryAppendRowViewData( _
    ByVal rowView As obj_RowViewItem, _
    ByVal columnCount As Long, _
    ByRef valueBlock As Variant, _
    ByVal styleSegments As Collection, _
    ByVal plannedRows As Long, _
    ByRef ioCurrentOutputRow As Long _
) As Boolean
    Dim rowModel As obj_Row
    Dim spacerIndex As Long

    If rowView Is Nothing Then
        private_TryAppendRowViewData = True
        Exit Function
    End If

    If Not rowView.IsVisible() Then
        private_TryAppendRowViewData = True
        Exit Function
    End If

    If Not private_TryAppendBannerBlock( _
        rowView.Banner, "rowbanner", columnCount, valueBlock, styleSegments, plannedRows, ioCurrentOutputRow) Then Exit Function

    If ioCurrentOutputRow >= plannedRows Then
        private_TryAppendRowViewData = True
        Exit Function
    End If

    Set rowModel = rowView.Row
    If rowModel Is Nothing Then
        VBA.MsgBox "TableList: row view item has no row model.", VBA.vbExclamation
        Exit Function
    End If

    ioCurrentOutputRow = ioCurrentOutputRow + 1
    rowModel.CopyToMatrixRow valueBlock, ioCurrentOutputRow, columnCount
#If ENALBE_STYLES Then
    private_AddStyleSegment styleSegments, "data", columnCount, ioCurrentOutputRow, ioCurrentOutputRow
#End If

    For spacerIndex = 1 To rowView.SpacerRowsAfter
        If ioCurrentOutputRow >= plannedRows Then Exit For

        ioCurrentOutputRow = ioCurrentOutputRow + 1
#If ENALBE_STYLES Then
        private_AddStyleSegment styleSegments, "spacer", columnCount, ioCurrentOutputRow, ioCurrentOutputRow
#End If
    Next spacerIndex

    private_TryAppendRowViewData = True
End Function

Private Function private_TryAppendBannerBlock( _
    ByVal bannerView As obj_BannerViewItem, _
    ByVal styleKind As String, _
    ByVal columnCount As Long, _
    ByRef valueBlock As Variant, _
    ByVal styleSegments As Collection, _
    ByVal plannedRows As Long, _
    ByRef ioCurrentOutputRow As Long _
) As Boolean
    Dim bannerRows As Long
    Dim writeStart As Long
    Dim writeEnd As Long
    Dim bannerModel As obj_Banner
    Dim rowOffset As Long

    bannerRows = private_GetBannerRenderRows(bannerView)
    If bannerRows <= 0 Then
        private_TryAppendBannerBlock = True
        Exit Function
    End If

    Set bannerModel = bannerView.Model
    writeStart = ioCurrentOutputRow + 1

    For rowOffset = 1 To bannerRows
        If ioCurrentOutputRow >= plannedRows Then Exit For

        ioCurrentOutputRow = ioCurrentOutputRow + 1

        If Not bannerModel Is Nothing Then
            If rowOffset = 1 Then
                valueBlock(ioCurrentOutputRow, 1) = bannerModel.Header
            ElseIf rowOffset = 2 Then
                valueBlock(ioCurrentOutputRow, 1) = bannerModel.Message
            End If
        End If
    Next rowOffset

    writeEnd = ioCurrentOutputRow
#If ENALBE_STYLES Then
    If writeEnd >= writeStart Then
        private_AddStyleSegment styleSegments, styleKind, columnCount, writeStart, writeEnd
    End If
#End If

    private_TryAppendBannerBlock = True
End Function

Private Function private_GetBannerRenderRows(ByVal bannerView As obj_BannerViewItem) As Long
    Dim spanRows As Long

    If bannerView Is Nothing Then Exit Function
    If Not bannerView.IsVisible() Then Exit Function

    If Not bannerView.Presentation Is Nothing Then
        spanRows = bannerView.Presentation.SpanRows
    End If

    If spanRows <= 0 Then spanRows = 2
    private_GetBannerRenderRows = spanRows
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

    ' Merge adjacent segments with same style+width to reduce style operations later.
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

    ' Group by (style kind + column count) and union ranges.
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

        If Not ex_ControlPartsRuntime.m_RegisterControlPart( _
            ws, _
            "tablelist", _
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
        Case "tablebanner"
            private_MapStyleKindToControlPart = "itembanner"
        Case "rowbanner"
            private_MapStyleKindToControlPart = "rowbanner"
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

        Case "tablebanner"
            backColor = VBA.RGB(45, 74, 104)
            fontColor = VBA.RGB(245, 251, 255)
            borderColor = VBA.RGB(26, 43, 61)
            fontSize = 10
            fontBold = True

        Case "rowbanner"
            backColor = VBA.RGB(52, 86, 118)
            fontColor = VBA.RGB(240, 248, 255)
            borderColor = VBA.RGB(33, 57, 82)
            fontSize = 10
            fontBold = False

        Case Else
            VBA.MsgBox "TableList: unsupported style segment kind '" & styleKind & "'.", VBA.vbExclamation
            Exit Function
    End Select

    private_TryResolveStylePreset = True
End Function

Private Function private_GetAvailableColumnCount() As Long
    If m_ColEnd <= 0 Or m_ColStart <= 0 Then Exit Function
    private_GetAvailableColumnCount = m_ColEnd - m_ColStart + 1
End Function

Private Function private_TryResolveTableViewItem(ByVal rawItem As Variant, ByRef outTableView As obj_TableViewItem) As Boolean
    Dim tableModel As obj_TableDynamic

    If Not VBA.IsObject(rawItem) Then
        VBA.MsgBox "TableList: itemsSource entry must be an object.", VBA.vbExclamation
        Exit Function
    End If

    Select Case VBA.LCase$(VBA.TypeName(rawItem))
        Case "obj_tableviewitem"
            Set outTableView = rawItem
            private_TryResolveTableViewItem = True

        Case "obj_tabledynamic", "obj_table"
            If Not private_TryResolveTableModelFromAny(rawItem, tableModel) Then Exit Function
            Set outTableView = private_CreateTableViewFromModel(tableModel)
            If outTableView Is Nothing Then Exit Function
            private_TryResolveTableViewItem = True

        Case Else
            VBA.MsgBox "TableList: unsupported item type '" & VBA.TypeName(rawItem) & _
                   "'. Expected obj_TableViewItem, obj_TableDynamic or obj_Table.", VBA.vbExclamation
    End Select
End Function

Private Function private_TryResolveTableModelFromAny(ByVal tableItem As Variant, ByRef outTable As obj_TableDynamic) As Boolean
    Dim fixedTable As obj_Table

    If Not VBA.IsObject(tableItem) Then
        VBA.MsgBox "TableList: itemsSource entry must be an object of type obj_TableDynamic or obj_Table.", VBA.vbExclamation
        Exit Function
    End If

    Select Case VBA.LCase$(VBA.TypeName(tableItem))
        Case "obj_tabledynamic"
            Set outTable = tableItem
            private_TryResolveTableModelFromAny = True

        Case "obj_table"
            Set fixedTable = tableItem
            Set outTable = private_ConvertFixedTableToDynamic(fixedTable)
            If outTable Is Nothing Then Exit Function
            private_TryResolveTableModelFromAny = True

        Case Else
            VBA.MsgBox "TableList: unsupported table model type '" & VBA.TypeName(tableItem) & _
                   "'. Expected obj_TableDynamic or obj_Table.", VBA.vbExclamation
    End Select
End Function

Private Function private_CreateTableViewFromModel(ByVal tableModel As obj_TableDynamic) As obj_TableViewItem
    Dim tableView As obj_TableViewItem

    If tableModel Is Nothing Then
        VBA.MsgBox "TableList: table model is not specified.", VBA.vbExclamation
        Exit Function
    End If

    Set tableView = New obj_TableViewItem
    Set tableView.Model = tableModel
    tableView.ItemVisible = True

    Set private_CreateTableViewFromModel = tableView
End Function

Private Function private_TryResolveRowViewItem(ByVal rawItem As Variant, ByRef outRowView As obj_RowViewItem) As Boolean
    Dim rowModel As obj_Row

    If Not VBA.IsObject(rawItem) Then
        VBA.MsgBox "TableList: row item must be an object.", VBA.vbExclamation
        Exit Function
    End If

    Select Case VBA.LCase$(VBA.TypeName(rawItem))
        Case "obj_rowviewitem"
            Set outRowView = rawItem
            private_TryResolveRowViewItem = True

        Case "obj_row"
            Set rowModel = rawItem
            Set outRowView = private_CreateRowViewFromModel(rowModel)
            If outRowView Is Nothing Then Exit Function
            private_TryResolveRowViewItem = True

        Case Else
            VBA.MsgBox "TableList: unsupported row item type '" & VBA.TypeName(rawItem) & _
                   "'. Expected obj_RowViewItem or obj_Row.", VBA.vbExclamation
    End Select
End Function

Private Function private_CreateRowViewFromModel(ByVal rowModel As obj_Row) As obj_RowViewItem
    Dim rowView As obj_RowViewItem

    If rowModel Is Nothing Then
        VBA.MsgBox "TableList: row model is not specified.", VBA.vbExclamation
        Exit Function
    End If

    Set rowView = New obj_RowViewItem
    Set rowView.Row = rowModel
    rowView.RowVisible = True

    Set private_CreateRowViewFromModel = rowView
End Function

Private Function private_ConvertFixedTableToDynamic(ByVal fixedTable As obj_Table) As obj_TableDynamic
    Dim dynamicTable As obj_TableDynamic
    Dim sourceColumns As Collection
    Dim sourceRows As Collection
    Dim sourceColumn As obj_Column
    Dim sourceRow As obj_Row
    Dim targetColumn As obj_Column
    Dim targetRow As obj_Row
    Dim colIndex As Long

    If fixedTable Is Nothing Then
        VBA.MsgBox "TableList: fixed table model is not specified.", VBA.vbExclamation
        Exit Function
    End If

    Set dynamicTable = New obj_TableDynamic
    dynamicTable.SectionTitle = fixedTable.SectionTitle

    Set sourceColumns = fixedTable.Columns
    For Each sourceColumn In sourceColumns
        Set targetColumn = New obj_Column
        targetColumn.Position = sourceColumn.Position
        targetColumn.Name = sourceColumn.Name
        If Not dynamicTable.AddColumn(targetColumn) Then Exit Function
    Next sourceColumn

    Set sourceRows = fixedTable.Rows
    For Each sourceRow In sourceRows
        Set targetRow = New obj_Row
        For colIndex = 1 To dynamicTable.ColumnCount
            targetRow.AddCell sourceRow.GetCell(colIndex)
        Next colIndex
        If Not dynamicTable.AddRow(targetRow) Then Exit Function
    Next sourceRow

    Set private_ConvertFixedTableToDynamic = dynamicTable
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
