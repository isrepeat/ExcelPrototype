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

' //
' // Interface
' //
Private Sub obj_IControl_Configure(ByVal controlNode As Object)
    m_IsConfigured = False
    Set m_TableItems = Nothing

    If controlNode Is Nothing Then
        MsgBox "TableList: control node is not specified.", vbExclamation
        Exit Sub
    End If

    m_ControlName = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "name"))
    If Len(m_ControlName) = 0 Then m_ControlName = "tablelist"

    m_ItemsSourceRaw = Trim$(CStr(ex_XmlCore.m_NodeAttrText(controlNode, "itemsSource")))
    If Len(m_ItemsSourceRaw) = 0 Then
        MsgBox "TableList: itemsSource is not specified for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    m_ItemVisibilityRaw = Trim$(CStr(ex_XmlCore.m_NodeAttrText(controlNode, "itemVisibility")))

    m_LayoutSheet = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "__layoutSheet"))
    If Len(m_LayoutSheet) = 0 Then
        MsgBox "TableList: runtime layout sheet is missing for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    If Not mp_TryReadLayoutLongAttr(controlNode, "__layoutRowStart", m_RowStart, True) Then Exit Sub
    If Not mp_TryReadLayoutLongAttr(controlNode, "__layoutColStart", m_ColStart, True) Then Exit Sub
    If Not mp_TryReadLayoutLongAttr(controlNode, "__layoutRowEnd", m_RowEnd, True) Then Exit Sub
    If Not mp_TryReadLayoutLongAttr(controlNode, "__layoutColEnd", m_ColEnd, True) Then Exit Sub

    If m_RowStart <= 0 Or m_ColStart <= 0 Then
        MsgBox "TableList: invalid row/column start for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    If m_RowEnd < m_RowStart Then
        MsgBox "TableList: control '" & m_ControlName & "' has invalid spanRows range.", vbExclamation
        Exit Sub
    End If

    If m_ColEnd < m_ColStart Then
        MsgBox "TableList: control '" & m_ControlName & "' has invalid spanCells range.", vbExclamation
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
        MsgBox "TableList: control '" & m_ControlName & "' is not configured.", vbExclamation
        Exit Sub
    End If

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then
        MsgBox "TableList: workbook is not specified for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    Set ws = mp_GetWorksheetByName(wb, m_LayoutSheet)
    If ws Is Nothing Then
        MsgBox "TableList: sheet '" & m_LayoutSheet & "' was not found for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    If m_TableItems Is Nothing Then
        MsgBox "TableList: itemsSource is not resolved for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    ' Build in-memory first, then write once to minimize COM overhead.
    If Not mp_TryBuildRenderBuffer(valueBlock, styleSegments) Then Exit Sub
    If IsEmpty(valueBlock) Then Exit Sub

    Set targetRange = ws.Range( _
        ws.Cells(m_RowStart, m_ColStart), _
        ws.Cells(m_RowStart + UBound(valueBlock, 1) - 1, m_ColStart + UBound(valueBlock, 2) - 1))

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

' //
' // Internal
' //
Private Function mp_TryApplyItemVisibilityFilter(ByRef tableItems As Collection) As Boolean
    Dim filteredItems As Collection
    Dim tableItem As Variant
    Dim isVisible As Boolean

    If Len(m_ItemVisibilityRaw) = 0 Then
        mp_TryApplyItemVisibilityFilter = True
        Exit Function
    End If

    If tableItems Is Nothing Then
        MsgBox "TableList: itemsSource is not resolved for control '" & m_ControlName & "'.", vbExclamation
        Exit Function
    End If

    Set filteredItems = New Collection

    For Each tableItem In tableItems
        If Not IsObject(tableItem) Then
            MsgBox "TableList: itemsSource entry must be an object for itemVisibility evaluation in control '" & m_ControlName & "'.", vbExclamation
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
            MsgBox "TableList: runtime layout attribute '" & attrName & "' is missing for control '" & m_ControlName & "'.", vbExclamation
            Exit Function
        End If

        outValue = 0
        mp_TryReadLayoutLongAttr = True
        Exit Function
    End If

    If Not IsNumeric(rawText) Then
        MsgBox "TableList: runtime layout attribute '" & attrName & "' must be numeric for control '" & m_ControlName & "'.", vbExclamation
        Exit Function
    End If

    outValue = CLng(rawText)
    mp_TryReadLayoutLongAttr = True
End Function

Private Function mp_TryBuildRenderBuffer(ByRef outValueBlock As Variant, ByRef outStyleSegments As Collection) As Boolean
    Dim tableItem As Variant
    Dim tableView As obj_TableViewItem
    Dim availableCols As Long
    Dim maxRows As Long
    Dim plannedRows As Long
    Dim rowsForItem As Long
    Dim currentOutputRow As Long

    availableCols = mp_GetAvailableColumnCount()
    maxRows = m_RowEnd - m_RowStart + 1

    If availableCols <= 0 Or maxRows <= 0 Then
        MsgBox "TableList: invalid render bounds for control '" & m_ControlName & "'.", vbExclamation
        Exit Function
    End If

    plannedRows = 0

    ' Pass 1: estimate output size up-front to allocate matrix once.
    For Each tableItem In m_TableItems
        If plannedRows >= maxRows Then Exit For

        Set tableView = Nothing
        If Not mp_TryResolveTableViewItem(tableItem, tableView) Then Exit Function
        If tableView Is Nothing Then GoTo ContinueEstimate

        If Not mp_TryEstimateTableOutputRows(tableView, availableCols, rowsForItem) Then Exit Function
        plannedRows = plannedRows + rowsForItem
        If plannedRows > maxRows Then
            plannedRows = maxRows
            Exit For
        End If

ContinueEstimate:
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

    ' Pass 2: fill matrix sequentially.
    For Each tableItem In m_TableItems
        If currentOutputRow >= plannedRows Then Exit For

        Set tableView = Nothing
        If Not mp_TryResolveTableViewItem(tableItem, tableView) Then Exit Function
        If tableView Is Nothing Then GoTo ContinueWrite

        If Not mp_TryWriteTableItemToBuffer( _
            tableView, outValueBlock, outStyleSegments, availableCols, plannedRows, currentOutputRow) Then Exit Function

ContinueWrite:
    Next tableItem

    mp_TryBuildRenderBuffer = True
End Function

Private Function mp_TryEstimateTableOutputRows( _
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
        mp_TryEstimateTableOutputRows = True
        Exit Function
    End If

    If Not tableView.m_IsVisible() Then
        mp_TryEstimateTableOutputRows = True
        Exit Function
    End If

    Set tableModel = tableView.Model
    If tableModel Is Nothing Then
        MsgBox "TableList: table view has no model.", vbExclamation
        Exit Function
    End If

    If tableModel.ColumnCount <= 0 Then
        MsgBox "TableList: table item has no columns.", vbExclamation
        Exit Function
    End If

    If tableModel.ColumnCount > availableCols Then
        MsgBox "TableList: control '" & m_ControlName & "' requires " & CStr(tableModel.ColumnCount) & _
               " columns, but span provides only " & CStr(availableCols) & ".", vbExclamation
        Exit Function
    End If

    outRows = outRows + mp_GetBannerRenderRows(tableView.Banner)

    ' section + header
    outRows = outRows + 2

    Set rowItems = tableView.RowItems
    If Not rowItems Is Nothing And rowItems.Count > 0 Then
        For Each rowItemRaw In rowItems
            Set rowView = Nothing
            If Not mp_TryResolveRowViewItem(rowItemRaw, rowView) Then Exit Function
            If rowView Is Nothing Then GoTo ContinueRowEstimate
            If Not rowView.m_IsVisible() Then GoTo ContinueRowEstimate

            outRows = outRows + mp_GetBannerRenderRows(rowView.Banner)
            outRows = outRows + 1
            outRows = outRows + rowView.SpacerRowsAfter

ContinueRowEstimate:
        Next rowItemRaw
    Else
        outRows = outRows + tableModel.RowCount
    End If

    ' Spacer after table
    outRows = outRows + 1

    mp_TryEstimateTableOutputRows = True
End Function

Private Function mp_TryWriteTableItemToBuffer( _
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
        mp_TryWriteTableItemToBuffer = True
        Exit Function
    End If

    If Not tableView.m_IsVisible() Then
        mp_TryWriteTableItemToBuffer = True
        Exit Function
    End If

    Set tableModel = tableView.Model
    If tableModel Is Nothing Then
        MsgBox "TableList: table view has no model.", vbExclamation
        Exit Function
    End If

    If tableModel.ColumnCount <= 0 Then
        MsgBox "TableList: table item has no columns.", vbExclamation
        Exit Function
    End If

    If tableModel.ColumnCount > availableCols Then
        MsgBox "TableList: control '" & m_ControlName & "' requires " & CStr(tableModel.ColumnCount) & _
               " columns, but span provides only " & CStr(availableCols) & ".", vbExclamation
        Exit Function
    End If

    If Not mp_TryAppendBannerBlock( _
        tableView.Banner, "tablebanner", tableModel.ColumnCount, valueBlock, styleSegments, plannedRows, ioCurrentOutputRow) Then Exit Function
    If ioCurrentOutputRow >= plannedRows Then
        mp_TryWriteTableItemToBuffer = True
        Exit Function
    End If

    ' Section row
    ioCurrentOutputRow = ioCurrentOutputRow + 1
    valueBlock(ioCurrentOutputRow, 1) = tableModel.SectionTitle
#If ENALBE_STYLES Then
    mp_AddStyleSegment styleSegments, "section", tableModel.ColumnCount, ioCurrentOutputRow, ioCurrentOutputRow
#End If
    If ioCurrentOutputRow >= plannedRows Then
        mp_TryWriteTableItemToBuffer = True
        Exit Function
    End If

    ' Header row
    ioCurrentOutputRow = ioCurrentOutputRow + 1
    tokens = Split(tableModel.HeaderText, "|")
    For colOffset = 1 To tableModel.ColumnCount
        If colOffset - 1 <= UBound(tokens) Then
            valueBlock(ioCurrentOutputRow, colOffset) = Trim$(CStr(tokens(colOffset - 1)))
        End If
    Next colOffset
#If ENALBE_STYLES Then
    mp_AddStyleSegment styleSegments, "header", tableModel.ColumnCount, ioCurrentOutputRow, ioCurrentOutputRow
#End If
    If ioCurrentOutputRow >= plannedRows Then
        mp_TryWriteTableItemToBuffer = True
        Exit Function
    End If

    Set rowItems = tableView.RowItems

    If Not rowItems Is Nothing And rowItems.Count > 0 Then
        For Each rowItemRaw In rowItems
            If ioCurrentOutputRow >= plannedRows Then Exit For

            Set rowView = Nothing
            If Not mp_TryResolveRowViewItem(rowItemRaw, rowView) Then Exit Function
            If rowView Is Nothing Then GoTo ContinueRowView

            If Not mp_TryAppendRowViewData( _
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
                rowModel.m_CopyToMatrixRow valueBlock, ioCurrentOutputRow, tableModel.ColumnCount
            Next rowRaw
            On Error GoTo 0
            writeEnd = ioCurrentOutputRow
#If ENALBE_STYLES Then
            If writeEnd >= writeStart Then
                mp_AddStyleSegment styleSegments, "data", tableModel.ColumnCount, writeStart, writeEnd
            End If
#End If
        End If
    End If

    If ioCurrentOutputRow >= plannedRows Then
        mp_TryWriteTableItemToBuffer = True
        Exit Function
    End If

    ' Spacer row
    ioCurrentOutputRow = ioCurrentOutputRow + 1
#If ENALBE_STYLES Then
    mp_AddStyleSegment styleSegments, "spacer", tableModel.ColumnCount, ioCurrentOutputRow, ioCurrentOutputRow
#End If

    mp_TryWriteTableItemToBuffer = True
    Exit Function

EH_INVALID_ROW:
    On Error GoTo 0
    MsgBox "TableList: unsupported row object in table rows. Expected obj_Row.", vbExclamation
End Function

Private Function mp_TryAppendRowViewData( _
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
        mp_TryAppendRowViewData = True
        Exit Function
    End If

    If Not rowView.m_IsVisible() Then
        mp_TryAppendRowViewData = True
        Exit Function
    End If

    If Not mp_TryAppendBannerBlock( _
        rowView.Banner, "rowbanner", columnCount, valueBlock, styleSegments, plannedRows, ioCurrentOutputRow) Then Exit Function

    If ioCurrentOutputRow >= plannedRows Then
        mp_TryAppendRowViewData = True
        Exit Function
    End If

    Set rowModel = rowView.Row
    If rowModel Is Nothing Then
        MsgBox "TableList: row view item has no row model.", vbExclamation
        Exit Function
    End If

    ioCurrentOutputRow = ioCurrentOutputRow + 1
    rowModel.m_CopyToMatrixRow valueBlock, ioCurrentOutputRow, columnCount
#If ENALBE_STYLES Then
    mp_AddStyleSegment styleSegments, "data", columnCount, ioCurrentOutputRow, ioCurrentOutputRow
#End If

    For spacerIndex = 1 To rowView.SpacerRowsAfter
        If ioCurrentOutputRow >= plannedRows Then Exit For

        ioCurrentOutputRow = ioCurrentOutputRow + 1
#If ENALBE_STYLES Then
        mp_AddStyleSegment styleSegments, "spacer", columnCount, ioCurrentOutputRow, ioCurrentOutputRow
#End If
    Next spacerIndex

    mp_TryAppendRowViewData = True
End Function

Private Function mp_TryAppendBannerBlock( _
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

    bannerRows = mp_GetBannerRenderRows(bannerView)
    If bannerRows <= 0 Then
        mp_TryAppendBannerBlock = True
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
        mp_AddStyleSegment styleSegments, styleKind, columnCount, writeStart, writeEnd
    End If
#End If

    mp_TryAppendBannerBlock = True
End Function

Private Function mp_GetBannerRenderRows(ByVal bannerView As obj_BannerViewItem) As Long
    Dim spanRows As Long

    If bannerView Is Nothing Then Exit Function
    If Not bannerView.m_IsVisible() Then Exit Function

    If Not bannerView.Presentation Is Nothing Then
        spanRows = bannerView.Presentation.SpanRows
    End If

    If spanRows <= 0 Then spanRows = 2
    mp_GetBannerRenderRows = spanRows
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
        Case "tablebanner"
            mp_MapStyleKindToControlPart = "itembanner"
        Case "rowbanner"
            mp_MapStyleKindToControlPart = "rowbanner"
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

        Case "tablebanner"
            backColor = RGB(45, 74, 104)
            fontColor = RGB(245, 251, 255)
            borderColor = RGB(26, 43, 61)
            fontSize = 10
            fontBold = True

        Case "rowbanner"
            backColor = RGB(52, 86, 118)
            fontColor = RGB(240, 248, 255)
            borderColor = RGB(33, 57, 82)
            fontSize = 10
            fontBold = False

        Case Else
            MsgBox "TableList: unsupported style segment kind '" & styleKind & "'.", vbExclamation
            Exit Function
    End Select

    mp_TryResolveStylePreset = True
End Function

Private Function mp_GetAvailableColumnCount() As Long
    If m_ColEnd <= 0 Or m_ColStart <= 0 Then Exit Function
    mp_GetAvailableColumnCount = m_ColEnd - m_ColStart + 1
End Function

Private Function mp_TryResolveTableViewItem(ByVal rawItem As Variant, ByRef outTableView As obj_TableViewItem) As Boolean
    Dim tableModel As obj_TableDynamic

    If Not IsObject(rawItem) Then
        MsgBox "TableList: itemsSource entry must be an object.", vbExclamation
        Exit Function
    End If

    Select Case LCase$(TypeName(rawItem))
        Case "obj_tableviewitem"
            Set outTableView = rawItem
            mp_TryResolveTableViewItem = True

        Case "obj_tabledynamic", "obj_table"
            If Not mp_TryResolveTableModelFromAny(rawItem, tableModel) Then Exit Function
            Set outTableView = mp_CreateTableViewFromModel(tableModel)
            If outTableView Is Nothing Then Exit Function
            mp_TryResolveTableViewItem = True

        Case Else
            MsgBox "TableList: unsupported item type '" & TypeName(rawItem) & _
                   "'. Expected obj_TableViewItem, obj_TableDynamic or obj_Table.", vbExclamation
    End Select
End Function

Private Function mp_TryResolveTableModelFromAny(ByVal tableItem As Variant, ByRef outTable As obj_TableDynamic) As Boolean
    Dim fixedTable As obj_Table

    If Not IsObject(tableItem) Then
        MsgBox "TableList: itemsSource entry must be an object of type obj_TableDynamic or obj_Table.", vbExclamation
        Exit Function
    End If

    Select Case LCase$(TypeName(tableItem))
        Case "obj_tabledynamic"
            Set outTable = tableItem
            mp_TryResolveTableModelFromAny = True

        Case "obj_table"
            Set fixedTable = tableItem
            Set outTable = mp_ConvertFixedTableToDynamic(fixedTable)
            If outTable Is Nothing Then Exit Function
            mp_TryResolveTableModelFromAny = True

        Case Else
            MsgBox "TableList: unsupported table model type '" & TypeName(tableItem) & _
                   "'. Expected obj_TableDynamic or obj_Table.", vbExclamation
    End Select
End Function

Private Function mp_CreateTableViewFromModel(ByVal tableModel As obj_TableDynamic) As obj_TableViewItem
    Dim tableView As obj_TableViewItem

    If tableModel Is Nothing Then
        MsgBox "TableList: table model is not specified.", vbExclamation
        Exit Function
    End If

    Set tableView = New obj_TableViewItem
    Set tableView.Model = tableModel
    tableView.ItemVisible = True

    Set mp_CreateTableViewFromModel = tableView
End Function

Private Function mp_TryResolveRowViewItem(ByVal rawItem As Variant, ByRef outRowView As obj_RowViewItem) As Boolean
    Dim rowModel As obj_Row

    If Not IsObject(rawItem) Then
        MsgBox "TableList: row item must be an object.", vbExclamation
        Exit Function
    End If

    Select Case LCase$(TypeName(rawItem))
        Case "obj_rowviewitem"
            Set outRowView = rawItem
            mp_TryResolveRowViewItem = True

        Case "obj_row"
            Set rowModel = rawItem
            Set outRowView = mp_CreateRowViewFromModel(rowModel)
            If outRowView Is Nothing Then Exit Function
            mp_TryResolveRowViewItem = True

        Case Else
            MsgBox "TableList: unsupported row item type '" & TypeName(rawItem) & _
                   "'. Expected obj_RowViewItem or obj_Row.", vbExclamation
    End Select
End Function

Private Function mp_CreateRowViewFromModel(ByVal rowModel As obj_Row) As obj_RowViewItem
    Dim rowView As obj_RowViewItem

    If rowModel Is Nothing Then
        MsgBox "TableList: row model is not specified.", vbExclamation
        Exit Function
    End If

    Set rowView = New obj_RowViewItem
    Set rowView.Row = rowModel
    rowView.RowVisible = True

    Set mp_CreateRowViewFromModel = rowView
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
        MsgBox "TableList: fixed table model is not specified.", vbExclamation
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

Private Function mp_GetWorksheetByName(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set mp_GetWorksheetByName = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function
