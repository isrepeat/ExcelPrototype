VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_TableListControlVM"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
#Const CELL_BUTTON_VIEW_ENABLED = False

Implements obj_IControl

#Const ENALBE_STYLES = True
#Const ENALBE_BORDERS = True

Private m_ControlBase As obj_ControlBase
Private m_ControlName As String
Private m_ItemsSourceRaw As String
Private m_ItemVisibilityRaw As String
#If CELL_BUTTON_VIEW_ENABLED Then
' Private m_CellButtonClickRaw As String
' Private m_CellButtonClickMacroRef As String
' Private m_CellButtonClickCallbackContext As Object
' Private m_CellButtonPayloadById As Object
' Private m_RuntimeControlKey As String
#End If
Private m_LayoutSheetName As String
Private m_RowStart As Long
Private m_ColStart As Long
Private m_RowEnd As Long
Private m_ColEnd As Long
Private m_TableItems As Collection
Private m_IsConfigured As Boolean
Private m_Page As obj_IPage
Private m_IsDisposed As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    obj_IControl_Dispose
    On Error GoTo 0
End Sub

' //
' // Interface
' //
Private Function obj_IControl_Initialize(ByVal page As obj_IPage) As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & TypeName(Me) & ".Initialize"
#End If
    m_IsDisposed = False
    m_IsConfigured = False
    Set m_Page = page
    obj_IControl_Initialize = True
End Function

Private Sub obj_IControl_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Err.Clear
    Set m_ControlBase = Nothing
    Set m_TableItems = Nothing
#If CELL_BUTTON_VIEW_ENABLED Then
'     Set m_CellButtonClickCallbackContext = Nothing
'     Set m_CellButtonPayloadById = Nothing
#End If
    Set m_Page = Nothing
    On Error GoTo 0
End Sub

Private Sub obj_IControl_Configure(ByVal controlNode As Object)
    Dim pageBase As obj_PageBase
#If CELL_BUTTON_VIEW_ENABLED Then
'     Dim dataContext As Object
#End If

    m_IsConfigured = False
    Set m_TableItems = Nothing
    Set m_ControlBase = Nothing
#If CELL_BUTTON_VIEW_ENABLED Then
'     Set m_CellButtonClickCallbackContext = Nothing
'     Set m_CellButtonPayloadById = Nothing
'     m_CellButtonClickRaw = VBA.vbNullString
'     m_CellButtonClickMacroRef = VBA.vbNullString
'     m_RuntimeControlKey = VBA.vbNullString
#End If

    If m_Page Is Nothing Then Exit Sub
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Sub
    Set m_ControlBase = New obj_ControlBase
    If Not m_ControlBase.Initialize(m_Page) Then Exit Sub
    If Not m_ControlBase.Configure(pageBase, controlNode, "TableList", "tablelist", m_ControlName) Then Exit Sub

    m_ItemsSourceRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "itemsSource")))
    If VBA.Len(m_ItemsSourceRaw) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: itemsSource is not specified for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    m_ItemVisibilityRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "itemVisibility")))
#If CELL_BUTTON_VIEW_ENABLED Then
'     m_CellButtonClickRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "cellButtonClick")))

'     Set dataContext = m_ControlBase.DataContext
'     If dataContext Is Nothing Then Set dataContext = m_Page
'     Set m_CellButtonClickCallbackContext = dataContext
'     If VBA.Len(m_CellButtonClickRaw) > 0 Then
'         If Not private_TryResolveCallbackRef(m_CellButtonClickRaw, dataContext, m_CellButtonClickMacroRef) Then Exit Sub
'     End If
#End If

    m_LayoutSheetName = VBA.Trim$(ex_XmlCore.fn_NodeAttrText(controlNode, "__layoutSheetName"))
    If VBA.Len(m_LayoutSheetName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: runtime layout sheet is missing for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    If Not private_TryReadLayoutLongAttr(controlNode, "__layoutRowStart", m_RowStart, True) Then Exit Sub
    If Not private_TryReadLayoutLongAttr(controlNode, "__layoutColStart", m_ColStart, True) Then Exit Sub
    If Not private_TryReadLayoutLongAttr(controlNode, "__layoutRowEnd", m_RowEnd, True) Then Exit Sub
    If Not private_TryReadLayoutLongAttr(controlNode, "__layoutColEnd", m_ColEnd, True) Then Exit Sub

    If m_RowStart <= 0 Or m_ColStart <= 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: invalid row/column start for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    If m_RowEnd < m_RowStart Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: control '" & m_ControlName & "' has invalid spanRows range."
#End If
        Exit Sub
    End If

    If m_ColEnd < m_ColStart Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: control '" & m_ControlName & "' has invalid spanColls range."
#End If
        Exit Sub
    End If

    Set pageBase = m_ControlBase.PageBase
    If pageBase Is Nothing Then Exit Sub
    If Not ex_RuntimeSourceResolver.fn_TryResolveItemsSource(pageBase.RuntimeSources, m_ItemsSourceRaw, m_TableItems) Then Exit Sub
    If Not private_TryApplyItemVisibilityFilter(m_TableItems) Then Exit Sub

#If CELL_BUTTON_VIEW_ENABLED Then
'     m_RuntimeControlKey = "tablelist|" & VBA.LCase$(VBA.Trim$(m_LayoutSheetName & "|" & m_ControlName))
#End If
    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render()
    Dim ws As Worksheet
    Dim valueBlock As Variant
    Dim targetRange As Range
    Dim styleSegments As Collection
#If CELL_BUTTON_VIEW_ENABLED Then
'     Dim cellButtonActions As Collection
#End If
    Dim page As obj_PageBase

    If Not m_IsConfigured Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: control '" & m_ControlName & "' is not configured."
#End If
        Exit Sub
    End If

    Set page = Nothing
    If Not m_ControlBase Is Nothing Then Set page = m_ControlBase.PageBase
    If page Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: page is not specified for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    Set ws = private_GetWorksheetByName(page, m_LayoutSheetName)
    If ws Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: sheet '" & m_LayoutSheetName & "' was not found for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

#If CELL_BUTTON_VIEW_ENABLED Then
'     private_DeleteExistingCellButtonShapes ws
'     Set m_CellButtonPayloadById = Nothing
#End If

    If m_TableItems Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: itemsSource is not resolved for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    ' Build in-memory first, then write once to minimize COM overhead.
    If Not private_TryBuildRenderBuffer(valueBlock, styleSegments) Then Exit Sub
    If IsEmpty(valueBlock) Then Exit Sub

    Set targetRange = ws.Range( _
        ws.Cells(m_RowStart, m_ColStart), _
        ws.Cells(m_RowStart + UBound(valueBlock, 1) - 1, m_ColStart + UBound(valueBlock, 2) - 1))

    targetRange.Value2 = valueBlock

    If Not private_TryRegisterControlPartSegments(ws, styleSegments) Then Exit Sub

#If ENALBE_STYLES Then
    private_ApplyStyleSegments ws, styleSegments
#End If
#If CELL_BUTTON_VIEW_ENABLED Then
'     If Not private_TryRenderCellButtonShapes(ws, cellButtonActions) Then Exit Sub
#End If
End Sub

Private Function obj_IControl_Measure( _
    ByVal controlNode As Object, _
    ByRef outSpanRows As Long, _
    ByRef outSpanColls As Long, _
    Optional ByVal dataContext As Object _
) As Boolean
    obj_IControl_Measure = private_TryMeasureNode(controlNode, outSpanRows, outSpanColls)
End Function

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(attrName))
        Case "itemssource", "itemvisibility"
            obj_IControl_SupportsAttribute = True
#If CELL_BUTTON_VIEW_ENABLED Then
'         Case "cellbuttonclick"
'             obj_IControl_SupportsAttribute = True
#End If
    End Select
End Function

Private Function obj_IControl_IsConfigured() As Boolean
    obj_IControl_IsConfigured = m_IsConfigured
End Function

' //
' // API
' //
#If CELL_BUTTON_VIEW_ENABLED Then
' Public Function RuntimeHandleCellButtonClick(Optional ByVal actionId As Variant) As Boolean
'     Dim payload As Variant
'     Dim payloadObject As Object
'     Dim actionKey As String

'     If VBA.Len(VBA.Trim$(m_CellButtonClickMacroRef)) = 0 Then
'         RuntimeHandleCellButtonClick = True
'         Exit Function
'     End If

'     actionKey = VBA.Trim$(VBA.CStr(actionId))
'     If VBA.Len(actionKey) = 0 Then Exit Function
'     If m_CellButtonPayloadById Is Nothing Then Exit Function
'     If Not m_CellButtonPayloadById.Exists(actionKey) Then Exit Function

'     If IsObject(m_CellButtonPayloadById(actionKey)) Then
'         Set payloadObject = m_CellButtonPayloadById(actionKey)
'         RuntimeHandleCellButtonClick = rt_Bridge.fn_RunCallback(m_CellButtonClickMacroRef, m_CellButtonClickCallbackContext, payloadObject)
'     Else
'         payload = m_CellButtonPayloadById(actionKey)
'         RuntimeHandleCellButtonClick = rt_Bridge.fn_RunCallback(m_CellButtonClickMacroRef, m_CellButtonClickCallbackContext, payload)
'     End If
' End Function
#End If

' //
' // Internal
' //
Private Function private_TryMeasureNode( _
    ByVal controlNode As Object, _
    ByRef outSpanRows As Long, _
    ByRef outSpanColls As Long _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim controlName As String
    Dim itemsSourceRaw As String
    Dim itemVisibilityRaw As String
    Dim tableItems As Collection
    Dim tableItem As Variant
    Dim tableViewItem As obj_TableViewItem
    Dim rowsForItem As Long

    outSpanRows = 1
    outSpanColls = 1

    If controlNode Is Nothing Then Exit Function
    If m_Page Is Nothing Then Exit Function

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If pageBase.RuntimeSources Is Nothing Then Exit Function

    controlName = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "name")))
    If VBA.Len(controlName) = 0 Then controlName = "tablelist"

    itemsSourceRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "itemsSource")))
    If VBA.Len(itemsSourceRaw) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: itemsSource is not specified for control '" & controlName & "'."
#End If
        Exit Function
    End If

    itemVisibilityRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "itemVisibility")))

    If Not ex_RuntimeSourceResolver.fn_TryResolveItemsSource(pageBase.RuntimeSources, itemsSourceRaw, tableItems) Then Exit Function
    If Not private_TryApplyItemVisibilityFilterRaw(tableItems, itemVisibilityRaw, controlName) Then Exit Function

    outSpanRows = 0
    outSpanColls = 0

    If tableItems Is Nothing Then
        private_TryMeasureNode = True
        Exit Function
    End If

    For Each tableItem In tableItems
        Set tableViewItem = Nothing
        If Not private_TryResolveTableViewItem(tableItem, tableViewItem) Then Exit Function
        If tableViewItem Is Nothing Then GoTo ContinueMeasure
        If Not private_TryEstimateTableOutputRows(tableViewItem, 1000000, rowsForItem) Then Exit Function
        outSpanRows = outSpanRows + rowsForItem
        If tableViewItem.ColumnCount > outSpanColls Then outSpanColls = tableViewItem.ColumnCount
ContinueMeasure:
    Next tableItem

    private_TryMeasureNode = True
End Function

Private Function private_TryApplyItemVisibilityFilter(ByRef tableItems As Collection) As Boolean
    private_TryApplyItemVisibilityFilter = private_TryApplyItemVisibilityFilterRaw(tableItems, m_ItemVisibilityRaw, m_ControlName)
End Function

Private Function private_TryApplyItemVisibilityFilterRaw( _
    ByRef tableItems As Collection, _
    ByVal itemVisibilityRaw As String, _
    ByVal controlName As String _
) As Boolean
    Dim filteredItems As Collection
    Dim tableItem As Variant
    Dim isVisible As Boolean

    If VBA.Len(itemVisibilityRaw) = 0 Then
        private_TryApplyItemVisibilityFilterRaw = True
        Exit Function
    End If

    If tableItems Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: itemsSource is not resolved for control '" & controlName & "'."
#End If
        Exit Function
    End If

    Set filteredItems = New Collection

    For Each tableItem In tableItems
        If Not IsObject(tableItem) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "TableList: itemsSource entry must be an object for itemVisibility evaluation in control '" & controlName & "'."
#End If
            Exit Function
        End If

        If Not ex_BindingRuntime.fn_TryResolveVisibilityBinding(itemVisibilityRaw, tableItem, isVisible) Then Exit Function
        If isVisible Then filteredItems.Add tableItem
    Next tableItem

    Set tableItems = filteredItems
    private_TryApplyItemVisibilityFilterRaw = True
End Function

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
            ex_Core.fn_Diagnostic_LogError "TableList: runtime layout attribute '" & attrName & "' is missing for control '" & m_ControlName & "'."
#End If
            Exit Function
        End If

        outValue = 0
        private_TryReadLayoutLongAttr = True
        Exit Function
    End If

    If Not VBA.IsNumeric(rawText) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: runtime layout attribute '" & attrName & "' must be numeric for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    outValue = VBA.CLng(rawText)
    private_TryReadLayoutLongAttr = True
End Function

Private Function private_TryBuildRenderBuffer( _
    ByRef outValueBlock As Variant, _
    ByRef outStyleSegments As Collection _
) As Boolean
    Dim tableItem As Variant
    Dim tableViewItem As obj_TableViewItem
    Dim availableCols As Long
    Dim maxRows As Long
    Dim plannedRows As Long
    Dim rowsForItem As Long
    Dim currentOutputRow As Long

    availableCols = private_GetAvailableColumnCount()
    maxRows = m_RowEnd - m_RowStart + 1

    If availableCols <= 0 Or maxRows <= 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: invalid render bounds for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    plannedRows = 0

    ' Pass 1: estimate output size up-front to allocate matrix once.
    For Each tableItem In m_TableItems
        Set tableViewItem = Nothing
        If Not private_TryResolveTableViewItem(tableItem, tableViewItem) Then Exit Function
        If tableViewItem Is Nothing Then GoTo ContinueEstimate

        If Not private_TryEstimateTableOutputRows(tableViewItem, availableCols, rowsForItem) Then Exit Function
        plannedRows = plannedRows + rowsForItem
        If plannedRows > maxRows Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "TableList: insufficient layout bounds for control '" & m_ControlName & "'. RequiredRows=" & VBA.CStr(plannedRows) & ", AvailableRows=" & VBA.CStr(maxRows) & "."
#End If
            VBA.MsgBox "PrototypeNew: table list control '" & m_ControlName & "' does not fit into allocated bounds. Required rows: " & VBA.CStr(plannedRows) & ", available rows: " & VBA.CStr(maxRows) & ". Increase spanRows or container size.", VBA.vbExclamation, "PrototypeNew / Table layout"
            Exit Function
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
#If CELL_BUTTON_VIEW_ENABLED Then
'     Set outCellButtonActions = New Collection
#End If

    currentOutputRow = 0

    ' Pass 2: fill matrix sequentially.
    For Each tableItem In m_TableItems
        If currentOutputRow >= plannedRows Then Exit For

        Set tableViewItem = Nothing
        If Not private_TryResolveTableViewItem(tableItem, tableViewItem) Then Exit Function
        If tableViewItem Is Nothing Then GoTo ContinueWrite

        If Not private_TryWriteTableItemToBuffer( _
            tableViewItem, outValueBlock, outStyleSegments, availableCols, plannedRows, currentOutputRow) Then Exit Function

ContinueWrite:
    Next tableItem

    private_TryBuildRenderBuffer = True
End Function

Private Function private_TryEstimateTableOutputRows( _
    ByVal tableViewItem As obj_TableViewItem, _
    ByVal availableCols As Long, _
    ByRef outRows As Long _
) As Boolean
    Dim tableDynamic As obj_TableDynamic
    Dim rowViewItems As list__obj_RowViewItem
    Dim rowViewItemRaw As Variant
    Dim rowViewItem As obj_RowViewItem
    Dim rowViewItemIndex As Long

    outRows = 0

    If tableViewItem Is Nothing Then
        private_TryEstimateTableOutputRows = True
        Exit Function
    End If

    If Not tableViewItem.IsVisible() Then
        private_TryEstimateTableOutputRows = True
        Exit Function
    End If

    Set tableDynamic = tableViewItem.Model
    If tableDynamic Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: table view has no model."
#End If
        Exit Function
    End If

    If tableDynamic.ColumnCount <= 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: table item has no columns."
#End If
        Exit Function
    End If

    If tableDynamic.ColumnCount > availableCols Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: control '" & m_ControlName & "' requires " & VBA.CStr(tableDynamic.ColumnCount) & _
               " columns, but span provides only " & VBA.CStr(availableCols) & "."
#End If
        VBA.MsgBox "PrototypeNew: table list control '" & m_ControlName & "' does not fit into allocated bounds. Required columns: " & VBA.CStr(tableDynamic.ColumnCount) & ", available columns: " & VBA.CStr(availableCols) & ". Increase spanColls or container size.", VBA.vbExclamation, "PrototypeNew / Table layout"
        Exit Function
    End If

    outRows = outRows + private_GetBannerRenderRows(tableViewItem.Banner)

    ' section + header
    outRows = outRows + 2

    Set rowViewItems = tableViewItem.RowViewItems
    If Not rowViewItems Is Nothing And rowViewItems.Count > 0 Then
        For rowViewItemIndex = 1 To rowViewItems.Count
            Set rowViewItemRaw = rowViewItems.Item(rowViewItemIndex)
            Set rowViewItem = Nothing
            If Not private_TryResolveRowViewItem(rowViewItemRaw, rowViewItem) Then Exit Function
            If rowViewItem Is Nothing Then GoTo ContinueRowEstimate
            If Not rowViewItem.IsVisible() Then GoTo ContinueRowEstimate

            outRows = outRows + private_GetBannerRenderRows(rowViewItem.Banner)
            outRows = outRows + 1
            outRows = outRows + rowViewItem.SpacerRowsAfter

ContinueRowEstimate:
        Next rowViewItemIndex
    Else
        outRows = outRows + tableDynamic.RowCount
    End If

    ' Spacer after table
    outRows = outRows + 1

    private_TryEstimateTableOutputRows = True
End Function

Private Function private_TryWriteTableItemToBuffer( _
    ByVal tableViewItem As obj_TableViewItem, _
    ByRef valueBlock As Variant, _
    ByVal styleSegments As Collection, _
    ByVal availableCols As Long, _
    ByVal plannedRows As Long, _
    ByRef ioCurrentOutputRow As Long _
) As Boolean
    Dim tableDynamic As obj_TableDynamic
    Dim rowViewItems As list__obj_RowViewItem
    Dim rowViewItemRaw As Variant
    Dim rowViewItem As obj_RowViewItem
    Dim tableRows As list__obj_Row
    Dim sourceRow As obj_Row
    Dim row As obj_Row
    Dim colOffset As Long
    Dim tokens As Variant
    Dim writeStart As Long
    Dim writeEnd As Long
    Dim rowViewItemIndex As Long
    Dim tableRowIndex As Long

    If tableViewItem Is Nothing Then
        private_TryWriteTableItemToBuffer = True
        Exit Function
    End If

    If Not tableViewItem.IsVisible() Then
        private_TryWriteTableItemToBuffer = True
        Exit Function
    End If

    Set tableDynamic = tableViewItem.Model
    If tableDynamic Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: table view has no model."
#End If
        Exit Function
    End If

    If tableDynamic.ColumnCount <= 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: table item has no columns."
#End If
        Exit Function
    End If

    If tableDynamic.ColumnCount > availableCols Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: control '" & m_ControlName & "' requires " & VBA.CStr(tableDynamic.ColumnCount) & _
               " columns, but span provides only " & VBA.CStr(availableCols) & "."
#End If
        Exit Function
    End If

    If Not private_TryAppendBannerBlock( _
        tableViewItem.Banner, "tablebanner", tableDynamic.ColumnCount, valueBlock, styleSegments, plannedRows, ioCurrentOutputRow) Then Exit Function
    If ioCurrentOutputRow >= plannedRows Then
        private_TryWriteTableItemToBuffer = True
        Exit Function
    End If

    ' Section row
    ioCurrentOutputRow = ioCurrentOutputRow + 1
    valueBlock(ioCurrentOutputRow, 1) = tableDynamic.SectionTitle
#If ENALBE_STYLES Then
    private_AddStyleSegment styleSegments, "section", tableDynamic.ColumnCount, ioCurrentOutputRow, ioCurrentOutputRow
#End If
    If ioCurrentOutputRow >= plannedRows Then
        private_TryWriteTableItemToBuffer = True
        Exit Function
    End If

    ' Header row
    ioCurrentOutputRow = ioCurrentOutputRow + 1
    tokens = VBA.Split(tableDynamic.HeaderText, "|")
    For colOffset = 1 To tableDynamic.ColumnCount
        If colOffset - 1 <= UBound(tokens) Then
            valueBlock(ioCurrentOutputRow, colOffset) = VBA.Trim$(VBA.CStr(tokens(colOffset - 1)))
        End If
    Next colOffset
#If ENALBE_STYLES Then
    private_AddStyleSegment styleSegments, "header", tableDynamic.ColumnCount, ioCurrentOutputRow, ioCurrentOutputRow
#End If
    If ioCurrentOutputRow >= plannedRows Then
        private_TryWriteTableItemToBuffer = True
        Exit Function
    End If

    Set rowViewItems = tableViewItem.RowViewItems

    If Not rowViewItems Is Nothing And rowViewItems.Count > 0 Then
        For rowViewItemIndex = 1 To rowViewItems.Count
            If ioCurrentOutputRow >= plannedRows Then Exit For
            Set rowViewItemRaw = rowViewItems.Item(rowViewItemIndex)

            Set rowViewItem = Nothing
            If Not private_TryResolveRowViewItem(rowViewItemRaw, rowViewItem) Then Exit Function
            If rowViewItem Is Nothing Then GoTo ContinueRowView

            If Not private_TryAppendRowViewData( _
                rowViewItem, tableDynamic.ColumnCount, valueBlock, styleSegments, plannedRows, ioCurrentOutputRow) Then Exit Function

ContinueRowView:
        Next rowViewItemIndex
    Else
        Set tableRows = tableDynamic.Rows
        If Not tableRows Is Nothing Then
            writeStart = ioCurrentOutputRow + 1
            For tableRowIndex = 1 To tableRows.Count
                If ioCurrentOutputRow >= plannedRows Then Exit For
                Set sourceRow = tableRows.Item(tableRowIndex)
                If sourceRow Is Nothing Then GoTo ContinueTableRow

                ioCurrentOutputRow = ioCurrentOutputRow + 1
                Set row = sourceRow
                row.CopyToMatrixRow valueBlock, ioCurrentOutputRow, tableDynamic.ColumnCount
#If CELL_BUTTON_VIEW_ENABLED Then
'                 private_CollectCellButtonActions row, ioCurrentOutputRow, tableDynamic.ColumnCount, cellButtonActions
#End If
ContinueTableRow:
            Next tableRowIndex
            writeEnd = ioCurrentOutputRow
#If ENALBE_STYLES Then
            If writeEnd >= writeStart Then
                private_AddStyleSegment styleSegments, "data", tableDynamic.ColumnCount, writeStart, writeEnd
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
    private_AddStyleSegment styleSegments, "spacer", tableDynamic.ColumnCount, ioCurrentOutputRow, ioCurrentOutputRow
#End If

    private_TryWriteTableItemToBuffer = True
End Function

Private Function private_TryAppendRowViewData( _
    ByVal rowViewItem As obj_RowViewItem, _
    ByVal columnCount As Long, _
    ByRef valueBlock As Variant, _
    ByVal styleSegments As Collection, _
    ByVal plannedRows As Long, _
    ByRef ioCurrentOutputRow As Long _
) As Boolean
    Dim row As obj_Row
    Dim spacerIndex As Long

    If rowViewItem Is Nothing Then
        private_TryAppendRowViewData = True
        Exit Function
    End If

    If Not rowViewItem.IsVisible() Then
        private_TryAppendRowViewData = True
        Exit Function
    End If

    If Not private_TryAppendBannerBlock( _
        rowViewItem.Banner, "rowbanner", columnCount, valueBlock, styleSegments, plannedRows, ioCurrentOutputRow) Then Exit Function

    If ioCurrentOutputRow >= plannedRows Then
        private_TryAppendRowViewData = True
        Exit Function
    End If

    Set row = rowViewItem.Model
    If row Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: row view item has no row model."
#End If
        Exit Function
    End If

    ioCurrentOutputRow = ioCurrentOutputRow + 1
    row.CopyToMatrixRow valueBlock, ioCurrentOutputRow, columnCount
#If CELL_BUTTON_VIEW_ENABLED Then
'     private_CollectCellButtonActions row, ioCurrentOutputRow, columnCount, cellButtonActions
#End If
#If ENALBE_STYLES Then
    private_AddStyleSegment styleSegments, "data", columnCount, ioCurrentOutputRow, ioCurrentOutputRow
#End If

    For spacerIndex = 1 To rowViewItem.SpacerRowsAfter
        If ioCurrentOutputRow >= plannedRows Then Exit For

        ioCurrentOutputRow = ioCurrentOutputRow + 1
#If ENALBE_STYLES Then
        private_AddStyleSegment styleSegments, "spacer", columnCount, ioCurrentOutputRow, ioCurrentOutputRow
#End If
    Next spacerIndex

    private_TryAppendRowViewData = True
End Function

#If CELL_BUTTON_VIEW_ENABLED Then
' Private Sub private_CollectCellButtonActions( _
'     ByVal rowObj As obj_Row, _
'     ByVal relativeRow As Long, _
'     ByVal columnCount As Long, _
'     ByVal cellButtonActions As Collection _
' )
'     Dim colIndex As Long
'     Dim cellObj As obj_Cell
'     Dim actionInfo As Object

'     If rowObj Is Nothing Then Exit Sub
'     If cellButtonActions Is Nothing Then Exit Sub
'     If relativeRow <= 0 Or columnCount <= 0 Then Exit Sub

'     For colIndex = 1 To columnCount
'         Set cellObj = Nothing
'         If Not rowObj.TryGetCellAt(colIndex, cellObj) Then GoTo ContinueCell
'         If cellObj Is Nothing Then GoTo ContinueCell
'         If Not cellObj.IsButtonView Then GoTo ContinueCell

'         Set actionInfo = VBA.CreateObject("Scripting.Dictionary")
'         actionInfo.CompareMode = 1
'         actionInfo("RelativeRow") = relativeRow
'         actionInfo("RelativeCol") = colIndex
'         Set actionInfo("Cell") = cellObj
'         cellButtonActions.Add actionInfo

' ContinueCell:
'     Next colIndex
' End Sub
#End If

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
    Dim banner As obj_Banner
    Dim rowOffset As Long

    bannerRows = private_GetBannerRenderRows(bannerView)
    If bannerRows <= 0 Then
        private_TryAppendBannerBlock = True
        Exit Function
    End If

    Set banner = bannerView.Model
    writeStart = ioCurrentOutputRow + 1

    For rowOffset = 1 To bannerRows
        If ioCurrentOutputRow >= plannedRows Then Exit For

        ioCurrentOutputRow = ioCurrentOutputRow + 1

        If Not banner Is Nothing Then
            If rowOffset = 1 Then
                valueBlock(ioCurrentOutputRow, 1) = banner.Header
            ElseIf rowOffset = 2 Then
                valueBlock(ioCurrentOutputRow, 1) = banner.Message
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

#If CELL_BUTTON_VIEW_ENABLED Then
' Private Function private_TryRenderCellButtonShapes(ByVal ws As Worksheet, ByVal cellButtonActions As Collection) As Boolean
'     Dim pageBase As obj_PageBase
'     Dim actionInfo As Object
'     Dim actionId As Long

'     If ws Is Nothing Then Exit Function
'     If cellButtonActions Is Nothing Then
'         private_TryRenderCellButtonShapes = True
'         Exit Function
'     End If
'     If cellButtonActions.Count <= 0 Then
'         private_TryRenderCellButtonShapes = True
'         Exit Function
'     End If

'     If VBA.Len(VBA.Trim$(m_CellButtonClickMacroRef)) = 0 Then
' #If LOGGING_DEBUG_ENABLED Then
'         ex_Core.fn_Diagnostic_LogError "TableList: ButtonView cells found, but cellButtonClick is not configured for control '" & m_ControlName & "'."
' #End If
'         private_TryRenderCellButtonShapes = True
'         Exit Function
'     End If

'     If m_Page Is Nothing Then Exit Function
'     Set pageBase = m_Page.GetPageBase()
'     If pageBase Is Nothing Then Exit Function
'     If VBA.Len(VBA.Trim$(m_RuntimeControlKey)) = 0 Then Exit Function
'     If Not pageBase.RegisterControl(m_RuntimeControlKey, Me) Then Exit Function

'     Set m_CellButtonPayloadById = VBA.CreateObject("Scripting.Dictionary")
'     m_CellButtonPayloadById.CompareMode = 1

'     actionId = 0
'     For Each actionInfo In cellButtonActions
'         If actionInfo Is Nothing Then GoTo ContinueAction
'         actionId = actionId + 1
'         If Not private_TryRenderOneCellButtonShape(ws, pageBase, actionInfo, actionId) Then Exit Function
' ContinueAction:
'     Next actionInfo

'     private_TryRenderCellButtonShapes = True
' End Function

' Private Function private_TryRenderOneCellButtonShape( _
'     ByVal ws As Worksheet, _
'     ByVal pageBase As obj_PageBase, _
'     ByVal actionInfo As Object, _
'     ByVal actionId As Long _
' ) As Boolean
'     Dim cellObj As obj_Cell
'     Dim targetCell As Range
'     Dim shp As Shape
'     Dim shapeName As String
'     Dim actionKey As String
'     Dim captionText As String
'     Dim callbackMacroRef As String
'     Dim metaMap As Object
'     Dim relativeRow As Long
'     Dim relativeCol As Long

'     If ws Is Nothing Then Exit Function
'     If pageBase Is Nothing Then Exit Function
'     If actionInfo Is Nothing Then Exit Function
'     If actionId <= 0 Then Exit Function

'     relativeRow = VBA.CLng(actionInfo("RelativeRow"))
'     relativeCol = VBA.CLng(actionInfo("RelativeCol"))
'     If relativeRow <= 0 Or relativeCol <= 0 Then Exit Function

'     Set cellObj = Nothing
'     On Error Resume Next
'     Set cellObj = actionInfo("Cell")
'     On Error GoTo 0
'     If cellObj Is Nothing Then Exit Function

'     Set targetCell = ws.Cells(m_RowStart + relativeRow - 1, m_ColStart + relativeCol - 1)
'     If targetCell Is Nothing Then Exit Function
'     If targetCell.Width <= 0# Or targetCell.Height <= 0# Then Exit Function

'     actionKey = VBA.CStr(actionId)
'     If Not private_TryStoreCellButtonPayload(actionKey, cellObj) Then Exit Function

'     shapeName = private_BuildCellButtonShapeName(actionId)
'     If VBA.Len(shapeName) = 0 Then Exit Function

'     Set shp = private_GetUiShapeByName(ws, shapeName)
'     If shp Is Nothing Then
'         Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, targetCell.Left, targetCell.Top, targetCell.Width, targetCell.Height)
'         shp.Name = shapeName
'     Else
'         shp.Left = targetCell.Left
'         shp.Top = targetCell.Top
'         shp.Width = targetCell.Width
'         shp.Height = targetCell.Height
'     End If
'     shp.Placement = xlMoveAndSize

'     callbackMacroRef = private_GetRuntimeCallbackMacroRef()
'     If VBA.Len(callbackMacroRef) = 0 Then Exit Function
'     If Not private_TryAssignShapeOnActionIfChanged(shp, callbackMacroRef) Then Exit Function
'     If Not pageBase.RegisterShapeRoute(shp.Name, m_RuntimeControlKey, "RuntimeHandleCellButtonClick", True, actionId) Then Exit Function

'     captionText = cellObj.Value
'     On Error Resume Next
'     shp.TextFrame2.TextRange.Text = captionText
'     shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
'     shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
'     shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = VBA.RGB(248, 250, 252)
'     shp.TextFrame2.TextRange.Font.Size = 10
'     shp.TextFrame.Characters.Text = captionText
'     shp.TextFrame.HorizontalAlignment = xlHAlignLeft
'     shp.TextFrame.VerticalAlignment = xlVAlignCenter
'     shp.Fill.ForeColor.RGB = VBA.RGB(5, 79, 35)
'     shp.Fill.Transparency = 0.05
'     shp.Line.ForeColor.RGB = VBA.RGB(16, 185, 129)
'     shp.Line.Weight = 1
'     On Error GoTo 0

'     Set metaMap = VBA.CreateObject("Scripting.Dictionary")
'     metaMap.CompareMode = 1
'     metaMap("pn.control") = m_ControlName
'     metaMap("pn.part") = "cellButton"
'     metaMap("pn.actionId") = actionKey
'     If Not ex_ShapeMetaRuntime.fn_TrySetShapeMetaValues(shp, metaMap) Then Exit Function

'     private_TryRenderOneCellButtonShape = True
' End Function

' Private Function private_TryStoreCellButtonPayload(ByVal actionKey As String, ByVal cellObj As obj_Cell) As Boolean
'     Dim payload As Variant
'     Dim payloadObject As Object

'     If m_CellButtonPayloadById Is Nothing Then Exit Function
'     If cellObj Is Nothing Then Exit Function
'     actionKey = VBA.Trim$(actionKey)
'     If VBA.Len(actionKey) = 0 Then Exit Function

'     If cellObj.ButtonActionArgIsObject Then
'         Set payloadObject = cellObj.ButtonActionArg
'         Set m_CellButtonPayloadById(actionKey) = payloadObject
'     Else
'         payload = cellObj.ButtonActionArg
'         m_CellButtonPayloadById(actionKey) = payload
'     End If

'     private_TryStoreCellButtonPayload = True
' End Function
#End If

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
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "TableList: unsupported style segment kind '" & styleKind & "'."
#End If
            Exit Function
    End Select

    private_TryResolveStylePreset = True
End Function

Private Function private_GetAvailableColumnCount() As Long
    If m_ColEnd <= 0 Or m_ColStart <= 0 Then Exit Function
    private_GetAvailableColumnCount = m_ColEnd - m_ColStart + 1
End Function

Private Function private_TryResolveTableViewItem(ByVal rawItem As Variant, ByRef outTableView As obj_TableViewItem) As Boolean
    Dim tableDynamic As obj_TableDynamic

    If Not IsObject(rawItem) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: itemsSource entry must be an object."
#End If
        Exit Function
    End If

    Select Case VBA.LCase$(TypeName(rawItem))
        Case "obj_tableviewitem"
            Set outTableView = rawItem
            private_TryResolveTableViewItem = True

        Case "obj_tabledynamic", "obj_table"
            If Not private_TryResolveTableModelFromAny(rawItem, tableDynamic) Then Exit Function
            Set outTableView = private_CreateTableViewFromModel(tableDynamic)
            If outTableView Is Nothing Then Exit Function
            private_TryResolveTableViewItem = True

        Case Else
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "TableList: unsupported item type '" & TypeName(rawItem) & _
                   "'. Expected obj_TableViewItem, obj_TableDynamic or obj_Table."
#End If
    End Select
End Function

Private Function private_TryResolveTableModelFromAny(ByVal tableItem As Variant, ByRef outTable As obj_TableDynamic) As Boolean
    Dim fixedTable As obj_Table

    If Not IsObject(tableItem) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: itemsSource entry must be an object of type obj_TableDynamic or obj_Table."
#End If
        Exit Function
    End If

    Select Case VBA.LCase$(TypeName(tableItem))
        Case "obj_tabledynamic"
            Set outTable = tableItem
            private_TryResolveTableModelFromAny = True

        Case "obj_table"
            Set fixedTable = tableItem
            Set outTable = private_ConvertFixedTableToDynamic(fixedTable)
            If outTable Is Nothing Then Exit Function
            private_TryResolveTableModelFromAny = True

        Case Else
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "TableList: unsupported table model type '" & TypeName(tableItem) & _
                   "'. Expected obj_TableDynamic or obj_Table."
#End If
    End Select
End Function

Private Function private_CreateTableViewFromModel(ByVal tableDynamic As obj_TableDynamic) As obj_TableViewItem
    Dim tableViewItem As obj_TableViewItem

    If tableDynamic Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: table model is not specified."
#End If
        Exit Function
    End If

    Set tableViewItem = New obj_TableViewItem
    If Not tableViewItem.Initialize(tableDynamic) Then Exit Function
    tableViewItem.ItemVisible = True

    Set private_CreateTableViewFromModel = tableViewItem
End Function

Private Function private_TryResolveRowViewItem(ByVal rawItem As Variant, ByRef outRowView As obj_RowViewItem) As Boolean
    Dim row As obj_Row

    If Not IsObject(rawItem) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: row item must be an object."
#End If
        Exit Function
    End If

    Select Case VBA.LCase$(TypeName(rawItem))
        Case "obj_rowviewitem"
            Set outRowView = rawItem
            private_TryResolveRowViewItem = True

        Case "obj_row"
            Set row = rawItem
            Set outRowView = private_CreateRowViewFromModel(row)
            If outRowView Is Nothing Then Exit Function
            private_TryResolveRowViewItem = True

        Case Else
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "TableList: unsupported row item type '" & TypeName(rawItem) & _
                   "'. Expected obj_RowViewItem or obj_Row."
#End If
    End Select
End Function

Private Function private_CreateRowViewFromModel(ByVal row As obj_Row) As obj_RowViewItem
    Dim rowViewItem As obj_RowViewItem

    If row Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: row model is not specified."
#End If
        Exit Function
    End If

    Set rowViewItem = New obj_RowViewItem
    If Not rowViewItem.Initialize(row) Then Exit Function
    rowViewItem.RowVisible = True

    Set private_CreateRowViewFromModel = rowViewItem
End Function

Private Function private_ConvertFixedTableToDynamic(ByVal fixedTable As obj_Table) As obj_TableDynamic
    Dim tableDynamic As obj_TableDynamic
    Dim sourceColumns As list__obj_Column
    Dim sourceRows As list__obj_Row
    Dim sourceColumn As obj_Column
    Dim sourceRow As obj_Row
    Dim targetColumn As obj_Column
    Dim targetRow As obj_Row
    Dim sourceAliases As Collection
    Dim aliasItem As Variant
    Dim colIndex As Long
    Dim sourceColumnIndex As Long
    Dim sourceRowIndex As Long

    If fixedTable Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: fixed table model is not specified."
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
        Set sourceAliases = sourceColumn.Aliases
        If Not sourceAliases Is Nothing Then
            For Each aliasItem In sourceAliases
                If Not targetColumn.AddAlias(VBA.CStr(aliasItem)) Then Exit Function
            Next aliasItem
        End If
        If Not tableDynamic.PushColumn(targetColumn) Then Exit Function
ContinueSourceColumn:
    Next sourceColumnIndex

    Set sourceRows = fixedTable.Rows
    For sourceRowIndex = 1 To sourceRows.Count
        Set sourceRow = sourceRows.Item(sourceRowIndex)
        If sourceRow Is Nothing Then GoTo ContinueSourceRow
        Set targetRow = New obj_Row
        For colIndex = 1 To tableDynamic.ColumnCount
            targetRow.PushCellRaw sourceRow.GetCellValue(colIndex)
        Next colIndex
        If Not tableDynamic.PushRow(targetRow) Then Exit Function
ContinueSourceRow:
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

#If CELL_BUTTON_VIEW_ENABLED Then
' Private Sub private_DeleteExistingCellButtonShapes(ByVal ws As Worksheet)
'     Dim shapePrefix As String
'     Dim i As Long
'     Dim shp As Shape

'     If ws Is Nothing Then Exit Sub
'     shapePrefix = VBA.LCase$(private_GetCellButtonShapePrefix())
'     If VBA.Len(shapePrefix) = 0 Then Exit Sub

'     On Error Resume Next
'     For i = ws.Shapes.Count To 1 Step -1
'         Set shp = ws.Shapes.Item(i)
'         If Not shp Is Nothing Then
'             If VBA.Left$(VBA.LCase$(VBA.Trim$(shp.Name)), VBA.Len(shapePrefix)) = shapePrefix Then
'                 shp.Delete
'             End If
'         End If
'     Next i
'     On Error GoTo 0
' End Sub

' Private Function private_BuildCellButtonShapeName(ByVal actionId As Long) As String
'     If actionId <= 0 Then Exit Function
'     private_BuildCellButtonShapeName = private_GetCellButtonShapePrefix() & VBA.CStr(actionId)
' End Function

' Private Function private_GetCellButtonShapePrefix() As String
'     Dim normalizedName As String

'     normalizedName = private_NormalizeNamePart(m_ControlName)
'     If VBA.Len(normalizedName) = 0 Then normalizedName = "tablelist"
'     private_GetCellButtonShapePrefix = "tblbtn_" & normalizedName & "_"
' End Function

' Private Function private_GetUiShapeByName(ByVal ws As Worksheet, ByVal shapeName As String) As Shape
'     If ws Is Nothing Then Exit Function
'     shapeName = VBA.Trim$(shapeName)
'     If VBA.Len(shapeName) = 0 Then Exit Function

'     On Error Resume Next
'     Set private_GetUiShapeByName = ws.Shapes(shapeName)
'     On Error GoTo 0
' End Function

' Private Function private_TryAssignShapeOnActionIfChanged(ByVal shp As Shape, ByVal macroRef As String) As Boolean
'     Dim currentMacroRef As String

'     If shp Is Nothing Then Exit Function
'     macroRef = VBA.Trim$(macroRef)
'     If VBA.Len(macroRef) = 0 Then
'         private_TryAssignShapeOnActionIfChanged = True
'         Exit Function
'     End If

'     On Error Resume Next
'     currentMacroRef = VBA.Trim$(VBA.CStr(shp.OnAction))
'     If Err.Number <> 0 Then
'         Err.Clear
'         currentMacroRef = VBA.vbNullString
'     End If
'     On Error GoTo 0

'     If VBA.StrComp(currentMacroRef, macroRef, VBA.vbBinaryCompare) <> 0 Then
'         On Error GoTo EH_SET
'         shp.OnAction = macroRef
'         On Error GoTo 0
'     End If

'     private_TryAssignShapeOnActionIfChanged = True
'     Exit Function

' EH_SET:
'     On Error GoTo 0
' End Function

' Private Function private_TryResolveCallbackRef( _
'     ByVal rawText As String, _
'     ByVal dataContext As Object, _
'     ByRef outCallbackRef As String _
' ) As Boolean
'     Dim resolvedValue As Variant

'     outCallbackRef = VBA.vbNullString
'     rawText = VBA.Trim$(rawText)
'     If VBA.Len(rawText) = 0 Then
'         private_TryResolveCallbackRef = True
'         Exit Function
'     End If

'     If Not ex_BindingRuntime.fn_TryResolveValueBinding(rawText, dataContext, resolvedValue) Then Exit Function
'     If IsObject(resolvedValue) Then
' #If LOGGING_DEBUG_ENABLED Then
'         ex_Core.fn_Diagnostic_LogError "TableList: callback binding must resolve to scalar value for control '" & m_ControlName & "'."
' #End If
'         Exit Function
'     End If

'     outCallbackRef = VBA.Trim$(VBA.CStr(resolvedValue))
'     If VBA.Len(outCallbackRef) = 0 Then
' #If LOGGING_DEBUG_ENABLED Then
'         ex_Core.fn_Diagnostic_LogError "TableList: callback binding resolved to empty value for control '" & m_ControlName & "'."
' #End If
'         Exit Function
'     End If

'     private_TryResolveCallbackRef = True
' End Function

' Private Function private_GetRuntimeCallbackMacroRef() As String
'     private_GetRuntimeCallbackMacroRef = private_QualifyMacroName("rt_Bridge.fn_OnShapeClick")
' End Function

' Private Function private_QualifyMacroName(ByVal macroName As String) As String
'     Dim wbName As String

'     macroName = VBA.Trim$(macroName)
'     If VBA.Len(macroName) = 0 Then Exit Function
'     If VBA.InStr(1, macroName, "!", VBA.vbBinaryCompare) > 0 Then
'         private_QualifyMacroName = macroName
'         Exit Function
'     End If

'     wbName = ThisWorkbook.Name
'     wbName = VBA.Replace$(wbName, "'", "''")
'     private_QualifyMacroName = "'" & wbName & "'!" & macroName
' End Function

' Private Function private_NormalizeNamePart(ByVal rawText As String) As String
'     Dim i As Long
'     Dim ch As String
'     Dim outText As String

'     rawText = VBA.Trim$(rawText)
'     For i = 1 To VBA.Len(rawText)
'         ch = VBA.Mid$(rawText, i, 1)
'         If (ch >= "A" And ch <= "Z") Or _
'            (ch >= "a" And ch <= "z") Or _
'            (ch >= "0" And ch <= "9") Or _
'            ch = "_" Then
'             outText = outText & ch
'         Else
'             outText = outText & "_"
'         End If
'     Next i

'     If VBA.Len(outText) = 0 Then outText = "x"
'     private_NormalizeNamePart = VBA.Left$(outText, 80)
' End Function
#End If

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
