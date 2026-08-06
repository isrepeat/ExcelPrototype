VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_TableListControlVM"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Implements obj_IControl

#Const ENALBE_STYLES = True
#Const ENALBE_BORDERS = True

Private m_ControlBase As obj_ControlBase
Private m_ControlName As String
Private m_ItemsSourceRaw As String
Private m_ItemVisibilityRaw As String
Private m_RenderAsListObject As Boolean
Private m_MergeSectionCells As Boolean
Private m_TableNameRaw As String
Private m_RuntimeTableName As String
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
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & TypeName(Me) & ".Class_Terminate"
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
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & TypeName(Me) & ".Initialize"
#End If
    m_IsDisposed = False
    m_IsConfigured = False
    Set m_Page = page
    obj_IControl_Initialize = True
End Function

Private Sub obj_IControl_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Err.Clear
    Set m_ControlBase = Nothing
    Set m_TableItems = Nothing
    m_RuntimeTableName = VBA.vbNullString
    Set m_Page = Nothing
    On Error GoTo 0
End Sub

Private Sub obj_IControl_Configure(ByVal controlNode As Object)
    Dim pageBase As obj_PageBase

    m_IsConfigured = False
    Set m_TableItems = Nothing
    Set m_ControlBase = Nothing
    m_RenderAsListObject = False
    m_MergeSectionCells = False
    m_TableNameRaw = VBA.vbNullString
    m_RuntimeTableName = VBA.vbNullString

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
    If Not private_TryReadOptionalBooleanAttr(controlNode, "renderAsListObject", False, m_RenderAsListObject) Then Exit Sub
    If Not private_TryReadOptionalBooleanAttr(controlNode, "mergeSectionCells", False, m_MergeSectionCells) Then Exit Sub
    m_TableNameRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "tableName")))

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

    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render()
    Dim ws As Worksheet
    Dim valueBlock As Variant
    Dim targetRange As Range
    Dim styleSegments As Collection
    Dim rowCount As Long
    Dim columnCount As Long
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

    If m_TableItems Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: itemsSource is not resolved for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    ' Сначала собираем таблицы в памяти: valueBlock содержит значения,
    ' а styleSegments размечает строки/ячейки по смыслу внутри этой матрицы.
    If Not private_TryBuildRenderBuffer(valueBlock, styleSegments) Then Exit Sub
    If IsEmpty(valueBlock) Then Exit Sub
    rowCount = UBound(valueBlock, 1)
    columnCount = UBound(valueBlock, 2)

    Set targetRange = ws.Range( _
        ws.Cells(m_RowStart, m_ColStart), _
        ws.Cells(m_RowStart + rowCount - 1, m_ColStart + columnCount - 1))

    If m_RenderAsListObject Then
        If Not private_TryDeleteIntersectingTables(ws, ws.Range(ws.Cells(m_RowStart, m_ColStart), ws.Cells(m_RowEnd, m_ColEnd))) Then Exit Sub
    End If
    ' При повторном partial render section предыдущего результата уже может
    ' быть объединён. Сначала возвращаем прямоугольную сетку, иначе Excel не
    ' позволит записать новый двумерный valueBlock в targetRange.
    If m_MergeSectionCells Then targetRange.UnMerge
    targetRange.Value2 = valueBlock

    If m_RenderAsListObject Then
        If Not private_TryCreateRenderedListObject(ws, styleSegments) Then Exit Sub
    End If
    If Not private_TryRegisterControlColumnAliasSegments(ws, valueBlock, rowCount, columnCount, styleSegments) Then Exit Sub
    If Not private_TryRegisterControlSourceAliasSegments(ws, rowCount, styleSegments) Then Exit Sub

    ' Размеченным ранее смысловым частям назначаются реальные диапазоны
    ' ячеек листа, затем эти Range публикуются как controlPart для XML style pipeline.
    If Not private_TryRegisterControlPartSegments(ws, styleSegments) Then Exit Sub

#If ENALBE_STYLES Then
    private_ApplyStyleSegments ws, styleSegments
#End If

    If m_MergeSectionCells Then
        ' Merge выполняется после style pipeline: оформление уже назначено
        ' всему section-range, а объединение влияет только на отображение
        ' длинного заголовка и не меняет координаты следующих строк.
        private_MergeSectionRanges ws, styleSegments
    End If
End Sub

Private Function obj_IControl_Measure( _
    ByVal controlNode As Object, _
    ByRef outSpanRows As Long, _
    ByRef outSpanColls As Long, _
    Optional ByVal dataContext As Object _
) As Boolean
    obj_IControl_Measure = private_TryMeasureNode(controlNode, outSpanRows, outSpanColls)
End Function

Private Function private_TryRegisterControlSourceAliasSegments( _
    ByVal ws As Worksheet, _
    ByVal rowCount As Long, _
    ByVal styleSegments As Collection _
) As Boolean
    Dim visibleTables As Collection
    Dim headerRows As Collection
    Dim tableItem As Variant
    Dim tableViewItem As obj_TableViewItem
    Dim tableDynamic As obj_TableDynamic
    Dim segment As Object
    Dim tableIndex As Long
    Dim relativeStartRow As Long
    Dim relativeEndRow As Long
    Dim tableRange As Range

    If ws Is Nothing Or rowCount <= 0 Then Exit Function
    Set visibleTables = New Collection
    For Each tableItem In m_TableItems
        Set tableViewItem = Nothing
        If Not private_TryResolveTableViewItem(tableItem, tableViewItem) Then Exit Function
        If Not tableViewItem Is Nothing Then
            If tableViewItem.IsVisible() Then
                Set tableDynamic = tableViewItem.Model
                If Not tableDynamic Is Nothing Then visibleTables.Add tableDynamic
            End If
        End If
    Next tableItem

    Set headerRows = New Collection
    For Each segment In styleSegments
        If VBA.StrComp(VBA.CStr(segment("StyleKind")), "header", VBA.vbTextCompare) = 0 Then
            headerRows.Add VBA.CLng(segment("RowStart"))
        End If
    Next segment
    If headerRows.Count <> visibleTables.Count Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: source alias registration cannot map tables to header segments."
#End If
        Exit Function
    End If

    For tableIndex = 1 To visibleTables.Count
        Set tableDynamic = visibleTables.Item(tableIndex)
        relativeStartRow = VBA.CLng(headerRows.Item(tableIndex)) - 1
        If tableIndex < headerRows.Count Then
            relativeEndRow = VBA.CLng(headerRows.Item(tableIndex + 1)) - 2
        Else
            relativeEndRow = rowCount
        End If
        Set tableRange = ws.Range( _
            ws.Cells(m_RowStart + relativeStartRow - 1, m_ColStart), _
            ws.Cells(m_RowStart + relativeEndRow - 1, m_ColStart + tableDynamic.ColumnCount - 1))
        ' Старые/обычные TableDynamic могут не иметь source metadata.
        ' Они продолжают рендериться, просто селекторы sourceAlias к ним
        ' неприменимы.
        If VBA.Len(VBA.Trim$(tableDynamic.SourceAlias)) > 0 Or _
           VBA.Len(VBA.Trim$(tableDynamic.SourceAliasTemplate)) > 0 Then
            If Not ex_ControlPartsRuntime.fn_RegisterControlSourceAlias( _
                ws, "tablelist", m_ControlName, tableDynamic.SourceAlias, _
                tableDynamic.SourceAliasTemplate, tableRange) Then Exit Function
        End If
    Next tableIndex

    private_TryRegisterControlSourceAliasSegments = True
End Function

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(attrName))
        Case "itemssource", "itemvisibility", "renderaslistobject", "mergesectioncells", "tablename"
            obj_IControl_SupportsAttribute = True
    End Select
End Function

Private Function obj_IControl_IsConfigured() As Boolean
    obj_IControl_IsConfigured = m_IsConfigured
End Function

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
    ' Коллекция сегментов создается вместе с матрицей значений.
    ' Дальше writer-ы добавляют в нее section/header/data/diff/etc.
    Set outStyleSegments = New Collection
#End If

    currentOutputRow = 0

    ' Pass 2: fill matrix sequentially. Каждая таблица сама добавляет
    ' свои строки в valueBlock и соответствующие styleSegments.
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
    Dim rowStyleKind As String

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
                rowViewItem, tableDynamic, valueBlock, styleSegments, plannedRows, ioCurrentOutputRow) Then Exit Function

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
                private_CoerceDateCellsInMatrixRow _
                    valueBlock, ioCurrentOutputRow, tableDynamic
#If ENALBE_STYLES Then
                ' Обычная строка получает StyleKind=data; diff-строки получают
                ' более точный StyleKind из row.Desc: diffadded, diffmodified, ...
                rowStyleKind = private_ResolveDataRowStyleKind(row)
                private_AddStyleSegment styleSegments, rowStyleKind, tableDynamic.ColumnCount, ioCurrentOutputRow, ioCurrentOutputRow
                private_AddCellMetadataStyleSegments styleSegments, row, tableDynamic.ColumnCount, ioCurrentOutputRow
#End If
ContinueTableRow:
            Next tableRowIndex
            writeEnd = ioCurrentOutputRow
#If ENALBE_STYLES Then
            If writeEnd >= writeStart Then
                ' Строки с обычным стилем уже сгруппированы выше вместе с diff-строками.
                private_AddColumnFormatStyleSegments styleSegments, tableDynamic, writeStart, writeEnd
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
    ByVal tableDynamic As obj_TableDynamic, _
    ByRef valueBlock As Variant, _
    ByVal styleSegments As Collection, _
    ByVal plannedRows As Long, _
    ByRef ioCurrentOutputRow As Long _
) As Boolean
    Dim row As obj_Row
    Dim columnCount As Long
    Dim spacerIndex As Long
    Dim rowStyleKind As String

    If rowViewItem Is Nothing Then
        private_TryAppendRowViewData = True
        Exit Function
    End If

    If Not rowViewItem.IsVisible() Then
        private_TryAppendRowViewData = True
        Exit Function
    End If
    If tableDynamic Is Nothing Then Exit Function
    columnCount = tableDynamic.ColumnCount
    If columnCount <= 0 Then Exit Function

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
    private_CoerceDateCellsInMatrixRow _
        valueBlock, ioCurrentOutputRow, tableDynamic
#If ENALBE_STYLES Then
    ' RowViewItem идет тем же путем: значение пишется в valueBlock,
    ' а смысл строки/ячеек фиксируется в styleSegments.
    rowStyleKind = private_ResolveDataRowStyleKind(row)
    private_AddStyleSegment styleSegments, rowStyleKind, columnCount, ioCurrentOutputRow, ioCurrentOutputRow
    private_AddCellMetadataStyleSegments styleSegments, row, columnCount, ioCurrentOutputRow
    private_AddColumnFormatStyleSegments styleSegments, tableDynamic, ioCurrentOutputRow, ioCurrentOutputRow
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

Private Sub private_CoerceDateCellsInMatrixRow( _
    ByRef valueBlock As Variant, _
    ByVal matrixRow As Long, _
    ByVal tableDynamic As obj_TableDynamic _
)
    Dim colIndex As Long
    Dim colObj As obj_Column
    Dim rawValue As Variant
    Dim parsedDate As Date

    If tableDynamic Is Nothing Then Exit Sub
    If tableDynamic.Columns Is Nothing Then Exit Sub
    If matrixRow <= 0 Then Exit Sub

    For colIndex = 1 To tableDynamic.ColumnCount
        Set colObj = tableDynamic.Columns.Item(colIndex)
        If colObj Is Nothing Then GoTo ContinueColumn
        If VBA.InStr(1, colObj.FormatKind, "date", VBA.vbTextCompare) = 0 Then _
            GoTo ContinueColumn

        rawValue = valueBlock(matrixRow, colIndex)
        If private_TryCoerceDateSerial(rawValue, parsedDate) Then
            ' Value2 должен получить число, а не строку с похожим видом даты.
            ' Тогда дата сохраняет тип при копировании и группируется фильтром.
            valueBlock(matrixRow, colIndex) = VBA.CDbl(parsedDate)
        End If
ContinueColumn:
    Next colIndex
End Sub

Private Function private_TryCoerceDateSerial( _
    ByVal rawValue As Variant, _
    ByRef outDate As Date _
) As Boolean
    Dim valueText As String
    Dim parts As Variant
    Dim dayValue As Long
    Dim monthValue As Long
    Dim yearValue As Long
    Dim numericValue As Double

    On Error GoTo CleanFail
    If VBA.IsError(rawValue) Or VBA.IsNull(rawValue) Or VBA.IsEmpty(rawValue) Then _
        Exit Function
    valueText = VBA.Trim$(VBA.CStr(rawValue))
    If VBA.Len(valueText) = 0 Then Exit Function

    ' Сначала разбираем канонический формат явно, независимо от региональных
    ' настроек Windows/Excel. DateSerial дополнительно валидируем обратным чтением.
    parts = VBA.Split(valueText, ".")
    If UBound(parts) = 2 Then
        If VBA.IsNumeric(parts(0)) And VBA.IsNumeric(parts(1)) And _
           VBA.IsNumeric(parts(2)) Then
            dayValue = VBA.CLng(parts(0))
            monthValue = VBA.CLng(parts(1))
            yearValue = VBA.CLng(parts(2))
            If yearValue < 100 Then yearValue = 2000 + yearValue
            If dayValue >= 1 And dayValue <= 31 And _
               monthValue >= 1 And monthValue <= 12 Then
                outDate = VBA.DateSerial(yearValue, monthValue, dayValue)
                If VBA.Day(outDate) = dayValue And _
                   VBA.Month(outDate) = monthValue And _
                   VBA.Year(outDate) = yearValue Then
                    private_TryCoerceDateSerial = True
                    Exit Function
                End If
            End If
        End If
    End If

    ' Источники Comparing могут уже содержать serial date. Не прогоняем его
    ' через locale-зависимый CDate(String), если достаточно числового значения.
    If VBA.IsNumeric(valueText) Then
        numericValue = VBA.CDbl(valueText)
        If numericValue > 0 Then
            outDate = VBA.CDate(numericValue)
            private_TryCoerceDateSerial = True
        End If
    End If
    Exit Function

CleanFail:
    private_TryCoerceDateSerial = False
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
    ByVal rowEnd As Long, _
    Optional ByVal columnStart As Long = 1, _
    Optional ByVal columnEnd As Long = 0 _
)
    Dim segment As Object
    Dim lastSegment As Object

    If styleSegments Is Nothing Then Exit Sub
    If rowEnd < rowStart Then Exit Sub
    If columnCount <= 0 Then Exit Sub
    If columnStart <= 0 Then columnStart = 1
    If columnEnd <= 0 Then columnEnd = columnCount
    If columnEnd < columnStart Then Exit Sub
    If columnStart > columnCount Then Exit Sub
    If columnEnd > columnCount Then columnEnd = columnCount

    ' Segment хранит относительные координаты внутри valueBlock, не адреса Excel.
    ' Абсолютный Range строится позже через private_BuildSegmentRange.
    ' Merge adjacent segments with same style+width to reduce style operations later.
    If styleSegments.Count > 0 Then
        Set lastSegment = styleSegments(styleSegments.Count)
        If VBA.StrComp(VBA.CStr(lastSegment("StyleKind")), styleKind, VBA.vbTextCompare) = 0 Then
            If VBA.CLng(lastSegment("ColumnCount")) = columnCount Then
                If VBA.CLng(lastSegment("ColumnStart")) = columnStart Then
                    If VBA.CLng(lastSegment("ColumnEnd")) = columnEnd Then
                        If VBA.CLng(lastSegment("RowEnd")) + 1 = rowStart Then
                            lastSegment("RowEnd") = rowEnd
                            Exit Sub
                        End If
                    End If
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
    segment("ColumnStart") = columnStart
    segment("ColumnEnd") = columnEnd

    styleSegments.Add segment
End Sub

Private Sub private_ApplyStyleSegments(ByVal ws As Worksheet, ByVal styleSegments As Collection)
    Dim groupedRanges As Object
    Dim segment As Object
    Dim segmentRange As Range
    Dim groupedKey As String
    Dim groupedRange As Range
    Dim key As Variant
    Dim keyParts As Variant
    Dim styleKind As String
    Dim columnCount As Long
    Dim columnStart As Long
    Dim columnEnd As Long
    Dim backColor As Long
    Dim fontColor As Long
    Dim borderColor As Long
    Dim fontSize As Double
    Dim fontBold As Boolean
    Dim passIndex As Long
    Dim applyChangedCellStyle As Boolean

    If ws Is Nothing Then Exit Sub
    If styleSegments Is Nothing Then Exit Sub

    ' Group by (style kind + column count) and union ranges.
    Set groupedRanges = VBA.CreateObject("Scripting.Dictionary")
    groupedRanges.CompareMode = 1

    For Each segment In styleSegments
        styleKind = VBA.LCase$(VBA.CStr(segment("StyleKind")))
        columnCount = VBA.CLng(segment("ColumnCount"))
        columnStart = VBA.CLng(segment("ColumnStart"))
        columnEnd = VBA.CLng(segment("ColumnEnd"))

        Set segmentRange = private_BuildSegmentRange( _
            ws, _
            VBA.CLng(segment("RowStart")), _
            VBA.CLng(segment("RowEnd")), _
            columnCount, _
            columnStart, _
            columnEnd)
        If segmentRange Is Nothing Then GoTo ContinueSegment

        groupedKey = styleKind & "|" & VBA.CStr(columnCount) & "|" & VBA.CStr(columnStart) & "|" & VBA.CStr(columnEnd)

        If groupedRanges.Exists(groupedKey) Then
            Set groupedRange = groupedRanges(groupedKey)
            Set groupedRanges(groupedKey) = Application.Union(groupedRange, segmentRange)
        Else
            groupedRanges.Add groupedKey, segmentRange
        End If

ContinueSegment:
    Next segment

    ' Строковые diff-стили наносим первым проходом, а подсветку конкретных
    ' измененных ячеек вторым. Иначе Excel может перетереть cell-style стилем строки.
    For passIndex = 1 To 2
        applyChangedCellStyle = (passIndex = 2)
        For Each key In groupedRanges.Keys
            keyParts = VBA.Split(VBA.CStr(key), "|")
            If UBound(keyParts) < 0 Then GoTo ContinueGroup
            styleKind = VBA.CStr(keyParts(0))
            If ( _
                VBA.StrComp(styleKind, "diffchangedcell", VBA.vbTextCompare) = 0 Or _
                VBA.StrComp(styleKind, "diffchangedmodifiedoldcell", VBA.vbTextCompare) = 0 Or _
                VBA.StrComp(styleKind, "diffchangedmodifiednewcell", VBA.vbTextCompare) = 0 _
            ) <> applyChangedCellStyle Then GoTo ContinueGroup

            ' tag-* не является встроенным стилем renderer-а. Сегмент нужен
            ' только для публикации controlPart и оформляется XML pipeline-ом.
            If VBA.Left$(VBA.LCase$(styleKind), 4) = "tag-" Then GoTo ContinueGroup

            Set groupedRange = groupedRanges(VBA.CStr(key))
            If groupedRange Is Nothing Then GoTo ContinueGroup

            If VBA.StrComp(styleKind, "datelike", VBA.vbTextCompare) = 0 Then GoTo ContinueGroup

            If Not private_TryResolveStylePreset(styleKind, backColor, fontColor, borderColor, fontSize, fontBold) Then Exit Sub

            private_ApplyRowStyle groupedRange, backColor, fontColor, borderColor, fontSize, fontBold

ContinueGroup:
        Next key
    Next passIndex
End Sub


Private Function private_TryRegisterControlColumnAliasSegments( _
    ByVal ws As Worksheet, _
    ByRef valueBlock As Variant, _
    ByVal rowCount As Long, _
    ByVal columnCount As Long, _
    ByVal styleSegments As Collection _
) As Boolean
    Dim segment As Object
    Dim styleKind As String
    Dim visibleTables As Collection
    Dim tableItem As Variant
    Dim tableViewItem As obj_TableViewItem
    Dim tableDynamic As obj_TableDynamic
    Dim headerSegmentIndex As Long
    Dim tableColumnIndex As Long
    Dim colObj As obj_Column
    Dim sourceAliases As Collection
    Dim aliasItem As Variant
    Dim registeredAliasKeys As Object
    Dim columnStart As Long
    Dim columnEnd As Long
    Dim colIndex As Long
    Dim columnRange As Range
    Dim absCol As Long

    If ws Is Nothing Then Exit Function
    If IsEmpty(valueBlock) Then Exit Function
    If rowCount <= 0 Or columnCount <= 0 Then Exit Function
    If styleSegments Is Nothing Then
        private_TryRegisterControlColumnAliasSegments = True
        Exit Function
    End If

    Set visibleTables = New Collection
    For Each tableItem In m_TableItems
        Set tableViewItem = Nothing
        If Not private_TryResolveTableViewItem(tableItem, tableViewItem) Then Exit Function
        If tableViewItem Is Nothing Then GoTo ContinueTableItem
        If Not tableViewItem.IsVisible() Then GoTo ContinueTableItem

        Set tableDynamic = tableViewItem.Model
        If tableDynamic Is Nothing Then GoTo ContinueTableItem
        visibleTables.Add tableDynamic

ContinueTableItem:
    Next tableItem

    Set registeredAliasKeys = VBA.CreateObject("Scripting.Dictionary")
    registeredAliasKeys.CompareMode = 1

    For Each segment In styleSegments
        styleKind = VBA.LCase$(VBA.Trim$(VBA.CStr(segment("StyleKind"))))
        If VBA.StrComp(styleKind, "header", VBA.vbTextCompare) <> 0 Then GoTo ContinueSegment

        headerSegmentIndex = headerSegmentIndex + 1
        If headerSegmentIndex > visibleTables.Count Then GoTo ContinueSegment

        Set tableDynamic = visibleTables(headerSegmentIndex)
        If tableDynamic Is Nothing Then GoTo ContinueSegment
        If tableDynamic.ColumnCount <= 0 Then GoTo ContinueSegment

        columnStart = VBA.CLng(segment("ColumnStart"))
        columnEnd = VBA.CLng(segment("ColumnEnd"))

        For colIndex = columnStart To columnEnd
            If colIndex <= 0 Or colIndex > columnCount Then GoTo ContinueColumn

            tableColumnIndex = colIndex - columnStart + 1
            If tableColumnIndex <= 0 Or tableColumnIndex > tableDynamic.ColumnCount Then GoTo ContinueColumn

            Set colObj = Nothing
            Set colObj = tableDynamic.Columns.Item(tableColumnIndex)
            If colObj Is Nothing Then GoTo ContinueColumn

            absCol = m_ColStart + colIndex - 1
            Set columnRange = ws.Range( _
                ws.Cells(m_RowStart, absCol), _
                ws.Cells(m_RowStart + rowCount - 1, absCol))
            If columnRange Is Nothing Then GoTo ContinueColumn

            If Not private_TryRegisterControlColumnAliasOne( _
                ws, _
                columnRange, _
                registeredAliasKeys, _
                absCol, _
                colObj.Name) Then Exit Function

            Set sourceAliases = Nothing
            Set sourceAliases = colObj.Aliases
            If Not sourceAliases Is Nothing Then
                For Each aliasItem In sourceAliases
                    If Not private_TryRegisterControlColumnAliasOne( _
                        ws, _
                        columnRange, _
                        registeredAliasKeys, _
                        absCol, _
                        VBA.CStr(aliasItem)) Then Exit Function
                Next aliasItem
            End If

ContinueColumn:
        Next colIndex

ContinueSegment:
    Next segment

    private_TryRegisterControlColumnAliasSegments = True
End Function

Private Function private_TryRegisterControlColumnAliasOne( _
    ByVal ws As Worksheet, _
    ByVal columnRange As Range, _
    ByVal registeredAliasKeys As Object, _
    ByVal absCol As Long, _
    ByVal aliasText As String _
) As Boolean
    Dim normalizedAlias As String
    Dim registerKey As String

    If ws Is Nothing Then Exit Function
    If columnRange Is Nothing Then Exit Function

    normalizedAlias = VBA.LCase$(VBA.Trim$(aliasText))
    If VBA.Len(normalizedAlias) = 0 Then
        private_TryRegisterControlColumnAliasOne = True
        Exit Function
    End If

    registerKey = VBA.CStr(absCol) & "|" & normalizedAlias
    If Not registeredAliasKeys Is Nothing Then
        If registeredAliasKeys.Exists(registerKey) Then
            private_TryRegisterControlColumnAliasOne = True
            Exit Function
        End If
    End If

    If Not ex_ControlPartsRuntime.fn_RegisterControlColumnAlias( _
        ws, _
        "tablelist", _
        m_ControlName, _
        normalizedAlias, _
        columnRange) Then Exit Function

    If Not registeredAliasKeys Is Nothing Then
        registeredAliasKeys(registerKey) = True
    End If

    private_TryRegisterControlColumnAliasOne = True
End Function

Private Function private_TryRegisterControlPartSegments(ByVal ws As Worksheet, ByVal styleSegments As Collection) As Boolean
    Dim segment As Object
    Dim partName As String
    Dim segmentRange As Range

    If ws Is Nothing Then Exit Function
    If styleSegments Is Nothing Then
        private_TryRegisterControlPartSegments = True
        Exit Function
    End If

    ' Здесь styleSegments становятся внешними selector-частями:
    ' StyleKind=data -> part=rows, StyleKind=diffadded -> part=diffadded.
    For Each segment In styleSegments
        partName = private_MapStyleKindToControlPart(VBA.CStr(segment("StyleKind")))
        If VBA.Len(partName) = 0 Then GoTo ContinueSegment

        Set segmentRange = private_BuildSegmentRange( _
            ws, _
            VBA.CLng(segment("RowStart")), _
            VBA.CLng(segment("RowEnd")), _
            VBA.CLng(segment("ColumnCount")), _
            VBA.CLng(segment("ColumnStart")), _
            VBA.CLng(segment("ColumnEnd")))
        If segmentRange Is Nothing Then GoTo ContinueSegment

        If Not ex_ControlPartsRuntime.fn_RegisterControlPart( _
            ws, _
            "tablelist", _
            m_ControlName, _
            partName, _
            segmentRange) Then Exit Function

        ' Любая diff-строка одновременно является общей строкой данных.
        ' Поэтому part=rows покрывает и обычные data-строки, и diff-строки.
        If VBA.StrComp(partName, "rows", VBA.vbTextCompare) = 0 Or _
           VBA.Left$(VBA.LCase$(VBA.Trim$(partName)), 4) = "diff" Then
            If Not ex_ControlPartsRuntime.fn_RegisterControlPart( _
                ws, _
                "tablelist", _
                m_ControlName, _
                "rows", _
                segmentRange) Then Exit Function
        End If

ContinueSegment:
    Next segment

    private_TryRegisterControlPartSegments = True
End Function

Private Function private_MapStyleKindToControlPart(ByVal styleKind As String) As String
    ' StyleKind - внутреннее имя сегмента при сборке valueBlock.
    ' ControlPart - публичное имя, которым пользуется XML selector.
    styleKind = VBA.LCase$(VBA.Trim$(styleKind))
    If VBA.Left$(styleKind, 4) = "tag-" Then
        private_MapStyleKindToControlPart = styleKind
        Exit Function
    End If

    Select Case styleKind
        Case "section"
            private_MapStyleKindToControlPart = "section"
        Case "header"
            private_MapStyleKindToControlPart = "header"
        Case "data"
            private_MapStyleKindToControlPart = "rows"
        Case "diffadded"
            private_MapStyleKindToControlPart = "diffadded"
        Case "diffdeleted"
            private_MapStyleKindToControlPart = "diffdeleted"
        Case "diffduplicateleft"
            private_MapStyleKindToControlPart = "diffduplicateleft"
        Case "diffduplicateright"
            private_MapStyleKindToControlPart = "diffduplicateright"
        Case "diffmodified"
            private_MapStyleKindToControlPart = "diffmodified"
        Case "diffmodifiedold"
            private_MapStyleKindToControlPart = "diffmodifiedold"
        Case "diffmodifiednew"
            private_MapStyleKindToControlPart = "diffmodifiednew"
        Case "diffaddedmoved"
            private_MapStyleKindToControlPart = "diffaddedmoved"
        Case "diffdeletedmoved"
            private_MapStyleKindToControlPart = "diffdeletedmoved"
        Case "diffchangedcell"
            private_MapStyleKindToControlPart = "diffchangedcell"
        Case "diffchangedmodifiedoldcell"
            private_MapStyleKindToControlPart = "diffchangedmodifiedoldcell"
        Case "diffchangedmodifiednewcell"
            private_MapStyleKindToControlPart = "diffchangedmodifiednewcell"
        Case "diffellipsis"
            private_MapStyleKindToControlPart = "diffellipsis"
        Case "spacer"
            private_MapStyleKindToControlPart = "spacer"
        Case "tablebanner"
            private_MapStyleKindToControlPart = "itembanner"
        Case "rowbanner"
            private_MapStyleKindToControlPart = "rowbanner"
        Case "datelike"
            private_MapStyleKindToControlPart = "datelike"
    End Select
End Function

Private Function private_ResolveDataRowStyleKind(ByVal rowObj As obj_Row) As String
    Dim rowDesc As String

    private_ResolveDataRowStyleKind = "data"
    If rowObj Is Nothing Then Exit Function

    rowDesc = VBA.LCase$(VBA.Trim$(rowObj.Desc))
    If VBA.InStr(1, rowDesc, "diff:added-moved", VBA.vbTextCompare) > 0 Then
        private_ResolveDataRowStyleKind = "diffaddedmoved"
    ElseIf VBA.InStr(1, rowDesc, "diff:deleted-moved", VBA.vbTextCompare) > 0 Then
        private_ResolveDataRowStyleKind = "diffdeletedmoved"
    ElseIf VBA.InStr(1, rowDesc, "diff:modified-old", VBA.vbTextCompare) > 0 Then
        private_ResolveDataRowStyleKind = "diffmodifiedold"
    ElseIf VBA.InStr(1, rowDesc, "diff:modified-new", VBA.vbTextCompare) > 0 Then
        private_ResolveDataRowStyleKind = "diffmodifiednew"
    ElseIf VBA.InStr(1, rowDesc, "diff:added", VBA.vbTextCompare) > 0 Then
        private_ResolveDataRowStyleKind = "diffadded"
    ElseIf VBA.InStr(1, rowDesc, "diff:deleted", VBA.vbTextCompare) > 0 Then
        private_ResolveDataRowStyleKind = "diffdeleted"
    ElseIf VBA.InStr(1, rowDesc, "diff:duplicate-left", VBA.vbTextCompare) > 0 Then
        private_ResolveDataRowStyleKind = "diffduplicateleft"
    ElseIf VBA.InStr(1, rowDesc, "diff:duplicate-right", VBA.vbTextCompare) > 0 Then
        private_ResolveDataRowStyleKind = "diffduplicateright"
    ElseIf VBA.InStr(1, rowDesc, "diff:modified", VBA.vbTextCompare) > 0 Then
        private_ResolveDataRowStyleKind = "diffmodified"
    ElseIf VBA.InStr(1, rowDesc, "diff:ellipsis", VBA.vbTextCompare) > 0 Then
        private_ResolveDataRowStyleKind = "diffellipsis"
    End If
End Function

Private Sub private_AddCellMetadataStyleSegments( _
    ByVal styleSegments As Collection, _
    ByVal rowObj As obj_Row, _
    ByVal columnCount As Long, _
    ByVal relativeRow As Long _
)
    Dim colIndex As Long
    Dim cellObj As obj_Cell
    Dim cellDesc As String
    Dim cellTags As Collection
    Dim tagItem As Variant

    If styleSegments Is Nothing Then Exit Sub
    If rowObj Is Nothing Then Exit Sub
    If columnCount <= 0 Then Exit Sub
    If relativeRow <= 0 Then Exit Sub

    ' Помимо стиля всей diff-строки могут быть точечные cell-сегменты.
    ' Они регистрируются в той же коллекции, но с ColumnStart=ColumnEnd.
    For colIndex = 1 To columnCount
        Set cellObj = Nothing
        If Not rowObj.TryGetCellAt(colIndex, cellObj) Then GoTo ContinueCell
        If cellObj Is Nothing Then GoTo ContinueCell

        cellDesc = VBA.LCase$(VBA.Trim$(cellObj.Desc))
        If VBA.InStr(1, cellDesc, "diff:changed-modified-old", VBA.vbTextCompare) > 0 Then
            private_AddStyleSegment styleSegments, "diffchangedmodifiedoldcell", columnCount, relativeRow, relativeRow, colIndex, colIndex
        ElseIf VBA.InStr(1, cellDesc, "diff:changed-modified-new", VBA.vbTextCompare) > 0 Then
            private_AddStyleSegment styleSegments, "diffchangedmodifiednewcell", columnCount, relativeRow, relativeRow, colIndex, colIndex
        ElseIf VBA.InStr(1, cellDesc, "diff:changed", VBA.vbTextCompare) > 0 Then
            private_AddStyleSegment styleSegments, "diffchangedcell", columnCount, relativeRow, relativeRow, colIndex, colIndex
        End If

        ' Теги модели транслируются в универсальные controlPart-сегменты.
        ' Сам TableList не знает их предметного смысла и не задаёт им стиль.
        Set cellTags = cellObj.Tags
        If Not cellTags Is Nothing Then
            For Each tagItem In cellTags
                private_AddStyleSegment styleSegments, _
                    "tag-" & VBA.CStr(tagItem), columnCount, _
                    relativeRow, relativeRow, colIndex, colIndex
            Next tagItem
        End If
ContinueCell:
    Next colIndex
End Sub

Private Sub private_AddColumnFormatStyleSegments( _
    ByVal styleSegments As Collection, _
    ByVal tableDynamic As obj_TableDynamic, _
    ByVal relativeRowStart As Long, _
    ByVal relativeRowEnd As Long _
)
    Dim colIndex As Long
    Dim colObj As obj_Column
    Dim formatKind As String

    If styleSegments Is Nothing Then Exit Sub
    If tableDynamic Is Nothing Then Exit Sub
    If relativeRowStart <= 0 Or relativeRowEnd < relativeRowStart Then Exit Sub
    If tableDynamic.Columns Is Nothing Then Exit Sub

    ' Формат колонки тоже описывается как segment, потому что XML pipeline
    ' может адресовать его тем же механизмом controlPart, например datelike.
    For colIndex = 1 To tableDynamic.ColumnCount
        Set colObj = Nothing
        Set colObj = tableDynamic.Columns.Item(colIndex)
        If colObj Is Nothing Then GoTo ContinueColumn

        formatKind = VBA.LCase$(VBA.Trim$(colObj.FormatKind))
        If VBA.InStr(1, formatKind, "date", VBA.vbTextCompare) > 0 Then
            private_AddStyleSegment styleSegments, "datelike", tableDynamic.ColumnCount, relativeRowStart, relativeRowEnd, colIndex, colIndex
        End If

ContinueColumn:
    Next colIndex
End Sub

Private Function private_BuildSegmentRange( _
    ByVal ws As Worksheet, _
    ByVal relativeRowStart As Long, _
    ByVal relativeRowEnd As Long, _
    ByVal columnCount As Long, _
    Optional ByVal relativeColStart As Long = 1, _
    Optional ByVal relativeColEnd As Long = 0 _
) As Range
    Dim absRowStart As Long
    Dim absRowEnd As Long
    Dim absColStart As Long
    Dim absColEnd As Long

    If ws Is Nothing Then Exit Function
    If relativeRowStart <= 0 Or relativeRowEnd < relativeRowStart Then Exit Function
    If columnCount <= 0 Then Exit Function
    If relativeColStart <= 0 Then relativeColStart = 1
    If relativeColEnd <= 0 Then relativeColEnd = columnCount
    If relativeColEnd < relativeColStart Then Exit Function
    If relativeColStart > columnCount Then Exit Function
    If relativeColEnd > columnCount Then relativeColEnd = columnCount

    absRowStart = m_RowStart + relativeRowStart - 1
    absRowEnd = m_RowStart + relativeRowEnd - 1
    absColStart = m_ColStart + relativeColStart - 1
    absColEnd = m_ColStart + relativeColEnd - 1

    Set private_BuildSegmentRange = ws.Range( _
        ws.Cells(absRowStart, absColStart), _
        ws.Cells(absRowEnd, absColEnd))
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

        Case "diffadded"
            backColor = VBA.RGB(31, 86, 27)
            fontColor = VBA.RGB(245, 245, 245)
            borderColor = VBA.RGB(10, 10, 10)
            fontSize = 10
            fontBold = False

        Case "diffdeleted"
            backColor = VBA.RGB(204, 0, 0)
            fontColor = VBA.RGB(245, 245, 245)
            borderColor = VBA.RGB(10, 10, 10)
            fontSize = 10
            fontBold = False

        Case "diffduplicateleft"
            backColor = VBA.RGB(218, 165, 32)
            fontColor = VBA.RGB(26, 26, 26)
            borderColor = VBA.RGB(10, 10, 10)
            fontSize = 10
            fontBold = False

        Case "diffduplicateright"
            backColor = VBA.RGB(218, 165, 32)
            fontColor = VBA.RGB(26, 26, 26)
            borderColor = VBA.RGB(10, 10, 10)
            fontSize = 10
            fontBold = False

        Case "diffmodified"
            backColor = VBA.RGB(145, 31, 135)
            fontColor = VBA.RGB(245, 245, 245)
            borderColor = VBA.RGB(10, 10, 10)
            fontSize = 10
            fontBold = False

        Case "diffmodifiedold"
            backColor = VBA.RGB(128, 10, 77)
            fontColor = VBA.RGB(245, 245, 245)
            borderColor = VBA.RGB(10, 10, 10)
            fontSize = 10
            fontBold = False

        Case "diffmodifiednew"
            backColor = VBA.RGB(185, 5, 209)
            fontColor = VBA.RGB(245, 245, 245)
            borderColor = VBA.RGB(10, 10, 10)
            fontSize = 10
            fontBold = False

        Case "diffaddedmoved"
            backColor = VBA.RGB(78, 100, 36)
            fontColor = VBA.RGB(245, 245, 245)
            borderColor = VBA.RGB(10, 10, 10)
            fontSize = 10
            fontBold = False

        Case "diffdeletedmoved"
            backColor = VBA.RGB(107, 35, 35)
            fontColor = VBA.RGB(245, 245, 245)
            borderColor = VBA.RGB(10, 10, 10)
            fontSize = 10
            fontBold = False

        Case "diffchangedcell"
            backColor = VBA.RGB(224, 116, 214)
            fontColor = VBA.RGB(245, 245, 245)
            borderColor = VBA.RGB(10, 10, 10)
            fontSize = 10
            fontBold = False

        Case "diffchangedmodifiedoldcell"
            backColor = VBA.RGB(245, 245, 245)
            fontColor = VBA.RGB(26, 26, 26)
            borderColor = VBA.RGB(10, 10, 10)
            fontSize = 10
            fontBold = False

        Case "diffchangedmodifiednewcell"
            backColor = VBA.RGB(234, 109, 251)
            fontColor = VBA.RGB(26, 26, 26)
            borderColor = VBA.RGB(10, 10, 10)
            fontSize = 10
            fontBold = False

        Case "diffellipsis"
            backColor = VBA.RGB(33, 33, 33)
            fontColor = VBA.RGB(190, 190, 190)
            borderColor = VBA.RGB(10, 10, 10)
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
        targetColumn.FormatKind = sourceColumn.FormatKind
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

Private Function private_TryCreateRenderedListObject(ByVal ws As Worksheet, ByVal styleSegments As Collection) As Boolean
    Dim tableRange As Range
    Dim tableObj As ListObject
    Dim targetTableName As String

    If ws Is Nothing Then Exit Function

    If Not private_TryResolveListObjectRange(ws, styleSegments, tableRange) Then Exit Function
    If tableRange Is Nothing Then
        private_TryCreateRenderedListObject = True
        Exit Function
    End If

    On Error GoTo EH_TABLE
    Set tableObj = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    On Error GoTo 0

    targetTableName = private_BuildTableName(ws)
    If VBA.Len(targetTableName) > 0 Then
        On Error Resume Next
        tableObj.Name = targetTableName
        On Error GoTo 0
    End If

    On Error Resume Next
    tableObj.TableStyle = "TableStyleMedium2"
    tableObj.ShowAutoFilter = True
    tableObj.ShowTableStyleRowStripes = False
    tableObj.ShowTableStyleColumnStripes = False
    On Error GoTo 0

    m_RuntimeTableName = VBA.Trim$(tableObj.Name)
    private_TryCreateRenderedListObject = True
    Exit Function

EH_TABLE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "TableList: failed to create Excel table for control '" & m_ControlName & "': " & Err.Description
#End If
End Function

Private Function private_TryResolveListObjectRange( _
    ByVal ws As Worksheet, _
    ByVal styleSegments As Collection, _
    ByRef outRange As Range _
) As Boolean
    Dim segment As Object
    Dim styleKind As String
    Dim headerRowStart As Long
    Dim headerRowEnd As Long
    Dim headerColumnStart As Long
    Dim headerColumnEnd As Long
    Dim bodyRowEnd As Long
    Dim headerCount As Long

    If ws Is Nothing Then Exit Function
    If styleSegments Is Nothing Then
        private_TryResolveListObjectRange = True
        Exit Function
    End If

    For Each segment In styleSegments
        styleKind = VBA.LCase$(VBA.Trim$(VBA.CStr(segment("StyleKind"))))
        Select Case styleKind
            Case "header"
                headerCount = headerCount + 1
                If headerCount = 1 Then
                    headerRowStart = VBA.CLng(segment("RowStart"))
                    headerRowEnd = VBA.CLng(segment("RowEnd"))
                    headerColumnStart = VBA.CLng(segment("ColumnStart"))
                    headerColumnEnd = VBA.CLng(segment("ColumnEnd"))
                    bodyRowEnd = headerRowEnd
                End If

            Case "data", "diffadded", "diffdeleted", "diffduplicateleft", "diffduplicateright", "diffmodified", "diffmodifiedold", "diffmodifiednew", "diffaddedmoved", "diffdeletedmoved", "diffchangedcell", "diffchangedmodifiedoldcell", "diffchangedmodifiednewcell", "diffellipsis", "rowbanner"
                If headerCount = 1 Then
                    If VBA.CLng(segment("RowEnd")) > bodyRowEnd Then bodyRowEnd = VBA.CLng(segment("RowEnd"))
                End If
        End Select
    Next segment

    If headerCount <> 1 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "TableList: renderAsListObject requires exactly one header segment for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    Set outRange = ws.Range( _
        ws.Cells(m_RowStart + headerRowStart - 1, m_ColStart + headerColumnStart - 1), _
        ws.Cells(m_RowStart + bodyRowEnd - 1, m_ColStart + headerColumnEnd - 1))

    private_TryResolveListObjectRange = True
End Function

Private Function private_TryDeleteIntersectingTables(ByVal ws As Worksheet, ByVal boundsRange As Range) As Boolean
    Dim idx As Long
    Dim tableObj As ListObject

    If ws Is Nothing Then Exit Function
    If boundsRange Is Nothing Then Exit Function

    On Error GoTo EH_DELETE
    For idx = ws.ListObjects.Count To 1 Step -1
        Set tableObj = ws.ListObjects(idx)
        If Not tableObj Is Nothing Then
            If Not Application.Intersect(tableObj.Range, boundsRange) Is Nothing Then
                tableObj.Delete
            End If
        End If
    Next idx

    private_TryDeleteIntersectingTables = True
    Exit Function

EH_DELETE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "TableList: failed to delete intersecting tables for control '" & m_ControlName & "': " & Err.Description
#End If
End Function

Private Function private_BuildTableName(ByVal ws As Worksheet) As String
    Dim baseName As String

    baseName = VBA.Trim$(m_TableNameRaw)
    If VBA.Len(baseName) = 0 Then baseName = "tablelist" & m_ControlName
    baseName = private_SanitizeTableName(baseName)
    If VBA.Len(baseName) = 0 Then baseName = "tableListResult"
    private_BuildTableName = private_BuildUniqueTableName(ws, baseName)
End Function

Private Function private_BuildUniqueTableName(ByVal ws As Worksheet, ByVal baseName As String) As String
    Dim candidate As String
    Dim suffixIndex As Long

    If ws Is Nothing Then Exit Function
    baseName = VBA.Left$(VBA.Trim$(baseName), 240)
    If VBA.Len(baseName) = 0 Then baseName = "tableListResult"

    candidate = baseName
    suffixIndex = 1
    Do While private_TableNameExists(ws, candidate)
        suffixIndex = suffixIndex + 1
        candidate = VBA.Left$(baseName, 240 - VBA.Len(VBA.CStr(suffixIndex))) & VBA.CStr(suffixIndex)
    Loop

    private_BuildUniqueTableName = candidate
End Function

Private Function private_TableNameExists(ByVal ws As Worksheet, ByVal tableName As String) As Boolean
    Dim tableObj As ListObject

    If ws Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(tableName)) = 0 Then Exit Function

    For Each tableObj In ws.ListObjects
        If VBA.StrComp(tableObj.Name, tableName, VBA.vbTextCompare) = 0 Then
            private_TableNameExists = True
            Exit Function
        End If
    Next tableObj
End Function

Private Function private_SanitizeTableName(ByVal valueText As String) As String
    Dim i As Long
    Dim ch As String
    Dim outName As String

    valueText = VBA.Trim$(valueText)
    For i = 1 To VBA.Len(valueText)
        ch = VBA.Mid$(valueText, i, 1)
        If (ch >= "A" And ch <= "Z") Or _
           (ch >= "a" And ch <= "z") Or _
           (ch >= "0" And ch <= "9") Or _
           ch = "_" Then
            outName = outName & ch
        Else
            outName = outName & "_"
        End If
    Next i

    If VBA.Len(outName) = 0 Then Exit Function
    If Not ((VBA.Left$(outName, 1) >= "A" And VBA.Left$(outName, 1) <= "Z") Or _
            (VBA.Left$(outName, 1) >= "a" And VBA.Left$(outName, 1) <= "z") Or _
            VBA.Left$(outName, 1) = "_") Then
        outName = "tbl_" & outName
    End If

    If VBA.Len(outName) > 255 Then outName = VBA.Left$(outName, 255)
    private_SanitizeTableName = outName
End Function

Private Sub private_MergeSectionRanges( _
    ByVal ws As Worksheet, _
    ByVal styleSegments As Collection _
)
    Dim segment As Object
    Dim relativeRow As Long
    Dim sectionRange As Range

    If ws Is Nothing Then Exit Sub
    If styleSegments Is Nothing Then Exit Sub

    On Error Resume Next
    For Each segment In styleSegments
        If VBA.StrComp( _
            VBA.CStr(segment("StyleKind")), _
            "section", _
            VBA.vbTextCompare) = 0 Then
            For relativeRow = VBA.CLng(segment("RowStart")) To _
                VBA.CLng(segment("RowEnd"))
                Set sectionRange = ws.Range( _
                    ws.Cells(m_RowStart + relativeRow - 1, m_ColStart), _
                    ws.Cells( _
                        m_RowStart + relativeRow - 1, _
                        m_ColStart + VBA.CLng(segment("ColumnCount")) - 1))
                If sectionRange.Cells.CountLarge > 1 Then sectionRange.Merge
            Next relativeRow
        End If
    Next segment
    On Error GoTo 0
End Sub

Private Function private_TryReadOptionalBooleanAttr( _
    ByVal controlNode As Object, _
    ByVal attrName As String, _
    ByVal defaultValue As Boolean, _
    ByRef outValue As Boolean _
) As Boolean
    Dim rawText As String

    outValue = defaultValue
    If controlNode Is Nothing Then Exit Function

    rawText = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, attrName)))
    If VBA.Len(rawText) = 0 Then
        private_TryReadOptionalBooleanAttr = True
        Exit Function
    End If

    Select Case VBA.LCase$(rawText)
        Case "true", "1", "yes"
            outValue = True
        Case "false", "0", "no"
            outValue = False
        Case Else
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "TableList: attribute '" & attrName & "' must be boolean for control '" & m_ControlName & "'."
#End If
            Exit Function
    End Select

    private_TryReadOptionalBooleanAttr = True
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
