Attribute VB_Name = "ex_OutputFormattingPipeline"
Option Explicit

Private Const TIMELINE_MIN_COLUMN_WIDTH As Double = 5#
Private Const TIMELINE_MAX_COLUMN_WIDTH As Double = 30#
Private Const STYLE_STAGE_BANNERS As String = "banners"
Private Const DEFAULT_SEGMENT_HIGHLIGHT_COLOR As String = "#66CCFF"

Public Sub m_ApplySheetPipeline( _
    ByVal ws As Worksheet, _
    Optional ByVal resultFieldRanges As Collection = Nothing, _
    Optional ByVal cfgStyles As Object = Nothing, _
    Optional ByVal kindRanges As Object = Nothing, _
    Optional ByVal activeModeKey As String = vbNullString, _
    Optional ByVal autoHeightOnly As Boolean = False, _
    Optional ByVal runtimeLayers As Collection = Nothing _
)
    Dim resolvedModeKey As String
    Dim pipeline As Collection
    Dim stagePipeline As Collection
    Dim runtimeLayer As obj_StyleLayer
    Dim stageLayer As obj_StyleLayer

    If ws Is Nothing Then Exit Sub

    resolvedModeKey = Trim$(activeModeKey)
    If Len(resolvedModeKey) = 0 Then
        resolvedModeKey = ex_ConfigProfilesManager.m_GetActiveModeKey(ws_Dev)
    End If

    Set pipeline = ex_StylePipelineEngine.m_BuildColumnStylesPipeline( _
        resultFieldRanges, _
        cfgStyles, _
        resolvedModeKey, _
        ThisWorkbook, _
        ws.Name _
    )
    If Not runtimeLayers Is Nothing Then
        For Each runtimeLayer In runtimeLayers
            If Not runtimeLayer Is Nothing Then
                ex_StylePipelineEngine.m_AddLayer pipeline, runtimeLayer
            End If
        Next runtimeLayer
    End If
    ex_StylePipelineEngine.m_ApplyColumnStylesPipeline ws, resultFieldRanges, pipeline, resolvedModeKey, kindRanges, autoHeightOnly

    If mp_HasBannerKinds(kindRanges) Then
        Set stagePipeline = ex_StylePipelineEngine.m_CreatePipeline()
        For Each stageLayer In ex_StylePipelineEngine.m_LoadSheetPipelineLayers(ws.Name, ThisWorkbook, STYLE_STAGE_BANNERS)
            If Not stageLayer Is Nothing Then
                ex_StylePipelineEngine.m_AddLayer stagePipeline, stageLayer
            End If
        Next stageLayer
        If stagePipeline.Count > 0 Then
            ex_StylePipelineEngine.m_ApplyColumnStylesPipeline ws, resultFieldRanges, stagePipeline, resolvedModeKey, kindRanges, autoHeightOnly
        End If

        ' Banner fontColor rules are applied at row level and can overwrite rich-text colors.
        ' Reapply per-segment highlight after final style pass.
        mp_ReapplyBannerHighlightSegments ws, kindRanges
    End If
End Sub

Private Sub mp_ReapplyBannerHighlightSegments(ByVal ws As Worksheet, ByVal kindRanges As Object)
    Dim itemsMapObj As Object
    Dim sourceKey As Variant
    Dim itemsSourceObj As Object
    Dim bannerItems As Collection
    Dim bannerItem As Variant
    Dim bannerObj As Object
    Dim segmentItems As Object
    Dim kindEntries As Collection
    Dim kindCursors As Object
    Dim bannerKind As String
    Dim bodyText As String
    Dim displayText As String
    Dim currentIndex As Long
    Dim targetRow As Long
    Dim targetCol As Long
    Dim bannerCell As Range

    If ws Is Nothing Then Exit Sub
    If kindRanges Is Nothing Then Exit Sub

    If Not ex_ResultLayoutItemsRt.m_TryGetItemsSourcesMap(ws, itemsMapObj) Then Exit Sub
    If itemsMapObj Is Nothing Then Exit Sub

    Set kindCursors = CreateObject("Scripting.Dictionary")
    kindCursors.CompareMode = 1

    For Each sourceKey In itemsMapObj.Keys
        Set itemsSourceObj = Nothing
        On Error Resume Next
        Set itemsSourceObj = itemsMapObj(sourceKey)
        On Error GoTo 0

        If itemsSourceObj Is Nothing Then GoTo ContinueSource
        If TypeName(itemsSourceObj) <> "Collection" Then GoTo ContinueSource

        Set bannerItems = itemsSourceObj
        If bannerItems.Count = 0 Then GoTo ContinueSource

        For Each bannerItem In bannerItems
            If Not IsObject(bannerItem) Then GoTo ContinueItem
            Set bannerObj = bannerItem

            If Not mp_TryGetObjectStringValue(bannerObj, "Kind", bannerKind) Then GoTo ContinueItem
            bannerKind = LCase$(Trim$(bannerKind))
            If Len(bannerKind) = 0 Then GoTo ContinueItem
            If Not kindRanges.Exists(bannerKind) Then GoTo ContinueItem

            Set kindEntries = kindRanges(bannerKind)
            If kindEntries Is Nothing Then GoTo ContinueItem

            currentIndex = 1
            If kindCursors.Exists(bannerKind) Then currentIndex = CLng(kindCursors(bannerKind))
            If currentIndex < 1 Or currentIndex > kindEntries.Count Then GoTo ContinueItem

            If Not mp_TryResolveBannerCellFromKindEntry(kindEntries(currentIndex), targetRow, targetCol) Then GoTo ContinueItem
            kindCursors(bannerKind) = currentIndex + 1

            Set bannerCell = ws.Cells(targetRow, targetCol)
            displayText = CStr(bannerCell.Value)
            If Len(displayText) = 0 Then
                If mp_TryGetObjectStringValue(bannerObj, "Body", bodyText) Then
                    displayText = bodyText
                End If
            End If
            ex_Messaging.m_ApplyBannerAutoHeightForRange ws, bannerCell, displayText, bannerKind

            If Not mp_TryGetObjectValue(bannerObj, "HighlightSegments", segmentItems) Then GoTo ContinueItem
            If segmentItems Is Nothing Then GoTo ContinueItem
            If TypeName(segmentItems) <> "Collection" Then GoTo ContinueItem
            If segmentItems.Count = 0 Then GoTo ContinueItem

            If Not mp_TryGetObjectStringValue(bannerObj, "Body", bodyText) Then
                bodyText = displayText
            End If

            mp_ApplyCellHighlightSegments ws.Cells(targetRow, targetCol), bodyText, segmentItems
ContinueItem:
        Next bannerItem
ContinueSource:
    Next sourceKey
End Sub

Private Function mp_TryResolveBannerCellFromKindEntry( _
    ByVal rangeEntry As Variant, _
    ByRef outRow As Long, _
    ByRef outCol As Long _
) As Boolean
    Dim entryObj As Object

    If Not IsObject(rangeEntry) Then Exit Function

    Set entryObj = rangeEntry
    If entryObj Is Nothing Then Exit Function

    On Error Resume Next
    outRow = CLng(entryObj("RowStart"))
    outCol = CLng(entryObj("ColStart"))
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If outRow <= 0 Then Exit Function
    If outCol <= 0 Then outCol = 1
    mp_TryResolveBannerCellFromKindEntry = True
End Function

Private Sub mp_ApplyCellHighlightSegments( _
    ByVal targetCell As Range, _
    ByVal fullText As String, _
    ByVal highlightSegments As Object _
)
    Dim segment As Variant
    Dim segmentStart As Long
    Dim segmentLength As Long
    Dim colorHex As String
    Dim fontColor As Long

    If targetCell Is Nothing Then Exit Sub
    If Len(fullText) = 0 Then Exit Sub
    If highlightSegments Is Nothing Then Exit Sub
    If TypeName(highlightSegments) <> "Collection" Then Exit Sub
    If highlightSegments.Count = 0 Then Exit Sub

    For Each segment In highlightSegments
        If Not IsObject(segment) Then GoTo NextSegment

        On Error Resume Next
        segmentStart = CLng(segment("Start"))
        segmentLength = CLng(segment("Length"))
        colorHex = CStr(segment("ColorHex"))
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo NextSegment
        End If
        On Error GoTo 0

        If segmentLength <= 0 Then GoTo NextSegment
        If segmentStart < 1 Then GoTo NextSegment
        If segmentStart > Len(fullText) Then GoTo NextSegment
        If segmentStart + segmentLength - 1 > Len(fullText) Then
            segmentLength = Len(fullText) - segmentStart + 1
            If segmentLength <= 0 Then GoTo NextSegment
        End If

        If Len(Trim$(colorHex)) = 0 Then colorHex = DEFAULT_SEGMENT_HIGHLIGHT_COLOR
        If Not ex_XmlCore.m_TryParseColor(colorHex, fontColor) Then
            If Not ex_XmlCore.m_TryParseColor(DEFAULT_SEGMENT_HIGHLIGHT_COLOR, fontColor) Then GoTo NextSegment
        End If

        targetCell.Characters(segmentStart, segmentLength).Font.Color = fontColor
NextSegment:
    Next segment
End Sub

Private Function mp_TryGetObjectValue(ByVal source As Object, ByVal memberName As String, ByRef outObject As Object) As Boolean
    Set outObject = Nothing
    If source Is Nothing Then Exit Function
    If Len(Trim$(memberName)) = 0 Then Exit Function

    On Error Resume Next
    Set outObject = source(memberName)
    If Err.Number = 0 Then
        mp_TryGetObjectValue = Not (outObject Is Nothing)
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function mp_TryGetObjectStringValue(ByVal source As Object, ByVal memberName As String, ByRef outValue As String) As Boolean
    outValue = vbNullString
    If source Is Nothing Then Exit Function
    If Len(Trim$(memberName)) = 0 Then Exit Function

    On Error Resume Next
    outValue = CStr(source(memberName))
    If Err.Number = 0 Then
        mp_TryGetObjectStringValue = True
    Else
        Err.Clear
        outValue = vbNullString
    End If
    On Error GoTo 0
End Function

Private Function mp_HasBannerKinds(ByVal kindRanges As Object) As Boolean
    If kindRanges Is Nothing Then Exit Function

    mp_HasBannerKinds = _
        kindRanges.Exists("warningbanner") Or _
        kindRanges.Exists("errorbanner") Or _
        kindRanges.Exists("notebanner") Or _
        kindRanges.Exists("warningbannertitle") Or _
        kindRanges.Exists("errorbannertitle") Or _
        kindRanges.Exists("notebannertitle") Or _
        kindRanges.Exists("warningbannerlayout") Or _
        kindRanges.Exists("errorbannerlayout") Or _
        kindRanges.Exists("notebannerlayout") Or _
        kindRanges.Exists("warningbannerlayouttitle") Or _
        kindRanges.Exists("errorbannerlayouttitle") Or _
        kindRanges.Exists("notebannerlayouttitle")
End Function

Public Sub m_FormatAsTable(ByVal ws As Worksheet, ByVal startRow As Long, ByVal rowCount As Long, ByVal colCount As Long)
    Dim headerRange As Range
    Dim allRange As Range

    Set headerRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, colCount))
    Set allRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + rowCount - 1, colCount))

    allRange.Font.Name = "Segoe UI"
    allRange.Font.Size = 10
    headerRange.Font.Bold = True
    allRange.HorizontalAlignment = xlCenter
    allRange.VerticalAlignment = xlCenter
    allRange.EntireColumn.AutoFit
    allRange.AutoFilter

    ws.Activate
    ActiveWindow.FreezePanes = False
End Sub

Public Sub m_ApplyComparingStyleLayers( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal rowCount As Long, _
    ByVal colCount As Long, _
    ByRef baseStyle As t_BaseSheetStyle, _
    ByRef outputStyle As t_OutputSheetStyle, _
    ByVal hasOutputStyle As Boolean _
)
    Dim layerOrder As Variant
    Dim layerName As Variant
    Dim fullRowCount As Long

    If ws Is Nothing Then Exit Sub
    If Not ex_SheetStylesXmlProvider.m_GetLayerOrder(hasOutputStyle, layerOrder, ThisWorkbook) Then Exit Sub

    fullRowCount = startRow + rowCount - 1

    For Each layerName In layerOrder
        Select Case CStr(layerName)
            Case ex_SheetStylesXmlProvider.LAYER_BASE
                ex_SheetStylesXmlProvider.m_ApplyBaseLayer ws, fullRowCount, colCount, baseStyle
            Case ex_SheetStylesXmlProvider.LAYER_OUTPUT
                mp_ApplyOutputStyleToResult ws, startRow, rowCount, colCount, outputStyle
                ex_SheetStylesXmlProvider.m_ApplyStatusLayer ws, startRow, rowCount, colCount, outputStyle
        End Select
    Next layerName
End Sub

Public Sub m_ApplyViewZoneWrapText( _
    ByVal ws As Worksheet, _
    ByVal viewStartRow As Long, _
    ByVal viewEndRow As Long, _
    ByVal viewColCount As Long, _
    Optional ByVal wrapEnabled As Boolean = True, _
    Optional ByVal excludedRows As Collection = Nothing, _
    Optional ByVal excludedRows2 As Collection = Nothing _
)
    Dim targetRange As Range
    Dim rowIndex As Long
    Dim rowRange As Range
    Dim excludedMap As Object
    Dim excludedMap2 As Object

    If ws Is Nothing Then Exit Sub
    If viewStartRow < 1 Then Exit Sub
    If viewEndRow < viewStartRow Then Exit Sub
    If viewColCount < 1 Then Exit Sub

    Set targetRange = ws.Range(ws.Cells(viewStartRow, 1), ws.Cells(viewEndRow, viewColCount))
    targetRange.WrapText = wrapEnabled

    If excludedRows Is Nothing Then Exit Sub

    Set excludedMap = mp_BuildRowsMap(excludedRows)
    Set excludedMap2 = mp_BuildRowsMap(excludedRows2)

    For rowIndex = viewStartRow To viewEndRow
        If mp_RowIsInMap(excludedMap, rowIndex) Then
            Set rowRange = ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, viewColCount))
            rowRange.WrapText = False
        ElseIf mp_RowIsInMap(excludedMap2, rowIndex) Then
            Set rowRange = ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, viewColCount))
            rowRange.WrapText = False
        End If
    Next rowIndex
End Sub

Public Sub m_ApplyTimelineDataRowsHeight( _
    ByVal ws As Worksheet, _
    ByVal viewStartRow As Long, _
    ByVal viewEndRow As Long, _
    ByVal viewColCount As Long, _
    ByVal headerRows As Collection, _
    ByVal sectionRows As Collection, _
    Optional ByVal dataRowHeight As Double = 32 _
)
    Dim rowIndex As Long
    Dim headerRowsMap As Object
    Dim sectionRowsMap As Object
    Dim lastDataRow As Long

    If ws Is Nothing Then Exit Sub
    If viewStartRow < 1 Then Exit Sub
    If viewEndRow < viewStartRow Then Exit Sub
    If viewColCount < 1 Then Exit Sub
    If dataRowHeight <= 0 Then Exit Sub

    Set headerRowsMap = mp_BuildRowsMap(headerRows)
    Set sectionRowsMap = mp_BuildRowsMap(sectionRows)

    lastDataRow = mp_GetLastUsedRow(ws)
    If lastDataRow <= 0 Then Exit Sub
    If lastDataRow < viewStartRow Then Exit Sub
    If lastDataRow > viewEndRow Then lastDataRow = viewEndRow

    For rowIndex = viewStartRow To lastDataRow
        If mp_RowIsInMap(headerRowsMap, rowIndex) Then GoTo ContinueRow
        If mp_RowIsInMap(sectionRowsMap, rowIndex) Then GoTo ContinueRow
        ws.Rows(rowIndex).RowHeight = dataRowHeight
ContinueRow:
    Next rowIndex
End Sub

Public Sub m_ApplyTimelineHeaderRowsLayout( _
    ByVal ws As Worksheet, _
    ByVal headerRows As Collection, _
    ByVal viewColCount As Long _
)
    Dim rowId As Variant
    Dim rowIndex As Long
    Dim lastCol As Long

    If ws Is Nothing Then Exit Sub
    If headerRows Is Nothing Then Exit Sub
    If viewColCount < 1 Then Exit Sub

    For Each rowId In headerRows
        rowIndex = CLng(rowId)
        If rowIndex < 1 Then GoTo ContinueRow

        lastCol = mp_GetLastUsedColumnInRow(ws, rowIndex)
        If lastCol < 1 Then GoTo ContinueRow
        If lastCol > viewColCount Then lastCol = viewColCount

        ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, lastCol)).WrapText = True
        ws.Rows(rowIndex).AutoFit
ContinueRow:
    Next rowId
End Sub

Public Sub m_ApplyTimelineRowLayoutsByStyle( _
    ByVal ws As Worksheet, _
    ByVal viewStartRow As Long, _
    ByVal viewEndRow As Long, _
    ByVal viewColCount As Long, _
    ByVal headerRows As Collection, _
    ByVal sectionRows As Collection, _
    ByRef style As t_OutputSheetStyle, _
    Optional ByVal dataRowHeight As Double = 32 _
)
    Dim rowIndex As Long
    Dim rowRange As Range
    Dim lastDataRow As Long
    Dim headerRowsMap As Object
    Dim sectionRowsMap As Object

    If ws Is Nothing Then Exit Sub
    If viewStartRow < 1 Then Exit Sub
    If viewEndRow < viewStartRow Then Exit Sub
    If viewColCount < 1 Then Exit Sub

    lastDataRow = mp_GetLastUsedRow(ws)
    If lastDataRow <= 0 Then Exit Sub
    If lastDataRow < viewStartRow Then Exit Sub
    If lastDataRow > viewEndRow Then lastDataRow = viewEndRow

    Set headerRowsMap = mp_BuildRowsMap(headerRows)
    Set sectionRowsMap = mp_BuildRowsMap(sectionRows)

    ' Width should be applied first because overflow/auto-height calculations depend on it.
    mp_ApplyWidthToRowCollection ws, headerRows, viewColCount, style.HeaderWidth
    mp_ApplyWidthToRowCollection ws, sectionRows, viewColCount, style.SectionWidth
    mp_ApplyWidthToContentRows ws, viewStartRow, lastDataRow, viewColCount, headerRowsMap, sectionRowsMap, style.ContentWidth
    mp_ClampTimelineViewColumnWidths ws, viewColCount, TIMELINE_MIN_COLUMN_WIDTH, TIMELINE_MAX_COLUMN_WIDTH

    mp_ApplyOverflowAndHeightToRowCollection ws, headerRows, viewColCount, style.HeaderOverflow, style.HeaderAutoHeight, style.RowHeight
    mp_ApplyOverflowAndHeightToRowCollection ws, sectionRows, viewColCount, style.SectionOverflow, style.SectionAutoHeight, style.RowHeight

    For rowIndex = viewStartRow To lastDataRow
        If mp_RowIsInMap(headerRowsMap, rowIndex) Then GoTo ContinueContentRow
        If mp_RowIsInMap(sectionRowsMap, rowIndex) Then GoTo ContinueContentRow

        Set rowRange = ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, viewColCount))
        mp_ApplyOverflowToRange rowRange, style.ContentOverflow
        If style.ContentAutoHeight Then
            ws.Rows(rowIndex).AutoFit
        ElseIf dataRowHeight > 0 Then
            ws.Rows(rowIndex).RowHeight = dataRowHeight
        End If
ContinueContentRow:
    Next rowIndex
End Sub

Private Sub mp_ClampTimelineViewColumnWidths( _
    ByVal ws As Worksheet, _
    ByVal viewColCount As Long, _
    ByVal minWidth As Double, _
    ByVal maxWidth As Double _
)
    Dim colIndex As Long
    Dim currentWidth As Double

    If ws Is Nothing Then Exit Sub
    If viewColCount < 1 Then Exit Sub
    If minWidth <= 0 Then Exit Sub
    If maxWidth < minWidth Then Exit Sub

    For colIndex = 1 To viewColCount
        currentWidth = ws.Columns(colIndex).ColumnWidth
        If currentWidth < minWidth Then
            ws.Columns(colIndex).ColumnWidth = minWidth
        ElseIf currentWidth > maxWidth Then
            ws.Columns(colIndex).ColumnWidth = maxWidth
        End If
    Next colIndex
End Sub

Public Sub m_ApplyPartialMatchRowsAutoHeight( _
    ByVal ws As Worksheet, _
    ByVal partialMatchRowRanges As Collection _
)
    Dim i As Long
    Dim itemObj As Object
    Dim rowStart As Long
    Dim rowEnd As Long
    Dim rowIndex As Long
    Dim lastCol As Long
    Dim rowRange As Range

    If ws Is Nothing Then Exit Sub
    If partialMatchRowRanges Is Nothing Then Exit Sub
    If partialMatchRowRanges.Count = 0 Then Exit Sub

    For i = 1 To partialMatchRowRanges.Count
        Set itemObj = partialMatchRowRanges(i)
        If itemObj Is Nothing Then GoTo ContinueItem
        If Not itemObj.Exists("RowStart") Then GoTo ContinueItem
        If Not itemObj.Exists("RowEnd") Then GoTo ContinueItem

        rowStart = CLng(itemObj("RowStart"))
        rowEnd = CLng(itemObj("RowEnd"))
        If rowStart <= 0 Then GoTo ContinueItem
        If rowEnd < rowStart Then GoTo ContinueItem

        For rowIndex = rowStart To rowEnd
            lastCol = mp_GetLastUsedColumnInRow(ws, rowIndex)
            If lastCol < 1 Then GoTo ContinueRow
            Set rowRange = ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, lastCol))
            rowRange.WrapText = True
            ws.Rows(rowIndex).AutoFit
ContinueRow:
        Next rowIndex
ContinueItem:
    Next i
End Sub

Private Sub mp_ApplyOutputStyleToResult( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal rowCount As Long, _
    ByVal colCount As Long, _
    ByRef style As t_OutputSheetStyle _
)
    Dim targetRange As Range
    Dim headerRange As Range
    Dim outputColumnsRange As Range

    Set targetRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + rowCount - 1, colCount))
    Set headerRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, colCount))
    Set outputColumnsRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount))

    targetRange.Interior.Pattern = xlSolid
    targetRange.Interior.Color = style.ContentBackColor
    targetRange.Font.Color = style.ContentColor
    targetRange.Font.Name = style.FontName
    targetRange.Font.Size = style.FontSize
    targetRange.RowHeight = style.RowHeight
    targetRange.HorizontalAlignment = style.HorizontalAlignment
    targetRange.VerticalAlignment = style.VerticalAlignment
    mp_ApplyOverflowToRange targetRange, style.ContentOverflow

    headerRange.Interior.Pattern = xlSolid
    headerRange.Interior.Color = style.HeaderBackColor
    headerRange.Font.Color = style.HeaderColor
    headerRange.Font.Bold = style.HeaderBold
    mp_ApplyOverflowToRange headerRange, style.HeaderOverflow

    If style.ContentWidth > 0 Then outputColumnsRange.ColumnWidth = style.ContentWidth
    If style.HeaderWidth > 0 Then outputColumnsRange.ColumnWidth = style.HeaderWidth

    If style.ContentAutoHeight Then
        targetRange.EntireRow.AutoFit
    End If
    If style.HeaderAutoHeight Then
        ws.Rows(startRow).AutoFit
    End If
End Sub

Private Sub mp_ApplyTimelineOutputStyle(ByVal ws As Worksheet, ByVal headerRows As Collection, ByVal sectionRows As Collection, ByRef style As t_OutputSheetStyle)
    Dim usedRows As Long
    Dim usedCols As Long
    Dim usedRange As Range
    Dim rowId As Variant
    Dim rowIndex As Long
    Dim lastCol As Long
    Dim titleRange As Range
    Dim sectionFillRange As Range
    Dim sectionTitleCols As Long

    If ws Is Nothing Then Exit Sub
    If Not ex_SheetStylesXmlProvider.m_GetUsedRangeSize(ws, usedRows, usedCols) Then Exit Sub

    Set usedRange = ws.Range(ws.Cells(1, 1), ws.Cells(usedRows, usedCols))

    usedRange.Interior.Pattern = xlSolid
    usedRange.Interior.Color = style.ContentBackColor
    usedRange.Font.Name = style.FontName
    usedRange.Font.Size = style.FontSize
    usedRange.Font.Color = style.ContentColor
    usedRange.HorizontalAlignment = style.HorizontalAlignment
    usedRange.VerticalAlignment = style.VerticalAlignment
    ws.Rows("1:" & CStr(usedRows)).RowHeight = style.RowHeight

    ' Header/section rows should not wrap before AutoFit, otherwise column widths can be over-expanded.
    For Each rowId In headerRows
        rowIndex = CLng(rowId)
        lastCol = mp_GetLastUsedColumnInRow(ws, rowIndex)
        If lastCol > 0 Then
            ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, lastCol)).WrapText = False
        End If
    Next rowId

    For Each rowId In sectionRows
        rowIndex = CLng(rowId)
        lastCol = mp_GetLastUsedColumnInRow(ws, rowIndex)
        If lastCol > 0 Then
            ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, lastCol)).WrapText = False
        End If
    Next rowId

    usedRange.EntireColumn.AutoFit

    For Each rowId In sectionRows
        rowIndex = CLng(rowId)
        Set sectionFillRange = ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, usedCols))
        sectionTitleCols = style.SectionMergeColumns
        If sectionTitleCols < 1 Then sectionTitleCols = 1
        If sectionTitleCols > usedCols Then sectionTitleCols = usedCols
        Set titleRange = ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, sectionTitleCols))
        titleRange.UnMerge
        titleRange.Merge
        titleRange.HorizontalAlignment = style.HorizontalAlignment
        titleRange.VerticalAlignment = style.VerticalAlignment
        sectionFillRange.Interior.Pattern = xlSolid
        sectionFillRange.Interior.Color = style.SectionBackColor
        titleRange.Font.Bold = style.SectionBold
        titleRange.Font.Color = style.SectionColor
    Next rowId

    For Each rowId In headerRows
        rowIndex = CLng(rowId)
        lastCol = mp_GetLastUsedColumnInRow(ws, rowIndex)
        If lastCol > 0 Then
            Set titleRange = ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, lastCol))
            titleRange.Interior.Pattern = xlSolid
            titleRange.Interior.Color = style.HeaderBackColor
            titleRange.Font.Bold = style.HeaderBold
            titleRange.Font.Color = style.HeaderColor
        End If
    Next rowId
End Sub

Private Function mp_GetLastUsedColumnInRow(ByVal ws As Worksheet, ByVal rowIndex As Long) As Long
    If ws Is Nothing Then Exit Function
    If rowIndex <= 0 Then Exit Function

    mp_GetLastUsedColumnInRow = ws.Cells(rowIndex, ws.Columns.Count).End(xlToLeft).Column
    If mp_GetLastUsedColumnInRow = 1 Then
        If Len(Trim$(CStr(ws.Cells(rowIndex, 1).Value))) = 0 Then
            mp_GetLastUsedColumnInRow = 0
        End If
    End If
End Function

Private Sub mp_ApplyOverflowAndHeightToRowCollection( _
    ByVal ws As Worksheet, _
    ByVal rowsCollection As Collection, _
    ByVal viewColCount As Long, _
    ByVal overflowStyle As String, _
    ByVal autoHeight As Boolean, _
    ByVal fixedRowHeight As Double _
)
    Dim rowId As Variant
    Dim rowIndex As Long
    Dim lastCol As Long
    Dim rowRange As Range

    If ws Is Nothing Then Exit Sub
    If rowsCollection Is Nothing Then Exit Sub
    If viewColCount < 1 Then Exit Sub

    For Each rowId In rowsCollection
        rowIndex = CLng(rowId)
        If rowIndex < 1 Then GoTo ContinueRow

        lastCol = mp_GetLastUsedColumnInRow(ws, rowIndex)
        If lastCol < 1 Then GoTo ContinueRow
        If lastCol > viewColCount Then lastCol = viewColCount

        Set rowRange = ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, lastCol))
        mp_ApplyOverflowToRange rowRange, overflowStyle
        If autoHeight Then
            ws.Rows(rowIndex).AutoFit
        ElseIf fixedRowHeight > 0 Then
            ws.Rows(rowIndex).RowHeight = fixedRowHeight
        End If
ContinueRow:
    Next rowId
End Sub

Private Sub mp_ApplyOverflowToRange(ByVal targetRange As Range, ByVal overflowStyle As String)
    Dim normalized As String

    If targetRange Is Nothing Then Exit Sub

    normalized = LCase$(Trim$(overflowStyle))
    If Len(normalized) = 0 Then normalized = "clip"

    Select Case normalized
        Case "wrap"
            targetRange.WrapText = True
            targetRange.ShrinkToFit = False
        Case "shrink"
            targetRange.WrapText = False
            targetRange.ShrinkToFit = True
        Case Else
            targetRange.WrapText = False
            targetRange.ShrinkToFit = False
    End Select
End Sub

Private Sub mp_ApplyWidthToRowCollection( _
    ByVal ws As Worksheet, _
    ByVal rowsCollection As Collection, _
    ByVal viewColCount As Long, _
    ByVal widthValue As Double _
)
    Dim columnsMap As Object
    Dim rowId As Variant
    Dim colKey As Variant
    Dim rowIndex As Long
    Dim lastCol As Long
    Dim colIndex As Long

    If ws Is Nothing Then Exit Sub
    If rowsCollection Is Nothing Then Exit Sub
    If viewColCount < 1 Then Exit Sub
    If widthValue <= 0 Then Exit Sub

    Set columnsMap = CreateObject("Scripting.Dictionary")
    columnsMap.CompareMode = 0

    For Each rowId In rowsCollection
        rowIndex = CLng(rowId)
        If rowIndex < 1 Then GoTo ContinueRow

        lastCol = mp_GetLastUsedColumnInRow(ws, rowIndex)
        If lastCol < 1 Then GoTo ContinueRow
        If lastCol > viewColCount Then lastCol = viewColCount

        For colIndex = 1 To lastCol
            columnsMap(CStr(colIndex)) = True
        Next colIndex
ContinueRow:
    Next rowId

    For Each colKey In columnsMap.Keys
        ws.Columns(CLng(colKey)).ColumnWidth = widthValue
    Next colKey
End Sub

Private Sub mp_ApplyWidthToContentRows( _
    ByVal ws As Worksheet, _
    ByVal viewStartRow As Long, _
    ByVal viewEndRow As Long, _
    ByVal viewColCount As Long, _
    ByVal headerRowsMap As Object, _
    ByVal sectionRowsMap As Object, _
    ByVal widthValue As Double _
)
    Dim columnsMap As Object
    Dim colKey As Variant
    Dim rowIndex As Long
    Dim lastCol As Long
    Dim colIndex As Long

    If ws Is Nothing Then Exit Sub
    If viewStartRow < 1 Then Exit Sub
    If viewEndRow < viewStartRow Then Exit Sub
    If viewColCount < 1 Then Exit Sub
    If widthValue <= 0 Then Exit Sub

    Set columnsMap = CreateObject("Scripting.Dictionary")
    columnsMap.CompareMode = 0

    For rowIndex = viewStartRow To viewEndRow
        If mp_RowIsInMap(headerRowsMap, rowIndex) Then GoTo ContinueRow
        If mp_RowIsInMap(sectionRowsMap, rowIndex) Then GoTo ContinueRow

        lastCol = mp_GetLastUsedColumnInRow(ws, rowIndex)
        If lastCol < 1 Then GoTo ContinueRow
        If lastCol > viewColCount Then lastCol = viewColCount

        For colIndex = 1 To lastCol
            columnsMap(CStr(colIndex)) = True
        Next colIndex
ContinueRow:
    Next rowIndex

    For Each colKey In columnsMap.Keys
        ws.Columns(CLng(colKey)).ColumnWidth = widthValue
    Next colKey
End Sub

Private Function mp_RowIsListed(ByVal rowsCollection As Collection, ByVal rowIndex As Long) As Boolean
    Dim itemValue As Variant

    If rowsCollection Is Nothing Then Exit Function

    For Each itemValue In rowsCollection
        If CLng(itemValue) = rowIndex Then
            mp_RowIsListed = True
            Exit Function
        End If
    Next itemValue
End Function

Private Function mp_BuildRowsMap(ByVal rowsCollection As Collection) As Object
    Dim rowsMap As Object
    Dim itemValue As Variant

    Set rowsMap = CreateObject("Scripting.Dictionary")
    rowsMap.CompareMode = 0

    If rowsCollection Is Nothing Then
        Set mp_BuildRowsMap = rowsMap
        Exit Function
    End If

    For Each itemValue In rowsCollection
        rowsMap(CStr(CLng(itemValue))) = True
    Next itemValue

    Set mp_BuildRowsMap = rowsMap
End Function

Private Function mp_RowIsInMap(ByVal rowsMap As Object, ByVal rowIndex As Long) As Boolean
    If rowsMap Is Nothing Then Exit Function
    mp_RowIsInMap = rowsMap.Exists(CStr(rowIndex))
End Function

Private Function mp_GetLastUsedRow(ByVal ws As Worksheet) As Long
    Dim lastCell As Range

    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If lastCell Is Nothing Then Exit Function
    mp_GetLastUsedRow = lastCell.Row
End Function
