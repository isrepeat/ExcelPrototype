Attribute VB_Name = "ex_OutputFormattingPipeline"
Option Explicit

Private Const TIMELINE_MIN_COLUMN_WIDTH As Double = 5#
Private Const TIMELINE_MAX_COLUMN_WIDTH As Double = 30#

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

Public Sub m_ApplyTimelineStyleLayers( _
    ByVal ws As Worksheet, _
    ByVal headerRows As Collection, _
    ByVal sectionRows As Collection, _
    Optional ByVal resultFieldRanges As Collection = Nothing _
)
    Dim workflowSteps As Collection
    Dim stepName As Variant
    Dim activeModeKey As String
    Dim pageName As String
    Dim rowKindRanges As Object

    If ws Is Nothing Then Exit Sub
    pageName = Trim$(ws.Name)
    activeModeKey = ex_ConfigProfilesManager.m_GetActiveModeKey(ws_Dev)
    If Not ex_StylePipelineEngine.m_GetRenderWorkflowStepOrder(pageName, "personalCardTimeline", workflowSteps, ThisWorkbook) Then
        Err.Raise vbObjectError + 1737, "ex_OutputFormattingPipeline", _
            "Failed to resolve style workflow steps for workflow 'personalCardTimeline' and page '" & pageName & "'."
    End If

    Set rowKindRanges = mp_BuildTimelineRowKindRanges(headerRows, sectionRows, resultFieldRanges)

    For Each stepName In workflowSteps
        Select Case LCase$(Trim$(CStr(stepName)))
            Case "base"
                mp_ApplyTimelineStageRules ws, activeModeKey, "base", resultFieldRanges, rowKindRanges
            Case "output"
                mp_ApplyTimelineStageRules ws, activeModeKey, "output", resultFieldRanges, rowKindRanges
            Case Else
                Err.Raise vbObjectError + 1734, "ex_OutputFormattingPipeline", _
                    "Unsupported workflow step '" & CStr(stepName) & "' in workflow 'personalCardTimeline'."
        End Select
    Next stepName
End Sub

Private Sub mp_ApplyTimelineStageRules( _
    ByVal ws As Worksheet, _
    ByVal activeModeKey As String, _
    ByVal stageName As String, _
    Optional ByVal resultFieldRanges As Collection = Nothing, _
    Optional ByVal rowKindRanges As Object = Nothing _
)
    Dim stagePipeline As Collection

    If ws Is Nothing Then Exit Sub
    If Len(Trim$(stageName)) = 0 Then Exit Sub

    Set stagePipeline = ex_StylePipelineEngine.m_LoadSheetPipelineStageLayers(ws.Name, stageName, ThisWorkbook)
    If stagePipeline Is Nothing Then Exit Sub
    If stagePipeline.Count = 0 Then Exit Sub

    ex_StylePipelineEngine.m_ApplyColumnStylesPipeline ws, resultFieldRanges, stagePipeline, activeModeKey, rowKindRanges
End Sub

Public Sub m_ApplyOutputPanelLayers( _
    ByVal ws As Worksheet, _
    ByRef outputStyle As t_OutputSheetStyle, _
    ByVal hasOutputStyle As Boolean, _
    ByVal viewStartRow As Long, _
    ByVal viewEndRow As Long, _
    ByVal viewColCount As Long _
)
    If ws Is Nothing Then Exit Sub
    If Not hasOutputStyle Then Exit Sub

    ex_OutputPanel.m_RenderForSheet ws, outputStyle
    ex_OutputPanel.m_ApplyFixedWidthViewZoneLayer ws, outputStyle, viewStartRow, viewEndRow, viewColCount
End Sub

Public Sub m_ApplyConfigNoteStyleLayer( _
    ByVal ws As Worksheet, _
    ByVal resultFieldRanges As Collection, _
    ByVal cfgNotes As Object _
)
    Dim styleTagsByMapKey As Object
    Dim activeModeKey As String
    Dim pipeline As Collection

    If ws Is Nothing Then Exit Sub
    If resultFieldRanges Is Nothing Then Exit Sub
    If cfgNotes Is Nothing Then Exit Sub

    activeModeKey = ex_ConfigProfilesManager.m_GetActiveModeKey(ws_Dev)
    Set styleTagsByMapKey = ex_ConfigProfilesManager.m_GetActiveProfileStyleTagsByKey(ws_Dev)
    Set pipeline = ex_StylePipelineEngine.m_BuildColumnStylesPipeline( _
        resultFieldRanges, _
        cfgNotes, _
        styleTagsByMapKey, _
        activeModeKey, _
        ThisWorkbook, _
        ws.Name _
    )

    ex_StylePipelineEngine.m_ApplyColumnStylesPipeline ws, resultFieldRanges, pipeline, activeModeKey
End Sub

Public Sub m_ApplyTimelinePostLayoutStyleLayers( _
    ByVal ws As Worksheet, _
    ByVal resultFieldRanges As Collection, _
    ByVal cfgNotes As Object _
)
    Dim workflowSteps As Collection
    Dim stepName As Variant
    Dim activeModeKey As String
    Dim pageName As String
    Dim stagePipeline As Collection
    Dim rowKindRanges As Object
    Dim configNoteStylesApplied As Boolean

    If ws Is Nothing Then Exit Sub

    pageName = Trim$(ws.Name)
    activeModeKey = ex_ConfigProfilesManager.m_GetActiveModeKey(ws_Dev)
    If Not ex_StylePipelineEngine.m_GetRenderWorkflowStepOrder(pageName, "personalCardPostLayout", workflowSteps, ThisWorkbook) Then
        Err.Raise vbObjectError + 1738, "ex_OutputFormattingPipeline", _
            "Failed to resolve style workflow steps for workflow 'personalCardPostLayout' and page '" & pageName & "'."
    End If

    For Each stepName In workflowSteps
        Select Case LCase$(Trim$(CStr(stepName)))
            Case "confignotestyles"
                m_ApplyConfigNoteStyleLayer ws, resultFieldRanges, cfgNotes
                configNoteStylesApplied = True
            Case "postlayout"
                If Not configNoteStylesApplied Then
                    m_ApplyConfigNoteStyleLayer ws, resultFieldRanges, cfgNotes
                    configNoteStylesApplied = True
                End If
                Set rowKindRanges = mp_BuildConfigNoteRowKindRanges(resultFieldRanges, cfgNotes)
                Set stagePipeline = ex_StylePipelineEngine.m_LoadSheetPipelineStageLayers(pageName, "postLayout", ThisWorkbook)
                ex_StylePipelineEngine.m_ApplyColumnStylesPipeline ws, resultFieldRanges, stagePipeline, activeModeKey, rowKindRanges
            Case Else
                Err.Raise vbObjectError + 1735, "ex_OutputFormattingPipeline", _
                    "Unsupported workflow step '" & CStr(stepName) & "' in workflow 'personalCardPostLayout'."
        End Select
    Next stepName
End Sub

Public Sub m_ApplyTimelinePostWarningsStyleLayers( _
    ByVal ws As Worksheet, _
    ByVal partialMatchRowRanges As Collection _
)
    Dim workflowSteps As Collection
    Dim stepName As Variant
    Dim activeModeKey As String
    Dim pageName As String
    Dim stagePipeline As Collection
    Dim rowKindRanges As Object

    If ws Is Nothing Then Exit Sub

    pageName = Trim$(ws.Name)
    activeModeKey = ex_ConfigProfilesManager.m_GetActiveModeKey(ws_Dev)
    If Not ex_StylePipelineEngine.m_GetRenderWorkflowStepOrder(pageName, "personalCardPostWarnings", workflowSteps, ThisWorkbook) Then
        Err.Raise vbObjectError + 1739, "ex_OutputFormattingPipeline", _
            "Failed to resolve style workflow steps for workflow 'personalCardPostWarnings' and page '" & pageName & "'."
    End If

    For Each stepName In workflowSteps
        Select Case LCase$(Trim$(CStr(stepName)))
            Case "partialmatchautoheight", "postwarnings"
            Case Else
                Err.Raise vbObjectError + 1736, "ex_OutputFormattingPipeline", _
                    "Unsupported workflow step '" & CStr(stepName) & "' in workflow 'personalCardPostWarnings'."
        End Select

        If rowKindRanges Is Nothing Then
            Set rowKindRanges = mp_BuildPartialMatchRowKindRanges(partialMatchRowRanges)
        End If

        Set stagePipeline = ex_StylePipelineEngine.m_LoadSheetPipelineStageLayers(pageName, LCase$(Trim$(CStr(stepName))), ThisWorkbook)
        ex_StylePipelineEngine.m_ApplyColumnStylesPipeline ws, Nothing, stagePipeline, activeModeKey, rowKindRanges
    Next stepName
End Sub

Private Function mp_BuildTimelineRowKindRanges( _
    ByVal headerRows As Collection, _
    ByVal sectionRows As Collection, _
    ByVal resultFieldRanges As Collection _
) As Object
    Dim result As Object
    Dim headerRowsMap As Object
    Dim sectionRowsMap As Object
    Dim contentRowsMap As Object
    Dim target As Object
    Dim rowStart As Long
    Dim rowEnd As Long
    Dim rowIndex As Long
    Dim rowKey As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1

    Set headerRowsMap = mp_BuildRowsMap(headerRows)
    Set sectionRowsMap = mp_BuildRowsMap(sectionRows)
    Set contentRowsMap = CreateObject("Scripting.Dictionary")
    contentRowsMap.CompareMode = 0

    If Not resultFieldRanges Is Nothing Then
        For Each target In resultFieldRanges
            If target Is Nothing Then GoTo ContinueTarget
            If Not target.Exists("RowStart") Then GoTo ContinueTarget
            If Not target.Exists("RowEnd") Then GoTo ContinueTarget

            rowStart = CLng(target("RowStart"))
            rowEnd = CLng(target("RowEnd"))
            If rowStart <= 0 Then GoTo ContinueTarget
            If rowEnd < rowStart Then rowEnd = rowStart

            For rowIndex = rowStart To rowEnd
                rowKey = CStr(rowIndex)
                If headerRowsMap.Exists(rowKey) Then GoTo ContinueRow
                If sectionRowsMap.Exists(rowKey) Then GoTo ContinueRow
                contentRowsMap(rowKey) = True
ContinueRow:
            Next rowIndex
ContinueTarget:
        Next target
    End If

    Set result("header") = mp_RowsMapToRangeCollection(headerRowsMap)
    Set result("section") = mp_RowsMapToRangeCollection(sectionRowsMap)
    Set result("content") = mp_RowsMapToRangeCollection(contentRowsMap)
    Set mp_BuildTimelineRowKindRanges = result
End Function

Private Function mp_RowsMapToRangeCollection(ByVal rowsMap As Object) As Collection
    Dim result As Collection
    Dim keys() As Long
    Dim keyValue As Variant
    Dim i As Long
    Dim count As Long
    Dim rangeItem As Object
    Dim runStart As Long
    Dim runEnd As Long

    Set result = New Collection
    If rowsMap Is Nothing Then
        Set mp_RowsMapToRangeCollection = result
        Exit Function
    End If
    If rowsMap.Count = 0 Then
        Set mp_RowsMapToRangeCollection = result
        Exit Function
    End If

    ReDim keys(1 To rowsMap.Count)
    For Each keyValue In rowsMap.Keys
        count = count + 1
        keys(count) = CLng(keyValue)
    Next keyValue
    If count = 0 Then
        Set mp_RowsMapToRangeCollection = result
        Exit Function
    End If

    mp_SortLongArray keys

    runStart = keys(1)
    runEnd = runStart
    For i = 2 To UBound(keys)
        If keys(i) = runEnd + 1 Then
            runEnd = keys(i)
        Else
            Set rangeItem = CreateObject("Scripting.Dictionary")
            rangeItem.CompareMode = 1
            rangeItem("RowStart") = runStart
            rangeItem("RowEnd") = runEnd
            result.Add rangeItem

            runStart = keys(i)
            runEnd = runStart
        End If
    Next i

    Set rangeItem = CreateObject("Scripting.Dictionary")
    rangeItem.CompareMode = 1
    rangeItem("RowStart") = runStart
    rangeItem("RowEnd") = runEnd
    result.Add rangeItem

    Set mp_RowsMapToRangeCollection = result
End Function

Private Sub mp_SortLongArray(ByRef values() As Long)
    Dim i As Long
    Dim j As Long
    Dim tmp As Long

    If UBound(values) <= LBound(values) Then Exit Sub

    For i = LBound(values) To UBound(values) - 1
        For j = i + 1 To UBound(values)
            If values(j) < values(i) Then
                tmp = values(i)
                values(i) = values(j)
                values(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Function mp_BuildConfigNoteRowKindRanges( _
    ByVal resultFieldRanges As Collection, _
    ByVal cfgNotes As Object _
) As Object
    Dim result As Object
    Dim configNoteRanges As Collection
    Dim dedupe As Object
    Dim target As Object
    Dim mapKey As String
    Dim noteText As String
    Dim rowStart As Long
    Dim rowEnd As Long
    Dim dedupeKey As String
    Dim rowItem As Object

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1
    Set configNoteRanges = New Collection
    Set dedupe = CreateObject("Scripting.Dictionary")
    dedupe.CompareMode = 1

    If Not resultFieldRanges Is Nothing And Not cfgNotes Is Nothing Then
        For Each target In resultFieldRanges
            If target Is Nothing Then GoTo ContinueTarget
            mapKey = Trim$(CStr(target("MapKey")))
            If Len(mapKey) = 0 Then GoTo ContinueTarget
            If Not cfgNotes.Exists(mapKey) Then GoTo ContinueTarget
            noteText = Trim$(CStr(cfgNotes(mapKey)))
            If Len(noteText) = 0 Then GoTo ContinueTarget

            rowStart = CLng(target("RowStart"))
            rowEnd = CLng(target("RowEnd"))
            If rowStart <= 0 Then GoTo ContinueTarget
            If rowEnd < rowStart Then rowEnd = rowStart

            dedupeKey = CStr(rowStart) & ":" & CStr(rowEnd)
            If dedupe.Exists(dedupeKey) Then GoTo ContinueTarget
            dedupe(dedupeKey) = True

            Set rowItem = CreateObject("Scripting.Dictionary")
            rowItem.CompareMode = 1
            rowItem("RowStart") = rowStart
            rowItem("RowEnd") = rowEnd
            configNoteRanges.Add rowItem
ContinueTarget:
        Next target
    End If

    Set result("confignote") = configNoteRanges
    Set mp_BuildConfigNoteRowKindRanges = result
End Function

Private Function mp_BuildPartialMatchRowKindRanges(ByVal partialMatchRowRanges As Collection) As Object
    Dim result As Object
    Dim partialRanges As Collection
    Dim rowItem As Variant

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1
    Set partialRanges = New Collection

    If Not partialMatchRowRanges Is Nothing Then
        For Each rowItem In partialMatchRowRanges
            partialRanges.Add rowItem
        Next rowItem
    End If

    Set result("partialmatch") = partialRanges
    Set mp_BuildPartialMatchRowKindRanges = result
End Function

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
