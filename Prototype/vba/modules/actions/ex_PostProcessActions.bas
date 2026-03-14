Attribute VB_Name = "ex_PostProcessActions"
Option Explicit

Private Const POST_PROCESS_STYLE_STAGE_NAME As String = "postProcess"
Private Const POST_PROCESS_HEADER_LAYER_ID As String = "pc-postprocess-header"
Private Const POST_PROCESS_FOOTER_LAYER_ID As String = "pc-postprocess-footer"
Private Const POST_PROCESS_MEASURE_SIDE_MARGIN As Double = 3
Private Const POST_PROCESS_MEASURE_VERTICAL_MARGIN As Double = 1
Private Const POST_PROCESS_MEASURE_EXTRA_HEIGHT_BASE As Double = 14
Private Const POST_PROCESS_MEASURE_EXTRA_HEIGHT_MIN As Double = 8
Private Const POST_PROCESS_MEASURE_EXTRA_HEIGHT_FONT_FACTOR As Double = 1
Private Const POST_PROCESS_MEASURE_HEIGHT_ROUND_PAD As Double = 1
Private Const POST_PROCESS_HEADER_SCROLL_CONTEXT_ROWS As Long = 4
Private Const POST_PROCESS_FOOTER_SCROLL_CONTEXT_ROWS As Long = 4
Private Const POST_PROCESS_HEADER_ANCHOR_MAX_OFFSET_ROWS As Long = 5
Private Const POST_PROCESS_HEADER_ROW_KIND As String = "postprocessheader"
Private Const POST_PROCESS_FOOTER_ROW_KIND As String = "postprocessfooter"
Private Const POST_PROCESS_HEADER_ANCHOR_NAME As String = "__pcPostProcessSingleHeader"
Private Const POST_PROCESS_FOOTER_ANCHOR_NAME As String = "__pcPostProcessSingleFooter"
Private Const RUNTIME_POINTER_PREFIX_NAME As String = "Name:"
Private Const BANNER_TYPE_WARNING As String = "TYPE_WARNING"
Private Const BANNER_TYPE_ERROR As String = "TYPE_ERROR"
Private Const BANNER_TYPE_NOTE As String = "TYPE_NOTE"
Private Const BANNER_TITLE_WARNING As String = "WARNING"
Private Const BANNER_TITLE_ERROR As String = "ERROR"
Private Const BANNER_TITLE_NOTE As String = "NOTE"
Private Const BANNER_KIND_WARNING As String = "warningbanner"
Private Const BANNER_KIND_ERROR As String = "errorbanner"
Private Const BANNER_KIND_NOTE As String = "notebanner"
Private Const DEFER_OP_APPEND_HEADER_TEXT As String = "append_single_header_text"
Private Const DEFER_OP_APPEND_FOOTER_TEXT As String = "append_single_footer_text"
Private Const DEFER_OP_APPEND_HEADER_ROW As String = "append_header_row"
Private Const DEFER_OP_APPEND_FOOTER_ROW As String = "append_footer_row"
Private Const DEFER_OP_SHOW_BANNER_AT_CELL As String = "show_banner_at_cell"
Private Const DEFER_OP_SHOW_BANNER_AFTER_BANNER As String = "show_banner_after_banner"
Private Const DEFER_OP_SHOW_BANNER_AT_TABLE As String = "show_banner_at_table"
Private Const DEFER_OP_SHOW_BANNER_BEFORE_ROW As String = "show_banner_before_row"
Private Const DEFER_OP_APPEND_RESULT_BLOCK As String = "append_result_block"
Private Const DEFER_ROW_ANCHOR_PREFIX As String = "__pcDeferredRow_"
Private Const RUNTIME_ROW_ANCHOR_PREFIX As String = "__pcRuntimeRow_"
Private Const RESULT_BLOCK_ANCHOR_PREFIX As String = "__pcResultBlock_"
Private Const RESULT_BLOCK_ANCHOR_MAX_INDEX As Long = 9999
Private Const RESULT_BLOCK_LAST_ANCHOR_NAME As String = "__pcResultBlockLast"

Private Type t_PostProcessHeaderStyle
    Columns As Long
    Overflow As String
    BackColor As Long
    FontColor As Long
    FontSize As Double
    RowHeight As Double
    MinRowHeight As Double
    AutoHeight As Boolean
    AutoHeightMarginTop As Double
    AutoHeightMarginBottom As Double
End Type

Private Type t_PostProcessFooterStyle
    Columns As Long
    Overflow As String
    BackColor As Long
    FontColor As Long
    FontSize As Double
    RowHeight As Double
    MinRowHeight As Double
    AutoHeight As Boolean
    AutoHeightMarginTop As Double
    AutoHeightMarginBottom As Double
End Type

Private g_PostProcessHeaderSheetKey As String
Private g_PostProcessHeaderNextInsertRow As Long
Private g_PostProcessHeaderRowIndex As Long
Private g_PostProcessHeaderHasAppended As Boolean
Private g_PostProcessFooterSheetKey As String
Private g_PostProcessFooterRowIndex As Long
Private g_PostProcessFooterHasAppended As Boolean
Private g_ResultBlockSheetKey As String
Private g_ResultBlockNextInsertRow As Long
Private g_ResultBlockLastRowIndex As Long
Private g_ResultBlockCount As Long
Private g_RuntimeDataBySheetAndKey As Object
Private g_DeferredSingleHeaderTextBySheet As Object
Private g_DeferredSingleFooterTextBySheet As Object
Private g_DeferredRowAnchorSeqBySheet As Object

Public Sub m_HighlightRow( _
    ByVal rowRef As obj_ResultRow, _
    Optional ByVal colorHex As String = "#FFF2CC" _
)
    Dim colorValue As Long
    Dim rowRange As Range
    Dim ws As Worksheet
    Dim usedCols As Long
    Dim rowIndex As Long

    If rowRef Is Nothing Then Exit Sub
    If Len(Trim$(colorHex)) = 0 Then colorHex = "#FFF2CC"
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    usedCols = mp_GetLastUsedColumn(ws)
    If usedCols <= 0 Then usedCols = 1

    If Not ex_XmlCore.m_TryParseColor(colorHex, colorValue) Then
        Err.Raise vbObjectError + 1650, "ex_PostProcessActions", "Invalid highlight color: " & colorHex
    End If

    rowIndex = mp_ResolveAnchoredRowIndex(ws, rowRef, "row highlight")
    Set rowRange = ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, usedCols))
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
    Dim rowIndex As Long

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

    rowIndex = mp_ResolveAnchoredRowIndex(ws, rowRef, "row cell highlight")
    Set targetCell = ws.Cells(rowIndex, targetCol)
    targetCell.Interior.Pattern = xlSolid
    targetCell.Interior.Color = colorValue
End Sub

Public Function m_RowToText( _
    ByVal rowRef As obj_ResultRow, _
    ByVal separatorText As String _
) As String
    m_RowToText = mp_GetRowText(rowRef, separatorText)
End Function

Public Function m_TextAppend( _
    ByVal baseText As String, _
    ByVal appendText As String, _
    Optional ByVal separatorText As String = vbNullString _
) As String
    If Len(appendText) = 0 Then
        m_TextAppend = CStr(baseText)
        Exit Function
    End If

    If Len(baseText) = 0 Then
        m_TextAppend = CStr(appendText)
    ElseIf Len(separatorText) = 0 Then
        m_TextAppend = CStr(baseText) & CStr(appendText)
    Else
        m_TextAppend = CStr(baseText) & CStr(separatorText) & CStr(appendText)
    End If
End Function

Public Function m_SetText(ByVal textValue As String) As String
    m_SetText = CStr(textValue)
End Function

Public Sub m_SetRowCellText( _
    ByVal rowRef As obj_ResultRow, _
    ByVal columnRef As String, _
    ByVal cellText As String _
)
    Dim targetCell As Range

    Set targetCell = mp_GetTargetCellForRowRef(rowRef, columnRef)
    targetCell.Value = CStr(cellText)
End Sub

Public Sub m_AppendRowCellText( _
    ByVal rowRef As obj_ResultRow, _
    ByVal columnRef As String, _
    ByVal appendText As String, _
    Optional ByVal separatorText As String = vbLf _
)
    Dim targetCell As Range
    Dim currentText As String

    If Len(appendText) = 0 Then Exit Sub

    Set targetCell = mp_GetTargetCellForRowRef(rowRef, columnRef)
    currentText = CStr(targetCell.Value)
    targetCell.Value = m_TextAppend(currentText, CStr(appendText), separatorText)
End Sub

Public Sub m_AppendToOwnerRowCell( _
    ByVal rowRef As obj_ResultRow, _
    ByVal ownerColumnRef As String, _
    ByVal targetColumnRef As String, _
    ByVal appendText As String, _
    Optional ByVal separatorText As String = vbLf _
)
    Dim ws As Worksheet
    Dim ownerCol As Long
    Dim targetCol As Long
    Dim ownerRowIndex As Long
    Dim probeRow As Long
    Dim targetCell As Range
    Dim currentText As String
    Dim sourceRowIndex As Long

    If rowRef Is Nothing Then
        Err.Raise vbObjectError + 1674, "ex_PostProcessActions", "Row reference is required for owner row append."
    End If

    If Len(appendText) = 0 Then Exit Sub
    Set ws = ActiveSheet
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1675, "ex_PostProcessActions", "Active sheet is not available for owner row append."
    End If

    If Not mp_TryResolveColumnIndexInRow(rowRef, ownerColumnRef, ownerCol) Then
        Err.Raise vbObjectError + 1676, "ex_PostProcessActions", "Unknown owner column reference '" & ownerColumnRef & "'."
    End If
    If Not mp_TryResolveColumnIndexInRow(rowRef, targetColumnRef, targetCol) Then
        Err.Raise vbObjectError + 1677, "ex_PostProcessActions", "Unknown target column reference '" & targetColumnRef & "'."
    End If

    sourceRowIndex = mp_ResolveAnchoredRowIndex(ws, rowRef, "owner row append")
    For probeRow = sourceRowIndex To 1 Step -1
        If Len(Trim$(CStr(ws.Cells(probeRow, ownerCol).Value))) > 0 Then
            ownerRowIndex = probeRow
            Exit For
        End If
    Next probeRow

    If ownerRowIndex <= 0 Then
        Err.Raise vbObjectError + 1678, "ex_PostProcessActions", "Unable to resolve owner row by column '" & ownerColumnRef & "' from row " & CStr(sourceRowIndex) & "."
    End If

    Set targetCell = ws.Cells(ownerRowIndex, targetCol)
    currentText = CStr(targetCell.Value)
    targetCell.Value = m_TextAppend(currentText, CStr(appendText), separatorText)
End Sub


Public Sub m_AddNote( _
    ByVal rowRef As obj_ResultRow, _
    ByVal noteText As String _
)
    Dim noteCell As Range
    Dim ws As Worksheet
    Dim rowIndex As Long

    If rowRef Is Nothing Then Exit Sub
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    rowIndex = mp_ResolveAnchoredRowIndex(ws, rowRef, "note add")
    Set noteCell = ws.Cells(rowIndex, 1)
    On Error Resume Next
    If Not noteCell.Comment Is Nothing Then noteCell.Comment.Delete
    On Error GoTo 0
    noteCell.AddComment noteText
End Sub

Public Sub m_ResetPostProcessHeaderCursor(Optional ByVal targetSheet As Worksheet)
    Dim ws As Worksheet

    g_PostProcessHeaderNextInsertRow = 0
    g_PostProcessHeaderRowIndex = 0
    g_PostProcessHeaderHasAppended = False
    g_ResultBlockSheetKey = vbNullString
    g_ResultBlockNextInsertRow = 0
    g_ResultBlockLastRowIndex = 0
    g_ResultBlockCount = 0
    If targetSheet Is Nothing Then
        g_PostProcessHeaderSheetKey = vbNullString
        m_ClearRuntimeData
        Set g_DeferredSingleHeaderTextBySheet = Nothing
        Set g_DeferredSingleFooterTextBySheet = Nothing
        Set g_DeferredRowAnchorSeqBySheet = Nothing
        ex_RenderQueue.m_ClearAll
        On Error Resume Next
        Set ws = ActiveSheet
        On Error GoTo 0
        If Not ws Is Nothing Then
            mp_ClearPreviousSinglePostProcessHeader ws
            mp_ClearPreviousResultBlocks ws
        End If
    Else
        g_PostProcessHeaderSheetKey = mp_BuildSheetKey(targetSheet)
        mp_ClearRuntimeDataForSheet targetSheet
        mp_ClearDeferredRenderSheetState targetSheet
        mp_SetDeferredRenderActive targetSheet, False
        mp_ClearPreviousSinglePostProcessHeader targetSheet
        mp_ClearPreviousResultBlocks targetSheet
    End If
End Sub

Public Sub m_ResetPostProcessFooterCursor(Optional ByVal targetSheet As Worksheet)
    Dim ws As Worksheet

    g_PostProcessFooterRowIndex = 0
    g_PostProcessFooterHasAppended = False
    g_ResultBlockSheetKey = vbNullString
    g_ResultBlockNextInsertRow = 0
    g_ResultBlockLastRowIndex = 0
    g_ResultBlockCount = 0
    If targetSheet Is Nothing Then
        g_PostProcessFooterSheetKey = vbNullString
        m_ClearRuntimeData
        Set g_DeferredSingleHeaderTextBySheet = Nothing
        Set g_DeferredSingleFooterTextBySheet = Nothing
        Set g_DeferredRowAnchorSeqBySheet = Nothing
        ex_RenderQueue.m_ClearAll
        On Error Resume Next
        Set ws = ActiveSheet
        On Error GoTo 0
        If Not ws Is Nothing Then
            mp_ClearPreviousSinglePostProcessFooter ws
            mp_ClearPreviousResultBlocks ws
        End If
    Else
        g_PostProcessFooterSheetKey = mp_BuildSheetKey(targetSheet)
        mp_ClearRuntimeDataForSheet targetSheet
        mp_ClearDeferredRenderSheetState targetSheet
        mp_SetDeferredRenderActive targetSheet, False
        mp_ClearPreviousSinglePostProcessFooter targetSheet
        mp_ClearPreviousResultBlocks targetSheet
    End If
End Sub

Public Sub m_BeginDeferredRender(Optional ByVal targetSheet As Worksheet = Nothing)
    Dim ws As Worksheet

    If targetSheet Is Nothing Then
        Set ws = ActiveSheet
    Else
        Set ws = targetSheet
    End If
    If ws Is Nothing Then Exit Sub

    mp_EnsureDeferredStores
    mp_ClearDeferredRenderSheetState ws
    ex_RenderQueue.m_BeginForSheet ws
End Sub

Public Sub m_CommitDeferredRender(Optional ByVal targetSheet As Worksheet = Nothing)
    Dim ws As Worksheet
    Dim queue As Collection
    Dim op As Object
    Dim phase As Long
    Dim prevSheet As Worksheet
    Dim prevSheetName As String
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    If targetSheet Is Nothing Then
        Set ws = ActiveSheet
    Else
        Set ws = targetSheet
    End If
    If ws Is Nothing Then Exit Sub

    If Not mp_IsDeferredRenderActiveForSheet(ws) Then Exit Sub
    Set queue = mp_GetDeferredQueueForSheet(ws)
    If queue Is Nothing Then
        mp_ClearDeferredRenderSheetState ws
        mp_SetDeferredRenderActive ws, False
        Exit Sub
    End If

    On Error Resume Next
    Set prevSheet = ActiveSheet
    If Not prevSheet Is Nothing Then prevSheetName = prevSheet.Name
    On Error GoTo 0

    On Error GoTo EH
    mp_SetDeferredRenderActive ws, False

    On Error Resume Next
    ws.Activate
    On Error GoTo EH

    For phase = 1 To 3
        For Each op In queue
            If mp_GetDeferredOperationPhase(op) = phase Then
                mp_ApplyDeferredOperation ws, op
            End If
        Next op
    Next phase

    mp_ClearDeferredRenderSheetState ws

    If Len(prevSheetName) > 0 Then
        On Error Resume Next
        ThisWorkbook.Worksheets(prevSheetName).Activate
        On Error GoTo 0
    End If
    Exit Sub

EH:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description

    On Error Resume Next
    mp_SetDeferredRenderActive ws, False
    mp_ClearDeferredRenderSheetState ws
    If Len(prevSheetName) > 0 Then ThisWorkbook.Worksheets(prevSheetName).Activate
    On Error GoTo 0

    Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub m_EndDeferredRender(Optional ByVal targetSheet As Worksheet = Nothing)
    Dim ws As Worksheet

    If targetSheet Is Nothing Then
        Set ws = ActiveSheet
    Else
        Set ws = targetSheet
    End If
    If ws Is Nothing Then Exit Sub

    mp_ClearDeferredRenderSheetState ws
    mp_SetDeferredRenderActive ws, False
End Sub

Public Sub m_AppendPostProcessHeaderText(ByVal postProcessHeaderText As String)
    Dim ws As Worksheet
    Dim insertRow As Long
    Dim endCol As Long
    Dim postProcessHeaderStyle As t_PostProcessHeaderStyle
    Dim postProcessHeaderRange As Range
    Dim sheetKey As String

    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    If mp_IsDeferredRenderActiveForSheet(ws) Then
        mp_QueueDeferredOperation ws, DEFER_OP_APPEND_HEADER_ROW, Array(CStr(postProcessHeaderText))
        Exit Sub
    End If

    If Not mp_TryLoadPostProcessHeaderStyle(postProcessHeaderStyle) Then
        Err.Raise vbObjectError + 1673, "ex_PostProcessActions", "Unable to apply postProcessHeader text: invalid postProcess header style."
    End If

    sheetKey = mp_BuildSheetKey(ws)
    If StrComp(g_PostProcessHeaderSheetKey, sheetKey, vbTextCompare) <> 0 Then
        g_PostProcessHeaderSheetKey = sheetKey
        g_PostProcessHeaderNextInsertRow = 0
    End If

    If g_PostProcessHeaderNextInsertRow <= 0 Then
        insertRow = mp_GetPostProcessHeaderInsertStartRow(ws)
    Else
        insertRow = g_PostProcessHeaderNextInsertRow
    End If
    If insertRow > ws.Rows.Count Then insertRow = ws.Rows.Count

    ws.Rows(insertRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    endCol = postProcessHeaderStyle.Columns
    If endCol < 1 Then endCol = 1
    If endCol > ws.Columns.Count Then endCol = ws.Columns.Count

    Set postProcessHeaderRange = ws.Range(ws.Cells(insertRow, 1), ws.Cells(insertRow, endCol))
    If postProcessHeaderRange.MergeCells Then postProcessHeaderRange.UnMerge

    postProcessHeaderRange.Cells(1, 1).Value = postProcessHeaderText
    mp_ApplyPostProcessHeaderKindStyle ws, insertRow

    mp_ApplyPostProcessHeaderRowHeight ws, postProcessHeaderRange, postProcessHeaderText, postProcessHeaderStyle
    g_PostProcessHeaderRowIndex = insertRow
    g_PostProcessHeaderHasAppended = True
    mp_SetPostProcessHeaderAnchors ws, insertRow
    g_PostProcessHeaderNextInsertRow = insertRow + 1
End Sub

Public Sub m_AppendToSinglePostProcessHeaderText( _
    ByVal appendText As String, _
    Optional ByVal separatorText As String = vbLf _
)
    Dim ws As Worksheet
    Dim sheetKey As String
    Dim postProcessHeaderStyle As t_PostProcessHeaderStyle
    Dim postProcessHeaderRange As Range
    Dim currentText As String
    Dim mergedText As String

    If Len(appendText) = 0 Then Exit Sub
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    If mp_IsDeferredRenderActiveForSheet(ws) Then
        mp_QueueDeferredOperation ws, DEFER_OP_APPEND_HEADER_TEXT, Array(CStr(appendText), CStr(separatorText))
        mp_AppendDeferredSingleHeaderText ws, appendText, separatorText
        Exit Sub
    End If

    sheetKey = mp_BuildSheetKey(ws)
    If StrComp(g_PostProcessHeaderSheetKey, sheetKey, vbTextCompare) <> 0 Then
        g_PostProcessHeaderSheetKey = sheetKey
        g_PostProcessHeaderNextInsertRow = 0
        g_PostProcessHeaderRowIndex = 0
        g_PostProcessHeaderHasAppended = False
    End If

    If Not mp_TryLoadPostProcessHeaderStyle(postProcessHeaderStyle) Then
        Err.Raise vbObjectError + 1728, "ex_PostProcessActions", "Unable to apply single postProcessHeader text: invalid postProcess header style."
    End If

    Set postProcessHeaderRange = mp_GetOrCreateSinglePostProcessHeaderRange(ws, postProcessHeaderStyle)
    If postProcessHeaderRange Is Nothing Then
        Err.Raise vbObjectError + 1741, "ex_PostProcessActions", "Unable to resolve single postProcessHeader range on sheet '" & ws.Name & "'."
    End If

    If g_PostProcessHeaderHasAppended Then
        currentText = CStr(postProcessHeaderRange.Cells(1, 1).Value)
    Else
        currentText = vbNullString
    End If
    mergedText = m_TextAppend(currentText, appendText, separatorText)
    postProcessHeaderRange.Cells(1, 1).Value = mergedText
    mp_ApplyPostProcessHeaderKindStyle ws, postProcessHeaderRange.Row

    mp_ApplyPostProcessHeaderRowHeight ws, postProcessHeaderRange, mergedText, postProcessHeaderStyle
    g_PostProcessHeaderHasAppended = True
End Sub

Public Function m_AppendResultBlock( _
    ByVal blockText As String, _
    Optional ByVal titleText As String = vbNullString, _
    Optional ByVal blockType As String = BANNER_TYPE_NOTE, _
    Optional ByVal gapRowsBefore As Long = 0, _
    Optional ByVal gapRowsAfter As Long = 1 _
) As String
    Dim ws As Worksheet
    Dim sheetKey As String
    Dim insertRow As Long
    Dim rowsToInsert As Long
    Dim blockRow As Long
    Dim bannerCols As Long
    Dim bannerRows As Long
    Dim bannerKind As String
    Dim bannerRangeAddress As String
    Dim bodyLines As Collection
    Dim insertedRange As Range
    Dim anchorName As String
    Dim nextBlockIndex As Long

    blockText = Trim$(blockText)
    If Len(blockText) = 0 Then Exit Function

    Set ws = ActiveSheet
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1782, "ex_PostProcessActions", "Active sheet is not available for result block render."
    End If

    If mp_IsDeferredRenderActiveForSheet(ws) Then
        mp_QueueDeferredOperation ws, DEFER_OP_APPEND_RESULT_BLOCK, Array(CStr(blockText), CStr(titleText), CStr(blockType), CLng(gapRowsBefore), CLng(gapRowsAfter))
        m_AppendResultBlock = blockText
        Exit Function
    End If

    mp_ValidateBannerGapRows gapRowsBefore, gapRowsAfter, "result block"
    bannerKind = mp_MapBannerTypeToKind(blockType)
    titleText = Trim$(titleText)

    sheetKey = mp_BuildSheetKey(ws)
    If StrComp(g_ResultBlockSheetKey, sheetKey, vbTextCompare) <> 0 Then
        g_ResultBlockSheetKey = sheetKey
        g_ResultBlockNextInsertRow = 0
        g_ResultBlockLastRowIndex = 0
        g_ResultBlockCount = 0
    End If

    If g_ResultBlockNextInsertRow <= 0 Then
        insertRow = mp_GetPostProcessHeaderInsertStartRow(ws)
    Else
        insertRow = g_ResultBlockNextInsertRow
    End If
    If insertRow < 1 Then insertRow = 1
    If insertRow > ws.Rows.Count Then insertRow = ws.Rows.Count

    mp_GetWarningBannerDimensions bannerCols, bannerRows, bannerKind
    rowsToInsert = gapRowsBefore + bannerRows + gapRowsAfter
    If rowsToInsert <= 0 Then rowsToInsert = 1
    If insertRow + rowsToInsert - 1 > ws.Rows.Count Then
        rowsToInsert = ws.Rows.Count - insertRow + 1
    End If

    ws.Rows(CStr(insertRow) & ":" & CStr(insertRow + rowsToInsert - 1)).Insert Shift:=xlDown
    mp_UnmergeRowsSafe ws, insertRow, gapRowsBefore

    blockRow = insertRow + gapRowsBefore
    If blockRow > ws.Rows.Count Then blockRow = ws.Rows.Count

    bannerRangeAddress = "A" & CStr(blockRow) & ":" & mp_ToColumnLetter(bannerCols) & CStr(blockRow + bannerRows - 1)

    Set bodyLines = New Collection
    bodyLines.Add blockText
    nextBlockIndex = g_ResultBlockCount + 1
    ex_Messaging.m_RenderBanner ws, titleText, bodyLines, bannerRangeAddress, bannerKind, mp_BuildResultBlockIdentity(sheetKey, nextBlockIndex, blockText)

    mp_UnmergeRowsSafe ws, blockRow + bannerRows, gapRowsAfter
    Set insertedRange = ws.Range(ws.Cells(insertRow, 1), ws.Cells(insertRow + rowsToInsert - 1, bannerCols))
    anchorName = mp_NextResultBlockAnchorName(ws)
    If Len(anchorName) > 0 Then
        mp_SetNamedRangeAnchor ws, anchorName, insertedRange
    End If

    g_ResultBlockCount = nextBlockIndex
    g_ResultBlockLastRowIndex = blockRow
    g_ResultBlockNextInsertRow = insertRow + rowsToInsert
    If g_ResultBlockNextInsertRow > ws.Rows.Count Then g_ResultBlockNextInsertRow = ws.Rows.Count
    mp_SetNamedRowAnchor ws, RESULT_BLOCK_LAST_ANCHOR_NAME, blockRow
    m_AppendResultBlock = blockText
End Function

Public Function m_GetLastResultBlockCellRef(Optional ByVal targetSheet As Worksheet = Nothing) As String
    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim sheetKey As String

    If targetSheet Is Nothing Then
        Set ws = ActiveSheet
    Else
        Set ws = targetSheet
    End If
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1783, "ex_PostProcessActions", "Active sheet is not available for result block pointer read."
    End If
    If mp_IsDeferredRenderActiveForSheet(ws) Then
        m_GetLastResultBlockCellRef = RUNTIME_POINTER_PREFIX_NAME & RESULT_BLOCK_LAST_ANCHOR_NAME
        Exit Function
    End If

    sheetKey = mp_BuildSheetKey(ws)
    If StrComp(g_ResultBlockSheetKey, sheetKey, vbTextCompare) <> 0 Then
        If Not mp_TryGetNamedRowAnchor(ws, RESULT_BLOCK_LAST_ANCHOR_NAME, rowIndex) Then
            Err.Raise vbObjectError + 1784, "ex_PostProcessActions", "Result block pointer is not available on sheet '" & ws.Name & "'."
        End If
        m_GetLastResultBlockCellRef = "Cell:A" & CStr(rowIndex)
        Exit Function
    End If
    rowIndex = g_ResultBlockLastRowIndex
    If rowIndex < 1 Or rowIndex > ws.Rows.Count Then
        If Not mp_TryGetNamedRowAnchor(ws, RESULT_BLOCK_LAST_ANCHOR_NAME, rowIndex) Then
            Err.Raise vbObjectError + 1784, "ex_PostProcessActions", "Result block pointer is not available on sheet '" & ws.Name & "'."
        End If
    End If

    m_GetLastResultBlockCellRef = "Cell:A" & CStr(rowIndex)
End Function

Public Function m_GetSinglePostProcessHeaderText(Optional ByVal targetSheet As Worksheet = Nothing) As String
    Dim ws As Worksheet
    Dim headerRowIndex As Long
    Dim deferredText As String

    If targetSheet Is Nothing Then
        Set ws = ActiveSheet
    Else
        Set ws = targetSheet
    End If
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1729, "ex_PostProcessActions", "Active sheet is not available for postProcessHeader read."
    End If
    If mp_IsDeferredRenderActiveForSheet(ws) Then
        If mp_TryGetDeferredSingleHeaderText(ws, deferredText) Then
            m_GetSinglePostProcessHeaderText = deferredText
            Exit Function
        End If
    End If

    If Not mp_TryGetCachedSinglePostProcessHeaderRowIndex(ws, headerRowIndex) Then
        Err.Raise vbObjectError + 1731, "ex_PostProcessActions", "Single postProcessHeader row is not available on sheet '" & ws.Name & "'."
    End If
    If headerRowIndex <= 0 Then
        Err.Raise vbObjectError + 1731, "ex_PostProcessActions", "Single postProcessHeader row is not available on sheet '" & ws.Name & "'."
    End If

    m_GetSinglePostProcessHeaderText = CStr(ws.Cells(headerRowIndex, 1).Value)
End Function

Public Sub m_AppendPostProcessFooterText(ByVal postProcessFooterText As String)
    Dim ws As Worksheet
    Dim startRow As Long
    Dim endCol As Long
    Dim postProcessFooterStyle As t_PostProcessFooterStyle
    Dim postProcessFooterRange As Range
    Dim sheetKey As String

    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    If mp_IsDeferredRenderActiveForSheet(ws) Then
        mp_QueueDeferredOperation ws, DEFER_OP_APPEND_FOOTER_ROW, Array(CStr(postProcessFooterText))
        Exit Sub
    End If

    If Not mp_TryLoadPostProcessFooterStyle(postProcessFooterStyle) Then
        Err.Raise vbObjectError + 1651, "ex_PostProcessActions", "Unable to apply postProcessFooter text: invalid postProcess footer style."
    End If

    startRow = mp_GetLastUsedRow(ws) + 2
    If startRow < 1 Then startRow = 1

    endCol = postProcessFooterStyle.Columns
    If endCol < 1 Then endCol = 1
    If endCol > ws.Columns.Count Then endCol = ws.Columns.Count

    Set postProcessFooterRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, endCol))
    If postProcessFooterRange.MergeCells Then postProcessFooterRange.UnMerge

    postProcessFooterRange.Cells(1, 1).Value = postProcessFooterText
    mp_ApplyPostProcessFooterKindStyle ws, startRow

    mp_ApplyPostProcessFooterRowHeight ws, postProcessFooterRange, postProcessFooterText, postProcessFooterStyle

    sheetKey = mp_BuildSheetKey(ws)
    g_PostProcessFooterSheetKey = sheetKey
    g_PostProcessFooterRowIndex = startRow
    g_PostProcessFooterHasAppended = True
End Sub

Public Sub m_AppendToSinglePostProcessFooterText( _
    ByVal appendText As String, _
    Optional ByVal separatorText As String = vbLf _
)
    Dim ws As Worksheet
    Dim postProcessFooterStyle As t_PostProcessFooterStyle
    Dim postProcessFooterRange As Range
    Dim currentText As String
    Dim mergedText As String

    If Len(appendText) = 0 Then Exit Sub
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    If mp_IsDeferredRenderActiveForSheet(ws) Then
        mp_QueueDeferredOperation ws, DEFER_OP_APPEND_FOOTER_TEXT, Array(CStr(appendText), CStr(separatorText))
        mp_AppendDeferredSingleFooterText ws, appendText, separatorText
        Exit Sub
    End If

    If Not mp_TryLoadPostProcessFooterStyle(postProcessFooterStyle) Then
        Err.Raise vbObjectError + 1682, "ex_PostProcessActions", "Unable to apply single postProcessFooter text: invalid postProcess footer style."
    End If

    Set postProcessFooterRange = mp_GetOrCreateSinglePostProcessFooterRange(ws, postProcessFooterStyle)
    If g_PostProcessFooterHasAppended Then
        currentText = CStr(postProcessFooterRange.Cells(1, 1).Value)
    Else
        currentText = vbNullString
    End If
    mergedText = m_TextAppend(currentText, appendText, separatorText)
    postProcessFooterRange.Cells(1, 1).Value = mergedText

    mp_ApplyPostProcessFooterRowHeight ws, postProcessFooterRange, mergedText, postProcessFooterStyle
    g_PostProcessFooterHasAppended = True
End Sub

Public Function m_GetSinglePostProcessFooterText(Optional ByVal targetSheet As Worksheet = Nothing) As String
    Dim ws As Worksheet
    Dim footerRowIndex As Long
    Dim deferredText As String

    If targetSheet Is Nothing Then
        Set ws = ActiveSheet
    Else
        Set ws = targetSheet
    End If
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1701, "ex_PostProcessActions", "Active sheet is not available for postProcessFooter read."
    End If
    If mp_IsDeferredRenderActiveForSheet(ws) Then
        If mp_TryGetDeferredSingleFooterText(ws, deferredText) Then
            m_GetSinglePostProcessFooterText = deferredText
            Exit Function
        End If
    End If

    If Not mp_TryGetCachedSinglePostProcessFooterRowIndex(ws, footerRowIndex) Then
        Err.Raise vbObjectError + 1702, "ex_PostProcessActions", "Single postProcessFooter row is not available on sheet '" & ws.Name & "'."
    End If
    If footerRowIndex <= 0 Then
        Err.Raise vbObjectError + 1703, "ex_PostProcessActions", "Single postProcessFooter row not found on sheet '" & ws.Name & "'."
    End If

    m_GetSinglePostProcessFooterText = CStr(ws.Cells(footerRowIndex, 1).Value)
End Function

Public Sub m_ClearRuntimeData(Optional ByVal targetSheet As Worksheet = Nothing)
    If targetSheet Is Nothing Then
        Set g_RuntimeDataBySheetAndKey = Nothing
    Else
        mp_ClearRuntimeDataForSheet targetSheet
    End If
End Sub

Public Sub m_SetRuntimeData( _
    ByVal dataKey As String, _
    ByVal dataValue As String, _
    Optional ByVal targetSheet As Worksheet = Nothing _
)
    Dim ws As Worksheet
    Dim cache As Object
    Dim runtimeKey As String

    If targetSheet Is Nothing Then
        Set ws = ActiveSheet
    Else
        Set ws = targetSheet
    End If
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1704, "ex_PostProcessActions", "Active sheet is not available for runtime data write."
    End If

    dataKey = mp_NormalizeRuntimeDataKey(dataKey)
    If Len(dataKey) = 0 Then
        Err.Raise vbObjectError + 1705, "ex_PostProcessActions", "Runtime data key cannot be empty."
    End If

    Set cache = mp_EnsureRuntimeDataCache()
    runtimeKey = mp_BuildRuntimeDataFullKey(ws, dataKey)
    cache(runtimeKey) = CStr(dataValue)
End Sub

Public Function m_GetRuntimeData( _
    ByVal dataKey As String, _
    Optional ByVal defaultValue As String = vbNullString, _
    Optional ByVal targetSheet As Worksheet = Nothing _
) As String
    Dim ws As Worksheet
    Dim cache As Object
    Dim runtimeKey As String

    If targetSheet Is Nothing Then
        Set ws = ActiveSheet
    Else
        Set ws = targetSheet
    End If
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1706, "ex_PostProcessActions", "Active sheet is not available for runtime data read."
    End If

    dataKey = mp_NormalizeRuntimeDataKey(dataKey)
    If Len(dataKey) = 0 Then
        Err.Raise vbObjectError + 1707, "ex_PostProcessActions", "Runtime data key cannot be empty."
    End If

    Set cache = g_RuntimeDataBySheetAndKey
    If cache Is Nothing Then
        m_GetRuntimeData = defaultValue
        Exit Function
    End If

    runtimeKey = mp_BuildRuntimeDataFullKey(ws, dataKey)
    If cache.Exists(runtimeKey) Then
        m_GetRuntimeData = CStr(cache(runtimeKey))
    Else
        m_GetRuntimeData = defaultValue
    End If
End Function

Public Function m_GetRuntimeDataEntriesByPrefix( _
    ByVal dataKeyPrefix As String, _
    Optional ByVal targetSheet As Worksheet = Nothing _
) As Object
    Dim ws As Worksheet
    Dim cache As Object
    Dim result As Object
    Dim sheetKey As String
    Dim keyPrefixNormalized As String
    Dim runtimePrefix As String
    Dim key As Variant
    Dim fullKey As String
    Dim dataKey As String

    If targetSheet Is Nothing Then
        Set ws = ActiveSheet
    Else
        Set ws = targetSheet
    End If
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1785, "ex_PostProcessActions", "Active sheet is not available for runtime data prefix read."
    End If

    keyPrefixNormalized = mp_NormalizeRuntimeDataKey(dataKeyPrefix)
    If Len(keyPrefixNormalized) = 0 Then
        Err.Raise vbObjectError + 1786, "ex_PostProcessActions", "Runtime data key prefix cannot be empty."
    End If

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1 ' vbTextCompare

    Set cache = g_RuntimeDataBySheetAndKey
    If cache Is Nothing Then
        Set m_GetRuntimeDataEntriesByPrefix = result
        Exit Function
    End If
    If cache.Count = 0 Then
        Set m_GetRuntimeDataEntriesByPrefix = result
        Exit Function
    End If

    sheetKey = mp_BuildSheetKey(ws)
    runtimePrefix = sheetKey & "|" & keyPrefixNormalized

    For Each key In cache.Keys
        fullKey = CStr(key)
        If StrComp(Left$(fullKey, Len(runtimePrefix)), runtimePrefix, vbTextCompare) = 0 Then
            dataKey = Mid$(fullKey, Len(sheetKey) + 2) ' skip "<sheetKey>|"
            result(dataKey) = CStr(cache(fullKey))
        End If
    Next key

    Set m_GetRuntimeDataEntriesByPrefix = result
End Function

Public Function m_GetSinglePostProcessFooterCellRef(Optional ByVal targetSheet As Worksheet = Nothing) As String
    Dim ws As Worksheet
    Dim footerRowIndex As Long

    If targetSheet Is Nothing Then
        Set ws = ActiveSheet
    Else
        Set ws = targetSheet
    End If
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1708, "ex_PostProcessActions", "Active sheet is not available for footer cell reference read."
    End If
    If mp_IsDeferredRenderActiveForSheet(ws) Then
        If mp_HasDeferredSingleFooterText(ws) Then
            m_GetSinglePostProcessFooterCellRef = RUNTIME_POINTER_PREFIX_NAME & POST_PROCESS_FOOTER_ANCHOR_NAME
            Exit Function
        End If
    End If

    If Not mp_TryGetCachedSinglePostProcessFooterRowIndex(ws, footerRowIndex) Then
        Err.Raise vbObjectError + 1709, "ex_PostProcessActions", "Single postProcessFooter row is not available on sheet '" & ws.Name & "'."
    End If
    If footerRowIndex <= 0 Then
        Err.Raise vbObjectError + 1710, "ex_PostProcessActions", "Single postProcessFooter row not found on sheet '" & ws.Name & "'."
    End If

    m_GetSinglePostProcessFooterCellRef = "Cell:A" & CStr(footerRowIndex)
End Function

Public Function m_GetSinglePostProcessHeaderCellRef(Optional ByVal targetSheet As Worksheet = Nothing) As String
    Dim ws As Worksheet
    Dim headerRowIndex As Long

    If targetSheet Is Nothing Then
        Set ws = ActiveSheet
    Else
        Set ws = targetSheet
    End If
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1732, "ex_PostProcessActions", "Active sheet is not available for header cell reference read."
    End If
    If mp_IsDeferredRenderActiveForSheet(ws) Then
        If mp_HasDeferredSingleHeaderText(ws) Then
            m_GetSinglePostProcessHeaderCellRef = RUNTIME_POINTER_PREFIX_NAME & POST_PROCESS_HEADER_ANCHOR_NAME
            Exit Function
        End If
    End If

    If Not mp_TryGetCachedSinglePostProcessHeaderRowIndex(ws, headerRowIndex) Then
        Err.Raise vbObjectError + 1734, "ex_PostProcessActions", "Single postProcessHeader row is not available on sheet '" & ws.Name & "'."
    End If
    If headerRowIndex <= 0 Then
        Err.Raise vbObjectError + 1734, "ex_PostProcessActions", "Single postProcessHeader row is not available on sheet '" & ws.Name & "'."
    End If

    m_GetSinglePostProcessHeaderCellRef = "Cell:A" & CStr(headerRowIndex)
End Function

Public Sub m_ScrollToPostProcessResults(Optional ByVal targetSheet As Worksheet = Nothing)
    Dim ws As Worksheet

    If targetSheet Is Nothing Then
        Set ws = ActiveSheet
    Else
        Set ws = targetSheet
    End If
    If ws Is Nothing Then Exit Sub

    If Not mp_HasPostProcessFooterForSheet(ws) Then Exit Sub
    Call mp_TryScrollToSinglePostProcessFooter(ws)
End Sub

Public Sub m_ScrollToSinglePostProcessHeader(Optional ByVal targetSheet As Worksheet = Nothing)
    Call mp_TryScrollToSinglePostProcessHeader(targetSheet)
End Sub

Public Sub m_ScrollToSinglePostProcessFooter(Optional ByVal targetSheet As Worksheet = Nothing)
    Call mp_TryScrollToSinglePostProcessFooter(targetSheet)
End Sub

Private Function mp_TryScrollToSinglePostProcessHeader(Optional ByVal targetSheet As Worksheet = Nothing) As Boolean
    Dim ws As Worksheet
    Dim headerRowIndex As Long
    Dim targetScrollRow As Long

    On Error GoTo SafeExit

    If targetSheet Is Nothing Then
        Set ws = ActiveSheet
    Else
        Set ws = targetSheet
    End If
    If ws Is Nothing Then Exit Function

    If Not mp_TryGetCachedSinglePostProcessHeaderRowIndex(ws, headerRowIndex) Then Exit Function

    targetScrollRow = headerRowIndex - POST_PROCESS_HEADER_SCROLL_CONTEXT_ROWS
    If targetScrollRow < 1 Then targetScrollRow = 1

    ws.Activate
    Application.Goto ws.Cells(targetScrollRow, 1), True
    mp_TryScrollToSinglePostProcessHeader = True

SafeExit:
End Function

Private Function mp_TryScrollToSinglePostProcessFooter(Optional ByVal targetSheet As Worksheet = Nothing) As Boolean
    Dim ws As Worksheet
    Dim footerRowIndex As Long
    Dim targetScrollRow As Long

    On Error GoTo SafeExit

    If targetSheet Is Nothing Then
        Set ws = ActiveSheet
    Else
        Set ws = targetSheet
    End If
    If ws Is Nothing Then Exit Function

    If Not mp_TryGetCachedSinglePostProcessFooterRowIndex(ws, footerRowIndex) Then Exit Function

    targetScrollRow = footerRowIndex - POST_PROCESS_FOOTER_SCROLL_CONTEXT_ROWS
    If targetScrollRow < 1 Then targetScrollRow = 1

    ws.Activate
    Application.Goto ws.Cells(targetScrollRow, 1), True
    mp_TryScrollToSinglePostProcessFooter = True

SafeExit:
End Function

Public Sub m_RaiseError(ByVal errorText As String)
    errorText = Trim$(errorText)
    If Len(errorText) = 0 Then errorText = "PostProcess script error."
    Err.Raise vbObjectError + 1712, "ex_PostProcessActions", errorText
End Sub

Public Function m_ShowLogicError(ByVal errorText As String) As String
    errorText = Trim$(errorText)
    If Len(errorText) = 0 Then errorText = "PostProcess logic error."
    MsgBox errorText, vbExclamation
    m_ShowLogicError = errorText
End Function

Public Function m_ShowBannerAtCell( _
    ByVal bannerType As String, _
    Optional ByVal titleText As String = vbNullString, _
    Optional ByVal bannerText As String = vbNullString, _
    Optional ByVal topLeftCellRef As String = "A1", _
    Optional ByVal gapRowsBefore As Long = 0, _
    Optional ByVal gapRowsAfter As Long = 0 _
) As String
    Dim ws As Worksheet
    Dim bannerKind As String

    bannerText = Trim$(bannerText)
    If Len(bannerText) = 0 Then Exit Function
    bannerKind = mp_MapBannerTypeToKind(bannerType)
    titleText = mp_ResolveBannerTitle(titleText, bannerType)
    mp_ValidateBannerGapRows gapRowsBefore, gapRowsAfter, "banner at cell"

    Set ws = ActiveSheet
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1726, "ex_PostProcessActions", "Active sheet is not available for banner at cell."
    End If
    If mp_IsDeferredRenderActiveForSheet(ws) Then
        mp_QueueDeferredOperation ws, DEFER_OP_SHOW_BANNER_AT_CELL, Array(CStr(bannerType), CStr(titleText), CStr(bannerText), CStr(topLeftCellRef), CLng(gapRowsBefore), CLng(gapRowsAfter))
        m_ShowBannerAtCell = bannerText
        Exit Function
    End If

    ex_Messaging.m_RenderTextBannerAtCell ws, bannerText, topLeftCellRef, titleText, bannerKind, gapRowsBefore, gapRowsAfter
    m_ShowBannerAtCell = bannerText
End Function

Public Function m_ShowBannerAfterBanner( _
    ByVal bannerType As String, _
    Optional ByVal titleText As String = vbNullString, _
    Optional ByVal bannerText As String = vbNullString, _
    Optional ByVal afterBannerIndex As Long = 1, _
    Optional ByVal gapRowsBefore As Long = 0, _
    Optional ByVal gapRowsAfter As Long = 1 _
) As String
    Dim ws As Worksheet
    Dim bannerKind As String

    bannerText = Trim$(bannerText)
    If Len(bannerText) = 0 Then Exit Function
    bannerKind = mp_MapBannerTypeToKind(bannerType)
    titleText = mp_ResolveBannerTitle(titleText, bannerType)
    If afterBannerIndex < 0 Then
        Err.Raise vbObjectError + 1727, "ex_PostProcessActions", "Banner index cannot be negative for 'after banner' placement."
    End If

    Set ws = ActiveSheet
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1729, "ex_PostProcessActions", "Active sheet is not available for 'after banner' placement."
    End If
    mp_ValidateBannerGapRows gapRowsBefore, gapRowsAfter, "after banner"
    If mp_IsDeferredRenderActiveForSheet(ws) Then
        mp_QueueDeferredOperation ws, DEFER_OP_SHOW_BANNER_AFTER_BANNER, Array(CStr(bannerType), CStr(titleText), CStr(bannerText), CLng(afterBannerIndex), CLng(gapRowsBefore), CLng(gapRowsAfter))
        m_ShowBannerAfterBanner = bannerText
        Exit Function
    End If

    ex_Messaging.m_RenderTextBannerAfterBanner ws, bannerText, afterBannerIndex, titleText, bannerKind, gapRowsBefore, gapRowsAfter
    m_ShowBannerAfterBanner = bannerText
End Function

Public Function m_ShowBannerAtTable( _
    ByVal bannerType As String, _
    ByVal tableRef As String, _
    Optional ByVal titleText As String = vbNullString, _
    Optional ByVal bannerText As String = vbNullString, _
    Optional ByVal positionText As String = "before", _
    Optional ByVal gapRowsBefore As Long = 0, _
    Optional ByVal gapRowsAfter As Long = 1 _
) As String
    Dim ws As Worksheet
    Dim bannerKind As String

    bannerText = Trim$(bannerText)
    If Len(bannerText) = 0 Then Exit Function
    bannerKind = mp_MapBannerTypeToKind(bannerType)
    titleText = mp_ResolveBannerTitle(titleText, bannerType)
    mp_ValidateBannerGapRows gapRowsBefore, gapRowsAfter, "table banner"
    tableRef = Trim$(tableRef)
    If Len(tableRef) = 0 Then
        Err.Raise vbObjectError + 1730, "ex_PostProcessActions", "Table reference is required for table banner placement."
    End If

    Set ws = ActiveSheet
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1731, "ex_PostProcessActions", "Active sheet is not available for table banner placement."
    End If
    If mp_IsDeferredRenderActiveForSheet(ws) Then
        mp_QueueDeferredOperation ws, DEFER_OP_SHOW_BANNER_AT_TABLE, Array(CStr(bannerType), CStr(tableRef), CStr(titleText), CStr(bannerText), CStr(positionText), CLng(gapRowsBefore), CLng(gapRowsAfter))
        m_ShowBannerAtTable = bannerText
        Exit Function
    End If

    ex_Messaging.m_RenderTextBannerAtTable ws, bannerText, tableRef, positionText, titleText, bannerKind, gapRowsBefore, gapRowsAfter
    m_ShowBannerAtTable = bannerText
End Function

Public Function m_GetRowIndex(ByVal rowRef As Object) As Long
    Dim sourceRowRef As obj_ResultRow
    Dim ws As Worksheet

    If rowRef Is Nothing Then
        Err.Raise vbObjectError + 1718, "ex_PostProcessActions", "Row reference is required for row index read."
    End If
    If Not TypeOf rowRef Is obj_ResultRow Then
        Err.Raise vbObjectError + 1719, "ex_PostProcessActions", "Row reference must be obj_ResultRow for row index read."
    End If
    Set sourceRowRef = rowRef

    Set ws = ActiveSheet
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1744, "ex_PostProcessActions", "Active sheet is not available for row index read."
    End If

    m_GetRowIndex = mp_ResolveAnchoredRowIndex(ws, sourceRowRef, "row index read")
End Function

Private Function m_GetBannerRangeAboveRow( _
    ByVal bannerType As String, _
    ByVal rowRef As Object, _
    Optional ByVal gapRowsText As String = "1" _
) As String
    Dim sourceRowRef As obj_ResultRow
    Dim gapRows As Long
    Dim bannerCols As Long
    Dim bannerRows As Long
    Dim startRow As Long
    Dim bannerKind As String
    Dim sourceRowIndex As Long
    Dim ws As Worksheet

    If rowRef Is Nothing Then
        Err.Raise vbObjectError + 1714, "ex_PostProcessActions", "Row reference is required for banner range."
    End If
    If Not TypeOf rowRef Is obj_ResultRow Then
        Err.Raise vbObjectError + 1715, "ex_PostProcessActions", "Row reference must be obj_ResultRow for banner range."
    End If
    Set sourceRowRef = rowRef

    gapRowsText = Trim$(gapRowsText)
    If Len(gapRowsText) = 0 Then gapRowsText = "1"
    If Not ex_XmlCore.m_TryParseLong(gapRowsText, gapRows) Then
        Err.Raise vbObjectError + 1716, "ex_PostProcessActions", "Gap rows must be integer for banner range."
    End If
    If gapRows < 0 Then
        Err.Raise vbObjectError + 1717, "ex_PostProcessActions", "Gap rows cannot be negative for banner range."
    End If

    bannerKind = mp_MapBannerTypeToKind(bannerType)
    mp_GetWarningBannerDimensions bannerCols, bannerRows, bannerKind
    Set ws = ActiveSheet
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1745, "ex_PostProcessActions", "Active sheet is not available for banner range."
    End If
    sourceRowIndex = mp_ResolveAnchoredRowIndex(ws, sourceRowRef, "banner range build")
    startRow = sourceRowIndex - gapRows - bannerRows
    If startRow < 1 Then startRow = 1

    m_GetBannerRangeAboveRow = "A" & CStr(startRow) & ":" & mp_ToColumnLetter(bannerCols) & CStr(startRow + bannerRows - 1)
End Function

Private Sub mp_ValidateBannerGapRows( _
    ByVal gapRowsBefore As Long, _
    ByVal gapRowsAfter As Long, _
    ByVal contextName As String _
)
    If gapRowsBefore < 0 Then
        Err.Raise vbObjectError + 1771, "ex_PostProcessActions", "Gap rows before cannot be negative for " & contextName & "."
    End If
    If gapRowsAfter < 0 Then
        Err.Raise vbObjectError + 1772, "ex_PostProcessActions", "Gap rows after cannot be negative for " & contextName & "."
    End If
End Sub

Public Function m_ShowBannerBeforeRowIndex( _
    ByVal bannerType As String, _
    Optional ByVal titleText As String = vbNullString, _
    Optional ByVal bannerText As String = vbNullString, _
    Optional ByVal rowIndex As Long = 0, _
    Optional ByVal gapRowsBefore As Long = 0, _
    Optional ByVal gapRowsAfter As Long = 1 _
) As String
    Dim ws As Worksheet
    Dim bannerCols As Long
    Dim bannerRows As Long
    Dim rowsToInsert As Long
    Dim bannerRangeAddress As String
    Dim existingRangeAddress As String
    Dim existingRange As Range
    Dim bannerKind As String
    Dim bannerStartRow As Long
    Dim existingStartRow As Long
    Dim existingRowCount As Long
    Dim existingEndRow As Long

    bannerText = Trim$(bannerText)
    If Len(bannerText) = 0 Then Exit Function
    bannerKind = mp_MapBannerTypeToKind(bannerType)
    titleText = mp_ResolveBannerTitle(titleText, bannerType)

    If rowIndex < 1 Then
        Err.Raise vbObjectError + 1720, "ex_PostProcessActions", "Row index must be >= 1 for banner insert."
    End If

    mp_ValidateBannerGapRows gapRowsBefore, gapRowsAfter, "banner insert"

    Set ws = ActiveSheet
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1724, "ex_PostProcessActions", "Active sheet is not available for banner insert."
    End If
    If mp_IsDeferredRenderActiveForSheet(ws) Then
        mp_QueueDeferredOperation ws, DEFER_OP_SHOW_BANNER_BEFORE_ROW, Array(CStr(bannerType), CStr(titleText), CStr(bannerText), CLng(rowIndex), CLng(gapRowsBefore), CLng(gapRowsAfter))
        m_ShowBannerBeforeRowIndex = bannerText
        Exit Function
    End If

    bannerStartRow = rowIndex + gapRowsBefore

    If ex_Messaging.m_TryGetBannerRangeAddressByText(ws, bannerText, existingRangeAddress) Then
        On Error Resume Next
        Set existingRange = ws.Range(existingRangeAddress)
        On Error GoTo 0

        If Not existingRange Is Nothing Then
            ' Если текущий баннер уже стоит перед нужной строкой,
            ' просто обновляем его содержимое на месте.
            If existingRange.Row = bannerStartRow Then
                ex_Messaging.m_RenderTextBanner ws, bannerText, titleText, existingRangeAddress, bannerKind
                m_ShowBannerBeforeRowIndex = bannerText
                Exit Function
            End If

            ' Если якорь остался от предыдущего прогона в другой строке,
            ' удаляем старые строки банера целиком, чтобы не оставался пустой "дубликат".
            existingStartRow = existingRange.Row
            existingRowCount = existingRange.Rows.Count
            existingEndRow = existingStartRow + existingRowCount - 1

            If existingRowCount > 0 Then
                On Error Resume Next
                existingRange.UnMerge
                ws.Rows(CStr(existingStartRow) & ":" & CStr(existingEndRow)).Delete Shift:=xlUp
                On Error GoTo 0

                ' Если удалили строки выше целевой owner-строки, сдвигаем rowIndex вверх.
                If existingStartRow < rowIndex Then
                    rowIndex = rowIndex - existingRowCount
                    If rowIndex < 1 Then rowIndex = 1
                End If
            End If
        End If
    End If

    mp_GetWarningBannerDimensions bannerCols, bannerRows, bannerKind
    rowsToInsert = gapRowsBefore + bannerRows + gapRowsAfter
    If rowsToInsert > 0 Then
        ws.Rows(CStr(rowIndex) & ":" & CStr(rowIndex + rowsToInsert - 1)).Insert Shift:=xlDown
        mp_UnmergeRowsSafe ws, rowIndex, gapRowsBefore
        mp_UnmergeRowsSafe ws, rowIndex + gapRowsBefore + bannerRows, gapRowsAfter
    End If

    bannerRangeAddress = "A" & CStr(rowIndex + gapRowsBefore) & ":" & mp_ToColumnLetter(bannerCols) & CStr(rowIndex + gapRowsBefore + bannerRows - 1)
    ex_Messaging.m_RenderTextBanner ws, bannerText, titleText, bannerRangeAddress, bannerKind
    m_ShowBannerBeforeRowIndex = bannerText
End Function

Private Sub mp_UnmergeRowsSafe(ByVal ws As Worksheet, ByVal startRow As Long, ByVal rowCount As Long)
    Dim endRow As Long

    If ws Is Nothing Then Exit Sub
    If rowCount <= 0 Then Exit Sub
    If startRow < 1 Then startRow = 1
    If startRow > ws.Rows.Count Then Exit Sub

    endRow = startRow + rowCount - 1
    If endRow > ws.Rows.Count Then endRow = ws.Rows.Count
    If endRow < startRow Then Exit Sub

    On Error Resume Next
    ws.Rows(CStr(startRow) & ":" & CStr(endRow)).UnMerge
    ws.Rows(CStr(startRow) & ":" & CStr(endRow)).RowHeight = ws.StandardHeight
    On Error GoTo 0
End Sub

Private Function mp_NormalizeBannerType(ByVal bannerType As String) As String
    Dim normalized As String

    normalized = UCase$(Trim$(bannerType))
    If Len(normalized) = 0 Then normalized = BANNER_TYPE_WARNING

    Select Case normalized
        Case BANNER_TYPE_WARNING, BANNER_TYPE_ERROR, BANNER_TYPE_NOTE
            mp_NormalizeBannerType = normalized
        Case BANNER_TITLE_WARNING
            mp_NormalizeBannerType = BANNER_TYPE_WARNING
        Case BANNER_TITLE_ERROR
            mp_NormalizeBannerType = BANNER_TYPE_ERROR
        Case BANNER_TITLE_NOTE
            mp_NormalizeBannerType = BANNER_TYPE_NOTE
        Case UCase$(BANNER_KIND_WARNING)
            mp_NormalizeBannerType = BANNER_TYPE_WARNING
        Case UCase$(BANNER_KIND_ERROR)
            mp_NormalizeBannerType = BANNER_TYPE_ERROR
        Case UCase$(BANNER_KIND_NOTE)
            mp_NormalizeBannerType = BANNER_TYPE_NOTE
        Case Else
            Err.Raise vbObjectError + 1737, "ex_PostProcessActions", "Unsupported banner type '" & bannerType & "'. Allowed: TYPE_WARNING, TYPE_ERROR, TYPE_NOTE."
    End Select
End Function

Private Function mp_MapBannerTypeToKind(ByVal bannerType As String) As String
    Select Case mp_NormalizeBannerType(bannerType)
        Case BANNER_TYPE_ERROR
            mp_MapBannerTypeToKind = BANNER_KIND_ERROR
        Case BANNER_TYPE_NOTE
            mp_MapBannerTypeToKind = BANNER_KIND_NOTE
        Case Else
            mp_MapBannerTypeToKind = BANNER_KIND_WARNING
    End Select
End Function

Private Function mp_ResolveBannerTitle(ByVal titleText As String, ByVal bannerType As String) As String
    titleText = Trim$(titleText)
    If Len(titleText) > 0 Then
        mp_ResolveBannerTitle = titleText
        Exit Function
    End If

    Select Case mp_NormalizeBannerType(bannerType)
        Case BANNER_TYPE_ERROR
            mp_ResolveBannerTitle = BANNER_TITLE_ERROR
        Case BANNER_TYPE_NOTE
            mp_ResolveBannerTitle = BANNER_TITLE_NOTE
        Case Else
            mp_ResolveBannerTitle = BANNER_TITLE_WARNING
    End Select
End Function

Public Function m_GetRelativeDayOfMonth(ByVal dayOffsetText As String) As String
    Dim dayOffset As Long
    dayOffsetText = Trim$(dayOffsetText)
    If Not ex_XmlCore.m_TryParseLong(dayOffsetText, dayOffset) Then
        Err.Raise vbObjectError + 1687, "ex_PostProcessActions", "Day offset must be integer."
    End If
    m_GetRelativeDayOfMonth = Format$(DateAdd("d", dayOffset, Date), "dd")
End Function

Public Function m_GetRelativeRowCellText( _
    ByVal rowRef As Object, _
    ByVal rowOffsetText As String, _
    ByVal columnRef As String _
) As String
    Dim sourceRowRef As obj_ResultRow
    Dim ws As Worksheet
    Dim rowOffset As Long
    Dim targetCol As Long
    Dim targetRow As Long
    Dim sourceRowIndex As Long

    If rowRef Is Nothing Then
        Err.Raise vbObjectError + 1688, "ex_PostProcessActions", "Row reference is required for relative row read."
    End If
    If Not TypeOf rowRef Is obj_ResultRow Then
        Err.Raise vbObjectError + 1695, "ex_PostProcessActions", "Row reference must be obj_ResultRow for relative row read."
    End If
    Set sourceRowRef = rowRef

    rowOffsetText = Trim$(rowOffsetText)
    If Not ex_XmlCore.m_TryParseLong(rowOffsetText, rowOffset) Then
        Err.Raise vbObjectError + 1689, "ex_PostProcessActions", "Row offset must be integer."
    End If

    If Not mp_TryResolveColumnIndexInRow(sourceRowRef, columnRef, targetCol) Then
        Err.Raise vbObjectError + 1690, "ex_PostProcessActions", "Unknown column reference '" & columnRef & "' for relative row read."
    End If

    Set ws = ActiveSheet
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1691, "ex_PostProcessActions", "Active sheet is not available for relative row read."
    End If

    sourceRowIndex = mp_ResolveAnchoredRowIndex(ws, sourceRowRef, "relative row cell read")
    targetRow = sourceRowIndex + rowOffset
    If targetRow < 1 Then Exit Function
    If targetRow > ws.Rows.Count Then Exit Function

    m_GetRelativeRowCellText = Trim$(CStr(ws.Cells(targetRow, targetCol).Value))
End Function

Public Function m_GetRelativeRow( _
    ByVal rowRef As Object, _
    ByVal rowOffsetText As String _
) As obj_ResultRow
    Dim sourceRowRef As obj_ResultRow
    Dim ws As Worksheet
    Dim rowOffset As Long
    Dim targetRow As Long
    Dim sourceColumns As Collection
    Dim i As Long
    Dim colObj As obj_ResultColumn
    Dim valueText As String
    Dim resultRow As obj_ResultRow
    Dim hasTargetRow As Boolean
    Dim sourceRowIndex As Long

    If rowRef Is Nothing Then
        Err.Raise vbObjectError + 1692, "ex_PostProcessActions", "Row reference is required for relative row read."
    End If
    If Not TypeOf rowRef Is obj_ResultRow Then
        Err.Raise vbObjectError + 1696, "ex_PostProcessActions", "Row reference must be obj_ResultRow for relative row read."
    End If
    Set sourceRowRef = rowRef

    rowOffsetText = Trim$(rowOffsetText)
    If Not ex_XmlCore.m_TryParseLong(rowOffsetText, rowOffset) Then
        Err.Raise vbObjectError + 1693, "ex_PostProcessActions", "Row offset must be integer."
    End If

    Set ws = ActiveSheet
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1694, "ex_PostProcessActions", "Active sheet is not available for relative row read."
    End If

    sourceRowIndex = mp_ResolveAnchoredRowIndex(ws, sourceRowRef, "relative row read")
    targetRow = sourceRowIndex + rowOffset
    hasTargetRow = (targetRow >= 1 And targetRow <= ws.Rows.Count)

    Set resultRow = New obj_ResultRow
    If hasTargetRow Then
        resultRow.Initialize targetRow
        mp_AssignRuntimeRowAnchor ws, resultRow, targetRow
    Else
        resultRow.Initialize sourceRowIndex
    End If

    Set sourceColumns = sourceRowRef.Columns
    For i = 1 To sourceColumns.Count
        Set colObj = sourceColumns(i)
        If hasTargetRow Then
            valueText = Trim$(CStr(ws.Cells(targetRow, i).Value))
        Else
            valueText = vbNullString
        End If
        resultRow.SetValue colObj.Alias, colObj.MapKey, valueText
    Next i

    Set m_GetRelativeRow = resultRow
End Function

Private Function mp_TryLoadPostProcessHeaderStyle(ByRef outStyle As t_PostProcessHeaderStyle) As Boolean
    If Not mp_TryLoadPostProcessStyleByLayer( _
        POST_PROCESS_HEADER_LAYER_ID, _
        outStyle.Columns, _
        outStyle.Overflow, _
        outStyle.BackColor, _
        outStyle.FontColor, _
        outStyle.FontSize, _
        outStyle.RowHeight, _
        outStyle.MinRowHeight, _
        outStyle.AutoHeight, _
        outStyle.AutoHeightMarginTop, _
        outStyle.AutoHeightMarginBottom _
    ) Then Exit Function

    mp_TryLoadPostProcessHeaderStyle = True
End Function

Private Function mp_TryLoadPostProcessFooterStyle(ByRef outStyle As t_PostProcessFooterStyle) As Boolean
    If Not mp_TryLoadPostProcessStyleByLayer( _
        POST_PROCESS_FOOTER_LAYER_ID, _
        outStyle.Columns, _
        outStyle.Overflow, _
        outStyle.BackColor, _
        outStyle.FontColor, _
        outStyle.FontSize, _
        outStyle.RowHeight, _
        outStyle.MinRowHeight, _
        outStyle.AutoHeight, _
        outStyle.AutoHeightMarginTop, _
        outStyle.AutoHeightMarginBottom _
    ) Then Exit Function

    mp_TryLoadPostProcessFooterStyle = True
End Function

Private Function mp_TryLoadPostProcessStyleByLayer( _
    ByVal layerId As String, _
    ByRef outColumns As Long, _
    ByRef outOverflow As String, _
    ByRef outBackColor As Long, _
    ByRef outFontColor As Long, _
    ByRef outFontSize As Double, _
    ByRef outRowHeight As Double, _
    ByRef outMinRowHeight As Double, _
    ByRef outAutoHeight As Boolean, _
    ByRef outAutoHeightMarginTop As Double, _
    ByRef outAutoHeightMarginBottom As Double _
) As Boolean
    Dim ws As Worksheet
    Dim pageName As String
    Dim stageLayers As Collection
    Dim layerObj As obj_StyleLayer
    Dim ruleObj As obj_StyleRule
    Dim declarations As Object
    Dim textValue As String

    Set ws = ActiveSheet
    If ws Is Nothing Then
        MsgBox "Active sheet is not available for post-process style loading.", vbExclamation
        Exit Function
    End If
    pageName = Trim$(ws.Name)
    If Len(pageName) = 0 Then
        MsgBox "Active sheet name is empty for post-process style loading.", vbExclamation
        Exit Function
    End If

    Set stageLayers = ex_StylePipelineEngine.m_LoadSheetPipelineLayers(pageName, ThisWorkbook, POST_PROCESS_STYLE_STAGE_NAME)
    If stageLayers Is Nothing Or stageLayers.Count = 0 Then
        MsgBox "StylePipeline has no stage '" & POST_PROCESS_STYLE_STAGE_NAME & "' for page '" & pageName & "'.", vbExclamation
        Exit Function
    End If

    Set layerObj = ex_StylePipelineEngine.m_GetLayer(stageLayers, layerId)
    If layerObj Is Nothing Then
        MsgBox "StylePipeline stage '" & POST_PROCESS_STYLE_STAGE_NAME & "' must contain layer '" & layerId & "' for page '" & pageName & "'.", vbExclamation
        Exit Function
    End If
    If layerObj.RuleCount <= 0 Then
        MsgBox "StylePipeline layer '" & layerId & "' must contain at least one rule with style declarations.", vbExclamation
        Exit Function
    End If

    Set ruleObj = layerObj.Rules(1)
    If ruleObj Is Nothing Then
        MsgBox "StylePipeline layer '" & layerId & "' has invalid first rule.", vbExclamation
        Exit Function
    End If
    Set declarations = ruleObj.Declarations
    If declarations Is Nothing Then
        MsgBox "StylePipeline layer '" & layerId & "' has empty declarations.", vbExclamation
        Exit Function
    End If

    outRowHeight = 0
    outMinRowHeight = 0
    outAutoHeight = False

    textValue = mp_ReadRequiredDeclText(declarations, "mergeColumns", layerId)
    If Len(textValue) = 0 Then Exit Function
    If Not ex_XmlCore.m_TryParseLong(textValue, outColumns) Then
        MsgBox "Invalid declaration 'mergeColumns' in style layer '" & layerId & "': expected integer.", vbExclamation
        Exit Function
    End If
    If outColumns < 1 Then
        MsgBox "Invalid declaration 'mergeColumns' in style layer '" & layerId & "': must be >= 1.", vbExclamation
        Exit Function
    End If

    outOverflow = LCase$(mp_ReadRequiredDeclText(declarations, "overflow", layerId))
    If Len(outOverflow) = 0 Then Exit Function
    Select Case outOverflow
        Case "wrap", "clip", "shrink"
        Case Else
            MsgBox "Invalid declaration 'overflow' in style layer '" & layerId & "': expected wrap, clip, or shrink.", vbExclamation
            Exit Function
    End Select

    textValue = mp_ReadRequiredDeclText(declarations, "backColor", layerId)
    If Len(textValue) = 0 Then Exit Function
    If Not ex_XmlCore.m_TryParseColor(textValue, outBackColor) Then
        MsgBox "Invalid declaration 'backColor' in style layer '" & layerId & "'.", vbExclamation
        Exit Function
    End If

    textValue = mp_ReadRequiredDeclText(declarations, "fontColor", layerId)
    If Len(textValue) = 0 Then Exit Function
    If Not ex_XmlCore.m_TryParseColor(textValue, outFontColor) Then
        MsgBox "Invalid declaration 'fontColor' in style layer '" & layerId & "'.", vbExclamation
        Exit Function
    End If

    textValue = mp_ReadRequiredDeclText(declarations, "fontSize", layerId)
    If Len(textValue) = 0 Then Exit Function
    If Not ex_XmlCore.m_TryParseDouble(textValue, outFontSize, True) Then
        MsgBox "Invalid declaration 'fontSize' in style layer '" & layerId & "': expected number.", vbExclamation
        Exit Function
    End If
    If outFontSize <= 0 Then
        MsgBox "Invalid declaration 'fontSize' in style layer '" & layerId & "': must be > 0.", vbExclamation
        Exit Function
    End If

    If Not mp_TryReadOptionalDeclPositiveDouble(declarations, "rowHeight", layerId, outRowHeight) Then Exit Function
    If Not mp_TryReadOptionalDeclPositiveDouble(declarations, "minRowHeight", layerId, outMinRowHeight) Then Exit Function
    If Not mp_TryReadOptionalDeclBoolean(declarations, "autoHeight", layerId, outAutoHeight) Then Exit Function

    If Not mp_TryReadOptionalDeclNonNegativeDouble(declarations, "customAutoHeight-margin-top", layerId, outAutoHeightMarginTop) Then Exit Function
    If Not mp_TryReadOptionalDeclNonNegativeDouble(declarations, "customAutoHeight-margin-bottom", layerId, outAutoHeightMarginBottom) Then Exit Function

    mp_TryLoadPostProcessStyleByLayer = True
End Function

Private Function mp_ReadRequiredDeclText( _
    ByVal declarations As Object, _
    ByVal propertyName As String, _
    ByVal layerId As String _
) As String
    Dim key As String

    key = LCase$(Trim$(propertyName))
    If Len(key) = 0 Then Exit Function
    If declarations Is Nothing Then Exit Function

    If Not declarations.Exists(key) Then
        MsgBox "Missing declaration '" & propertyName & "' in style layer '" & layerId & "'.", vbExclamation
        Exit Function
    End If

    mp_ReadRequiredDeclText = Trim$(CStr(declarations(key)))
End Function

Private Function mp_TryReadOptionalDeclNonNegativeDouble( _
    ByVal declarations As Object, _
    ByVal propertyName As String, _
    ByVal layerId As String, _
    ByRef outValue As Double _
) As Boolean
    Dim key As String
    Dim textValue As String

    outValue = 0
    key = LCase$(Trim$(propertyName))
    If Len(key) = 0 Then
        mp_TryReadOptionalDeclNonNegativeDouble = True
        Exit Function
    End If
    If declarations Is Nothing Then
        mp_TryReadOptionalDeclNonNegativeDouble = True
        Exit Function
    End If
    If Not declarations.Exists(key) Then
        mp_TryReadOptionalDeclNonNegativeDouble = True
        Exit Function
    End If

    textValue = Trim$(CStr(declarations(key)))
    If Not ex_XmlCore.m_TryParseDouble(textValue, outValue, True) Then
        MsgBox "Invalid declaration '" & propertyName & "' in style layer '" & layerId & "': expected number >= 0.", vbExclamation
        Exit Function
    End If
    If outValue < 0 Then
        MsgBox "Invalid declaration '" & propertyName & "' in style layer '" & layerId & "': expected number >= 0.", vbExclamation
        Exit Function
    End If

    mp_TryReadOptionalDeclNonNegativeDouble = True
End Function

Private Function mp_TryReadOptionalDeclPositiveDouble( _
    ByVal declarations As Object, _
    ByVal propertyName As String, _
    ByVal layerId As String, _
    ByRef outValue As Double _
) As Boolean
    Dim key As String
    Dim textValue As String

    outValue = 0
    key = LCase$(Trim$(propertyName))
    If Len(key) = 0 Then
        mp_TryReadOptionalDeclPositiveDouble = True
        Exit Function
    End If
    If declarations Is Nothing Then
        mp_TryReadOptionalDeclPositiveDouble = True
        Exit Function
    End If
    If Not declarations.Exists(key) Then
        mp_TryReadOptionalDeclPositiveDouble = True
        Exit Function
    End If

    textValue = Trim$(CStr(declarations(key)))
    If Not ex_XmlCore.m_TryParseDouble(textValue, outValue, True) Then
        MsgBox "Invalid declaration '" & propertyName & "' in style layer '" & layerId & "': expected number > 0.", vbExclamation
        Exit Function
    End If
    If outValue <= 0 Then
        MsgBox "Invalid declaration '" & propertyName & "' in style layer '" & layerId & "': expected number > 0.", vbExclamation
        Exit Function
    End If

    mp_TryReadOptionalDeclPositiveDouble = True
End Function

Private Function mp_TryReadOptionalDeclBoolean( _
    ByVal declarations As Object, _
    ByVal propertyName As String, _
    ByVal layerId As String, _
    ByRef outValue As Boolean _
) As Boolean
    Dim key As String
    Dim textValue As String

    outValue = False
    key = LCase$(Trim$(propertyName))
    If Len(key) = 0 Then
        mp_TryReadOptionalDeclBoolean = True
        Exit Function
    End If
    If declarations Is Nothing Then
        mp_TryReadOptionalDeclBoolean = True
        Exit Function
    End If
    If Not declarations.Exists(key) Then
        mp_TryReadOptionalDeclBoolean = True
        Exit Function
    End If

    textValue = Trim$(CStr(declarations(key)))
    If Not ex_XmlCore.m_TryParseBoolean(textValue, outValue) Then
        MsgBox "Invalid declaration '" & propertyName & "' in style layer '" & layerId & "': expected true/false.", vbExclamation
        Exit Function
    End If

    mp_TryReadOptionalDeclBoolean = True
End Function

Private Sub mp_ApplyPostProcessHeaderKindStyle( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long _
)
    Dim stageLayers As Collection
    Dim headerLayer As obj_StyleLayer
    Dim headerPipeline As Collection
    Dim rowKindRanges As Object
    Dim headerRows As Collection
    Dim emptyTargets As Collection

    If ws Is Nothing Then Exit Sub
    If rowIndex < 1 Then Exit Sub

    Set stageLayers = ex_StylePipelineEngine.m_LoadSheetPipelineLayers(ws.Name, ThisWorkbook, POST_PROCESS_STYLE_STAGE_NAME)
    If stageLayers Is Nothing Or stageLayers.Count = 0 Then
        Err.Raise vbObjectError + 1737, "ex_PostProcessActions", _
            "StylePipeline has no stage '" & POST_PROCESS_STYLE_STAGE_NAME & "' for page '" & ws.Name & "'."
    End If

    Set headerLayer = ex_StylePipelineEngine.m_GetLayer(stageLayers, POST_PROCESS_HEADER_LAYER_ID)
    If headerLayer Is Nothing Then
        Err.Raise vbObjectError + 1738, "ex_PostProcessActions", _
            "StylePipeline stage '" & POST_PROCESS_STYLE_STAGE_NAME & "' must contain layer '" & POST_PROCESS_HEADER_LAYER_ID & "' for page '" & ws.Name & "'."
    End If

    Set headerPipeline = ex_StylePipelineEngine.m_CreatePipeline()
    ex_StylePipelineEngine.m_AddLayer headerPipeline, headerLayer

    Set rowKindRanges = CreateObject("Scripting.Dictionary")
    rowKindRanges.CompareMode = 1
    Set headerRows = New Collection
    headerRows.Add CLng(rowIndex)
    Set rowKindRanges(POST_PROCESS_HEADER_ROW_KIND) = headerRows

    Set emptyTargets = New Collection
    ex_StylePipelineEngine.m_ApplyColumnStylesPipeline ws, emptyTargets, headerPipeline, vbNullString, rowKindRanges
End Sub

Private Sub mp_ApplyPostProcessFooterKindStyle( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long _
)
    Dim stageLayers As Collection
    Dim footerLayer As obj_StyleLayer
    Dim footerPipeline As Collection
    Dim rowKindRanges As Object
    Dim footerRows As Collection
    Dim emptyTargets As Collection

    If ws Is Nothing Then Exit Sub
    If rowIndex < 1 Then Exit Sub

    Set stageLayers = ex_StylePipelineEngine.m_LoadSheetPipelineLayers(ws.Name, ThisWorkbook, POST_PROCESS_STYLE_STAGE_NAME)
    If stageLayers Is Nothing Or stageLayers.Count = 0 Then
        Err.Raise vbObjectError + 1739, "ex_PostProcessActions", _
            "StylePipeline has no stage '" & POST_PROCESS_STYLE_STAGE_NAME & "' for page '" & ws.Name & "'."
    End If

    Set footerLayer = ex_StylePipelineEngine.m_GetLayer(stageLayers, POST_PROCESS_FOOTER_LAYER_ID)
    If footerLayer Is Nothing Then
        Err.Raise vbObjectError + 1740, "ex_PostProcessActions", _
            "StylePipeline stage '" & POST_PROCESS_STYLE_STAGE_NAME & "' must contain layer '" & POST_PROCESS_FOOTER_LAYER_ID & "' for page '" & ws.Name & "'."
    End If

    Set footerPipeline = ex_StylePipelineEngine.m_CreatePipeline()
    ex_StylePipelineEngine.m_AddLayer footerPipeline, footerLayer

    Set rowKindRanges = CreateObject("Scripting.Dictionary")
    rowKindRanges.CompareMode = 1
    Set footerRows = New Collection
    footerRows.Add CLng(rowIndex)
    Set rowKindRanges(POST_PROCESS_FOOTER_ROW_KIND) = footerRows

    Set emptyTargets = New Collection
    ex_StylePipelineEngine.m_ApplyColumnStylesPipeline ws, emptyTargets, footerPipeline, vbNullString, rowKindRanges
End Sub

Private Sub mp_ApplyPostProcessHeaderRowHeight( _
    ByVal ws As Worksheet, _
    ByVal postProcessHeaderRange As Range, _
    ByVal postProcessHeaderText As String, _
    ByRef postProcessHeaderStyle As t_PostProcessHeaderStyle _
)
    If ws Is Nothing Then Exit Sub
    If postProcessHeaderRange Is Nothing Then Exit Sub

    ex_SheetHelpers.m_ApplySingleRowTextAutoHeight _
        ws, _
        postProcessHeaderRange, _
        postProcessHeaderText, _
        postProcessHeaderStyle.RowHeight, _
        postProcessHeaderStyle.MinRowHeight, _
        postProcessHeaderStyle.AutoHeight, _
        (StrComp(postProcessHeaderStyle.Overflow, "wrap", vbTextCompare) = 0), _
        postProcessHeaderStyle.AutoHeightMarginTop, _
        postProcessHeaderStyle.AutoHeightMarginBottom, _
        POST_PROCESS_MEASURE_SIDE_MARGIN, _
        POST_PROCESS_MEASURE_VERTICAL_MARGIN, _
        mp_GetMeasureExtraHeight(postProcessHeaderStyle.FontSize), _
        POST_PROCESS_MEASURE_HEIGHT_ROUND_PAD, _
        postProcessHeaderStyle.FontSize
End Sub

Private Sub mp_ApplyPostProcessFooterRowHeight( _
    ByVal ws As Worksheet, _
    ByVal postProcessFooterRange As Range, _
    ByVal postProcessFooterText As String, _
    ByRef postProcessFooterStyle As t_PostProcessFooterStyle _
)
    If ws Is Nothing Then Exit Sub
    If postProcessFooterRange Is Nothing Then Exit Sub

    ex_SheetHelpers.m_ApplySingleRowTextAutoHeight _
        ws, _
        postProcessFooterRange, _
        postProcessFooterText, _
        postProcessFooterStyle.RowHeight, _
        postProcessFooterStyle.MinRowHeight, _
        postProcessFooterStyle.AutoHeight, _
        (StrComp(postProcessFooterStyle.Overflow, "wrap", vbTextCompare) = 0), _
        postProcessFooterStyle.AutoHeightMarginTop, _
        postProcessFooterStyle.AutoHeightMarginBottom, _
        POST_PROCESS_MEASURE_SIDE_MARGIN, _
        POST_PROCESS_MEASURE_VERTICAL_MARGIN, _
        mp_GetMeasureExtraHeight(postProcessFooterStyle.FontSize), _
        POST_PROCESS_MEASURE_HEIGHT_ROUND_PAD, _
        postProcessFooterStyle.FontSize
End Sub

Private Function mp_MeasurePostProcessHeaderTextHeight( _
    ByVal ws As Worksheet, _
    ByVal postProcessHeaderRange As Range, _
    ByVal postProcessHeaderText As String, _
    ByVal fontSize As Double _
) As Double
    mp_MeasurePostProcessHeaderTextHeight = ex_SheetHelpers.m_MeasureTextHeight( _
        ws, _
        postProcessHeaderRange, _
        postProcessHeaderText, _
        POST_PROCESS_MEASURE_SIDE_MARGIN, _
        POST_PROCESS_MEASURE_VERTICAL_MARGIN, _
        fontSize, _
        CStr(postProcessHeaderRange.Font.Name) _
    ) + mp_GetMeasureExtraHeight(fontSize)
End Function

Private Function mp_MeasurePostProcessFooterTextHeight( _
    ByVal ws As Worksheet, _
    ByVal postProcessFooterRange As Range, _
    ByVal postProcessFooterText As String, _
    ByVal fontSize As Double _
) As Double
    mp_MeasurePostProcessFooterTextHeight = ex_SheetHelpers.m_MeasureTextHeight( _
        ws, _
        postProcessFooterRange, _
        postProcessFooterText, _
        POST_PROCESS_MEASURE_SIDE_MARGIN, _
        POST_PROCESS_MEASURE_VERTICAL_MARGIN, _
        fontSize, _
        CStr(postProcessFooterRange.Font.Name) _
    ) + mp_GetMeasureExtraHeight(fontSize)
End Function

Private Function mp_GetFirstUsedRow(ByVal ws As Worksheet) As Long
    Dim firstUsedCell As Range

    On Error GoTo ExitFn
    Set firstUsedCell = ws.Cells.Find(What:="*", After:=ws.Cells(ws.Rows.Count, ws.Columns.Count), SearchOrder:=xlByRows, SearchDirection:=xlNext)
    If Not firstUsedCell Is Nothing Then mp_GetFirstUsedRow = firstUsedCell.Row
ExitFn:
End Function

Private Function mp_GetPostProcessHeaderInsertStartRow(ByVal ws As Worksheet) As Long
    Dim outputViewStartRow As Long

    If ws Is Nothing Then
        Err.Raise vbObjectError + 1735, "ex_PostProcessActions", "Target sheet is not available for postProcessHeader insert start row."
    End If

    outputViewStartRow = ex_SheetStylesXmlProvider.m_GetOutputViewStartRow(ThisWorkbook)
    If outputViewStartRow < 1 Then
        Err.Raise vbObjectError + 1736, "ex_PostProcessActions", "Unable to resolve valid output view start row for postProcessHeader."
    End If

    mp_GetPostProcessHeaderInsertStartRow = outputViewStartRow
End Function

Private Function mp_IsRowBlankSegment( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal endCol As Long _
) As Boolean
    Dim probeRange As Range

    If ws Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > ws.Rows.Count Then Exit Function
    If endCol < 1 Then endCol = 1
    If endCol > ws.Columns.Count Then endCol = ws.Columns.Count

    Set probeRange = ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, endCol))
    mp_IsRowBlankSegment = (Application.WorksheetFunction.CountA(probeRange) = 0)
End Function

Private Function mp_IsSinglePostProcessTextRowShape( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal endCol As Long _
) As Boolean
    Dim probeRange As Range

    If ws Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > ws.Rows.Count Then Exit Function

    If endCol < 1 Then endCol = 1
    If endCol > ws.Columns.Count Then endCol = ws.Columns.Count

    Set probeRange = ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, endCol))
    If Application.WorksheetFunction.CountA(probeRange) <> 1 Then Exit Function
    If Len(Trim$(CStr(ws.Cells(rowIndex, 1).Value))) = 0 Then Exit Function

    mp_IsSinglePostProcessTextRowShape = True
End Function

Private Function mp_IsHeaderAnchorRowValid( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal endCol As Long _
) As Boolean
    If Not mp_IsSinglePostProcessTextRowShape(ws, rowIndex, endCol) Then Exit Function
    mp_IsHeaderAnchorRowValid = True
End Function

Private Function mp_IsFooterAnchorRowValid( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal endCol As Long _
) As Boolean
    If Not mp_IsSinglePostProcessTextRowShape(ws, rowIndex, endCol) Then Exit Function
    mp_IsFooterAnchorRowValid = True
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

Private Function mp_GetMeasureExtraHeight(ByVal fontSize As Double) As Double
    Dim dynamicExtra As Double

    dynamicExtra = fontSize * POST_PROCESS_MEASURE_EXTRA_HEIGHT_FONT_FACTOR
    If dynamicExtra < POST_PROCESS_MEASURE_EXTRA_HEIGHT_MIN Then
        dynamicExtra = POST_PROCESS_MEASURE_EXTRA_HEIGHT_MIN
    End If

    mp_GetMeasureExtraHeight = POST_PROCESS_MEASURE_EXTRA_HEIGHT_BASE + dynamicExtra
End Function

Private Function mp_RoundUpMeasuredHeight(ByVal measuredHeight As Double) As Double
    mp_RoundUpMeasuredHeight = ex_SheetHelpers.m_RoundUpMeasuredHeight(measuredHeight, POST_PROCESS_MEASURE_HEIGHT_ROUND_PAD)
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
    Dim ws As Worksheet
    Dim rowIndex As Long

    If rowRef Is Nothing Then
        Err.Raise vbObjectError + 1671, "ex_PostProcessActions", "Row reference is required for live cell text parsing."
    End If
    If Not mp_TryResolveColumnIndexInRow(rowRef, columnRef, targetCol) Then
        Err.Raise vbObjectError + 1672, "ex_PostProcessActions", "Unknown row cell reference '" & columnRef & "' for live cell text parsing."
    End If

    Set ws = ActiveSheet
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1746, "ex_PostProcessActions", "Active sheet is not available for live cell text parsing."
    End If
    rowIndex = mp_ResolveAnchoredRowIndex(ws, rowRef, "live cell text parsing")
    Set targetCell = mp_GetRowCellRange(rowIndex, targetCol)
    mp_GetRowCellLiveText = CStr(targetCell.Value)
End Function

Private Function mp_BuildSheetKey(ByVal ws As Worksheet) As String
    If ws Is Nothing Then Exit Function
    mp_BuildSheetKey = CStr(ws.Parent.Name) & "|" & CStr(ws.Name)
End Function

Private Function mp_EnsureRuntimeDataCache() As Object
    If g_RuntimeDataBySheetAndKey Is Nothing Then
        Set g_RuntimeDataBySheetAndKey = CreateObject("Scripting.Dictionary")
        g_RuntimeDataBySheetAndKey.CompareMode = 1 ' vbTextCompare
    End If
    Set mp_EnsureRuntimeDataCache = g_RuntimeDataBySheetAndKey
End Function

Private Sub mp_ClearRuntimeDataForSheet(ByVal ws As Worksheet)
    Dim cache As Object
    Dim key As Variant
    Dim prefix As String
    Dim keysToRemove As Collection

    If ws Is Nothing Then Exit Sub
    Set cache = g_RuntimeDataBySheetAndKey
    If cache Is Nothing Then Exit Sub
    If cache.Count = 0 Then Exit Sub

    prefix = mp_BuildSheetKey(ws) & "|"
    Set keysToRemove = New Collection

    For Each key In cache.Keys
        If StrComp(Left$(CStr(key), Len(prefix)), prefix, vbTextCompare) = 0 Then
            keysToRemove.Add CStr(key)
        End If
    Next key

    For Each key In keysToRemove
        cache.Remove CStr(key)
    Next key
End Sub

Private Function mp_NormalizeRuntimeDataKey(ByVal dataKey As String) As String
    dataKey = Trim$(dataKey)
    If Len(dataKey) = 0 Then Exit Function
    If InStr(1, dataKey, "|", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 1711, "ex_PostProcessActions", "Runtime data key cannot contain '|' character."
    End If
    mp_NormalizeRuntimeDataKey = LCase$(dataKey)
End Function

Private Function mp_BuildRuntimeDataFullKey(ByVal ws As Worksheet, ByVal dataKey As String) As String
    If ws Is Nothing Then Exit Function
    dataKey = mp_NormalizeRuntimeDataKey(dataKey)
    If Len(dataKey) = 0 Then Exit Function
    mp_BuildRuntimeDataFullKey = mp_BuildSheetKey(ws) & "|" & dataKey
End Function

Private Sub mp_EnsureDeferredStores()
    If g_DeferredSingleHeaderTextBySheet Is Nothing Then
        Set g_DeferredSingleHeaderTextBySheet = CreateObject("Scripting.Dictionary")
        g_DeferredSingleHeaderTextBySheet.CompareMode = 1 ' vbTextCompare
    End If
    If g_DeferredSingleFooterTextBySheet Is Nothing Then
        Set g_DeferredSingleFooterTextBySheet = CreateObject("Scripting.Dictionary")
        g_DeferredSingleFooterTextBySheet.CompareMode = 1 ' vbTextCompare
    End If
    If g_DeferredRowAnchorSeqBySheet Is Nothing Then
        Set g_DeferredRowAnchorSeqBySheet = CreateObject("Scripting.Dictionary")
        g_DeferredRowAnchorSeqBySheet.CompareMode = 1 ' vbTextCompare
    End If
End Sub

Private Function mp_BuildDeferredSheetKey(ByVal ws As Worksheet) As String
    mp_BuildDeferredSheetKey = mp_BuildSheetKey(ws)
End Function

Private Function mp_GetDeferredQueueForSheet(ByVal ws As Worksheet) As Collection
    If ws Is Nothing Then Exit Function
    Set mp_GetDeferredQueueForSheet = ex_RenderQueue.m_GetOrCreateQueueForSheet(ws)
End Function

Private Sub mp_SetDeferredRenderActive(ByVal ws As Worksheet, ByVal isActive As Boolean)
    If ws Is Nothing Then Exit Sub
    ex_RenderQueue.m_SetActiveForSheet ws, isActive
End Sub

Private Function mp_IsDeferredRenderActiveForSheet(ByVal ws As Worksheet) As Boolean
    mp_IsDeferredRenderActiveForSheet = ex_RenderQueue.m_IsActiveForSheet(ws)
End Function

Private Sub mp_ClearDeferredRenderSheetState(ByVal ws As Worksheet)
    Dim sheetKey As String

    If ws Is Nothing Then Exit Sub
    sheetKey = mp_BuildDeferredSheetKey(ws)
    If Len(sheetKey) = 0 Then Exit Sub

    ex_RenderQueue.m_ClearQueueForSheet ws
    If Not g_DeferredSingleHeaderTextBySheet Is Nothing Then
        If g_DeferredSingleHeaderTextBySheet.Exists(sheetKey) Then g_DeferredSingleHeaderTextBySheet.Remove sheetKey
    End If
    If Not g_DeferredSingleFooterTextBySheet Is Nothing Then
        If g_DeferredSingleFooterTextBySheet.Exists(sheetKey) Then g_DeferredSingleFooterTextBySheet.Remove sheetKey
    End If
    If Not g_DeferredRowAnchorSeqBySheet Is Nothing Then
        If g_DeferredRowAnchorSeqBySheet.Exists(sheetKey) Then g_DeferredRowAnchorSeqBySheet.Remove sheetKey
    End If
    mp_ClearDeferredRowAnchors ws
End Sub

Private Sub mp_QueueDeferredOperation( _
    ByVal ws As Worksheet, _
    ByVal opType As String, _
    ByVal args As Variant _
)
    Dim queue As Collection
    Dim op As Object
    Dim rowIndex As Long
    Dim anchorName As String

    If ws Is Nothing Then Exit Sub
    opType = LCase$(Trim$(opType))
    If Len(opType) = 0 Then Exit Sub

    Set queue = mp_GetDeferredQueueForSheet(ws)
    If queue Is Nothing Then Exit Sub

    Set op = CreateObject("Scripting.Dictionary")
    op.CompareMode = 1
    op("Type") = opType
    op("Args") = args

    If opType = DEFER_OP_SHOW_BANNER_BEFORE_ROW Then
        If IsArray(args) Then
            If UBound(args) >= 3 Then
                If mp_TryCoerceLong(args(3), rowIndex) Then
                    If rowIndex >= 1 And rowIndex <= ws.Rows.Count Then
                        anchorName = mp_NextDeferredRowAnchorName(ws)
                        If Len(anchorName) > 0 Then
                            mp_SetNamedRowAnchor ws, anchorName, rowIndex
                            op("AnchorName") = anchorName
                        End If
                    End If
                End If
            End If
        End If
    End If

    queue.Add op
End Sub

Private Sub mp_ApplyDeferredOperation(ByVal ws As Worksheet, ByVal op As Object)
    Dim args As Variant
    Dim opType As String
    Dim rowIndexValue As Long
    Dim anchoredRowIndex As Long
    Dim anchorName As String

    If ws Is Nothing Then Exit Sub
    If op Is Nothing Then Exit Sub
    If Not op.Exists("Type") Then Exit Sub

    opType = LCase$(Trim$(CStr(op("Type"))))
    If op.Exists("Args") Then args = op("Args")

    Select Case opType
        Case DEFER_OP_APPEND_HEADER_TEXT
            m_AppendToSinglePostProcessHeaderText CStr(args(0)), CStr(args(1))
        Case DEFER_OP_APPEND_FOOTER_TEXT
            m_AppendToSinglePostProcessFooterText CStr(args(0)), CStr(args(1))
        Case DEFER_OP_APPEND_HEADER_ROW
            m_AppendPostProcessHeaderText CStr(args(0))
        Case DEFER_OP_APPEND_FOOTER_ROW
            m_AppendPostProcessFooterText CStr(args(0))
        Case DEFER_OP_SHOW_BANNER_AT_CELL
            Call m_ShowBannerAtCell(CStr(args(0)), CStr(args(1)), CStr(args(2)), CStr(args(3)), mp_RequireLongArg(args, 4, "banner at cell"), mp_RequireLongArg(args, 5, "banner at cell"))
        Case DEFER_OP_SHOW_BANNER_AFTER_BANNER
            Call m_ShowBannerAfterBanner(CStr(args(0)), CStr(args(1)), CStr(args(2)), mp_RequireLongArg(args, 3, "after banner"), mp_RequireLongArg(args, 4, "after banner"), mp_RequireLongArg(args, 5, "after banner"))
        Case DEFER_OP_SHOW_BANNER_AT_TABLE
            Call m_ShowBannerAtTable(CStr(args(0)), CStr(args(1)), CStr(args(2)), CStr(args(3)), CStr(args(4)), mp_RequireLongArg(args, 5, "table banner"), mp_RequireLongArg(args, 6, "table banner"))
        Case DEFER_OP_SHOW_BANNER_BEFORE_ROW
            rowIndexValue = mp_RequireLongArg(args, 3, "banner insert")
            If op.Exists("AnchorName") Then
                anchorName = CStr(op("AnchorName"))
                If Len(anchorName) > 0 Then
                    If mp_TryGetNamedRowAnchor(ws, anchorName, anchoredRowIndex) Then rowIndexValue = anchoredRowIndex
                    mp_ClearNamedRowAnchor ws, anchorName
                End If
            End If
            Call m_ShowBannerBeforeRowIndex(CStr(args(0)), CStr(args(1)), CStr(args(2)), rowIndexValue, mp_RequireLongArg(args, 4, "banner insert"), mp_RequireLongArg(args, 5, "banner insert"))
        Case DEFER_OP_APPEND_RESULT_BLOCK
            Call m_AppendResultBlock(CStr(args(0)), CStr(args(1)), CStr(args(2)), mp_RequireLongArg(args, 3, "result block"), mp_RequireLongArg(args, 4, "result block"))
    End Select
End Sub

Private Function mp_RequireLongArg( _
    ByVal args As Variant, _
    ByVal argIndex As Long, _
    ByVal contextName As String _
) As Long
    Dim parsedValue As Long

    If Not IsArray(args) Then
        Err.Raise vbObjectError + 1773, "ex_PostProcessActions", "Deferred arguments are not array for " & contextName & "."
    End If
    If argIndex < LBound(args) Or argIndex > UBound(args) Then
        Err.Raise vbObjectError + 1774, "ex_PostProcessActions", "Deferred argument index is out of range for " & contextName & "."
    End If
    If Not mp_TryCoerceLong(args(argIndex), parsedValue) Then
        Err.Raise vbObjectError + 1775, "ex_PostProcessActions", "Deferred argument #" & CStr(argIndex) & " must be integer for " & contextName & "."
    End If

    mp_RequireLongArg = parsedValue
End Function

Private Function mp_TryCoerceLong(ByVal valueRef As Variant, ByRef outValue As Long) As Boolean
    Dim textValue As String

    On Error GoTo Fail

    If IsObject(valueRef) Then Exit Function
    If IsNull(valueRef) Or IsEmpty(valueRef) Then Exit Function

    If VarType(valueRef) = vbString Then
        textValue = Trim$(CStr(valueRef))
        If Len(textValue) = 0 Then Exit Function
        If Not ex_XmlCore.m_TryParseLong(textValue, outValue) Then Exit Function
        mp_TryCoerceLong = True
        Exit Function
    End If

    outValue = CLng(valueRef)
    mp_TryCoerceLong = True
    Exit Function

Fail:
    mp_TryCoerceLong = False
End Function

Private Function mp_GetDeferredOperationPhase(ByVal op As Object) As Long
    Dim opType As String

    mp_GetDeferredOperationPhase = 2
    If op Is Nothing Then Exit Function
    If Not op.Exists("Type") Then Exit Function

    opType = LCase$(Trim$(CStr(op("Type"))))
    Select Case opType
        Case DEFER_OP_APPEND_HEADER_TEXT, DEFER_OP_APPEND_HEADER_ROW
            mp_GetDeferredOperationPhase = 1
        Case DEFER_OP_APPEND_FOOTER_TEXT, DEFER_OP_APPEND_FOOTER_ROW
            mp_GetDeferredOperationPhase = 3
        Case Else
            mp_GetDeferredOperationPhase = 2
    End Select
End Function

Private Function mp_NextDeferredRowAnchorName(ByVal ws As Worksheet) As String
    mp_NextDeferredRowAnchorName = mp_NextSequentialRowAnchorName(ws, DEFER_ROW_ANCHOR_PREFIX)
End Function

Private Function mp_NextRuntimeRowAnchorName(ByVal ws As Worksheet) As String
    mp_NextRuntimeRowAnchorName = mp_NextSequentialRowAnchorName(ws, RUNTIME_ROW_ANCHOR_PREFIX)
End Function

Private Function mp_NextSequentialRowAnchorName( _
    ByVal ws As Worksheet, _
    ByVal anchorPrefix As String _
) As String
    Dim sheetKey As String
    Dim nextSeq As Long

    If ws Is Nothing Then Exit Function
    anchorPrefix = Trim$(anchorPrefix)
    If Len(anchorPrefix) = 0 Then Exit Function
    mp_EnsureDeferredStores
    sheetKey = mp_BuildDeferredSheetKey(ws)
    If Len(sheetKey) = 0 Then Exit Function

    If g_DeferredRowAnchorSeqBySheet.Exists(sheetKey) Then
        nextSeq = CLng(g_DeferredRowAnchorSeqBySheet(sheetKey)) + 1
    Else
        nextSeq = 1
    End If

    g_DeferredRowAnchorSeqBySheet(sheetKey) = CStr(nextSeq)
    mp_NextSequentialRowAnchorName = anchorPrefix & CStr(nextSeq)
End Function

Private Sub mp_ClearDeferredRowAnchors(ByVal ws As Worksheet)
    mp_ClearRowAnchorsByPrefix ws, DEFER_ROW_ANCHOR_PREFIX
    mp_ClearRowAnchorsByPrefix ws, RUNTIME_ROW_ANCHOR_PREFIX
End Sub

Private Sub mp_ClearRowAnchorsByPrefix(ByVal ws As Worksheet, ByVal anchorPrefix As String)
    Dim i As Long
    Dim entry As Name
    Dim localName As String
    Dim namePos As Long
    Dim localPrefix As String

    If ws Is Nothing Then Exit Sub
    localPrefix = LCase$(Trim$(anchorPrefix))
    If Len(localPrefix) = 0 Then Exit Sub

    On Error Resume Next
    For i = ws.Names.Count To 1 Step -1
        Set entry = ws.Names(i)
        localName = CStr(entry.Name)
        namePos = InStrRev(localName, "!", vbBinaryCompare)
        If namePos > 0 Then localName = Mid$(localName, namePos + 1)
        If LCase$(Left$(localName, Len(localPrefix))) = localPrefix Then
            entry.Delete
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub mp_AppendDeferredSingleHeaderText( _
    ByVal ws As Worksheet, _
    ByVal appendText As String, _
    ByVal separatorText As String _
)
    Dim sheetKey As String
    Dim currentText As String

    If ws Is Nothing Then Exit Sub
    mp_EnsureDeferredStores
    sheetKey = mp_BuildDeferredSheetKey(ws)
    If Len(sheetKey) = 0 Then Exit Sub

    If g_DeferredSingleHeaderTextBySheet.Exists(sheetKey) Then
        currentText = CStr(g_DeferredSingleHeaderTextBySheet(sheetKey))
    Else
        currentText = vbNullString
    End If

    g_DeferredSingleHeaderTextBySheet(sheetKey) = m_TextAppend(currentText, appendText, separatorText)
End Sub

Private Sub mp_AppendDeferredSingleFooterText( _
    ByVal ws As Worksheet, _
    ByVal appendText As String, _
    ByVal separatorText As String _
)
    Dim sheetKey As String
    Dim currentText As String

    If ws Is Nothing Then Exit Sub
    mp_EnsureDeferredStores
    sheetKey = mp_BuildDeferredSheetKey(ws)
    If Len(sheetKey) = 0 Then Exit Sub

    If g_DeferredSingleFooterTextBySheet.Exists(sheetKey) Then
        currentText = CStr(g_DeferredSingleFooterTextBySheet(sheetKey))
    Else
        currentText = vbNullString
    End If

    g_DeferredSingleFooterTextBySheet(sheetKey) = m_TextAppend(currentText, appendText, separatorText)
End Sub

Private Function mp_TryGetDeferredSingleHeaderText(ByVal ws As Worksheet, ByRef outText As String) As Boolean
    Dim sheetKey As String

    If ws Is Nothing Then Exit Function
    If g_DeferredSingleHeaderTextBySheet Is Nothing Then Exit Function
    sheetKey = mp_BuildDeferredSheetKey(ws)
    If Len(sheetKey) = 0 Then Exit Function
    If Not g_DeferredSingleHeaderTextBySheet.Exists(sheetKey) Then Exit Function

    outText = CStr(g_DeferredSingleHeaderTextBySheet(sheetKey))
    mp_TryGetDeferredSingleHeaderText = True
End Function

Private Function mp_TryGetDeferredSingleFooterText(ByVal ws As Worksheet, ByRef outText As String) As Boolean
    Dim sheetKey As String

    If ws Is Nothing Then Exit Function
    If g_DeferredSingleFooterTextBySheet Is Nothing Then Exit Function
    sheetKey = mp_BuildDeferredSheetKey(ws)
    If Len(sheetKey) = 0 Then Exit Function
    If Not g_DeferredSingleFooterTextBySheet.Exists(sheetKey) Then Exit Function

    outText = CStr(g_DeferredSingleFooterTextBySheet(sheetKey))
    mp_TryGetDeferredSingleFooterText = True
End Function

Private Function mp_HasDeferredSingleHeaderText(ByVal ws As Worksheet) As Boolean
    Dim textValue As String
    If mp_TryGetDeferredSingleHeaderText(ws, textValue) Then
        mp_HasDeferredSingleHeaderText = (Len(textValue) > 0)
    End If
End Function

Private Function mp_HasDeferredSingleFooterText(ByVal ws As Worksheet) As Boolean
    Dim textValue As String
    If mp_TryGetDeferredSingleFooterText(ws, textValue) Then
        mp_HasDeferredSingleFooterText = (Len(textValue) > 0)
    End If
End Function

Private Function mp_HasPostProcessFooterForSheet(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    If Not g_PostProcessFooterHasAppended Then Exit Function
    If g_PostProcessFooterRowIndex < 1 Then Exit Function
    If StrComp(g_PostProcessFooterSheetKey, mp_BuildSheetKey(ws), vbTextCompare) <> 0 Then Exit Function
    mp_HasPostProcessFooterForSheet = True
End Function

Private Function mp_TryGetNamedRowAnchor( _
    ByVal ws As Worksheet, _
    ByVal anchorName As String, _
    ByRef outRowIndex As Long _
) As Boolean
    Dim namedEntry As Name
    Dim anchorCell As Range

    If ws Is Nothing Then Exit Function
    anchorName = Trim$(anchorName)
    If Len(anchorName) = 0 Then Exit Function

    On Error Resume Next
    Set namedEntry = ws.Names(anchorName)
    On Error GoTo 0
    If namedEntry Is Nothing Then Exit Function

    On Error Resume Next
    Set anchorCell = namedEntry.RefersToRange
    On Error GoTo 0
    If anchorCell Is Nothing Then Exit Function

    outRowIndex = anchorCell.Row
    If outRowIndex < 1 Or outRowIndex > ws.Rows.Count Then Exit Function
    mp_TryGetNamedRowAnchor = True
End Function

Private Sub mp_ClearNamedRowAnchor( _
    ByVal ws As Worksheet, _
    ByVal anchorName As String _
)
    If ws Is Nothing Then Exit Sub
    anchorName = Trim$(anchorName)
    If Len(anchorName) = 0 Then Exit Sub

    On Error Resume Next
    ws.Names(anchorName).Delete
    ThisWorkbook.Names(anchorName).Delete
    On Error GoTo 0
End Sub

Private Sub mp_SetNamedRowAnchor( _
    ByVal ws As Worksheet, _
    ByVal anchorName As String, _
    ByVal rowIndex As Long _
)
    Dim refersToText As String
    Dim resolvedRowIndex As Long

    If ws Is Nothing Then Exit Sub
    anchorName = Trim$(anchorName)
    If Len(anchorName) = 0 Then Exit Sub
    If rowIndex < 1 Or rowIndex > ws.Rows.Count Then Exit Sub

    mp_ClearNamedRowAnchor ws, anchorName
    refersToText = "=" & ws.Cells(rowIndex, 1).Address(True, True, xlA1, True)

    On Error Resume Next
    ws.Names.Add Name:=anchorName, RefersTo:=refersToText
    On Error GoTo 0

    If Not mp_TryGetNamedRowAnchor(ws, anchorName, resolvedRowIndex) Then
        On Error Resume Next
        ws.Names.Add Name:=ws.Name & "!" & anchorName, RefersTo:=refersToText
        On Error GoTo 0
    End If
End Sub

Private Sub mp_ClearNamedRangeAnchor( _
    ByVal ws As Worksheet, _
    ByVal anchorName As String _
)
    If ws Is Nothing Then Exit Sub
    anchorName = Trim$(anchorName)
    If Len(anchorName) = 0 Then Exit Sub

    On Error Resume Next
    ws.Names(anchorName).Delete
    ThisWorkbook.Names(anchorName).Delete
    On Error GoTo 0
End Sub

Private Sub mp_SetNamedRangeAnchor( _
    ByVal ws As Worksheet, _
    ByVal anchorName As String, _
    ByVal targetRange As Range _
)
    Dim refersToText As String

    If ws Is Nothing Then Exit Sub
    If targetRange Is Nothing Then Exit Sub
    anchorName = Trim$(anchorName)
    If Len(anchorName) = 0 Then Exit Sub

    mp_ClearNamedRangeAnchor ws, anchorName
    refersToText = "=" & targetRange.Address(True, True, xlA1, True)

    On Error Resume Next
    ws.Names.Add Name:=anchorName, RefersTo:=refersToText
    If Err.Number <> 0 Then
        Err.Clear
        ws.Names.Add Name:=ws.Name & "!" & anchorName, RefersTo:=refersToText
    End If
    On Error GoTo 0
End Sub

Private Function mp_NextResultBlockAnchorName(ByVal ws As Worksheet) As String
    Dim idx As Long
    Dim anchorName As String
    Dim namedEntry As Name

    If ws Is Nothing Then Exit Function

    For idx = 1 To RESULT_BLOCK_ANCHOR_MAX_INDEX
        anchorName = RESULT_BLOCK_ANCHOR_PREFIX & Format$(idx, "0000")
        Set namedEntry = Nothing
        On Error Resume Next
        Set namedEntry = ws.Names(anchorName)
        On Error GoTo 0
        If namedEntry Is Nothing Then
            mp_NextResultBlockAnchorName = anchorName
            Exit Function
        End If
    Next idx
End Function

Private Function mp_BuildResultBlockIdentity( _
    ByVal sheetKey As String, _
    ByVal blockIndex As Long, _
    ByVal blockText As String _
) As String
    mp_BuildResultBlockIdentity = "__resultblock__" & CStr(sheetKey) & "_" & CStr(blockIndex) & "_" & CStr(blockText)
End Function

Private Sub mp_ClearPreviousResultBlocks(ByVal ws As Worksheet)
    Dim namedEntry As Name
    Dim normalizedName As String
    Dim refersRange As Range
    Dim blocks As Collection
    Dim namesToDelete As Collection
    Dim entry As Object
    Dim maxRow As Long
    Dim maxIndex As Long
    Dim i As Long
    Dim deleteName As Variant
    Dim rowStart As Long
    Dim rowEnd As Long

    If ws Is Nothing Then Exit Sub

    Set blocks = New Collection
    Set namesToDelete = New Collection

    For Each namedEntry In ws.Names
        normalizedName = LCase$(CStr(namedEntry.Name))
        If InStr(1, normalizedName, LCase$(RESULT_BLOCK_ANCHOR_PREFIX), vbBinaryCompare) > 0 Then
            namesToDelete.Add CStr(namedEntry.Name)
            Set refersRange = Nothing
            On Error Resume Next
            Set refersRange = namedEntry.RefersToRange
            On Error GoTo 0
            If Not refersRange Is Nothing Then
                Set entry = CreateObject("Scripting.Dictionary")
                entry.CompareMode = 1
                entry("RowStart") = CLng(refersRange.Row)
                entry("RowEnd") = CLng(refersRange.Row + refersRange.Rows.Count - 1)
                blocks.Add entry
            End If
        End If
    Next namedEntry

    Do While blocks.Count > 0
        maxIndex = 0
        maxRow = -1
        For i = 1 To blocks.Count
            Set entry = blocks(i)
            If CLng(entry("RowStart")) > maxRow Then
                maxRow = CLng(entry("RowStart"))
                maxIndex = i
            End If
        Next i
        If maxIndex <= 0 Then Exit Do

        Set entry = blocks(maxIndex)
        rowStart = CLng(entry("RowStart"))
        rowEnd = CLng(entry("RowEnd"))
        If rowStart < 1 Then rowStart = 1
        If rowEnd > ws.Rows.Count Then rowEnd = ws.Rows.Count
        If rowEnd >= rowStart Then
            On Error Resume Next
            ws.Rows(CStr(rowStart) & ":" & CStr(rowEnd)).Delete Shift:=xlUp
            On Error GoTo 0
        End If
        blocks.Remove maxIndex
    Loop

    For Each deleteName In namesToDelete
        On Error Resume Next
        ws.Names(CStr(deleteName)).Delete
        ThisWorkbook.Names(CStr(deleteName)).Delete
        On Error GoTo 0
    Next deleteName

    mp_ClearNamedRowAnchor ws, RESULT_BLOCK_LAST_ANCHOR_NAME
End Sub

Private Sub mp_SetPostProcessHeaderAnchors(ByVal ws As Worksheet, ByVal rowIndex As Long)
    If ws Is Nothing Then Exit Sub
    mp_SetNamedRowAnchor ws, POST_PROCESS_HEADER_ANCHOR_NAME, rowIndex
End Sub

Private Sub mp_SetPostProcessFooterAnchors(ByVal ws As Worksheet, ByVal rowIndex As Long)
    If ws Is Nothing Then Exit Sub
    mp_SetNamedRowAnchor ws, POST_PROCESS_FOOTER_ANCHOR_NAME, rowIndex
End Sub

Private Function mp_TryGetPostProcessHeaderAnchorRow( _
    ByVal ws As Worksheet, _
    ByRef outRowIndex As Long _
) As Boolean
    mp_TryGetPostProcessHeaderAnchorRow = mp_TryGetNamedRowAnchor(ws, POST_PROCESS_HEADER_ANCHOR_NAME, outRowIndex)
End Function

Private Function mp_TryGetPostProcessFooterAnchorRow( _
    ByVal ws As Worksheet, _
    ByRef outRowIndex As Long _
) As Boolean
    mp_TryGetPostProcessFooterAnchorRow = mp_TryGetNamedRowAnchor(ws, POST_PROCESS_FOOTER_ANCHOR_NAME, outRowIndex)
End Function

Private Sub mp_ClearPostProcessHeaderAnchors(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    mp_ClearNamedRowAnchor ws, POST_PROCESS_HEADER_ANCHOR_NAME
End Sub

Private Sub mp_ClearPostProcessFooterAnchors(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    mp_ClearNamedRowAnchor ws, POST_PROCESS_FOOTER_ANCHOR_NAME
End Sub

Private Sub mp_ClearPreviousSinglePostProcessHeader(ByVal ws As Worksheet)
    Dim headerRow As Long
    Dim deleteStartRow As Long
    Dim deleteEndRow As Long

    If ws Is Nothing Then Exit Sub

    If mp_TryGetPostProcessHeaderAnchorRow(ws, headerRow) Then
        deleteStartRow = headerRow - 1
        deleteEndRow = headerRow + 1
        If deleteStartRow < 1 Then deleteStartRow = 1
        If deleteEndRow > ws.Rows.Count Then deleteEndRow = ws.Rows.Count
        If deleteEndRow >= deleteStartRow Then
            ws.Rows(CStr(deleteStartRow) & ":" & CStr(deleteEndRow)).Delete Shift:=xlUp
        End If
    End If

    mp_ClearPostProcessHeaderAnchors ws
End Sub

Private Sub mp_ClearPreviousSinglePostProcessFooter(ByVal ws As Worksheet)
    Dim footerRow As Long
    Dim usedCols As Long
    Dim footerRange As Range

    If ws Is Nothing Then Exit Sub

    If mp_TryGetPostProcessFooterAnchorRow(ws, footerRow) Then
        usedCols = mp_GetLastUsedColumn(ws)
        If usedCols < 1 Then usedCols = 1
        Set footerRange = ws.Range(ws.Cells(footerRow, 1), ws.Cells(footerRow, usedCols))
        footerRange.UnMerge
        footerRange.ClearContents
    End If

    mp_ClearPostProcessFooterAnchors ws
End Sub

Private Function mp_TryGetCachedSinglePostProcessHeaderRowIndex( _
    ByVal ws As Worksheet, _
    ByRef outRowIndex As Long _
) As Boolean
    Dim sheetKey As String
    Dim rowIndex As Long
    Dim viewStartRow As Long
    Dim probeEndCol As Long

    If ws Is Nothing Then Exit Function
    viewStartRow = mp_GetPostProcessHeaderInsertStartRow(ws)
    probeEndCol = mp_GetLastUsedColumn(ws)
    If probeEndCol < 1 Then probeEndCol = 1

    sheetKey = mp_BuildSheetKey(ws)
    If StrComp(g_PostProcessHeaderSheetKey, sheetKey, vbTextCompare) = 0 Then
        rowIndex = g_PostProcessHeaderRowIndex
        If rowIndex >= viewStartRow _
            And rowIndex <= (viewStartRow + POST_PROCESS_HEADER_ANCHOR_MAX_OFFSET_ROWS) _
            And rowIndex <= ws.Rows.Count Then
            If mp_IsHeaderAnchorRowValid(ws, rowIndex, probeEndCol) Then
                outRowIndex = rowIndex
                mp_TryGetCachedSinglePostProcessHeaderRowIndex = True
                Exit Function
            End If
        End If
    End If

    If mp_TryGetPostProcessHeaderAnchorRow(ws, rowIndex) Then
        If rowIndex < viewStartRow Then Exit Function
        If rowIndex > (viewStartRow + POST_PROCESS_HEADER_ANCHOR_MAX_OFFSET_ROWS) Then Exit Function
        If Not mp_IsHeaderAnchorRowValid(ws, rowIndex, probeEndCol) Then
            mp_ClearPostProcessHeaderAnchors ws
            Exit Function
        End If
        g_PostProcessHeaderSheetKey = sheetKey
        g_PostProcessHeaderRowIndex = rowIndex
        g_PostProcessHeaderNextInsertRow = rowIndex + 2
        If g_PostProcessHeaderNextInsertRow > ws.Rows.Count Then g_PostProcessHeaderNextInsertRow = ws.Rows.Count
        outRowIndex = rowIndex
        mp_TryGetCachedSinglePostProcessHeaderRowIndex = True
    End If
End Function

Private Function mp_TryGetCachedSinglePostProcessFooterRowIndex( _
    ByVal ws As Worksheet, _
    ByRef outRowIndex As Long _
) As Boolean
    Dim sheetKey As String
    Dim rowIndex As Long
    Dim probeEndCol As Long

    If ws Is Nothing Then Exit Function
    probeEndCol = mp_GetLastUsedColumn(ws)
    If probeEndCol < 1 Then probeEndCol = 1

    sheetKey = mp_BuildSheetKey(ws)
    If StrComp(g_PostProcessFooterSheetKey, sheetKey, vbTextCompare) = 0 Then
        rowIndex = g_PostProcessFooterRowIndex
        If rowIndex >= 1 And rowIndex <= ws.Rows.Count Then
            If mp_IsFooterAnchorRowValid(ws, rowIndex, probeEndCol) Then
                outRowIndex = rowIndex
                mp_TryGetCachedSinglePostProcessFooterRowIndex = True
                Exit Function
            End If
        End If
    End If

    If mp_TryGetPostProcessFooterAnchorRow(ws, rowIndex) Then
        If Not mp_IsFooterAnchorRowValid(ws, rowIndex, probeEndCol) Then
            mp_ClearPostProcessFooterAnchors ws
            Exit Function
        End If
        g_PostProcessFooterSheetKey = sheetKey
        g_PostProcessFooterRowIndex = rowIndex
        outRowIndex = rowIndex
        mp_TryGetCachedSinglePostProcessFooterRowIndex = True
    End If
End Function

Private Function mp_GetOrCreateSinglePostProcessHeaderRange( _
    ByVal ws As Worksheet, _
    ByRef postProcessHeaderStyle As t_PostProcessHeaderStyle _
) As Range
    Dim sheetKey As String
    Dim viewStartRow As Long
    Dim targetRow As Long
    Dim existingRow As Long
    Dim insertRow As Long
    Dim endCol As Long
    Dim headerRange As Range

    If ws Is Nothing Then Exit Function

    sheetKey = mp_BuildSheetKey(ws)
    If StrComp(g_PostProcessHeaderSheetKey, sheetKey, vbTextCompare) <> 0 Then
        g_PostProcessHeaderSheetKey = sheetKey
        g_PostProcessHeaderNextInsertRow = 0
        g_PostProcessHeaderRowIndex = 0
        g_PostProcessHeaderHasAppended = False
    End If
    viewStartRow = mp_GetPostProcessHeaderInsertStartRow(ws)

    endCol = postProcessHeaderStyle.Columns
    If endCol < 1 Then endCol = 1
    If endCol > ws.Columns.Count Then endCol = ws.Columns.Count

    If g_PostProcessHeaderRowIndex > 0 Then
        targetRow = g_PostProcessHeaderRowIndex
        If targetRow > ws.Rows.Count Then targetRow = ws.Rows.Count
    ElseIf mp_TryGetPostProcessHeaderAnchorRow(ws, existingRow) Then
        targetRow = existingRow
    End If
    If targetRow >= viewStartRow Then
        If Not mp_IsHeaderAnchorRowValid(ws, targetRow, endCol) Then
            targetRow = 0
        End If
    End If
    If targetRow > (viewStartRow + POST_PROCESS_HEADER_ANCHOR_MAX_OFFSET_ROWS) Then targetRow = 0
    If targetRow < viewStartRow Then
        insertRow = mp_GetPostProcessHeaderInsertStartRow(ws)
        If insertRow < 1 Then insertRow = 1
        If insertRow > ws.Rows.Count Then insertRow = ws.Rows.Count
        If insertRow > ws.Rows.Count - 2 Then insertRow = ws.Rows.Count - 2
        If insertRow < 1 Then insertRow = 1

        ws.Rows(CStr(insertRow) & ":" & CStr(insertRow + 2)).Insert Shift:=xlDown
        targetRow = insertRow + 1
    End If

    If targetRow < 1 Then targetRow = 1
    If targetRow > ws.Rows.Count Then targetRow = ws.Rows.Count

    g_PostProcessHeaderRowIndex = targetRow
    g_PostProcessHeaderNextInsertRow = targetRow + 2
    If g_PostProcessHeaderNextInsertRow > ws.Rows.Count Then g_PostProcessHeaderNextInsertRow = ws.Rows.Count
    mp_SetPostProcessHeaderAnchors ws, targetRow

    Set headerRange = ws.Range(ws.Cells(targetRow, 1), ws.Cells(targetRow, endCol))
    If headerRange.MergeCells Then headerRange.UnMerge
    mp_ApplyPostProcessHeaderKindStyle ws, targetRow

    Set mp_GetOrCreateSinglePostProcessHeaderRange = headerRange
End Function

Private Function mp_GetOrCreateSinglePostProcessFooterRange( _
    ByVal ws As Worksheet, _
    ByRef postProcessFooterStyle As t_PostProcessFooterStyle _
) As Range
    Dim sheetKey As String
    Dim targetRow As Long
    Dim endCol As Long
    Dim footerRange As Range

    If ws Is Nothing Then Exit Function

    sheetKey = mp_BuildSheetKey(ws)
    If StrComp(g_PostProcessFooterSheetKey, sheetKey, vbTextCompare) <> 0 Then
        g_PostProcessFooterSheetKey = sheetKey
        g_PostProcessFooterRowIndex = 0
        g_PostProcessFooterHasAppended = False
    End If

    endCol = postProcessFooterStyle.Columns
    If endCol < 1 Then endCol = 1
    If endCol > ws.Columns.Count Then endCol = ws.Columns.Count

    If g_PostProcessFooterRowIndex > 0 Then
        targetRow = g_PostProcessFooterRowIndex
        If targetRow > ws.Rows.Count Then targetRow = ws.Rows.Count
    ElseIf mp_TryGetPostProcessFooterAnchorRow(ws, targetRow) Then
        ' anchor resolved
    Else
        targetRow = mp_GetLastUsedRow(ws) + 2
        If targetRow < 1 Then targetRow = 1
    End If
    If targetRow > 0 Then
        If Not mp_IsFooterAnchorRowValid(ws, targetRow, endCol) Then
            targetRow = mp_GetLastUsedRow(ws) + 2
            If targetRow < 1 Then targetRow = 1
        End If
    End If

    If targetRow < 1 Then targetRow = 1
    If targetRow > ws.Rows.Count Then targetRow = ws.Rows.Count
    g_PostProcessFooterRowIndex = targetRow
    mp_SetPostProcessFooterAnchors ws, targetRow

    Set footerRange = ws.Range(ws.Cells(targetRow, 1), ws.Cells(targetRow, endCol))
    If footerRange.MergeCells Then footerRange.UnMerge
    mp_ApplyPostProcessFooterKindStyle ws, targetRow

    Set mp_GetOrCreateSinglePostProcessFooterRange = footerRange
End Function

Private Sub mp_GetWarningBannerDimensions( _
    ByRef outColumns As Long, _
    ByRef outRows As Long, _
    Optional ByVal bannerKind As String = "warningbanner" _
)
    Dim bannerStyle As ex_SheetStylesXmlProvider.t_ErrorBannerStyle
    Dim normalizedKind As String

    normalizedKind = LCase$(Trim$(bannerKind))
    If normalizedKind = "errorbanner" Then
        If ex_SheetStylesXmlProvider.m_GetErrorBannerStyle(bannerStyle, ThisWorkbook) Then
            outColumns = bannerStyle.Columns
            outRows = bannerStyle.Rows
        ElseIf ex_SheetStylesXmlProvider.m_GetWarningBannerStyle(bannerStyle, ThisWorkbook) Then
            outColumns = bannerStyle.Columns
            outRows = bannerStyle.Rows
        End If
    Else
        If ex_SheetStylesXmlProvider.m_GetWarningBannerStyle(bannerStyle, ThisWorkbook) Then
            outColumns = bannerStyle.Columns
            outRows = bannerStyle.Rows
        ElseIf ex_SheetStylesXmlProvider.m_GetErrorBannerStyle(bannerStyle, ThisWorkbook) Then
            outColumns = bannerStyle.Columns
            outRows = bannerStyle.Rows
        End If
    End If

    If outColumns < 1 Then outColumns = 8
    ' Banner layout is single-row (title + body in one cell with blank line).
    outRows = 1
End Sub

Private Function mp_ToColumnLetter(ByVal columnIndex As Long) As String
    Dim n As Long
    Dim remainder As Long

    If columnIndex < 1 Then columnIndex = 1
    n = columnIndex

    Do While n > 0
        remainder = (n - 1) Mod 26
        mp_ToColumnLetter = Chr$(65 + remainder) & mp_ToColumnLetter
        n = (n - 1) \ 26
    Loop
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

Private Function mp_GetTargetCellForRowRef( _
    ByVal rowRef As obj_ResultRow, _
    ByVal columnRef As String _
) As Range
    Dim targetCol As Long
    Dim ws As Worksheet
    Dim rowIndex As Long

    If rowRef Is Nothing Then
        Err.Raise vbObjectError + 1679, "ex_PostProcessActions", "Row reference is required for row cell write."
    End If

    columnRef = Trim$(columnRef)
    If Len(columnRef) = 0 Then
        Err.Raise vbObjectError + 1680, "ex_PostProcessActions", "Column reference is empty for row cell write."
    End If

    If Not mp_TryResolveColumnIndexInRow(rowRef, columnRef, targetCol) Then
        Err.Raise vbObjectError + 1681, "ex_PostProcessActions", "Unknown row cell reference '" & columnRef & "' for row cell write."
    End If

    Set ws = ActiveSheet
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1747, "ex_PostProcessActions", "Active sheet is not available for row cell write."
    End If

    rowIndex = mp_ResolveAnchoredRowIndex(ws, rowRef, "row cell write")
    Set mp_GetTargetCellForRowRef = mp_GetRowCellRange(rowIndex, targetCol)
End Function

Private Function mp_ResolveAnchoredRowIndex( _
    ByVal ws As Worksheet, _
    ByVal rowRef As obj_ResultRow, _
    ByVal operationName As String _
) As Long
    Dim anchorName As String
    Dim resolvedRowIndex As Long

    If ws Is Nothing Then
        Err.Raise vbObjectError + 1748, "ex_PostProcessActions", "Active sheet is not available for " & operationName & "."
    End If
    If rowRef Is Nothing Then
        Err.Raise vbObjectError + 1749, "ex_PostProcessActions", "Row reference is required for " & operationName & "."
    End If

    anchorName = Trim$(rowRef.RowAnchorName)
    If Len(anchorName) = 0 Then
        Err.Raise vbObjectError + 1750, "ex_PostProcessActions", "Row anchor is not defined for " & operationName & "."
    End If
    If Not mp_TryGetNamedRowAnchor(ws, anchorName, resolvedRowIndex) Then
        Err.Raise vbObjectError + 1751, "ex_PostProcessActions", "Row anchor '" & anchorName & "' is not found for " & operationName & "."
    End If

    mp_ResolveAnchoredRowIndex = resolvedRowIndex
End Function

Private Sub mp_AssignRuntimeRowAnchor( _
    ByVal ws As Worksheet, _
    ByVal rowRef As obj_ResultRow, _
    ByVal rowIndex As Long _
)
    Dim anchorName As String

    If ws Is Nothing Then Exit Sub
    If rowRef Is Nothing Then Exit Sub
    If rowIndex < 1 Then Exit Sub
    If rowIndex > ws.Rows.Count Then Exit Sub

    anchorName = mp_NextRuntimeRowAnchorName(ws)
    If Len(anchorName) = 0 Then
        Err.Raise vbObjectError + 1752, "ex_PostProcessActions", "Unable to allocate runtime row anchor."
    End If
    mp_SetNamedRowAnchor ws, anchorName, rowIndex
    rowRef.RowAnchorName = anchorName
End Sub
