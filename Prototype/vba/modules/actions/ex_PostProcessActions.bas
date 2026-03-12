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
Private g_RuntimeDataBySheetAndKey As Object

Public Sub m_HighlightRow( _
    ByVal rowRef As obj_ResultRow, _
    Optional ByVal colorHex As String = "#FFF2CC" _
)
    Dim colorValue As Long
    Dim rowRange As Range
    Dim ws As Worksheet
    Dim usedCols As Long

    If rowRef Is Nothing Then Exit Sub
    If Len(Trim$(colorHex)) = 0 Then colorHex = "#FFF2CC"
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    usedCols = mp_GetLastUsedColumn(ws)
    If usedCols <= 0 Then usedCols = 1

    If Not ex_XmlCore.m_TryParseColor(colorHex, colorValue) Then
        Err.Raise vbObjectError + 1650, "ex_PostProcessActions", "Invalid highlight color: " & colorHex
    End If

    Set rowRange = ws.Range(ws.Cells(rowRef.RowIndex, 1), ws.Cells(rowRef.RowIndex, usedCols))
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

    Set targetCell = ws.Cells(rowRef.RowIndex, targetCol)
    targetCell.Interior.Pattern = xlSolid
    targetCell.Interior.Color = colorValue
End Sub

Public Function m_RegexIsMatch( _
    ByVal textValue As String, _
    ByVal regexPattern As String _
) As Boolean
    Dim rx As Object

    Set rx = mp_CreateRegex(regexPattern)
    m_RegexIsMatch = rx.Test(CStr(textValue))
End Function

Public Function m_RegexFirstMatch( _
    ByVal textValue As String, _
    ByVal regexPattern As String _
) As String
    Dim rx As Object
    Dim matches As Object

    Set rx = mp_CreateRegex(regexPattern)
    Set matches = rx.Execute(CStr(textValue))
    If matches.Count > 0 Then
        m_RegexFirstMatch = CStr(matches(0).Value)
    End If
End Function

Public Function m_RowToText( _
    ByVal rowRef As obj_ResultRow, _
    ByVal separatorText As String _
) As String
    m_RowToText = mp_GetRowText(rowRef, separatorText)
End Function

Public Function m_RowCellRegexIsMatch( _
    ByVal rowRef As obj_ResultRow, _
    ByVal columnRef As String, _
    ByVal regexPattern As String _
) As Boolean
    m_RowCellRegexIsMatch = m_RegexIsMatch(mp_GetRowCellLiveText(rowRef, columnRef), regexPattern)
End Function

Public Function m_RowCellRegexFirstMatch( _
    ByVal rowRef As obj_ResultRow, _
    ByVal columnRef As String, _
    ByVal regexPattern As String _
) As String
    m_RowCellRegexFirstMatch = m_RegexFirstMatch(mp_GetRowCellLiveText(rowRef, columnRef), regexPattern)
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

    For probeRow = rowRef.RowIndex To 1 Step -1
        If Len(Trim$(CStr(ws.Cells(probeRow, ownerCol).Value))) > 0 Then
            ownerRowIndex = probeRow
            Exit For
        End If
    Next probeRow

    If ownerRowIndex <= 0 Then
        Err.Raise vbObjectError + 1678, "ex_PostProcessActions", "Unable to resolve owner row by column '" & ownerColumnRef & "' from row " & CStr(rowRef.RowIndex) & "."
    End If

    Set targetCell = ws.Cells(ownerRowIndex, targetCol)
    currentText = CStr(targetCell.Value)
    targetCell.Value = m_TextAppend(currentText, CStr(appendText), separatorText)
End Sub

Public Sub m_EmphasizeRowCellTextByRegex( _
    ByVal rowRef As obj_ResultRow, _
    ByVal columnRef As String, _
    ByVal regexPattern As String, _
    Optional ByVal fontColorHex As String = "#FF0000", _
    Optional ByVal uppercaseMatches As String = "false" _
)
    Dim targetCell As Range
    Dim targetCol As Long
    Dim originalText As String
    Dim transformedText As String
    Dim rx As Object
    Dim matches As Object
    Dim i As Long
    Dim matchObj As Object
    Dim colorValue As Long
    Dim makeUpper As Boolean

    If rowRef Is Nothing Then
        Err.Raise vbObjectError + 1664, "ex_PostProcessActions", "Row reference is required for regex text emphasis."
    End If

    If Not mp_TryResolveColumnIndexInRow(rowRef, columnRef, targetCol) Then
        Err.Raise vbObjectError + 1665, "ex_PostProcessActions", "Unknown row cell reference '" & columnRef & "' for regex text emphasis."
    End If
    Set targetCell = mp_GetRowCellRange(rowRef.RowIndex, targetCol)
    originalText = CStr(targetCell.Value)

    If Len(Trim$(fontColorHex)) = 0 Then fontColorHex = "#FF0000"
    If Not ex_XmlCore.m_TryParseColor(fontColorHex, colorValue) Then
        Err.Raise vbObjectError + 1666, "ex_PostProcessActions", "Invalid regex emphasis color: " & fontColorHex
    End If
    makeUpper = mp_ParseRequiredBoolean(uppercaseMatches, "uppercaseMatches")

    Set rx = mp_CreateRegex(regexPattern, True)
    Set matches = rx.Execute(originalText)
    If matches Is Nothing Or matches.Count = 0 Then Exit Sub

    If makeUpper Then
        transformedText = originalText
        For i = 0 To matches.Count - 1
            Set matchObj = matches(i)
            If matchObj.Length > 0 Then
                transformedText = Left$(transformedText, matchObj.FirstIndex) & UCase$(Mid$(transformedText, matchObj.FirstIndex + 1, matchObj.Length)) & Mid$(transformedText, matchObj.FirstIndex + matchObj.Length + 1)
            End If
        Next i
        targetCell.Value = transformedText
    End If

    For i = 0 To matches.Count - 1
        Set matchObj = matches(i)
        If matchObj.Length > 0 Then
            targetCell.Characters(matchObj.FirstIndex + 1, matchObj.Length).Font.Color = colorValue
            targetCell.Characters(matchObj.FirstIndex + 1, matchObj.Length).Font.Bold = True
        End If
    Next i
End Sub

Public Sub m_AddNote( _
    ByVal rowRef As obj_ResultRow, _
    ByVal noteText As String _
)
    Dim noteCell As Range
    Dim ws As Worksheet

    If rowRef Is Nothing Then Exit Sub
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    Set noteCell = ws.Cells(rowRef.RowIndex, 1)
    On Error Resume Next
    If Not noteCell.Comment Is Nothing Then noteCell.Comment.Delete
    On Error GoTo 0
    noteCell.AddComment noteText
End Sub

Public Sub m_ResetPostProcessHeaderCursor(Optional ByVal targetSheet As Worksheet)
    g_PostProcessHeaderNextInsertRow = 0
    g_PostProcessHeaderRowIndex = 0
    g_PostProcessHeaderHasAppended = False
    If targetSheet Is Nothing Then
        g_PostProcessHeaderSheetKey = vbNullString
        m_ClearRuntimeData
    Else
        g_PostProcessHeaderSheetKey = mp_BuildSheetKey(targetSheet)
        mp_ClearRuntimeDataForSheet targetSheet
    End If
End Sub

Public Sub m_ResetPostProcessFooterCursor(Optional ByVal targetSheet As Worksheet)
    g_PostProcessFooterRowIndex = 0
    g_PostProcessFooterHasAppended = False
    If targetSheet Is Nothing Then
        g_PostProcessFooterSheetKey = vbNullString
        m_ClearRuntimeData
    Else
        g_PostProcessFooterSheetKey = mp_BuildSheetKey(targetSheet)
        mp_ClearRuntimeDataForSheet targetSheet
    End If
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

    If Not mp_TryLoadPostProcessHeaderStyle(postProcessHeaderStyle) Then
        Err.Raise vbObjectError + 1673, "ex_PostProcessActions", "Unable to apply postProcessHeader text: invalid '/sheetStyles/postProcessHeaderStyle'."
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

    sheetKey = mp_BuildSheetKey(ws)
    If StrComp(g_PostProcessHeaderSheetKey, sheetKey, vbTextCompare) <> 0 Then
        g_PostProcessHeaderSheetKey = sheetKey
        g_PostProcessHeaderNextInsertRow = 0
        g_PostProcessHeaderRowIndex = 0
        g_PostProcessHeaderHasAppended = False
    End If

    If Not mp_TryLoadPostProcessHeaderStyle(postProcessHeaderStyle) Then
        Err.Raise vbObjectError + 1728, "ex_PostProcessActions", "Unable to apply single postProcessHeader text: invalid '/sheetStyles/postProcessHeaderStyle'."
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

Public Function m_GetSinglePostProcessHeaderText(Optional ByVal targetSheet As Worksheet = Nothing) As String
    Dim ws As Worksheet
    Dim headerRowIndex As Long

    If targetSheet Is Nothing Then
        Set ws = ActiveSheet
    Else
        Set ws = targetSheet
    End If
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1729, "ex_PostProcessActions", "Active sheet is not available for postProcessHeader read."
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

    If Not mp_TryLoadPostProcessFooterStyle(postProcessFooterStyle) Then
        Err.Raise vbObjectError + 1651, "ex_PostProcessActions", "Unable to apply postProcessFooter text: invalid '/sheetStyles/postProcessFooterStyle'."
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

    If Not mp_TryLoadPostProcessFooterStyle(postProcessFooterStyle) Then
        Err.Raise vbObjectError + 1682, "ex_PostProcessActions", "Unable to apply single postProcessFooter text: invalid '/sheetStyles/postProcessFooterStyle'."
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

    If targetSheet Is Nothing Then
        Set ws = ActiveSheet
    Else
        Set ws = targetSheet
    End If
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1701, "ex_PostProcessActions", "Active sheet is not available for postProcessFooter read."
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

Public Function m_ShowWarningBanner( _
    ByVal warningText As String, _
    Optional ByVal titleText As String = "WARNING", _
    Optional ByVal bannerRangeAddress As String = vbNullString _
) As String
    Dim ws As Worksheet

    warningText = Trim$(warningText)
    If Len(warningText) = 0 Then Exit Function

    Set ws = ActiveSheet
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1713, "ex_PostProcessActions", "Active sheet is not available for warning banner."
    End If

    ex_Messaging.m_RenderWarningBanner ws, warningText, titleText, bannerRangeAddress
    m_ShowWarningBanner = warningText
End Function

Public Function m_GetRowIndex(ByVal rowRef As Object) As String
    Dim sourceRowRef As obj_ResultRow

    If rowRef Is Nothing Then
        Err.Raise vbObjectError + 1718, "ex_PostProcessActions", "Row reference is required for row index read."
    End If
    If Not TypeOf rowRef Is obj_ResultRow Then
        Err.Raise vbObjectError + 1719, "ex_PostProcessActions", "Row reference must be obj_ResultRow for row index read."
    End If
    Set sourceRowRef = rowRef

    m_GetRowIndex = CStr(sourceRowRef.RowIndex)
End Function

Public Function m_GetWarningBannerRangeAboveRow( _
    ByVal rowRef As Object, _
    Optional ByVal gapRowsText As String = "1" _
) As String
    Dim sourceRowRef As obj_ResultRow
    Dim gapRows As Long
    Dim bannerCols As Long
    Dim bannerRows As Long
    Dim startRow As Long

    If rowRef Is Nothing Then
        Err.Raise vbObjectError + 1714, "ex_PostProcessActions", "Row reference is required for warning banner range."
    End If
    If Not TypeOf rowRef Is obj_ResultRow Then
        Err.Raise vbObjectError + 1715, "ex_PostProcessActions", "Row reference must be obj_ResultRow for warning banner range."
    End If
    Set sourceRowRef = rowRef

    gapRowsText = Trim$(gapRowsText)
    If Len(gapRowsText) = 0 Then gapRowsText = "1"
    If Not ex_XmlCore.m_TryParseLong(gapRowsText, gapRows) Then
        Err.Raise vbObjectError + 1716, "ex_PostProcessActions", "Gap rows must be integer for warning banner range."
    End If
    If gapRows < 0 Then
        Err.Raise vbObjectError + 1717, "ex_PostProcessActions", "Gap rows cannot be negative for warning banner range."
    End If

    mp_GetWarningBannerDimensions bannerCols, bannerRows
    startRow = sourceRowRef.RowIndex - gapRows - bannerRows
    If startRow < 1 Then startRow = 1

    m_GetWarningBannerRangeAboveRow = "A" & CStr(startRow) & ":" & mp_ToColumnLetter(bannerCols) & CStr(startRow + bannerRows - 1)
End Function

Public Function m_ShowWarningBannerBeforeRowIndex( _
    ByVal warningText As String, _
    Optional ByVal titleText As String = "WARNING", _
    Optional ByVal rowIndexText As String = vbNullString, _
    Optional ByVal gapRowsText As String = "0" _
) As String
    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim gapRows As Long
    Dim bannerCols As Long
    Dim bannerRows As Long
    Dim rowsToInsert As Long
    Dim bannerRangeAddress As String

    warningText = Trim$(warningText)
    If Len(warningText) = 0 Then Exit Function

    rowIndexText = Trim$(rowIndexText)
    If Not ex_XmlCore.m_TryParseLong(rowIndexText, rowIndex) Then
        Err.Raise vbObjectError + 1720, "ex_PostProcessActions", "Row index must be integer for warning banner insert."
    End If
    If rowIndex < 1 Then
        Err.Raise vbObjectError + 1721, "ex_PostProcessActions", "Row index must be >= 1 for warning banner insert."
    End If

    gapRowsText = Trim$(gapRowsText)
    If Len(gapRowsText) = 0 Then gapRowsText = "0"
    If Not ex_XmlCore.m_TryParseLong(gapRowsText, gapRows) Then
        Err.Raise vbObjectError + 1722, "ex_PostProcessActions", "Gap rows must be integer for warning banner insert."
    End If
    If gapRows < 0 Then
        Err.Raise vbObjectError + 1723, "ex_PostProcessActions", "Gap rows cannot be negative for warning banner insert."
    End If

    Set ws = ActiveSheet
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1724, "ex_PostProcessActions", "Active sheet is not available for warning banner insert."
    End If

    mp_GetWarningBannerDimensions bannerCols, bannerRows
    rowsToInsert = bannerRows + gapRows
    If rowsToInsert > 0 Then
        ws.Rows(CStr(rowIndex) & ":" & CStr(rowIndex + rowsToInsert - 1)).Insert Shift:=xlDown
    End If

    bannerRangeAddress = "A" & CStr(rowIndex) & ":" & mp_ToColumnLetter(bannerCols) & CStr(rowIndex + bannerRows - 1)
    ex_Messaging.m_RenderWarningBanner ws, warningText, titleText, bannerRangeAddress
    m_ShowWarningBannerBeforeRowIndex = warningText
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

    targetRow = sourceRowRef.RowIndex + rowOffset
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

    targetRow = sourceRowRef.RowIndex + rowOffset
    hasTargetRow = (targetRow >= 1 And targetRow <= ws.Rows.Count)

    Set resultRow = New obj_ResultRow
    If hasTargetRow Then
        resultRow.Initialize targetRow
    Else
        resultRow.Initialize sourceRowRef.RowIndex
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
    Dim targetRow As Long
    Dim measuredHeight As Double
    Dim standardRowHeight As Double

    If ws Is Nothing Then Exit Sub
    If postProcessHeaderRange Is Nothing Then Exit Sub

    targetRow = postProcessHeaderRange.Row
    If targetRow <= 0 Then Exit Sub

    If postProcessHeaderStyle.RowHeight > 0 Then
        ws.Rows(targetRow).RowHeight = postProcessHeaderStyle.RowHeight
        Exit Sub
    End If

    standardRowHeight = ws.StandardHeight
    If standardRowHeight <= 0 Then standardRowHeight = 15

    If Not postProcessHeaderStyle.AutoHeight Or StrComp(postProcessHeaderStyle.Overflow, "wrap", vbTextCompare) <> 0 Then
        If postProcessHeaderStyle.MinRowHeight > 0 Then
            ws.Rows(targetRow).RowHeight = postProcessHeaderStyle.MinRowHeight
        Else
            ws.Rows(targetRow).RowHeight = standardRowHeight
        End If
        Exit Sub
    End If

    measuredHeight = mp_MeasurePostProcessHeaderTextHeight(ws, postProcessHeaderRange, postProcessHeaderText, postProcessHeaderStyle.FontSize)
    measuredHeight = measuredHeight + postProcessHeaderStyle.AutoHeightMarginTop + postProcessHeaderStyle.AutoHeightMarginBottom
    If measuredHeight < postProcessHeaderStyle.MinRowHeight Then measuredHeight = postProcessHeaderStyle.MinRowHeight

    If measuredHeight <= 0 Then
        ws.Rows(targetRow).RowHeight = standardRowHeight
    Else
        ws.Rows(targetRow).RowHeight = mp_RoundUpMeasuredHeight(measuredHeight)
    End If
End Sub

Private Sub mp_ApplyPostProcessFooterRowHeight( _
    ByVal ws As Worksheet, _
    ByVal postProcessFooterRange As Range, _
    ByVal postProcessFooterText As String, _
    ByRef postProcessFooterStyle As t_PostProcessFooterStyle _
)
    Dim targetRow As Long
    Dim measuredHeight As Double
    Dim standardRowHeight As Double

    If ws Is Nothing Then Exit Sub
    If postProcessFooterRange Is Nothing Then Exit Sub

    targetRow = postProcessFooterRange.Row
    If targetRow <= 0 Then Exit Sub

    If postProcessFooterStyle.RowHeight > 0 Then
        ws.Rows(targetRow).RowHeight = postProcessFooterStyle.RowHeight
        Exit Sub
    End If

    standardRowHeight = ws.StandardHeight
    If standardRowHeight <= 0 Then standardRowHeight = 15

    If Not postProcessFooterStyle.AutoHeight Or StrComp(postProcessFooterStyle.Overflow, "wrap", vbTextCompare) <> 0 Then
        If postProcessFooterStyle.MinRowHeight > 0 Then
            ws.Rows(targetRow).RowHeight = postProcessFooterStyle.MinRowHeight
        Else
            ws.Rows(targetRow).RowHeight = standardRowHeight
        End If
        Exit Sub
    End If

    measuredHeight = mp_MeasurePostProcessFooterTextHeight(ws, postProcessFooterRange, postProcessFooterText, postProcessFooterStyle.FontSize)
    measuredHeight = measuredHeight + postProcessFooterStyle.AutoHeightMarginTop + postProcessFooterStyle.AutoHeightMarginBottom
    If measuredHeight < postProcessFooterStyle.MinRowHeight Then measuredHeight = postProcessFooterStyle.MinRowHeight

    If measuredHeight <= 0 Then
        ws.Rows(targetRow).RowHeight = standardRowHeight
    Else
        ws.Rows(targetRow).RowHeight = mp_RoundUpMeasuredHeight(measuredHeight)
    End If
End Sub

Private Function mp_MeasurePostProcessHeaderTextHeight( _
    ByVal ws As Worksheet, _
    ByVal postProcessHeaderRange As Range, _
    ByVal postProcessHeaderText As String, _
    ByVal fontSize As Double _
) As Double
    Dim textBoxShape As Object

    On Error GoTo EH
    If ws Is Nothing Then Exit Function
    If postProcessHeaderRange Is Nothing Then Exit Function
    If Len(postProcessHeaderText) = 0 Then Exit Function

    Set textBoxShape = ws.Shapes.AddTextbox(1, postProcessHeaderRange.Left, postProcessHeaderRange.Top, postProcessHeaderRange.Width, 8)
    textBoxShape.Line.Visible = 0
    textBoxShape.Fill.Visible = 0
    textBoxShape.TextFrame2.MarginLeft = POST_PROCESS_MEASURE_SIDE_MARGIN
    textBoxShape.TextFrame2.MarginRight = POST_PROCESS_MEASURE_SIDE_MARGIN
    textBoxShape.TextFrame2.MarginTop = POST_PROCESS_MEASURE_VERTICAL_MARGIN
    textBoxShape.TextFrame2.MarginBottom = POST_PROCESS_MEASURE_VERTICAL_MARGIN
    textBoxShape.TextFrame2.WordWrap = -1
    textBoxShape.TextFrame2.AutoSize = 1
    textBoxShape.TextFrame2.TextRange.Text = postProcessHeaderText
    textBoxShape.TextFrame2.TextRange.Font.Size = fontSize
    textBoxShape.TextFrame2.TextRange.Font.Name = CStr(postProcessHeaderRange.Font.Name)

    mp_MeasurePostProcessHeaderTextHeight = textBoxShape.Height + mp_GetMeasureExtraHeight(fontSize)

Cleanup:
    On Error Resume Next
    If Not textBoxShape Is Nothing Then textBoxShape.Delete
    On Error GoTo 0
    Exit Function

EH:
    mp_MeasurePostProcessHeaderTextHeight = 0
    Resume Cleanup
End Function

Private Function mp_MeasurePostProcessFooterTextHeight( _
    ByVal ws As Worksheet, _
    ByVal postProcessFooterRange As Range, _
    ByVal postProcessFooterText As String, _
    ByVal fontSize As Double _
) As Double
    Dim textBoxShape As Object

    On Error GoTo EH
    If ws Is Nothing Then Exit Function
    If postProcessFooterRange Is Nothing Then Exit Function
    If Len(postProcessFooterText) = 0 Then Exit Function

    Set textBoxShape = ws.Shapes.AddTextbox(1, postProcessFooterRange.Left, postProcessFooterRange.Top, postProcessFooterRange.Width, 8)
    textBoxShape.Line.Visible = 0
    textBoxShape.Fill.Visible = 0
    textBoxShape.TextFrame2.MarginLeft = POST_PROCESS_MEASURE_SIDE_MARGIN
    textBoxShape.TextFrame2.MarginRight = POST_PROCESS_MEASURE_SIDE_MARGIN
    textBoxShape.TextFrame2.MarginTop = POST_PROCESS_MEASURE_VERTICAL_MARGIN
    textBoxShape.TextFrame2.MarginBottom = POST_PROCESS_MEASURE_VERTICAL_MARGIN
    textBoxShape.TextFrame2.WordWrap = -1
    textBoxShape.TextFrame2.AutoSize = 1
    textBoxShape.TextFrame2.TextRange.Text = postProcessFooterText
    textBoxShape.TextFrame2.TextRange.Font.Size = fontSize
    textBoxShape.TextFrame2.TextRange.Font.Name = CStr(postProcessFooterRange.Font.Name)

    mp_MeasurePostProcessFooterTextHeight = textBoxShape.Height + mp_GetMeasureExtraHeight(fontSize)

Cleanup:
    On Error Resume Next
    If Not textBoxShape Is Nothing Then textBoxShape.Delete
    On Error GoTo 0
    Exit Function

EH:
    mp_MeasurePostProcessFooterTextHeight = 0
    Resume Cleanup
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
    If Not mp_IsRowBlankSegment(ws, rowIndex - 1, endCol) Then Exit Function
    If Not mp_IsRowBlankSegment(ws, rowIndex + 1, endCol) Then Exit Function
    mp_IsHeaderAnchorRowValid = True
End Function

Private Function mp_IsFooterAnchorRowValid( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal endCol As Long _
) As Boolean
    If Not mp_IsSinglePostProcessTextRowShape(ws, rowIndex, endCol) Then Exit Function
    If Not mp_IsRowBlankSegment(ws, rowIndex - 1, endCol) Then Exit Function
    If Not mp_IsRowBlankSegment(ws, rowIndex + 1, endCol) Then Exit Function
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
    Dim roundedHeight As Double

    If measuredHeight <= 0 Then Exit Function
    roundedHeight = Int(measuredHeight + 0.999)
    mp_RoundUpMeasuredHeight = roundedHeight + POST_PROCESS_MEASURE_HEIGHT_ROUND_PAD
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

    If rowRef Is Nothing Then
        Err.Raise vbObjectError + 1671, "ex_PostProcessActions", "Row reference is required for live cell text parsing."
    End If
    If Not mp_TryResolveColumnIndexInRow(rowRef, columnRef, targetCol) Then
        Err.Raise vbObjectError + 1672, "ex_PostProcessActions", "Unknown row cell reference '" & columnRef & "' for live cell text parsing."
    End If

    Set targetCell = mp_GetRowCellRange(rowRef.RowIndex, targetCol)
    mp_GetRowCellLiveText = CStr(targetCell.Value)
End Function

Private Function mp_CreateRegex( _
    ByVal regexPattern As String, _
    Optional ByVal globalMatches As Boolean = False _
) As Object
    Dim rx As Object

    regexPattern = Trim$(regexPattern)
    If Len(regexPattern) = 0 Then
        Err.Raise vbObjectError + 1661, "ex_PostProcessActions", "Regex pattern is empty."
    End If

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = globalMatches
    rx.IgnoreCase = True
    rx.MultiLine = True

    On Error GoTo PatternErr
    rx.Pattern = regexPattern
    On Error GoTo 0

    Set mp_CreateRegex = rx
    Exit Function

PatternErr:
    Err.Raise vbObjectError + 1662, "ex_PostProcessActions", "Invalid regex pattern '" & regexPattern & "': " & Err.Description
End Function

Private Function mp_ParseRequiredBoolean(ByVal valueText As String, ByVal fieldName As String) As Boolean
    Dim parsedValue As Boolean

    valueText = Trim$(valueText)
    If Not ex_XmlCore.m_TryParseBoolean(valueText, parsedValue) Then
        Err.Raise vbObjectError + 1667, "ex_PostProcessActions", "Invalid boolean for '" & fieldName & "': '" & valueText & "'."
    End If

    mp_ParseRequiredBoolean = parsedValue
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
    On Error GoTo 0
End Sub

Private Sub mp_SetNamedRowAnchor( _
    ByVal ws As Worksheet, _
    ByVal anchorName As String, _
    ByVal rowIndex As Long _
)
    If ws Is Nothing Then Exit Sub
    anchorName = Trim$(anchorName)
    If Len(anchorName) = 0 Then Exit Sub
    If rowIndex < 1 Or rowIndex > ws.Rows.Count Then Exit Sub

    mp_ClearNamedRowAnchor ws, anchorName

    On Error Resume Next
    ws.Names.Add Name:=anchorName, RefersTo:="=" & ws.Cells(rowIndex, 1).Address(True, True, xlA1, True)
    On Error GoTo 0
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

    If mp_TryGetNamedRowAnchor(ws, POST_PROCESS_HEADER_ANCHOR_NAME, rowIndex) Then
        If rowIndex < viewStartRow Then Exit Function
        If rowIndex > (viewStartRow + POST_PROCESS_HEADER_ANCHOR_MAX_OFFSET_ROWS) Then Exit Function
        If Not mp_IsHeaderAnchorRowValid(ws, rowIndex, probeEndCol) Then
            mp_ClearNamedRowAnchor ws, POST_PROCESS_HEADER_ANCHOR_NAME
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

    If mp_TryGetNamedRowAnchor(ws, POST_PROCESS_FOOTER_ANCHOR_NAME, rowIndex) Then
        If Not mp_IsFooterAnchorRowValid(ws, rowIndex, probeEndCol) Then
            mp_ClearNamedRowAnchor ws, POST_PROCESS_FOOTER_ANCHOR_NAME
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
    ElseIf mp_TryGetNamedRowAnchor(ws, POST_PROCESS_HEADER_ANCHOR_NAME, existingRow) Then
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
    mp_SetNamedRowAnchor ws, POST_PROCESS_HEADER_ANCHOR_NAME, targetRow

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
    ElseIf mp_TryGetNamedRowAnchor(ws, POST_PROCESS_FOOTER_ANCHOR_NAME, targetRow) Then
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
    mp_SetNamedRowAnchor ws, POST_PROCESS_FOOTER_ANCHOR_NAME, targetRow

    Set footerRange = ws.Range(ws.Cells(targetRow, 1), ws.Cells(targetRow, endCol))
    If footerRange.MergeCells Then footerRange.UnMerge
    mp_ApplyPostProcessFooterKindStyle ws, targetRow

    Set mp_GetOrCreateSinglePostProcessFooterRange = footerRange
End Function

Private Sub mp_GetWarningBannerDimensions(ByRef outColumns As Long, ByRef outRows As Long)
    Dim bannerStyle As ex_SheetStylesXmlProvider.t_ErrorBannerStyle

    If ex_SheetStylesXmlProvider.m_GetWarningBannerStyle(bannerStyle, ThisWorkbook) Then
        outColumns = bannerStyle.Columns
        outRows = bannerStyle.Rows
    ElseIf ex_SheetStylesXmlProvider.m_GetErrorBannerStyle(bannerStyle, ThisWorkbook) Then
        outColumns = bannerStyle.Columns
        outRows = bannerStyle.Rows
    End If

    If outColumns < 1 Then outColumns = 8
    If outRows < 1 Then outRows = 3
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

    Set mp_GetTargetCellForRowRef = mp_GetRowCellRange(rowRef.RowIndex, targetCol)
End Function
