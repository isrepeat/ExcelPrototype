Attribute VB_Name = "ex_ConfigTableStore"
Option Explicit

Private Const DEV_CONFIG_TABLE_NAME As String = "tblDevConfig"
Private Const DEV_CONFIG_HEADER_ROW As Long = 1
Private Const DEV_CONFIG_MARKER_COL As Long = 1
Private Const DEV_CONFIG_KEY_COL As Long = 2
Private Const DEV_CONFIG_VALUE_COL As Long = 3
Private Const DEV_CONFIG_STYLES_COL As Long = 4
Private Const DEV_CONFIG_COL_COUNT As Long = 4
Private Const DEV_HEADER_STYLES As String = "Styles"
Private Const DEV_MARKER_SYMBOL As String = "#"
Private Const DEV_MARKER_HEADER As String = ".."
Private Const DEV_MARKER_PREFIX As String = "#MARKER:"
Private Const DEV_MARKER_SECTION As String = "#MARKER:SECTION"
Private Const DEV_MARKER_SPACER As String = "#MARKER:SPACER"
Private Const MIN_MARKER_WIDTH_UNITS As Double = 4#
Private Const MIN_CONFIG_DATA_COL_WIDTH_POINTS As Double = 16#

Public Function m_GetConfigTable(ByVal ws As Worksheet, Optional ByVal createIfMissing As Boolean = False) As ListObject
    Dim tbl As ListObject

    On Error Resume Next
    Set tbl = ws.ListObjects(DEV_CONFIG_TABLE_NAME)
    On Error GoTo 0

    If Not tbl Is Nothing Then
        m_EnsureConfigTableLayout ws, tbl
    End If

    If tbl Is Nothing And createIfMissing Then
        Set tbl = m_CreateConfigTable(ws)
    End If

    Set m_GetConfigTable = tbl
End Function

Public Function m_GetTableDataRowCount(ByVal tbl As ListObject) As Long
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    m_GetTableDataRowCount = tbl.DataBodyRange.Rows.Count
End Function

Public Sub m_ResizeConfigTableRows(ByVal ws As Worksheet, ByVal tbl As ListObject, ByVal rowCount As Long)
    Dim topRow As Long
    Dim leftCol As Long
    Dim bottomRow As Long
    Dim rightCol As Long
    Dim resizeRange As Range

    If rowCount < 0 Then rowCount = 0

    topRow = tbl.HeaderRowRange.Row
    leftCol = tbl.Range.Column
    rightCol = leftCol + DEV_CONFIG_COL_COUNT - 1
    bottomRow = topRow + rowCount

    Set resizeRange = ws.Range(ws.Cells(topRow, leftCol), ws.Cells(bottomRow, rightCol))
    tbl.Resize resizeRange
End Sub

Public Sub m_ClearConfigDataArea(ByVal ws As Worksheet, ByVal tbl As ListObject)
    Dim rowCount As Long
    Dim topRow As Long
    Dim leftCol As Long
    Dim rightCol As Long
    Dim clearRange As Range

    rowCount = m_GetTableDataRowCount(tbl)
    If rowCount <= 0 Then Exit Sub

    topRow = tbl.HeaderRowRange.Row + 1
    leftCol = tbl.Range.Column
    rightCol = leftCol + DEV_CONFIG_COL_COUNT - 1

    Set clearRange = ws.Range(ws.Cells(topRow, leftCol), ws.Cells(topRow + rowCount - 1, rightCol))
    clearRange.Clear
End Sub

Public Sub m_AutoFitConfigColumnsWithinStableZone( _
    ByVal ws As Worksheet, _
    ByVal firstCol As Long, _
    ByVal colCount As Long, _
    Optional ByVal markerRelativeCol As Long = 1)

    Dim stableStartCol As Long
    Dim bufferCol As Long
    Dim markerAbsCol As Long
    Dim beforeWidth As Double
    Dim afterWidth As Double
    Dim bufferWidthBefore As Double
    Dim minBufferWidthUnits As Double
    Dim minBufferWidthPoints As Double
    Dim allowedGrowth As Double
    Dim maxWidth As Double

    If ws Is Nothing Then
        MsgBox "Worksheet is not specified for config column layout.", vbExclamation
        Exit Sub
    End If
    If firstCol < 1 Then
        MsgBox "Invalid config layout bounds: first column must be >= 1.", vbExclamation
        Exit Sub
    End If
    If colCount < 1 Then
        MsgBox "Invalid config layout bounds: column count must be >= 1.", vbExclamation
        Exit Sub
    End If

    markerAbsCol = firstCol + markerRelativeCol - 1
    If markerRelativeCol < 1 Or markerRelativeCol > colCount Then
        markerAbsCol = firstCol
    End If
    beforeWidth = mp_GetColumnsWidthPoints(ws, firstCol, colCount)

    If Not mp_TryGetStableZoneColumns(ws, stableStartCol, bufferCol) Then Exit Sub
    If bufferCol < firstCol + colCount - 1 Then
        MsgBox "Invalid UI layout: stable zone starts before the end of config columns.", vbExclamation
        Exit Sub
    End If

    bufferWidthBefore = ws.Columns(bufferCol).Width
    minBufferWidthUnits = mp_GetStableZoneMinBufferWidthUnits()
    If minBufferWidthUnits <= 0 Then Exit Sub

    minBufferWidthPoints = mp_GetColumnWidthPointsForUnits(ws.Columns(bufferCol), minBufferWidthUnits)
    If minBufferWidthPoints <= 0 Then
        MsgBox "Failed to evaluate minimum divider width for stable zone buffer column.", vbExclamation
        Exit Sub
    End If

    ws.Columns(firstCol).Resize(, colCount).AutoFit

    If ws.Columns(markerAbsCol).ColumnWidth < MIN_MARKER_WIDTH_UNITS Then
        ws.Columns(markerAbsCol).ColumnWidth = MIN_MARKER_WIDTH_UNITS
    End If

    allowedGrowth = bufferWidthBefore - minBufferWidthPoints
    If allowedGrowth < 0 Then allowedGrowth = 0
    maxWidth = beforeWidth + allowedGrowth

    afterWidth = mp_GetColumnsWidthPoints(ws, firstCol, colCount)
    If afterWidth <= maxWidth + 0.1 Then Exit Sub

    mp_ShrinkConfigColumnsToMaxWidth ws, firstCol, colCount, markerAbsCol, maxWidth

    If ws.Columns(markerAbsCol).ColumnWidth < MIN_MARKER_WIDTH_UNITS Then
        ws.Columns(markerAbsCol).ColumnWidth = MIN_MARKER_WIDTH_UNITS
    End If
End Sub

Public Function m_ScaleConfigColumnsToStableTarget( _
    ByVal ws As Worksheet, _
    ByVal firstCol As Long, _
    ByVal colCount As Long, _
    ByVal targetStableZoneLeft As Double _
) As Boolean

    Dim stableStartCol As Long
    Dim bufferCol As Long
    Dim currentStableLeft As Double
    Dim overflowPoints As Double
    Dim configStartLeft As Double
    Dim configEndCol As Long
    Dim gapStartCol As Long
    Dim gapEndCol As Long
    Dim gapWidth As Double
    Dim availableWidth As Double
    Dim requestedTotal As Double
    Dim scaleK As Double
    Dim colIndex As Long
    Dim srcWidths() As Double
    Dim targetPoints As Double

    If ws Is Nothing Then Exit Function
    If firstCol < 1 Or colCount < 1 Then Exit Function
    If targetStableZoneLeft < 0 Then Exit Function

    If Not mp_TryGetStableZoneColumns(ws, stableStartCol, bufferCol) Then Exit Function

    currentStableLeft = ws.Cells(1, stableStartCol).Left
    overflowPoints = currentStableLeft - targetStableZoneLeft
    If overflowPoints <= 0.1 Then Exit Function

    configStartLeft = ws.Cells(1, firstCol).Left
    configEndCol = firstCol + colCount - 1
    gapStartCol = configEndCol + 1
    gapEndCol = stableStartCol - 1
    If gapStartCol <= gapEndCol Then
        gapWidth = ws.Range(ws.Cells(1, gapStartCol), ws.Cells(1, gapEndCol)).Width
    End If

    ' Scale only config columns while preserving widths of gap/buffer columns before stable zone.
    availableWidth = targetStableZoneLeft - configStartLeft - gapWidth
    If availableWidth <= 0 Then Exit Function

    requestedTotal = mp_GetColumnsWidthPoints(ws, firstCol, colCount)
    If requestedTotal <= 0 Then Exit Function
    If requestedTotal <= availableWidth + 0.1 Then Exit Function

    scaleK = availableWidth / requestedTotal
    If scaleK <= 0 Then Exit Function

    ReDim srcWidths(1 To colCount)
    For colIndex = 1 To colCount
        srcWidths(colIndex) = ws.Columns(firstCol + colIndex - 1).Width
    Next colIndex

    For colIndex = 1 To colCount
        targetPoints = srcWidths(colIndex) * scaleK
        If targetPoints <= 0.1 Then targetPoints = 0.1
        If Not mp_SetColumnWidthByPoints(ws.Columns(firstCol + colIndex - 1), targetPoints) Then
            Exit Function
        End If
    Next colIndex

    m_ScaleConfigColumnsToStableTarget = True
End Function

Public Function m_EqualizeConfigColumnsToStableTarget( _
    ByVal ws As Worksheet, _
    ByVal firstCol As Long, _
    ByVal colCount As Long, _
    ByVal targetStableZoneLeft As Double _
) As Boolean
    m_EqualizeConfigColumnsToStableTarget = m_ScaleConfigColumnsToStableTarget(ws, firstCol, colCount, targetStableZoneLeft)
End Function

Public Sub m_ApplyConfigMarkerStyles(ByVal tbl As ListObject)
    Dim rowCount As Long
    Dim i As Long
    Dim keyText As String
    Dim markerKind As String
    Dim rowRange As Range

    If tbl Is Nothing Then Exit Sub
    rowCount = m_GetTableDataRowCount(tbl)

    If rowCount > 0 Then
        For i = 1 To rowCount
            Set rowRange = tbl.DataBodyRange.Cells(i, 1).Resize(1, DEV_CONFIG_COL_COUNT)
            keyText = Trim$(CStr(rowRange.Cells(1, DEV_CONFIG_KEY_COL).Value))
            markerKind = Trim$(CStr(rowRange.Cells(1, DEV_CONFIG_MARKER_COL).Value))
            If StrComp(markerKind, DEV_MARKER_SYMBOL, vbTextCompare) <> 0 And Not m_IsMarkerKey(keyText) Then GoTo NextRow

            rowRange.Cells(1, DEV_CONFIG_MARKER_COL).Value = DEV_MARKER_SYMBOL
            If Len(Trim$(CStr(rowRange.Cells(1, DEV_CONFIG_KEY_COL).Value))) = 0 Then
                rowRange.Cells(1, DEV_CONFIG_VALUE_COL).Value = vbNullString
                rowRange.Cells(1, DEV_CONFIG_STYLES_COL).Value = vbNullString
            End If
NextRow:
        Next i
    End If

    mp_ApplyConfigStylesFromPipeline tbl
End Sub

Private Sub mp_ApplyConfigStylesFromPipeline(ByVal tbl As ListObject)
    Dim ws As Worksheet
    Dim rowKindRanges As Object
    Dim allRows As Collection
    Dim headerRows As Collection
    Dim dataRows As Collection
    Dim markerRows As Collection
    Dim rowCount As Long
    Dim i As Long
    Dim rowIndex As Long
    Dim markerKind As String
    Dim keyText As String

    If tbl Is Nothing Then Exit Sub
    Set ws = tbl.Parent
    If ws Is Nothing Then Exit Sub

    Set rowKindRanges = CreateObject("Scripting.Dictionary")
    rowKindRanges.CompareMode = 1

    Set allRows = New Collection
    Set headerRows = New Collection
    Set dataRows = New Collection
    Set markerRows = New Collection

    rowIndex = tbl.HeaderRowRange.Row
    allRows.Add rowIndex
    headerRows.Add rowIndex

    rowCount = m_GetTableDataRowCount(tbl)
    For i = 1 To rowCount
        rowIndex = tbl.DataBodyRange.Row + i - 1
        allRows.Add rowIndex

        markerKind = Trim$(CStr(tbl.DataBodyRange.Cells(i, DEV_CONFIG_MARKER_COL).Value))
        keyText = Trim$(CStr(tbl.DataBodyRange.Cells(i, DEV_CONFIG_KEY_COL).Value))
        If StrComp(markerKind, DEV_MARKER_SYMBOL, vbTextCompare) = 0 Or m_IsMarkerKey(keyText) Then
            markerRows.Add rowIndex
        Else
            dataRows.Add rowIndex
        End If
    Next i

    Set rowKindRanges("configall") = allRows
    Set rowKindRanges("configheader") = headerRows
    Set rowKindRanges("configdata") = dataRows
    Set rowKindRanges("configmarker") = markerRows

    ex_OutputFormattingPipeline.m_ApplySheetPipeline ws, Nothing, Nothing, rowKindRanges
    m_AutoFitConfigColumnsWithinStableZone ws, tbl.Range.Column, DEV_CONFIG_COL_COUNT, DEV_CONFIG_MARKER_COL
End Sub

Public Sub m_NormalizeLegacyMarkerEntry(ByRef entries As Variant, ByVal rowIndex As Long)
    Dim markerText As String
    Dim keyText As String
    Dim valueText As String

    markerText = Trim$(CStr(entries(rowIndex, DEV_CONFIG_MARKER_COL)))
    keyText = Trim$(CStr(entries(rowIndex, DEV_CONFIG_KEY_COL)))
    valueText = CStr(entries(rowIndex, DEV_CONFIG_VALUE_COL))

    If StrComp(markerText, DEV_MARKER_SYMBOL, vbTextCompare) = 0 Then Exit Sub

    If StrComp(Left$(keyText, Len(DEV_MARKER_SECTION)), DEV_MARKER_SECTION, vbTextCompare) = 0 Then
        entries(rowIndex, DEV_CONFIG_MARKER_COL) = DEV_MARKER_SYMBOL
        entries(rowIndex, DEV_CONFIG_KEY_COL) = valueText
        entries(rowIndex, DEV_CONFIG_VALUE_COL) = vbNullString
        entries(rowIndex, DEV_CONFIG_STYLES_COL) = vbNullString
        Exit Sub
    End If

    If StrComp(Left$(keyText, Len(DEV_MARKER_SPACER)), DEV_MARKER_SPACER, vbTextCompare) = 0 Then
        entries(rowIndex, DEV_CONFIG_MARKER_COL) = DEV_MARKER_SYMBOL
        entries(rowIndex, DEV_CONFIG_KEY_COL) = vbNullString
        entries(rowIndex, DEV_CONFIG_VALUE_COL) = vbNullString
        entries(rowIndex, DEV_CONFIG_STYLES_COL) = vbNullString
    End If
End Sub

Private Function m_CreateConfigTable(ByVal ws As Worksheet) As ListObject
    Dim lastRow As Long
    Dim rangeToTable As Range

    If Trim$(CStr(ws.Cells(DEV_CONFIG_HEADER_ROW, DEV_CONFIG_MARKER_COL).Value)) <> DEV_MARKER_HEADER Then ws.Cells(DEV_CONFIG_HEADER_ROW, DEV_CONFIG_MARKER_COL).Value = DEV_MARKER_HEADER
    If Trim$(CStr(ws.Cells(DEV_CONFIG_HEADER_ROW, DEV_CONFIG_KEY_COL).Value)) <> "Key" Then ws.Cells(DEV_CONFIG_HEADER_ROW, DEV_CONFIG_KEY_COL).Value = "Key"
    If Trim$(CStr(ws.Cells(DEV_CONFIG_HEADER_ROW, DEV_CONFIG_STYLES_COL).Value)) <> DEV_HEADER_STYLES Then ws.Cells(DEV_CONFIG_HEADER_ROW, DEV_CONFIG_STYLES_COL).Value = DEV_HEADER_STYLES

    lastRow = m_GetLastConfigRow(ws)
    If lastRow < DEV_CONFIG_HEADER_ROW Then lastRow = DEV_CONFIG_HEADER_ROW

    Set rangeToTable = ws.Range(ws.Cells(DEV_CONFIG_HEADER_ROW, DEV_CONFIG_MARKER_COL), ws.Cells(lastRow, DEV_CONFIG_STYLES_COL))

    On Error Resume Next
    Set m_CreateConfigTable = ws.ListObjects.Add(xlSrcRange, rangeToTable, , xlYes)
    If Err.Number <> 0 Then
        MsgBox "Failed to create config table on sheet '" & ws.Name & "': " & Err.Description, vbExclamation
        Err.Clear
        On Error GoTo 0
        Set m_CreateConfigTable = Nothing
        Exit Function
    End If
    On Error GoTo 0

    m_EnsureConfigTableLayout ws, m_CreateConfigTable

    On Error Resume Next
    m_CreateConfigTable.Name = DEV_CONFIG_TABLE_NAME
    ex_ConfigProvider.m_RefreshConfigTitle ws
    On Error GoTo 0
End Function

Private Sub m_EnsureConfigTableLayout(ByVal ws As Worksheet, ByVal tbl As ListObject)
    Dim rowCount As Long
    Dim i As Long
    Dim oldData As Variant
    Dim migrated() As Variant
    Dim styleColIndex As Long

    If tbl.HeaderRowRange.Row <> DEV_CONFIG_HEADER_ROW Then
        Set tbl = mp_RecreateTableAtHeaderRow(ws, tbl)
        If tbl Is Nothing Then Exit Sub
    End If

    rowCount = m_GetTableDataRowCount(tbl)
    If tbl.ListColumns.Count = DEV_CONFIG_COL_COUNT Then
        tbl.HeaderRowRange.Cells(1, DEV_CONFIG_MARKER_COL).Value = DEV_MARKER_HEADER
        tbl.HeaderRowRange.Cells(1, DEV_CONFIG_KEY_COL).Value = "Key"
        If Trim$(CStr(tbl.HeaderRowRange.Cells(1, DEV_CONFIG_STYLES_COL).Value)) <> DEV_HEADER_STYLES Then
            tbl.HeaderRowRange.Cells(1, DEV_CONFIG_STYLES_COL).Value = DEV_HEADER_STYLES
        End If
        On Error Resume Next
        ex_ConfigProvider.m_RefreshConfigTitle ws
        On Error GoTo 0
        Exit Sub
    End If

    If tbl.ListColumns.Count = 2 Then
        If rowCount > 0 Then
            oldData = tbl.DataBodyRange.Cells(1, 1).Resize(rowCount, 2).Value
            ReDim migrated(1 To rowCount, 1 To DEV_CONFIG_COL_COUNT)
            styleColIndex = tbl.Range.Column + 2
            For i = 1 To rowCount
                migrated(i, DEV_CONFIG_MARKER_COL) = vbNullString
                migrated(i, DEV_CONFIG_KEY_COL) = CStr(oldData(i, 1))
                migrated(i, DEV_CONFIG_VALUE_COL) = CStr(oldData(i, 2))
                migrated(i, DEV_CONFIG_STYLES_COL) = CStr(ws.Cells(tbl.HeaderRowRange.Row + i, styleColIndex).Value)
                m_NormalizeLegacyMarkerEntry migrated, i
            Next i
        End If

        tbl.Resize ws.Range(ws.Cells(tbl.HeaderRowRange.Row, tbl.Range.Column), ws.Cells(tbl.HeaderRowRange.Row + rowCount, tbl.Range.Column + DEV_CONFIG_COL_COUNT - 1))
        tbl.HeaderRowRange.Cells(1, DEV_CONFIG_MARKER_COL).Value = DEV_MARKER_HEADER
        tbl.HeaderRowRange.Cells(1, DEV_CONFIG_KEY_COL).Value = "Key"
        tbl.HeaderRowRange.Cells(1, DEV_CONFIG_STYLES_COL).Value = DEV_HEADER_STYLES
        On Error Resume Next
        ex_ConfigProvider.m_RefreshConfigTitle ws
        On Error GoTo 0
        If rowCount > 0 Then tbl.DataBodyRange.Cells(1, 1).Resize(rowCount, DEV_CONFIG_COL_COUNT).Value = migrated
        Exit Sub
    End If

    MsgBox "Unsupported config table layout in '" & DEV_CONFIG_TABLE_NAME & "' (columns: " & CStr(tbl.ListColumns.Count) & ").", vbExclamation
End Sub

Private Function mp_RecreateTableAtHeaderRow(ByVal ws As Worksheet, ByVal tbl As ListObject) As ListObject
    Dim tableName As String
    Dim colCount As Long
    Dim rowCount As Long
    Dim leftCol As Long
    Dim sourceRange As Range
    Dim targetRange As Range
    Dim headerValues As Variant
    Dim bodyValues As Variant

    If ws Is Nothing Or tbl Is Nothing Then Exit Function

    tableName = tbl.Name
    colCount = tbl.ListColumns.Count
    rowCount = m_GetTableDataRowCount(tbl)
    leftCol = tbl.Range.Column

    Set sourceRange = tbl.Range
    headerValues = tbl.HeaderRowRange.Cells(1, 1).Resize(1, colCount).Value
    If rowCount > 0 Then
        bodyValues = tbl.DataBodyRange.Cells(1, 1).Resize(rowCount, colCount).Value
    End If

    On Error Resume Next
    tbl.Unlist
    If Err.Number <> 0 Then
        MsgBox "Failed to move config table '" & tableName & "' to row " & CStr(DEV_CONFIG_HEADER_ROW) & ": " & Err.Description, vbExclamation
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    sourceRange.Clear

    Set targetRange = ws.Range( _
        ws.Cells(DEV_CONFIG_HEADER_ROW, leftCol), _
        ws.Cells(DEV_CONFIG_HEADER_ROW + rowCount, leftCol + colCount - 1) _
    )
    targetRange.Clear
    targetRange.Cells(1, 1).Resize(1, colCount).Value = headerValues
    If rowCount > 0 Then
        targetRange.Cells(2, 1).Resize(rowCount, colCount).Value = bodyValues
    End If

    On Error Resume Next
    Set mp_RecreateTableAtHeaderRow = ws.ListObjects.Add(xlSrcRange, targetRange, , xlYes)
    If Err.Number <> 0 Then
        MsgBox "Failed to recreate config table '" & tableName & "' at row " & CStr(DEV_CONFIG_HEADER_ROW) & ": " & Err.Description, vbExclamation
        Err.Clear
        On Error GoTo 0
        Set mp_RecreateTableAtHeaderRow = Nothing
        Exit Function
    End If
    On Error GoTo 0

    On Error Resume Next
    mp_RecreateTableAtHeaderRow.Name = tableName
    On Error GoTo 0
End Function

Private Function m_GetLastConfigRow(ByVal ws As Worksheet) As Long
    Dim lastKey As Long
    Dim lastValue As Long

    lastKey = ws.Cells(ws.Rows.Count, DEV_CONFIG_MARKER_COL).End(xlUp).Row
    lastValue = ws.Cells(ws.Rows.Count, DEV_CONFIG_VALUE_COL).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, DEV_CONFIG_STYLES_COL).End(xlUp).Row > lastValue Then
        lastValue = ws.Cells(ws.Rows.Count, DEV_CONFIG_STYLES_COL).End(xlUp).Row
    End If

    m_GetLastConfigRow = lastKey
    If lastValue > m_GetLastConfigRow Then m_GetLastConfigRow = lastValue
End Function

Private Function m_IsMarkerKey(ByVal keyText As String) As Boolean
    keyText = Trim$(keyText)
    If Len(keyText) < Len(DEV_MARKER_PREFIX) Then Exit Function
    m_IsMarkerKey = (StrComp(Left$(keyText, Len(DEV_MARKER_PREFIX)), DEV_MARKER_PREFIX, vbTextCompare) = 0)
End Function

Private Sub mp_ShrinkConfigColumnsToMaxWidth( _
    ByVal ws As Worksheet, _
    ByVal firstCol As Long, _
    ByVal colCount As Long, _
    ByVal markerAbsCol As Long, _
    ByVal maxWidth As Double)

    Dim endCol As Long
    Dim pass As Long
    Dim colIndex As Long
    Dim totalWidth As Double
    Dim overflow As Double
    Dim currentPoints As Double
    Dim minPoints As Double
    Dim targetPoints As Double
    Dim markerMinPoints As Double

    endCol = firstCol + colCount - 1
    markerMinPoints = mp_GetColumnWidthPointsForUnits(ws.Columns(markerAbsCol), MIN_MARKER_WIDTH_UNITS)
    If markerMinPoints <= 0 Then markerMinPoints = MIN_CONFIG_DATA_COL_WIDTH_POINTS

    For pass = 1 To 3
        totalWidth = mp_GetColumnsWidthPoints(ws, firstCol, colCount)
        overflow = totalWidth - maxWidth
        If overflow <= 0.1 Then Exit Sub

        For colIndex = endCol To firstCol Step -1
            currentPoints = ws.Columns(colIndex).Width
            minPoints = MIN_CONFIG_DATA_COL_WIDTH_POINTS
            If colIndex = markerAbsCol Then minPoints = markerMinPoints

            If currentPoints > minPoints + 0.1 Then
                targetPoints = currentPoints - overflow
                If targetPoints < minPoints Then targetPoints = minPoints
                mp_SetColumnWidthByPoints ws.Columns(colIndex), targetPoints

                totalWidth = mp_GetColumnsWidthPoints(ws, firstCol, colCount)
                overflow = totalWidth - maxWidth
                If overflow <= 0.1 Then Exit Sub
            End If
        Next colIndex
    Next pass
End Sub

Private Function mp_GetColumnsWidthPoints(ByVal ws As Worksheet, ByVal firstCol As Long, ByVal colCount As Long) As Double
    If ws Is Nothing Then Exit Function
    If firstCol < 1 Or colCount < 1 Then Exit Function
    mp_GetColumnsWidthPoints = ws.Columns(firstCol).Resize(, colCount).Width
End Function

Private Function mp_GetColumnWidthPointsForUnits(ByVal colRange As Range, ByVal widthUnits As Double) As Double
    Dim prevUnits As Double

    If colRange Is Nothing Then Exit Function
    If widthUnits <= 0 Then Exit Function

    On Error GoTo EH
    prevUnits = colRange.ColumnWidth
    colRange.ColumnWidth = widthUnits
    mp_GetColumnWidthPointsForUnits = colRange.Width
    colRange.ColumnWidth = prevUnits
    Exit Function
EH:
    On Error Resume Next
    If prevUnits > 0 Then colRange.ColumnWidth = prevUnits
    On Error GoTo 0
    mp_GetColumnWidthPointsForUnits = 0
End Function

Private Function mp_TryGetStableZoneColumns(ByVal ws As Worksheet, ByRef stableStartCol As Long, ByRef bufferCol As Long) As Boolean
    Dim stableColText As String

    stableColText = Trim$(ex_UiXmlProvider.m_GetLayoutAttribute("stableZone", "startCol", ThisWorkbook))
    If Len(stableColText) = 0 Then
        MsgBox "UI layout must define /uiDefinition/layout/stableZone@startCol in DevUI.xml.", vbExclamation
        Exit Function
    End If

    If Not mp_TryResolveColumnIndex(ws, stableColText, stableStartCol) Then
        MsgBox "Invalid stable zone startCol in DevUI.xml: '" & stableColText & "'.", vbExclamation
        Exit Function
    End If

    If stableStartCol <= 1 Then
        MsgBox "Layout stable zone startCol must be greater than column A.", vbExclamation
        Exit Function
    End If

    bufferCol = stableStartCol - 1
    mp_TryGetStableZoneColumns = True
End Function

Private Function mp_GetStableZoneMinBufferWidthUnits() As Double
    Dim widthText As String

    widthText = Trim$(ex_UiXmlProvider.m_GetLayoutAttribute("stableZone", "minBufferWidth", ThisWorkbook))
    If Len(widthText) = 0 Then
        MsgBox "UI layout must define /uiDefinition/layout/stableZone@minBufferWidth in DevUI.xml.", vbExclamation
        Exit Function
    End If

    If Not IsNumeric(widthText) Then
        MsgBox "Invalid stable zone minBufferWidth in DevUI.xml: '" & widthText & "'.", vbExclamation
        Exit Function
    End If

    mp_GetStableZoneMinBufferWidthUnits = CDbl(widthText)
    If mp_GetStableZoneMinBufferWidthUnits <= 0 Then
        MsgBox "Invalid stable zone minBufferWidth in DevUI.xml: value must be > 0.", vbExclamation
        mp_GetStableZoneMinBufferWidthUnits = 0
    End If
End Function

Private Function mp_TryResolveColumnIndex(ByVal ws As Worksheet, ByVal valueText As String, ByRef outColumnIndex As Long) As Boolean
    Dim parsed As Long
    Dim refRange As Range

    valueText = Trim$(valueText)
    If Len(valueText) = 0 Then Exit Function

    If IsNumeric(valueText) Then
        parsed = CLng(valueText)
        If parsed > 0 Then
            outColumnIndex = parsed
            mp_TryResolveColumnIndex = True
            Exit Function
        End If
    End If

    On Error Resume Next
    Set refRange = ws.Range(valueText)
    If refRange Is Nothing Then Set refRange = ws.Range(valueText & "1")
    If refRange Is Nothing Then Set refRange = ws.Columns(valueText & ":" & valueText)
    On Error GoTo 0

    If refRange Is Nothing Then Exit Function

    outColumnIndex = refRange.Column
    mp_TryResolveColumnIndex = (outColumnIndex > 0)
End Function

Private Function mp_SetColumnWidthByPoints(ByVal colRange As Range, ByVal targetPoints As Double) As Boolean
    Dim i As Long
    Dim currentPoints As Double
    Dim currentWidthUnits As Double
    Dim slope As Double
    Dim deltaPoints As Double

    If colRange Is Nothing Then Exit Function
    If targetPoints <= 0 Then Exit Function

    On Error GoTo EH
    currentPoints = colRange.Width
    currentWidthUnits = colRange.ColumnWidth
    If currentPoints <= 0 Or currentWidthUnits <= 0 Then Exit Function

    colRange.ColumnWidth = currentWidthUnits * (targetPoints / currentPoints)

    For i = 1 To 8
        currentPoints = colRange.Width
        deltaPoints = targetPoints - currentPoints
        If Abs(deltaPoints) < 0.1 Then
            mp_SetColumnWidthByPoints = True
            Exit Function
        End If

        currentWidthUnits = colRange.ColumnWidth
        If currentWidthUnits <= 0 Then Exit For
        slope = currentPoints / currentWidthUnits
        If slope <= 0 Then Exit For

        colRange.ColumnWidth = currentWidthUnits + (deltaPoints / slope)
        If colRange.ColumnWidth < 0.1 Then colRange.ColumnWidth = 0.1
    Next i

    mp_SetColumnWidthByPoints = (Abs(targetPoints - colRange.Width) < 0.5)
    Exit Function
EH:
    mp_SetColumnWidthByPoints = False
End Function
