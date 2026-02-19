Attribute VB_Name = "ex_ConfigTableStore"
Option Explicit

Private Const DEV_CONFIG_TABLE_NAME As String = "tblDevConfig"
Private Const DEV_CONFIG_HEADER_ROW As Long = 2
Private Const DEV_CONFIG_MARKER_COL As Long = 1
Private Const DEV_CONFIG_KEY_COL As Long = 2
Private Const DEV_CONFIG_VALUE_COL As Long = 3
Private Const DEV_CONFIG_NOTE_COL As Long = 4
Private Const DEV_CONFIG_COL_COUNT As Long = 4
Private Const DEV_MARKER_SYMBOL As String = "#"
Private Const DEV_MARKER_HEADER As String = ".."
Private Const DEV_MARKER_PREFIX As String = "#MARKER:"
Private Const DEV_MARKER_SECTION As String = "#MARKER:SECTION"
Private Const DEV_MARKER_SPACER As String = "#MARKER:SPACER"
Private Const DEV_COLOR_BG As Long = &H1E1E1E
Private Const DEV_COLOR_TEXT As Long = &HEBEBEB
Private Const DEV_COLOR_BORDER As Long = &H505050
Private Const DEV_COLOR_NOTE_TEXT As Long = &HA8A8A8
Private Const THEME_BG As Long = &H262626
Private Const THEME_TEXT As Long = &HEBEBEB
Private Const THEME_BORDER As Long = &H0

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

Public Sub m_ApplySheetThemeToFormerTableTail(ByVal ws As Worksheet, ByVal tbl As ListObject, ByVal previousRowCount As Long, ByVal newRowCount As Long)
    Dim topRow As Long
    Dim leftCol As Long
    Dim rightCol As Long
    Dim previousBottom As Long
    Dim newBottom As Long
    Dim tailRange As Range

    If previousRowCount <= newRowCount Then Exit Sub

    topRow = tbl.HeaderRowRange.Row
    leftCol = tbl.Range.Column
    rightCol = leftCol + DEV_CONFIG_COL_COUNT - 1
    previousBottom = topRow + previousRowCount
    newBottom = topRow + newRowCount
    If previousBottom <= newBottom Then Exit Sub

    Set tailRange = ws.Range(ws.Cells(newBottom + 1, leftCol), ws.Cells(previousBottom, rightCol))
    With tailRange
        .Interior.Pattern = xlSolid
        .Interior.Color = THEME_BG
        .Font.Color = THEME_TEXT
        .Borders.LineStyle = xlContinuous
        .Borders.Color = THEME_BORDER
        .Borders.Weight = xlThin
    End With
End Sub

Public Sub m_ApplyConfigTableDarkTheme(ByVal tbl As ListObject)
    Dim targetRange As Range
    Dim bodyRange As Range

    On Error Resume Next
    tbl.TableStyle = vbNullString
    tbl.ShowTableStyleColumnStripes = False
    tbl.ShowTableStyleRowStripes = False
    tbl.ShowTableStyleFirstColumn = False
    tbl.ShowTableStyleLastColumn = False
    On Error GoTo 0

    Set targetRange = tbl.Range
    With targetRange
        .Interior.Pattern = xlSolid
        .Interior.Color = DEV_COLOR_BG
        .Font.Color = DEV_COLOR_TEXT
        .Font.Bold = False
        .Borders.LineStyle = xlContinuous
        .Borders.Color = DEV_COLOR_BORDER
        .Borders.Weight = xlThin
    End With

    Set bodyRange = tbl.DataBodyRange
    If Not bodyRange Is Nothing Then
        bodyRange.Font.Bold = False
        bodyRange.Columns(DEV_CONFIG_NOTE_COL).Font.Color = DEV_COLOR_NOTE_TEXT
    End If

    tbl.HeaderRowRange.Font.Bold = True
    tbl.HeaderRowRange.Cells(1, DEV_CONFIG_NOTE_COL).Font.Color = DEV_COLOR_TEXT
    tbl.Range.EntireColumn.AutoFit
    If tbl.ListColumns(DEV_CONFIG_MARKER_COL).Range.ColumnWidth < 4 Then
        tbl.ListColumns(DEV_CONFIG_MARKER_COL).Range.ColumnWidth = 4
    End If
End Sub

Public Sub m_ApplyConfigMarkerStyles(ByVal tbl As ListObject)
    Dim rowCount As Long
    Dim i As Long
    Dim keyText As String
    Dim markerKind As String
    Dim rowRange As Range

    rowCount = m_GetTableDataRowCount(tbl)
    If rowCount <= 0 Then Exit Sub

    For i = 1 To rowCount
        Set rowRange = tbl.DataBodyRange.Cells(i, 1).Resize(1, DEV_CONFIG_COL_COUNT)
        keyText = Trim$(CStr(rowRange.Cells(1, DEV_CONFIG_KEY_COL).Value))
        markerKind = Trim$(CStr(rowRange.Cells(1, DEV_CONFIG_MARKER_COL).Value))
        If StrComp(markerKind, DEV_MARKER_SYMBOL, vbTextCompare) <> 0 And Not m_IsMarkerKey(keyText) Then GoTo NextRow

        rowRange.Cells(1, DEV_CONFIG_MARKER_COL).Value = DEV_MARKER_SYMBOL
        rowRange.Interior.Pattern = xlSolid
        rowRange.Interior.Color = RGB(45, 45, 45)
        rowRange.Font.Color = DEV_COLOR_TEXT
        rowRange.Font.Bold = False
        rowRange.Cells(1, DEV_CONFIG_MARKER_COL).Font.Color = DEV_COLOR_TEXT

        If Len(Trim$(CStr(rowRange.Cells(1, DEV_CONFIG_KEY_COL).Value))) > 0 Then
            rowRange.Cells(1, DEV_CONFIG_KEY_COL).Font.Bold = True
            rowRange.Cells(1, DEV_CONFIG_KEY_COL).Font.Color = RGB(245, 245, 245)
        Else
            rowRange.Cells(1, DEV_CONFIG_VALUE_COL).Value = vbNullString
            rowRange.Cells(1, DEV_CONFIG_NOTE_COL).Value = vbNullString
        End If
NextRow:
    Next i
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
        entries(rowIndex, DEV_CONFIG_NOTE_COL) = vbNullString
        Exit Sub
    End If

    If StrComp(Left$(keyText, Len(DEV_MARKER_SPACER)), DEV_MARKER_SPACER, vbTextCompare) = 0 Then
        entries(rowIndex, DEV_CONFIG_MARKER_COL) = DEV_MARKER_SYMBOL
        entries(rowIndex, DEV_CONFIG_KEY_COL) = vbNullString
        entries(rowIndex, DEV_CONFIG_VALUE_COL) = vbNullString
        entries(rowIndex, DEV_CONFIG_NOTE_COL) = vbNullString
    End If
End Sub

Private Function m_CreateConfigTable(ByVal ws As Worksheet) As ListObject
    Dim lastRow As Long
    Dim rangeToTable As Range

    If Trim$(CStr(ws.Cells(DEV_CONFIG_HEADER_ROW, DEV_CONFIG_MARKER_COL).Value)) <> DEV_MARKER_HEADER Then ws.Cells(DEV_CONFIG_HEADER_ROW, DEV_CONFIG_MARKER_COL).Value = DEV_MARKER_HEADER
    If Trim$(CStr(ws.Cells(DEV_CONFIG_HEADER_ROW, DEV_CONFIG_KEY_COL).Value)) <> "Key" Then ws.Cells(DEV_CONFIG_HEADER_ROW, DEV_CONFIG_KEY_COL).Value = "Key"
    If Trim$(CStr(ws.Cells(DEV_CONFIG_HEADER_ROW, DEV_CONFIG_NOTE_COL).Value)) <> "Note" Then ws.Cells(DEV_CONFIG_HEADER_ROW, DEV_CONFIG_NOTE_COL).Value = "Note"

    lastRow = m_GetLastConfigRow(ws)
    If lastRow < DEV_CONFIG_HEADER_ROW Then lastRow = DEV_CONFIG_HEADER_ROW

    Set rangeToTable = ws.Range(ws.Cells(DEV_CONFIG_HEADER_ROW, DEV_CONFIG_MARKER_COL), ws.Cells(lastRow, DEV_CONFIG_NOTE_COL))

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
    Dim noteColIndex As Long

    rowCount = m_GetTableDataRowCount(tbl)
    If tbl.ListColumns.Count = DEV_CONFIG_COL_COUNT Then
        tbl.HeaderRowRange.Cells(1, DEV_CONFIG_MARKER_COL).Value = DEV_MARKER_HEADER
        tbl.HeaderRowRange.Cells(1, DEV_CONFIG_KEY_COL).Value = "Key"
        tbl.HeaderRowRange.Cells(1, DEV_CONFIG_NOTE_COL).Value = "Note"
        On Error Resume Next
        ex_ConfigProvider.m_RefreshConfigTitle ws
        On Error GoTo 0
        Exit Sub
    End If

    If tbl.ListColumns.Count = 2 Then
        If rowCount > 0 Then
            oldData = tbl.DataBodyRange.Cells(1, 1).Resize(rowCount, 2).Value
            ReDim migrated(1 To rowCount, 1 To DEV_CONFIG_COL_COUNT)
            noteColIndex = tbl.Range.Column + 2
            For i = 1 To rowCount
                migrated(i, DEV_CONFIG_MARKER_COL) = vbNullString
                migrated(i, DEV_CONFIG_KEY_COL) = CStr(oldData(i, 1))
                migrated(i, DEV_CONFIG_VALUE_COL) = CStr(oldData(i, 2))
                migrated(i, DEV_CONFIG_NOTE_COL) = CStr(ws.Cells(tbl.HeaderRowRange.Row + i, noteColIndex).Value)
                m_NormalizeLegacyMarkerEntry migrated, i
            Next i
        End If

        tbl.Resize ws.Range(ws.Cells(tbl.HeaderRowRange.Row, tbl.Range.Column), ws.Cells(tbl.HeaderRowRange.Row + rowCount, tbl.Range.Column + DEV_CONFIG_COL_COUNT - 1))
        tbl.HeaderRowRange.Cells(1, DEV_CONFIG_MARKER_COL).Value = DEV_MARKER_HEADER
        tbl.HeaderRowRange.Cells(1, DEV_CONFIG_KEY_COL).Value = "Key"
        tbl.HeaderRowRange.Cells(1, DEV_CONFIG_NOTE_COL).Value = "Note"
        On Error Resume Next
        ex_ConfigProvider.m_RefreshConfigTitle ws
        On Error GoTo 0
        If rowCount > 0 Then tbl.DataBodyRange.Cells(1, 1).Resize(rowCount, DEV_CONFIG_COL_COUNT).Value = migrated
        Exit Sub
    End If

    MsgBox "Unsupported config table layout in '" & DEV_CONFIG_TABLE_NAME & "' (columns: " & CStr(tbl.ListColumns.Count) & ").", vbExclamation
End Sub

Private Function m_GetLastConfigRow(ByVal ws As Worksheet) As Long
    Dim lastKey As Long
    Dim lastValue As Long

    lastKey = ws.Cells(ws.Rows.Count, DEV_CONFIG_MARKER_COL).End(xlUp).Row
    lastValue = ws.Cells(ws.Rows.Count, DEV_CONFIG_VALUE_COL).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, DEV_CONFIG_NOTE_COL).End(xlUp).Row > lastValue Then
        lastValue = ws.Cells(ws.Rows.Count, DEV_CONFIG_NOTE_COL).End(xlUp).Row
    End If

    m_GetLastConfigRow = lastKey
    If lastValue > m_GetLastConfigRow Then m_GetLastConfigRow = lastValue
End Function

Private Function m_IsMarkerKey(ByVal keyText As String) As Boolean
    keyText = Trim$(keyText)
    If Len(keyText) < Len(DEV_MARKER_PREFIX) Then Exit Function
    m_IsMarkerKey = (StrComp(Left$(keyText, Len(DEV_MARKER_PREFIX)), DEV_MARKER_PREFIX, vbTextCompare) = 0)
End Function
