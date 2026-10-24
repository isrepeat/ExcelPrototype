Option Explicit

Private m_filterInputActive As Boolean
Private m_filterInputCell As Range
Private m_filterInputAddress As String
Private m_filterInputHadFormula As Boolean
Private m_filterInputFormula As Variant
Private m_filterInputValue As Variant
Private m_filterInputInteriorColor As Long
Private m_filterInputInteriorPattern As Long
Private m_filterTable As ListObject
Private m_filterFieldIndex As Long
Private m_filterHeaderCell As Range

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_AppendTenTableRows()
    Const ROW_COUNT As Long = 10
    Dim tableObj As ListObject
    Dim templateRow As Range
    Dim addedRows As Range
    Dim templateCell As Range
    Dim originalTableRows As Long
    Dim previousCalculation As XlCalculation
    Dim previousEnableEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim stateChanged As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim errorSource As String

    On Error GoTo EH

    ' Validate the target before changing the table or application state.
    If Not TypeOf Application.ActiveSheet Is Worksheet Then _
        Err.Raise 5, "PERSONAL.Table.WorksheetRequired"
    Set tableObj = Application.ActiveCell.ListObject
    If tableObj Is Nothing Then _
        Err.Raise 5, "PERSONAL.Table.SelectCellInsideTable"
    If tableObj.DataBodyRange Is Nothing Then _
        Err.Raise 5, "PERSONAL.Table.TemplateRowRequired"
    If tableObj.Parent.ProtectContents Then _
        Err.Raise 1004, "PERSONAL.Table.WorksheetProtected"
    If tableObj.Range.Row + tableObj.Range.Rows.Count + ROW_COUNT - 1 > tableObj.Parent.Rows.Count Then _
        Err.Raise 6, "PERSONAL.Table.InsufficientWorksheetRows"
    Set templateRow = tableObj.DataBodyRange.Rows(tableObj.ListRows.Count)
    originalTableRows = tableObj.Range.Rows.Count
    Set addedRows = templateRow.Offset(1, 0).Resize(ROW_COUNT)

    ' Expand once instead of rebuilding table dependencies for each new row.
    previousEnableEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    previousCalculation = Application.Calculation
    stateChanged = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    addedRows.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    tableObj.Resize tableObj.Range.Cells(1, 1).Resize( _
        originalTableRows + ROW_COUNT, tableObj.ListColumns.Count)
    Set addedRows = templateRow.Offset(1, 0).Resize(ROW_COUNT)

    ' Apply the template to the entire block, keeping constant fields empty.
    templateRow.Copy
    addedRows.PasteSpecial Paste:=xlPasteFormats
    addedRows.PasteSpecial Paste:=xlPasteValidation
    addedRows.ClearContents
    addedRows.RowHeight = templateRow.RowHeight
    For Each templateCell In templateRow.Cells
        If templateCell.HasFormula Then
            addedRows.Columns(templateCell.Column - templateRow.Column + 1).FormulaR1C1 = _
                templateCell.FormulaR1C1
        End If
    Next templateCell
    ex_Core.fn_Diagnostic_WriteLog "TABLE_ROWS_APPENDED | Table=" & tableObj.Name & _
        " | Count=" & VBA.CStr(ROW_COUNT)

CleanExit:
    If stateChanged Then
        Application.CutCopyMode = False
        Application.Calculation = previousCalculation
        Application.EnableEvents = previousEnableEvents
        Application.ScreenUpdating = previousScreenUpdating
    End If
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    errorSource = VBA.Err.Source
    ex_Core.fn_Diagnostic_WriteLog "TABLE_ROWS_APPEND_ERROR | Source=" & errorSource & _
        " | Description=" & errorDescription
    ex_WindowsUi.fn_ShowMessage errorSource & VBA.vbCrLf & errorDescription, VBA.vbExclamation
    Resume CleanExit
End Sub

Public Sub fn_RemoveEmptyTableTail()
    Const SCAN_ROWS As Long = 256
    Dim tableObj As ListObject
    Dim bodyRange As Range
    Dim scanRange As Range
    Dim removedRange As Range
    Dim sheetFormatRow As Range
    Dim constantCells As Range
    Dim constantArea As Range
    Dim scanEnd As Long
    Dim scanStart As Long
    Dim lastDataRow As Long
    Dim lastConstantRow As Long
    Dim removeCount As Long
    Dim specialCellsError As Long
    Dim previousCalculation As XlCalculation
    Dim previousEnableEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim stateChanged As Boolean
    Dim errorDescription As String
    Dim errorSource As String

    On Error GoTo EH

    ' Only constant input values prevent removal; calculated columns are templates.
    If Not TypeOf Application.ActiveSheet Is Worksheet Then _
        Err.Raise 5, "PERSONAL.Table.WorksheetRequired"
    Set tableObj = Application.ActiveCell.ListObject
    If tableObj Is Nothing Then _
        Err.Raise 5, "PERSONAL.Table.SelectCellInsideTable"
    Set bodyRange = tableObj.DataBodyRange
    If bodyRange Is Nothing Then Exit Sub
    If tableObj.Parent.ProtectContents Then _
        Err.Raise 1004, "PERSONAL.Table.WorksheetProtected"
    lastDataRow = bodyRange.Row - 1
    scanEnd = bodyRange.Rows.Count
    ' Inspect the tail in small blocks instead of enumerating constants in the whole table.
    Do While scanEnd > 0
        scanStart = scanEnd - SCAN_ROWS + 1
        If scanStart < 1 Then scanStart = 1
        Set scanRange = bodyRange.Rows(scanStart).Resize( _
            scanEnd - scanStart + 1, bodyRange.Columns.Count)
        Set constantCells = Nothing
        On Error Resume Next
        Set constantCells = scanRange.SpecialCells(xlCellTypeConstants)
        specialCellsError = VBA.Err.Number
        VBA.Err.Clear
        On Error GoTo EH
        If specialCellsError <> 0 And specialCellsError <> 1004 Then _
            Err.Raise specialCellsError, "PERSONAL.Table.FindConstantCells"
        If Not constantCells Is Nothing Then
            Set constantCells = Application.Intersect(scanRange, constantCells)
        End If
        If Not constantCells Is Nothing Then
            For Each constantArea In constantCells.Areas
                lastConstantRow = constantArea.Row + constantArea.Rows.Count - 1
                If lastConstantRow > lastDataRow Then lastDataRow = lastConstantRow
            Next constantArea
            Exit Do
        End If
        scanEnd = scanStart - 1
    Loop
    removeCount = bodyRange.Row + bodyRange.Rows.Count - 1 - lastDataRow
    If removeCount = 0 Then Exit Sub

    ' Shrink without shifting worksheet cells; retain native deletion for totals or an empty table.
    previousEnableEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    previousCalculation = Application.Calculation
    stateChanged = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Set removedRange = bodyRange.Rows(bodyRange.Rows.Count - removeCount + 1).Resize( _
        removeCount, bodyRange.Columns.Count)
    If tableObj.ShowTotals Or removeCount = bodyRange.Rows.Count Or _
            bodyRange.Row + bodyRange.Rows.Count > tableObj.Parent.Rows.Count Then
        removedRange.Delete Shift:=xlUp
    Else
        Set sheetFormatRow = bodyRange.Rows(bodyRange.Rows.Count).Offset(1, 0)
        tableObj.Resize tableObj.Range.Resize(tableObj.Range.Rows.Count - removeCount)
        removedRange.ClearContents
        ' Resize drops table-style fills; restore the surrounding worksheet format explicitly.
        sheetFormatRow.Copy
        removedRange.PasteSpecial Paste:=xlPasteFormats
        removedRange.RowHeight = sheetFormatRow.RowHeight
    End If
    ex_Core.fn_Diagnostic_WriteLog "TABLE_TAIL_REMOVED | Table=" & tableObj.Name & _
        " | Count=" & VBA.CStr(removeCount)

CleanExit:
    If stateChanged Then
        Application.CutCopyMode = False
        Application.Calculation = previousCalculation
        Application.EnableEvents = previousEnableEvents
        Application.ScreenUpdating = previousScreenUpdating
    End If
    Exit Sub
EH:
    errorDescription = VBA.Err.Description
    errorSource = VBA.Err.Source
    ex_Core.fn_Diagnostic_WriteLog "TABLE_TAIL_REMOVE_ERROR | Source=" & errorSource & _
        " | Description=" & errorDescription
    ex_WindowsUi.fn_ShowMessage errorSource & VBA.vbCrLf & errorDescription, VBA.vbExclamation
    Resume CleanExit
End Sub

Public Sub fn_RecalculateWorkbook()
    Application.Calculate
End Sub

Public Sub fn_RecalculateActiveSheet()
    ActiveSheet.Calculate
End Sub

Public Sub fn_ToggleFirstTwoRows()
    On Error GoTo ExitPoint

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    With ActiveWindow
        .FreezePanes = False
        .SplitRow = 0
        .SplitColumn = 0
    End With

    If ws.Rows("1:2").Hidden Then
        ws.Rows("1:2").Hidden = False
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        ws.Range("A3").Select
        ActiveWindow.FreezePanes = True

    Else
        ws.Rows("1:2").Hidden = True
        ActiveWindow.ScrollRow = 3
        ActiveWindow.ScrollColumn = 1
    End If

ExitPoint:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub fn_DatePlusOne()
    If IsDate(ActiveCell.value) Then
        ActiveCell.value = CDate(ActiveCell.value) + 1
    End If
End Sub

Public Sub fn_DateMinusOne()
    If IsDate(ActiveCell.value) Then
        ActiveCell.value = CDate(ActiveCell.value) - 1
    End If
End Sub

Public Sub fn_FilterContainsCurrentColumn()
    Dim tableObj As ListObject
    Dim columnIndex As Long
    Dim headerCell As Range
    Dim inputCell As Range

    ex_Core.fn_Diagnostic_WriteLog "START | Workbook=" & ActiveWorkbook.Name & _
        " | Sheet=" & ActiveSheet.Name & " | Cell=" & ActiveCell.Address(False, False)

    ' Use the free cell immediately above the active column header as the input field.
    On Error Resume Next
    Set tableObj = ActiveCell.ListObject
    On Error GoTo EH
    If Not tableObj Is Nothing Then
        columnIndex = ActiveCell.Column - tableObj.Range.Columns(1).Column + 1
        If columnIndex < 1 Or columnIndex > tableObj.ListColumns.Count Then Exit Sub
        Set headerCell = tableObj.HeaderRowRange.Cells(1, columnIndex)
        If headerCell.Row = 1 Then
            Set inputCell = tableObj.Parent.Cells( _
                tableObj.Range.Row + tableObj.Range.Rows.Count, _
                headerCell.Column)
        Else
            Set inputCell = headerCell.Offset(-1, 0)
        End If
        ex_Core.fn_Diagnostic_WriteLog "TABLE | Name=" & tableObj.Name & _
            " | Range=" & tableObj.Range.Address(False, False) & _
            " | Field=" & VBA.CStr(columnIndex) & _
            " | Header=" & tableObj.ListColumns(columnIndex).Name
        private_Filter_BeginFilterInput tableObj, columnIndex, headerCell, inputCell
        Exit Sub
    End If

    ex_WindowsUi.fn_ShowMessage "Select a cell inside an Excel Table before using Ctrl+Q.", _
        VBA.vbExclamation, "Filter contains"
    Exit Sub
EH:
    ex_Core.fn_Diagnostic_WriteLog "ERROR | Number=" & VBA.CStr(VBA.Err.Number) & _
        " | Description=" & VBA.Err.Description
    ex_WindowsUi.fn_ShowMessage "Filtering failed: [" & VBA.CStr(VBA.Err.Number) & "] " & _
        VBA.Err.Description, VBA.vbExclamation, "Filter contains"
End Sub

Public Sub fn_HandleFilterInputCellChange( _
    ByVal sheetObject As Object, _
    ByVal target As Range _
)
    If Not private_Filter_IsPendingFilterInput(sheetObject, target) Then Exit Sub
    private_Filter_ApplyPendingFilter
End Sub

Public Sub fn_HandleFilterInputSelectionChange( _
    ByVal sheetObject As Object, _
    ByVal target As Range _
)
    If Not private_Filter_IsPendingFilterSheet(sheetObject) Then Exit Sub
    If target.CountLarge <> 1 Then Exit Sub
    If VBA.StrComp(target.Address(False, False), m_filterInputAddress, _
            VBA.vbTextCompare) = 0 Then Exit Sub

    If VBA.Len(private_Filter_NormalizeFilterQuery(m_filterInputCell.Value2)) = 0 Then
        private_Filter_CancelPendingFilterInput
    Else
        private_Filter_ApplyPendingFilter
    End If
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Filter {
' --------------------------------------
Private Sub private_Filter_BeginFilterInput( _
    ByVal tableObj As ListObject, _
    ByVal fieldIndex As Long, _
    ByVal headerCell As Range, _
    ByVal inputCell As Range _
)
    private_Filter_CancelPendingFilterInput

    If inputCell.HasFormula Or Not IsEmpty(inputCell.Value2) Then
        ex_WindowsUi.fn_ShowMessage "The cell above the table header must be empty for Ctrl+Q filtering.", _
            VBA.vbExclamation, "Filter contains"
        Exit Sub
    End If

    Set m_filterInputCell = inputCell
    Set m_filterTable = tableObj
    Set m_filterHeaderCell = headerCell
    m_filterFieldIndex = fieldIndex
    m_filterInputAddress = inputCell.Address(False, False)
    m_filterInputHadFormula = inputCell.HasFormula
    If m_filterInputHadFormula Then
        m_filterInputFormula = inputCell.Formula
    Else
        m_filterInputValue = inputCell.Value2
    End If
    m_filterInputInteriorColor = inputCell.Interior.Color
    m_filterInputInteriorPattern = inputCell.Interior.Pattern

    m_filterInputActive = True
    inputCell.Interior.Pattern = xlSolid
    inputCell.Interior.Color = RGB(16, 72, 97)
    ' Goto moves the active cell without cancelling a pending copy operation.
    Application.Goto inputCell, False
    ex_Core.fn_Diagnostic_WriteLog "INPUT_STARTED | Address=" & m_filterInputAddress
End Sub

Private Function private_Filter_IsPendingFilterSheet(ByVal sheetObject As Object) As Boolean
    If Not m_filterInputActive Then Exit Function
    If m_filterInputCell Is Nothing Then Exit Function
    private_Filter_IsPendingFilterSheet = sheetObject Is m_filterInputCell.Worksheet
End Function

Private Function private_Filter_IsPendingFilterInput( _
    ByVal sheetObject As Object, _
    ByVal target As Range _
) As Boolean
    If Not private_Filter_IsPendingFilterSheet(sheetObject) Then Exit Function
    If target.CountLarge <> 1 Then Exit Function
    private_Filter_IsPendingFilterInput = VBA.StrComp( _
        target.Address(False, False), m_filterInputAddress, VBA.vbTextCompare) = 0
End Function

Private Sub private_Filter_ApplyPendingFilter()
    Dim filterDate As Date
    Dim query As String
    Dim previousEnableEvents As Boolean

    On Error GoTo EH
    If Not m_filterInputActive Then Exit Sub
    query = private_Filter_NormalizeFilterQuery(VBA.CStr(m_filterInputCell.Value2))
    previousEnableEvents = Application.EnableEvents
    Application.EnableEvents = False
    If VBA.Len(query) > 0 Then
        If private_Filter_IsDateFilterColumn(m_filterTable, m_filterFieldIndex) And _
                (VBA.IsDate(query) Or private_Filter_IsExcelDateSerial(query)) Then
            If private_Filter_IsExcelDateSerial(query) Then
                filterDate = VBA.DateValue(VBA.CDate(VBA.CDbl(query)))
            Else
                filterDate = VBA.DateValue(VBA.CDate(query))
            End If
            m_filterTable.Range.AutoFilter Field:=m_filterFieldIndex, _
                Criteria1:=">=" & VBA.CStr(VBA.CLng(filterDate)), _
                Operator:=xlAnd, _
                Criteria2:="<" & VBA.CStr(VBA.CLng(filterDate) + 1)
            ex_Core.fn_Diagnostic_WriteLog "DATE_FILTER_APPLIED | Field=" & _
                VBA.CStr(m_filterFieldIndex) & " | Date=" & _
                VBA.Format$(filterDate, "yyyy-mm-dd")
        Else
            m_filterTable.Range.AutoFilter Field:=m_filterFieldIndex, _
                Criteria1:="*" & query & "*"
            ex_Core.fn_Diagnostic_WriteLog "TEXT_FILTER_APPLIED | Field=" & _
                VBA.CStr(m_filterFieldIndex) & " | Query=" & query
        End If
    End If
    private_Filter_RestoreFilterInputCell
    m_filterHeaderCell.Select

CleanExit:
    Application.EnableEvents = previousEnableEvents
    private_Filter_ClearPendingFilterInput
    Exit Sub
EH:
    ex_Core.fn_Diagnostic_WriteLog "FILTER_ERROR | Number=" & VBA.CStr(VBA.Err.Number) & _
        " | Description=" & VBA.Err.Description
    private_Filter_RestoreFilterInputCell
    Resume CleanExit
End Sub

Private Function private_Filter_IsExcelDateSerial(ByVal query As String) As Boolean
    Dim serialValue As Double

    If Not VBA.IsNumeric(query) Then Exit Function
    serialValue = VBA.CDbl(query)
    private_Filter_IsExcelDateSerial = serialValue >= 1 And serialValue <= 2958465
End Function

Private Function private_Filter_IsDateFilterColumn( _
    ByVal tableObj As ListObject, _
    ByVal fieldIndex As Long _
) As Boolean
    Dim columnCell As Range
    Dim numberFormat As String

    If tableObj.DataBodyRange Is Nothing Then Exit Function
    For Each columnCell In tableObj.ListColumns(fieldIndex).DataBodyRange.Cells
        If Not VBA.IsError(columnCell.Value) Then
            If VBA.Len(VBA.CStr(columnCell.Value2)) > 0 Then
                numberFormat = VBA.Trim$(columnCell.NumberFormat)
                If VBA.StrComp(numberFormat, "General", VBA.vbTextCompare) = 0 Then _
                    Exit Function
                private_Filter_IsDateFilterColumn = VBA.IsDate(columnCell.Value)
                Exit Function
            End If
        End If
    Next columnCell
End Function

Private Sub private_Filter_CancelPendingFilterInput()
    If Not m_filterInputActive Then Exit Sub
    On Error Resume Next
    Application.EnableEvents = False
    private_Filter_RestoreFilterInputCell
    Application.EnableEvents = True
    private_Filter_ClearPendingFilterInput
End Sub

Private Sub private_Filter_RestoreFilterInputCell()
    If m_filterInputCell Is Nothing Then Exit Sub
    If m_filterInputHadFormula Then
        m_filterInputCell.Formula = m_filterInputFormula
    Else
        m_filterInputCell.Value2 = m_filterInputValue
    End If
    m_filterInputCell.Interior.Color = m_filterInputInteriorColor
    m_filterInputCell.Interior.Pattern = m_filterInputInteriorPattern
End Sub

Private Sub private_Filter_ClearPendingFilterInput()
    m_filterInputActive = False
    m_filterInputAddress = VBA.vbNullString
    m_filterInputHadFormula = False
    m_filterInputFormula = Empty
    m_filterInputValue = Empty
    m_filterInputInteriorColor = 0
    m_filterInputInteriorPattern = xlNone
    m_filterFieldIndex = 0
    Set m_filterInputCell = Nothing
    Set m_filterTable = Nothing
    Set m_filterHeaderCell = Nothing
End Sub

Private Function private_Filter_NormalizeFilterQuery(ByVal textValue As String) As String
    ' Normalize line breaks, non-breaking spaces, and repeated spaces.
    textValue = Replace(textValue, vbCr, " ")
    textValue = Replace(textValue, vbLf, " ")
    textValue = Replace(textValue, ChrW(160), " ")
    textValue = Trim(textValue)
    Do While InStr(textValue, "  ") > 0
        textValue = Replace(textValue, "  ", " ")
    Loop
    private_Filter_NormalizeFilterQuery = textValue
End Function
' --------------------------------------
' } // namespace Filter
' --------------------------------------