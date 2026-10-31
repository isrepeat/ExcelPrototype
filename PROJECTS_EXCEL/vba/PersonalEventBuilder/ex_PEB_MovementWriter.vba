Attribute VB_Name = "ex_PEB_MovementWriter"
Option Explicit

' Keep the temporary switch for checking row sizing independently of export formatting.
#Const MOVEMENT_AUTOFIT_ENABLED = True

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Apply( _
    ByVal plan As Object, _
    ByVal target As ListObject, _
    ByRef diagnostic As String, _
    Optional ByRef exportDiff As Object _
) As Boolean
    Dim entries As Collection
    Dim links As Object
    Dim action As Object
    Dim entry As Object
    Dim update As Object
    Dim cell As Range
    Dim key As Variant
    Dim fields As Object
    Dim actual As Object
    Dim previousEvents As Boolean
    Dim previousScreen As Boolean
    Dim previousWindow As Window
    Dim previousSheet As Object
    Dim previousAutoFill As Boolean
    Dim previousCalculation As XlCalculation
    Dim settingsCaptured As Boolean
    Dim changed As Boolean
    Dim failure As String
    Dim rollbackError As String
    Dim settingsTemplate As String
    Dim rollbackTemplate As String
    Dim restoredText As String
    Dim stage As String
    Dim startedAt As Double
    Dim totalStartedAt As Double
    Dim originalAddress As String
    Dim extension As Range
    Dim extensionSnapshot As Object
    Dim addedRows As Long
    Dim pendingDiff As Object
    Dim tvoRollbackError As String
    Dim filterStates As Collection
    Dim filterError As String

    On Error GoTo EH
    totalStartedAt = VBA.Timer
    diagnostic = VBA.vbNullString
    Set exportDiff = Nothing
    Set entries = New Collection
    private_Stage "Write.Validate", stage, startedAt
    If plan Is Nothing Then
        ex_PEB_MovementLinks.fn_Fail "MissingPlan", ""
    End If
    If Not plan("IsReady") Then
        ex_PEB_MovementLinks.fn_Fail "InvalidMovementPlan", target.Name
    End If
    If plan("TargetStamp") <> ex_PEB_MovementLinks.fn_StructureStamp(target) Or _
            plan("PlanStamp") <> ex_PEB_MovementLinks.fn_PlanStamp(plan) Then
        ex_PEB_MovementLinks.fn_Fail "PlanChangedSincePreview", target.Name
    End If
    settingsTemplate = ex_PEB_Export.fn_Config("text.MovementBatchSettingsError")
    rollbackTemplate = ex_PEB_Export.fn_Config("text.MovementBatchRollbackError")
    restoredText = ex_PEB_Export.fn_Config("text.MovementBatchRestored")
    Set links = ex_PEB_MovementLinks.fn_Headers()
    For Each key In links.Keys
        Set cell = target.ListColumns(links(key)).Range.Cells(1, 1)
    Next key

    ' The caller builds this plan immediately before applying it, without yielding.
    private_Stage "Write.Capture", stage, startedAt
    originalAddress = target.Range.Address
    For Each action In plan("Actions")
        If action("Status") = "Planned" Then
            If action.Exists("PlaceholderWrites") Then
                addedRows = addedRows + 1
            End If
            If action("Operation") <> "close" Then
                If Not action("IsUpdate") Or action("TargetIndex") > target.ListRows.Count Then
                    addedRows = addedRows + 1
                End If
            End If
        End If
    Next action
    If addedRows > 0 Then
        Set extension = private_Extension(target, addedRows)
        Set extensionSnapshot = ex_PEB_MovementLinks.fn_ReadRange(extension)
    End If
    Set actual = VBA.CreateObject("Scripting.Dictionary")
    For Each action In plan("Actions")
        If action("Status") = "Planned" And action.Exists("PlaceholderWrites") Then
            Set fields = action("PlaceholderWrites")
            For Each key In fields.Keys
                private_Capture target, action("PlaceholderIndex"), VBA.CStr(key), fields(key), entries, actual, action
            Next key
        End If
        If action("Status") = "Planned" And action("IsUpdate") Then
            For Each update In action("UpdateCells")
                private_Capture target, update("RowIndex"), update("Header"), update("After"), _
                    entries, actual, action, True, update("Before")
            Next update
        ElseIf action("Status") = "Planned" Then
            If action("Operation") <> "append" Then
                Set fields = action("ClosingWrites")
                For Each key In fields.Keys
                    private_Capture target, action("ClosingIndex"), VBA.CStr(key), fields(key), entries, actual, action
                Next key
                private_Capture target, action("ClosingIndex"), links("ClosingId"), action("RowId"), entries, actual, action
                private_Capture target, action("ClosingIndex"), links("ClosingHash"), action("ClosingHash"), entries, actual, action
            End If
            If action("Operation") <> "close" Then
                Set fields = action("Writes")
                For Each key In fields.Keys
                    private_Capture target, action("TargetIndex"), VBA.CStr(key), fields(key), entries, actual, action
                Next key
                private_Capture target, action("TargetIndex"), links("OpeningId"), action("RowId"), entries, actual, action
                private_Capture target, action("TargetIndex"), links("OpeningHash"), action("OpeningHash"), entries, actual, action
            End If
        ElseIf action("Status") <> "AlreadyApplied" Then
            ex_PEB_MovementLinks.fn_Fail "InvalidMovementPlan", action("RowId")
        End If
    Next action
    Set pendingDiff = ex_PEB_MovementDiff.fn_Build(plan, target, entries)
    ex_PEB_MovementTvo.fn_AddDiff pendingDiff, plan
    previousEvents = Application.EnableEvents
    previousScreen = Application.ScreenUpdating
    Set previousWindow = Application.ActiveWindow
    Set previousSheet = Application.ActiveSheet
    previousAutoFill = Application.AutoCorrect.AutoFillFormulasInLists
    previousCalculation = Application.Calculation
    settingsCaptured = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.AutoCorrect.AutoFillFormulasInLists = False
    private_SuspendFilters plan, target, filterStates
    If entries.Count = 0 And ex_PEB_MovementTvo.fn_ChangeCount(plan) = 0 Then
        ex_PEB_MovementDiff.fn_Seal pendingDiff
        If Not ex_PEB_HistoryRevisions.fn_RecordExport(plan("Actions"), target.Parent.Parent.FullName, "Movement", diagnostic) Then
            GoTo Cleanup
        End If
        fn_Apply = True
        GoTo Cleanup
    End If

    private_Stage "Write.Write", stage, startedAt
    changed = True
    If addedRows > 0 Then
        target.Resize target.Range.Resize(target.Range.Rows.Count + addedRows)
    End If
    private_Stage "Write.FormatText", stage, startedAt
    private_FormatEntries target, entries, True
    private_Stage "Write.Values", stage, startedAt
    private_WriteValues target, entries
    private_Stage "Write.RestoreFormats", stage, startedAt
    private_FormatEntries target, entries, False
    private_Stage "Write.Verify", stage, startedAt
    For Each entry In entries
        private_Verify target, entry, False
    Next entry
    ex_PEB_MovementTvo.fn_Apply plan
#If MOVEMENT_AUTOFIT_ENABLED Then
    private_Stage "Write.AutoFitRows", stage, startedAt
    private_AutoFitRows target, entries
#End If
    private_Stage "Write.Journal", stage, startedAt
    ex_PEB_MovementDiff.fn_Seal pendingDiff
    If Not ex_PEB_HistoryRevisions.fn_RecordExport(plan("Actions"), target.Parent.Parent.FullName, "Movement", diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    fn_Apply = True
Cleanup:
    private_Stage "Write.RestoreSettings", stage, startedAt
    filterError = ex_TableFilters.fn_Restore(filterStates)
    If Len(filterError) > 0 Then
        diagnostic = diagnostic & VBA.vbCrLf & Replace$( _
            ex_PEB_Export.fn_Config("text.HistoryFilterRestoreFailed"), "{details}", filterError)
    End If
    If settingsCaptured Then
        On Error Resume Next
        VBA.Err.Clear
        Application.AutoCorrect.AutoFillFormulasInLists = previousAutoFill
        If VBA.Err.Number <> 0 Then
            failure = failure & " AutoFill: " & VBA.Err.Description
        End If
        VBA.Err.Clear
        Application.Calculation = previousCalculation
        If VBA.Err.Number <> 0 Then
            failure = failure & " Calculation: " & VBA.Err.Description
        End If
        VBA.Err.Clear
        ' Restore focus while events and screen updates are still suspended.
        If Not previousWindow Is Nothing Then
            If Not Application.ActiveWindow Is previousWindow Then
                previousWindow.Activate
            End If
        End If
        If Not previousSheet Is Nothing Then
            If Not Application.ActiveSheet Is previousSheet Then
                previousSheet.Activate
            End If
        End If
        If VBA.Err.Number <> 0 Then
            failure = failure & " Focus: " & VBA.Err.Description
        End If
        VBA.Err.Clear
        Application.EnableEvents = previousEvents
        If VBA.Err.Number <> 0 Then
            failure = failure & " EnableEvents: " & VBA.Err.Description
        End If
        VBA.Err.Clear
        Application.ScreenUpdating = previousScreen
        If VBA.Err.Number <> 0 Then
            failure = failure & " ScreenUpdating: " & VBA.Err.Description
        End If
        On Error GoTo 0
    End If
    If VBA.Len(failure) > 0 Then
        diagnostic = diagnostic & VBA.vbCrLf & VBA.Replace$(settingsTemplate, "{error}", failure)
    End If
    If fn_Apply Then
        Set exportDiff = pendingDiff
    End If
    private_Stage VBA.vbNullString, stage, startedAt
    ex_Core.fn_Diagnostic_WritePerf "Movement.Write.TotalIncludingCalculation", totalStartedAt
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    ex_Core.fn_Diagnostic_WriteError "Movement." & stage, VBA.Err.Number, VBA.Err.Source, diagnostic
    If changed Then
        private_Stage "Write.Rollback", stage, startedAt
        tvoRollbackError = ex_PEB_MovementTvo.fn_Restore(plan)
        If private_Restore(target, entries, originalAddress, extensionSnapshot, rollbackError) Then
            If Len(tvoRollbackError) = 0 Then
                diagnostic = diagnostic & VBA.vbCrLf & restoredText
            Else
                diagnostic = diagnostic & VBA.vbCrLf & Replace$(rollbackTemplate, "{error}", tvoRollbackError)
            End If
        Else
            diagnostic = diagnostic & VBA.vbCrLf & VBA.Replace$(rollbackTemplate, "{error}", rollbackError & tvoRollbackError)
        End If
    End If
    Resume Cleanup
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Sub private_SuspendFilters( _
    ByVal plan As Object, _
    ByVal target As ListObject, _
    ByRef states As Collection _
)
    Dim tables As New Collection
    Dim book As Workbook
    Dim sheet As Worksheet
    Dim table As ListObject
    Dim diagnostic As String
    Dim historyPath As String

    tables.Add target
    tables.Add plan("TvoPlan")("Table")
    historyPath = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("history.path"))
    For Each book In Application.Workbooks
        If StrComp(book.FullName, historyPath, vbTextCompare) = 0 Then
            For Each sheet In book.Worksheets
                For Each table In sheet.ListObjects
                    tables.Add table
                Next table
            Next sheet
        End If
    Next book
    If Not ex_TableFilters.fn_Suspend(tables, states, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
End Sub

Private Function private_Extension( _
    ByVal target As ListObject, _
    ByVal addedRows As Long _
) As Range
    Dim area As Range
    Dim other As ListObject
    Dim merged As Variant

    If target.ShowTotals Then
        ex_PEB_MovementLinks.fn_Fail "AppendWithTotalsUnsupported", target.Name
    End If
    If target.Range.Row + target.Range.Rows.Count + addedRows - 1 > target.Parent.Rows.Count Then
        ex_PEB_MovementLinks.fn_Fail "InsufficientWorksheetRows", target.Name
    End If
    Set area = target.Range.Rows(target.Range.Rows.Count).Offset(1).Resize(addedRows)
    merged = area.MergeCells
    If VBA.IsNull(merged) Then
        ex_PEB_MovementLinks.fn_Fail "MergedExpansionCell", area.Address
    End If
    If VBA.CBool(merged) Or Application.WorksheetFunction.CountA(area) <> 0 Then
        ex_PEB_MovementLinks.fn_Fail "OccupiedExpansionCell", area.Address
    End If
    For Each other In target.Parent.ListObjects
        If Not other Is target Then
            If Not Application.Intersect(area, other.Range) Is Nothing Then
                ex_PEB_MovementLinks.fn_Fail "ExpansionIntersectsTable", other.Name
            End If
        End If
    Next other
    Set private_Extension = area
End Function

Private Sub private_Capture( _
    ByVal target As ListObject, _
    ByVal rowIndex As Long, _
    ByVal header As String, _
    ByVal expected As Variant, _
    ByVal entries As Collection, _
    ByVal seen As Object, _
    ByVal action As Object, _
    Optional ByVal replacing As Boolean = False, _
    Optional ByVal original As Variant _
)
    Dim cell As Range
    Dim entry As Object

    Set cell = target.HeaderRowRange.Cells(1, target.ListColumns(header).Index).Offset(rowIndex)
    If seen.Exists(cell.Address) Then
        Set entry = seen(cell.Address)
        If replacing Then
            If ex_PEB_MovementLinks.fn_ValueToken(entry("Expected")) <> ex_PEB_MovementLinks.fn_ValueToken(original) Then
                ex_PEB_MovementLinks.fn_Fail "ConflictingPlannedWrite", cell.Address
            End If
        ElseIf VBA.Len(ex_PEB_Export.fn_Text(entry("Expected"))) > 0 Then
            ex_PEB_MovementLinks.fn_Fail "ConflictingPlannedWrite", cell.Address
        End If
        entry("Expected") = expected
        Set entry("Action") = action
        Exit Sub
    End If
    If cell.HasFormula Then
        ex_PEB_MovementLinks.fn_Fail "FormulaOverwrite", header
    End If
    If replacing Then
        If ex_PEB_MovementLinks.fn_ValueToken(cell.Value2) <> ex_PEB_MovementLinks.fn_ValueToken(original) Then
            ex_PEB_MovementLinks.fn_Fail "TargetChangedBeforeWrite", header
        End If
    ElseIf VBA.Len(ex_PEB_Export.fn_Text(cell.Value2)) > 0 Then
        ex_PEB_MovementLinks.fn_Fail "PopulatedCellOverwrite", header
    End If
    Set entry = VBA.CreateObject("Scripting.Dictionary")
    entry.Add "Address", cell.Address
    entry.Add "Row", cell.Row
    entry.Add "Column", cell.Column
    entry.Add "Value", cell.Value2
    entry.Add "Format", cell.NumberFormat
    entry.Add "WriteFormat", cell.NumberFormat
#If MOVEMENT_AUTOFIT_ENABLED Then
    entry.Add "RowHeight", cell.EntireRow.RowHeight
#End If
    entry.Add "Expected", expected
    entry.Add "Header", header
    entry.Add "TableRow", rowIndex
    entry.Add "Action", action
    entries.Add entry
    seen.Add cell.Address, entry
End Sub

Private Sub private_WriteValues( _
    ByVal target As ListObject, _
    ByVal entries As Collection _
)
    Dim columns As Object
    Dim rows As Object
    Dim entry As Object
    Dim column As Variant
    Dim rowIds As Variant
    Dim rowId As Variant
    Dim values() As Variant
    Dim value As Variant
    Dim index As Long
    Dim cursor As Long
    Dim first As Long
    Dim last As Long
    Dim count As Long
    Dim writeCount As Long

    Set columns = VBA.CreateObject("Scripting.Dictionary")
    For Each entry In entries
        column = CStr(entry("Column"))
        If Not columns.Exists(column) Then Set columns(column) = VBA.CreateObject("Scripting.Dictionary")
        Set rows = columns(column)
        Set rows(CStr(entry("Row"))) = entry
    Next entry

    ' Write only contiguous runs of planned cells; untouched cells and formulas stay outside each range.
    For Each column In columns.Keys
        Set rows = columns(column)
        rowIds = rows.Keys
        For index = LBound(rowIds) + 1 To UBound(rowIds)
            rowId = rowIds(index)
            cursor = index - 1
            Do While cursor >= LBound(rowIds)
                If CLng(rowIds(cursor)) <= CLng(rowId) Then Exit Do
                rowIds(cursor + 1) = rowIds(cursor)
                cursor = cursor - 1
            Loop
            rowIds(cursor + 1) = rowId
        Next index
        first = LBound(rowIds)
        Do While first <= UBound(rowIds)
            last = first
            Do While last < UBound(rowIds)
                If CLng(rowIds(last + 1)) <> CLng(rowIds(last)) + 1 Then Exit Do
                last = last + 1
            Loop
            count = last - first + 1
            ReDim values(1 To count, 1 To 1)
            For index = 1 To count
                Set entry = rows(rowIds(first + index - 1))
                value = entry("Expected")
                If VBA.VarType(value) = VBA.vbDate Then value = VBA.CDbl(value)
                values(index, 1) = value
            Next index
            target.Parent.Cells(CLng(rowIds(first)), CLng(column)).Resize(count, 1).Value2 = values
            writeCount = writeCount + 1
            first = last + 1
        Loop
    Next column
    ex_Core.fn_Diagnostic_WriteLog "MOVEMENT_VALUE_BATCH | Cells=" & CStr(entries.Count) & " | Writes=" & CStr(writeCount)
End Sub

#If MOVEMENT_AUTOFIT_ENABLED Then
Private Sub private_AutoFitRows( _
    ByVal target As ListObject, _
    ByVal entries As Collection _
)
    Dim rows As Object
    Dim entry As Object
    Dim cell As Range
    Dim key As Variant

    ' Fit only changed rows after restoring their final number formats.
    Set rows = VBA.CreateObject("Scripting.Dictionary")
    For Each entry In entries
        Set cell = target.Parent.Range(entry("Address"))
        rows(CStr(cell.Row)) = True
    Next entry
    For Each key In rows.Keys
        target.Parent.Rows(CLng(key)).AutoFit
    Next key
End Sub
#End If

Private Sub private_FormatEntries( _
    ByVal target As ListObject, _
    ByVal entries As Collection, _
    ByVal textOnly As Boolean _
)
    Dim groups As Object
    Dim entry As Object
    Dim format As String
    Dim addresses As String
    Dim key As Variant
    Dim cell As Range

    If Not textOnly Then
        ' Restore each cell's captured format instead of applying formats to disjoint areas.
        For Each entry In entries
            Set cell = target.Parent.Range(entry("Address"))
            cell.NumberFormat = entry("WriteFormat")
        Next entry
        Exit Sub
    End If
    Set groups = VBA.CreateObject("Scripting.Dictionary")
    For Each entry In entries
        Set cell = target.Parent.Range(entry("Address"))
        entry("WriteFormat") = cell.NumberFormat
    Next entry
    For Each entry In entries
        If Not textOnly Or VBA.VarType(entry("Expected")) = VBA.vbString Then
            format = entry("WriteFormat")
            If textOnly Then format = "@"
            If Not groups.Exists(format) Then groups.Add format, VBA.vbNullString
            addresses = groups(format)
            ' Keep each multi-area address below Excel's Range argument limit.
            If VBA.Len(addresses) + VBA.Len(entry("Address")) + 1 > 200 Then
                target.Parent.Range(addresses).NumberFormat = format
                addresses = VBA.vbNullString
            End If
            If VBA.Len(addresses) > 0 Then addresses = addresses & ","
            groups(format) = addresses & entry("Address")
        End If
    Next entry
    For Each key In groups.Keys
        If VBA.Len(groups(key)) > 0 Then target.Parent.Range(groups(key)).NumberFormat = VBA.CStr(key)
    Next key
End Sub

Private Sub private_Verify( _
    ByVal target As ListObject, _
    ByVal entry As Object, _
    ByVal restored As Boolean _
)
    Dim cell As Range
    Dim expected As Variant

    Set cell = target.Parent.Range(entry("Address"))
    If cell.HasFormula Then
        private_VerificationLog target, entry, restored, "Formula", False, cell.Formula
        ex_PEB_MovementLinks.fn_Fail "WrittenPlanVerificationFailed", cell.Address
    End If
    If restored Then
        expected = entry("Value")
    Else
        expected = entry("Expected")
    End If
    If ex_PEB_MovementLinks.fn_ValueToken(expected) <> ex_PEB_MovementLinks.fn_ValueToken(cell.Value2) Then
        private_VerificationLog target, entry, restored, "Value", expected, cell.Value2
        ex_PEB_MovementLinks.fn_Fail "WrittenPlanVerificationFailed", cell.Address
    End If
End Sub

Private Sub private_VerificationLog( _
    ByVal target As ListObject, _
    ByVal entry As Object, _
    ByVal restored As Boolean, _
    ByVal kind As String, _
    ByVal expected As Variant, _
    ByVal actual As Variant _
)
    Dim action As Object
    Dim expectedText As String
    Dim actualText As String

    Set action = entry("Action")
    expectedText = Replace$(Replace$(CStr(expected), vbCr, "\r"), vbLf, "\n")
    actualText = Replace$(Replace$(CStr(actual), vbCr, "\r"), vbLf, "\n")
    ex_Core.fn_Diagnostic_WriteLog "MOVEMENT_VERIFY_MISMATCH | Kind=" & kind & _
        " | Restored=" & CStr(restored) & " | Sheet=" & target.Parent.Name & _
        " | Cell=" & CStr(entry("Address")) & " | Header=" & CStr(entry("Header")) & _
        " | RowId=" & CStr(action("RowId")) & " | Person=" & CStr(action("PersonName")) & _
        " | ExpectedType=" & CStr(VarType(expected)) & " | ActualType=" & CStr(VarType(actual)) & _
        " | ExpectedLength=" & CStr(Len(CStr(expected))) & " | ActualLength=" & CStr(Len(CStr(actual))) & _
        " | Expected=" & expectedText & " | Actual=" & actualText
End Sub

Private Function private_Restore( _
    ByVal target As ListObject, _
    ByVal entries As Collection, _
    ByVal originalAddress As String, _
    ByVal extensionSnapshot As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim entry As Object
    Dim cell As Range
    Dim area As Range
    Dim formats As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long

    ' Attempt every cell even when one restoration fails.
    On Error Resume Next
    target.Resize target.Parent.Range(originalAddress)
    If VBA.Err.Number <> 0 Then
        diagnostic = diagnostic & " " & VBA.Err.Description
    End If
    For Each entry In entries
        VBA.Err.Clear
        Set cell = target.Parent.Range(entry("Address"))
        cell.NumberFormat = "@"
        cell.Value2 = entry("Value")
        cell.NumberFormat = entry("Format")
#If MOVEMENT_AUTOFIT_ENABLED Then
        cell.EntireRow.RowHeight = entry("RowHeight")
#End If
        If VBA.Err.Number <> 0 Then
            diagnostic = diagnostic & " " & entry("Address") & ": " & VBA.Err.Description
        End If
    Next entry
    For Each entry In entries
        VBA.Err.Clear
        private_Verify target, entry, True
        If VBA.Err.Number <> 0 Then
            diagnostic = diagnostic & " " & entry("Address") & ": " & VBA.Err.Description
        End If
    Next entry
    If Not extensionSnapshot Is Nothing Then
        VBA.Err.Clear
        Set area = target.Parent.Range(extensionSnapshot("Address"))
        area.ClearContents
        formats = extensionSnapshot("Formats")
        If VBA.Err.Number <> 0 Then
            diagnostic = diagnostic & " " & VBA.Err.Description
        End If
        For rowIndex = 1 To area.Rows.Count
            For columnIndex = 1 To area.Columns.Count
                VBA.Err.Clear
                area.Cells(rowIndex, columnIndex).NumberFormat = formats(rowIndex, columnIndex)
                If VBA.Err.Number <> 0 Then
                    diagnostic = diagnostic & " " & VBA.Err.Description
                End If
            Next columnIndex
        Next rowIndex
    End If
    On Error GoTo 0
    private_Restore = (VBA.Len(diagnostic) = 0)
End Function

Private Sub private_Stage( _
    ByVal nextStage As String, _
    ByRef stage As String, _
    ByRef startedAt As Double _
)
    If VBA.Len(stage) > 0 Then
        ex_Core.fn_Diagnostic_WritePerf "Movement." & stage, startedAt
    End If
    stage = nextStage
    startedAt = VBA.Timer
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------