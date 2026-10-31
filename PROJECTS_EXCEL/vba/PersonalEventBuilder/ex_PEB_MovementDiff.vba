Attribute VB_Name = "ex_PEB_MovementDiff"
Option Explicit

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
Public Function fn_Build( _
    ByVal plan As Object, _
    ByVal target As ListObject, _
    ByVal entries As Collection _
) As Object
    Dim report As Object
    Dim rows As New Collection
    Dim links As Object
    Dim internalHeaders As Object
    Dim changedEvents As Object
    Dim action As Object
    Dim entry As Object
    Dim row As Object
    Dim key As Variant
    Dim added As Long
    Dim updated As Long
    Dim unchanged As Long
    Dim summary As String
    Dim originalRows As Long

    Set report = VBA.CreateObject("Scripting.Dictionary")
    originalRows = target.ListRows.Count
    Set links = ex_PEB_MovementLinks.fn_Headers()
    Set internalHeaders = VBA.CreateObject("Scripting.Dictionary")
    For Each key In links.Keys
        internalHeaders(VBA.CStr(links(key))) = True
    Next key
    Set changedEvents = VBA.CreateObject("Scripting.Dictionary")
    For Each entry In entries
        If Not internalHeaders.Exists(entry("Header")) Then
            If ex_PEB_MovementLinks.fn_ValueToken(entry("Value")) <> _
                    ex_PEB_MovementLinks.fn_ValueToken(entry("Expected")) Then
                Set action = entry("Action")
                Set row = VBA.CreateObject("Scripting.Dictionary")
                row("PersonName") = action("PersonName")
                row("EventRowId") = action("RowId")
                row("EventName") = action("EventName")
                row("Worksheet") = target.Parent.Name
                row("TargetRow") = VBA.CStr(entry("Row"))
                row("Field") = entry("Header")
                fn_CaptureUndo row, target, entry("TableRow"), entry("Value"), entry("Expected"), entry("Format"), action
                row("Before") = private_Text(entry("Value"), VBA.CStr(entry("Format")))
                row("After") = private_Text(entry("Expected"))
                row("Action") = ex_PEB_Export.fn_Config("movementDiff.updated")
                If entry("TableRow") > originalRows Then
                    row("Action") = ex_PEB_Export.fn_Config("movementDiff.added")
                End If
                rows.Add row
                changedEvents(action("RowId")) = True
            End If
        End If
    Next entry
    For Each action In plan("Actions")
        If Not changedEvents.Exists(action("RowId")) Then
            unchanged = unchanged + 1
        ElseIf action("IsUpdate") Or action("Operation") = "close" Then
            updated = updated + 1
        Else
            added = added + 1
        End If
    Next action
    summary = ex_PEB_Export.fn_Config("movementDiff.summary")
    summary = VBA.Replace$(summary, "{added}", VBA.CStr(added))
    summary = VBA.Replace$(summary, "{updated}", VBA.CStr(updated))
    summary = VBA.Replace$(summary, "{unchanged}", VBA.CStr(unchanged))
    summary = VBA.Replace$(summary, "{time}", VBA.Format$(VBA.Now, "dd.mm.yyyy hh:nn:ss"))
    report.Add "Summary", summary
    report.Add "Added", added
    report.Add "Updated", updated
    report.Add "Unchanged", unchanged
    report.Add "Target", target.Parent.Parent.FullName & " / " & target.Parent.Name & " / " & target.Name
    report.Add "AbsenceSheet", target.Parent.Name
    report.Add "TvoSheet", ex_PEB_Export.fn_Config("reporter.tvo.source.sheet")
    report.Add "Rows", rows
    Set report("Workbook") = target.Parent.Parent
    Set fn_Build = report
End Function

Public Function fn_Result( _
    ByVal report As Object, _
    ByRef result As obj_LookupResult, _
    ByRef diagnostic As String, _
    Optional ByVal worksheetName As String = "" _
) As Boolean
    Dim records As New Collection
    Dim columns As New Collection
    Dim row As Object
    Dim record As obj_LookupRecord
    Dim column As Object
    Dim key As Variant
    Dim index As Long

    For Each row In report("Rows")
        index = index + 1
        row("DiffId") = VBA.CStr(index)
        row("DiffTags") = VBA.vbNullString
        If row.Exists("UndoSelected") Then
            If row("UndoSelected") Then
                row("DiffTags") = "diffUndoPending"
            End If
        End If
        If row.Exists("UndoApplied") Then
            If row("UndoApplied") Then
                row("DiffTags") = "diffUndoApplied"
            End If
        End If
        If Len(worksheetName) > 0 Then
            If row("Worksheet") <> worksheetName Then
                GoTo NextDiffRow
            End If
        End If
        Set record = New obj_LookupRecord
        If Not record.Initialize(VBA.CStr(index), row) Then
            Exit Function
        End If
        records.Add record
NextDiffRow:
    Next row
    For Each key In VBA.Split(ex_PEB_Export.fn_Config("movementDiff.fields"), "|")
        Set column = VBA.CreateObject("Scripting.Dictionary")
        column.Add "name", key
        column.Add "caption", ex_PEB_Export.fn_Config("movementDiff.columns." & key)
        columns.Add column
    Next key
    Set result = New obj_LookupResult
    fn_Result = result.Initialize(records, columns, False, diagnostic, "DiffTags")
End Function

Public Sub fn_CaptureUndo( _
    ByVal row As Object, _
    ByVal table As ListObject, _
    ByVal tableRow As Long, _
    ByVal before As Variant, _
    ByVal after As Variant, _
    ByVal numberFormat As String, _
    ByVal action As Object _
)
    row("UndoTable") = table.Name
    row("UndoRow") = tableRow
    row("UndoBefore") = before
    row("UndoAfter") = after
    row("UndoFormat") = numberFormat
    row("UndoRevision") = action("Revision")
    row("UndoSide") = "Opening"
    If row("Worksheet") = ex_PEB_Export.fn_Config("reporter.tvo.source.sheet") Then
        row("UndoSide") = "TVO/" & VBA.CStr(tableRow)
    ElseIf action.Exists("ClosingIndex") Then
        If action("ClosingIndex") = tableRow Then
            row("UndoSide") = "Closing"
        End If
    End If
    row("UndoSelected") = False
    row("UndoApplied") = False
End Sub

Public Sub fn_Seal(ByVal report As Object)
    Dim row As Object
    Dim table As ListObject
    Dim stamps As Object
    Dim key As String

    Set stamps = VBA.CreateObject("Scripting.Dictionary")
    For Each row In report("Rows")
        Set table = report("Workbook").Worksheets(row("Worksheet")).ListObjects(row("UndoTable"))
        row("UndoStructure") = ex_PEB_MovementLinks.fn_StructureStamp(table)
        key = row("Worksheet") & "/" & row("UndoTable") & "/" & row("UndoRow")
        If Not stamps.Exists(key) Then
            stamps(key) = private_RowStamp(table.ListRows(row("UndoRow")).Range)
        End If
        row("UndoStamp") = stamps(key)
    Next row
End Sub

Public Function fn_Toggle( _
    ByVal report As Object, _
    ByVal selection As Range, _
    ByVal area As Range _
) As Boolean
    Dim table As ListObject
    Dim row As Object
    Dim index As Long
    Dim key As String

    If selection Is Nothing Or area Is Nothing Then
        Exit Function
    End If
    If selection.CountLarge <> 1 Then
        Exit Function
    End If
    If Not selection.Parent Is area.Parent Then
        Exit Function
    End If
    If Application.Intersect(selection, area) Is Nothing Then
        Exit Function
    End If
    For Each table In area.Parent.ListObjects
        If Not table.DataBodyRange Is Nothing Then
            If Not Application.Intersect(selection, table.DataBodyRange) Is Nothing Then
                index = selection.Row - table.DataBodyRange.Row + 1
                key = VBA.CStr(table.DataBodyRange.Cells(index, _
                    table.ListColumns(ex_PEB_Export.fn_Config("movementDiff.columns.DiffId")).Index).Value2)
                If Not VBA.IsNumeric(key) Then
                    Exit Function
                End If
                index = VBA.CLng(key)
                If index < 1 Or index > report("Rows").Count Then
                    Exit Function
                End If
                Set row = report("Rows")(index)
                If row("UndoApplied") Then
                    Exit Function
                End If
                row("UndoSelected") = Not row("UndoSelected")
                fn_Toggle = True
                Exit Function
            End If
        End If
    Next table
End Function

Public Function fn_SelectedCount(ByVal report As Object) As Long
    Dim row As Object

    For Each row In report("Rows")
        If row.Exists("UndoSelected") Then
            If row("UndoSelected") Then
                fn_SelectedCount = fn_SelectedCount + 1
            End If
        End If
    Next row
End Function

Public Function fn_ApplyUndo( _
    ByVal report As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim book As Workbook
    Dim candidate As Workbook
    Dim history As Workbook
    Dim journal As ListObject
    Dim table As ListObject
    Dim row As Object
    Dim cell As Range
    Dim entry As Object
    Dim entries As New Collection
    Dim count As Long
    Dim journalCount As Long
    Dim data(1 To 1, 1 To 12) As Variant
    Dim previousEvents As Boolean
    Dim previousScreen As Boolean
    Dim previousCalculation As XlCalculation
    Dim previousFill As Boolean
    Dim captured As Boolean
    Dim changed As Boolean
    Dim failure As String
    Dim rollbackError As String
    Dim settingsError As String
    Dim affected As Object
    Dim undoCount As Long
    Dim newStamps As Object
    Dim key As String
    Dim appliedText As String
    Dim successText As String
    Dim journalRow As ListRow
    Dim column As Long
    Dim filterTables As New Collection
    Dim filterStates As Collection
    Dim filterError As String

    On Error GoTo EH_UNDO
    diagnostic = VBA.vbNullString
    appliedText = ex_PEB_Export.fn_Config("movementDiff.undo.applied")
    successText = ex_PEB_Export.fn_Config("movementDiff.undo.success")
    If fn_SelectedCount(report) = 0 Then
        Exit Function
    End If
    Set book = report("Workbook")
    ' Require the original open workbook instance, including after close/reopen or Save As.
    For Each candidate In Application.Workbooks
        If candidate Is book Then
            count = count + 1
        End If
        If VBA.StrComp(candidate.FullName, ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("history.path")), VBA.vbTextCompare) = 0 Then
            Set history = candidate
        End If
    Next candidate
    If count <> 1 Or history Is Nothing Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("movementDiff.undo.unavailable")
    End If
    If book.ReadOnly Or book.Date1904 Or book.MultiUserEditing Or history.ReadOnly Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("movementDiff.undo.unavailable")
    End If
    Set journal = ex_PEB_MovementConflicts.fn_Table(history, "history.conflicts")
    journalCount = journal.ListRows.Count
    undoCount = fn_SelectedCount(report)
    private_CheckJournal journal, undoCount
    Set affected = VBA.CreateObject("Scripting.Dictionary")
    Set newStamps = VBA.CreateObject("Scripting.Dictionary")
    filterTables.Add journal
    ' Validate and capture every selected field before the first mutation.
    For Each row In report("Rows")
        If row("UndoSelected") Then
            Set table = book.Worksheets(row("Worksheet")).ListObjects(row("UndoTable"))
            filterTables.Add table
            If table.Parent.ProtectContents Or Not table.ShowHeaders Then
                Err.Raise 5, , ex_PEB_Export.fn_Config("movementDiff.undo.unavailable")
            End If
            If row("UndoStructure") <> ex_PEB_MovementLinks.fn_StructureStamp(table) Then
                Err.Raise 5, , ex_PEB_Export.fn_Config("movementDiff.undo.stale")
            End If
            If row("UndoStamp") <> private_RowStamp(table.ListRows(row("UndoRow")).Range) Then
                Err.Raise 5, , ex_PEB_Export.fn_Config("movementDiff.undo.stale")
            End If
            Set cell = table.DataBodyRange.Cells(row("UndoRow"), table.ListColumns(row("Field")).Index)
            If cell.HasFormula Or ex_PEB_MovementLinks.fn_ValueToken(cell.Value2) <> _
                    ex_PEB_MovementLinks.fn_ValueToken(row("UndoAfter")) Then
                Err.Raise 5, , ex_PEB_Export.fn_Config("movementDiff.undo.stale")
            End If
            Set entry = VBA.CreateObject("Scripting.Dictionary")
            Set entry("Cell") = cell
            Set entry("Row") = row
            entry("Value") = cell.Value2
            entry("Format") = cell.NumberFormat
            entries.Add entry
            affected(row("Worksheet") & "/" & row("UndoTable") & "/" & row("UndoRow")) = True
        End If
    Next row
    previousEvents = Application.EnableEvents
    previousScreen = Application.ScreenUpdating
    previousCalculation = Application.Calculation
    previousFill = Application.AutoCorrect.AutoFillFormulasInLists
    captured = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.AutoCorrect.AutoFillFormulasInLists = False
    If Not ex_TableFilters.fn_Suspend(filterTables, filterStates, filterError) Then
        Err.Raise 5, , filterError
    End If
    changed = True
    For Each entry In entries
        Set row = entry("Row")
        Set cell = entry("Cell")
        private_Write cell, row("UndoBefore"), entry("Format")
        data(1, 1) = row("EventRowId") & "/DiffRevert/" & VBA.CStr(journal.ListRows.Count + 1)
        data(1, 2) = row("EventRowId")
        data(1, 3) = row("UndoRevision")
        data(1, 4) = "Movement"
        data(1, 5) = book.FullName
        data(1, 6) = row("UndoSide")
        data(1, 7) = row("Field")
        data(1, 8) = ex_PEB_MovementLinks.fn_ValueToken(row("UndoAfter"))
        data(1, 9) = ex_PEB_MovementLinks.fn_ValueToken(row("UndoBefore"))
        data(1, 10) = "DiffRevert"
        data(1, 11) = VBA.Format$(VBA.Now, "yyyy-mm-dd hh:nn:ss")
        data(1, 12) = Application.UserName & " / " & VBA.Environ$("USERNAME")
        Set journalRow = journal.ListRows.Add(AlwaysInsert:=False)
        With journalRow.Range
            .NumberFormat = "@"
            .Value2 = data
        End With
        For column = 1 To 12
            If ex_PEB_MovementLinks.fn_ValueToken(journalRow.Range.Cells(1, column).Value2) <> _
                    ex_PEB_MovementLinks.fn_ValueToken(data(1, column)) Then
                Err.Raise 5, , ex_PEB_Export.fn_Config("movementDiff.undo.stale")
            End If
        Next column
    Next entry
    For Each entry In entries
        Set row = entry("Row")
        If ex_PEB_MovementLinks.fn_ValueToken(entry("Cell").Value2) <> ex_PEB_MovementLinks.fn_ValueToken(row("UndoBefore")) Then
            Err.Raise 5, , ex_PEB_Export.fn_Config("movementDiff.undo.stale")
        End If
    Next entry
    ' Refresh stamps for remaining fields only after successful writes and journal recording.
    For Each row In report("Rows")
        key = row("Worksheet") & "/" & row("UndoTable") & "/" & row("UndoRow")
        If affected.Exists(key) And Not newStamps.Exists(key) Then
            Set table = book.Worksheets(row("Worksheet")).ListObjects(row("UndoTable"))
            newStamps(key) = private_RowStamp(table.ListRows(row("UndoRow")).Range)
        End If
    Next row
    For Each row In report("Rows")
        key = row("Worksheet") & "/" & row("UndoTable") & "/" & row("UndoRow")
        If newStamps.Exists(key) Then
            row("UndoStamp") = newStamps(key)
        End If
    Next row
    For Each entry In entries
        Set row = entry("Row")
        row("UndoSelected") = False
        row("UndoApplied") = True
        row("Action") = appliedText
    Next entry
    fn_ApplyUndo = True
    diagnostic = VBA.Replace$(successText, "{count}", VBA.CStr(entries.Count))
CleanUndo:
    filterError = ex_TableFilters.fn_Restore(filterStates)
    If Len(filterError) > 0 Then
        diagnostic = diagnostic & vbCrLf & Replace$( _
            ex_PEB_Export.fn_Config("text.HistoryFilterRestoreFailed"), "{details}", filterError)
    End If
    If captured Then
        On Error Resume Next
        VBA.Err.Clear
        Application.AutoCorrect.AutoFillFormulasInLists = previousFill
        If VBA.Err.Number <> 0 Then
            settingsError = settingsError & " " & VBA.Err.Description
        End If
        VBA.Err.Clear
        Application.Calculation = previousCalculation
        If VBA.Err.Number <> 0 Then
            settingsError = settingsError & " " & VBA.Err.Description
        End If
        VBA.Err.Clear
        Application.EnableEvents = previousEvents
        If VBA.Err.Number <> 0 Then
            settingsError = settingsError & " " & VBA.Err.Description
        End If
        VBA.Err.Clear
        Application.ScreenUpdating = previousScreen
        If VBA.Err.Number <> 0 Then
            settingsError = settingsError & " " & VBA.Err.Description
        End If
        On Error GoTo 0
    End If
    If VBA.Len(settingsError) > 0 Then
        diagnostic = diagnostic & VBA.vbCrLf & VBA.Replace$( _
            ex_PEB_Export.fn_Config("text.MovementBatchSettingsError"), "{error}", settingsError)
    End If
    Exit Function
EH_UNDO:
    failure = VBA.Err.Description
    If changed Then
        On Error Resume Next
        For Each entry In entries
            VBA.Err.Clear
            private_Write entry("Cell"), entry("Value"), entry("Format")
            If VBA.Err.Number <> 0 Then
                rollbackError = rollbackError & " " & VBA.Err.Description
            End If
        Next entry
        Do While journal.ListRows.Count > journalCount
            VBA.Err.Clear
            journal.ListRows(journal.ListRows.Count).Delete
            If VBA.Err.Number <> 0 Then
                rollbackError = rollbackError & " " & VBA.Err.Description
                Exit Do
            End If
        Loop
        On Error GoTo 0
    End If
    diagnostic = VBA.Replace$(ex_PEB_Export.fn_Config("movementDiff.undo.failed"), "{details}", failure)
    If VBA.Len(rollbackError) > 0 Then
        diagnostic = diagnostic & VBA.vbCrLf & VBA.Replace$( _
            ex_PEB_Export.fn_Config("text.MovementBatchRollbackError"), "{error}", rollbackError)
    End If
    GoTo CleanUndo
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Sub private_CheckJournal( _
    ByVal table As ListObject, _
    ByVal count As Long _
)
    Dim area As Range
    Dim other As ListObject

    If table.ShowTotals Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("movementDiff.undo.unavailable")
    End If
    Set area = table.Range.Rows(table.Range.Rows.Count).Offset(1).Resize(count)
    If VBA.IsNull(area.MergeCells) Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("movementDiff.undo.unavailable")
    End If
    If area.MergeCells Or Application.WorksheetFunction.CountA(area) <> 0 Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("movementDiff.undo.unavailable")
    End If
    For Each other In table.Parent.ListObjects
        If Not other Is table Then
            If Not Application.Intersect(area, other.Range) Is Nothing Then
                Err.Raise 5, , ex_PEB_Export.fn_Config("movementDiff.undo.unavailable")
            End If
        End If
    Next other
End Sub

Private Function private_RowStamp(ByVal area As Range) As String
    Dim values As Variant
    Dim formulas As Variant
    Dim index As Long
    Dim stamp As String

    values = area.Value2
    formulas = area.Formula
    If area.Columns.Count = 1 Then
        stamp = private_CellStamp(values, formulas)
    Else
        For index = 1 To area.Columns.Count
            stamp = stamp & private_CellStamp(values(1, index), formulas(1, index))
        Next index
    End If
    private_RowStamp = ex_PEB_MovementLinks.fn_Hash(stamp)
End Function

Private Function private_CellStamp( _
    ByVal value As Variant, _
    ByVal formula As Variant _
) As String
    Dim token As String

    ' Calculated results may change during our own reversal; fingerprint the formula instead.
    If VBA.VarType(formula) = VBA.vbString Then
        If VBA.Left$(VBA.CStr(formula), 1) = "=" Then
            private_CellStamp = ex_PEB_MovementLinks.fn_Token("f:" & VBA.CStr(formula))
            Exit Function
        End If
    End If
    If VBA.IsError(value) Then
        token = "e:" & VBA.CStr(value)
    Else
        token = ex_PEB_MovementLinks.fn_ValueToken(value)
    End If
    private_CellStamp = ex_PEB_MovementLinks.fn_Token(token) & ex_PEB_MovementLinks.fn_Token(VBA.CStr(formula))
End Function

Private Sub private_Write( _
    ByVal cell As Range, _
    ByVal value As Variant, _
    ByVal numberFormat As String _
)
    cell.NumberFormat = "@"
    cell.Value2 = value
    cell.NumberFormat = numberFormat
End Sub

Private Function private_Text( _
    ByVal value As Variant, _
    Optional ByVal numberFormat As String = "" _
) As String
    If VBA.IsError(value) Then
        private_Text = ex_PEB_Export.fn_Config("movementDiff.error")
    ElseIf VBA.IsNull(value) Or VBA.IsEmpty(value) Then
        private_Text = ex_PEB_Export.fn_Config("movementDiff.empty")
    ElseIf VBA.Len(VBA.CStr(value)) = 0 Then
        private_Text = ex_PEB_Export.fn_Config("movementDiff.empty")
    ElseIf VBA.VarType(value) = VBA.vbDate Then
        private_Text = VBA.Format$(value, ex_PEB_Export.fn_Config("export.Movement.dateFormat"))
    ElseIf VBA.IsNumeric(value) And (VBA.InStr(1, numberFormat, "yy", VBA.vbTextCompare) > 0 Or _
            VBA.InStr(1, numberFormat, "dd", VBA.vbTextCompare) > 0) Then
        private_Text = VBA.Format$(VBA.CDate(value), ex_PEB_Export.fn_Config("export.Movement.dateFormat"))
    Else
        private_Text = VBA.CStr(value)
    End If
End Function
' --------------------------------------
' } // namespace Private
' --------------------------------------