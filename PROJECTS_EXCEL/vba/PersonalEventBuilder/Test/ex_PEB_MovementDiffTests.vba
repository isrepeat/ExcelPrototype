Attribute VB_Name = "ex_PEB_MovementDiffTests"
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
Public Function fn_Run(ByRef diagnostic As String) As Boolean
    Dim book As Workbook
    Dim previousBook As Workbook
    Dim target As ListObject
    Dim plan As Object
    Dim actions As New Collection
    Dim entries As New Collection
    Dim changed As Object
    Dim added As Object
    Dim internalOnly As Object
    Dim links As Object
    Dim report As Object
    Dim rows As Collection
    Dim result As obj_LookupResult
    Dim display As ListObject
    Dim sheet As Worksheet
    Dim tags As obj_IUiRowTags

    On Error GoTo EH
    Set previousBook = Application.ActiveWorkbook
    Set book = Application.Workbooks.Add
    book.Worksheets(1).Range("A1").Value2 = "Fixture"
    Set target = book.Worksheets(1).ListObjects.Add(1, book.Worksheets(1).Range("A1:A2"), , 1)
    Set changed = private_Action("Changed", True, "Planned")
    Set added = private_Action("Added", False, "Planned")
    Set internalOnly = private_Action("Internal", True, "Planned")
    actions.Add changed
    actions.Add added
    actions.Add private_Action("Confirmed", False, "AlreadyApplied")
    actions.Add internalOnly
    entries.Add private_Entry(changed, "Destination", "Manual value", "Generated value", 1)
    entries.Add private_Entry(changed, "Unchanged", "Same", "Same", 1)
    entries.Add private_Entry(changed, "Cleared", "Previous value", VBA.vbNullString, 1)
    entries.Add private_Entry(added, "Person", Empty, "New person", 2)
    Set links = ex_PEB_MovementLinks.fn_Headers()
    entries.Add private_Entry(internalOnly, links("OpeningHash"), "Old hash", "New hash", 1)
    Set plan = VBA.CreateObject("Scripting.Dictionary")
    plan.Add "Actions", actions
    Set report = ex_PEB_MovementDiff.fn_Build(plan, target, entries)
    Set rows = report("Rows")
    If rows.Count <> 3 Or report("Added") <> 1 Or report("Updated") <> 1 Or report("Unchanged") <> 2 Then
        Err.Raise 5, , "Incorrect Movement diff rows or event counters."
    End If
    If rows(1)("Before") <> "Manual value" Or rows(1)("After") <> "Generated value" Then
        Err.Raise 5, , "Movement diff did not retain the actual pre-write cell value."
    End If
    If rows(2)("After") <> ex_PEB_Export.fn_Config("movementDiff.empty") Then
        Err.Raise 5, , "Cleared Movement values must be explicit."
    End If
    If rows(3)("Action") <> ex_PEB_Export.fn_Config("movementDiff.added") Then
        Err.Raise 5, , "Appended Movement rows must be identified."
    End If
    If Not ex_PEB_MovementDiff.fn_Result(report, result, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    If result.RowCount <> 3 Then
        Err.Raise 5, , "Movement diff table lost report rows."
    End If
    If rows(1)("UndoBefore") <> "Manual value" Or rows(1)("UndoAfter") <> "Generated value" Then
        Err.Raise 5, , "Undo values must retain raw capture independently of display text."
    End If
    Set sheet = book.Worksheets.Add
    sheet.Range("A1").Value2 = ex_PEB_Export.fn_Config("movementDiff.columns.DiffId")
    sheet.Range("A2").Value2 = "2"
    sheet.Range("A3").Value2 = "1"
    sheet.Range("A4").Value2 = "3"
    Set display = sheet.ListObjects.Add(1, sheet.Range("A1:A4"), , 1)
    ' A physically reordered display resolves by stable identity rather than source ordinal.
    If Not ex_PEB_MovementDiff.fn_Toggle(report, sheet.Range("A2"), display.Range) Then
        Err.Raise 5, , "Cannot select a displayed diff row."
    End If
    If Not rows(2)("UndoSelected") Or rows(1)("UndoSelected") Then
        Err.Raise 5, , "Undo selection addressed the wrong source row after sorting."
    End If
    If ex_PEB_MovementDiff.fn_SelectedCount(report) <> 1 Then
        Err.Raise 5, , "Incorrect undo selection count."
    End If
    result.Dispose
    If Not ex_PEB_MovementDiff.fn_Result(report, result, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    Set tags = result
    If tags.RowTagsAt(2) <> "diffUndoPending" Then
        Err.Raise 5, , "Selected undo rows must expose the pending style tag."
    End If
    If Not ex_PEB_MovementDiff.fn_Toggle(report, sheet.Range("A2"), display.Range) Then
        Err.Raise 5, , "Cannot clear an undo selection."
    End If
    If ex_PEB_MovementDiff.fn_SelectedCount(report) <> 0 Then
        Err.Raise 5, , "Repeated undo selection must clear the mark."
    End If
    rows(2)("UndoApplied") = True
    If ex_PEB_MovementDiff.fn_Toggle(report, sheet.Range("A2"), display.Range) Then
        Err.Raise 5, , "An applied undo must not be selected again."
    End If
    If ex_PEB_MovementDiff.fn_Toggle(report, sheet.Range("A1"), display.Range) Then
        Err.Raise 5, , "Headers must not be selectable for undo."
    End If
    fn_Run = True
Cleanup:
    On Error Resume Next
    Set tags = Nothing
    If Not result Is Nothing Then
        result.Dispose
    End If
    If Not book Is Nothing Then
        book.Close False
    End If
    If Not previousBook Is Nothing Then
        previousBook.Activate
    End If
    On Error GoTo 0
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_Action( _
    ByVal identifier As String, _
    ByVal isUpdate As Boolean, _
    ByVal status As String _
) As Object
    Dim action As Object

    Set action = VBA.CreateObject("Scripting.Dictionary")
    action.Add "RowId", identifier
    action.Add "Revision", 1
    action.Add "PersonName", identifier
    action.Add "EventName", "Fixture"
    action.Add "Operation", "append"
    action.Add "IsUpdate", isUpdate
    action.Add "Status", status
    Set private_Action = action
End Function

Private Function private_Entry( _
    ByVal action As Object, _
    ByVal header As String, _
    ByVal before As Variant, _
    ByVal after As Variant, _
    ByVal rowIndex As Long _
) As Object
    Dim entry As Object

    Set entry = VBA.CreateObject("Scripting.Dictionary")
    entry.Add "Action", action
    entry.Add "Header", header
    entry.Add "Value", before
    entry.Add "Expected", after
    entry.Add "Row", rowIndex + 1
    entry.Add "TableRow", rowIndex
    entry.Add "Format", "@"
    Set private_Entry = entry
End Function
' --------------------------------------
' } // namespace Private
' --------------------------------------