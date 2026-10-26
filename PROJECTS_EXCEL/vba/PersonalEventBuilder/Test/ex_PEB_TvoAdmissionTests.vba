Attribute VB_Name = "ex_PEB_TvoAdmissionTests"
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
    Dim sheet As Worksheet
    Dim target As ListObject
    Dim table As ListObject
    Dim fields As Variant
    Dim data() As Variant
    Dim columns As Object
    Dim values As Object
    Dim actor As Object
    Dim action As Object
    Dim plan As Object
    Dim duties As New Collection
    Dim emptyDuties As New Collection
    Dim actions As New Collection
    Dim writes As Object
    Dim operation As Variant
    Dim index As Long
    Dim startDate As Double
    Dim started As Boolean
    Dim errorNumber As Long
    Dim errorText As String

    On Error GoTo EH
    startDate = 1
    If ex_PEB_MovementTvo.fn_IsStarted(Empty, Empty, startDate, "Fixture") Or startDate <> 0 Then
        Err.Raise 5, , "Admission without acceptance must not start duties."
    End If
    If ex_PEB_MovementTvo.fn_IsStarted(vbNullString, DateSerial(2026, 10, 21), startDate, "Fixture") Then
        Err.Raise 5, , "A date without an acceptance order must not start duties."
    End If
    If Not ex_PEB_MovementTvo.fn_IsStarted("304", DateSerial(2026, 10, 21), startDate, "Fixture") Then
        Err.Raise 5, , "An acceptance order and date must start duties."
    End If
    If startDate <> CDbl(DateSerial(2026, 10, 21)) Then
        Err.Raise 5, , "Incorrect acceptance date."
    End If
    On Error Resume Next
    Err.Clear
    started = ex_PEB_MovementTvo.fn_IsStarted("304", Empty, startDate, "Fixture")
    errorNumber = Err.Number
    errorText = Err.Description
    Err.Clear
    On Error GoTo EH
    If errorNumber = 0 Or InStr(errorText, "TvoMissingStartDate") = 0 Then
        Err.Raise 5, , "An accepted assignment without a date must still be rejected."
    End If

    Set previousBook = Application.ActiveWorkbook
    Set book = Application.Workbooks.Add
    Set sheet = book.Worksheets(1)
    sheet.Range("A1").Value2 = "Fixture"
    Set target = sheet.ListObjects.Add(1, sheet.Range("A1:A2"), , 1)
    Set sheet = book.Worksheets.Add
    sheet.Name = ex_PEB_Export.fn_Config("reporter.tvo.source.sheet")
    fields = Split("PersonRank|PersonName|PersonId|OwnPositionCode|PositionCode|StartDate|EndDate|StartOrder|EndOrder|StartBasis|EndBasis|PlannedEndDate|AccessOrder", "|")
    ReDim data(1 To 2, 1 To UBound(fields) + 1)
    Set columns = CreateObject("Scripting.Dictionary")
    For index = 0 To UBound(fields)
        data(1, index + 1) = ex_PEB_Export.fn_Config("returnTvo.table.headers." & fields(index))
        columns(fields(index)) = index + 1
    Next index
    data(2, columns("PersonId")) = "AdmissionId"
    data(2, columns("PersonName")) = "Admitted actor"
    data(2, columns("OwnPositionCode")) = "ActorStaffPosition"
    data(2, columns("PositionCode")) = "OwnerPosition"
    data(2, columns("AccessOrder")) = "304"
    sheet.Range("A1").Resize(2, UBound(fields) + 1).Value2 = data
    Set table = sheet.ListObjects.Add(1, sheet.Range("A1").Resize(2, UBound(fields) + 1), , 1)
    table.Name = ex_PEB_Export.fn_Config("returnTvo.table.name")
    Set actor = CreateObject("Scripting.Dictionary")
    actor("PersonId") = "AdmissionId"
    actor("PersonName") = "Admitted actor"
    actor("PositionCode") = "OwnerPosition"
    duties.Add actor
    Set values = CreateObject("Scripting.Dictionary")
    values("PersonId") = "OwnerId"
    values("PersonName") = "Owner"
    values("PositionCode") = "OwnerPosition"
    values("EventDate") = DateSerial(2026, 10, 21)
    values("AcceptedOrderNumber") = "304"
    Set values("MetaTvoTables") = duties
    Set action = CreateObject("Scripting.Dictionary")
    action("RowId") = "FixtureAdmission"
    action("Status") = "Planned"
    action("EventStatus") = "Active"
    action("IsUpdate") = False
    Set action("TvoValues") = values
    Set writes = CreateObject("Scripting.Dictionary")
    Set action("Writes") = writes
    Set action("ClosingWrites") = writes
    actions.Add action
    Set plan = CreateObject("Scripting.Dictionary")
    Set plan("Actions") = actions
    For Each operation In Array("append", "close", "transition")
        action("Operation") = operation
        ex_PEB_MovementTvo.fn_Prepare plan, target
        If ex_PEB_MovementTvo.fn_ChangeCount(plan) <> 0 Or plan("TvoPlan")("FinalCount") <> 1 Then
            Err.Raise 5, , "Admission was closed, promoted or duplicated: " & operation
        End If
    Next operation
    Set values("MetaTvoTables") = emptyDuties
    action("Operation") = "transition"
    ex_PEB_MovementTvo.fn_Prepare plan, target
    If ex_PEB_MovementTvo.fn_ChangeCount(plan) <> 0 Then
        Err.Raise 5, , "Admission must not restore duties from legacy columns."
    End If
    Set values("MetaTvoTables") = duties

    ' Another accepted actor must remain usable even when the same post has an admission row.
    table.ListRows.Add
    table.DataBodyRange.Cells(2, columns("PersonId")).Value2 = "ActiveId"
    table.DataBodyRange.Cells(2, columns("PersonName")).Value2 = "Accepted actor"
    table.DataBodyRange.Cells(2, columns("OwnPositionCode")).Value2 = "ActiveStaffPosition"
    table.DataBodyRange.Cells(2, columns("PositionCode")).Value2 = "OwnerPosition"
    table.DataBodyRange.Cells(2, columns("StartOrder")).Value2 = "304"
    table.DataBodyRange.Cells(2, columns("StartDate")).Value2 = CDbl(DateSerial(2026, 10, 21))
    action("Operation") = "append"
    ex_PEB_MovementTvo.fn_Prepare plan, target
    If ex_PEB_MovementTvo.fn_ChangeCount(plan) <> 0 Then
        Err.Raise 5, , "Admission metadata must not replace an accepted actor."
    End If
    action("Operation") = "close"
    ex_PEB_MovementTvo.fn_Prepare plan, target
    If ex_PEB_MovementTvo.fn_ChangeCount(plan) <> 2 Then
        Err.Raise 5, , "Returning must close only the accepted actor's date and order."
    End If
    fn_Run = True
Cleanup:
    On Error Resume Next
    If Not book Is Nothing Then
        book.Close False
    End If
    If Not previousBook Is Nothing Then
        previousBook.Activate
    End If
    Exit Function
EH:
    diagnostic = Err.Description
    Resume Cleanup
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------