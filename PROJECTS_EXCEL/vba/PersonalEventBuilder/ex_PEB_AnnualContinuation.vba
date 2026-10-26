Attribute VB_Name = "ex_PEB_AnnualContinuation"
Option Explicit

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Enabled(ByVal modeId As String) As Boolean
    fn_Enabled = InStr(1, "|" & ex_PEB_Export.fn_Config("lookup.AnnualContinuation.modes") & "|", "|" & modeId & "|", vbBinaryCompare) > 0
End Function

Public Function fn_Fields( _
    ByVal personId As String, _
    ByVal personName As String, _
    ByVal orderDate As Variant _
) As Object
    Dim snapshot As obj_TableSnapshotSet
    Dim source As obj_TableSource
    Dim query As obj_TableQuery
    Dim table As obj_DataTable
    Dim fields As Object
    Dim headers As Object
    Dim columns As Object
    Dim field As Variant
    Dim cacheKey As String
    Dim diagnostic As String
    Dim failure As String
    Dim row As Long
    Dim hospitalRow As Long
    Dim vacationRow As Long
    Dim departure As Double
    Dim latest As Double
    Dim cutoff As Double
    Dim planned As Double
    Dim annualStart As Double
    Dim hospitalStatus As String
    Dim annualStatus As String
    Dim rx As Object
    Dim matches As Object
    Dim basis As String

    On Error GoTo EH
    Set snapshot = ex_TableSnapshots.fn_Current()
    If snapshot Is Nothing Then Err.Raise 5, , "Annual continuation requires a loaded Movement snapshot."
    cutoff = ex_PEB_MovementSource.fn_OrderDate(orderDate)
    If cutoff <= 0 Then Err.Raise 5, , "Annual continuation requires an accepted order date."
    cacheKey = "AnnualContinuation|" & personId & "|" & ex_TableQuery.fn_NormalizeText(personName) & "|" & CStr(cutoff)
    If snapshot.Prepared.Exists(cacheKey) Then
        Set fn_Fields = snapshot.Prepared(cacheKey)
        Exit Function
    End If
    Set fields = CreateObject("Scripting.Dictionary")
    fields("Absence.DurationDays") = vbNullString
    fields("Absence.Destination") = vbNullString
    fields("Absence.DepartureDate") = Format$(CDate(cutoff), "dd.mm.yyyy")
    fields("Continuation.DocumentNumber") = vbNullString
    fields("Continuation.DocumentDate") = vbNullString
    fields("Absence.TicketNumber") = vbNullString
    fields("Continuation.TicketOrder") = vbNullString
    Set headers = CreateObject("Scripting.Dictionary")
    For Each field In Array("EventName", "DepartureDate", "ArrivalDate", "Destination", "DepartureBasis")
        headers(field) = ex_PEB_Export.fn_Config("export.Movement.headers." & field)
    Next field
    headers("PlannedArrivalDate") = ex_PEB_Export.fn_Config("lookup.ReturnTickets.arrivalHeader")
    headers("TicketNumber") = ex_PEB_Export.fn_Config("lookup.ReturnTickets.ticketHeader")
    headers("TicketOrder") = ex_PEB_Export.fn_Config("lookup.ReturnTickets.orderHeader")
    Set source = New obj_TableSource
    Set query = New obj_TableQuery
    If Not source.Initialize() Or Not query.Initialize() Then Err.Raise 5, , "Cannot initialize annual continuation query."
    source.WorkbookPath = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("validation.Movement.source.path"))
    source.SheetName = ex_PEB_Export.fn_Config("validation.Movement.source.sheet")
    source.RangeAddress = ex_PEB_Export.fn_Config("validation.Movement.source.range")
    For Each field In headers.Keys
        query.AddColumn headers(field)
    Next field
    If Len(Trim$(personId)) = 0 Then Err.Raise 5, , "Annual continuation candidate is missing Personnel.TaxId."
    query.AddCondition ex_PEB_Export.fn_Config("export.Movement.headers.PersonId"), QueryEquals, personId, QueryText, True
    If ex_TableQuery.fn_NormalizeText(personId) = ex_TableQuery.fn_NormalizeText(ex_PEB_Export.fn_Config("lookup.ReturnTickets.temporaryIpn")) Then
        If Len(Trim$(personName)) = 0 Then Err.Raise 5, , "Annual continuation candidate is missing Personnel.FullName."
        query.AddCondition ex_PEB_Export.fn_Config("lookup.ReturnTickets.nameHeader"), QueryEquals, personName, QueryText, True
    End If
    If Not snapshot.TryQuery(source, query, table, diagnostic) Then Err.Raise 5, , diagnostic
    Set columns = CreateObject("Scripting.Dictionary")
    For Each field In headers.Keys
        columns(field) = ex_TableQuery.fn_HeaderIndex(table.Headers, headers(field))
    Next field
    hospitalStatus = ex_PEB_Export.fn_Config("export.Movement.status.HospitalTreatment")
    annualStatus = ex_PEB_Export.fn_Config("export.Movement.status.AnnualVacation")

    ' Only the latest period as of the order may supply a hospital interruption.
    For row = 1 To table.RowCount
        departure = ex_PEB_MovementSource.fn_OrderDate(table.ValueAt(row, columns("DepartureDate")))
        If departure > latest And departure <= cutoff Then
            latest = departure
            hospitalRow = 0
            If CStr(table.ValueAt(row, columns("EventName"))) = hospitalStatus Then hospitalRow = row
        ElseIf departure = latest And departure > 0 And departure <= cutoff Then
            Err.Raise 5, , "Ambiguous latest Movement period for annual continuation: " & personId
        End If
    Next row
    If hospitalRow > 0 Then
        basis = CStr(table.ValueAt(hospitalRow, columns("DepartureBasis")))
        basis = Replace$(Replace$(basis, ChrW$(160), " "), ChrW$(8239), " ")
        Set rx = CreateObject("VBScript.RegExp")
        rx.Pattern = ex_PEB_Export.fn_Config("lookup.AnnualContinuation.cardPattern")
        rx.IgnoreCase = True
        rx.Global = True
        Set matches = rx.Execute(basis)
        If matches.Count > 1 Then Err.Raise 5, , "Ambiguous medical card in hospital departure basis: " & personId
        If matches.Count = 1 Then
            fields("Continuation.DocumentNumber") = matches(0).SubMatches(0)
            If Len(matches(0).SubMatches(1)) > 0 Then
                fields("Continuation.DocumentDate") = Format$(ex_PEB_History.fn_Date(matches(0).SubMatches(1)), "dd.mm.yyyy")
            End If
        End If
        For row = 1 To table.RowCount
            If CStr(table.ValueAt(row, columns("EventName"))) = annualStatus Then
                departure = ex_PEB_MovementSource.fn_OrderDate(table.ValueAt(row, columns("DepartureDate")))
                If departure > 0 And departure < latest Then
                    If ex_PEB_MovementSource.fn_OrderDate(table.ValueAt(row, columns("ArrivalDate"))) = latest Then
                        If vacationRow > 0 Then Err.Raise 5, , "Ambiguous interrupted annual vacation: " & personId
                        vacationRow = row
                        annualStart = departure
                    End If
                End If
            End If
        Next row
    End If
    If vacationRow > 0 Then
        fields("Absence.TicketNumber") = table.ValueAt(vacationRow, columns("TicketNumber"))
        fields("Continuation.TicketOrder") = table.ValueAt(vacationRow, columns("TicketOrder"))
        planned = ex_PEB_MovementSource.fn_OrderDate(table.ValueAt(vacationRow, columns("PlannedArrivalDate")))
        ' Planned arrival already includes travel; the interruption day remains unused.
        If planned > latest And latest > annualStart Then
            fields("Absence.DurationDays") = CStr(DateDiff("d", CDate(latest), CDate(planned)))
            fields("Absence.Destination") = table.ValueAt(vacationRow, columns("Destination"))
        End If
    End If
    snapshot.Prepared.Add cacheKey, fields
    Set fn_Fields = fields
Cleanup:
    If Not table Is Nothing Then table.Dispose
    If Not query Is Nothing Then query.Dispose
    If Not source Is Nothing Then source.Dispose
    If Len(failure) > 0 Then
        On Error GoTo 0
        Err.Raise 5, , failure
    End If
    Exit Function
EH:
    failure = Err.Description
    Resume Cleanup
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------