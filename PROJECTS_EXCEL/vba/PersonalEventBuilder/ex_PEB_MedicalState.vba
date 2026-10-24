Attribute VB_Name = "ex_PEB_MedicalState"
Option Explicit

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_TryApply( _
    ByVal rows As Object, _
    ByVal orderNumber As String, _
    ByVal orderDate As Date, _
    ByVal queryService As obj_TableQueryService, _
    ByRef warnings As Long, _
    ByRef diagnostic As String _
) As Boolean
    Dim events As Collection
    Dim storedEvent As Object
    Dim plan As Object
    Dim completed As Object
    Dim people As Object
    Dim latest As Object
    Dim fields As Object
    Dim state As Object
    Dim key As Variant
    Dim personId As String
    Dim note As String
    Dim status As String
    Dim text As String
    Dim pending As Object
    Dim startedAt As Double
    Dim stageAt As Double

    On Error GoTo EH
    startedAt = VBA.Timer
    stageAt = VBA.Timer
    Set completed = VBA.CreateObject("Scripting.Dictionary")
    Set people = VBA.CreateObject("Scripting.Dictionary")
    For Each key In rows.Keys
        Set fields = rows(key)
        personId = private_Id(fields("PersonId"))
        If VBA.Len(personId) > 0 Then people(personId) = True
    Next key
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.State.People | People=" & VBA.CStr(people.Count), stageAt
    stageAt = VBA.Timer
    ' A people-scoped read permits a legitimately empty history result.
    If Not ex_PEB_History.fn_ReadOrder(orderNumber, orderDate, events, diagnostic, False, people, False) Then Err.Raise 5, , diagnostic
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.History.Read | Events=" & VBA.CStr(events.Count), stageAt
    stageAt = VBA.Timer
    For Each storedEvent In events
        If storedEvent("RowType") = "Event" And storedEvent("Status") = "Active" Then
            If storedEvent("AcceptedOrderNumber") = orderNumber And _
                    ex_PEB_History.fn_Date(storedEvent("AcceptedOrderDate")) = orderDate And _
                    ex_PEB_History.fn_Date(storedEvent("EventDate")) = orderDate Then
                If ex_PEB_MovementRegistry.fn_TryEvent(storedEvent("EventId"), plan) Then
                    If plan("targetStatus") = private_Config("history.targetStatus") Then
                        personId = private_Id(storedEvent("PersonId"))
                        If VBA.Len(personId) > 0 Then completed(personId) = True
                    End If
                End If
            End If
        End If
    Next storedEvent
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.History.Match | People=" & VBA.CStr(completed.Count), stageAt
    stageAt = VBA.Timer
    Set pending = VBA.CreateObject("Scripting.Dictionary")
    For Each key In people.Keys
        If Not completed.Exists(key) Then pending(key) = True
    Next key
    If Not private_ReadMovement(orderDate, pending, queryService, latest, diagnostic) Then Err.Raise 5, , diagnostic
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.Movement | People=" & VBA.CStr(pending.Count), stageAt
    stageAt = VBA.Timer
    For Each key In rows.Keys
        Set fields = rows(key)
        personId = private_Id(fields("PersonId"))
        note = VBA.vbNullString
        If completed.Exists(personId) Then
            note = private_Config("text.State.Recorded")
            note = VBA.Replace$(note, "{number}", orderNumber)
            note = VBA.Replace$(note, "{date}", VBA.Format$(orderDate, "dd.mm.yyyy"))
            If VBA.Len(fields("RowTags")) = 0 Then fields("RowTags") = "medicalRecorded"
        ElseIf latest.Exists(personId) Then
            Set state = latest(personId)
            If VBA.Len(ex_PEB_Export.fn_Text(state("ArrivalDate"))) = 0 Then
                status = ex_PEB_Export.fn_Text(state("EventName"))
                If VBA.StrComp(status, ex_PEB_Export.fn_Config("export.Movement.status.HospitalTreatment"), VBA.vbTextCompare) = 0 Then
                    note = private_Config("text.State.Hospital")
                ElseIf VBA.StrComp(status, ex_PEB_Export.fn_Config("export.Movement.status.MedicalLeave"), VBA.vbTextCompare) = 0 Then
                    note = private_Config("text.State.MedicalLeave")
                    text = private_DateText(state("PlannedReturnDate"))
                    If VBA.Len(text) > 0 Then
                        note = note & " " & VBA.Replace$(private_Config("text.State.Until"), "{date}", text)
                    End If
                End If
                If VBA.Len(note) > 0 Then
                    note = VBA.Replace$(note, "{from}", private_DateText(state("DepartureDate")))
                    If fields("RowTags") <> "medicalReview" Then warnings = warnings + 1
                    fields("RowTags") = "medicalConflict"
                End If
            End If
        End If
        fields.Add "Notes", note
    Next key
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.State.Notes | Rows=" & VBA.CStr(rows.Count), stageAt
    fn_TryApply = True
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.State.Total | Success=True", startedAt
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.State.Total | Success=False", startedAt
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_ReadMovement( _
    ByVal orderDate As Date, _
    ByVal people As Object, _
    ByVal queryService As obj_TableQueryService, _
    ByRef latest As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim source As obj_TableSource
    Dim query As obj_TableQuery
    Dim data As obj_DataTable
    Dim headers As Object
    Dim indexes As Object
    Dim state As Object
    Dim previous As Object
    Dim field As Variant
    Dim personId As String
    Dim row As Long
    Dim dateKey As Double
    Dim arrivalKey As Double
    Dim replaceState As Boolean
    Dim values As Variant
    Dim stageAt As Double

    On Error GoTo EH
    stageAt = VBA.Timer
    Set latest = VBA.CreateObject("Scripting.Dictionary")
    If people.Count = 0 Then
        ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.Movement.Skip | People=0", stageAt
        private_ReadMovement = True
        Exit Function
    End If
    Set headers = VBA.CreateObject("Scripting.Dictionary")
    For Each field In VBA.Array("PersonId", "EventName", "DepartureDate", "ArrivalDate")
        headers.Add field, ex_PEB_Export.fn_Config("export.Movement.headers." & field)
    Next field
    headers.Add "PlannedReturnDate", private_Config("movement.plannedHeader")
    Set source = New obj_TableSource
    Set query = New obj_TableQuery
    If Not source.Initialize() Or Not query.Initialize() Then
        Err.Raise 5, , VBA.Replace$(private_Config("text.InitFailed"), "{value}", "obj_TableSource / obj_TableQuery")
    End If
    source.WorkbookPath = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("validation.Movement.source.path"))
    source.SheetName = ex_PEB_Export.fn_Config("validation.Movement.source.sheet")
    source.RangeAddress = ex_PEB_Export.fn_Config("validation.Movement.source.range")
    For Each field In headers.Keys
        query.AddColumn headers(field)
    Next field
    If Not queryService.TryExecute(source, query, data, diagnostic) Then GoTo Cleanup
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.Movement.Query | Rows=" & VBA.CStr(data.RowCount), stageAt
    stageAt = VBA.Timer
    Set indexes = VBA.CreateObject("Scripting.Dictionary")
    For Each field In headers.Keys
        indexes.Add field, ex_TableQuery.fn_HeaderIndex(data.Headers, headers(field))
    Next field
    values = data.Values
    For row = 1 To data.RowCount
        personId = private_Id(values(row, indexes("PersonId")))
        If people.Exists(personId) Then
            dateKey = ex_PEB_MovementSource.fn_OrderDate(values(row, indexes("DepartureDate")))
            arrivalKey = ex_PEB_MovementSource.fn_OrderDate(values(row, indexes("ArrivalDate")))
            If dateKey > 0 And dateKey <= VBA.CDbl(orderDate) Then
                replaceState = True
                If latest.Exists(personId) Then
                    Set previous = latest(personId)
                    replaceState = (dateKey > previous("DateKey") Or _
                        (dateKey = previous("DateKey") And arrivalKey <= previous("ArrivalKey")))
                End If
                If replaceState Then
                    Set state = VBA.CreateObject("Scripting.Dictionary")
                    For Each field In headers.Keys
                        state.Add field, values(row, indexes(field))
                    Next field
                    state.Add "DateKey", dateKey
                    state.Add "ArrivalKey", arrivalKey
                    If latest.Exists(personId) Then latest.Remove personId
                    latest.Add personId, state
                End If
            End If
        End If
    Next row
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.Movement.Index | People=" & VBA.CStr(latest.Count), stageAt
    private_ReadMovement = True
Cleanup:
    If Not data Is Nothing Then data.Dispose
    If Not query Is Nothing Then query.Dispose
    If Not source Is Nothing Then source.Dispose
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function

Private Function private_Id(ByVal value As Variant) As String
    Dim text As String

    text = ex_PEB_Export.fn_Text(value)
    text = VBA.Replace$(text, VBA.ChrW(160), "")
    text = VBA.Replace$(text, " ", "")
    text = VBA.Replace$(text, VBA.vbTab, "")
    text = VBA.Replace$(text, VBA.vbCr, "")
    text = VBA.Replace$(text, VBA.vbLf, "")
    If VBA.Left$(text, 1) = "'" Then text = VBA.Mid$(text, 2)
    If text Like "##########" Then private_Id = text
End Function

Private Function private_DateText(ByVal value As Variant) As String
    Dim dateKey As Double

    private_DateText = ex_PEB_Export.fn_Text(value)
    If VBA.Len(private_DateText) = 0 Then Exit Function
    If VBA.VarType(value) = VBA.vbDate Or VBA.IsNumeric(value) Then
        dateKey = VBA.CDbl(value)
        If dateKey >= 1 And dateKey <= 2958465 Then
            private_DateText = VBA.Format$(VBA.CDate(dateKey), "dd.mm.yyyy")
        End If
    End If
End Function

Private Function private_Config(ByVal key As String) As String
    private_Config = ex_PEB_Export.fn_Config("validation.MedicalReports." & key)
End Function
' --------------------------------------
' } // namespace Private
' --------------------------------------