Attribute VB_Name = "ex_PEB_MovementPlanner"
Option Explicit

#Const TVO_DEPRECATED = True

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
Public Function fn_BuildOrder( _
    ByVal orderNumber As String, _
    ByVal orderDate As Variant, _
    ByVal target As ListObject, _
    ByRef plan As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim events As Collection
    Dim snapshotStarted As Boolean
    Dim startedAt As Double

    On Error GoTo EH
    Set plan = Nothing
    diagnostic = VBA.vbNullString
    ex_PEB_Export.fn_BeginConfigSnapshot
    snapshotStarted = True
    startedAt = VBA.Timer
    If ex_PEB_History.fn_ReadOrder(orderNumber, orderDate, events, diagnostic, True, Nothing, False) Then
        ex_Core.fn_Diagnostic_WritePerf "Movement.Preview.ReadEvents", startedAt
        fn_BuildOrder = fn_Build(events, target, plan, diagnostic)
    End If
Cleanup:
    If snapshotStarted Then
        ex_PEB_Export.fn_EndConfigSnapshot
    End If
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function

Public Function fn_Build( _
    ByVal events As Collection, _
    ByVal target As ListObject, _
    ByRef plan As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim pebExportService As obj_PEB_ExportService
    Dim rows As Collection
    Dim ordered As Collection
    Dim actions As Collection
    Dim blocked As Object
    Dim ids As Object
    Dim headers As Object
    Dim stored As Object
    Dim values As Object
    Dim action As Object
    Dim row As Object
    Dim personId As String
    Dim conflicts As Long
    Dim message As String
    Dim orderIdentity As String
    Dim identity As String
    Dim targetStamp As String
    Dim people As Object
    Dim eventPeople As Object
    Dim startedAt As Double
    Dim nextTargetIndex As Long

    On Error GoTo EH
    Set plan = Nothing
    diagnostic = VBA.vbNullString
    If target Is Nothing Then
        private_Fail "MissingTarget", ""
    End If
    If events Is Nothing Then
        private_Fail "MissingEvents", ""
    End If
    If events.Count = 0 Then
        private_Fail "MissingEvents", ""
    End If
    If target.Parent.Parent.Date1904 Then
        private_Fail "UnsupportedDateSystem", ""
    End If
    startedAt = VBA.Timer
    ex_Core.fn_Diagnostic_WriteLog "MOVEMENT_PREVIEW_STARTED | Table=" & target.Name & " | Rows=" & target.ListRows.Count & " | Columns=" & target.ListColumns.Count
    If Not ex_Core.fn_Diagnostic_Flush(True) Then
        Debug.Print "Movement diagnostic flush failed."
    End If
    Set people = VBA.CreateObject("Scripting.Dictionary")
    people.CompareMode = VBA.vbTextCompare
    Set eventPeople = VBA.CreateObject("Scripting.Dictionary")
    For Each stored In events
        personId = ex_PEB_Export.fn_Text(stored("PersonId"))
        people(personId) = True
        eventPeople(ex_PEB_Export.fn_Text(stored("RowId"))) = personId
    Next stored
    targetStamp = ex_PEB_MovementLinks.fn_StructureStamp(target)
    ex_Core.fn_Diagnostic_WritePerf "Movement.Preview.Snapshot", startedAt
    startedAt = VBA.Timer
    Set rows = private_Read(target, headers, people, eventPeople)
    nextTargetIndex = target.ListRows.Count
    ex_Core.fn_Diagnostic_WriteLog "MOVEMENT_MODEL | People=" & people.Count & " | SelectedRows=" & rows.Count & " | TotalRows=" & target.ListRows.Count
    ex_Core.fn_Diagnostic_WritePerf "Movement.Preview.Model", startedAt
    startedAt = VBA.Timer
    Set ordered = New Collection
    Set actions = New Collection
    Set blocked = VBA.CreateObject("Scripting.Dictionary")
    blocked.CompareMode = VBA.vbTextCompare
    Set ids = VBA.CreateObject("Scripting.Dictionary")
    ' Validate and sort the complete order before calculating changes.
    For Each stored In events
        identity = ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(stored, "RowId"))
        If VBA.Len(identity) = 0 Then
            private_Fail "MissingRowId", ""
        End If
        If ids.Exists(identity) Then
            private_Fail "DuplicateRowId", identity
        End If
        ids.Add identity, True
        If stored("RowType") <> "Event" Then
            private_Fail "NotAnEvent", identity
        End If
        identity = ex_PEB_Export.fn_Text(stored("AcceptedOrderNumber")) & "|" & _
            VBA.Format$(ex_PEB_History.fn_Date(stored("AcceptedOrderDate")), "yyyy-mm-dd")
        If VBA.Len(orderIdentity) = 0 Then
            orderIdentity = identity
        End If
        If orderIdentity <> identity Then
            private_Fail "MixedOrders", identity
        End If
        private_Insert ordered, stored
    Next stored
    ex_PEB_EventRoutes.fn_Module_Dispose
    Set pebExportService = New obj_PEB_ExportService
    If Not pebExportService.Initialize() Then
        private_Fail "InitializationFailed", ""
    End If
    For Each stored In ordered
        ex_Core.fn_Diagnostic_WriteLog "MOVEMENT_EVENT | RowId=" & stored("RowId")
        Set action = VBA.CreateObject("Scripting.Dictionary")
        action.Add "RowId", stored("RowId")
        action.Add "Revision", stored("Revision")
        action.Add "EventStatus", stored("Status")
        action.Add "PersonId", stored("PersonId")
        action.Add "PersonName", stored("PersonName")
        action.Add "EventName", stored("EventLabel")
        action.Add "EventDate", ex_PEB_History.fn_Date(stored("EventDate"))
        action.Add "RowOrder", VBA.CLng(stored("RowOrder"))
        action.Add "HistoryStamp", ex_PEB_MovementLinks.fn_EventStamp(stored)
        action.Add "IsUpdate", False
        action.Add "Status", "Conflict"
        action.Add "Diagnostic", VBA.vbNullString
        personId = ex_PEB_Export.fn_Text(stored("PersonId"))
        message = VBA.vbNullString
        If blocked.Exists(personId) Then
            message = private_Message("PreviousConflict", VBA.CStr(blocked(personId)))
        ElseIf Not pebExportService.TryPrepareHistoryMovement(stored, values, message) Then
            blocked.Add personId, stored("RowId")
        ElseIf Not private_Apply(rows, headers, values, action, nextTargetIndex, message) Then
            blocked.Add personId, stored("RowId")
        Else
            action.Add "TvoValues", values
        End If
        If VBA.Len(message) > 0 Then
            action("Diagnostic") = message
            conflicts = conflicts + 1
        End If
        actions.Add action
    Next stored
    Set ids = VBA.CreateObject("Scripting.Dictionary")
    ids.CompareMode = VBA.vbTextCompare
    ' Validate final periods after all events; a later arrival may close a new departure.
    For Each action In actions
        personId = action("PersonId")
        If Not ids.Exists(personId) Then
            ids.Add personId, True
            If Not blocked.Exists(personId) Then
                message = VBA.vbNullString
                If Not private_TryValidatePeriods(rows, personId, message) Then
                    For Each row In actions
                        If VBA.StrComp(row("PersonId"), personId, VBA.vbTextCompare) = 0 Then
                            row("Status") = "Conflict"
                            row("Diagnostic") = message
                            conflicts = conflicts + 1
                        End If
                    Next row
                End If
            End If
        End If
    Next action
    Set plan = VBA.CreateObject("Scripting.Dictionary")
    plan.Add "WorkbookPath", target.Parent.Parent.FullName
    plan.Add "WorksheetName", target.Parent.Name
    plan.Add "TableName", target.Name
    plan.Add "OrderIdentity", orderIdentity
    plan.Add "OrderNumber", ordered(1)("AcceptedOrderNumber")
    plan.Add "OrderDate", ordered(1)("AcceptedOrderDate")
    plan.Add "Actions", actions
    plan.Add "ConflictCount", conflicts
    plan.Add "IsReady", (conflicts = 0)
    plan.Add "TargetStamp", targetStamp
    ex_PEB_MovementTvo.fn_Prepare plan, target
    plan.Add "PlanStamp", ex_PEB_MovementLinks.fn_PlanStamp(plan)
    ex_Core.fn_Diagnostic_WritePerf "Movement.Preview.Events", startedAt
    startedAt = VBA.Timer
    identity = ex_PEB_MovementLinks.fn_StructureStamp(target)
    If targetStamp <> identity Then
        private_Fail "TargetChangedDuringPreview", target.Name
    End If
    ex_Core.fn_Diagnostic_WritePerf "Movement.Preview.CheckUnchanged", startedAt
    fn_Build = True
Cleanup:
    If Not pebExportService Is Nothing Then
        pebExportService.Dispose
    End If
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    ex_Core.fn_Diagnostic_WriteError "Movement.Preview", VBA.Err.Number, VBA.Err.Source, diagnostic
    Set plan = Nothing
    Resume Cleanup
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_Read( _
    ByVal target As ListObject, _
    ByRef headers As Object, _
    ByVal people As Object, _
    ByVal eventPeople As Object _
) As Collection
    Dim rows As Collection
    Dim row As Object
    Dim cells As Object
    Dim formulas As Object
    Dim column As ListColumn
    Dim index As Long
    Dim key As Variant
    Dim links As Object
    Dim present As Long
    Dim openingIds As Object
    Dim closingIds As Object
    Dim values As Variant
    Dim formulaValues As Variant
    Dim masks As Variant
    Dim names() As String
    Dim columnIndex As Long
    Dim personValues As Variant
    Dim personValue As Variant
    Dim personHeader As String
    Dim snapshot As Object
    Dim linkedId As String
    Dim linkCells As Object
    Dim openingValues As Variant
    Dim openingStamps As Variant
    Dim closingValues As Variant
    Dim closingStamps As Variant

    Set rows = New Collection
    ReDim names(1 To target.ListColumns.Count)
    Set headers = VBA.CreateObject("Scripting.Dictionary")
    headers.CompareMode = VBA.vbTextCompare
    For Each column In target.ListColumns
        headers.Add column.Name, False
        names(column.Index) = column.Name
    Next column
    Set links = ex_PEB_MovementLinks.fn_Headers()
    Set openingIds = VBA.CreateObject("Scripting.Dictionary")
    Set closingIds = VBA.CreateObject("Scripting.Dictionary")
    For Each key In links.Keys
        If headers.Exists(links(key)) Then
            present = present + 1
        End If
    Next key
    If present <> links.Count Then
        private_Fail "IncompleteLinkSchema", target.Name
    End If
    For Each key In VBA.Array("PersonId", "EventName", "DepartureDate", "ArrivalDate")
        private_RequireHeader headers, ex_PEB_Export.fn_Config("export.Movement.headers." & key)
    Next key
    If target.ListRows.Count = 0 Then
        Set private_Read = rows
        Exit Function
    End If
    personHeader = ex_PEB_Export.fn_Config("export.Movement.headers.PersonId")
    personValues = target.ListColumns(personHeader).DataBodyRange.Value2
    Set linkCells = VBA.CreateObject("Scripting.Dictionary")
    openingValues = target.ListColumns(links("OpeningId")).DataBodyRange.Value2
    openingStamps = target.ListColumns(links("OpeningHash")).DataBodyRange.Value2
    closingValues = target.ListColumns(links("ClosingId")).DataBodyRange.Value2
    closingStamps = target.ListColumns(links("ClosingHash")).DataBodyRange.Value2
    For index = 1 To target.ListRows.Count
        ' Check link uniqueness globally without creating a model for unrelated people.
        If target.ListRows.Count = 1 Then
            linkCells(links("OpeningId")) = openingValues
            linkCells(links("OpeningHash")) = openingStamps
            linkCells(links("ClosingId")) = closingValues
            linkCells(links("ClosingHash")) = closingStamps
        Else
            linkCells(links("OpeningId")) = openingValues(index, 1)
            linkCells(links("OpeningHash")) = openingStamps(index, 1)
            linkCells(links("ClosingId")) = closingValues(index, 1)
            linkCells(links("ClosingHash")) = closingStamps(index, 1)
        End If
        private_ValidateLink linkCells, links("OpeningId"), links("OpeningHash"), openingIds
        private_ValidateLink linkCells, links("ClosingId"), links("ClosingHash"), closingIds
        If target.ListRows.Count = 1 Then
            personValue = personValues
        Else
            personValue = personValues(index, 1)
        End If
        If VBA.IsError(personValue) Then
            private_Fail "InvalidCell", personHeader & "/" & index
        End If
        For Each key In VBA.Array("OpeningId", "ClosingId")
            linkedId = ex_PEB_Export.fn_Text(linkCells(links(key)))
            If eventPeople.Exists(linkedId) Then
                If VBA.StrComp(eventPeople(linkedId), ex_PEB_Export.fn_Text(personValue), VBA.vbTextCompare) <> 0 Then
                    private_Fail "AppliedPersonChanged", linkedId
                End If
            End If
        Next key
        If Not people.Exists(ex_PEB_Export.fn_Text(personValue)) Then
            GoTo NextHistoryRow
        End If
        ' Materialize only the selected people's history, retaining physical row numbers.
        Set snapshot = ex_PEB_MovementLinks.fn_ReadRange(target.DataBodyRange.Rows(index))
        values = snapshot("Values")
        formulaValues = snapshot("Formulas")
        masks = snapshot("FormulaMasks")
        Set row = VBA.CreateObject("Scripting.Dictionary")
        Set cells = VBA.CreateObject("Scripting.Dictionary")
        Set formulas = VBA.CreateObject("Scripting.Dictionary")
        cells.CompareMode = VBA.vbTextCompare
        formulas.CompareMode = VBA.vbTextCompare
        For columnIndex = 1 To UBound(names)
            If VBA.IsError(values(1, columnIndex)) Then
                private_Fail "InvalidCell", names(columnIndex) & "/" & index
            End If
            cells.Add names(columnIndex), values(1, columnIndex)
            If masks(1, columnIndex) Then
                formulas.Add names(columnIndex), formulaValues(1, columnIndex)
                headers(names(columnIndex)) = True
            End If
        Next columnIndex
        For Each key In links.Keys
            If formulas.Exists(links(key)) Then
                private_Fail "FormulaInLinkColumn", links(key)
            End If
        Next key
        row.Add "Cells", cells
        row.Add "Formulas", formulas
        row.Add "SourceRow", index
        row.Add "TargetRow", index
        row.Add "OriginRowId", VBA.vbNullString
        row.Add "ClosingRowId", VBA.vbNullString
        rows.Add row
NextHistoryRow:
    Next index
    Set private_Read = rows
End Function

Private Sub private_Insert( _
    ByVal ordered As Collection, _
    ByVal stored As Object _
)
    ex_PEB_MovementValidation.fn_Insert ordered, stored
End Sub

Private Function private_Apply( _
    ByRef rows As Collection, _
    ByVal headers As Object, _
    ByVal values As Object, _
    ByVal action As Object, _
    ByRef nextTargetIndex As Long, _
    ByRef diagnostic As String _
) As Boolean
    Dim prefix As String
    Dim operation As String
    Dim openHeader As String
    Dim writes As Object
    Dim closingWrites As Object
    Dim identities As Object
    Dim joins As Object
    Dim row As Object
    Dim cells As Object
    Dim candidate As Collection
    Dim selected As Long
    Dim index As Long
    Dim key As Variant
    Dim personHeader As String
    Dim departureHeader As String
    Dim eventDate As Date
    Dim links As Object
    Dim semanticPeriods As Collection
    Dim retainedClosings As Object
    Dim semanticIndex As Long

    On Error GoTo EH
    prefix = "profiles." & values("MovementPlanId")
    operation = ex_PEB_MovementRegistry.fn_Value(prefix & ".operation")
    Set writes = private_Mappings(prefix & ".fields", values)
    Set identities = private_Mappings(prefix & ".identity", values)
    Set closingWrites = writes
    If operation = "append" Then
        Set joins = private_Mappings(prefix & ".guard.join", values)
        openHeader = ex_PEB_MovementRegistry.fn_Value(prefix & ".guard.openHeader")
    ElseIf operation = "close" Or operation = "transition" Then
        Set joins = private_Mappings(prefix & ".join", values)
        openHeader = ex_PEB_MovementRegistry.fn_Value(prefix & ".openHeader")
        If operation = "transition" Then
            Set closingWrites = private_Mappings("profiles." & ex_PEB_MovementRegistry.fn_Value(prefix & ".closeProfile") & ".fields", values)
        End If
    Else
        private_Fail "UnknownOperation", operation
    End If
    personHeader = ex_PEB_Export.fn_Config("export.Movement.headers.PersonId")
    departureHeader = ex_PEB_Export.fn_Config("export.Movement.headers.DepartureDate")
    eventDate = ex_PEB_History.fn_Date(values("EventDate"))
    Set links = ex_PEB_MovementLinks.fn_Headers()
    If operation = "close" Then
        private_PreserveArrivalFoodSupport rows, values("RowId"), closingWrites, links
    End If
    If values("EventId") = "MoveFromAwolToBusinessTrip" Then
        ' Enroll on return without overwriting the original AWOL removal date.
        closingWrites(ex_PEB_Export.fn_Config("export.Movement.headers.ArrivalFoodSupport")) = eventDate
    End If
    action.Add "Operation", operation
    action.Add "Writes", writes
    action.Add "ClosingWrites", closingWrites
    action.Add "GeneratedWrites", private_CopyFields(writes)
    action.Add "GeneratedClosingWrites", private_CopyFields(closingWrites)
    action.Add "Joins", joins
    action.Add "OpeningHash", ex_PEB_MovementLinks.fn_FieldsHash(writes)
    action.Add "ClosingHash", ex_PEB_MovementLinks.fn_FieldsHash(closingWrites)
    If values("Status") = "Cancelled" Then
        private_Cancel rows, action, closingWrites, joins, links
        private_Apply = True
        Exit Function
    End If
    If values("EventId") = "MoveFromAwolToBusinessTrip" Then
        private_PlanReturnPlaceholder rows, headers, action, links, nextTargetIndex
    End If
    If private_Existing(rows, values, action, writes, closingWrites, joins, links, headers, nextTargetIndex) Then
        If action.Exists("PlaceholderWrites") And action("Status") = "AlreadyApplied" Then
            action("Status") = "Planned"
            action("IsUpdate") = True
        End If
        private_Apply = True
        Exit Function
    End If
    ' Common semantic rules run on a detached model; target/link checks remain below.
    Set semanticPeriods = private_StatePeriods(rows, values("PersonId"))
    Set retainedClosings = VBA.CreateObject("Scripting.Dictionary")
    semanticIndex = ex_PEB_MovementValidation.fn_Apply(semanticPeriods, values, retainedClosings)
    If semanticIndex > 0 Then
        selected = semanticPeriods(semanticIndex)("ModelIndex")
    End If
    For Each key In writes.Keys
        private_RequireHeader headers, VBA.CStr(key)
        If operation <> "close" And headers(key) Then
            private_Fail "FormulaColumnWrite", VBA.CStr(key)
        End If
    Next key
    For Each key In closingWrites.Keys
        private_RequireHeader headers, VBA.CStr(key)
    Next key
    For Each key In identities.Keys
        private_RequireHeader headers, VBA.CStr(key)
    Next key
    For Each key In joins.Keys
        private_RequireHeader headers, VBA.CStr(key)
    Next key
    private_RequireHeader headers, openHeader
    For index = 1 To rows.Count
        Set row = rows(index)
        Set cells = row("Cells")
        If private_Matches(cells, identities) Then
            private_Fail "ExistingIdentityRequiresLink", VBA.CStr(row("SourceRow"))
        End If
        If VBA.StrComp(ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(cells, personHeader)), values("PersonId"), VBA.vbTextCompare) = 0 Then
            For Each key In VBA.Array(departureHeader, openHeader)
                If VBA.Len(ex_PEB_Export.fn_Text(cells(key))) > 0 Then
                    If private_Date(cells(key)) = eventDate Then
                        If VBA.Len(row("OriginRowId")) = 0 And VBA.Len(row("ClosingRowId")) = 0 Then
                            ' The unique selected period may start and end on the same date.
                            If index <> selected Or operation = "append" Or VBA.CStr(key) <> departureHeader Then
                                private_Fail "AmbiguousHistoricalSameDay", VBA.CStr(row("SourceRow"))
                            End If
                        End If
                    End If
                End If
            Next key
        End If
    Next index
    Set candidate = New Collection
    For index = 1 To rows.Count
        If operation <> "append" And index = selected Then
            candidate.Add private_CopyRow(rows(index))
        Else
            candidate.Add rows(index)
        End If
    Next index
    action.Add "SourceRow", 0
    action.Add "ClosingIndex", selected
    If operation <> "append" Then
        Set row = candidate(selected)
        action("ClosingIndex") = row("TargetRow")
        Set cells = row("Cells")
        If eventDate < private_Date(ex_PEB_Export.fn_Required(cells, departureHeader)) Then
            private_Fail "ArrivalBeforeDeparture", VBA.CStr(row("SourceRow"))
        End If
        For Each key In closingWrites.Keys
            If row("Formulas").Exists(key) Then
                private_Fail "FormulaOverwrite", VBA.CStr(key)
            End If
            If VBA.Len(ex_PEB_Export.fn_Text(ex_PEB_Export.fn_Required(cells, VBA.CStr(key)))) > 0 Then
                private_Fail "PopulatedCellOverwrite", VBA.CStr(key)
            End If
            cells(key) = closingWrites(key)
        Next key
        row("ClosingRowId") = values("RowId")
        If VBA.Len(ex_PEB_Export.fn_Text(cells(links("ClosingId")))) > 0 Then
            private_Fail "ClosingLinkAlreadyPopulated", VBA.CStr(row("SourceRow"))
        End If
        cells(links("ClosingId")) = values("RowId")
        cells(links("ClosingHash")) = action("ClosingHash")
        action("SourceRow") = row("SourceRow")
    End If
    If operation <> "close" Then
        Set row = VBA.CreateObject("Scripting.Dictionary")
        Set cells = VBA.CreateObject("Scripting.Dictionary")
        cells.CompareMode = VBA.vbTextCompare
        For Each key In headers.Keys
            cells.Add VBA.CStr(key), VBA.vbNullString
        Next key
        For Each key In writes.Keys
            cells(key) = writes(key)
        Next key
        cells(links("OpeningId")) = values("RowId")
        cells(links("OpeningHash")) = action("OpeningHash")
        If Not cells.Exists(openHeader) Then
            cells.Add openHeader, VBA.vbNullString
        End If
        row.Add "Cells", cells
        row.Add "Formulas", VBA.CreateObject("Scripting.Dictionary")
        row.Add "SourceRow", 0
        row.Add "TargetRow", nextTargetIndex + 1
        row.Add "OriginRowId", values("RowId")
        row.Add "ClosingRowId", VBA.vbNullString
        candidate.Add row
    End If
    action.Add "TargetIndex", selected
    If operation <> "close" Then
        nextTargetIndex = nextTargetIndex + 1
        action("TargetIndex") = nextTargetIndex
    Else
        action("TargetIndex") = action("ClosingIndex")
    End If
    Set rows = candidate
    action("Status") = "Planned"
    private_Apply = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    ex_Core.fn_Diagnostic_WriteError "Movement.Preview.Event", VBA.Err.Number, VBA.Err.Source, diagnostic
End Function

Private Sub private_PreserveArrivalFoodSupport( _
    ByVal rows As Collection, _
    ByVal rowId As String, _
    ByVal closingWrites As Object, _
    ByVal links As Object _
)
    Dim row As Object
    Dim cells As Object
    Dim foodSupportHeader As String

    foodSupportHeader = ex_PEB_Export.fn_Config("export.Movement.headers.ArrivalFoodSupport")
    For Each row In rows
        Set cells = row("Cells")
        If VBA.StrComp(ex_PEB_Export.fn_Text(cells(links("ClosingId"))), rowId, VBA.vbTextCompare) = 0 Then
            If VBA.Len(ex_PEB_Export.fn_Text(cells(foodSupportHeader))) > 0 Then
                ' Keep the approved food-support boundary when only the arrival date is corrected.
                closingWrites(foodSupportHeader) = cells(foodSupportHeader)
            End If
            Exit Sub
        End If
    Next row
End Sub

Private Sub private_PlanReturnPlaceholder( _
    ByVal rows As Collection, _
    ByVal headers As Object, _
    ByVal action As Object, _
    ByVal links As Object, _
    ByRef nextTargetIndex As Long _
)
    Dim row As Object
    Dim cells As Object
    Dim writes As Object
    Dim closingWrites As Object
    Dim key As Variant
    Dim identifier As String
    Dim eventHeader As String
    Dim personHeader As String
    Dim statusName As String

    ' The closed placeholder has its own stable link; its dates can later be edited manually.
    identifier = action("RowId") & "/Return"
    eventHeader = ex_PEB_Export.fn_Config("export.Movement.headers.EventName")
    personHeader = ex_PEB_Export.fn_Config("export.Movement.headers.PersonId")
    statusName = ex_PEB_Export.fn_Config("export.Movement.status.ReturnToOtherUnit")
    For Each row In rows
        Set cells = row("Cells")
        If cells(links("OpeningId")) = identifier Then
            If cells(personHeader) <> action("PersonId") Or cells(eventHeader) <> statusName Then
                private_Fail "AppliedMovementChanged", identifier
            End If
            Exit Sub
        End If
    Next row

    ' Create one closed zero-length period before the new business-trip period.
    Set writes = private_CopyFields(action("Writes"))
    Set closingWrites = action("ClosingWrites")
    writes(eventHeader) = statusName
    For Each key In closingWrites.Keys
        writes(key) = closingWrites(key)
    Next key
    writes(ex_PEB_Export.fn_Config("export.Movement.headers.TicketDocument")) = VBA.vbNullString
    For Each key In VBA.Array("DepartureDate", "DepartureFoodSupport", "ArrivalDate", "ArrivalFoodSupport")
        writes(ex_PEB_Export.fn_Config("export.Movement.headers." & key)) = action("EventDate")
    Next key
    Set cells = VBA.CreateObject("Scripting.Dictionary")
    cells.CompareMode = VBA.vbTextCompare
    For Each key In headers.Keys
        cells.Add VBA.CStr(key), VBA.vbNullString
    Next key
    For Each key In writes.Keys
        private_RequireHeader headers, VBA.CStr(key)
        If headers(key) Then
            private_Fail "FormulaColumnWrite", VBA.CStr(key)
        End If
        cells(key) = writes(key)
    Next key
    writes.Add links("OpeningId"), identifier
    writes.Add links("OpeningHash"), ex_PEB_MovementLinks.fn_FieldsHash(writes)
    cells(links("OpeningId")) = identifier
    cells(links("OpeningHash")) = writes(links("OpeningHash"))
    nextTargetIndex = nextTargetIndex + 1
    Set row = VBA.CreateObject("Scripting.Dictionary")
    row.Add "Cells", cells
    row.Add "Formulas", VBA.CreateObject("Scripting.Dictionary")
    row.Add "SourceRow", 0
    row.Add "TargetRow", nextTargetIndex
    row.Add "OriginRowId", identifier
    row.Add "ClosingRowId", VBA.vbNullString
    rows.Add row
    action.Add "PlaceholderIndex", nextTargetIndex
    action.Add "PlaceholderWrites", writes
End Sub

Private Sub private_ValidateLink( _
    ByVal cells As Object, _
    ByVal idHeader As String, _
    ByVal stampHeader As String, _
    ByVal ids As Object _
)
    Dim identifier As String
    Dim stamp As String

    identifier = ex_PEB_Export.fn_Text(cells(idHeader))
    stamp = ex_PEB_Export.fn_Text(cells(stampHeader))
    If (VBA.Len(identifier) = 0) <> (VBA.Len(stamp) = 0) Then
        private_Fail "IncompleteEventLink", idHeader
    End If
    If VBA.Len(identifier) > 0 Then
        If ids.Exists(identifier) Then
            private_Fail "DuplicateEventLink", identifier
        End If
        ids.Add identifier, True
    End If
End Sub

Private Function private_Existing( _
    ByVal rows As Collection, _
    ByVal values As Object, _
    ByVal action As Object, _
    ByVal writes As Object, _
    ByVal closingWrites As Object, _
    ByVal joins As Object, _
    ByVal links As Object, _
    ByVal headers As Object, _
    ByRef nextTargetIndex As Long _
) As Boolean
    Dim row As Object
    Dim opening As Object
    Dim closing As Object
    Dim operation As String
    Dim identifier As String
    Dim updates As Collection
    Dim entry As Object
    Dim addedOpening As Object

    identifier = values("RowId")
    operation = action("Operation")
    For Each row In rows
        If VBA.StrComp(ex_PEB_Export.fn_Text(row("Cells")(links("OpeningId"))), identifier, VBA.vbBinaryCompare) = 0 Then
            If Not opening Is Nothing Then
                private_Fail "DuplicateEventLink", identifier
            End If
            Set opening = row
        End If
        If VBA.StrComp(ex_PEB_Export.fn_Text(row("Cells")(links("ClosingId"))), identifier, VBA.vbBinaryCompare) = 0 Then
            If Not closing Is Nothing Then
                private_Fail "DuplicateEventLink", identifier
            End If
            Set closing = row
        End If
    Next row
    If opening Is Nothing And closing Is Nothing Then
        Exit Function
    End If
    Select Case operation
        Case "append"
            If opening Is Nothing Or Not closing Is Nothing Then
                private_Fail "UnexpectedEventLinks", identifier
            End If
        Case "close"
            If closing Is Nothing Or Not opening Is Nothing Then
                private_Fail "UnexpectedEventLinks", identifier
            End If
        Case "transition"
            If closing Is Nothing Then
                private_Fail "IncompleteTransitionLinks", identifier
            End If
            If Not opening Is Nothing Then
                If opening Is closing Then
                    private_Fail "InvalidTransitionLinks", identifier
                End If
            End If
    End Select
    Set updates = New Collection
    If Not opening Is Nothing Then
        private_PlanBasisSide opening, writes, links("OpeningHash"), action("OpeningHash"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.DepartureBasis"), identifier, updates, action, "Opening"
    End If
    If Not closing Is Nothing Then
        private_PlanBasisSide closing, closingWrites, links("ClosingHash"), action("ClosingHash"), _
            ex_PEB_Export.fn_Config("export.Movement.headers.ArrivalBasis"), identifier, updates, action, "Closing"
        If Not private_Matches(closing("Cells"), joins) Then
            private_Fail "AppliedMovementChanged", identifier
        End If
    End If
    ' Commit the candidate only after both transition sides passed validation.
    If ex_PEB_MovementConflicts.fn_Pending(identifier) Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.ConflictPending")
    End If
    ' An exported arrival already owns the closing side; add only the new departure.
    If operation = "transition" And opening Is Nothing Then
        Set addedOpening = private_PlanOpening(rows, headers, writes, links, action, nextTargetIndex + 1, updates)
    End If
    For Each entry In updates
        For Each row In rows
            If row("TargetRow") = entry("RowIndex") Then
                row("Cells")(entry("Header")) = entry("After")
                Exit For
            End If
        Next row
    Next entry
    If Not addedOpening Is Nothing Then
        rows.Add addedOpening
        Set opening = addedOpening
        nextTargetIndex = nextTargetIndex + 1
    End If
    If Not opening Is Nothing Then
        opening("OriginRowId") = identifier
        action.Add "TargetIndex", opening("TargetRow")
    Else
        action.Add "TargetIndex", closing("TargetRow")
    End If
    If Not closing Is Nothing Then
        closing("ClosingRowId") = identifier
        action.Add "ClosingIndex", closing("TargetRow")
    End If
    action.Add "UpdateCells", updates
    action("IsUpdate") = (updates.Count > 0)
    If action("IsUpdate") Then
        action("Status") = "Planned"
    Else
        action("Status") = "AlreadyApplied"
    End If
    private_Existing = True
End Function

Private Function private_PlanOpening( _
    ByVal rows As Collection, _
    ByVal headers As Object, _
    ByVal writes As Object, _
    ByVal links As Object, _
    ByVal action As Object, _
    ByVal targetIndex As Long, _
    ByVal updates As Collection _
) As Object
    Dim row As Object
    Dim cells As Object
    Dim entry As Object
    Dim key As Variant
    Dim personHeader As String
    Dim departureHeader As String

    ' Reject an unlinked duplicate instead of creating a second departure.
    personHeader = ex_PEB_Export.fn_Config("export.Movement.headers.PersonId")
    departureHeader = ex_PEB_Export.fn_Config("export.Movement.headers.DepartureDate")
    For Each row In rows
        Set cells = row("Cells")
        If VBA.StrComp(ex_PEB_Export.fn_Text(cells(personHeader)), action("PersonId"), VBA.vbTextCompare) = 0 Then
            If VBA.Len(ex_PEB_Export.fn_Text(cells(departureHeader))) > 0 Then
                If private_Date(cells(departureHeader)) = action("EventDate") And _
                        row("OriginRowId") <> action("RowId") & "/Return" Then
                    private_Fail "ExistingIdentityRequiresLink", VBA.CStr(row("SourceRow"))
                End If
            End If
        End If
    Next row

    ' Build a detached row and journal its writes together with the closing updates.
    Set cells = VBA.CreateObject("Scripting.Dictionary")
    cells.CompareMode = VBA.vbTextCompare
    For Each key In headers.Keys
        cells.Add VBA.CStr(key), VBA.vbNullString
    Next key
    For Each key In writes.Keys
        private_RequireHeader headers, VBA.CStr(key)
        If headers(key) Then
            private_Fail "FormulaColumnWrite", VBA.CStr(key)
        End If
        cells(key) = writes(key)
    Next key
    cells(links("OpeningId")) = action("RowId")
    cells(links("OpeningHash")) = action("OpeningHash")
    For Each key In cells.Keys
        If writes.Exists(key) Or key = links("OpeningId") Or key = links("OpeningHash") Then
            Set entry = private_UpdateCell(targetIndex, VBA.CStr(key), VBA.vbNullString, cells(key))
            updates.Add entry
        End If
    Next key
    Set row = VBA.CreateObject("Scripting.Dictionary")
    row.Add "Cells", cells
    row.Add "Formulas", VBA.CreateObject("Scripting.Dictionary")
    row.Add "SourceRow", 0
    row.Add "TargetRow", targetIndex
    row.Add "OriginRowId", action("RowId")
    row.Add "ClosingRowId", VBA.vbNullString
    Set private_PlanOpening = row
End Function

Private Sub private_Cancel( _
    ByVal rows As Collection, _
    ByVal action As Object, _
    ByVal fields As Object, _
    ByVal joins As Object, _
    ByVal links As Object _
)
    Dim row As Object
    Dim selected As Object
    Dim cells As Object
    Dim key As Variant
    Dim updates As Collection
    Dim entry As Object

    Set updates = New Collection
    For Each row In rows
        Set cells = row("Cells")
        If cells(links("OpeningId")) = action("RowId") Then
            private_Fail "CancellationOpeningNotSupported", action("RowId")
        End If
        If cells(links("ClosingId")) = action("RowId") Then
            If Not selected Is Nothing Then
                private_Fail "DuplicateEventLink", action("RowId")
            End If
            Set selected = row
        End If
    Next row
    If selected Is Nothing Then
        action("Operation") = "cancel"
        action("Status") = "AlreadyApplied"
        action.Add "UpdateCells", updates
        Exit Sub
    End If
    If action("Operation") <> "close" Then
        private_Fail "CancellationOperationNotSupported", action("RowId")
    End If
    Set cells = selected("Cells")
    If cells(links("ClosingHash")) <> ex_PEB_MovementLinks.fn_FieldsHash(fields, cells) Then
        private_Fail "AppliedMovementChanged", action("RowId")
    End If
    If Not private_Matches(cells, joins) Then
        private_Fail "AppliedMovementChanged", action("RowId")
    End If
    For Each key In fields.Keys
        If selected("Formulas").Exists(key) Then
            private_Fail "AppliedMovementFormulaChanged", action("RowId")
        End If
        updates.Add private_UpdateCell(selected("TargetRow"), VBA.CStr(key), cells(key), VBA.vbNullString)
    Next key
    For Each key In VBA.Array("ClosingId", "ClosingHash")
        updates.Add private_UpdateCell(selected("TargetRow"), links(key), cells(links(key)), VBA.vbNullString)
    Next key
    For Each entry In updates
        cells(entry("Header")) = entry("After")
    Next entry
    selected("ClosingRowId") = VBA.vbNullString
    action.Add "ClosingIndex", selected("TargetRow")
    action.Add "TargetIndex", selected("TargetRow")
    action.Add "UpdateCells", updates
    action("IsUpdate") = True
    action("Operation") = "cancel"
    action("Status") = "Planned"
End Sub

Private Sub private_PlanBasisSide( _
    ByVal row As Object, _
    ByVal fields As Object, _
    ByVal hashHeader As String, _
    ByVal newHash As String, _
    ByVal basisHeader As String, _
    ByVal identifier As String, _
    ByVal updates As Collection, _
    ByVal action As Object, _
    ByVal side As String _
)
    Dim cells As Object
    Dim key As Variant
    Dim entry As Object

    Set cells = row("Cells")
    If Not ex_PEB_MovementConflicts.fn_Resolve(action, side, fields, cells, basisHeader, hashHeader) Then
        If cells(hashHeader) <> ex_PEB_MovementLinks.fn_FieldsHash(fields, cells) Then
            private_Fail "AppliedMovementChanged", identifier
        End If
    End If
    For Each key In fields.Keys
        If row("Formulas").Exists(key) Then
            private_Fail "AppliedMovementFormulaChanged", identifier
        End If
        If ex_PEB_MovementLinks.fn_ValueToken(cells(key)) <> ex_PEB_MovementLinks.fn_ValueToken(fields(key)) Then
            If VBA.StrComp(VBA.CStr(key), basisHeader, VBA.vbTextCompare) <> 0 Then
                If Not ex_PEB_MovementConflicts.fn_Overwrite() Then
#If TVO_DEPRECATED Then
                    If Not ex_PEB_MovementRegistry.fn_IsTvoHeader(VBA.CStr(key)) Then
                        private_Fail "AppliedEventChanged", identifier & "/" & VBA.CStr(key)
                    End If
#Else
                    private_Fail "AppliedEventChanged", identifier & "/" & VBA.CStr(key)
#End If
                End If
            End If
            Set entry = private_UpdateCell(row("TargetRow"), VBA.CStr(key), cells(key), fields(key))
            updates.Add entry
        End If
    Next key
    newHash = ex_PEB_MovementLinks.fn_FieldsHash(fields)
    If cells(hashHeader) <> newHash Then
        Set entry = private_UpdateCell(row("TargetRow"), hashHeader, cells(hashHeader), newHash)
        updates.Add entry
    End If
End Sub

Private Function private_CopyFields(ByVal fields As Object) As Object
    Dim copy As Object
    Dim key As Variant

    Set copy = VBA.CreateObject("Scripting.Dictionary")
    For Each key In fields.Keys
        copy.Add key, fields(key)
    Next key
    Set private_CopyFields = copy
End Function

Private Function private_UpdateCell( _
    ByVal rowIndex As Long, _
    ByVal header As String, _
    ByVal before As Variant, _
    ByVal after As Variant _
) As Object
    Dim entry As Object

    Set entry = VBA.CreateObject("Scripting.Dictionary")
    entry.Add "RowIndex", rowIndex
    entry.Add "Header", header
    entry.Add "Before", before
    entry.Add "After", after
    Set private_UpdateCell = entry
End Function

Private Sub private_RequireHeader( _
    ByVal headers As Object, _
    ByVal name As String _
)
    If Not headers.Exists(name) Then
        private_Fail "MissingHeader", name
    End If
End Sub

Private Function private_Mappings( _
    ByVal prefix As String, _
    ByVal values As Object _
) As Object
    Dim result As Object
    Dim index As Long
    Dim field As String
    Dim header As String
    Dim value As Variant

    Set result = VBA.CreateObject("Scripting.Dictionary")
    result.CompareMode = VBA.vbTextCompare
    For index = 1 To ex_PEB_MovementRegistry.fn_Count(prefix & ".count")
        field = ex_PEB_MovementRegistry.fn_Value(prefix & "." & index & ".field")
        header = ex_PEB_MovementRegistry.fn_Value(prefix & "." & index & ".header")
        If VBA.Right$(prefix, 7) = ".fields" Then
            value = ex_PEB_MovementRegistry.fn_WriteValue(values, field, header)
            If VBA.VarType(value) = VBA.vbString Then
                If VBA.Len(value) > 32767 Then
                    private_Fail "MovementValueTooLong", header
                End If
            End If
        Else
            value = ex_PEB_Export.fn_Required(values, field)
            If VBA.Len(ex_PEB_Export.fn_Text(value)) = 0 Then
                private_Fail "EmptyMatchField", field
            End If
        End If
        result.Add header, value
    Next index
    Set private_Mappings = result
End Function

Private Function private_Matches( _
    ByVal cells As Object, _
    ByVal mappings As Object _
) As Boolean
    Dim key As Variant
    Dim actual As Variant

    For Each key In mappings.Keys
        actual = ex_PEB_Export.fn_Required(cells, VBA.CStr(key))
        If VBA.CStr(key) = ex_PEB_Export.fn_Config("export.Movement.headers.Destination") Then
            If VBA.StrComp( _
                    ex_PEB_MovementRegistry.fn_Destination(actual), _
                    ex_PEB_MovementRegistry.fn_Destination(mappings(key)), _
                    VBA.vbTextCompare) <> 0 Then Exit Function
        ElseIf VBA.VarType(mappings(key)) = VBA.vbDate Then
            If VBA.Len(ex_PEB_Export.fn_Text(actual)) = 0 Then
                Exit Function
            End If
            If private_Date(actual) <> mappings(key) Then
                Exit Function
            End If
        ElseIf VBA.StrComp(ex_PEB_Export.fn_Text(actual), ex_PEB_Export.fn_Text(mappings(key)), VBA.vbTextCompare) <> 0 Then
            Exit Function
        End If
    Next key
    private_Matches = True
End Function

Private Function private_CopyRow(ByVal source As Object) As Object
    Dim result As Object
    Dim cells As Object
    Dim key As Variant

    Set result = VBA.CreateObject("Scripting.Dictionary")
    For Each key In source.Keys
        If Not VBA.IsObject(source(key)) Then
            result.Add VBA.CStr(key), source(key)
        End If
    Next key
    Set cells = VBA.CreateObject("Scripting.Dictionary")
    cells.CompareMode = VBA.vbTextCompare
    For Each key In source("Cells").Keys
        cells.Add VBA.CStr(key), source("Cells")(key)
    Next key
    result.Add "Cells", cells
    result.Add "Formulas", source("Formulas")
    Set private_CopyRow = result
End Function

Private Sub private_ValidatePeriods( _
    ByVal rows As Collection, _
    ByVal personId As String _
)
    Dim periods As Collection

    Set periods = private_StatePeriods(rows, personId)
    ex_PEB_MovementValidation.fn_CheckPeriods periods
End Sub

Private Function private_StatePeriods( _
    ByVal rows As Collection, _
    ByVal personId As String _
) As Collection
    Dim result As New Collection
    Dim row As Object
    Dim period As Object
    Dim cells As Object
    Dim headers As Object
    Dim field As Variant
    Dim personHeader As String
    Dim index As Long

    Set headers = VBA.CreateObject("Scripting.Dictionary")
    For Each field In VBA.Array("PersonId", "EventName", "Destination", "DepartureDate", "ArrivalDate", "DepartureOrder", "ArrivalOrder")
        headers(field) = ex_PEB_Export.fn_Config("export.Movement.headers." & field)
    Next field
    personHeader = headers("PersonId")
    For Each row In rows
        index = index + 1
        Set cells = row("Cells")
        If VBA.StrComp(ex_PEB_Export.fn_Text(cells(personHeader)), personId, VBA.vbTextCompare) = 0 Then
            Set period = VBA.CreateObject("Scripting.Dictionary")
            For Each field In headers.Keys
                period(field) = cells(headers(field))
            Next field
            period("ModelIndex") = index
            result.Add period
        End If
    Next row
    Set private_StatePeriods = result
End Function

Private Function private_TryValidatePeriods( _
    ByVal rows As Collection, _
    ByVal personId As String, _
    ByRef diagnostic As String _
) As Boolean
    On Error GoTo EH
    private_ValidatePeriods rows, personId
    private_TryValidatePeriods = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
End Function

Private Function private_Date(ByVal value As Variant) As Date
    If VBA.VarType(value) <> VBA.vbString And VBA.IsNumeric(value) Then
        If VBA.CDbl(value) < 61 Or VBA.CDbl(value) > 2958465# Then
            private_Fail "InvalidMovementDate", VBA.CStr(value)
        End If
        private_Date = VBA.CDate(value)
    Else
        private_Date = ex_PEB_History.fn_Date(value)
    End If
End Function

Private Function private_Message( _
    ByVal code As String, _
    ByVal detail As String _
) As String
    private_Message = VBA.Replace$(VBA.Replace$(ex_PEB_Export.fn_Config("text.MovementPlanError"), "{code}", code), "{detail}", detail)
End Function

Private Sub private_Fail( _
    ByVal code As String, _
    ByVal detail As String _
)
    Err.Raise 5, , private_Message(code, detail)
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------