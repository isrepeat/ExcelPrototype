VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PEB_CandidateEnricher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_absenceProfile As obj_LookupProfile
Private m_latestByName As Object
Private m_absenceTable As obj_DataTable
Private m_absenceRecords As Object
Private m_returnTickets As obj_PEB_ReturnTickets
Private m_generation As Long

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize(ByRef diagnostic As String) As Boolean
    If m_isInitialized Or m_isDisposed Then Exit Function
    Set m_absenceProfile = New obj_LookupProfile
    If Not m_absenceProfile.Initialize("PersonalEventBuilder::lookup.Absence", diagnostic) Then
        Me.Dispose
        Exit Function
    End If
    If VBA.Len(m_absenceProfile.Source.RangeAddress) = 0 Or VBA.LCase$(m_absenceProfile.Source.RangeAddress) = "auto" Then
        diagnostic = "Explicit configuration range required: PersonalEventBuilder::lookup.Absence.source.range"
        Me.Dispose
        Exit Function
    End If
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
    If Not m_absenceProfile Is Nothing Then m_absenceProfile.Dispose
    Set m_returnTickets = Nothing
    Set m_absenceProfile = Nothing
    Set m_latestByName = Nothing
    Set m_absenceTable = Nothing
    Set m_absenceRecords = Nothing
End Sub

Public Sub InvalidateCache()
    Set m_latestByName = Nothing
    Set m_absenceTable = Nothing
    Set m_absenceRecords = Nothing
    Set m_returnTickets = Nothing
End Sub

Public Function PrepareSnapshot( _
    ByVal snapshot As obj_TableSnapshotSet, _
    ByRef diagnostic As String _
) As Boolean
    PrepareSnapshot = private_LoadLatest(diagnostic, snapshot)
End Function

Public Function TryEnrich( _
    ByVal lookupService As obj_LookupService, _
    ByVal candidates As Collection, _
    ByVal orderNumber As String, _
    ByVal orderDate As Variant, _
    ByVal modeId As String, _
    ByRef output As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim record As obj_LookupRecord
    Dim supplement As obj_LookupRecord
    Dim fields As Object
    Dim sequence As Variant
    Dim parsedOrderDate As Date
    Dim hasOrderDate As Boolean
    Dim useOrderIncomingDate As Boolean
    Dim returnMode As Boolean
    Dim copy As obj_LookupRecord
    Dim name As String
    Dim value As Variant
    Dim detail As Variant
    Dim snapshot As obj_TableSnapshotSet
    Dim stageStartedAt As Double
    Dim personId As Variant
    Dim personName As Variant
    Dim documents As Collection

    On Error GoTo EH
    stageStartedAt = VBA.Timer
    Set snapshot = ex_TableSnapshots.fn_Current()
    If Not snapshot Is Nothing Then
        If m_generation <> ex_TableSnapshots.fn_Generation() Then
            Me.InvalidateCache
        End If
        m_generation = ex_TableSnapshots.fn_Generation()
    End If
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Candidate enricher is not initialized."
    returnMode = VBA.InStr(1, "|" & ex_PEB_Export.fn_Config("lookup.ReturnTickets.modes") & "|", "|" & modeId & "|", VBA.vbTextCompare) > 0
    useOrderIncomingDate = VBA.InStr(1, "|" & ex_PEB_Export.fn_Config("lookup.IncomingDate.orderDateModes") & "|", "|" & modeId & "|", VBA.vbTextCompare) > 0
    orderNumber = VBA.Trim$(orderNumber)
    If VBA.Len(orderNumber) > 0 Then
        If orderNumber Like "*[!0-9]*" Then Err.Raise 5, , "Order number must contain digits only."
    End If
    If VBA.Len(private_Text(orderDate)) > 0 Then
        If Not private_TryDate(orderDate, parsedOrderDate) Then Err.Raise 5, , "Order date must be a valid date in dd.mm.yyyy format."
        hasOrderDate = True
    End If

    ' Load and index the extension once, only when candidates exist.
    ex_Core.fn_Diagnostic_WritePerf "Candidates.Enrich.Configure", stageStartedAt
    stageStartedAt = VBA.Timer
    If candidates.Count > 0 Then
        If m_latestByName Is Nothing Then
            If Not private_LoadLatest(diagnostic) Then Exit Function
        End If
        Set output = New Collection
        For Each record In candidates
            Set copy = New obj_LookupRecord
            If Not copy.Initialize(record.Key, record.Fields) Then Err.Raise 5, , "Cannot clone candidate."
            For Each detail In record.DetailNames
                If Not copy.SetDetail(VBA.CStr(detail), record.GetDetail(VBA.CStr(detail))) Then
                    Err.Raise 5, , "Cannot copy candidate details."
                End If
            Next detail
            If Not record.TryGetValue("Personnel.FullName", value) Then Err.Raise 5, , "Candidate name not found."
            name = ex_TableQuery.fn_NormalizeText(private_Text(value))
            If VBA.Len(name) > 0 And m_latestByName.Exists(name) Then
                If Not copy.TryMerge(private_AbsenceRecord(name), False) Then Err.Raise 5, , "Cannot merge absence fields."
            End If
            output.Add copy
        Next record
    Else
        Set output = New Collection
    End If

    ' Always clear derived tickets before choosing the source for this mode.
    ex_Core.fn_Diagnostic_WritePerf "Candidates.Enrich.Absence | Rows=" & VBA.CStr(output.Count), stageStartedAt
    stageStartedAt = VBA.Timer
    For Each record In output
        Set fields = VBA.CreateObject("Scripting.Dictionary")
        fields("Absence.Type") = VBA.vbNullString
        fields("Absence.Destination") = VBA.vbNullString
        fields("Absence.DepartureDate") = VBA.vbNullString
        fields("Absence.DurationDays") = VBA.vbNullString
        fields("Absence.TicketNumber") = VBA.vbNullString
        fields("Absence.TicketDate") = VBA.vbNullString
        Set supplement = New obj_LookupRecord
        If Not supplement.Initialize(record.Key, fields) Then Err.Raise 5, , "Cannot initialize empty extension fields."
        If Not record.TryMerge(supplement, False) Then Err.Raise 5, , "Cannot fill empty extension fields."
        Set fields = VBA.CreateObject("Scripting.Dictionary")
        fields("Absence.TicketNumber") = VBA.vbNullString
        fields("Absence.TicketDate") = VBA.vbNullString
        fields("Absence.PlannedArrivalDate") = VBA.vbNullString
        fields("Event.IncomingDate") = VBA.vbNullString
        If hasOrderDate And useOrderIncomingDate Then
            fields("Event.IncomingDate") = VBA.Format$(parsedOrderDate, "dd.mm.yyyy")
        End If
        fields("Event.Date") = VBA.vbNullString
        Set supplement = New obj_LookupRecord
        If Not supplement.Initialize(record.Key, fields) Then
            Err.Raise 5, , "Cannot initialize empty ticket fields."
        End If
        If Not record.TryMerge(supplement, True) Then
            Err.Raise 5, , "Cannot clear previous ticket fields."
        End If
        sequence = VBA.vbNullString
        If Not returnMode And record.TryGetValue("Absence.TicketSequence", sequence) Then
            If private_IsPositiveInteger(private_Text(sequence)) And hasOrderDate And VBA.Len(orderNumber) > 0 Then
                Set fields = VBA.CreateObject("Scripting.Dictionary")
                fields("Absence.TicketNumber") = VBA.CStr(VBA.Year(parsedOrderDate)) & "/" & orderNumber & "/" & private_Text(sequence)
                fields("Absence.TicketDate") = VBA.Format$(parsedOrderDate, "dd.mm.yyyy")
                Set supplement = New obj_LookupRecord
                If Not supplement.Initialize(record.Key, fields) Then Err.Raise 5, , "Cannot initialize ticket fields."
                If Not record.TryMerge(supplement, True) Then Err.Raise 5, , "Cannot apply ticket fields."
            End If
        End If
    Next record
    ex_Core.fn_Diagnostic_WritePerf "Candidates.Enrich.DerivedFields", stageStartedAt
    stageStartedAt = VBA.Timer
    If returnMode Then
        If Not snapshot Is Nothing Then
            Set m_returnTickets = snapshot.Prepared("ReturnTickets")
        End If
        If m_returnTickets Is Nothing Then
            Set m_returnTickets = New obj_PEB_ReturnTickets
            If Not m_returnTickets.Initialize() Then
                Err.Raise 5, , "Cannot initialize return ticket provider."
            End If
        End If
        If Not m_returnTickets.TryEnrich(output, diagnostic) Then
            Set output = Nothing
            Exit Function
        End If
    End If
    ex_Core.fn_Diagnostic_WritePerf "Candidates.Enrich.ReturnTickets | Enabled=" & VBA.CStr(returnMode), stageStartedAt
    stageStartedAt = VBA.Timer
    If InStr(1, "|" & ex_PEB_Export.fn_Config("lookup.ReturnHospital.modes") & "|", "|" & modeId & "|", vbTextCompare) > 0 Or _
            ex_PEB_AnnualContinuation.fn_Enabled(modeId) Then
        If Not snapshot Is Nothing Then
            Set m_returnTickets = snapshot.Prepared("ReturnTickets")
        End If
        If m_returnTickets Is Nothing Then
            Set m_returnTickets = New obj_PEB_ReturnTickets
            If Not m_returnTickets.Initialize() Then Err.Raise 5, , "Cannot initialize Movement provider."
        End If
        If Not m_returnTickets.TryEnrichHospital(output, diagnostic) Then
            Set output = Nothing
            Exit Function
        End If
    End If
    ex_Core.fn_Diagnostic_WritePerf "Candidates.Enrich.ReturnHospital", stageStartedAt
    stageStartedAt = VBA.Timer
    If ex_PEB_AnnualContinuation.fn_Enabled(modeId) Then
        For Each record In output
            If Not record.TryGetValue("Personnel.TaxId", personId) Then Err.Raise 5, , "Annual continuation candidate is missing Personnel.TaxId."
            If Not record.TryGetValue("Personnel.FullName", personName) Then Err.Raise 5, , "Annual continuation candidate is missing Personnel.FullName."
            Set fields = ex_PEB_AnnualContinuation.fn_Fields(CStr(personId), CStr(personName), orderDate)
            Set supplement = New obj_LookupRecord
            If Not supplement.Initialize(record.Key, fields) Then Err.Raise 5, , "Cannot initialize annual continuation fields."
            If Not record.TryMerge(supplement, True) Then Err.Raise 5, , "Cannot apply annual continuation fields."
            Set fields = CreateObject("Scripting.Dictionary")
            If Not record.TryGetValue("Continuation.TicketOrder", value) Then Err.Raise 5, , "Missing interrupted vacation order."
            fields("Absence.TicketDate") = m_returnTickets.TicketDate(value)
            Set supplement = New obj_LookupRecord
            If Not supplement.Initialize(record.Key, fields) Then Err.Raise 5, , "Cannot initialize original vacation ticket date."
            If Not record.TryMerge(supplement, True) Then Err.Raise 5, , "Cannot apply original vacation ticket date."
            Set fields = CreateObject("Scripting.Dictionary")
            fields("SourceAlias") = ex_PEB_Export.fn_Config("lookup.ReturnTickets.sourceAlias")
            fields("Document.Name") = ex_PEB_Export.fn_Config("extensions.role.MedicalCard.document")
            If Not record.TryGetValue("Continuation.DocumentNumber", value) Then Err.Raise 5, , "Missing continuation medical card number field."
            fields("DocumentNumber") = value
            fields("DocumentDate") = Format$(parsedOrderDate, "dd.mm.yyyy")
            fields("DocumentDateSourceAlias") = ex_PEB_Export.fn_Config("legend.orderDate.sourceAlias")
            Set supplement = New obj_LookupRecord
            If Not supplement.Initialize(record.Key & "|MedicalCard", fields) Then Err.Raise 5, , "Cannot initialize continuation medical card."
            Set documents = New Collection
            documents.Add supplement
            If Not record.SetDetail("Document", documents) Then Err.Raise 5, , "Cannot attach continuation medical card."
        Next record
    End If
    ex_Core.fn_Diagnostic_WritePerf "Candidates.Enrich.AnnualContinuation", stageStartedAt
    TryEnrich = True
    Exit Function
EH:
    Set output = Nothing
    diagnostic = "Enrich candidates: " & VBA.Err.Description
End Function

' //
' // Private
' //
Private Function private_LoadLatest( _
    ByRef diagnostic As String, _
    Optional ByVal snapshot As obj_TableSnapshotSet = Nothing _
) As Boolean
    Dim tableQueryService As obj_TableQueryService
    Dim tableQuery As obj_TableQuery
    Dim dataTable As obj_DataTable
    Dim byName As Object
    Dim datesByName As Object
    Dim column As Object
    Dim name As String
    Dim value As Variant
    Dim departureDate As Date
    Dim rowIndex As Long
    Dim prepared As Object

    On Error GoTo EH
    If snapshot Is Nothing Then
        Set snapshot = ex_TableSnapshots.fn_Current()
    End If
    If Not snapshot Is Nothing Then
        If snapshot.Prepared.Exists("Absence.Latest") Then
            Set prepared = snapshot.Prepared("Absence.Latest")
            Set m_latestByName = prepared("latest")
            Set m_absenceTable = prepared("table")
            Set m_absenceRecords = prepared("records")
            private_LoadLatest = True
            Exit Function
        End If
    End If
    Set tableQueryService = New obj_TableQueryService
    Set tableQuery = New obj_TableQuery
    If Not tableQueryService.Initialize() Or Not tableQuery.Initialize() Then Err.Raise 5, , "Cannot initialize extension query."
    tableQueryService.Backend = QueryAdo
    For Each column In m_absenceProfile.Columns
        tableQuery.AddColumn column("header")
    Next column
    If snapshot Is Nothing Then
        If Not tableQueryService.TryExecute(m_absenceProfile.Source, tableQuery, dataTable, diagnostic) Then
            GoTo Cleanup
        End If
    Else
        If Not snapshot.TryQuery(m_absenceProfile.Source, tableQuery, dataTable, diagnostic) Then
            GoTo Cleanup
        End If
    End If
    Set byName = VBA.CreateObject("Scripting.Dictionary")
    byName.CompareMode = VBA.vbTextCompare
    Set datesByName = VBA.CreateObject("Scripting.Dictionary")
    For rowIndex = 1 To dataTable.RowCount
        name = ex_TableQuery.fn_NormalizeText(private_Text(dataTable.ValueAt(rowIndex, _
            ex_TableQuery.fn_HeaderIndex(dataTable.Headers, m_absenceProfile.HeaderFor("Personnel.FullName")))))
        If VBA.Len(name) > 0 Then
            value = dataTable.ValueAt(rowIndex, ex_TableQuery.fn_HeaderIndex(dataTable.Headers, m_absenceProfile.HeaderFor("Absence.DepartureDate")))
            If Not private_TryDate(value, departureDate) Then Err.Raise 5, , "Invalid departure date in extension row: " & VBA.CStr(rowIndex)
            ' Equal dates use the last physical row in the source table.
            If Not byName.Exists(name) Then
                byName(name) = Array(rowIndex, departureDate)
                datesByName(name) = VBA.CDbl(departureDate)
            ElseIf VBA.CDbl(departureDate) >= datesByName(name) Then
                byName(name) = Array(rowIndex, departureDate)
                datesByName(name) = VBA.CDbl(departureDate)
            End If
        End If
    Next rowIndex
    Set m_latestByName = byName
    Set m_absenceTable = dataTable
    Set dataTable = Nothing
    Set m_absenceRecords = CreateObject("Scripting.Dictionary")
    m_absenceRecords.CompareMode = vbTextCompare
    If Not snapshot Is Nothing Then
        Set prepared = VBA.CreateObject("Scripting.Dictionary")
        Set prepared("latest") = m_latestByName
        Set prepared("table") = m_absenceTable
        Set prepared("records") = m_absenceRecords
        snapshot.Prepared.Add "Absence.Latest", prepared
    End If
    private_LoadLatest = True
Cleanup:
    If Not tableQueryService Is Nothing Then tableQueryService.Dispose
    If Not tableQuery Is Nothing Then tableQuery.Dispose
    If Not dataTable Is Nothing Then dataTable.Dispose
    Exit Function
EH:
    diagnostic = "Load absence extension: " & VBA.Err.Description
    Resume Cleanup
End Function

Private Function private_AbsenceRecord(ByVal name As String) As obj_LookupRecord
    Dim fields As Object
    Dim column As Object
    Dim record As obj_LookupRecord
    Dim location As Variant
    Dim value As Variant

    If m_absenceRecords.Exists(name) Then
        Set private_AbsenceRecord = m_absenceRecords(name)
        Exit Function
    End If
    location = m_latestByName(name)
    Set fields = CreateObject("Scripting.Dictionary")
    For Each column In m_absenceProfile.Columns
        value = m_absenceTable.ValueAt(location(0), ex_TableQuery.fn_HeaderIndex(m_absenceTable.Headers, column("header")))
        If IsError(value) Then Err.Raise 5, , "Excel error in candidate extension column: " & column("header")
        If IsNull(value) Or IsEmpty(value) Then value = vbNullString
        fields(column("name")) = value
    Next column
    fields("Absence.DepartureDate") = Format$(location(1), "dd.mm.yyyy")
    Set record = New obj_LookupRecord
    If Not record.Initialize(name, fields) Then Err.Raise 5, , "Cannot initialize candidate absence record."
    m_absenceRecords.Add name, record
    Set private_AbsenceRecord = record
End Function

Private Function private_TryDate( _
    ByVal value As Variant, _
    ByRef parsedDate As Date _
) As Boolean
    Dim text As String
    Dim parts As Variant
    Dim day As Long
    Dim month As Long
    Dim year As Long

    On Error GoTo EH
    If VBA.IsError(value) Or VBA.IsNull(value) Or VBA.IsEmpty(value) Then Exit Function
    If VBA.VarType(value) = VBA.vbDate Or VBA.IsNumeric(value) Then
        parsedDate = VBA.CDate(value)
    Else
        text = VBA.Trim$(VBA.CStr(value))
        parts = VBA.Split(text, ".")
        If UBound(parts) <> 2 Then Exit Function
        If Not VBA.IsNumeric(parts(0)) Or Not VBA.IsNumeric(parts(1)) Or Not VBA.IsNumeric(parts(2)) Then Exit Function
        day = VBA.CLng(parts(0))
        month = VBA.CLng(parts(1))
        year = VBA.CLng(parts(2))
        parsedDate = VBA.DateSerial(year, month, day)
        If VBA.Day(parsedDate) <> day Or VBA.Month(parsedDate) <> month Or VBA.Year(parsedDate) <> year Then Exit Function
    End If
    private_TryDate = True
EH:
End Function

Private Function private_IsPositiveInteger(ByVal text As String) As Boolean
    If VBA.Len(text) = 0 Then Exit Function
    If text Like "*[!0-9]*" Then Exit Function
    private_IsPositiveInteger = VBA.Val(text) > 0
End Function

Private Function private_Text(ByVal value As Variant) As String
    If VBA.IsError(value) Then Err.Raise 5, , "Excel error in absence field."
    If VBA.IsNull(value) Or VBA.IsEmpty(value) Then Exit Function
    private_Text = VBA.Trim$(VBA.CStr(value))
End Function