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
Private m_latestRecords As Collection

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
    Set m_absenceProfile = Nothing
    Set m_latestRecords = Nothing
End Sub

Public Sub InvalidateCache()
    Set m_latestRecords = Nothing
End Sub

Public Function TryEnrich( _
    ByVal lookupService As obj_LookupService, _
    ByVal candidates As Collection, _
    ByVal orderNumber As String, _
    ByVal orderDate As Variant, _
    ByRef output As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim record As obj_LookupRecord
    Dim supplement As obj_LookupRecord
    Dim fields As Object
    Dim sequence As Variant
    Dim parsedOrderDate As Date
    Dim hasOrderDate As Boolean

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Candidate enricher is not initialized."
    orderNumber = VBA.Trim$(orderNumber)
    If VBA.Len(orderNumber) > 0 Then
        If orderNumber Like "*[!0-9]*" Then Err.Raise 5, , "Order number must contain digits only."
    End If
    If VBA.Len(private_Text(orderDate)) > 0 Then
        If Not private_TryDate(orderDate, parsedOrderDate) Then Err.Raise 5, , "Order date must be a valid date in dd.mm.yyyy format."
        hasOrderDate = True
    End If

    ' Load and index the extension once, only when candidates exist.
    If candidates.Count > 0 Then
        If m_latestRecords Is Nothing Then
            If Not private_LoadLatest(diagnostic) Then Exit Function
        End If
        If Not lookupService.TryExtend(candidates, m_latestRecords, "Personnel.FullName", "error", False, output, diagnostic) Then Exit Function
    Else
        Set output = New Collection
    End If

    ' Missing extensions remain blank; tickets use the shared order fields.
    For Each record In output
        Set fields = VBA.CreateObject("Scripting.Dictionary")
        fields("Absence.Type") = VBA.vbNullString
        fields("Absence.Destination") = VBA.vbNullString
        fields("Absence.DepartureDate") = VBA.vbNullString
        fields("Absence.TicketNumber") = VBA.vbNullString
        fields("Absence.TicketDate") = VBA.vbNullString
        Set supplement = New obj_LookupRecord
        If Not supplement.Initialize(record.Key, fields) Then Err.Raise 5, , "Cannot initialize empty extension fields."
        If Not record.TryMerge(supplement, False) Then Err.Raise 5, , "Cannot fill empty extension fields."
        sequence = VBA.vbNullString
        If record.TryGetValue("Absence.TicketSequence", sequence) Then
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
    TryEnrich = True
    Exit Function
EH:
    Set output = Nothing
    diagnostic = "Enrich candidates: " & VBA.Err.Description
End Function

' //
' // Private
' //
Private Function private_LoadLatest(ByRef diagnostic As String) As Boolean
    Dim tableQueryService As obj_TableQueryService
    Dim tableQuery As obj_TableQuery
    Dim dataTable As obj_DataTable
    Dim byName As Object
    Dim datesByName As Object
    Dim fields As Object
    Dim column As Object
    Dim record As obj_LookupRecord
    Dim name As String
    Dim key As Variant
    Dim value As Variant
    Dim departureDate As Date
    Dim rowIndex As Long
    Dim rows As Collection

    On Error GoTo EH
    Set tableQueryService = New obj_TableQueryService
    Set tableQuery = New obj_TableQuery
    If Not tableQueryService.Initialize() Or Not tableQuery.Initialize() Then Err.Raise 5, , "Cannot initialize extension query."
    tableQueryService.Backend = QueryAdo
    For Each column In m_absenceProfile.Columns
        tableQuery.AddColumn column("header")
    Next column
    If Not tableQueryService.TryExecute(m_absenceProfile.Source, tableQuery, dataTable, diagnostic) Then GoTo Cleanup
    Set byName = VBA.CreateObject("Scripting.Dictionary")
    byName.CompareMode = VBA.vbTextCompare
    Set datesByName = VBA.CreateObject("Scripting.Dictionary")
    For rowIndex = 1 To dataTable.RowCount
        name = ex_TableQuery.fn_NormalizeText(private_Text(dataTable.ValueAt(rowIndex, _
            ex_TableQuery.fn_HeaderIndex(dataTable.Headers, m_absenceProfile.HeaderFor("Personnel.FullName")))))
        If VBA.Len(name) > 0 Then
            value = dataTable.ValueAt(rowIndex, ex_TableQuery.fn_HeaderIndex(dataTable.Headers, m_absenceProfile.HeaderFor("Absence.DepartureDate")))
            If Not private_TryDate(value, departureDate) Then Err.Raise 5, , "Invalid departure date in extension row: " & VBA.CStr(rowIndex)
            Set fields = VBA.CreateObject("Scripting.Dictionary")
            For Each column In m_absenceProfile.Columns
                value = dataTable.ValueAt(rowIndex, ex_TableQuery.fn_HeaderIndex(dataTable.Headers, column("header")))
                If VBA.IsError(value) Then Err.Raise 5, , "Excel error in extension column: " & column("header")
                If VBA.IsNull(value) Or VBA.IsEmpty(value) Then value = VBA.vbNullString
                fields(column("name")) = value
            Next column
            fields("Absence.DepartureDate") = VBA.Format$(departureDate, "dd.mm.yyyy")
            Set record = New obj_LookupRecord
            If Not record.Initialize(name, fields) Then Err.Raise 5, , "Cannot initialize absence record."
            ' Equal dates use the last physical row in the source table.
            If Not byName.Exists(name) Then
                Set byName(name) = record
                datesByName(name) = VBA.CDbl(departureDate)
            ElseIf VBA.CDbl(departureDate) >= datesByName(name) Then
                Set byName(name) = record
                datesByName(name) = VBA.CDbl(departureDate)
            End If
        End If
    Next rowIndex
    Set rows = New Collection
    For Each key In byName.Keys
        rows.Add byName(key)
    Next key
    Set m_latestRecords = rows
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