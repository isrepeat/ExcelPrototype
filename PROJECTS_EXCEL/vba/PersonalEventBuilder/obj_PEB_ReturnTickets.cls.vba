VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PEB_ReturnTickets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_byIpn As Object
Private m_byName As Object
Private m_orderDates As Object
Private m_orderYear As Long
Private m_hospitalByIpn As Object
Private m_hospitalByName As Object
Private m_hospitalNames As Object
Private m_loadingSnapshot As obj_TableSnapshotSet

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
Public Function Initialize() As Boolean
    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isInitialized = False
    m_isDisposed = True
    Me.InvalidateCache
End Sub

Public Sub InvalidateCache()
    Set m_hospitalNames = Nothing
    Set m_hospitalByIpn = Nothing
    Set m_hospitalByName = Nothing
    Set m_byIpn = Nothing
    Set m_byName = Nothing
    Set m_orderDates = Nothing
End Sub

Public Function PrepareSnapshot( _
    ByVal snapshot As obj_TableSnapshotSet, _
    ByRef diagnostic As String _
) As Boolean
    On Error GoTo EH
    Set m_loadingSnapshot = snapshot
    private_LoadMovement
    private_LoadOrders
    private_LoadHospital
    PrepareSnapshot = True
Cleanup:
    Set m_loadingSnapshot = Nothing
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function

Public Function TryEnrich( _
    ByVal records As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim record As obj_LookupRecord
    Dim supplement As obj_LookupRecord
    Dim fields As Object
    Dim index As Object
    Dim scalar As Variant
    Dim key As String
    Dim pair As Variant
    Dim orderNumber As String
    Dim temporaryIpn As String
    Dim errorNumber As Long
    Dim errorText As String
    Dim stage As String

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Return ticket provider is not initialized."
    End If
    If records.Count = 0 Then
        TryEnrich = True
        Exit Function
    End If
    stage = "MovementSnapshot"
    private_LoadMovement
    stage = "OrderDates"
    private_LoadOrders
    stage = "CandidateTickets"
    temporaryIpn = ex_TableQuery.fn_NormalizeText(ex_PEB_Export.fn_Config("lookup.ReturnTickets.temporaryIpn"))
    For Each record In records
        Set fields = VBA.CreateObject("Scripting.Dictionary")
        fields("Absence.TicketNumber") = VBA.vbNullString
        fields("Absence.TicketDate") = VBA.vbNullString
        fields("Absence.Destination") = VBA.vbNullString
        fields("Absence.PlannedArrivalDate") = VBA.vbNullString
        scalar = VBA.vbNullString
        If Not record.TryGetValue("Personnel.TaxId", scalar) Then
            Err.Raise 5, , "Return ticket candidate is missing Personnel.TaxId."
        End If
        key = ex_TableQuery.fn_NormalizeText(private_Text(scalar))
        Set index = m_byIpn
        If VBA.StrComp(key, temporaryIpn, VBA.vbTextCompare) = 0 Then
            scalar = VBA.vbNullString
            If Not record.TryGetValue("Personnel.FullName", scalar) Then
                Err.Raise 5, , "Return ticket candidate is missing Personnel.FullName."
            End If
            key = ex_TableQuery.fn_NormalizeText(private_Text(scalar))
            Set index = m_byName
        End If
        If VBA.Len(key) > 0 And index.Exists(key) Then
            pair = index(key)
            fields("Absence.Destination") = pair(2)
            If Len(private_Text(pair(3))) > 0 Then
                If Not IsDate(pair(3)) And Not IsNumeric(pair(3)) Then
                    Err.Raise 5, , "Invalid planned arrival date for Movement candidate: " & key
                End If
                fields("Absence.PlannedArrivalDate") = Format$(CDate(pair(3)), "dd.mm.yyyy")
            End If
            orderNumber = private_OrderNumber(pair(1))
            If m_orderDates.Exists(orderNumber) Then
                fields("Absence.TicketNumber") = pair(0)
                fields("Absence.TicketDate") = VBA.Format$(m_orderDates(orderNumber), "dd.mm.yyyy")
            End If
        End If
        Set supplement = New obj_LookupRecord
        If Not supplement.Initialize(record.Key, fields) Then
            Err.Raise 5, , "Cannot initialize return ticket fields."
        End If
        If Not record.TryMerge(supplement, True) Then
            Err.Raise 5, , "Cannot apply return ticket fields."
        End If
    Next record
    TryEnrich = True
    Exit Function
EH:
    errorNumber = VBA.Err.Number
    errorText = VBA.Err.Description
    diagnostic = VBA.Replace$(ex_PEB_Export.fn_Config("text.ReturnTicketsError"), "{details}", errorText)
    ex_Core.fn_Diagnostic_WriteLog "RETURN_TICKETS_FAILED | Stage=" & stage & _
        " | Number=" & VBA.CStr(errorNumber) & " | Description=" & errorText
End Function

Public Function TryEnrichHospital( _
    ByVal records As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim record As obj_LookupRecord
    Dim supplement As obj_LookupRecord
    Dim fields As Object
    Dim index As Object
    Dim value As Variant
    Dim key As String
    Dim field As Variant
    Dim cached As Object
    Dim rx As Object
    Dim matches As Object

    On Error GoTo EH_HOSPITAL
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Movement provider is not initialized."
    If records.Count > 0 Then private_LoadHospital
    For Each record In records
        Set fields = CreateObject("Scripting.Dictionary")
        For Each field In Array("Hospital", "HospitalShortName", "Document", "DocumentNumber")
            fields("Movement." & field) = vbNullString
        Next field
        If Not record.TryGetValue("Personnel.TaxId", value) Then Err.Raise 5, , "Hospital candidate is missing Personnel.TaxId."
        key = ex_TableQuery.fn_NormalizeText(private_Text(value))
        Set index = m_hospitalByIpn
        If key = ex_TableQuery.fn_NormalizeText(ex_PEB_Export.fn_Config("lookup.ReturnTickets.temporaryIpn")) Then
            If Not record.TryGetValue("Personnel.FullName", value) Then Err.Raise 5, , "Hospital candidate is missing Personnel.FullName."
            key = ex_TableQuery.fn_NormalizeText(private_Text(value))
            Set index = m_hospitalByName
        End If
        If Len(key) > 0 And index.Exists(key) Then
            Set cached = index(key)
            ' Parse only the matched person's basis, once per cached source row.
            If Not cached.Exists("Parsed") Then
                value = ex_TableQuery.fn_NormalizeText(cached("Movement.Hospital"))
                If m_hospitalNames.Exists(value) Then
                    cached("Movement.Hospital") = m_hospitalNames(value)(0)
                    cached("Movement.HospitalShortName") = m_hospitalNames(value)(1)
                End If
                If rx Is Nothing Then
                    Set rx = CreateObject("VBScript.RegExp")
                    rx.Pattern = ex_PEB_Export.fn_Config("lookup.ReturnHospital.cardPattern")
                    rx.IgnoreCase = True
                End If
                ' VBScript RegExp does not treat nonbreaking spaces as whitespace.
                value = VBA.Replace$(cached("Basis"), VBA.ChrW$(160), " ")
                value = VBA.Replace$(value, VBA.ChrW$(8239), " ")
                Set matches = rx.Execute(value)
                If matches.Count > 0 Then
                    cached("Movement.Document") = ex_PEB_Export.fn_Config("lookup.ReturnHospital.cardDocument")
                    cached("Movement.DocumentNumber") = matches(0).SubMatches(0)
                End If
                cached("Parsed") = True
            End If
            For Each field In fields.Keys
                fields(field) = cached(field)
            Next field
        End If
        Set supplement = New obj_LookupRecord
        If Not supplement.Initialize(record.Key, fields) Then Err.Raise 5, , "Cannot initialize hospital return fields."
        If Not record.TryMerge(supplement, True) Then Err.Raise 5, , "Cannot apply hospital return fields."
    Next record
    TryEnrichHospital = True
    Exit Function
EH_HOSPITAL:
    diagnostic = "Hospital return: " & Err.Description
End Function

Public Function TicketDate(ByVal orderNumber As Variant) As String
    Dim key As String

    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Movement provider is not initialized."
    private_LoadOrders
    key = private_OrderNumber(orderNumber)
    If m_orderDates.Exists(key) Then TicketDate = Format$(m_orderDates(key), "dd.mm.yyyy")
End Function

' //
' // Private
' //
Private Sub private_LoadHospital()
    Dim table As obj_DataTable
    Dim institutions As obj_DataTable
    Dim names As Object
    Dim fields As Object
    Dim destination As String
    Dim key As String
    Dim row As Long
    Dim ipnColumn As Long
    Dim nameColumn As Long
    Dim destinationColumn As Long
    Dim basisColumn As Long
    Dim fullColumn As Long
    Dim shortColumn As Long
    Dim ipn As String
    Dim name As String

    If Not m_hospitalByIpn Is Nothing Then
        Exit Sub
    End If
    Set institutions = private_ReadSaved("lookup.Hospital")
    fullColumn = ex_TableQuery.fn_HeaderIndex(institutions.Headers, ex_PEB_Export.fn_Config("lookup.Hospital.columns.1.header"))
    shortColumn = ex_TableQuery.fn_HeaderIndex(institutions.Headers, ex_PEB_Export.fn_Config("lookup.Hospital.columns.2.header"))
    Set names = CreateObject("Scripting.Dictionary")
    names.CompareMode = vbTextCompare
    For row = 1 To institutions.RowCount
        key = ex_TableQuery.fn_NormalizeText(private_Text(institutions.ValueAt(row, shortColumn)))
        If Len(key) > 0 Then names(key) = Array(private_Text(institutions.ValueAt(row, fullColumn)), private_Text(institutions.ValueAt(row, shortColumn)))
    Next row
    institutions.Dispose
    Set m_hospitalNames = names
    Set table = private_ReadSaved("lookup.ReturnTickets")
    ipnColumn = ex_TableQuery.fn_HeaderIndex(table.Headers, ex_PEB_Export.fn_Config("lookup.ReturnTickets.ipnHeader"))
    nameColumn = ex_TableQuery.fn_HeaderIndex(table.Headers, ex_PEB_Export.fn_Config("lookup.ReturnTickets.nameHeader"))
    destinationColumn = ex_TableQuery.fn_HeaderIndex(table.Headers, ex_PEB_Export.fn_Config("lookup.ReturnTickets.destinationHeader"))
    basisColumn = ex_TableQuery.fn_HeaderIndex(table.Headers, Replace$(ex_PEB_Export.fn_Config("lookup.ReturnHospital.basisHeader"), ".", "#"))
    Set m_hospitalByIpn = CreateObject("Scripting.Dictionary")
    m_hospitalByIpn.CompareMode = vbTextCompare
    Set m_hospitalByName = CreateObject("Scripting.Dictionary")
    m_hospitalByName.CompareMode = vbTextCompare
    ' Retain one complete source row per person, including an empty card number.
    For row = table.RowCount To 1 Step -1
        ipn = ex_TableQuery.fn_NormalizeText(private_Text(table.ValueAt(row, ipnColumn)))
        name = ex_TableQuery.fn_NormalizeText(private_Text(table.ValueAt(row, nameColumn)))
        If (Len(ipn) = 0 Or m_hospitalByIpn.Exists(ipn)) And _
            (Len(name) = 0 Or m_hospitalByName.Exists(name)) Then GoTo NextHospitalRow
        Set fields = CreateObject("Scripting.Dictionary")
        destination = private_Text(table.ValueAt(row, destinationColumn))
        fields("Movement.Hospital") = destination
        fields("Movement.HospitalShortName") = vbNullString
        fields("Movement.Document") = vbNullString
        fields("Movement.DocumentNumber") = vbNullString
        fields("Basis") = private_Text(table.ValueAt(row, basisColumn))
        If Len(ipn) > 0 And Not m_hospitalByIpn.Exists(ipn) Then m_hospitalByIpn.Add ipn, fields
        If Len(name) > 0 And Not m_hospitalByName.Exists(name) Then m_hospitalByName.Add name, fields
NextHospitalRow:
    Next row
    table.Dispose
End Sub

Private Sub private_LoadMovement()
    Dim table As obj_DataTable
    Dim row As Long
    Dim ipnColumn As Long
    Dim nameColumn As Long
    Dim ticketColumn As Long
    Dim orderColumn As Long
    Dim destinationColumn As Long
    Dim arrivalColumn As Long
    Dim arrivalValue As Variant
    Dim ipn As String
    Dim name As String
    Dim ticket As String
    Dim orderNumber As String
    Dim destination As String
    Dim prefix As String

    prefix = "lookup.ReturnTickets"
    If Not m_byIpn Is Nothing Then
        Exit Sub
    End If
    Set table = private_ReadSaved(prefix)
    ipnColumn = ex_TableQuery.fn_HeaderIndex(table.Headers, ex_PEB_Export.fn_Config(prefix & ".ipnHeader"))
    nameColumn = ex_TableQuery.fn_HeaderIndex(table.Headers, ex_PEB_Export.fn_Config(prefix & ".nameHeader"))
    ticketColumn = ex_TableQuery.fn_HeaderIndex(table.Headers, ex_PEB_Export.fn_Config(prefix & ".ticketHeader"))
    destinationColumn = ex_TableQuery.fn_HeaderIndex(table.Headers, ex_PEB_Export.fn_Config(prefix & ".destinationHeader"))
    arrivalColumn = ex_TableQuery.fn_HeaderIndex( _
        table.Headers, _
        VBA.Replace$(ex_PEB_Export.fn_Config(prefix & ".arrivalHeader"), ".", "#") _
    )
    ' ACE exposes periods in Excel column headers as hash characters.
    orderColumn = ex_TableQuery.fn_HeaderIndex( _
        table.Headers, _
        VBA.Replace$(ex_PEB_Export.fn_Config(prefix & ".orderHeader"), ".", "#") _
    )
    Set m_byIpn = VBA.CreateObject("Scripting.Dictionary")
    m_byIpn.CompareMode = VBA.vbTextCompare
    Set m_byName = VBA.CreateObject("Scripting.Dictionary")
    m_byName.CompareMode = VBA.vbTextCompare

    ' Keep the destination and ticket values from the same last complete physical row.
    For row = table.RowCount To 1 Step -1
        ticket = private_Text(table.ValueAt(row, ticketColumn))
        orderNumber = private_Text(table.ValueAt(row, orderColumn))
        If VBA.Len(ticket) > 0 And VBA.Len(orderNumber) > 0 Then
            destination = private_Text(table.ValueAt(row, destinationColumn))
            arrivalValue = table.ValueAt(row, arrivalColumn)
            ipn = ex_TableQuery.fn_NormalizeText(private_Text(table.ValueAt(row, ipnColumn)))
            name = ex_TableQuery.fn_NormalizeText(private_Text(table.ValueAt(row, nameColumn)))
            If VBA.Len(ipn) > 0 And Not m_byIpn.Exists(ipn) Then
                m_byIpn.Add ipn, VBA.Array(ticket, orderNumber, destination, arrivalValue)
            End If
            If VBA.Len(name) > 0 And Not m_byName.Exists(name) Then
                m_byName.Add name, VBA.Array(ticket, orderNumber, destination, arrivalValue)
            End If
        End If
    Next row
    table.Dispose
End Sub

Private Sub private_LoadOrders()
    Dim table As obj_DataTable
    Dim numberColumn As Long
    Dim dateColumn As Long
    Dim row As Long
    Dim number As String
    Dim dateValue As Variant
    Dim orderDate As Date

    If Not m_orderDates Is Nothing And m_orderYear = VBA.Year(VBA.Date) Then
        Exit Sub
    End If
    Set table = private_ReadSaved("orders")
    numberColumn = ex_TableQuery.fn_HeaderIndex(table.Headers, ex_PEB_Export.fn_Config("orders.numberHeader"))
    dateColumn = ex_TableQuery.fn_HeaderIndex(table.Headers, ex_PEB_Export.fn_Config("orders.dateHeader"))
    Set m_orderDates = VBA.CreateObject("Scripting.Dictionary")
    m_orderYear = VBA.Year(VBA.Date)
    For row = 1 To table.RowCount
        number = private_OrderNumber(table.ValueAt(row, numberColumn))
        dateValue = table.ValueAt(row, dateColumn)
        If VBA.Len(number) > 0 And (VBA.IsDate(dateValue) Or VBA.IsNumeric(dateValue)) Then
            orderDate = VBA.CDate(dateValue)
            If VBA.Year(orderDate) = m_orderYear Then
                If m_orderDates.Exists(number) Then
                    If m_orderDates(number) <> orderDate Then
                        Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.OrderDuplicate"), "{key}", number)
                    End If
                Else
                    m_orderDates.Add number, orderDate
                End If
            End If
        End If
    Next row
    table.Dispose
End Sub

Private Function private_ReadSaved(ByVal prefix As String) As obj_DataTable
    Dim source As obj_TableSource
    Dim query As obj_TableQuery
    Dim service As obj_TableQueryService
    Dim table As obj_DataTable
    Dim diagnostic As String
    Dim failure As String

    On Error GoTo EH
    Set source = New obj_TableSource
    Set query = New obj_TableQuery
    Set service = New obj_TableQueryService
    If Not source.Initialize() Or Not query.Initialize() Or Not service.Initialize() Then Err.Raise 5, , "Cannot initialize snapshot lookup."
    source.WorkbookPath = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config(prefix & ".source.path"))
    source.SheetName = ex_PEB_Export.fn_Config(prefix & ".source.sheet")
    source.RangeAddress = ex_PEB_Export.fn_Config(prefix & ".source.range")
    If m_loadingSnapshot Is Nothing Then
        If Not service.TryExecute(source, query, table, diagnostic) Then Err.Raise 5, , diagnostic
    Else
        If Not m_loadingSnapshot.TryQuery(source, query, table, diagnostic) Then Err.Raise 5, , diagnostic
    End If
    Set private_ReadSaved = table
Cleanup:
    If Not source Is Nothing Then source.Dispose
    If Not query Is Nothing Then query.Dispose
    If Not service Is Nothing Then service.Dispose
    If VBA.Len(failure) > 0 Then
        On Error GoTo 0
        Err.Raise 5, , failure
    End If
    Exit Function
EH:
    failure = VBA.Err.Description
    Resume Cleanup
End Function

Private Function private_OrderNumber(ByVal value As Variant) As String
    Dim text As String

    text = VBA.Replace$(private_Text(value), " ", VBA.vbNullString)
    If VBA.Len(text) > 0 And VBA.IsNumeric(text) Then
        text = VBA.CStr(VBA.CLng(text))
    End If
    private_OrderNumber = text
End Function

Private Function private_Text(ByVal value As Variant) As String
    If VBA.IsError(value) Then
        Err.Raise 5, , "Excel error in return ticket source."
    End If
    If VBA.IsNull(value) Or VBA.IsEmpty(value) Then
        Exit Function
    End If
    private_Text = VBA.Trim$(VBA.CStr(value))
End Function