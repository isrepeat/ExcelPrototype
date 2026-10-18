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
Private m_movementStamp As String
Private m_ordersStamp As String
Private m_orderYear As Long

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
    Set m_byIpn = Nothing
    Set m_byName = Nothing
    Set m_orderDates = Nothing
    m_movementStamp = VBA.vbNullString
    m_ordersStamp = VBA.vbNullString
End Sub

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
            fields("Absence.PlannedArrivalDate") = pair(3)
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

' //
' // Private
' //
Private Sub private_LoadMovement()
    Dim table As obj_DataTable
    Dim stamp As String
    Dim row As Long
    Dim ipnColumn As Long
    Dim nameColumn As Long
    Dim ticketColumn As Long
    Dim orderColumn As Long
    Dim destinationColumn As Long
    Dim arrivalColumn As Long
    Dim arrivalValue As Variant
    Dim arrivalDate As String
    Dim ipn As String
    Dim name As String
    Dim ticket As String
    Dim orderNumber As String
    Dim destination As String
    Dim prefix As String

    prefix = "lookup.ReturnTickets"
    stamp = private_Stamp(ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config(prefix & ".source.path")))
    If Not m_byIpn Is Nothing And stamp = m_movementStamp Then
        Exit Sub
    End If
    m_movementStamp = VBA.vbNullString
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
            arrivalDate = VBA.vbNullString
            If VBA.Len(private_Text(arrivalValue)) > 0 Then
                If Not VBA.IsDate(arrivalValue) And Not VBA.IsNumeric(arrivalValue) Then
                    Err.Raise 5, , "Invalid planned arrival date in Movement snapshot: row " & VBA.CStr(row)
                End If
                arrivalDate = VBA.Format$(VBA.CDate(arrivalValue), "dd.mm.yyyy")
            End If
            ipn = ex_TableQuery.fn_NormalizeText(private_Text(table.ValueAt(row, ipnColumn)))
            name = ex_TableQuery.fn_NormalizeText(private_Text(table.ValueAt(row, nameColumn)))
            If VBA.Len(ipn) > 0 And Not m_byIpn.Exists(ipn) Then
                m_byIpn.Add ipn, VBA.Array(ticket, orderNumber, destination, arrivalDate)
            End If
            If VBA.Len(name) > 0 And Not m_byName.Exists(name) Then
                m_byName.Add name, VBA.Array(ticket, orderNumber, destination, arrivalDate)
            End If
        End If
    Next row
    table.Dispose
    m_movementStamp = stamp
End Sub

Private Sub private_LoadOrders()
    Dim table As obj_DataTable
    Dim stamp As String
    Dim numberColumn As Long
    Dim dateColumn As Long
    Dim row As Long
    Dim number As String
    Dim dateValue As Variant
    Dim orderDate As Date

    stamp = private_Stamp(ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("orders.source.path")))
    If Not m_orderDates Is Nothing And stamp = m_ordersStamp And m_orderYear = VBA.Year(VBA.Date) Then
        Exit Sub
    End If
    m_ordersStamp = VBA.vbNullString
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
    m_ordersStamp = stamp
End Sub

Private Function private_ReadSaved(ByVal prefix As String) As obj_DataTable
    Dim source As obj_TableSource
    Dim query As obj_TableQuery
    Dim service As obj_TableQueryService
    Dim table As obj_DataTable
    Dim fileSystem As Object
    Dim original As String
    Dim snapshot As String
    Dim stamp As String
    Dim diagnostic As String
    Dim errorNumber As Long
    Dim errorText As String

    On Error GoTo EH
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    original = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config(prefix & ".source.path"))
    stamp = private_Stamp(original)
    snapshot = fileSystem.BuildPath(fileSystem.GetSpecialFolder(2).Path, fileSystem.GetTempName & "." & fileSystem.GetExtensionName(original))
    ' Copy disk contents, including when the source workbook has unsaved live edits.
    fileSystem.CopyFile original, snapshot, False
    Set source = New obj_TableSource
    Set query = New obj_TableQuery
    Set service = New obj_TableQueryService
    If Not source.Initialize() Or Not query.Initialize() Or Not service.Initialize() Then
        Err.Raise 5, , "Cannot initialize saved Movement lookup."
    End If
    source.WorkbookPath = snapshot
    source.SheetName = ex_PEB_Export.fn_Config(prefix & ".source.sheet")
    source.RangeAddress = ex_PEB_Export.fn_Config(prefix & ".source.range")
    If VBA.LCase$(source.RangeAddress) = "auto" Then
        Err.Raise 5, , "Saved Movement lookup requires an explicit range."
    End If
    service.Backend = QueryAdo
    If Not service.TryExecute(source, query, table, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    If private_Stamp(original) <> stamp Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.ReturnTicketsSourceChanged")
    End If
Cleanup:
    On Error Resume Next
    If Not service Is Nothing Then
        service.Dispose
    End If
    If Not query Is Nothing Then
        query.Dispose
    End If
    If Not source Is Nothing Then
        source.Dispose
    End If
    If VBA.Len(snapshot) > 0 Then
        If fileSystem.FileExists(snapshot) Then
            fileSystem.DeleteFile snapshot
        End If
    End If
    If VBA.Err.Number <> 0 And errorNumber = 0 Then
        errorNumber = VBA.Err.Number
        errorText = VBA.Err.Description
    End If
    On Error GoTo 0
    If errorNumber <> 0 Then
        If Not table Is Nothing Then
            table.Dispose
        End If
        Err.Raise errorNumber, , errorText
    End If
    Set private_ReadSaved = table
    Exit Function
EH:
    errorNumber = VBA.Err.Number
    errorText = VBA.Err.Description
    Resume Cleanup
End Function

Private Function private_Stamp(ByVal path As String) As String
    Dim fileSystem As Object
    Dim file As Object

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FileExists(path) Then
        Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.ReturnTicketsSourceMissing"), "{path}", path)
    End If
    Set file = fileSystem.GetFile(path)
    private_Stamp = VBA.CStr(file.DateLastModified) & "|" & VBA.CStr(file.Size)
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