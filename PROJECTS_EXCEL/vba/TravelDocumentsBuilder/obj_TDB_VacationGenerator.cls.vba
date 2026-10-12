VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_TDB_VacationGenerator"
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_bindingContext As obj_UiBindingContext
Private m_lookupService As obj_LookupService
Private m_config As Object
Private m_form As Object
Private m_fileSystem As Object
Private m_ordersConnection As Object
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
Public Function Initialize( _
    ByVal bindingContext As obj_UiBindingContext, _
    ByVal lookupService As obj_LookupService _
) As Boolean
    Dim keys As String
    Dim key As Variant
    Dim valueText As String
    Dim kind As Variant
    Dim part As Variant

    If m_isInitialized Or m_isDisposed Then Exit Function
    If bindingContext Is Nothing Or lookupService Is Nothing Then Exit Function
    Set m_bindingContext = bindingContext
    Set m_lookupService = lookupService
    Set m_config = VBA.CreateObject("Scripting.Dictionary")
    Set m_fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    keys = "registry.workbook,registry.sheet,registry.table,registry.columns,registry.arrivalFormula," & _
        "orders.path,orders.rangePrefix,orders.numberHeader,orders.dateHeader," & _
        "kind.ids,abroad.yes,abroad.no,abroad.word,status.active,status.cancelled,status.mark," & _
        "donation.event,donation.orderMissing,donation.orderPrefix,months,datePattern,name.active,name.cancelled," & _
        "duration.connector,duration.vacation.one,duration.vacation.few,duration.vacation.many," & _
        "duration.donation.one,duration.donation.few,duration.donation.many,duration.donation.suffix," & _
        "duration.road.one,duration.road.few,duration.road.many,duration.road.suffix," & _
        "position.reservedPrefix,position.removedPrefix,position.reservedText,position.removedText," & _
        "message.Required,message.Invalid,message.Workbook,message.Table,message.Column,message.Duplicate," & _
        "message.NotFound,message.Exists,message.Restore,message.Replace,message.Path,message.Marker," & _
        "message.Order,message.Tvo,message.Unsupported,message.Success,message.Failed,message.Rollback," & _
        "message.Title,limits.maxDays"
    For Each key In VBA.Split(keys, ",")
        If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("generation." & VBA.CStr(key), valueText) Then Exit Function
        m_config.Add VBA.CStr(key), valueText
    Next key
    For Each kind In VBA.Split(m_config("kind.ids"), ",")
        For Each part In VBA.Array("input", "word", "event")
            key = "kind." & VBA.CStr(kind) & "." & VBA.CStr(part)
            If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("generation." & key, valueText) Then Exit Function
            m_config.Add key, valueText
        Next part
    Next kind
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
    On Error Resume Next
    If Not m_ordersConnection Is Nothing Then m_ordersConnection.Close
    Set m_ordersConnection = Nothing
    Set m_bindingContext = Nothing
    Set m_lookupService = Nothing
    Set m_config = Nothing
    Set m_form = Nothing
    Set m_fileSystem = Nothing
    On Error GoTo 0
End Sub

Public Function TryGenerate( _
    ByVal isUpdate As Boolean, _
    ByRef diagnostic As String _
) As Boolean
    Dim ticketsWorkbook As Workbook
    Dim ticketsTable As ListObject
    Dim mainRow As ListRow
    Dim donationRow As ListRow
    Dim row As ListRow
    Dim orderNo As String
    Dim orderDate As Date
    Dim departureDate As Date
    Dim mainReturn As Date
    Dim finalReturn As Date
    Dim vacationDays As Long
    Dim roadDays As Long
    Dim donationDays As Long
    Dim ticketNo As String
    Dim kindWord As String
    Dim kindEvent As String
    Dim donationOrder As String
    Dim documentPath As String
    Dim matchedPath As String
    Dim temporaryPath As String
    Dim backupPath As String
    Dim templatePath As String
    Dim outputFolder As String
    Dim arrivalFormula As String
    Dim originalData As Variant
    Dim originalFormats As Object
    Dim originalCount As Long
    Dim tableChanged As Boolean
    Dim filePublished As Boolean
    Dim eventsBefore As Boolean
    Dim updatingBefore As Boolean
    Dim guarded As Boolean
    Dim restoreRecord As Boolean
    Dim key As Variant
    Dim index As Long
    Dim rollbackError As String

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    On Error GoTo EH
    diagnostic = VBA.vbNullString
    private_ReadForm
    If m_form("DocumentType") <> private_DocumentCaption() Then private_Fail "Unsupported", m_form("DocumentType")
    For Each key In VBA.Array("PersonName", "PersonId", "PersonRank", "PersonPositionCode", "DocumentTarget", _
            "OrderReference", "DepartureDate", "VacationKind", "VacationAbroad", "VacationStatus", "TemplatePath", "OutputFolderPath")
        If VBA.Len(m_form(key)) = 0 Then private_Fail "Required", private_FieldCaption(VBA.CStr(key))
    Next key
    departureDate = private_Date(m_form("DepartureDate"))
    vacationDays = private_Days("VacationDays", False)
    roadDays = private_Days("RoadDays", True)
    donationDays = private_Days("DonationDays", True)
    private_MapKind kindWord, kindEvent
    If m_form("VacationAbroad") <> m_config("abroad.yes") And m_form("VacationAbroad") <> m_config("abroad.no") Then
        private_Fail "Invalid", private_FieldCaption("VacationAbroad")
    End If
    If m_form("VacationStatus") <> m_config("status.active") And m_form("VacationStatus") <> m_config("status.cancelled") Then
        private_Fail "Invalid", private_FieldCaption("VacationStatus")
    End If
    private_ResolveTvo
    private_ResolveOrder m_form("OrderReference"), orderNo, orderDate, False
    mainReturn = VBA.DateAdd("d", vacationDays + roadDays, departureDate)
    finalReturn = VBA.DateAdd("d", donationDays, mainReturn)
    If donationDays > 0 Then
        private_ResolveOrder VBA.Format$(mainReturn, "dd.mm.yyyy"), donationOrder, orderDate, True
        If VBA.Len(donationOrder) = 0 Then donationOrder = m_config("donation.orderMissing")
        donationOrder = m_config("donation.orderPrefix") & " " & donationOrder
        private_ResolveOrder m_form("OrderReference"), orderNo, orderDate, False
    End If

    ' Validate the workbook, schema and existing records before changing data.
    Set ticketsWorkbook = private_OpenWorkbook()
    If ticketsWorkbook.ReadOnly Then private_Fail "Workbook", ticketsWorkbook.Name
    Set ticketsTable = private_Table(ticketsWorkbook)
    For Each key In VBA.Split(m_config("registry.columns"), "|")
        index = private_Column(ticketsTable, VBA.CStr(key))
    Next key
    private_FindMain ticketsTable, orderNo, mainRow
    If isUpdate And mainRow Is Nothing Then private_Fail "NotFound", m_form("PersonId") & " / " & orderNo
    If mainRow Is Nothing Then
        ticketNo = private_NextNumber(ticketsTable, orderNo, orderDate)
    Else
        ticketNo = private_Cell(mainRow, ticketsTable, 15)
        If VBA.Len(ticketNo) = 0 Then private_Fail "Invalid", private_Header(15)
    End If
    private_FindDonation ticketsTable, ticketNo, donationRow
    templatePath = private_Path(m_form("TemplatePath"))
    outputFolder = private_Path(m_form("OutputFolderPath"))
    If Not m_fileSystem.FileExists(templatePath) Then private_Fail "Path", templatePath
    If Not m_fileSystem.FolderExists(outputFolder) Then private_Fail "Path", outputFolder
    matchedPath = private_FindDocument(outputFolder, ticketNo)
    If isUpdate And VBA.Len(matchedPath) = 0 Then private_Fail "NotFound", ticketNo
    If Not isUpdate And Not mainRow Is Nothing Then
        If VBA.Len(matchedPath) > 0 Then private_Fail "Exists", matchedPath
        If Not private_Confirm("Restore", ticketNo) Then Exit Function
        restoreRecord = True
    End If
    documentPath = private_DocumentPath(outputFolder, ticketNo)
    If VBA.StrComp(templatePath, documentPath, VBA.vbTextCompare) = 0 Then private_Fail "Path", templatePath
    If VBA.StrComp(templatePath, matchedPath, VBA.vbTextCompare) = 0 Then private_Fail "Path", templatePath
    If m_fileSystem.FileExists(documentPath) And VBA.StrComp(documentPath, matchedPath, VBA.vbTextCompare) <> 0 Then
        private_Fail "Exists", documentPath
    End If
    If Not isUpdate And Not restoreRecord And VBA.Len(matchedPath) > 0 Then
        If Not private_Confirm("Replace", matchedPath) Then Exit Function
    End If
    arrivalFormula = private_ArrivalFormula(ticketsTable)
    temporaryPath = m_fileSystem.BuildPath(outputFolder, m_fileSystem.GetTempName & ".docx")
    private_GenerateWord templatePath, temporaryPath, ticketNo, orderDate, departureDate, finalReturn, _
        vacationDays, roadDays, donationDays, kindWord

    ' Snapshot the table so updated and deleted rows can also be restored.
    originalCount = ticketsTable.ListRows.Count
    Set originalFormats = VBA.CreateObject("Scripting.Dictionary")
    If originalCount > 0 Then
        originalData = ticketsTable.DataBodyRange.Formula
        For Each row In ticketsTable.ListRows
            For index = 1 To ticketsTable.ListColumns.Count
                originalFormats.Add VBA.CStr(row.Index) & "/" & VBA.CStr(index), row.Range.Cells(1, index).NumberFormat
            Next index
        Next row
    End If
    eventsBefore = Application.EnableEvents
    updatingBefore = Application.ScreenUpdating
    guarded = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    tableChanged = True
    If mainRow Is Nothing Then Set mainRow = private_AddRow(ticketsTable, donationRow)
    private_WriteRow ticketsTable, mainRow, kindEvent, m_form("DocumentTarget"), orderNo, orderDate, _
        departureDate, vacationDays, roadDays, ticketNo, arrivalFormula
    If donationDays > 0 Then
        If donationRow Is Nothing Then Set donationRow = private_AddRow(ticketsTable, mainRow)
        private_WriteRow ticketsTable, donationRow, m_config("donation.event"), VBA.vbNullString, donationOrder, _
            finalReturn, mainReturn, donationDays, 0, ticketNo, arrivalFormula
    End If

    ' Publish the prepared DOCX while keeping a backup of the previous file.
    If VBA.Len(matchedPath) > 0 Then
        backupPath = m_fileSystem.BuildPath(outputFolder, m_fileSystem.GetTempName & ".bak")
        m_fileSystem.MoveFile matchedPath, backupPath
    End If
    m_fileSystem.MoveFile temporaryPath, documentPath
    filePublished = True
    If donationDays = 0 And Not donationRow Is Nothing Then donationRow.Delete
    If VBA.Len(backupPath) > 0 Then m_fileSystem.DeleteFile backupPath
    tableChanged = False
    filePublished = False
    backupPath = VBA.vbNullString
    Application.EnableEvents = eventsBefore
    Application.ScreenUpdating = updatingBefore
    guarded = False
    m_bindingContext.SetValue "Form", "ReturnDate", VBA.Format$(finalReturn, "dd.mm.yyyy")
    m_bindingContext.SetValue "Form", "TicketNo", ticketNo
    m_bindingContext.SetValue "Form", "DocumentPath", documentPath
    m_bindingContext.SetValue "Form", "GenerationStatus", private_Message("Success", documentPath)
    ex_Core.fn_Diagnostic_WriteLog "TDB_GENERATED | Ticket=" & ticketNo & " | Path=" & documentPath
    TryGenerate = True
    Exit Function
EH:
    diagnostic = private_Message("Failed", VBA.Err.Description)
    ex_Core.fn_Diagnostic_WriteLog "TDB_GENERATION_FAILED | " & diagnostic
    On Error Resume Next
    If filePublished Then
        VBA.Err.Clear
        m_fileSystem.DeleteFile documentPath
        If VBA.Err.Number <> 0 Then rollbackError = rollbackError & VBA.Err.Description & VBA.vbCrLf
    End If
    If VBA.Len(backupPath) > 0 Then
        VBA.Err.Clear
        If m_fileSystem.FileExists(backupPath) Then m_fileSystem.MoveFile backupPath, matchedPath
        If VBA.Err.Number <> 0 Then rollbackError = rollbackError & VBA.Err.Description & VBA.vbCrLf
    End If
    If tableChanged Then
        VBA.Err.Clear
        private_RestoreTable ticketsTable, originalCount, originalData, originalFormats
        If VBA.Err.Number <> 0 Then rollbackError = rollbackError & VBA.Err.Description & VBA.vbCrLf
    End If
    If VBA.Len(temporaryPath) > 0 Then
        VBA.Err.Clear
        If m_fileSystem.FileExists(temporaryPath) Then m_fileSystem.DeleteFile temporaryPath
        If VBA.Err.Number <> 0 Then rollbackError = rollbackError & VBA.Err.Description & VBA.vbCrLf
    End If
    If guarded Then
        Application.EnableEvents = eventsBefore
        Application.ScreenUpdating = updatingBefore
    End If
    If VBA.Len(rollbackError) > 0 Then diagnostic = diagnostic & VBA.vbCrLf & private_Message("Rollback", rollbackError)
    On Error GoTo 0
End Function
' //
' // Private
' //
Private Sub private_ReadForm()
    Dim key As Variant
    Dim value As Variant
    Dim objectValue As Object
    Dim isObject As Boolean

    Set m_form = VBA.CreateObject("Scripting.Dictionary")
    For Each key In VBA.Array("DocumentType", "PersonName", "PersonId", "PersonRank", "PersonPositionCode", _
            "DocumentTarget", "DepartureDate", "OrderReference", "VacationKind", "VacationDays", "RoadDays", _
            "DonationDays", "VacationAbroad", "VacationStatus", "TvoLookup", "TemplatePath", "OutputFolderPath")
        If Not m_bindingContext.TryGetValue("Form", VBA.CStr(key), value, objectValue, isObject) Then
            private_Fail "Required", private_FieldCaption(VBA.CStr(key))
        End If
        If isObject Or VBA.IsError(value) Or VBA.IsNull(value) Then private_Fail "Invalid", private_FieldCaption(VBA.CStr(key))
        m_form.Add VBA.CStr(key), ex_TableQuery.fn_NormalizeText(VBA.CStr(value))
    Next key
    Do While VBA.InStr(m_form("PersonName"), "  ") > 0
        m_form("PersonName") = VBA.Replace$(m_form("PersonName"), "  ", " ")
    Loop
End Sub

Private Function private_DocumentCaption() As String
    If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("text.VacationTicket.DocumentFormTitle", private_DocumentCaption) Then
        VBA.Err.Raise VBA.vbObjectError + 2400
    End If
End Function

Private Function private_FieldCaption(ByVal field As String) As String
    Dim captionKey As String

    captionKey = field
    Select Case field
        Case "PersonName", "PersonId", "PersonRank", "PersonPositionCode", "DocumentType"
            captionKey = field & "Label"
        Case "DocumentTarget", "DepartureDate"
            captionKey = "VacationTicket." & field & "Label"
    End Select
    If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("text." & captionKey, private_FieldCaption) Then
        VBA.Err.Raise VBA.vbObjectError + 2400
    End If
End Function

Private Function private_Message( _
    ByVal key As String, _
    ByVal value As String _
) As String
    private_Message = VBA.Replace$(m_config("message." & key), "{value}", value)
End Function

Private Sub private_Fail( _
    ByVal key As String, _
    ByVal value As String _
)
    VBA.Err.Raise VBA.vbObjectError + 2401, "obj_TDB_VacationGenerator", private_Message(key, value)
End Sub

Private Function private_Confirm( _
    ByVal key As String, _
    ByVal value As String _
) As Boolean
    private_Confirm = ex_WindowsUi.fn_ShowMessage( _
        private_Message(key, value), _
        VBA.vbYesNo + VBA.vbQuestion, _
        m_config("message.Title") _
    ) = VBA.vbYes
End Function

Private Function private_Date(ByVal text As String) As Date
    Dim parts As Variant
    Dim parsed As Date

    parts = VBA.Split(text, ".")
    If UBound(parts) <> 2 Then private_Fail "Invalid", text
    If Not text Like "##.##.####" Then private_Fail "Invalid", text
    parsed = VBA.DateSerial(VBA.CInt(parts(2)), VBA.CInt(parts(1)), VBA.CInt(parts(0)))
    If VBA.Format$(parsed, "dd.mm.yyyy") <> text Then private_Fail "Invalid", text
    private_Date = parsed
End Function

Private Function private_Days( _
    ByVal field As String, _
    ByVal optionalValue As Boolean _
) As Long
    Dim text As String

    text = m_form(field)
    If VBA.Len(text) = 0 And optionalValue Then Exit Function
    If VBA.Len(text) = 0 Then private_Fail "Required", private_FieldCaption(field)
    If text Like "*[!0-9]*" Then private_Fail "Invalid", private_FieldCaption(field)
    If VBA.CDbl(text) > VBA.CDbl(m_config("limits.maxDays")) Then private_Fail "Invalid", private_FieldCaption(field)
    private_Days = VBA.CLng(text)
End Function

Private Sub private_MapKind( _
    ByRef word As String, _
    ByRef eventText As String _
)
    Dim kind As Variant

    For Each kind In VBA.Split(m_config("kind.ids"), ",")
        If VBA.StrComp(m_form("VacationKind"), m_config("kind." & kind & ".input"), VBA.vbTextCompare) = 0 Then
            word = m_config("kind." & kind & ".word")
            eventText = m_config("kind." & kind & ".event")
            Exit Sub
        End If
    Next kind
    private_Fail "Invalid", private_FieldCaption("VacationKind")
End Sub

Private Sub private_ResolveTvo()
    Dim result As obj_LookupResult
    Dim record As obj_LookupRecord
    Dim diagnostic As String
    Dim name As Variant
    Dim ipn As Variant
    Dim position As Variant
    Dim count As Long

    m_form("TvoName") = VBA.vbNullString
    m_form("TvoIpn") = VBA.vbNullString
    m_form("TvoPosition") = VBA.vbNullString
    If VBA.Len(m_form("TvoLookup")) = 0 Then Exit Sub
    If Not m_lookupService.TrySearch(m_form("TvoLookup"), result, diagnostic) Then private_Fail "Tvo", diagnostic
    For Each record In result.Records
        If Not record.TryGetValue("Personnel.FullName", name) Then private_Fail "Tvo", m_form("TvoLookup")
        If Not record.TryGetValue("Personnel.TaxId", ipn) Then private_Fail "Tvo", m_form("TvoLookup")
        If VBA.StrComp(VBA.CStr(name), m_form("TvoLookup"), VBA.vbTextCompare) = 0 Or _
                VBA.StrComp(VBA.CStr(ipn), m_form("TvoLookup"), VBA.vbTextCompare) = 0 Then
            count = count + 1
            If Not record.TryGetValue("Personnel.PositionCode", position) Then private_Fail "Tvo", m_form("TvoLookup")
            If VBA.Len(VBA.CStr(position)) = 0 Then private_Fail "Tvo", m_form("TvoLookup")
            m_form("TvoName") = VBA.CStr(name)
            m_form("TvoIpn") = VBA.CStr(ipn)
            m_form("TvoPosition") = VBA.CStr(position)
        End If
    Next record
    If count <> 1 Or result.HasMore Then private_Fail "Tvo", m_form("TvoLookup")
End Sub

Private Sub private_ResolveOrder( _
    ByVal reference As String, _
    ByRef orderNo As String, _
    ByRef orderDate As Date, _
    ByVal allowMissing As Boolean _
)
    Dim year As Long
    Dim inputDate As Date
    Dim isDateReference As Boolean
    Dim tableRef As String
    Dim recordset As Object
    Dim candidateNo As String
    Dim candidateDate As Date
    Dim count As Long
    Dim value As Variant
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo EH
    orderNo = VBA.vbNullString
    isDateReference = VBA.InStr(reference, ".") > 0
    If isDateReference Then
        inputDate = private_Date(reference)
        year = VBA.Year(inputDate)
    Else
        year = VBA.Year(VBA.Date)
        reference = VBA.Replace$(reference, "№", VBA.vbNullString)
        reference = VBA.Trim$(reference)
        If VBA.Len(reference) = 0 Or reference Like "*[!0-9]*" Then private_Fail "Order", reference
    End If
    If Not ex_TravelDocumentsBuilder.fn_TryGetConfig( _
            m_config("orders.rangePrefix") & VBA.CStr(year), tableRef) Then private_Fail "Order", VBA.CStr(year)
    If m_ordersConnection Is Nothing Then
        Set m_ordersConnection = VBA.CreateObject("ADODB.Connection")
        m_ordersConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""" & _
            VBA.Replace$(private_Path(m_config("orders.path")), """", """""") & _
            """;Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1;ReadOnly=True"";"
    End If
    Set recordset = VBA.CreateObject("ADODB.Recordset")
    recordset.Open "SELECT * FROM " & tableRef, m_ordersConnection, 0, 1
    Do While Not recordset.EOF
        value = recordset.Fields(m_config("orders.numberHeader")).Value
        candidateNo = VBA.vbNullString
        If Not VBA.IsNull(value) Then candidateNo = VBA.Trim$(VBA.Replace$(VBA.CStr(value), "№", VBA.vbNullString))
        value = recordset.Fields(m_config("orders.dateHeader")).Value
        If VBA.IsDate(value) Then
            candidateDate = VBA.DateValue(value)
            If (isDateReference And candidateDate = inputDate) Or (Not isDateReference And candidateNo = reference) Then
                If VBA.Len(candidateNo) = 0 Or candidateNo Like "*[!0-9]*" Then private_Fail "Order", candidateNo
                count = count + 1
                orderNo = candidateNo
                orderDate = candidateDate
            End If
        End If
        recordset.MoveNext
    Loop
    recordset.Close
    If count > 1 Or (count = 0 And Not allowMissing) Then private_Fail "Order", reference
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    On Error Resume Next
    If Not recordset Is Nothing Then recordset.Close
    On Error GoTo 0
    VBA.Err.Raise errorNumber, "Orders", errorDescription
End Sub

Private Function private_OpenWorkbook() As Workbook
    Dim book As Workbook

    For Each book In Application.Workbooks
        If VBA.StrComp(book.Name, m_config("registry.workbook"), VBA.vbTextCompare) = 0 Then
            Set private_OpenWorkbook = book
            Exit Function
        End If
    Next book
    private_Fail "Workbook", m_config("registry.workbook")
End Function

Private Function private_Table(ByVal book As Workbook) As ListObject
    Dim sheet As Worksheet
    Dim table As ListObject

    For Each sheet In book.Worksheets
        If sheet.Name = m_config("registry.sheet") Then
            If sheet.ProtectContents Then private_Fail "Table", sheet.Name
            For Each table In sheet.ListObjects
                If table.Name = m_config("registry.table") Then
                    Set private_Table = table
                    Exit Function
                End If
            Next table
        End If
    Next sheet
    private_Fail "Table", m_config("registry.sheet") & "/" & m_config("registry.table")
End Function

Private Function private_Header(ByVal index As Long) As String
    Dim headers As Variant

    headers = VBA.Split(m_config("registry.columns"), "|")
    If UBound(headers) <> 19 Then private_Fail "Invalid", "generation.registry.columns"
    private_Header = headers(index)
End Function

Private Function private_Column( _
    ByVal table As ListObject, _
    ByVal header As String _
) As Long
    Dim column As ListColumn

    For Each column In table.ListColumns
        If column.Name = header Then
            private_Column = column.Index
            Exit Function
        End If
    Next column
    private_Fail "Column", header
End Function

Private Function private_Cell( _
    ByVal row As ListRow, _
    ByVal table As ListObject, _
    ByVal index As Long _
) As String
    private_Cell = ex_TableQuery.fn_NormalizeText(VBA.CStr(row.Range.Cells(1, private_Column(table, private_Header(index))).Value2))
End Function

Private Sub private_FindMain( _
    ByVal table As ListObject, _
    ByVal orderNo As String, _
    ByRef matched As ListRow _
)
    Dim row As ListRow

    For Each row In table.ListRows
        If private_Cell(row, table, 2) = m_form("PersonId") And private_Cell(row, table, 6) = orderNo And _
                private_Cell(row, table, 4) <> m_config("donation.event") Then
            If Not matched Is Nothing Then private_Fail "Duplicate", m_form("PersonId") & "/" & orderNo
            Set matched = row
        End If
    Next row
End Sub

Private Sub private_FindDonation( _
    ByVal table As ListObject, _
    ByVal ticketNo As String, _
    ByRef matched As ListRow _
)
    Dim row As ListRow
    Dim orderReference As String

    For Each row In table.ListRows
        If private_Cell(row, table, 2) = m_form("PersonId") And private_Cell(row, table, 15) = ticketNo And _
                private_Cell(row, table, 4) = m_config("donation.event") Then
            orderReference = private_Cell(row, table, 6)
            If VBA.Len(orderReference) = 0 Or _
                    VBA.Left$(orderReference, VBA.Len(m_config("donation.orderPrefix")) + 1) = m_config("donation.orderPrefix") & " " Then
                If Not matched Is Nothing Then private_Fail "Duplicate", ticketNo
                Set matched = row
            End If
        End If
    Next row
End Sub

Private Function private_NextNumber( _
    ByVal table As ListObject, _
    ByVal orderNo As String, _
    ByVal orderDate As Date _
) As String
    Dim row As ListRow
    Dim prefix As String
    Dim number As String
    Dim sequence As String
    Dim maximum As Long

    prefix = VBA.CStr(VBA.Year(orderDate)) & "/" & orderNo & "/"
    For Each row In table.ListRows
        If private_Cell(row, table, 6) = orderNo Then
            number = private_Cell(row, table, 15)
            If VBA.Left$(number, VBA.Len(prefix)) <> prefix Then private_Fail "Invalid", number
            sequence = VBA.Mid$(number, VBA.Len(prefix) + 1)
            If VBA.Len(sequence) = 0 Or sequence Like "*[!0-9]*" Then private_Fail "Invalid", number
            If VBA.CDbl(sequence) >= 2147483647# Then private_Fail "Invalid", number
            If VBA.CLng(sequence) > maximum Then maximum = VBA.CLng(sequence)
        End If
    Next row
    private_NextNumber = prefix & VBA.CStr(maximum + 1)
End Function

Private Function private_AddRow( _
    ByVal table As ListObject, _
    ByVal excluded As ListRow _
) As ListRow
    Dim row As ListRow
    Dim cell As Range
    Dim emptyRow As Boolean

    For Each row In table.ListRows
        If Not row Is excluded Then
            emptyRow = True
            For Each cell In row.Range.Cells
                If Not cell.HasFormula And VBA.Len(VBA.CStr(cell.Value2)) > 0 Then emptyRow = False
            Next cell
            If emptyRow Then
                Set private_AddRow = row
                Exit Function
            End If
        End If
    Next row
    Set private_AddRow = table.ListRows.Add
End Function

Private Function private_ArrivalFormula(ByVal table As ListObject) As String
    Dim row As ListRow
    Dim cell As Range

    For Each row In table.ListRows
        Set cell = row.Range.Cells(1, private_Column(table, private_Header(11)))
        If cell.HasFormula Then
            private_ArrivalFormula = cell.Formula
            Exit Function
        End If
    Next row
    private_ArrivalFormula = m_config("registry.arrivalFormula")
    If VBA.Len(private_ArrivalFormula) = 0 Then private_Fail "Required", private_Header(11)
End Function

Private Function private_Position(ByVal code As String) As String
    Dim normalized As String

    normalized = VBA.UCase$(VBA.Replace$(code, " ", VBA.vbNullString))
    normalized = VBA.Replace$(normalized, VBA.ChrW$(1040), "A")
    normalized = VBA.Replace$(normalized, VBA.ChrW$(1042), "B")
    If VBA.Left$(normalized, VBA.Len(m_config("position.reservedPrefix"))) = m_config("position.reservedPrefix") Then
        private_Position = m_config("position.reservedText")
    ElseIf VBA.Left$(normalized, VBA.Len(m_config("position.removedPrefix"))) = m_config("position.removedPrefix") Then
        private_Position = m_config("position.removedText")
    Else
        private_Position = code
    End If
End Function

Private Sub private_WriteRow( _
    ByVal table As ListObject, _
    ByVal row As ListRow, _
    ByVal eventText As String, _
    ByVal destination As String, _
    ByVal orderNo As String, _
    ByVal foodDate As Date, _
    ByVal departure As Date, _
    ByVal duration As Long, _
    ByVal road As Long, _
    ByVal ticketNo As String, _
    ByVal arrivalFormula As String _
)
    Dim values As Variant
    Dim indexes As Variant
    Dim index As Long
    Dim cell As Range
    Dim status As String

    If m_form("VacationStatus") = m_config("status.cancelled") Then status = m_config("status.mark")
    indexes = VBA.Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 15, 16, 17, 18, 19)
    values = VBA.Array(m_form("PersonRank"), m_form("PersonName"), m_form("PersonId"), _
        private_Position(m_form("PersonPositionCode")), eventText, destination, orderNo, foodDate, departure, _
        duration, road, ticketNo, m_form("TvoName"), m_form("TvoIpn"), private_Position(m_form("TvoPosition")), status)
    For index = 0 To UBound(indexes)
        Set cell = row.Range.Cells(1, private_Column(table, private_Header(VBA.CLng(indexes(index)))))
        If VBA.VarType(values(index)) = VBA.vbString Then
            cell.NumberFormat = "@"
        ElseIf VBA.VarType(values(index)) = VBA.vbDate Then
            cell.NumberFormat = "dd.mm.yyyy"
        End If
        If indexes(index) = 9 Or indexes(index) = 10 Then
            cell.NumberFormat = "0"
            If values(index) = 0 Then
                cell.Value2 = VBA.vbNullString
            Else
                cell.Value2 = values(index)
            End If
        Else
            cell.Value = values(index)
        End If
    Next index
    row.Range.Cells(1, private_Column(table, private_Header(11))).Formula = arrivalFormula
End Sub

Private Sub private_RestoreTable( _
    ByVal table As ListObject, _
    ByVal count As Long, _
    ByVal data As Variant, _
    ByVal formats As Object _
)
    Dim row As ListRow
    Dim index As Long

    Do While table.ListRows.Count > count
        table.ListRows(table.ListRows.Count).Delete
    Loop
    Do While table.ListRows.Count < count
        table.ListRows.Add
    Loop
    If count = 0 Then Exit Sub
    For Each row In table.ListRows
        For index = 1 To table.ListColumns.Count
            row.Range.Cells(1, index).NumberFormat = formats(VBA.CStr(row.Index) & "/" & VBA.CStr(index))
        Next index
    Next row
    table.DataBodyRange.Formula = data
End Sub

Private Function private_Path(ByVal path As String) As String
    If VBA.Left$(path, 1) = "\" And VBA.Left$(path, 2) <> "\\" Then
        private_Fail "Path", path
    End If
    If VBA.Left$(path, 2) <> "\\" And VBA.Mid$(path, 2, 2) <> ":\" Then
        path = m_fileSystem.BuildPath(ThisWorkbook.Path, path)
    End If
    private_Path = m_fileSystem.GetAbsolutePathName(path)
End Function

Private Function private_DocumentPath( _
    ByVal folder As String, _
    ByVal ticketNo As String _
) As String
    Dim name As String
    Dim char As Variant
    Dim key As String

    key = "name.active"
    If m_form("VacationStatus") = m_config("status.cancelled") Then key = "name.cancelled"
    name = VBA.Replace$(m_config(key), "{TicketNo}", ticketNo)
    name = VBA.Replace$(name, "{FIO}", m_form("PersonName"))
    name = VBA.Replace$(name, "{IPN}", m_form("PersonId"))
    For Each char In VBA.Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        name = VBA.Replace$(name, VBA.CStr(char), "_")
    Next char
    private_DocumentPath = m_fileSystem.BuildPath(folder, name & ".docx")
End Function

Private Function private_FindDocument( _
    ByVal folder As String, _
    ByVal ticketNo As String _
) As String
    Dim file As Object
    Dim safeNumber As String

    safeNumber = VBA.Replace$(ticketNo, "/", "_")
    For Each file In m_fileSystem.GetFolder(folder).Files
        If VBA.LCase$(m_fileSystem.GetExtensionName(file.Name)) = "docx" And VBA.Left$(file.Name, 2) <> "~$" Then
            If VBA.InStr(1, file.Name, " " & safeNumber & " ", VBA.vbTextCompare) > 0 And _
                    VBA.InStr(1, file.Name, "(" & m_form("PersonId") & ")", VBA.vbTextCompare) > 0 Then
                If VBA.Len(private_FindDocument) > 0 Then private_Fail "Duplicate", ticketNo
                private_FindDocument = file.Path
            End If
        End If
    Next file
End Function

Private Function private_WordDate(ByVal dateValue As Date) As String
    Dim months As Variant
    Dim result As String

    months = VBA.Split(m_config("months"), "|")
    If UBound(months) <> 11 Then private_Fail "Invalid", "generation.months"
    result = VBA.Replace$(m_config("datePattern"), "{dd}", VBA.Format$(dateValue, "dd"))
    result = VBA.Replace$(result, "{month}", months(VBA.Month(dateValue) - 1))
    private_WordDate = VBA.Replace$(result, "{yyyy}", VBA.Format$(dateValue, "yyyy"))
End Function

Private Function private_CountForm( _
    ByVal count As Long, _
    ByVal category As String _
) As String
    Dim suffix As String

    suffix = "many"
    If count Mod 100 < 11 Or count Mod 100 > 14 Then
        If count Mod 10 = 1 Then suffix = "one"
        If count Mod 10 >= 2 And count Mod 10 <= 4 Then suffix = "few"
    End If
    private_CountForm = VBA.CStr(count) & " " & m_config("duration." & category & "." & suffix)
End Function

Private Sub private_GenerateWord( _
    ByVal templatePath As String, _
    ByVal path As String, _
    ByVal ticketNo As String, _
    ByVal orderDate As Date, _
    ByVal departure As Date, _
    ByVal finalReturn As Date, _
    ByVal vacationDays As Long, _
    ByVal roadDays As Long, _
    ByVal donationDays As Long, _
    ByVal kindWord As String _
)
    Dim word As Object
    Dim document As Object
    Dim range As Object
    Dim names As Variant
    Dim values As Variant
    Dim parts As Variant
    Dim initials As String
    Dim duration As String
    Dim abroad As String
    Dim index As Long
    Dim found As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo EH
    parts = VBA.Split(m_form("PersonName"), " ")
    If UBound(parts) < 2 Then private_Fail "Invalid", private_FieldCaption("PersonName")
    initials = m_form("PersonRank") & " " & parts(0) & VBA.ChrW$(160) & _
        VBA.Left$(parts(1), 1) & "." & VBA.Left$(parts(2), 1) & "."
    duration = private_CountForm(vacationDays, "vacation")
    If donationDays > 0 Then duration = duration & " " & m_config("duration.connector") & " " & _
        private_CountForm(donationDays, "donation") & " " & m_config("duration.donation.suffix")
    If roadDays > 0 Then duration = duration & " " & m_config("duration.connector") & " " & _
        private_CountForm(roadDays, "road") & " " & m_config("duration.road.suffix")
    abroad = VBA.vbCr
    If m_form("VacationAbroad") = m_config("abroad.yes") Then abroad = m_config("abroad.word")
    names = VBA.Array("TicketDate", "TicketNum", "PersonalLine", "VacationKind", "VacationPlace", _
        "VacationDuration", "DateFrom", "DateTo", "PersonalInitials", "DateArrival", "VacationAbroadText", "FooterSpacer")
    values = VBA.Array(private_WordDate(orderDate), ticketNo, m_form("PersonRank") & " " & m_form("PersonName"), _
        kindWord, m_form("DocumentTarget"), duration, private_WordDate(departure), private_WordDate(finalReturn), _
        initials, private_WordDate(VBA.DateAdd("d", 1, finalReturn)), abroad, VBA.vbCr)
    Set word = VBA.CreateObject("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    word.AutomationSecurity = 3
    Set document = word.Documents.Open(templatePath, False, True, False)
    For index = 0 To UBound(names)
        found = False
        Set range = document.Content.Duplicate
        With range.Find
            .ClearFormatting
            .Text = "<" & names(index) & "></" & names(index) & ">"
            .Forward = True
            .Wrap = 0
            .MatchWildcards = False
            .MatchCase = True
            .MatchWholeWord = False
            .Format = False
        End With
        Do While range.Find.Execute
            found = True
            range.Text = VBA.CStr(values(index))
            range.Collapse 0
            range.End = document.Content.End
        Loop
        If Not found Then private_Fail "Marker", VBA.CStr(names(index))
    Next index
    document.SaveAs2 path, 12
    document.Close False
    word.Quit
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    On Error Resume Next
    If Not document Is Nothing Then document.Close False
    If Not word Is Nothing Then word.Quit
    On Error GoTo 0
    VBA.Err.Raise errorNumber, "Word", errorDescription
End Sub