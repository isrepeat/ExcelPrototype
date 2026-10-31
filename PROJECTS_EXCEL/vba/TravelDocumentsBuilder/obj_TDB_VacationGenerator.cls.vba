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
Private m_orderResolver As obj_OrderResolver
Private m_operationId As String
Private m_stage As String
Private m_chain As Collection
Private m_wordSession As Object
Private m_wordAnchors As Object
Private m_wordTemplatePath As String
Private m_wordSourcePath As String

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
    Set m_orderResolver = New obj_OrderResolver
    If Not m_orderResolver.Initialize("TravelDocumentsBuilder::orders", "TravelDocumentsBuilder::text", valueText) Then
        ex_WindowsUi.fn_ShowMessage valueText, VBA.vbExclamation
        Exit Function
    End If
    Set m_config = VBA.CreateObject("Scripting.Dictionary")
    keys = "registry.fieldOrder," & _
        "kind.ids,abroad.yes,abroad.no,abroad.word,status.active,status.cancelled," & _
        "donation.event,donation.orderMissing,donation.orderPrefix,months,datePattern," & _
        "duration.connector,duration.vacation.one,duration.vacation.few,duration.vacation.many," & _
        "duration.donation.one,duration.donation.few,duration.donation.many,duration.donation.suffix," & _
        "duration.road.one,duration.road.few,duration.road.many,duration.road.suffix," & _
        "position.reservedPrefix,position.removedPrefix,position.reservedText,position.removedText," & _
        "message.Required,message.Invalid,message.Column,message.Marker,message.Order,message.Unsupported," & _
        "message.Success,message.Failed,message.Rollback,limits.maxDays"
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
    Set m_wordSession = Nothing
    Set m_wordAnchors = Nothing
    m_wordTemplatePath = VBA.vbNullString
    m_wordSourcePath = VBA.vbNullString
    If Not m_orderResolver Is Nothing Then m_orderResolver.Dispose
    Set m_orderResolver = Nothing
    Set m_bindingContext = Nothing
    Set m_lookupService = Nothing
    Set m_config = Nothing
    Set m_form = Nothing
    Set m_chain = Nothing
    On Error GoTo 0
End Sub

Public Function TrySave( _
    ByVal isUpdate As Boolean, _
    ByRef diagnostic As String, _
    Optional ByVal operationId As String = "" _
) As Boolean
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
    Dim tableChanged As Boolean
    Dim eventsBefore As Boolean
    Dim updatingBefore As Boolean
    Dim guarded As Boolean
    Dim revisions As ListObject
    Dim transaction As Object
    Dim recoveryToken As String
    Dim protection As obj_TDB_HistoryProtection
    Dim protectionDiagnostic As String
    Dim historyRows As Object
    Dim historyRow As Object
    Dim identifier As String
    Dim metadata As Object
    Dim parentId As String
    Dim positionCode As String
    Dim revision As Long
    Dim key As Variant
    Dim index As Long
    Dim rollbackError As String
    Dim failureNumber As Long
    Dim failureSource As String
    Dim failureDescription As String
    Dim filterTables As New Collection
    Dim filterStates As Collection
    Dim filterError As String

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    On Error GoTo EH
    diagnostic = VBA.vbNullString
    If ex_TableSnapshots.fn_Active() Is Nothing Then
        diagnostic = ex_TravelDocumentsBuilder.fn_Config("text.DependenciesRequired")
        Exit Function
    End If
    m_operationId = operationId
    private_LogStage "ReadForm"
    private_ReadForm
    For Each key In m_form.Keys
        private_Log "TDB_FORM_NORMALIZED", "Field=" & key & " | Value=" & ex_TravelDocumentsBuilder.fn_LogValue(m_form(key))
    Next key
    For Each key In m_config.Keys
        private_Log "TDB_GENERATION_CONFIG", "Key=" & key & " | Value=" & ex_TravelDocumentsBuilder.fn_LogValue(m_config(key))
    Next key
    private_LogStage "ValidateForm"
    If m_form("DocumentType") <> private_DocumentCaption() Then private_Fail "Unsupported", m_form("DocumentType")
    For Each key In VBA.Array("PersonName", "PersonId", "PersonRank", "PersonPositionCode", "DocumentTarget", _
            "OrderReference", "ReportNo", "DepartureDate", "VacationKind", "VacationAbroad", "VacationStatus")
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
    private_LogStage "ResolveTvoAndOrder"
    private_ReadChain
    private_ResolveOrder m_form("OrderReference"), orderNo, orderDate, False
    mainReturn = ex_TDB_History.fn_ReturnDate(departureDate, vacationDays, roadDays)
    finalReturn = ex_TDB_History.fn_ReturnDate(departureDate, vacationDays, roadDays, donationDays)
    If donationDays > 0 Then
        private_ResolveOrder VBA.Format$(mainReturn, "dd.mm.yyyy"), donationOrder, orderDate, True
        If VBA.Len(donationOrder) = 0 Then donationOrder = m_config("donation.orderMissing")
        donationOrder = m_config("donation.orderPrefix") & " " & donationOrder
        private_ResolveOrder m_form("OrderReference"), orderNo, orderDate, False
    End If

    ' Validate the workbook, schema and existing records before changing data.
    private_LogStage "ResolveRegistry"
    ex_TDB_History.fn_Open ticketsTable, revisions
    identifier = m_form("HistoryRowId")
    If isUpdate And VBA.Len(identifier) = 0 Then ex_TDB_History.fn_Fail "SelectRequired"
    If Not isUpdate And VBA.Len(identifier) > 0 Then ex_TDB_History.fn_Fail "NewRequired"
    Set transaction = ex_TDB_History.fn_Begin(ticketsTable, revisions, identifier, m_form("HistoryStamp"))
    Set historyRows = ex_TDB_History.fn_Read(ticketsTable)
    For Each key In historyRows.Keys
        Set historyRow = historyRows(key)
        If historyRow("RowType") = "Event" And historyRow("RowId") <> identifier Then
            If VBA.CStr(historyRow(private_Header(2))) = m_form("PersonId") And _
                    VBA.CStr(historyRow(private_Header(6))) = orderNo And ex_TDB_History.fn_Date(historyRow("OrderDate")) = orderDate And _
                    VBA.CStr(historyRow(private_Header(4))) = m_form("VacationKind") And ex_TDB_History.fn_Date(historyRow(private_Header(8))) = departureDate Then
                ex_TDB_History.fn_Fail "Duplicate"
            End If
        End If
    Next key
    If isUpdate Then
        Set mainRow = ex_TDB_History.fn_Find(ticketsTable, identifier)
        ticketNo = private_Cell(mainRow, ticketsTable, 12)
        revision = VBA.CLng(mainRow.Range.Cells(1, ticketsTable.ListColumns("Revision").Index).Value2)
    Else
        identifier = ex_TDB_History.fn_NewId()
        ticketNo = private_NextNumber(ticketsTable, orderNo, orderDate)
        revision = 1
    End If
    If VBA.Len(ticketNo) = 0 Then private_Fail "Invalid", private_Header(12)
    For Each key In historyRows.Keys
        Set historyRow = historyRows(key)
        If historyRow("ParentRowId") = identifier And historyRow("RowType") = "Donation" Then
            Set donationRow = ex_TDB_History.fn_Find(ticketsTable, VBA.CStr(historyRow("RowId")))
        End If
    Next key
    private_Log "TDB_RESOLVED", "Ticket=" & ticketNo & " | Order=" & orderNo & _
        " | OrderDate=" & VBA.Format$(orderDate, "yyyy-mm-dd") & " | MainReturn=" & VBA.Format$(mainReturn, "yyyy-mm-dd") & _
        " | FinalReturn=" & VBA.Format$(finalReturn, "yyyy-mm-dd") & " | DonationOrder=" & ex_TravelDocumentsBuilder.fn_LogValue(donationOrder)
    If Not mainRow Is Nothing Then private_LogRow ticketsTable, mainRow, "BeforeMain"
    If Not donationRow Is Nothing Then private_LogRow ticketsTable, donationRow, "BeforeDonation"

    ' Recheck both tables before the history-only transaction.
    ex_TDB_History.fn_AssertCurrent ticketsTable, revisions, transaction
    Set protection = New obj_TDB_HistoryProtection
    If Not protection.Initialize(ticketsTable, revisions) Then ex_TDB_History.fn_Fail "ReadOnly"
    eventsBefore = Application.EnableEvents
    updatingBefore = Application.ScreenUpdating
    guarded = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    protection.Unprotect
    filterTables.Add ticketsTable
    filterTables.Add revisions
    If Not ex_TableFilters.fn_Suspend(filterTables, filterStates, filterError) Then VBA.Err.Raise 5, , filterError
    tableChanged = True
    private_LogStage "WriteRegistry"
    If mainRow Is Nothing Then Set mainRow = private_AddRow(ticketsTable, donationRow)
    private_WriteRow ticketsTable, mainRow, m_form("VacationKind"), m_form("DocumentTarget"), orderNo, orderDate, _
        departureDate, vacationDays, roadDays, ticketNo, finalReturn
    Set metadata = VBA.CreateObject("Scripting.Dictionary")
    metadata.Add "RowId", identifier
    metadata.Add "ParentRowId", VBA.vbNullString
    metadata.Add "RowType", "Event"
    metadata.Add "Revision", revision
    metadata.Add "PositionCode", m_form("PersonPositionCode")
    metadata.Add "UnitName", m_form("PersonUnit")
    metadata.Add "VacationKind", m_form("VacationKind")
    metadata.Add "VacationAbroad", m_form("VacationAbroad")
    metadata.Add "OrderDate", orderDate
    metadata.Add "ReportNumber", m_form("ReportNo")
    metadata.Add "Basis", m_form("Purpose")
    private_LogStage "WriteEventMetadata"
    ex_TDB_History.fn_Write ticketsTable, mainRow, metadata
    If donationDays > 0 Then
        If donationRow Is Nothing Then Set donationRow = private_AddRow(ticketsTable, mainRow)
        private_WriteRow ticketsTable, donationRow, m_form("VacationKind"), VBA.vbNullString, donationOrder, _
            finalReturn, mainReturn, donationDays, 0, ticketNo, finalReturn
        metadata("RowType") = "Donation"
        metadata("ParentRowId") = identifier
        metadata("RowId") = VBA.CStr(donationRow.Range.Cells(1, ticketsTable.ListColumns("RowId").Index).Value2)
        If VBA.Len(metadata("RowId")) = 0 Then metadata("RowId") = ex_TDB_History.fn_NewId()
        ex_TDB_History.fn_Write ticketsTable, donationRow, metadata
    ElseIf Not donationRow Is Nothing Then
        donationRow.Delete
    End If

    ' Replace only this event's TVO children, retaining stable IDs from the loaded chain.
    For index = ticketsTable.ListRows.Count To 1 Step -1
        Set row = ticketsTable.ListRows(index)
        key = VBA.CStr(row.Range.Cells(1, ticketsTable.ListColumns("RowId").Index).Value2)
        If historyRows.Exists(key) Then
            Set historyRow = historyRows(key)
            If historyRow("RowType") = "TVO" Then
                Set metadata = historyRow
                Do While metadata("RowType") <> "Event"
                    Set metadata = historyRows(metadata("ParentRowId"))
                Loop
                If metadata("RowId") = identifier Then row.Delete
            End If
        End If
    Next index
    parentId = identifier
    positionCode = m_form("PersonPositionCode")
    For index = 1 To m_chain.Count
        Set metadata = m_chain(index)
        If VBA.Len(metadata("RowId")) = 0 Then metadata("RowId") = ex_TDB_History.fn_NewId()
        metadata("ParentRowId") = parentId
        metadata("RowType") = "TVO"
        metadata("DutyPositionCode") = positionCode
        metadata("Revision") = revision
        metadata("Status") = "Active"
        If m_form("VacationStatus") = m_config("status.cancelled") Then metadata("Status") = "Cancelled"
        Set row = ticketsTable.ListRows.Add
        ex_TDB_History.fn_Write ticketsTable, row, metadata
        parentId = metadata("RowId")
        positionCode = metadata("PositionCode")
    Next index
    private_LogStage "CommitHistoryRevision"
    ex_TDB_History.fn_Commit ticketsTable, revisions, identifier, transaction, m_form("ChangeReason")
    If revisions.ListRows.Count > VBA.CLng(transaction("RevisionsCount")) Then
        private_LogStage "WriteHistoryRecovery"
        recoveryToken = ex_TDB_HistoryRecovery.fn_Prepare(ticketsTable, revisions, identifier, transaction)
    End If
    filterError = ex_TableFilters.fn_Restore(filterStates)
    If Not protection.TryProtect(protectionDiagnostic) Then VBA.Err.Raise 5, , protectionDiagnostic
    If VBA.Len(recoveryToken) > 0 Then ex_TDB_HistoryRecovery.fn_Mark ticketsTable.Parent.Parent, recoveryToken, "COMMIT"

    tableChanged = False
    protection.Dispose
    Application.EnableEvents = eventsBefore
    Application.ScreenUpdating = updatingBefore
    guarded = False
    m_bindingContext.SetValue "Form", "HistoryRowId", identifier
    m_bindingContext.SetValue "Form", "ReturnDate", VBA.Format$(finalReturn, "dd.mm.yyyy")
    m_bindingContext.SetValue "Form", "TicketNo", ticketNo
    m_bindingContext.SetValue "Form", "GenerationStatusStyle", "generationSuccess"
    m_bindingContext.SetValue "Form", "GenerationStatus", private_Message("Success", ticketNo)
    If Len(filterError) > 0 Then
        m_bindingContext.SetValue "Form", "GenerationStatus", private_Message("Success", ticketNo) & vbLf & filterError
    End If
    private_Log "TDB_HISTORY_SAVED", "Ticket=" & ticketNo & " | RowId=" & identifier
    TrySave = True
    Exit Function
EH:
    failureNumber = VBA.Err.Number
    failureSource = VBA.Err.Source
    failureDescription = VBA.Err.Description
    private_Log "TDB_GENERATION_ERROR", "Stage=" & m_stage & " | Number=" & VBA.CStr(failureNumber) & _
        " | Source=" & ex_TravelDocumentsBuilder.fn_LogValue(failureSource) & " | Description=" & ex_TravelDocumentsBuilder.fn_LogValue(failureDescription) & _
        " | TableChanged=" & VBA.CStr(tableChanged)
    diagnostic = private_Message("Failed", failureDescription)
    ex_Core.fn_Diagnostic_WriteLog "TDB_GENERATION_FAILED | " & diagnostic
    On Error Resume Next
    If tableChanged Then
        VBA.Err.Clear
        protection.Unprotect
        If VBA.Err.Number <> 0 Then rollbackError = rollbackError & VBA.Err.Description & VBA.vbCrLf
        VBA.Err.Clear
        ex_TableFilters.fn_ClearCaptured filterStates
        If VBA.Err.Number <> 0 Then rollbackError = rollbackError & VBA.Err.Description & VBA.vbCrLf
        VBA.Err.Clear
        ex_TDB_History.fn_Rollback ticketsTable, revisions, transaction
        If VBA.Err.Number <> 0 Then rollbackError = rollbackError & VBA.Err.Description & VBA.vbCrLf
    End If
    filterError = ex_TableFilters.fn_Restore(filterStates)
    If Len(filterError) > 0 Then rollbackError = rollbackError & filterError & vbCrLf
    If Not protection Is Nothing Then
        If Not protection.TryProtect(protectionDiagnostic) Then rollbackError = rollbackError & protectionDiagnostic & VBA.vbCrLf
        protection.Dispose
    End If
    If guarded Then
        Application.EnableEvents = eventsBefore
        Application.ScreenUpdating = updatingBefore
    End If
    If VBA.Len(rollbackError) > 0 Then diagnostic = diagnostic & VBA.vbCrLf & private_Message("Rollback", rollbackError)
    On Error GoTo 0
    private_Log "TDB_ROLLBACK_FINISHED", "Errors=" & ex_TravelDocumentsBuilder.fn_LogValue(rollbackError)
End Function

Public Function TryRender( _
    ByVal root As Object, _
    ByVal group As Collection, _
    ByVal templatePath As String, _
    ByVal outputPath As String, _
    ByVal signatory As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim mapping As Variant
    Dim pair As Variant
    Dim row As Object
    Dim vacationDays As Long
    Dim roadDays As Long
    Dim donationDays As Long
    Dim departure As Date
    Dim finalReturn As Date
    Dim kindWord As String
    Dim kindEvent As String

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    On Error GoTo Failed
    m_operationId = "WORD/" & root("RowId") & "/R" & root("Revision")
    Set m_form = VBA.CreateObject("Scripting.Dictionary")
    For Each mapping In VBA.Split(ex_TDB_History.fn_Config("formFields"), "|")
        pair = VBA.Split(mapping, "=")
        m_form(pair(0)) = VBA.CStr(root(pair(1)))
    Next mapping
    m_form("DonationDays") = "0"
    For Each row In group
        If row("RowType") = "Donation" Then m_form("DonationDays") = VBA.CStr(row(ex_TDB_History.fn_Config("durationField")))
    Next row
    vacationDays = private_Days("VacationDays", False)
    roadDays = private_Days("RoadDays", True)
    donationDays = private_Days("DonationDays", True)
    departure = ex_TDB_History.fn_Date(root(ex_TDB_History.fn_Config("departureField")))
    finalReturn = ex_TDB_History.fn_Date(root("PlannedReturnDate"))
    If finalReturn <> ex_TDB_History.fn_ReturnDate(departure, vacationDays, roadDays, donationDays) Then ex_TDB_History.fn_Fail "InvalidRows"
    private_MapKind kindWord, kindEvent
    If m_form("VacationAbroad") <> m_config("abroad.yes") And m_form("VacationAbroad") <> m_config("abroad.no") Then
        private_Fail "Invalid", private_FieldCaption("VacationAbroad")
    End If
    private_GenerateWord templatePath, outputPath, VBA.CStr(root(ex_TDB_History.fn_Config("ticketField"))), _
        ex_TDB_History.fn_Date(root("OrderDate")), departure, finalReturn, vacationDays, roadDays, donationDays, kindWord, signatory
    TryRender = True
    Exit Function
Failed:
    diagnostic = VBA.Err.Description
End Function

' //
' // Private
' //
Private Sub private_Log( _
    ByVal eventName As String, _
    ByVal details As String _
)
    ex_Core.fn_Diagnostic_WriteLog eventName & " | Operation=" & m_operationId & " | " & details
End Sub

Private Sub private_LogStage(ByVal stageName As String)
    m_stage = stageName
    private_Log "TDB_STAGE", "Name=" & stageName
End Sub

Private Sub private_LogRow( _
    ByVal table As ListObject, _
    ByVal row As ListRow, _
    ByVal action As String _
)
    Dim columnIndex As Long
    Dim cell As Range

    For columnIndex = 1 To table.ListColumns.Count
        Set cell = row.Range.Cells(1, columnIndex)
        private_Log "TDB_REGISTRY_CELL", "Action=" & action & " | Row=" & VBA.CStr(row.Index) & _
            " | Column=" & ex_TravelDocumentsBuilder.fn_LogValue(table.ListColumns(columnIndex).Name) & _
            " | Value=" & ex_TravelDocumentsBuilder.fn_LogValue(cell.Value2) & _
            " | Formula=" & ex_TravelDocumentsBuilder.fn_LogValue(cell.Formula) & _
            " | Format=" & ex_TravelDocumentsBuilder.fn_LogValue(cell.NumberFormat)
    Next columnIndex
End Sub

Private Sub private_ReadForm()
    Dim key As Variant
    Dim value As Variant
    Dim objectValue As Object
    Dim isObject As Boolean

    Set m_form = VBA.CreateObject("Scripting.Dictionary")
    For Each key In VBA.Array("HistoryRowId", "HistoryStamp", "ChangeReason", "PersonUnit", "DocumentType", "PersonName", "PersonId", "PersonRank", "PersonPosition", "PersonPositionCode", _
            "DocumentTarget", "DepartureDate", "OrderReference", "VacationKind", "VacationDays", "RoadDays", _
            "DonationDays", "VacationAbroad", "VacationStatus", "ReportNo", "Purpose", "TvoName", "TvoIpn", "TvoPosition", "TemplatePath", "OutputFolderPath")
        If Not m_bindingContext.TryGetValue("Form", VBA.CStr(key), value, objectValue, isObject) Then
            private_Fail "Required", private_FieldCaption(VBA.CStr(key))
        End If
        If isObject Or VBA.IsError(value) Or VBA.IsNull(value) Then private_Fail "Invalid", private_FieldCaption(VBA.CStr(key))
        If VBA.CStr(key) = "HistoryStamp" Then
            m_form.Add VBA.CStr(key), VBA.CStr(value)
        ElseIf VBA.CStr(key) = "Purpose" Then
            m_form.Add VBA.CStr(key), VBA.Trim$(VBA.CStr(value))
        Else
            m_form.Add VBA.CStr(key), ex_TableQuery.fn_NormalizeText(VBA.CStr(value))
        End If
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
        Case "DocumentTarget", "DepartureDate", "Purpose"
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

Private Function private_Date(ByVal text As String) As Date
    private_Date = ex_TDB_History.fn_Date(text)
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

Private Sub private_ReadChain()
    Dim source As Object
    Dim value As Variant
    Dim isObject As Boolean
    Dim row As Object
    Dim resolved As Object
    Dim people As Object
    Dim positions As Object
    Dim preserved As Object
    Dim field As Variant
    Dim mapping As Variant
    Dim pair As Variant

    Set m_chain = New Collection
    If Not m_bindingContext.TryGetValue("Data", "TvoChain", value, source, isObject) Then ex_TDB_History.fn_Fail "InvalidChain"
    If Not isObject Then ex_TDB_History.fn_Fail "InvalidChain"
    Set people = VBA.CreateObject("Scripting.Dictionary")
    Set positions = VBA.CreateObject("Scripting.Dictionary")
    people.CompareMode = VBA.vbTextCompare
    positions.CompareMode = VBA.vbTextCompare
    people.Add m_form("PersonId"), True
    positions.Add m_form("PersonPositionCode"), True
    For Each row In source
        Set resolved = ex_TDB_History.fn_Actor(m_lookupService, VBA.CStr(row(ex_TDB_History.fn_Config("personIdField"))), _
            VBA.CStr(row(ex_TDB_History.fn_Config("personNameField"))), VBA.CStr(row("PositionCode")))
        Set preserved = ex_TDB_History.fn_NewRow()
        For Each field In row.Keys
            preserved(field) = row(field)
        Next field
        For Each mapping In VBA.Split(ex_TDB_History.fn_Config("actorFields"), "|")
            pair = VBA.Split(mapping, "=")
            preserved(pair(1)) = resolved(pair(1))
        Next mapping
        Set resolved = preserved
        If people.Exists(resolved(ex_TDB_History.fn_Config("personIdField"))) Or positions.Exists(resolved("PositionCode")) Then
            ex_TDB_History.fn_Fail "InvalidChain"
        End If
        people.Add resolved(ex_TDB_History.fn_Config("personIdField")), True
        positions.Add resolved("PositionCode"), True
        m_chain.Add resolved
    Next row
    If VBA.Len(m_form("TvoName") & m_form("TvoIpn") & m_form("TvoPosition")) > 0 Then
        If m_chain.Count > 0 Then ex_TDB_History.fn_Fail "PendingTvo"
        Set resolved = ex_TDB_History.fn_Actor(m_lookupService, m_form("TvoIpn"), m_form("TvoName"), m_form("TvoPosition"))
        If people.Exists(resolved(ex_TDB_History.fn_Config("personIdField"))) Or positions.Exists(resolved("PositionCode")) Then
            ex_TDB_History.fn_Fail "InvalidChain"
        End If
        m_chain.Add resolved
    End If
End Sub

Private Sub private_ResolveOrder( _
    ByVal reference As String, _
    ByRef orderNo As String, _
    ByRef orderDate As Date, _
    ByVal allowMissing As Boolean _
)
    Dim matches As Collection
    Dim record As obj_LookupRecord
    Dim diagnostic As String
    Dim value As Variant

    orderNo = VBA.vbNullString
    reference = VBA.Trim$(VBA.Replace$(reference, "№", VBA.vbNullString))
    If VBA.Len(reference) = 0 Then private_Fail "Order", reference
    If Not m_orderResolver.TrySearch(reference, VBA.vbNullString, matches, diagnostic) Then
        private_Fail "Order", diagnostic
    End If
    If matches.Count > 1 Or (matches.Count = 0 And Not allowMissing) Then private_Fail "Order", reference
    If matches.Count = 0 Then Exit Sub
    Set record = matches(1)
    If Not record.TryGetValue("OrderNumber", value) Then private_Fail "Order", reference
    orderNo = VBA.CStr(value)
    If Not record.TryGetValue("OrderDate", value) Then private_Fail "Order", reference
    orderDate = VBA.CDate(value)
End Sub

Private Function private_Header(ByVal index As Long) As String
    Dim headers As Variant

    headers = VBA.Split(m_config("registry.fieldOrder"), "|")
    If UBound(headers) <> 18 Then private_Fail "Invalid", "generation.registry.fieldOrder"
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
        If VBA.CStr(row.Range.Cells(1, table.ListColumns("RowType").Index).Value2) = "Event" Then
            If VBA.Left$(private_Cell(row, table, 12), VBA.Len(prefix)) = prefix Then
                number = private_Cell(row, table, 12)
                sequence = VBA.Mid$(number, VBA.Len(prefix) + 1)
                If VBA.Len(sequence) = 0 Or sequence Like "*[!0-9]*" Then private_Fail "Invalid", number
                If VBA.CDbl(sequence) >= 2147483647# Then private_Fail "Invalid", number
                If VBA.CLng(sequence) > maximum Then maximum = VBA.CLng(sequence)
            End If
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
    ByVal plannedReturn As Date _
)
    Dim values As Variant
    Dim indexes As Variant
    Dim index As Long
    Dim status As String
    Dim fields As Object

    status = "Active"
    If m_form("VacationStatus") = m_config("status.cancelled") Then status = "Cancelled"
    indexes = VBA.Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 16, 17, 11)
    values = VBA.Array(m_form("PersonRank"), m_form("PersonName"), m_form("PersonId"), _
        m_form("PersonPosition"), eventText, destination, orderNo, foodDate, departure, _
        duration, road, ticketNo, status, m_form("Purpose"), plannedReturn)
    Set fields = VBA.CreateObject("Scripting.Dictionary")
    For index = 0 To UBound(indexes)
        fields.Add private_Header(VBA.CLng(indexes(index))), values(index)
        private_Log "TDB_REGISTRY_WRITE", "Row=" & VBA.CStr(row.Index) & _
            " | Column=" & ex_TravelDocumentsBuilder.fn_LogValue(private_Header(VBA.CLng(indexes(index)))) & _
            " | IntendedValue=" & ex_TravelDocumentsBuilder.fn_LogValue(values(index))
    Next index
    ex_TDB_History.fn_Write table, row, fields
    row.Range.Cells(1, private_Column(table, private_Header(17))).WrapText = True
End Sub

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
    ByVal kindWord As String, _
    ByVal signatory As Object _
)
    Dim word As Object
    Dim document As Object
    Dim names As Variant
    Dim values As Variant
    Dim parts As Variant
    Dim initials As String
    Dim duration As String
    Dim abroad As String
    Dim index As Long
    Dim bookmarkName As Variant
    Dim target As Object
    Dim documentStartedAt As Double
    Dim stageStartedAt As Double
    Dim stage As String
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo EH
    documentStartedAt = VBA.Timer
    stage = "Document.Prepare"
    stageStartedAt = documentStartedAt
    parts = VBA.Split(m_form("PersonName"), " ")
    If UBound(parts) < 2 Then private_Fail "Invalid", private_FieldCaption("PersonName")
    initials = m_form("PersonRank") & " " & parts(0) & VBA.ChrW$(160) & _
        VBA.Left$(parts(1), 1) & "." & VBA.Left$(parts(2), 1) & "."
    duration = private_CountForm(vacationDays, "vacation")
    If donationDays > 0 Then duration = duration & " " & m_config("duration.connector") & " " & _
        private_CountForm(donationDays, "donation") & " " & m_config("duration.donation.suffix")
    If roadDays > 0 Then duration = duration & " " & m_config("duration.connector") & " " & _
        private_CountForm(roadDays, "road") & " " & m_config("duration.road.suffix")
    abroad = VBA.vbNullString
    If m_form("VacationAbroad") = m_config("abroad.yes") Then abroad = m_config("abroad.word")
    names = VBA.Array("TicketDate", "TicketNum", "PersonalLine", "VacationKind", "VacationPlace", _
        "VacationDuration", "DateFrom", "DateTo", "PersonalInitials", "DateArrival", "VacationAbroadText", "FooterSpacer", _
        "SignatoryPosition", "SignatoryRank", "SignatoryName")
    values = VBA.Array(private_WordDate(orderDate), ticketNo, m_form("PersonRank") & " " & m_form("PersonName"), _
        kindWord, m_form("DocumentTarget"), duration, private_WordDate(departure), private_WordDate(finalReturn), _
        initials, private_WordDate(VBA.DateAdd("d", 1, finalReturn)), abroad, VBA.vbNullString, _
        signatory("SignatoryPosition"), signatory("SignatoryRank"), signatory("SignatoryName"))
    For index = 0 To UBound(names)
        private_Log "TDB_DOCX_VALUE", "Marker=" & names(index) & " | Value=" & ex_TravelDocumentsBuilder.fn_LogValue(values(index))
    Next index
    private_WordPerf stage, stageStartedAt
    stage = "Session.Start"
    stageStartedAt = VBA.Timer
    If m_wordSession Is Nothing Then
        Set m_wordSession = ex_TDB_WordSession.fn_Acquire()
    End If
    Set word = m_wordSession
    word.ScreenUpdating = False
    private_WordPerf stage, stageStartedAt
    stage = "Document.Create"
    stageStartedAt = VBA.Timer
    If m_wordAnchors Is Nothing Then
        Set document = private_OpenWordDocument(word, templatePath)
        private_WordPerf stage, stageStartedAt
        stage = "Template.IndexMarkers"
        stageStartedAt = VBA.Timer
        private_IndexWordMarkers document, names
        private_WordPerf stage, stageStartedAt
        stage = "Template.SaveAnchors"
        stageStartedAt = VBA.Timer
        m_wordTemplatePath = VBA.Left$(templatePath, VBA.InStrRev(templatePath, "\")) & "a.docx"
        document.SaveAs2 m_wordTemplatePath, 12
        m_wordSourcePath = templatePath
        private_WordPerf stage, stageStartedAt
    Else
        If VBA.StrComp(templatePath, m_wordSourcePath, VBA.vbTextCompare) <> 0 Then private_Fail "Invalid", "TemplatePath"
        Set document = private_OpenWordDocument(word, m_wordTemplatePath)
        private_WordPerf stage, stageStartedAt
    End If
    stage = "Document.ReplaceMarkers"
    stageStartedAt = VBA.Timer
    ' Bookmarks retain their positions when earlier replacements change the text length.
    For index = 0 To UBound(names)
        For Each bookmarkName In m_wordAnchors(VBA.CStr(names(index)))
            Set target = document.Bookmarks(VBA.CStr(bookmarkName)).Range
            target.Text = VBA.CStr(values(index))
        Next bookmarkName
    Next index
    private_WordPerf stage, stageStartedAt
    stage = "Document.Save"
    stageStartedAt = VBA.Timer
    private_Log "TDB_DOCX_SAVE", "PathLength=" & VBA.CStr(VBA.Len(path)) & " | Path=" & path
    document.SaveAs2 path, 12
    private_WordPerf stage, stageStartedAt
    stage = "Document.Close"
    stageStartedAt = VBA.Timer
    document.Close False
    Set document = Nothing
    private_WordPerf stage, stageStartedAt
    private_WordPerf "Document.Total", documentStartedAt
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    On Error Resume Next
    private_Log "TDB_WORD_FAILED", "Stage=" & stage & " | Path=" & path & " | Number=" & VBA.CStr(errorNumber) & " | Error=" & errorDescription
    private_WordPerf stage & ".Failed", stageStartedAt
    If Not document Is Nothing Then document.Close False
    Set m_wordSession = Nothing
    Set m_wordAnchors = Nothing
    m_wordTemplatePath = VBA.vbNullString
    m_wordSourcePath = VBA.vbNullString
    private_WordPerf "Document.FailedTotal", documentStartedAt
    On Error GoTo 0
    VBA.Err.Raise errorNumber, "Word", errorDescription
End Sub

Private Sub private_WordPerf( _
    ByVal stage As String, _
    ByVal startedAt As Double _
)
    ex_Core.fn_Diagnostic_WritePerf "TDB.Word." & stage & " | Operation=" & m_operationId, startedAt
End Sub

Private Function private_OpenWordDocument( _
    ByVal word As Object, _
    ByVal path As String _
) As Object
    Set private_OpenWordDocument = word.Documents.Open( _
        FileName:=path, _
        ConfirmConversions:=False, _
        ReadOnly:=True, _
        AddToRecentFiles:=False, _
        Visible:=False _
    )
End Function

Private Sub private_IndexWordMarkers( _
    ByVal document As Object, _
    ByVal names As Variant _
)
    Dim story As Object
    Dim storyRange As Object
    Dim searchRange As Object
    Dim occurrences As Collection
    Dim existing As Object
    Dim prefix As String
    Dim bookmarkName As String
    Dim marker As String
    Dim index As Long
    Dim duplicate As Boolean
    Dim searchStart As Long
    Dim storyEnd As Long
    Dim matchStart As Long
    Dim matchEnd As Long

    Set m_wordAnchors = VBA.CreateObject("Scripting.Dictionary")
    prefix = "TDB_" & VBA.Mid$(ex_TDB_History.fn_NewId(), 3, 8) & "_"
    For index = 0 To UBound(names)
        Set occurrences = New Collection
        m_wordAnchors.Add VBA.CStr(names(index)), occurrences
    Next index
    ' Index the frozen template once, including first-page and linked section stories.
    For Each story In document.StoryRanges
        Set storyRange = story
        Do
            For index = 0 To UBound(names)
                marker = "{" & names(index) & "}"
                Set occurrences = m_wordAnchors(VBA.CStr(names(index)))
                searchStart = storyRange.Start
                storyEnd = storyRange.End
                Do While searchStart < storyEnd
                    Set searchRange = storyRange.Duplicate
                    searchRange.SetRange searchStart, storyEnd
                    searchRange.Find.ClearFormatting
                    If Not searchRange.Find.Execute(marker, True, False, False, False, False, True, 0, False) Then Exit Do
                    ' Capture numeric bounds before Bookmarks.Add can affect a COM Range argument.
                    matchStart = searchRange.Start
                    matchEnd = searchRange.End
                    If matchStart < searchStart Or matchEnd <= searchStart Or matchEnd > storyEnd Then
                        private_Log "TDB_WORD_INDEX_BOUNDARY", "Marker=" & marker & " | Start=" & VBA.CStr(searchStart) & _
                            " | Match=" & VBA.CStr(matchStart) & "/" & VBA.CStr(matchEnd)
                        private_Fail "Marker", VBA.CStr(names(index))
                    End If
                    duplicate = False
                    For Each existing In searchRange.Bookmarks
                        If VBA.Left$(existing.Name, VBA.Len(prefix)) = prefix Then duplicate = True
                    Next existing
                    If Not duplicate Then
                        bookmarkName = prefix & VBA.CStr(index) & "_" & VBA.CStr(occurrences.Count + 1)
                        ' Consume the marker before anchoring: Word can repeat an unchanged Find match.
                        ' Equal-length padding preserves offsets and is replaced through the bookmark later.
                        searchRange.Text = VBA.Space$(matchEnd - matchStart)
                        document.Bookmarks.Add bookmarkName, searchRange.Duplicate
                        occurrences.Add bookmarkName
                    End If
                    searchStart = matchEnd
                Loop
            Next index
            Set storyRange = storyRange.NextStoryRange
        Loop Until storyRange Is Nothing
    Next story
    For index = 0 To UBound(names)
        Set occurrences = m_wordAnchors(VBA.CStr(names(index)))
        If occurrences.Count = 0 And names(index) <> "FooterSpacer" Then private_Fail "Marker", VBA.CStr(names(index))
    Next index
End Sub