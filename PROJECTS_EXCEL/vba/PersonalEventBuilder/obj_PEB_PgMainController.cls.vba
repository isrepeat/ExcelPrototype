VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PEB_PgMainController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_pageBase As obj_PageBase
Private m_lookupProfile As obj_LookupProfile
Private m_lookupService As obj_LookupService
Private m_lookupResult As obj_LookupResult
Private m_messages As Object
Private m_candidateEnricher As obj_PEB_CandidateEnricher
Private m_orderResolver As obj_OrderResolver
Private m_ordersByCaption As Object
Private m_pendingOrder As obj_LookupRecord
Private m_acceptedOrderNumber As String
Private m_acceptedOrderDate As Variant
Private m_modeFields As Collection
Private m_modeId As String
Private m_modeDirection As String
Private m_eventType As String
Private m_modeConfig As Object

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
Public Function Initialize(ByVal pageBase As obj_PageBase) As Boolean
    Dim messageKey As Variant
    Dim messageText As String
    Dim diagnostic As String
    Dim defaultMode As String

    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    Set m_pageBase = pageBase
    If m_pageBase Is Nothing Then
        ex_WindowsUi.fn_ShowMessage "The PersonalEventBuilder page controller requires a page base.", _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    Set m_messages = VBA.CreateObject("Scripting.Dictionary")
    For Each messageKey In VBA.Array("PersonnelError", "PersonnelMinimum", "PersonnelEmpty", _
            "PersonnelCount", "PersonnelMore", "PersonnelSelected", "PersonnelActionError", _
            "OrderReady", "OrderEmpty", "OrderFoundOne", "OrderFoundMany", "OrderChooseRequired", _
            "OrderChoiceFormat", "OrderAccepted", "OrderAppliedRefreshError")
        If Not ex_Core.fn_TryGetWorkbookConfigValue("PersonalEventBuilder::text." & messageKey, messageText) Then
            ex_WindowsUi.fn_ShowMessage "Required configuration text not found: PersonalEventBuilder::text." & messageKey, vbExclamation, "Configuration"
            Exit Function
        End If
        m_messages(messageKey) = messageText
    Next messageKey
    Set m_lookupProfile = New obj_LookupProfile
    Set m_ordersByCaption = VBA.CreateObject("Scripting.Dictionary")
    Set m_orderResolver = New obj_OrderResolver
    If Not m_orderResolver.Initialize("PersonalEventBuilder::orders", "PersonalEventBuilder::text", diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Configuration"
        Exit Function
    End If
    If Not m_lookupProfile.Initialize("PersonalEventBuilder::lookup.Personnel", diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Personnel lookup"
        Exit Function
    End If
    If VBA.Len(m_lookupProfile.Source.RangeAddress) = 0 Or VBA.LCase$(m_lookupProfile.Source.RangeAddress) = "auto" Then
        ex_WindowsUi.fn_ShowMessage "Explicit configuration range required: PersonalEventBuilder::lookup.Personnel.source.range", vbExclamation, "Configuration"
        Exit Function
    End If
    Set m_lookupService = New obj_LookupService
    If Not m_lookupService.Initialize(m_lookupProfile) Then Exit Function
    Set m_candidateEnricher = New obj_PEB_CandidateEnricher
    If Not m_candidateEnricher.Initialize(diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Personnel lookup"
        Exit Function
    End If
    If Not m_lookupService.TrySearch(vbNullString, m_lookupResult, diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Personnel lookup"
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", m_lookupResult) Then Exit Function
    If Not private_CaptureModeConfig(diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Configuration"
        Exit Function
    End If
    If Not private_ReadModeConfig("modes.default", defaultMode, diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Configuration"
        Exit Function
    End If
    If Not private_LoadMode(defaultMode, diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Configuration"
        Exit Function
    End If
    Initialize = True
    m_isInitialized = Initialize
    If Not Me.SearchOrdersHandler() Then
        Me.Dispose
        Initialize = False
    End If
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    If Not m_lookupService Is Nothing Then m_lookupService.Dispose
    If Not m_lookupResult Is Nothing Then m_lookupResult.Dispose
    If Not m_lookupProfile Is Nothing Then m_lookupProfile.Dispose
    If Not m_candidateEnricher Is Nothing Then m_candidateEnricher.Dispose
    If Not m_orderResolver Is Nothing Then m_orderResolver.Dispose
    Set m_orderResolver = Nothing
    Set m_ordersByCaption = Nothing
    Set m_pendingOrder = Nothing
    m_acceptedOrderNumber = vbNullString
    m_acceptedOrderDate = Empty
    Set m_candidateEnricher = Nothing
    Set m_messages = Nothing
    Set m_lookupService = Nothing
    Set m_lookupProfile = Nothing
    Set m_lookupResult = Nothing
    Set m_modeFields = Nothing
    Set m_modeConfig = Nothing
    Set m_pageBase = Nothing
End Sub

Public Function SearchPersonnelHandler() As Boolean
    Dim text As Variant
    Dim diagnostic As String
    Dim result As obj_LookupResult
    Dim updates As Object
    Dim name As Variant
    Dim status As String
    Dim startedAt As Double

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not private_TryReadFormValue("PersonName", text) Then Exit Function
    Set updates = VBA.CreateObject("Scripting.Dictionary")
    For Each name In VBA.Array("AbsenceType", "Destination", "PersonId", "PersonRank", _
            "PositionCode", "PersonPosition", "DepartureDate", "TicketNumber", "TicketDate")
        updates("Form." & name) = vbNullString
    Next name
    updates("Form.EventDate") = vbNullString
    If Not m_pageBase.BindingContext.SetValue("Data", "SelectedPersonnel", vbNullString) Then Exit Function
    If Not m_pageBase.BindingContext.TryApplyValues(updates, diagnostic) Then GoTo Failure
    If Not m_lookupService.TrySearch(VBA.CStr(text), result, diagnostic) Then
        If Not m_lookupService.TrySearch(vbNullString, result, status) Then GoTo Failure
        If Not private_TryEnrichResult(result, status) Then GoTo Failure
        Set m_lookupResult = result
        m_pageBase.BindingContext.SetObject "Data", "PersonnelCandidates", result
        m_pageBase.BindingContext.SetValue "Text", "PersonnelStatus", m_messages("PersonnelError")
        GoTo Failure
    End If
    If Not private_TryEnrichResult(result, diagnostic) Then
        If m_lookupService.TrySearch(vbNullString, result, status) Then
            If Not private_TryEnrichResult(result, status) Then GoTo Failure
            Set m_lookupResult = result
            m_pageBase.BindingContext.SetObject "Data", "PersonnelCandidates", result
        End If
        GoTo Failure
    End If
    Set m_lookupResult = result
    ex_Core.fn_Diagnostic_WriteLog "LOOKUP_STAGE_STARTED | Name=PublishCandidates"
    startedAt = VBA.Timer
    If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", result) Then Exit Function
    ex_Core.fn_Diagnostic_WritePerf "Lookup.PublishCandidates | Rows=" & VBA.CStr(result.RowCount), startedAt
    If VBA.Len(ex_TableQuery.fn_NormalizeText(VBA.CStr(text))) < m_lookupProfile.MinChars Then
        status = VBA.Replace$(m_messages("PersonnelMinimum"), "{minChars}", VBA.CStr(m_lookupProfile.MinChars))
    ElseIf result.RowCount = 0 Then
        status = m_messages("PersonnelEmpty")
    Else
        status = VBA.Replace$(m_messages("PersonnelCount"), "{count}", VBA.CStr(result.RowCount))
        If result.HasMore Then status = VBA.Replace$(m_messages("PersonnelMore"), "{count}", VBA.CStr(result.RowCount))
    End If
    SearchPersonnelHandler = m_pageBase.BindingContext.SetValue("Text", "PersonnelStatus", status)
    If SearchPersonnelHandler Then SearchPersonnelHandler = private_SetActionStatus(status, False)
    Exit Function
Failure:
    private_SetActionStatus VBA.Replace$(m_messages("PersonnelActionError"), "{details}", diagnostic), True
    ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Personnel lookup"
End Function

Public Function SelectPersonnelHandler(ByVal selected As Object) As Boolean
    Dim record As obj_LookupRecord
    Dim diagnostic As String
    Dim previousScreenUpdating As Boolean
    Dim startedAt As Double

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If selected Is Nothing Then Exit Function
    If Not TypeOf selected Is obj_LookupRecord Then Exit Function
    previousScreenUpdating = Application.ScreenUpdating
    startedAt = VBA.Timer
    On Error GoTo EH
    Application.ScreenUpdating = False
    Set record = selected
    If Not private_ApplyModeRecord(record, diagnostic) Then
        private_SetActionStatus VBA.Replace$(m_messages("PersonnelActionError"), "{details}", diagnostic), True
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Personnel lookup"
        GoTo CleanExit
    End If
    SelectPersonnelHandler = m_pageBase.BindingContext.SetValue("Text", "PersonnelStatus", m_messages("PersonnelSelected"))
    If SelectPersonnelHandler Then SelectPersonnelHandler = private_SetActionStatus(m_messages("PersonnelSelected"), False)
CleanExit:
    Application.ScreenUpdating = previousScreenUpdating
    ex_Core.fn_Diagnostic_WritePerf "Personnel.ApplySelection", startedAt
    Exit Function
EH:
    private_SetActionStatus VBA.Replace$(m_messages("PersonnelActionError"), "{details}", VBA.Err.Description), True
    Resume CleanExit
End Function

Public Function SearchOrdersHandler() As Boolean
    Dim query As Variant
    Dim year As Variant
    Dim matches As Collection
    Dim displayMatches As Collection
    Dim options As Collection
    Dim todayOptions As Collection
    Dim record As obj_LookupRecord
    Dim caption As String
    Dim diagnostic As String
    Dim status As String
    Dim orderDate As Variant
    Dim isError As Boolean

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    On Error GoTo EH
    If Not private_TryReadFormValue("OrderQuery", query) Then Exit Function
    If Not private_TryReadFormValue("OrderYear", year) Then Exit Function
    Set m_pendingOrder = Nothing
    If Not m_pageBase.BindingContext.SetValue("Form", "OrderChoice", vbNullString) Then Exit Function

    ' Dropdown contents depend only on today and neighboring rows.
    If Not m_orderResolver.TryGetNearby(displayMatches, diagnostic) Then GoTo Failed
    Set m_ordersByCaption = VBA.CreateObject("Scripting.Dictionary")
    Set options = New Collection
    Set todayOptions = New Collection
    For Each record In displayMatches
        caption = private_OrderText(record, "OrderChoiceFormat")
        options.Add caption
        Set m_ordersByCaption(caption) = record
        If record.TryGetValue("OrderDate", orderDate) Then
            If VBA.CDate(orderDate) = VBA.Date Then todayOptions.Add caption
        End If
    Next record
    If Not m_pageBase.BindingContext.SetObject("Data", "OrderOptions", options) Then Exit Function
    If Not m_pageBase.BindingContext.SetObject("Data", "TodayOrderOptions", todayOptions) Then Exit Function
    ex_Core.fn_Diagnostic_WriteLog "ORDER_DROPDOWN_NEARBY | Items=" & VBA.CStr(options.Count)

    ' Input resolution never changes the dropdown contents.
    If Not m_orderResolver.TrySearch(VBA.CStr(query), VBA.CStr(year), matches, diagnostic) Then GoTo Failed
    If VBA.Len(VBA.Trim$(VBA.CStr(query))) = 0 Then
        status = m_messages("OrderReady")
    ElseIf matches.Count = 1 Then
        Set m_pendingOrder = matches(1)
        If Not m_pageBase.BindingContext.SetValue("Form", "OrderChoice", private_OrderText(m_pendingOrder, "OrderChoiceFormat")) Then Exit Function
        SearchOrdersHandler = private_ApplyResolvedOrder()
        Exit Function
    ElseIf matches.Count > 1 Then
        status = VBA.Replace$(m_messages("OrderFoundMany"), "{count}", VBA.CStr(matches.Count))
    Else
        status = m_messages("OrderEmpty")
        isError = True
    End If
    SearchOrdersHandler = private_SetActionStatus(status, isError)
    Exit Function
EH:
    diagnostic = VBA.Err.Description
Failed:
    SearchOrdersHandler = private_SetActionStatus(diagnostic, True)
End Function

Public Function OrderQueryChangedHandler() As Boolean
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    ' Typed input and dropdown values use the same resolver.
    OrderQueryChangedHandler = Me.SearchOrdersHandler()
End Function


Private Function private_ApplyResolvedOrder() As Boolean
    Dim number As Variant
    Dim orderDate As Variant
    Dim updates As Object
    Dim diagnostic As String
    Dim acceptedText As String
    Dim selectedValue As Variant
    Dim selectedObject As Object
    Dim isObject As Boolean
    Dim selectedKey As String
    Dim selectedRecord As obj_LookupRecord
    Dim record As obj_LookupRecord

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If m_pendingOrder Is Nothing Then
        private_ApplyResolvedOrder = private_SetActionStatus(m_messages("OrderChooseRequired"), True)
        Exit Function
    End If
    On Error GoTo EH
    If Not m_pendingOrder.TryGetValue("OrderNumber", number) Then Exit Function
    If Not m_pendingOrder.TryGetValue("OrderDate", orderDate) Then Exit Function
    acceptedText = private_OrderText(m_pendingOrder, "OrderAccepted")
    If VBA.Len(m_acceptedOrderNumber) > 0 Then
        If m_acceptedOrderNumber = VBA.CStr(number) And VBA.CDate(m_acceptedOrderDate) = VBA.CDate(orderDate) Then
            private_ApplyResolvedOrder = private_SetActionStatus(acceptedText, False)
            Exit Function
        End If
    End If
    Set updates = VBA.CreateObject("Scripting.Dictionary")
    updates("Form.AcceptedOrderNumber") = number
    updates("Form.AcceptedOrderDate") = orderDate
    updates("Form.AcceptedOrderText") = acceptedText

    ' Apply only a uniquely resolved pair; invalid and ambiguous input keeps the previous pair.
    If Not m_pageBase.BindingContext.TryApplyValues(updates, diagnostic) Then GoTo Failed
    m_acceptedOrderNumber = VBA.CStr(number)
    m_acceptedOrderDate = orderDate

    ' Refresh derived tickets and restore the previously selected personnel record.
    If m_pageBase.BindingContext.TryGetValue("Data", "SelectedPersonnel", selectedValue, selectedObject, isObject) Then
        If isObject Then
            If TypeOf selectedObject Is obj_LookupRecord Then
                Set selectedRecord = selectedObject
                selectedKey = selectedRecord.Key
            End If
        End If
    End If
    If Not Me.SearchPersonnelHandler() Then
        private_ApplyResolvedOrder = private_SetActionStatus(m_messages("OrderAppliedRefreshError"), True)
        Exit Function
    End If
    If VBA.Len(selectedKey) > 0 Then
        For Each record In m_lookupResult.Records
            If record.Key = selectedKey Then
                If Not m_pageBase.BindingContext.SetObject("Data", "SelectedPersonnel", record) Then Exit Function
                If Not Me.SelectPersonnelHandler(record) Then Exit Function
                Exit For
            End If
        Next record
    End If
    private_ApplyResolvedOrder = private_SetActionStatus(acceptedText, False)
    Exit Function
EH:
    diagnostic = VBA.Err.Description
Failed:
    private_ApplyResolvedOrder = private_SetActionStatus(diagnostic, True)
End Function

Public Function RefreshPersonnelHandler() As Boolean
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    ex_TableDataRuntime.fn_ResetCaches
    m_lookupService.InvalidateCache
    m_candidateEnricher.InvalidateCache
    m_orderResolver.InvalidateCache
    RefreshPersonnelHandler = Me.SearchPersonnelHandler()
End Function

Public Function ArrivalTreatmentHandler() As Boolean
    ArrivalTreatmentHandler = private_SwitchMode("ArrivalTreatment")
End Function

Public Function ArrivalAnnualLeaveHandler() As Boolean
    ArrivalAnnualLeaveHandler = private_SwitchMode("ArrivalAnnualLeave")
End Function

Public Function ArrivalMedicalLeaveHandler() As Boolean
    ArrivalMedicalLeaveHandler = private_SwitchMode("ArrivalMedicalLeave")
End Function

Public Function ArrivalFamilyLeaveHandler() As Boolean
    ArrivalFamilyLeaveHandler = private_SwitchMode("ArrivalFamilyLeave")
End Function

Public Function DepartureTreatmentHandler() As Boolean
    DepartureTreatmentHandler = private_SwitchMode("DepartureTreatment")
End Function

Public Function DepartureAnnualLeaveHandler() As Boolean
    DepartureAnnualLeaveHandler = private_SwitchMode("DepartureAnnualLeave")
End Function

Public Function DepartureMedicalLeaveHandler() As Boolean
    DepartureMedicalLeaveHandler = private_SwitchMode("DepartureMedicalLeave")
End Function

Public Function DepartureFamilyLeaveHandler() As Boolean
    DepartureFamilyLeaveHandler = private_SwitchMode("DepartureFamilyLeave")
End Function

' //
' // Private
' //
Private Function private_CaptureModeConfig(ByRef diagnostic As String) As Boolean
    Dim sheet As Worksheet
    Dim table As ListObject
    Dim configTable As ListObject
    Dim values As Variant
    Dim keyColumn As Long
    Dim rowIndex As Long
    Dim key As String
    Dim prefix As String

    On Error GoTo EH
    prefix = "PersonalEventBuilder::"
    For Each sheet In ThisWorkbook.Worksheets
        For Each table In sheet.ListObjects
            If table.Name = "tbConfig" Then
                If Not configTable Is Nothing Then Err.Raise 5, , "Duplicate configuration table tbConfig."
                Set configTable = table
            End If
        Next table
    Next sheet
    If configTable Is Nothing Then Err.Raise 5, , "Configuration table tbConfig was not found."
    If configTable.DataBodyRange Is Nothing Then Err.Raise 5, , "Configuration table tbConfig is empty."
    keyColumn = configTable.ListColumns("Key").Index
    If keyColumn >= configTable.ListColumns.Count Then Err.Raise 5, , "Configuration values must follow Key."
    values = configTable.DataBodyRange.Value2
    Set m_modeConfig = VBA.CreateObject("Scripting.Dictionary")
    m_modeConfig.CompareMode = VBA.vbTextCompare
    ' Read the worksheet once; mode switches use the in-memory snapshot.
    For rowIndex = 1 To UBound(values, 1)
        key = VBA.Trim$(VBA.CStr(values(rowIndex, keyColumn)))
        If VBA.Left$(key, VBA.Len(prefix)) = prefix Then
            key = VBA.Mid$(key, VBA.Len(prefix) + 1)
            If m_modeConfig.Exists(key) Then Err.Raise 5, , "Duplicate configuration key: " & prefix & key
            m_modeConfig.Add key, VBA.Trim$(VBA.CStr(values(rowIndex, keyColumn + 1)))
        End If
    Next rowIndex
    private_CaptureModeConfig = True
    Exit Function
EH:
    diagnostic = "Mode configuration: " & VBA.Err.Description
End Function

Private Function private_ReadModeConfig( _
    ByVal key As String, _
    ByRef value As String, _
    ByRef diagnostic As String _
) As Boolean
    value = VBA.vbNullString
    If Not m_modeConfig Is Nothing Then
        If m_modeConfig.Exists(key) Then
            value = m_modeConfig(key)
            private_ReadModeConfig = (VBA.Len(value) > 0)
        End If
    End If
    If Not private_ReadModeConfig Then diagnostic = "Required configuration key not found: PersonalEventBuilder::" & key
End Function

Private Function private_LoadMode( _
    ByVal modeId As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim direction As String
    Dim formKind As String
    Dim caption As String
    Dim countText As String
    Dim fieldName As String
    Dim captionKey As String
    Dim recordName As String
    Dim sourceAlias As String
    Dim fieldCaption As String
    Dim eventType As String
    Dim eventDateLabel As String
    Dim text As String
    Dim mode As Variant
    Dim field As Variant
    Dim index As Long
    Dim prefix As String
    Dim columns As New Collection
    Dim column As Object
    Dim emptyRecords As New Collection
    Dim emptyResult As obj_LookupResult
    Dim modes As Variant

    On Error GoTo EH
    modes = VBA.Array("ArrivalTreatment", _
            "ArrivalAnnualLeave", _
            "ArrivalMedicalLeave", _
            "ArrivalFamilyLeave", _
            "DepartureTreatment", _
            "DepartureAnnualLeave", _
            "DepartureMedicalLeave", _
            "DepartureFamilyLeave")
    If Not private_ReadModeConfig("modes." & modeId & ".direction", direction, diagnostic) Then Exit Function
    If Not private_ReadModeConfig("modes." & modeId & ".form", formKind, diagnostic) Then Exit Function
    If Not private_ReadModeConfig("modes." & modeId & ".caption", caption, diagnostic) Then Exit Function
    If direction <> "Arrival" And direction <> "Departure" Then Err.Raise 5, , "Invalid mode direction: " & modeId
    If formKind <> "Leave" And formKind <> "Treatment" Then Err.Raise 5, , "Invalid mode form: " & modeId
    If Not private_ReadModeConfig("forms." & formKind & ".fields.count", countText, diagnostic) Then Exit Function
    If Not VBA.IsNumeric(countText) Then Err.Raise 5, , "Invalid mode field count: " & formKind
    If VBA.CDbl(countText) < 1 Or VBA.CDbl(countText) > 16 Or VBA.CDbl(countText) <> VBA.Fix(VBA.CDbl(countText)) Then
        Err.Raise 5, , "Mode field count must be an integer from 1 to 16."
    End If
    If Not private_ReadModeConfig("text." & direction & "Type", eventType, diagnostic) Then Exit Function
    If Not private_ReadModeConfig("text.Mode" & direction & "DateLabel", eventDateLabel, diagnostic) Then Exit Function
    For index = 1 To VBA.CLng(countText)
        prefix = "forms." & formKind & ".fields." & VBA.CStr(index)
        If Not private_ReadModeConfig(prefix & ".name", fieldName, diagnostic) Then Exit Function
        If Not private_ReadModeConfig(prefix & ".captionKey", captionKey, diagnostic) Then Exit Function
        If VBA.Left$(captionKey, VBA.Len("PersonalEventBuilder::")) <> "PersonalEventBuilder::" Then
            diagnostic = "Mode caption key must use PersonalEventBuilder:: prefix: " & captionKey
            Exit Function
        End If
        If Not private_ReadModeConfig(VBA.Mid$(captionKey, VBA.Len("PersonalEventBuilder::") + 1), fieldCaption, diagnostic) Then
            diagnostic = "Required caption key not found: " & captionKey
            Exit Function
        End If
        If fieldName = "EventDate" Then fieldCaption = eventDateLabel
        recordName = VBA.vbNullString
        sourceAlias = VBA.vbNullString
        If m_modeConfig.Exists(prefix & ".record") Then recordName = m_modeConfig(prefix & ".record")
        If m_modeConfig.Exists(prefix & ".sourceAlias") Then sourceAlias = m_modeConfig(prefix & ".sourceAlias")
        Set column = VBA.CreateObject("Scripting.Dictionary")
        column("formField") = fieldName
        column("apply") = (VBA.Len(recordName) > 0)
        If VBA.Len(recordName) = 0 Then recordName = "Manual." & fieldName
        column("name") = recordName
        column("caption") = fieldCaption
        column("sourceAlias") = sourceAlias
        columns.Add column
    Next index
    For Each field In VBA.Array("EventType", _
            "Destination", _
            "PersonRank", _
            "PersonName", _
            "PersonId", _
            "PositionCode", _
            "PersonPosition", _
            "ReportRank", _
            "ReportName", _
            "ReportPositionCode", _
            "IncomingNumber", _
            "IncomingDate", _
            "EventDate", _
            "TicketNumber", _
            "TicketDate", _
            "Hospital", _
            "HospitalShortName", _
            "Document", _
            "DocumentDate", _
            "DocumentNumber")
        If Not private_ReadModeConfig("text.Mode" & VBA.CStr(field) & "Label", text, diagnostic) Then Exit Function
        If Not m_pageBase.BindingContext.SetValue("Text", "Mode" & VBA.CStr(field) & "Label", text) Then Exit Function
        If Not m_pageBase.BindingContext.SetValue("Form", VBA.CStr(field), VBA.vbNullString) Then Exit Function
    Next field
    If Not m_pageBase.BindingContext.SetValue("Text", "ModeEventDateLabel", eventDateLabel) Then Exit Function
    For Each mode In modes
        If Not private_ReadModeConfig("modes." & VBA.CStr(mode) & ".caption", text, diagnostic) Then Exit Function
        If Not m_pageBase.BindingContext.SetValue("Text", VBA.CStr(mode), text) Then Exit Function
        text = "secondaryButton"
        If VBA.CStr(mode) = modeId Then text = "activeModeButton"
        If Not m_pageBase.BindingContext.SetValue("Settings", VBA.CStr(mode) & "Style", text) Then Exit Function
    Next mode
    If Not m_pageBase.BindingContext.SetValue("Form", "ModeId", modeId) Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Form", "ModeDirection", direction) Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Form", "ModeForm", formKind) Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Form", "EventType", eventType) Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Data", "SelectedPersonnel", VBA.vbNullString) Then Exit Function
    Set emptyResult = New obj_LookupResult
    If Not emptyResult.Initialize(emptyRecords, columns, False, diagnostic) Then Exit Function
    If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", emptyResult) Then Exit Function
    Set m_lookupResult = emptyResult
    Set m_modeFields = columns
    m_modeId = modeId
    m_modeDirection = direction
    m_eventType = eventType
    private_LoadMode = True
    Exit Function
EH:
    diagnostic = "Mode configuration: " & VBA.Err.Description
End Function

Private Function private_SwitchMode(ByVal modeId As String) As Boolean
    Dim diagnostic As String
    Dim caption As String
    Dim status As String

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If modeId = m_modeId Then
        private_SwitchMode = True
        Exit Function
    End If
    If Not private_LoadMode(modeId, diagnostic) Then GoTo Failed
    If Not private_ReadModeConfig("modes." & modeId & ".caption", caption, diagnostic) Then GoTo Failed
    If Not private_ReadModeConfig("text.ModeSelected", status, diagnostic) Then GoTo Failed
    If Not private_SetActionStatus(VBA.Replace$(status, "{mode}", caption), False) Then Exit Function
    private_SwitchMode = m_pageBase.UpdatePage()
    Exit Function
Failed:
    private_SetActionStatus diagnostic, True
    ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Event mode"
End Function

Private Function private_PrepareModeRecords( _
    ByVal records As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim record As obj_LookupRecord
    Dim supplement As obj_LookupRecord
    Dim fields As Object
    Dim eventDate As Variant

    For Each record In records
        eventDate = VBA.vbNullString
        If m_modeDirection = "Departure" Then record.TryGetValue "Absence.DepartureDate", eventDate
        Set fields = VBA.CreateObject("Scripting.Dictionary")
        fields("Event.Type") = m_eventType
        fields("Event.Date") = eventDate
        Set supplement = New obj_LookupRecord
        If Not supplement.Initialize(record.Key, fields) Then Exit Function
        If Not record.TryMerge(supplement, True) Then
            diagnostic = "Cannot apply event mode fields."
            Exit Function
        End If
    Next record
    private_PrepareModeRecords = True
End Function

Private Function private_ApplyModeRecord( _
    ByVal record As obj_LookupRecord, _
    ByRef diagnostic As String _
) As Boolean
    Dim column As Object
    Dim value As Variant
    Dim updates As Object

    Set updates = VBA.CreateObject("Scripting.Dictionary")
    For Each column In m_modeFields
        If column("apply") Then
            value = VBA.vbNullString
            If Not record.TryGetValue(column("name"), value) Then
                diagnostic = "Candidate field not found: " & column("name")
                Exit Function
            End If
            updates("Form." & column("formField")) = value
        End If
    Next column
    private_ApplyModeRecord = m_pageBase.BindingContext.TryApplyValues(updates, diagnostic)
End Function

Private Function private_TryEnrichResult( _
    ByRef result As obj_LookupResult, _
    ByRef diagnostic As String _
) As Boolean
    Dim records As Collection
    Dim enrichedResult As obj_LookupResult

    If Not m_candidateEnricher.TryEnrich(m_lookupService, result.Records, _
            m_acceptedOrderNumber, m_acceptedOrderDate, records, diagnostic) Then Exit Function
    Set enrichedResult = New obj_LookupResult
    If Not private_PrepareModeRecords(records, diagnostic) Then Exit Function
    If Not enrichedResult.Initialize(records, m_modeFields, result.HasMore, diagnostic) Then Exit Function
    Set result = enrichedResult
    private_TryEnrichResult = True
End Function

Private Function private_SetActionStatus( _
    ByVal text As String, _
    ByVal isError As Boolean _
) As Boolean
    Dim updates As Object
    Dim diagnostic As String

    Set updates = VBA.CreateObject("Scripting.Dictionary")
    updates("Form.ActionStatus") = text
    updates("Form.ActionStatusStyle") = "actionSuccess"
    If isError Then updates("Form.ActionStatusStyle") = "actionError"
    private_SetActionStatus = m_pageBase.BindingContext.TryApplyValues(updates, diagnostic)
End Function

Private Function private_OrderText( _
    ByVal record As obj_LookupRecord, _
    ByVal messageKey As String _
) As String
    Dim number As Variant
    Dim orderDate As Variant

    If Not record.TryGetValue("OrderNumber", number) Then Exit Function
    If Not record.TryGetValue("OrderDate", orderDate) Then Exit Function
    private_OrderText = VBA.Replace$(m_messages(messageKey), "{number}", VBA.CStr(number))
    private_OrderText = VBA.Replace$(private_OrderText, "{date}", VBA.Format$(VBA.CDate(orderDate), "dd.mm.yyyy"))
End Function

Private Function private_TryReadFormValue( _
    ByVal fieldName As String, _
    ByRef outValue As Variant _
) As Boolean
    Dim outObject As Object
    Dim isObject As Boolean

    If m_pageBase Is Nothing Then Exit Function
    private_TryReadFormValue = m_pageBase.BindingContext.TryGetValue( _
        "Form", fieldName, outValue, outObject, isObject)
End Function