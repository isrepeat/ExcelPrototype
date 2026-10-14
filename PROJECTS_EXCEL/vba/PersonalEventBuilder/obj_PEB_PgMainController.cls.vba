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
Private m_orderResolver As obj_PEB_OrderResolver
Private m_ordersByCaption As Object
Private m_pendingOrder As obj_LookupRecord
Private m_acceptedOrderNumber As String
Private m_acceptedOrderDate As Variant

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
    Set m_orderResolver = New obj_PEB_OrderResolver
    If Not m_orderResolver.Initialize(diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Configuration"
        Exit Function
    End If
    If Not m_lookupProfile.Initialize("PersonalEventBuilder::lookup.Personnel", diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Personnel lookup"
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
    If Not m_pageBase.BindingContext.SetValue("Data", "SelectedPersonnel", vbNullString) Then Exit Function
    If Not m_pageBase.BindingContext.TryApplyValues(updates, diagnostic) Then GoTo Failure
    If Not m_lookupService.TrySearch(VBA.CStr(text), result, diagnostic) Then
        If Not m_lookupService.TrySearch(vbNullString, result, status) Then GoTo Failure
        Set m_lookupResult = result
        m_pageBase.BindingContext.SetObject "Data", "PersonnelCandidates", result
        m_pageBase.BindingContext.SetValue "Text", "PersonnelStatus", m_messages("PersonnelError")
        GoTo Failure
    End If
    If Not private_TryEnrichResult(result, diagnostic) Then
        If m_lookupService.TrySearch(vbNullString, result, status) Then
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

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If selected Is Nothing Then Exit Function
    If Not TypeOf selected Is obj_LookupRecord Then Exit Function
    Set record = selected
    If Not m_lookupProfile.TryApply(record, m_pageBase.BindingContext, diagnostic) Then
        private_SetActionStatus VBA.Replace$(m_messages("PersonnelActionError"), "{details}", diagnostic), True
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Personnel lookup"
        Exit Function
    End If
    SelectPersonnelHandler = m_pageBase.BindingContext.SetValue("Text", "PersonnelStatus", m_messages("PersonnelSelected"))
    If SelectPersonnelHandler Then SelectPersonnelHandler = private_SetActionStatus(m_messages("PersonnelSelected"), False)
End Function

Public Function SearchOrdersHandler() As Boolean
    Dim query As Variant
    Dim year As Variant
    Dim choice As Variant
    Dim matches As Collection
    Dim options As Collection
    Dim record As obj_LookupRecord
    Dim caption As String
    Dim diagnostic As String
    Dim status As String
    Dim updates As Object
    Dim todayOptions As Collection
    Dim orderDate As Variant
    Dim showSuggestions As Boolean
    Dim suggestionDiagnostic As String
    Dim displayMatches As Collection
    Dim resolvedCaption As String

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    On Error GoTo EH
    If Not private_TryReadFormValue("OrderQuery", query) Then Exit Function
    If Not private_TryReadFormValue("OrderYear", year) Then Exit Function
    If Not private_TryReadFormValue("OrderChoice", choice) Then Exit Function
    Set m_pendingOrder = Nothing
    Set m_ordersByCaption = VBA.CreateObject("Scripting.Dictionary")
    Set options = New Collection
    Set todayOptions = New Collection
    Set updates = VBA.CreateObject("Scripting.Dictionary")
    updates("Form.OrderChoice") = vbNullString
    updates("Form.OrderQuery") = VBA.CStr(query)
    If Not m_pageBase.BindingContext.TryApplyValues(updates, diagnostic) Then GoTo Failed
    If Not m_pageBase.BindingContext.SetObject("Data", "OrderOptions", options) Then Exit Function
    If Not m_pageBase.BindingContext.SetObject("Data", "TodayOrderOptions", todayOptions) Then Exit Function
    If Not m_orderResolver.TrySearch(VBA.CStr(query), VBA.CStr(year), matches, diagnostic) Then
        status = diagnostic
        showSuggestions = True
    ElseIf matches.Count = 0 And VBA.Len(VBA.Trim$(VBA.CStr(query))) > 0 Then
        status = m_messages("OrderEmpty")
        showSuggestions = True
    End If
    If showSuggestions Then
        ' Keep the error visible, but offer nearby orders as explicit alternatives.
        If Not m_orderResolver.TrySearch(vbNullString, VBA.CStr(year), matches, suggestionDiagnostic) Then
            If VBA.Len(diagnostic) = 0 Then diagnostic = suggestionDiagnostic
            GoTo Failed
        End If
        ex_Core.fn_Diagnostic_WriteLog "ORDER_LOOKUP_SUGGESTIONS | Items=" & VBA.CStr(matches.Count)
    End If

    Set displayMatches = matches
    If Not showSuggestions And VBA.InStr(1, VBA.CStr(query), "/") > 0 And matches.Count = 1 Then
        ' Resolve the pair independently of the nearby orders offered by the dropdown.
        If Not m_orderResolver.TrySearch(vbNullString, VBA.CStr(year), displayMatches, diagnostic) Then GoTo Failed
    End If
    For Each record In displayMatches
        caption = private_OrderText(record, "OrderChoiceFormat")
        options.Add caption
        If record.TryGetValue("OrderDate", orderDate) Then
            If VBA.CDate(orderDate) = VBA.Date Then todayOptions.Add caption
        End If
        Set m_ordersByCaption(caption) = record
    Next record
    If Not showSuggestions And matches.Count = 1 Then
        resolvedCaption = private_OrderText(matches(1), "OrderChoiceFormat")
        If Not m_ordersByCaption.Exists(resolvedCaption) Then
            options.Add resolvedCaption
            Set m_ordersByCaption(resolvedCaption) = matches(1)
        End If
    End If
    If showSuggestions Then
        ' Suggestions require an explicit dropdown selection before acceptance.
        Set m_pendingOrder = Nothing
    ElseIf matches.Count = 1 And VBA.Len(VBA.Trim$(VBA.CStr(query))) > 0 Then
        Set m_pendingOrder = matches(1)
        choice = resolvedCaption
        status = m_messages("OrderFoundOne")
    ElseIf matches.Count > 0 Then
        If m_ordersByCaption.Exists(VBA.CStr(choice)) Then Set m_pendingOrder = m_ordersByCaption(VBA.CStr(choice))
        status = VBA.Replace$(m_messages("OrderFoundMany"), "{count}", VBA.CStr(matches.Count))
    ElseIf VBA.Len(VBA.Trim$(VBA.CStr(query))) = 0 Then
        status = m_messages("OrderReady")
    Else
        status = m_messages("OrderEmpty")
    End If
    If VBA.Len(VBA.Trim$(VBA.CStr(query))) = 0 Then status = m_messages("OrderReady")
    If Not m_pageBase.BindingContext.SetObject("Data", "OrderOptions", options) Then Exit Function
    If Not m_pageBase.BindingContext.SetObject("Data", "TodayOrderOptions", todayOptions) Then Exit Function
    If Not m_pendingOrder Is Nothing Then
        updates("Form.OrderChoice") = VBA.CStr(choice)
    End If
    If Not m_pageBase.BindingContext.TryApplyValues(updates, diagnostic) Then GoTo Failed
    If Not showSuggestions And matches.Count = 1 And VBA.Len(VBA.Trim$(VBA.CStr(query))) > 0 Then
        SearchOrdersHandler = private_ApplyResolvedOrder()
        Exit Function
    End If
    SearchOrdersHandler = private_SetActionStatus(status, showSuggestions)
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

' //
' // Private
' //
Private Function private_TryEnrichResult( _
    ByRef result As obj_LookupResult, _
    ByRef diagnostic As String _
) As Boolean
    Dim records As Collection
    Dim enrichedResult As obj_LookupResult

    If Not m_candidateEnricher.TryEnrich(m_lookupService, result.Records, _
            m_acceptedOrderNumber, m_acceptedOrderDate, records, diagnostic) Then Exit Function
    Set enrichedResult = New obj_LookupResult
    If Not enrichedResult.Initialize(records, m_lookupProfile.DisplayColumns, result.HasMore, diagnostic) Then Exit Function
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