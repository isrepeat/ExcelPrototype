VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_TDB_PgMainController"
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
Private m_placeResult As obj_LookupResult
Private m_placeChainView As obj_LookupResult
Private m_tvoChainView As obj_LookupResult
Private m_lookupIsPlace As Boolean
Private m_messages As Object
Private m_historyResult As obj_LookupResult
Private m_orderResolver As obj_OrderResolver
Private m_orderOptions As Object
Private m_acceptedOrderNumber As String
Private m_acceptedOrderDate As Date
Private m_historyOrderNumber As String
Private m_historyOrderDate As Date
Private m_previousAutoFill As Variant

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
    Dim key As Variant
    Dim valueText As String
    Dim diagnostic As String
    Dim dependencyLoader As obj_TDB_DependencyLoader
    Dim spans As Variant
    Dim column As Object
    Dim columnIndex As Long
    Dim emptyChain As New Collection

    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    Set m_pageBase = pageBase
    If m_pageBase Is Nothing Then
        Exit Function
    End If
    Set dependencyLoader = New obj_TDB_DependencyLoader
    If Not dependencyLoader.Initialize() Then
        Exit Function
    End If
    ex_TableSnapshots.fn_Register dependencyLoader, ex_TravelDocumentsBuilder.fn_Config("text.DependenciesRequired")
    Set m_messages = VBA.CreateObject("Scripting.Dictionary")
    For Each key In VBA.Array("PersonnelError", "PersonnelMinimum", "PersonnelEmpty", _
            "PersonnelCount", "PersonnelMore", "PersonnelSelected")
        valueText = ex_TravelDocumentsBuilder.fn_Config("text." & VBA.CStr(key))
        m_messages(VBA.CStr(key)) = valueText
    Next key
    Set m_lookupProfile = New obj_LookupProfile
    If Not m_lookupProfile.Initialize("TravelDocumentsBuilder::lookup.Personnel", diagnostic) Then
        ex_Core.fn_Diagnostic_WriteLog "TDB_INITIALIZE_FAILED | Stage=LookupProfile | " & diagnostic
        ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation
        Exit Function
    End If
    spans = VBA.Split(ex_TravelDocumentsBuilder.fn_Config("candidates.personnel.columnSpans"), "|")
    If UBound(spans) + 1 <> m_lookupProfile.DisplayColumns.Count Then VBA.Err.Raise 5, , "Candidate column span count mismatch."
    For columnIndex = 1 To m_lookupProfile.DisplayColumns.Count
        Set column = m_lookupProfile.DisplayColumns(columnIndex)
        column("columnSpan") = VBA.CLng(spans(columnIndex - 1))
    Next columnIndex
    If VBA.LCase$(m_lookupProfile.Source.RangeAddress) = "auto" Then
        VBA.Err.Raise 5, , "Explicit configuration range required: TravelDocumentsBuilder::lookup.Personnel.source.range"
    End If
    Set m_lookupService = New obj_LookupService
    If Not m_lookupService.Initialize(m_lookupProfile) Then
        ex_Core.fn_Diagnostic_WriteLog "TDB_INITIALIZE_FAILED | Stage=LookupService"
        Exit Function
    End If
    Set m_orderResolver = New obj_OrderResolver
    If Not m_orderResolver.Initialize("TravelDocumentsBuilder::orders", "TravelDocumentsBuilder::text", diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation
        Exit Function
    End If
    Set m_orderOptions = VBA.CreateObject("Scripting.Dictionary")
    m_orderOptions.CompareMode = VBA.vbTextCompare
    If Not m_lookupService.TrySearch(VBA.vbNullString, m_lookupResult, diagnostic) Then
        ex_Core.fn_Diagnostic_WriteLog "TDB_INITIALIZE_FAILED | Stage=InitialSearch | " & diagnostic
        ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation
        Exit Function
    End If
    If Not private_ValidateCandidateBounds(m_lookupResult, "candidates.personnel.columnOffset") Then Exit Function
    If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", m_lookupResult) Then
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetObject("Data", "HistoryCandidates", m_lookupResult) Then
        Exit Function
    End If
    If Not private_UpdateFieldState() Then
        Exit Function
    End If
    If Not private_PublishPlaces() Then Exit Function
    If Not private_PublishChain(emptyChain) Then Exit Function
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub

    m_isDisposed = True
    m_isInitialized = False
    ex_TDB_WordSession.fn_Module_Dispose
    If Not m_lookupService Is Nothing Then m_lookupService.Dispose
    If Not m_lookupResult Is Nothing Then m_lookupResult.Dispose
    If Not m_placeResult Is Nothing Then m_placeResult.Dispose
    Set m_placeResult = Nothing
    If Not m_placeChainView Is Nothing Then m_placeChainView.Dispose
    Set m_placeChainView = Nothing
    If Not m_tvoChainView Is Nothing Then m_tvoChainView.Dispose
    Set m_tvoChainView = Nothing
    If Not m_lookupProfile Is Nothing Then m_lookupProfile.Dispose
    If Not m_historyResult Is Nothing Then m_historyResult.Dispose
    Set m_historyResult = Nothing
    If Not m_orderResolver Is Nothing Then m_orderResolver.Dispose
    Set m_orderResolver = Nothing
    Set m_orderOptions = Nothing
    Set m_messages = Nothing
    Set m_lookupService = Nothing
    Set m_lookupResult = Nothing
    Set m_lookupProfile = Nothing
    Set m_pageBase = Nothing
End Sub

Public Function SearchPersonnelHandler() As Boolean
    Dim searchText As Variant
    Dim result As obj_LookupResult
    Dim diagnostic As String
    Dim statusText As String
    Dim isTvo As Variant

    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    m_lookupIsPlace = False
    If Not m_pageBase.BindingContext.SetValue("Settings", "CandidateColumnOffset", _
            ex_TravelDocumentsBuilder.fn_Config("candidates.personnel.columnOffset")) Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Settings", "CandidateSelectionMode", _
            ex_TravelDocumentsBuilder.fn_Config("candidates.selectionMode")) Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Form", "CandidatesVisibility", "Visible") Then
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetValue("Form", "HistoryVisibility", "Collapsed") Then
        Exit Function
    End If
    If ex_TableSnapshots.fn_Active() Is Nothing Then
        private_SetActionStatus ex_TravelDocumentsBuilder.fn_Config("text.DependenciesRequired"), True
        Exit Function
    End If
    If Not private_TryReadFormValue("PersonnelQuery", searchText) Then
        Exit Function
    End If
    If Not private_TryReadFormValue("LookupIsTvo", isTvo) Then
        Exit Function
    End If
    If VBA.CBool(isTvo) Then
        If Not private_ClearTvo(True) Then
            Exit Function
        End If
    End If
    If Not m_pageBase.BindingContext.SetValue("Data", "SelectedPersonnel", VBA.vbNullString) Then
        Exit Function
    End If
    If Not m_lookupService.TrySearch(VBA.CStr(searchText), result, diagnostic) Then
        ex_Core.fn_Diagnostic_WriteLog "TDB_PERSONNEL_SEARCH_FAILED | " & diagnostic
        private_SetActionStatus diagnostic, True
        Exit Function
    End If
    Set m_lookupResult = result
    If Not private_ValidateCandidateBounds(result, "candidates.personnel.columnOffset") Then Exit Function
    If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", result) Then
        Exit Function
    End If
    statusText = private_BuildPersonnelStatus(VBA.CStr(searchText), result)
    SearchPersonnelHandler = private_SetActionStatus(statusText, False)
End Function

Public Function PersonNameChangedHandler() As Boolean
    Dim personName As Variant
    Dim fieldName As Variant
    Dim autoFill As Variant

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not private_TryReadFormValue("AutoFillPersonnel", autoFill) Then Exit Function
    If Not VBA.CBool(autoFill) Then
        PersonNameChangedHandler = True
        Exit Function
    End If
    If Not private_TryReadFormValue("PersonName", personName) Then Exit Function
    For Each fieldName In VBA.Array("PersonId", "PersonRank", "PersonPosition", "PersonUnit", "PersonPositionCode")
        If Not m_pageBase.BindingContext.SetValue("Form", VBA.CStr(fieldName), VBA.vbNullString) Then Exit Function
    Next fieldName
    If Not m_pageBase.BindingContext.SetValue("Form", "PersonnelQuery", VBA.CStr(personName)) Then Exit Function
    If Not private_SetLookupTarget(False) Then Exit Function
    If Not private_UpdateFieldState() Then Exit Function
    PersonNameChangedHandler = Me.SearchPersonnelHandler()
End Function

Public Function TvoNameChangedHandler() As Boolean
    Dim name As Variant
    Dim autoFill As Variant

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not private_TryReadFormValue("AutoFillPersonnel", autoFill) Then Exit Function
    If Not VBA.CBool(autoFill) Then
        TvoNameChangedHandler = True
        Exit Function
    End If
    If Not private_TryReadFormValue("TvoName", name) Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Form", "PersonnelQuery", VBA.CStr(name)) Then Exit Function
    If Not private_SetLookupTarget(True) Then Exit Function
    TvoNameChangedHandler = Me.SearchPersonnelHandler()
End Function

Public Function VacationAbroadChangedHandler() As Boolean
    Dim field As Variant

    On Error GoTo Failed
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    For Each field In ex_TDB_Places.fn_FormFields()
        If Not m_pageBase.BindingContext.SetValue("Form", VBA.CStr(field), VBA.vbNullString) Then Exit Function
    Next field
    If Not m_pageBase.BindingContext.SetValue("Data", "SelectedPersonnel", VBA.vbNullString) Then Exit Function
    If Not private_UpdatePlaceFields() Then Exit Function
    VacationAbroadChangedHandler = Me.PlaceChangedHandler()
    Exit Function
Failed:
    private_SetActionStatus VBA.Err.Description, True
End Function

Public Function PlaceChangedHandler() As Boolean
    Dim query As Variant
    Dim vacationKind As Variant
    Dim result As obj_LookupResult
    Dim status As String
    Dim autoFill As Variant
    Dim field As Variant
    Dim abroad As Variant
    Dim isAbroad As Boolean
    Dim prefix As String

    On Error GoTo Failed
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not private_TryReadFormValue("VacationKind", vacationKind) Then Exit Function
    If ex_TravelDocumentsBuilder.fn_IsMedicalCompany(VBA.CStr(vacationKind)) Then
        PlaceChangedHandler = True
        Exit Function
    End If
    If Not private_TryReadFormValue("DocumentTarget", query) Then Exit Function
    If Not private_TryReadFormValue("AutoFillPersonnel", autoFill) Then Exit Function
    If VBA.CBool(autoFill) Then
        For Each field In ex_TDB_Places.fn_FormFields()
            If Not m_pageBase.BindingContext.SetValue("Form", VBA.CStr(field), VBA.vbNullString) Then Exit Function
        Next field
    Else
        query = VBA.vbNullString
    End If
    If Not private_TryReadFormValue("VacationAbroad", abroad) Then Exit Function
    isAbroad = ex_TDB_Places.fn_IsAbroad(VBA.CStr(abroad))
    prefix = ex_TDB_Places.fn_Prefix(isAbroad)
    Set result = ex_TDB_Places.fn_Search(VBA.CStr(query), isAbroad)
    If Not private_ValidateCandidateBounds(result, prefix & ".columnOffset") Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Settings", "CandidateColumnOffset", _
            ex_TravelDocumentsBuilder.fn_Config(prefix & ".columnOffset")) Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Data", "SelectedPersonnel", VBA.vbNullString) Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Settings", "CandidateSelectionMode", "Confirm") Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Form", "HistoryVisibility", "Collapsed") Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Form", "CandidatesVisibility", "Visible") Then Exit Function
    If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", result) Then Exit Function
    m_lookupIsPlace = True
    If Not m_placeResult Is Nothing Then m_placeResult.Dispose
    Set m_placeResult = result
    If VBA.Len(ex_TableQuery.fn_NormalizeText(VBA.CStr(query))) < VBA.CLng(ex_TravelDocumentsBuilder.fn_Config(prefix & ".minChars")) Then
        status = VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config("lookup.Places.text.Minimum"), _
            "{minChars}", ex_TravelDocumentsBuilder.fn_Config(prefix & ".minChars"))
    ElseIf result.RowCount = 0 Then
        status = ex_TravelDocumentsBuilder.fn_Config("lookup.Places.text.Empty")
    ElseIf result.HasMore Then
        status = VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config("lookup.Places.text.More"), _
            "{count}", VBA.CStr(result.RowCount))
    Else
        status = ex_TravelDocumentsBuilder.fn_Config("lookup.Places.text.Choose")
    End If
    PlaceChangedHandler = private_SetActionStatus(status, False)
    Exit Function
Failed:
    private_SetActionStatus VBA.Err.Description, True
End Function

Public Function SelectPersonnelHandler(ByVal selected As Object) As Boolean
    Dim diagnostic As String
    Dim autoFillPersonnel As Variant
    Dim address As Variant
    Dim mapping As Variant
    Dim pair As Variant
    Dim placeValue As Variant
    Dim updates As Object
    Dim abroad As Variant
    Dim prefix As String

    If selected Is Nothing Or Not TypeOf selected Is obj_LookupRecord Then
        Exit Function
    End If
    If m_lookupIsPlace Then
        If Not private_TryReadFormValue("AutoFillPersonnel", autoFillPersonnel) Then Exit Function
        If Not VBA.CBool(autoFillPersonnel) Then
            SelectPersonnelHandler = True
            Exit Function
        End If
        If Not selected.TryGetValue("Address", address) Then Exit Function
        Set updates = VBA.CreateObject("Scripting.Dictionary")
        updates("Form.DocumentTarget") = address
        For Each mapping In ex_TDB_Places.fn_FormFields()
            updates("Form." & VBA.CStr(mapping)) = VBA.vbNullString
        Next mapping
        If Not selected.TryGetValue("Abroad", abroad) Then Exit Function
        prefix = ex_TDB_Places.fn_Prefix(ex_TDB_Places.fn_IsAbroad(VBA.CStr(abroad)))
        For Each mapping In VBA.Split(ex_TravelDocumentsBuilder.fn_Config(prefix & ".formFields"), "|")
            pair = VBA.Split(mapping, "=")
            If Not selected.TryGetValue(VBA.CStr(pair(0)), placeValue) Then Exit Function
            updates("Form." & pair(1)) = placeValue
        Next mapping
        If Not m_pageBase.BindingContext.TryApplyValues(updates, diagnostic) Then Exit Function
        SelectPersonnelHandler = private_SetActionStatus(ex_TravelDocumentsBuilder.fn_Config("lookup.Places.text.Selected"), False)
        Exit Function
    End If
    If Not private_TryReadFormValue("AutoFillPersonnel", autoFillPersonnel) Then
        Exit Function
    End If
    If Not VBA.CBool(autoFillPersonnel) Then
        SelectPersonnelHandler = True
        Exit Function
    End If
    If Not private_ApplySelectedPersonnel(selected, diagnostic) Then
        ex_Core.fn_Diagnostic_WriteLog "TDB_PERSONNEL_APPLY_FAILED | " & diagnostic
        private_SetActionStatus diagnostic, True
        Exit Function
    End If
    SelectPersonnelHandler = private_SetActionStatus(m_messages("PersonnelSelected"), False)
    If SelectPersonnelHandler Then
        SelectPersonnelHandler = private_UpdateFieldState()
    End If
End Function

Public Function RefreshPersonnelHandler() As Boolean
    Dim emptyResult As obj_LookupResult
    Dim diagnostic As String
    Dim statusText As String

    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    On Error GoTo Failed
    If Not ex_TableSnapshots.fn_Refresh(diagnostic) Then
        GoTo Failed
    End If
    m_lookupService.InvalidateCache
    m_orderResolver.InvalidateCache
    m_orderOptions.RemoveAll
    If Not Me.SearchOrdersHandler() Then
        Exit Function
    End If
    If Not m_lookupService.TrySearch(VBA.vbNullString, emptyResult, diagnostic) Then
        private_SetActionStatus diagnostic, True
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetValue("Data", "SelectedPersonnel", VBA.vbNullString) Then
        Exit Function
    End If
    If Not private_ValidateCandidateBounds(emptyResult, "candidates.personnel.columnOffset") Then Exit Function
    If Not m_pageBase.BindingContext.SetObject("Data", "PersonnelCandidates", emptyResult) Then
        Exit Function
    End If
    If Not m_lookupResult Is Nothing Then
        m_lookupResult.Dispose
    End If
    Set m_lookupResult = emptyResult
    m_lookupIsPlace = False
    If Not m_pageBase.BindingContext.SetValue("Settings", "CandidateColumnOffset", _
            ex_TravelDocumentsBuilder.fn_Config("candidates.personnel.columnOffset")) Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Settings", "CandidateSelectionMode", _
            ex_TravelDocumentsBuilder.fn_Config("candidates.selectionMode")) Then Exit Function
    statusText = private_BuildPersonnelStatus(VBA.vbNullString, emptyResult)
    RefreshPersonnelHandler = private_SetActionStatus(statusText, False)
    ex_Core.fn_StatusBar_Show statusText
    Exit Function
Failed:
    If VBA.Len(diagnostic) = 0 Then
        diagnostic = VBA.Err.Description
    End If
    ex_Core.fn_StatusBar_Show diagnostic
    private_SetActionStatus diagnostic, True
End Function

Public Function VacationKindChangedHandler() As Boolean
    Dim vacationKind As Variant
    Dim documentType As Variant
    Dim prefix As String
    Dim visibility As String
    Dim destination As Variant
    Dim fixedDestination As String

    If Not private_TryReadFormValue("VacationKind", vacationKind) Then Exit Function
    If Not private_TryReadFormValue("DocumentType", documentType) Then Exit Function
    If Not ex_TravelDocumentsBuilder.fn_TryGetDocumentPrefix(VBA.CStr(documentType), prefix) Then Exit Function
    visibility = "Collapsed"
    If prefix = "VacationTicket" And ex_TravelDocumentsBuilder.fn_IsTreatment(VBA.CStr(vacationKind)) Then
        visibility = "Visible"
    End If
    If Not m_pageBase.BindingContext.SetValue("Form", "TreatmentVisibility", visibility) Then Exit Function
    fixedDestination = ex_TravelDocumentsBuilder.fn_Config("generation.medicalCompany.destination")
    visibility = "Visible"
    If prefix = "VacationTicket" And ex_TravelDocumentsBuilder.fn_IsMedicalCompany(VBA.CStr(vacationKind)) Then
        visibility = "Collapsed"
        If Not m_pageBase.BindingContext.SetValue("Form", "DocumentTarget", fixedDestination) Then Exit Function
    Else
        If Not private_TryReadFormValue("DocumentTarget", destination) Then Exit Function
        If VBA.CStr(destination) = fixedDestination Then
            If Not m_pageBase.BindingContext.SetValue("Form", "DocumentTarget", VBA.vbNullString) Then Exit Function
        End If
    End If
    If Not m_pageBase.BindingContext.SetValue("Form", "DestinationVisibility", visibility) Then Exit Function
    VacationKindChangedHandler = private_UpdatePlaceFields()
End Function

Public Function DocumentTypeChangedHandler() As Boolean
    Dim documentType As Variant
    Dim prefix As String
    Dim key As Variant
    Dim valueText As String

    If Not private_TryReadFormValue("DocumentType", documentType) Then Exit Function
    If Not ex_TravelDocumentsBuilder.fn_TryGetDocumentPrefix(VBA.CStr(documentType), prefix) Then Exit Function
    For Each key In VBA.Array("DocumentFormTitle", "DocumentTargetLabel", _
            "DepartureDateLabel", "ReturnDateLabel", "PurposeLabel")
        If Not ex_Core.fn_TryGetWorkbookConfigValue( _
                "TravelDocumentsBuilder::text." & prefix & "." & VBA.CStr(key), valueText) Then Exit Function
        If Not m_pageBase.BindingContext.SetValue("Text", VBA.CStr(key), valueText) Then Exit Function
    Next key
    ' Bindings refresh captions; rebuilding the page during a click disposes the active handler.
    DocumentTypeChangedHandler = Me.VacationKindChangedHandler()
End Function

Public Function ClearFormHandler() As Boolean
    Dim updates As Object
    Dim key As Variant
    Dim valueText As String
    Dim diagnostic As String
    Dim emptyChain As New Collection
    Dim currentAutoFill As Variant
    Dim emptyPlaces As New Collection

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    Set updates = VBA.CreateObject("Scripting.Dictionary")
    For Each key In VBA.Array("PersonName", "PersonId", "PersonRank", "PersonPosition", _
            "PersonUnit", "PersonPositionCode", "PersonnelQuery", "DocumentTarget", "DepartureDate", _
            "PlaceRegion", "PlaceType", "PlaceName", "PlaceDistrict", _
            "PlaceAbroad", "PlaceCountry", "PlaceCountryEn", "PlaceRegionEn", "PlaceNameEn", _
            "VlkNo", "VlkDate", "LongTreatmentDays", "TreatmentLeaveDays", _
            "ReturnDate", "Purpose", "ReportNo", "OrderReference", "TvoName", "TvoIpn", "TvoPosition", _
            "TicketNo", "DocumentPath", "GenerationStatus", "HistoryRowId", "HistoryStamp", "ChangeReason")
        updates("Form." & VBA.CStr(key)) = VBA.vbNullString
    Next key
    For Each key In VBA.Array("VacationKind", "VacationAbroad", "VacationStatus", _
            "VacationDays", "RoadDays", "DonationDays")
        If Not ex_TravelDocumentsBuilder.fn_TryGetConfig( _
                "generation.defaults." & VBA.CStr(key), valueText) Then Exit Function
        updates("Form." & VBA.CStr(key)) = valueText
    Next key
    If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("documentTypes.default", valueText) Then Exit Function
    updates("Form.DocumentType") = valueText
    updates("Form.LookupIsTvo") = False
    If Not private_TryReadFormValue("AutoFillPersonnel", currentAutoFill) Then Exit Function
    updates("Form.AutoFillPersonnel") = currentAutoFill
    If Not VBA.IsEmpty(m_previousAutoFill) Then updates("Form.AutoFillPersonnel") = m_previousAutoFill
    If VBA.Len(m_acceptedOrderNumber) > 0 Then
        updates("Form.OrderReference") = private_OrderCaption(m_acceptedOrderNumber, m_acceptedOrderDate)
    End If
    updates("Text.HistoryEditMode") = ex_TDB_History.fn_Config("text.NewMode")
    updates("Text.SaveHistory") = ex_TravelDocumentsBuilder.fn_Config("text.SaveHistory")
    updates("Form.GenerationStatusStyle") = "generationSuccess"
    updates("Form.CandidatesVisibility") = "Visible"
    updates("Form.HistoryVisibility") = "Collapsed"
    If Not m_pageBase.BindingContext.TryApplyValues(updates, diagnostic) Then
        private_SetActionStatus diagnostic, True
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetValue("Data", "SelectedHistory", VBA.vbNullString) Then Exit Function
    m_previousAutoFill = Empty
    If Not private_UpdateFieldState() Then Exit Function
    If Not Me.DocumentTypeChangedHandler() Then Exit Function
    If Not m_pageBase.BindingContext.SetObject("Data", "TvoChain", emptyChain) Then Exit Function
    If Not m_pageBase.BindingContext.SetObject("Data", "PlaceChain", emptyPlaces) Then Exit Function
    If Not private_PublishPlaces() Then Exit Function
    If Not private_PublishChain(emptyChain) Then Exit Function
    If ex_TableSnapshots.fn_Active() Is Nothing Then
        ClearFormHandler = True
        Exit Function
    End If
    ClearFormHandler = Me.SearchPersonnelHandler()
End Function

Public Function SaveHistoryHandler() As Boolean
    Dim identifier As Variant

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not private_TryReadFormValue("HistoryRowId", identifier) Then Exit Function
    SaveHistoryHandler = private_SaveEvent(VBA.Len(VBA.CStr(identifier)) > 0)
End Function

Public Function ExportWordHandler() As Boolean
    Dim orderWordExport As obj_TDB_OrderWordExport
    Dim documentType As Variant
    Dim prefix As String
    Dim templatePath As String
    Dim outputRoot As String
    Dim diagnostic As String
    Dim succeeded As Boolean

    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    On Error GoTo Failed
    If VBA.Len(m_acceptedOrderNumber) = 0 Or m_acceptedOrderDate = 0 Then
        private_SetActionStatus ex_TDB_History.fn_Config("text.ChooseOrder"), True
        Exit Function
    End If
    If Not private_TryReadFormValue("DocumentType", documentType) Then
        Exit Function
    End If
    If Not ex_TravelDocumentsBuilder.fn_TryGetDocumentPrefix(VBA.CStr(documentType), prefix) Then
        Exit Function
    End If
    If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("generation.templates." & prefix, templatePath) Then
        Exit Function
    End If
    If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("generation.defaults.OutputFolderPath", outputRoot) Then
        Exit Function
    End If
    Set orderWordExport = New obj_TDB_OrderWordExport
    If Not orderWordExport.Initialize(m_pageBase.BindingContext, m_lookupService) Then
        VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("word.text.Initialize")
    End If
    succeeded = orderWordExport.TryExport(m_acceptedOrderNumber, m_acceptedOrderDate, VBA.CStr(templatePath), VBA.CStr(outputRoot), diagnostic)
    orderWordExport.Dispose
    m_pageBase.BindingContext.SetValue "Form", "GenerationStatusStyle", "generationError"
    If succeeded Then
        m_pageBase.BindingContext.SetValue "Form", "GenerationStatusStyle", "generationSuccess"
    End If
    m_pageBase.BindingContext.SetValue "Form", "GenerationStatus", diagnostic
    ex_Core.fn_StatusBar_Show diagnostic
    ExportWordHandler = succeeded
    Exit Function
Failed:
    diagnostic = VBA.Err.Description
    If Not orderWordExport Is Nothing Then
        orderWordExport.Dispose
    End If
    m_pageBase.BindingContext.SetValue "Form", "GenerationStatusStyle", "generationError"
    m_pageBase.BindingContext.SetValue "Form", "GenerationStatus", diagnostic
End Function

Public Function LoadHistoryHandler() As Boolean
    Dim result As obj_LookupResult
    Dim status As String

    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    On Error GoTo Failed
    If VBA.Len(m_acceptedOrderNumber) = 0 Or m_acceptedOrderDate = 0 Then
        private_SetActionStatus ex_TDB_History.fn_Config("text.ChooseOrder"), True
        Exit Function
    End If
    Set result = ex_TDB_History.fn_List(m_acceptedOrderNumber, m_acceptedOrderDate)
    If Not m_pageBase.BindingContext.SetObject("Data", "HistoryCandidates", result) Then
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetValue("Data", "SelectedHistory", VBA.vbNullString) Then
        Exit Function
    End If
    If Not m_historyResult Is Nothing Then
        m_historyResult.Dispose
    End If
    Set m_historyResult = result
    m_historyOrderNumber = m_acceptedOrderNumber
    m_historyOrderDate = m_acceptedOrderDate
    If Not m_pageBase.BindingContext.SetValue("Form", "CandidatesVisibility", "Collapsed") Then
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetValue("Form", "HistoryVisibility", "Visible") Then
        Exit Function
    End If
    status = ex_TDB_History.fn_Config("text.SelectHint")
    If result.RowCount = 0 Then
        status = ex_TDB_History.fn_Config("text.EmptyOrder")
    End If
    LoadHistoryHandler = private_SetActionStatus(status, False)
    Exit Function
Failed:
    private_SetActionStatus VBA.Err.Description, True
End Function

Public Function SelectHistoryHandler(ByVal selected As Object) As Boolean
    Dim identifier As Variant
    Dim record As obj_LookupRecord
    Dim chain As Collection
    Dim stamp As Variant
    Dim value As Variant
    Dim savedAutoFill As Variant

    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    On Error GoTo Failed
    If Not TypeOf selected Is obj_LookupRecord Then
        Exit Function
    End If
    Set record = selected
    If VBA.Len(m_acceptedOrderNumber) = 0 Or m_historyOrderNumber <> m_acceptedOrderNumber Or _
            m_historyOrderDate <> m_acceptedOrderDate Then ex_TDB_History.fn_Fail "OrderStale"
    If Not record.TryGetValue(ex_TDB_History.fn_Config("orderNumberField"), value) Then
        Exit Function
    End If
    If VBA.CStr(value) <> m_acceptedOrderNumber Then
        ex_TDB_History.fn_Fail "OrderStale"
    End If
    If Not record.TryGetValue("OrderDate", value) Then
        Exit Function
    End If
    If VBA.CStr(value) <> VBA.Format$(m_acceptedOrderDate, "dd.mm.yyyy") Then
        ex_TDB_History.fn_Fail "OrderStale"
    End If
    If Not record.TryGetValue("RowId", identifier) Then
        Exit Function
    End If
    If Not record.TryGetValue("HistoryStamp", stamp) Then
        Exit Function
    End If
    If VBA.IsEmpty(m_previousAutoFill) Then
        If Not private_TryReadFormValue("AutoFillPersonnel", savedAutoFill) Then
            Exit Function
        End If
    End If
    Set chain = ex_TDB_History.fn_Load(VBA.CStr(identifier), m_pageBase.BindingContext, VBA.CStr(stamp))
    If VBA.IsEmpty(m_previousAutoFill) Then
        m_previousAutoFill = savedAutoFill
    End If
    If Not m_pageBase.BindingContext.SetValue("Form", "AutoFillPersonnel", False) Then
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetValue("Data", "SelectedPersonnel", VBA.vbNullString) Then
        Exit Function
    End If
    If Not private_SetLookupTarget(False) Then
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetValue("Text", "HistoryEditMode", VBA.Replace$(ex_TDB_History.fn_Config("text.Editing"), "{id}", VBA.CStr(identifier))) Then
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetValue("Text", "SaveHistory", ex_TravelDocumentsBuilder.fn_Config("text.UpdateHistory")) Then
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetValue("Form", "DocumentType", ex_TravelDocumentsBuilder.fn_Config("text.VacationTicket.DocumentFormTitle")) Then
        Exit Function
    End If
    If Not Me.DocumentTypeChangedHandler() Then
        Exit Function
    End If
    If Not private_PublishChain(chain) Then
        Exit Function
    End If
    If Not private_PublishPlaces() Then Exit Function
    If Not Me.VacationDatesChangedHandler() Then
        Exit Function
    End If
    If Not private_UpdateFieldState() Then
        Exit Function
    End If
    SelectHistoryHandler = private_SetActionStatus(ex_TDB_History.fn_Config("text.Loaded"), False)
    Exit Function
Failed:
    private_SetActionStatus VBA.Err.Description, True
End Function

Public Function CloseHistoryHandler() As Boolean
    Dim query As Variant

    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetValue("Form", "HistoryVisibility", "Collapsed") Then
        Exit Function
    End If
    If Not m_lookupResult Is Nothing Then
        If Not private_TryReadFormValue("PersonnelQuery", query) Then
            Exit Function
        End If
        If Not private_SetActionStatus(private_BuildPersonnelStatus(VBA.CStr(query), m_lookupResult), False) Then
            Exit Function
        End If
    End If
    CloseHistoryHandler = m_pageBase.BindingContext.SetValue("Form", "CandidatesVisibility", "Visible")
End Function

Public Function AutoFillChangedHandler() As Boolean
    Dim autoFill As Variant
    Dim selected As Object
    Dim scalar As Variant
    Dim isObject As Boolean

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not private_TryReadFormValue("AutoFillPersonnel", autoFill) Then Exit Function
    If VBA.VarType(autoFill) <> VBA.vbBoolean Then Exit Function
    If VBA.CBool(autoFill) Then
        If Not m_pageBase.BindingContext.TryGetValue("Data", "SelectedPersonnel", scalar, selected, isObject) Then Exit Function
        If isObject Then
            If Not Me.SelectPersonnelHandler(selected) Then Exit Function
        End If
    End If
    If Not private_UpdateFieldState() Then Exit Function
    If Not private_UpdatePlaceFields() Then Exit Function
    AutoFillChangedHandler = m_pageBase.FlushUpdates()
End Function

Public Function SearchOrdersHandler() As Boolean
    Dim query As Variant
    Dim year As Variant
    Dim matches As Collection
    Dim record As obj_LookupRecord
    Dim diagnostic As String

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    On Error GoTo Failed
    If ex_TableSnapshots.fn_Active() Is Nothing Then
        VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("text.DependenciesRequired")
    End If
    If Not private_TryReadFormValue("OrderQuery", query) Then Exit Function
    If Not private_TryReadFormValue("OrderYear", year) Then Exit Function
    query = VBA.Trim$(VBA.CStr(query))
    year = VBA.Trim$(VBA.CStr(year))
    ' A complete order reference or a date is resolved without the number-only year hint.
    If VBA.Len(VBA.CStr(query)) = 0 Or VBA.CStr(query) Like "*[!0-9]*" Then year = VBA.vbNullString
    m_acceptedOrderNumber = VBA.vbNullString
    m_acceptedOrderDate = 0
    If Not m_pageBase.BindingContext.SetValue("Form", "AcceptedOrderNumber", VBA.vbNullString) Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Form", "AcceptedOrderDate", VBA.vbNullString) Then Exit Function
    If Not Me.CloseHistoryHandler() Then Exit Function
    If Not private_UpdateFieldState() Then Exit Function
    If VBA.Len(VBA.CStr(year)) > 0 Then
        If Not VBA.CStr(year) Like "####" Then VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("text.OrderInvalidYear")
        If VBA.CLng(year) < 1000 Then VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("text.OrderInvalidYear")
    End If
    If m_orderOptions.Exists(VBA.CStr(query)) Then
        Set record = m_orderOptions(VBA.CStr(query))
        SearchOrdersHandler = private_AcceptOrder(record)
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(VBA.CStr(query))) = 0 Then
        If Not m_orderResolver.TryGetNearby(matches, diagnostic) Then GoTo Failed
    Else
        If Not m_orderResolver.TrySearch(VBA.CStr(query), VBA.CStr(year), matches, diagnostic) Then GoTo Failed
    End If
    If Not private_PublishOrders(matches) Then Exit Function
    If matches.Count = 1 And VBA.Len(VBA.Trim$(VBA.CStr(query))) > 0 Then
        Set record = matches(1)
        SearchOrdersHandler = private_AcceptOrder(record)
    Else
        diagnostic = ex_TDB_History.fn_Config("text.ChooseOrder")
        If matches.Count = 0 Then diagnostic = ex_TDB_History.fn_Config("text.OrderEmpty")
        SearchOrdersHandler = private_SetActionStatus(diagnostic, True)
    End If
    Exit Function
Failed:
    If VBA.Len(diagnostic) = 0 Then diagnostic = VBA.Err.Description
    SearchOrdersHandler = private_SetActionStatus(diagnostic, True)
End Function

Public Function AddPlaceHandler() As Boolean
    Dim scalar As Variant
    Dim chain As Object
    Dim isObject As Boolean
    Dim row As Object
    Dim vacationKind As Variant

    On Error GoTo Failed
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not private_TryReadFormValue("VacationKind", vacationKind) Then Exit Function
    If ex_TravelDocumentsBuilder.fn_IsMedicalCompany(VBA.CStr(vacationKind)) Then
        VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("placeChain.text.Fixed")
    End If
    If Not m_pageBase.BindingContext.TryGetValue("Data", "PlaceChain", scalar, chain, isObject) Then Exit Function
    If Not isObject Then ex_TDB_History.fn_Fail "InvalidChain"
    If chain.Count >= VBA.CLng(ex_TravelDocumentsBuilder.fn_Config("placeChain.maxItems")) Then
        VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("placeChain.text.Limit")
    End If
    Set row = ex_TDB_PlaceChain.fn_Draft(m_pageBase.BindingContext)
    ex_TDB_PlaceChain.fn_ClearDraft m_pageBase.BindingContext
    chain.Add row
    AddPlaceHandler = private_PublishPlaces()
    Exit Function
Failed:
    private_SetActionStatus VBA.Err.Description, True
End Function

Public Function RemovePlaceHandler() As Boolean
    Dim scalar As Variant
    Dim chain As Object
    Dim isObject As Boolean
    Dim row As Object
    Dim field As Variant
    Dim updates As Object
    Dim diagnostic As String

    On Error GoTo Failed
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not m_pageBase.BindingContext.TryGetValue("Data", "PlaceChain", scalar, chain, isObject) Then Exit Function
    If Not isObject Then ex_TDB_History.fn_Fail "InvalidChain"
    If chain.Count = 0 Then
        RemovePlaceHandler = True
        Exit Function
    End If
    For Each field In VBA.Split("DocumentTarget|" & ex_TravelDocumentsBuilder.fn_Config("placeChain.fields"), "|")
        If Not private_TryReadFormValue(VBA.CStr(field), scalar) Then Exit Function
        If VBA.Len(VBA.CStr(scalar)) > 0 Then VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("placeChain.text.Pending")
    Next field
    Set row = chain(chain.Count)
    Set updates = VBA.CreateObject("Scripting.Dictionary")
    updates("Form.DocumentTarget") = row("Destination")
    updates("Form.VacationAbroad") = row("PlaceAbroad")
    For Each field In ex_TDB_Places.fn_FormFields()
        updates("Form." & field) = row(field)
    Next field
    If Not m_pageBase.BindingContext.TryApplyValues(updates, diagnostic) Then VBA.Err.Raise 5, , diagnostic
    chain.Remove chain.Count
    If Not private_UpdatePlaceFields() Then Exit Function
    RemovePlaceHandler = private_PublishPlaces()
    Exit Function
Failed:
    private_SetActionStatus VBA.Err.Description, True
End Function

Public Function AddTvoHandler() As Boolean
    Dim name As Variant
    Dim personId As Variant
    Dim position As Variant
    Dim value As Variant
    Dim chain As Object
    Dim isObject As Boolean
    Dim row As Object
    Dim actor As Object

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    On Error GoTo Failed
    If Not private_TryReadFormValue("TvoName", name) Then Exit Function
    If Not private_TryReadFormValue("TvoIpn", personId) Then Exit Function
    If Not private_TryReadFormValue("TvoPosition", position) Then Exit Function
    Set actor = ex_TDB_History.fn_Actor(m_lookupService, VBA.CStr(personId), VBA.CStr(name), VBA.CStr(position))
    If Not m_pageBase.BindingContext.TryGetValue("Data", "TvoChain", value, chain, isObject) Then Exit Function
    If Not isObject Then Exit Function
    For Each row In chain
        If row(ex_TDB_History.fn_Config("personIdField")) = personId Or row("PositionCode") = position Then ex_TDB_History.fn_Fail "InvalidChain"
    Next row
    If Not private_TryReadFormValue("PersonId", value) Then Exit Function
    If VBA.CStr(value) = VBA.CStr(personId) Then ex_TDB_History.fn_Fail "InvalidChain"
    If Not private_TryReadFormValue("PersonPositionCode", value) Then Exit Function
    If VBA.CStr(value) = VBA.CStr(position) Then ex_TDB_History.fn_Fail "InvalidChain"
    chain.Add actor
    If Not private_ClearTvo() Then Exit Function
    AddTvoHandler = private_PublishChain(chain)
    Exit Function
Failed:
    private_SetActionStatus VBA.Err.Description, True
End Function

Public Function RemoveTvoHandler() As Boolean
    Dim value As Variant
    Dim chain As Object
    Dim isObject As Boolean

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not m_pageBase.BindingContext.TryGetValue("Data", "TvoChain", value, chain, isObject) Then Exit Function
    If Not isObject Then Exit Function
    If chain.Count > 0 Then chain.Remove chain.Count
    RemoveTvoHandler = private_PublishChain(chain)
End Function

Public Function VacationDatesChangedHandler() As Boolean
    Dim value As Variant
    Dim parts As Variant
    Dim departure As Date
    Dim totalDays As Long
    Dim key As Variant
    Dim maximum As String
    Dim result As String
    Dim text As String

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    On Error GoTo InvalidInput
    If Not private_TryReadFormValue("DepartureDate", value) Then GoTo InvalidInput
    text = VBA.Trim$(VBA.CStr(value))
    If Not text Like "##.##.####" Then GoTo InvalidInput
    parts = VBA.Split(text, ".")
    departure = VBA.DateSerial(VBA.CInt(parts(2)), VBA.CInt(parts(1)), VBA.CInt(parts(0)))
    If VBA.Format$(departure, "dd.mm.yyyy") <> text Then GoTo InvalidInput
    If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("generation.limits.maxDays", maximum) Then GoTo InvalidInput
    For Each key In VBA.Array("VacationDays", "RoadDays", "DonationDays")
        If Not private_TryReadFormValue(VBA.CStr(key), value) Then GoTo InvalidInput
        text = VBA.Trim$(VBA.CStr(value))
        If VBA.Len(text) = 0 Then
            If VBA.CStr(key) = "VacationDays" Then GoTo InvalidInput
            text = "0"
        End If
        If text Like "*[!0-9]*" Then GoTo InvalidInput
        If VBA.CDbl(text) > VBA.CDbl(maximum) Then GoTo InvalidInput
        totalDays = totalDays + VBA.CLng(text)
    Next key
    result = VBA.Format$(ex_TDB_History.fn_ReturnDate(departure, totalDays, 0), "dd.mm.yyyy")
InvalidInput:
    VacationDatesChangedHandler = m_pageBase.BindingContext.SetValue("Form", "ReturnDate", result)
End Function

' //
' // Private
' //
Private Function private_SetActionStatus( _
    ByVal text As String, _
    ByVal isError As Boolean _
) As Boolean
    Dim updates As Object
    Dim diagnostic As String

    Set updates = VBA.CreateObject("Scripting.Dictionary")
    updates("Form.GenerationStatus") = text
    updates("Form.GenerationStatusStyle") = "generationSuccess"
    If isError Then updates("Form.GenerationStatusStyle") = "generationError"
    private_SetActionStatus = m_pageBase.BindingContext.TryApplyValues(updates, diagnostic)
    If Not private_SetActionStatus Then ex_WindowsUi.fn_ShowMessage text, VBA.vbExclamation
End Function

Private Function private_OrderCaption( _
    ByVal number As String, _
    ByVal orderDate As Date _
) As String
    Dim text As String

    text = ex_TravelDocumentsBuilder.fn_Config("orders.caption")
    text = VBA.Replace$(text, "{number}", number)
    private_OrderCaption = VBA.Replace$(text, "{date}", VBA.Format$(orderDate, "dd.mm.yyyy"))
End Function

Private Function private_PublishOrders(ByVal matches As Collection) As Boolean
    Dim options As New Collection
    Dim record As obj_LookupRecord
    Dim number As Variant
    Dim orderDate As Variant
    Dim caption As String

    m_orderOptions.RemoveAll
    For Each record In matches
        If Not record.TryGetValue("OrderNumber", number) Then Exit Function
        If Not record.TryGetValue("OrderDate", orderDate) Then Exit Function
        caption = private_OrderCaption(VBA.CStr(number), VBA.CDate(orderDate))
        If m_orderOptions.Exists(caption) Then ex_TDB_History.fn_Fail "OrderStale"
        m_orderOptions.Add caption, record
        options.Add caption
    Next record
    private_PublishOrders = m_pageBase.BindingContext.SetObject("Data", "OrderOptions", options)
End Function

Private Function private_AcceptOrder( _
    ByVal record As obj_LookupRecord, _
    Optional ByVal preserveEdit As Boolean = False _
) As Boolean
    Dim number As Variant
    Dim orderDate As Variant
    Dim identifier As Variant
    Dim updates As Object
    Dim diagnostic As String

    If Not record.TryGetValue("OrderNumber", number) Then Exit Function
    If Not record.TryGetValue("OrderDate", orderDate) Then Exit Function
    If Not private_TryReadFormValue("HistoryRowId", identifier) Then Exit Function
    If Not preserveEdit And VBA.Len(VBA.CStr(identifier)) > 0 Then
        If m_historyOrderNumber <> VBA.CStr(number) Or m_historyOrderDate <> VBA.CDate(orderDate) Then
            If Not Me.ClearFormHandler() Then Exit Function
        End If
    End If
    m_acceptedOrderNumber = VBA.CStr(number)
    m_acceptedOrderDate = VBA.CDate(orderDate)
    Set updates = VBA.CreateObject("Scripting.Dictionary")
    updates("Form.AcceptedOrderNumber") = m_acceptedOrderNumber
    updates("Form.AcceptedOrderDate") = VBA.Format$(m_acceptedOrderDate, "dd.mm.yyyy")
    updates("Form.OrderQuery") = private_OrderCaption(m_acceptedOrderNumber, m_acceptedOrderDate)
    updates("Form.OrderYear") = VBA.vbNullString
    updates("Form.OrderReference") = private_OrderCaption(m_acceptedOrderNumber, m_acceptedOrderDate)
    If Not preserveEdit Then
        updates("Form.GenerationStatusStyle") = "generationSuccess"
        updates("Form.GenerationStatus") = VBA.Replace$(ex_TDB_History.fn_Config("text.OrderAccepted"), "{order}", private_OrderCaption(m_acceptedOrderNumber, m_acceptedOrderDate)) & _
            VBA.vbCrLf & private_BuildPersonnelStatus(VBA.vbNullString, m_lookupResult)
    End If
    If Not m_pageBase.BindingContext.TryApplyValues(updates, diagnostic) Then VBA.Err.Raise 5, , diagnostic
    If preserveEdit Then
        m_historyOrderNumber = m_acceptedOrderNumber
        m_historyOrderDate = m_acceptedOrderDate
    End If
    private_AcceptOrder = private_UpdateFieldState()
End Function

Private Function private_UpdateFieldState() As Boolean
    Dim autoFill As Variant
    Dim value As Variant
    Dim field As Variant
    Dim readOnly As Boolean
    Dim labelStyle As String
    Dim style As String
    Dim updates As Object
    Dim diagnostic As String

    If Not private_TryReadFormValue("AutoFillPersonnel", autoFill) Then Exit Function
    If VBA.VarType(autoFill) <> VBA.vbBoolean Then Exit Function
    Set updates = VBA.CreateObject("Scripting.Dictionary")
    If Not private_TryReadFormValue("HistoryRowId", value) Then Exit Function
    updates("Form.ChangeReasonVisibility") = "Collapsed"
    If VBA.Len(VBA.CStr(value)) > 0 Then
        updates("Form.ChangeReasonVisibility") = "Visible"
    End If
    For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("forms.readOnly.fields"), "|")
        If Not private_TryReadFormValue(VBA.CStr(field), value) Then Exit Function
        readOnly = VBA.CBool(autoFill) And VBA.Len(VBA.CStr(value)) > 0
        labelStyle = ex_TravelDocumentsBuilder.fn_Config("forms.sourceStyles.label")
        style = ex_TravelDocumentsBuilder.fn_Config("forms.sourceStyles.default")
        If readOnly Then
            labelStyle = ex_TravelDocumentsBuilder.fn_Config("forms.readOnly.labelStyle")
            style = ex_TravelDocumentsBuilder.fn_Config("forms.sourceStyles.Personnel")
        End If
        updates("Form." & field & "LabelStyle") = labelStyle
        updates("Form." & field & "ReadOnly") = readOnly
        updates("Form." & field & "Style") = style
    Next field
    private_UpdateFieldState = m_pageBase.BindingContext.TryApplyValues(updates, diagnostic)
    If Not private_UpdateFieldState Then VBA.Err.Raise 5, , diagnostic
End Function

Private Function private_PublishPlaces() As Boolean
    Dim scalar As Variant
    Dim chain As Object
    Dim isObject As Boolean
    Dim records As New Collection
    Dim columns As New Collection
    Dim column As Object
    Dim row As Object
    Dim fields As Object
    Dim lookupRecord As obj_LookupRecord
    Dim result As New obj_LookupResult
    Dim index As Long
    Dim line As String
    Dim diagnostic As String
    Dim visibility As String

    If Not m_pageBase.BindingContext.TryGetValue("Data", "PlaceChain", scalar, chain, isObject) Then Exit Function
    If Not isObject Then ex_TDB_History.fn_Fail "InvalidChain"
    Set column = VBA.CreateObject("Scripting.Dictionary")
    column.Add "name", "Address"
    column.Add "caption", ex_TravelDocumentsBuilder.fn_Config("text.PlaceChainTitle")
    column.Add "columnSpan", 3
    columns.Add column
    For index = 1 To chain.Count
        Set row = chain(index)
        line = ex_TravelDocumentsBuilder.fn_Config("placeChain.itemTemplate")
        line = VBA.Replace$(line, "{index}", VBA.CStr(index))
        line = VBA.Replace$(line, "{address}", VBA.CStr(row("Destination")))
        Set fields = VBA.CreateObject("Scripting.Dictionary")
        fields.Add "Address", line
        Set lookupRecord = New obj_LookupRecord
        If Not lookupRecord.Initialize(VBA.CStr(index), fields) Then ex_TDB_History.fn_Fail "InvalidChain"
        records.Add lookupRecord
    Next index
    If Not result.Initialize(records, columns, False, diagnostic) Then VBA.Err.Raise 5, , diagnostic
    If Not m_pageBase.BindingContext.SetObject("Data", "PlaceChainView", result) Then Exit Function
    If Not m_placeChainView Is Nothing Then m_placeChainView.Dispose
    Set m_placeChainView = result
    visibility = "Collapsed"
    If chain.Count > 0 Then visibility = "Visible"
    private_PublishPlaces = m_pageBase.BindingContext.SetValue("Form", "PlaceChainVisibility", visibility)
End Function

Private Function private_PublishChain(ByVal chain As Collection) As Boolean
    Dim row As Object
    Dim text As String
    Dim line As String
    Dim index As Long
    Dim records As New Collection
    Dim columns As New Collection
    Dim column As Object
    Dim fields As Object
    Dim lookupRecord As obj_LookupRecord
    Dim result As New obj_LookupResult
    Dim diagnostic As String
    Dim visibility As String

    Set column = VBA.CreateObject("Scripting.Dictionary")
    column.Add "name", "Summary"
    column.Add "caption", ex_TravelDocumentsBuilder.fn_Config("text.TvoChainTitle")
    column.Add "columnSpan", 3
    columns.Add column
    For index = 1 To chain.Count
        Set row = chain(index)
        line = ex_TDB_History.fn_Config("text.ChainItem")
        line = VBA.Replace$(line, "{index}", VBA.CStr(index))
        line = VBA.Replace$(line, "{name}", VBA.CStr(row(ex_TDB_History.fn_Config("personNameField"))))
        line = VBA.Replace$(line, "{position}", VBA.CStr(row("PositionCode")))
        If VBA.Len(text) > 0 Then text = text & VBA.vbLf
        text = text & line
        Set fields = VBA.CreateObject("Scripting.Dictionary")
        fields.Add "Summary", line
        Set lookupRecord = New obj_LookupRecord
        If Not lookupRecord.Initialize(VBA.CStr(index), fields) Then ex_TDB_History.fn_Fail "InvalidChain"
        records.Add lookupRecord
    Next index
    If chain.Count = 0 Then text = ex_TDB_History.fn_Config("text.ChainEmpty")
    If Not result.Initialize(records, columns, False, diagnostic) Then VBA.Err.Raise 5, , diagnostic
    If Not m_pageBase.BindingContext.SetObject("Data", "TvoChainView", result) Then Exit Function
    If Not m_tvoChainView Is Nothing Then m_tvoChainView.Dispose
    Set m_tvoChainView = result
    visibility = "Collapsed"
    If chain.Count > 0 Then visibility = "Visible"
    If Not m_pageBase.BindingContext.SetValue("Form", "TvoChainVisibility", visibility) Then Exit Function
    private_PublishChain = m_pageBase.BindingContext.SetValue("Text", "TvoChainSummary", text)
End Function

Private Function private_ClearTvo(Optional ByVal preserveName As Boolean = False) As Boolean
    Dim key As Variant

    For Each key In VBA.Array("TvoIpn", "TvoPosition")
        If Not m_pageBase.BindingContext.SetValue("Form", VBA.CStr(key), VBA.vbNullString) Then Exit Function
    Next key
    If Not m_pageBase.BindingContext.SetValue("Data", "SelectedPersonnel", VBA.vbNullString) Then Exit Function
    If Not preserveName Then
        If Not m_pageBase.BindingContext.SetValue("Form", "TvoName", VBA.vbNullString) Then Exit Function
    End If
    private_ClearTvo = private_UpdateFieldState()
End Function

Private Function private_SetLookupTarget(ByVal isTvo As Boolean) As Boolean
    Dim previous As Variant

    If Not private_TryReadFormValue("LookupIsTvo", previous) Then Exit Function
    If VBA.CBool(previous) = isTvo Then
        private_SetLookupTarget = True
        Exit Function
    End If
    If Not m_pageBase.BindingContext.SetValue("Form", "LookupIsTvo", isTvo) Then Exit Function
    private_SetLookupTarget = m_pageBase.BindingContext.SetValue("Data", "SelectedPersonnel", VBA.vbNullString)
End Function

Private Function private_UpdatePlaceFields() As Boolean
    Dim autoFill As Variant
    Dim vacationKind As Variant
    Dim visibility As String
    Dim abroad As Variant
    Dim domesticVisibility As String
    Dim foreignVisibility As String
    Dim prefix As String
    Dim field As Variant
    Dim emptyPlaces As New Collection

    If Not private_TryReadFormValue("AutoFillPersonnel", autoFill) Then Exit Function
    If Not private_TryReadFormValue("VacationKind", vacationKind) Then Exit Function
    visibility = "Collapsed"
    If ex_TravelDocumentsBuilder.fn_IsMedicalCompany(VBA.CStr(vacationKind)) Then
        If Not m_pageBase.BindingContext.SetObject("Data", "PlaceChain", emptyPlaces) Then Exit Function
        If Not private_PublishPlaces() Then Exit Function
        For Each field In ex_TDB_Places.fn_FormFields()
            If Not m_pageBase.BindingContext.SetValue("Form", VBA.CStr(field), VBA.vbNullString) Then Exit Function
        Next field
    ElseIf Not VBA.CBool(autoFill) Then
        visibility = "Visible"
    End If
    If Not private_TryReadFormValue("VacationAbroad", abroad) Then Exit Function
    prefix = ex_TDB_Places.fn_Prefix(ex_TDB_Places.fn_IsAbroad(VBA.CStr(abroad)))
    If Not m_pageBase.BindingContext.SetValue("Text", "PlaceRegion", ex_TravelDocumentsBuilder.fn_Config(prefix & ".columns.Region")) Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Text", "PlaceName", ex_TravelDocumentsBuilder.fn_Config(prefix & ".columns.Name")) Then Exit Function
    domesticVisibility = "Collapsed"
    foreignVisibility = "Collapsed"
    If visibility = "Visible" Then
        If ex_TDB_Places.fn_IsAbroad(VBA.CStr(abroad)) Then
            foreignVisibility = "Visible"
        Else
            domesticVisibility = "Visible"
        End If
    End If
    If Not m_pageBase.BindingContext.SetValue("Form", "DomesticPlaceFieldsVisibility", domesticVisibility) Then Exit Function
    If Not m_pageBase.BindingContext.SetValue("Form", "ForeignPlaceFieldsVisibility", foreignVisibility) Then Exit Function
    private_UpdatePlaceFields = m_pageBase.BindingContext.SetValue("Form", "PlaceFieldsVisibility", visibility)
End Function

Private Function private_ValidateCandidateBounds( _
    ByVal lookupResult As obj_LookupResult, _
    ByVal offsetKey As String _
) As Boolean
    Dim tableSource As obj_IUiTableSource
    Dim rawTable As obj_UiRawTable
    Dim columnOffset As Long
    Dim maxColumns As Long
    Dim message As String

    Set tableSource = lookupResult
    Set rawTable = tableSource.GetTable(1)
    columnOffset = VBA.CLng(ex_TravelDocumentsBuilder.fn_Config(offsetKey))
    maxColumns = VBA.CLng(ex_TravelDocumentsBuilder.fn_Config("candidates.layout.columns"))
    If columnOffset < 0 Or maxColumns < 1 Or columnOffset + rawTable.ColumnCount > maxColumns Then
        message = ex_TravelDocumentsBuilder.fn_Config("candidates.layout.overflowMessage")
        message = VBA.Replace$(message, "{columns}", VBA.CStr(rawTable.ColumnCount))
        message = VBA.Replace$(message, "{offset}", VBA.CStr(columnOffset))
        message = VBA.Replace$(message, "{available}", VBA.CStr(maxColumns - columnOffset))
        ex_WindowsUi.fn_ShowMessage message, VBA.vbExclamation
        private_SetActionStatus message, True
        Exit Function
    End If
    private_ValidateCandidateBounds = True
End Function

Private Function private_ApplySelectedPersonnel( _
    ByVal record As obj_LookupRecord, _
    ByRef diagnostic As String _
) As Boolean
    Dim isTvo As Variant
    Dim countText As String
    Dim field As String
    Dim target As String
    Dim value As Variant
    Dim updates As Object
    Dim index As Long

    If Not private_TryReadFormValue("LookupIsTvo", isTvo) Then Exit Function
    If Not VBA.CBool(isTvo) Then
        private_ApplySelectedPersonnel = m_lookupProfile.TryApply(record, m_pageBase.BindingContext, diagnostic)
        Exit Function
    End If
    If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("lookup.Tvo.apply.count", countText) Then Exit Function
    Set updates = VBA.CreateObject("Scripting.Dictionary")
    For index = 1 To VBA.CLng(countText)
        If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("lookup.Tvo.apply." & index & ".field", field) Then Exit Function
        If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("lookup.Tvo.apply." & index & ".target", target) Then Exit Function
        If Not record.TryGetValue(field, value) Then
            diagnostic = field
            Exit Function
        End If
        updates(target) = value
    Next index
    private_ApplySelectedPersonnel = m_pageBase.BindingContext.TryApplyValues(updates, diagnostic)
End Function

Private Function private_SaveEvent(ByVal isUpdate As Boolean) As Boolean
    Dim vacationGenerator As obj_TDB_VacationGenerator
    Dim diagnostic As String
    Dim operationId As String
    Dim key As Variant
    Dim value As Variant
    Dim objectValue As Object
    Dim isObject As Boolean
    Dim matches As Collection
    Dim orderRecord As obj_LookupRecord

    Static attemptNumber As Long

    If Not m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    On Error GoTo EH
    If VBA.Len(m_acceptedOrderNumber) = 0 Or m_acceptedOrderDate = 0 Then
        private_SetActionStatus ex_TDB_History.fn_Config("text.ChooseOrder"), True
        Exit Function
    End If
    If Not private_TryReadFormValue("OrderReference", value) Then
        Exit Function
    End If
    If Not m_orderResolver.TrySearch(VBA.CStr(value), VBA.vbNullString, matches, diagnostic) Then
        VBA.Err.Raise 5, , diagnostic
    End If
    If matches.Count <> 1 Then
        ex_TDB_History.fn_Fail "OrderStale"
    End If
    Set orderRecord = matches(1)
    If Not orderRecord.TryGetValue("OrderNumber", value) Then
        Exit Function
    End If
    If VBA.CStr(value) <> m_acceptedOrderNumber Then
        ex_TDB_History.fn_Fail "OrderStale"
    End If
    If Not orderRecord.TryGetValue("OrderDate", value) Then
        Exit Function
    End If
    If VBA.CDate(value) <> m_acceptedOrderDate Then
        ex_TDB_History.fn_Fail "OrderStale"
    End If
    If isUpdate Then
        If VBA.Len(m_acceptedOrderNumber) = 0 Or m_acceptedOrderNumber <> m_historyOrderNumber Or _
                m_acceptedOrderDate <> m_historyOrderDate Then ex_TDB_History.fn_Fail "OrderStale"
    End If
    attemptNumber = attemptNumber + 1
    operationId = VBA.Format$(VBA.Now, "yyyymmdd-hhnnss") & "-" & _
        VBA.CStr(VBA.CLng(VBA.Timer * 1000#)) & "-" & VBA.CStr(attemptNumber)
    ex_Core.fn_Diagnostic_WriteLog "TDB_ACTION | Operation=" & operationId & _
        " | Update=" & VBA.CStr(isUpdate) & " | Workbook=" & ex_TravelDocumentsBuilder.fn_LogValue(ThisWorkbook.FullName)
    For Each key In VBA.Array("DocumentType", "PersonName", "PersonId", "PersonRank", "PersonPositionCode", _
            "PersonPosition", "PersonUnit", "PersonnelQuery", "AutoFillPersonnel", "LookupIsTvo", _
            "DocumentTarget", "DepartureDate", "ReturnDate", "OrderReference", "Purpose", "ReportNo", "VacationKind", _
            "VacationDays", "RoadDays", "DonationDays", "VacationAbroad", "VacationStatus", _
            "TvoName", "TvoIpn", "TvoPosition", "TicketNo", "DocumentPath")
        If m_pageBase.BindingContext.TryGetValue("Form", VBA.CStr(key), value, objectValue, isObject) Then
            ex_Core.fn_Diagnostic_WriteLog "TDB_FORM_RAW | Operation=" & operationId & " | Field=" & key & _
                " | Type=" & VBA.TypeName(value) & " | Value=" & ex_TravelDocumentsBuilder.fn_LogValue(value)
        Else
            ex_Core.fn_Diagnostic_WriteLog "TDB_FORM_RAW | Operation=" & operationId & " | Field=" & key & " | Missing=True"
        End If
    Next key
    If Not ex_Core.fn_Diagnostic_Flush() Then
        ex_Core.fn_Diagnostic_WriteLog "TDB_LOG_FLUSH_FAILED | Operation=" & operationId & " | Phase=BeforeGeneration"
    End If
    Set vacationGenerator = New obj_TDB_VacationGenerator
    If vacationGenerator.Initialize(m_pageBase.BindingContext, m_lookupService) Then
        private_SaveEvent = vacationGenerator.TrySave(isUpdate, diagnostic, operationId)
    Else
        ex_Core.fn_Diagnostic_WriteLog "TDB_INITIALIZATION_FAILED | Operation=" & operationId
    End If
    vacationGenerator.Dispose
    If private_SaveEvent Then
        If private_TryReadFormValue("HistoryRowId", value) Then
            Set objectValue = ex_TDB_History.fn_Load(VBA.CStr(value), m_pageBase.BindingContext)
            private_PublishChain objectValue
            If Not private_PublishPlaces() Then Exit Function
            If private_TryReadFormValue("OrderReference", value) Then
                If Not m_orderResolver.TrySearch(VBA.CStr(value), VBA.vbNullString, matches, diagnostic) Then
                    VBA.Err.Raise 5, , diagnostic
                End If
                If matches.Count <> 1 Then
                    ex_TDB_History.fn_Fail "OrderStale"
                End If
                Set orderRecord = matches(1)
                If Not private_AcceptOrder(orderRecord, True) Then
                    Exit Function
                End If
            End If
            If VBA.IsEmpty(m_previousAutoFill) Then
                If Not private_TryReadFormValue("AutoFillPersonnel", m_previousAutoFill) Then
                    Exit Function
                End If
            End If
            m_pageBase.BindingContext.SetValue "Form", "AutoFillPersonnel", False
            m_pageBase.BindingContext.SetValue "Text", "SaveHistory", ex_TravelDocumentsBuilder.fn_Config("text.UpdateHistory")
            If Not private_TryReadFormValue("HistoryRowId", value) Then
                Exit Function
            End If
            m_pageBase.BindingContext.SetValue "Text", "HistoryEditMode", VBA.Replace$(ex_TDB_History.fn_Config("text.Editing"), "{id}", VBA.CStr(value))
            private_UpdateFieldState
        End If
    End If
    ex_Core.fn_Diagnostic_WriteLog "TDB_ACTION_FINISHED | Operation=" & operationId & _
        " | Success=" & VBA.CStr(private_SaveEvent) & " | Diagnostic=" & ex_TravelDocumentsBuilder.fn_LogValue(diagnostic)
    If Not ex_Core.fn_Diagnostic_Flush() Then
        ex_Core.fn_Diagnostic_WriteLog "TDB_LOG_FLUSH_FAILED | Operation=" & operationId & " | Phase=AfterGeneration"
    End If
    If Not private_SaveEvent And VBA.Len(diagnostic) > 0 Then
        m_pageBase.BindingContext.SetValue "Form", "GenerationStatusStyle", "generationError"
        m_pageBase.BindingContext.SetValue "Form", "GenerationStatus", diagnostic
    End If
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    m_pageBase.BindingContext.SetValue "Form", "GenerationStatusStyle", "generationError"
    m_pageBase.BindingContext.SetValue "Form", "GenerationStatus", diagnostic
    If Not vacationGenerator Is Nothing Then
        vacationGenerator.Dispose
    End If
    ex_Core.fn_Diagnostic_WriteLog "TDB_GENERATION_HANDLER_FAILED | Operation=" & operationId & _
        " | Diagnostic=" & ex_TravelDocumentsBuilder.fn_LogValue(diagnostic)
    If Not ex_Core.fn_Diagnostic_Flush() Then
        ex_Core.fn_Diagnostic_WriteLog "TDB_LOG_FLUSH_FAILED | Operation=" & operationId & " | Phase=HandlerError"
    End If
    private_SetActionStatus diagnostic, True
End Function

Private Function private_BuildPersonnelStatus( _
    ByVal searchText As String, _
    ByVal result As obj_LookupResult _
) As String
    If ex_TableSnapshots.fn_Active() Is Nothing Then
        private_BuildPersonnelStatus = ex_TravelDocumentsBuilder.fn_Config("text.DependenciesRequired")
        Exit Function
    End If
    If VBA.Len(m_acceptedOrderNumber) = 0 Or m_acceptedOrderDate = 0 Then
        private_BuildPersonnelStatus = ex_TDB_History.fn_Config("text.ChooseOrder")
        Exit Function
    End If
    If VBA.Len(ex_TableQuery.fn_NormalizeText(searchText)) < m_lookupProfile.MinChars Then
        private_BuildPersonnelStatus = VBA.Replace$( _
            m_messages("PersonnelMinimum"), "{minChars}", VBA.CStr(m_lookupProfile.MinChars))
    ElseIf result.RowCount = 0 Then
        private_BuildPersonnelStatus = m_messages("PersonnelEmpty")
    ElseIf result.HasMore Then
        private_BuildPersonnelStatus = VBA.Replace$(m_messages("PersonnelMore"), _
            "{count}", VBA.CStr(result.RowCount))
    Else
        private_BuildPersonnelStatus = VBA.Replace$(m_messages("PersonnelCount"), _
            "{count}", VBA.CStr(result.RowCount))
    End If
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