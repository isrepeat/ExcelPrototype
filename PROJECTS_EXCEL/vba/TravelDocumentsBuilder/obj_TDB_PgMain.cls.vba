VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_TDB_PgMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IPage

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_pageBase As obj_PageBase
Private m_controller As obj_TDB_PgMainController

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Interface
' //
Private Function obj_IPage_Initialize(ByVal profileId As String) As Boolean
    obj_IPage_Initialize = Me.Initialize(profileId)
End Function

Private Function obj_IPage_Render() As Boolean
    If m_pageBase Is Nothing Then Exit Function
    obj_IPage_Render = m_pageBase.Render()
End Function

Private Function obj_IPage_HandleCellChange(ByVal target As Range) As Boolean
    If target Is Nothing Then Exit Function
    obj_IPage_HandleCellChange = True
End Function

Private Sub obj_IPage_Dispose()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize(ByVal profileId As String) As Boolean
    If m_isDisposed Or m_isInitialized Then Exit Function

    Set m_pageBase = New obj_PageBase
    If Not m_pageBase.Initialize(profileId, ex_TravelDocumentsBuilder.fn_Config("ui.folder")) Then Exit Function
    If Not private_TryRegisterBindings() Then Exit Function
    Set m_controller = New obj_TDB_PgMainController
    If Not m_controller.Initialize(m_pageBase) Then Exit Function
    If Not private_TryRegisterCommands() Then Exit Function
    If Not m_controller.DocumentTypeChangedHandler() Then
        Exit Function
    End If
    Initialize = True
    m_isInitialized = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub

    m_isDisposed = True
    m_isInitialized = False
    If Not m_controller Is Nothing Then m_controller.Dispose
    Set m_controller = Nothing
    If Not m_pageBase Is Nothing Then m_pageBase.Dispose
    Set m_pageBase = Nothing
End Sub

' //
' // Private
' //
Private Function private_TryRegisterBindings() As Boolean
    Dim bindingContext As obj_UiBindingContext
    Dim key As Variant
    Dim valueText As String
    Dim defaultDocumentType As String
    Dim items As Variant
    Dim departureDates As Collection
    Dim daysAhead As Double
    Dim daysBefore As Double
    Dim offset As Long
    Dim chain As New Collection
    Dim placeChain As New Collection
    Dim orderOptions As New Collection
    Dim todayOrderOptions As New Collection
    Dim basisChain As New Collection

    If m_pageBase Is Nothing Then
        Exit Function
    End If
    Set bindingContext = m_pageBase.BindingContext
    If bindingContext Is Nothing Then
        Exit Function
    End If
    If Not bindingContext.SetObject("Data", "TvoChain", chain) Then
        Exit Function
    End If
    If Not bindingContext.SetObject("Data", "PlaceChain", placeChain) Then Exit Function
    If Not bindingContext.SetObject("Data", "BasisChain", basisChain) Then
        Exit Function
    End If
    For Each key In VBA.Split("BasisDraftId|" & ex_TravelDocumentsBuilder.fn_Config("basisChain.fields"), "|")
        If Not bindingContext.SetValue("Form", VBA.CStr(key), VBA.vbNullString) Then
            Exit Function
        End If
    Next key
    If Not bindingContext.SetValue("Form", "BasisTextVisibility", "Visible") Then
        Exit Function
    End If
    If Not bindingContext.SetValue("Form", "BasisInstitutionQuery", VBA.vbNullString) Then Exit Function
    If Not bindingContext.SetValue("Form", "BasisOwnText", VBA.vbNullString) Then Exit Function
    If Not bindingContext.SetValue("Form", "DeleteEventVisibility", "Collapsed") Then Exit Function
    If Not bindingContext.SetValue("Text", "DeleteEvent", ex_TravelDocumentsBuilder.fn_Config("text.DeleteEvent")) Then Exit Function
    If Not bindingContext.SetValue("Data", "BasisTypes", VBA.vbNullString) Then Exit Function
    If Not ex_TDB_BasisChain.fn_RefreshInstitutions(bindingContext) Then
        Exit Function
    End If
    For Each key In VBA.Split("BasisType|BasisInstitution|AddBasis|RemoveBasis|BasisTitle|" & ex_TravelDocumentsBuilder.fn_Config("basisChain.editorFields"), "|")
        If Not bindingContext.SetValue("Text", VBA.CStr(key), ex_TravelDocumentsBuilder.fn_Config("text." & key)) Then
            Exit Function
        End If
    Next key
    If Not bindingContext.SetValue("Text", "BasisInstitutionQuery", ex_TravelDocumentsBuilder.fn_Config("text.BasisInstitutionQuery")) Then Exit Function
    If Not bindingContext.SetValue("Form", "PlaceChainVisibility", "Collapsed") Then Exit Function
    If Not bindingContext.SetValue("Form", "TvoChainVisibility", "Collapsed") Then Exit Function
    If Not bindingContext.SetValue("Text", "TvoChainSummary", ex_TDB_History.fn_Config("text.ChainEmpty")) Then
        Exit Function
    End If
    If Not bindingContext.SetValue("Form", "CandidatesVisibility", "Visible") Then
        Exit Function
    End If
    If Not bindingContext.SetValue("Form", "HistoryVisibility", "Collapsed") Then
        Exit Function
    End If
    If Not bindingContext.SetValue("Data", "SelectedHistory", VBA.vbNullString) Then
        Exit Function
    End If
    If Not bindingContext.SetObject("Data", "OrderOptions", orderOptions) Then
        Exit Function
    End If
    If Not bindingContext.SetObject("Data", "TodayOrderOptions", todayOrderOptions) Then Exit Function
    If Not bindingContext.SetValue("Text", "HistoryEditMode", ex_TDB_History.fn_Config("text.NewMode")) Then
        Exit Function
    End If
    For Each key In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("forms.readOnly.fields"), "|")
        If Not bindingContext.SetValue("Form", VBA.CStr(key) & "ReadOnly", False) Then
            Exit Function
        End If
        If Not bindingContext.SetValue("Form", VBA.CStr(key) & "Style", ex_TravelDocumentsBuilder.fn_Config("forms.sourceStyles.default")) Then
            Exit Function
        End If
        If Not bindingContext.SetValue("Form", VBA.CStr(key) & "LabelStyle", ex_TravelDocumentsBuilder.fn_Config("forms.sourceStyles.label")) Then
            Exit Function
        End If
    Next key
    valueText = ex_TravelDocumentsBuilder.fn_Config("candidates.selectionMode")
    If valueText <> "Focus" And valueText <> "Confirm" Then
        VBA.Err.Raise 5, , "Invalid configuration key: TravelDocumentsBuilder::candidates.selectionMode"
    End If
    If valueText = "Confirm" Then
        If Not ex_Core.fn_TryGetWorkbookConfigValue("ThisWorkbook::ui.confirmSelectionKey", valueText) Then
            VBA.Err.Raise 5, , "Required configuration key not found: ThisWorkbook::ui.confirmSelectionKey"
        End If
        If VBA.Len(VBA.Trim$(valueText)) = 0 Then
            VBA.Err.Raise 5, , "Empty confirmation hotkey."
        End If
    End If
    If Not bindingContext.SetValue("Settings", "CandidateSelectionMode", ex_TravelDocumentsBuilder.fn_Config("candidates.selectionMode")) Then
        Exit Function
    End If
    If Not bindingContext.SetValue("Settings", "CandidateColumnOffset", ex_TravelDocumentsBuilder.fn_Config("candidates.personnel.columnOffset")) Then Exit Function
    If Not bindingContext.SetValue("Settings", "CandidateMaxColumns", ex_TravelDocumentsBuilder.fn_Config("candidates.layout.columns")) Then Exit Function
    If Not bindingContext.SetValue("Text", "CandidateOverflow", ex_TravelDocumentsBuilder.fn_Config("candidates.layout.overflowMessage")) Then Exit Function
    If Not bindingContext.SetValue("Form", "LookupIsTvo", False) Then
        Exit Function
    End If
    If Not bindingContext.SetValue("Form", "GenerationStatusStyle", "generationSuccess") Then
        Exit Function
    End If
    If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("dates.daysAhead", valueText) Then
        Exit Function
    End If
    If Not ex_Helpers.fn_TryParseInvariantNumber(valueText, daysAhead) Then
        Exit Function
    End If
    If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("dates.daysBefore", valueText) Then
        Exit Function
    End If
    If Not ex_Helpers.fn_TryParseInvariantNumber(valueText, daysBefore) Then
        Exit Function
    End If
    If daysAhead < 0 Or daysAhead > 31 Or daysAhead <> VBA.Fix(daysAhead) Or _
            daysBefore < 0 Or daysBefore > 31 Or daysBefore <> VBA.Fix(daysBefore) Then
        If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("text.InvalidDateRange", valueText) Then
            Exit Function
        End If
        ex_WindowsUi.fn_ShowMessage valueText, VBA.vbExclamation
        Exit Function
    End If
    Set departureDates = New Collection
    For offset = -VBA.CLng(daysBefore) To VBA.CLng(daysAhead)
        departureDates.Add VBA.Format$(VBA.DateAdd("d", offset, VBA.Date), "dd.mm.yyyy")
    Next offset
    If Not bindingContext.SetObject("Data", "DepartureDates", departureDates) Then
        Exit Function
    End If
    If Not bindingContext.SetValue("Data", "CurrentDate", VBA.Format$(VBA.Date, "dd.mm.yyyy")) Then
        Exit Function
    End If

    If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("documentTypes.default", defaultDocumentType) Then
        Exit Function
    End If
    If Not bindingContext.SetValue("Form", "DocumentType", defaultDocumentType) Then
        Exit Function
    End If
    If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("documentTypes.items", valueText) Then
        Exit Function
    End If
    If Not bindingContext.SetValue("Data", "DocumentTypes", valueText) Then
        Exit Function
    End If
    For Each key In VBA.Array("PersonName", "PersonId", "PersonRank", "PersonPosition", _
            "PersonUnit", "PersonPositionCode", "PersonnelQuery", "DocumentTarget", "DepartureDate", "ReturnDate", "Purpose", _
            "OrderReference", "ReportNo", "TvoName", "TvoIpn", "TvoPosition", "TicketNo", "DocumentPath", "GenerationStatus", _
            "WeaponLine", "PlaceRegion", "PlaceType", "PlaceName", "PlaceDistrict", _
            "PlaceCountry", "PlaceCountryEn", "PlaceRegionEn", "PlaceNameEn", _
            "VlkNo", "VlkDate", "LongTreatmentDays", "TreatmentLeaveDays", _
            "HistoryRowId", "HistoryStamp", "ChangeReason", "OrderQuery", "OrderYear", "AcceptedOrderNumber", "AcceptedOrderDate")
        If Not bindingContext.SetValue("Form", VBA.CStr(key), VBA.vbNullString) Then
            Exit Function
        End If
    Next key
    If Not bindingContext.SetValue("Form", "PlaceAbroad", VBA.vbNullString) Then Exit Function
    If Not bindingContext.SetValue("Form", "DomesticPlaceFieldsVisibility", "Collapsed") Then Exit Function
    If Not bindingContext.SetValue("Form", "ForeignPlaceFieldsVisibility", "Collapsed") Then Exit Function
    For Each key In VBA.Array("PlaceKind", "VacationKind", "VacationAbroad", "VacationStatus", "VacationDays", _
            "RoadDays", "DonationDays", "TripKind")
        If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("generation.defaults." & VBA.CStr(key), valueText) Then
            Exit Function
        End If
        If Not bindingContext.SetValue("Form", VBA.CStr(key), valueText) Then
            Exit Function
        End If
    Next key
    If Not bindingContext.SetValue("Form", "ChangeReasonVisibility", "Collapsed") Then Exit Function
    If Not bindingContext.SetValue("Form", "PlaceFieldsVisibility", "Collapsed") Then Exit Function
    If Not bindingContext.SetValue("Form", "TreatmentVisibility", "Collapsed") Then Exit Function
    If Not bindingContext.SetValue("Form", "DestinationVisibility", "Visible") Then Exit Function
    If Not bindingContext.SetValue("Form", "VacationVisibility", "Visible") Then
        Exit Function
    End If
    If Not bindingContext.SetValue("Form", "VacationDaysReadOnly", False) Then Exit Function
    If Not bindingContext.SetValue("Form", "TripEditableVisibility", "Visible") Then Exit Function
    If Not bindingContext.SetValue("Form", "TripVisibility", "Collapsed") Then
        Exit Function
    End If
    For Each key In VBA.Array("PlaceKinds", "VacationKinds", "AbroadOptions", "StatusOptions", "TripKinds")
        If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("generation.items." & VBA.CStr(key), valueText) Then
            Exit Function
        End If
        If VBA.CStr(key) = "VacationKinds" Then
            items = VBA.Split(valueText, "|")
            If Not bindingContext.SetValue("Data", VBA.CStr(key), items) Then
                Exit Function
            End If
        Else
            If Not bindingContext.SetValue("Data", VBA.CStr(key), valueText) Then
                Exit Function
            End If
        End If
    Next key
    For Each key In VBA.Array("WeaponLine", "PlaceKind", "PlaceRegion", "PlaceType", "PlaceName", "PlaceDistrict", _
            "PlaceCountry", "PlaceCountryEn", "PlaceRegionEn", "PlaceNameEn", _
            "VacationKind", "TripKind", "OrderReference", "VacationDays", "RoadDays", "DonationDays", _
            "VlkNo", "VlkDate", "LongTreatmentDays", "TreatmentLeaveDays", _
            "VacationAbroad", "VacationStatus", "TvoName", "TvoIpn", "TvoPosition", _
            "SaveHistory", "ExportWord", "ClearForm", "ReportNo", "TicketNo", "LastStatus", _
            "LoadHistory", "AddTvo", "RemoveTvo", "AddPlace", "RemovePlace", "PlaceChainTitle", _
            "ChangeReason", "TvoChainTitle", "OrderQueryLabel", "OrderYearLabel", "CloseHistory")
        If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("text." & VBA.CStr(key), valueText) Then
            Exit Function
        End If
        If Not bindingContext.SetValue("Text", VBA.CStr(key), valueText) Then
            Exit Function
        End If
    Next key
    If Not bindingContext.SetValue("Data", "SelectedPersonnel", VBA.vbNullString) Then
        Exit Function
    End If
    If Not bindingContext.SetValue("Form", "AutoFillPersonnel", True) Then
        Exit Function
    End If
    For Each key In VBA.Array("DocumentTypeLabel", "PersonNameLabel", _
            "PersonIdLabel", "PersonRankLabel", "PersonPositionLabel", "PersonUnitLabel", _
            "RefreshPersonnel", "AutoFillPersonnelLabel", "Title", "PersonnelSection", _
            "DocumentSection", "SearchSection", "PersonPositionCodeLabel", _
            "CheckboxChecked", "CheckboxUnchecked", "ActionsSection", "ToggleTheme", "ReloadPage")
        If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("text." & VBA.CStr(key), valueText) Then
            Exit Function
        End If
        If Not bindingContext.SetValue("Text", VBA.CStr(key), valueText) Then
            Exit Function
        End If
    Next key
    If Not ex_Core.fn_TryGetWorkbookConfigValue( _
            "TravelDocumentsBuilder::text.PersonnelPrompt", valueText) Then Exit Function
    If Not bindingContext.SetValue("Form", "GenerationStatus", valueText) Then
        Exit Function
    End If
    If Not bindingContext.SetValue("Settings", "Theme", ex_TravelDocumentsBuilder.fn_Config("ui.theme")) Then Exit Function
    private_TryRegisterBindings = private_ApplyDocumentMode(defaultDocumentType)
End Function

Private Function private_TryRegisterCommands() As Boolean
    Dim bindingContext As obj_UiBindingContext

    If m_pageBase Is Nothing Or m_controller Is Nothing Then Exit Function
    Set bindingContext = m_pageBase.BindingContext
    If bindingContext Is Nothing Then Exit Function
    If Not private_TryRegisterCommand(bindingContext, "ToggleTheme", "ToggleThemeHandler") Then Exit Function
    If Not private_TryRegisterCommand(bindingContext, "ReloadPage", "ReloadPageHandler") Then Exit Function
    If Not private_TryRegisterCommand( _
            bindingContext, "PersonNameChanged", "PersonNameChangedHandler") Then Exit Function
    If Not private_TryRegisterCommand( _
            bindingContext, "SelectPersonnel", "SelectPersonnelHandler") Then Exit Function
    If Not private_TryRegisterCommand( _
            bindingContext, "RefreshPersonnel", "RefreshPersonnelHandler") Then Exit Function
    If Not private_TryRegisterCommand( _
            bindingContext, "DocumentTypeChanged", "DocumentTypeChangedHandler") Then Exit Function
    If Not private_TryRegisterCommand( _
            bindingContext, "TripKindChanged", "TripKindChangedHandler") Then Exit Function
    If Not private_TryRegisterCommand( _
            bindingContext, "SaveHistory", "SaveHistoryHandler") Then Exit Function
    If Not private_TryRegisterCommand( _
            bindingContext, "ExportWord", "ExportWordHandler") Then Exit Function
    If Not private_TryRegisterCommand( _
            bindingContext, "ClearForm", "ClearFormHandler") Then Exit Function
    If Not private_TryRegisterCommand(bindingContext, "LoadHistory", "LoadHistoryHandler") Then Exit Function
    If Not private_TryRegisterCommand(bindingContext, "SelectHistory", "SelectHistoryHandler") Then Exit Function
    If Not private_TryRegisterCommand(bindingContext, "AddTvo", "AddTvoHandler") Then Exit Function
    If Not private_TryRegisterCommand(bindingContext, "RemoveTvo", "RemoveTvoHandler") Then Exit Function
    If Not private_TryRegisterCommand(bindingContext, "SearchOrders", "SearchOrdersHandler") Then Exit Function
    If Not private_TryRegisterCommand(bindingContext, "AutoFillChanged", "AutoFillChangedHandler") Then Exit Function
    If Not private_TryRegisterCommand(bindingContext, "CloseHistory", "CloseHistoryHandler") Then Exit Function
    If Not private_TryRegisterCommand(bindingContext, "TvoNameChanged", "TvoNameChangedHandler") Then Exit Function
    If Not private_TryRegisterCommand(bindingContext, "VacationKindChanged", "VacationKindChangedHandler") Then Exit Function
    If Not private_TryRegisterCommand(bindingContext, "PlaceChanged", "PlaceChangedHandler") Then Exit Function
    If Not private_TryRegisterCommand(bindingContext, "PlaceKindChanged", "VacationAbroadChangedHandler") Then Exit Function
    If Not private_TryRegisterCommand(bindingContext, "VacationAbroadChanged", "VacationAbroadChangedHandler") Then Exit Function
    If Not private_TryRegisterCommand(bindingContext, "AddPlace", "AddPlaceHandler") Then Exit Function
    If Not private_TryRegisterCommand(bindingContext, "BasisInstitutionChanged", "BasisInstitutionChangedHandler") Then
        Exit Function
    End If
    If Not private_TryRegisterCommand(bindingContext, "BasisInstitutionQueryChanged", "BasisInstitutionQueryChangedHandler") Then Exit Function
    If Not private_TryRegisterCommand(bindingContext, "BasisTypeChanged", "BasisTypeChangedHandler") Then
        Exit Function
    End If
    If Not private_TryRegisterCommand(bindingContext, "BasisDraftChanged", "BasisDraftChangedHandler") Then
        Exit Function
    End If
    If Not private_TryRegisterCommand(bindingContext, "DeleteEvent", "DeleteEventHandler") Then Exit Function
    If Not private_TryRegisterCommand(bindingContext, "AddBasis", "AddBasisHandler") Then
        Exit Function
    End If
    If Not private_TryRegisterCommand(bindingContext, "RemoveBasis", "RemoveBasisHandler") Then
        Exit Function
    End If
    If Not private_TryRegisterCommand(bindingContext, "RemovePlace", "RemovePlaceHandler") Then Exit Function
    private_TryRegisterCommands = True
End Function

Private Function private_TryRegisterCommand( _
    ByVal bindingContext As obj_UiBindingContext, _
    ByVal commandKey As String, _
    ByVal methodName As String _
) As Boolean
    Dim command As obj_UiCommand

    If bindingContext Is Nothing Or m_controller Is Nothing Then Exit Function
    Set command = New obj_UiCommand
    If Not command.Initialize(m_controller, methodName) Then Exit Function
    private_TryRegisterCommand = bindingContext.SetObject("Commands", commandKey, command)
End Function

Private Function private_ApplyDocumentMode(ByVal documentType As String) As Boolean
    Dim bindingContext As obj_UiBindingContext
    Dim prefix As String
    Dim key As Variant
    Dim valueText As String

    If m_pageBase Is Nothing Then Exit Function
    Set bindingContext = m_pageBase.BindingContext
    If bindingContext Is Nothing Then Exit Function

    If Not ex_TravelDocumentsBuilder.fn_TryGetDocumentPrefix(documentType, prefix) Then Exit Function
    For Each key In VBA.Array("DocumentFormTitle", "DocumentTargetLabel", _
            "DepartureDateLabel", "ReturnDateLabel", "PurposeLabel")
        If Not ex_Core.fn_TryGetWorkbookConfigValue( _
                "TravelDocumentsBuilder::text." & prefix & "." & VBA.CStr(key), valueText) Then Exit Function
        If Not bindingContext.SetValue("Text", VBA.CStr(key), valueText) Then Exit Function
    Next key
    private_ApplyDocumentMode = True
End Function