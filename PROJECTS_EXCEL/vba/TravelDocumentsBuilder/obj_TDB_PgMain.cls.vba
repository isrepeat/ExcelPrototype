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
    If m_isDisposed Or m_isInitialized Then Exit Function

    Set m_pageBase = New obj_PageBase
    If Not m_pageBase.Initialize(profileId, VBA.Trim$(profileId)) Then Exit Function
    If Not private_TryRegisterBindings() Then Exit Function
    Set m_controller = New obj_TDB_PgMainController
    If Not m_controller.Initialize(m_pageBase) Then Exit Function
    If Not private_TryRegisterCommands() Then Exit Function
    obj_IPage_Initialize = True
    m_isInitialized = True
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

Public Sub Dispose()
    If m_isDisposed Then Exit Sub

    m_isDisposed = True
    m_isInitialized = False
    If Not m_controller Is Nothing Then m_controller.Dispose
    Set m_controller = Nothing
    If Not m_pageBase Is Nothing Then m_pageBase.Dispose
    Set m_pageBase = Nothing
End Sub

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

    If m_pageBase Is Nothing Then Exit Function
    Set bindingContext = m_pageBase.BindingContext
    If bindingContext Is Nothing Then Exit Function
    If Not bindingContext.SetValue("Form", "LookupIsTvo", False) Then Exit Function
    If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("text.LookupTarget.Personnel", valueText) Then Exit Function
    If Not bindingContext.SetValue("Text", "LookupTargetCaption", valueText) Then Exit Function
    If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("text.ClearTvo", valueText) Then Exit Function
    If Not bindingContext.SetValue("Text", "ClearTvo", valueText) Then Exit Function
    If Not bindingContext.SetValue("Form", "GenerationStatusStyle", "generationSuccess") Then Exit Function
    If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("dates.daysAhead", valueText) Then Exit Function
    If Not ex_Helpers.fn_TryParseInvariantNumber(valueText, daysAhead) Then Exit Function
    If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("dates.daysBefore", valueText) Then Exit Function
    If Not ex_Helpers.fn_TryParseInvariantNumber(valueText, daysBefore) Then Exit Function
    If daysAhead < 0 Or daysAhead > 31 Or daysAhead <> VBA.Fix(daysAhead) Or _
            daysBefore < 0 Or daysBefore > 31 Or daysBefore <> VBA.Fix(daysBefore) Then
        If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("text.InvalidDateRange", valueText) Then Exit Function
        ex_WindowsUi.fn_ShowMessage valueText, VBA.vbExclamation
        Exit Function
    End If
    Set departureDates = New Collection
    For offset = -VBA.CLng(daysBefore) To VBA.CLng(daysAhead)
        departureDates.Add VBA.Format$(VBA.DateAdd("d", offset, VBA.Date), "dd.mm.yyyy")
    Next offset
    If Not bindingContext.SetObject("Data", "DepartureDates", departureDates) Then Exit Function
    If Not bindingContext.SetValue("Data", "CurrentDate", VBA.Format$(VBA.Date, "dd.mm.yyyy")) Then Exit Function

    If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("documentTypes.default", defaultDocumentType) Then Exit Function
    If Not bindingContext.SetValue("Form", "DocumentType", defaultDocumentType) Then Exit Function
    If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("documentTypes.items", valueText) Then Exit Function
    If Not bindingContext.SetValue("Data", "DocumentTypes", valueText) Then Exit Function
    For Each key In VBA.Array("PersonName", "PersonId", "PersonRank", "PersonPosition", _
            "PersonUnit", "PersonPositionCode", "PersonnelQuery", "DocumentTarget", "DepartureDate", "ReturnDate", "Purpose", _
            "OrderReference", "TvoName", "TvoIpn", "TvoPosition", "TicketNo", "DocumentPath", "GenerationStatus")
        If Not bindingContext.SetValue("Form", VBA.CStr(key), VBA.vbNullString) Then Exit Function
    Next key
    For Each key In VBA.Array("VacationKind", "VacationAbroad", "VacationStatus", "VacationDays", _
            "RoadDays", "DonationDays", "TemplatePath", "OutputFolderPath")
        If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("generation.defaults." & VBA.CStr(key), valueText) Then Exit Function
        If Not bindingContext.SetValue("Form", VBA.CStr(key), valueText) Then Exit Function
    Next key
    For Each key In VBA.Array("VacationKinds", "AbroadOptions", "StatusOptions")
        If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("generation.items." & VBA.CStr(key), valueText) Then Exit Function
        If VBA.CStr(key) = "VacationKinds" Then
            items = VBA.Split(valueText, "|")
            If Not bindingContext.SetValue("Data", VBA.CStr(key), items) Then Exit Function
        Else
            If Not bindingContext.SetValue("Data", VBA.CStr(key), valueText) Then Exit Function
        End If
    Next key
    For Each key In VBA.Array("VacationKind", "OrderReference", "VacationDays", "RoadDays", "DonationDays", _
            "VacationAbroad", "VacationStatus", "TvoName", "TvoIpn", "TvoPosition", "TemplatePath", "OutputFolderPath", _
            "CreateDocument", "UpdateDocument", "ClearForm", "TicketNo", "DocumentPath")
        If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("text." & VBA.CStr(key), valueText) Then Exit Function
        If Not bindingContext.SetValue("Text", VBA.CStr(key), valueText) Then Exit Function
    Next key
    If Not bindingContext.SetValue("Data", "SelectedPersonnel", VBA.vbNullString) Then Exit Function
    If Not bindingContext.SetValue("Form", "AutoFillPersonnel", True) Then Exit Function
    For Each key In VBA.Array("DocumentTypeLabel", "PersonnelPrompt", "PersonNameLabel", _
            "PersonIdLabel", "PersonRankLabel", "PersonPositionLabel", "PersonUnitLabel", _
            "RefreshPersonnel", "AutoFillPersonnelLabel", "Title", "PersonnelSection", _
            "DocumentSection", "SearchSection", "SearchPersonnel", "SearchLabel", "PersonPositionCodeLabel", _
            "CheckboxChecked", "CheckboxUnchecked", "ActionsSection")
        If Not ex_TravelDocumentsBuilder.fn_TryGetConfig("text." & VBA.CStr(key), valueText) Then Exit Function
        If Not bindingContext.SetValue("Text", VBA.CStr(key), valueText) Then Exit Function
    Next key
    If Not ex_Core.fn_TryGetWorkbookConfigValue( _
            "TravelDocumentsBuilder::text.PersonnelPrompt", valueText) Then Exit Function
    If Not bindingContext.SetValue("Text", "PersonnelStatus", valueText) Then Exit Function
    private_TryRegisterBindings = private_ApplyDocumentMode(defaultDocumentType)
End Function

Private Function private_TryRegisterCommands() As Boolean
    Dim bindingContext As obj_UiBindingContext

    If m_pageBase Is Nothing Or m_controller Is Nothing Then Exit Function
    Set bindingContext = m_pageBase.BindingContext
    If bindingContext Is Nothing Then Exit Function
    If Not private_TryRegisterCommand( _
            bindingContext, "SearchPersonnel", "SearchPersonnelHandler") Then Exit Function
    If Not private_TryRegisterCommand( _
            bindingContext, "PersonNameChanged", "PersonNameChangedHandler") Then Exit Function
    If Not private_TryRegisterCommand( _
            bindingContext, "SelectPersonnel", "SelectPersonnelHandler") Then Exit Function
    If Not private_TryRegisterCommand( _
            bindingContext, "RefreshPersonnel", "RefreshPersonnelHandler") Then Exit Function
    If Not private_TryRegisterCommand( _
            bindingContext, "DocumentTypeChanged", "DocumentTypeChangedHandler") Then Exit Function
    If Not private_TryRegisterCommand( _
            bindingContext, "CreateDocument", "CreateDocumentHandler") Then Exit Function
    If Not private_TryRegisterCommand( _
            bindingContext, "UpdateDocument", "UpdateDocumentHandler") Then Exit Function
    If Not private_TryRegisterCommand( _
            bindingContext, "ClearForm", "ClearFormHandler") Then Exit Function
    If Not private_TryRegisterCommand( _
            bindingContext, "VacationDatesChanged", "VacationDatesChangedHandler") Then Exit Function
    If Not private_TryRegisterCommand(bindingContext, "ToggleLookupTarget", "ToggleLookupTargetHandler") Then Exit Function
    If Not private_TryRegisterCommand(bindingContext, "FocusTvo", "FocusTvoHandler") Then Exit Function
    If Not private_TryRegisterCommand(bindingContext, "FocusPersonnel", "FocusPersonnelHandler") Then Exit Function
    If Not private_TryRegisterCommand(bindingContext, "ClearTvo", "ClearTvoHandler") Then Exit Function
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