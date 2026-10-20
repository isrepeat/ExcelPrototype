VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PEB_PgMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IPage
Implements obj_IUiBindingListener

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_pageBase As obj_PageBase
Private m_controller As obj_PEB_PgMainController
Private m_fieldDefaultStyles As Object
Private m_dynamicStylesEnabled As Boolean

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
    Dim diagnostic As String

    If m_pageBase Is Nothing Then Exit Function
    If ex_PEB_Testing.fn_IsRunning() Then
        ex_WindowsUi.fn_ShowMessage ex_PEB_Export.fn_Config("text.TestsBusy"), vbExclamation, ex_PEB_Export.fn_Config("text.TestTitle")
        Exit Function
    End If
    If Not m_controller.TryRefreshSourceLegend(diagnostic) Then
        ex_WindowsUi.fn_ShowMessage diagnostic, vbExclamation, "Source legend"
        Exit Function
    End If
    obj_IPage_Render = m_pageBase.Render()
End Function

Private Function obj_IPage_HandleCellChange(ByVal target As Range) As Boolean
    If target Is Nothing Then Exit Function
    obj_IPage_HandleCellChange = True
End Function

Private Sub obj_IPage_Dispose()
    Me.Dispose
End Sub

Private Sub obj_IUiBindingListener_OnValueChanged( _
    ByVal sourceName As String, _
    ByVal bindingPath As String _
)
    Dim defaultStyle As String
    Dim value As Variant
    Dim valueObject As Object
    Dim isObject As Boolean

    If m_isDisposed Or m_pageBase Is Nothing Then Exit Sub
    If sourceName <> "Form" Then Exit Sub
    If Not m_dynamicStylesEnabled Then Exit Sub
    If m_fieldDefaultStyles Is Nothing Then Exit Sub
    If Not m_fieldDefaultStyles.Exists(bindingPath) Then Exit Sub
    If Not m_pageBase.BindingContext.TryGetValue(sourceName, bindingPath, value, valueObject, isObject) Then Exit Sub
    If isObject Then Exit Sub
    If VBA.IsNull(value) Or VBA.IsEmpty(value) Then
        value = VBA.vbNullString
    End If
    If VBA.Len(VBA.Trim$(VBA.CStr(value))) > 0 Then Exit Sub
    defaultStyle = m_fieldDefaultStyles(bindingPath)
    If Not m_pageBase.BindingContext.SetValue("Form", bindingPath & "Style", defaultStyle) Then
        Err.Raise 5, , "Cannot reset form source style: " & bindingPath
    End If
End Sub

' //
' // API
' //
Public Function Initialize(ByVal profileId As String) As Boolean
    Dim uiFolder As String

    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    Set m_pageBase = New obj_PageBase
    uiFolder = ex_PEB_Export.fn_Config("ui.folder")
    If Not m_pageBase.Initialize(profileId, uiFolder) Then
        Exit Function
    End If
    If Not private_TryRegisterBindings() Then
        Exit Function
    End If
    m_pageBase.BindingContext.Subscribe Me
    Set m_controller = New obj_PEB_PgMainController
    If Not m_controller.Initialize(m_pageBase) Then
        Exit Function
    End If
    If Not private_TryRegisterCommands() Then
        Exit Function
    End If
    Initialize = True
    m_isInitialized = Initialize
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    If Not m_pageBase Is Nothing Then m_pageBase.BindingContext.Unsubscribe Me
    Set m_fieldDefaultStyles = Nothing
    If Not m_controller Is Nothing Then m_controller.Dispose
    Set m_controller = Nothing
    If Not m_pageBase Is Nothing Then m_pageBase.Dispose
    Set m_pageBase = Nothing
End Sub

' //
' // Private
' //
Private Function private_TryRegisterBindings() As Boolean
    Dim uiBindingContext As obj_UiBindingContext
    Dim textKey As Variant
    Dim configuredText As String
    Dim confirmSelectionKey As String
    Dim orderOptions As Collection
    Dim wordContext As Object
    Dim fieldName As Variant
    Dim defaultStyle As String
    Dim ordinaryStyle As String
    Dim lookupStyle As String
    Dim lookupFields As String

    If m_pageBase Is Nothing Then Exit Function
    Set uiBindingContext = m_pageBase.BindingContext
    If uiBindingContext Is Nothing Then Exit Function
    configuredText = ex_PEB_Export.fn_Config("forms.dynamicStylesEnabled")
    If configuredText <> "true" And configuredText <> "false" Then Err.Raise 5, , "Invalid forms.dynamicStylesEnabled."
    m_dynamicStylesEnabled = (configuredText = "true")
    If Not uiBindingContext.SetValue("Settings", "CandidateTextWarningLimit", ex_PEB_Export.fn_Config("candidates.textWarningLimit")) Then Exit Function
    If Not uiBindingContext.SetValue("Settings", "CandidateTextWarningColumn", ex_PEB_Export.fn_Config("candidates.textWarningColumn")) Then Exit Function
    Set m_fieldDefaultStyles = VBA.CreateObject("Scripting.Dictionary")
    ordinaryStyle = ex_PEB_Export.fn_Config("forms.sourceStyles.default")
    lookupStyle = ex_PEB_Export.fn_Config("forms.lookupDefaultStyle")
    lookupFields = "|" & ex_PEB_Export.fn_Config("forms.lookupFields") & "|"
    For Each fieldName In VBA.Split(ex_PEB_Export.fn_Config("forms.styledFields"), "|")
        defaultStyle = ordinaryStyle
        If VBA.InStr(1, lookupFields, "|" & VBA.CStr(fieldName) & "|", VBA.vbBinaryCompare) > 0 Then
            defaultStyle = lookupStyle
        End If
        m_fieldDefaultStyles.Add VBA.CStr(fieldName), defaultStyle
        If Not uiBindingContext.SetValue("Form", VBA.CStr(fieldName) & "Style", defaultStyle) Then Exit Function
    Next fieldName
    Set wordContext = VBA.CreateObject("Scripting.Dictionary")
    wordContext.CompareMode = VBA.vbBinaryCompare
    If Not uiBindingContext.SetObject("Data", "WordContext", wordContext) Then Exit Function
    If Not ex_Core.fn_TryGetWorkbookConfigValue("PersonalEventBuilder::candidates.selectionMode", configuredText) Then
        ex_WindowsUi.fn_ShowMessage "Required configuration key not found: PersonalEventBuilder::candidates.selectionMode", vbExclamation, "Configuration"
        Exit Function
    End If
    If configuredText <> "Focus" And configuredText <> "Confirm" Then
        ex_WindowsUi.fn_ShowMessage "Invalid configuration value: PersonalEventBuilder::candidates.selectionMode", vbExclamation, "Configuration"
        Exit Function
    End If
    If configuredText = "Confirm" Then
        If Not ex_Core.fn_TryGetWorkbookConfigValue("ThisWorkbook::ui.confirmSelectionKey", confirmSelectionKey) Then
            ex_WindowsUi.fn_ShowMessage "Required configuration key not found: ThisWorkbook::ui.confirmSelectionKey", vbExclamation, "Configuration"
            Exit Function
        End If
        If VBA.Len(VBA.Trim$(confirmSelectionKey)) = 0 Then
            ex_WindowsUi.fn_ShowMessage "Empty configuration key: ThisWorkbook::ui.confirmSelectionKey", vbExclamation, "Configuration"
            Exit Function
        End If
    End If
    If Not uiBindingContext.SetValue("Settings", "CandidateSelectionMode", configuredText) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "AbsenceType", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "WordPreviewText", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "HistoryVisibility", "Collapsed") Then Exit Function
    If Not uiBindingContext.SetValue("Form", "ConflictsVisibility", "Collapsed") Then Exit Function
    If Not uiBindingContext.SetValue("Form", "ConflictColumn", 0) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "CandidatesVisibility", "Collapsed") Then Exit Function
    If Not uiBindingContext.SetValue("Form", "OrderQuery", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "OrderYear", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "OrderChoice", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "AcceptedOrderNumber", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "AcceptedOrderDate", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "ActionStatusStyle", "actionSuccess") Then Exit Function
    If Not ex_Core.fn_TryGetWorkbookConfigValue("PersonalEventBuilder::personnel.autoFill.default", configuredText) Then
        ex_WindowsUi.fn_ShowMessage "Required configuration key not found: PersonalEventBuilder::personnel.autoFill.default", vbExclamation, "Configuration"
        Exit Function
    End If
    If configuredText <> "true" And configuredText <> "false" Then
        ex_WindowsUi.fn_ShowMessage "Invalid configuration value: PersonalEventBuilder::personnel.autoFill.default", vbExclamation, "Configuration"
        Exit Function
    End If
    If Not uiBindingContext.SetValue("Form", "AutoFillPersonnel", configuredText = "true") Then Exit Function
    Set orderOptions = New Collection
    If Not uiBindingContext.SetObject("Data", "OrderOptions", orderOptions) Then Exit Function
    Set orderOptions = New Collection
    If Not uiBindingContext.SetObject("Data", "TodayOrderOptions", orderOptions) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "Destination", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "PersonRank", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "PersonName", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "PersonId", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "PositionCode", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "PersonPosition", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "DepartureDate", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "TicketNumber", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "TicketDate", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Data", "SelectedPersonnel", vbNullString) Then Exit Function
    For Each textKey In VBA.Array("PersonNameLabel", "PersonIdLabel", "PersonRankLabel", _
            "PersonPositionLabel", "PositionCodeLabel", "AbsenceTypeLabel", "DestinationLabel", _
            "DepartureDateLabel", "TicketNumberLabel", "TicketDateLabel", _
            "OrderQueryLabel", "OrderYearLabel", "LastStatus", _
            "PersonnelSection", "ActionsSection", "RefreshPersonnel", _
            "AutoFillPersonnelLabel", "CheckboxChecked", "CheckboxUnchecked", _
            "SourceLegendSample", "SourceLegendSeparator", "SourceLegendOrderDate", "ResetForm", _
            "DummyMode", "ExtendedModes", "EnvironmentSwitch", "ReinitializePage", _
            "LoadTest", "NextTest", _
            "PreviewWord", "SaveHistory", "ShowHistory", "HistoryEditMode", "CloseTable", "ExportWord", "ExportMovement", _
            "EventVersions", "CancelEvent", "RestoreEvent")
        If Not ex_Core.fn_TryGetWorkbookConfigValue("PersonalEventBuilder::text." & textKey, configuredText) Then
            ex_WindowsUi.fn_ShowMessage "Required configuration text not found: PersonalEventBuilder::text." & textKey, vbExclamation, "Configuration"
            Exit Function
        End If
        If Not uiBindingContext.SetValue("Text", VBA.CStr(textKey), configuredText) Then Exit Function
    Next textKey
    configuredText = ex_PEB_Export.fn_Config("environment.name")
    If configuredText <> "TEST" And configuredText <> "PROD" Then
        ex_WindowsUi.fn_ShowMessage "Invalid environment.name: " & configuredText, vbExclamation, "Configuration"
        Exit Function
    End If
    If configuredText = "TEST" Then
        If Not uiBindingContext.SetValue("Text", "TestScenarioItems", ex_PEB_Export.fn_Config("text.TestScenarioItems")) Then Exit Function
    Else
        If Not uiBindingContext.SetValue("Text", "TestScenarioItems", VBA.vbNullString) Then Exit Function
    End If
    If Not uiBindingContext.SetValue("Settings", "TestMode", configuredText = "TEST") Then Exit Function
    If Not uiBindingContext.SetValue("Settings", "TestsRunning", False) Then Exit Function
    If Not uiBindingContext.SetValue("Settings", "TestReadOnly", configuredText <> "TEST") Then Exit Function
    If configuredText = "TEST" Then
        If Not uiBindingContext.SetValue("Settings", "TestButtonStyle", "secondaryButton") Then Exit Function
    Else
        If Not uiBindingContext.SetValue("Settings", "TestButtonStyle", "disabledButton") Then Exit Function
    End If
    If Not uiBindingContext.SetValue("Form", "TestScenario", VBA.vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "ExtendedMode", VBA.vbNullString) Then Exit Function
    If Not ex_Core.fn_TryGetWorkbookConfigValue("PersonalEventBuilder::text.OrderUnaccepted", configuredText) Then
        ex_WindowsUi.fn_ShowMessage "Required configuration text not found: PersonalEventBuilder::text.OrderUnaccepted", vbExclamation, "Configuration"
        Exit Function
    End If
    If Not uiBindingContext.SetValue("Form", "AcceptedOrderText", configuredText) Then Exit Function
    If Not ex_Core.fn_TryGetWorkbookConfigValue("PersonalEventBuilder::text.OrderReady", configuredText) Then
        ex_WindowsUi.fn_ShowMessage "Required configuration text not found: PersonalEventBuilder::text.OrderReady", vbExclamation, "Configuration"
        Exit Function
    End If
    If Not uiBindingContext.SetValue("Form", "ActionStatus", configuredText) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "EventType", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "Destination", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "PersonRank", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "PersonName", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "PersonId", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "PositionCode", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "PersonPosition", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "ReportRank", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "ReportName", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "ReportPositionCode", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "IncomingNumber", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "IncomingDate", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "EventDate", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "TicketNumber", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "TicketDate", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "Hospital", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "HospitalShortName", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "Document", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "DocumentDate", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "DocumentNumber", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "ModeId", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "ModeDirection", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "ModeForm", vbNullString) Then Exit Function
    private_TryRegisterBindings = True
End Function

Private Function private_TryRegisterCommands() As Boolean
    Dim lookupCommand As obj_UiCommand
    Dim uiBindingContext As obj_UiBindingContext
    Dim orderCommand As Variant

    If m_controller Is Nothing Or m_pageBase Is Nothing Then Exit Function
    Set uiBindingContext = m_pageBase.BindingContext
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "SearchPersonnelHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "SearchPersonnel", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "SelectPersonnelHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "SelectPersonnel", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "RefreshPersonnelHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "RefreshPersonnel", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "SelectHistoryHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "SelectHistory", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "SelectConflictHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "SelectConflict", lookupCommand) Then Exit Function
    For Each orderCommand In VBA.Array("SearchOrders", "OrderQueryChanged", "AutoFillPersonnelChanged", "PersonNameChanged", _
            "PreviewWord", "SaveHistory", "ShowHistory", "ExportWord", "ExportMovement", "ReportNameChanged", "HospitalChanged", "DocumentChanged", "EnvironmentSwitch", _
            "LoadTest", "NextTest", "ResetForm", "ReinitializePage", "ExtendedModeChanged", "CloseCandidates", "CloseHistory", _
            "EventVersions", "CancelEvent", "RestoreEvent", _
            "DummyMode11", "DummyMode12", "DummyMode13", "DummyMode14", _
            "DummyMode21", "DummyMode22", "DummyMode23", "DummyMode24")
        Set lookupCommand = New obj_UiCommand
        If Not lookupCommand.Initialize(m_controller, VBA.CStr(orderCommand) & "Handler") Then Exit Function
        If Not uiBindingContext.SetObject("Commands", VBA.CStr(orderCommand), lookupCommand) Then Exit Function
    Next orderCommand
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "FromHospitalHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "FromHospital", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "FromAnnualMainVacationHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "FromAnnualMainVacation", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "FromVacationForCuringHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "FromVacationForCuring", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "FromVacationForFamilyHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "FromVacationForFamily", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "FromOutpatientHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "FromOutpatient", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "FromMedicalCompanyHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "FromMedicalCompany", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "FromBusinessTripHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "FromBusinessTrip", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "ToHospitalHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "ToHospital", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "ToAnnualMainVacationHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "ToAnnualMainVacation", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "ToVacationForCuringHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "ToVacationForCuring", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "ToVacationForFamilyHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "ToVacationForFamily", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "ToOutpatientHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "ToOutpatient", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "ToMedicalCompanyHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "ToMedicalCompany", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "ToBusinessTripHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "ToBusinessTrip", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "HospitalToLeaveHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "HospitalToLeave", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "MedicalLeaveExtendHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "MedicalLeaveExtend", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "OutpatientToMedicalLeaveHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "OutpatientToMedicalLeave", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "MedCompanyToMedicalLeaveHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "MedCompanyToMedicalLeave", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "ToBusinessTripAwolHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "ToBusinessTripAwol", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "AnnualToMedicalLeaveHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "AnnualToMedicalLeave", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "FamilyToMedicalLeaveHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "FamilyToMedicalLeave", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "MedicalLeaveToHospitalHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "MedicalLeaveToHospital", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "MedicalLeaveToAnnualHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "MedicalLeaveToAnnual", lookupCommand) Then Exit Function
    private_TryRegisterCommands = True
End Function