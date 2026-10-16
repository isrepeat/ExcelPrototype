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

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_pageBase As obj_PageBase
Private m_controller As obj_PEB_PgMainController

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
    Dim uiFolder As String

    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    Set m_pageBase = New obj_PageBase
    uiFolder = ex_PEB_Export.fn_Config("ui.folder")
    If Not m_pageBase.Initialize(profileId, uiFolder) Then Exit Function
    If Not private_TryRegisterBindings() Then Exit Function
    Set m_controller = New obj_PEB_PgMainController
    If Not m_controller.Initialize(m_pageBase) Then Exit Function
    If Not private_TryRegisterCommands() Then Exit Function
    If ex_PEB_Export.fn_Config("environment.name") = "TEST" Then
        If Not m_pageBase.BindingContext.SetValue("Form", "ActionStatus", VBA.vbNullString) Then Exit Function
    End If
    obj_IPage_Initialize = True
    m_isInitialized = obj_IPage_Initialize
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
    If Not obj_IPage_Render Then Exit Function
    If ex_PEB_Export.fn_Config("environment.name") = "TEST" Then
        On Error GoTo EH_TEST_PANEL
        If Not ex_PEB_Testing.fn_Open() Then Err.Raise 5, , "Cannot open the test scenario panel."
    End If
    Exit Function
EH_TEST_PANEL:
    ex_WindowsUi.fn_ShowMessage VBA.Err.Description, vbExclamation, ex_PEB_Export.fn_Config("text.TestTitle")
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
Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
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
    Dim uiBindingContext As obj_UiBindingContext
    Dim rawTable As obj_UiRawTable
    Dim previews As Collection
    Dim textKey As Variant
    Dim configuredText As String
    Dim confirmSelectionKey As String
    Dim orderOptions As Collection

    If m_pageBase Is Nothing Then Exit Function
    Set uiBindingContext = m_pageBase.BindingContext
    If uiBindingContext Is Nothing Then Exit Function
    Set rawTable = New obj_UiRawTable
    If Not rawTable.InitializeEmpty(VBA.Split(ex_PEB_Export.fn_Config("test.ui.catalogCaptions"), "|")) Then Exit Function
    If Not uiBindingContext.SetObject("Data", "TestCatalog", rawTable) Then Exit Function
    Set rawTable = New obj_UiRawTable
    If Not rawTable.InitializeEmpty(VBA.Split(ex_PEB_Export.fn_Config("test.panel.captions"), "|"), _
            VBA.vbNullString, VBA.Split(ex_PEB_Export.fn_Config("test.panel.fields"), "|")) Then Exit Function
    If Not uiBindingContext.SetObject("Data", "TestResults", rawTable) Then Exit Function
    Set previews = New Collection
    If Not uiBindingContext.SetObject("Data", "TestPreviews", previews) Then Exit Function
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
            "SourceLegendSample", "SourceLegendSeparator", _
            "DummyMode", "ExtendedModeItems", "EnvironmentSwitch", _
            "TestScenarioItems", "LoadTest", "NextTest", "OpenTests", "TestsStop", "TestsPageTitle", "TestsRefresh", _
            "TestsResults", "PreviewWord", "ExportWord", "ExportMovement")
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
    If Not uiBindingContext.SetValue("Settings", "TestMode", configuredText = "TEST") Then Exit Function
    If Not uiBindingContext.SetValue("Settings", "TestsRunning", False) Then Exit Function
    If Not uiBindingContext.SetValue("Settings", "TestReadOnly", configuredText <> "TEST") Then Exit Function
    If configuredText = "TEST" Then
        If Not uiBindingContext.SetValue("Settings", "TestButtonStyle", "secondaryButton") Then Exit Function
    Else
        If Not uiBindingContext.SetValue("Settings", "TestButtonStyle", "disabledButton") Then Exit Function
    End If
    If Not uiBindingContext.SetValue("Form", "TestScenario", VBA.vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "ExtendedMode", ex_PEB_Export.fn_Config("text.ExtendedModes")) Then Exit Function
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
    For Each orderCommand In VBA.Array("SearchOrders", "OrderQueryChanged", "AutoFillPersonnelChanged", "PersonNameChanged", _
            "PreviewWord", "ExportWord", "ExportMovement", "ReportNameChanged", "HospitalChanged", "EnvironmentSwitch", _
            "LoadTest", "NextTest", "OpenTests", "CancelTests", "TestPreview", _
            "DummyMode11", "DummyMode12", "DummyMode13", "DummyMode14", _
            "DummyMode21", "DummyMode22", "DummyMode23", "DummyMode24")
        Set lookupCommand = New obj_UiCommand
        If Not lookupCommand.Initialize(m_controller, VBA.CStr(orderCommand) & "Handler") Then Exit Function
        If Not uiBindingContext.SetObject("Commands", VBA.CStr(orderCommand), lookupCommand) Then Exit Function
    Next orderCommand
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "ArrivalTreatmentHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "ArrivalTreatment", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "ArrivalAnnualLeaveHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "ArrivalAnnualLeave", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "ArrivalMedicalLeaveHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "ArrivalMedicalLeave", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "ArrivalFamilyLeaveHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "ArrivalFamilyLeave", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "DepartureTreatmentHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "DepartureTreatment", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "DepartureAnnualLeaveHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "DepartureAnnualLeave", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "DepartureMedicalLeaveHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "DepartureMedicalLeave", lookupCommand) Then Exit Function
    Set lookupCommand = New obj_UiCommand
    If Not lookupCommand.Initialize(m_controller, "DepartureFamilyLeaveHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "DepartureFamilyLeave", lookupCommand) Then Exit Function
    private_TryRegisterCommands = True
End Function