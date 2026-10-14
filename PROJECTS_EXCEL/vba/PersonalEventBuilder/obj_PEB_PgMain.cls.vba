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
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    Set m_pageBase = New obj_PageBase
    If Not m_pageBase.Initialize(profileId, VBA.Trim$(profileId)) Then Exit Function
    If Not private_TryRegisterBindings() Then Exit Function
    Set m_controller = New obj_PEB_PgMainController
    If Not m_controller.Initialize(m_pageBase) Then Exit Function
    If Not private_TryRegisterCommands() Then Exit Function
    obj_IPage_Initialize = True
    m_isInitialized = obj_IPage_Initialize
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
    Dim textKey As Variant
    Dim configuredText As String
    Dim confirmSelectionKey As String

    If m_pageBase Is Nothing Then Exit Function
    Set uiBindingContext = m_pageBase.BindingContext
    If uiBindingContext Is Nothing Then Exit Function
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
    If Not uiBindingContext.SetValue("Form", "EventNote", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "EventDate", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "PersonName", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "PersonId", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "PersonRank", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "PersonPosition", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "PersonUnit", vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Data", "SelectedPersonnel", vbNullString) Then Exit Function
    For Each textKey In VBA.Array("PersonNameLabel", "PersonIdLabel", "PersonRankLabel", _
            "PersonPositionLabel", "PersonUnitLabel", "EventNoteLabel", "EventDateLabel", _
            "PersonnelSection", "ActionsSection", "RefreshPersonnel")
        If Not ex_Core.fn_TryGetWorkbookConfigValue("PersonalEventBuilder::text." & textKey, configuredText) Then
            ex_WindowsUi.fn_ShowMessage "Required configuration text not found: PersonalEventBuilder::text." & textKey, vbExclamation, "Configuration"
            Exit Function
        End If
        If Not uiBindingContext.SetValue("Text", VBA.CStr(textKey), configuredText) Then Exit Function
    Next textKey
    private_TryRegisterBindings = True
End Function

Private Function private_TryRegisterCommands() As Boolean
    Dim lookupCommand As obj_UiCommand
    Dim uiBindingContext As obj_UiBindingContext

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
    private_TryRegisterCommands = True
End Function