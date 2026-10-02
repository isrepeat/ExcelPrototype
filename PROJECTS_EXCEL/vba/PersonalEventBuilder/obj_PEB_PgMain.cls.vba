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

Private m_pageBase As obj_PageBase
Private m_controller As obj_PEB_PgMainController
Private m_isDisposed As Boolean

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    obj_IPage_Dispose
End Sub

' //
' // Interface
' //
Private Function obj_IPage_Initialize(ByVal profileId As String) As Boolean
    m_isDisposed = False
    Set m_pageBase = New obj_PageBase
    If Not m_pageBase.Initialize(profileId, VBA.Trim$(profileId)) Then Exit Function
    If Not private_TryRegisterBindings() Then Exit Function
    Set m_controller = New obj_PEB_PgMainController
    If Not m_controller.Initialize(m_pageBase) Then Exit Function
    If Not private_TryRegisterCommands() Then Exit Function
    obj_IPage_Initialize = True
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
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
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

    If m_pageBase Is Nothing Then Exit Function
    Set uiBindingContext = m_pageBase.BindingContext
    If uiBindingContext Is Nothing Then Exit Function
    If Not uiBindingContext.SetValue("Text", "Title", "PersonalEventBuilder") Then Exit Function
    If Not uiBindingContext.SetValue("Text", "Reset", "Reset") Then Exit Function
    If Not uiBindingContext.SetValue("Text", "UpdatePage", "Update page") Then Exit Function
    If Not uiBindingContext.SetValue("Text", "GenerateTables", "Generate tables") Then Exit Function
    If Not uiBindingContext.SetValue("Form", "EventName", VBA.vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Form", "Category", "Meeting") Then Exit Function
    If Not uiBindingContext.SetValue("Form", "Notes", VBA.vbNullString) Then Exit Function
    If Not uiBindingContext.SetValue("Data", "EventTypes", "Meeting,Training,Leave") Then Exit Function
    If Not uiBindingContext.SetValue("Resources", "PrimaryButton", "primaryButton") Then Exit Function
    If Not uiBindingContext.SetValue("Resources", "PageTitle", "pageTitle") Then Exit Function
    private_TryRegisterBindings = True
End Function

Private Function private_TryRegisterCommands() As Boolean
    Dim resetCommand As obj_UiCommand
    Dim updatePageCommand As obj_UiCommand
    Dim generateTablesCommand As obj_UiCommand
    Dim formChangedCommand As obj_UiCommand
    Dim submitFormCommand As obj_UiCommand
    Dim uiBindingContext As obj_UiBindingContext

    If m_controller Is Nothing Then Exit Function
    If m_pageBase Is Nothing Then Exit Function
    Set uiBindingContext = m_pageBase.BindingContext
    If uiBindingContext Is Nothing Then Exit Function
    Set resetCommand = New obj_UiCommand
    If Not resetCommand.Initialize(m_controller, "ResetCommandHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "ResetCommand", resetCommand) Then Exit Function
    Set updatePageCommand = New obj_UiCommand
    If Not updatePageCommand.Initialize(m_controller, "UpdatePageCommandHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "UpdatePageCommand", updatePageCommand) Then Exit Function
    Set generateTablesCommand = New obj_UiCommand
    If Not generateTablesCommand.Initialize(m_controller, "GenerateTablesCommandHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "GenerateTablesCommand", generateTablesCommand) Then Exit Function
    Set formChangedCommand = New obj_UiCommand
    If Not formChangedCommand.Initialize(m_controller, "FormChangedCommandHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "FormChangedCommand", formChangedCommand) Then Exit Function
    Set submitFormCommand = New obj_UiCommand
    If Not submitFormCommand.Initialize(m_controller, "SubmitFormCommandHandler") Then Exit Function
    If Not uiBindingContext.SetObject("Commands", "SubmitFormCommand", submitFormCommand) Then Exit Function
    private_TryRegisterCommands = True
End Function