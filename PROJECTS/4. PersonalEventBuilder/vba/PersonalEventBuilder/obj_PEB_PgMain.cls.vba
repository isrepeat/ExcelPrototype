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

' --------------------------------------
' namespace API {
' --------------------------------------
Private Function obj_IPage_Initialize(ByVal profileId As String) As Boolean
    Set m_pageBase = New obj_PageBase
    If Not m_pageBase.fn_Initialize(profileId, "ui\" & VBA.Trim$(profileId)) Then Exit Function
    If Not private_TryRegisterBindings() Then Exit Function
    Set m_controller = New obj_PEB_PgMainController
    If Not m_controller.fn_Initialize(m_pageBase) Then Exit Function
    If Not private_TryRegisterCommands() Then Exit Function
    obj_IPage_Initialize = True
End Function

Private Function obj_IPage_Render() As Boolean
    If m_pageBase Is Nothing Then Exit Function
    obj_IPage_Render = m_pageBase.fn_Render()
End Function

Private Function obj_IPage_HandleCellChange(ByVal target As Range) As Boolean
    If target Is Nothing Then Exit Function
    obj_IPage_HandleCellChange = True
End Function

Private Sub obj_IPage_Dispose()
    If Not m_controller Is Nothing Then m_controller.fn_Dispose
    Set m_controller = Nothing
    If Not m_pageBase Is Nothing Then m_pageBase.fn_Dispose
    Set m_pageBase = Nothing
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_TryRegisterBindings() As Boolean
    Dim uiBindingContext As obj_UiBindingContext

    If m_pageBase Is Nothing Then Exit Function
    Set uiBindingContext = m_pageBase.fn_BindingContext
    If uiBindingContext Is Nothing Then Exit Function
    If Not uiBindingContext.fn_SetValue("Text", "Title", "PersonalEventBuilder") Then Exit Function
    If Not uiBindingContext.fn_SetValue("Text", "HelloWorld", "Hello World") Then Exit Function
    If Not uiBindingContext.fn_SetValue("Text", "UpdatePage", "Update page") Then Exit Function
    If Not uiBindingContext.fn_SetValue("Resources", "PrimaryButton", "primaryButton") Then Exit Function
    If Not uiBindingContext.fn_SetValue("Resources", "PageTitle", "pageTitle") Then Exit Function
    private_TryRegisterBindings = True
End Function

Private Function private_TryRegisterCommands() As Boolean
    Dim helloWorldCommand As obj_UiCommand
    Dim updatePageCommand As obj_UiCommand
    Dim uiBindingContext As obj_UiBindingContext

    If m_controller Is Nothing Then Exit Function
    If m_pageBase Is Nothing Then Exit Function
    Set uiBindingContext = m_pageBase.fn_BindingContext
    If uiBindingContext Is Nothing Then Exit Function
    Set helloWorldCommand = New obj_UiCommand
    If Not helloWorldCommand.fn_Initialize(m_controller, "fn_HelloWorld") Then Exit Function
    If Not uiBindingContext.fn_SetObject("Commands", "HelloWorld", helloWorldCommand) Then Exit Function
    Set updatePageCommand = New obj_UiCommand
    If Not updatePageCommand.fn_Initialize(m_controller, "fn_UpdatePage") Then Exit Function
    If Not uiBindingContext.fn_SetObject("Commands", "UpdatePage", updatePageCommand) Then Exit Function
    private_TryRegisterCommands = True
End Function