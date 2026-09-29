Option Explicit

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_Initialize()
    Dim profileId As String
    Dim uiBindingContext As obj_UiBindingContext

    If Not ex_Core.fn_TryGetWorkbookProfileId(profileId) Then Exit Sub
    If Not private_TryCreateBindingContext(uiBindingContext) Then Exit Sub
    ex_Core.fn_Diagnostic_WriteLog "INITIALIZE_STARTED | Workbook=" & ThisWorkbook.Name
    ex_UiRenderer.fn_RenderPages "ui\" & profileId, uiBindingContext
    ex_Core.fn_Diagnostic_WriteLog "INITIALIZE_COMPLETED | Workbook=" & ThisWorkbook.Name
End Sub

Public Sub fn_HelloWorld()
    ex_Core.fn_Diagnostic_WriteLog "HELLO_WORLD_CLICKED"
    VBA.MsgBox "Hello World from PersonalEventBuilder.", VBA.vbInformation, "PersonalEventBuilder"
End Sub

Public Sub fn_UpdatePage()
    Dim previousScreenUpdating As Boolean
    Dim profileId As String
    Dim uiBindingContext As obj_UiBindingContext

    previousScreenUpdating = Application.ScreenUpdating
    On Error GoTo EH
    If Not ex_Core.fn_TryGetWorkbookProfileId(profileId) Then Exit Sub
    If Not private_TryCreateBindingContext(uiBindingContext) Then Exit Sub
    Application.ScreenUpdating = False
    ex_UiRenderer.fn_RenderActivePage "ui\" & profileId, uiBindingContext
CleanExit:
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub
EH:
    VBA.MsgBox "The page cannot be updated: " & VBA.Err.Description, _
        VBA.vbExclamation, "PersonalEventBuilder"
    Resume CleanExit
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_TryCreateBindingContext(ByRef outUiBindingContext As obj_UiBindingContext) As Boolean
    Dim helloWorldCommand As obj_UiCommand
    Dim updatePageCommand As obj_UiCommand

    Set outUiBindingContext = New obj_UiBindingContext
    If Not outUiBindingContext.fn_Initialize() Then Exit Function
    If Not outUiBindingContext.fn_SetValue("Text", "Title", "PersonalEventBuilder") Then Exit Function
    If Not outUiBindingContext.fn_SetValue("Text", "HelloWorld", "Hello World") Then Exit Function
    If Not outUiBindingContext.fn_SetValue("Text", "UpdatePage", "Update page") Then Exit Function
    If Not outUiBindingContext.fn_SetValue("Resources", "PrimaryButton", "primaryButton") Then Exit Function
    If Not outUiBindingContext.fn_SetValue("Resources", "PageTitle", "pageTitle") Then Exit Function

    Set helloWorldCommand = New obj_UiCommand
    If Not helloWorldCommand.fn_Initialize("ex_PersonalEventBuilder.fn_HelloWorld") Then Exit Function
    If Not outUiBindingContext.fn_SetObject("Commands", "HelloWorld", helloWorldCommand) Then Exit Function
    Set updatePageCommand = New obj_UiCommand
    If Not updatePageCommand.fn_Initialize("ex_PersonalEventBuilder.fn_UpdatePage") Then Exit Function
    If Not outUiBindingContext.fn_SetObject("Commands", "UpdatePage", updatePageCommand) Then Exit Function
    private_TryCreateBindingContext = True
End Function