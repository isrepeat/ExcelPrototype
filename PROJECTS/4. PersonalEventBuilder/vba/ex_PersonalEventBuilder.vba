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
    ex_Core.fn_Diagnostic_WriteLog "INITIALIZE_STARTED | Workbook=" & ThisWorkbook.Name
    ex_UiRenderer.fn_RenderPages
    ex_Core.fn_Diagnostic_WriteLog "INITIALIZE_COMPLETED | Workbook=" & ThisWorkbook.Name
End Sub

Public Sub fn_HelloWorld()
    ex_Core.fn_Diagnostic_WriteLog "HELLO_WORLD_CLICKED"
    VBA.MsgBox "Hello World from PersonalEventBuilder.", VBA.vbInformation, _
        "PersonalEventBuilder"
End Sub

Public Sub fn_UpdatePage()
    Dim previousScreenUpdating As Boolean

    previousScreenUpdating = Application.ScreenUpdating
    On Error GoTo EH
    Application.ScreenUpdating = False
    ex_Core.fn_Diagnostic_WriteLog "UPDATE_PAGE_CLICKED | Sheet=" & _
        Application.ActiveSheet.Name
    ex_UiRuntime.fn_RenderActivePage
CleanExit:
    Application.ScreenUpdating = previousScreenUpdating
    Exit Sub
EH:
    ex_Core.fn_Diagnostic_WriteLog "UPDATE_PAGE_ERROR | Number=" & _
        VBA.CStr(VBA.Err.Number) & " | Description=" & VBA.Err.Description
    VBA.MsgBox "The page cannot be updated: " & VBA.Err.Description, _
        VBA.vbExclamation, "PersonalEventBuilder"
    Resume CleanExit
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------