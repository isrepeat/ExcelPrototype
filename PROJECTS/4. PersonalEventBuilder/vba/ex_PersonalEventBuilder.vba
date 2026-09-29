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
    ex_Core.fn_Diagnostic_WriteLog "UPDATE_PAGE_CLICKED | Sheet=" & _
        Application.ActiveSheet.Name
    ex_UiRuntime.fn_RenderActivePage
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------