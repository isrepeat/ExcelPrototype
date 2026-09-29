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
' --------------------------------------
' } // namespace API
' --------------------------------------