Option Explicit

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    ex_UiPageManager.fn_Module_Dispose
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_Initialize()
    Dim profileId As String

    If Not ex_Core.fn_TryGetWorkbookProfileId(profileId) Then Exit Sub
    ex_Core.fn_Diagnostic_WriteLog "INITIALIZE_STARTED | Workbook=" & ThisWorkbook.Name
    If Not ex_UiPageManager.fn_ShowPage("PersonalEventBuilder", profileId) Then Exit Sub
    ex_Core.fn_Diagnostic_WriteLog "INITIALIZE_COMPLETED | Workbook=" & ThisWorkbook.Name
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------