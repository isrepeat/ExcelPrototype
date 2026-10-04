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
Public Function fn_Initialize() As Boolean
    Dim profileId As String

    If Not ex_Core.fn_TryGetWorkbookProfileId(profileId) Then
        Exit Function
    End If
    If VBA.StrComp(profileId, "PersonnelAtDisposalDaysCalculation", VBA.vbTextCompare) <> 0 Then
        Exit Function
    End If
    If Not ex_UiPageManager.fn_ShowPage("PersonnelAtDisposalDaysCalculation", profileId) Then
        Exit Function
    End If

    ex_Core.fn_Diagnostic_WriteLog "INITIALIZE_COMPLETED | Profile=" & profileId
    fn_Initialize = True
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------