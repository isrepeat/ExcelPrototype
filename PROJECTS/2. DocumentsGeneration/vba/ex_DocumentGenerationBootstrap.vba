Option Explicit

' --------------------------------------
' namespace API {
' --------------------------------------
' Builds form runtime routes after the workbook opens or reloads.
Public Sub fn_Initialize()
    Dim candidatesConfig As Object

    ex_Config.fn_ResetCache
    If Not ex_VacationTicketGeneration.fn_TryInitializeUiRuntime() Then Exit Sub
    If Not ex_VacationTicketGeneration.fn_TryGetCandidatesConfig(candidatesConfig) Then Exit Sub
    If Not ex_Candidates.fn_Configure(candidatesConfig) Then Exit Sub
    ex_CellChangeRouter.fn_Reset
    If Not ex_Candidates.fn_RegisterRoutes() Then
        VBA.MsgBox "Candidate routes were not initialized.", _
            VBA.vbExclamation, "Document Generation"
        Exit Sub
    End If
    ex_Helpers.LogDebug "Document Generation cell routes initialized"
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------