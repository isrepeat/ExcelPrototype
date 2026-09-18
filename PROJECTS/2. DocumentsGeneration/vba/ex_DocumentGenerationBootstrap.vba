Option Explicit

' --------------------------------------
' namespace API {
' --------------------------------------
' Собирает runtime-маршруты функций формы после открытия или hot reload книги.
Public Sub fn_Initialize()
    If Not ex_VacationTicketGeneration.fn_TryInitializeUiRuntime() Then Exit Sub
    ex_CellChangeRouter.fn_Reset
    If Not ex_PersonnelCandidates.fn_RegisterRoutes() Then
        VBA.MsgBox "Personnel candidate routes were not initialized.", _
            VBA.vbExclamation, "Document Generation"
        Exit Sub
    End If
    ex_Helpers.LogDebug "Document Generation cell routes initialized"
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------