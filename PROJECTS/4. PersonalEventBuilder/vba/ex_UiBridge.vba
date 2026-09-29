Option Explicit

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_OnShapeClick()
    Dim callbackName As String

    If Not ex_UiBindings.fn_TryGetCallback(VBA.CStr(Application.Caller), callbackName) Then
        VBA.MsgBox "No action is registered for the selected button.", _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Sub
    End If
    Application.Run "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & _
        "'!" & callbackName
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------