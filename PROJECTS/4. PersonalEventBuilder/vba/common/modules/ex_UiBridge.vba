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
Public Sub fn_OnShapeClick()
    Dim callbackName As String

    If Not ex_UiBindings.fn_TryGetCallback(VBA.CStr(Application.Caller), callbackName) Then
        ex_Core.fn_Diagnostic_WriteLog "UI_CLICK_BINDING_MISSING | Shape=" & _
            VBA.CStr(Application.Caller)
        VBA.MsgBox "No action is registered for the selected button.", _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Sub
    End If
    ex_Core.fn_Diagnostic_WriteLog "UI_CLICK | Shape=" & VBA.CStr(Application.Caller) & _
        " | Callback=" & callbackName
    Application.Run "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & _
        "'!" & callbackName
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------