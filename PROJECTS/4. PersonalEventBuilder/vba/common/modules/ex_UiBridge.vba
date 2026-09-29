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
    Dim uiCommand As obj_UiCommand
    Dim shapeName As String

    shapeName = VBA.CStr(Application.Caller)
    If Not ex_UiBindings.fn_TryGetCommand(shapeName, uiCommand) Then
        VBA.MsgBox "No command is registered for the selected button.", _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Sub
    End If
    ex_Core.fn_Diagnostic_WriteLog "UI_CLICK | Shape=" & shapeName & _
        " | Command=" & uiCommand.fn_CallbackName
    uiCommand.fn_Execute
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------