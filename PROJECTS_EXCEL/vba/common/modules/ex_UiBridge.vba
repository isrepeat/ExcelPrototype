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
    Dim runtimeContext As Object
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim uiCommand As obj_UiCommand
    Dim shapeName As String

    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    On Error GoTo EH
    shapeName = VBA.CStr(Application.Caller)
    If ex_UiBindings.fn_HandleSelectShapeClick(shapeName) Then GoTo CleanExit
    ex_UiBindings.fn_CollapseSelectControls
    If Not ex_UiBindings.fn_TryGetCommand(shapeName, uiCommand) Then
        VBA.MsgBox "No command is registered for the selected button.", _
            VBA.vbExclamation, "PersonalEventBuilder"
        GoTo CleanExit
    End If
    ex_Core.fn_Diagnostic_WriteLog "UI_CLICK | Shape=" & shapeName & _
        " | Command=" & uiCommand.CallbackName
    uiCommand.Execute
CleanExit:
    Set uiCommand = Nothing
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Set uiCommand = Nothing
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Err.Raise errorNumber, "ex_UiBridge.fn_OnShapeClick", errorDescription
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------