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
    Dim context As obj_UiRenderContext
    Dim shapeName As String
    Dim diagnostic As String

    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    On Error GoTo EH
    shapeName = VBA.CStr(Application.Caller)
    If TypeOf Application.ActiveSheet Is Worksheet Then
        If ex_UiRuntime.fn_TryGetContext(Application.ActiveSheet, context) Then
            context.Router.DispatchShape shapeName
            If Not context.FlushLayout(diagnostic) Then VBA.Err.Raise VBA.vbObjectError + 2231, , diagnostic
        End If
    End If
CleanExit:
    Set context = Nothing
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    Set context = Nothing
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    VBA.Err.Raise errorNumber, "ex_UiBridge.fn_OnShapeClick", errorDescription
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------