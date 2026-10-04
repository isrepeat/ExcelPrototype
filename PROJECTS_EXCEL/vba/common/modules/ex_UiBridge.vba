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
    Dim targetWorksheet As Worksheet
    Dim shapeName As String
    Dim diagnostic As String

    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    On Error GoTo EH
    shapeName = VBA.CStr(Application.Caller)
    If TypeOf Application.ActiveSheet Is Worksheet Then
        Set targetWorksheet = Application.ActiveSheet
        If ex_UiRuntime.fn_TryGetContext(targetWorksheet, context) Then
            context.Router.DispatchShape shapeName
            ' The command can replace and dispose the page context during rendering.
            Set context = Nothing
            If ex_UiRuntime.fn_TryGetContext(targetWorksheet, context) Then
                If Not context.FlushLayout(diagnostic) Then
                    VBA.Err.Raise VBA.vbObjectError + 2231, , diagnostic
                End If
            End If
        End If
    End If
CleanExit:
    Set context = Nothing
    Set targetWorksheet = Nothing
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    Set context = Nothing
    Set targetWorksheet = Nothing
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    VBA.Err.Raise errorNumber, "ex_UiBridge.fn_OnShapeClick", errorDescription
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------