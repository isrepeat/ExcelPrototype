Attribute VB_Name = "ex_UiBindings"
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
Public Function fn_HandleSelection(ByVal target As Range) As Boolean
    Dim context As obj_UiRenderContext
    Dim previousEvents As Boolean
    Dim diagnostic As String

    If target Is Nothing Then Exit Function
    If Not ex_UiRuntime.fn_TryGetContext(target.Parent, context) Then Exit Function
    previousEvents = Application.EnableEvents
    On Error GoTo EH
    Application.EnableEvents = False
    context.Router.Broadcast "dismiss"
    fn_HandleSelection = context.Router.DispatchSelection(target)
    If Not context.FlushLayout(diagnostic) Then VBA.Err.Raise 5, , diagnostic
CleanExit:
    Application.EnableEvents = previousEvents
    Exit Function
EH:
    ex_WindowsUi.fn_ShowMessage "Cannot select table row: " & VBA.Err.Description, vbExclamation, "UI"
    Resume CleanExit
End Function

Public Function fn_HandleCellChange(ByVal target As Range) As Boolean
    Dim context As obj_UiRenderContext
    Dim previousEvents As Boolean
    Dim diagnostic As String

    If target Is Nothing Then Exit Function
    If Not ex_UiRuntime.fn_TryGetContext(target.Parent, context) Then Exit Function
    previousEvents = Application.EnableEvents
    On Error GoTo EH
    Application.EnableEvents = False
    fn_HandleCellChange = context.Router.DispatchCells(target)
    If Not context.FlushLayout(diagnostic) Then VBA.Err.Raise VBA.vbObjectError + 2230, , diagnostic
CleanExit:
    Application.EnableEvents = previousEvents
    Exit Function
EH:
    ex_WindowsUi.fn_ShowMessage "Cannot dispatch cell change: " & VBA.Err.Description, VBA.vbExclamation, "UI"
    Resume CleanExit
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------