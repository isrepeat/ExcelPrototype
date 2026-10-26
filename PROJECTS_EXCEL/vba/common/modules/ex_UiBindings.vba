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
Public Function fn_HandleSelection( _
    ByVal target As Range, _
    Optional ByVal kind As String = "selection" _
) As Boolean
    Dim context As obj_UiRenderContext
    Dim previousEvents As Boolean
    Dim diagnostic As String
    Dim previousScreenUpdating As Boolean
    Dim busyEntered As Boolean
    Dim diagnosticBatchEntered As Boolean

    If target Is Nothing Then Exit Function
    If Not ex_UiRuntime.fn_TryGetContext(target.Parent, context) Then Exit Function
    previousEvents = Application.EnableEvents
    On Error GoTo EH
    previousScreenUpdating = Application.ScreenUpdating
    ex_Core.fn_Diagnostic_BeginBatch
    diagnosticBatchEntered = True
    If kind = "confirmSelection" Then
        busyEntered = ex_UiBusy.fn_Begin()
        Application.ScreenUpdating = False
    End If
    Application.EnableEvents = False
    context.Router.Broadcast "dismiss"
    fn_HandleSelection = context.Router.DispatchSelection(target, kind)
    ' A focus command can rebuild the page and dispose the original context.
    Set context = Nothing
    If ex_UiRuntime.fn_TryGetContext(target.Parent, context) Then
        If Not context.FlushLayout(diagnostic) Then
            VBA.Err.Raise 5, , diagnostic
        End If
    End If
CleanExit:
    Application.EnableEvents = previousEvents
    If kind = "confirmSelection" Then Application.ScreenUpdating = previousScreenUpdating
    If busyEntered Then ex_UiBusy.fn_End
    If diagnosticBatchEntered Then ex_Core.fn_Diagnostic_EndBatch
    Exit Function
EH:
    ex_WindowsUi.fn_ShowMessage "Cannot select table row: " & VBA.Err.Description, vbExclamation, "UI"
    Resume CleanExit
End Function

Public Function fn_HandleCellChange(ByVal target As Range) As Boolean
    Dim context As obj_UiRenderContext
    Dim previousEvents As Boolean
    Dim diagnostic As String
    Dim previousScreenUpdating As Boolean
    Dim busyEntered As Boolean
    Dim diagnosticBatchEntered As Boolean
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    If target Is Nothing Then Exit Function
    If Not ex_UiRuntime.fn_TryGetContext(target.Parent, context) Then Exit Function
    previousEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    On Error GoTo EH
    ex_Core.fn_Diagnostic_BeginBatch
    diagnosticBatchEntered = True
    busyEntered = ex_UiBusy.fn_Begin()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    fn_HandleCellChange = context.Router.DispatchCells(target)
    ' A field change can also replace the entire page context.
    Set context = Nothing
    If ex_UiRuntime.fn_TryGetContext(target.Parent, context) Then
        If Not context.FlushLayout(diagnostic) Then
            VBA.Err.Raise VBA.vbObjectError + 2230, , diagnostic
        End If
    End If
CleanExit:
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreenUpdating
    If busyEntered Then ex_UiBusy.fn_End
    If diagnosticBatchEntered Then ex_Core.fn_Diagnostic_EndBatch
    Exit Function
EH:
    errorNumber = VBA.Err.Number
    errorSource = VBA.Err.Source
    errorDescription = VBA.Err.Description
    ex_Core.fn_Diagnostic_WriteError "UI.CellChange", errorNumber, errorSource, errorDescription
    ex_WindowsUi.fn_ShowMessage "Cannot dispatch cell change: " & errorDescription, VBA.vbExclamation, "UI"
    Resume CleanExit
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------