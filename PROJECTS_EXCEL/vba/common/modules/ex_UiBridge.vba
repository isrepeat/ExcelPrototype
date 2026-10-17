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
    Dim previousScreenUpdating As Boolean
    Dim busyEntered As Boolean
    Dim startedAt As Double
    Dim stageStartedAt As Double

    startedAt = VBA.Timer
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    previousScreenUpdating = Application.ScreenUpdating
    On Error GoTo EH
    busyEntered = ex_UiBusy.fn_Begin()
    ' Hide intermediate changes until the command and layout refresh complete.
    Application.ScreenUpdating = False
    shapeName = VBA.CStr(Application.Caller)
    If TypeOf Application.ActiveSheet Is Worksheet Then
        Set targetWorksheet = Application.ActiveSheet
        If ex_UiRuntime.fn_TryGetContext(targetWorksheet, context) Then
            stageStartedAt = VBA.Timer
            context.Router.DispatchShape shapeName
            ex_Core.fn_Diagnostic_WritePerf "Ui.Click.Dispatch | Shape=" & shapeName, stageStartedAt
            ' The command can replace and dispose the page context during rendering.
            Set context = Nothing
            If ex_UiRuntime.fn_TryGetContext(targetWorksheet, context) Then
                stageStartedAt = VBA.Timer
                If Not context.FlushLayout(diagnostic) Then
                    VBA.Err.Raise VBA.vbObjectError + 2231, , diagnostic
                End If
                ex_Core.fn_Diagnostic_WritePerf "Ui.Click.Flush | Shape=" & shapeName, stageStartedAt
            End If
        End If
    End If
CleanExit:
    Application.ScreenUpdating = previousScreenUpdating
    If busyEntered Then ex_UiBusy.fn_End
    Set context = Nothing
    Set targetWorksheet = Nothing
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    ex_Core.fn_Diagnostic_WritePerf "Ui.Click.Total | Shape=" & shapeName, startedAt
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    Application.ScreenUpdating = previousScreenUpdating
    If busyEntered Then ex_UiBusy.fn_End
    Set context = Nothing
    Set targetWorksheet = Nothing
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    VBA.Err.Raise errorNumber, "ex_UiBridge.fn_OnShapeClick", errorDescription
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------