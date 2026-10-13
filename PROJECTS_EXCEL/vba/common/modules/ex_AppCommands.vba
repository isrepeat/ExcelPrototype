Option Explicit

Private Const DIAGNOSTIC_MODE_IMMEDIATE As String = "Immediate"
Private Const DIAGNOSTIC_MODE_BUFFERED As String = "Buffered"

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
Public Sub fn_RegisterHotkeys()
    If Not ex_AppHotkeys.fn_Register("^%l", "ex_AppCommands.fn_ToggleDiagnosticMode") Then VBA.Err.Raise 5, , "Hotkey registration failed: Ctrl+Alt+L"
    If Not ex_AppHotkeys.fn_Register("^%u", "ex_AppCommands.fn_UpdateConfig") Then VBA.Err.Raise 5, , "Hotkey registration failed: Ctrl+Alt+U"
    If Not ex_AppHotkeys.fn_Register("^r", "ex_AppCommands.fn_UpdatePage") Then VBA.Err.Raise 5, , "Hotkey registration failed: Ctrl+R"
End Sub

Public Sub fn_UpdateConfig()
    Dim runtimeContext As Object
    Dim diagnostic As String
    Dim errorNumber As Long
    Dim errorDescription As String

    If Not Application.ActiveWorkbook Is ThisWorkbook Then Exit Sub
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    On Error GoTo EH
    If ex_ConfigRefresh.fn_TryUpdate(diagnostic) Then
        ex_Core.fn_StatusBar_Show diagnostic
    Else
        ex_WindowsUi.fn_ShowMessage diagnostic, VBA.vbExclamation
    End If
Cleanup:
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    VBA.Err.Raise errorNumber, "ex_AppCommands.fn_UpdateConfig", errorDescription
End Sub

Public Sub fn_UpdatePage()
    Dim runtimeContext As Object
    Dim errorNumber As Long
    Dim errorDescription As String

    If Not Application.ActiveWorkbook Is ThisWorkbook Then Exit Sub
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    On Error GoTo EH_UPDATE
    If Not ex_UiRuntime.fn_UpdateActivePage() Then
        ex_Core.fn_Diagnostic_WriteLog "HOTKEY_PAGE_UPDATE_FAILED"
    End If
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH_UPDATE:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    VBA.Err.Raise errorNumber, "fn_UpdatePage", errorDescription
End Sub

Public Sub fn_ToggleDiagnosticMode()
    Dim runtimeContext As Object
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim currentMode As String
    Dim nextMode As String

    If Not Application.ActiveWorkbook Is ThisWorkbook Then Exit Sub
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    On Error GoTo EH_TOGGLE
    currentMode = ex_Core.fn_Diagnostic_GetMode()
    If VBA.StrComp(currentMode, DIAGNOSTIC_MODE_BUFFERED, VBA.vbTextCompare) = 0 Then
        nextMode = DIAGNOSTIC_MODE_IMMEDIATE
    Else
        nextMode = DIAGNOSTIC_MODE_BUFFERED
    End If

    If Not ex_Core.fn_Diagnostic_SetMode(nextMode) Then
        ex_WindowsUi.fn_ShowMessage "Could not change diagnostic logging mode.", _
            VBA.vbExclamation, "Diagnostic logging"
        GoTo CleanToggle
    End If
    ex_Core.fn_Diagnostic_WriteLog "DIAGNOSTIC_MODE_CHANGED | Mode=" & nextMode
    ex_Core.fn_StatusBar_Show "Diagnostic logging mode: " & nextMode & " | Save the workbook to keep this setting."
CleanToggle:
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH_TOGGLE:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    VBA.Err.Raise errorNumber, "fn_ToggleDiagnosticMode", errorDescription
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------