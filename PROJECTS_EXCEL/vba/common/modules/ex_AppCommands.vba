Option Explicit

Private Const DIAGNOSTIC_MODE_IMMEDIATE As String = "Immediate"
Private Const DIAGNOSTIC_MODE_BUFFERED As String = "Buffered"
Private m_reinitializeAt As Date

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    If m_reinitializeAt = 0 Then Exit Sub
    Application.OnTime m_reinitializeAt, private_ReinitializeCallback(), Schedule:=False
    m_reinitializeAt = 0
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_RegisterHotkeys()
    Dim confirmSelectionKey As String

    If Not ex_AppHotkeys.fn_Register("^%l", "ex_AppCommands.fn_ToggleDiagnosticMode") Then VBA.Err.Raise 5, , "Hotkey registration failed: Ctrl+Alt+L"
    If Not ex_AppHotkeys.fn_Register("^%u", "ex_AppCommands.fn_UpdateConfig") Then VBA.Err.Raise 5, , "Hotkey registration failed: Ctrl+Alt+U"
    If Not ex_AppHotkeys.fn_Register("^r", "ex_AppCommands.fn_UpdatePage") Then VBA.Err.Raise 5, , "Hotkey registration failed: Ctrl+R"
    If ex_Core.fn_TryGetWorkbookConfigValue("ThisWorkbook::ui.confirmSelectionKey", confirmSelectionKey) Then
        If VBA.Len(VBA.Trim$(confirmSelectionKey)) = 0 Then
            VBA.Err.Raise 5, , "Empty configuration key: ThisWorkbook::ui.confirmSelectionKey"
        End If
        If Not ex_AppHotkeys.fn_Register(confirmSelectionKey, "ex_AppCommands.fn_ConfirmSelection") Then
            VBA.Err.Raise 5, , "Hotkey registration failed: " & confirmSelectionKey
        End If
    End If
End Sub

Public Sub fn_ConfirmSelection()
    Dim runtimeContext As Object
    Dim target As Range
    Dim errorNumber As Long
    Dim errorDescription As String

    If Not Application.ActiveWorkbook Is ThisWorkbook Then Exit Sub
    If Not TypeOf Application.Selection Is Range Then Exit Sub
    Set target = Application.Selection
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    On Error GoTo EH_CONFIRM
    ex_UiBindings.fn_HandleSelection target, "confirmSelection"
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH_CONFIRM:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    VBA.Err.Raise errorNumber, "ex_AppCommands.fn_ConfirmSelection", errorDescription
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

Public Function fn_RequestReinitialize() As Boolean
    If m_reinitializeAt <> 0 Then
        fn_RequestReinitialize = True
        Exit Function
    End If
    m_reinitializeAt = VBA.Now + VBA.TimeSerial(0, 0, 1)
    Application.OnTime m_reinitializeAt, private_ReinitializeCallback()
    fn_RequestReinitialize = True
End Function

Public Sub fn_ReinitializePage()
    Dim runtimeContext As Object
    Dim initializerName As String
    Dim previousEvents As Boolean
    Dim previousUpdating As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    m_reinitializeAt = 0
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    previousEvents = Application.EnableEvents
    previousUpdating = Application.ScreenUpdating
    On Error GoTo EH_REINITIALIZE
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    initializerName = ex_WorkbookCallbacks.fn_Resolve("ThisWorkbook::initializer")
    ex_Core.fn_Diagnostic_WriteLog "CONFIG_REINITIALIZE_STARTED"
    ex_TableDataRuntime.fn_ResetCaches
    ex_RuntimePaths.fn_Module_Dispose
    If Not VBA.CBool(Application.Run(initializerName)) Then
        VBA.Err.Raise 5, , "Page reinitialization failed: " & initializerName
    End If
    ex_Core.fn_Diagnostic_WriteLog "CONFIG_REINITIALIZE_COMPLETED"
CleanupReinitialize:
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousUpdating
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH_REINITIALIZE:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    ex_Core.fn_Diagnostic_WriteLog "CONFIG_REINITIALIZE_FAILED | Number=" & VBA.CStr(errorNumber) & _
        " | Description=" & errorDescription
    ex_WindowsUi.fn_ShowMessage errorDescription, VBA.vbExclamation, "Configuration"
    Resume CleanupReinitialize
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

Private Function private_ReinitializeCallback() As String
    private_ReinitializeCallback = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!ex_AppCommands.fn_ReinitializePage"
End Function