Option Explicit

Private m_filterInputEvents As obj_FilterInputEvents

' --------------------------------------
' namespace Events {
' --------------------------------------
Private Sub Workbook_Open()
    Dim startedAt As Double
    Dim errorNumber As Long
    Dim errorDescription As String

    startedAt = VBA.Timer
    On Error GoTo EH
    ex_Core.fn_Diagnostic_WriteLog "PERSONAL_WORKBOOK_OPEN_STARTED | Workbook=" & _
        ThisWorkbook.Name & " | ExcelVersion=" & Application.Version & _
        " | Workbooks=" & VBA.CStr(Application.Workbooks.Count)
    private_InitializeFilterInputEvents
    ex_Core.fn_Diagnostic_WriteLog "PERSONAL_WORKBOOK_OPEN_STAGE | Name=InputEventsInitialized"
    BindKeys
    ex_Core.fn_Diagnostic_WriteLog "PERSONAL_WORKBOOK_OPEN_COMPLETED | ElapsedMs=" & _
        private_FormatElapsedMilliseconds(startedAt)
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    ex_Core.fn_Diagnostic_WriteLog "PERSONAL_WORKBOOK_OPEN_ERROR | Number=" & _
        VBA.CStr(errorNumber) & " | Description=" & errorDescription
    VBA.Err.Raise errorNumber, "PERSONAL.Workbook_Open", errorDescription
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim startedAt As Double

    startedAt = VBA.Timer
    ex_Core.fn_Diagnostic_WriteLog "PERSONAL_WORKBOOK_BEFORE_CLOSE_STARTED | Workbook=" & _
        ThisWorkbook.Name
    private_TryUnbindKey "^%r"
    private_TryUnbindKey "^+%r"
    private_TryUnbindKey "^%d"
    ex_Core.fn_Diagnostic_WriteLog "PERSONAL_WORKBOOK_BEFORE_CLOSE_COMPLETED | ElapsedMs=" & _
        private_FormatElapsedMilliseconds(startedAt)
End Sub

' --------------------------------------
' } // namespace Events
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub BindKeys()
    Dim startedAt As Double
    Dim boundCount As Long
    Dim failedCount As Long
    Dim toggleKeyBound As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    startedAt = VBA.Timer
    ex_Core.fn_Diagnostic_WriteLog "PERSONAL_HOTKEY_BINDINGS_STARTED"
    On Error GoTo EH
    private_InitializeFilterInputEvents

    If private_TryBindKey("^q", _
            "ex_ShortcutsHandlers.fn_FilterContainsCurrentColumn") Then
        boundCount = boundCount + 1
    Else
        failedCount = failedCount + 1
    End If
    If private_TryBindKey("^r", _
            "ex_ShortcutsHandlers.fn_RecalculateActiveSheet") Then
        boundCount = boundCount + 1
    Else
        failedCount = failedCount + 1
    End If
    If private_TryBindKey("^%r", "ex_Core.fn_ReloadActiveWorkbookVba") Then
        boundCount = boundCount + 1
    Else
        failedCount = failedCount + 1
    End If
    If private_TryBindKey("^+%r", "ex_Core.fn_ReloadActiveWorkbookVbaDeferred") Then
        boundCount = boundCount + 1
    Else
        failedCount = failedCount + 1
    End If
    If private_TryBindKey("^%d", "ex_Core.fn_ClearActiveWorkbookVba") Then
        boundCount = boundCount + 1
    Else
        failedCount = failedCount + 1
    End If

    If private_TryBindKey("%{PGUP}", _
            "ex_ShortcutsHandlers.fn_DatePlusOne") Then
        boundCount = boundCount + 1
    Else
        failedCount = failedCount + 1
    End If
    If private_TryBindKey("%{PGDN}", _
            "ex_ShortcutsHandlers.fn_DateMinusOne") Then
        boundCount = boundCount + 1
    Else
        failedCount = failedCount + 1
    End If

    ' EN: Ctrl + `
    toggleKeyBound = private_TryBindKey("^`", _
        "ex_ShortcutsHandlers.fn_ToggleFirstTwoRows")
    If toggleKeyBound Then
        boundCount = boundCount + 1
    Else
        failedCount = failedCount + 1
    End If

    ' RU/UKR fallback: Ctrl + '
    If Not toggleKeyBound Then
        toggleKeyBound = private_TryBindKey("^'", _
            "ex_ShortcutsHandlers.fn_ToggleFirstTwoRows")
        If toggleKeyBound Then
            boundCount = boundCount + 1
        Else
            failedCount = failedCount + 1
        End If
    End If

    ' RU fallback: Ctrl + ¸
    If Not toggleKeyBound Then
        toggleKeyBound = private_TryBindKey("^¸", _
            "ex_ShortcutsHandlers.fn_ToggleFirstTwoRows")
        If toggleKeyBound Then
            boundCount = boundCount + 1
        Else
            failedCount = failedCount + 1
        End If
    End If

    ex_Core.fn_Diagnostic_WriteLog "PERSONAL_HOTKEY_BINDINGS_COMPLETED | Bound=" & _
        VBA.CStr(boundCount) & " | FailedAttempts=" & VBA.CStr(failedCount) & _
        " | ElapsedMs=" & private_FormatElapsedMilliseconds(startedAt)
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    ex_Core.fn_Diagnostic_WriteLog "PERSONAL_HOTKEY_BINDINGS_ERROR | Number=" & _
        VBA.CStr(errorNumber) & " | Description=" & errorDescription
End Sub

' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Runtime {
' --------------------------------------
' Recreates runtime event handlers and global keyboard bindings after a code update.
Public Sub fn_ReloadRuntime()
    Dim startedAt As Double
    Dim errorNumber As Long
    Dim errorDescription As String

    startedAt = VBA.Timer
    ex_Core.fn_Diagnostic_WriteLog "PERSONAL_RUNTIME_RELOAD_STARTED"
    On Error GoTo EH
    Set m_filterInputEvents = Nothing
    private_InitializeFilterInputEvents
    BindKeys
    ex_Core.fn_Diagnostic_WriteLog "PERSONAL_RUNTIME_RELOAD_COMPLETED | ElapsedMs=" & _
        private_FormatElapsedMilliseconds(startedAt)
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    ex_Core.fn_Diagnostic_WriteLog "PERSONAL_RUNTIME_RELOAD_ERROR | Number=" & _
        VBA.CStr(errorNumber) & " | Description=" & errorDescription
    VBA.Err.Raise errorNumber, "PERSONAL.fn_ReloadRuntime", errorDescription
End Sub
' --------------------------------------
' } // namespace Runtime
' --------------------------------------

Private Sub private_InitializeFilterInputEvents()
    If m_filterInputEvents Is Nothing Then
        ex_Core.fn_Diagnostic_WriteLog "PERSONAL_INPUT_EVENTS_CREATE_STARTED"
        Set m_filterInputEvents = New obj_FilterInputEvents
        Set m_filterInputEvents.ExcelApplication = Application
        ex_Core.fn_Diagnostic_WriteLog "PERSONAL_INPUT_EVENTS_CREATE_COMPLETED"
    End If
End Sub

Private Function private_TryBindKey( _
    ByVal keySequence As String, _
    ByVal macroName As String _
) As Boolean
    On Error GoTo EH
    Application.OnKey keySequence, private_GlobalMacroRef(macroName)
    private_TryBindKey = True
    Exit Function
EH:
    ex_Core.fn_Diagnostic_WriteLog "PERSONAL_HOTKEY_BIND_ERROR | Key=" & _
        keySequence & " | Macro=" & macroName & " | Number=" & _
        VBA.CStr(VBA.Err.Number) & " | Description=" & VBA.Err.Description
End Function

Private Sub private_TryUnbindKey(ByVal keySequence As String)
    On Error GoTo EH
    Application.OnKey keySequence
    ex_Core.fn_Diagnostic_WriteLog "PERSONAL_HOTKEY_UNBOUND | Key=" & keySequence
    Exit Sub
EH:
    ex_Core.fn_Diagnostic_WriteLog "PERSONAL_HOTKEY_UNBIND_ERROR | Key=" & _
        keySequence & " | Number=" & VBA.CStr(VBA.Err.Number) & _
        " | Description=" & VBA.Err.Description
End Sub

Private Function private_FormatElapsedMilliseconds(ByVal startedAt As Double) As String
    Dim elapsedSeconds As Double

    elapsedSeconds = VBA.Timer - startedAt
    If elapsedSeconds < 0 Then elapsedSeconds = elapsedSeconds + 86400#
    private_FormatElapsedMilliseconds = VBA.Format$(elapsedSeconds * 1000#, "0.0")
End Function

Private Function private_GlobalMacroRef(ByVal macroName As String) As String
    private_GlobalMacroRef = "'" & Replace$(ThisWorkbook.Name, "'", "''") & _
        "'!" & macroName
End Function