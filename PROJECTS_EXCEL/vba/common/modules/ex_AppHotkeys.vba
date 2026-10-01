Option Explicit

Private Const DIAGNOSTIC_MODE_IMMEDIATE As String = "Immediate"
Private Const DIAGNOSTIC_MODE_BUFFERED As String = "Buffered"
Private Const DIAGNOSTIC_MODE_HOTKEY As String = "^%l"
Private Const PERSONAL_WORKBOOK_NAME As String = "PERSONAL.XLSB"
Private Const PERSONAL_RESTORE_MACRO As String = "ex_Core.fn_RestorePersonalHotkeys"
Private m_hotkeyHandlers As Object
Private m_isActive As Boolean
Private m_isBrokerManaged As Boolean

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_PrepareReload()
    Dim keySequence As Variant
    If m_isBrokerManaged Then
        If Not private_TryDeactivateThroughPersonal() Then _
            Err.Raise vbObjectError + 2210, "fn_PrepareReload", "Could not detach the PERSONAL hotkey broker."
    ElseIf Not m_hotkeyHandlers Is Nothing Then
        For Each keySequence In m_hotkeyHandlers.Keys
            Application.OnKey VBA.CStr(keySequence)
        Next keySequence
    End If
    m_isActive = False
    m_isBrokerManaged = False
End Sub

Public Sub fn_Module_Dispose()
    fn_Deactivate
    Set m_hotkeyHandlers = Nothing
    m_isBrokerManaged = False
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Register( _
    ByVal keySequence As String, _
    ByVal macroName As String _
) As Boolean
    keySequence = VBA.Trim$(keySequence)
    macroName = VBA.Trim$(macroName)
    If VBA.Len(keySequence) = 0 Or VBA.Len(macroName) = 0 Then Exit Function

    private_EnsureRegistry
    m_hotkeyHandlers(keySequence) = macroName
    If Not m_isActive Then
        fn_Register = True
    ElseIf m_isBrokerManaged Then
        fn_Register = private_TryActivateThroughPersonal
    Else
        fn_Register = private_TryBind(keySequence, macroName)
    End If
End Function

Public Sub fn_Activate()
    Dim keySequence As Variant

    private_EnsureRegistry
    If private_TryActivateThroughPersonal Then
        m_isBrokerManaged = True
        m_isActive = True
        ex_Core.fn_Diagnostic_WriteLog "HOTKEY_BROKER_ACTIVATED | KeyCount=" & _
            VBA.CStr(m_hotkeyHandlers.Count)
        Exit Sub
    End If
    For Each keySequence In m_hotkeyHandlers.Keys
        private_TryBind VBA.CStr(keySequence), VBA.CStr(m_hotkeyHandlers(keySequence))
    Next keySequence
    m_isBrokerManaged = False
    m_isActive = True
    ex_Core.fn_Diagnostic_WriteLog "HOTKEY_LOCAL_ACTIVATED | KeyCount=" & _
        VBA.CStr(m_hotkeyHandlers.Count)
End Sub

Public Sub fn_Deactivate()
    Dim keySequence As Variant

    If m_isBrokerManaged Then
        private_TryDeactivateThroughPersonal
        ex_Core.fn_Diagnostic_WriteLog "HOTKEY_BROKER_DEACTIVATED"
    ElseIf Not m_hotkeyHandlers Is Nothing Then
        On Error Resume Next
        For Each keySequence In m_hotkeyHandlers.Keys
            Application.OnKey VBA.CStr(keySequence)
        Next keySequence
        On Error GoTo 0
        private_TryRestorePersonalBindings
        ex_Core.fn_Diagnostic_WriteLog "HOTKEY_LOCAL_DEACTIVATED"
    End If
    m_isActive = False
    m_isBrokerManaged = False
End Sub

Public Sub fn_ToggleDiagnosticMode()
    Dim runtimeContext As Object
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim currentMode As String
    Dim nextMode As String

    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    On Error GoTo EH_TOGGLE
    currentMode = ex_Core.fn_Diagnostic_GetMode()
    If VBA.StrComp(currentMode, DIAGNOSTIC_MODE_BUFFERED, VBA.vbTextCompare) = 0 Then
        nextMode = DIAGNOSTIC_MODE_IMMEDIATE
    Else
        nextMode = DIAGNOSTIC_MODE_BUFFERED
    End If

    If Not ex_Core.fn_Diagnostic_SetMode(nextMode) Then
        VBA.MsgBox "Could not change diagnostic logging mode.", _
            VBA.vbExclamation, "Diagnostic logging"
        GoTo CleanToggle
    End If
    ex_Core.fn_Diagnostic_WriteLog "DIAGNOSTIC_MODE_CHANGED | Mode=" & nextMode
    VBA.MsgBox "Diagnostic logging mode: " & nextMode, _
        VBA.vbInformation, "Diagnostic logging"
CleanToggle:
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH_TOGGLE:
    errorNumber = Err.Number
    errorDescription = Err.Description
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Err.Raise errorNumber, "fn_ToggleDiagnosticMode", errorDescription
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Sub private_EnsureRegistry()
    If m_hotkeyHandlers Is Nothing Then
        Set m_hotkeyHandlers = VBA.CreateObject("Scripting.Dictionary")
        m_hotkeyHandlers.CompareMode = VBA.vbTextCompare
        m_hotkeyHandlers(DIAGNOSTIC_MODE_HOTKEY) = _
            "ex_AppHotkeys.fn_ToggleDiagnosticMode"
    End If
End Sub

Private Function private_TryBind( _
    ByVal keySequence As String, _
    ByVal macroName As String _
) As Boolean
    Dim macroReference As String

    On Error GoTo EH
    macroReference = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & _
        "'!" & macroName
    Application.OnKey keySequence, macroReference
    private_TryBind = True
EH:
End Function

Private Function private_TryActivateThroughPersonal() As Boolean
    Dim macroReference As String
    Dim resultValue As Variant

    If m_hotkeyHandlers Is Nothing Then Exit Function
    On Error GoTo EH
    macroReference = "'" & PERSONAL_WORKBOOK_NAME & "'!" & _
        "ex_Core.fn_HotkeyBroker_Activate"
    resultValue = Application.Run( _
        macroReference, ThisWorkbook.FullName, private_SerializeBindings)
    private_TryActivateThroughPersonal = VBA.CBool(resultValue)
    Exit Function
EH:
End Function

Private Function private_TryDeactivateThroughPersonal() As Boolean
    Dim macroReference As String
    Dim resultValue As Variant

    On Error GoTo EH
    macroReference = "'" & PERSONAL_WORKBOOK_NAME & "'!" & _
        "ex_Core.fn_HotkeyBroker_Deactivate"
    resultValue = Application.Run(macroReference, ThisWorkbook.FullName)
    private_TryDeactivateThroughPersonal = VBA.CBool(resultValue)
    Exit Function
EH:
End Function

Private Sub private_TryRestorePersonalBindings()
    Dim targetWorkbook As Workbook

    For Each targetWorkbook In Application.Workbooks
        If VBA.StrComp(targetWorkbook.Name, PERSONAL_WORKBOOK_NAME, _
                VBA.vbTextCompare) = 0 Then
            On Error Resume Next
            Application.Run "'" & PERSONAL_WORKBOOK_NAME & "'!" & _
                PERSONAL_RESTORE_MACRO
            On Error GoTo 0
            Exit Sub
        End If
    Next targetWorkbook
End Sub

Private Function private_SerializeBindings() As String
    Dim keySequence As Variant
    Dim resultText As String

    For Each keySequence In m_hotkeyHandlers.Keys
        If VBA.Len(resultText) > 0 Then resultText = resultText & VBA.ChrW$(31)
        resultText = resultText & VBA.CStr(keySequence) & VBA.ChrW$(30) & _
            VBA.CStr(m_hotkeyHandlers(keySequence))
    Next keySequence
    private_SerializeBindings = resultText
End Function
' --------------------------------------
' } // namespace Private
' --------------------------------------