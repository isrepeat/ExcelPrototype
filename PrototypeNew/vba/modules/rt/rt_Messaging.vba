Attribute VB_Name = "rt_Messaging"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False

' Runtime messaging: живет в rt_* ядре и не переимпортируется при обычном обновлении кода.

Private g_ScheduledHideAt As Date
Private g_ScheduledHideMacro As String
Private g_StatusBarMessage As String

Public Sub m_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:rt_Messaging.m_Module_Dispose"
#End If
End Sub

' //
' // API
' //
Public Sub m_ShowStatusBarFor3s(ByVal messageText As String)
    m_ShowStatusBar messageText, 3
End Sub


Public Sub m_ShowStatusBarNotice(ByVal messageText As String, Optional ByVal timeoutSeconds As Long = 3)
    m_ShowStatusBar VBA.CStr(messageText), timeoutSeconds
End Sub


Public Sub m_ShowStatusBarSuccess(ByVal messageText As String, Optional ByVal timeoutSeconds As Long = 3)
    m_ShowStatusBar "OK: " & VBA.CStr(messageText), timeoutSeconds
End Sub


Public Sub m_ShowStatusBarWarning(ByVal messageText As String, Optional ByVal timeoutSeconds As Long = 3)
    m_ShowStatusBar "Warning: " & VBA.CStr(messageText), timeoutSeconds
End Sub


Public Sub m_ShowStatusBarError(ByVal messageText As String, Optional ByVal timeoutSeconds As Long = 3)
    m_ShowStatusBar "Error: " & VBA.CStr(messageText), timeoutSeconds
End Sub


Public Sub m_ShowStatusBar(ByVal messageText As String, Optional ByVal timeoutSeconds As Long = 3)
    Dim hideMacroRef As String
    Dim hideAt As Date
    Dim wbMacroPrefix As String

    messageText = VBA.Trim$(messageText)
    If VBA.Len(messageText) = 0 Then
        m_HideStatusBarNow
        Exit Sub
    End If

    If timeoutSeconds <= 0 Then timeoutSeconds = 3

    g_StatusBarMessage = VBA.CStr(messageText)
    private_ApplyNativeStatusBar
    private_LogStatusBarMessage "show", g_StatusBarMessage, timeoutSeconds

    private_TryCancelScheduledHide

    wbMacroPrefix = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!"
    hideMacroRef = wbMacroPrefix & "rt_Messaging.m_HideStatusBarScheduled"
    hideAt = VBA.DateAdd("s", timeoutSeconds, private_GetNextOnTimeTick())

    On Error GoTo EH_SCHEDULE
    Application.OnTime hideAt, hideMacroRef
    g_ScheduledHideAt = hideAt
    g_ScheduledHideMacro = hideMacroRef
    Exit Sub

EH_SCHEDULE:
    g_ScheduledHideAt = 0#
    g_ScheduledHideMacro = VBA.vbNullString
    g_StatusBarMessage = "Error: Failed to schedule status bar hide: " & Err.Description
    private_ApplyNativeStatusBar
    private_LogStatusBarMessage "schedule-failed", g_StatusBarMessage
End Sub


Public Sub m_HideStatusBarNow()
    private_TryCancelScheduledHide
    g_StatusBarMessage = VBA.vbNullString
    private_ApplyNativeStatusBar
    private_LogStatusBarMessage "hide-now", VBA.vbNullString
End Sub


Public Sub m_HideStatusBarScheduled()
    g_ScheduledHideAt = 0#
    g_ScheduledHideMacro = VBA.vbNullString
    g_StatusBarMessage = VBA.vbNullString
    private_ApplyNativeStatusBar
    private_LogStatusBarMessage "hide-scheduled", VBA.vbNullString
End Sub

' Callstack[1]: VBA.ImmediateWindow -> rt_Messaging.m_TryGetStatusBarMessage
Public Function m_TryGetStatusBarMessage(ByRef outMessage As String) As Boolean
    outMessage = g_StatusBarMessage
    m_TryGetStatusBarMessage = True
End Function

' //
' // Internal
' //
Private Sub private_ApplyNativeStatusBar()
    If VBA.Len(VBA.Trim$(g_StatusBarMessage)) = 0 Then
        Application.StatusBar = False
        Exit Sub
    End If
    On Error Resume Next
    Application.StatusBar = "PrototypeNew: " & g_StatusBarMessage
    Err.Clear
    On Error GoTo 0
End Sub


Private Sub private_TryCancelScheduledHide()
    If g_ScheduledHideAt <= 0# Then Exit Sub
    If VBA.Len(VBA.Trim$(g_ScheduledHideMacro)) = 0 Then Exit Sub

    On Error Resume Next
    Application.OnTime EarliestTime:=g_ScheduledHideAt, Procedure:=g_ScheduledHideMacro, Schedule:=False
    Err.Clear
    On Error GoTo 0

    g_ScheduledHideAt = 0#
    g_ScheduledHideMacro = VBA.vbNullString
End Sub


Private Sub private_LogStatusBarMessage( _
    ByVal actionName As String, _
    ByVal messageText As String, _
    Optional ByVal timeoutSeconds As Long = 0 _
)
    On Error Resume Next
    ex_Core.m_Diagnostic_LogStatusBarMessage actionName, messageText, timeoutSeconds
    Err.Clear
    On Error GoTo 0
End Sub


Private Function private_GetNextOnTimeTick() As Date
    Dim nowValue As Date

    nowValue = VBA.Now
    private_GetNextOnTimeTick = VBA.DateSerial(VBA.Year(nowValue), VBA.Month(nowValue), VBA.Day(nowValue)) + _
                           VBA.TimeSerial(VBA.Hour(nowValue), VBA.Minute(nowValue), VBA.Second(nowValue) + 1)
End Function
