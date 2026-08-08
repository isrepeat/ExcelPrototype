Attribute VB_Name = "rt_Messaging"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False

' Runtime messaging: живет в rt_* ядре и не переимпортируется при обычном обновлении кода.

Private g_StatusBarMessage As String
Private g_ScheduledHideAt As Date
Private g_ScheduledHideMacro As String

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:rt_Messaging.fn_Module_Dispose"
#End If
    ' При hot-update обязательно снимаем отложенный hide,
    ' чтобы OnTime не стрелял в момент remove/import модулей.
    fn_CancelDeferredTasks
    g_StatusBarMessage = VBA.vbNullString
    On Error Resume Next
    Application.StatusBar = False
    Err.Clear
    On Error GoTo 0
End Sub


Public Sub fn_CancelDeferredTasks()
    private_CancelScheduledHide
End Sub

' //
' // API
' //
Public Sub fn_ShowStatusBarFor3s(ByVal messageText As String)
    fn_ShowStatusBar messageText, 3
End Sub


Public Sub fn_ShowStatusBarNotice(ByVal messageText As String, Optional ByVal timeoutSeconds As Long = 3)
    fn_ShowStatusBar VBA.CStr(messageText), timeoutSeconds
End Sub


Public Sub fn_ShowStatusBarSuccess(ByVal messageText As String, Optional ByVal timeoutSeconds As Long = 3)
    fn_ShowStatusBar "OK: " & VBA.CStr(messageText), timeoutSeconds
End Sub


Public Sub fn_ShowStatusBarWarning(ByVal messageText As String, Optional ByVal timeoutSeconds As Long = 3)
    fn_ShowStatusBar "Warning: " & VBA.CStr(messageText), timeoutSeconds
End Sub


Public Sub fn_ShowStatusBarError(ByVal messageText As String, Optional ByVal timeoutSeconds As Long = 3)
    fn_ShowStatusBar "Error: " & VBA.CStr(messageText), timeoutSeconds
End Sub

Public Sub fn_ShowStatusBarProgress( _
    ByVal processCaption As String, _
    ByVal processedCount As Long, _
    ByVal totalCount As Long _
)
    Dim progressPercent As Long

    If totalCount <= 0 Then
        progressPercent = 0
    Else
        progressPercent = VBA.CLng( _
            (VBA.CDbl(processedCount) / VBA.CDbl(totalCount)) * 100#)
    End If
    If progressPercent < 0 Then progressPercent = 0
    If progressPercent > 100 Then progressPercent = 100

    private_CancelScheduledHide
    g_StatusBarMessage = VBA.Trim$(processCaption) & ": " & _
        VBA.CStr(progressPercent) & "%"
    private_ApplyNativeStatusBar
    private_LogStatusBarMessage "progress", g_StatusBarMessage, 0
End Sub


Public Sub fn_ShowStatusBar(ByVal messageText As String, Optional ByVal timeoutSeconds As Long = 3)
    Dim hideMacroRef As String
    Dim hideAt As Date
    Dim wbMacroPrefix As String
    Dim scheduleErrorNumber As Long
    Dim scheduleErrorDescription As String

    messageText = VBA.Trim$(messageText)
    If VBA.Len(messageText) = 0 Then
        fn_HideStatusBarNow
        Exit Sub
    End If

    If timeoutSeconds <= 0 Then timeoutSeconds = 3

    g_StatusBarMessage = VBA.CStr(messageText)
    private_ApplyNativeStatusBar
    private_LogStatusBarMessage "show", g_StatusBarMessage, timeoutSeconds

    private_CancelScheduledHide

    wbMacroPrefix = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!"
    hideMacroRef = wbMacroPrefix & "rt_Messaging.fn_HideStatusBarScheduled"
    hideAt = VBA.DateAdd("s", timeoutSeconds, private_GetNextOnTimeTick())

    On Error GoTo EH_SCHEDULE
    Application.OnTime EarliestTime:=hideAt, Procedure:=hideMacroRef
    g_ScheduledHideAt = hideAt
    g_ScheduledHideMacro = hideMacroRef
    Exit Sub

EH_SCHEDULE:
    scheduleErrorNumber = Err.Number
    scheduleErrorDescription = Err.Description
    private_ClearScheduledHide
    g_StatusBarMessage = _
        "Error: Failed to schedule status bar hide: [" & _
        VBA.CStr(scheduleErrorNumber) & "] " & scheduleErrorDescription
    private_ApplyNativeStatusBar
    private_LogStatusBarMessage "schedule-failed", g_StatusBarMessage
End Sub


Public Sub fn_HideStatusBarNow()
    private_CancelScheduledHide
    g_StatusBarMessage = VBA.vbNullString
    private_ApplyNativeStatusBar
    private_LogStatusBarMessage "hide-now", VBA.vbNullString
End Sub


Public Sub fn_HideStatusBarScheduled()
    ' Callback уже удалён Excel из очереди OnTime до входа в процедуру.
    private_ClearScheduledHide
    g_StatusBarMessage = VBA.vbNullString
    private_ApplyNativeStatusBar
    private_LogStatusBarMessage "hide-scheduled", VBA.vbNullString
End Sub

' Callstack[1]: VBA.ImmediateWindow -> rt_Messaging.fn_TryGetStatusBarMessage
Public Function fn_TryGetStatusBarMessage(ByRef outMessage As String) As Boolean
    outMessage = g_StatusBarMessage
    fn_TryGetStatusBarMessage = True
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


Private Sub private_CancelScheduledHide()
    Dim cancelErrorNumber As Long
    Dim cancelErrorDescription As String

    If g_ScheduledHideAt <= 0# Or _
        VBA.Len(VBA.Trim$(g_ScheduledHideMacro)) = 0 Then
        private_ClearScheduledHide
        Exit Sub
    End If

    On Error GoTo EH_CANCEL
    Application.OnTime _
        EarliestTime:=g_ScheduledHideAt, _
        Procedure:=g_ScheduledHideMacro, _
        Schedule:=False
    private_ClearScheduledHide
    Exit Sub

EH_CANCEL:
    cancelErrorNumber = Err.Number
    cancelErrorDescription = Err.Description
    Err.Raise cancelErrorNumber, _
        "rt_Messaging.private_CancelScheduledHide", _
        "Не удалось отменить отложенное скрытие status bar '" & _
        g_ScheduledHideMacro & "': " & cancelErrorDescription

End Sub


Private Sub private_ClearScheduledHide()
    g_ScheduledHideAt = 0#
    g_ScheduledHideMacro = VBA.vbNullString
End Sub


Private Sub private_LogStatusBarMessage( _
    ByVal actionName As String, _
    ByVal messageText As String, _
    Optional ByVal timeoutSeconds As Long = 0 _
)
    On Error Resume Next
    ex_Core.fn_Diagnostic_LogStatusBarMessage actionName, messageText, timeoutSeconds
    Err.Clear
    On Error GoTo 0
End Sub


Private Function private_GetNextOnTimeTick() As Date
    Dim nowValue As Date

    nowValue = VBA.Now
    private_GetNextOnTimeTick = VBA.DateSerial(VBA.Year(nowValue), VBA.Month(nowValue), VBA.Day(nowValue)) + _
                           VBA.TimeSerial(VBA.Hour(nowValue), VBA.Minute(nowValue), VBA.Second(nowValue) + 1)
End Function
