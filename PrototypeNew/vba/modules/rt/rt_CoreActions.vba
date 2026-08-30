Attribute VB_Name = "rt_CoreActions"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
#Const RUNTIME_SNAPSHOTS_ENABLED = False

Private Const LAST_CODE_UPDATE_NAME As String = _
    "__PrototypeLastCodeUpdateAt"

Private g_ScheduledUpdateAt As Date
Private g_ScheduledUpdateMacro As String

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:rt_CoreActions.fn_Module_Dispose"
#End If
    fn_DisposeForLifecycle True
End Sub


Public Sub fn_DisposeForLifecycle(ByVal isWorkbookClosing As Boolean)
    Dim cancelErrorNumber As Long
    Dim cancelErrorDescription As String

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "lifecycle:method-enter " & _
        "method='rt_CoreActions.fn_DisposeForLifecycle' " & _
        "isWorkbookClosing='" & _
        VBA.LCase$(VBA.CStr(isWorkbookClosing)) & "'"
#End If
    If g_ScheduledUpdateAt <= 0# Or _
        VBA.Len(VBA.Trim$(g_ScheduledUpdateMacro)) = 0 Then
        private_ClearScheduledUpdate
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo _
            "lifecycle:method-exit " & _
            "method='rt_CoreActions.fn_DisposeForLifecycle' " & _
            "result='true' reason='no-scheduled-update'"
#End If
        Exit Sub
    End If

    ' Во время Update Code этот callback уже исполняется: Excel удалил его из
    ' очереди OnTime, поэтому отменять его нельзя. На закрытии такого допущения
    ' нет — любой сохранённый callback обязан быть явно снят.
    If Not isWorkbookClosing And g_ScheduledUpdateAt <= VBA.Now Then
        private_ClearScheduledUpdate
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo _
            "lifecycle:method-exit " & _
            "method='rt_CoreActions.fn_DisposeForLifecycle' " & _
            "result='true' reason='callback-already-running'"
#End If
        Exit Sub
    End If

    On Error GoTo EH_CANCEL
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "lifecycle:call-enter " & _
        "caller='rt_CoreActions.fn_DisposeForLifecycle' " & _
        "callee='Application.OnTime:cancel' macro='" & _
        VBA.Replace$(g_ScheduledUpdateMacro, "'", "''") & "'"
#End If
    Application.OnTime _
        EarliestTime:=g_ScheduledUpdateAt, _
        Procedure:=g_ScheduledUpdateMacro, _
        Schedule:=False
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "lifecycle:call-exit " & _
        "caller='rt_CoreActions.fn_DisposeForLifecycle' " & _
        "callee='Application.OnTime:cancel'"
#End If
    private_ClearScheduledUpdate
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "lifecycle:method-exit " & _
        "method='rt_CoreActions.fn_DisposeForLifecycle' result='true'"
#End If
    Exit Sub

EH_CANCEL:
    cancelErrorNumber = Err.Number
    cancelErrorDescription = Err.Description
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError _
        "lifecycle:method-error " & _
        "method='rt_CoreActions.fn_DisposeForLifecycle' " & _
        "callee='Application.OnTime:cancel' errNumber='" & _
        VBA.CStr(cancelErrorNumber) & "' err='" & _
        VBA.Replace$(cancelErrorDescription, "'", "''") & "'"
#End If
    Err.Raise cancelErrorNumber, _
        "rt_CoreActions.fn_DisposeForLifecycle", _
        "Не удалось отменить отложенный Update Code '" & _
        g_ScheduledUpdateMacro & "': " & cancelErrorDescription
End Sub

' //
' // API
' //
Public Sub fn_UpdateCodeFullAndRerender()
    private_QueueSafeCoreUpdate _
        "ex_Core.fn_Dev_UpdateAllModules", "full"
End Sub


Public Sub fn_UpdateCodeDateAndRerender()
    private_QueueSafeCoreUpdate _
        "ex_Core.fn_Dev_UpdateCodeByDate", "date"
End Sub


Public Sub fn_UpdateCodeSizeAndRerender()
    private_QueueSafeCoreUpdate _
        "ex_Core.fn_Dev_UpdateCodeBySize", "size"
End Sub


Public Sub fn_RerenderLastPageAfterUpdate()
    Dim restoredPagesCount As Long

    ex_Core.fn_Dev_MarkRuntimeStateRestoreStarted
    ex_HelpersSheet.fn_SetBusyCursor True
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "core-actions:rerender-after-update start"
#End If

    On Error GoTo EH_RERENDER
    ' Timestamp хранится в workbook Name, поскольку hot-import сбрасывает
    ' статические поля VBA до того, как Main будет создан заново.
    If Not private_TrySaveLastCodeUpdateAt(VBA.Now) Then
        ex_HelpersSheet.fn_SetBusyCursor False
        rt_Messaging.fn_ShowStatusBarError _
            "Failed to save code update timestamp.", 6
        Exit Sub
    End If
#If RUNTIME_SNAPSHOTS_ENABLED Then
    If Not rt_RestoreManager.fn_RestoreRuntimeState("after-update", restoredPagesCount) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "core-actions:rerender-after-update restore-runtime-failed"
#End If
        rt_Messaging.fn_ShowStatusBarError "Failed to restore runtime state after update.", 6
        ex_HelpersSheet.fn_SetBusyCursor False
        Exit Sub
    End If
#Else
    ' Cold open и восстановление после hot-import используют один initialization
    ' pipeline, поэтому routes, hotkeys и logging policy не расходятся.
    If Not rt_Lifecycle.fn_InitializeRuntime( _
        "rt_CoreActions.fn_RerenderLastPageAfterUpdate") Then
        rt_Lifecycle.fn_DisposeRuntime False, _
            "rerender-after-update:initialize-failed"
        ex_HelpersSheet.fn_SetBusyCursor False
        rt_Messaging.fn_ShowStatusBarError _
            "Failed to create Main page after update.", 6
        Exit Sub
    End If
    restoredPagesCount = 1
#End If

    ex_HelpersSheet.fn_SetBusyCursor False
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "core-actions:rerender-after-update done restoredPages=" & VBA.CStr(restoredPagesCount)
#End If
    rt_Messaging.fn_ShowStatusBarSuccess "Update completed. Restored pages: " & VBA.CStr(restoredPagesCount) & ".", 1
    Exit Sub

EH_RERENDER:
    ex_HelpersSheet.fn_SetBusyCursor False
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "core-actions:rerender-after-update exception err='" & VBA.Replace$(Err.Description, "'", "''") & "'"
#End If
    rt_Messaging.fn_ShowStatusBarError "Failed to restore runtime state after update: " & Err.Description, 6
End Sub

Public Function fn_GetLastCodeUpdateCaption() As String
    Dim updateName As Name
    Dim rawValue As String

    On Error Resume Next
    Set updateName = ThisWorkbook.Names(LAST_CODE_UPDATE_NAME)
    On Error GoTo 0
    If updateName Is Nothing Then
        fn_GetLastCodeUpdateCaption = "Код ещё не обновлялся"
        Exit Function
    End If

    rawValue = VBA.CStr(updateName.RefersTo)
    If VBA.Left$(rawValue, 2) = "=""" And _
        VBA.Right$(rawValue, 1) = """" Then
        rawValue = VBA.Mid$(rawValue, 3, VBA.Len(rawValue) - 3)
        rawValue = VBA.Replace$(rawValue, """""", """")
    End If
    rawValue = VBA.Trim$(rawValue)
    If VBA.Len(rawValue) = 0 Then
        fn_GetLastCodeUpdateCaption = "Код ещё не обновлялся"
    Else
        fn_GetLastCodeUpdateCaption = "Код обновлён: " & rawValue
    End If
End Function

Private Function private_TrySaveLastCodeUpdateAt( _
    ByVal updatedAt As Date _
) As Boolean
    Dim timestampText As String
    Dim refersToText As String

    On Error GoTo EH
    timestampText = VBA.Format$(updatedAt, "dd.mm.yyyy HH:nn:ss")
    refersToText = "=""" & _
        VBA.Replace$(timestampText, """", """""") & """"
    ThisWorkbook.Names.Add _
        Name:=LAST_CODE_UPDATE_NAME, _
        RefersTo:=refersToText, _
        Visible:=False
    private_TrySaveLastCodeUpdateAt = True
    Exit Function

EH:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError _
        "core-actions:last-update-save-failed err='" & _
        VBA.Replace$(Err.Description, "'", "''") & "'"
#End If
End Function


' //
' // Internal
' //
Private Function private_GetNextOnTimeTick() As Date
    Dim nowValue As Date

    ' Механизм OnTime в Excel работает с точностью до секунды.
    ' Планирование на текущее время может иногда падать, если округленное значение уже оказалось в прошлом.
    nowValue = VBA.Now
    private_GetNextOnTimeTick = VBA.DateSerial(VBA.Year(nowValue), VBA.Month(nowValue), VBA.Day(nowValue)) + _
                           VBA.TimeSerial(VBA.Hour(nowValue), VBA.Minute(nowValue), VBA.Second(nowValue) + 1)
End Function


Private Sub private_QueueSafeCoreUpdate(ByVal coreMethod As String, ByVal updateKind As String)
    Dim updateMethod As String
    Dim macroRef As String
    Dim scheduleAt As Date
    Dim scheduleErrorNumber As Long
    Dim scheduleErrorDescription As String
    Dim queueStage As String

    updateMethod = VBA.Trim$(coreMethod)
    If VBA.Len(updateMethod) = 0 Then
        rt_Messaging.fn_ShowStatusBarWarning "Safe update method is not specified.", 5
        Exit Sub
    End If

    If VBA.InStr(1, updateMethod, "!", VBA.vbBinaryCompare) > 0 Then
        macroRef = updateMethod
    Else
        macroRef = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!" & updateMethod
    End If
    updateKind = VBA.LCase$(VBA.Trim$(updateKind))
    If VBA.Len(updateKind) = 0 Then updateKind = "unknown"

    scheduleAt = private_GetNextOnTimeTick()

    On Error GoTo EH_QUEUE
    queueStage = "cancel-existing"
    private_CancelPendingUpdate
    queueStage = "schedule"
    Application.OnTime EarliestTime:=scheduleAt, Procedure:=macroRef
    g_ScheduledUpdateAt = scheduleAt
    g_ScheduledUpdateMacro = macroRef
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "core-actions:redirect-safe-update queued kind='" & VBA.Replace$(updateKind, "'", "''") & "' macro='" & VBA.Replace$(macroRef, "'", "''") & "'"
#End If
    rt_Messaging.fn_ShowStatusBarNotice "Update start: safe task has been queued.", 1
    Exit Sub

EH_QUEUE:
    scheduleErrorNumber = Err.Number
    scheduleErrorDescription = Err.Description
    If VBA.StrComp(queueStage, "schedule", VBA.vbBinaryCompare) = 0 Then
        private_ClearScheduledUpdate
    End If
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "core-actions:redirect-safe-update ontime-failed kind='" & VBA.Replace$(updateKind, "'", "''") & "' err='" & VBA.Replace$(scheduleErrorDescription, "'", "''") & "'"
#End If
    rt_Messaging.fn_ShowStatusBarError _
        "Failed to schedule safe update: [" & _
        VBA.CStr(scheduleErrorNumber) & "] " & _
        scheduleErrorDescription, 6
End Sub


Private Sub private_CancelPendingUpdate()
    If g_ScheduledUpdateAt <= 0# Or _
        VBA.Len(VBA.Trim$(g_ScheduledUpdateMacro)) = 0 Then
        private_ClearScheduledUpdate
        Exit Sub
    End If

    Application.OnTime _
        EarliestTime:=g_ScheduledUpdateAt, _
        Procedure:=g_ScheduledUpdateMacro, _
        Schedule:=False
    private_ClearScheduledUpdate
End Sub


Private Sub private_ClearScheduledUpdate()
    g_ScheduledUpdateAt = 0#
    g_ScheduledUpdateMacro = VBA.vbNullString
End Sub
