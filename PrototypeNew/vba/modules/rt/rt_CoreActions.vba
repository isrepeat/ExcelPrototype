Attribute VB_Name = "rt_CoreActions"
Option Explicit

Private g_ScheduledUpdateAt As Date
Private g_ScheduledUpdateMacro As String
Private g_PendingUpdateMacroRef As String
Private g_IsRunningScheduledUpdate As Boolean

' //
' // API
' //
Public Sub m_UpdateCodeFullAndRerender()
    private_QueueSafeCoreUpdate "ex_Core.m_Dev_UpdateAllModules", "full"
End Sub


Public Sub m_UpdateCodeDateAndRerender()
    private_QueueSafeCoreUpdate "ex_Core.m_Dev_UpdateCodeByDate", "date"
End Sub


Public Sub m_UpdateCodeSizeAndRerender()
    private_QueueSafeCoreUpdate "ex_Core.m_Dev_UpdateCodeBySize", "size"
End Sub


Public Sub m_RerenderLastPageAfterUpdate()
    Dim restoredPagesCount As Long

    ex_HelpersSheet.m_SetBusyCursor True
    ex_Core.m_Diagnostic_LogInfo "core-actions:rerender-after-update start"

    On Error GoTo EH_RERENDER
    If Not rt_Snapshots.m_RestorePageSnapshots(True, "after-update", restoredPagesCount) Then
        ex_Core.m_Diagnostic_LogError "core-actions:rerender-after-update restore-pages-failed"
        rt_Messaging.m_ShowStatusBarError "Failed to restore pages after update.", 6
        ex_HelpersSheet.m_SetBusyCursor False
        Exit Sub
    End If

    If Not rt_Snapshots.m_RestoreRuntimeGlobalsSnapshot() Then
        ex_Core.m_Diagnostic_LogError "core-actions:rerender-after-update restore-runtime-globals-failed"
        rt_Messaging.m_ShowStatusBarError "Failed to restore runtime globals after update.", 6
        ex_HelpersSheet.m_SetBusyCursor False
        Exit Sub
    End If

    ex_HelpersSheet.m_SetBusyCursor False
    ex_Core.m_Diagnostic_LogInfo "core-actions:rerender-after-update done restoredPages=" & VBA.CStr(restoredPagesCount)
    rt_Messaging.m_ShowStatusBarSuccess "Update completed. Restored pages: " & VBA.CStr(restoredPagesCount) & "; runtime state refreshed.", 1
    Exit Sub

EH_RERENDER:
    ex_HelpersSheet.m_SetBusyCursor False
    ex_Core.m_Diagnostic_LogError "core-actions:rerender-after-update exception err='" & VBA.Replace$(Err.Description, "'", "''") & "'"
    rt_Messaging.m_ShowStatusBarError "Failed to restore runtime state after update: " & Err.Description, 6
End Sub


Public Sub m_RunScheduledUpdateAndRerender()
    Dim updateMacroRef As String

    If g_IsRunningScheduledUpdate Then
        ex_Core.m_Diagnostic_LogError "core-actions:run-scheduled reentry-blocked"
        Exit Sub
    End If

    g_ScheduledUpdateAt = 0#
    g_ScheduledUpdateMacro = VBA.vbNullString

    updateMacroRef = VBA.Trim$(g_PendingUpdateMacroRef)
    g_PendingUpdateMacroRef = VBA.vbNullString

    If VBA.Len(updateMacroRef) = 0 Then
        ex_HelpersSheet.m_SetBusyCursor False
        rt_Messaging.m_ShowStatusBarWarning "Scheduled update method was not found.", 5
        Exit Sub
    End If

    g_IsRunningScheduledUpdate = True
    On Error GoTo EH_RUN
    Application.Run updateMacroRef
    m_RerenderLastPageAfterUpdate
    g_IsRunningScheduledUpdate = False
    Exit Sub

EH_RUN:
    g_IsRunningScheduledUpdate = False
    ex_HelpersSheet.m_SetBusyCursor False
    rt_Messaging.m_ShowStatusBarError "Failed to run update: " & Err.Description, 6
End Sub

' //
' // Internal
' //
Private Sub private_ScheduleUpdateAndRerender(ByVal devToolsMethod As String)
    Dim updateMacroRef As String
    Dim wbMacroPrefix As String
    Dim updateAt As Date
    Dim updateMethod As String

    ' Snapshot-сценарий:
    ' 1) сериализуем страницы (включая runtime контролов внутри payload страницы) в CustomXMLPart;
    ' 2) запускаем update;
    ' 3) после update восстанавливаем страницы и их runtime-состояние.

    updateMethod = VBA.Trim$(devToolsMethod)
    If VBA.Len(updateMethod) = 0 Then
        rt_Messaging.m_ShowStatusBarWarning "Update method is not specified.", 5
        Exit Sub
    End If
    If g_IsRunningScheduledUpdate Then
        ex_Core.m_Diagnostic_LogError "core-actions:schedule-update skipped reason='update-is-running' method='" & VBA.Replace$(updateMethod, "'", "''") & "'"
        Exit Sub
    End If
    ex_Core.m_Diagnostic_LogInfo "core-actions:schedule-update start method='" & VBA.Replace$(updateMethod, "'", "''") & "'"

    ex_HelpersSheet.m_SetBusyCursor True

    If Not rt_Snapshots.m_SavePageSnapshots() Then
        ex_Core.m_Diagnostic_LogError "core-actions:schedule-update save-page-snapshots-failed"
        ex_HelpersSheet.m_SetBusyCursor False
        Exit Sub
    End If

    If Not rt_Snapshots.m_SaveRuntimeGlobalsSnapshot() Then
        ex_Core.m_Diagnostic_LogError "core-actions:schedule-update save-runtime-globals-failed"
        ex_HelpersSheet.m_SetBusyCursor False
        Exit Sub
    End If

    wbMacroPrefix = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!"
    If VBA.InStr(1, updateMethod, "!", VBA.vbBinaryCompare) > 0 Then
        g_PendingUpdateMacroRef = updateMethod
    Else
        g_PendingUpdateMacroRef = wbMacroPrefix & updateMethod
    End If
    updateMacroRef = wbMacroPrefix & "rt_CoreActions.m_RunScheduledUpdateAndRerender"

    ' Если пользователь часто кликает подряд, отменяем отложенные старые задачи.
    private_TryCancelScheduledTask g_ScheduledUpdateAt, g_ScheduledUpdateMacro

    updateAt = private_GetNextOnTimeTick()

    On Error GoTo EH_SCHEDULE
    Application.OnTime updateAt, updateMacroRef
    g_ScheduledUpdateAt = updateAt
    g_ScheduledUpdateMacro = updateMacroRef
    ex_Core.m_Diagnostic_LogInfo "core-actions:schedule-update queued macro='" & VBA.Replace$(updateMacroRef, "'", "''") & "'"
    rt_Messaging.m_ShowStatusBarNotice "Update start: task has been queued.", 1
    Exit Sub

EH_SCHEDULE:
    ex_HelpersSheet.m_SetBusyCursor False
    g_PendingUpdateMacroRef = VBA.vbNullString
    ex_Core.m_Diagnostic_LogError "core-actions:schedule-update ontime-failed err='" & VBA.Replace$(Err.Description, "'", "''") & "'"
    rt_Messaging.m_ShowStatusBarError "Failed to schedule update: " & Err.Description, 6
End Sub

Private Function private_GetNextOnTimeTick() As Date
    Dim nowValue As Date

    ' Механизм OnTime в Excel работает с точностью до секунды.
    ' Планирование на текущее время может иногда падать, если округленное значение уже оказалось в прошлом.
    nowValue = VBA.Now
    private_GetNextOnTimeTick = VBA.DateSerial(VBA.Year(nowValue), VBA.Month(nowValue), VBA.Day(nowValue)) + _
                           VBA.TimeSerial(VBA.Hour(nowValue), VBA.Minute(nowValue), VBA.Second(nowValue) + 1)
End Function


Private Sub private_TryCancelScheduledTask(ByVal scheduledAt As Date, ByVal macroRef As String)
    If VBA.Len(VBA.Trim$(macroRef)) = 0 Then Exit Sub
    If scheduledAt <= 0# Then Exit Sub

    On Error Resume Next
    Application.OnTime EarliestTime:=scheduledAt, Procedure:=macroRef, Schedule:=False
    Err.Clear
    On Error GoTo 0
End Sub


Private Sub private_QueueSafeCoreUpdate(ByVal coreMethod As String, ByVal updateKind As String)
    Dim updateMethod As String
    Dim macroRef As String
    Dim scheduleAt As Date

    updateMethod = VBA.Trim$(coreMethod)
    If VBA.Len(updateMethod) = 0 Then
        rt_Messaging.m_ShowStatusBarWarning "Safe update method is not specified.", 5
        Exit Sub
    End If

    If VBA.InStr(1, updateMethod, "!", VBA.vbBinaryCompare) > 0 Then
        macroRef = updateMethod
    Else
        macroRef = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!" & updateMethod
    End If
    updateKind = VBA.LCase$(VBA.Trim$(updateKind))
    If VBA.Len(updateKind) = 0 Then updateKind = "unknown"

    private_TryCancelScheduledTask g_ScheduledUpdateAt, g_ScheduledUpdateMacro
    scheduleAt = private_GetNextOnTimeTick()

    On Error GoTo EH_QUEUE
    Application.OnTime EarliestTime:=scheduleAt, Procedure:=macroRef
    g_ScheduledUpdateAt = scheduleAt
    g_ScheduledUpdateMacro = macroRef
    ex_Core.m_Diagnostic_LogInfo "core-actions:redirect-safe-update queued kind='" & VBA.Replace$(updateKind, "'", "''") & "' macro='" & VBA.Replace$(macroRef, "'", "''") & "'"
    rt_Messaging.m_ShowStatusBarNotice "Update start: safe task has been queued.", 1
    Exit Sub

EH_QUEUE:
    ex_Core.m_Diagnostic_LogError "core-actions:redirect-safe-update ontime-failed kind='" & VBA.Replace$(updateKind, "'", "''") & "' err='" & VBA.Replace$(Err.Description, "'", "''") & "'"
    rt_Messaging.m_ShowStatusBarError "Failed to schedule safe update: " & Err.Description, 6
End Sub
