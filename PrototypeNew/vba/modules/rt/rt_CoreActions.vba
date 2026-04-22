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
    private_QueueSafeCoreUpdate "ex_Core.dev_UpdateAllModules", "full"
End Sub


Public Sub m_UpdateCodeDateAndRerender()
    private_QueueSafeCoreUpdate "ex_Core.dev_UpdateCodeByDate", "date"
End Sub


Public Sub m_UpdateCodeSizeAndRerender()
    private_QueueSafeCoreUpdate "ex_Core.dev_UpdateCodeBySize", "size"
End Sub


Public Sub m_RerenderLastPageAfterUpdate()
    Dim restoredPagesCount As Long

    ex_HelpersSheet.m_SetBusyCursor True
    ex_Core.m_LogInfo "core-actions:rerender-after-update start"

    On Error GoTo EH_RERENDER
    If Not rt_Snapshots.m_RestorePageSnapshots(True, "after-update", restoredPagesCount) Then
        ex_Core.m_LogError "core-actions:rerender-after-update restore-pages-failed"
        rt_Messaging.m_ShowStatusBarError "Не удалось восстановить страницы после обновления.", 6
        ex_HelpersSheet.m_SetBusyCursor False
        Exit Sub
    End If

    If Not rt_Snapshots.m_RestoreRuntimeGlobalsSnapshot() Then
        ex_Core.m_LogError "core-actions:rerender-after-update restore-runtime-globals-failed"
        rt_Messaging.m_ShowStatusBarError "Не удалось восстановить runtime-глобалы после обновления.", 6
        ex_HelpersSheet.m_SetBusyCursor False
        Exit Sub
    End If

    ex_HelpersSheet.m_SetBusyCursor False
    ex_Core.m_LogInfo "core-actions:rerender-after-update done restoredPages=" & VBA.CStr(restoredPagesCount)
    rt_Messaging.m_ShowStatusBarSuccess "Обновление завершено. Восстановлено страниц: " & VBA.CStr(restoredPagesCount) & "; runtime-состояние актуализировано.", 1
    Exit Sub

EH_RERENDER:
    ex_HelpersSheet.m_SetBusyCursor False
    ex_Core.m_LogError "core-actions:rerender-after-update exception err='" & VBA.Replace$(Err.Description, "'", "''") & "'"
    rt_Messaging.m_ShowStatusBarError "Ошибка восстановления runtime-состояния после обновления: " & Err.Description, 6
End Sub


Public Sub m_RunScheduledUpdateAndRerender()
    Dim updateMacroRef As String

    If g_IsRunningScheduledUpdate Then
        ex_Core.m_LogError "core-actions:run-scheduled reentry-blocked"
        Exit Sub
    End If

    g_ScheduledUpdateAt = 0#
    g_ScheduledUpdateMacro = VBA.vbNullString

    updateMacroRef = VBA.Trim$(g_PendingUpdateMacroRef)
    g_PendingUpdateMacroRef = VBA.vbNullString
    updateMacroRef = private_NormalizeUpdateMacroRef(updateMacroRef)

    If VBA.Len(updateMacroRef) = 0 Then
        ex_HelpersSheet.m_SetBusyCursor False
        rt_Messaging.m_ShowStatusBarWarning "Не найден запланированный метод обновления.", 5
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
    rt_Messaging.m_ShowStatusBarError "Ошибка запуска обновления: " & Err.Description, 6
End Sub

' //
' // Internal
' //
Private Sub private_ScheduleUpdateAndRerender(ByVal devToolsMethod As String)
    Dim updateMacroRef As String
    Dim wbMacroPrefix As String
    Dim updateAt As Date
    Dim normalizedMethod As String

    ' Snapshot-сценарий:
    ' 1) сериализуем страницы (включая runtime контролов внутри payload страницы) в CustomXMLPart;
    ' 2) запускаем update;
    ' 3) после update восстанавливаем страницы и их runtime-состояние.

    devToolsMethod = VBA.Trim$(devToolsMethod)
    normalizedMethod = private_NormalizeUpdateMethodName(devToolsMethod)
    If VBA.Len(normalizedMethod) = 0 Then
        rt_Messaging.m_ShowStatusBarWarning "Не указан метод обновления.", 5
        Exit Sub
    End If
    If g_IsRunningScheduledUpdate Then
        ex_Core.m_LogError "core-actions:schedule-update skipped reason='update-is-running' method='" & VBA.Replace$(normalizedMethod, "'", "''") & "'"
        Exit Sub
    End If
    ex_Core.m_LogInfo "core-actions:schedule-update start method='" & VBA.Replace$(normalizedMethod, "'", "''") & "'"

    ex_HelpersSheet.m_SetBusyCursor True

    If Not rt_Snapshots.m_SavePageSnapshots() Then
        ex_Core.m_LogError "core-actions:schedule-update save-page-snapshots-failed"
        ex_HelpersSheet.m_SetBusyCursor False
        Exit Sub
    End If

    If Not rt_Snapshots.m_SaveRuntimeGlobalsSnapshot() Then
        ex_Core.m_LogError "core-actions:schedule-update save-runtime-globals-failed"
        ex_HelpersSheet.m_SetBusyCursor False
        Exit Sub
    End If

    wbMacroPrefix = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!"
    g_PendingUpdateMacroRef = wbMacroPrefix & normalizedMethod
    updateMacroRef = wbMacroPrefix & "rt_CoreActions.m_RunScheduledUpdateAndRerender"

    ' Если пользователь часто кликает подряд, отменяем отложенные старые задачи.
    private_TryCancelScheduledTask g_ScheduledUpdateAt, g_ScheduledUpdateMacro

    updateAt = private_GetNextOnTimeTick()

    On Error GoTo EH_SCHEDULE
    Application.OnTime updateAt, updateMacroRef
    g_ScheduledUpdateAt = updateAt
    g_ScheduledUpdateMacro = updateMacroRef
    ex_Core.m_LogInfo "core-actions:schedule-update queued macro='" & VBA.Replace$(updateMacroRef, "'", "''") & "'"
    rt_Messaging.m_ShowStatusBarNotice "Начало обновления: задача поставлена в очередь.", 1
    Exit Sub

EH_SCHEDULE:
    ex_HelpersSheet.m_SetBusyCursor False
    g_PendingUpdateMacroRef = VBA.vbNullString
    ex_Core.m_LogError "core-actions:schedule-update ontime-failed err='" & VBA.Replace$(Err.Description, "'", "''") & "'"
    rt_Messaging.m_ShowStatusBarError "Не удалось запланировать обновление: " & Err.Description, 6
End Sub


Private Function private_NormalizeUpdateMethodName(ByVal methodName As String) As String
    methodName = VBA.Trim$(methodName)
    If VBA.Len(methodName) = 0 Then Exit Function

    Select Case VBA.LCase$(methodName)
        Case "ex_core.dev_updateallmodules"
            private_NormalizeUpdateMethodName = "ex_Core.dev_UpdateAllModules"
        Case "ex_core.dev_updateallmodulesunsafe"
            private_NormalizeUpdateMethodName = "ex_Core.dev_UpdateAllModules"
        Case "ex_core.dev_updatecodebydate"
            private_NormalizeUpdateMethodName = "ex_Core.dev_UpdateCodeByDate"
        Case "ex_core.dev_updatecodebydateunsafe"
            private_NormalizeUpdateMethodName = "ex_Core.dev_UpdateCodeByDate"
        Case "ex_core.dev_updatecodebysize"
            private_NormalizeUpdateMethodName = "ex_Core.dev_UpdateCodeBySize"
        Case "ex_core.dev_updatecodebysizeunsafe"
            private_NormalizeUpdateMethodName = "ex_Core.dev_UpdateCodeBySize"
        Case Else
            private_NormalizeUpdateMethodName = methodName
    End Select
End Function


Private Function private_NormalizeUpdateMacroRef(ByVal macroRef As String) As String
    Dim bangPos As Long
    Dim workbookPrefix As String
    Dim methodName As String

    macroRef = VBA.Trim$(macroRef)
    If VBA.Len(macroRef) = 0 Then Exit Function

    bangPos = VBA.InStr(1, macroRef, "!", VBA.vbBinaryCompare)
    If bangPos <= 0 Then
        private_NormalizeUpdateMacroRef = private_NormalizeUpdateMethodName(macroRef)
        Exit Function
    End If

    workbookPrefix = VBA.Left$(macroRef, bangPos)
    methodName = VBA.Mid$(macroRef, bangPos + 1)
    methodName = private_NormalizeUpdateMethodName(methodName)
    If VBA.Len(methodName) = 0 Then Exit Function

    private_NormalizeUpdateMacroRef = workbookPrefix & methodName
End Function


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
    Dim normalizedMethod As String
    Dim macroRef As String
    Dim scheduleAt As Date

    normalizedMethod = private_NormalizeUpdateMethodName(coreMethod)
    If VBA.Len(VBA.Trim$(normalizedMethod)) = 0 Then
        rt_Messaging.m_ShowStatusBarWarning "Не указан безопасный метод обновления.", 5
        Exit Sub
    End If

    macroRef = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!" & normalizedMethod
    updateKind = VBA.LCase$(VBA.Trim$(updateKind))
    If VBA.Len(updateKind) = 0 Then updateKind = "unknown"

    private_TryCancelScheduledTask g_ScheduledUpdateAt, g_ScheduledUpdateMacro
    scheduleAt = private_GetNextOnTimeTick()

    On Error GoTo EH_QUEUE
    Application.OnTime EarliestTime:=scheduleAt, Procedure:=macroRef
    g_ScheduledUpdateAt = scheduleAt
    g_ScheduledUpdateMacro = macroRef
    ex_Core.m_LogInfo "core-actions:redirect-safe-update queued kind='" & VBA.Replace$(updateKind, "'", "''") & "' macro='" & VBA.Replace$(macroRef, "'", "''") & "'"
    rt_Messaging.m_ShowStatusBarNotice "Начало обновления: безопасная задача поставлена в очередь.", 1
    Exit Sub

EH_QUEUE:
    ex_Core.m_LogError "core-actions:redirect-safe-update ontime-failed kind='" & VBA.Replace$(updateKind, "'", "''") & "' err='" & VBA.Replace$(Err.Description, "'", "''") & "'"
    rt_Messaging.m_ShowStatusBarError "Не удалось запланировать безопасное обновление: " & Err.Description, 6
End Sub
