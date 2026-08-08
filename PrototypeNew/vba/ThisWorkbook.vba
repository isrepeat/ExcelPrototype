Option Explicit
#Const LOGGING_DEBUG_ENABLED = True

Private Sub Workbook_Open()
    Dim openErrorNumber As Long
    Dim openErrorDescription As String
    Dim startupCleanupError As String
    Dim activeCallee As String

    On Error GoTo EH

    ex_Core.fn_Diagnostic_BeginLifecycleLogging
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "startup:method-enter method='ThisWorkbook.Workbook_Open'"
#End If

    ' При каждом открытии runtime и лист Main создаются заново. Сохранённое
    ' состояние страниц намеренно не восстанавливается.
    activeCallee = "rt_Lifecycle.fn_InitializeRuntime"
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "startup:call-enter caller='ThisWorkbook.Workbook_Open' " & _
        "callee='rt_Lifecycle.fn_InitializeRuntime'"
#End If
    If Not rt_Lifecycle.fn_InitializeRuntime( _
        "ThisWorkbook.Workbook_Open:main-create") Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError _
            "startup:call-failed caller='ThisWorkbook.Workbook_Open' " & _
            "callee='rt_Lifecycle.fn_InitializeRuntime' result='false'"
#End If
        If Not private_TryCleanupFailedStartup(startupCleanupError) Then
            VBA.MsgBox _
                "Инициализация PrototypeNew остановлена. Дополнительно не удалось " & _
                "очистить частично созданный runtime: " & startupCleanupError, _
                VBA.vbExclamation, "PrototypeNew / запуск"
        End If
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError _
            "startup:method-exit method='ThisWorkbook.Workbook_Open' " & _
            "result='false'"
#End If
        Exit Sub
    End If
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "startup:call-exit caller='ThisWorkbook.Workbook_Open' " & _
        "callee='rt_Lifecycle.fn_InitializeRuntime'"
    ex_Core.fn_Diagnostic_LogInfo _
        "startup:method-exit method='ThisWorkbook.Workbook_Open' result='true'"
#End If

    Exit Sub
EH:
    openErrorNumber = Err.Number
    openErrorDescription = Err.Description
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError _
        "startup:method-error method='ThisWorkbook.Workbook_Open' callee='" & _
        VBA.Replace$(activeCallee, "'", "''") & "' errNumber='" & _
        VBA.CStr(openErrorNumber) & "' err='" & _
        VBA.Replace$(openErrorDescription, "'", "''") & "'"
#End If
    If Not private_TryCleanupFailedStartup(startupCleanupError) Then
        openErrorDescription = openErrorDescription & VBA.vbCrLf & _
            "Ошибка cleanup частично созданного runtime: " & _
            startupCleanupError
    End If
    VBA.MsgBox "Не удалось инициализировать PrototypeNew: [" & _
        VBA.CStr(openErrorNumber) & "] " & openErrorDescription, _
        VBA.vbExclamation, "PrototypeNew / запуск"
End Sub


Private Function private_TryCleanupFailedStartup( _
    ByRef outErrorDescription As String _
) As Boolean
    outErrorDescription = VBA.vbNullString
    On Error GoTo EH_CLEANUP

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "startup:method-enter " & _
        "method='ThisWorkbook.private_TryCleanupFailedStartup'"
#End If

    rt_Lifecycle.fn_DisposeRuntime True, "workbook-open-failed"
    private_TryCleanupFailedStartup = True
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "startup:method-exit " & _
        "method='ThisWorkbook.private_TryCleanupFailedStartup' result='true'"
#End If
    Exit Function

EH_CLEANUP:
    outErrorDescription = "[" & VBA.CStr(Err.Number) & "] " & _
        Err.Description
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError _
        "startup:method-error " & _
        "method='ThisWorkbook.private_TryCleanupFailedStartup' err='" & _
        VBA.Replace$(outErrorDescription, "'", "''") & "'"
#End If
End Function

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim closeErrorNumber As Long
    Dim closeErrorDescription As String
    Dim discardedUnsavedChanges As Boolean
    Dim closeTeardownCommitted As Boolean
    Dim activeCallee As String

    On Error GoTo EH_BEFORE_CLOSE

    ex_Core.fn_Diagnostic_BeginLifecycleLogging
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "shutdown:method-enter method='ThisWorkbook.Workbook_BeforeClose'"
#End If

    ' DoEvents нужен длительным операциям (например WORD search) для кнопки
    ' отмены. Но он также позволяет пользователю закрыть книгу внутри активного
    ' метода controller. Уничтожать этот controller из его же call stack нельзя:
    ' Excel/VBE может аварийно завершить весь общий процесс Excel.
    activeCallee = "rt_Bridge.fn_IsDispatchingAny"
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "shutdown:call-enter caller='ThisWorkbook.Workbook_BeforeClose' " & _
        "callee='rt_Bridge.fn_IsDispatchingAny'"
#End If
    If rt_Bridge.fn_IsDispatchingAny() Then
        Cancel = True
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo _
            "shutdown:call-exit caller='ThisWorkbook.Workbook_BeforeClose' " & _
            "callee='rt_Bridge.fn_IsDispatchingAny' result='true'"
        ex_Core.fn_Diagnostic_LogError _
            "shutdown:close-cancelled reason='runtime-dispatch-active'"
        ex_Core.fn_Diagnostic_LogInfo _
            "shutdown:method-exit method='ThisWorkbook.Workbook_BeforeClose' " & _
            "result='false' reason='runtime-dispatch-active'"
#End If
        VBA.MsgBox _
            "Сейчас выполняется операция PrototypeNew. Дождитесь её завершения " & _
            "или отмените поиск кнопкой «Скасувати пошук», затем закройте книгу.", _
            VBA.vbExclamation, "PrototypeNew / закрытие книги"
        Exit Sub
    End If
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "shutdown:call-exit caller='ThisWorkbook.Workbook_BeforeClose' " & _
        "callee='rt_Bridge.fn_IsDispatchingAny' result='false'"
#End If

    ' BeforeClose вызывается до стандартного Excel prompt. Сначала фиксируем
    ' решение пользователя; только после этого начинается необратимый dispose.
    ' Иначе выбор Cancel в штатном prompt оставил бы открытую книгу без runtime.
    activeCallee = "ThisWorkbook.private_TryCommitCloseDecision"
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "shutdown:call-enter caller='ThisWorkbook.Workbook_BeforeClose' " & _
        "callee='ThisWorkbook.private_TryCommitCloseDecision'"
#End If
    If Not private_TryCommitCloseDecision( _
        Cancel, discardedUnsavedChanges) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo _
            "shutdown:call-exit caller='ThisWorkbook.Workbook_BeforeClose' " & _
            "callee='ThisWorkbook.private_TryCommitCloseDecision' result='false'"
        ex_Core.fn_Diagnostic_LogInfo _
            "shutdown:method-exit method='ThisWorkbook.Workbook_BeforeClose' " & _
            "result='false' reason='close-not-committed'"
#End If
        Exit Sub
    End If
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "shutdown:call-exit caller='ThisWorkbook.Workbook_BeforeClose' " & _
        "callee='ThisWorkbook.private_TryCommitCloseDecision' result='true'"
#End If

    ' После решения пользователя закрыть книгу выполняем единый dispose:
    ' отменяем callbacks и освобождаем module/class runtime state.
    closeTeardownCommitted = True
    activeCallee = "rt_Lifecycle.fn_DisposeRuntime"
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "shutdown:call-enter caller='ThisWorkbook.Workbook_BeforeClose' " & _
        "callee='rt_Lifecycle.fn_DisposeRuntime'"
#End If
    rt_Lifecycle.fn_DisposeRuntime True, "workbook-before-close"
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "shutdown:call-exit caller='ThisWorkbook.Workbook_BeforeClose' " & _
        "callee='rt_Lifecycle.fn_DisposeRuntime'"
    ex_Core.fn_Diagnostic_LogInfo _
        "shutdown:method-exit method='ThisWorkbook.Workbook_BeforeClose' " & _
        "result='true'"
#End If
    Exit Sub

EH_BEFORE_CLOSE:
    closeErrorNumber = Err.Number
    closeErrorDescription = Err.Description
    On Error Resume Next
    If Not closeTeardownCommitted And discardedUnsavedChanges Then _
        ThisWorkbook.Saved = False
    On Error GoTo 0

    ' До commit runtime цел, поэтому закрытие можно отменить. После commit книгу
    ' обязательно выгружаем, чтобы не оставить пользователю полуживую сессию.
    Cancel = Not closeTeardownCommitted
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError _
        "shutdown:method-error method='ThisWorkbook.Workbook_BeforeClose' " & _
        "callee='" & VBA.Replace$(activeCallee, "'", "''") & _
        "' errNumber='" & VBA.CStr(closeErrorNumber) & "' err='" & _
        VBA.Replace$(closeErrorDescription, "'", "''") & "'"
#End If
    If closeTeardownCommitted Then
        VBA.MsgBox "При завершении runtime PrototypeNew возникла ошибка: [" & _
            VBA.CStr(closeErrorNumber) & "] " & closeErrorDescription & VBA.vbCrLf & _
            "Книга будет закрыта, чтобы не оставлять повреждённую runtime-сессию.", _
            VBA.vbExclamation, "PrototypeNew / закрытие книги"
    Else
        VBA.MsgBox "Не удалось подготовить безопасное закрытие PrototypeNew: [" & _
            VBA.CStr(closeErrorNumber) & "] " & closeErrorDescription & VBA.vbCrLf & _
            "Закрытие книги отменено; runtime не уничтожался.", _
            VBA.vbExclamation, "PrototypeNew / закрытие книги"
    End If
End Sub


Private Function private_TryCommitCloseDecision( _
    ByRef Cancel As Boolean, _
    ByRef outDiscardedUnsavedChanges As Boolean _
) As Boolean
    Dim userChoice As VbMsgBoxResult
    Dim saveErrorNumber As Long
    Dim saveErrorDescription As String

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "shutdown:method-enter " & _
        "method='ThisWorkbook.private_TryCommitCloseDecision'"
#End If
    outDiscardedUnsavedChanges = False
    If ThisWorkbook.Saved Then
        private_TryCommitCloseDecision = True
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo _
            "shutdown:method-exit " & _
            "method='ThisWorkbook.private_TryCommitCloseDecision' " & _
            "result='true' decision='already-saved'"
#End If
        Exit Function
    End If

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "shutdown:checkpoint " & _
        "method='ThisWorkbook.private_TryCommitCloseDecision' " & _
        "step='show-save-prompt'"
#End If
    userChoice = VBA.MsgBox( _
        "Сохранить изменения в книге «" & ThisWorkbook.Name & "»?", _
        VBA.vbYesNoCancel Or VBA.vbQuestion, _
        "PrototypeNew / закрытие книги")

    Select Case userChoice
        Case VBA.vbCancel
            Cancel = True
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogInfo _
                "shutdown:method-exit " & _
                "method='ThisWorkbook.private_TryCommitCloseDecision' " & _
                "result='false' decision='cancel'"
#End If
            Exit Function

        Case VBA.vbYes
            On Error GoTo EH_SAVE
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogInfo _
                "shutdown:call-enter " & _
                "caller='ThisWorkbook.private_TryCommitCloseDecision' " & _
                "callee='ThisWorkbook.Save'"
#End If
            ThisWorkbook.Save
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogInfo _
                "shutdown:call-exit " & _
                "caller='ThisWorkbook.private_TryCommitCloseDecision' " & _
                "callee='ThisWorkbook.Save'"
#End If
            If Not ThisWorkbook.Saved Then
                Err.Raise VBA.vbObjectError + 9211, _
                    "ThisWorkbook.private_TryCommitCloseDecision", _
                    "Excel не подтвердил сохранение книги."
            End If

        Case VBA.vbNo
            ' Подавляем последующий стандартный prompt Excel. Если cleanup
            ' завершится ошибкой, BeforeClose вернёт Saved=False обратно.
            outDiscardedUnsavedChanges = True
            ThisWorkbook.Saved = True
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogInfo _
                "shutdown:checkpoint " & _
                "method='ThisWorkbook.private_TryCommitCloseDecision' " & _
                "decision='discard'"
#End If

        Case Else
            Cancel = True
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError _
                "shutdown:method-exit " & _
                "method='ThisWorkbook.private_TryCommitCloseDecision' " & _
                "result='false' decision='unsupported'"
#End If
            Exit Function
    End Select

    private_TryCommitCloseDecision = True
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "shutdown:method-exit " & _
        "method='ThisWorkbook.private_TryCommitCloseDecision' result='true'"
#End If
    Exit Function

EH_SAVE:
    saveErrorNumber = Err.Number
    saveErrorDescription = Err.Description
    Cancel = True
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError _
        "shutdown:method-error " & _
        "method='ThisWorkbook.private_TryCommitCloseDecision' " & _
        "callee='ThisWorkbook.Save' errNumber='" & _
        VBA.CStr(saveErrorNumber) & "' err='" & _
        VBA.Replace$(saveErrorDescription, "'", "''") & "'"
#End If
    On Error Resume Next
    VBA.MsgBox _
        "Не удалось сохранить книгу перед закрытием: [" & _
        VBA.CStr(saveErrorNumber) & "] " & saveErrorDescription, _
        VBA.vbExclamation, "PrototypeNew / закрытие книги"
    On Error GoTo 0
End Function

Private Sub Workbook_Activate()
    On Error GoTo EH_WORKBOOK_ACTIVATE
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogEventInfo _
        "event:method-enter method='ThisWorkbook.Workbook_Activate' " & _
        "enableEvents='" & _
        VBA.LCase$(VBA.CStr(Application.EnableEvents)) & "'"
#End If
    rt_Bridge.fn_OnSheetActivate ActiveSheet
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogEventInfo _
        "event:method-exit method='ThisWorkbook.Workbook_Activate' result='true'"
#End If
    Exit Sub

EH_WORKBOOK_ACTIVATE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogEventError _
        "event:method-error method='ThisWorkbook.Workbook_Activate' " & _
        "errNumber='" & VBA.CStr(Err.Number) & "' err='" & _
        VBA.Replace$(Err.Description, "'", "''") & "'"
#End If
End Sub

Private Sub Workbook_Deactivate()
    Dim syncOk As Boolean

    On Error GoTo EH_WORKBOOK_DEACTIVATE
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogEventInfo _
        "event:method-enter method='ThisWorkbook.Workbook_Deactivate'"
#End If
    syncOk = rt_HotkeyRuntime.fn_ActivatePageHotkeys(VBA.vbNullString)
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogEventInfo _
        "event:method-exit method='ThisWorkbook.Workbook_Deactivate' result='" & _
        VBA.LCase$(VBA.CStr(syncOk)) & "'"
#End If
    Exit Sub

EH_WORKBOOK_DEACTIVATE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogEventError _
        "event:method-error method='ThisWorkbook.Workbook_Deactivate' " & _
        "errNumber='" & VBA.CStr(Err.Number) & "' err='" & _
        VBA.Replace$(Err.Description, "'", "''") & "'"
#End If
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo EH_SHEET_CHANGE
    rt_Bridge.fn_OnSheetChange Sh, Target
    Exit Sub

EH_SHEET_CHANGE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "PrototypeNew: Workbook_SheetChange failed: " & Err.Description
#End If
End Sub

Private Sub Workbook_SheetSelectionChange( _
    ByVal Sh As Object, _
    ByVal Target As Range _
)
    On Error GoTo EH_SHEET_SELECTION_CHANGE
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogEventInfo _
        "event:method-enter " & _
        "method='ThisWorkbook.Workbook_SheetSelectionChange'"
#End If
    rt_Bridge.fn_OnSheetSelectionChange Sh, Target
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogEventInfo _
        "event:method-exit " & _
        "method='ThisWorkbook.Workbook_SheetSelectionChange' result='true'"
#End If
    Exit Sub

EH_SHEET_SELECTION_CHANGE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogEventError _
        "event:method-error " & _
        "method='ThisWorkbook.Workbook_SheetSelectionChange' errNumber='" & _
        VBA.CStr(Err.Number) & "' err='" & _
        VBA.Replace$(Err.Description, "'", "''") & "'"
#End If
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error GoTo EH_SHEET_ACTIVATE
    rt_Bridge.fn_OnSheetActivate Sh
    Exit Sub

EH_SHEET_ACTIVATE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "PrototypeNew: Workbook_SheetActivate failed: " & Err.Description
#End If
End Sub

Private Sub Workbook_SheetBeforeDelete(ByVal Sh As Object)
    Dim ws As Worksheet
    Dim sheetName As String

    If Not TypeOf Sh Is Worksheet Then Exit Sub
    If rt_PageManager.fn_IsPageRemovalInProgress() Then Exit Sub
    Set ws = Sh

    On Error Resume Next
    sheetName = VBA.LCase$(VBA.Trim$(VBA.CStr(ws.Name)))
    Err.Clear
    On Error GoTo 0

    ' Временные листы участвуют только в сценариях reset/restore.
    ' Не пробрасываем их в PageManager, чтобы не трогать runtime-реестр страниц.
    If VBA.StrComp(sheetName, "__startup_tmp__", VBA.vbTextCompare) = 0 Then Exit Sub
    If VBA.StrComp(sheetName, "__restore_tmp__", VBA.vbTextCompare) = 0 Then Exit Sub

    Call ex_HelpersSheet.fn_RemovePageByWorksheet(ws)
End Sub

' //
' // API
' //
' Callstack[1]: ThisWorkbook.Workbook_Open -> ThisWorkbook.m_ResetWorkbookAndCreateMainPage
Public Function m_ResetWorkbookAndCreateMainPage( _
    Optional ByVal renderReason As String = "ThisWorkbook.m_ResetWorkbookAndCreateMainPage", _
    Optional ByVal showErrorUi As Boolean = True _
) As Boolean
    Dim resetErrorNumber As Long
    Dim resetErrorDescription As String

    On Error GoTo EH_RESET_MAIN
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "startup:method-enter " & _
        "method='ThisWorkbook.m_ResetWorkbookAndCreateMainPage'"
#End If
    m_ResetWorkbookAndCreateMainPage = _
        private_ResetWorkbookAndCreateMainPage(renderReason, showErrorUi)
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "startup:method-exit " & _
        "method='ThisWorkbook.m_ResetWorkbookAndCreateMainPage' result='" & _
        VBA.LCase$(VBA.CStr(m_ResetWorkbookAndCreateMainPage)) & "'"
#End If
    Exit Function

EH_RESET_MAIN:
    resetErrorNumber = Err.Number
    resetErrorDescription = Err.Description
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError _
        "startup:method-error " & _
        "method='ThisWorkbook.m_ResetWorkbookAndCreateMainPage' " & _
        "errNumber='" & VBA.CStr(resetErrorNumber) & "' err='" & _
        VBA.Replace$(resetErrorDescription, "'", "''") & "'"
#End If
    Err.Raise resetErrorNumber, _
        "ThisWorkbook.m_ResetWorkbookAndCreateMainPage", _
        resetErrorDescription
End Function


Private Function private_ResetWorkbookAndCreateMainPage( _
    Optional ByVal renderReason As String = "ThisWorkbook.private_ResetWorkbookAndCreateMainPage", _
    Optional ByVal showErrorUi As Boolean = True _
) As Boolean
    Dim wb As Workbook
    Dim previousMainWs As Worksheet
    Dim createdMainWs As Worksheet
    Dim cleanupWs As Worksheet
    Dim previousMainBackupName As String
    Dim createdPage As obj_IPage
    Dim createdPageBase As obj_PageBase
    Dim isPageCreated As Boolean
    Dim isMainRendered As Boolean
    Dim previousDisplayAlerts As Boolean
    Dim previousEnableEvents As Boolean
    Dim applicationStateCaptured As Boolean
    Dim createErrorDescription As String
    Dim cleanupErrorDescription As String
    Dim worksheetIndex As Long
    Dim activeStep As String

    Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    On Error GoTo EH_CREATE
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "startup:method-enter " & _
        "method='ThisWorkbook.private_ResetWorkbookAndCreateMainPage'"
#End If

    activeStep = "capture-application-state"
    previousDisplayAlerts = Application.DisplayAlerts
    previousEnableEvents = Application.EnableEvents
    applicationStateCaptured = True
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    ' Старый runtime отделяем от сохранённых листов, но сами листы пока не
    ' удаляем. Они являются rollback-копией до успешного render нового Main.
    activeStep = "rt_PageManager.fn_DisposeAllPages"
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "startup:call-enter " & _
        "caller='ThisWorkbook.private_ResetWorkbookAndCreateMainPage' " & _
        "callee='rt_PageManager.fn_DisposeAllPages'"
#End If
    rt_PageManager.fn_DisposeAllPages
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "startup:call-exit " & _
        "caller='ThisWorkbook.private_ResetWorkbookAndCreateMainPage' " & _
        "callee='rt_PageManager.fn_DisposeAllPages'"
#End If

    ' Чтобы новый page сразу получил окончательное имя Main (включая все
    ' runtime registry keys), старый Main лишь временно переименовываем.
    On Error Resume Next
    Set previousMainWs = wb.Worksheets("Main")
    Err.Clear
    On Error GoTo EH_CREATE
    If Not previousMainWs Is Nothing Then
        previousMainBackupName = private_BuildUniqueWorksheetName( _
            wb, "__startup_old_main__")
        If VBA.Len(previousMainBackupName) = 0 Then
            Err.Raise VBA.vbObjectError + 9201, _
                "ThisWorkbook.private_ResetWorkbookAndCreateMainPage", _
                "Не удалось подобрать rollback-имя для существующего листа Main."
        End If
        previousMainWs.Name = previousMainBackupName
    End If

    activeStep = "New obj_PageMain"
    Set createdPage = New obj_PageMain
    If createdPage Is Nothing Then
        Err.Raise VBA.vbObjectError + 9202, _
            "ThisWorkbook.private_ResetWorkbookAndCreateMainPage", _
            "Не удалось создать объект страницы Main."
    End If

    activeStep = "rt_PageManager.fn_CreatePage"
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "startup:call-enter " & _
        "caller='ThisWorkbook.private_ResetWorkbookAndCreateMainPage' " & _
        "callee='rt_PageManager.fn_CreatePage'"
#End If
    If Not rt_PageManager.fn_CreatePage( _
        createdPage, "ui\MainUI.xml", "Main") Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError _
            "startup:call-failed " & _
            "caller='ThisWorkbook.private_ResetWorkbookAndCreateMainPage' " & _
            "callee='rt_PageManager.fn_CreatePage' result='false'"
#End If
        GoTo EH_CREATE
    End If
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "startup:call-exit " & _
        "caller='ThisWorkbook.private_ResetWorkbookAndCreateMainPage' " & _
        "callee='rt_PageManager.fn_CreatePage'"
#End If
    isPageCreated = True
    Set createdPageBase = createdPage.GetPageBase()
    If createdPageBase Is Nothing Then GoTo EH_CREATE
    Set createdMainWs = createdPageBase.Worksheet
    If createdMainWs Is Nothing Then GoTo EH_CREATE

    activeStep = "rt_PageManager.fn_RenderPage"
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "startup:call-enter " & _
        "caller='ThisWorkbook.private_ResetWorkbookAndCreateMainPage' " & _
        "callee='rt_PageManager.fn_RenderPage'"
#End If
    If Not rt_PageManager.fn_RenderPage(createdPage, renderReason) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError _
            "startup:call-failed " & _
            "caller='ThisWorkbook.private_ResetWorkbookAndCreateMainPage' " & _
            "callee='rt_PageManager.fn_RenderPage' result='false'"
#End If
        GoTo EH_CREATE
    End If
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "startup:call-exit " & _
        "caller='ThisWorkbook.private_ResetWorkbookAndCreateMainPage' " & _
        "callee='rt_PageManager.fn_RenderPage'"
#End If
    isMainRendered = True
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "startup:new-main-rendered"
#End If

    ' Только успешный render является commit-point. До него ни один исходный
    ' лист не удалялся, поэтому binding/config ошибка не разрушает workbook.
    activeStep = "delete-old-worksheets"
    For worksheetIndex = wb.Worksheets.Count To 1 Step -1
        Set cleanupWs = wb.Worksheets(worksheetIndex)
        If Not cleanupWs Is createdMainWs Then cleanupWs.Delete
    Next worksheetIndex
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "startup:checkpoint " & _
        "method='ThisWorkbook.private_ResetWorkbookAndCreateMainPage' " & _
        "step='old-worksheets-removed'"
#End If

    Application.DisplayAlerts = previousDisplayAlerts
    Application.EnableEvents = previousEnableEvents
    private_ResetWorkbookAndCreateMainPage = True
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "startup:method-exit " & _
        "method='ThisWorkbook.private_ResetWorkbookAndCreateMainPage' " & _
        "result='true'"
#End If
    Exit Function

EH_CREATE:
    createErrorDescription = Err.Description
    If VBA.Len(VBA.Trim$(createErrorDescription)) = 0 Then
        createErrorDescription = "Операция создания или рендера Main вернула False без VBA-ошибки."
    End If
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError _
        "startup:method-error " & _
        "method='ThisWorkbook.private_ResetWorkbookAndCreateMainPage' " & _
        "step='" & VBA.Replace$(activeStep, "'", "''") & "' err='" & _
        VBA.Replace$(createErrorDescription, "'", "''") & "'"
#End If
    On Error Resume Next
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' До commit-point новый Main можно безопасно удалить: сохранённые листы
    ' всё ещё существуют, поэтому Excel никогда не остаётся без worksheet.
    If Not isMainRendered And _
        Not createdPage Is Nothing And isPageCreated Then
        Call rt_PageManager.fn_RemovePage(createdPage, True)
    End If

    ' Возвращаем исходному Main его имя, если transaction не дошла до render.
    If Not isMainRendered And Not previousMainWs Is Nothing Then
        If Not private_WorksheetNameExists(wb, "Main") Then
            previousMainWs.Name = "Main"
        End If
    End If

    If applicationStateCaptured Then
        Application.DisplayAlerts = previousDisplayAlerts
        Application.EnableEvents = previousEnableEvents
    End If
    On Error GoTo 0

    If isMainRendered Then
        ' Новый Main уже полностью работоспособен. Ошибка могла возникнуть лишь
        ' при удалении одного из старых листов; не разбираем успешный runtime.
        cleanupErrorDescription = createErrorDescription
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError _
            "PrototypeNew: Main rendered, but old worksheets cleanup failed: " & _
            cleanupErrorDescription
#End If
        If showErrorUi Then
            VBA.MsgBox "Страница Main создана, но не удалось удалить один из старых листов: " & _
                cleanupErrorDescription, VBA.vbExclamation, _
                "PrototypeNew / запуск"
        End If
        private_ResetWorkbookAndCreateMainPage = True
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo _
            "startup:method-exit " & _
            "method='ThisWorkbook.private_ResetWorkbookAndCreateMainPage' " & _
            "result='true' warning='old-worksheet-cleanup-failed'"
#End If
        Exit Function
    End If

    If showErrorUi Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to create default main page: " & createErrorDescription
#End If
        VBA.MsgBox "Не удалось создать страницу Main: " & _
            createErrorDescription, VBA.vbExclamation, _
            "PrototypeNew / запуск"
    End If
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError _
        "startup:method-exit " & _
        "method='ThisWorkbook.private_ResetWorkbookAndCreateMainPage' " & _
        "result='false'"
#End If
End Function


Private Function private_BuildUniqueWorksheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim i As Long
    Dim suffix As String
    Dim candidate As String

    If wb Is Nothing Then Exit Function

    baseName = VBA.Trim$(baseName)
    If VBA.Len(baseName) = 0 Then baseName = "tmp_sheet"
    If VBA.Len(baseName) > 31 Then baseName = VBA.Left$(baseName, 31)

    If Not private_WorksheetNameExists(wb, baseName) Then
        private_BuildUniqueWorksheetName = baseName
        Exit Function
    End If

    For i = 1 To 9999
        suffix = "_" & VBA.CStr(i)
        candidate = VBA.Left$(baseName, 31 - VBA.Len(suffix)) & suffix
        If VBA.Len(candidate) = 0 Then candidate = "tmp" & suffix
        If Not private_WorksheetNameExists(wb, candidate) Then
            private_BuildUniqueWorksheetName = candidate
            Exit Function
        End If
    Next i
End Function


Private Function private_WorksheetNameExists(ByVal wb As Workbook, ByVal worksheetName As String) As Boolean
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function

    worksheetName = VBA.Trim$(worksheetName)
    If VBA.Len(worksheetName) = 0 Then Exit Function

    On Error Resume Next
    Set ws = wb.Worksheets(worksheetName)
    private_WorksheetNameExists = Not ws Is Nothing
    Err.Clear
    On Error GoTo 0
End Function
