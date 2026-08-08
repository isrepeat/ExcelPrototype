Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const RUNTIME_SNAPSHOTS_ENABLED = False

Private Sub Workbook_Open()
    Dim restoredPagesCount As Long
    Dim restoredOk As Boolean
    Dim openErrorNumber As Long
    Dim openErrorDescription As String
    Dim startupCleanupError As String

    On Error GoTo EH

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "startup:workbook-open-enter"
#End If

#If RUNTIME_SNAPSHOTS_ENABLED Then
    restoredOk = rt_RestoreManager.fn_RestoreRuntimeState( _
        "Workbook_Open", restoredPagesCount)
    If restoredOk And restoredPagesCount > 0 Then Exit Sub
#End If

    If Not rt_Lifecycle.fn_InitializeRuntime( _
        "ThisWorkbook.Workbook_Open:main-create") Then
        If Not private_TryCleanupFailedStartup(startupCleanupError) Then
            VBA.MsgBox _
                "Инициализация PrototypeNew остановлена. Дополнительно не удалось " & _
                "очистить частично созданный runtime: " & startupCleanupError, _
                VBA.vbExclamation, "PrototypeNew / запуск"
        End If
        Exit Sub
    End If
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "startup:workbook-open-done"
#End If

    Exit Sub
EH:
    openErrorNumber = Err.Number
    openErrorDescription = Err.Description
    If Not private_TryCleanupFailedStartup(startupCleanupError) Then
        openErrorDescription = openErrorDescription & VBA.vbCrLf & _
            "Ошибка cleanup частично созданного runtime: " & _
            startupCleanupError
    End If
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "PrototypeNew: Workbook_Open failed: [" & _
        VBA.CStr(openErrorNumber) & "] " & openErrorDescription
#End If
    VBA.MsgBox "Не удалось инициализировать PrototypeNew: [" & _
        VBA.CStr(openErrorNumber) & "] " & openErrorDescription, _
        VBA.vbExclamation, "PrototypeNew / запуск"
End Sub


Private Function private_TryCleanupFailedStartup( _
    ByRef outErrorDescription As String _
) As Boolean
    outErrorDescription = VBA.vbNullString
    On Error GoTo EH_CLEANUP

    rt_Lifecycle.fn_DisposeRuntime True, "workbook-open-failed"
    private_TryCleanupFailedStartup = True
    Exit Function

EH_CLEANUP:
    outErrorDescription = "[" & VBA.CStr(Err.Number) & "] " & _
        Err.Description
End Function

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim previousEnableEvents As Boolean
    Dim enableEventsCaptured As Boolean
    Dim closeErrorNumber As Long
    Dim closeErrorDescription As String
    Dim discardedUnsavedChanges As Boolean
    Dim closeTeardownCommitted As Boolean

    On Error GoTo EH_BEFORE_CLOSE

#If LOGGING_DEBUG_ENABLED Then
    ' Первый checkpoint должен появиться до любого обращения к OnKey/OnTime/COM.
    ' Если его нет, Workbook_BeforeClose вообще не был вызван.
    ex_Core.fn_Diagnostic_LogInfo "shutdown:before-close-enter"
#End If

    ' DoEvents нужен длительным операциям (например WORD search) для кнопки
    ' отмены. Но он также позволяет пользователю закрыть книгу внутри активного
    ' метода controller. Уничтожать этот controller из его же call stack нельзя:
    ' Excel/VBE может аварийно завершить весь общий процесс Excel.
    If rt_Bridge.fn_IsDispatchingAny() Then
        Cancel = True
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError _
            "shutdown:close-cancelled reason='runtime-dispatch-active'"
#End If
        VBA.MsgBox _
            "Сейчас выполняется операция PrototypeNew. Дождитесь её завершения " & _
            "или отмените поиск кнопкой «Скасувати пошук», затем закройте книгу.", _
            VBA.vbExclamation, "PrototypeNew / закрытие книги"
        Exit Sub
    End If

    ' BeforeClose вызывается до стандартного Excel prompt. Сначала фиксируем
    ' решение пользователя; только после этого начинается необратимый dispose.
    ' Иначе выбор Cancel в штатном prompt оставил бы открытую книгу без runtime.
    If Not private_TryCommitCloseDecision( _
        Cancel, discardedUnsavedChanges) Then Exit Sub

    previousEnableEvents = Application.EnableEvents
    enableEventsCaptured = True
    Application.EnableEvents = False
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "shutdown:events-disabled"
#End If

    ' Фаза prepare только отменяет внешние callbacks. До её завершения runtime
    ' не разрушается, поэтому ошибка ещё может безопасно отменить закрытие.
    rt_Lifecycle.fn_PrepareRuntimeDispose True, "workbook-before-close"
    closeTeardownCommitted = True

    ' После prepare закрытие необратимо: оставлять книгу открытой с частично
    ' освобождённым graph опаснее, чем завершить unload с явной диагностикой.
    rt_Lifecycle.fn_DisposePreparedRuntime True, "workbook-before-close"

    Application.EnableEvents = previousEnableEvents
#If LOGGING_DEBUG_ENABLED Then
    ' Этот checkpoint отделяет VBA-cleanup от последующей native save/unload
    ' фазы Excel. Если падение случится позже, cleanup уже был завершён.
    ex_Core.fn_Diagnostic_LogInfo "shutdown:before-close-done"
#End If
    Exit Sub

EH_BEFORE_CLOSE:
    closeErrorNumber = Err.Number
    closeErrorDescription = Err.Description
    On Error Resume Next
    If enableEventsCaptured Then Application.EnableEvents = previousEnableEvents
    If Not closeTeardownCommitted And discardedUnsavedChanges Then _
        ThisWorkbook.Saved = False
    On Error GoTo 0

    ' До commit runtime цел, поэтому закрытие можно отменить. После commit книгу
    ' обязательно выгружаем, чтобы не оставить пользователю полуживую сессию.
    Cancel = Not closeTeardownCommitted
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "PrototypeNew: Workbook_BeforeClose cleanup failed: [" & _
        VBA.CStr(closeErrorNumber) & "] " & closeErrorDescription
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

    outDiscardedUnsavedChanges = False
    If ThisWorkbook.Saved Then
        private_TryCommitCloseDecision = True
        Exit Function
    End If

    userChoice = VBA.MsgBox( _
        "Сохранить изменения в книге «" & ThisWorkbook.Name & "»?", _
        VBA.vbYesNoCancel Or VBA.vbQuestion, _
        "PrototypeNew / закрытие книги")

    Select Case userChoice
        Case VBA.vbCancel
            Cancel = True
            Exit Function

        Case VBA.vbYes
            On Error GoTo EH_SAVE
            ThisWorkbook.Save
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

        Case Else
            Cancel = True
            Exit Function
    End Select

    private_TryCommitCloseDecision = True
    Exit Function

EH_SAVE:
    saveErrorNumber = Err.Number
    saveErrorDescription = Err.Description
    Cancel = True
    On Error Resume Next
    VBA.MsgBox _
        "Не удалось сохранить книгу перед закрытием: [" & _
        VBA.CStr(saveErrorNumber) & "] " & saveErrorDescription, _
        VBA.vbExclamation, "PrototypeNew / закрытие книги"
    On Error GoTo 0
End Function

Private Sub Workbook_Activate()
    On Error GoTo EH_WORKBOOK_ACTIVATE
    rt_Bridge.fn_OnSheetActivate ActiveSheet
    Exit Sub

EH_WORKBOOK_ACTIVATE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "PrototypeNew: Workbook_Activate failed: " & Err.Description
#End If
End Sub

Private Sub Workbook_Deactivate()
    Dim syncOk As Boolean

    On Error GoTo EH_WORKBOOK_DEACTIVATE
    syncOk = rt_HotkeyRuntime.fn_ActivatePageHotkeys(VBA.vbNullString)
    Exit Sub

EH_WORKBOOK_DEACTIVATE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "PrototypeNew: Workbook_Deactivate failed: " & Err.Description
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
    rt_Bridge.fn_OnSheetSelectionChange Sh, Target
    Exit Sub

EH_SHEET_SELECTION_CHANGE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError _
        "PrototypeNew: Workbook_SheetSelectionChange failed: " & _
        Err.Description
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
    m_ResetWorkbookAndCreateMainPage = private_ResetWorkbookAndCreateMainPage(renderReason, showErrorUi)
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

    Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    On Error GoTo EH_CREATE
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "startup:main-reset-enter"
#End If

    previousDisplayAlerts = Application.DisplayAlerts
    previousEnableEvents = Application.EnableEvents
    applicationStateCaptured = True
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    ' Старый runtime отделяем от сохранённых листов, но сами листы пока не
    ' удаляем. Они являются rollback-копией до успешного render нового Main.
    rt_PageManager.fn_DisposeAllPages
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "startup:old-runtime-disposed"
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

    Set createdPage = New obj_PageMain
    If createdPage Is Nothing Then
        Err.Raise VBA.vbObjectError + 9202, _
            "ThisWorkbook.private_ResetWorkbookAndCreateMainPage", _
            "Не удалось создать объект страницы Main."
    End If

    If Not rt_PageManager.fn_CreatePage( _
        createdPage, "ui\MainUI.xml", "Main") Then GoTo EH_CREATE
    isPageCreated = True
    Set createdPageBase = createdPage.GetPageBase()
    If createdPageBase Is Nothing Then GoTo EH_CREATE
    Set createdMainWs = createdPageBase.Worksheet
    If createdMainWs Is Nothing Then GoTo EH_CREATE

    If Not rt_PageManager.fn_RenderPage(createdPage, renderReason) Then GoTo EH_CREATE
    isMainRendered = True
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "startup:new-main-rendered"
#End If

    ' Только успешный render является commit-point. До него ни один исходный
    ' лист не удалялся, поэтому binding/config ошибка не разрушает workbook.
    For worksheetIndex = wb.Worksheets.Count To 1 Step -1
        Set cleanupWs = wb.Worksheets(worksheetIndex)
        If Not cleanupWs Is createdMainWs Then cleanupWs.Delete
    Next worksheetIndex
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "startup:old-worksheets-removed"
#End If

    Application.DisplayAlerts = previousDisplayAlerts
    Application.EnableEvents = previousEnableEvents
    private_ResetWorkbookAndCreateMainPage = True
    Exit Function

EH_CREATE:
    createErrorDescription = Err.Description
    If VBA.Len(VBA.Trim$(createErrorDescription)) = 0 Then
        createErrorDescription = "Операция создания или рендера Main вернула False без VBA-ошибки."
    End If
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
