Attribute VB_Name = "rt_RestoreManager"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const RUNTIME_GLOBALS_NS As String = "urn:excelprototype:runtime-globals:v1"
Private Const RUNTIME_GLOBALS_ROOT As String = "runtimeGlobals"
Private Const RUNTIME_GLOBALS_MODULE_NODE As String = "module"
Private Const RUNTIME_GLOBALS_MODULE_NAME_ATTR As String = "name"
Private Const RUNTIME_GLOBALS_MODULE_SNAPSHOT_NODE As String = "snapshot"
Private Const RUNTIME_GLOBALS_ACTIVE_SHEET_ATTR As String = "activeSheetName"
Private Const MODULE_NAME_PAGE_MANAGER As String = "rt_PageManager"

' Callstack[1]: ex_Core.private_Dev_TryPrepareRuntimeForHotUpdate -> private_Dev_TryRunModuleDisposers -> Application.Run(rt_RestoreManager.fn_Module_Dispose)
Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:rt_RestoreManager.fn_Module_Dispose"
#End If
End Sub

' //
' // API
' //
' Callstack[1]: ThisWorkbook.Workbook_BeforeClose -> rt_RestoreManager.fn_SaveRuntimeState
' Callstack[2]: rt_CoreActions.private_ScheduleUpdateAndRerender -> rt_RestoreManager.fn_SaveRuntimeState
' Callstack[3]: ex_Core.private_Dev_TryRunSafeUpdateByMode -> private_Dev_TryRunRuntimeBooleanFunction("rt_RestoreManager","fn_SaveRuntimeState")
' Callstack[4]: rt_RestoreManager.private_TryFallbackRestoreByResettingMainPage -> rt_RestoreManager.fn_SaveRuntimeState
Public Function fn_SaveRuntimeState() As Boolean
    ' Единая точка сохранения runtime-состояния.
    ' Сейчас сохраняется модульный snapshot rt_PageManager, но формат рассчитан
    ' на добавление других runtime-модулей без изменения внешнего API.
    fn_SaveRuntimeState = private_TrySaveRuntimeGlobalsSnapshot()
End Function

' Callstack[1]: ThisWorkbook.Workbook_Open -> rt_RestoreManager.fn_RestoreRuntimeState
' Callstack[2]: rt_CoreActions.fn_RerenderLastPageAfterUpdate -> rt_RestoreManager.fn_RestoreRuntimeState
' Callstack[3]: rt_RestoreManager.fn_RunDeferredRuntimeStateRestore -> rt_RestoreManager.fn_RestoreRuntimeState
Public Function fn_RestoreRuntimeState( _
    Optional ByVal reasonText As String = VBA.vbNullString, _
    Optional ByRef outRestoredPagesCount As Long = 0 _
) As Boolean
    Static isRuntimeStateRestoreRunning As Boolean
    Dim restoredActiveWorksheetName As String

    reasonText = VBA.Trim$(reasonText)
    If VBA.Len(reasonText) = 0 Then reasonText = "unknown"

    outRestoredPagesCount = 0
    If isRuntimeStateRestoreRunning Then Exit Function

    isRuntimeStateRestoreRunning = True
    On Error GoTo EH_RESTORE

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "restore-manager:restore-runtime-state start reason='" & VBA.Replace$(reasonText, "'", "''") & "'"
#End If

    ' Быстрый путь: если активный лист уже привязан к runtime-странице,
    ' не делаем жесткий fallback/reset, а просто пытаемся подтянуть runtime state.
    If private_HasRuntimePageForActiveWorksheet() Then
        If Not private_TryRestoreRuntimeGlobalsSnapshot(restoredActiveWorksheetName) Then GoTo RestoreFailed
        If Not private_TryGetPagesCount(outRestoredPagesCount) Then GoTo RestoreFailed
        Call private_TryActivateSavedWorksheetAfterRestore(reasonText, restoredActiveWorksheetName)
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "restore-manager:restore-runtime-state done reason='" & VBA.Replace$(reasonText, "'", "''") & "' restoredPages=" & VBA.CStr(outRestoredPagesCount)
#End If
        fn_RestoreRuntimeState = True
        GoTo Cleanup
    End If

    If Not private_TryRestoreRuntimeGlobalsSnapshot(restoredActiveWorksheetName) Then GoTo RestoreFailed
    If Not private_TryGetPagesCount(outRestoredPagesCount) Then GoTo RestoreFailed
    If outRestoredPagesCount <= 0 Then GoTo RestoreFailed
    Call private_TryActivateSavedWorksheetAfterRestore(reasonText, restoredActiveWorksheetName)

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "restore-manager:restore-runtime-state done reason='" & VBA.Replace$(reasonText, "'", "''") & "' restoredPages=" & VBA.CStr(outRestoredPagesCount)
#End If
    fn_RestoreRuntimeState = True

Cleanup:
    isRuntimeStateRestoreRunning = False
    Exit Function

RestoreFailed:
    If private_TryFallbackRestoreByResettingMainPage(reasonText, outRestoredPagesCount) Then
        fn_RestoreRuntimeState = True
        GoTo Cleanup
    End If
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "restore-manager:restore-runtime-state failed reason='" & VBA.Replace$(reasonText, "'", "''") & "' restoredPages=" & VBA.CStr(outRestoredPagesCount)
#End If
    GoTo Cleanup

EH_RESTORE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "restore-manager:restore-runtime-state exception reason='" & VBA.Replace$(reasonText, "'", "''") & "' err='" & VBA.Replace$(Err.Description, "'", "''") & "'"
#End If
    Resume Cleanup
End Function

' Callstack[1]: ex_Core.private_Dev_QueueRuntimeStateRestoreAfterUpdate -> Application.OnTime -> rt_RestoreManager.fn_RunDeferredRuntimeStateRestore
Public Sub fn_RunDeferredRuntimeStateRestore()
    Dim restoredPagesCount As Long

    Call fn_RestoreRuntimeState("deferred:on-time", restoredPagesCount)
End Sub

' Callstack[1]: rt_RestoreManager.private_TryDeserializeRuntimeModuleSnapshot(rt_PageManager) -> rt_PageManager.fn_TryDeserializeModuleSnapshot -> rt_RestoreManager.fn_TryPrepareWorkbookForRestore
Public Function fn_TryPrepareWorkbookForRestore(ByRef outTemporaryWorksheet As Worksheet) As Boolean
    Dim wb As Workbook
    Dim tmpName As String

    Set outTemporaryWorksheet = Nothing
    Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    ' Перед восстановлением очищаем runtime-реестр страниц, чтобы избежать
    ' смешивания старых инстансов с восстанавливаемыми.
    rt_PageManager.fn_DisposeAllPages

    On Error GoTo EH_RESET
    Application.DisplayAlerts = False
    ' Приводим workbook к "чистому" состоянию с одним листом-заглушкой.
    ' На него будет опираться создание новых page-листов при deserialize.
    Do While wb.Worksheets.Count > 1
        wb.Worksheets(1).Delete
    Loop
    Set outTemporaryWorksheet = wb.Worksheets(1)
    Application.DisplayAlerts = True

    tmpName = "__restore_tmp__"
    On Error Resume Next
    outTemporaryWorksheet.Name = tmpName
    Err.Clear
    On Error GoTo 0

    fn_TryPrepareWorkbookForRestore = True
    Exit Function

EH_RESET:
    Application.DisplayAlerts = True
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "RestoreManager: failed to reset workbook before restore: " & Err.Description
#End If
End Function

' Callstack[1]: rt_RestoreManager.private_TryDeserializeRuntimeModuleSnapshot(rt_PageManager) -> rt_PageManager.fn_TryDeserializeModuleSnapshot -> rt_RestoreManager.fn_TryFinalizeWorkbookAfterRestore
Public Function fn_TryFinalizeWorkbookAfterRestore(ByVal temporaryWorksheet As Worksheet) As Boolean
    Dim wb As Workbook

    fn_TryFinalizeWorkbookAfterRestore = True
    If temporaryWorksheet Is Nothing Then Exit Function

    Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function
    If wb.Worksheets.Count <= 1 Then Exit Function

    ' Если при восстановлении были созданы страницы, временный лист уже лишний.
    ' Удаляем его в самом конце, когда все связи и рендер уже завершены.
    On Error Resume Next
    Application.DisplayAlerts = False
    temporaryWorksheet.Delete
    Application.DisplayAlerts = True
    If Err.Number <> 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "RestoreManager: failed to finalize restore temporary worksheet: " & Err.Description
#End If
        Err.Clear
        fn_TryFinalizeWorkbookAfterRestore = False
    End If
    On Error GoTo 0
End Function

' Callstack[1]: rt_RestoreManager.private_TryDeserializeRuntimeModuleSnapshot(rt_PageManager) -> rt_PageManager.fn_TryDeserializeModuleSnapshot -> rt_RestoreManager.fn_TryRestoreSerializableCollectionState
Public Function fn_TryRestoreSerializableCollectionState( _
    ByVal serializableItems As Collection, _
    Optional ByVal ownerName As String = VBA.vbNullString _
) As Boolean
    Dim item As Variant
    Dim serializableItem As obj_ISerializable
    Dim typeName As String

    fn_TryRestoreSerializableCollectionState = True
    ownerName = VBA.Trim$(ownerName)
    If VBA.Len(ownerName) = 0 Then ownerName = "unknown"

    If serializableItems Is Nothing Then Exit Function

    ' Вторая фаза восстановления:
    ' 1) объекты уже созданы и десериализованы;
    ' 2) здесь каждый объект достраивает внутренние ссылки/состояние в TryRestoreState.
    For Each item In serializableItems
        Set serializableItem = Nothing
        typeName = "unknown"

        On Error Resume Next
        typeName = VBA.TypeName(item)
        Set serializableItem = item
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "RestoreManager: collection item does not implement obj_ISerializable. owner='" & VBA.Replace$(ownerName, "'", "''") & "' type='" & VBA.Replace$(typeName, "'", "''") & "'."
#End If
            fn_TryRestoreSerializableCollectionState = False
            Exit Function
        End If
        On Error GoTo 0

        If serializableItem Is Nothing Then GoTo ContinueItem
        If serializableItem.TryRestoreState() Then GoTo ContinueItem

#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "RestoreManager: TryRestoreState failed. owner='" & VBA.Replace$(ownerName, "'", "''") & "' type='" & VBA.Replace$(typeName, "'", "''") & "'."
#End If
        fn_TryRestoreSerializableCollectionState = False
        Exit Function

ContinueItem:
    Next item
End Function

' //
' // Internal
' //
Private Function private_TrySaveRuntimeGlobalsSnapshot() As Boolean
    Dim dom As Object
    Dim rootNode As Object
    Dim partObj As Object
    Dim activeWorksheetName As String

    ' Корневой snapshot runtime содержит набор "module" узлов.
    ' Каждый модуль сам сериализует свой внутренний state в XML-строку.
    If Not ex_Core.fn_CustomXmlPartStore_TryCreateEmptyDom(RUNTIME_GLOBALS_ROOT, RUNTIME_GLOBALS_NS, dom) Then Exit Function
    Set rootNode = dom.DocumentElement
    If rootNode Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "RestoreManager: runtime globals root node is missing."
#End If
        Exit Function
    End If

    activeWorksheetName = VBA.vbNullString
    If private_TryGetActiveWorksheetName(activeWorksheetName) Then
        rootNode.setAttribute RUNTIME_GLOBALS_ACTIVE_SHEET_ATTR, activeWorksheetName
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "restore-manager:save-active-sheet sheet='" & VBA.Replace$(activeWorksheetName, "'", "''") & "'"
#End If
    End If

    If Not private_TryAppendModuleSnapshot(rootNode, MODULE_NAME_PAGE_MANAGER) Then Exit Function

    If Not ex_Core.fn_CustomXmlPartStore_TryFindPartByNamespace(RUNTIME_GLOBALS_NS, partObj) Then Exit Function
    If Not ex_Core.fn_CustomXmlPartStore_TrySaveDom(dom, partObj) Then Exit Function

    private_TrySaveRuntimeGlobalsSnapshot = True
End Function

Private Function private_TryRestoreRuntimeGlobalsSnapshot(ByRef outActiveWorksheetName As String) As Boolean
    Dim partObj As Object
    Dim dom As Object
    Dim rootNode As Object
    Dim moduleNodes As Object
    Dim moduleNode As Object
    Dim snapshotNode As Object
    Dim moduleName As String
    Dim snapshotXml As String
    Dim attrValue As Variant

    outActiveWorksheetName = VBA.vbNullString

    If Not ex_Core.fn_CustomXmlPartStore_TryFindPartByNamespace(RUNTIME_GLOBALS_NS, partObj) Then Exit Function
    If partObj Is Nothing Then
        ' Отсутствие runtime snapshot на первой загрузке — нормальный сценарий.
        private_TryRestoreRuntimeGlobalsSnapshot = True
        Exit Function
    End If

    If Not ex_Core.fn_CustomXmlPartStore_TryLoadPartDom(partObj, dom) Then Exit Function
    Set rootNode = dom.DocumentElement
    If rootNode Is Nothing Then
        private_TryRestoreRuntimeGlobalsSnapshot = True
        Exit Function
    End If
    On Error Resume Next
    attrValue = rootNode.getAttribute(RUNTIME_GLOBALS_ACTIVE_SHEET_ATTR)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        attrValue = VBA.vbNullString
    End If
    On Error GoTo 0
    If Not IsNull(attrValue) Then
        outActiveWorksheetName = VBA.Trim$(VBA.CStr(attrValue))
    End If

    Set moduleNodes = rootNode.selectNodes("*[local-name()='" & RUNTIME_GLOBALS_MODULE_NODE & "']")
    If moduleNodes Is Nothing Then
        private_TryRestoreRuntimeGlobalsSnapshot = True
        Exit Function
    End If

    For Each moduleNode In moduleNodes
        attrValue = VBA.vbNullString
        On Error Resume Next
        attrValue = moduleNode.getAttribute(RUNTIME_GLOBALS_MODULE_NAME_ATTR)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            attrValue = VBA.vbNullString
        End If
        On Error GoTo 0
        If IsNull(attrValue) Then
            moduleName = VBA.vbNullString
        Else
            moduleName = VBA.Trim$(VBA.CStr(attrValue))
        End If
        If VBA.Len(moduleName) = 0 Then GoTo ContinueModule

        snapshotXml = VBA.vbNullString
        Set snapshotNode = moduleNode.selectSingleNode("*[local-name()='" & RUNTIME_GLOBALS_MODULE_SNAPSHOT_NODE & "']")
        If Not snapshotNode Is Nothing Then
            snapshotXml = VBA.CStr(snapshotNode.Text)
        End If

        ' Module-level deserialize выполняется строго через диспетчер по имени модуля.
        If Not private_TryDeserializeRuntimeModuleSnapshot(moduleName, snapshotXml) Then Exit Function
ContinueModule:
    Next moduleNode

    private_TryRestoreRuntimeGlobalsSnapshot = True
End Function

Private Function private_TryAppendModuleSnapshot(ByVal rootNode As Object, ByVal moduleName As String) As Boolean
    Dim snapshotXml As String
    Dim moduleNode As Object
    Dim snapshotNode As Object
    Dim dom As Object

    If rootNode Is Nothing Then Exit Function
    moduleName = VBA.Trim$(moduleName)
    If VBA.Len(moduleName) = 0 Then Exit Function

    snapshotXml = VBA.vbNullString
    If Not private_TrySerializeRuntimeModuleSnapshot(moduleName, snapshotXml) Then Exit Function

    Set dom = rootNode.OwnerDocument
    If dom Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "RestoreManager: runtime globals DOM owner is not available."
#End If
        Exit Function
    End If

    Set moduleNode = dom.createNode(1, RUNTIME_GLOBALS_MODULE_NODE, RUNTIME_GLOBALS_NS)
    moduleNode.setAttribute RUNTIME_GLOBALS_MODULE_NAME_ATTR, moduleName

    Set snapshotNode = dom.createElement(RUNTIME_GLOBALS_MODULE_SNAPSHOT_NODE)
    snapshotNode.Text = VBA.CStr(snapshotXml)
    moduleNode.appendChild snapshotNode

    rootNode.appendChild moduleNode
    private_TryAppendModuleSnapshot = True
End Function

Private Function private_TrySerializeRuntimeModuleSnapshot(ByVal moduleName As String, ByRef outSnapshotXml As String) As Boolean
    outSnapshotXml = VBA.vbNullString
    moduleName = VBA.LCase$(VBA.Trim$(moduleName))

    ' Таблица диспетчеризации сериализации runtime-модулей.
    Select Case moduleName
        Case VBA.LCase$(MODULE_NAME_PAGE_MANAGER)
            private_TrySerializeRuntimeModuleSnapshot = rt_PageManager.fn_TrySerializeModuleSnapshot(outSnapshotXml)

        Case Else
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogInfo "restore-manager: serialize skipped for unknown module '" & moduleName & "'"
#End If
            ' Неизвестные модули пропускаем мягко, чтобы не ломать совместимость.
            private_TrySerializeRuntimeModuleSnapshot = True
    End Select
End Function

Private Function private_TryDeserializeRuntimeModuleSnapshot(ByVal moduleName As String, ByVal snapshotXml As String) As Boolean
    moduleName = VBA.LCase$(VBA.Trim$(moduleName))

    ' Таблица диспетчеризации десериализации runtime-модулей.
    Select Case moduleName
        Case VBA.LCase$(MODULE_NAME_PAGE_MANAGER)
            private_TryDeserializeRuntimeModuleSnapshot = rt_PageManager.fn_TryDeserializeModuleSnapshot(snapshotXml)

        Case Else
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogInfo "restore-manager: deserialize skipped for unknown module '" & moduleName & "'"
#End If
            ' Неизвестные модули пропускаем мягко, чтобы старые snapshot-части
            ' не блокировали запуск новой версии runtime.
            private_TryDeserializeRuntimeModuleSnapshot = True
    End Select
End Function

Private Function private_HasRuntimePageForActiveWorksheet() As Boolean
    Dim ws As Worksheet
    Dim page As obj_IPage

    On Error Resume Next
    If TypeOf Application.ActiveSheet Is Worksheet Then
        Set ws = Application.ActiveSheet
    End If
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    private_HasRuntimePageForActiveWorksheet = rt_PageManager.fn_TryGetPageByWorksheet(ws, page)
End Function

Private Function private_TryGetPagesCount(ByRef outPagesCount As Long) As Boolean
    outPagesCount = 0
    private_TryGetPagesCount = rt_PageManager.fn_TryGetPagesCount(outPagesCount)
End Function

Private Function private_TryGetActiveWorksheetName(ByRef outWorksheetName As String) As Boolean
    Dim ws As Worksheet

    outWorksheetName = VBA.vbNullString

    On Error Resume Next
    If TypeOf Application.ActiveSheet Is Worksheet Then
        Set ws = Application.ActiveSheet
    End If
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If ws Is Nothing Then Exit Function
    outWorksheetName = VBA.Trim$(VBA.CStr(ws.Name))
    private_TryGetActiveWorksheetName = (VBA.Len(outWorksheetName) > 0)
End Function

Private Function private_ShouldApplySavedWorksheetFocusAfterRestore(ByVal reasonText As String) As Boolean
    reasonText = VBA.LCase$(VBA.Trim$(reasonText))
    Select Case reasonText
        Case "after-update", "deferred:on-time"
            private_ShouldApplySavedWorksheetFocusAfterRestore = True
    End Select
End Function

Private Sub private_TryActivateSavedWorksheetAfterRestore( _
    ByVal reasonText As String, _
    ByVal worksheetName As String _
)
    Dim ws As Worksheet
    Dim activeSheetName As String
    Dim errDescription As String

    If Not private_ShouldApplySavedWorksheetFocusAfterRestore(reasonText) Then Exit Sub

    worksheetName = VBA.Trim$(worksheetName)
    If VBA.Len(worksheetName) = 0 Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(worksheetName)
    If Err.Number <> 0 Then
        errDescription = Err.Description
        Err.Clear
        On Error GoTo 0
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "restore-manager:restore-active-sheet lookup-failed sheet='" & VBA.Replace$(worksheetName, "'", "''") & "' err='" & VBA.Replace$(errDescription, "'", "''") & "'"
#End If
        Exit Sub
    End If
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Activate
    If Err.Number <> 0 Then
        errDescription = Err.Description
        Err.Clear
        On Error GoTo 0
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "restore-manager:restore-active-sheet activate-failed sheet='" & VBA.Replace$(worksheetName, "'", "''") & "' err='" & VBA.Replace$(errDescription, "'", "''") & "'"
#End If
        Exit Sub
    End If
    On Error GoTo 0

    If Not private_TryGetActiveWorksheetName(activeSheetName) Then Exit Sub
    If VBA.StrComp(activeSheetName, worksheetName, VBA.vbTextCompare) <> 0 Then Exit Sub

    Call rt_PageManager.fn_TryRestoreLastRenderedWorksheetName(worksheetName)
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "restore-manager:restore-active-sheet done sheet='" & VBA.Replace$(worksheetName, "'", "''") & "' reason='" & VBA.Replace$(reasonText, "'", "''") & "'"
#End If
End Sub

Private Function private_TryFallbackRestoreByResettingMainPage( _
    ByVal reasonText As String, _
    ByRef outRestoredPagesCount As Long _
) As Boolean
    Dim saveRuntimeOk As Boolean
    Dim errDescription As String

    reasonText = VBA.Trim$(reasonText)
    If VBA.Len(reasonText) = 0 Then reasonText = "unknown"

    ' Жесткий fallback: если restore runtime state не удался,
    ' пересоздаем Main страницу и сразу сохраняем новый checkpoint.
    On Error Resume Next
    If Not ThisWorkbook.m_ResetWorkbookAndCreateMainPage("rt_RestoreManager:restore-runtime-state:fallback-main-reset", False) Then
        If Err.Number <> 0 Then
            errDescription = Err.Description
            Err.Clear
            On Error GoTo 0
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "restore-manager:restore-runtime-state fallback-main-reset failed reason='" & VBA.Replace$(reasonText, "'", "''") & "' err='" & VBA.Replace$(errDescription, "'", "''") & "'"
#End If
            Exit Function
        End If
        On Error GoTo 0
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "restore-manager:restore-runtime-state fallback-main-reset failed reason='" & VBA.Replace$(reasonText, "'", "''") & "' err='returned-false'"
#End If
        Exit Function
    End If
    If Err.Number <> 0 Then
        errDescription = Err.Description
        Err.Clear
        On Error GoTo 0
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "restore-manager:restore-runtime-state fallback-main-reset failed reason='" & VBA.Replace$(reasonText, "'", "''") & "' err='" & VBA.Replace$(errDescription, "'", "''") & "'"
#End If
        Exit Function
    End If
    On Error GoTo 0

    outRestoredPagesCount = 1
    saveRuntimeOk = fn_SaveRuntimeState()

    If Not saveRuntimeOk Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "restore-manager:restore-runtime-state fallback-checkpoint-failed reason='" & VBA.Replace$(reasonText, "'", "''") & "' saveRuntime=" & VBA.CStr(saveRuntimeOk)
#End If
    End If

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "restore-manager:restore-runtime-state fallback-main-reset done reason='" & VBA.Replace$(reasonText, "'", "''") & "' restoredPages=" & VBA.CStr(outRestoredPagesCount)
#End If
    private_TryFallbackRestoreByResettingMainPage = True
End Function
