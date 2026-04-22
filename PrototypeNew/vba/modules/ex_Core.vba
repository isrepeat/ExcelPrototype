' Должен быть вставлен во внутренний модуль книги .xlsm
Option Explicit

#Const CORE_ENABLE_STATUS_BAR_LOGGING = True
#Const CORE_FORCE_NATIVE_STATUS_BAR = True
#Const CORE_ENABLE_SELF_LOGGING = True
#Const CORE_ENABLE_LOGGING = True

Private Const BASE_DIR As String = "vba\\"
Private Const IMPORT_CACHE_FILE As String = ".devtools_import_cache.txt"
Private Const CORE_LOG_FILE_REL_PATH As String = "Logs\\core.log"
Private Const ENABLE_CLASS_IMPORT_VALIDATION As Boolean = False
Private Const MAX_IMPORT_RECURSION_DEPTH As Long = 4
Private Const COMP_TYPE_MODULE As String = "module"
Private Const COMP_TYPE_CLASS As String = "class"
Private Const COMP_TYPE_SHEET As String = "sheet"
Private Const COMP_TYPE_WORKBOOK As String = "workbook"
Private Const MAX_VBA_COMPONENT_NAME_LEN As Long = 31
Private Const UPDATE_MODE_FULL As Long = 1
Private Const UPDATE_MODE_DATE As Long = 2
Private Const UPDATE_MODE_SIZE As Long = 3
Private Const CORE_COMPONENT_NAME As String = "ex_Core"
Private Const PATTERN_ALL_COMPONENTS As String = ".+"
Private Const PATTERN_MAIN_COMPONENTS As String = "^(?!rt_).+"
Private Const PATTERN_RUNTIME_COMPONENTS As String = "^rt_.+"
Private Const PATTERN_EXCLUDE_CORE As String = "^ex_core$"
Private g_QueuedBridgeUpdateAt As Date
Private g_QueuedBridgeUpdateMacro As String
Private g_QueuedRuntimeStateRestoreAt As Date
Private g_QueuedRuntimeStateRestoreMacro As String

'==========================
' Публичные процедуры
'==========================
' --- Загрузка модулей

' //
' // API
' //
Public Sub dev_OpenCoreModule()
    Dim comp As Object
    Dim cp As Object

    On Error GoTo EH

    Application.VBE.MainWindow.Visible = True

    ' Закрываем все открытые окна кода перед открытием целевого модуля.
    On Error Resume Next
    For Each cp In Application.VBE.CodePanes
        cp.Window.Close
    Next cp
    On Error GoTo EH

    Set comp = ThisWorkbook.VBProject.VBComponents(CORE_COMPONENT_NAME)

    If comp Is Nothing Then
        private_ShowStatusWarning "Модуль ex_Core не найден в проекте VBA.", True, 6
        Exit Sub
    End If

    comp.Activate
    private_ShowStatusNotice "Открыт модуль ex_Core в редакторе VBA.", True, 2
    Exit Sub

EH:
    private_ShowStatusError "Не удалось открыть модуль ex_Core: " & Err.Description, True, 6
End Sub


Public Sub dev_RemoveAllModulesAndClasses()
    On Error GoTo EH
    Application.ScreenUpdating = False

    private_RemoveAllModulesAndClasses PATTERN_ALL_COMPONENTS, PATTERN_EXCLUDE_CORE

    Application.ScreenUpdating = True
    private_ShowStatusSuccess "Модули и классы удалены, объектные модули очищены (ex_Core сохранен).", True, 3
    private_LogCoreSelfEvent "remove-modules-classes: success"
    Exit Sub

EH:
    Application.ScreenUpdating = True
    private_ShowStatusError "Ошибка удаления модулей/классов: " & Err.Description, True, 6
    private_LogCoreSelfEvent "remove-modules-classes: fail: " & Err.Description
End Sub

' Callstack[1]: VBA.Macros(ex_Core.dev_UpdateAllModules) -> ex_Core.dev_UpdateAllModules
' Callstack[2]: DevUI(onClick: ex_Core.dev_UpdateAllModules) -> ex_Core.dev_UpdateAllModules
Public Sub dev_UpdateAllModules()
    ' Если обновление запущено из bridge-click dispatch, переносим запуск на следующий тик.
    ' Это защищает от reentry: нельзя безопасно переимпортировать модули в середине обработки клика.
    If private_TryQueueRuntimeUpdateWhenBridgeDispatch("full") Then Exit Sub
    If private_TryRunSafeUpdateByMode(UPDATE_MODE_FULL, PATTERN_ALL_COMPONENTS, PATTERN_EXCLUDE_CORE, True, "full") Then Exit Sub
    private_ShowStatusError "Безопасное обновление (full) не завершено. Проверьте core.log.", True, 6
End Sub


' Callstack[1]: VBA.Macros(ex_Core.dev_UpdateCodeByDate) -> ex_Core.dev_UpdateCodeByDate
Public Sub dev_UpdateCodeByDate()
    ' Та же логика, что и для full: если мы внутри bridge-dispatch, только очередь через OnTime.
    If private_TryQueueRuntimeUpdateWhenBridgeDispatch("date") Then Exit Sub
    If private_TryRunSafeUpdateByMode(UPDATE_MODE_DATE, PATTERN_MAIN_COMPONENTS, PATTERN_EXCLUDE_CORE, False, "date") Then Exit Sub
    private_ShowStatusError "Безопасное обновление (date) не завершено. Проверьте core.log.", True, 6
End Sub


' Callstack[1]: VBA.Macros(ex_Core.dev_UpdateCodeBySize) -> ex_Core.dev_UpdateCodeBySize
Public Sub dev_UpdateCodeBySize()
    ' Та же логика, что и для full/date: отложенный запуск только если идет dispatch клика.
    If private_TryQueueRuntimeUpdateWhenBridgeDispatch("size") Then Exit Sub
    If private_TryRunSafeUpdateByMode(UPDATE_MODE_SIZE, PATTERN_MAIN_COMPONENTS, PATTERN_EXCLUDE_CORE, False, "size") Then Exit Sub
    private_ShowStatusError "Безопасное обновление (size) не завершено. Проверьте core.log.", True, 6
End Sub


' Отдельная процедура для ядра рантайма (rt_*):
' Инкрементальные обновления по дате/размеру эти модули не затрагивают.
Public Sub dev_UpdateRuntimeCore()
    private_UpdateCodeByRegex PATTERN_RUNTIME_COMPONENTS, PATTERN_EXCLUDE_CORE, UPDATE_MODE_FULL, True
End Sub


' --- Логирование
Public Sub m_LogInfo(ByVal messageText As String)
    private_LogCoreEvent VBA.CStr(messageText)
End Sub


Public Sub m_LogError(ByVal messageText As String)
    messageText = VBA.Trim$(VBA.CStr(messageText))
    If VBA.Len(messageText) = 0 Then Exit Sub
    private_LogCoreEvent "error: " & messageText
End Sub


Public Sub m_LogStatusBarMessage( _
    ByVal actionName As String, _
    ByVal messageText As String, _
    Optional ByVal timeoutSeconds As Long = 0 _
)
    private_LogStatusBarEvent actionName, messageText, timeoutSeconds
End Sub


Public Sub m_ClearCoreLog()
    private_ClearCoreLogFile
End Sub

' //
' // Internal
' //
'==========================
' Приватные процедуры
'==========================
' --- Загрузка модулей (ядро импорта)
Private Function private_TryQueueRuntimeUpdateWhenBridgeDispatch(ByVal updateKind As String) As Boolean
    Dim isBridgeDispatching As Boolean
    Dim callResult As Variant
    Dim bridgeComponent As Object
    Dim coreMethod As String
    Dim macroRef As String
    Dim scheduleAt As Date
    Dim errDescription As String

    updateKind = VBA.LCase$(VBA.Trim$(updateKind))
    If VBA.Len(updateKind) = 0 Then Exit Function

    ' Эту очередь используем только в одном случае:
    ' пользователь нажал кнопку обновления из UI, и мы сейчас внутри rt_Bridge dispatch.
    ' Тогда переносим обновление на OnTime, чтобы сначала завершить текущий click pipeline.
    ' Если rt_Bridge еще не загружен (bootstrap/cold start), это не ошибка:
    ' просто выполняем обновление сразу, без очереди.
    Set bridgeComponent = private_TryGetComponentByName("rt_Bridge")
    If bridgeComponent Is Nothing Then Exit Function

    If Not private_TryRunRuntimeNoArgMember("rt_Bridge", "m_IsDispatchingClick", callResult, True) Then
        private_LogCoreSelfEvent "queued-runtime-update-failed: kind='" & VBA.Replace$(updateKind, "'", "''") & "' err='bridge-dispatch-state-read-failed'"
        Exit Function
    End If

    On Error Resume Next
    isBridgeDispatching = VBA.CBool(callResult)
    If Err.Number <> 0 Then
        errDescription = Err.Description
        Err.Clear
        On Error GoTo 0
        private_LogCoreSelfEvent "queued-runtime-update-failed: kind='" & VBA.Replace$(updateKind, "'", "''") & "' err='bridge-dispatch-state-cast-failed: " & VBA.Replace$(errDescription, "'", "''") & "'"
        Exit Function
    End If
    On Error GoTo 0

    If Not isBridgeDispatching Then Exit Function

    Select Case updateKind
        Case "full"
            coreMethod = "dev_UpdateAllModules"
        Case "date"
            coreMethod = "dev_UpdateCodeByDate"
        Case "size"
            coreMethod = "dev_UpdateCodeBySize"
        Case Else
            Exit Function
    End Select

    macroRef = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!ex_Core." & coreMethod
    scheduleAt = private_GetNextOnTimeTick()

    On Error Resume Next
    ' Держим только последнюю отложенную задачу обновления.
    If g_QueuedBridgeUpdateAt > 0# And VBA.Len(VBA.Trim$(g_QueuedBridgeUpdateMacro)) > 0 Then
        Application.OnTime EarliestTime:=g_QueuedBridgeUpdateAt, Procedure:=g_QueuedBridgeUpdateMacro, Schedule:=False
        Err.Clear
    End If

    Application.OnTime EarliestTime:=scheduleAt, Procedure:=macroRef
    If Err.Number <> 0 Then
        errDescription = Err.Description
        Err.Clear
        On Error GoTo 0
        private_LogCoreSelfEvent "queued-runtime-update-failed: kind='" & VBA.Replace$(updateKind, "'", "''") & "' err='" & VBA.Replace$(errDescription, "'", "''") & "'"
        Exit Function
    End If
    On Error GoTo 0

    g_QueuedBridgeUpdateAt = scheduleAt
    g_QueuedBridgeUpdateMacro = macroRef

    private_LogCoreSelfEvent "queued-runtime-update: kind='" & VBA.Replace$(updateKind, "'", "''") & "'"
    private_TryQueueRuntimeUpdateWhenBridgeDispatch = True
End Function


' Callstack[1]: ex_Core.dev_UpdateAllModules -> private_TryRunSafeUpdateByMode
' Callstack[2]: ex_Core.dev_UpdateCodeByDate -> private_TryRunSafeUpdateByMode
' Callstack[3]: ex_Core.dev_UpdateCodeBySize -> private_TryRunSafeUpdateByMode
Private Function private_TryRunSafeUpdateByMode( _
    ByVal updateMode As Long, _
    ByVal includeComponentPattern As String, _
    ByVal excludeComponentPattern As String, _
    ByVal useNativeStatus As Boolean, _
    ByVal operationName As String _
) As Boolean
    Dim bootstrapMode As String
    Dim savePagesOk As Boolean
    Dim saveRuntimeOk As Boolean

    ' Этап 0. Нормализуем имя операции для логов.
    operationName = VBA.LCase$(VBA.Trim$(operationName))
    If VBA.Len(operationName) = 0 Then operationName = "unknown"

    private_LogCoreSelfEvent "safe-update:start op='" & operationName & "'"

    ' Этап 1. Гарантируем, что runtime-пайплайн вообще доступен.
    ' Важно: rt_Snapshots/rt_PageManager могут отсутствовать (например, после частичного импорта/сброса проекта),
    ' поэтому сохранить snapshot "до любых действий" не всегда возможно.
    ' bootstrap сначала поднимает минимально нужные runtime-компоненты.
    If Not private_TryBootstrapRuntimePipeline(bootstrapMode) Then
        private_LogCoreSelfEvent "safe-update:fail op='" & operationName & "' reason='runtime-bootstrap-failed'"
        Exit Function
    End If

    ' Если bootstrap был "full", массовый импорт уже выполнен.
    ' Упрощенный путь: не делаем sync-recovery в этом стеке, а сразу планируем
    ' единое deferred-восстановление snapshots/globals на следующий тик OnTime.
    ' Это выравнивает flow с обычной веткой (save/import -> deferred restore).
    If VBA.StrComp(bootstrapMode, "full", VBA.vbTextCompare) = 0 Then
        private_LogCoreSelfEvent "safe-update:deferred op='" & operationName & "' reason='full-bootstrap-was-required'"
        Call private_QueueRuntimeStateRestoreAfterUpdate("safe-update:bootstrap:" & operationName)
        private_LogCoreSelfEvent "safe-update:done op='" & operationName & "'"
        private_TryRunSafeUpdateByMode = True
        Exit Function
    End If

    ' Этап 2. Runtime валиден -> сохраняем page snapshots перед целевым update.
    ' Если save не удался, безопасный update прекращаем (без перехода в unsafe fallback).
    If Not private_TryRunRuntimeBooleanFunction("rt_Snapshots", "m_SavePageSnapshots", savePagesOk) Then
        private_LogCoreSelfEvent "safe-update:fail op='" & operationName & "' reason='save-pages-call-failed'"
        Exit Function
    End If
    If Not savePagesOk Then
        private_LogCoreSelfEvent "safe-update:fail op='" & operationName & "' reason='save-pages-returned-false'"
        Exit Function
    End If

    ' Этап 3. Сохраняем runtime-глобалы (не только страницы, но и состояние runtime-модулей).
    If Not private_TryRunRuntimeBooleanFunction("rt_Snapshots", "m_SaveRuntimeGlobalsSnapshot", saveRuntimeOk) Then
        private_LogCoreSelfEvent "safe-update:fail op='" & operationName & "' reason='save-runtime-call-failed'"
        Exit Function
    End If
    If Not saveRuntimeOk Then
        private_LogCoreSelfEvent "safe-update:fail op='" & operationName & "' reason='save-runtime-returned-false'"
        Exit Function
    End If

    ' Этап 4. Выполняем фактический импорт/обновление модулей.
    private_UpdateCodeByRegex includeComponentPattern, excludeComponentPattern, updateMode, useNativeStatus

    ' Этап 5. Не делаем синхронный restore прямо здесь.
    ' Восстановление целиком переносится в deferred-путь на следующий тик OnTime,
    ' чтобы избежать двойного рендера (sync restore + deferred restore) после hot-update.
    Call private_QueueRuntimeStateRestoreAfterUpdate("safe-update:" & operationName)
    private_LogCoreSelfEvent "safe-update:done op='" & operationName & "'"
    private_TryRunSafeUpdateByMode = True
End Function


' Callstack[1]: ex_Core.dev_UpdateAllModules -> private_TryBootstrapRuntimePipeline
' Callstack[2]: ex_Core.dev_UpdateCodeByDate -> private_TryBootstrapRuntimePipeline
' Callstack[3]: ex_Core.dev_UpdateCodeBySize -> private_TryBootstrapRuntimePipeline
Private Function private_TryBootstrapRuntimePipeline(ByRef outBootstrapMode As String) As Boolean
    outBootstrapMode = VBA.vbNullString

    ' Нормальный путь: runtime-компоненты уже на месте, ничего доп. не делаем.
    If private_AreSafeUpdateRuntimeComponentsPresent() Then
        outBootstrapMode = "none"
        private_TryBootstrapRuntimePipeline = True
        Exit Function
    End If

    ' Аварийный путь: runtime невалиден.
    ' Для rt_Snapshots нужны зависимости из ex_*/obj_*, поэтому поднимаем полный набор.
    outBootstrapMode = "full"
    private_LogCoreSelfEvent "runtime-update-pipeline-bootstrap: start scope='all-components'"
    private_UpdateCodeByRegex PATTERN_ALL_COMPONENTS, PATTERN_EXCLUDE_CORE, UPDATE_MODE_FULL, True

    If Not private_AreSafeUpdateRuntimeComponentsPresent() Then
        private_LogCoreSelfEvent "runtime-update-pipeline-bootstrap: fail component-not-found-after-import"
        Exit Function
    End If

    private_LogCoreSelfEvent "runtime-update-pipeline-bootstrap: done scope='all-components'"
    private_TryBootstrapRuntimePipeline = True
End Function


Private Function private_AreSafeUpdateRuntimeComponentsPresent() As Boolean
    ' Минимальный контракт для safe-update: если чего-то из этого нет,
    ' snapshot-сценарий нельзя считать надежным.
    If private_TryGetComponentByName("rt_Snapshots") Is Nothing Then Exit Function
    If private_TryGetComponentByName("rt_PageManager") Is Nothing Then Exit Function
    If private_TryGetComponentByName("ex_CustomXmlPartStore") Is Nothing Then Exit Function
    If private_TryGetComponentByName("ex_HelpersSheet") Is Nothing Then Exit Function
    If private_TryGetComponentByName("obj_PageBase") Is Nothing Then Exit Function
    If private_TryGetComponentByName("obj_IPage") Is Nothing Then Exit Function
    If private_TryGetComponentByName("obj_ISerializable") Is Nothing Then Exit Function
    private_AreSafeUpdateRuntimeComponentsPresent = True
End Function


' Callstack[1]: ex_Core.private_TryRunSafeUpdateByMode -> private_QueueRuntimeStateRestoreAfterUpdate
Private Sub private_QueueRuntimeStateRestoreAfterUpdate(ByVal reasonText As String)
    Dim macroRef As String
    Dim scheduleAt As Date
    Dim errDescription As String

    reasonText = VBA.Trim$(reasonText)
    If VBA.Len(reasonText) = 0 Then reasonText = "unknown"

    macroRef = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!rt_Snapshots.m_RunDeferredRuntimeStateRestore"
    scheduleAt = private_GetNextOnTimeTick()

    On Error Resume Next
    ' Держим только последнюю задачу deferred restore.
    If g_QueuedRuntimeStateRestoreAt > 0# And VBA.Len(VBA.Trim$(g_QueuedRuntimeStateRestoreMacro)) > 0 Then
        Application.OnTime EarliestTime:=g_QueuedRuntimeStateRestoreAt, Procedure:=g_QueuedRuntimeStateRestoreMacro, Schedule:=False
        Err.Clear
    End If

    Application.OnTime EarliestTime:=scheduleAt, Procedure:=macroRef
    If Err.Number <> 0 Then
        errDescription = Err.Description
        Err.Clear
        On Error GoTo 0
        private_LogCoreSelfEvent "runtime-state-restore-queue-failed reason='" & VBA.Replace$(reasonText, "'", "''") & "' err='" & VBA.Replace$(errDescription, "'", "''") & "'"
        Exit Sub
    End If
    On Error GoTo 0

    g_QueuedRuntimeStateRestoreAt = scheduleAt
    g_QueuedRuntimeStateRestoreMacro = macroRef
    private_LogCoreSelfEvent "runtime-state-restore-queued reason='" & VBA.Replace$(reasonText, "'", "''") & "'"
End Sub


Private Sub private_UpdateCodeByRegex( _
    ByVal includeComponentPattern As String, _
    Optional ByVal excludeComponentPattern As String = VBA.vbNullString, _
    Optional ByVal updateMode As Long = UPDATE_MODE_FULL, _
    Optional ByVal useNativeStatus As Boolean = False _
)
    private_UpdateCodeCore updateMode, useNativeStatus, includeComponentPattern, excludeComponentPattern
End Sub


Private Sub private_UpdateCodeCore( _
    ByVal updateMode As Long, _
    ByVal useNativeStatus As Boolean, _
    Optional ByVal includeComponentPattern As String = VBA.vbNullString, _
    Optional ByVal excludeComponentPattern As String = VBA.vbNullString _
)
    Dim basePath As String
    Dim cachePath As String
    Dim prevCache As Object
    Dim nextCache As Object
    Dim incrementalMode As Boolean
    Dim stageName As String
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim fullErrorText As String

    stageName = "init"
    incrementalMode = (updateMode <> UPDATE_MODE_FULL)

    private_ShowStatusNotice "Запущено обновление кода...", useNativeStatus, 1
    private_LogCoreSelfEvent "update-start"

    basePath = ThisWorkbook.Path & "\\" & BASE_DIR
    If VBA.Len(Dir(basePath, vbDirectory)) = 0 Then
        private_ShowStatusWarning "Путь книги пустой или папка vba не найдена. Сначала сохраните файл.", useNativeStatus, 6
        private_LogCoreSelfEvent "update-stop: vba-folder-not-found"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo EH

    stageName = "load-cache"
    cachePath = basePath & IMPORT_CACHE_FILE
    Set prevCache = private_LoadImportCache(cachePath)
    Set nextCache = private_CreateDictionary()

    If Not incrementalMode Then
        stageName = "remove-imported-by-scope"
        private_RemoveImportedModulesByScope includeComponentPattern, excludeComponentPattern
    End If

    stageName = "import-folder"
    private_ImportFolder basePath, updateMode, prevCache, nextCache, includeComponentPattern, excludeComponentPattern
    If ENABLE_CLASS_IMPORT_VALIDATION Then
        stageName = "validate-class-imports"
        private_ValidateClassImports basePath
    End If

    If incrementalMode Then
        stageName = "remove-stale"
        private_RemoveStaleImportedComponentsByScope prevCache, nextCache, includeComponentPattern, excludeComponentPattern
    End If
    stageName = "preserve-out-of-scope-cache"
    private_PreserveOutOfScopeCacheRecords prevCache, nextCache, includeComponentPattern, excludeComponentPattern
    stageName = "save-cache"
    private_SaveImportCache cachePath, nextCache

    Application.ScreenUpdating = True
    private_ShowCodeUpdatedNotice useNativeStatus
    private_LogCoreSelfEvent "update-done"
    Exit Sub

EH:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    fullErrorText = "Ошибка обновления кода на этапе '" & stageName & "': [" & errSource & " #" & VBA.CStr(errNumber) & "] " & errDescription

    Application.ScreenUpdating = True
    private_ShowStatusError fullErrorText, useNativeStatus, 6

    ' Статус-бар часто обрезает длинный текст ошибки импорта.
    ' Показываем полную диагностику (включая список файлов) отдельным окном.
    VBA.MsgBox fullErrorText, VBA.vbExclamation, "PrototypeNew: Update Code"

    ' Логируем ошибку напрямую в core.log, даже если CORE_ENABLE_SELF_LOGGING = False.
    private_LogCoreEvent "update-fail: stage='" & stageName & "' err='" & VBA.Replace$(errDescription, "'", "''") & "'"
End Sub


Private Sub private_PreserveOutOfScopeCacheRecords( _
    ByVal prevCache As Object, _
    ByVal nextCache As Object, _
    Optional ByVal includeComponentPattern As String = VBA.vbNullString, _
    Optional ByVal excludeComponentPattern As String = VBA.vbNullString _
)
    Dim key As Variant
    Dim rec As Object
    Dim componentName As String

    If prevCache Is Nothing Then Exit Sub
    If nextCache Is Nothing Then Exit Sub

    For Each key In prevCache.Keys
        If nextCache.Exists(VBA.CStr(key)) Then GoTo ContinueKey

        Set rec = prevCache(VBA.CStr(key))
        If rec Is Nothing Then GoTo ContinueKey
        If Not rec.Exists("Name") Then GoTo ContinueKey

        componentName = VBA.CStr(rec("Name"))
    If private_ShouldProcessComponentByScope(componentName, includeComponentPattern, excludeComponentPattern) Then GoTo ContinueKey

        nextCache.Add VBA.CStr(key), rec
ContinueKey:
    Next key
End Sub


Private Sub private_ValidateClassImports(ByVal rootPath As String)
    Dim fso As Object
    Dim failed As String

    If Dir(rootPath, vbDirectory) = "" Then
        Err.Raise VBA.vbObjectError + 1006, "private_ValidateClassImports", "VBA root folder not found: " & rootPath
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    private_ValidateClassImportsRecursive fso.GetFolder(rootPath), 0, failed

    If VBA.Len(failed) > 0 Then
        Err.Raise VBA.vbObjectError + 1007, "private_ValidateClassImports", "Class import validation failed:" & failed
    End If
End Sub


Private Sub private_ValidateClassImportsRecursive( _
    ByVal folderObj As Object, _
    ByVal depth As Long, _
    ByRef failed As String _
)
    Dim fileObj As Object
    Dim subFolder As Object
    Dim compType As String
    Dim fallbackName As String
    Dim className As String
    Dim vbComp As Object

    If folderObj Is Nothing Then Exit Sub
    If depth > MAX_IMPORT_RECURSION_DEPTH Then Exit Sub

    For Each fileObj In folderObj.Files
        If Not private_TryResolveFileComponentType(VBA.CStr(fileObj.Name), compType, fallbackName) Then GoTo ContinueFile
        If VBA.StrComp(compType, COMP_TYPE_CLASS, VBA.vbTextCompare) <> 0 Then GoTo ContinueFile

        className = private_GetComponentNameFromSource(VBA.CStr(fileObj.Path))
        Set vbComp = Nothing
        On Error Resume Next
        Set vbComp = ThisWorkbook.VBProject.VBComponents(className)
        On Error GoTo 0

        If vbComp Is Nothing Then
            failed = failed & VBA.vbCrLf & "- missing class: " & className
        ElseIf vbComp.Type <> 2 Then ' модуль не является модулем класса
            failed = failed & VBA.vbCrLf & "- wrong component type for class '" & className & "': " & VBA.CStr(vbComp.Type)
        End If

ContinueFile:
    Next fileObj

    For Each subFolder In folderObj.SubFolders
        private_ValidateClassImportsRecursive subFolder, depth + 1, failed
    Next subFolder
End Sub


Private Sub private_ShowCodeUpdatedNotice(ByVal useNativeStatus As Boolean)
    private_ShowStatusSuccess "Код обновлен.", useNativeStatus, 1
End Sub


'==========================
' Управление модулями
'==========================
Private Sub private_RemoveAllModulesAndClasses( _
    Optional ByVal includeComponentPattern As String = VBA.vbNullString, _
    Optional ByVal excludeComponentPattern As String = VBA.vbNullString _
)
    Dim prj As Object
    Dim comp As Object
    Dim names() As String
    Dim docNames() As String
    Dim n As Long
    Dim d As Long
    Dim i As Long

    Set prj = ThisWorkbook.VBProject

    For Each comp In prj.VBComponents
        Select Case comp.Type
            Case 1, 2 ' стандартный модуль, модуль класса
                If private_ShouldProcessComponentByScope(VBA.CStr(comp.Name), includeComponentPattern, excludeComponentPattern) Then
                    n = n + 1
                    ReDim Preserve names(1 To n)
                    names(n) = comp.Name
                End If

            Case 100 ' модуль документа (книга/листы)
                If private_ShouldProcessComponentByScope(VBA.CStr(comp.Name), includeComponentPattern, excludeComponentPattern) Then
                    d = d + 1
                    ReDim Preserve docNames(1 To d)
                    docNames(d) = VBA.CStr(comp.Name)
                End If
        End Select
    Next comp

    For i = 1 To n
        On Error GoTo EH_REMOVE
        prj.VBComponents.Remove prj.VBComponents(names(i))
        On Error GoTo 0
    Next i

    For i = 1 To d
        On Error GoTo EH_CLEAR_DOC
        private_ClearDocumentModuleCode prj.VBComponents(docNames(i))
        On Error GoTo 0
    Next i

    Exit Sub

EH_REMOVE:
    Err.Raise VBA.vbObjectError + 1008, "private_RemoveAllModulesAndClasses", _
              "Failed to remove component '" & names(i) & "': " & Err.Description

EH_CLEAR_DOC:
    Err.Raise VBA.vbObjectError + 1011, "private_RemoveAllModulesAndClasses", _
              "Failed to clear document module '" & docNames(i) & "': " & Err.Description
End Sub


Private Sub private_RemoveImportedModulesByScope( _
    Optional ByVal includeComponentPattern As String = VBA.vbNullString, _
    Optional ByVal excludeComponentPattern As String = VBA.vbNullString _
)
    Dim prj As Object
    Dim comp As Object
    Dim names() As String
    Dim n As Long
    Dim i As Long

    Set prj = ThisWorkbook.VBProject

    For Each comp In prj.VBComponents
        If comp.Type <> 100 Then ' модуль документа
            If private_ShouldProcessComponentByScope(VBA.CStr(comp.Name), includeComponentPattern, excludeComponentPattern) Then
                n = n + 1
                ReDim Preserve names(1 To n)
                names(n) = VBA.CStr(comp.Name)
            End If
        End If
    Next comp

    For i = 1 To n
        On Error GoTo EH_REMOVE
        prj.VBComponents.Remove prj.VBComponents(names(i))
        On Error GoTo 0
    Next i

    Exit Sub

EH_REMOVE:
    Err.Raise VBA.vbObjectError + 1004, "private_RemoveImportedModulesByScope", _
              "Failed to remove component '" & names(i) & "': " & Err.Description
End Sub


Private Sub private_ImportFolder( _
    ByVal folderPath As String, _
    ByVal updateMode As Long, _
    ByVal prevCache As Object, _
    ByVal nextCache As Object, _
    Optional ByVal includeComponentPattern As String = VBA.vbNullString, _
    Optional ByVal excludeComponentPattern As String = VBA.vbNullString _
)
    Dim fso As Object
    Dim rootFolder As Object
    Dim failed As String

    If Dir(folderPath, vbDirectory) = "" Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rootFolder = fso.GetFolder(folderPath)
    private_ImportFolderRecursive rootFolder, 0, failed, updateMode, prevCache, nextCache, includeComponentPattern, excludeComponentPattern

    If VBA.Len(failed) > 0 Then
        Err.Raise VBA.vbObjectError + 1001, "private_ImportFolder", "Import failed for file(s):" & failed
    End If
End Sub


Private Sub private_ImportFolderRecursive( _
    ByVal folderObj As Object, _
    ByVal depth As Long, _
    ByRef failed As String, _
    ByVal updateMode As Long, _
    ByVal prevCache As Object, _
    ByVal nextCache As Object, _
    Optional ByVal includeComponentPattern As String = VBA.vbNullString, _
    Optional ByVal excludeComponentPattern As String = VBA.vbNullString _
)
    Dim fileObj As Object
    Dim subFolder As Object
    Dim importPath As String
    Dim fileName As String
    Dim componentName As String
    Dim errText As String
    Dim sourceText As String
    Dim fileDateStamp As String
    Dim fileSizeStamp As String
    Dim cacheKey As String
    Dim compType As String
    Dim fallbackName As String
    Dim componentNameForCache As String
    Dim sourceEncodingIsUtf8 As Boolean
    Dim incrementalMode As Boolean
    Dim shouldProcess As Boolean

    If folderObj Is Nothing Then Exit Sub
    If depth > MAX_IMPORT_RECURSION_DEPTH Then Exit Sub

    incrementalMode = (updateMode <> UPDATE_MODE_FULL)

    For Each fileObj In folderObj.Files
        fileName = VBA.CStr(fileObj.Name)

        If private_TryResolveFileComponentType(fileName, compType, fallbackName) Then
            importPath = VBA.CStr(fileObj.Path)
            On Error GoTo EH_IMPORT_FILE

            fileDateStamp = private_BuildFileDateStampFromFileObject(fileObj)
            fileSizeStamp = private_BuildFileSizeStampFromFileObject(fileObj)
            cacheKey = private_NormalizeCacheKey(importPath)
            sourceText = VBA.vbNullString
            sourceEncodingIsUtf8 = private_HasUtf8MarkerBeforeVba(fileName)

            shouldProcess = private_ShouldProcessComponentByScope(fallbackName, includeComponentPattern, excludeComponentPattern)
            If Not shouldProcess Then GoTo ContinueNextFile

            Select Case VBA.LCase$(compType)
                Case COMP_TYPE_MODULE, COMP_TYPE_CLASS
                    If incrementalMode Then
                        If private_TryGetCachedComponentNameByMode(prevCache, cacheKey, compType, fileDateStamp, fileSizeStamp, updateMode, componentName) Then
                            If private_IsComponentPresentForType(componentName, compType) Then
                                private_SetCacheRecord nextCache, cacheKey, compType, componentName, fileDateStamp, fileSizeStamp
                                GoTo ContinueNextFile
                            End If
                        End If
                    End If

                    sourceText = private_ReadAllText(importPath, sourceEncodingIsUtf8)
                    componentName = private_GetComponentNameFromSourceText(sourceText, fallbackName)
                    private_EnsureValidComponentNameLength componentName, importPath

                    private_RemoveComponentIfExists componentName
                    If VBA.StrComp(compType, COMP_TYPE_MODULE, VBA.vbTextCompare) = 0 Then
                        private_ImportStandardModuleFromSource componentName, importPath, sourceText
                    Else
                        private_ImportClassModuleFromSource componentName, importPath, sourceText
                    End If
                    private_SetCacheRecord nextCache, cacheKey, compType, componentName, fileDateStamp, fileSizeStamp

                Case COMP_TYPE_SHEET
                    componentNameForCache = private_ResolveSheetCodeName(fallbackName)
                    If VBA.Len(componentNameForCache) = 0 Then GoTo ContinueNextFile

                    If incrementalMode Then
                        If private_IsCacheRecordCurrentByMode(prevCache, cacheKey, COMP_TYPE_SHEET, componentNameForCache, fileDateStamp, fileSizeStamp, updateMode) Then
                            If private_IsComponentPresentForType(componentNameForCache, COMP_TYPE_SHEET) Then
                                private_SetCacheRecord nextCache, cacheKey, COMP_TYPE_SHEET, componentNameForCache, fileDateStamp, fileSizeStamp
                                GoTo ContinueNextFile
                            End If
                        End If
                    End If

                    sourceText = private_ReadAllText(importPath, sourceEncodingIsUtf8)
                    If private_UpdateSheetModule(componentNameForCache, importPath, sourceText) Then
                        private_SetCacheRecord nextCache, cacheKey, COMP_TYPE_SHEET, componentNameForCache, fileDateStamp, fileSizeStamp
                    End If

                Case COMP_TYPE_WORKBOOK
                    componentNameForCache = private_FindWorkbookComponentName()
                    If VBA.Len(componentNameForCache) = 0 Then GoTo ContinueNextFile

                    If incrementalMode Then
                        If private_IsCacheRecordCurrentByMode(prevCache, cacheKey, COMP_TYPE_WORKBOOK, componentNameForCache, fileDateStamp, fileSizeStamp, updateMode) Then
                            If private_IsComponentPresentForType(componentNameForCache, COMP_TYPE_WORKBOOK) Then
                                private_SetCacheRecord nextCache, cacheKey, COMP_TYPE_WORKBOOK, componentNameForCache, fileDateStamp, fileSizeStamp
                                GoTo ContinueNextFile
                            End If
                        End If
                    End If

                    sourceText = private_ReadAllText(importPath, sourceEncodingIsUtf8)
                    If private_UpdateWorkbookModuleFromText(componentNameForCache, sourceText) Then
                        private_SetCacheRecord nextCache, cacheKey, COMP_TYPE_WORKBOOK, componentNameForCache, fileDateStamp, fileSizeStamp
                    End If
            End Select
            On Error GoTo 0
        End If

ContinueNextFile:
    Next fileObj

    For Each subFolder In folderObj.SubFolders
        private_ImportFolderRecursive subFolder, depth + 1, failed, updateMode, prevCache, nextCache, includeComponentPattern, excludeComponentPattern
    Next subFolder

    Exit Sub

EH_IMPORT_FILE:
    errText = VBA.CStr(Err.Number) & ": " & Err.Description
    failed = failed & VBA.vbCrLf & "- " & importPath & " (" & errText & ")"
    Err.Clear
    On Error GoTo 0
    GoTo ContinueNextFile
End Sub


Private Sub private_EnsureValidComponentNameLength(ByVal componentName As String, ByVal importPath As String)
    If VBA.Len(componentName) <= MAX_VBA_COMPONENT_NAME_LEN Then Exit Sub
    Err.Raise VBA.vbObjectError + 1010, "private_EnsureValidComponentNameLength", _
              "VBA component name '" & componentName & "' is too long (" & VBA.CStr(VBA.Len(componentName)) & _
              "). Maximum allowed is " & VBA.CStr(MAX_VBA_COMPONENT_NAME_LEN) & ". File: " & importPath
End Sub


Private Sub private_ImportStandardModuleFromSource( _
    ByVal componentName As String, _
    ByVal importPath As String, _
    Optional ByVal preloadedSourceText As String = VBA.vbNullString _
)
    Dim vbComp As Object
    Dim cm As Object
    Dim sourceText As String
    Dim cleanCode As String

    If VBA.Len(VBA.Trim$(componentName)) = 0 Then
        Err.Raise VBA.vbObjectError + 1009, "private_ImportStandardModuleFromSource", "Standard module name is empty for: " & importPath
    End If

    If VBA.Len(preloadedSourceText) > 0 Then
        sourceText = preloadedSourceText
    Else
        sourceText = private_ReadAllText(importPath)
    End If
    cleanCode = private_ExtractCodeBody(sourceText)

    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(1) ' стандартный модуль (vbext_ct_StdModule)
    vbComp.Name = componentName
    Set cm = vbComp.CodeModule
    cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString cleanCode
End Sub


Private Sub private_ImportClassModuleFromSource( _
    ByVal componentName As String, _
    ByVal importPath As String, _
    Optional ByVal preloadedSourceText As String = VBA.vbNullString _
)
    Dim vbComp As Object
    Dim cm As Object
    Dim sourceText As String
    Dim cleanCode As String

    If VBA.Len(VBA.Trim$(componentName)) = 0 Then
        Err.Raise VBA.vbObjectError + 1005, "private_ImportClassModuleFromSource", "Class module name is empty for: " & importPath
    End If

    If VBA.Len(preloadedSourceText) > 0 Then
        sourceText = preloadedSourceText
    Else
        sourceText = private_ReadAllText(importPath)
    End If
    cleanCode = private_ExtractCodeBody(sourceText)

    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(2) ' модуль класса (vbext_ct_ClassModule)
    vbComp.Name = componentName
    Set cm = vbComp.CodeModule
    cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString cleanCode
End Sub


Private Function private_ExtractCodeBody(ByVal sourceText As String) As String
    Dim lines As Variant
    Dim i As Long
    Dim lineText As String
    Dim trimmed As String
    Dim outText As String

    sourceText = VBA.Replace(sourceText, VBA.vbCrLf, VBA.vbLf)
    sourceText = VBA.Replace(sourceText, VBA.vbCr, VBA.vbLf)
    lines = VBA.Split(sourceText, VBA.vbLf)

    For i = LBound(lines) To UBound(lines)
        lineText = VBA.CStr(lines(i))
        ' Удаляем служебный BOM/непечатаемый префикс, если он присутствует.
        lineText = VBA.Replace(lineText, ChrW$(65279), VBA.vbNullString)
        lineText = VBA.Replace(lineText, ChrW$(160), " ")
        trimmed = VBA.Trim$(lineText)

        If VBA.StrComp(VBA.Left$(trimmed, 8), "VERSION ", VBA.vbTextCompare) = 0 Then GoTo ContinueLine
        If VBA.StrComp(trimmed, "BEGIN", VBA.vbTextCompare) = 0 Then GoTo ContinueLine
        If VBA.StrComp(trimmed, "END", VBA.vbTextCompare) = 0 Then GoTo ContinueLine
        If VBA.StrComp(VBA.Left$(trimmed, 10), "Attribute ", VBA.vbTextCompare) = 0 Then GoTo ContinueLine
        ' Строки метаданных класса из заголовка экспортированного .cls не являются корректными инструкциями VBA.
        If VBA.StrComp(VBA.Left$(trimmed, 10), "MultiUse =", VBA.vbTextCompare) = 0 Then GoTo ContinueLine
        If VBA.StrComp(VBA.Left$(trimmed, 13), "Persistable =", VBA.vbTextCompare) = 0 Then GoTo ContinueLine
        If VBA.StrComp(VBA.Left$(trimmed, 20), "DataBindingBehavior =", VBA.vbTextCompare) = 0 Then GoTo ContinueLine
        If VBA.StrComp(VBA.Left$(trimmed, 19), "DataSourceBehavior =", VBA.vbTextCompare) = 0 Then GoTo ContinueLine
        If VBA.StrComp(VBA.Left$(trimmed, 21), "MTSTransactionMode =", VBA.vbTextCompare) = 0 Then GoTo ContinueLine

        If VBA.Len(outText) > 0 Then outText = outText & VBA.vbCrLf
        outText = outText & lineText

ContinueLine:
    Next i

    private_ExtractCodeBody = outText
End Function


Private Sub private_RemoveComponentIfExists(ByVal componentName As String)
    Dim vbComp As Object

    If VBA.Len(componentName) = 0 Then Exit Sub

    Set vbComp = private_TryGetComponentByName(componentName)
    If vbComp Is Nothing Then Exit Sub

    ThisWorkbook.VBProject.VBComponents.Remove vbComp
End Sub


Private Function private_TryGetComponentByName(ByVal componentName As String) As Object
    On Error Resume Next
    Set private_TryGetComponentByName = ThisWorkbook.VBProject.VBComponents(componentName)
    On Error GoTo 0
End Function


Private Function private_IsComponentPresentForType(ByVal componentName As String, ByVal compType As String) As Boolean
    Dim vbComp As Object

    Set vbComp = private_TryGetComponentByName(componentName)
    If vbComp Is Nothing Then Exit Function

    Select Case VBA.LCase$(compType)
        Case COMP_TYPE_MODULE
            private_IsComponentPresentForType = (vbComp.Type = 1) ' стандартный модуль
        Case COMP_TYPE_CLASS
            private_IsComponentPresentForType = (vbComp.Type = 2) ' модуль класса
        Case COMP_TYPE_SHEET, COMP_TYPE_WORKBOOK
            private_IsComponentPresentForType = (vbComp.Type = 100) ' модуль документа
    End Select
End Function


Private Function private_GetComponentNameFromSource(ByVal importPath As String) As String
    Dim fileName As String
    Dim dotPos As Long
    Dim compType As String
    Dim sourceText As String
    Dim fallbackName As String

    fileName = VBA.Mid$(importPath, VBA.InStrRev(importPath, "\") + 1)
    If Not private_TryResolveFileComponentType(fileName, compType, fallbackName) Then
        dotPos = VBA.InStrRev(fileName, ".")
        If dotPos > 1 Then
            fallbackName = VBA.Left$(fileName, dotPos - 1)
        Else
            fallbackName = fileName
        End If
    End If

    sourceText = private_ReadAllText(importPath)
    private_GetComponentNameFromSource = private_GetComponentNameFromSourceText(sourceText, fallbackName)
End Function


Private Function private_GetComponentNameFromSourceText(ByVal sourceText As String, ByVal fallbackName As String) As String
    Dim attrPos As Long
    Dim quoteStart As Long
    Dim quoteEnd As Long

    private_GetComponentNameFromSourceText = fallbackName

    attrPos = VBA.InStr(1, sourceText, "Attribute VB_Name", VBA.vbTextCompare)
    If attrPos = 0 Then Exit Function

    quoteStart = VBA.InStr(attrPos, sourceText, """")
    If quoteStart = 0 Then Exit Function

    quoteEnd = VBA.InStr(quoteStart + 1, sourceText, """")
    If quoteEnd <= quoteStart Then Exit Function

    private_GetComponentNameFromSourceText = VBA.Mid$(sourceText, quoteStart + 1, quoteEnd - quoteStart - 1)
End Function


Private Function private_EndsWith(ByVal value As String, ByVal suffix As String) As Boolean
    private_EndsWith = (VBA.LCase$(VBA.Right$(value, VBA.Len(suffix))) = VBA.LCase$(suffix))
End Function


Private Function private_ShouldProcessComponentByScope( _
    ByVal componentName As String, _
    Optional ByVal includeComponentPattern As String = VBA.vbNullString, _
    Optional ByVal excludeComponentPattern As String = VBA.vbNullString _
) As Boolean
    private_ShouldProcessComponentByScope = private_MatchesIncludeExclude(componentName, includeComponentPattern, excludeComponentPattern)
End Function


Private Function private_MatchesIncludeExclude( _
    ByVal valueText As String, _
    ByVal includePattern As String, _
    ByVal excludePattern As String _
) As Boolean
    valueText = VBA.Trim$(VBA.CStr(valueText))
    includePattern = VBA.Trim$(VBA.CStr(includePattern))
    excludePattern = VBA.Trim$(VBA.CStr(excludePattern))

    If VBA.Len(valueText) = 0 Then
        private_MatchesIncludeExclude = (VBA.Len(includePattern) = 0)
        Exit Function
    End If

    If VBA.Len(includePattern) > 0 Then
        If Not private_IsRegexMatch(valueText, includePattern) Then Exit Function
    End If

    If VBA.Len(excludePattern) > 0 Then
        If private_IsRegexMatch(valueText, excludePattern) Then Exit Function
    End If

    private_MatchesIncludeExclude = True
End Function


Private Function private_IsRegexMatch(ByVal valueText As String, ByVal regexPattern As String) As Boolean
    Dim re As Object

    regexPattern = VBA.Trim$(VBA.CStr(regexPattern))
    If VBA.Len(regexPattern) = 0 Then Exit Function

    On Error GoTo EH

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = regexPattern

    private_IsRegexMatch = re.Test(VBA.CStr(valueText))
    Exit Function

EH:
    Err.Raise VBA.vbObjectError + 1013, "private_IsRegexMatch", "Некорректный regex '" & regexPattern & "': " & Err.Description
End Function


Private Function private_TryResolveFileComponentType( _
    ByVal fileName As String, _
    ByRef outCompType As String, _
    ByRef outFallbackName As String _
) As Boolean
    Dim normalizedName As String
    Dim baseName As String

    normalizedName = VBA.LCase$(VBA.Trim$(fileName))
    outCompType = VBA.vbNullString
    outFallbackName = VBA.vbNullString

    If private_EndsWith(normalizedName, ".utf8.vba") Then
        baseName = VBA.Left$(fileName, VBA.Len(fileName) - VBA.Len(".utf8.vba"))
        normalizedName = VBA.LCase$(VBA.Trim$(baseName))
    ElseIf private_EndsWith(normalizedName, ".vba") Then
        baseName = VBA.Left$(fileName, VBA.Len(fileName) - VBA.Len(".vba"))
        normalizedName = VBA.LCase$(VBA.Trim$(baseName))
    Else
        Exit Function
    End If

    If VBA.StrComp(normalizedName, "thisworkbook", VBA.vbTextCompare) = 0 Then
        outCompType = COMP_TYPE_WORKBOOK
        outFallbackName = "ThisWorkbook"
    ElseIf VBA.Left$(normalizedName, 3) = "ws_" Then
        outCompType = COMP_TYPE_SHEET
        outFallbackName = VBA.Mid$(baseName, 4)
    ElseIf VBA.Left$(normalizedName, 3) = "ex_" Or VBA.Left$(normalizedName, 3) = "rt_" Then
        outCompType = COMP_TYPE_MODULE
        outFallbackName = baseName
    ElseIf VBA.Left$(normalizedName, 4) = "obj_" And private_EndsWith(normalizedName, ".cls") Then
        outCompType = COMP_TYPE_CLASS
        outFallbackName = VBA.Left$(baseName, VBA.Len(baseName) - VBA.Len(".cls"))
    End If

    private_TryResolveFileComponentType = (VBA.Len(VBA.Trim$(outCompType)) > 0 And VBA.Len(VBA.Trim$(outFallbackName)) > 0)
End Function


Private Function private_HasUtf8MarkerBeforeVba(ByVal fileName As String) As Boolean
    private_HasUtf8MarkerBeforeVba = private_EndsWith(VBA.LCase$(VBA.Trim$(fileName)), ".utf8.vba")
End Function


Private Function private_CreateDictionary() As Object
    Set private_CreateDictionary = CreateObject("Scripting.Dictionary")
    private_CreateDictionary.CompareMode = 1
End Function


Private Function private_NormalizeCacheKey(ByVal filePath As String) As String
    private_NormalizeCacheKey = VBA.LCase$(VBA.Replace$(VBA.CStr(filePath), "/", "\"))
End Function


Private Function private_BuildFileStamp(ByVal filePath As String) As String
    Dim fso As Object
    Dim fileObj As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Exit Function
    Set fileObj = fso.GetFile(filePath)
    private_BuildFileStamp = private_BuildFileDateStampFromFileObject(fileObj) & ":" & private_BuildFileSizeStampFromFileObject(fileObj)
End Function


Private Function private_BuildFileDateStampFromFileObject(ByVal fileObj As Object) As String
    private_BuildFileDateStampFromFileObject = VBA.CStr(VBA.CDbl(fileObj.DateLastModified))
End Function


Private Function private_BuildFileSizeStampFromFileObject(ByVal fileObj As Object) As String
    private_BuildFileSizeStampFromFileObject = VBA.CStr(VBA.CLng(fileObj.Size))
End Function


Private Function private_IsCacheRecordCurrentByMode( _
    ByVal cache As Object, _
    ByVal cacheKey As String, _
    ByVal compType As String, _
    ByVal componentName As String, _
    ByVal fileDateStamp As String, _
    ByVal fileSizeStamp As String, _
    ByVal updateMode As Long _
) As Boolean
    Dim rec As Object

    If cache Is Nothing Then Exit Function
    If Not cache.Exists(cacheKey) Then Exit Function
    Set rec = cache(cacheKey)
    If rec Is Nothing Then Exit Function

    If VBA.StrComp(VBA.CStr(rec("Type")), compType, VBA.vbTextCompare) <> 0 Then Exit Function
    If VBA.StrComp(VBA.CStr(rec("Name")), componentName, VBA.vbTextCompare) <> 0 Then Exit Function
    If Not private_IsCacheRecordMatchByMode(rec, fileDateStamp, fileSizeStamp, updateMode) Then Exit Function

    private_IsCacheRecordCurrentByMode = True
End Function


Private Function private_TryGetCachedComponentNameByMode( _
    ByVal cache As Object, _
    ByVal cacheKey As String, _
    ByVal compType As String, _
    ByVal fileDateStamp As String, _
    ByVal fileSizeStamp As String, _
    ByVal updateMode As Long, _
    ByRef outComponentName As String _
) As Boolean
    Dim rec As Object

    If cache Is Nothing Then Exit Function
    If Not cache.Exists(cacheKey) Then Exit Function

    Set rec = cache(cacheKey)
    If rec Is Nothing Then Exit Function
    If VBA.StrComp(VBA.CStr(rec("Type")), compType, VBA.vbTextCompare) <> 0 Then Exit Function
    If Not private_IsCacheRecordMatchByMode(rec, fileDateStamp, fileSizeStamp, updateMode) Then Exit Function

    outComponentName = VBA.CStr(rec("Name"))
    private_TryGetCachedComponentNameByMode = (VBA.Len(outComponentName) > 0)
End Function


Private Function private_IsCacheRecordMatchByMode( _
    ByVal rec As Object, _
    ByVal fileDateStamp As String, _
    ByVal fileSizeStamp As String, _
    ByVal updateMode As Long _
) As Boolean
    Dim cachedDateStamp As String
    Dim cachedSizeStamp As String

    If rec Is Nothing Then Exit Function

    cachedDateStamp = VBA.vbNullString
    cachedSizeStamp = VBA.vbNullString

    If rec.Exists("DateStamp") Then
        cachedDateStamp = VBA.CStr(rec("DateStamp"))
    End If
    If rec.Exists("SizeStamp") Then
        cachedSizeStamp = VBA.CStr(rec("SizeStamp"))
    End If

    If VBA.Len(cachedDateStamp) = 0 And rec.Exists("Stamp") Then
        cachedDateStamp = VBA.CStr(rec("Stamp"))
    End If

    Select Case updateMode
        Case UPDATE_MODE_DATE
            If VBA.StrComp(cachedDateStamp, fileDateStamp, VBA.vbBinaryCompare) <> 0 Then Exit Function

        Case UPDATE_MODE_SIZE
            If VBA.StrComp(cachedSizeStamp, fileSizeStamp, VBA.vbBinaryCompare) <> 0 Then Exit Function

        Case Else
            If VBA.StrComp(cachedDateStamp, fileDateStamp, VBA.vbBinaryCompare) <> 0 Then Exit Function
            If VBA.StrComp(cachedSizeStamp, fileSizeStamp, VBA.vbBinaryCompare) <> 0 Then Exit Function
    End Select

    private_IsCacheRecordMatchByMode = True
End Function


Private Sub private_SetCacheRecord( _
    ByVal cache As Object, _
    ByVal cacheKey As String, _
    ByVal compType As String, _
    ByVal componentName As String, _
    ByVal fileDateStamp As String, _
    ByVal fileSizeStamp As String _
)
    Dim rec As Object

    If cache Is Nothing Then Exit Sub

    Set rec = private_CreateDictionary()
    rec("Type") = compType
    rec("Name") = componentName
    rec("DateStamp") = fileDateStamp
    rec("SizeStamp") = fileSizeStamp

    If cache.Exists(cacheKey) Then
        cache.Remove cacheKey
    End If
    cache.Add cacheKey, rec
End Sub


Private Function private_LoadImportCache(ByVal cachePath As String) As Object
    Dim cache As Object
    Dim lineText As String
    Dim parts() As String
    Dim f As Integer
    Dim legacyStamp As String
    Dim sepPos As Long
    Dim fileDateStamp As String
    Dim fileSizeStamp As String

    Set cache = private_CreateDictionary()
    If VBA.Len(Dir(cachePath)) = 0 Then
        Set private_LoadImportCache = cache
        Exit Function
    End If

    f = FreeFile
    Open cachePath For Input As #f
    Do While Not EOF(f)
        Line Input #f, lineText
        If VBA.Len(VBA.Trim$(lineText)) = 0 Then GoTo ContinueLoop
        parts = VBA.Split(lineText, "|")
        If UBound(parts) < 3 Then GoTo ContinueLoop

        fileDateStamp = VBA.vbNullString
        fileSizeStamp = VBA.vbNullString

        If UBound(parts) >= 3 Then
            fileDateStamp = VBA.CStr(parts(3))
        End If
        If UBound(parts) >= 4 Then
            fileSizeStamp = VBA.CStr(parts(4))
        ElseIf VBA.Len(fileDateStamp) > 0 Then
            ' Старый формат: дата и размер в одном поле через двоеточие.
            legacyStamp = fileDateStamp
            sepPos = VBA.InStrRev(legacyStamp, ":")
            If sepPos > 0 Then
                fileDateStamp = VBA.Left$(legacyStamp, sepPos - 1)
                fileSizeStamp = VBA.Mid$(legacyStamp, sepPos + 1)
            End If
        End If

        private_SetCacheRecord cache, VBA.CStr(parts(0)), VBA.CStr(parts(1)), VBA.CStr(parts(2)), fileDateStamp, fileSizeStamp
ContinueLoop:
    Loop
    Close #f

    Set private_LoadImportCache = cache
End Function


Private Sub private_SaveImportCache(ByVal cachePath As String, ByVal cache As Object)
    Dim f As Integer
    Dim key As Variant
    Dim rec As Object

    If cache Is Nothing Then Exit Sub

    f = FreeFile
    Open cachePath For Output As #f
    For Each key In cache.Keys
        Set rec = cache(VBA.CStr(key))
        Print #f, VBA.CStr(key) & "|" & VBA.CStr(rec("Type")) & "|" & VBA.CStr(rec("Name")) & "|" & VBA.CStr(rec("DateStamp")) & "|" & VBA.CStr(rec("SizeStamp"))
    Next key
    Close #f
End Sub


Private Sub private_RemoveStaleImportedComponentsByScope( _
    ByVal prevCache As Object, _
    ByVal nextCache As Object, _
    Optional ByVal includeComponentPattern As String = VBA.vbNullString, _
    Optional ByVal excludeComponentPattern As String = VBA.vbNullString _
)
    Dim key As Variant
    Dim rec As Object
    Dim compType As String
    Dim componentName As String

    If prevCache Is Nothing Then Exit Sub
    If nextCache Is Nothing Then Exit Sub

    For Each key In prevCache.Keys
        If Not nextCache.Exists(VBA.CStr(key)) Then
            Set rec = prevCache(VBA.CStr(key))
            If Not rec Is Nothing Then
                compType = VBA.CStr(rec("Type"))
                componentName = VBA.CStr(rec("Name"))
                If Not private_ShouldProcessComponentByScope(componentName, includeComponentPattern, excludeComponentPattern) Then GoTo ContinueKey
                    If VBA.StrComp(compType, COMP_TYPE_MODULE, VBA.vbTextCompare) = 0 Or _
                       VBA.StrComp(compType, COMP_TYPE_CLASS, VBA.vbTextCompare) = 0 Then
                        private_RemoveComponentIfExists componentName
                    End If
            End If
        End If
ContinueKey:
    Next key
End Sub


Private Function private_GetNextOnTimeTick() As Date
    Dim nowValue As Date

    ' Excel.OnTime планирует с точностью до секунды, поэтому даем +1 сек от текущего времени.
    nowValue = VBA.Now
    private_GetNextOnTimeTick = VBA.DateSerial(VBA.Year(nowValue), VBA.Month(nowValue), VBA.Day(nowValue)) + _
                                VBA.TimeSerial(VBA.Hour(nowValue), VBA.Minute(nowValue), VBA.Second(nowValue) + 1)
End Function


' Callstack[1]: ex_Core.private_TryRunSafeUpdateByMode -> private_TryRunRuntimeBooleanFunction
Private Function private_TryRunRuntimeBooleanFunction( _
    ByVal moduleName As String, _
    ByVal functionName As String, _
    ByRef outResult As Boolean _
) As Boolean
    Dim callResult As Variant
    Dim errDescription As String

    outResult = False
    If Not private_TryRunRuntimeNoArgMember(moduleName, functionName, callResult) Then Exit Function

    On Error Resume Next
    outResult = VBA.CBool(callResult)
    If Err.Number <> 0 Then
        errDescription = Err.Description
        Err.Clear
        On Error GoTo 0
        private_LogCoreSelfEvent "runtime-call-failed: module='" & VBA.Replace$(moduleName, "'", "''") & "' function='" & VBA.Replace$(functionName, "'", "''") & "' err='bool-cast-failed: " & VBA.Replace$(errDescription, "'", "''") & "'"
        Exit Function
    End If
    On Error GoTo 0

    private_TryRunRuntimeBooleanFunction = True
End Function


' Callstack[1]: ex_Core.private_TryQueueRuntimeUpdateWhenBridgeDispatch -> private_TryRunRuntimeNoArgMember
' Callstack[2]: ex_Core.private_TryRunRuntimeBooleanFunction -> private_TryRunRuntimeNoArgMember
Private Function private_TryRunRuntimeNoArgMember( _
    ByVal moduleName As String, _
    ByVal memberName As String, _
    ByRef outResult As Variant, _
    Optional ByVal suppressFailureLog As Boolean = False _
) As Boolean
    Dim macroRef As String
    Dim unqualifiedMacroRef As String
    Dim errDescriptionQualified As String
    Dim errDescriptionUnqualified As String
    Dim runtimeComponent As Object

    outResult = Empty
    moduleName = VBA.Trim$(moduleName)
    memberName = VBA.Trim$(memberName)
    If VBA.Len(moduleName) = 0 Then Exit Function
    If VBA.Len(memberName) = 0 Then Exit Function

    Set runtimeComponent = private_TryGetComponentByName(moduleName)
    If runtimeComponent Is Nothing Then
        If Not suppressFailureLog Then
            private_LogCoreSelfEvent "runtime-call-failed: module='" & VBA.Replace$(moduleName, "'", "''") & "' member='" & VBA.Replace$(memberName, "'", "''") & "' err='component is missing'"
        End If
        Exit Function
    End If

    macroRef = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!" & moduleName & "." & memberName
    unqualifiedMacroRef = moduleName & "." & memberName

    On Error Resume Next
    outResult = Application.Run(macroRef)
    If Err.Number = 0 Then
        private_TryRunRuntimeNoArgMember = True
        On Error GoTo 0
        Exit Function
    End If
    errDescriptionQualified = Err.Description
    Err.Clear

    outResult = Application.Run(unqualifiedMacroRef)
    If Err.Number = 0 Then
        private_TryRunRuntimeNoArgMember = True
        On Error GoTo 0
        Exit Function
    End If

    errDescriptionUnqualified = Err.Description
    Err.Clear
    On Error GoTo 0

    If Not suppressFailureLog Then
        private_LogCoreSelfEvent "runtime-call-failed: module='" & VBA.Replace$(moduleName, "'", "''") & "' member='" & VBA.Replace$(memberName, "'", "''") & "' qualifiedErr='" & VBA.Replace$(errDescriptionQualified, "'", "''") & "' unqualifiedErr='" & VBA.Replace$(errDescriptionUnqualified, "'", "''") & "'"
    End If
End Function


'==========================
' Обновление модулей листов
'==========================
Private Function private_UpdateSheetModule( _
    ByVal sheetName As String, _
    ByVal sheetCodePath As String, _
    Optional ByVal preloadedCodeText As String = VBA.vbNullString _
) As Boolean
    Dim vbProj As Object
    Dim vbComp As Object
    Dim cm As Object
    Dim codeText As String

    Set vbProj = ThisWorkbook.VBProject
    If Not private_SheetModuleExists(vbProj, sheetName) Then Exit Function

    If VBA.Len(preloadedCodeText) > 0 Then
        codeText = preloadedCodeText
    Else
        If VBA.Len(private_BuildFileStamp(sheetCodePath)) = 0 Then Exit Function
        codeText = private_ReadAllText(sheetCodePath)
    End If

    Set vbComp = vbProj.VBComponents(sheetName)
    Set cm = vbComp.CodeModule

    cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString codeText
    private_UpdateSheetModule = True
End Function


Private Function private_ResolveSheetCodeName(ByVal fileStem As String) As String
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(fileStem)
    On Error GoTo 0

    If Not ws Is Nothing Then
        private_ResolveSheetCodeName = ws.CodeName
    Else
        private_ResolveSheetCodeName = fileStem
    End If
End Function


Private Function private_SheetModuleExists(ByVal vbProj As Object, ByVal sheetName As String) As Boolean
    Dim vbComp As Object
    On Error Resume Next
    Set vbComp = vbProj.VBComponents(sheetName)
    private_SheetModuleExists = Not vbComp Is Nothing
    On Error GoTo 0
End Function


Private Function private_FindWorkbookComponentName() As String
    Dim vbProj As Object
    Dim vbComp As Object
    Dim nameCandidates(1 To 4) As String
    Dim i As Long

    Set vbProj = ThisWorkbook.VBProject
    nameCandidates(1) = "wb_Host"
    nameCandidates(2) = "ThisWorkbook"
    nameCandidates(3) = "ЭтаКнига"
    nameCandidates(4) = "ЦяКнига"

    For i = LBound(nameCandidates) To UBound(nameCandidates)
        On Error Resume Next
        Set vbComp = vbProj.VBComponents(nameCandidates(i))
        On Error GoTo 0
        If Not vbComp Is Nothing Then
            private_FindWorkbookComponentName = nameCandidates(i)
            Exit Function
        End If
    Next i
End Function


Private Function private_UpdateWorkbookModuleFromText( _
    ByVal workbookComponentName As String, _
    ByVal codeText As String _
) As Boolean
    Dim vbProj As Object
    Dim vbComp As Object
    Dim cm As Object

    If VBA.Len(VBA.Trim$(workbookComponentName)) = 0 Then Exit Function
    If VBA.Len(codeText) = 0 Then Exit Function

    Set vbProj = ThisWorkbook.VBProject
    On Error Resume Next
    Set vbComp = vbProj.VBComponents(workbookComponentName)
    On Error GoTo 0
    If vbComp Is Nothing Then Exit Function

    Set cm = vbComp.CodeModule
    cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString codeText

    private_UpdateWorkbookModuleFromText = True
End Function


Private Function private_ReadAllText( _
    ByVal filePath As String, _
    Optional ByVal preferUtf8 As Boolean = False _
) As String
    If preferUtf8 Then
        On Error GoTo FallbackLegacy
        private_ReadAllText = private_ReadAllTextByCharset(filePath, "utf-8")
        If VBA.Left$(private_ReadAllText, 1) = ChrW$(65279) Then
            private_ReadAllText = VBA.Mid$(private_ReadAllText, 2)
        End If
        Exit Function
    End If

    On Error GoTo FallbackLegacy
    private_ReadAllText = private_ReadAllTextByCharset(filePath, "utf-8")
    If VBA.Left$(private_ReadAllText, 1) = ChrW$(65279) Then
        private_ReadAllText = VBA.Mid$(private_ReadAllText, 2)
    End If
    Exit Function

FallbackLegacy:
    Err.Clear
    private_ReadAllText = private_ReadAllTextLegacy(filePath)
End Function


Private Function private_ReadAllTextByCharset(ByVal filePath As String, ByVal charsetName As String) As String
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' текстовый поток
    stream.Mode = 3 ' режим чтение/запись
    stream.Charset = charsetName
    stream.Open
    stream.LoadFromFile filePath
    private_ReadAllTextByCharset = stream.ReadText(-1)
    stream.Close
End Function


Private Function private_ReadAllTextLegacy(ByVal filePath As String) As String
    Dim f As Integer
    Dim text As String

    f = FreeFile
    Open filePath For Input As #f
    text = Input$(LOF(f), f)
    Close #f

    private_ReadAllTextLegacy = text
End Function


Private Sub private_ShowStatusNotice(ByVal messageText As String, ByVal useNativeStatus As Boolean, Optional ByVal timeoutSeconds As Long = 3)
    If private_UseNativeStatus(useNativeStatus) Then
        private_ShowNativeStatus messageText
    Else
        If Not private_TryShowRtStatus("m_ShowStatusBarNotice", messageText, timeoutSeconds) Then
            private_ShowNativeStatus messageText
        End If
    End If
End Sub


Private Sub private_ShowStatusSuccess(ByVal messageText As String, ByVal useNativeStatus As Boolean, Optional ByVal timeoutSeconds As Long = 3)
    If private_UseNativeStatus(useNativeStatus) Then
        private_ShowNativeStatus messageText
    Else
        If Not private_TryShowRtStatus("m_ShowStatusBarSuccess", messageText, timeoutSeconds) Then
            private_ShowNativeStatus messageText
        End If
    End If
End Sub


Private Sub private_ShowStatusWarning(ByVal messageText As String, ByVal useNativeStatus As Boolean, Optional ByVal timeoutSeconds As Long = 3)
    If private_UseNativeStatus(useNativeStatus) Then
        private_ShowNativeStatus "Внимание: " & messageText
    Else
        If Not private_TryShowRtStatus("m_ShowStatusBarWarning", messageText, timeoutSeconds) Then
            private_ShowNativeStatus "Внимание: " & messageText
        End If
    End If
End Sub


Private Sub private_ShowStatusError(ByVal messageText As String, ByVal useNativeStatus As Boolean, Optional ByVal timeoutSeconds As Long = 3)
    If private_UseNativeStatus(useNativeStatus) Then
        private_ShowNativeStatus "Ошибка: " & messageText
    Else
        If Not private_TryShowRtStatus("m_ShowStatusBarError", messageText, timeoutSeconds) Then
            private_ShowNativeStatus "Ошибка: " & messageText
        End If
    End If
End Sub


Private Function private_UseNativeStatus(ByVal useNativeStatus As Boolean) As Boolean
#If CORE_FORCE_NATIVE_STATUS_BAR Then
    private_UseNativeStatus = True
#Else
    private_UseNativeStatus = useNativeStatus
#End If
End Function


Private Function private_TryShowRtStatus(ByVal methodName As String, ByVal messageText As String, ByVal timeoutSeconds As Long) As Boolean
    Dim macroRef As String
    Dim errDescription As String

    methodName = VBA.Trim$(methodName)
    If VBA.Len(methodName) = 0 Then Exit Function

    macroRef = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!rt_Messaging." & methodName

    On Error Resume Next
    Application.Run macroRef, VBA.CStr(messageText), VBA.CLng(timeoutSeconds)
    If Err.Number = 0 Then
        private_TryShowRtStatus = True
    Else
        errDescription = Err.Description
        Err.Clear
        private_LogCoreSelfEvent "rt-messaging-call-failed: " & methodName & ": " & errDescription
    End If
    On Error GoTo 0
End Function


Private Sub private_ShowNativeStatus(ByVal messageText As String)
    Dim statusText As String
    Dim errDescription As String

    messageText = VBA.Trim$(VBA.CStr(messageText))
    If VBA.Len(messageText) = 0 Then Exit Sub

    statusText = "PrototypeNew: " & messageText

    On Error Resume Next
    Application.StatusBar = statusText
    If Err.Number <> 0 Then
        errDescription = Err.Description
        Err.Clear
        On Error GoTo 0
        private_LogCoreSelfEvent "native-status-failed: message='" & VBA.Replace$(messageText, "'", "''") & "' err='" & VBA.Replace$(errDescription, "'", "''") & "'"
        Exit Sub
    End If
    On Error GoTo 0

    private_LogStatusBarEvent "native-show", messageText
End Sub


' --- Логирование в файл (внутренняя логика)
Private Sub private_LogStatusBarEvent( _
    ByVal actionName As String, _
    ByVal messageText As String, _
    Optional ByVal timeoutSeconds As Long = 0 _
)
#If Not CORE_ENABLE_STATUS_BAR_LOGGING Then
    Exit Sub
#Else
    Dim logLine As String

    actionName = VBA.Trim$(VBA.CStr(actionName))
    If VBA.Len(actionName) = 0 Then actionName = "event"

    logLine = "status-bar-" & actionName
    If timeoutSeconds > 0 Then logLine = logLine & ": timeout=" & VBA.CStr(timeoutSeconds)

    messageText = VBA.Trim$(VBA.CStr(messageText))
    If VBA.Len(messageText) > 0 Then
        logLine = logLine & " message='" & VBA.Replace$(messageText, "'", "''") & "'"
    End If
    
    private_LogCoreSelfEvent logLine
#End If
End Sub


Private Sub private_LogCoreSelfEvent(ByVal messageText As String)
#If Not CORE_ENABLE_SELF_LOGGING Then
    Exit Sub
#Else
    private_LogCoreEvent messageText
#End If
End Sub


Private Sub private_LogCoreEvent(ByVal messageText As String)
#If Not CORE_ENABLE_LOGGING Then
    Exit Sub
#Else
    Dim logPath As String
    Dim folderPath As String
    Dim fso As Object
    Dim stream As Object
    Dim lineText As String

    messageText = VBA.Trim$(VBA.CStr(messageText))
    If VBA.Len(messageText) = 0 Then Exit Sub
    If VBA.Len(VBA.Trim$(ThisWorkbook.Path)) = 0 Then Exit Sub

    logPath = ThisWorkbook.Path & "\\" & CORE_LOG_FILE_REL_PATH
    folderPath = VBA.Left$(logPath, VBA.InStrRev(logPath, "\\") - 1)

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        If VBA.Len(folderPath) > 0 Then
            If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
        End If
    End If

    lineText = VBA.Format$(VBA.Now, "yyyy-mm-dd hh:nn:ss") & " | " & messageText
    Set stream = fso.OpenTextFile(logPath, 8, True) ' режим добавления в конец файла (ForAppending)
    If Not stream Is Nothing Then
        stream.WriteLine lineText
        stream.Close
    End If
    Err.Clear
    On Error GoTo 0
#End If
End Sub


Private Sub private_ClearCoreLogFile()
#If Not CORE_ENABLE_LOGGING Then
    Exit Sub
#Else
    Dim logPath As String
    Dim folderPath As String
    Dim fso As Object
    Dim stream As Object

    If VBA.Len(VBA.Trim$(ThisWorkbook.Path)) = 0 Then Exit Sub

    logPath = ThisWorkbook.Path & "\\" & CORE_LOG_FILE_REL_PATH
    folderPath = VBA.Left$(logPath, VBA.InStrRev(logPath, "\\") - 1)

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        If VBA.Len(folderPath) > 0 Then
            If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
        End If
    End If

    Set stream = fso.OpenTextFile(logPath, 2, True) ' режим перезаписи файла (ForWriting)
    If Not stream Is Nothing Then stream.Close
    Err.Clear
    On Error GoTo 0
#End If
End Sub


Private Sub private_ClearDocumentModuleCode(ByVal vbComp As Object)
    Dim cm As Object

    If vbComp Is Nothing Then Exit Sub
    Set cm = vbComp.CodeModule
    If cm Is Nothing Then Exit Sub
    If cm.CountOfLines <= 0 Then Exit Sub

    cm.DeleteLines 1, cm.CountOfLines
End Sub
