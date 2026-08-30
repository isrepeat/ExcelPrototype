' Должен быть вставлен во внутренний модуль книги .xlsm
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const RUNTIME_SNAPSHOTS_ENABLED = False
#Const LOGGING_VERBOSE_ENABLED = False

#Const CORE_ENABLE_STATUS_BAR_LOGGING = True
#Const CORE_FORCE_NATIVE_STATUS_BAR = True
#Const CORE_ENABLE_SELF_LOGGING = True

Private Const BASE_DIR As String = "vba\\"
Private Const IMPORT_CACHE_FILE As String = ".devtools_import_cache.txt"
Private Const CORE_LOG_FILE_REL_PATH As String = "Logs\\core.log"
Private Const ENABLE_CLASS_IMPORT_VALIDATION As Boolean = False
Private Const MAX_IMPORT_RECURSION_DEPTH As Long = 4
Private Const COMP_TYPE_MODULE As String = "module"
Private Const COMP_TYPE_CLASS As String = "class"
Private Const COMP_TYPE_SHEET As String = "sheet"
Private Const COMP_TYPE_WORKBOOK As String = "workbook"
Private Const VB_COMPONENT_TYPE_DOCUMENT As Long = 100
Private Const MAX_VBA_COMPONENT_NAME_LEN As Long = 31
Private Const UPDATE_MODE_FULL As Long = 1
Private Const UPDATE_MODE_DATE As Long = 2
Private Const UPDATE_MODE_SIZE As Long = 3
Private Const CORE_COMPONENT_NAME As String = "ex_Core"
Private Const PATTERN_ALL_COMPONENTS As String = ".+"
Private Const PATTERN_MAIN_COMPONENTS As String = "^(?!rt_).+"
Private Const PATTERN_RUNTIME_COMPONENTS As String = "^rt_.+"
Private Const PATTERN_EXCLUDE_CORE As String = "^ex_core$"

Private Const SETTINGS_FILE_NAME As String = "Settings.xml"
Private Const SETTINGS_ROOT_NODE As String = "Settings"
Private Const SETTINGS_FLAGS_NODE As String = "Flags"
Private Const SETTINGS_FLAG_IS_LOGGING_ENABLED As String = "IsLoggingEnabled"
Private Const SETTINGS_FLAG_IS_LOGGING_ENABLED_DEFAULT As Boolean = True
Private Const SETTINGS_FLAG_IS_LOGGING_MAIN_PAGE_ONLY As String = "IsLoggingMainPageOnly"
Private Const SETTINGS_FLAG_IS_LOGGING_MAIN_PAGE_ONLY_DEFAULT As Boolean = True
Private Const SETTINGS_FLAG_IS_LOGGING_VERBOSE_ENABLED As String = "IsLoggingVerboseEnabled"
Private Const SETTINGS_FLAG_IS_LOGGING_VERBOSE_ENABLED_DEFAULT As Boolean = False
Private Const MAIN_PAGE_WORKSHEET_NAME As String = "Main"

Private g_QueuedRuntimeStateRestoreAt As Date
Private g_QueuedRuntimeStateRestoreMacro As String

Private g_FileCacheMap As Object
Private g_GlobalItemsSourceMap As Object
Private g_GlobalObjectSourceMap As Object
Private g_IsLoggingRuntimeStateInitialized As Boolean
Private g_IsLoggingRuntimeActive As Boolean

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:ex_Core.fn_Module_Dispose"
#End If
    Call fn_CancelDeferredTasks

    ' После hot-update политика будет заново применена новой startup-сессией.
    ' До этого момента не блокируем диагностические сообщения update pipeline.
    g_IsLoggingRuntimeStateInitialized = False
    g_IsLoggingRuntimeActive = True

    On Error Resume Next
    Set g_FileCacheMap = Nothing
    Call private_RuntimeSource_ResetStorage
    On Error GoTo 0
End Sub


Public Sub fn_CancelDeferredTasks()
    ' Отложенный restore ссылается на ThisWorkbook в имени
    ' макроса. Если Excel остается запущенным из-за другой книги, такой OnTime
    ' после закрытия может повторно открыть PrototypeNew.
    Call private_Dev_CancelQueuedRuntimeStateRestore("lifecycle:ex_Core.fn_CancelDeferredTasks")
End Sub


' //
' // API
' //
' --------------------------------------
'  namespace Dev {
' --------------------------------------
Public Sub fn_Dev_OpenCoreModule()
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
        private_ShowStatusWarning "Module ex_Core was not found in the VBA project.", True, 6
        Exit Sub
    End If

    comp.Activate
    private_ShowStatusNotice "Module ex_Core was opened in the VBA editor.", True, 2
    Exit Sub

EH:
    private_ShowStatusError "Failed to open module ex_Core: " & Err.Description, True, 6
End Sub


Public Sub fn_Dev_RemoveAllModulesAndClasses()
    On Error GoTo EH
    Application.ScreenUpdating = False

    private_Dev_RemoveAllModulesAndClasses PATTERN_ALL_COMPONENTS, PATTERN_EXCLUDE_CORE

    Application.ScreenUpdating = True
    private_ShowStatusSuccess "Modules and classes were removed; document modules were cleared (ex_Core preserved).", True, 3
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent "remove-modules-classes: success"
#End If
    Exit Sub

EH:
    Application.ScreenUpdating = True
    private_ShowStatusError "Failed to remove modules/classes: " & Err.Description, True, 6
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent "remove-modules-classes: fail: " & Err.Description
#End If
End Sub


' Callstack[1]: VBA.Macros(ex_Core.fn_Dev_UpdateAllModules) -> ex_Core.fn_Dev_UpdateAllModules
' Callstack[2]: DevUI(onClick: ex_Core.fn_Dev_UpdateAllModules) -> ex_Core.fn_Dev_UpdateAllModules
Public Sub fn_Dev_UpdateAllModules()
    If private_Dev_TryRunSafeUpdateByMode(UPDATE_MODE_FULL, PATTERN_ALL_COMPONENTS, PATTERN_EXCLUDE_CORE, True, "full") Then Exit Sub
    private_ShowStatusError "Safe update (full) did not complete. Check core.log.", True, 6
End Sub


' Callstack[1]: VBA.Macros(ex_Core.fn_Dev_UpdateCodeByDate) -> ex_Core.fn_Dev_UpdateCodeByDate
Public Sub fn_Dev_UpdateCodeByDate()
    If private_Dev_TryRunSafeUpdateByMode(UPDATE_MODE_DATE, PATTERN_MAIN_COMPONENTS, PATTERN_EXCLUDE_CORE, False, "date") Then Exit Sub
    private_ShowStatusError "Safe update (date) did not complete. Check core.log.", True, 6
End Sub


' Callstack[1]: VBA.Macros(ex_Core.fn_Dev_UpdateCodeBySize) -> ex_Core.fn_Dev_UpdateCodeBySize
Public Sub fn_Dev_UpdateCodeBySize()
    If private_Dev_TryRunSafeUpdateByMode(UPDATE_MODE_SIZE, PATTERN_MAIN_COMPONENTS, PATTERN_EXCLUDE_CORE, False, "size") Then Exit Sub
    private_ShowStatusError "Safe update (size) did not complete. Check core.log.", True, 6
End Sub


' Отдельная процедура для ядра рантайма (rt_*):
' Инкрементальные обновления по дате/размеру эти модули не затрагивают.
Public Sub fn_Dev_UpdateRuntimeCore()
    If private_Dev_UpdateCodeByRegex(PATTERN_RUNTIME_COMPONENTS, PATTERN_EXCLUDE_CORE, UPDATE_MODE_FULL, True) Then Exit Sub
    private_ShowStatusError "Runtime core update did not complete. Check core.log.", True, 6
End Sub


' Callstack[1]: DevUI(onClick: ex_Core.fn_Dev_ToggleLogging) -> ex_Core.fn_Dev_ToggleLogging
Public Sub fn_Dev_ToggleLogging()
    Dim isEnabled As Boolean

    If Not fn_Settings_TryToggleFlagBoolean(SETTINGS_FLAG_IS_LOGGING_ENABLED, SETTINGS_FLAG_IS_LOGGING_ENABLED_DEFAULT, isEnabled, True) Then
        private_ShowStatusError "Failed to update EnableLogging in Settings.xml.", True, 6
        Exit Sub
    End If

    If isEnabled Then
        private_ShowStatusSuccess "Logging is enabled (Settings.xml).", True, 3
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "settings:enable-logging=true"
#End If
    Else
        private_ShowStatusWarning "Logging is disabled (Settings.xml).", True, 3
    End If

    ' Перерисовываем только кнопку по уже записанному значению Settings.xml.
    ' Полный render активной страницы после Clear Pages мог завершиться неуспешно,
    ' а прежний Call игнорировал Boolean-результат и оставлял stale caption.
    If Not private_TryRefreshRuntimeStaticControl("ToggleLogging") Then
        VBA.MsgBox "PrototypeNew: logging state was changed, but the 'ToggleLogging' button could not be refreshed.", _
            VBA.vbExclamation, "PrototypeNew / Logging"
    End If
End Sub


Public Sub fn_Dev_ToggleMainPageLogging()
    Dim isMainPageOnly As Boolean
    Dim activeSheetName As String

    If Not fn_Settings_TryToggleFlagBoolean( _
        SETTINGS_FLAG_IS_LOGGING_MAIN_PAGE_ONLY, _
        SETTINGS_FLAG_IS_LOGGING_MAIN_PAGE_ONLY_DEFAULT, _
        isMainPageOnly, _
        True) Then
        private_ShowStatusError "Failed to update IsLoggingMainPageOnly in Settings.xml.", True, 6
        Exit Sub
    End If

    On Error Resume Next
    activeSheetName = VBA.CStr(Application.ActiveSheet.Name)
    On Error GoTo 0
    Call fn_Diagnostic_ApplyLoggingPagePolicy(activeSheetName)

    If Not private_TryRefreshRuntimeStaticControl("ToggleMainPageLogging") Then
        VBA.MsgBox "PrototypeNew: logging page policy was changed, but the toggle button could not be refreshed.", _
            VBA.vbExclamation, "PrototypeNew / Logging"
    End If
End Sub


Public Sub fn_Dev_ToggleVerboseLogging()
    Dim isVerboseEnabled As Boolean

    If Not fn_Settings_TryToggleFlagBoolean( _
        SETTINGS_FLAG_IS_LOGGING_VERBOSE_ENABLED, _
        SETTINGS_FLAG_IS_LOGGING_VERBOSE_ENABLED_DEFAULT, _
        isVerboseEnabled, _
        True) Then
        private_ShowStatusError "Failed to update IsLoggingVerboseEnabled in Settings.xml.", True, 6
        Exit Sub
    End If

    If isVerboseEnabled Then
        private_ShowStatusSuccess "Verbose logging is enabled (Settings.xml).", True, 3
    Else
        private_ShowStatusWarning "Verbose logging is disabled (Settings.xml).", True, 3
    End If

    If Not private_TryRefreshRuntimeStaticControl("ToggleVerboseLogging") Then
        VBA.MsgBox "PrototypeNew: verbose logging state was changed, but the 'ToggleVerboseLogging' button could not be refreshed.", _
            VBA.vbExclamation, "PrototypeNew / Logging"
    End If
End Sub


Public Sub fn_Dev_ClearLogs()
    private_Diagnostic_ClearCoreLogFile
    private_ShowStatusSuccess "Log file was cleared.", True, 3
End Sub

Public Sub fn_Dev_OpenLogs()
    Dim logPath As String
    Dim commandText As String
    Dim shellRunner As Object
    Dim errorText As String

    If VBA.Len(VBA.Trim$(ThisWorkbook.Path)) = 0 Then
        VBA.MsgBox "Cannot open the log because the workbook has not been saved.", vbExclamation, "PrototypeNew / Logging"
        Exit Sub
    End If

    logPath = ThisWorkbook.Path & "\\" & CORE_LOG_FILE_REL_PATH
    If VBA.Len(VBA.Dir$(logPath, VBA.vbNormal)) = 0 Then private_Diagnostic_ClearCoreLogFile

    On Error GoTo EH
    commandText = "notepad.exe """ & logPath & """"
    Set shellRunner = VBA.CreateObject("WScript.Shell")
    shellRunner.Run commandText, VBA.vbNormalFocus, False
    Set shellRunner = Nothing
    Exit Sub
EH:
    errorText = Err.Description
    Set shellRunner = Nothing
    VBA.MsgBox "Failed to open the log file in Notepad: " & errorText, vbExclamation, "PrototypeNew / Logging"
End Sub
' --------------------------------------
'  } // namespace Dev
' --------------------------------------

' --------------------------------------
'  namespace RuntimeSource {
' --------------------------------------
' Callstack[1]: External startup/init -> ex_Core.fn_RuntimeSource_SetGlobalItemsSource
' Глобальные sources живут в ex_Core, чтобы избежать зависимости Settings/Diagnostic логики от rt_PageManager.
Public Function fn_RuntimeSource_SetGlobalItemsSource(ByVal sourceKey As String, ByVal items As Collection) As Boolean
    Dim normalizedKey As String

    normalizedKey = private_RuntimeSource_NormalizeKey(sourceKey)
    If VBA.Len(normalizedKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "RuntimeSource: global items source key is empty."
#End If
        Exit Function
    End If
    If items Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "RuntimeSource: global items source collection is not specified for key '" & normalizedKey & "'."
#End If
        Exit Function
    End If

    private_RuntimeSource_EnsureStorage
    Set g_GlobalItemsSourceMap(normalizedKey) = items
    fn_RuntimeSource_SetGlobalItemsSource = True
End Function


' Callstack[1]: External lifecycle/reset -> ex_Core.fn_RuntimeSource_RemoveGlobalItemsSource
Public Function fn_RuntimeSource_RemoveGlobalItemsSource(ByVal sourceKey As String) As Boolean
    Dim normalizedKey As String

    normalizedKey = private_RuntimeSource_NormalizeKey(sourceKey)
    If VBA.Len(normalizedKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "RuntimeSource: global items source key is empty."
#End If
        Exit Function
    End If

    private_RuntimeSource_EnsureStorage
    If g_GlobalItemsSourceMap.Exists(normalizedKey) Then
        g_GlobalItemsSourceMap.Remove normalizedKey
    End If
    fn_RuntimeSource_RemoveGlobalItemsSource = True
End Function


' Callstack[1]: ex_RuntimeSourceResolver.fn_TryResolveItemsSource -> ex_Core.fn_RuntimeSource_TryGetGlobalItemsSourceByKey
Public Function fn_RuntimeSource_TryGetGlobalItemsSourceByKey( _
    ByVal sourceKey As String, _
    ByRef outItems As Collection, _
    Optional ByVal allowMissing As Boolean = False _
) As Boolean
    Dim normalizedKey As String

    Set outItems = Nothing
    normalizedKey = private_RuntimeSource_NormalizeKey(sourceKey)

    If VBA.Len(normalizedKey) = 0 Then
        If allowMissing Then
            fn_RuntimeSource_TryGetGlobalItemsSourceByKey = True
            Exit Function
        End If
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "RuntimeSource: global items source key is empty."
#End If
        Exit Function
    End If

    private_RuntimeSource_EnsureStorage
    If g_GlobalItemsSourceMap.Exists(normalizedKey) Then
        Set outItems = g_GlobalItemsSourceMap(normalizedKey)
        fn_RuntimeSource_TryGetGlobalItemsSourceByKey = True
        Exit Function
    End If

    If allowMissing Then
        fn_RuntimeSource_TryGetGlobalItemsSourceByKey = True
        Exit Function
    End If

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "RuntimeSource: global items source '" & normalizedKey & "' is not registered."
#End If
End Function


' Callstack[1]: External startup/init -> ex_Core.fn_RuntimeSource_SetGlobalObjectSource
Public Function fn_RuntimeSource_SetGlobalObjectSource(ByVal sourceKey As String, ByVal sourceObject As Object) As Boolean
    Dim normalizedKey As String

    normalizedKey = private_RuntimeSource_NormalizeKey(sourceKey)
    If VBA.Len(normalizedKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "RuntimeSource: global object source key is empty."
#End If
        Exit Function
    End If
    If sourceObject Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "RuntimeSource: global object source is not specified for key '" & normalizedKey & "'."
#End If
        Exit Function
    End If

    private_RuntimeSource_EnsureStorage
    Set g_GlobalObjectSourceMap(normalizedKey) = sourceObject
    fn_RuntimeSource_SetGlobalObjectSource = True
End Function


' Callstack[1]: External lifecycle/reset -> ex_Core.fn_RuntimeSource_RemoveGlobalObjectSource
Public Function fn_RuntimeSource_RemoveGlobalObjectSource(ByVal sourceKey As String) As Boolean
    Dim normalizedKey As String

    normalizedKey = private_RuntimeSource_NormalizeKey(sourceKey)
    If VBA.Len(normalizedKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "RuntimeSource: global object source key is empty."
#End If
        Exit Function
    End If

    private_RuntimeSource_EnsureStorage
    If g_GlobalObjectSourceMap.Exists(normalizedKey) Then
        g_GlobalObjectSourceMap.Remove normalizedKey
    End If
    fn_RuntimeSource_RemoveGlobalObjectSource = True
End Function


' Callstack[1]: ex_RuntimeSourceResolver.fn_TryResolveObjectSource -> ex_Core.fn_RuntimeSource_TryGetGlobalObjectSourceByKey
Public Function fn_RuntimeSource_TryGetGlobalObjectSourceByKey( _
    ByVal sourceKey As String, _
    ByRef outObject As Object, _
    Optional ByVal allowMissing As Boolean = False _
) As Boolean
    Dim normalizedKey As String

    Set outObject = Nothing
    normalizedKey = private_RuntimeSource_NormalizeKey(sourceKey)

    If VBA.Len(normalizedKey) = 0 Then
        If allowMissing Then
            fn_RuntimeSource_TryGetGlobalObjectSourceByKey = True
            Exit Function
        End If
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "RuntimeSource: global object source key is empty."
#End If
        Exit Function
    End If

    private_RuntimeSource_EnsureStorage

    If g_GlobalObjectSourceMap.Exists(normalizedKey) Then
        Set outObject = g_GlobalObjectSourceMap(normalizedKey)
        fn_RuntimeSource_TryGetGlobalObjectSourceByKey = True
        Exit Function
    End If

    If allowMissing Then
        fn_RuntimeSource_TryGetGlobalObjectSourceByKey = True
        Exit Function
    End If

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "RuntimeSource: global object source '" & normalizedKey & "' is not registered."
#End If
End Function
' --------------------------------------
'  } // namespace RuntimeSource
' --------------------------------------

' --------------------------------------
'  namespace Settings {
' --------------------------------------
' Callstack[1]: ex_Core.fn_Settings_TryToggleFlagBoolean -> ex_Core.fn_Settings_TryGetFlagBoolean
' Callstack[2]: ex_Core.private_Diagnostic_LogCoreEvent -> ex_Core.fn_Settings_TryGetFlagBoolean
Public Function fn_Settings_TryGetFlagBoolean( _
    ByVal flagName As String, _
    ByVal defaultValue As Boolean, _
    ByRef outValue As Boolean, _
    Optional ByVal showErrorUi As Boolean = True _
) As Boolean
    Dim flagsMap As Object
    Dim rawValue As Variant
    Dim parsedValue As Boolean
    Dim normalizedFlagName As String

    outValue = defaultValue
    normalizedFlagName = fn_Helpers_NormilizeString(flagName)
    If VBA.Len(normalizedFlagName) = 0 Then
        If showErrorUi Then ex_Core.fn_Diagnostic_LogError "Settings: flag name is empty."
        Exit Function
    End If

    If Not private_Settings_TryEnsureFlagsMapCurrent(flagsMap, showErrorUi) Then Exit Function
    If flagsMap Is Nothing Then Exit Function

    If Not flagsMap.Exists(normalizedFlagName) Then
        fn_Settings_TryGetFlagBoolean = True
        Exit Function
    End If

    rawValue = flagsMap(normalizedFlagName)
    If fn_Helpers_TryParseBooleanText(VBA.CStr(rawValue), parsedValue) Then
        outValue = parsedValue
    ElseIf showErrorUi Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Settings: flag '" & normalizedFlagName & "' has non-boolean value '" & VBA.CStr(rawValue) & "'."
#End If
    End If

    fn_Settings_TryGetFlagBoolean = True
End Function


' Callstack[1]: External caller (macro/runtime action) -> ex_Core.fn_Settings_TrySetFlagBoolean
Public Function fn_Settings_TrySetFlagBoolean( _
    ByVal flagName As String, _
    ByVal newValue As Boolean, _
    Optional ByVal showErrorUi As Boolean = True _
) As Boolean
    Dim normalizedFlagName As String

    normalizedFlagName = fn_Helpers_NormilizeString(flagName)
    If VBA.Len(normalizedFlagName) = 0 Then
        If showErrorUi Then ex_Core.fn_Diagnostic_LogError "Settings: flag name is empty."
        Exit Function
    End If

    fn_Settings_TrySetFlagBoolean = private_Settings_TryWriteFlagBoolean(normalizedFlagName, VBA.CBool(newValue), showErrorUi)
End Function


' Callstack[1]: ex_Core.fn_Dev_ToggleLogging -> ex_Core.fn_Settings_TryToggleFlagBoolean
Public Function fn_Settings_TryToggleFlagBoolean( _
    ByVal flagName As String, _
    ByVal defaultValue As Boolean, _
    ByRef outValue As Boolean, _
    Optional ByVal showErrorUi As Boolean = True _
) As Boolean
    Dim currentValue As Boolean
    Dim nextValue As Boolean
    Dim normalizedFlagName As String

    outValue = defaultValue
    normalizedFlagName = fn_Helpers_NormilizeString(flagName)
    If VBA.Len(normalizedFlagName) = 0 Then
        If showErrorUi Then ex_Core.fn_Diagnostic_LogError "Settings: flag name is empty."
        Exit Function
    End If

    If Not fn_Settings_TryGetFlagBoolean(normalizedFlagName, defaultValue, currentValue, showErrorUi) Then Exit Function
    nextValue = Not currentValue
    If Not private_Settings_TryWriteFlagBoolean(normalizedFlagName, nextValue, showErrorUi) Then Exit Function

    outValue = nextValue

    fn_Settings_TryToggleFlagBoolean = True
End Function


' Callstack[1]: External caller (debug/runtime action) -> ex_Core.fn_Settings_TryGetDom
' forceReload очищает только file-text cache; сам файл остается источником истины.
Public Function fn_Settings_TryGetDom( _
    ByRef outDom As Object, _
    Optional ByVal forceReload As Boolean = False, _
    Optional ByVal showErrorUi As Boolean = True _
) As Boolean
    Dim partObj As Object
    Dim settingsPath As String

    Set outDom = Nothing
    If VBA.CBool(forceReload) Then
        If private_Settings_TryResolveFilePath(settingsPath, False) Then
            Call private_FileCache_RemoveFileText(settingsPath)
        End If
    End If

    If Not private_Settings_TryGetSettingsDom(outDom, partObj, True, showErrorUi) Then Exit Function
    fn_Settings_TryGetDom = True
End Function


' Callstack[1]: ex_RuntimeSourceResolver.fn_TryResolveObjectSource(GlobalRuntimeSource='settings') -> ex_Core.fn_Settings_TryGetObjectSource
' Возвращает snapshot-объект настроек из Settings.xml; чтение XML переиспользует общий file-cache по DateLastModified.
Public Function fn_Settings_TryGetObjectSource( _
    ByRef outObjectSource As Object, _
    Optional ByVal showErrorUi As Boolean = True _
) As Boolean
    Dim settingsMap As Object

    Set outObjectSource = Nothing

    Set settingsMap = VBA.CreateObject("Scripting.Dictionary")
    settingsMap.CompareMode = 1
    ' Объект строится из общего file-cache; XML читаем только при изменении DateLastModified у Settings.xml.
    If Not private_Settings_TryFillObjectSourceMap(settingsMap, showErrorUi) Then Exit Function

    Set outObjectSource = settingsMap
    fn_Settings_TryGetObjectSource = True
End Function
' --------------------------------------
'  } // namespace Settings
' --------------------------------------

' --------------------------------------
'  namespace Diagnostic {
' --------------------------------------
Public Sub fn_Diagnostic_LogInfo(ByVal messageText As String)
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreEvent VBA.CStr(messageText)
#End If
End Sub


Public Sub fn_Diagnostic_LogError(ByVal messageText As String)
    messageText = VBA.Trim$(VBA.CStr(messageText))
    If VBA.Len(messageText) = 0 Then Exit Sub
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreEvent "error: " & messageText
#End If
End Sub


' Событийная трассировка должна быть доступна и на generated-страницах даже
' при IsLoggingMainPageOnly=true. Master switch IsLoggingEnabled сохраняется.
Public Sub fn_Diagnostic_LogEventInfo(ByVal messageText As String)
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreEvent VBA.CStr(messageText), True
#End If
End Sub


Public Sub fn_Diagnostic_LogEventError(ByVal messageText As String)
    messageText = VBA.Trim$(VBA.CStr(messageText))
    If VBA.Len(messageText) = 0 Then Exit Sub
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreEvent "error: " & messageText, True
#End If
End Sub


Public Sub fn_Diagnostic_LogWarning(ByVal messageText As String)
    messageText = VBA.Trim$(VBA.CStr(messageText))
    If VBA.Len(messageText) = 0 Then Exit Sub
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreEvent "warning: " & messageText
#End If
End Sub


Public Sub fn_Diagnostic_LogVerbose(ByVal messageText As String)
    Dim isVerboseEnabled As Boolean

    If Not fn_Settings_TryGetFlagBoolean( _
        SETTINGS_FLAG_IS_LOGGING_VERBOSE_ENABLED, _
        SETTINGS_FLAG_IS_LOGGING_VERBOSE_ENABLED_DEFAULT, _
        isVerboseEnabled, _
        False) Then Exit Sub
    If Not isVerboseEnabled Then Exit Sub

#If LOGGING_DEBUG_ENABLED Then
    ' Verbose использует ту же page-policy, что Info/Error/Warning.
    private_Diagnostic_LogCoreEvent "verbose: " & VBA.CStr(messageText)
#End If
End Sub


' Startup/shutdown диагностика не должна зависеть от активного листа: падение
' часто происходит после удаления или деактивации страницы, когда page-policy
' уже не соответствует фактическому lifecycle-контексту.
Public Sub fn_Diagnostic_BeginLifecycleLogging()
    g_IsLoggingRuntimeActive = True
    g_IsLoggingRuntimeStateInitialized = True
End Sub


Public Sub fn_Diagnostic_LogStatusBarMessage( _
    ByVal actionName As String, _
    ByVal messageText As String, _
    Optional ByVal timeoutSeconds As Long = 0 _
)
    private_Diagnostic_LogStatusBarEvent actionName, messageText, timeoutSeconds
End Sub


' Применяет только runtime-фильтр. Значение IsLoggingEnabled остаётся
' пользовательским master switch и не перезаписывается при переходах по листам.
Public Sub fn_Diagnostic_ApplyLoggingPagePolicy(ByVal worksheetName As String)
    Dim isMainPageOnly As Boolean

    If Not fn_Settings_TryGetFlagBoolean( _
        SETTINGS_FLAG_IS_LOGGING_MAIN_PAGE_ONLY, _
        SETTINGS_FLAG_IS_LOGGING_MAIN_PAGE_ONLY_DEFAULT, _
        isMainPageOnly, _
        False) Then
        ' Ошибка чтения Settings не должна молча включать ограничительный режим.
        g_IsLoggingRuntimeActive = True
        g_IsLoggingRuntimeStateInitialized = True
        Exit Sub
    End If

    If isMainPageOnly Then
        g_IsLoggingRuntimeActive = (VBA.StrComp( _
            VBA.Trim$(worksheetName), _
            MAIN_PAGE_WORKSHEET_NAME, _
            VBA.vbTextCompare) = 0)
    Else
        g_IsLoggingRuntimeActive = True
    End If
    g_IsLoggingRuntimeStateInitialized = True
End Sub


Public Sub fn_Diagnostic_ClearCoreLog()
    private_Diagnostic_ClearCoreLogFile
End Sub
' --------------------------------------
'  } // namespace Diagnostic
' --------------------------------------

' --------------------------------------
'  namespace CustomXmlPartStore {
' --------------------------------------
Public Function fn_CustomXmlPartStore_TryFindPartByNamespace( _
    ByVal namespaceUri As String, _
    ByRef outPart As Object _
) As Boolean
    Dim parts As Object

    namespaceUri = VBA.Trim$(namespaceUri)
    If VBA.Len(namespaceUri) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "CustomXmlPartStore: namespace is empty."
#End If
        Exit Function
    End If

    On Error GoTo EH_FIND
    Set parts = ThisWorkbook.CustomXMLParts.SelectByNamespace(namespaceUri)
    On Error GoTo 0

    If Not parts Is Nothing Then
        If parts.Count > 0 Then
            Set outPart = parts(1)
        End If
    End If

    fn_CustomXmlPartStore_TryFindPartByNamespace = True
    Exit Function

EH_FIND:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "CustomXmlPartStore: failed to find XML part by namespace '" & namespaceUri & "': " & Err.Description
#End If
End Function


Public Function fn_CustomXmlPartStore_TryLoadDomFromXml( _
    ByVal xmlText As String, _
    ByRef outDom As Object _
) As Boolean
    Dim dom As Object

    Set dom = VBA.CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False
    dom.validateOnParse = False
    dom.setProperty "SelectionLanguage", "XPath"

    If Not dom.LoadXML(VBA.CStr(xmlText)) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "CustomXmlPartStore: failed to parse XML."
#End If
        Exit Function
    End If

    Set outDom = dom
    fn_CustomXmlPartStore_TryLoadDomFromXml = True
End Function


Public Function fn_CustomXmlPartStore_TryCreateEmptyDom( _
    ByVal rootNodeName As String, _
    ByVal namespaceUri As String, _
    ByRef outDom As Object _
) As Boolean
    Dim xmlText As String

    rootNodeName = VBA.Trim$(rootNodeName)
    namespaceUri = VBA.Trim$(namespaceUri)

    If VBA.Len(rootNodeName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "CustomXmlPartStore: root node name is empty."
#End If
        Exit Function
    End If
    If VBA.Len(namespaceUri) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "CustomXmlPartStore: namespace is empty."
#End If
        Exit Function
    End If

    xmlText = "<?xml version=""1.0"" encoding=""UTF-8""?>" & _
              "<" & rootNodeName & " xmlns=""" & namespaceUri & """></" & rootNodeName & ">"

    If Not fn_CustomXmlPartStore_TryLoadDomFromXml(xmlText, outDom) Then Exit Function
    fn_CustomXmlPartStore_TryCreateEmptyDom = True
End Function


Public Function fn_CustomXmlPartStore_TryLoadPartDom( _
    ByVal partObj As Object, _
    ByRef outDom As Object _
) As Boolean
    If partObj Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "CustomXmlPartStore: part is not specified."
#End If
        Exit Function
    End If

    If Not fn_CustomXmlPartStore_TryLoadDomFromXml(VBA.CStr(partObj.XML), outDom) Then Exit Function
    fn_CustomXmlPartStore_TryLoadPartDom = True
End Function


Public Function fn_CustomXmlPartStore_TrySaveDom( _
    ByVal dom As Object, _
    ByVal existingPart As Object _
) As Boolean
    Dim xmlText As String

    If dom Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "CustomXmlPartStore: DOM is not specified."
#End If
        Exit Function
    End If

    xmlText = VBA.CStr(dom.XML)
    If VBA.Len(VBA.Trim$(xmlText)) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "CustomXmlPartStore: state XML is empty."
#End If
        Exit Function
    End If

    On Error GoTo EH_SAVE
    If Not existingPart Is Nothing Then existingPart.Delete
    ThisWorkbook.CustomXMLParts.Add xmlText
    fn_CustomXmlPartStore_TrySaveDom = True
    Exit Function

EH_SAVE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "CustomXmlPartStore: failed to persist state XML: " & Err.Description
#End If
End Function
' --------------------------------------
'  } // namespace CustomXmlPartStore
' --------------------------------------

' --------------------------------------
'  namespace Helpers {
' --------------------------------------
Public Function fn_Helpers_NormilizeString(ByVal value As String) As String
    fn_Helpers_NormilizeString = VBA.Trim$(VBA.CStr(value))
End Function


Public Function fn_Helpers_TryParseBooleanText(ByVal textValue As String, ByRef outValue As Boolean) As Boolean
    textValue = VBA.LCase$(fn_Helpers_NormilizeString(textValue))

    Select Case textValue
        Case "true", "1", "yes", "on"
            outValue = True
            fn_Helpers_TryParseBooleanText = True
        Case "false", "0", "no", "off"
            outValue = False
            fn_Helpers_TryParseBooleanText = True
    End Select
End Function


Public Function fn_Helpers_BoolToText(ByVal value As Boolean) As String
    If value Then
        fn_Helpers_BoolToText = "true"
    Else
        fn_Helpers_BoolToText = "false"
    End If
End Function


Public Function fn_Helpers_EndsWith(ByVal value As String, ByVal suffix As String) As Boolean
    fn_Helpers_EndsWith = (VBA.LCase$(VBA.Right$(value, VBA.Len(suffix))) = VBA.LCase$(suffix))
End Function


Public Function fn_Helpers_IsRegexMatch(ByVal valueText As String, ByVal regexPattern As String) As Boolean
    Dim re As Object

    regexPattern = fn_Helpers_NormilizeString(regexPattern)
    If VBA.Len(regexPattern) = 0 Then Exit Function

    On Error GoTo EH

    Set re = VBA.CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = regexPattern

    fn_Helpers_IsRegexMatch = re.Test(VBA.CStr(valueText))
    Exit Function

EH:
    Err.Raise VBA.vbObjectError + 1013, "fn_Helpers_IsRegexMatch", "Некорректный regex '" & regexPattern & "': " & Err.Description
End Function


Public Function fn_Helpers_TryGetFileText( _
    ByVal filePath As String, _
    ByRef outText As String, _
    Optional ByVal allowMissing As Boolean = True, _
    Optional ByVal showErrorUi As Boolean = True _
) As Boolean
    Dim fileStamp As String
    Dim cacheKey As String
    Dim entry As Object

    outText = VBA.vbNullString
    filePath = fn_Helpers_NormilizeString(filePath)
    If VBA.Len(filePath) = 0 Then
        fn_Helpers_TryGetFileText = allowMissing
        Exit Function
    End If

    ' Общий read-through cache:
    ' 1) читаем stamp файла,
    ' 2) возвращаем cached text только при совпадении stamp,
    ' 3) при промахе читаем файл и обновляем cache entry.
    cacheKey = private_FileCache_BuildFileTextKey(filePath)
    If Not private_FileCache_TryGetFileStamp(filePath, fileStamp, showErrorUi) Then
        Call private_FileCache_Remove(cacheKey)
        If allowMissing And VBA.Len(VBA.Dir(filePath)) = 0 Then
            fn_Helpers_TryGetFileText = True
            Exit Function
        End If
        Exit Function
    End If

    If private_FileCache_TryGet(cacheKey, entry, True) Then
        If Not entry Is Nothing Then
            If entry.Exists("Stamp") Then
                If VBA.StrComp(VBA.CStr(entry("Stamp")), fileStamp, VBA.vbBinaryCompare) = 0 Then
                    If entry.Exists("Text") Then outText = VBA.CStr(entry("Text"))
                    fn_Helpers_TryGetFileText = True
                    Exit Function
                End If
            End If
        End If
    End If

    If Not private_FileCache_TryReadTextFile(filePath, outText, showErrorUi) Then Exit Function
    If Not private_FileCache_SetFileTextEntry(filePath, outText, fileStamp) Then Exit Function
    fn_Helpers_TryGetFileText = True
End Function
' --------------------------------------
'  } // namespace Helpers
' --------------------------------------

' //
' // Internal
' //
' --------------------------------------
'  namespace Settings {
' --------------------------------------
Private Function private_Settings_TryGetSettingsDom( _
    ByRef outDom As Object, _
    ByRef outPart As Object, _
    Optional ByVal createIfMissing As Boolean = True, _
    Optional ByVal showErrorUi As Boolean = True _
) As Boolean
    Dim settingsPath As String
    Dim settingsXmlText As String
    Dim hasStructureChanges As Boolean

    Set outDom = Nothing
    Set outPart = Nothing

    If Not private_Settings_TryResolveFilePath(settingsPath, showErrorUi) Then Exit Function
    If createIfMissing Then
        If Not private_Settings_TryEnsureTemplateExists(settingsPath, showErrorUi) Then Exit Function
    Else
        If VBA.Len(VBA.Dir(settingsPath)) = 0 Then
            If showErrorUi Then
#If LOGGING_DEBUG_ENABLED Then
                ex_Core.fn_Diagnostic_LogError "Settings: file '" & settingsPath & "' was not found."
#End If
            End If
            Exit Function
        End If
    End If

    If Not fn_Helpers_TryGetFileText(settingsPath, settingsXmlText, createIfMissing, showErrorUi) Then Exit Function
    If VBA.Len(VBA.Trim$(settingsXmlText)) = 0 Then
        If Not createIfMissing Then Exit Function
        settingsXmlText = private_Settings_BuildTemplateXml
    End If

    ' Settings всегда читаются из файла (через FileCache), поэтому объект не stale:
    ' при изменении DateLastModified helper автоматически перечитает XML с диска.
    If Not fn_CustomXmlPartStore_TryLoadDomFromXml(settingsXmlText, outDom) Then Exit Function
    If Not private_Settings_TryEnsureSettingsStructure(outDom, hasStructureChanges, showErrorUi) Then Exit Function
    If hasStructureChanges Then
        If Not private_Settings_TryWriteSettingsDomToFile(settingsPath, outDom, showErrorUi) Then Exit Function
    End If

    private_Settings_TryGetSettingsDom = True
End Function


Private Function private_Settings_TryEnsureSettingsStructure( _
    ByVal settingsDom As Object, _
    ByRef outIsChanged As Boolean, _
    ByVal showErrorUi As Boolean _
) As Boolean
    Dim settingsNode As Object
    Dim flagsNode As Object

    outIsChanged = False
    If settingsDom Is Nothing Then Exit Function

    Set settingsNode = settingsDom.selectSingleNode("/*[local-name()='" & SETTINGS_ROOT_NODE & "']")
    If settingsNode Is Nothing Then
        If settingsDom.DocumentElement Is Nothing Then
            Set settingsNode = settingsDom.createElement(SETTINGS_ROOT_NODE)
            settingsDom.appendChild settingsNode
            outIsChanged = True
        Else
            If showErrorUi Then
#If LOGGING_DEBUG_ENABLED Then
                ex_Core.fn_Diagnostic_LogError "Settings: unexpected root node '" & VBA.CStr(settingsDom.DocumentElement.baseName) & "'. Expected '" & SETTINGS_ROOT_NODE & "'."
#End If
            End If
            Exit Function
        End If
    End If

    Set flagsNode = settingsNode.selectSingleNode("*[local-name()='" & SETTINGS_FLAGS_NODE & "']")
    If flagsNode Is Nothing Then
        Set flagsNode = settingsDom.createElement(SETTINGS_FLAGS_NODE)
        settingsNode.appendChild flagsNode
        outIsChanged = True
    End If

    If Not private_Settings_TryEnsureDefaultFlagNode( _
        settingsDom, _
        flagsNode, _
        SETTINGS_FLAG_IS_LOGGING_ENABLED, _
        fn_Helpers_BoolToText(SETTINGS_FLAG_IS_LOGGING_ENABLED_DEFAULT), _
        outIsChanged, _
        showErrorUi) Then Exit Function

    If Not private_Settings_TryEnsureDefaultFlagNode( _
        settingsDom, _
        flagsNode, _
        SETTINGS_FLAG_IS_LOGGING_MAIN_PAGE_ONLY, _
        fn_Helpers_BoolToText(SETTINGS_FLAG_IS_LOGGING_MAIN_PAGE_ONLY_DEFAULT), _
        outIsChanged, _
        showErrorUi) Then Exit Function

    If Not private_Settings_TryEnsureDefaultFlagNode( _
        settingsDom, _
        flagsNode, _
        SETTINGS_FLAG_IS_LOGGING_VERBOSE_ENABLED, _
        fn_Helpers_BoolToText(SETTINGS_FLAG_IS_LOGGING_VERBOSE_ENABLED_DEFAULT), _
        outIsChanged, _
        showErrorUi) Then Exit Function

    private_Settings_TryEnsureSettingsStructure = True
End Function


Private Function private_Settings_TryEnsureDefaultFlagNode( _
    ByVal settingsDom As Object, _
    ByVal flagsNode As Object, _
    ByVal flagName As String, _
    ByVal defaultValue As String, _
    ByRef outIsChanged As Boolean, _
    ByVal showErrorUi As Boolean _
) As Boolean
    Dim normalizedFlagName As String
    Dim normalizedDefaultValue As String
    Dim flagNode As Object

    normalizedFlagName = fn_Helpers_NormilizeString(flagName)
    normalizedDefaultValue = fn_Helpers_NormilizeString(defaultValue)
    If VBA.Len(normalizedFlagName) = 0 Then Exit Function
    If VBA.Len(normalizedDefaultValue) = 0 Then Exit Function
    If settingsDom Is Nothing Then Exit Function
    If flagsNode Is Nothing Then Exit Function

    Set flagNode = flagsNode.selectSingleNode("*[local-name()='" & normalizedFlagName & "']")
    If flagNode Is Nothing Then
        On Error GoTo EH_CREATE_FLAG
        Set flagNode = settingsDom.createElement(normalizedFlagName)
        On Error GoTo 0
        flagNode.Text = normalizedDefaultValue
        flagsNode.appendChild flagNode
        outIsChanged = True
    ElseIf VBA.Len(VBA.Trim$(VBA.CStr(flagNode.Text))) = 0 Then
        flagNode.Text = normalizedDefaultValue
        outIsChanged = True
    End If

    private_Settings_TryEnsureDefaultFlagNode = True
    Exit Function

EH_CREATE_FLAG:
    If showErrorUi Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Settings: invalid default flag name '" & normalizedFlagName & "' for XML node."
#End If
    End If
    On Error GoTo 0
End Function


Private Function private_Settings_TryEnsureFlagsMapCurrent( _
    ByRef outFlagsMap As Object, _
    ByVal showErrorUi As Boolean _
) As Boolean
    Dim settingsDom As Object
    Dim settingsPart As Object

    Set outFlagsMap = Nothing

    If Not private_Settings_TryGetSettingsDom(settingsDom, settingsPart, True, showErrorUi) Then Exit Function
    If Not private_Settings_TryBuildFlagsMapFromDom(settingsDom, outFlagsMap, showErrorUi) Then Exit Function
    private_Settings_TryEnsureFlagsMapCurrent = True
End Function


Private Function private_Settings_TryBuildFlagsMapFromDom( _
    ByVal settingsDom As Object, _
    ByRef outFlagsMap As Object, _
    ByVal showErrorUi As Boolean _
) As Boolean
    Dim flagsNode As Object
    Dim childNode As Object
    Dim flagKey As String
    Dim rawText As String
    Dim parsedValue As Boolean
    Dim hasStructureChanges As Boolean

    Set outFlagsMap = Nothing
    If settingsDom Is Nothing Then Exit Function
    If Not private_Settings_TryEnsureSettingsStructure(settingsDom, hasStructureChanges, showErrorUi) Then Exit Function

    Set flagsNode = settingsDom.selectSingleNode("/*[local-name()='" & SETTINGS_ROOT_NODE & "']/*[local-name()='" & SETTINGS_FLAGS_NODE & "']")
    If flagsNode Is Nothing Then Exit Function

    Set outFlagsMap = VBA.CreateObject("Scripting.Dictionary")
    outFlagsMap.CompareMode = 1

    For Each childNode In flagsNode.ChildNodes
        If Not childNode Is Nothing Then
            If VBA.CLng(childNode.nodeType) = 1 Then
                flagKey = fn_Helpers_NormilizeString(VBA.CStr(childNode.baseName))
                If VBA.Len(flagKey) > 0 Then
                    rawText = VBA.Trim$(VBA.CStr(childNode.Text))
                    If fn_Helpers_TryParseBooleanText(rawText, parsedValue) Then
                        outFlagsMap(flagKey) = parsedValue
                    Else
                        outFlagsMap(flagKey) = rawText
                    End If
                End If
            End If
        End If
    Next childNode

    private_Settings_TryBuildFlagsMapFromDom = True
End Function


Private Function private_Settings_TryWriteFlagBoolean( _
    ByVal normalizedFlagName As String, _
    ByVal newValue As Boolean, _
    ByVal showErrorUi As Boolean _
) As Boolean
    private_Settings_TryWriteFlagBoolean = private_Settings_TryWriteFlagText(normalizedFlagName, fn_Helpers_BoolToText(VBA.CBool(newValue)), showErrorUi)
End Function


Private Function private_Settings_TryWriteFlagText( _
    ByVal normalizedFlagName As String, _
    ByVal newValue As String, _
    ByVal showErrorUi As Boolean _
) As Boolean
    Dim settingsDom As Object
    Dim settingsPart As Object
    Dim flagNode As Object
    Dim settingsPath As String
    Dim normalizedValue As String

    normalizedFlagName = fn_Helpers_NormilizeString(normalizedFlagName)
    normalizedValue = fn_Helpers_NormilizeString(newValue)
    If VBA.Len(normalizedFlagName) = 0 Then Exit Function
    If VBA.Len(normalizedValue) = 0 Then Exit Function

    If Not private_Settings_TryGetSettingsDom(settingsDom, settingsPart, True, showErrorUi) Then Exit Function
    If Not private_Settings_TryGetOrCreateFlagNode(settingsDom, normalizedFlagName, flagNode, showErrorUi) Then Exit Function

    flagNode.Text = normalizedValue

    If Not private_Settings_TryResolveFilePath(settingsPath, showErrorUi) Then Exit Function
    If Not private_Settings_TryWriteSettingsDomToFile(settingsPath, settingsDom, showErrorUi) Then Exit Function

    private_Settings_TryWriteFlagText = True
End Function


Private Function private_Settings_TryGetOrCreateFlagNode( _
    ByVal settingsDom As Object, _
    ByVal normalizedFlagName As String, _
    ByRef outFlagNode As Object, _
    ByVal showErrorUi As Boolean _
) As Boolean
    Dim flagsNode As Object
    Dim childNode As Object
    Dim hasStructureChanges As Boolean

    Set outFlagNode = Nothing
    If settingsDom Is Nothing Then Exit Function

    normalizedFlagName = fn_Helpers_NormilizeString(normalizedFlagName)
    If VBA.Len(normalizedFlagName) = 0 Then Exit Function

    If Not private_Settings_TryEnsureSettingsStructure(settingsDom, hasStructureChanges, showErrorUi) Then Exit Function
    Set flagsNode = settingsDom.selectSingleNode("/*[local-name()='" & SETTINGS_ROOT_NODE & "']/*[local-name()='" & SETTINGS_FLAGS_NODE & "']")
    If flagsNode Is Nothing Then Exit Function

    For Each childNode In flagsNode.ChildNodes
        If Not childNode Is Nothing Then
            If VBA.CLng(childNode.nodeType) = 1 Then
                If VBA.StrComp(fn_Helpers_NormilizeString(VBA.CStr(childNode.baseName)), normalizedFlagName, VBA.vbTextCompare) = 0 Then
                    Set outFlagNode = childNode
                    private_Settings_TryGetOrCreateFlagNode = True
                    Exit Function
                End If
            End If
        End If
    Next childNode

    On Error GoTo EH_CREATE_FLAG
    Set outFlagNode = settingsDom.createElement(normalizedFlagName)
    On Error GoTo 0
    flagsNode.appendChild outFlagNode
    private_Settings_TryGetOrCreateFlagNode = True
    Exit Function

EH_CREATE_FLAG:
    If showErrorUi Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Settings: invalid flag name '" & normalizedFlagName & "' for XML node."
#End If
    End If
    On Error GoTo 0
End Function


Private Function private_Settings_TryFillObjectSourceMap( _
    ByVal settingsMap As Object, _
    ByVal showErrorUi As Boolean _
) As Boolean
    Dim flagsMap As Object
    Dim key As Variant

    If settingsMap Is Nothing Then Exit Function

    If Not private_Settings_TryEnsureFlagsMapCurrent(flagsMap, showErrorUi) Then Exit Function
    If flagsMap Is Nothing Then
        private_Settings_TryFillObjectSourceMap = True
        Exit Function
    End If

    For Each key In flagsMap.Keys
        settingsMap(VBA.CStr(key)) = flagsMap(VBA.CStr(key))
    Next key
    private_Settings_TryFillObjectSourceMap = True
End Function


Private Function private_Settings_TryResolveFilePath(ByRef outSettingsPath As String, ByVal showErrorUi As Boolean) As Boolean
    Dim workbookPath As String

    workbookPath = VBA.Trim$(VBA.CStr(ThisWorkbook.Path))
    If VBA.Len(workbookPath) = 0 Then
        If showErrorUi Then ex_Core.fn_Diagnostic_LogError "Settings: workbook is not saved. Save .xlsm first."
        Exit Function
    End If

    outSettingsPath = workbookPath & "\\" & SETTINGS_FILE_NAME
    private_Settings_TryResolveFilePath = True
End Function


Private Function private_Settings_BuildTemplateXml() As String
    private_Settings_BuildTemplateXml = _
        "<?xml version=""1.0"" encoding=""UTF-8""?>" & VBA.vbCrLf & _
        "<Settings>" & VBA.vbCrLf & _
        "  <Flags>" & VBA.vbCrLf & _
    "    <" & SETTINGS_FLAG_IS_LOGGING_ENABLED & ">" & fn_Helpers_BoolToText(SETTINGS_FLAG_IS_LOGGING_ENABLED_DEFAULT) & "</" & SETTINGS_FLAG_IS_LOGGING_ENABLED & ">" & VBA.vbCrLf & _
    "    <" & SETTINGS_FLAG_IS_LOGGING_MAIN_PAGE_ONLY & ">" & fn_Helpers_BoolToText(SETTINGS_FLAG_IS_LOGGING_MAIN_PAGE_ONLY_DEFAULT) & "</" & SETTINGS_FLAG_IS_LOGGING_MAIN_PAGE_ONLY & ">" & VBA.vbCrLf & _
    "    <" & SETTINGS_FLAG_IS_LOGGING_VERBOSE_ENABLED & ">" & fn_Helpers_BoolToText(SETTINGS_FLAG_IS_LOGGING_VERBOSE_ENABLED_DEFAULT) & "</" & SETTINGS_FLAG_IS_LOGGING_VERBOSE_ENABLED & ">" & VBA.vbCrLf & _
        "  </Flags>" & VBA.vbCrLf & _
        "</Settings>"
End Function


Private Function private_Settings_TryEnsureTemplateExists(ByVal settingsPath As String, ByVal showErrorUi As Boolean) As Boolean
    settingsPath = VBA.Trim$(settingsPath)
    If VBA.Len(settingsPath) = 0 Then Exit Function

    If VBA.Len(VBA.Dir(settingsPath)) > 0 Then
        private_Settings_TryEnsureTemplateExists = True
        Exit Function
    End If

    private_Settings_TryEnsureTemplateExists = private_FileCache_SetFileText(settingsPath, private_Settings_BuildTemplateXml(), showErrorUi)
End Function


Private Function private_Settings_TryWriteSettingsDomToFile( _
    ByVal settingsPath As String, _
    ByVal settingsDom As Object, _
    ByVal showErrorUi As Boolean _
) As Boolean
    Dim xmlText As String

    If settingsDom Is Nothing Then Exit Function
    settingsPath = VBA.Trim$(settingsPath)
    If VBA.Len(settingsPath) = 0 Then Exit Function

    xmlText = VBA.CStr(settingsDom.XML)
    If VBA.Len(VBA.Trim$(xmlText)) = 0 Then Exit Function

    private_Settings_TryWriteSettingsDomToFile = private_FileCache_SetFileText(settingsPath, xmlText, showErrorUi)
End Function
' --------------------------------------
'  } // namespace Settings
' --------------------------------------

' --------------------------------------
'  namespace FileCache {
' --------------------------------------
Private Sub private_FileCache_EnsureStorage()
    If g_FileCacheMap Is Nothing Then
        Set g_FileCacheMap = VBA.CreateObject("Scripting.Dictionary")
        g_FileCacheMap.CompareMode = 1
    End If
End Sub


Private Function private_FileCache_NormalizeKey(ByVal cacheKey As String) As String
    private_FileCache_NormalizeKey = VBA.LCase$(VBA.Replace$(VBA.Trim$(cacheKey), "/", "\"))
End Function


Private Function private_FileCache_BuildFileTextKey(ByVal filePath As String) As String
    private_FileCache_BuildFileTextKey = "file.text|" & private_FileCache_NormalizeKey(filePath)
End Function


Private Function private_FileCache_TryGetFileStamp( _
    ByVal filePath As String, _
    ByRef outStamp As String, _
    ByVal showErrorUi As Boolean _
) As Boolean
    Dim fso As Object
    Dim fileObj As Object

    filePath = VBA.Trim$(filePath)
    outStamp = VBA.vbNullString
    If VBA.Len(filePath) = 0 Then Exit Function
    If VBA.Len(VBA.Dir(filePath)) = 0 Then Exit Function

    On Error GoTo EH_STAMP
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    Set fileObj = fso.GetFile(filePath)
    outStamp = VBA.CStr(VBA.CDbl(fileObj.DateLastModified))
    private_FileCache_TryGetFileStamp = True
    Exit Function

EH_STAMP:
    If showErrorUi Then ex_Core.fn_Diagnostic_LogError "FileCache: failed to read file modified date '" & filePath & "': " & Err.Description
End Function


Private Function private_FileCache_TryReadTextFile( _
    ByVal filePath As String, _
    ByRef outText As String, _
    ByVal showErrorUi As Boolean _
) As Boolean
    Dim f As Integer

    outText = VBA.vbNullString
    filePath = VBA.Trim$(filePath)
    If VBA.Len(filePath) = 0 Then Exit Function
    If VBA.Len(VBA.Dir(filePath)) = 0 Then Exit Function

    On Error GoTo EH_READ
    f = VBA.FreeFile
    Open filePath For Input As #f
    outText = VBA.Input$(VBA.LOF(f), f)
    Close #f
    private_FileCache_TryReadTextFile = True
    Exit Function

EH_READ:
    On Error Resume Next
    If f > 0 Then Close #f
    On Error GoTo 0
    If showErrorUi Then ex_Core.fn_Diagnostic_LogError "FileCache: failed to read file '" & filePath & "': " & Err.Description
End Function


Private Function private_FileCache_TryWriteTextFile( _
    ByVal filePath As String, _
    ByVal textValue As String, _
    ByVal showErrorUi As Boolean _
) As Boolean
    Dim f As Integer

    filePath = VBA.Trim$(filePath)
    If VBA.Len(filePath) = 0 Then Exit Function

    On Error GoTo EH_WRITE
    f = VBA.FreeFile
    Open filePath For Output As #f
    Print #f, textValue
    Close #f
    private_FileCache_TryWriteTextFile = True
    Exit Function

EH_WRITE:
    On Error Resume Next
    If f > 0 Then Close #f
    On Error GoTo 0
    If showErrorUi Then ex_Core.fn_Diagnostic_LogError "FileCache: failed to write file '" & filePath & "': " & Err.Description
End Function


Private Function private_FileCache_SetFileTextEntry( _
    ByVal filePath As String, _
    ByVal textValue As String, _
    ByVal fileStamp As String _
) As Boolean
    Dim cacheKey As String
    Dim entry As Object

    cacheKey = private_FileCache_BuildFileTextKey(filePath)
    If VBA.Len(cacheKey) = 0 Then Exit Function

    Set entry = VBA.CreateObject("Scripting.Dictionary")
    entry.CompareMode = 1
    entry("Text") = VBA.CStr(textValue)
    entry("Stamp") = VBA.CStr(fileStamp)

    private_FileCache_SetFileTextEntry = private_FileCache_Set(cacheKey, entry)
End Function


Private Function private_FileCache_SetFileText( _
    ByVal filePath As String, _
    ByVal textValue As String, _
    Optional ByVal showErrorUi As Boolean = True _
) As Boolean
    Dim fileStamp As String

    filePath = VBA.Trim$(filePath)
    If VBA.Len(filePath) = 0 Then Exit Function

    If Not private_FileCache_TryWriteTextFile(filePath, textValue, showErrorUi) Then Exit Function

    If Not private_FileCache_TryGetFileStamp(filePath, fileStamp, False) Then
        fileStamp = VBA.vbNullString
    End If
    private_FileCache_SetFileText = private_FileCache_SetFileTextEntry(filePath, textValue, fileStamp)
End Function


Private Sub private_FileCache_RemoveFileText(ByVal filePath As String)
    Dim cacheKey As String

    filePath = VBA.Trim$(filePath)
    If VBA.Len(filePath) = 0 Then Exit Sub
    cacheKey = private_FileCache_BuildFileTextKey(filePath)
    Call private_FileCache_Remove(cacheKey)
End Sub


Private Function private_FileCache_TryGet( _
    ByVal cacheKey As String, _
    ByRef outEntry As Object, _
    Optional ByVal allowMissing As Boolean = True _
) As Boolean
    cacheKey = private_FileCache_NormalizeKey(cacheKey)
    Set outEntry = Nothing

    If VBA.Len(cacheKey) = 0 Then
        private_FileCache_TryGet = allowMissing
        Exit Function
    End If

    private_FileCache_EnsureStorage
    If g_FileCacheMap.Exists(cacheKey) Then Set outEntry = g_FileCacheMap(cacheKey)
    private_FileCache_TryGet = True
End Function


Private Function private_FileCache_Set(ByVal cacheKey As String, ByVal entry As Object) As Boolean
    cacheKey = private_FileCache_NormalizeKey(cacheKey)
    If VBA.Len(cacheKey) = 0 Then Exit Function
    If entry Is Nothing Then Exit Function

    private_FileCache_EnsureStorage
    Set g_FileCacheMap(cacheKey) = entry
    private_FileCache_Set = True
End Function


Private Sub private_FileCache_Remove(ByVal cacheKey As String)
    cacheKey = private_FileCache_NormalizeKey(cacheKey)
    If VBA.Len(cacheKey) = 0 Then Exit Sub

    private_FileCache_EnsureStorage
    If g_FileCacheMap.Exists(cacheKey) Then g_FileCacheMap.Remove cacheKey
End Sub
' --------------------------------------
'  } // namespace FileCache
' --------------------------------------

' --------------------------------------
'  namespace RuntimeSource {
' --------------------------------------
Private Sub private_RuntimeSource_EnsureStorage()
    If g_GlobalItemsSourceMap Is Nothing Then
        Set g_GlobalItemsSourceMap = VBA.CreateObject("Scripting.Dictionary")
        g_GlobalItemsSourceMap.CompareMode = 1
    End If

    If g_GlobalObjectSourceMap Is Nothing Then
        Set g_GlobalObjectSourceMap = VBA.CreateObject("Scripting.Dictionary")
        g_GlobalObjectSourceMap.CompareMode = 1
    End If
End Sub


Private Sub private_RuntimeSource_ResetStorage()
    Dim sourceKey As Variant

    If Not g_GlobalItemsSourceMap Is Nothing Then
        For Each sourceKey In g_GlobalItemsSourceMap.Keys
            Set g_GlobalItemsSourceMap(sourceKey) = Nothing
        Next sourceKey

        g_GlobalItemsSourceMap.RemoveAll
        Set g_GlobalItemsSourceMap = Nothing
    End If

    If Not g_GlobalObjectSourceMap Is Nothing Then
        For Each sourceKey In g_GlobalObjectSourceMap.Keys
            Set g_GlobalObjectSourceMap(sourceKey) = Nothing
        Next sourceKey

        g_GlobalObjectSourceMap.RemoveAll
        Set g_GlobalObjectSourceMap = Nothing
    End If
End Sub


Private Function private_RuntimeSource_NormalizeKey(ByVal sourceKey As String) As String
    private_RuntimeSource_NormalizeKey = VBA.LCase$(VBA.Trim$(sourceKey))
End Function
' --------------------------------------
'  } // namespace RuntimeSource
' --------------------------------------

' --------------------------------------
'  namespace Dev {
' --------------------------------------
' Callstack[1]: ex_Core.fn_Dev_UpdateAllModules -> private_Dev_TryRunSafeUpdateByMode
' Callstack[2]: ex_Core.fn_Dev_UpdateCodeByDate -> private_Dev_TryRunSafeUpdateByMode
' Callstack[3]: ex_Core.fn_Dev_UpdateCodeBySize -> private_Dev_TryRunSafeUpdateByMode
' Главный сценарий безопасного обновления модулей:
' 1) убеждаемся, что runtime-компоненты доступны;
' 2) сохраняем runtime-состояние (snapshot);
' 3) освобождаем runtime-объекты и отменяем конфликтующие deferred-задачи;
' 4) выполняем remove/import модулей;
' 5) после успешного обновления запускаем отложенное восстановление состояния.
Private Function private_Dev_TryRunSafeUpdateByMode( _
    ByVal updateMode As Long, _
    ByVal includeComponentPattern As String, _
    ByVal excludeComponentPattern As String, _
    ByVal useNativeStatus As Boolean, _
    ByVal operationName As String _
) As Boolean
    Dim saveRuntimeOk As Boolean
    Dim updateOk As Boolean
    Dim safeErrorNumber As Long
    Dim safeErrorDescription As String
    Dim missingRuntimeComponents As String
    Dim presentRuntimeComponentsCount As Long
    Dim isCoreOnlyProject As Boolean

    On Error GoTo EH_SAFE_UPDATE

    ' Этап 0. Нормализуем имя операции для логов.
    operationName = VBA.LCase$(VBA.Trim$(operationName))
    If VBA.Len(operationName) = 0 Then operationName = "unknown"
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent "safe-update:start op='" & operationName & "'"
#End If

    ' Этап 1. ex_Core одновременно является автономным initial installer.
    ' Если runtime отсутствует целиком, Full Update импортирует его без dispose:
    ' освобождать в такой книге ещё нечего. Частично установленный runtime — уже
    ' ошибка, которую нельзя молча считать первичной установкой.
    If Not private_Dev_TryGetRuntimeContractState( _
        missingRuntimeComponents, presentRuntimeComponentsCount) Then
        isCoreOnlyProject = private_Dev_IsCoreOnlyProject()
        If isCoreOnlyProject And _
            VBA.StrComp(operationName, "full", VBA.vbTextCompare) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
            private_Diagnostic_LogCoreSelfEvent _
                "safe-update:initial-install-start op='full'"
#End If
            updateOk = private_Dev_UpdateCodeByRegex( _
                includeComponentPattern, excludeComponentPattern, _
                updateMode, useNativeStatus)
            If Not updateOk Then Exit Function

            missingRuntimeComponents = VBA.vbNullString
            presentRuntimeComponentsCount = 0
            If Not private_Dev_TryGetRuntimeContractState( _
                missingRuntimeComponents, presentRuntimeComponentsCount) Then
                VBA.MsgBox _
                    "Initial runtime installation is incomplete. Missing VBA components:" & _
                    VBA.vbCrLf & missingRuntimeComponents, _
                    VBA.vbExclamation, "PrototypeNew / Update Code"
                Exit Function
            End If

            private_Dev_QueueRuntimeStateRestoreAfterUpdate _
                "initial-install:full"
#If LOGGING_DEBUG_ENABLED Then
            private_Diagnostic_LogCoreSelfEvent _
                "safe-update:initial-install-done op='full'"
#End If
            private_Diagnostic_ClearCoreLogFile
            private_Dev_TryRunSafeUpdateByMode = True
            Exit Function
        End If

        private_Dev_ShowInvalidRuntimeContract _
            missingRuntimeComponents, isCoreOnlyProject
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "safe-update:fail op='" & operationName & "' reason='runtime-contract-invalid'"
#End If
        Exit Function
    End If

    ' Этап 2. Runtime валиден -> сохраняем runtime state перед целевым update.
#If RUNTIME_SNAPSHOTS_ENABLED Then
    If Not private_Dev_TryRunRuntimeBooleanFunction("rt_RestoreManager", "fn_SaveRuntimeState", saveRuntimeOk) Then
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "safe-update:fail op='" & operationName & "' reason='save-runtime-state-call-failed'"
#End If
        Exit Function
    End If
    If Not saveRuntimeOk Then
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "safe-update:fail op='" & operationName & "' reason='save-runtime-state-returned-false'"
#End If
        Exit Function
    End If
#End If

    ' Этап 3. Перед hot-import отменяем висящие deferred restore и освобождаем runtime-ссылки.
    If Not private_Dev_TryPrepareRuntimeForHotUpdate(operationName) Then
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "safe-update:fail op='" & operationName & "' reason='runtime-prepare-failed'"
#End If
        Exit Function
    End If

    ' Этап 4. Выполняем фактический импорт/обновление модулей.
    updateOk = private_Dev_UpdateCodeByRegex(includeComponentPattern, excludeComponentPattern, updateMode, useNativeStatus)
    If Not updateOk Then
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "safe-update:fail op='" & operationName & "' reason='update-import-failed'"
#End If
        Exit Function
    End If

    ' Этап 5. Не делаем синхронный restore прямо здесь.
    ' Восстановление целиком переносится в deferred-путь на следующий тик OnTime,
    ' чтобы избежать двойного рендера (sync restore + deferred restore) после hot-update.
    Call private_Dev_QueueRuntimeStateRestoreAfterUpdate("safe-update:" & operationName)
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent "safe-update:done op='" & operationName & "'"
#End If
    ' Full import is a clean diagnostic boundary. Clear only after the import
    ' succeeded and deferred runtime restoration was queued successfully.
    If VBA.StrComp(operationName, "full", VBA.vbTextCompare) = 0 Then
        private_Diagnostic_ClearCoreLogFile
    End If
    private_Dev_TryRunSafeUpdateByMode = True
    Exit Function

EH_SAFE_UPDATE:
    safeErrorNumber = Err.Number
    safeErrorDescription = Err.Description
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent _
        "safe-update:exception op='" & _
        VBA.Replace$(operationName, "'", "''") & "' err='" & _
        VBA.Replace$(safeErrorDescription, "'", "''") & "'"
#End If
    VBA.MsgBox _
        "Update Code stopped: [" & VBA.CStr(safeErrorNumber) & "] " & _
        safeErrorDescription, VBA.vbExclamation, _
        "PrototypeNew / Update Code"
End Function


Private Function private_Dev_TryGetRuntimeContractState( _
    ByRef outMissingComponents As String, _
    ByRef outPresentComponentsCount As Long _
) As Boolean
    outMissingComponents = VBA.vbNullString
    outPresentComponentsCount = 0

    private_Dev_AppendRuntimeComponentState _
        outMissingComponents, outPresentComponentsCount, "rt_Lifecycle"
    private_Dev_AppendRuntimeComponentState _
        outMissingComponents, outPresentComponentsCount, "rt_CoreActions"
    private_Dev_AppendRuntimeComponentState _
        outMissingComponents, outPresentComponentsCount, "rt_RestoreManager"
    private_Dev_AppendRuntimeComponentState _
        outMissingComponents, outPresentComponentsCount, "rt_PageManager"
    private_Dev_AppendRuntimeComponentState _
        outMissingComponents, outPresentComponentsCount, "ex_HelpersSheet"
    private_Dev_AppendRuntimeComponentState _
        outMissingComponents, outPresentComponentsCount, "obj_PageBase"
    private_Dev_AppendRuntimeComponentState _
        outMissingComponents, outPresentComponentsCount, "obj_IPage"
    private_Dev_AppendRuntimeComponentState _
        outMissingComponents, outPresentComponentsCount, "obj_ISerializable"

    private_Dev_TryGetRuntimeContractState = _
        (VBA.Len(outMissingComponents) = 0)
End Function


Private Function private_Dev_IsCoreOnlyProject() As Boolean
    Dim vbComponent As Object
    Dim componentName As String

    ' Документные модули ThisWorkbook/Worksheet всегда существуют и не
    ' считаются установленным runtime. Любой другой компонент кроме ex_Core
    ' означает частичную либо стороннюю сборку, которую Full Update не чинит
    ' автоматически как initial install.
    For Each vbComponent In ThisWorkbook.VBProject.VBComponents
        If VBA.CLng(vbComponent.Type) <> VB_COMPONENT_TYPE_DOCUMENT Then
            componentName = VBA.Trim$(VBA.CStr(vbComponent.Name))
            If VBA.StrComp( _
                componentName, CORE_COMPONENT_NAME, _
                VBA.vbTextCompare) <> 0 Then Exit Function
        End If
    Next vbComponent

    private_Dev_IsCoreOnlyProject = True
End Function


Private Sub private_Dev_AppendRuntimeComponentState( _
    ByRef missingComponents As String, _
    ByRef presentComponentsCount As Long, _
    ByVal componentName As String _
)
    Dim runtimeComponent As Object

    Set runtimeComponent = private_Dev_TryGetComponentByName(componentName)
    If Not runtimeComponent Is Nothing Then
        presentComponentsCount = presentComponentsCount + 1
        Exit Sub
    End If
    missingComponents = missingComponents & "- " & componentName & VBA.vbCrLf
End Sub


Private Sub private_Dev_ShowInvalidRuntimeContract( _
    ByVal missingComponents As String, _
    ByVal isCoreOnlyProject As Boolean _
)
    Dim guidanceText As String

    If isCoreOnlyProject Then
        guidanceText = _
            "Run Full Update to perform the initial runtime installation."
    Else
        guidanceText = _
            "The VBA project contains a partial runtime installation. " & _
            "Full Update was not started automatically."
    End If

    VBA.MsgBox _
        "Update Code cannot continue. Missing required VBA components:" & _
        VBA.vbCrLf & missingComponents & VBA.vbCrLf & guidanceText, _
        VBA.vbExclamation, "PrototypeNew / Update Code"
End Sub


' Callstack[1]: ex_Core.private_Dev_TryRunSafeUpdateByMode -> private_Dev_QueueRuntimeStateRestoreAfterUpdate
Private Sub private_Dev_QueueRuntimeStateRestoreAfterUpdate(ByVal reasonText As String)
    Dim macroRef As String
    Dim scheduleAt As Date
    Dim errDescription As String
    Dim queueStage As String

    reasonText = VBA.Trim$(reasonText)
    If VBA.Len(reasonText) = 0 Then reasonText = "unknown"

#If RUNTIME_SNAPSHOTS_ENABLED Then
    macroRef = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!rt_RestoreManager.fn_RunDeferredRuntimeStateRestore"
#Else
    ' Hot-update освобождает rt_PageManager и тем самым удаляет runtime routes.
    ' Без snapshots на следующем тике создаём чистую Main, иначе на листе
    ' останутся только Shape с OnAction, но без соответствующего page instance.
    macroRef = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!rt_CoreActions.fn_RerenderLastPageAfterUpdate"
#End If
    scheduleAt = private_Dev_GetNextOnTimeTick()

    On Error GoTo EH_QUEUE
    queueStage = "cancel-existing"
    ' Держим только последнюю задачу deferred restore.
    If g_QueuedRuntimeStateRestoreAt > 0# And VBA.Len(VBA.Trim$(g_QueuedRuntimeStateRestoreMacro)) > 0 Then
        Application.OnTime EarliestTime:=g_QueuedRuntimeStateRestoreAt, Procedure:=g_QueuedRuntimeStateRestoreMacro, Schedule:=False
    End If
    g_QueuedRuntimeStateRestoreAt = 0#
    g_QueuedRuntimeStateRestoreMacro = VBA.vbNullString

    queueStage = "schedule"
    Application.OnTime EarliestTime:=scheduleAt, Procedure:=macroRef
    g_QueuedRuntimeStateRestoreAt = scheduleAt
    g_QueuedRuntimeStateRestoreMacro = macroRef
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent "runtime-state-restore-queued reason='" & VBA.Replace$(reasonText, "'", "''") & "'"
#End If
    Exit Sub

EH_QUEUE:
    errDescription = Err.Description
    If VBA.StrComp(queueStage, "schedule", VBA.vbBinaryCompare) = 0 Then
        g_QueuedRuntimeStateRestoreAt = 0#
        g_QueuedRuntimeStateRestoreMacro = VBA.vbNullString
    End If
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent "runtime-state-restore-queue-failed reason='" & VBA.Replace$(reasonText, "'", "''") & "' err='" & VBA.Replace$(errDescription, "'", "''") & "'"
#End If
    Err.Raise VBA.vbObjectError + 9213, _
        "ex_Core.private_Dev_QueueRuntimeStateRestoreAfterUpdate", _
        "Failed to schedule runtime restoration: " & _
        errDescription
End Sub


Public Sub fn_Dev_MarkRuntimeStateRestoreStarted()
    ' Excel удаляет callback из очереди непосредственно перед вызовом макроса.
    ' Обнуляем локальную регистрацию, чтобы lifecycle не пытался отменить уже
    ' исполняемую задачу.
    g_QueuedRuntimeStateRestoreAt = 0#
    g_QueuedRuntimeStateRestoreMacro = VBA.vbNullString
End Sub


Private Function private_Dev_TryPrepareRuntimeForHotUpdate(ByVal operationName As String) As Boolean
    Dim macroRef As String

    operationName = VBA.Trim$(operationName)
    If VBA.Len(operationName) = 0 Then operationName = "unknown"

    On Error GoTo EH_PREPARE
    ' ex_Core должен компилироваться и работать в книге, где он является
    ' единственным стандартным модулем. Поэтому dependency на lifecycle здесь
    ' строковая и строго адресует один известный public API, а не ищет disposer-ы.
    macroRef = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & _
        "'!rt_Lifecycle.fn_DisposeRuntime"
    Application.Run macroRef, False, _
        "safe-update:prepare:" & operationName

#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent "runtime-update-prepare:done op='" & VBA.Replace$(operationName, "'", "''") & "'"
#End If
    private_Dev_TryPrepareRuntimeForHotUpdate = True
    Exit Function

EH_PREPARE:
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent _
        "runtime-update-prepare:failed op='" & _
        VBA.Replace$(operationName, "'", "''") & "' err='" & _
        VBA.Replace$(Err.Description, "'", "''") & "'"
#End If
End Function


Private Sub private_Dev_CancelQueuedRuntimeStateRestore(Optional ByVal reasonText As String = VBA.vbNullString)
    Dim errNumber As Long
    Dim errDescription As String

    reasonText = VBA.Trim$(reasonText)
    If VBA.Len(reasonText) = 0 Then reasonText = "unknown"

    If g_QueuedRuntimeStateRestoreAt > 0# And VBA.Len(VBA.Trim$(g_QueuedRuntimeStateRestoreMacro)) > 0 Then
        On Error GoTo EH_CANCEL
        Application.OnTime EarliestTime:=g_QueuedRuntimeStateRestoreAt, Procedure:=g_QueuedRuntimeStateRestoreMacro, Schedule:=False
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "runtime-state-restore-cancelled reason='" & VBA.Replace$(reasonText, "'", "''") & "'"
#End If
    End If

    g_QueuedRuntimeStateRestoreAt = 0#
    g_QueuedRuntimeStateRestoreMacro = VBA.vbNullString
    Exit Sub

EH_CANCEL:
    errNumber = Err.Number
    errDescription = Err.Description
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent "runtime-state-restore-cancel-failed reason='" & VBA.Replace$(reasonText, "'", "''") & "' err='" & VBA.Replace$(errDescription, "'", "''") & "'"
#End If
    Err.Raise errNumber, _
        "ex_Core.private_Dev_CancelQueuedRuntimeStateRestore", _
        "Failed to cancel deferred runtime restoration: " & _
        errDescription
End Sub


Private Function private_Dev_UpdateCodeByRegex( _
    ByVal includeComponentPattern As String, _
    Optional ByVal excludeComponentPattern As String = VBA.vbNullString, _
    Optional ByVal updateMode As Long = UPDATE_MODE_FULL, _
    Optional ByVal useNativeStatus As Boolean = False _
) As Boolean
    private_Dev_UpdateCodeByRegex = private_Dev_UpdateCodeCore(updateMode, useNativeStatus, includeComponentPattern, excludeComponentPattern)
End Function


Private Function private_Dev_UpdateCodeCore( _
    ByVal updateMode As Long, _
    ByVal useNativeStatus As Boolean, _
    Optional ByVal includeComponentPattern As String = VBA.vbNullString, _
    Optional ByVal excludeComponentPattern As String = VBA.vbNullString _
) As Boolean
    Dim basePath As String
    Dim cachePath As String
    Dim prevCache As Object
    Dim nextCache As Object
    Dim stageName As String
    Dim updatedComponents As Object
    Dim updatedComponentsCount As Long
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim fullErrorText As String
    Dim previousScreenUpdating As Boolean
    Dim screenUpdatingCaptured As Boolean

    ' Низкоуровневый движок импорта: обновляет компоненты in-place, удаляет
    ' только stale-компоненты, обслуживает кеш и выполняет валидацию.
    ' Любая ошибка импорта является финальным результатом текущей операции.
    ' stageName нужен для точной диагностики места падения.
    stageName = "init"
    private_ShowStatusNotice "Code update started...", useNativeStatus, 1
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent "update-start"
#End If

    basePath = ThisWorkbook.Path & "\\" & BASE_DIR
    If VBA.Len(VBA.Dir(basePath, vbDirectory)) = 0 Then
        private_ShowStatusWarning "Workbook path is empty or 'vba' folder was not found. Save the workbook first.", useNativeStatus, 6
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "update-stop: vba-folder-not-found"
#End If
        private_Dev_UpdateCodeCore = False
        Exit Function
    End If

    On Error GoTo EH
    previousScreenUpdating = Application.ScreenUpdating
    screenUpdatingCaptured = True
    Application.ScreenUpdating = False

    stageName = "load-cache"
    cachePath = basePath & IMPORT_CACHE_FILE
    Set prevCache = private_Dev_LoadImportCache(cachePath)
    Set nextCache = private_Dev_CreateDictionary()
    Set updatedComponents = private_Dev_CreateDictionary()

    stageName = "import-folder"
    private_Dev_ImportFolder basePath, updateMode, prevCache, nextCache, includeComponentPattern, excludeComponentPattern, updatedComponents
    If ENABLE_CLASS_IMPORT_VALIDATION Then
        stageName = "validate-class-imports"
        private_Dev_ValidateClassImports basePath
    End If

    ' Существующие standard/class компоненты обновляются in-place. Удаляем
    ' только файлы, которые действительно исчезли из исходного дерева.
    ' Массовый Remove/Add при full update менял COM/type identity классов и мог
    ' повреждать compiled state сохраняемого VBA-проекта.
    stageName = "remove-stale"
    private_Dev_RemoveStaleImportedComponentsByScope prevCache, nextCache, includeComponentPattern, excludeComponentPattern
    stageName = "preserve-out-of-scope-cache"
    private_Dev_PreserveOutOfScopeCacheRecords prevCache, nextCache, includeComponentPattern, excludeComponentPattern
    stageName = "save-cache"
    private_Dev_SaveImportCache cachePath, nextCache

    If screenUpdatingCaptured Then
        Application.ScreenUpdating = previousScreenUpdating
    End If
    updatedComponentsCount = private_Dev_GetDictionaryCount(updatedComponents)
    private_Dev_LogUpdatedComponents updatedComponents, updateMode
    private_Dev_ShowCodeUpdatedNotice useNativeStatus, updatedComponentsCount
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent "update-done"
#End If
    private_Dev_UpdateCodeCore = True
    Exit Function

EH:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    ' Сохраняем причину в module-state для полной диагностики ошибки импорта.
    fullErrorText = "Code update failed at stage '" & stageName & "': [" & errSource & " #" & VBA.CStr(errNumber) & "] " & errDescription

    If screenUpdatingCaptured Then
        Application.ScreenUpdating = previousScreenUpdating
    End If
    private_ShowStatusError fullErrorText, useNativeStatus, 6

    ' Статус-бар часто обрезает длинный текст ошибки импорта.
    ' Пишем полную диагностику (включая список файлов) в лог.
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError fullErrorText
#End If

    ' Логируем ошибку напрямую в core.log, даже если CORE_ENABLE_SELF_LOGGING = False.
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreEvent "update-fail: stage='" & stageName & "' err='" & VBA.Replace$(errDescription, "'", "''") & "'"
#End If
    private_Dev_UpdateCodeCore = False
End Function


Private Sub private_Dev_PreserveOutOfScopeCacheRecords( _
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
    If private_Dev_ShouldProcessComponentByScope(componentName, includeComponentPattern, excludeComponentPattern) Then GoTo ContinueKey

        nextCache.Add VBA.CStr(key), rec
ContinueKey:
    Next key
End Sub


Private Sub private_Dev_ValidateClassImports(ByVal rootPath As String)
    Dim fso As Object
    Dim failed As String

    If VBA.Dir(rootPath, vbDirectory) = "" Then
        Err.Raise VBA.vbObjectError + 1006, "private_Dev_ValidateClassImports", "VBA root folder not found: " & rootPath
    End If

    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    private_Dev_ValidateClassImportsRecursive fso.GetFolder(rootPath), 0, failed

    If VBA.Len(failed) > 0 Then
        Err.Raise VBA.vbObjectError + 1007, "private_Dev_ValidateClassImports", "Class import validation failed:" & failed
    End If
End Sub


Private Sub private_Dev_ValidateClassImportsRecursive( _
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
        If Not private_Dev_TryResolveFileComponentType(VBA.CStr(fileObj.Name), compType, fallbackName) Then GoTo ContinueFile
        If VBA.StrComp(compType, COMP_TYPE_CLASS, VBA.vbTextCompare) <> 0 Then GoTo ContinueFile

        className = private_Dev_GetComponentNameFromSource(VBA.CStr(fileObj.Path))
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
        private_Dev_ValidateClassImportsRecursive subFolder, depth + 1, failed
    Next subFolder
End Sub


Private Sub private_Dev_ShowCodeUpdatedNotice(ByVal useNativeStatus As Boolean, ByVal updatedComponentsCount As Long)
    If updatedComponentsCount < 0 Then updatedComponentsCount = 0
    private_ShowStatusSuccess "Code updated. Updated modules: " & VBA.CStr(updatedComponentsCount) & ".", useNativeStatus, 1
End Sub

Private Function private_Dev_GetDictionaryCount(ByVal dict As Object) As Long
    If dict Is Nothing Then Exit Function
    On Error Resume Next
    private_Dev_GetDictionaryCount = VBA.CLng(dict.Count)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub private_Dev_RegisterUpdatedComponent(ByVal updatedComponents As Object, ByVal componentName As String)
    If updatedComponents Is Nothing Then Exit Sub

    componentName = VBA.Trim$(componentName)
    If VBA.Len(componentName) = 0 Then Exit Sub
    If updatedComponents.Exists(componentName) Then Exit Sub

    updatedComponents.Add componentName, True
End Sub

Private Sub private_Dev_LogUpdatedComponents(ByVal updatedComponents As Object, ByVal updateMode As Long)
    Dim componentName As Variant
    Dim updatedCount As Long
    Dim modeName As String

    updatedCount = private_Dev_GetDictionaryCount(updatedComponents)
    modeName = private_Dev_GetUpdateModeName(updateMode)

#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent "update-modules: mode='" & modeName & "' count=" & VBA.CStr(updatedCount)
    If updatedComponents Is Nothing Then Exit Sub
    For Each componentName In updatedComponents.Keys
        private_Diagnostic_LogCoreSelfEvent "update-module: mode='" & modeName & "' name='" & VBA.Replace$(VBA.CStr(componentName), "'", "''") & "'"
    Next componentName
#End If
End Sub

Private Function private_Dev_GetUpdateModeName(ByVal updateMode As Long) As String
    Select Case updateMode
        Case UPDATE_MODE_FULL
            private_Dev_GetUpdateModeName = "full"
        Case UPDATE_MODE_DATE
            private_Dev_GetUpdateModeName = "date"
        Case UPDATE_MODE_SIZE
            private_Dev_GetUpdateModeName = "size"
        Case Else
            private_Dev_GetUpdateModeName = "unknown"
    End Select
End Function


Private Sub private_Dev_RemoveAllModulesAndClasses( _
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
                If private_Dev_ShouldProcessComponentByScope(VBA.CStr(comp.Name), includeComponentPattern, excludeComponentPattern) Then
                    n = n + 1
                    ReDim Preserve names(1 To n)
                    names(n) = comp.Name
                End If

            Case 100 ' модуль документа (книга/листы)
                If private_Dev_ShouldProcessComponentByScope(VBA.CStr(comp.Name), includeComponentPattern, excludeComponentPattern) Then
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
        private_Dev_ClearDocumentModuleCode prj.VBComponents(docNames(i))
        On Error GoTo 0
    Next i

    Exit Sub

EH_REMOVE:
    Err.Raise VBA.vbObjectError + 1008, "private_Dev_RemoveAllModulesAndClasses", _
              "Failed to remove component '" & names(i) & "': " & Err.Description

EH_CLEAR_DOC:
    Err.Raise VBA.vbObjectError + 1011, "private_Dev_RemoveAllModulesAndClasses", _
              "Failed to clear document module '" & docNames(i) & "': " & Err.Description
End Sub


Private Sub private_Dev_ImportFolder( _
    ByVal folderPath As String, _
    ByVal updateMode As Long, _
    ByVal prevCache As Object, _
    ByVal nextCache As Object, _
    Optional ByVal includeComponentPattern As String = VBA.vbNullString, _
    Optional ByVal excludeComponentPattern As String = VBA.vbNullString, _
    Optional ByVal updatedComponents As Object = Nothing _
)
    Dim fso As Object
    Dim rootFolder As Object
    Dim failed As String
    Dim importPass As Long

    If VBA.Dir(folderPath, vbDirectory) = "" Then Exit Sub

    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    Set rootFolder = fso.GetFolder(folderPath)

    ' Глобальные два прохода по всему дереву:
    ' 1) сначала все компоненты кроме интерфейсов obj_I*;
    ' 2) затем интерфейсы.
    ' Важно: при импорте VBA сразу проверяет/компилирует сигнатуры членов класса.
    ' Поэтому если интерфейс ссылается на тип (например As obj_PageBase), а этот
    ' тип еще не импортирован, импорт падает с "User-defined type not defined".
    ' Так интерфейсы всегда импортируются после потенциально зависимых классов
    ' даже если они лежат в разных подпапках.
    For importPass = 1 To 2
        private_Dev_ImportFolderRecursive rootFolder, 0, failed, updateMode, prevCache, nextCache, includeComponentPattern, excludeComponentPattern, importPass, updatedComponents
    Next importPass

    If VBA.Len(failed) > 0 Then
        Err.Raise VBA.vbObjectError + 1001, "private_Dev_ImportFolder", "Import failed for file(s):" & failed
    End If
End Sub


Private Sub private_Dev_ImportFolderRecursive( _
    ByVal folderObj As Object, _
    ByVal depth As Long, _
    ByRef failed As String, _
    ByVal updateMode As Long, _
    ByVal prevCache As Object, _
    ByVal nextCache As Object, _
    Optional ByVal includeComponentPattern As String = VBA.vbNullString, _
    Optional ByVal excludeComponentPattern As String = VBA.vbNullString, _
    Optional ByVal importPass As Long = 1, _
    Optional ByVal updatedComponents As Object = Nothing _
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
    Dim incrementalMode As Boolean
    Dim shouldProcess As Boolean
    If folderObj Is Nothing Then Exit Sub
    If depth > MAX_IMPORT_RECURSION_DEPTH Then Exit Sub

    ' Проходим файлы "мягко": ошибка отдельного файла не роняет обход мгновенно,
    ' а копится в списке failed. Это позволяет получить полную картину проблем за pass
    ' и вывести полную диагностику после завершения прохода.
    incrementalMode = (updateMode <> UPDATE_MODE_FULL)

    For Each fileObj In folderObj.Files
        fileName = VBA.CStr(fileObj.Name)
        If Not private_Dev_ShouldImportFileInPass(fileName, importPass) Then GoTo ContinueNextFile

        If private_Dev_TryResolveFileComponentType(fileName, compType, fallbackName) Then
            importPath = VBA.CStr(fileObj.Path)
            On Error GoTo EH_IMPORT_FILE

            fileDateStamp = private_Dev_BuildFileDateStampFromFileObject(fileObj)
            fileSizeStamp = private_Dev_BuildFileSizeStampFromFileObject(fileObj)
            cacheKey = private_Dev_NormalizeCacheKey(importPath)
            sourceText = VBA.vbNullString

            shouldProcess = private_Dev_ShouldProcessComponentByScope(fallbackName, includeComponentPattern, excludeComponentPattern)
            If Not shouldProcess Then GoTo ContinueNextFile

            Select Case VBA.LCase$(compType)
                Case COMP_TYPE_MODULE, COMP_TYPE_CLASS
                    If incrementalMode Then
                        If private_Dev_TryGetCachedComponentNameByMode(prevCache, cacheKey, compType, fileDateStamp, fileSizeStamp, updateMode, componentName) Then
                            If private_Dev_IsComponentPresentForType(componentName, compType) Then
                                private_Dev_SetCacheRecord nextCache, cacheKey, compType, componentName, fileDateStamp, fileSizeStamp
                                GoTo ContinueNextFile
                            End If
                        End If
                    End If

                    sourceText = private_Dev_ReadAllText(importPath)
                    componentName = private_Dev_GetComponentNameFromSourceText(sourceText, fallbackName)
                    private_Dev_EnsureValidComponentNameLength componentName, importPath

                    If VBA.StrComp(compType, COMP_TYPE_MODULE, VBA.vbTextCompare) = 0 Then
                        private_Dev_ImportStandardModuleFromSource componentName, importPath, sourceText
                    Else
                        private_Dev_ImportClassModuleFromSource componentName, importPath, sourceText
                    End If
                    private_Dev_SetCacheRecord nextCache, cacheKey, compType, componentName, fileDateStamp, fileSizeStamp
                    private_Dev_RegisterUpdatedComponent updatedComponents, componentName

                Case COMP_TYPE_SHEET
                    componentNameForCache = private_Dev_ResolveSheetCodeName(fallbackName)
                    If VBA.Len(componentNameForCache) = 0 Then GoTo ContinueNextFile

                    If incrementalMode Then
                        If private_Dev_IsCacheRecordCurrentByMode(prevCache, cacheKey, COMP_TYPE_SHEET, componentNameForCache, fileDateStamp, fileSizeStamp, updateMode) Then
                            If private_Dev_IsComponentPresentForType(componentNameForCache, COMP_TYPE_SHEET) Then
                                private_Dev_SetCacheRecord nextCache, cacheKey, COMP_TYPE_SHEET, componentNameForCache, fileDateStamp, fileSizeStamp
                                GoTo ContinueNextFile
                            End If
                        End If
                    End If

                    sourceText = private_Dev_ReadAllText(importPath)
                    If private_Dev_UpdateSheetModule(componentNameForCache, importPath, sourceText) Then
                        private_Dev_SetCacheRecord nextCache, cacheKey, COMP_TYPE_SHEET, componentNameForCache, fileDateStamp, fileSizeStamp
                        private_Dev_RegisterUpdatedComponent updatedComponents, componentNameForCache
                    End If

                Case COMP_TYPE_WORKBOOK
                    componentNameForCache = private_Dev_FindWorkbookComponentName()
                    If VBA.Len(componentNameForCache) = 0 Then GoTo ContinueNextFile

                    If incrementalMode Then
                        If private_Dev_IsCacheRecordCurrentByMode(prevCache, cacheKey, COMP_TYPE_WORKBOOK, componentNameForCache, fileDateStamp, fileSizeStamp, updateMode) Then
                            If private_Dev_IsComponentPresentForType(componentNameForCache, COMP_TYPE_WORKBOOK) Then
                                private_Dev_SetCacheRecord nextCache, cacheKey, COMP_TYPE_WORKBOOK, componentNameForCache, fileDateStamp, fileSizeStamp
                                GoTo ContinueNextFile
                            End If
                        End If
                    End If

                    sourceText = private_Dev_ReadAllText(importPath)
                    If private_Dev_UpdateWorkbookModuleFromText(componentNameForCache, sourceText) Then
                        private_Dev_SetCacheRecord nextCache, cacheKey, COMP_TYPE_WORKBOOK, componentNameForCache, fileDateStamp, fileSizeStamp
                        private_Dev_RegisterUpdatedComponent updatedComponents, componentNameForCache
                    End If
            End Select
            On Error GoTo 0
        End If

ContinueNextFile:
    ' Важно сбрасывать per-file handler на каждой итерации,
    ' иначе ошибка из одного файла может "залипнуть" и исказить картину следующего.
        On Error GoTo 0
    Next fileObj

    On Error GoTo 0
    For Each subFolder In folderObj.SubFolders
        private_Dev_ImportFolderRecursive subFolder, depth + 1, failed, updateMode, prevCache, nextCache, includeComponentPattern, excludeComponentPattern, importPass, updatedComponents
    Next subFolder

    Exit Sub

EH_IMPORT_FILE:
    errText = VBA.CStr(Err.Number) & ": " & Err.Description
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent "update-import-file-failed: path='" & VBA.Replace$(importPath, "'", "''") & "' err='" & VBA.Replace$(errText, "'", "''") & "'"
#End If
    failed = failed & VBA.vbCrLf & "- " & importPath & " (" & errText & ")"
    ' Продолжаем обход, чтобы собрать все проблемные файлы в одном отчете.
    ' Финальная агрегированная ошибка поднимается после завершения прохода.
    Err.Clear
    On Error GoTo 0
    GoTo ContinueNextFile
End Sub


Private Function private_Dev_ShouldImportFileInPass(ByVal fileName As String, ByVal importPass As Long) As Boolean
    Dim normalizedName As String
    Dim baseName As String
    Dim componentStem As String
    Dim markerChar As String
    Dim isInterfaceClass As Boolean

    normalizedName = VBA.LCase$(VBA.Trim$(fileName))
    If fn_Helpers_EndsWith(normalizedName, ".utf8.vba") Then
        baseName = VBA.Left$(fileName, VBA.Len(fileName) - VBA.Len(".utf8.vba"))
    ElseIf fn_Helpers_EndsWith(normalizedName, ".vba") Then
        baseName = VBA.Left$(fileName, VBA.Len(fileName) - VBA.Len(".vba"))
    Else
        private_Dev_ShouldImportFileInPass = False
        Exit Function
    End If

    normalizedName = VBA.LCase$(VBA.Trim$(baseName))
    isInterfaceClass = False
    If fn_Helpers_EndsWith(normalizedName, ".cls") Then
        componentStem = VBA.Left$(baseName, VBA.Len(baseName) - VBA.Len(".cls"))
        If VBA.Left$(componentStem, 5) = "obj_I" Then
            markerChar = VBA.Mid$(componentStem, 6, 1)
            If markerChar >= "A" And markerChar <= "Z" Then
                isInterfaceClass = True
            End If
        End If
    End If

    If importPass <= 1 Then
        private_Dev_ShouldImportFileInPass = Not isInterfaceClass
    Else
        private_Dev_ShouldImportFileInPass = isInterfaceClass
    End If
End Function


Private Sub private_Dev_EnsureValidComponentNameLength(ByVal componentName As String, ByVal importPath As String)
    If VBA.Len(componentName) <= MAX_VBA_COMPONENT_NAME_LEN Then Exit Sub
    Err.Raise VBA.vbObjectError + 1010, "private_Dev_EnsureValidComponentNameLength", _
              "VBA component name '" & componentName & "' is too long (" & VBA.CStr(VBA.Len(componentName)) & _
              "). Maximum allowed is " & VBA.CStr(MAX_VBA_COMPONENT_NAME_LEN) & ". File: " & importPath
End Sub


Private Sub private_Dev_ImportStandardModuleFromSource( _
    ByVal componentName As String, _
    ByVal importPath As String, _
    Optional ByVal preloadedSourceText As String = VBA.vbNullString _
)
    Dim vbComp As Object
    Dim cm As Object
    Dim sourceText As String
    Dim cleanCode As String

    If VBA.Len(VBA.Trim$(componentName)) = 0 Then
        Err.Raise VBA.vbObjectError + 1009, "private_Dev_ImportStandardModuleFromSource", "Standard module name is empty for: " & importPath
    End If

    If VBA.Len(preloadedSourceText) > 0 Then
        sourceText = preloadedSourceText
    Else
        sourceText = private_Dev_ReadAllText(importPath)
    End If
    cleanCode = private_Dev_ExtractCodeBody(sourceText)

    Set vbComp = private_Dev_TryGetComponentByName(componentName)
    If vbComp Is Nothing Then
        Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(1) ' vbext_ct_StdModule
        vbComp.Name = componentName
    ElseIf vbComp.Type <> 1 Then
        Err.Raise VBA.vbObjectError + 1016, _
            "private_Dev_ImportStandardModuleFromSource", _
            "Existing component '" & componentName & _
            "' is not a standard module."
    End If
    Set cm = vbComp.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString cleanCode
End Sub


Private Sub private_Dev_ImportClassModuleFromSource( _
    ByVal componentName As String, _
    ByVal importPath As String, _
    Optional ByVal preloadedSourceText As String = VBA.vbNullString _
)
    Dim vbComp As Object
    Dim cm As Object
    Dim sourceText As String
    Dim cleanCode As String

    If VBA.Len(VBA.Trim$(componentName)) = 0 Then
        Err.Raise VBA.vbObjectError + 1005, "private_Dev_ImportClassModuleFromSource", "Class module name is empty for: " & importPath
    End If

    If VBA.Len(preloadedSourceText) > 0 Then
        sourceText = preloadedSourceText
    Else
        sourceText = private_Dev_ReadAllText(importPath)
    End If
    cleanCode = private_Dev_ExtractCodeBody(sourceText)

    Set vbComp = private_Dev_TryGetComponentByName(componentName)
    If vbComp Is Nothing Then
        Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(2) ' vbext_ct_ClassModule
        vbComp.Name = componentName
    ElseIf vbComp.Type <> 2 Then
        Err.Raise VBA.vbObjectError + 1017, _
            "private_Dev_ImportClassModuleFromSource", _
            "Existing component '" & componentName & _
            "' is not a class module."
    End If
    Set cm = vbComp.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString cleanCode
End Sub


Private Function private_Dev_ExtractCodeBody(ByVal sourceText As String) As String
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
        lineText = VBA.Replace(lineText, VBA.ChrW$(65279), VBA.vbNullString)
        lineText = VBA.Replace(lineText, VBA.ChrW$(160), " ")
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

    private_Dev_ExtractCodeBody = outText
End Function


Private Sub private_Dev_RemoveComponentIfExists(ByVal componentName As String)
    Dim vbComp As Object

    If VBA.Len(componentName) = 0 Then Exit Sub

    Set vbComp = private_Dev_TryGetComponentByName(componentName)
    If vbComp Is Nothing Then Exit Sub

    ThisWorkbook.VBProject.VBComponents.Remove vbComp

    ' DoEvents внутри VBIDE mutation создаёт окно reentrancy, в котором можно
    ' закрыть книгу с незавершённым import call stack. Если VBE ещё удерживает
    ' stale-компонент, безопасный update сам повторится через OnTime.
    Set vbComp = private_Dev_TryGetComponentByName(componentName)
    If vbComp Is Nothing Then Exit Sub

    Err.Raise VBA.vbObjectError + 1015, "private_Dev_RemoveComponentIfExists", _
              "Component '" & componentName & "' is still present after remove operation."
End Sub


Private Function private_Dev_TryGetComponentByName(ByVal componentName As String) As Object
    On Error Resume Next
    Set private_Dev_TryGetComponentByName = ThisWorkbook.VBProject.VBComponents(componentName)
    On Error GoTo 0
End Function


Private Function private_Dev_IsComponentPresentForType(ByVal componentName As String, ByVal compType As String) As Boolean
    Dim vbComp As Object

    Set vbComp = private_Dev_TryGetComponentByName(componentName)
    If vbComp Is Nothing Then Exit Function

    Select Case VBA.LCase$(compType)
        Case COMP_TYPE_MODULE
            private_Dev_IsComponentPresentForType = (vbComp.Type = 1) ' стандартный модуль
        Case COMP_TYPE_CLASS
            private_Dev_IsComponentPresentForType = (vbComp.Type = 2) ' модуль класса
        Case COMP_TYPE_SHEET, COMP_TYPE_WORKBOOK
            private_Dev_IsComponentPresentForType = (vbComp.Type = 100) ' модуль документа
    End Select
End Function


Private Function private_Dev_GetComponentNameFromSource(ByVal importPath As String) As String
    Dim fileName As String
    Dim dotPos As Long
    Dim compType As String
    Dim sourceText As String
    Dim fallbackName As String

    fileName = VBA.Mid$(importPath, VBA.InStrRev(importPath, "\") + 1)
    If Not private_Dev_TryResolveFileComponentType(fileName, compType, fallbackName) Then
        dotPos = VBA.InStrRev(fileName, ".")
        If dotPos > 1 Then
            fallbackName = VBA.Left$(fileName, dotPos - 1)
        Else
            fallbackName = fileName
        End If
    End If

    sourceText = private_Dev_ReadAllText(importPath)
    private_Dev_GetComponentNameFromSource = private_Dev_GetComponentNameFromSourceText(sourceText, fallbackName)
End Function


Private Function private_Dev_GetComponentNameFromSourceText(ByVal sourceText As String, ByVal fallbackName As String) As String
    Dim attrPos As Long
    Dim quoteStart As Long
    Dim quoteEnd As Long

    private_Dev_GetComponentNameFromSourceText = fallbackName

    attrPos = VBA.InStr(1, sourceText, "Attribute VB_Name", VBA.vbTextCompare)
    If attrPos = 0 Then Exit Function

    quoteStart = VBA.InStr(attrPos, sourceText, """")
    If quoteStart = 0 Then Exit Function

    quoteEnd = VBA.InStr(quoteStart + 1, sourceText, """")
    If quoteEnd <= quoteStart Then Exit Function

    private_Dev_GetComponentNameFromSourceText = VBA.Mid$(sourceText, quoteStart + 1, quoteEnd - quoteStart - 1)
End Function


Private Function private_Dev_ShouldProcessComponentByScope( _
    ByVal componentName As String, _
    Optional ByVal includeComponentPattern As String = VBA.vbNullString, _
    Optional ByVal excludeComponentPattern As String = VBA.vbNullString _
) As Boolean
    private_Dev_ShouldProcessComponentByScope = private_Dev_MatchesIncludeExclude(componentName, includeComponentPattern, excludeComponentPattern)
End Function


Private Function private_Dev_MatchesIncludeExclude( _
    ByVal valueText As String, _
    ByVal includePattern As String, _
    ByVal excludePattern As String _
) As Boolean
    valueText = VBA.Trim$(VBA.CStr(valueText))
    includePattern = VBA.Trim$(VBA.CStr(includePattern))
    excludePattern = VBA.Trim$(VBA.CStr(excludePattern))

    If VBA.Len(valueText) = 0 Then
        private_Dev_MatchesIncludeExclude = (VBA.Len(includePattern) = 0)
        Exit Function
    End If

    If VBA.Len(includePattern) > 0 Then
        If Not fn_Helpers_IsRegexMatch(valueText, includePattern) Then Exit Function
    End If

    If VBA.Len(excludePattern) > 0 Then
        If fn_Helpers_IsRegexMatch(valueText, excludePattern) Then Exit Function
    End If

    private_Dev_MatchesIncludeExclude = True
End Function


Private Function private_Dev_TryResolveFileComponentType( _
    ByVal fileName As String, _
    ByRef outCompType As String, _
    ByRef outFallbackName As String _
) As Boolean
    Dim normalizedName As String
    Dim baseName As String

    normalizedName = VBA.LCase$(VBA.Trim$(fileName))
    outCompType = VBA.vbNullString
    outFallbackName = VBA.vbNullString

    If fn_Helpers_EndsWith(normalizedName, ".utf8.vba") Then
        baseName = VBA.Left$(fileName, VBA.Len(fileName) - VBA.Len(".utf8.vba"))
        normalizedName = VBA.LCase$(VBA.Trim$(baseName))
    ElseIf fn_Helpers_EndsWith(normalizedName, ".vba") Then
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
    ElseIf fn_Helpers_EndsWith(normalizedName, ".cls") Then
        outCompType = COMP_TYPE_CLASS
        outFallbackName = VBA.Left$(baseName, VBA.Len(baseName) - VBA.Len(".cls"))
    End If

    private_Dev_TryResolveFileComponentType = (VBA.Len(VBA.Trim$(outCompType)) > 0 And VBA.Len(VBA.Trim$(outFallbackName)) > 0)
End Function


Private Function private_Dev_CreateDictionary() As Object
    Set private_Dev_CreateDictionary = VBA.CreateObject("Scripting.Dictionary")
    private_Dev_CreateDictionary.CompareMode = 1
End Function


Private Function private_Dev_NormalizeCacheKey(ByVal filePath As String) As String
    private_Dev_NormalizeCacheKey = VBA.LCase$(VBA.Replace$(VBA.CStr(filePath), "/", "\"))
End Function


Private Function private_Dev_BuildFileStamp(ByVal filePath As String) As String
    Dim fso As Object
    Dim fileObj As Object

    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Exit Function
    Set fileObj = fso.GetFile(filePath)
    private_Dev_BuildFileStamp = private_Dev_BuildFileDateStampFromFileObject(fileObj) & ":" & private_Dev_BuildFileSizeStampFromFileObject(fileObj)
End Function


Private Function private_Dev_BuildFileDateStampFromFileObject(ByVal fileObj As Object) As String
    private_Dev_BuildFileDateStampFromFileObject = VBA.CStr(VBA.CDbl(fileObj.DateLastModified))
End Function


Private Function private_Dev_BuildFileSizeStampFromFileObject(ByVal fileObj As Object) As String
    private_Dev_BuildFileSizeStampFromFileObject = VBA.CStr(VBA.CLng(fileObj.Size))
End Function


Private Function private_Dev_IsCacheRecordCurrentByMode( _
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
    If Not private_Dev_IsCacheRecordMatchByMode(rec, fileDateStamp, fileSizeStamp, updateMode) Then Exit Function

    private_Dev_IsCacheRecordCurrentByMode = True
End Function


Private Function private_Dev_TryGetCachedComponentNameByMode( _
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
    If Not private_Dev_IsCacheRecordMatchByMode(rec, fileDateStamp, fileSizeStamp, updateMode) Then Exit Function

    outComponentName = VBA.CStr(rec("Name"))
    private_Dev_TryGetCachedComponentNameByMode = (VBA.Len(outComponentName) > 0)
End Function


Private Function private_Dev_IsCacheRecordMatchByMode( _
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

    private_Dev_IsCacheRecordMatchByMode = True
End Function


Private Sub private_Dev_SetCacheRecord( _
    ByVal cache As Object, _
    ByVal cacheKey As String, _
    ByVal compType As String, _
    ByVal componentName As String, _
    ByVal fileDateStamp As String, _
    ByVal fileSizeStamp As String _
)
    Dim rec As Object

    If cache Is Nothing Then Exit Sub

    Set rec = private_Dev_CreateDictionary()
    rec("Type") = compType
    rec("Name") = componentName
    rec("DateStamp") = fileDateStamp
    rec("SizeStamp") = fileSizeStamp

    If cache.Exists(cacheKey) Then
        cache.Remove cacheKey
    End If
    cache.Add cacheKey, rec
End Sub


Private Function private_Dev_LoadImportCache(ByVal cachePath As String) As Object
    Dim cache As Object
    Dim cacheText As String
    Dim normalizedText As String
    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim parts() As String
    Dim fileDateStamp As String
    Dim fileSizeStamp As String

    Set cache = private_Dev_CreateDictionary()
    If Not fn_Helpers_TryGetFileText(cachePath, cacheText, True, True) Then
        Err.Raise vbObjectError + 3100, "private_Dev_LoadImportCache", "Failed to load import cache from '" & cachePath & "'."
    End If

    If VBA.Len(VBA.Trim$(cacheText)) = 0 Then
        Set private_Dev_LoadImportCache = cache
        Exit Function
    End If

    normalizedText = VBA.Replace$(cacheText, VBA.vbCr, VBA.vbNullString)
    lines = VBA.Split(normalizedText, VBA.vbLf)
    For i = LBound(lines) To UBound(lines)
        lineText = lines(i)
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
        End If

        private_Dev_SetCacheRecord cache, VBA.CStr(parts(0)), VBA.CStr(parts(1)), VBA.CStr(parts(2)), fileDateStamp, fileSizeStamp
ContinueLoop:
    Next i

    Set private_Dev_LoadImportCache = cache
End Function


Private Sub private_Dev_SaveImportCache(ByVal cachePath As String, ByVal cache As Object)
    Dim key As Variant
    Dim rec As Object
    Dim cacheText As String

    If cache Is Nothing Then Exit Sub

    For Each key In cache.Keys
        Set rec = cache(VBA.CStr(key))
        If VBA.Len(cacheText) > 0 Then cacheText = cacheText & VBA.vbCrLf
        cacheText = cacheText & VBA.CStr(key) & "|" & VBA.CStr(rec("Type")) & "|" & VBA.CStr(rec("Name")) & "|" & VBA.CStr(rec("DateStamp")) & "|" & VBA.CStr(rec("SizeStamp"))
    Next key

    If Not private_FileCache_SetFileText(cachePath, cacheText, True) Then
        Err.Raise vbObjectError + 3101, "private_Dev_SaveImportCache", "Failed to save import cache to '" & cachePath & "'."
    End If
End Sub


Private Sub private_Dev_RemoveStaleImportedComponentsByScope( _
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
                If Not private_Dev_ShouldProcessComponentByScope(componentName, includeComponentPattern, excludeComponentPattern) Then GoTo ContinueKey
                    If VBA.StrComp(compType, COMP_TYPE_MODULE, VBA.vbTextCompare) = 0 Or _
                       VBA.StrComp(compType, COMP_TYPE_CLASS, VBA.vbTextCompare) = 0 Then
                        private_Dev_RemoveComponentIfExists componentName
                    End If
            End If
        End If
ContinueKey:
    Next key
End Sub


Private Function private_Dev_GetNextOnTimeTick() As Date
    Dim nowValue As Date

    ' Excel.OnTime планирует с точностью до секунды, поэтому даем +1 сек от текущего времени.
    nowValue = VBA.Now
    private_Dev_GetNextOnTimeTick = VBA.DateSerial(VBA.Year(nowValue), VBA.Month(nowValue), VBA.Day(nowValue)) + _
                                VBA.TimeSerial(VBA.Hour(nowValue), VBA.Minute(nowValue), VBA.Second(nowValue) + 1)
End Function

' Callstack[1]: ex_Core.private_Dev_TryRunSafeUpdateByMode -> private_Dev_TryRunRuntimeBooleanFunction
Private Function private_Dev_TryRunRuntimeBooleanFunction( _
    ByVal moduleName As String, _
    ByVal functionName As String, _
    ByRef outResult As Boolean _
) As Boolean
    Dim callResult As Variant
    Dim errDescription As String

    outResult = False
    If Not private_Dev_TryRunRuntimeNoArgMember(moduleName, functionName, callResult) Then Exit Function

    On Error Resume Next
    outResult = VBA.CBool(callResult)
    If Err.Number <> 0 Then
        errDescription = Err.Description
        Err.Clear
        On Error GoTo 0
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "runtime-call-failed: module='" & VBA.Replace$(moduleName, "'", "''") & "' function='" & VBA.Replace$(functionName, "'", "''") & "' err='bool-cast-failed: " & VBA.Replace$(errDescription, "'", "''") & "'"
#End If
        Exit Function
    End If
    On Error GoTo 0

    private_Dev_TryRunRuntimeBooleanFunction = True
End Function

' Callstack[1]: ex_Core.private_Dev_TryRunRuntimeBooleanFunction -> private_Dev_TryRunRuntimeNoArgMember
Private Function private_Dev_TryRunRuntimeNoArgMember( _
    ByVal moduleName As String, _
    ByVal memberName As String, _
    ByRef outResult As Variant, _
    Optional ByVal suppressFailureLog As Boolean = False _
) As Boolean
    Dim macroRef As String
    Dim errDescriptionQualified As String
    Dim runtimeComponent As Object

    outResult = Empty
    moduleName = VBA.Trim$(moduleName)
    memberName = VBA.Trim$(memberName)
    If VBA.Len(moduleName) = 0 Then Exit Function
    If VBA.Len(memberName) = 0 Then Exit Function

    Set runtimeComponent = private_Dev_TryGetComponentByName(moduleName)
    If runtimeComponent Is Nothing Then
        If Not suppressFailureLog Then
#If LOGGING_DEBUG_ENABLED Then
            private_Diagnostic_LogCoreSelfEvent "runtime-call-failed: module='" & VBA.Replace$(moduleName, "'", "''") & "' member='" & VBA.Replace$(memberName, "'", "''") & "' err='component is missing'"
#End If
        End If
        Exit Function
    End If

    macroRef = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!" & moduleName & "." & memberName

    On Error Resume Next
    outResult = Application.Run(macroRef)
    If Err.Number = 0 Then
        private_Dev_TryRunRuntimeNoArgMember = True
        On Error GoTo 0
        Exit Function
    End If
    errDescriptionQualified = Err.Description
    Err.Clear
    On Error GoTo 0

    If Not suppressFailureLog Then
#If LOGGING_DEBUG_ENABLED Then
        ' Runtime-вызов никогда не должен разрешаться через другую открытую
        ' книгу: unqualified Application.Run нарушает изоляцию VBA-проектов.
        private_Diagnostic_LogCoreSelfEvent "runtime-call-failed: module='" & VBA.Replace$(moduleName, "'", "''") & "' member='" & VBA.Replace$(memberName, "'", "''") & "' err='" & VBA.Replace$(errDescriptionQualified, "'", "''") & "'"
#End If
    End If
End Function


Private Function private_Dev_UpdateSheetModule( _
    ByVal sheetName As String, _
    ByVal sheetCodePath As String, _
    Optional ByVal preloadedCodeText As String = VBA.vbNullString _
) As Boolean
    Dim vbProj As Object
    Dim vbComp As Object
    Dim cm As Object
    Dim codeText As String

    Set vbProj = ThisWorkbook.VBProject
    If Not private_Dev_SheetModuleExists(vbProj, sheetName) Then Exit Function

    If VBA.Len(preloadedCodeText) > 0 Then
        codeText = preloadedCodeText
    Else
        If VBA.Len(private_Dev_BuildFileStamp(sheetCodePath)) = 0 Then Exit Function
        codeText = private_Dev_ReadAllText(sheetCodePath)
    End If

    Set vbComp = vbProj.VBComponents(sheetName)
    Set cm = vbComp.CodeModule

    cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString codeText
    private_Dev_UpdateSheetModule = True
End Function


Private Function private_Dev_ResolveSheetCodeName(ByVal fileStem As String) As String
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(fileStem)
    On Error GoTo 0

    If Not ws Is Nothing Then
        private_Dev_ResolveSheetCodeName = ws.CodeName
    Else
        private_Dev_ResolveSheetCodeName = fileStem
    End If
End Function


Private Function private_Dev_SheetModuleExists(ByVal vbProj As Object, ByVal sheetName As String) As Boolean
    Dim vbComp As Object
    On Error Resume Next
    Set vbComp = vbProj.VBComponents(sheetName)
    private_Dev_SheetModuleExists = Not vbComp Is Nothing
    On Error GoTo 0
End Function


Private Function private_Dev_FindWorkbookComponentName() As String
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
            private_Dev_FindWorkbookComponentName = nameCandidates(i)
            Exit Function
        End If
    Next i
End Function


Private Function private_Dev_UpdateWorkbookModuleFromText( _
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

    private_Dev_UpdateWorkbookModuleFromText = True
End Function


Private Function private_Dev_ReadAllText(ByVal filePath As String) As String
    private_Dev_ReadAllText = private_Dev_ReadAllTextByCharset(filePath, "utf-8")
    If VBA.Left$(private_Dev_ReadAllText, 1) = VBA.ChrW$(65279) Then
        private_Dev_ReadAllText = VBA.Mid$(private_Dev_ReadAllText, 2)
    End If
End Function


Private Function private_Dev_ReadAllTextByCharset(ByVal filePath As String, ByVal charsetName As String) As String
    Dim stream As Object

    Set stream = VBA.CreateObject("ADODB.Stream")
    stream.Type = 2 ' текстовый поток
    stream.Mode = 3 ' режим чтение/запись
    stream.Charset = charsetName
    stream.Open
    stream.LoadFromFile filePath
    private_Dev_ReadAllTextByCharset = stream.ReadText(-1)
    stream.Close
End Function


Private Sub private_Dev_ClearDocumentModuleCode(ByVal vbComp As Object)
    Dim cm As Object

    If vbComp Is Nothing Then Exit Sub
    Set cm = vbComp.CodeModule
    If cm Is Nothing Then Exit Sub
    If cm.CountOfLines <= 0 Then Exit Sub

    cm.DeleteLines 1, cm.CountOfLines
End Sub
' --------------------------------------
'  } // namespace Dev
' --------------------------------------

Private Sub private_ShowStatusNotice(ByVal messageText As String, ByVal useNativeStatus As Boolean, Optional ByVal timeoutSeconds As Long = 3)
    If private_UseNativeStatus(useNativeStatus) Then
        private_ShowNativeStatus messageText
    Else
        If Not private_TryShowRtStatus("fn_ShowStatusBarNotice", messageText, timeoutSeconds) Then
            private_ShowNativeStatus messageText
        End If
    End If
End Sub


Private Sub private_ShowStatusSuccess(ByVal messageText As String, ByVal useNativeStatus As Boolean, Optional ByVal timeoutSeconds As Long = 3)
    If private_UseNativeStatus(useNativeStatus) Then
        private_ShowNativeStatus messageText
    Else
        If Not private_TryShowRtStatus("fn_ShowStatusBarSuccess", messageText, timeoutSeconds) Then
            private_ShowNativeStatus messageText
        End If
    End If
End Sub


Private Sub private_ShowStatusWarning(ByVal messageText As String, ByVal useNativeStatus As Boolean, Optional ByVal timeoutSeconds As Long = 3)
    If private_UseNativeStatus(useNativeStatus) Then
        private_ShowNativeStatus "Warning: " & messageText
    Else
        If Not private_TryShowRtStatus("fn_ShowStatusBarWarning", messageText, timeoutSeconds) Then
            private_ShowNativeStatus "Warning: " & messageText
        End If
    End If
End Sub


Private Sub private_ShowStatusError(ByVal messageText As String, ByVal useNativeStatus As Boolean, Optional ByVal timeoutSeconds As Long = 3)
    If private_UseNativeStatus(useNativeStatus) Then
        private_ShowNativeStatus "Error: " & messageText
    Else
        If Not private_TryShowRtStatus("fn_ShowStatusBarError", messageText, timeoutSeconds) Then
            private_ShowNativeStatus "Error: " & messageText
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
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "rt-messaging-call-failed: " & methodName & ": " & errDescription
#End If
    End If
    On Error GoTo 0
End Function


Private Function private_TryRefreshRuntimeStaticControl( _
    ByVal controlName As String _
) As Boolean
    Dim macroRef As String
    Dim callResult As Variant

    controlName = VBA.Trim$(controlName)
    If VBA.Len(controlName) = 0 Then Exit Function

    ' String-bound runtime boundary сохраняет ex_Core автономным initial
    ' installer. Этот метод вызывается только после установки runtime.
    macroRef = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & _
        "'!ex_ControlRefreshRuntime.fn_TryRefreshStaticControl"

    On Error GoTo EH_REFRESH
    callResult = Application.Run(macroRef, controlName)
    private_TryRefreshRuntimeStaticControl = VBA.CBool(callResult)
    Exit Function

EH_REFRESH:
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent _
        "control-refresh-runtime-call-failed control='" & _
        VBA.Replace$(controlName, "'", "''") & "' err='" & _
        VBA.Replace$(Err.Description, "'", "''") & "'"
#End If
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
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "native-status-failed: message='" & VBA.Replace$(messageText, "'", "''") & "' err='" & VBA.Replace$(errDescription, "'", "''") & "'"
#End If
        Exit Sub
    End If
    On Error GoTo 0

    private_Diagnostic_LogStatusBarEvent "native-show", messageText
End Sub

' --------------------------------------
'  namespace Diagnostic {
' --------------------------------------
Private Sub private_Diagnostic_LogStatusBarEvent( _
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
    
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent logLine
#End If
#End If
End Sub


Private Sub private_Diagnostic_LogCoreSelfEvent(ByVal messageText As String)
#If Not CORE_ENABLE_SELF_LOGGING Then
    Exit Sub
#Else
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreEvent messageText
#End If
#End If
End Sub


Private Sub private_Diagnostic_LogCoreEvent( _
    ByVal messageText As String, _
    Optional ByVal ignoreRuntimePagePolicy As Boolean = False _
)
    Dim enableLogging As Boolean
    Dim logPath As String
    Dim folderPath As String
    Dim fso As Object
    Dim stream As Object
    Dim lineText As String

    If Not ignoreRuntimePagePolicy And _
        g_IsLoggingRuntimeStateInitialized Then
        If Not g_IsLoggingRuntimeActive Then Exit Sub
    End If
    If Not fn_Settings_TryGetFlagBoolean(SETTINGS_FLAG_IS_LOGGING_ENABLED, SETTINGS_FLAG_IS_LOGGING_ENABLED_DEFAULT, enableLogging, False) Then Exit Sub
    If Not enableLogging Then Exit Sub

    messageText = VBA.Trim$(VBA.CStr(messageText))
    If VBA.Len(messageText) = 0 Then Exit Sub

    If VBA.Len(VBA.Trim$(ThisWorkbook.Path)) = 0 Then Exit Sub

    logPath = ThisWorkbook.Path & "\\" & CORE_LOG_FILE_REL_PATH
    folderPath = VBA.Left$(logPath, VBA.InStrRev(logPath, "\\") - 1)

    On Error Resume Next
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
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
End Sub


Private Sub private_Diagnostic_ClearCoreLogFile()
    Dim logPath As String
    Dim folderPath As String
    Dim fso As Object
    Dim stream As Object

    If VBA.Len(VBA.Trim$(ThisWorkbook.Path)) = 0 Then Exit Sub

    logPath = ThisWorkbook.Path & "\\" & CORE_LOG_FILE_REL_PATH
    folderPath = VBA.Left$(logPath, VBA.InStrRev(logPath, "\\") - 1)

    On Error Resume Next
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        If VBA.Len(folderPath) > 0 Then
            If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
        End If
    End If

    Set stream = fso.OpenTextFile(logPath, 2, True) ' режим перезаписи файла (ForWriting)
    If Not stream Is Nothing Then stream.Close
    Err.Clear
    On Error GoTo 0
End Sub
' --------------------------------------
'  } // namespace Diagnostic
' --------------------------------------
