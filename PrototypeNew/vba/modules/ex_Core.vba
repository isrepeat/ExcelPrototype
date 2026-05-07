' Должен быть вставлен во внутренний модуль книги .xlsm
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
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
Private Const MAX_VBA_COMPONENT_NAME_LEN As Long = 31
Private Const UPDATE_MODE_FULL As Long = 1
Private Const UPDATE_MODE_DATE As Long = 2
Private Const UPDATE_MODE_SIZE As Long = 3
Private Const ERR_COMPONENT_STILL_PRESENT As Long = VBA.vbObjectError + 1015
' Retry нужен для transient-сценария: VBIDE иногда держит компонент "живым" еще 1 тик
' после Remove, и следующий OnTime-запуск проходит уже без правок исходников.
Private Const MAX_SAFE_UPDATE_RETRY_ATTEMPTS As Long = 1
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

Private g_QueuedBridgeUpdateAt As Date
Private g_QueuedBridgeUpdateMacro As String
Private g_QueuedRuntimeStateRestoreAt As Date
Private g_QueuedRuntimeStateRestoreMacro As String
Private g_QueuedSafeUpdateRetryAt As Date
Private g_QueuedSafeUpdateRetryMacro As String

Private g_FileCacheMap As Object
Private g_GlobalItemsSourceMap As Object
Private g_GlobalObjectSourceMap As Object
Private g_SafeUpdateRetryOperation As String
Private g_SafeUpdateRetryAttempts As Long
Private g_LastUpdateErrorNumber As Long
Private g_LastUpdateErrorSource As String
Private g_LastUpdateErrorDescription As String

Public Sub m_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:ex_Core.m_Module_Dispose"
#End If
    On Error Resume Next
    Set g_FileCacheMap = Nothing
    Set g_GlobalItemsSourceMap = Nothing
    Set g_GlobalObjectSourceMap = Nothing
    On Error GoTo 0
End Sub

' //
' // API
' //
' --------------------------------------
'  namespace Dev {
' --------------------------------------
Public Sub m_Dev_OpenCoreModule()
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


Public Sub m_Dev_RemoveAllModulesAndClasses()
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


' Callstack[1]: VBA.Macros(ex_Core.m_Dev_UpdateAllModules) -> ex_Core.m_Dev_UpdateAllModules
' Callstack[2]: DevUI(onClick: ex_Core.m_Dev_UpdateAllModules) -> ex_Core.m_Dev_UpdateAllModules
Public Sub m_Dev_UpdateAllModules()
    ' Если обновление запущено из bridge-click dispatch, переносим запуск на следующий тик.
    ' Это защищает от reentry: нельзя безопасно переимпортировать модули в середине обработки клика.
    If private_Dev_TryQueueRuntimeUpdateWhenBridgeDispatch("full") Then Exit Sub
    If private_Dev_TryRunSafeUpdateByMode(UPDATE_MODE_FULL, PATTERN_ALL_COMPONENTS, PATTERN_EXCLUDE_CORE, True, "full") Then Exit Sub
    private_ShowStatusError "Safe update (full) did not complete. Check core.log.", True, 6
End Sub


' Callstack[1]: VBA.Macros(ex_Core.m_Dev_UpdateCodeByDate) -> ex_Core.m_Dev_UpdateCodeByDate
Public Sub m_Dev_UpdateCodeByDate()
    ' Та же логика, что и для full: если мы внутри bridge-dispatch, только очередь через OnTime.
    If private_Dev_TryQueueRuntimeUpdateWhenBridgeDispatch("date") Then Exit Sub
    If private_Dev_TryRunSafeUpdateByMode(UPDATE_MODE_DATE, PATTERN_MAIN_COMPONENTS, PATTERN_EXCLUDE_CORE, False, "date") Then Exit Sub
    private_ShowStatusError "Safe update (date) did not complete. Check core.log.", True, 6
End Sub


' Callstack[1]: VBA.Macros(ex_Core.m_Dev_UpdateCodeBySize) -> ex_Core.m_Dev_UpdateCodeBySize
Public Sub m_Dev_UpdateCodeBySize()
    ' Та же логика, что и для full/date: отложенный запуск только если идет dispatch клика.
    If private_Dev_TryQueueRuntimeUpdateWhenBridgeDispatch("size") Then Exit Sub
    If private_Dev_TryRunSafeUpdateByMode(UPDATE_MODE_SIZE, PATTERN_MAIN_COMPONENTS, PATTERN_EXCLUDE_CORE, False, "size") Then Exit Sub
    private_ShowStatusError "Safe update (size) did not complete. Check core.log.", True, 6
End Sub


' Отдельная процедура для ядра рантайма (rt_*):
' Инкрементальные обновления по дате/размеру эти модули не затрагивают.
Public Sub m_Dev_UpdateRuntimeCore()
    If private_Dev_UpdateCodeByRegex(PATTERN_RUNTIME_COMPONENTS, PATTERN_EXCLUDE_CORE, UPDATE_MODE_FULL, True) Then Exit Sub
    private_ShowStatusError "Runtime core update did not complete. Check core.log.", True, 6
End Sub


' Callstack[1]: DevUI(onClick: ex_Core.m_Dev_ToggleLogging) -> ex_Core.m_Dev_ToggleLogging
Public Sub m_Dev_ToggleLogging()
    Dim isEnabled As Boolean

    If Not m_Settings_TryToggleFlagBoolean(SETTINGS_FLAG_IS_LOGGING_ENABLED, SETTINGS_FLAG_IS_LOGGING_ENABLED_DEFAULT, isEnabled, True) Then
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

    Call ex_HelpersSheet.m_TryRerenderActivePage("settings:toggle-logging")
End Sub
' --------------------------------------
'  } // namespace Dev
' --------------------------------------

' --------------------------------------
'  namespace RuntimeSource {
' --------------------------------------
' Callstack[1]: External bootstrap/init -> ex_Core.m_RuntimeSource_SetGlobalItemsSource
' Глобальные sources живут в ex_Core, чтобы избежать зависимости Settings/Diagnostic логики от rt_PageManager.
Public Function m_RuntimeSource_SetGlobalItemsSource(ByVal sourceKey As String, ByVal items As Collection) As Boolean
    Dim normalizedKey As String

    normalizedKey = private_RuntimeSource_NormalizeKey(sourceKey)
    If VBA.Len(normalizedKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "RuntimeSource: global items source key is empty."
#End If
        Exit Function
    End If
    If items Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "RuntimeSource: global items source collection is not specified for key '" & normalizedKey & "'."
#End If
        Exit Function
    End If

    private_RuntimeSource_EnsureStorage
    Set g_GlobalItemsSourceMap(normalizedKey) = items
    m_RuntimeSource_SetGlobalItemsSource = True
End Function


' Callstack[1]: External bootstrap/reset -> ex_Core.m_RuntimeSource_RemoveGlobalItemsSource
Public Function m_RuntimeSource_RemoveGlobalItemsSource(ByVal sourceKey As String) As Boolean
    Dim normalizedKey As String

    normalizedKey = private_RuntimeSource_NormalizeKey(sourceKey)
    If VBA.Len(normalizedKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "RuntimeSource: global items source key is empty."
#End If
        Exit Function
    End If

    private_RuntimeSource_EnsureStorage
    If g_GlobalItemsSourceMap.Exists(normalizedKey) Then
        g_GlobalItemsSourceMap.Remove normalizedKey
    End If
    m_RuntimeSource_RemoveGlobalItemsSource = True
End Function


' Callstack[1]: ex_RuntimeSourceResolver.m_TryResolveItemsSource -> ex_Core.m_RuntimeSource_TryGetGlobalItemsSourceByKey
Public Function m_RuntimeSource_TryGetGlobalItemsSourceByKey( _
    ByVal sourceKey As String, _
    ByRef outItems As Collection, _
    Optional ByVal allowMissing As Boolean = False _
) As Boolean
    Dim normalizedKey As String

    Set outItems = Nothing
    normalizedKey = private_RuntimeSource_NormalizeKey(sourceKey)

    If VBA.Len(normalizedKey) = 0 Then
        If allowMissing Then
            m_RuntimeSource_TryGetGlobalItemsSourceByKey = True
            Exit Function
        End If
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "RuntimeSource: global items source key is empty."
#End If
        Exit Function
    End If

    private_RuntimeSource_EnsureStorage
    If g_GlobalItemsSourceMap.Exists(normalizedKey) Then
        Set outItems = g_GlobalItemsSourceMap(normalizedKey)
        m_RuntimeSource_TryGetGlobalItemsSourceByKey = True
        Exit Function
    End If

    If allowMissing Then
        m_RuntimeSource_TryGetGlobalItemsSourceByKey = True
        Exit Function
    End If

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.m_Diagnostic_LogError "RuntimeSource: global items source '" & normalizedKey & "' is not registered."
#End If
End Function


' Callstack[1]: External bootstrap/init -> ex_Core.m_RuntimeSource_SetGlobalObjectSource
Public Function m_RuntimeSource_SetGlobalObjectSource(ByVal sourceKey As String, ByVal sourceObject As Object) As Boolean
    Dim normalizedKey As String

    normalizedKey = private_RuntimeSource_NormalizeKey(sourceKey)
    If VBA.Len(normalizedKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "RuntimeSource: global object source key is empty."
#End If
        Exit Function
    End If
    If sourceObject Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "RuntimeSource: global object source is not specified for key '" & normalizedKey & "'."
#End If
        Exit Function
    End If

    private_RuntimeSource_EnsureStorage
    Set g_GlobalObjectSourceMap(normalizedKey) = sourceObject
    m_RuntimeSource_SetGlobalObjectSource = True
End Function


' Callstack[1]: External bootstrap/reset -> ex_Core.m_RuntimeSource_RemoveGlobalObjectSource
Public Function m_RuntimeSource_RemoveGlobalObjectSource(ByVal sourceKey As String) As Boolean
    Dim normalizedKey As String

    normalizedKey = private_RuntimeSource_NormalizeKey(sourceKey)
    If VBA.Len(normalizedKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "RuntimeSource: global object source key is empty."
#End If
        Exit Function
    End If

    private_RuntimeSource_EnsureStorage
    If g_GlobalObjectSourceMap.Exists(normalizedKey) Then
        g_GlobalObjectSourceMap.Remove normalizedKey
    End If
    m_RuntimeSource_RemoveGlobalObjectSource = True
End Function


' Callstack[1]: ex_RuntimeSourceResolver.m_TryResolveObjectSource -> ex_Core.m_RuntimeSource_TryGetGlobalObjectSourceByKey
Public Function m_RuntimeSource_TryGetGlobalObjectSourceByKey( _
    ByVal sourceKey As String, _
    ByRef outObject As Object, _
    Optional ByVal allowMissing As Boolean = False _
) As Boolean
    Dim normalizedKey As String

    Set outObject = Nothing
    normalizedKey = private_RuntimeSource_NormalizeKey(sourceKey)

    If VBA.Len(normalizedKey) = 0 Then
        If allowMissing Then
            m_RuntimeSource_TryGetGlobalObjectSourceByKey = True
            Exit Function
        End If
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "RuntimeSource: global object source key is empty."
#End If
        Exit Function
    End If

    private_RuntimeSource_EnsureStorage

    If g_GlobalObjectSourceMap.Exists(normalizedKey) Then
        Set outObject = g_GlobalObjectSourceMap(normalizedKey)
        m_RuntimeSource_TryGetGlobalObjectSourceByKey = True
        Exit Function
    End If

    If allowMissing Then
        m_RuntimeSource_TryGetGlobalObjectSourceByKey = True
        Exit Function
    End If

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.m_Diagnostic_LogError "RuntimeSource: global object source '" & normalizedKey & "' is not registered."
#End If
End Function
' --------------------------------------
'  } // namespace RuntimeSource
' --------------------------------------

' --------------------------------------
'  namespace Settings {
' --------------------------------------
' Callstack[1]: ex_Core.m_Settings_TryToggleFlagBoolean -> ex_Core.m_Settings_TryGetFlagBoolean
' Callstack[2]: ex_Core.private_Diagnostic_LogCoreEvent -> ex_Core.m_Settings_TryGetFlagBoolean
Public Function m_Settings_TryGetFlagBoolean( _
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
    normalizedFlagName = m_Helpers_NormilizeString(flagName)
    If VBA.Len(normalizedFlagName) = 0 Then
        If showErrorUi Then ex_Core.m_Diagnostic_LogError "Settings: flag name is empty."
        Exit Function
    End If

    If Not private_Settings_TryEnsureFlagsMapCurrent(flagsMap, showErrorUi) Then Exit Function
    If flagsMap Is Nothing Then Exit Function

    If Not flagsMap.Exists(normalizedFlagName) Then
        m_Settings_TryGetFlagBoolean = True
        Exit Function
    End If

    rawValue = flagsMap(normalizedFlagName)
    If m_Helpers_TryParseBooleanText(VBA.CStr(rawValue), parsedValue) Then
        outValue = parsedValue
    ElseIf showErrorUi Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "Settings: flag '" & normalizedFlagName & "' has non-boolean value '" & VBA.CStr(rawValue) & "'."
#End If
    End If

    m_Settings_TryGetFlagBoolean = True
End Function


' Callstack[1]: External caller (macro/runtime action) -> ex_Core.m_Settings_TrySetFlagBoolean
Public Function m_Settings_TrySetFlagBoolean( _
    ByVal flagName As String, _
    ByVal newValue As Boolean, _
    Optional ByVal showErrorUi As Boolean = True _
) As Boolean
    Dim normalizedFlagName As String

    normalizedFlagName = m_Helpers_NormilizeString(flagName)
    If VBA.Len(normalizedFlagName) = 0 Then
        If showErrorUi Then ex_Core.m_Diagnostic_LogError "Settings: flag name is empty."
        Exit Function
    End If

    m_Settings_TrySetFlagBoolean = private_Settings_TryWriteFlagBoolean(normalizedFlagName, VBA.CBool(newValue), showErrorUi)
End Function


' Callstack[1]: ex_Core.m_Dev_ToggleLogging -> ex_Core.m_Settings_TryToggleFlagBoolean
Public Function m_Settings_TryToggleFlagBoolean( _
    ByVal flagName As String, _
    ByVal defaultValue As Boolean, _
    ByRef outValue As Boolean, _
    Optional ByVal showErrorUi As Boolean = True _
) As Boolean
    Dim currentValue As Boolean
    Dim nextValue As Boolean
    Dim normalizedFlagName As String

    outValue = defaultValue
    normalizedFlagName = m_Helpers_NormilizeString(flagName)
    If VBA.Len(normalizedFlagName) = 0 Then
        If showErrorUi Then ex_Core.m_Diagnostic_LogError "Settings: flag name is empty."
        Exit Function
    End If

    If Not m_Settings_TryGetFlagBoolean(normalizedFlagName, defaultValue, currentValue, showErrorUi) Then Exit Function
    nextValue = Not currentValue
    If Not private_Settings_TryWriteFlagBoolean(normalizedFlagName, nextValue, showErrorUi) Then Exit Function

    outValue = nextValue

    m_Settings_TryToggleFlagBoolean = True
End Function


' Callstack[1]: External caller (debug/runtime action) -> ex_Core.m_Settings_TryGetDom
' forceReload очищает только file-text cache; сам файл остается источником истины.
Public Function m_Settings_TryGetDom( _
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
    m_Settings_TryGetDom = True
End Function


' Callstack[1]: ex_RuntimeSourceResolver.m_TryResolveObjectSource(GlobalRuntimeSource='settings') -> ex_Core.m_Settings_TryGetObjectSource
' Возвращает snapshot-объект настроек из Settings.xml; чтение XML переиспользует общий file-cache по DateLastModified.
Public Function m_Settings_TryGetObjectSource( _
    ByRef outObjectSource As Object, _
    Optional ByVal showErrorUi As Boolean = True _
) As Boolean
    Dim settingsMap As Object

    Set outObjectSource = Nothing

    Set settingsMap = CreateObject("Scripting.Dictionary")
    settingsMap.CompareMode = 1
    ' Объект строится из общего file-cache; XML читаем только при изменении DateLastModified у Settings.xml.
    If Not private_Settings_TryFillObjectSourceMap(settingsMap, showErrorUi) Then Exit Function

    Set outObjectSource = settingsMap
    m_Settings_TryGetObjectSource = True
End Function
' --------------------------------------
'  } // namespace Settings
' --------------------------------------

' --------------------------------------
'  namespace Diagnostic {
' --------------------------------------
Public Sub m_Diagnostic_LogInfo(ByVal messageText As String)
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreEvent VBA.CStr(messageText)
#End If
End Sub


Public Sub m_Diagnostic_LogError(ByVal messageText As String)
    messageText = VBA.Trim$(VBA.CStr(messageText))
    If VBA.Len(messageText) = 0 Then Exit Sub
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreEvent "error: " & messageText
#End If
End Sub


Public Sub m_Diagnostic_LogWarning(ByVal messageText As String)
    messageText = VBA.Trim$(VBA.CStr(messageText))
    If VBA.Len(messageText) = 0 Then Exit Sub
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreEvent "warning: " & messageText
#End If
End Sub


Public Sub m_Diagnostic_LogStatusBarMessage( _
    ByVal actionName As String, _
    ByVal messageText As String, _
    Optional ByVal timeoutSeconds As Long = 0 _
)
    private_Diagnostic_LogStatusBarEvent actionName, messageText, timeoutSeconds
End Sub


Public Sub m_Diagnostic_ClearCoreLog()
    private_Diagnostic_ClearCoreLogFile
End Sub
' --------------------------------------
'  } // namespace Diagnostic
' --------------------------------------

' --------------------------------------
'  namespace CustomXmlPartStore {
' --------------------------------------
Public Function m_CustomXmlPartStore_TryFindPartByNamespace( _
    ByVal namespaceUri As String, _
    ByRef outPart As Object _
) As Boolean
    Dim parts As Object

    namespaceUri = VBA.Trim$(namespaceUri)
    If VBA.Len(namespaceUri) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "CustomXmlPartStore: namespace is empty."
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

    m_CustomXmlPartStore_TryFindPartByNamespace = True
    Exit Function

EH_FIND:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.m_Diagnostic_LogError "CustomXmlPartStore: failed to find XML part by namespace '" & namespaceUri & "': " & Err.Description
#End If
End Function


Public Function m_CustomXmlPartStore_TryLoadDomFromXml( _
    ByVal xmlText As String, _
    ByRef outDom As Object _
) As Boolean
    Dim dom As Object

    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False
    dom.validateOnParse = False
    dom.setProperty "SelectionLanguage", "XPath"

    If Not dom.LoadXML(VBA.CStr(xmlText)) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "CustomXmlPartStore: failed to parse XML."
#End If
        Exit Function
    End If

    Set outDom = dom
    m_CustomXmlPartStore_TryLoadDomFromXml = True
End Function


Public Function m_CustomXmlPartStore_TryCreateEmptyDom( _
    ByVal rootNodeName As String, _
    ByVal namespaceUri As String, _
    ByRef outDom As Object _
) As Boolean
    Dim xmlText As String

    rootNodeName = VBA.Trim$(rootNodeName)
    namespaceUri = VBA.Trim$(namespaceUri)

    If VBA.Len(rootNodeName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "CustomXmlPartStore: root node name is empty."
#End If
        Exit Function
    End If
    If VBA.Len(namespaceUri) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "CustomXmlPartStore: namespace is empty."
#End If
        Exit Function
    End If

    xmlText = "<?xml version=""1.0"" encoding=""UTF-8""?>" & _
              "<" & rootNodeName & " xmlns=""" & namespaceUri & """></" & rootNodeName & ">"

    If Not m_CustomXmlPartStore_TryLoadDomFromXml(xmlText, outDom) Then Exit Function
    m_CustomXmlPartStore_TryCreateEmptyDom = True
End Function


Public Function m_CustomXmlPartStore_TryLoadPartDom( _
    ByVal partObj As Object, _
    ByRef outDom As Object _
) As Boolean
    If partObj Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "CustomXmlPartStore: part is not specified."
#End If
        Exit Function
    End If

    If Not m_CustomXmlPartStore_TryLoadDomFromXml(VBA.CStr(partObj.XML), outDom) Then Exit Function
    m_CustomXmlPartStore_TryLoadPartDom = True
End Function


Public Function m_CustomXmlPartStore_TrySaveDom( _
    ByVal dom As Object, _
    ByVal existingPart As Object _
) As Boolean
    Dim xmlText As String

    If dom Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "CustomXmlPartStore: DOM is not specified."
#End If
        Exit Function
    End If

    xmlText = VBA.CStr(dom.XML)
    If VBA.Len(VBA.Trim$(xmlText)) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "CustomXmlPartStore: state XML is empty."
#End If
        Exit Function
    End If

    On Error GoTo EH_SAVE
    If Not existingPart Is Nothing Then existingPart.Delete
    ThisWorkbook.CustomXMLParts.Add xmlText
    m_CustomXmlPartStore_TrySaveDom = True
    Exit Function

EH_SAVE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.m_Diagnostic_LogError "CustomXmlPartStore: failed to persist state XML: " & Err.Description
#End If
End Function
' --------------------------------------
'  } // namespace CustomXmlPartStore
' --------------------------------------

' --------------------------------------
'  namespace Helpers {
' --------------------------------------
Public Function m_Helpers_NormilizeString(ByVal value As String) As String
    m_Helpers_NormilizeString = VBA.Trim$(VBA.CStr(value))
End Function


Public Function m_Helpers_TryParseBooleanText(ByVal textValue As String, ByRef outValue As Boolean) As Boolean
    textValue = VBA.LCase$(m_Helpers_NormilizeString(textValue))

    Select Case textValue
        Case "true", "1", "yes", "on"
            outValue = True
            m_Helpers_TryParseBooleanText = True
        Case "false", "0", "no", "off"
            outValue = False
            m_Helpers_TryParseBooleanText = True
    End Select
End Function


Public Function m_Helpers_BoolToText(ByVal value As Boolean) As String
    If value Then
        m_Helpers_BoolToText = "true"
    Else
        m_Helpers_BoolToText = "false"
    End If
End Function


Public Function m_Helpers_EndsWith(ByVal value As String, ByVal suffix As String) As Boolean
    m_Helpers_EndsWith = (VBA.LCase$(VBA.Right$(value, VBA.Len(suffix))) = VBA.LCase$(suffix))
End Function


Public Function m_Helpers_IsRegexMatch(ByVal valueText As String, ByVal regexPattern As String) As Boolean
    Dim re As Object

    regexPattern = m_Helpers_NormilizeString(regexPattern)
    If VBA.Len(regexPattern) = 0 Then Exit Function

    On Error GoTo EH

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = regexPattern

    m_Helpers_IsRegexMatch = re.Test(VBA.CStr(valueText))
    Exit Function

EH:
    Err.Raise VBA.vbObjectError + 1013, "m_Helpers_IsRegexMatch", "Некорректный regex '" & regexPattern & "': " & Err.Description
End Function


Public Function m_Helpers_TryGetFileText( _
    ByVal filePath As String, _
    ByRef outText As String, _
    Optional ByVal allowMissing As Boolean = True, _
    Optional ByVal showErrorUi As Boolean = True _
) As Boolean
    Dim fileStamp As String
    Dim cacheKey As String
    Dim entry As Object

    outText = VBA.vbNullString
    filePath = m_Helpers_NormilizeString(filePath)
    If VBA.Len(filePath) = 0 Then
        m_Helpers_TryGetFileText = allowMissing
        Exit Function
    End If

    ' Общий read-through cache:
    ' 1) читаем stamp файла,
    ' 2) возвращаем cached text только при совпадении stamp,
    ' 3) при промахе читаем файл и обновляем cache entry.
    cacheKey = private_FileCache_BuildFileTextKey(filePath)
    If Not private_FileCache_TryGetFileStamp(filePath, fileStamp, showErrorUi) Then
        Call private_FileCache_Remove(cacheKey)
        If allowMissing And VBA.Len(Dir(filePath)) = 0 Then
            m_Helpers_TryGetFileText = True
            Exit Function
        End If
        Exit Function
    End If

    If private_FileCache_TryGet(cacheKey, entry, True) Then
        If Not entry Is Nothing Then
            If entry.Exists("Stamp") Then
                If VBA.StrComp(VBA.CStr(entry("Stamp")), fileStamp, VBA.vbBinaryCompare) = 0 Then
                    If entry.Exists("Text") Then outText = VBA.CStr(entry("Text"))
                    m_Helpers_TryGetFileText = True
                    Exit Function
                End If
            End If
        End If
    End If

    If Not private_FileCache_TryReadTextFile(filePath, outText, showErrorUi) Then Exit Function
    If Not private_FileCache_SetFileTextEntry(filePath, outText, fileStamp) Then Exit Function
    m_Helpers_TryGetFileText = True
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
        If VBA.Len(Dir(settingsPath)) = 0 Then
            If showErrorUi Then
#If LOGGING_DEBUG_ENABLED Then
                ex_Core.m_Diagnostic_LogError "Settings: file '" & settingsPath & "' was not found."
#End If
            End If
            Exit Function
        End If
    End If

    If Not m_Helpers_TryGetFileText(settingsPath, settingsXmlText, createIfMissing, showErrorUi) Then Exit Function
    If VBA.Len(VBA.Trim$(settingsXmlText)) = 0 Then
        If Not createIfMissing Then Exit Function
        settingsXmlText = private_Settings_BuildTemplateXml
    End If

    ' Settings всегда читаются из файла (через FileCache), поэтому объект не stale:
    ' при изменении DateLastModified helper автоматически перечитает XML с диска.
    If Not m_CustomXmlPartStore_TryLoadDomFromXml(settingsXmlText, outDom) Then Exit Function
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
                ex_Core.m_Diagnostic_LogError "Settings: unexpected root node '" & VBA.CStr(settingsDom.DocumentElement.baseName) & "'. Expected '" & SETTINGS_ROOT_NODE & "'."
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
        m_Helpers_BoolToText(SETTINGS_FLAG_IS_LOGGING_ENABLED_DEFAULT), _
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

    normalizedFlagName = m_Helpers_NormilizeString(flagName)
    normalizedDefaultValue = m_Helpers_NormilizeString(defaultValue)
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
        ex_Core.m_Diagnostic_LogError "Settings: invalid default flag name '" & normalizedFlagName & "' for XML node."
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

    Set outFlagsMap = CreateObject("Scripting.Dictionary")
    outFlagsMap.CompareMode = 1

    For Each childNode In flagsNode.ChildNodes
        If Not childNode Is Nothing Then
            If VBA.CLng(childNode.nodeType) = 1 Then
                flagKey = m_Helpers_NormilizeString(VBA.CStr(childNode.baseName))
                If VBA.Len(flagKey) > 0 Then
                    rawText = VBA.Trim$(VBA.CStr(childNode.Text))
                    If m_Helpers_TryParseBooleanText(rawText, parsedValue) Then
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
    private_Settings_TryWriteFlagBoolean = private_Settings_TryWriteFlagText(normalizedFlagName, m_Helpers_BoolToText(VBA.CBool(newValue)), showErrorUi)
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

    normalizedFlagName = m_Helpers_NormilizeString(normalizedFlagName)
    normalizedValue = m_Helpers_NormilizeString(newValue)
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

    normalizedFlagName = m_Helpers_NormilizeString(normalizedFlagName)
    If VBA.Len(normalizedFlagName) = 0 Then Exit Function

    If Not private_Settings_TryEnsureSettingsStructure(settingsDom, hasStructureChanges, showErrorUi) Then Exit Function
    Set flagsNode = settingsDom.selectSingleNode("/*[local-name()='" & SETTINGS_ROOT_NODE & "']/*[local-name()='" & SETTINGS_FLAGS_NODE & "']")
    If flagsNode Is Nothing Then Exit Function

    For Each childNode In flagsNode.ChildNodes
        If Not childNode Is Nothing Then
            If VBA.CLng(childNode.nodeType) = 1 Then
                If VBA.StrComp(m_Helpers_NormilizeString(VBA.CStr(childNode.baseName)), normalizedFlagName, VBA.vbTextCompare) = 0 Then
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
        ex_Core.m_Diagnostic_LogError "Settings: invalid flag name '" & normalizedFlagName & "' for XML node."
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
        If showErrorUi Then ex_Core.m_Diagnostic_LogError "Settings: workbook is not saved. Save .xlsm first."
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
    "    <" & SETTINGS_FLAG_IS_LOGGING_ENABLED & ">" & m_Helpers_BoolToText(SETTINGS_FLAG_IS_LOGGING_ENABLED_DEFAULT) & "</" & SETTINGS_FLAG_IS_LOGGING_ENABLED & ">" & VBA.vbCrLf & _
        "  </Flags>" & VBA.vbCrLf & _
        "</Settings>"
End Function


Private Function private_Settings_TryEnsureTemplateExists(ByVal settingsPath As String, ByVal showErrorUi As Boolean) As Boolean
    settingsPath = VBA.Trim$(settingsPath)
    If VBA.Len(settingsPath) = 0 Then Exit Function

    If VBA.Len(Dir(settingsPath)) > 0 Then
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
        Set g_FileCacheMap = CreateObject("Scripting.Dictionary")
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
    If VBA.Len(Dir(filePath)) = 0 Then Exit Function

    On Error GoTo EH_STAMP
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileObj = fso.GetFile(filePath)
    outStamp = VBA.CStr(VBA.CDbl(fileObj.DateLastModified))
    private_FileCache_TryGetFileStamp = True
    Exit Function

EH_STAMP:
    If showErrorUi Then ex_Core.m_Diagnostic_LogError "FileCache: failed to read file modified date '" & filePath & "': " & Err.Description
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
    If VBA.Len(Dir(filePath)) = 0 Then Exit Function

    On Error GoTo EH_READ
    f = FreeFile
    Open filePath For Input As #f
    outText = Input$(LOF(f), f)
    Close #f
    private_FileCache_TryReadTextFile = True
    Exit Function

EH_READ:
    On Error Resume Next
    If f > 0 Then Close #f
    On Error GoTo 0
    If showErrorUi Then ex_Core.m_Diagnostic_LogError "FileCache: failed to read file '" & filePath & "': " & Err.Description
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
    f = FreeFile
    Open filePath For Output As #f
    Print #f, textValue
    Close #f
    private_FileCache_TryWriteTextFile = True
    Exit Function

EH_WRITE:
    On Error Resume Next
    If f > 0 Then Close #f
    On Error GoTo 0
    If showErrorUi Then ex_Core.m_Diagnostic_LogError "FileCache: failed to write file '" & filePath & "': " & Err.Description
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

    Set entry = CreateObject("Scripting.Dictionary")
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
        Set g_GlobalItemsSourceMap = CreateObject("Scripting.Dictionary")
        g_GlobalItemsSourceMap.CompareMode = 1
    End If

    If g_GlobalObjectSourceMap Is Nothing Then
        Set g_GlobalObjectSourceMap = CreateObject("Scripting.Dictionary")
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
Private Function private_Dev_TryQueueRuntimeUpdateWhenBridgeDispatch(ByVal updateKind As String) As Boolean
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
    Set bridgeComponent = private_Dev_TryGetComponentByName("rt_Bridge")
    If bridgeComponent Is Nothing Then Exit Function

    If Not private_Dev_TryRunRuntimeNoArgMember("rt_Bridge", "m_IsDispatchingClick", callResult, True) Then
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "queued-runtime-update-failed: kind='" & VBA.Replace$(updateKind, "'", "''") & "' err='bridge-dispatch-state-read-failed'"
#End If
        Exit Function
    End If

    On Error Resume Next
    isBridgeDispatching = VBA.CBool(callResult)
    If Err.Number <> 0 Then
        errDescription = Err.Description
        Err.Clear
        On Error GoTo 0
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "queued-runtime-update-failed: kind='" & VBA.Replace$(updateKind, "'", "''") & "' err='bridge-dispatch-state-cast-failed: " & VBA.Replace$(errDescription, "'", "''") & "'"
#End If
        Exit Function
    End If
    On Error GoTo 0

    If Not isBridgeDispatching Then Exit Function

    Select Case updateKind
        Case "full"
            coreMethod = "m_Dev_UpdateAllModules"
        Case "date"
            coreMethod = "m_Dev_UpdateCodeByDate"
        Case "size"
            coreMethod = "m_Dev_UpdateCodeBySize"
        Case Else
            Exit Function
    End Select

    macroRef = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!ex_Core." & coreMethod
    scheduleAt = private_Dev_GetNextOnTimeTick()

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
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "queued-runtime-update-failed: kind='" & VBA.Replace$(updateKind, "'", "''") & "' err='" & VBA.Replace$(errDescription, "'", "''") & "'"
#End If
        Exit Function
    End If
    On Error GoTo 0

    g_QueuedBridgeUpdateAt = scheduleAt
    g_QueuedBridgeUpdateMacro = macroRef

#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent "queued-runtime-update: kind='" & VBA.Replace$(updateKind, "'", "''") & "'"
#End If
    private_Dev_TryQueueRuntimeUpdateWhenBridgeDispatch = True
End Function


' Callstack[1]: ex_Core.m_Dev_UpdateAllModules -> private_Dev_TryRunSafeUpdateByMode
' Callstack[2]: ex_Core.m_Dev_UpdateCodeByDate -> private_Dev_TryRunSafeUpdateByMode
' Callstack[3]: ex_Core.m_Dev_UpdateCodeBySize -> private_Dev_TryRunSafeUpdateByMode
Private Function private_Dev_TryRunSafeUpdateByMode( _
    ByVal updateMode As Long, _
    ByVal includeComponentPattern As String, _
    ByVal excludeComponentPattern As String, _
    ByVal useNativeStatus As Boolean, _
    ByVal operationName As String _
) As Boolean
    Dim bootstrapMode As String
    Dim saveRuntimeOk As Boolean
    Dim updateOk As Boolean

    ' Этап 0. Нормализуем имя операции для логов.
    operationName = VBA.LCase$(VBA.Trim$(operationName))
    If VBA.Len(operationName) = 0 Then operationName = "unknown"
    If VBA.StrComp(g_SafeUpdateRetryOperation, operationName, VBA.vbTextCompare) <> 0 Then
        Call private_Dev_ResetSafeUpdateRetryState
    End If

#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent "safe-update:start op='" & operationName & "'"
#End If

    ' Этап 1. Гарантируем, что runtime-пайплайн вообще доступен.
    ' Важно: rt_RestoreManager/rt_PageManager могут отсутствовать (например, после частичного импорта/сброса проекта),
    ' поэтому сохранить snapshot "до любых действий" не всегда возможно.
    ' bootstrap сначала поднимает минимально нужные runtime-компоненты.
    If Not private_Dev_TryBootstrapRuntimePipeline(bootstrapMode) Then
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "safe-update:fail op='" & operationName & "' reason='runtime-bootstrap-failed'"
#End If
        Exit Function
    End If

    ' Если bootstrap был "full", массовый импорт уже выполнен.
    ' Упрощенный путь: не делаем sync-recovery в этом стеке, а сразу планируем
    ' единое deferred-восстановление snapshots/globals на следующий тик OnTime.
    ' Это выравнивает flow с обычной веткой (save/import -> deferred restore).
    If VBA.StrComp(bootstrapMode, "full", VBA.vbTextCompare) = 0 Then
        Call private_Dev_ResetSafeUpdateRetryState
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "safe-update:deferred op='" & operationName & "' reason='full-bootstrap-was-required'"
#End If
        Call private_Dev_QueueRuntimeStateRestoreAfterUpdate("safe-update:bootstrap:" & operationName)
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "safe-update:done op='" & operationName & "'"
#End If
        private_Dev_TryRunSafeUpdateByMode = True
        Exit Function
    End If

    ' Этап 2. Runtime валиден -> сохраняем runtime state перед целевым update.
    If Not private_Dev_TryRunRuntimeBooleanFunction("rt_RestoreManager", "m_SaveRuntimeState", saveRuntimeOk) Then
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
        If private_Dev_IsRetryableSafeUpdateFailure() Then
            If private_Dev_TryQueueSafeUpdateRetry(operationName, "component-still-present") Then
#If LOGGING_DEBUG_ENABLED Then
                private_Diagnostic_LogCoreSelfEvent "safe-update:deferred-retry op='" & operationName & "' reason='component-still-present'"
#End If
                private_ShowStatusWarning "Code update was retried automatically due to temporary VBA component lock.", useNativeStatus, 6
                private_Dev_TryRunSafeUpdateByMode = True
                Exit Function
            End If
#If LOGGING_DEBUG_ENABLED Then
            private_Diagnostic_LogCoreSelfEvent "safe-update:retry-exhausted op='" & operationName & "' reason='component-still-present'"
#End If
        End If
        Call private_Dev_ResetSafeUpdateRetryState
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "safe-update:fail op='" & operationName & "' reason='update-import-failed'"
#End If
        Exit Function
    End If

    ' Этап 5. Не делаем синхронный restore прямо здесь.
    ' Восстановление целиком переносится в deferred-путь на следующий тик OnTime,
    ' чтобы избежать двойного рендера (sync restore + deferred restore) после hot-update.
    Call private_Dev_ResetSafeUpdateRetryState
    Call private_Dev_QueueRuntimeStateRestoreAfterUpdate("safe-update:" & operationName)
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent "safe-update:done op='" & operationName & "'"
#End If
    private_Dev_TryRunSafeUpdateByMode = True
End Function


Private Function private_Dev_IsRetryableSafeUpdateFailure() As Boolean
    Dim errDescription As String

    ' Повторяем только узкий класс ошибок "still present after remove operation".
    ' Это симптом временного lock в VBIDE/COM, а не признак логической ошибки в коде.
    errDescription = VBA.Trim$(g_LastUpdateErrorDescription)
    If VBA.Len(errDescription) = 0 Then Exit Function
    If VBA.InStr(1, errDescription, "is still present after remove operation", VBA.vbTextCompare) = 0 Then Exit Function

    If g_LastUpdateErrorNumber = ERR_COMPONENT_STILL_PRESENT Then
        private_Dev_IsRetryableSafeUpdateFailure = True
        Exit Function
    End If

    ' В import-folder ошибка remove компонента агрегируется в vbObjectError+1001,
    ' поэтому допускаем retry и для обернутого случая.
    If g_LastUpdateErrorNumber = VBA.vbObjectError + 1001 Then
        If VBA.InStr(1, g_LastUpdateErrorSource, "private_Dev_ImportFolder", VBA.vbTextCompare) > 0 Then
            private_Dev_IsRetryableSafeUpdateFailure = True
            Exit Function
        End If
    End If
End Function


Private Function private_Dev_TryQueueSafeUpdateRetry(ByVal operationName As String, ByVal reasonText As String) As Boolean
    Dim macroRef As String
    Dim scheduleAt As Date
    Dim nextAttempt As Long
    Dim errDescription As String

    operationName = VBA.LCase$(VBA.Trim$(operationName))
    reasonText = VBA.Trim$(reasonText)
    If VBA.Len(operationName) = 0 Then Exit Function
    If VBA.Len(reasonText) = 0 Then reasonText = "unknown"

    ' Важно: retry не меняет исходники и не делает дополнительной очистки данных.
    ' Мы просто переносим повтор на следующий тик OnTime, чтобы lock успел отпуститься.
    nextAttempt = g_SafeUpdateRetryAttempts + 1
    If nextAttempt > MAX_SAFE_UPDATE_RETRY_ATTEMPTS Then Exit Function
    If Not private_Dev_TryResolveSafeUpdateRetryMacro(operationName, macroRef) Then
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "safe-update-retry-queue-failed op='" & operationName & "' err='macro-not-resolved'"
#End If
        Exit Function
    End If

    scheduleAt = private_Dev_GetNextOnTimeTick()

    On Error Resume Next
    If g_QueuedSafeUpdateRetryAt > 0# And VBA.Len(VBA.Trim$(g_QueuedSafeUpdateRetryMacro)) > 0 Then
        Application.OnTime EarliestTime:=g_QueuedSafeUpdateRetryAt, Procedure:=g_QueuedSafeUpdateRetryMacro, Schedule:=False
        Err.Clear
    End If

    Application.OnTime EarliestTime:=scheduleAt, Procedure:=macroRef
    If Err.Number <> 0 Then
        errDescription = Err.Description
        Err.Clear
        On Error GoTo 0
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "safe-update-retry-queue-failed op='" & operationName & "' err='" & VBA.Replace$(errDescription, "'", "''") & "'"
#End If
        Exit Function
    End If
    On Error GoTo 0

    g_QueuedSafeUpdateRetryAt = scheduleAt
    g_QueuedSafeUpdateRetryMacro = macroRef
    g_SafeUpdateRetryOperation = operationName
    g_SafeUpdateRetryAttempts = nextAttempt

#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent "safe-update-retry-queued op='" & operationName & "' attempt='" & VBA.CStr(nextAttempt) & "' reason='" & VBA.Replace$(reasonText, "'", "''") & "'"
#End If
    private_Dev_TryQueueSafeUpdateRetry = True
End Function


Private Function private_Dev_TryResolveSafeUpdateRetryMacro(ByVal operationName As String, ByRef outMacroRef As String) As Boolean
    Dim coreMethod As String

    outMacroRef = VBA.vbNullString
    operationName = VBA.LCase$(VBA.Trim$(operationName))
    If VBA.Len(operationName) = 0 Then Exit Function

    Select Case operationName
        Case "full"
            coreMethod = "m_Dev_UpdateAllModules"
        Case "date"
            coreMethod = "m_Dev_UpdateCodeByDate"
        Case "size"
            coreMethod = "m_Dev_UpdateCodeBySize"
        Case Else
            Exit Function
    End Select

    outMacroRef = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!ex_Core." & coreMethod
    private_Dev_TryResolveSafeUpdateRetryMacro = True
End Function


Private Sub private_Dev_ResetSafeUpdateRetryState()
    Call private_Dev_TryCancelQueuedSafeUpdateRetry
    g_SafeUpdateRetryOperation = VBA.vbNullString
    g_SafeUpdateRetryAttempts = 0
End Sub


Private Sub private_Dev_TryCancelQueuedSafeUpdateRetry()
    If g_QueuedSafeUpdateRetryAt > 0# And VBA.Len(VBA.Trim$(g_QueuedSafeUpdateRetryMacro)) > 0 Then
        On Error Resume Next
        Application.OnTime EarliestTime:=g_QueuedSafeUpdateRetryAt, Procedure:=g_QueuedSafeUpdateRetryMacro, Schedule:=False
        Err.Clear
        On Error GoTo 0
    End If

    g_QueuedSafeUpdateRetryAt = 0#
    g_QueuedSafeUpdateRetryMacro = VBA.vbNullString
End Sub


' Callstack[1]: ex_Core.m_Dev_UpdateAllModules -> private_Dev_TryBootstrapRuntimePipeline
' Callstack[2]: ex_Core.m_Dev_UpdateCodeByDate -> private_Dev_TryBootstrapRuntimePipeline
' Callstack[3]: ex_Core.m_Dev_UpdateCodeBySize -> private_Dev_TryBootstrapRuntimePipeline
Private Function private_Dev_TryBootstrapRuntimePipeline(ByRef outBootstrapMode As String) As Boolean
    outBootstrapMode = VBA.vbNullString

    ' Нормальный путь: runtime-компоненты уже на месте, ничего доп. не делаем.
    If private_Dev_AreSafeUpdateRuntimeComponentsPresent() Then
        outBootstrapMode = "none"
        private_Dev_TryBootstrapRuntimePipeline = True
        Exit Function
    End If

    ' Аварийный путь: runtime невалиден.
    ' Для rt_RestoreManager нужны зависимости из ex_*/obj_*, поэтому поднимаем полный набор.
    outBootstrapMode = "full"
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent "runtime-update-pipeline-bootstrap: start scope='all-components'"
#End If
    If Not private_Dev_UpdateCodeByRegex(PATTERN_ALL_COMPONENTS, PATTERN_EXCLUDE_CORE, UPDATE_MODE_FULL, True) Then
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "runtime-update-pipeline-bootstrap: fail import-failed"
#End If
        Exit Function
    End If

    If Not private_Dev_AreSafeUpdateRuntimeComponentsPresent() Then
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "runtime-update-pipeline-bootstrap: fail component-not-found-after-import"
#End If
        Exit Function
    End If

#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent "runtime-update-pipeline-bootstrap: done scope='all-components'"
#End If
    private_Dev_TryBootstrapRuntimePipeline = True
End Function


Private Function private_Dev_AreSafeUpdateRuntimeComponentsPresent() As Boolean
    ' Минимальный контракт для safe-update: если чего-то из этого нет,
    ' snapshot-сценарий нельзя считать надежным.
    If private_Dev_TryGetComponentByName("rt_RestoreManager") Is Nothing Then Exit Function
    If private_Dev_TryGetComponentByName("rt_PageManager") Is Nothing Then Exit Function
    If private_Dev_TryGetComponentByName("ex_HelpersSheet") Is Nothing Then Exit Function
    If private_Dev_TryGetComponentByName("obj_PageBase") Is Nothing Then Exit Function
    If private_Dev_TryGetComponentByName("obj_IPage") Is Nothing Then Exit Function
    If private_Dev_TryGetComponentByName("obj_ISerializable") Is Nothing Then Exit Function
    private_Dev_AreSafeUpdateRuntimeComponentsPresent = True
End Function


' Callstack[1]: ex_Core.private_Dev_TryRunSafeUpdateByMode -> private_Dev_QueueRuntimeStateRestoreAfterUpdate
Private Sub private_Dev_QueueRuntimeStateRestoreAfterUpdate(ByVal reasonText As String)
    Dim macroRef As String
    Dim scheduleAt As Date
    Dim errDescription As String

    reasonText = VBA.Trim$(reasonText)
    If VBA.Len(reasonText) = 0 Then reasonText = "unknown"

    macroRef = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!rt_RestoreManager.m_RunDeferredRuntimeStateRestore"
    scheduleAt = private_Dev_GetNextOnTimeTick()

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
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "runtime-state-restore-queue-failed reason='" & VBA.Replace$(reasonText, "'", "''") & "' err='" & VBA.Replace$(errDescription, "'", "''") & "'"
#End If
        Exit Sub
    End If
    On Error GoTo 0

    g_QueuedRuntimeStateRestoreAt = scheduleAt
    g_QueuedRuntimeStateRestoreMacro = macroRef
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent "runtime-state-restore-queued reason='" & VBA.Replace$(reasonText, "'", "''") & "'"
#End If
End Sub


Private Function private_Dev_TryPrepareRuntimeForHotUpdate(ByVal operationName As String) As Boolean
    operationName = VBA.Trim$(operationName)
    If VBA.Len(operationName) = 0 Then operationName = "unknown"

    ' На старте update убираем возможный "хвост" deferred restore,
    ' чтобы он не сработал, пока rt_* компоненты временно удалены.
    Call private_Dev_TryCancelQueuedRuntimeStateRestore("safe-update:prepare:" & operationName)

    ' Освобождаем runtime-ссылки через единый lifecycle-контракт модулей.
    ' rt_PageManager инкапсулирует dispose страниц внутри m_Module_Dispose.
    Call private_Dev_TryRunModuleDisposers("safe-update:prepare:" & operationName & ":dispose-modules")

    ' Даем завершиться Class_Terminate/освобождению COM-ссылок перед массовым remove/import.
    DoEvents

    ' Повторный проход module dispose после page dispose:
    ' часть модулей может освобождать ссылки только после того, как страницы уже закрыты.
    Call private_Dev_TryRunModuleDisposers("safe-update:prepare:" & operationName & ":post-dispose-pages")

    ' Сбрасываем глобальные runtime-sources, чтобы не удерживать старые class instances.
    Call private_RuntimeSource_ResetStorage

#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent "runtime-update-prepare:done op='" & VBA.Replace$(operationName, "'", "''") & "'"
#End If
    private_Dev_TryPrepareRuntimeForHotUpdate = True
End Function


Private Sub private_Dev_TryRunModuleDisposers(ByVal reasonText As String)
    Dim prj As Object
    Dim comp As Object
    Dim moduleName As String
    Dim macroRef As String
    Dim unqualifiedMacroRef As String
    Dim errDescriptionQualified As String
    Dim errDescriptionUnqualified As String

    reasonText = VBA.Trim$(reasonText)
    If VBA.Len(reasonText) = 0 Then reasonText = "unknown"

    Set prj = ThisWorkbook.VBProject
    If prj Is Nothing Then Exit Sub

    For Each comp In prj.VBComponents
        If comp Is Nothing Then GoTo ContinueComponent
        If VBA.CLng(comp.Type) <> 1 Then GoTo ContinueComponent ' Только стандартные модули.

        moduleName = VBA.Trim$(VBA.CStr(comp.Name))
        If VBA.Len(moduleName) = 0 Then GoTo ContinueComponent

        macroRef = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!" & moduleName & ".m_Module_Dispose"
        unqualifiedMacroRef = moduleName & ".m_Module_Dispose"

        On Error Resume Next
        Application.Run macroRef
        If Err.Number = 0 Then
            On Error GoTo 0
#If LOGGING_DEBUG_ENABLED Then
            private_Diagnostic_LogCoreSelfEvent "module-dispose-done module='" & VBA.Replace$(moduleName, "'", "''") & "' reason='" & VBA.Replace$(reasonText, "'", "''") & "'"
#End If
            GoTo ContinueComponent
        End If

        errDescriptionQualified = Err.Description
        Err.Clear

        Application.Run unqualifiedMacroRef
        If Err.Number = 0 Then
            On Error GoTo 0
#If LOGGING_DEBUG_ENABLED Then
            private_Diagnostic_LogCoreSelfEvent "module-dispose-done module='" & VBA.Replace$(moduleName, "'", "''") & "' reason='" & VBA.Replace$(reasonText, "'", "''") & "'"
#End If
            GoTo ContinueComponent
        End If

        errDescriptionUnqualified = Err.Description
        Err.Clear
        On Error GoTo 0
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "module-dispose-skip module='" & VBA.Replace$(moduleName, "'", "''") & "' reason='" & VBA.Replace$(reasonText, "'", "''") & "' qualifiedErr='" & VBA.Replace$(errDescriptionQualified, "'", "''") & "' unqualifiedErr='" & VBA.Replace$(errDescriptionUnqualified, "'", "''") & "'"
#End If
ContinueComponent:
    Next comp
End Sub


Private Sub private_Dev_TryCancelQueuedRuntimeStateRestore(Optional ByVal reasonText As String = VBA.vbNullString)
    Dim errDescription As String

    reasonText = VBA.Trim$(reasonText)
    If VBA.Len(reasonText) = 0 Then reasonText = "unknown"

    If g_QueuedRuntimeStateRestoreAt > 0# And VBA.Len(VBA.Trim$(g_QueuedRuntimeStateRestoreMacro)) > 0 Then
        On Error Resume Next
        Application.OnTime EarliestTime:=g_QueuedRuntimeStateRestoreAt, Procedure:=g_QueuedRuntimeStateRestoreMacro, Schedule:=False
        If Err.Number <> 0 Then
            errDescription = Err.Description
            Err.Clear
            On Error GoTo 0
#If LOGGING_DEBUG_ENABLED Then
            private_Diagnostic_LogCoreSelfEvent "runtime-state-restore-cancel-failed reason='" & VBA.Replace$(reasonText, "'", "''") & "' err='" & VBA.Replace$(errDescription, "'", "''") & "'"
#End If
        Else
            On Error GoTo 0
#If LOGGING_DEBUG_ENABLED Then
            private_Diagnostic_LogCoreSelfEvent "runtime-state-restore-cancelled reason='" & VBA.Replace$(reasonText, "'", "''") & "'"
#End If
        End If
    End If

    g_QueuedRuntimeStateRestoreAt = 0#
    g_QueuedRuntimeStateRestoreMacro = VBA.vbNullString
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
    Dim incrementalMode As Boolean
    Dim stageName As String
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim fullErrorText As String

    stageName = "init"
    incrementalMode = (updateMode <> UPDATE_MODE_FULL)
    Call private_Dev_ResetLastUpdateErrorState

    private_ShowStatusNotice "Code update started...", useNativeStatus, 1
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent "update-start"
#End If

    basePath = ThisWorkbook.Path & "\\" & BASE_DIR
    If VBA.Len(Dir(basePath, vbDirectory)) = 0 Then
        private_ShowStatusWarning "Workbook path is empty or 'vba' folder was not found. Save the workbook first.", useNativeStatus, 6
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "update-stop: vba-folder-not-found"
#End If
        private_Dev_UpdateCodeCore = False
        Exit Function
    End If

    ' Унифицированная очистка module-level ссылок перед любым hot-import.
    Call private_Dev_TryRunModuleDisposers("update-core:pre-import")

    Application.ScreenUpdating = False
    On Error GoTo EH

    stageName = "load-cache"
    cachePath = basePath & IMPORT_CACHE_FILE
    Set prevCache = private_Dev_LoadImportCache(cachePath)
    Set nextCache = private_Dev_CreateDictionary()

    If Not incrementalMode Then
        stageName = "remove-imported-by-scope"
        private_Dev_RemoveImportedModulesByScope includeComponentPattern, excludeComponentPattern
    End If

    stageName = "import-folder"
    private_Dev_ImportFolder basePath, updateMode, prevCache, nextCache, includeComponentPattern, excludeComponentPattern
    If ENABLE_CLASS_IMPORT_VALIDATION Then
        stageName = "validate-class-imports"
        private_Dev_ValidateClassImports basePath
    End If

    If incrementalMode Then
        stageName = "remove-stale"
        private_Dev_RemoveStaleImportedComponentsByScope prevCache, nextCache, includeComponentPattern, excludeComponentPattern
    End If
    stageName = "preserve-out-of-scope-cache"
    private_Dev_PreserveOutOfScopeCacheRecords prevCache, nextCache, includeComponentPattern, excludeComponentPattern
    stageName = "save-cache"
    private_Dev_SaveImportCache cachePath, nextCache

    Application.ScreenUpdating = True
    private_Dev_ShowCodeUpdatedNotice useNativeStatus
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreSelfEvent "update-done"
#End If
    private_Dev_UpdateCodeCore = True
    Exit Function

EH:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Call private_Dev_SetLastUpdateErrorState(errNumber, errSource, errDescription)
    fullErrorText = "Code update failed at stage '" & stageName & "': [" & errSource & " #" & VBA.CStr(errNumber) & "] " & errDescription

    Application.ScreenUpdating = True
    private_ShowStatusError fullErrorText, useNativeStatus, 6

    ' Статус-бар часто обрезает длинный текст ошибки импорта.
    ' Пишем полную диагностику (включая список файлов) в лог.
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.m_Diagnostic_LogError fullErrorText
#End If

    ' Логируем ошибку напрямую в core.log, даже если CORE_ENABLE_SELF_LOGGING = False.
#If LOGGING_DEBUG_ENABLED Then
    private_Diagnostic_LogCoreEvent "update-fail: stage='" & stageName & "' err='" & VBA.Replace$(errDescription, "'", "''") & "'"
#End If
    private_Dev_UpdateCodeCore = False
End Function


Private Sub private_Dev_ResetLastUpdateErrorState()
    g_LastUpdateErrorNumber = 0
    g_LastUpdateErrorSource = VBA.vbNullString
    g_LastUpdateErrorDescription = VBA.vbNullString
End Sub


Private Sub private_Dev_SetLastUpdateErrorState( _
    ByVal errNumber As Long, _
    ByVal errSource As String, _
    ByVal errDescription As String _
)
    g_LastUpdateErrorNumber = errNumber
    g_LastUpdateErrorSource = VBA.CStr(errSource)
    g_LastUpdateErrorDescription = VBA.CStr(errDescription)
End Sub


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

    If Dir(rootPath, vbDirectory) = "" Then
        Err.Raise VBA.vbObjectError + 1006, "private_Dev_ValidateClassImports", "VBA root folder not found: " & rootPath
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
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


Private Sub private_Dev_ShowCodeUpdatedNotice(ByVal useNativeStatus As Boolean)
    private_ShowStatusSuccess "Code updated.", useNativeStatus, 1
End Sub


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


Private Sub private_Dev_RemoveImportedModulesByScope( _
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
            If private_Dev_ShouldProcessComponentByScope(VBA.CStr(comp.Name), includeComponentPattern, excludeComponentPattern) Then
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
    Err.Raise VBA.vbObjectError + 1004, "private_Dev_RemoveImportedModulesByScope", _
              "Failed to remove component '" & names(i) & "': " & Err.Description
End Sub


Private Sub private_Dev_ImportFolder( _
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
    Dim importPass As Long

    If Dir(folderPath, vbDirectory) = "" Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
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
        private_Dev_ImportFolderRecursive rootFolder, 0, failed, updateMode, prevCache, nextCache, includeComponentPattern, excludeComponentPattern, importPass
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
    Optional ByVal importPass As Long = 1 _
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

                    private_Dev_RemoveComponentIfExists componentName
                    If VBA.StrComp(compType, COMP_TYPE_MODULE, VBA.vbTextCompare) = 0 Then
                        private_Dev_ImportStandardModuleFromSource componentName, importPath, sourceText
                    Else
                        private_Dev_ImportClassModuleFromSource componentName, importPath, sourceText
                    End If
                    private_Dev_SetCacheRecord nextCache, cacheKey, compType, componentName, fileDateStamp, fileSizeStamp

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
                    End If
            End Select
            On Error GoTo 0
        End If

ContinueNextFile:
    Next fileObj

    For Each subFolder In folderObj.SubFolders
        private_Dev_ImportFolderRecursive subFolder, depth + 1, failed, updateMode, prevCache, nextCache, includeComponentPattern, excludeComponentPattern, importPass
    Next subFolder

    Exit Sub

EH_IMPORT_FILE:
    errText = VBA.CStr(Err.Number) & ": " & Err.Description
    failed = failed & VBA.vbCrLf & "- " & importPath & " (" & errText & ")"
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
    If m_Helpers_EndsWith(normalizedName, ".utf8.vba") Then
        baseName = VBA.Left$(fileName, VBA.Len(fileName) - VBA.Len(".utf8.vba"))
    ElseIf m_Helpers_EndsWith(normalizedName, ".vba") Then
        baseName = VBA.Left$(fileName, VBA.Len(fileName) - VBA.Len(".vba"))
    Else
        private_Dev_ShouldImportFileInPass = False
        Exit Function
    End If

    normalizedName = VBA.LCase$(VBA.Trim$(baseName))
    isInterfaceClass = False
    If m_Helpers_EndsWith(normalizedName, ".cls") Then
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

    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(1) ' стандартный модуль (vbext_ct_StdModule)
    vbComp.Name = componentName
    Set cm = vbComp.CodeModule
    cm.DeleteLines 1, cm.CountOfLines
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

    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(2) ' модуль класса (vbext_ct_ClassModule)
    vbComp.Name = componentName
    Set cm = vbComp.CodeModule
    cm.DeleteLines 1, cm.CountOfLines
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

    private_Dev_ExtractCodeBody = outText
End Function


Private Sub private_Dev_RemoveComponentIfExists(ByVal componentName As String)
    Dim vbComp As Object
    Dim attempt As Long

    If VBA.Len(componentName) = 0 Then Exit Sub

    Set vbComp = private_Dev_TryGetComponentByName(componentName)
    If vbComp Is Nothing Then Exit Sub

    ThisWorkbook.VBProject.VBComponents.Remove vbComp

    ' После Remove Excel/VBE иногда освобождает компонент не мгновенно.
    ' Коротко "дожидаемся" исчезновения, чтобы не ловить конфликт имени на следующем Add.
    For attempt = 1 To 8
        Set vbComp = private_Dev_TryGetComponentByName(componentName)
        If vbComp Is Nothing Then Exit Sub
        DoEvents
    Next attempt

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
        If Not m_Helpers_IsRegexMatch(valueText, includePattern) Then Exit Function
    End If

    If VBA.Len(excludePattern) > 0 Then
        If m_Helpers_IsRegexMatch(valueText, excludePattern) Then Exit Function
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

    If m_Helpers_EndsWith(normalizedName, ".utf8.vba") Then
        baseName = VBA.Left$(fileName, VBA.Len(fileName) - VBA.Len(".utf8.vba"))
        normalizedName = VBA.LCase$(VBA.Trim$(baseName))
    ElseIf m_Helpers_EndsWith(normalizedName, ".vba") Then
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
    ElseIf m_Helpers_EndsWith(normalizedName, ".cls") Then
        outCompType = COMP_TYPE_CLASS
        outFallbackName = VBA.Left$(baseName, VBA.Len(baseName) - VBA.Len(".cls"))
    End If

    private_Dev_TryResolveFileComponentType = (VBA.Len(VBA.Trim$(outCompType)) > 0 And VBA.Len(VBA.Trim$(outFallbackName)) > 0)
End Function


Private Function private_Dev_CreateDictionary() As Object
    Set private_Dev_CreateDictionary = CreateObject("Scripting.Dictionary")
    private_Dev_CreateDictionary.CompareMode = 1
End Function


Private Function private_Dev_NormalizeCacheKey(ByVal filePath As String) As String
    private_Dev_NormalizeCacheKey = VBA.LCase$(VBA.Replace$(VBA.CStr(filePath), "/", "\"))
End Function


Private Function private_Dev_BuildFileStamp(ByVal filePath As String) As String
    Dim fso As Object
    Dim fileObj As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
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
    If Not m_Helpers_TryGetFileText(cachePath, cacheText, True, True) Then
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

' Callstack[1]: ex_Core.private_Dev_TryQueueRuntimeUpdateWhenBridgeDispatch -> private_Dev_TryRunRuntimeNoArgMember
' Callstack[2]: ex_Core.private_Dev_TryRunRuntimeBooleanFunction -> private_Dev_TryRunRuntimeNoArgMember
Private Function private_Dev_TryRunRuntimeNoArgMember( _
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
    unqualifiedMacroRef = moduleName & "." & memberName

    On Error Resume Next
    outResult = Application.Run(macroRef)
    If Err.Number = 0 Then
        private_Dev_TryRunRuntimeNoArgMember = True
        On Error GoTo 0
        Exit Function
    End If
    errDescriptionQualified = Err.Description
    Err.Clear

    outResult = Application.Run(unqualifiedMacroRef)
    If Err.Number = 0 Then
        private_Dev_TryRunRuntimeNoArgMember = True
        On Error GoTo 0
        Exit Function
    End If

    errDescriptionUnqualified = Err.Description
    Err.Clear
    On Error GoTo 0

    If Not suppressFailureLog Then
#If LOGGING_DEBUG_ENABLED Then
        private_Diagnostic_LogCoreSelfEvent "runtime-call-failed: module='" & VBA.Replace$(moduleName, "'", "''") & "' member='" & VBA.Replace$(memberName, "'", "''") & "' qualifiedErr='" & VBA.Replace$(errDescriptionQualified, "'", "''") & "' unqualifiedErr='" & VBA.Replace$(errDescriptionUnqualified, "'", "''") & "'"
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
    If VBA.Left$(private_Dev_ReadAllText, 1) = ChrW$(65279) Then
        private_Dev_ReadAllText = VBA.Mid$(private_Dev_ReadAllText, 2)
    End If
End Function


Private Function private_Dev_ReadAllTextByCharset(ByVal filePath As String, ByVal charsetName As String) As String
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
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
        private_ShowNativeStatus "Warning: " & messageText
    Else
        If Not private_TryShowRtStatus("m_ShowStatusBarWarning", messageText, timeoutSeconds) Then
            private_ShowNativeStatus "Warning: " & messageText
        End If
    End If
End Sub


Private Sub private_ShowStatusError(ByVal messageText As String, ByVal useNativeStatus As Boolean, Optional ByVal timeoutSeconds As Long = 3)
    If private_UseNativeStatus(useNativeStatus) Then
        private_ShowNativeStatus "Error: " & messageText
    Else
        If Not private_TryShowRtStatus("m_ShowStatusBarError", messageText, timeoutSeconds) Then
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


Private Sub private_Diagnostic_LogCoreEvent(ByVal messageText As String)
    Dim enableLogging As Boolean
    Dim logPath As String
    Dim folderPath As String
    Dim fso As Object
    Dim stream As Object
    Dim lineText As String

    If Not m_Settings_TryGetFlagBoolean(SETTINGS_FLAG_IS_LOGGING_ENABLED, SETTINGS_FLAG_IS_LOGGING_ENABLED_DEFAULT, enableLogging, False) Then Exit Sub
    If Not enableLogging Then Exit Sub

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
End Sub
' --------------------------------------
'  } // namespace Diagnostic
' --------------------------------------
