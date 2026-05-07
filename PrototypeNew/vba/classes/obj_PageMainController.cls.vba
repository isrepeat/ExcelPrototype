VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PageMainController"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const MODES_ROOT_REL_PATH As String = "modes"
Private Const MODE_PROFILES_FILE_SUFFIX As String = "Profiles.xml"
Private Const PERSONAL_CARD_SHEET_BASE_NAME As String = "PersonalCard"
Private Const MODE_ON_SELECT_MACRO As String = "OnConfigModeChanged"
Private Const PROFILE_ON_SELECT_MACRO As String = "OnConfigProfileChanged"
Private Const MODE_PICKER_CONTROL_NAME As String = "ConfigModePicker"
Private Const PROFILE_PICKER_CONTROL_NAME As String = "ConfigProfilePicker"
Private Const CONFIG_CONTROL_NAME As String = "DevConfig"
Private Const MODES_RUNTIME_KEY As String = "RuntimeItems.PageMain.ConfigModes"
Private Const PROFILES_RUNTIME_KEY As String = "RuntimeItems.PageMain.ConfigProfiles"
Private Const CONFIG_RUNTIME_KEY As String = "RuntimeItems.PageMain.Config"
Private Const CONTROLLER_RUNTIME_OBJECT_KEY As String = "RuntimeObjects.PageMain.Controller"

Private m_IsDisposed As Boolean
Private m_Page As obj_IPage
Private m_PageBase As obj_PageBase
Private m_ModeItemsProvider As obj_SIP_ModeFolders
Private m_ProfileItemsProvider As obj_SIP_ModeProfilesXml
Private m_SelectItemsProvidersReady As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Dispose
    On Error GoTo 0
End Sub


' //
' // API
' //
Public Function Initialize(ByVal page As obj_IPage) As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.Initialize"
    #End If
    If page Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: PageMainController initialization failed because page is not specified."
        #End If
        MsgBox "PrototypeNew: PageMainController initialization failed because page is not specified.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    m_IsDisposed = False
    Set m_Page = page
    Set m_PageBase = m_Page.GetPageBase()
    If m_PageBase Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: PageMainController initialization failed because page base is not specified."
        #End If
        MsgBox "PrototypeNew: PageMainController initialization failed because page base is not specified.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If
    If Not private_TryEnsureControllerObjectSourceBound(m_PageBase) Then Exit Function
    Initialize = True
End Function

Public Sub Dispose()
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.Dispose"
    #End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    If Not m_PageBase Is Nothing Then
        If Not m_PageBase.RuntimeSources Is Nothing Then
            Call m_PageBase.RuntimeSources.RemoveObjectSource(CONTROLLER_RUNTIME_OBJECT_KEY)
        End If
    End If
    Call ex_SelectItemsSourceProviders.fn_UnregisterProvider(MODES_RUNTIME_KEY)
    Call ex_SelectItemsSourceProviders.fn_UnregisterProvider(PROFILES_RUNTIME_KEY)
    Set m_ModeItemsProvider = Nothing
    Set m_ProfileItemsProvider = Nothing
    m_SelectItemsProvidersReady = False
    Set m_Page = Nothing
    Set m_PageBase = Nothing
    On Error GoTo 0
End Sub

Public Property Get RuntimeObjectSourceKey() As String
    RuntimeObjectSourceKey = CONTROLLER_RUNTIME_OBJECT_KEY
End Property

Public Function OnConfigModeChanged( _
    Optional ByVal notifyChange As Boolean = True, _
    Optional ByVal preferredPageBase As Object _
) As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.OnConfigModeChanged"
    #End If
    ' После смены режима нужно пересобрать runtime-источники.
    If Not private_TryPrepareModeProfileConfigRuntime(notifyChange, preferredPageBase) Then Exit Function
    OnConfigModeChanged = True
End Function


Public Function OnConfigModeDropDownOpened( _
    Optional ByVal preferredPageBase As Object _
) As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.OnConfigModeDropDownOpened"
    #End If
    Dim pageBase As obj_PageBase
    Dim modeOptions As Collection
    Dim usedCache As Boolean
    Dim existingModeItems As Collection

    ' DropDownOpened у mode-select обновляет только source режимов.
    ' Сам select после callback перечитывает source и перерисовывает dropdown.
    If Not private_TryResolvePageBase(pageBase, preferredPageBase) Then Exit Function
    If pageBase Is Nothing Then Exit Function

    If Not private_TryBuildModeSelectOptions(modeOptions, usedCache) Then Exit Function

    ' Если получили cache-hit и источник уже зарегистрирован на странице,
    ' повторная запись SetItemsSource не нужна.
    If usedCache Then
        If pageBase.RuntimeSources.TryGetItemsSourceByKey(MODES_RUNTIME_KEY, existingModeItems, True) Then
            If Not existingModeItems Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
                ex_Core.fn_Diagnostic_LogInfo "config-modes: skip-setitemsource reason='cache-hit-runtime-present'"
#End If
                OnConfigModeDropDownOpened = True
                Exit Function
            End If
        End If
    End If

    If Not private_TrySetItemsSource(MODES_RUNTIME_KEY, modeOptions, False, pageBase) Then Exit Function

    OnConfigModeDropDownOpened = True
End Function


Public Function OnConfigProfileDropDownOpened( _
    Optional ByVal preferredPageBase As Object _
) As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.OnConfigProfileDropDownOpened"
    #End If
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim modeOptions As Collection
    Dim profileOptions As Collection
    Dim selectedModeId As String
    Dim profileFilePath As String

    ' DropDownOpened у profile-select пересобирает:
    ' 1) актуальные режимы
    ' 2) профили для текущего выбранного режима
    If Not private_TryResolvePageBase(pageBase, preferredPageBase) Then Exit Function
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: worksheet is not specified for profile dropdown refresh."
        #End If
        MsgBox "PrototypeNew: worksheet is not specified for profile dropdown refresh.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    If Not private_TryBuildModeSelectOptions(modeOptions) Then Exit Function
    If Not private_TrySetItemsSource(MODES_RUNTIME_KEY, modeOptions, False, pageBase) Then Exit Function
    If Not private_TryResolveSelectedIdForControl(ws, MODE_PICKER_CONTROL_NAME, modeOptions, selectedModeId) Then Exit Function

    If Not private_TryBuildProfileSelectOptionsByMode(selectedModeId, profileOptions, profileFilePath) Then Exit Function
    If Not private_TrySetItemsSource(PROFILES_RUNTIME_KEY, profileOptions, False, pageBase) Then Exit Function

    OnConfigProfileDropDownOpened = True
End Function


Public Function OnConfigProfileChanged( _
    Optional ByVal notifyChange As Boolean = True, _
    Optional ByVal preferredPageBase As Object _
) As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.OnConfigProfileChanged"
    #End If
    ' После смены профиля также пересобираем runtime-состояние.
    If Not private_TryPrepareModeProfileConfigRuntime(notifyChange, preferredPageBase) Then Exit Function
    OnConfigProfileChanged = True
End Function


Public Sub SaveCurrentConfigProfile()
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.SaveCurrentConfigProfile"
    #End If
    If Not private_TrySaveCurrentConfigProfile() Then Exit Sub
End Sub

Public Function OpenPersonalCardPage() As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.OpenPersonalCardPage"
    #End If
    Dim sheetName As String
    Dim personalCardPage As obj_IPage
    Dim isPageCreated As Boolean

    On Error GoTo EH_OPEN

    sheetName = private_BuildUniqueWorksheetName(ThisWorkbook, PERSONAL_CARD_SHEET_BASE_NAME)
    If VBA.Len(sheetName) = 0 Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to allocate worksheet name for PersonalCard page."
        #End If
        MsgBox "PrototypeNew: failed to allocate worksheet name for PersonalCard page.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    Set personalCardPage = New obj_PagePersonalCard
    If personalCardPage Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to allocate PersonalCard page instance."
        #End If
        MsgBox "PrototypeNew: failed to allocate PersonalCard page instance.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    If m_Page Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: parent page is not specified for PersonalCard page creation."
        #End If
        MsgBox "PrototypeNew: parent page is not specified for PersonalCard page creation.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    If m_PageBase Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: page base is not specified for PersonalCard page creation."
        #End If
        MsgBox "PrototypeNew: page base is not specified for PersonalCard page creation.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    If m_PageBase.Worksheet Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: worksheet is not specified for PersonalCard page creation."
        #End If
        MsgBox "PrototypeNew: worksheet is not specified for PersonalCard page creation.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    If Not rt_PageManager.fn_CreatePage(personalCardPage, "ui\PersonalCardUI.xml", sheetName, m_Page) Then GoTo EH_CREATE
    isPageCreated = True

    If Not personalCardPage.RunPagePipeline() Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to run PersonalCard page pipeline."
        #End If
        MsgBox "PrototypeNew: failed to run PersonalCard page pipeline.", vbExclamation, "PrototypeNew / Config runtime"
        GoTo EH_CREATE
    End If

    If Not rt_PageManager.fn_RenderPage(personalCardPage, "pagemain:open-personalcard") Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to render PersonalCard page."
        #End If
        MsgBox "PrototypeNew: failed to render PersonalCard page.", vbExclamation, "PrototypeNew / Config runtime"
        GoTo EH_CREATE
    End If

    rt_Messaging.fn_ShowStatusBarSuccess "PersonalCard page has been created.", 3
    OpenPersonalCardPage = True
    Exit Function

EH_CREATE:
    On Error Resume Next
    If Not personalCardPage Is Nothing And isPageCreated Then
        Call rt_PageManager.fn_RemovePage(personalCardPage, True)
    End If
    On Error GoTo 0

    If Not OpenPersonalCardPage Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to create PersonalCard page."
        #End If
        MsgBox "PrototypeNew: failed to create PersonalCard page.", vbExclamation, "PrototypeNew / Config runtime"
    End If
    Exit Function

EH_OPEN:
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: exception in OpenPersonalCardPage: [" & VBA.CStr(Err.Number) & "] " & Err.Description
    #End If
    MsgBox "PrototypeNew: exception in OpenPersonalCardPage: [" & VBA.CStr(Err.Number) & "] " & Err.Description, vbExclamation, "PrototypeNew / Config runtime"
    Resume EH_CREATE
End Function


' //
' // Internal
' //
Private Function private_TryPrepareModeProfileConfigRuntime( _
    Optional ByVal notifyChange As Boolean = False, _
    Optional ByVal preferredPageBase As Object _
) As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.private_TryPrepareModeProfileConfigRuntime"
    #End If
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim modeOptions As Collection
    Dim profileOptions As Collection
    Dim selectedModeId As String
    Dim selectedProfileId As String
    Dim profileFilePath As String

    On Error GoTo EH_PREPARE_RUNTIME
    ' 1) Резолвим контекст текущей страницы.
    If Not private_TryResolvePageBase(pageBase, preferredPageBase) Then Exit Function
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: worksheet is not specified for mode/profile config runtime prepare."
        #End If
        MsgBox "PrototypeNew: worksheet is not specified for mode/profile config runtime prepare.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    ' 2) Формируем источник для списка режимов.
    If Not private_TryBuildModeSelectOptions(modeOptions) Then Exit Function
    If Not private_TrySetItemsSource(MODES_RUNTIME_KEY, modeOptions, False, pageBase) Then Exit Function

    ' 3) Для выбранного режима собираем список профилей.
    If Not private_TryResolveSelectedIdForControl(ws, MODE_PICKER_CONTROL_NAME, modeOptions, selectedModeId) Then Exit Function
    If Not private_TryBuildProfileSelectOptionsByMode(selectedModeId, profileOptions, profileFilePath) Then Exit Function
    If Not private_TrySetItemsSource(PROFILES_RUNTIME_KEY, profileOptions, False, pageBase) Then Exit Function

    ' 4) Загружаем config из выбранного профиля.
    If Not private_TryResolveSelectedIdForControl(ws, PROFILE_PICKER_CONTROL_NAME, profileOptions, selectedProfileId) Then Exit Function
    If Not private_TryRegisterConfigFromXmlProfile(profileFilePath, selectedProfileId, False, pageBase) Then Exit Function

    ' 5) По флагу обновляем UI.
    If notifyChange Then
        If Not private_TryRerenderPage(pageBase, "config:mode-profile-runtime") Then Exit Function
    End If

    private_TryPrepareModeProfileConfigRuntime = True
    On Error GoTo 0
    Exit Function

EH_PREPARE_RUNTIME:
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: exception in config runtime prepare: [" & VBA.CStr(Err.Number) & "] " & Err.Description
    #End If
    MsgBox "PrototypeNew: exception in config runtime prepare: [" & VBA.CStr(Err.Number) & "] " & Err.Description, vbExclamation, "PrototypeNew / Config runtime"
End Function


Private Function private_TrySaveCurrentConfigProfile( _
    Optional ByVal preferredPageBase As Object _
) As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.private_TrySaveCurrentConfigProfile"
    #End If
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim modeOptions As Collection
    Dim profileOptions As Collection
    Dim selectedModeId As String
    Dim selectedProfileId As String
    Dim profileFilePath As String
    Dim configControl As obj_ConfigControlVM
    Dim configEntries As Collection
    Dim dom As Object
    Dim profileNode As Object
    Dim generatedConfigNode As Object

    On Error GoTo EH_SAVE_PROFILE

    ' 1) Резолвим runtime-контекст страницы/листа.
    If Not private_TryResolvePageBase(pageBase, preferredPageBase) Then Exit Function
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: worksheet is not specified for config profile save."
        #End If
        MsgBox "PrototypeNew: worksheet is not specified for config profile save.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    ' 2) Определяем текущие выбранные mode/profile и путь к соответствующему Profiles.xml.
    If Not private_TryBuildModeSelectOptions(modeOptions) Then Exit Function
    If Not private_TryResolveSelectedIdForControl(ws, MODE_PICKER_CONTROL_NAME, modeOptions, selectedModeId) Then Exit Function
    If Not private_TryBuildProfileSelectOptionsByMode(selectedModeId, profileOptions, profileFilePath) Then Exit Function
    If Not private_TryResolveSelectedIdForControl(ws, PROFILE_PICKER_CONTROL_NAME, profileOptions, selectedProfileId) Then Exit Function

    ' 3) Резолвим Config-контрол.
    ' Здесь же читаем "плоскую" runtime-модель obj_ConfigEntry (Attr/Key/Value),
    ' она нужна для синхронизации RuntimeSources после успешного сохранения файла.
    If Not private_TryResolveConfigControl(pageBase, configControl) Then Exit Function
    If Not configControl.TryGetRenderedConfigEntries(configEntries) Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to read current entries from config control '" & CONFIG_CONTROL_NAME & "'."
        #End If
        MsgBox "PrototypeNew: failed to read current entries from config control '" & CONFIG_CONTROL_NAME & "'.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    ' 4) Грузим профильный XML и находим именно тот profile-node,
    ' который соответствует текущему выбору profileId.
    If Not private_TryLoadProfileDomAndNode(profileFilePath, selectedProfileId, dom, profileNode) Then Exit Function

    ' 5) Контрол формирует source-узел из текущего UI-рендера (без знания конкретного profile файла).
    ' Затем оркестратор переносит строки из source-узла в target profile-node.
    ' То есть ответственность разделена:
    ' - ConfigControl: "как представить текущие данные в XML"
    ' - Actions: "как применить этот XML к реальному профилю"
    If Not configControl.TryBuildRenderedConfigNode(dom, generatedConfigNode) Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to build source config node from control '" & CONFIG_CONTROL_NAME & "'."
        #End If
        MsgBox "PrototypeNew: failed to build source config node from control '" & CONFIG_CONTROL_NAME & "'.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If
    If Not private_TryReplaceProfileRowsFromSourceNode(profileNode, generatedConfigNode) Then
        Exit Function
    End If
    If Not private_TrySaveDomToFile(dom, profileFilePath) Then Exit Function

    ' 6) После успешного Save обновляем in-memory source,
    ' чтобы следующий render взял именно то состояние, которое ушло в файл.
    If Not private_TrySetItemsSource(CONFIG_RUNTIME_KEY, configEntries, False, pageBase) Then Exit Function

    rt_Messaging.fn_ShowStatusBarSuccess "Config profile '" & selectedProfileId & "' saved to '" & profileFilePath & "'.", 4
    private_TrySaveCurrentConfigProfile = True
    On Error GoTo 0
    Exit Function

EH_SAVE_PROFILE:
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: exception in config profile save: [" & VBA.CStr(Err.Number) & "] " & Err.Description
    #End If
    MsgBox "PrototypeNew: exception in config profile save: [" & VBA.CStr(Err.Number) & "] " & Err.Description, vbExclamation, "PrototypeNew / Config runtime"
End Function





Private Function private_TryRegisterConfigFromXmlProfile( _
    ByVal filePath As String, _
    ByVal profileKey As String, _
    Optional ByVal notifyChange As Boolean = False, _
    Optional ByVal preferredPageBase As Object _
) As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.private_TryRegisterConfigFromXmlProfile"
    #End If
    Dim dom As Object
    Dim profileNode As Object
    ' Загружаем DOM + выбранный профильный узел из внешнего Profiles.xml.
    If Not private_TryLoadProfileDomAndNode(filePath, profileKey, dom, profileNode) Then Exit Function

    ' Преобразуем XML-узел профиля в runtime-коллекцию и регистрируем в RuntimeSources.
    If Not private_TryRegisterConfigFromProfileNode(profileNode, notifyChange, preferredPageBase) Then Exit Function
    private_TryRegisterConfigFromXmlProfile = True
End Function


Private Function private_TryRegisterConfigFromProfileNode( _
    ByVal profileNode As Object, _
    Optional ByVal notifyChange As Boolean = False, _
    Optional ByVal preferredPageBase As Object _
) As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.private_TryRegisterConfigFromProfileNode"
    #End If
    Dim configTable As obj_ConfigTable
    Dim configEntries As list__obj_ConfigEntry
    Dim sourceItems As Collection
    Dim sourceConfigEntry As obj_ConfigEntry
    Dim i As Long

    ' Узел профиля обязателен.
    If profileNode Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: config profile node is not specified."
        #End If
        MsgBox "PrototypeNew: config profile node is not specified.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    ' Парсим XML profile-node в typed-модель obj_ConfigTable/obj_ConfigEntry.
    Set configTable = New obj_ConfigTable
    If Not configTable.TryLoadFromXmlNode(profileNode, True) Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to parse selected config profile node."
        #End If
        MsgBox "PrototypeNew: failed to parse selected config profile node.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    ' Готовим обычную Collection для runtime source map (совместимый формат источника).
    Set sourceItems = New Collection
    Set configEntries = configTable.Items
    If configEntries Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: config table entries collection is not initialized."
        #End If
        MsgBox "PrototypeNew: config table entries collection is not initialized.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    ' Важно: здесь используем индексный проход, а не For Each по list__*,
    ' чтобы не зависеть от NewEnum-атрибута после hot-import.
    For i = 1 To configEntries.Count
        Set sourceConfigEntry = configEntries.Item(i)
        If sourceConfigEntry Is Nothing Then GoTo ContinueSourceConfigEntry
        sourceItems.Add sourceConfigEntry
ContinueSourceConfigEntry:
    Next i

    ' Публикуем источник для Config-контрола.
    ' Если notifyChange=True, внешний слой может сделать немедленный rerender страницы.
    If Not private_TrySetItemsSource(CONFIG_RUNTIME_KEY, sourceItems, notifyChange, preferredPageBase) Then Exit Function
    private_TryRegisterConfigFromProfileNode = True
End Function


Private Function private_TryLoadProfileDomAndNode( _
    ByVal filePath As String, _
    ByVal profileKey As String, _
    ByRef outDom As Object, _
    ByRef outProfileNode As Object _
) As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.private_TryLoadProfileDomAndNode"
    #End If
    Dim normalizedFilePath As String
    Dim normalizedProfileKey As String
    Dim profileKeyLiteral As String
    Dim profileXPath As String

    Set outDom = Nothing
    Set outProfileNode = Nothing

    normalizedFilePath = VBA.Trim$(filePath)
    normalizedProfileKey = VBA.Trim$(profileKey)

    If VBA.Len(normalizedFilePath) = 0 Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: config profiles file path is empty."
        #End If
        MsgBox "PrototypeNew: config profiles file path is empty.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If
    If VBA.Len(normalizedProfileKey) = 0 Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: config profile key is empty."
        #End If
        MsgBox "PrototypeNew: config profile key is empty.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    ' 1) Загружаем XML документ режима (например <Mode>Profiles.xml).
    Set outDom = ex_XmlCore.fn_LoadDomByFilePath( _
        normalizedFilePath, _
        "PrototypeNew: config profiles file was not found: ", _
        "PrototypeNew: failed to parse config profiles file: ", _
        VBA.vbNullString)
    If outDom Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to load config profiles file '" & normalizedFilePath & "'."
        #End If
        MsgBox "PrototypeNew: failed to load config profiles file '" & normalizedFilePath & "'.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    profileKeyLiteral = ex_XmlCore.fn_XPathLiteral(normalizedProfileKey)
    ' 2) XPath подбирает профиль по id/name/key (атрибуты или дочерние теги),
    ' и одновременно гарантирует, что найденный узел действительно содержит
    ' конфиг-строки (item/row/entry/config) в себе или в потомках.
    profileXPath = "//*[" & _
                  "(" & _
                  "@id=" & profileKeyLiteral & " or @name=" & profileKeyLiteral & " or @key=" & profileKeyLiteral & " or " & _
                  "normalize-space(*[local-name()='id'][1])=" & profileKeyLiteral & " or " & _
                  "normalize-space(*[local-name()='name'][1])=" & profileKeyLiteral & " or " & _
                  "normalize-space(*[local-name()='key'][1])=" & profileKeyLiteral & _
                  ")" & _
                  " and " & _
                  "(.//*[local-name()='item' or local-name()='row' or local-name()='entry' or local-name()='config']" & _
                  " or *[local-name()='item' or local-name()='row' or local-name()='entry' or local-name()='config'])" & _
                  "]"

    ' 3) Возвращаем ровно один целевой узел профиля для чтения/перезаписи.
    Set outProfileNode = outDom.selectSingleNode(profileXPath)
    If outProfileNode Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: config profile '" & normalizedProfileKey & "' was not found in file '" & normalizedFilePath & "'."
        #End If
        MsgBox "PrototypeNew: config profile '" & normalizedProfileKey & "' was not found in file '" & normalizedFilePath & "'.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    private_TryLoadProfileDomAndNode = True
End Function


Private Function private_TryBuildModeSelectOptions( _
    ByRef outOptions As Collection, _
    Optional ByRef outUsedCache As Boolean = False _
) As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.private_TryBuildModeSelectOptions"
    #End If
    Set outOptions = Nothing
    outUsedCache = False

    ' Новый подход:
    ' 1) controller не знает деталей кеша/сканирования;
    ' 2) просто запрашивает items по providerKey;
    ' 3) provider+manager решают: cache-hit или rebuild.
    If Not private_TryEnsureSelectItemsProvidersRegistered() Then Exit Function
    If Not ex_SelectItemsSourceProviders.fn_TryResolveItemsByProviderKey(MODES_RUNTIME_KEY, outOptions, outUsedCache) Then Exit Function
    If outOptions Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: mode source provider returned empty collection."
        #End If
        MsgBox "PrototypeNew: mode source provider returned empty collection.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If
    If outOptions.Count = 0 Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: mode source provider returned no mode options."
        #End If
        MsgBox "PrototypeNew: mode source provider returned no mode options.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

#If LOGGING_DEBUG_ENABLED Then
    If outUsedCache Then
        ex_Core.fn_Diagnostic_LogInfo "config-modes: cache-hit count=" & VBA.CStr(outOptions.Count)
    Else
        ex_Core.fn_Diagnostic_LogInfo "config-modes: cache-refresh count=" & VBA.CStr(outOptions.Count)
    End If
#End If

    private_TryBuildModeSelectOptions = True
End Function


Private Function private_TryEnsureSelectItemsProvidersRegistered() As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.private_TryEnsureSelectItemsProvidersRegistered"
    #End If
    ' Регистрируем providers один раз за lifecycle модуля actions.
    ' Дальше все resolve идут через ex_SelectItemsSourceProviders.
    If m_SelectItemsProvidersReady Then
        private_TryEnsureSelectItemsProvidersRegistered = True
        Exit Function
    End If

    Set m_ModeItemsProvider = New obj_SIP_ModeFolders
    If m_ModeItemsProvider Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to create mode source provider."
        #End If
        MsgBox "PrototypeNew: failed to create mode source provider.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If
    If Not m_ModeItemsProvider.Initialize(MODES_RUNTIME_KEY, MODES_ROOT_REL_PATH, MODE_ON_SELECT_MACRO) Then Exit Function

    Set m_ProfileItemsProvider = New obj_SIP_ModeProfilesXml
    If m_ProfileItemsProvider Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to create profile source provider."
        #End If
        MsgBox "PrototypeNew: failed to create profile source provider.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If
    If Not m_ProfileItemsProvider.Initialize(PROFILES_RUNTIME_KEY, MODES_ROOT_REL_PATH, MODE_PROFILES_FILE_SUFFIX, PROFILE_ON_SELECT_MACRO) Then Exit Function

    If Not ex_SelectItemsSourceProviders.fn_RegisterProvider(m_ModeItemsProvider, True) Then Exit Function
    If Not ex_SelectItemsSourceProviders.fn_RegisterProvider(m_ProfileItemsProvider, True) Then Exit Function

    m_SelectItemsProvidersReady = True
    private_TryEnsureSelectItemsProvidersRegistered = True
End Function


Private Function private_TryBuildProfileSelectOptionsByMode( _
    ByVal modeId As String, _
    ByRef outOptions As Collection, _
    ByRef outProfilesFilePath As String _
) As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.private_TryBuildProfileSelectOptionsByMode"
    #End If
    Dim usedCache As Boolean

    Set outOptions = Nothing
    outProfilesFilePath = VBA.vbNullString

    ' Новый поток для профилей:
    ' 1) передаем provider-у текущий modeId;
    ' 2) cache manager сам решает cache-hit/cache-miss;
    ' 3) получаем и options, и фактический путь <Mode>Profiles.xml.
    If Not private_TryEnsureSelectItemsProvidersRegistered() Then Exit Function
    If m_ProfileItemsProvider Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: profile source provider is not initialized."
        #End If
        MsgBox "PrototypeNew: profile source provider is not initialized.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If
    If Not m_ProfileItemsProvider.SetCurrentModeId(modeId) Then Exit Function

    If Not ex_SelectItemsSourceProviders.fn_TryResolveItemsByProviderKey(PROFILES_RUNTIME_KEY, outOptions, usedCache) Then Exit Function
    If outOptions Is Nothing Then Exit Function
    If outOptions.Count = 0 Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: profile source provider returned no profile options for mode '" & modeId & "'."
        #End If
        MsgBox "PrototypeNew: profile source provider returned no profile options for mode '" & modeId & "'.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    outProfilesFilePath = VBA.Trim$(m_ProfileItemsProvider.CurrentProfilesFilePath)
    If VBA.Len(outProfilesFilePath) = 0 Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: profile source provider did not resolve profiles file path for mode '" & modeId & "'."
        #End If
        MsgBox "PrototypeNew: profile source provider did not resolve profiles file path for mode '" & modeId & "'.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

#If LOGGING_DEBUG_ENABLED Then
    If usedCache Then
        ex_Core.fn_Diagnostic_LogInfo "config-profiles: cache-hit mode='" & VBA.Replace$(VBA.Trim$(modeId), "'", "''") & "' count=" & VBA.CStr(outOptions.Count)
    Else
        ex_Core.fn_Diagnostic_LogInfo "config-profiles: cache-refresh mode='" & VBA.Replace$(VBA.Trim$(modeId), "'", "''") & "' count=" & VBA.CStr(outOptions.Count)
    End If
#End If

    private_TryBuildProfileSelectOptionsByMode = True
End Function


Private Function private_TryResolveSelectedIdForControl( _
    ByVal ws As Worksheet, _
    ByVal controlName As String, _
    ByVal options As Collection, _
    ByRef outSelectedId As String _
) As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.private_TryResolveSelectedIdForControl"
    #End If
    Dim storedId As String
    Dim firstId As String

    ' Сначала пробуем восстановить выбранный id из state store.
    outSelectedId = VBA.vbNullString
    If ws Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: worksheet is not specified for selectedId resolve."
        #End If
        MsgBox "PrototypeNew: worksheet is not specified for selectedId resolve.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If
    If options Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: options are not specified for control '" & controlName & "'."
        #End If
        MsgBox "PrototypeNew: options are not specified for control '" & controlName & "'.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If
    If options.Count = 0 Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: options are empty for control '" & controlName & "'."
        #End If
        MsgBox "PrototypeNew: options are empty for control '" & controlName & "'.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    If Not private_TryGetStoredSelectedIdForControl(ws, controlName, storedId) Then Exit Function
    If private_SelectOptionsContainsId(options, storedId) Then
        outSelectedId = VBA.Trim$(storedId)
        private_TryResolveSelectedIdForControl = True
        Exit Function
    End If

    ' Если не нашли — берем первый option и сохраняем его как выбранный.
    If Not private_TryGetFirstOptionId(options, firstId) Then Exit Function
    If VBA.Len(firstId) = 0 Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to resolve first option id for control '" & controlName & "'."
        #End If
        MsgBox "PrototypeNew: failed to resolve first option id for control '" & controlName & "'.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    outSelectedId = firstId
    If Not private_TrySetStoredSelectedIdForControl(ws, controlName, outSelectedId) Then Exit Function
    private_TryResolveSelectedIdForControl = True
End Function


Private Function private_SelectOptionsContainsId(ByVal options As Collection, ByVal optionId As String) As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.private_SelectOptionsContainsId"
    #End If
    Dim itemObj As Variant
    Dim normalizedId As String

    If options Is Nothing Then Exit Function

    normalizedId = VBA.LCase$(VBA.Trim$(optionId))
    If VBA.Len(normalizedId) = 0 Then Exit Function

    For Each itemObj In options
        If Not VBA.IsObject(itemObj) Then GoTo ContinueOptionContains
        If VBA.StrComp(VBA.TypeName(itemObj), "obj_SelectOption", VBA.vbTextCompare) <> 0 Then GoTo ContinueOptionContains

        If VBA.LCase$(VBA.Trim$(VBA.CStr(itemObj.Id))) = normalizedId Then
            private_SelectOptionsContainsId = True
            Exit Function
        End If
ContinueOptionContains:
    Next itemObj
End Function


Private Function private_TryGetFirstOptionId(ByVal options As Collection, ByRef outId As String) As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.private_TryGetFirstOptionId"
    #End If
    Dim itemObj As Variant

    outId = VBA.vbNullString
    If options Is Nothing Then Exit Function

    For Each itemObj In options
        If Not VBA.IsObject(itemObj) Then GoTo ContinueFirstOption
        If VBA.StrComp(VBA.TypeName(itemObj), "obj_SelectOption", VBA.vbTextCompare) <> 0 Then GoTo ContinueFirstOption

        outId = VBA.Trim$(VBA.CStr(itemObj.Id))
        private_TryGetFirstOptionId = True
        Exit Function
ContinueFirstOption:
    Next itemObj
End Function


Private Function private_TryGetStoredSelectedIdForControl( _
    ByVal ws As Worksheet, _
    ByVal controlName As String, _
    ByRef outSelectedId As String _
) As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.private_TryGetStoredSelectedIdForControl"
    #End If
    Dim selectStatic As obj_SelectControlVMStatic
    Dim selectKey As String

    outSelectedId = VBA.vbNullString
    If ws Is Nothing Then Exit Function

    ' Ключ хранения: "<SheetName>|<ControlName>".
    selectKey = VBA.LCase$(VBA.Trim$(ws.Name) & "|" & VBA.Trim$(controlName))
    If VBA.Len(selectKey) = 0 Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: select key is empty for control '" & controlName & "'."
        #End If
        MsgBox "PrototypeNew: select key is empty for control '" & controlName & "'.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    Set selectStatic = New obj_SelectControlVMStatic
    private_TryGetStoredSelectedIdForControl = selectStatic.TryGetSelectedId(selectKey, outSelectedId)
End Function


Private Function private_TrySetStoredSelectedIdForControl( _
    ByVal ws As Worksheet, _
    ByVal controlName As String, _
    ByVal selectedId As String _
) As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.private_TrySetStoredSelectedIdForControl"
    #End If
    Dim selectStatic As obj_SelectControlVMStatic
    Dim selectKey As String

    If ws Is Nothing Then Exit Function

    ' Ключ хранения: "<SheetName>|<ControlName>".
    selectKey = VBA.LCase$(VBA.Trim$(ws.Name) & "|" & VBA.Trim$(controlName))
    If VBA.Len(selectKey) = 0 Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: select key is empty for control '" & controlName & "'."
        #End If
        MsgBox "PrototypeNew: select key is empty for control '" & controlName & "'.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    Set selectStatic = New obj_SelectControlVMStatic
    private_TrySetStoredSelectedIdForControl = selectStatic.SetSelectedId(selectKey, VBA.Trim$(selectedId))
End Function


Private Function private_TryResolvePageBase( _
    ByRef outPageBase As obj_PageBase, _
    Optional ByVal preferredPageBase As Object _
) As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.private_TryResolvePageBase"
    #End If
    Set outPageBase = Nothing

    ' Приоритет: явно переданный page context.
    If Not preferredPageBase Is Nothing Then
        If TypeOf preferredPageBase Is obj_PageBase Then
            Set outPageBase = preferredPageBase
            If Not private_TryEnsureControllerObjectSourceBound(outPageBase) Then Exit Function
            private_TryResolvePageBase = True
            Exit Function
        End If

        If ex_HelpersSheet.fn_TryCastPageBase(preferredPageBase, outPageBase) Then
            If Not private_TryEnsureControllerObjectSourceBound(outPageBase) Then Exit Function
            private_TryResolvePageBase = True
            Exit Function
        End If

        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: preferred page runtime context has unsupported type '" & VBA.TypeName(preferredPageBase) & "'."
        #End If
        MsgBox "PrototypeNew: preferred page runtime context has unsupported type '" & VBA.TypeName(preferredPageBase) & "'.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    ' Основной fallback для page-owned controller: его собственный page context.
    If m_PageBase Is Nothing Then
        If Not m_Page Is Nothing Then
            Set m_PageBase = m_Page.GetPageBase()
        End If
    End If
    If Not m_PageBase Is Nothing Then
        Set outPageBase = m_PageBase
        If Not private_TryEnsureControllerObjectSourceBound(outPageBase) Then Exit Function
        private_TryResolvePageBase = True
        Exit Function
    End If

    ' Последний fallback: активная страница из runtime.
    If Not ex_HelpersSheet.fn_TryGetActivePageBase(outPageBase) Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: page runtime context is not resolved for active worksheet."
        #End If
        MsgBox "PrototypeNew: page runtime context is not resolved for active worksheet.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If
    If outPageBase Is Nothing Then Exit Function
    If Not private_TryEnsureControllerObjectSourceBound(outPageBase) Then Exit Function

    private_TryResolvePageBase = True
End Function


Private Function private_TryEnsureControllerObjectSourceBound(ByVal pageBase As obj_PageBase) As Boolean
    Dim runtimeSources As obj_PageRuntimeSources

    If pageBase Is Nothing Then Exit Function
    Set runtimeSources = pageBase.RuntimeSources
    If runtimeSources Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: page runtime sources are not specified for PageMainController."
        #End If
        MsgBox "PrototypeNew: page runtime sources are not specified for PageMainController.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    If Not runtimeSources.SetObjectSource(CONTROLLER_RUNTIME_OBJECT_KEY, Me) Then Exit Function
    private_TryEnsureControllerObjectSourceBound = True
End Function


Private Function private_TryResolveConfigControl( _
    ByVal pageBase As obj_PageBase, _
    ByRef outConfigControl As obj_ConfigControlVM _
) As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.private_TryResolveConfigControl"
    #End If
    Dim rawControl As Object

    Set outConfigControl = Nothing
    If pageBase Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: page base is not specified for config control resolve."
        #End If
        MsgBox "PrototypeNew: page base is not specified for config control resolve.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    If Not pageBase.TryGetRegisteredControlByName(CONFIG_CONTROL_NAME, rawControl) Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: config control '" & CONFIG_CONTROL_NAME & "' was not found in runtime registry."
        #End If
        MsgBox "PrototypeNew: config control '" & CONFIG_CONTROL_NAME & "' was not found in runtime registry.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If
    If rawControl Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: config control '" & CONFIG_CONTROL_NAME & "' runtime entry is empty."
        #End If
        MsgBox "PrototypeNew: config control '" & CONFIG_CONTROL_NAME & "' runtime entry is empty.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If
    If Not TypeOf rawControl Is obj_ConfigControlVM Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: config control '" & CONFIG_CONTROL_NAME & "' has unexpected type '" & VBA.TypeName(rawControl) & "'."
        #End If
        MsgBox "PrototypeNew: config control '" & CONFIG_CONTROL_NAME & "' has unexpected type '" & VBA.TypeName(rawControl) & "'.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    Set outConfigControl = rawControl
    private_TryResolveConfigControl = True
End Function


Private Function private_TryReplaceProfileRowsFromSourceNode( _
    ByVal targetProfileNode As Object, _
    ByVal sourceConfigNode As Object _
) As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.private_TryReplaceProfileRowsFromSourceNode"
    #End If
    Dim targetRowNodes As Object
    Dim sourceRowNodes As Object
    Dim rowIndex As Long
    Dim sourceRowNode As Object
    Dim clonedRowNode As Object

    If targetProfileNode Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: target profile node is not specified for save."
        #End If
        MsgBox "PrototypeNew: target profile node is not specified for save.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If
    If sourceConfigNode Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: source config node is not specified for save."
        #End If
        MsgBox "PrototypeNew: source config node is not specified for save.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    ' Этап A: очищаем у target-профиля все существующие row-узлы.
    ' Это делает операцию save "полной заменой", а не частичным merge.
    On Error GoTo EH_XML
    Set targetRowNodes = targetProfileNode.selectNodes("./*[local-name()='item' or local-name()='row' or local-name()='entry' or local-name()='config']")
    On Error GoTo 0

    If Not targetRowNodes Is Nothing Then
        For rowIndex = targetRowNodes.Length - 1 To 0 Step -1
            targetProfileNode.removeChild targetRowNodes.Item(rowIndex)
        Next rowIndex
    End If

    ' Этап B: читаем row-узлы из source, который сгенерировал ConfigControl.
    On Error GoTo EH_XML
    Set sourceRowNodes = sourceConfigNode.selectNodes("./*[local-name()='item' or local-name()='row' or local-name()='entry' or local-name()='config']")
    On Error GoTo 0
    If sourceRowNodes Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: source config node does not contain readable rows for save."
        #End If
        MsgBox "PrototypeNew: source config node does not contain readable rows for save.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If
    If sourceRowNodes.Length = 0 Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: source config node is empty and cannot replace profile rows."
        #End If
        MsgBox "PrototypeNew: source config node is empty and cannot replace profile rows.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    ' Этап C: переносим source-узлы в target.
    ' Используем cloneNode(True), чтобы перенос был независимым от исходного контейнера.
    For rowIndex = 0 To sourceRowNodes.Length - 1
        Set sourceRowNode = sourceRowNodes.Item(rowIndex)
        If sourceRowNode Is Nothing Then GoTo ContinueSourceRow

        ' В save-пайплайне source node строится тем же DOM, что и target profile node.
        Set clonedRowNode = sourceRowNode.cloneNode(True)
        If clonedRowNode Is Nothing Then
            #If LOGGING_DEBUG_ENABLED Then
                ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to clone source row node while updating profile."
            #End If
            MsgBox "PrototypeNew: failed to clone source row node while updating profile.", vbExclamation, "PrototypeNew / Config runtime"
            Exit Function
        End If
        targetProfileNode.appendChild clonedRowNode
ContinueSourceRow:
    Next rowIndex

    private_TryReplaceProfileRowsFromSourceNode = True
    Exit Function

EH_XML:
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to transfer source config rows into profile node: " & Err.Description
    #End If
    MsgBox "PrototypeNew: failed to transfer source config rows into profile node: " & Err.Description, vbExclamation, "PrototypeNew / Config runtime"
End Function


Private Function private_TrySaveDomToFile(ByVal dom As Object, ByVal filePath As String) As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.private_TrySaveDomToFile"
    #End If
    filePath = VBA.Trim$(filePath)
    If dom Is Nothing Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: DOM is not specified for file save."
        #End If
        MsgBox "PrototypeNew: DOM is not specified for file save.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If
    If VBA.Len(filePath) = 0 Then
        #If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PrototypeNew: file path is empty for profile save."
        #End If
        MsgBox "PrototypeNew: file path is empty for profile save.", vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    On Error GoTo EH_SAVE_DOM
    dom.Save filePath
    private_TrySaveDomToFile = True
    Exit Function

EH_SAVE_DOM:
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to write profile file '" & filePath & "': " & Err.Description
    #End If
    MsgBox "PrototypeNew: failed to write profile file '" & filePath & "': " & Err.Description, vbExclamation, "PrototypeNew / Config runtime"
End Function


Private Function private_TrySetItemsSource( _
    ByVal sourceKey As String, _
    ByVal items As Collection, _
    ByVal notifyChange As Boolean, _
    Optional ByVal preferredPageBase As Object _
) As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.private_TrySetItemsSource"
    #End If
    Dim pageBase As obj_PageBase
    Dim normalizedKey As String

    ' Приводим ключ к normalized-форме.
    If Not private_TryResolvePageBase(pageBase, preferredPageBase) Then Exit Function
    normalizedKey = VBA.LCase$(VBA.Trim$(sourceKey))

    If Not pageBase.RuntimeSources.SetItemsSource(normalizedKey, items) Then Exit Function

    ' По флагу делаем немедленный rerender.
    If notifyChange Then
        If Not private_TryRerenderPage(pageBase, "itemsSource:" & normalizedKey) Then Exit Function
    End If

    private_TrySetItemsSource = True
End Function


Private Function private_TryRerenderPage(ByVal pageBase As obj_PageBase, ByVal reason As String) As Boolean
    #If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageMainController.private_TryRerenderPage"
    #End If
    Dim pageRef As obj_IPage
    Dim ws As Worksheet

    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function

    ' Получаем страницу по worksheet и запускаем render.
    If Not rt_PageManager.fn_TryGetPageByWorksheet(ws, pageRef) Then Exit Function
    private_TryRerenderPage = rt_PageManager.fn_RenderPage(pageRef, reason)
End Function

Private Function private_BuildUniqueWorksheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim i As Long
    Dim suffix As String
    Dim candidate As String

    If wb Is Nothing Then Exit Function

    baseName = VBA.Trim$(baseName)
    If VBA.Len(baseName) = 0 Then baseName = "GeneratedPage"
    If VBA.Len(baseName) > 31 Then baseName = VBA.Left$(baseName, 31)

    If Not private_WorksheetNameExists(wb, baseName) Then
        private_BuildUniqueWorksheetName = baseName
        Exit Function
    End If

    For i = 1 To 9999
        suffix = "_" & VBA.CStr(i)
        candidate = VBA.Left$(baseName, 31 - VBA.Len(suffix)) & suffix
        If VBA.Len(candidate) = 0 Then candidate = "Page" & suffix
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
