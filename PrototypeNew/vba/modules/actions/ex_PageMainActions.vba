Attribute VB_Name = "ex_PageMainActions"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const MODES_ROOT_REL_PATH As String = "modes"
Private Const MODE_PROFILES_FILE_SUFFIX As String = "Profiles.xml"
Private Const MODE_PICKER_CONTROL_NAME As String = "ConfigModePicker"
Private Const PROFILE_PICKER_CONTROL_NAME As String = "ConfigProfilePicker"
Private Const CONFIG_CONTROL_NAME As String = "DevConfig"
Private Const MODES_RUNTIME_KEY As String = "RuntimeItems.PageMain.ConfigModes"
Private Const PROFILES_RUNTIME_KEY As String = "RuntimeItems.PageMain.ConfigProfiles"
Private Const CONFIG_RUNTIME_KEY As String = "RuntimeItems.PageMain.Config"

Public Sub m_Module_Dispose()
    private_LogEnter "m_Module_Dispose"
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:ex_PageMainActions.m_Module_Dispose"
#End If
End Sub


' //
' // API
' //
Public Function m_OnConfigModeChanged( _
    Optional ByVal notifyChange As Boolean = True, _
    Optional ByVal preferredPageBase As Object _
)
    private_LogEnter "m_OnConfigModeChanged"
    ' После смены режима нужно пересобрать runtime-источники.
    If Not private_TryPrepareModeProfileConfigRuntime(notifyChange, preferredPageBase) Then Exit Function
    m_OnConfigModeChanged = True
End Function


Public Function m_OnConfigProfileChanged( _
    Optional ByVal notifyChange As Boolean = True, _
    Optional ByVal preferredPageBase As Object _
)
    private_LogEnter "m_OnConfigProfileChanged"
    ' После смены профиля также пересобираем runtime-состояние.
    If Not private_TryPrepareModeProfileConfigRuntime(notifyChange, preferredPageBase) Then Exit Function
    m_OnConfigProfileChanged = True
End Function


Public Sub m_SaveCurrentConfigProfile()
    private_LogEnter "m_SaveCurrentConfigProfile"
    If Not private_TrySaveCurrentConfigProfile() Then Exit Sub
End Sub


' //
' // Internal
' //
Private Function private_TryPrepareModeProfileConfigRuntime( _
    Optional ByVal notifyChange As Boolean = False, _
    Optional ByVal preferredPageBase As Object _
) As Boolean
    private_LogEnter "private_TryPrepareModeProfileConfigRuntime"
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
        private_ReportRuntimeConfigError "PrototypeNew: worksheet is not specified for mode/profile config runtime prepare."
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
    private_ReportRuntimeConfigError "PrototypeNew: exception in config runtime prepare: [" & VBA.CStr(Err.Number) & "] " & Err.Description
End Function


Private Function private_TrySaveCurrentConfigProfile( _
    Optional ByVal preferredPageBase As Object _
) As Boolean
    private_LogEnter "private_TrySaveCurrentConfigProfile"
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
        private_ReportRuntimeConfigError "PrototypeNew: worksheet is not specified for config profile save."
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
        private_ReportRuntimeConfigError "PrototypeNew: failed to read current entries from config control '" & CONFIG_CONTROL_NAME & "'."
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
        private_ReportRuntimeConfigError "PrototypeNew: failed to build source config node from control '" & CONFIG_CONTROL_NAME & "'."
        Exit Function
    End If
    If Not private_TryReplaceProfileRowsFromSourceNode(profileNode, generatedConfigNode) Then
        Exit Function
    End If
    If Not private_TrySaveDomToFile(dom, profileFilePath) Then Exit Function

    ' 6) После успешного Save обновляем in-memory source,
    ' чтобы следующий render взял именно то состояние, которое ушло в файл.
    If Not private_TrySetItemsSource(CONFIG_RUNTIME_KEY, configEntries, False, pageBase) Then Exit Function

    rt_Messaging.m_ShowStatusBarSuccess "Config profile '" & selectedProfileId & "' saved to '" & profileFilePath & "'.", 4
    private_TrySaveCurrentConfigProfile = True
    On Error GoTo 0
    Exit Function

EH_SAVE_PROFILE:
    private_ReportRuntimeConfigError "PrototypeNew: exception in config profile save: [" & VBA.CStr(Err.Number) & "] " & Err.Description
End Function





Private Function private_TryRegisterConfigFromXmlProfile( _
    ByVal filePath As String, _
    ByVal profileKey As String, _
    Optional ByVal notifyChange As Boolean = False, _
    Optional ByVal preferredPageBase As Object _
) As Boolean
    private_LogEnter "private_TryRegisterConfigFromXmlProfile"
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
    private_LogEnter "private_TryRegisterConfigFromProfileNode"
    Dim configTable As obj_ConfigTable
    Dim configEntries As list__obj_ConfigEntry
    Dim sourceItems As Collection
    Dim sourceConfigEntry As obj_ConfigEntry
    Dim i As Long

    ' Узел профиля обязателен.
    If profileNode Is Nothing Then
        private_ReportRuntimeConfigError "PrototypeNew: config profile node is not specified."
        Exit Function
    End If

    ' Парсим XML profile-node в typed-модель obj_ConfigTable/obj_ConfigEntry.
    Set configTable = New obj_ConfigTable
    If Not configTable.TryLoadFromXmlNode(profileNode, True) Then
        private_ReportRuntimeConfigError "PrototypeNew: failed to parse selected config profile node."
        Exit Function
    End If

    ' Готовим обычную Collection для runtime source map (совместимый формат источника).
    Set sourceItems = New Collection
    Set configEntries = configTable.Items
    If configEntries Is Nothing Then
        private_ReportRuntimeConfigError "PrototypeNew: config table entries collection is not initialized."
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
    private_LogEnter "private_TryLoadProfileDomAndNode"
    Dim normalizedFilePath As String
    Dim normalizedProfileKey As String
    Dim profileKeyLiteral As String
    Dim profileXPath As String

    Set outDom = Nothing
    Set outProfileNode = Nothing

    normalizedFilePath = VBA.Trim$(filePath)
    normalizedProfileKey = VBA.Trim$(profileKey)

    If VBA.Len(normalizedFilePath) = 0 Then
        private_ReportRuntimeConfigError "PrototypeNew: config profiles file path is empty."
        Exit Function
    End If
    If VBA.Len(normalizedProfileKey) = 0 Then
        private_ReportRuntimeConfigError "PrototypeNew: config profile key is empty."
        Exit Function
    End If

    ' 1) Загружаем XML документ режима (например <Mode>Profiles.xml).
    Set outDom = ex_XmlCore.m_LoadDomByFilePath( _
        normalizedFilePath, _
        "PrototypeNew: config profiles file was not found: ", _
        "PrototypeNew: failed to parse config profiles file: ", _
        VBA.vbNullString)
    If outDom Is Nothing Then
        private_ReportRuntimeConfigError "PrototypeNew: failed to load config profiles file '" & normalizedFilePath & "'."
        Exit Function
    End If

    profileKeyLiteral = ex_XmlCore.m_XPathLiteral(normalizedProfileKey)
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
        private_ReportRuntimeConfigError "PrototypeNew: config profile '" & normalizedProfileKey & "' was not found in file '" & normalizedFilePath & "'."
        Exit Function
    End If

    private_TryLoadProfileDomAndNode = True
End Function


Private Function private_GetModesRootFolderPath() As String
    private_LogEnter "private_GetModesRootFolderPath"
    ' Режимы всегда ожидаются рядом с Excel-файлом: <WorkbookPath>\modes
    private_GetModesRootFolderPath = VBA.Trim$(ex_XmlCore.m_CombineBasePath(ThisWorkbook, MODES_ROOT_REL_PATH))
End Function


Private Function private_IsSafeModeId(ByVal modeId As String) As Boolean
    private_LogEnter "private_IsSafeModeId"
    ' Минимальная защита от path traversal и инъекций в имя режима.
    modeId = VBA.Trim$(modeId)
    If VBA.Len(modeId) = 0 Then Exit Function
    If VBA.InStr(1, modeId, "\", VBA.vbBinaryCompare) > 0 Then Exit Function
    If VBA.InStr(1, modeId, "/", VBA.vbBinaryCompare) > 0 Then Exit Function
    If VBA.InStr(1, modeId, ":", VBA.vbBinaryCompare) > 0 Then Exit Function
    If VBA.InStr(1, modeId, "..", VBA.vbBinaryCompare) > 0 Then Exit Function

    private_IsSafeModeId = True
End Function


Private Function private_BuildModeProfilesFilePath(ByVal modeId As String) As String
    private_LogEnter "private_BuildModeProfilesFilePath"
    Dim modesRootPath As String

    modeId = VBA.Trim$(modeId)
    ' Проверяем, что mode id безопасен для подстановки в путь.
    If Not private_IsSafeModeId(modeId) Then
        private_ReportRuntimeConfigError "PrototypeNew: unsafe mode id '" & modeId & "'."
        Exit Function
    End If

    modesRootPath = private_GetModesRootFolderPath()
    If VBA.Len(modesRootPath) = 0 Then
        private_ReportRuntimeConfigError "PrototypeNew: failed to resolve modes root path."
        Exit Function
    End If

    private_BuildModeProfilesFilePath = modesRootPath & "\" & modeId & "\" & modeId & MODE_PROFILES_FILE_SUFFIX
End Function


Private Function private_TryBuildModeSelectOptions(ByRef outOptions As Collection) As Boolean
    private_LogEnter "private_TryBuildModeSelectOptions"
    Dim modesRootPath As String
    Dim folderName As String
    Dim folderPath As String
    Dim folderAttr As Long
    Dim optionItem As obj_SelectOption

    ' Собираем SelectOption по подпапкам modes.
    Set outOptions = New Collection
    modesRootPath = private_GetModesRootFolderPath()
    If VBA.Len(modesRootPath) = 0 Then
        private_ReportRuntimeConfigError "PrototypeNew: failed to resolve modes root path."
        Exit Function
    End If

    If VBA.Len(Dir$(modesRootPath, vbDirectory)) = 0 Then
        private_ReportRuntimeConfigError "PrototypeNew: modes folder was not found: " & modesRootPath
        Exit Function
    End If

    folderName = Dir$(modesRootPath & "\*", vbDirectory)
    Do While VBA.Len(folderName) > 0
        If VBA.StrComp(folderName, ".", VBA.vbBinaryCompare) <> 0 And VBA.StrComp(folderName, "..", VBA.vbBinaryCompare) <> 0 Then
            folderPath = modesRootPath & "\" & folderName
            On Error Resume Next
            folderAttr = GetAttr(folderPath)
            If Err.Number = 0 Then
                If (folderAttr And vbDirectory) = vbDirectory Then
                    Set optionItem = private_CreateSelectOption(folderName, folderName, "ex_PageMainActions.m_OnConfigModeChanged")
                    outOptions.Add optionItem
                End If
            End If
            Err.Clear
            On Error GoTo 0
        End If

        folderName = Dir$
    Loop

    If outOptions.Count = 0 Then
        private_ReportRuntimeConfigError "PrototypeNew: no mode folders were found under '" & modesRootPath & "'."
        Exit Function
    End If

    private_TryBuildModeSelectOptions = True
End Function


Private Function private_TryBuildProfileSelectOptionsByMode( _
    ByVal modeId As String, _
    ByRef outOptions As Collection, _
    ByRef outProfilesFilePath As String _
) As Boolean
    private_LogEnter "private_TryBuildProfileSelectOptionsByMode"
    Dim dom As Object

    ' Строим путь <ModeName>\<ModeName>Profiles.xml.
    Set outOptions = Nothing
    outProfilesFilePath = VBA.Trim$(private_BuildModeProfilesFilePath(modeId))
    If VBA.Len(outProfilesFilePath) = 0 Then Exit Function

    If VBA.Len(Dir$(outProfilesFilePath)) = 0 Then
        private_ReportRuntimeConfigError "PrototypeNew: profiles file was not found for mode '" & modeId & "': " & outProfilesFilePath
        Exit Function
    End If

    ' Загружаем XML и извлекаем профили в select options.
    Set dom = ex_XmlCore.m_LoadDomByFilePath( _
        outProfilesFilePath, _
        "PrototypeNew: profiles file was not found: ", _
        "PrototypeNew: failed to parse profiles file: ", _
        VBA.vbNullString)
    If dom Is Nothing Then
        private_ReportRuntimeConfigError "PrototypeNew: failed to load profiles file '" & outProfilesFilePath & "'."
        Exit Function
    End If

    If Not private_TryCollectProfileSelectOptionsFromDom(dom, outOptions) Then Exit Function
    If outOptions Is Nothing Then Exit Function
    If outOptions.Count = 0 Then
        private_ReportRuntimeConfigError "PrototypeNew: profiles file '" & outProfilesFilePath & "' does not contain selectable profiles."
        Exit Function
    End If

    private_TryBuildProfileSelectOptionsByMode = True
End Function


Private Function private_TryCollectProfileSelectOptionsFromDom(ByVal dom As Object, ByRef outOptions As Collection) As Boolean
    private_LogEnter "private_TryCollectProfileSelectOptionsFromDom"
    Dim profileNodes As Object
    Dim profileNode As Object
    Dim seenIds As Object
    Dim optionItem As obj_SelectOption
    Dim optionId As String

    Set outOptions = New Collection
    Set seenIds = VBA.CreateObject("Scripting.Dictionary")
    seenIds.CompareMode = 1

    ' Шаг 1: профильные узлы ожидаемых имен.
    If dom Is Nothing Then
        private_ReportRuntimeConfigError "PrototypeNew: profiles DOM is not specified."
        Exit Function
    End If

    On Error GoTo EH_XML
    Set profileNodes = dom.selectNodes("//*[local-name()='profile' or local-name()='configProfile' or local-name()='preset' or local-name()='variant']")
    On Error GoTo 0

    If Not profileNodes Is Nothing Then
        For Each profileNode In profileNodes
            Set optionItem = Nothing
            If Not private_TryCreateProfileSelectOptionFromNode(profileNode, optionItem) Then Exit Function
            If optionItem Is Nothing Then GoTo ContinueProfileNodePrimary

            optionId = VBA.LCase$(VBA.Trim$(optionItem.Id))
            If VBA.Len(optionId) = 0 Then GoTo ContinueProfileNodePrimary
            If seenIds.Exists(optionId) Then GoTo ContinueProfileNodePrimary

            seenIds(optionId) = True
            outOptions.Add optionItem
ContinueProfileNodePrimary:
        Next profileNode
    End If

    If outOptions.Count > 0 Then
        private_TryCollectProfileSelectOptionsFromDom = True
        Exit Function
    End If

    ' Шаг 2 (fallback): любые узлы с id/name/key и config-строками.
    On Error GoTo EH_XML
    Set profileNodes = dom.selectNodes("//*[" & _
                                      "(@id or @name or @key)" & _
                                      " and " & _
                                      "(.//*[local-name()='item' or local-name()='row' or local-name()='entry' or local-name()='config']" & _
                                      " or *[local-name()='item' or local-name()='row' or local-name()='entry' or local-name()='config'])" & _
                                      "]")
    On Error GoTo 0

    If Not profileNodes Is Nothing Then
        For Each profileNode In profileNodes
            Set optionItem = Nothing
            If Not private_TryCreateProfileSelectOptionFromNode(profileNode, optionItem) Then Exit Function
            If optionItem Is Nothing Then GoTo ContinueProfileNodeFallback

            optionId = VBA.LCase$(VBA.Trim$(optionItem.Id))
            If VBA.Len(optionId) = 0 Then GoTo ContinueProfileNodeFallback
            If seenIds.Exists(optionId) Then GoTo ContinueProfileNodeFallback

            seenIds(optionId) = True
            outOptions.Add optionItem
ContinueProfileNodeFallback:
        Next profileNode
    End If

    private_TryCollectProfileSelectOptionsFromDom = True
    Exit Function

EH_XML:
    private_ReportRuntimeConfigError "PrototypeNew: failed to read profile list from XML: " & Err.Description
End Function


Private Function private_TryCreateProfileSelectOptionFromNode( _
    ByVal profileNode As Object, _
    ByRef outOption As obj_SelectOption _
) As Boolean
    private_LogEnter "private_TryCreateProfileSelectOptionFromNode"
    Dim profileId As String
    Dim captionText As String
    Dim onSelectText As String

    ' Читаем id профиля из атрибутов/дочерних узлов.
    Set outOption = Nothing
    If profileNode Is Nothing Then
        private_TryCreateProfileSelectOptionFromNode = True
        Exit Function
    End If

    profileId = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(profileNode, "id")))
    If VBA.Len(profileId) = 0 Then profileId = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(profileNode, "name")))
    If VBA.Len(profileId) = 0 Then profileId = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(profileNode, "key")))
    If VBA.Len(profileId) = 0 Then
        If Not private_TryReadChildNodeText(profileNode, "id", profileId) Then Exit Function
    End If
    If VBA.Len(profileId) = 0 Then
        If Not private_TryReadChildNodeText(profileNode, "name", profileId) Then Exit Function
    End If
    If VBA.Len(profileId) = 0 Then
        If Not private_TryReadChildNodeText(profileNode, "key", profileId) Then Exit Function
    End If
    If VBA.Len(profileId) = 0 Then
        private_TryCreateProfileSelectOptionFromNode = True
        Exit Function
    End If

    ' Читаем подпись для UI (caption/title/display/name).
    captionText = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(profileNode, "caption")))
    If VBA.Len(captionText) = 0 Then captionText = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(profileNode, "title")))
    If VBA.Len(captionText) = 0 Then captionText = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(profileNode, "display")))
    If VBA.Len(captionText) = 0 Then captionText = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(profileNode, "name")))
    If VBA.Len(captionText) = 0 Then
        If Not private_TryReadChildNodeText(profileNode, "caption", captionText) Then Exit Function
    End If
    If VBA.Len(captionText) = 0 Then
        If Not private_TryReadChildNodeText(profileNode, "title", captionText) Then Exit Function
    End If
    If VBA.Len(captionText) = 0 Then
        If Not private_TryReadChildNodeText(profileNode, "display", captionText) Then Exit Function
    End If
    If VBA.Len(captionText) = 0 Then
        If Not private_TryReadChildNodeText(profileNode, "name", captionText) Then Exit Function
    End If
    If VBA.Len(captionText) = 0 Then captionText = profileId

    ' При выборе профиля перезапускаем runtime pipeline.
    onSelectText = "ex_PageMainActions.m_OnConfigProfileChanged"
    Set outOption = private_CreateSelectOption(captionText, profileId, onSelectText)
    private_TryCreateProfileSelectOptionFromNode = True
End Function


Private Function private_TryReadChildNodeText( _
    ByVal parentNode As Object, _
    ByVal childLocalName As String, _
    ByRef outText As String _
) As Boolean
    private_LogEnter "private_TryReadChildNodeText"
    Dim childNode As Object

    ' Доступ к child через local-name(), чтобы не зависеть от namespace.
    outText = VBA.vbNullString
    If parentNode Is Nothing Then
        private_TryReadChildNodeText = True
        Exit Function
    End If

    On Error GoTo EH_XML
    Set childNode = parentNode.selectSingleNode("./*[local-name()='" & childLocalName & "']")
    On Error GoTo 0

    If Not childNode Is Nothing Then outText = VBA.Trim$(VBA.CStr(childNode.Text))
    private_TryReadChildNodeText = True
    Exit Function

EH_XML:
    private_ReportRuntimeConfigError "PrototypeNew: failed to read child node '" & childLocalName & "': " & Err.Description
End Function


Private Function private_TryResolveSelectedIdForControl( _
    ByVal ws As Worksheet, _
    ByVal controlName As String, _
    ByVal options As Collection, _
    ByRef outSelectedId As String _
) As Boolean
    private_LogEnter "private_TryResolveSelectedIdForControl"
    Dim storedId As String
    Dim firstId As String

    ' Сначала пробуем восстановить выбранный id из state store.
    outSelectedId = VBA.vbNullString
    If ws Is Nothing Then
        private_ReportRuntimeConfigError "PrototypeNew: worksheet is not specified for selectedId resolve."
        Exit Function
    End If
    If options Is Nothing Then
        private_ReportRuntimeConfigError "PrototypeNew: options are not specified for control '" & controlName & "'."
        Exit Function
    End If
    If options.Count = 0 Then
        private_ReportRuntimeConfigError "PrototypeNew: options are empty for control '" & controlName & "'."
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
        private_ReportRuntimeConfigError "PrototypeNew: failed to resolve first option id for control '" & controlName & "'."
        Exit Function
    End If

    outSelectedId = firstId
    If Not private_TrySetStoredSelectedIdForControl(ws, controlName, outSelectedId) Then Exit Function
    private_TryResolveSelectedIdForControl = True
End Function


Private Function private_SelectOptionsContainsId(ByVal options As Collection, ByVal optionId As String) As Boolean
    private_LogEnter "private_SelectOptionsContainsId"
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
    private_LogEnter "private_TryGetFirstOptionId"
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
    private_LogEnter "private_TryGetStoredSelectedIdForControl"
    Dim selectStatic As obj_SelectControlVMStatic
    Dim selectKey As String

    outSelectedId = VBA.vbNullString
    If ws Is Nothing Then Exit Function

    ' Ключ хранения: "<SheetName>|<ControlName>".
    selectKey = VBA.LCase$(VBA.Trim$(ws.Name) & "|" & VBA.Trim$(controlName))
    If VBA.Len(selectKey) = 0 Then
        private_ReportRuntimeConfigError "PrototypeNew: select key is empty for control '" & controlName & "'."
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
    private_LogEnter "private_TrySetStoredSelectedIdForControl"
    Dim selectStatic As obj_SelectControlVMStatic
    Dim selectKey As String

    If ws Is Nothing Then Exit Function

    ' Ключ хранения: "<SheetName>|<ControlName>".
    selectKey = VBA.LCase$(VBA.Trim$(ws.Name) & "|" & VBA.Trim$(controlName))
    If VBA.Len(selectKey) = 0 Then
        private_ReportRuntimeConfigError "PrototypeNew: select key is empty for control '" & controlName & "'."
        Exit Function
    End If

    Set selectStatic = New obj_SelectControlVMStatic
    private_TrySetStoredSelectedIdForControl = selectStatic.SetSelectedId(selectKey, VBA.Trim$(selectedId))
End Function


Private Function private_CreateSelectOption( _
    ByVal captionText As String, _
    ByVal idText As String, _
    ByVal onSelectMacro As String _
) As obj_SelectOption
    private_LogEnter "private_CreateSelectOption"
    Dim selectOption As obj_SelectOption

    ' Формируем объект, который читает obj_SelectControlVM.
    Set selectOption = New obj_SelectOption
    selectOption.Caption = VBA.CStr(captionText)
    selectOption.Id = VBA.CStr(idText)
    selectOption.OnSelect = VBA.CStr(onSelectMacro)

    Set private_CreateSelectOption = selectOption
End Function


Private Function private_TryResolvePageBase( _
    ByRef outPageBase As obj_PageBase, _
    Optional ByVal preferredPageBase As Object _
) As Boolean
    private_LogEnter "private_TryResolvePageBase"
    Set outPageBase = Nothing

    ' Приоритет: явно переданный page context.
    If Not preferredPageBase Is Nothing Then
        If TypeOf preferredPageBase Is obj_PageBase Then
            Set outPageBase = preferredPageBase
            private_TryResolvePageBase = True
            Exit Function
        End If

        If ex_HelpersSheet.m_TryCastPageBase(preferredPageBase, outPageBase) Then
            private_TryResolvePageBase = True
            Exit Function
        End If

        private_ReportRuntimeConfigError "PrototypeNew: preferred page runtime context has unsupported type '" & VBA.TypeName(preferredPageBase) & "'."
        Exit Function
    End If

    ' Fallback: активная страница из runtime.
    If Not ex_HelpersSheet.m_TryGetActivePageBase(outPageBase) Then
        private_ReportRuntimeConfigError "PrototypeNew: page runtime context is not resolved for active worksheet."
        Exit Function
    End If
    If outPageBase Is Nothing Then Exit Function

    private_TryResolvePageBase = True
End Function


Private Function private_TryResolveConfigControl( _
    ByVal pageBase As obj_PageBase, _
    ByRef outConfigControl As obj_ConfigControlVM _
) As Boolean
    private_LogEnter "private_TryResolveConfigControl"
    Dim rawControl As Object

    Set outConfigControl = Nothing
    If pageBase Is Nothing Then
        private_ReportRuntimeConfigError "PrototypeNew: page base is not specified for config control resolve."
        Exit Function
    End If

    If Not pageBase.TryGetRegisteredControlByName(CONFIG_CONTROL_NAME, rawControl) Then
        private_ReportRuntimeConfigError "PrototypeNew: config control '" & CONFIG_CONTROL_NAME & "' was not found in runtime registry."
        Exit Function
    End If
    If rawControl Is Nothing Then
        private_ReportRuntimeConfigError "PrototypeNew: config control '" & CONFIG_CONTROL_NAME & "' runtime entry is empty."
        Exit Function
    End If
    If Not TypeOf rawControl Is obj_ConfigControlVM Then
        private_ReportRuntimeConfigError "PrototypeNew: config control '" & CONFIG_CONTROL_NAME & "' has unexpected type '" & VBA.TypeName(rawControl) & "'."
        Exit Function
    End If

    Set outConfigControl = rawControl
    private_TryResolveConfigControl = True
End Function


Private Function private_TryReplaceProfileRowsFromSourceNode( _
    ByVal targetProfileNode As Object, _
    ByVal sourceConfigNode As Object _
) As Boolean
    private_LogEnter "private_TryReplaceProfileRowsFromSourceNode"
    Dim targetRowNodes As Object
    Dim sourceRowNodes As Object
    Dim rowIndex As Long
    Dim sourceRowNode As Object
    Dim clonedRowNode As Object

    If targetProfileNode Is Nothing Then
        private_ReportRuntimeConfigError "PrototypeNew: target profile node is not specified for save."
        Exit Function
    End If
    If sourceConfigNode Is Nothing Then
        private_ReportRuntimeConfigError "PrototypeNew: source config node is not specified for save."
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
        private_ReportRuntimeConfigError "PrototypeNew: source config node does not contain readable rows for save."
        Exit Function
    End If
    If sourceRowNodes.Length = 0 Then
        private_ReportRuntimeConfigError "PrototypeNew: source config node is empty and cannot replace profile rows."
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
            private_ReportRuntimeConfigError "PrototypeNew: failed to clone source row node while updating profile."
            Exit Function
        End If
        targetProfileNode.appendChild clonedRowNode
ContinueSourceRow:
    Next rowIndex

    private_TryReplaceProfileRowsFromSourceNode = True
    Exit Function

EH_XML:
    private_ReportRuntimeConfigError "PrototypeNew: failed to transfer source config rows into profile node: " & Err.Description
End Function


Private Function private_TrySaveDomToFile(ByVal dom As Object, ByVal filePath As String) As Boolean
    private_LogEnter "private_TrySaveDomToFile"
    filePath = VBA.Trim$(filePath)
    If dom Is Nothing Then
        private_ReportRuntimeConfigError "PrototypeNew: DOM is not specified for file save."
        Exit Function
    End If
    If VBA.Len(filePath) = 0 Then
        private_ReportRuntimeConfigError "PrototypeNew: file path is empty for profile save."
        Exit Function
    End If

    On Error GoTo EH_SAVE_DOM
    dom.Save filePath
    private_TrySaveDomToFile = True
    Exit Function

EH_SAVE_DOM:
    private_ReportRuntimeConfigError "PrototypeNew: failed to write profile file '" & filePath & "': " & Err.Description
End Function


Private Function private_TrySetItemsSource( _
    ByVal sourceKey As String, _
    ByVal items As Collection, _
    ByVal notifyChange As Boolean, _
    Optional ByVal preferredPageBase As Object _
) As Boolean
    private_LogEnter "private_TrySetItemsSource"
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
    private_LogEnter "private_TryRerenderPage"
    Dim pageRef As obj_IPage
    Dim ws As Worksheet

    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function

    ' Получаем страницу по worksheet и запускаем render.
    If Not rt_PageManager.m_TryGetPageByWorksheet(ws, pageRef) Then Exit Function
    private_TryRerenderPage = rt_PageManager.m_RenderPage(pageRef, reason)
End Function

Private Sub private_LogEnter(ByVal memberName As String)
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "enter:ex_PageMainActions." & VBA.Trim$(memberName)
#End If
End Sub

Private Sub private_ReportRuntimeConfigError(ByVal messageText As String)
    private_LogEnter "private_ReportRuntimeConfigError"
    messageText = VBA.Trim$(messageText)
    If VBA.Len(messageText) = 0 Then Exit Sub

    ' Ошибку фиксируем в логах и показываем пользователю.
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.m_Diagnostic_LogError messageText
#End If
    MsgBox messageText, vbExclamation, "PrototypeNew / Config runtime"
End Sub
