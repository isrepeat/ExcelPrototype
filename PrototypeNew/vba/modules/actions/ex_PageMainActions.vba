Attribute VB_Name = "ex_PageMainActions"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const MODES_ROOT_REL_PATH As String = "modes"
Private Const MODE_PROFILES_FILE_SUFFIX As String = "Profiles.xml"
Private Const MODE_PICKER_CONTROL_NAME As String = "ConfigModePicker"
Private Const PROFILE_PICKER_CONTROL_NAME As String = "ConfigProfilePicker"
Private Const MODES_RUNTIME_KEY As String = "RuntimeItems.PageMain.ConfigModes"
Private Const PROFILES_RUNTIME_KEY As String = "RuntimeItems.PageMain.ConfigProfiles"
Private Const CONFIG_RUNTIME_KEY As String = "RuntimeItems.PageMain.Config"

Public Sub m_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:ex_PageMainActions.m_Module_Dispose"
#End If
End Sub
' //
' // API
' //
Public Sub m_OnConfigModeChanged()
    ' После смены режима нужно пересобрать runtime-источники.
    If Not m_PrepareModeProfileConfigRuntime(True) Then Exit Sub
End Sub


Public Sub m_OnConfigProfileChanged()
    ' После смены профиля также пересобираем runtime-состояние.
    If Not m_PrepareModeProfileConfigRuntime(True) Then Exit Sub
End Sub


Public Function m_PrepareModeProfileConfigRuntime( _
    Optional ByVal notifyChange As Boolean = False, _
    Optional ByVal preferredPageBase As Object _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim modeOptions As Collection
    Dim profileOptions As Collection
    Dim selectedModeId As String
    Dim selectedProfileId As String
    Dim profileFilePath As String
    Dim prepareStep As String

    On Error GoTo EH_PREPARE_RUNTIME
    ' 1) Резолвим контекст текущей страницы.
    prepareStep = "resolve-page-context"
    If Not private_TryResolvePageBase(pageBase, preferredPageBase) Then Exit Function
    If pageBase Is Nothing Then Exit Function

    prepareStep = "resolve-worksheet"
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then
        private_ReportRuntimeConfigError "PrototypeNew: worksheet is not specified for mode/profile config runtime prepare."
        Exit Function
    End If

    ' 2) Формируем источник для списка режимов.
    prepareStep = "build-mode-options"
    If Not private_TryBuildModeSelectOptions(modeOptions) Then Exit Function
    prepareStep = "register-mode-items-source"
    If Not private_TrySetItemsSource(MODES_RUNTIME_KEY, modeOptions, False, pageBase) Then Exit Function

    ' 3) Для выбранного режима собираем список профилей.
    prepareStep = "resolve-selected-mode"
    If Not private_TryResolveSelectedIdForControl(ws, MODE_PICKER_CONTROL_NAME, modeOptions, selectedModeId) Then Exit Function
    prepareStep = "build-profile-options"
    If Not private_TryBuildProfileSelectOptionsByMode(selectedModeId, profileOptions, profileFilePath) Then Exit Function
    prepareStep = "register-profile-items-source"
    If Not private_TrySetItemsSource(PROFILES_RUNTIME_KEY, profileOptions, False, pageBase) Then Exit Function

    ' 4) Загружаем config из выбранного профиля.
    prepareStep = "resolve-selected-profile"
    If Not private_TryResolveSelectedIdForControl(ws, PROFILE_PICKER_CONTROL_NAME, profileOptions, selectedProfileId) Then Exit Function
    prepareStep = "register-config-from-profile"
    If Not m_RegisterConfigFromXmlProfile(profileFilePath, selectedProfileId, False, pageBase) Then Exit Function

    ' 5) По флагу обновляем UI.
    If notifyChange Then
        prepareStep = "rerender-page"
        If Not private_TryRerenderPage(pageBase, "config:mode-profile-runtime") Then Exit Function
    End If

    m_PrepareModeProfileConfigRuntime = True
    On Error GoTo 0
    Exit Function

EH_PREPARE_RUNTIME:
    private_ReportRuntimeConfigError "PrototypeNew: exception in config runtime prepare at step '" & prepareStep & "': [" & VBA.CStr(Err.Number) & "] " & Err.Description
End Function


Public Function m_RegisterConfigFromXmlProfile( _
    ByVal filePath As String, _
    ByVal profileKey As String, _
    Optional ByVal notifyChange As Boolean = False, _
    Optional ByVal preferredPageBase As Object _
) As Boolean
    Dim normalizedFilePath As String
    Dim normalizedProfileKey As String
    Dim dom As Object
    Dim profileNode As Object
    Dim profileKeyLiteral As String
    Dim profileXPath As String

    ' Нормализуем и валидируем вход.
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

    ' Загружаем XML с профилями.
    Set dom = ex_XmlCore.m_LoadDomByFilePath( _
        normalizedFilePath, _
        "PrototypeNew: config profiles file was not found: ", _
        "PrototypeNew: failed to parse config profiles file: ", _
        VBA.vbNullString)
    If dom Is Nothing Then
        private_ReportRuntimeConfigError "PrototypeNew: failed to load config profiles file '" & normalizedFilePath & "'."
        Exit Function
    End If

    ' Ищем профиль по id/name/key (атрибуты и дочерние элементы).
    profileKeyLiteral = ex_XmlCore.m_XPathLiteral(normalizedProfileKey)
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

    Set profileNode = dom.selectSingleNode(profileXPath)
    If profileNode Is Nothing Then
        private_ReportRuntimeConfigError "PrototypeNew: config profile '" & normalizedProfileKey & "' was not found in file '" & normalizedFilePath & "'."
        Exit Function
    End If

    ' Делегируем преобразование узла в runtime-коллекцию конфига.
    If Not m_RegisterConfigFromProfileNode(profileNode, notifyChange, preferredPageBase) Then Exit Function
    m_RegisterConfigFromXmlProfile = True
End Function


Public Function m_RegisterConfigFromProfileNode( _
    ByVal profileNode As Object, _
    Optional ByVal notifyChange As Boolean = False, _
    Optional ByVal preferredPageBase As Object _
) As Boolean
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

    ' Парсим профиль в модель таблицы конфига.
    Set configTable = New obj_ConfigTable
    If Not configTable.TryLoadFromXmlNode(profileNode, True) Then
        private_ReportRuntimeConfigError "PrototypeNew: failed to parse selected config profile node."
        Exit Function
    End If

    ' Готовим Collection для runtime source map.
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

    ' Регистрируем источник для Config-контрола.
    If Not private_TrySetItemsSource(CONFIG_RUNTIME_KEY, sourceItems, notifyChange, preferredPageBase) Then Exit Function
    m_RegisterConfigFromProfileNode = True
End Function

' //
' // Internal
' //
Private Sub private_ReportRuntimeConfigError(ByVal messageText As String)
    messageText = VBA.Trim$(messageText)
    If VBA.Len(messageText) = 0 Then Exit Sub

    ' Ошибку фиксируем в логах и показываем пользователю.
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.m_Diagnostic_LogError messageText
#End If
    MsgBox messageText, vbExclamation, "PrototypeNew / Config runtime"
End Sub


Private Function private_GetModesRootFolderPath() As String
    ' Режимы всегда ожидаются рядом с Excel-файлом: <WorkbookPath>\modes
    private_GetModesRootFolderPath = VBA.Trim$(ex_XmlCore.m_CombineBasePath(ThisWorkbook, MODES_ROOT_REL_PATH))
End Function


Private Function private_IsSafeModeId(ByVal modeId As String) As Boolean
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


Private Function private_TrySetItemsSource( _
    ByVal sourceKey As String, _
    ByVal items As Collection, _
    ByVal notifyChange As Boolean, _
    Optional ByVal preferredPageBase As Object _
) As Boolean
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
    Dim pageRef As obj_IPage
    Dim ws As Worksheet

    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function

    ' Получаем страницу по worksheet и запускаем render.
    If Not rt_PageManager.m_TryGetPageByWorksheet(ws, pageRef) Then Exit Function
    private_TryRerenderPage = rt_PageManager.m_RenderPage(pageRef, reason)
End Function
