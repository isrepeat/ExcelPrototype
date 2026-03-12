Attribute VB_Name = "ex_ConfigProfilesManager"
Option Explicit

' =============================================================================
' ex_ConfigProfilesManager
' =============================================================================
' Назначение:
' - читать/писать профили конфигурации из внешнего XML-файла
'   (пути задаются в `config\DevUI.xml` -> `dataSources/itemsSource[name='profilesFileByMode']`);
' - применять выбранный профиль к таблице `tblDevConfig` на листе Dev;
' - сохранять текущее состояние таблицы обратно в активный профиль;
' - поддерживать совместимость со старыми форматами (legacy row/marker layout);
' - синхронизировать визуальное оформление таблицы после загрузки профиля.
'
' Границы ответственности:
' - этот модуль работает с профильным XML и состоянием выбранных mode/profile;
' - бизнес-логика чтения значений конфигурации по ключу находится в ex_ConfigProvider.
' =============================================================================

Private Const PROFILES_NS As String = "urn:excelprototype:profiles"
Private Const DEV_CONFIG_TABLE_NAME As String = "tblDevConfig"
Private Const DEV_CONFIG_MARKER_COL As Long = 1
Private Const DEV_CONFIG_KEY_COL As Long = 2
Private Const DEV_CONFIG_VALUE_COL As Long = 3
Private Const DEV_CONFIG_STYLES_COL As Long = 4
Private Const DEV_CONFIG_COL_COUNT As Long = 4
Private Const DEV_MARKER_SYMBOL As String = "#"
Private Const DEV_MARKER_HEADER As String = ".."
Private Const DEV_MARKER_PREFIX As String = "#MARKER:"
Private Const DEV_MARKER_SECTION As String = "#MARKER:SECTION"
Private Const DEV_MARKER_SPACER As String = "#MARKER:SPACER"
Private Const DEV_COLOR_BG As Long = &H1E1E1E
Private Const DEV_COLOR_TEXT As Long = &HEBEBEB
Private Const DEV_COLOR_BORDER As Long = &H505050
Private Const THEME_BG As Long = &H262626
Private Const THEME_TEXT As Long = &HEBEBEB
Private Const THEME_BORDER As Long = &H0
Private Const PROFILE_CFG_COL_MARKER As String = "marker"
Private Const PROFILE_CFG_COL_KEY As String = "key"
Private Const PROFILE_CFG_COL_VALUE As String = "value"
Private Const PROFILE_CFG_COL_STYLES As String = "styles"
Private Const PFUI_UI_DEFINITION_REL_PATH As String = "config\DevUI.xml"
Private Const PFUI_UI_BLOCK_GROUP_NAME As String = "grpUiBlock"
Private Const PFUI_PROFILE_DROPDOWN_SHAPE As String = "btnCustomProfile"
Private Const PFUI_MODE_DROPDOWN_SHAPE As String = "btnCustomMode"
Private Const LEGACY_PROFILE_DROPDOWN_SHAPE As String = "ddProfile"
Private Const LEGACY_MODE_DROPDOWN_SHAPE As String = "ddMode"
Private Const PFUI_UPDATE_BUTTON_SHAPE As String = "btnUpdateCode"
Private Const PFUI_UPDATE_UI_BUTTON_SHAPE As String = "btnUpdateUI"
Private Const PFUI_CLEAR_BUTTON_SHAPE As String = "btnClear"
Private Const PFUI_MODE_BUTTON_SHAPE As String = "btnMode"
Private Const PFUI_PERSONAL_BUTTON_SHAPE As String = "btnPersonalCard"
Private Const PFUI_COMPARING_BUTTON_SHAPE As String = "btnComparing"
Private Const STATE_ACTIVE_MODE_KEY_PROP As String = "Settings.ActiveModeKey"
Private Const STATE_ACTIVE_PROFILE_PROP_PREFIX As String = "Settings.ActiveProfile."
Private Const STATE_PROFILE_FILE_SIGNATURE_PROP_PREFIX As String = "Settings.ProfileFileSignature."
Private Const XML_ATTR_LOCKED_WITH_PLACEHOLDER As String = "lockedWithPlaceholder"

' =============================================================================
' Public API (сверху по требованию рефакторинга)
' =============================================================================

' Применяет выбранный профиль в таблицу на листе Dev.
' Последовательность:
' 1) определяет активный профиль (аргумент или сохранённый activeProfile в state);
' 2) загружает DOM профилей из XML;
' 3) читает узлы профиля в внутренний массив строк;
' 4) перезаписывает таблицу `tblDevConfig` и обновляет заголовок профиля.
Public Sub m_ApplyProfileFromDev(Optional ByVal profileName As String = vbNullString)
    Dim ws As Worksheet
    Dim doc As Object
    Dim profileNode As Object
    Dim entries As Variant
    Dim lockedWithPlaceholder As Object
    Dim cfgTable As ListObject
    Dim profiles As Variant
    Dim prevEvents As Boolean
    Dim targetStableZoneLeft As Double
    Dim stepName As String

    On Error GoTo EH
    prevEvents = Application.EnableEvents
    stepName = "resolve-sheet"

    Set ws = ws_Dev
    targetStableZoneLeft = ex_CustomDropdown.m_GetStableZoneStartLeft(ws)

    stepName = "resolve-profile-name"
    If Len(profileName) = 0 Then
        profiles = mp_GetProfileNames(ws)
        If mp_ArrayHasItems(profiles) Then
            profileName = mp_GetSelectedProfileNameFromDropdown(ws, profiles, False)
        End If
    End If
    profileName = Trim$(profileName)
    If Len(profileName) = 0 Then Exit Sub

    stepName = "load-profile-dom"
    Set doc = mp_LoadProfilesDom(ws)
    stepName = "resolve-profile-node"
    Set profileNode = mp_GetProfileNode(doc, profileName, False)
    If profileNode Is Nothing Then
        MsgBox "Profile '" & profileName & "' was not found in config file: " & mp_GetProfilesFilePath(), vbExclamation
        Exit Sub
    End If

    stepName = "read-profile-entries"
    entries = ex_ProfilesEntriesMapper.m_ReadProfileEntries(ws, profileNode)
    Set lockedWithPlaceholder = mp_ReadLockedWithPlaceholder(profileNode)

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    stepName = "write-config-table"
    mp_WriteEntriesToConfigTable ws, entries

    stepName = "apply-config-styles"
    If Not mp_ApplyProfileConfigStyles(ws, profileNode, targetStableZoneLeft) Then GoTo EH

    stepName = "apply-locked-placeholders"
    Set cfgTable = ex_ConfigTableStore.m_GetConfigTable(ws, True)
    If cfgTable Is Nothing Then
        MsgBox "Config table '" & DEV_CONFIG_TABLE_NAME & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        GoTo EH
    End If
    mp_ApplyLockedPlaceholderCells ws, cfgTable, lockedWithPlaceholder

    stepName = "refresh-title"
    On Error Resume Next
    ex_ConfigProvider.m_RefreshConfigTitle ws, profileName
    On Error GoTo 0
    stepName = "apply-profile-ui"
    ex_ConfigProfilesManager.m_ApplyProfileUI ws, profileNode, profileName
    stepName = "apply-mode-visibility"
    mp_ApplyModeVisibility ws
    stepName = "save-profile-file-signature"
    mp_SaveAppliedProfileFileSignature ws, profileName
EH:
    If Err.Number <> 0 Then
        MsgBox "Failed to apply profile '" & profileName & "' at step '" & stepName & "': [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description, vbExclamation
        Err.Clear
    End If
    Application.ScreenUpdating = True
    Application.EnableEvents = prevEvents
End Sub

' UI-обёртка для кнопки "Apply profile" на Dev листе.
Public Sub m_ApplyProfile_Button()
    m_ApplyProfileFromDev
End Sub

' Сохраняет изменения из таблицы Dev в активный профиль.
' Сохранение выполняется только если изменённый диапазон пересекается с DataBodyRange таблицы.
' Это защищает от лишних XML-записей при несвязанных изменениях на листе.
Public Sub m_SaveEditsToProfile(ByVal ws As Worksheet, ByVal targetRange As Range, Optional ByVal profileName As String = vbNullString)
    Dim editRange As Range
    Dim dataRange As Range
    Dim doc As Object
    Dim profileNode As Object
    Dim tbl As ListObject
    Dim profiles As Variant

    Set tbl = ex_ConfigTableStore.m_GetConfigTable(ws, True)
    If tbl Is Nothing Then
        MsgBox "Config table '" & DEV_CONFIG_TABLE_NAME & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If

    If Len(profileName) = 0 Then
        profiles = mp_GetProfileNames(ws)
        If mp_ArrayHasItems(profiles) Then
            profileName = mp_GetSelectedProfileNameFromDropdown(ws, profiles, False)
        End If
    End If
    profileName = Trim$(profileName)
    If Len(profileName) = 0 Then Exit Sub

    Set dataRange = tbl.DataBodyRange
    If dataRange Is Nothing Then Exit Sub
    Set editRange = Intersect(targetRange, dataRange)
    If editRange Is Nothing Then Exit Sub

    Set doc = mp_LoadProfilesDom(ws)
    Set profileNode = mp_GetProfileNode(doc, profileName, True)
    If profileNode Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    ex_ProfilesEntriesMapper.m_WriteSheetValuesToProfile ws, doc, profileNode

    mp_SaveProfilesDom doc
    m_RefreshProfileValidation ws

    Application.ScreenUpdating = True
End Sub

' Legacy D1 validation не используется.
' Метод оставлен как стабильная точка вызова для старых мест интеграции.
Public Sub m_RefreshProfileValidation(Optional ByVal ws As Worksheet)
    If ws Is Nothing Then
        Set ws = ws_Dev
    End If
End Sub

Public Function m_GetActiveModeKey(Optional ByVal ws As Worksheet) As String
    Dim modeKey As String
    Dim defaultModeKey As String

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If

    On Error Resume Next
    modeKey = Trim$(mp_GetSelectedModeKey(ws))
    On Error GoTo 0

    If Len(modeKey) = 0 Then
        modeKey = Trim$(mp_GetStatePropText(STATE_ACTIVE_MODE_KEY_PROP, vbNullString))
    End If
    If Len(modeKey) = 0 Then
        defaultModeKey = Trim$(ex_UiXmlProvider.m_GetDefaultModeKey(ThisWorkbook))
        If Len(defaultModeKey) > 0 Then
            modeKey = defaultModeKey
        Else
            modeKey = Trim$(ex_UiXmlProvider.m_GetModeKeyByIndex(1, ThisWorkbook))
        End If
    End If

    m_GetActiveModeKey = modeKey
End Function

Public Function m_GetActiveModeName(Optional ByVal ws As Worksheet) As String
    Dim modeName As String
    Dim modeKey As String

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If

    On Error Resume Next
    modeName = Trim$(mp_GetSelectedModeName(ws))
    On Error GoTo 0

    modeKey = m_GetActiveModeKey(ws)

    If Len(modeName) = 0 Then
        modeName = Trim$(ex_UiXmlProvider.m_GetDropdownItemCaptionByKey(PFUI_MODE_DROPDOWN_SHAPE, modeKey, ThisWorkbook))
    End If
    If Len(modeName) = 0 Then
        modeName = modeKey
    End If

    m_GetActiveModeName = modeName
End Function

Public Function m_GetActiveProfileName(Optional ByVal ws As Worksheet) As String
    Dim profiles As Variant
    Dim profileName As String

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If

    profiles = mp_GetProfileNames(ws)
    If Not mp_ArrayHasItems(profiles) Then Exit Function

    profileName = Trim$(mp_GetSelectedProfileNameFromDropdown(ws, profiles, False))
    If Len(profileName) = 0 Then
        profileName = Trim$(CStr(profiles(LBound(profiles))))
    End If

    m_GetActiveProfileName = profileName
End Function

Public Function m_GetActiveProfileAttribute( _
    ByVal attrName As String, _
    Optional ByVal defaultValue As String = vbNullString, _
    Optional ByVal ws As Worksheet _
) As String
    Dim profileName As String
    Dim doc As Object
    Dim profileNode As Object
    Dim attrValue As String

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If

    attrName = Trim$(attrName)
    If Len(attrName) = 0 Then
        m_GetActiveProfileAttribute = defaultValue
        Exit Function
    End If

    profileName = Trim$(m_GetActiveProfileName(ws))
    If Len(profileName) = 0 Then
        m_GetActiveProfileAttribute = defaultValue
        Exit Function
    End If

    Set doc = mp_LoadProfilesDom(ws)
    If doc Is Nothing Then
        m_GetActiveProfileAttribute = defaultValue
        Exit Function
    End If

    Set profileNode = mp_GetProfileNode(doc, profileName, False)
    If profileNode Is Nothing Then
        m_GetActiveProfileAttribute = defaultValue
        Exit Function
    End If

    attrValue = Trim$(mp_NodeAttrText(profileNode, attrName))
    If Len(attrValue) = 0 Then attrValue = defaultValue
    m_GetActiveProfileAttribute = attrValue
End Function

Public Sub m_SetActiveModeKey(ByVal modeKey As String, Optional ByVal ws As Worksheet)
    Dim resolvedModeKey As String

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If

    modeKey = Trim$(modeKey)
    If Len(modeKey) = 0 Then Exit Sub

    resolvedModeKey = modeKey
    If Not mp_IsModeKeyValid(resolvedModeKey) Then
        MsgBox "Mode key '" & modeKey & "' is not available in control '" & PFUI_MODE_DROPDOWN_SHAPE & "'.", vbExclamation
        Exit Sub
    End If

    mp_SetStatePropText STATE_ACTIVE_MODE_KEY_PROP, resolvedModeKey
    ex_UiXmlProvider.m_SetDropdownContextValue "activeMode", resolvedModeKey
End Sub

Public Sub m_SetActiveProfileName(ByVal profileName As String, Optional ByVal modeKey As String = vbNullString, Optional ByVal ws As Worksheet)
    Dim resolvedModeKey As String
    Dim availableProfiles As Variant

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If

    profileName = Trim$(profileName)
    If Len(profileName) = 0 Then Exit Sub

    resolvedModeKey = Trim$(modeKey)
    If Len(resolvedModeKey) = 0 Then
        resolvedModeKey = Trim$(m_GetActiveModeKey(ws))
    End If
    If Len(resolvedModeKey) = 0 Then Exit Sub

    availableProfiles = ex_UiXmlProvider.m_GetDropdownItemsByName(PFUI_PROFILE_DROPDOWN_SHAPE, ThisWorkbook, resolvedModeKey)
    If mp_ArrayHasItems(availableProfiles) Then
        If Not mp_ArrayContains(availableProfiles, profileName) Then
            MsgBox "Profile '" & profileName & "' is not available for mode '" & resolvedModeKey & "'.", vbExclamation
            Exit Sub
        End If
    End If

    mp_SetStatePropText STATE_ACTIVE_PROFILE_PROP_PREFIX & mp_NormalizePropSuffix(resolvedModeKey), profileName
    ex_UiXmlProvider.m_SetDropdownContextValue "activeProfile", profileName
End Sub

Public Sub m_EnsureProfileDropdown(Optional ByVal ws As Worksheet)
    Dim profiles As Variant
    Dim profileName As String
    Dim savedProfileName As String

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If

    m_EnsureModeDropdown ws
    mp_ApplyModeVisibility ws

    profiles = mp_GetProfileNames(ws)
    m_RefreshProfileValidation ws

    If Not mp_ArrayHasItems(profiles) Then
        MsgBox "Profiles are missing or config file is empty: " & mp_GetProfilesFilePath(), vbExclamation
        Exit Sub
    End If

    savedProfileName = mp_GetSavedProfileNameForModeKey(mp_GetSelectedModeKey(ws))
    profileName = mp_GetSelectedProfileNameFromDropdown(ws, profiles, True, savedProfileName)
    If Len(profileName) = 0 Then
        Exit Sub
    End If

    m_ApplyProfileFromDev profileName
End Sub

Public Sub m_OnProfileChanged()
    Dim ws As Worksheet
    Dim profiles As Variant
    Dim profileName As String
    Dim targetStableZoneLeft As Double

    Set ws = ws_Dev
    targetStableZoneLeft = ex_CustomDropdown.m_GetStableZoneStartLeft(ws)
    profiles = mp_GetProfileNames(ws)
    If Not mp_ArrayHasItems(profiles) Then
        MsgBox "Profiles are missing or config file is empty: " & mp_GetProfilesFilePath(), vbExclamation
        Exit Sub
    End If

    profileName = mp_GetSelectedProfileNameFromDropdown(ws, profiles, True)
    If Len(profileName) = 0 Then Exit Sub

    m_ApplyProfileFromDev profileName
    mp_ReapplySelectedProfileUi ws
    m_SaveSelectionState ws
    mp_SyncDropdownContextState ws

    If targetStableZoneLeft >= 0 Then
        ex_CustomDropdown.m_StabilizeChooseModeAnchorX ws, targetStableZoneLeft
    End If
End Sub

Public Function m_ReapplyActiveProfileIfSourceChanged(Optional ByVal ws As Worksheet) As Boolean
    Dim modeKey As String
    Dim profileName As String
    Dim propName As String
    Dim currentSignature As String
    Dim savedSignature As String

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If
    If ws Is Nothing Then Exit Function

    modeKey = Trim$(mp_GetSelectedModeKey(ws))
    If Len(modeKey) = 0 Then
        modeKey = Trim$(m_GetActiveModeKey(ws))
    End If
    If Len(modeKey) = 0 Then Exit Function

    profileName = Trim$(m_GetActiveProfileName(ws))
    If Len(profileName) = 0 Then Exit Function

    currentSignature = mp_BuildProfilesFileSignature(modeKey)
    If Len(currentSignature) = 0 Then Exit Function

    propName = mp_BuildProfileFileSignaturePropName(modeKey, profileName)
    savedSignature = mp_GetStatePropText(propName, vbNullString)

    If Len(savedSignature) = 0 Then
        mp_SetStatePropText propName, currentSignature
        Exit Function
    End If

    If StrComp(savedSignature, currentSignature, vbBinaryCompare) <> 0 Then
        m_ApplyProfileFromDev profileName
        mp_SaveAppliedProfileFileSignature ws, profileName
        m_ReapplyActiveProfileIfSourceChanged = True
    End If
End Function

Public Sub m_OnModeChanged()
    Dim ws As Worksheet
    Dim profiles As Variant
    Dim targetStableZoneLeft As Double

    Set ws = ws_Dev
    targetStableZoneLeft = ex_CustomDropdown.m_GetStableZoneStartLeft(ws)

    m_EnsureModeDropdown ws
    mp_ApplyModeVisibility ws
    profiles = mp_GetProfileNames(ws)
    If Not mp_ArrayHasItems(profiles) Then Exit Sub
    m_EnsureProfileDropdown ws
    mp_ReapplySelectedProfileUi ws
    m_SaveSelectionState ws
    mp_SyncDropdownContextState ws

    If targetStableZoneLeft >= 0 Then
        ex_CustomDropdown.m_StabilizeChooseModeAnchorX ws, targetStableZoneLeft
    End If
End Sub

Public Sub m_RestoreSelectionState(Optional ByVal ws As Worksheet)
    Dim savedModeKey As String

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If

    m_EnsureModeDropdown ws
    savedModeKey = mp_GetStatePropText(STATE_ACTIVE_MODE_KEY_PROP, vbNullString)
    If Len(savedModeKey) > 0 Then
        mp_TrySelectModeByKey ws, savedModeKey
    End If

    mp_ApplyModeVisibility ws
    m_EnsureProfileDropdown ws
    mp_ReapplySelectedProfileUi ws
    mp_SyncDropdownContextState ws
End Sub

Public Sub m_SaveSelectionState(Optional ByVal ws As Worksheet)
    Dim modeKey As String
    Dim profileName As String
    Dim profiles As Variant

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If

    modeKey = Trim$(mp_GetSelectedModeKey(ws))
    If Len(modeKey) = 0 Then Exit Sub
    mp_SetStatePropText STATE_ACTIVE_MODE_KEY_PROP, modeKey

    profiles = mp_GetProfileNames(ws)
    If Not mp_ArrayHasItems(profiles) Then Exit Sub

    profileName = Trim$(mp_GetSelectedProfileNameFromDropdown(ws, profiles, False))
    If Len(profileName) = 0 Then Exit Sub

    mp_SetStatePropText STATE_ACTIVE_PROFILE_PROP_PREFIX & mp_NormalizePropSuffix(modeKey), profileName
    mp_SyncDropdownContextState ws
End Sub

Public Sub m_ResetDevUILayout(Optional ByVal ws As Worksheet)
    If ws Is Nothing Then
        Set ws = ws_Dev
    End If
    If ws Is Nothing Then Exit Sub

    ' Layer 1: hard reset worksheet dimensions to default baseline.
    mp_ResetWorksheetDimensions ws

    ' Layer 2: apply config-table column layout (includes AutoFit by existing logic).
    mp_ApplyConfigColumnsLayoutLayer ws

    ' Layer 3 (disabled): stable-zone compensation via buffer column.
    ' ex_CustomDropdown.m_StabilizeChooseModeAnchorX ws, ex_CustomDropdown.m_GetStableZoneStartLeft(ws)

    ex_CustomDropdown.m_InitDevTestDropdown ThisWorkbook
End Sub

Public Sub m_ResetDevUILayout_UI()
    m_ResetDevUILayout ws_Dev
End Sub

Public Sub m_EnsureProfileDropdown_UI()
    m_EnsureProfileDropdown ws_Dev
End Sub

Public Sub m_OpenProfilePicker_UI()
    m_EnsureProfileDropdown ws_Dev
End Sub

Public Sub m_OpenProfilePicker(Optional ByVal ws As Worksheet)
    If ws Is Nothing Then
        Set ws = ws_Dev
    End If
    m_EnsureProfileDropdown ws
End Sub

Public Sub m_SaveCurrentProfileToConfig(Optional ByVal ws As Worksheet)
    Dim profiles As Variant
    Dim profileName As String
    Dim doc As Object
    Dim profileNode As Object

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If

    profiles = mp_GetProfileNames(ws)
    If Not mp_ArrayHasItems(profiles) Then
        MsgBox "Profiles are missing or config file is empty: " & mp_GetProfilesFilePath(), vbExclamation
        Exit Sub
    End If

    profileName = mp_GetSelectedProfileNameFromDropdown(ws, profiles, False)
    If Len(profileName) = 0 Then Exit Sub

    If Not mp_ArrayContains(profiles, profileName) Then
        MsgBox "Profile '" & profileName & "' was not found in config file: " & mp_GetProfilesFilePath(), vbExclamation
        Exit Sub
    End If

    Set doc = mp_LoadProfilesDom(ws)
    Set profileNode = mp_GetProfileNode(doc, profileName, True)
    If profileNode Is Nothing Then
        MsgBox "Failed to access profile '" & profileName & "' in config file: " & mp_GetProfilesFilePath(), vbExclamation
        Exit Sub
    End If

    ex_ProfilesEntriesMapper.m_WriteSheetValuesToProfile ws, doc, profileNode
    mp_SaveProfilesDom doc
    m_SaveSelectionState ws
    On Error Resume Next
    ex_ConfigProvider.m_RefreshConfigTitle ws, profileName
    On Error GoTo 0
    ex_Messaging.m_ShowNotice "Profiles config saved: " & profileName
End Sub

' =============================================================================
' Internal helpers
' =============================================================================

' Создаёт новый профиль в DOM из текущего состояния таблицы конфигурации.
' Используется как "seed" при инициализации пустого файла профилей.
' Порядок строк в таблице сохраняется как порядок узлов <v> в XML.
Private Sub mp_SeedProfileFromSheet(ByVal doc As Object, ByVal ws As Worksheet)
    Dim profileName As String
    Dim root As Object
    Dim profileNode As Object
    Dim entries As Variant
    Dim i As Long
    Dim vNode As Object

    profileName = vbNullString
    entries = mp_GetProfileNames(ws)
    If mp_ArrayHasItems(entries) Then
        profileName = mp_GetSelectedProfileNameFromDropdown(ws, entries, False)
    End If
    If Len(profileName) = 0 Then
        profileName = "Default"
    End If

    Set root = doc.selectSingleNode("/p:profiles")
    If root Is Nothing Then Exit Sub

    Set profileNode = doc.createNode(1, "profile", PROFILES_NS)
    profileNode.setAttribute "name", profileName
    root.appendChild profileNode

    entries = ex_ProfilesEntriesMapper.m_ReadConfigTableEntries(ws)
    If Not mp_ArrayHasItems(entries) Then Exit Sub

    For i = LBound(entries, 1) To UBound(entries, 1)
        Set vNode = doc.createNode(1, "v", PROFILES_NS)
        If Len(Trim$(CStr(entries(i, DEV_CONFIG_MARKER_COL)))) > 0 Then
            vNode.setAttribute "type", CStr(entries(i, DEV_CONFIG_MARKER_COL))
        End If
        vNode.setAttribute "key", CStr(entries(i, DEV_CONFIG_KEY_COL))
        If Len(Trim$(CStr(entries(i, DEV_CONFIG_STYLES_COL)))) > 0 Then
            vNode.setAttribute "styles", CStr(entries(i, DEV_CONFIG_STYLES_COL))
        End If
        vNode.Text = CStr(entries(i, DEV_CONFIG_VALUE_COL))
        profileNode.appendChild vNode
    Next i

    mp_SaveProfilesDom doc
End Sub

Private Function mp_GetProfileNames(Optional ByVal ws As Worksheet) As Variant
    Dim modeKey As String

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If

    modeKey = m_GetActiveModeKey(ws)
    mp_GetProfileNames = ex_UiXmlProvider.m_GetDropdownItemsByName(PFUI_PROFILE_DROPDOWN_SHAPE, ThisWorkbook, modeKey)
End Function

Private Function mp_ArrayHasItems(ByVal values As Variant) As Boolean
    On Error GoTo EH
    If IsArray(values) Then
        mp_ArrayHasItems = (UBound(values) >= LBound(values))
    End If
    Exit Function
EH:
    mp_ArrayHasItems = False
End Function

Private Function mp_ArrayContains(ByVal values As Variant, ByVal needle As String) As Boolean
    Dim i As Long

    If Not mp_ArrayHasItems(values) Then Exit Function
    needle = Trim$(needle)
    If Len(needle) = 0 Then Exit Function

    For i = LBound(values) To UBound(values)
        If StrComp(CStr(values(i)), needle, vbTextCompare) = 0 Then
            mp_ArrayContains = True
            Exit Function
        End If
    Next i
End Function

Private Function mp_GetSelectedProfileNameFromDropdown( _
    ByVal ws As Worksheet, _
    ByVal profiles As Variant, _
    Optional ByVal syncItems As Boolean = False, _
    Optional ByVal preferredName As String = vbNullString _
) As String
    Dim modeKey As String
    Dim resolvedProfile As String
    Dim contextProfile As String

    modeKey = Trim$(mp_GetSelectedModeKey(ws))
    preferredName = Trim$(preferredName)

    If Len(preferredName) > 0 And mp_ArrayContains(profiles, preferredName) Then
        resolvedProfile = preferredName
    End If

    If Len(resolvedProfile) = 0 Then
        resolvedProfile = Trim$(mp_GetSavedProfileNameForModeKey(modeKey))
        If Len(resolvedProfile) = 0 Then
            contextProfile = Trim$(ex_UiXmlProvider.m_GetDropdownContextValue("activeProfile", vbNullString))
            If Len(contextProfile) > 0 Then resolvedProfile = contextProfile
        End If
        If Len(resolvedProfile) > 0 Then
            If Not mp_ArrayContains(profiles, resolvedProfile) Then
                resolvedProfile = vbNullString
            End If
        End If
    End If

    If Len(resolvedProfile) = 0 Then
        resolvedProfile = CStr(profiles(LBound(profiles)))
    End If

    If syncItems Then
        m_SetActiveProfileName resolvedProfile, modeKey, ws
    End If

    mp_GetSelectedProfileNameFromDropdown = resolvedProfile
End Function

Private Sub m_EnsureModeDropdown(ByVal ws As Worksheet)
    Dim modeKey As String

    modeKey = Trim$(mp_GetSelectedModeKey(ws))
    If Len(modeKey) = 0 Then
        MsgBox "Mode list is empty for control '" & PFUI_MODE_DROPDOWN_SHAPE & "' in config\DevUI.xml.", vbExclamation
        Exit Sub
    End If

    mp_SetStatePropText STATE_ACTIVE_MODE_KEY_PROP, modeKey
End Sub

Private Sub mp_TrySelectModeByName(ByVal ws As Worksheet, ByVal modeName As String)
    Dim modeKey As String

    modeName = Trim$(modeName)
    If Len(modeName) = 0 Then Exit Sub

    modeKey = mp_ResolveModeKeyByName(modeName)
    If Len(modeKey) = 0 Then Exit Sub
    mp_SetStatePropText STATE_ACTIVE_MODE_KEY_PROP, modeKey
End Sub

Private Sub mp_TrySelectModeByKey(ByVal ws As Worksheet, ByVal modeKey As String)
    modeKey = Trim$(modeKey)
    If Len(modeKey) = 0 Then Exit Sub
    If Not mp_IsModeKeyValid(modeKey) Then Exit Sub
    mp_SetStatePropText STATE_ACTIVE_MODE_KEY_PROP, modeKey
End Sub

Private Function mp_GetSelectedModeKey(ByVal ws As Worksheet) As String
    Dim modeKey As String
    Dim fallbackModeKey As String

    modeKey = Trim$(mp_GetStatePropText(STATE_ACTIVE_MODE_KEY_PROP, vbNullString))
    If Len(modeKey) > 0 Then
        If mp_IsModeKeyValid(modeKey) Then
            mp_GetSelectedModeKey = modeKey
            Exit Function
        End If
    End If

    fallbackModeKey = Trim$(ex_UiXmlProvider.m_GetDefaultModeKey(ThisWorkbook))
    If Len(fallbackModeKey) = 0 Then fallbackModeKey = Trim$(ex_UiXmlProvider.m_GetModeKeyByIndex(1, ThisWorkbook))
    If Len(fallbackModeKey) > 0 Then
        If mp_IsModeKeyValid(fallbackModeKey) Then
            mp_SetStatePropText STATE_ACTIVE_MODE_KEY_PROP, fallbackModeKey
            mp_GetSelectedModeKey = fallbackModeKey
            Exit Function
        End If
    End If
End Function

Private Function mp_GetSelectedModeName(ByVal ws As Worksheet) As String
    Dim modeKey As String
    Dim modeName As String

    modeKey = Trim$(mp_GetSelectedModeKey(ws))
    If Len(modeKey) = 0 Then Exit Function

    modeName = Trim$(mp_ResolveModeNameByKey(modeKey))
    If Len(modeName) = 0 Then modeName = modeKey
    mp_GetSelectedModeName = modeName
End Function

Private Function mp_ResolveModeKeyByName(ByVal modeName As String) As String
    Dim modeKey As String
    Dim existingCaption As String

    modeName = Trim$(modeName)
    If Len(modeName) = 0 Then Exit Function

    modeKey = Trim$(ex_UiXmlProvider.m_GetDropdownItemKeyByTarget(PFUI_MODE_DROPDOWN_SHAPE, modeName, ThisWorkbook))
    If Len(modeKey) = 0 Then
        existingCaption = Trim$(ex_UiXmlProvider.m_GetDropdownItemCaptionByKey(PFUI_MODE_DROPDOWN_SHAPE, modeName, ThisWorkbook))
        If Len(existingCaption) > 0 Then modeKey = modeName
    End If

    mp_ResolveModeKeyByName = modeKey
End Function

Private Function mp_ResolveModeNameByKey(ByVal modeKey As String) As String
    Dim modeName As String

    modeKey = Trim$(modeKey)
    If Len(modeKey) = 0 Then Exit Function

    modeName = Trim$(ex_UiXmlProvider.m_GetDropdownItemCaptionByKey(PFUI_MODE_DROPDOWN_SHAPE, modeKey, ThisWorkbook))
    If Len(modeName) = 0 Then modeName = modeKey

    mp_ResolveModeNameByKey = modeName
End Function

Private Sub mp_ApplyModeVisibility(ByVal ws As Worksheet)
    Dim profiles As Variant
    Dim profileName As String
    Dim doc As Object
    Dim profileNode As Object

    profiles = mp_GetProfileNames(ws)
    If Not mp_ArrayHasItems(profiles) Then Exit Sub

    profileName = mp_GetSelectedProfileNameFromDropdown(ws, profiles, False)
    If Len(profileName) = 0 Then
        profileName = CStr(profiles(LBound(profiles)))
    End If
    If Len(profileName) = 0 Then Exit Sub

    Set doc = mp_LoadProfilesDom(ws)
    Set profileNode = mp_GetProfileNode(doc, profileName, False)
    If profileNode Is Nothing Then Exit Sub

    ex_ConfigProfilesManager.m_ApplyModeVisibility ws, profileNode
End Sub

Private Sub mp_ReapplySelectedProfileUi(ByVal ws As Worksheet)
    Dim profiles As Variant
    Dim profileName As String
    Dim doc As Object
    Dim profileNode As Object

    profiles = mp_GetProfileNames(ws)
    If Not mp_ArrayHasItems(profiles) Then Exit Sub

    profileName = mp_GetSelectedProfileNameFromDropdown(ws, profiles, False)
    If Len(profileName) = 0 Then Exit Sub

    Set doc = mp_LoadProfilesDom(ws)
    Set profileNode = mp_GetProfileNode(doc, profileName, False)
    If profileNode Is Nothing Then Exit Sub

    ex_ConfigProfilesManager.m_ApplyProfileUI ws, profileNode, profileName
    ex_ConfigProfilesManager.m_ApplyModeVisibility ws, profileNode
End Sub

Private Function mp_ArrayLength(ByVal values As Variant) As Long
    If Not mp_ArrayHasItems(values) Then Exit Function
    mp_ArrayLength = UBound(values) - LBound(values) + 1
End Function

Private Function mp_IsModeKeyValid(ByVal modeKey As String) As Boolean
    Dim modeRecords As Variant
    Dim i As Long
    Dim recordKey As String

    modeKey = Trim$(modeKey)
    If Len(modeKey) = 0 Then Exit Function

    modeRecords = ex_UiXmlProvider.m_GetDropdownItemRecordsByControl(PFUI_MODE_DROPDOWN_SHAPE, ThisWorkbook)
    If Not mp_ArrayHasItems(modeRecords) Then Exit Function

    For i = LBound(modeRecords, 1) To UBound(modeRecords, 1)
        recordKey = Trim$(CStr(modeRecords(i, ex_UiXmlProvider.DROPDOWN_ITEM_COL_KEY)))
        If StrComp(recordKey, modeKey, vbTextCompare) = 0 Then
            mp_IsModeKeyValid = True
            Exit Function
        End If
    Next i
End Function

Private Function mp_FindProfileIndexByName(ByVal profiles As Variant, ByVal profileName As String) As Long
    Dim i As Long

    profileName = Trim$(profileName)
    If Len(profileName) = 0 Then Exit Function
    If Not mp_ArrayHasItems(profiles) Then Exit Function

    For i = LBound(profiles) To UBound(profiles)
        If StrComp(CStr(profiles(i)), profileName, vbTextCompare) = 0 Then
            mp_FindProfileIndexByName = i - LBound(profiles) + 1
            Exit Function
        End If
    Next i
End Function

Private Function mp_GetSavedProfileNameForModeKey(ByVal modeKey As String) As String
    Dim propName As String

    modeKey = Trim$(modeKey)
    If Len(modeKey) = 0 Then Exit Function

    propName = STATE_ACTIVE_PROFILE_PROP_PREFIX & mp_NormalizePropSuffix(modeKey)
    mp_GetSavedProfileNameForModeKey = mp_GetStatePropText(propName, vbNullString)
End Function

Private Function mp_NormalizePropSuffix(ByVal valueText As String) As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim resultText As String

    valueText = Trim$(valueText)
    If Len(valueText) = 0 Then
        mp_NormalizePropSuffix = "_"
        Exit Function
    End If

    For i = 1 To Len(valueText)
        ch = Mid$(valueText, i, 1)
        code = AscW(ch)
        If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Then
            resultText = resultText & ch
        Else
            resultText = resultText & "_"
        End If
    Next i

    If Len(resultText) = 0 Then resultText = "_"
    mp_NormalizePropSuffix = resultText
End Function

Private Sub mp_SyncDropdownContextState(ByVal ws As Worksheet)
    Dim modeKey As String
    Dim profileName As String

    modeKey = Trim$(mp_GetSelectedModeKey(ws))
    If Len(modeKey) > 0 Then
        ex_UiXmlProvider.m_SetDropdownContextValue "activeMode", modeKey
    End If

    profileName = Trim$(m_GetActiveProfileName(ws))
    If Len(profileName) > 0 Then
        ex_UiXmlProvider.m_SetDropdownContextValue "activeProfile", profileName
    End If
End Sub

Private Function mp_GetStatePropText(ByVal propName As String, ByVal defaultValue As String) As String
    On Error GoTo EH
    mp_GetStatePropText = CStr(ThisWorkbook.CustomDocumentProperties(propName).Value)
    Exit Function
EH:
    mp_GetStatePropText = defaultValue
End Function

Private Sub mp_SetStatePropText(ByVal propName As String, ByVal valueText As String)
    On Error GoTo AddProp
    ThisWorkbook.CustomDocumentProperties(propName).Value = CStr(valueText)
    Exit Sub
AddProp:
    ThisWorkbook.CustomDocumentProperties.Add _
        Name:=propName, _
        LinkToContent:=False, _
        Type:=msoPropertyTypeString, _
        Value:=CStr(valueText)
End Sub

Private Sub mp_ResetWorksheetDimensions(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Columns.ColumnWidth = 8
    ws.Rows.RowHeight = 16
    On Error GoTo 0
End Sub

Private Sub mp_ApplyConfigColumnsLayoutLayer(ByVal ws As Worksheet)
    Dim tbl As ListObject

    Set tbl = ex_ConfigTableStore.m_GetConfigTable(ws, True)
    If tbl Is Nothing Then
        MsgBox "Config table '" & DEV_CONFIG_TABLE_NAME & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If

    ex_ConfigTableStore.m_ApplyConfigMarkerStyles tbl
End Sub

Private Function mp_LoadProfilesDom(Optional ByVal ws As Worksheet) As Object
    Dim filePath As String
    Dim modeKey As String

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If

    modeKey = mp_GetSelectedModeKey(ws)
    filePath = mp_GetProfilesFilePath(modeKey)
    If Len(filePath) = 0 Then Exit Function

    Set mp_LoadProfilesDom = ex_ProfilesStore.m_LoadProfilesDom(filePath)
End Function

Private Sub mp_SaveProfilesDom(ByVal doc As Object)
    Dim filePath As String
    Dim modeKey As String

    modeKey = mp_GetSelectedModeKey(ws_Dev)
    filePath = mp_GetProfilesFilePath(modeKey)
    ex_ProfilesStore.m_SaveProfilesDom doc, filePath
End Sub

Private Sub mp_SaveAppliedProfileFileSignature(ByVal ws As Worksheet, ByVal profileName As String)
    Dim modeKey As String
    Dim signatureText As String
    Dim propName As String

    If ws Is Nothing Then Exit Sub

    modeKey = Trim$(mp_GetSelectedModeKey(ws))
    If Len(modeKey) = 0 Then
        modeKey = Trim$(m_GetActiveModeKey(ws))
    End If
    If Len(modeKey) = 0 Then Exit Sub

    profileName = Trim$(profileName)
    If Len(profileName) = 0 Then Exit Sub

    signatureText = mp_BuildProfilesFileSignature(modeKey)
    If Len(signatureText) = 0 Then Exit Sub

    propName = mp_BuildProfileFileSignaturePropName(modeKey, profileName)
    mp_SetStatePropText propName, signatureText
End Sub

Private Function mp_BuildProfileFileSignaturePropName(ByVal modeKey As String, ByVal profileName As String) As String
    mp_BuildProfileFileSignaturePropName = _
        STATE_PROFILE_FILE_SIGNATURE_PROP_PREFIX & _
        mp_NormalizePropSuffix(modeKey) & "." & mp_NormalizePropSuffix(profileName)
End Function

Private Function mp_BuildProfilesFileSignature(ByVal modeKey As String) As String
    Dim filePath As String
    Dim modifiedAt As Date
    Dim fileSize As Long

    filePath = Trim$(mp_GetProfilesFilePath(modeKey))
    If Len(filePath) = 0 Then Exit Function

    On Error GoTo EH
    modifiedAt = FileDateTime(filePath)
    fileSize = FileLen(filePath)
    mp_BuildProfilesFileSignature = Format$(modifiedAt, "yyyy-mm-dd hh:nn:ss") & "|" & CStr(fileSize)
    Exit Function
EH:
    mp_BuildProfilesFileSignature = vbNullString
End Function


Private Function mp_NodeAttrText(ByVal node As Object, ByVal attrName As String) As String
    On Error Resume Next
    mp_NodeAttrText = CStr(node.selectSingleNode("@*[local-name()='" & attrName & "']").Text)
    If Err.Number <> 0 Then
        Err.Clear
        mp_NodeAttrText = vbNullString
    End If
    On Error GoTo 0
End Function

' Перезаписывает таблицу целиком:
' - очищает старое содержимое;
' - ресайзит таблицу под новый объём;
' - накладывает маркерные метки и pipeline-стили.
Private Sub mp_WriteEntriesToConfigTable(ByVal ws As Worksheet, ByVal entries As Variant)
    Dim tbl As ListObject
    Dim rowCount As Long
    Dim values() As Variant
    Dim i As Long

    ' Table normalization/writes can fail if previous profile left sheet protected.
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    Set tbl = ex_ConfigTableStore.m_GetConfigTable(ws, True)
    If tbl Is Nothing Then
        MsgBox "Config table '" & DEV_CONFIG_TABLE_NAME & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If

    ex_ConfigTableStore.m_ClearConfigDataArea ws, tbl

    rowCount = mp_ArrayRowCount(entries)
    ex_ConfigTableStore.m_ResizeConfigTableRows ws, tbl, rowCount

    If rowCount > 0 Then
        ReDim values(1 To rowCount, 1 To DEV_CONFIG_COL_COUNT)
        For i = 1 To rowCount
            values(i, DEV_CONFIG_MARKER_COL) = CStr(entries(i, DEV_CONFIG_MARKER_COL))
            values(i, DEV_CONFIG_KEY_COL) = CStr(entries(i, DEV_CONFIG_KEY_COL))
            values(i, DEV_CONFIG_VALUE_COL) = CStr(entries(i, DEV_CONFIG_VALUE_COL))
            values(i, DEV_CONFIG_STYLES_COL) = CStr(entries(i, DEV_CONFIG_STYLES_COL))
        Next i

        tbl.DataBodyRange.Cells(1, 1).Resize(rowCount, DEV_CONFIG_COL_COUNT).Value = values
    End If

    ex_ConfigTableStore.m_ApplyConfigMarkerStyles tbl
End Sub

Private Sub mp_ApplyLockedPlaceholderCells(ByVal ws As Worksheet, ByVal tbl As ListObject, ByVal lockedWithPlaceholder As Object)
    Dim r As Long
    Dim markerText As String
    Dim keyText As String
    Dim hasLockedCells As Boolean
    Dim placeholderText As String
    Dim cell As Range
    Dim lockedRows As Collection

    If ws Is Nothing Then Exit Sub
    If tbl Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Unprotect
    On Error GoTo EH

    ws.EnableSelection = xlNoRestrictions
    ' Base policy for Dev sheet: unlock everything, then lock only explicit placeholder cells.
    ws.Cells.Locked = False
    Set lockedRows = New Collection

    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.Columns(DEV_CONFIG_VALUE_COL).NumberFormat = "General"

        For r = 1 To tbl.DataBodyRange.Rows.Count
            markerText = Trim$(CStr(tbl.DataBodyRange.Cells(r, DEV_CONFIG_MARKER_COL).Value))
            If StrComp(markerText, DEV_MARKER_SYMBOL, vbTextCompare) = 0 Then GoTo ContinueRow

            keyText = Trim$(CStr(tbl.DataBodyRange.Cells(r, DEV_CONFIG_KEY_COL).Value))
            If lockedWithPlaceholder Is Nothing Then GoTo ContinueRow
            If Not lockedWithPlaceholder.Exists(keyText) Then GoTo ContinueRow
            placeholderText = Trim$(CStr(lockedWithPlaceholder(keyText)))
            If Len(placeholderText) = 0 Then GoTo ContinueRow

            Set cell = tbl.DataBodyRange.Cells(r, DEV_CONFIG_VALUE_COL)
            cell.Locked = True
            cell.NumberFormat = mp_BuildLockedPlaceholderFormat(placeholderText)
            lockedRows.Add cell.Row
            hasLockedCells = True

ContinueRow:
        Next r
    End If

    If hasLockedCells Then
        mp_ApplyLockedPlaceholderStylePipeline ws, tbl, lockedRows
        ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True, AllowFiltering:=True
        ws.EnableSelection = xlUnlockedCells
    Else
        ws.EnableSelection = xlNoRestrictions
    End If
    Exit Sub

EH:
    MsgBox "Failed to apply locked-placeholder masking/protection on sheet '" & ws.Name & "': " & Err.Description, vbExclamation
End Sub

Private Sub mp_ApplyLockedPlaceholderStylePipeline(ByVal ws As Worksheet, ByVal tbl As ListObject, ByVal lockedRows As Collection)
    Dim rowKindRanges As Object
    Dim allRows As Collection
    Dim headerRows As Collection
    Dim dataRows As Collection
    Dim markerRows As Collection
    Dim rowCount As Long
    Dim i As Long
    Dim rowIndex As Long
    Dim markerKind As String
    Dim keyText As String

    If ws Is Nothing Then Exit Sub
    If tbl Is Nothing Then Exit Sub
    If lockedRows Is Nothing Then Exit Sub
    If lockedRows.Count = 0 Then Exit Sub

    Set rowKindRanges = CreateObject("Scripting.Dictionary")
    rowKindRanges.CompareMode = 1

    Set allRows = New Collection
    Set headerRows = New Collection
    Set dataRows = New Collection
    Set markerRows = New Collection

    rowIndex = tbl.HeaderRowRange.Row
    allRows.Add rowIndex
    headerRows.Add rowIndex

    If Not tbl.DataBodyRange Is Nothing Then
        rowCount = tbl.DataBodyRange.Rows.Count
        For i = 1 To rowCount
            rowIndex = tbl.DataBodyRange.Row + i - 1
            allRows.Add rowIndex

            markerKind = Trim$(CStr(tbl.DataBodyRange.Cells(i, DEV_CONFIG_MARKER_COL).Value))
            keyText = Trim$(CStr(tbl.DataBodyRange.Cells(i, DEV_CONFIG_KEY_COL).Value))
            If StrComp(markerKind, DEV_MARKER_SYMBOL, vbTextCompare) = 0 Or mp_IsMarkerKey(keyText) Then
                markerRows.Add rowIndex
            Else
                dataRows.Add rowIndex
            End If
        Next i
    End If

    Set rowKindRanges("configall") = allRows
    Set rowKindRanges("configheader") = headerRows
    Set rowKindRanges("configdata") = dataRows
    Set rowKindRanges("configmarker") = markerRows
    Set rowKindRanges("configlockedplaceholder") = lockedRows

    ex_OutputFormattingPipeline.m_ApplySheetPipeline ws, Nothing, Nothing, rowKindRanges
End Sub

Private Function mp_ReadLockedWithPlaceholder(ByVal profileNode As Object) As Object
    Dim result As Object
    Dim nodes As Object
    Dim node As Object
    Dim keyText As String
    Dim placeholderText As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1

    If profileNode Is Nothing Then
        Set mp_ReadLockedWithPlaceholder = result
        Exit Function
    End If

    Set nodes = profileNode.selectNodes("p:v")
    If nodes Is Nothing Then
        Set mp_ReadLockedWithPlaceholder = result
        Exit Function
    End If

    For Each node In nodes
        keyText = Trim$(mp_NodeAttrText(node, "key"))
        placeholderText = Trim$(mp_NodeAttrText(node, XML_ATTR_LOCKED_WITH_PLACEHOLDER))
        If Len(keyText) = 0 Then GoTo ContinueNode
        If Len(placeholderText) = 0 Then GoTo ContinueNode
        result(keyText) = placeholderText
ContinueNode:
    Next node

    Set mp_ReadLockedWithPlaceholder = result
End Function

Private Function mp_BuildLockedPlaceholderFormat(ByVal placeholderText As String) As String
    placeholderText = Replace(placeholderText, """", """""")
    mp_BuildLockedPlaceholderFormat = ";;;""" & placeholderText & """"
End Function

Private Function mp_IsMarkerKey(ByVal keyText As String) As Boolean
    keyText = Trim$(keyText)
    If Len(keyText) < Len(DEV_MARKER_PREFIX) Then Exit Function
    mp_IsMarkerKey = (StrComp(Left$(keyText, Len(DEV_MARKER_PREFIX)), DEV_MARKER_PREFIX, vbTextCompare) = 0)
End Function


Private Function mp_ArrayRowCount(ByVal values As Variant) As Long
    On Error GoTo EH
    If IsArray(values) Then
        mp_ArrayRowCount = UBound(values, 1) - LBound(values, 1) + 1
    End If
    Exit Function
EH:
    mp_ArrayRowCount = 0
End Function


Private Function mp_GetProfilesFilePath(Optional ByVal modeKey As String = vbNullString) As String
    Dim resolvedModeKey As String
    Dim defaultModeKey As String

    resolvedModeKey = Trim$(modeKey)
    If Len(resolvedModeKey) = 0 Then
        defaultModeKey = Trim$(ex_UiXmlProvider.m_GetDefaultModeKey(ThisWorkbook))
        If Len(defaultModeKey) > 0 Then resolvedModeKey = defaultModeKey
        On Error Resume Next
        resolvedModeKey = mp_GetSelectedModeKey(ws_Dev)
        On Error GoTo 0
    End If
    If Len(resolvedModeKey) = 0 Then resolvedModeKey = Trim$(ex_UiXmlProvider.m_GetModeKeyByIndex(1, ThisWorkbook))

    mp_GetProfilesFilePath = ex_ProfilesStore.m_GetProfilesFilePath(resolvedModeKey, ThisWorkbook)
End Function

Private Function mp_GetProfileNode(ByVal doc As Object, ByVal profileName As String, ByVal createIfMissing As Boolean) As Object
    Set mp_GetProfileNode = ex_ProfilesStore.m_GetProfileNode(doc, profileName, createIfMissing)
End Function

Private Function mp_ApplyProfileConfigStyles( _
    ByVal ws As Worksheet, _
    ByVal profileNode As Object, _
    Optional ByVal targetStableZoneLeft As Double = -1 _
) As Boolean
    Dim cfgNodes As Object
    Dim cfgNode As Object
    Dim tbl As ListObject
    Dim usedIds As Object
    Dim colId As String
    Dim relCol As Long
    Dim absCol As Long
    Dim widthText As String
    Dim alignText As String
    Dim overflowText As String
    Dim widthUnits As Double
    Dim horizontalAlign As Long
    Dim verticalAlign As Long
    Dim bodyRange As Range
    Dim normalizedOverflow As String
    Dim shouldAutoFitConfigRows As Boolean
    Dim didScale As Boolean

    If ws Is Nothing Then
        MsgBox "Failed to apply profile config styles: worksheet is not specified.", vbExclamation
        Exit Function
    End If
    If profileNode Is Nothing Then
        MsgBox "Failed to apply profile config styles: profile node is not specified.", vbExclamation
        Exit Function
    End If

    On Error Resume Next
    profileNode.OwnerDocument.setProperty "SelectionNamespaces", "xmlns:p='" & PROFILES_NS & "'"
    On Error GoTo 0

    Set cfgNodes = profileNode.selectNodes("p:styles/p:config/p:column")
    If cfgNodes Is Nothing Then
        mp_ApplyProfileConfigStyles = True
        Exit Function
    End If
    If cfgNodes.Length = 0 Then
        mp_ApplyProfileConfigStyles = True
        Exit Function
    End If

    Set tbl = ex_ConfigTableStore.m_GetConfigTable(ws, True)
    If tbl Is Nothing Then
        MsgBox "Config table '" & DEV_CONFIG_TABLE_NAME & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        Exit Function
    End If

    Set usedIds = CreateObject("Scripting.Dictionary")
    usedIds.CompareMode = 1

    For Each cfgNode In cfgNodes
        colId = LCase$(Trim$(mp_NodeAttrText(cfgNode, "id")))
        If Len(colId) = 0 Then
            MsgBox "Profile styles/config/column must define non-empty attribute 'id'.", vbExclamation
            Exit Function
        End If
        If usedIds.Exists(colId) Then
            MsgBox "Duplicate profile styles/config/column id: '" & colId & "'.", vbExclamation
            Exit Function
        End If
        usedIds(colId) = True

        If Not mp_TryResolveProfileConfigColumnId(colId, relCol) Then
            MsgBox "Unsupported profile styles/config/column@id value: '" & colId & "'. Allowed: marker, key, value, styles.", vbExclamation
            Exit Function
        End If
        If relCol < 1 Or relCol > tbl.ListColumns.Count Then
            MsgBox "Profile styles/config/column id '" & colId & "' is out of bounds for table '" & DEV_CONFIG_TABLE_NAME & "'.", vbExclamation
            Exit Function
        End If

        widthText = Trim$(mp_NodeAttrText(cfgNode, "width"))
        alignText = Trim$(mp_NodeAttrText(cfgNode, "align"))
        overflowText = Trim$(mp_NodeAttrText(cfgNode, "overflow"))
        If Len(widthText) = 0 And Len(alignText) = 0 And Len(overflowText) = 0 Then
            MsgBox "Profile styles/config/column id '" & colId & "' must define at least one attribute: width, align, or overflow.", vbExclamation
            Exit Function
        End If

        absCol = tbl.Range.Column + relCol - 1
        If Len(widthText) > 0 Then
            If Not mp_TryParseProfileConfigWidth(widthText, widthUnits) Then
                MsgBox "Invalid profile styles/config/column@width for id '" & colId & "': '" & widthText & "'. Expected positive number.", vbExclamation
                Exit Function
            End If
            ws.Columns(absCol).ColumnWidth = widthUnits
        End If

        On Error Resume Next
        Set bodyRange = tbl.ListColumns(relCol).DataBodyRange
        On Error GoTo 0
        If bodyRange Is Nothing Then GoTo NextCfgNode

        If Len(alignText) > 0 Then
            If Not mp_TryParseProfileConfigAlign(alignText, horizontalAlign, verticalAlign) Then
                MsgBox "Invalid profile styles/config/column@align for id '" & colId & "': '" & alignText & "'.", vbExclamation
                Exit Function
            End If
            bodyRange.HorizontalAlignment = horizontalAlign
            bodyRange.VerticalAlignment = verticalAlign
        End If

        If Len(overflowText) > 0 Then
            normalizedOverflow = LCase$(Trim$(overflowText))
            Select Case normalizedOverflow
                Case "wrap"
                    bodyRange.WrapText = True
                    bodyRange.ShrinkToFit = False
                Case "shrink"
                    bodyRange.WrapText = False
                    bodyRange.ShrinkToFit = True
                Case "clip"
                    bodyRange.WrapText = False
                    bodyRange.ShrinkToFit = False
                Case Else
                    MsgBox "Invalid profile styles/config/column@overflow for id '" & colId & "': '" & overflowText & "'. Allowed: clip, wrap, shrink.", vbExclamation
                    Exit Function
            End Select

            If normalizedOverflow = "wrap" Then
                shouldAutoFitConfigRows = True
            End If
        End If
NextCfgNode:
    Next cfgNode

    If shouldAutoFitConfigRows Then
        mp_AutoFitConfigRangeRows ws, tbl
    End If

    If targetStableZoneLeft >= 0 Then
        ex_CustomDropdown.m_StabilizeChooseModeAnchorX ws, targetStableZoneLeft
        didScale = ex_ConfigTableStore.m_ScaleConfigColumnsToStableTarget(ws, tbl.Range.Column, DEV_CONFIG_COL_COUNT, targetStableZoneLeft)
        ex_CustomDropdown.m_StabilizeChooseModeAnchorX ws, targetStableZoneLeft
        If didScale Then
            ex_Messaging.m_ShowNotice "WARNING: Config styles exceed available dynamic zone width. Columns were scaled proportionally to fit."
        End If
    End If

    mp_ApplyProfileConfigStyles = True
End Function

Private Function mp_TryResolveProfileConfigColumnId(ByVal colId As String, ByRef outColIndex As Long) As Boolean
    Select Case LCase$(Trim$(colId))
        Case PROFILE_CFG_COL_MARKER
            outColIndex = DEV_CONFIG_MARKER_COL
        Case PROFILE_CFG_COL_KEY
            outColIndex = DEV_CONFIG_KEY_COL
        Case PROFILE_CFG_COL_VALUE
            outColIndex = DEV_CONFIG_VALUE_COL
        Case PROFILE_CFG_COL_STYLES
            outColIndex = DEV_CONFIG_STYLES_COL
        Case Else
            Exit Function
    End Select

    mp_TryResolveProfileConfigColumnId = True
End Function

Private Function mp_TryParseProfileConfigWidth(ByVal widthText As String, ByRef outWidth As Double) As Boolean
    widthText = Trim$(widthText)
    If Len(widthText) = 0 Then Exit Function

    If Len(widthText) >= 2 Then
        If LCase$(Right$(widthText, 2)) = "px" Then
            widthText = Trim$(Left$(widthText, Len(widthText) - 2))
        End If
    End If

    If Not IsNumeric(widthText) Then Exit Function

    outWidth = CDbl(widthText)
    If outWidth <= 0 Then Exit Function

    mp_TryParseProfileConfigWidth = True
End Function

Private Function mp_TryParseProfileConfigAlign(ByVal alignText As String, ByRef outHorizontal As Long, ByRef outVertical As Long) As Boolean
    Dim normalized As String

    normalized = LCase$(Trim$(alignText))
    normalized = Replace(normalized, "_", "-")
    normalized = Replace(normalized, " ", vbNullString)

    Select Case normalized
        Case "tl": normalized = "top-left"
        Case "tc": normalized = "top-center"
        Case "tr": normalized = "top-right"
        Case "ml": normalized = "middle-left"
        Case "mc": normalized = "middle-center"
        Case "mr": normalized = "middle-right"
        Case "bl": normalized = "bottom-left"
        Case "bc": normalized = "bottom-center"
        Case "br": normalized = "bottom-right"
    End Select

    Select Case normalized
        Case "top-left"
            outHorizontal = xlLeft: outVertical = xlTop
        Case "top-center"
            outHorizontal = xlCenter: outVertical = xlTop
        Case "top-right"
            outHorizontal = xlRight: outVertical = xlTop
        Case "middle-left"
            outHorizontal = xlLeft: outVertical = xlCenter
        Case "middle-center"
            outHorizontal = xlCenter: outVertical = xlCenter
        Case "middle-right"
            outHorizontal = xlRight: outVertical = xlCenter
        Case "bottom-left"
            outHorizontal = xlLeft: outVertical = xlBottom
        Case "bottom-center"
            outHorizontal = xlCenter: outVertical = xlBottom
        Case "bottom-right"
            outHorizontal = xlRight: outVertical = xlBottom
        Case Else
            Exit Function
    End Select

    mp_TryParseProfileConfigAlign = True
End Function

Private Sub mp_AutoFitConfigRangeRows(ByVal ws As Worksheet, ByVal tbl As ListObject)
    Dim topRow As Long
    Dim bottomRow As Long

    If ws Is Nothing Then Exit Sub
    If tbl Is Nothing Then Exit Sub
    If tbl.Range Is Nothing Then Exit Sub

    topRow = tbl.Range.Row
    bottomRow = topRow + tbl.Range.Rows.Count - 1
    If bottomRow < topRow Then Exit Sub

    ws.Rows(CStr(topRow) & ":" & CStr(bottomRow)).AutoFit
End Sub

' =============================================================================
' Profile UI helpers (migrated from ex_ProfileUI)
' =============================================================================
Public Sub m_ApplyProfileUI(ByVal ws As Worksheet, ByVal profileNode As Object, Optional ByVal profileName As String = vbNullString)
    Dim uiNodes As Object
    Dim node As Object
    Dim shapeName As String
    Dim shp As Shape

    If ws Is Nothing Then
        MsgBox "Failed to apply profile UI: worksheet is not specified.", vbExclamation
        Exit Sub
    End If
    If profileNode Is Nothing Then
        MsgBox "Failed to apply profile UI: profile node is not specified.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    profileNode.OwnerDocument.setProperty "SelectionNamespaces", "xmlns:p='" & PROFILES_NS & "'"
    On Error GoTo 0

    Set uiNodes = profileNode.selectNodes("p:ui/p:control")
    If uiNodes Is Nothing Then Exit Sub
    If uiNodes.Length = 0 Then Exit Sub

    For Each node In uiNodes
        shapeName = Trim$(mp_NodeAttrText(node, "name"))
        If Len(shapeName) = 0 Then
            MsgBox "Profile UI contains shape entry without 'name' attribute.", vbExclamation
            Exit Sub
        End If

        Set shp = m_GetShapeByName(ws, shapeName)
        If shp Is Nothing Then
            If StrComp(shapeName, LEGACY_PROFILE_DROPDOWN_SHAPE, vbTextCompare) = 0 Then GoTo NextNode
            If StrComp(shapeName, LEGACY_MODE_DROPDOWN_SHAPE, vbTextCompare) = 0 Then GoTo NextNode
            If pfui_IsButtonShapeName(shapeName) Then GoTo NextNode
            MsgBox "Profile UI shape '" & shapeName & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
            Exit Sub
        End If

        If Not pfui_ApplyShapeVisible(node, shp) Then Exit Sub

        Set shp = Nothing
NextNode:
    Next node
End Sub

Public Sub m_EnsureUiControlsAbsolute(Optional ByVal ws As Worksheet)
    Dim shp As Shape

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If
    If ws Is Nothing Then
        MsgBox "Failed to apply absolute UI layout: worksheet is not specified.", vbExclamation
        Exit Sub
    End If

    For Each shp In ws.Shapes
        If pfui_IsManagedUiBlockShape(shp.Name) Then
            On Error GoTo EH_PLACEMENT
            shp.Placement = xlFreeFloating
            On Error GoTo 0
        End If
    Next shp

    Exit Sub
EH_PLACEMENT:
    MsgBox "Failed to set absolute placement for shape '" & shp.Name & "': " & Err.Description, vbExclamation
End Sub

Public Sub m_InitUiBlockLayoutAndGroup(Optional ByVal ws As Worksheet)
    Dim shp As Shape
    Dim names As Variant
    Dim groupShape As Shape
    Dim shapeName As Variant

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If
    If ws Is Nothing Then
        MsgBox "Failed to initialize UI block group: worksheet is not specified.", vbExclamation
        Exit Sub
    End If

    m_EnsureUiControlsAbsolute ws
    On Error GoTo EH_UNGROUP
    pfui_UngroupManagedUiShapes ws
    On Error GoTo 0

    names = Array(PFUI_MODE_DROPDOWN_SHAPE, PFUI_CLEAR_BUTTON_SHAPE, PFUI_MODE_BUTTON_SHAPE, PFUI_PERSONAL_BUTTON_SHAPE, PFUI_COMPARING_BUTTON_SHAPE)
    For Each shapeName In names
        Set shp = m_GetShapeByName(ws, CStr(shapeName))
        If shp Is Nothing Then
            MsgBox "Failed to initialize UI block group: shape '" & CStr(shapeName) & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
            Exit Sub
        End If
    Next shapeName

    On Error GoTo EH_GROUP
    Set groupShape = ws.Shapes.Range(names).Group
    groupShape.Name = PFUI_UI_BLOCK_GROUP_NAME
    Exit Sub

EH_UNGROUP:
    MsgBox "Failed to ungroup existing UI block shapes before creating '" & PFUI_UI_BLOCK_GROUP_NAME & "': " & Err.Description, vbExclamation
    Exit Sub
EH_GROUP:
    MsgBox "Failed to create group '" & PFUI_UI_BLOCK_GROUP_NAME & "'. Group " & PFUI_MODE_DROPDOWN_SHAPE & " + buttons manually if needed: " & Err.Description, vbExclamation
End Sub

Public Sub m_ApplyModeVisibility(ByVal ws As Worksheet, ByVal profileNode As Object)
    Dim uiDefDoc As Object
    Dim uiControlNodes As Object
    Dim uiNodes As Object

    If ws Is Nothing Then
        MsgBox "Failed to apply mode visibility: worksheet is not specified.", vbExclamation
        Exit Sub
    End If
    If profileNode Is Nothing Then
        MsgBox "Failed to apply mode visibility: profile node is not specified.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    profileNode.OwnerDocument.setProperty "SelectionNamespaces", "xmlns:p='" & PROFILES_NS & "'"
    On Error GoTo 0

    pfui_HideAllButtons ws

    Set uiDefDoc = pfui_LoadUiDefinitionDom()
    If uiDefDoc Is Nothing Then Exit Sub

    Set uiControlNodes = uiDefDoc.selectNodes("/p:uiDefinition/p:controls/p:control")
    If uiControlNodes Is Nothing Then
        MsgBox "Invalid UI definition format. Expected '/uiDefinition/controls/control'.", vbExclamation
        Exit Sub
    End If
    pfui_ApplyGlobalVisibilityFromUiControls ws, uiControlNodes

    Set uiNodes = profileNode.selectNodes("p:ui/p:control")
    If uiNodes Is Nothing Then Exit Sub
    If uiNodes.Length = 0 Then Exit Sub
    pfui_ApplyFilteredVisibilityFromNodes ws, uiNodes
End Sub

Public Function m_GetShapeByName(ByVal ws As Worksheet, ByVal shapeName As String) As Shape
    If ws Is Nothing Then Exit Function
    Set m_GetShapeByName = pfui_FindShapeInContainer(ws.Shapes, shapeName)
End Function

Private Sub pfui_ApplyFilteredVisibilityFromNodes(ByVal ws As Worksheet, ByVal nodes As Object)
    Dim node As Object
    Dim shapeName As String
    Dim shp As Shape

    For Each node In nodes
        shapeName = Trim$(mp_NodeAttrText(node, "name"))
        If Len(shapeName) = 0 Then
            MsgBox "UI visibility block contains shape entry without 'name'.", vbExclamation
            Exit Sub
        End If
        If Not pfui_IsButtonShapeName(shapeName) Then GoTo NextNode

        Set shp = m_GetShapeByName(ws, shapeName)
        If shp Is Nothing Then GoTo NextNode

        If pfui_IsShapeVisibleByFilters(node) Then shp.Visible = msoTrue

        Set shp = Nothing
NextNode:
    Next node
End Sub

Private Sub pfui_ApplyGlobalVisibilityFromUiControls(ByVal ws As Worksheet, ByVal controlNodes As Object)
    Dim node As Object
    Dim shapeName As String
    Dim visibleText As String
    Dim isGlobalVisible As Boolean
    Dim shp As Shape

    For Each node In controlNodes
        shapeName = Trim$(mp_NodeAttrText(node, "name"))
        If Len(shapeName) = 0 Then
            MsgBox "UI control entry contains no 'name' attribute.", vbExclamation
            Exit Sub
        End If
        If Not pfui_IsButtonShapeName(shapeName) Then GoTo NextNode

        visibleText = Trim$(mp_NodeAttrText(node, "globalVisible"))
        If Len(visibleText) = 0 Then GoTo NextNode
        If Not pfui_TryParseBoolean(visibleText, isGlobalVisible) Then
            MsgBox "Invalid boolean value for UI control attribute 'globalVisible' on '" & shapeName & "': " & visibleText, vbExclamation
            Exit Sub
        End If
        If Not isGlobalVisible Then GoTo NextNode

        Set shp = m_GetShapeByName(ws, shapeName)
        If shp Is Nothing Then GoTo NextNode
        shp.Visible = msoTrue
NextNode:
    Next node
End Sub

Private Function pfui_ApplyShapeVisible(ByVal node As Object, ByVal shp As Shape) As Boolean
    Dim valueText As String
    Dim valueBool As Boolean

    valueText = Trim$(mp_NodeAttrText(node, "visible"))
    If Len(valueText) = 0 Then
        shp.Visible = msoFalse
        pfui_ApplyShapeVisible = True
        Exit Function
    End If

    If Not pfui_TryParseBoolean(valueText, valueBool) Then
        MsgBox "Invalid boolean value for UI attribute 'visible' on shape '" & shp.Name & "': " & valueText, vbExclamation
        Exit Function
    End If

    shp.Visible = IIf(valueBool, msoTrue, msoFalse)
    pfui_ApplyShapeVisible = True
End Function

Private Function pfui_IsShapeVisibleByFilters(ByVal node As Object) As Boolean
    Dim visibleText As String
    Dim isBaseVisible As Boolean

    visibleText = Trim$(mp_NodeAttrText(node, "visible"))
    isBaseVisible = False
    If Len(visibleText) > 0 Then
        If Not pfui_TryParseBoolean(visibleText, isBaseVisible) Then
            MsgBox "Invalid boolean value for UI attribute 'visible' in mode filter block: " & visibleText, vbExclamation
            Exit Function
        End If
    End If
    If Not isBaseVisible Then Exit Function
    pfui_IsShapeVisibleByFilters = True
End Function

Private Function pfui_TryParseBoolean(ByVal valueText As String, ByRef result As Boolean) As Boolean
    Select Case LCase$(Trim$(valueText))
        Case "1", "true", "yes"
            result = True
            pfui_TryParseBoolean = True
        Case "0", "false", "no"
            result = False
            pfui_TryParseBoolean = True
    End Select
End Function

Private Sub pfui_HideAllButtons(ByVal ws As Worksheet)
    pfui_HideAllButtonsInContainer ws.Shapes
End Sub

Private Sub pfui_HideAllButtonsInContainer(ByVal shapeContainer As Object)
    Dim shp As Shape
    Dim groupItem As Shape

    For Each shp In shapeContainer
        If pfui_IsButtonShapeName(shp.Name) Then shp.Visible = msoFalse
        If shp.Type = msoGroup Then
            For Each groupItem In shp.GroupItems
                If pfui_IsButtonShapeName(groupItem.Name) Then groupItem.Visible = msoFalse
            Next groupItem
        End If
    Next shp
End Sub

Private Function pfui_FindShapeInContainer(ByVal shapeContainer As Object, ByVal shapeName As String) As Shape
    Dim shp As Shape
    Dim groupItem As Shape
    Dim normalized As String

    normalized = Trim$(shapeName)
    If Len(normalized) = 0 Then Exit Function

    For Each shp In shapeContainer
        If StrComp(shp.Name, normalized, vbTextCompare) = 0 Then
            Set pfui_FindShapeInContainer = shp
            Exit Function
        End If
        If shp.Type = msoGroup Then
            For Each groupItem In shp.GroupItems
                If StrComp(groupItem.Name, normalized, vbTextCompare) = 0 Then
                    Set pfui_FindShapeInContainer = groupItem
                    Exit Function
                End If
            Next groupItem
        End If
    Next shp
End Function

Private Sub pfui_UngroupManagedUiShapes(ByVal ws As Worksheet)
    Dim hasGroupsToUngroup As Boolean
    Dim i As Long
    Dim shp As Shape

    Do
        hasGroupsToUngroup = False
        For i = ws.Shapes.Count To 1 Step -1
            Set shp = ws.Shapes(i)
            If shp.Type = msoGroup Then
                If pfui_GroupContainsManagedShapes(shp) Then
                    shp.Ungroup
                    hasGroupsToUngroup = True
                    Exit For
                End If
            End If
        Next i
    Loop While hasGroupsToUngroup
End Sub

Private Function pfui_GroupContainsManagedShapes(ByVal groupShape As Shape) As Boolean
    Dim groupItem As Shape

    For Each groupItem In groupShape.GroupItems
        If pfui_IsManagedUiBlockShape(groupItem.Name) Then
            pfui_GroupContainsManagedShapes = True
            Exit Function
        End If
    Next groupItem
End Function

Private Function pfui_IsButtonShapeName(ByVal shapeName As String) As Boolean
    pfui_IsButtonShapeName = (LCase$(Left$(Trim$(shapeName), 3)) = "btn")
End Function

Private Function pfui_IsManagedUiBlockShape(ByVal shapeName As String) As Boolean
    Dim normalized As String

    normalized = Trim$(shapeName)
    If Len(normalized) = 0 Then Exit Function

    If StrComp(normalized, PFUI_PROFILE_DROPDOWN_SHAPE, vbTextCompare) = 0 Then
        pfui_IsManagedUiBlockShape = True
        Exit Function
    End If

    If StrComp(normalized, LEGACY_PROFILE_DROPDOWN_SHAPE, vbTextCompare) = 0 Then
        pfui_IsManagedUiBlockShape = True
        Exit Function
    End If

    If StrComp(normalized, PFUI_MODE_DROPDOWN_SHAPE, vbTextCompare) = 0 Then
        pfui_IsManagedUiBlockShape = True
        Exit Function
    End If

    If StrComp(normalized, LEGACY_MODE_DROPDOWN_SHAPE, vbTextCompare) = 0 Then
        pfui_IsManagedUiBlockShape = True
        Exit Function
    End If

    If pfui_IsButtonShapeName(normalized) Then
        pfui_IsManagedUiBlockShape = ( _
            StrComp(normalized, PFUI_UPDATE_BUTTON_SHAPE, vbTextCompare) <> 0 _
            And StrComp(normalized, PFUI_UPDATE_UI_BUTTON_SHAPE, vbTextCompare) <> 0)
    End If
End Function

Private Function pfui_LoadUiDefinitionDom() As Object
    Set pfui_LoadUiDefinitionDom = ex_XmlCore.m_LoadDomByRelativePath( _
        ThisWorkbook, _
        PFUI_UI_DEFINITION_REL_PATH, _
        PROFILES_NS, _
        "UI definition config file was not found: ", _
        "Failed to parse UI definition config file: ")
End Function
