Attribute VB_Name = "ex_ConfigProfilesManager"
Option Explicit

' =============================================================================
' ex_ConfigProfilesManager
' =============================================================================
' Назначение:
' - читать/писать профили конфигурации из внешнего XML-файла
'   (пути задаются в `config\UI.xml` -> `dataSources/profilesSource`);
' - применять выбранный профиль к таблице `tblDevConfig` на листе Dev;
' - сохранять текущее состояние таблицы обратно в активный профиль;
' - поддерживать совместимость со старыми форматами (legacy row/marker layout);
' - синхронизировать визуальное оформление таблицы после загрузки профиля.
'
' Границы ответственности:
' - этот модуль работает с профильным XML и dropdown `ddProfile`;
' - бизнес-логика чтения значений конфигурации по ключу находится в ex_ConfigProvider.
' =============================================================================

Private Const PRESETS_NS As String = "urn:excelprototype:presets"
Private Const MODE_PERSONAL_CARD As String = "Personal Card"
Private Const MODE_TABLES_COMPARING As String = "Comparing"
Private Const DEV_CONFIG_TABLE_NAME As String = "tblDevConfig"
Private Const DEV_CONFIG_HEADER_ROW As Long = 2
Private Const DEV_CONFIG_MARKER_COL As Long = 1
Private Const DEV_CONFIG_KEY_COL As Long = 2
Private Const DEV_CONFIG_VALUE_COL As Long = 3
Private Const DEV_CONFIG_NOTE_COL As Long = 4
Private Const DEV_CONFIG_COL_COUNT As Long = 4
Private Const DEV_MARKER_SYMBOL As String = "#"
Private Const DEV_MARKER_HEADER As String = ".."
Private Const DEV_MARKER_PREFIX As String = "#MARKER:"
Private Const DEV_MARKER_SECTION As String = "#MARKER:SECTION"
Private Const DEV_MARKER_SPACER As String = "#MARKER:SPACER"
Private Const DEV_COLOR_BG As Long = &H1E1E1E
Private Const DEV_COLOR_TEXT As Long = &HEBEBEB
Private Const DEV_COLOR_BORDER As Long = &H505050
Private Const DEV_COLOR_NOTE_TEXT As Long = &HA8A8A8
Private Const THEME_BG As Long = &H262626
Private Const THEME_TEXT As Long = &HEBEBEB
Private Const THEME_BORDER As Long = &H0
Private Const PFUI_UI_DEFINITION_REL_PATH As String = "config\UI.xml"
Private Const PFUI_UI_BLOCK_GROUP_NAME As String = "grpUiBlock"
Private Const PFUI_PROFILE_DROPDOWN_SHAPE As String = "ddProfile"
Private Const PFUI_MODE_DROPDOWN_SHAPE As String = "ddMode"
Private Const PFUI_UPDATE_BUTTON_SHAPE As String = "btnUpdateCode"
Private Const PFUI_CLEAR_BUTTON_SHAPE As String = "btnClear"
Private Const PFUI_MODE_BUTTON_SHAPE As String = "btnMode"
Private Const PFUI_PERSONAL_BUTTON_SHAPE As String = "btnPersonalCard"
Private Const PFUI_COMPARING_BUTTON_SHAPE As String = "btnComparing"
Private Const STATE_ACTIVE_MODE_PROP As String = "Settings.ActiveModeName"
Private Const STATE_ACTIVE_PROFILE_PROP_PREFIX As String = "Settings.ActiveProfile."

' =============================================================================
' Public API (сверху по требованию рефакторинга)
' =============================================================================

' Применяет выбранный профиль в таблицу на листе Dev.
' Последовательность:
' 1) определяет активный профиль (аргумент или `ddProfile`);
' 2) загружает DOM профилей из XML;
' 3) читает узлы профиля в внутренний массив строк;
' 4) перезаписывает таблицу `tblDevConfig` и обновляет заголовок профиля.
Public Sub m_ApplyProfileFromDev(Optional ByVal profileName As String = vbNullString)
    Dim ws As Worksheet
    Dim doc As Object
    Dim profileNode As Object
    Dim entries As Variant
    Dim profiles As Variant
    Dim prevEvents As Boolean

    On Error GoTo EH
    prevEvents = Application.EnableEvents

    Set ws = ws_Dev

    If Len(profileName) = 0 Then
        profiles = mp_GetProfileNames(ws)
        If mp_ArrayHasItems(profiles) Then
            profileName = mp_GetSelectedProfileNameFromDropdown(ws, profiles, False)
        End If
    End If
    profileName = Trim$(profileName)
    If Len(profileName) = 0 Then Exit Sub

    Set doc = mp_LoadPresetsDom(ws)
    Set profileNode = mp_GetProfileNode(doc, profileName, False)
    If profileNode Is Nothing Then
        MsgBox "Profile '" & profileName & "' was not found in config file: " & mp_GetProfilesFilePath(), vbExclamation
        Exit Sub
    End If

    entries = ex_ProfilesEntriesMapper.m_ReadProfileEntries(ws, profileNode)

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    mp_WriteEntriesToConfigTable ws, entries
    On Error Resume Next
    ex_ConfigProvider.m_RefreshConfigTitle ws, profileName
    On Error GoTo 0
    ex_ConfigProfilesManager.m_ApplyProfileUI ws, profileNode, profileName
    mp_ApplyModeVisibility ws
EH:
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

    Set doc = mp_LoadPresetsDom(ws)
    Set profileNode = mp_GetProfileNode(doc, profileName, True)
    If profileNode Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    ex_ProfilesEntriesMapper.m_WriteSheetValuesToProfile ws, doc, profileNode

    mp_SavePresetsDom doc
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

    savedProfileName = mp_GetSavedProfileNameForMode(mp_GetSelectedModeName(ws))
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

    Set ws = ws_Dev
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
End Sub

Public Sub m_OnModeChanged()
    Dim ws As Worksheet
    Dim profiles As Variant

    Set ws = ws_Dev
    m_EnsureModeDropdown ws
    mp_ApplyModeVisibility ws
    profiles = mp_GetProfileNames(ws)
    If Not mp_ArrayHasItems(profiles) Then Exit Sub
    m_EnsureProfileDropdown ws
    mp_ReapplySelectedProfileUi ws
    m_SaveSelectionState ws
End Sub

Public Sub m_RestoreSelectionState(Optional ByVal ws As Worksheet)
    Dim savedModeName As String

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If

    m_EnsureModeDropdown ws
    savedModeName = mp_GetStatePropText(STATE_ACTIVE_MODE_PROP, vbNullString)
    If Len(savedModeName) > 0 Then
        mp_TrySelectModeByName ws, savedModeName
    End If

    mp_ApplyModeVisibility ws
    m_EnsureProfileDropdown ws
    mp_ReapplySelectedProfileUi ws
End Sub

Public Sub m_SaveSelectionState(Optional ByVal ws As Worksheet)
    Dim modeName As String
    Dim profileName As String
    Dim profiles As Variant

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If

    modeName = Trim$(mp_GetSelectedModeName(ws))
    If Len(modeName) = 0 Then Exit Sub
    mp_SetStatePropText STATE_ACTIVE_MODE_PROP, modeName

    profiles = mp_GetProfileNames(ws)
    If Not mp_ArrayHasItems(profiles) Then Exit Sub

    profileName = Trim$(mp_GetSelectedProfileNameFromDropdown(ws, profiles, False))
    If Len(profileName) = 0 Then Exit Sub

    mp_SetStatePropText STATE_ACTIVE_PROFILE_PROP_PREFIX & mp_NormalizePropSuffix(modeName), profileName
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

    Set doc = mp_LoadPresetsDom(ws)
    Set profileNode = mp_GetProfileNode(doc, profileName, True)
    If profileNode Is Nothing Then
        MsgBox "Failed to access profile '" & profileName & "' in config file: " & mp_GetProfilesFilePath(), vbExclamation
        Exit Sub
    End If

    ex_ProfilesEntriesMapper.m_WriteSheetValuesToProfile ws, doc, profileNode
    mp_SavePresetsDom doc
    m_SaveSelectionState ws
    On Error Resume Next
    ex_ConfigProvider.m_RefreshConfigTitle ws, profileName
    On Error GoTo 0
    Application.StatusBar = "Profiles config saved: " & profileName
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

    Set root = doc.selectSingleNode("/p:presets")
    If root Is Nothing Then Exit Sub

    Set profileNode = doc.createNode(1, "profile", PRESETS_NS)
    profileNode.setAttribute "name", profileName
    root.appendChild profileNode

    entries = ex_ProfilesEntriesMapper.m_ReadConfigTableEntries(ws)
    If Not mp_ArrayHasItems(entries) Then Exit Sub

    For i = LBound(entries, 1) To UBound(entries, 1)
        Set vNode = doc.createNode(1, "v", PRESETS_NS)
        If Len(Trim$(CStr(entries(i, DEV_CONFIG_MARKER_COL)))) > 0 Then
            vNode.setAttribute "type", CStr(entries(i, DEV_CONFIG_MARKER_COL))
        End If
        vNode.setAttribute "key", CStr(entries(i, DEV_CONFIG_KEY_COL))
        If Len(Trim$(CStr(entries(i, DEV_CONFIG_NOTE_COL)))) > 0 Then
            vNode.setAttribute "note", CStr(entries(i, DEV_CONFIG_NOTE_COL))
        End If
        vNode.Text = CStr(entries(i, DEV_CONFIG_VALUE_COL))
        profileNode.appendChild vNode
    Next i

    mp_SavePresetsDom doc
End Sub

Private Function mp_GetProfileNames(Optional ByVal ws As Worksheet) As Variant
    Dim modeName As String

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If

    modeName = mp_GetSelectedModeName(ws)
    mp_GetProfileNames = ex_UiXmlProvider.m_GetDropdownItemsByName("ddProfile", ThisWorkbook, modeName)
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
    Dim cf As Object
    Dim selectedIndex As Long
    Dim previousIndex As Long
    Dim previousName As String
    Dim matchedIndex As Long

    Set cf = mp_GetProfileDropdownControl(ws)
    If cf Is Nothing Then
        MsgBox "Profile control 'ddProfile' was not found on sheet '" & ws.Name & "'.", vbExclamation
        Exit Function
    End If

    previousIndex = mp_GetControlIndex(cf)
    previousName = mp_GetControlItemByIndex(cf, previousIndex)

    If syncItems Then
        mp_SetDropdownItems cf, profiles
        matchedIndex = mp_FindProfileIndexByName(profiles, preferredName)
        If matchedIndex = 0 Then
            matchedIndex = mp_FindProfileIndexByName(profiles, previousName)
        End If
        If matchedIndex > 0 Then
            On Error Resume Next
            cf.Value = matchedIndex
            On Error GoTo 0
        ElseIf previousIndex >= 1 And previousIndex <= mp_ArrayLength(profiles) Then
            On Error Resume Next
            cf.Value = previousIndex
            On Error GoTo 0
        End If
    End If

    selectedIndex = mp_GetControlIndex(cf)
    If selectedIndex < 1 Or selectedIndex > mp_ArrayLength(profiles) Then
        selectedIndex = 1
        On Error Resume Next
        cf.Value = selectedIndex
        On Error GoTo 0
    End If

    mp_GetSelectedProfileNameFromDropdown = CStr(profiles(selectedIndex - 1))
End Function

Private Sub m_EnsureModeDropdown(ByVal ws As Worksheet)
    Dim cf As Object
    Dim selectedIndex As Long
    Dim modeNames As Variant

    Set cf = mp_GetModeDropdownControl(ws)
    If cf Is Nothing Then
        MsgBox "Mode control 'ddMode' was not found on sheet '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If

    modeNames = ex_UiXmlProvider.m_GetDropdownItemsByName("ddMode", ThisWorkbook)
    If Not mp_ArrayHasItems(modeNames) Then
        MsgBox "Mode control 'ddMode' has no <items> in config\UI.xml.", vbExclamation
        Exit Sub
    End If

    selectedIndex = mp_GetControlIndex(cf)
    mp_SetDropdownItems cf, modeNames

    If selectedIndex < 1 Or selectedIndex > mp_ArrayLength(modeNames) Then
        selectedIndex = 1
    End If

    On Error Resume Next
    cf.Value = selectedIndex
    On Error GoTo 0
End Sub

Private Sub mp_TrySelectModeByName(ByVal ws As Worksheet, ByVal modeName As String)
    Dim cf As Object
    Dim modeNames As Variant
    Dim targetIndex As Long

    modeName = Trim$(modeName)
    If Len(modeName) = 0 Then Exit Sub

    Set cf = mp_GetModeDropdownControl(ws)
    If cf Is Nothing Then Exit Sub

    modeNames = ex_UiXmlProvider.m_GetDropdownItemsByName("ddMode", ThisWorkbook)
    If Not mp_ArrayHasItems(modeNames) Then Exit Sub

    targetIndex = mp_FindProfileIndexByName(modeNames, modeName)
    If targetIndex <= 0 Then Exit Sub

    On Error Resume Next
    cf.Value = targetIndex
    On Error GoTo 0
End Sub

Private Function mp_GetSelectedModeName(ByVal ws As Worksheet) As String
    Dim cf As Object
    Dim selectedIndex As Long
    Dim selectedName As String

    Set cf = mp_GetModeDropdownControl(ws)
    If cf Is Nothing Then
        mp_GetSelectedModeName = MODE_PERSONAL_CARD
        Exit Function
    End If

    selectedIndex = mp_GetControlIndex(cf)
    selectedName = mp_GetControlItemByIndex(cf, selectedIndex)
    If Len(selectedName) = 0 Then
        mp_GetSelectedModeName = MODE_PERSONAL_CARD
    Else
        mp_GetSelectedModeName = selectedName
    End If
End Function

Private Function mp_GetModeDropdownControl(ByVal ws As Worksheet) As Object
    Dim shp As Shape

    Set shp = ex_ConfigProfilesManager.m_GetShapeByName(ws, "ddMode")
    If shp Is Nothing Then Exit Function

    On Error Resume Next
    Set mp_GetModeDropdownControl = shp.ControlFormat
    On Error GoTo 0
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

    Set doc = mp_LoadPresetsDom(ws)
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

    Set doc = mp_LoadPresetsDom(ws)
    Set profileNode = mp_GetProfileNode(doc, profileName, False)
    If profileNode Is Nothing Then Exit Sub

    ex_ConfigProfilesManager.m_ApplyProfileUI ws, profileNode, profileName
    ex_ConfigProfilesManager.m_ApplyModeVisibility ws, profileNode
End Sub

Private Function mp_GetControlItemByIndex(ByVal cf As Object, ByVal itemIndex As Long) As String
    On Error Resume Next
    If itemIndex >= 1 Then
        mp_GetControlItemByIndex = CStr(cf.List(itemIndex))
    End If
    On Error GoTo 0
End Function

Private Function mp_GetProfileDropdownControl(ByVal ws As Worksheet) As Object
    Dim shp As Shape

    Set shp = ex_ConfigProfilesManager.m_GetShapeByName(ws, "ddProfile")
    If shp Is Nothing Then Exit Function

    On Error Resume Next
    Set mp_GetProfileDropdownControl = shp.ControlFormat
    On Error GoTo 0
End Function

Private Sub mp_SetDropdownItems(ByVal cf As Object, ByVal profiles As Variant)
    Dim i As Long

    On Error Resume Next
    cf.RemoveAllItems
    If Err.Number <> 0 Then
        MsgBox "Failed to clear dropdown control items: " & Err.Description, vbExclamation
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    For i = LBound(profiles) To UBound(profiles)
        On Error Resume Next
        cf.AddItem CStr(profiles(i))
        If Err.Number <> 0 Then
            MsgBox "Failed to add dropdown item '" & CStr(profiles(i)) & "': " & Err.Description, vbExclamation
            Err.Clear
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
    Next i
End Sub

Private Function mp_GetControlIndex(ByVal cf As Object) As Long
    On Error Resume Next
    mp_GetControlIndex = CLng(cf.Value)
    If Err.Number <> 0 Then
        Err.Clear
        mp_GetControlIndex = 0
    End If
    On Error GoTo 0
End Function

Private Function mp_ArrayLength(ByVal values As Variant) As Long
    If Not mp_ArrayHasItems(values) Then Exit Function
    mp_ArrayLength = UBound(values) - LBound(values) + 1
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

Private Function mp_GetSavedProfileNameForMode(ByVal modeName As String) As String
    Dim propName As String

    modeName = Trim$(modeName)
    If Len(modeName) = 0 Then Exit Function

    propName = STATE_ACTIVE_PROFILE_PROP_PREFIX & mp_NormalizePropSuffix(modeName)
    mp_GetSavedProfileNameForMode = mp_GetStatePropText(propName, vbNullString)
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

Private Function mp_LoadPresetsDom(Optional ByVal ws As Worksheet) As Object
    Dim filePath As String
    Dim modeName As String

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If

    modeName = mp_GetSelectedModeName(ws)
    filePath = mp_GetProfilesFilePath(modeName)
    If Len(filePath) = 0 Then Exit Function

    Set mp_LoadPresetsDom = ex_ProfilesStore.m_LoadPresetsDom(filePath)
End Function

Private Sub mp_SavePresetsDom(ByVal doc As Object)
    Dim filePath As String
    Dim modeName As String

    modeName = mp_GetSelectedModeName(ws_Dev)
    filePath = mp_GetProfilesFilePath(modeName)
    ex_ProfilesStore.m_SavePresetsDom doc, filePath
End Sub


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
' - очищает старое содержимое/формат в прежних границах;
' - ресайзит таблицу под новый объём;
' - накладывает тему и маркерные стили.
Private Sub mp_WriteEntriesToConfigTable(ByVal ws As Worksheet, ByVal entries As Variant)
    Dim tbl As ListObject
    Dim previousRowCount As Long
    Dim rowCount As Long
    Dim values() As Variant
    Dim i As Long

    Set tbl = ex_ConfigTableStore.m_GetConfigTable(ws, True)
    If tbl Is Nothing Then
        MsgBox "Config table '" & DEV_CONFIG_TABLE_NAME & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If

    previousRowCount = ex_ConfigTableStore.m_GetTableDataRowCount(tbl)
    ex_ConfigTableStore.m_ClearConfigDataArea ws, tbl

    rowCount = mp_ArrayRowCount(entries)
    ex_ConfigTableStore.m_ResizeConfigTableRows ws, tbl, rowCount
    ex_ConfigTableStore.m_ApplySheetThemeToFormerTableTail ws, tbl, previousRowCount, rowCount

    If rowCount = 0 Then Exit Sub

    ReDim values(1 To rowCount, 1 To DEV_CONFIG_COL_COUNT)
    For i = 1 To rowCount
        values(i, DEV_CONFIG_MARKER_COL) = CStr(entries(i, DEV_CONFIG_MARKER_COL))
        values(i, DEV_CONFIG_KEY_COL) = CStr(entries(i, DEV_CONFIG_KEY_COL))
        values(i, DEV_CONFIG_VALUE_COL) = CStr(entries(i, DEV_CONFIG_VALUE_COL))
        values(i, DEV_CONFIG_NOTE_COL) = CStr(entries(i, DEV_CONFIG_NOTE_COL))
    Next i

    tbl.DataBodyRange.Cells(1, 1).Resize(rowCount, DEV_CONFIG_COL_COUNT).Value = values
    ex_ConfigTableStore.m_ApplyConfigTableDarkTheme tbl
    ex_ConfigTableStore.m_ApplyConfigMarkerStyles tbl
End Sub


Private Function mp_ArrayRowCount(ByVal values As Variant) As Long
    On Error GoTo EH
    If IsArray(values) Then
        mp_ArrayRowCount = UBound(values, 1) - LBound(values, 1) + 1
    End If
    Exit Function
EH:
    mp_ArrayRowCount = 0
End Function


Private Function mp_GetProfilesFilePath(Optional ByVal modeName As String = vbNullString) As String
    Dim resolvedMode As String

    resolvedMode = Trim$(modeName)
    If Len(resolvedMode) = 0 Then
        resolvedMode = MODE_PERSONAL_CARD
        On Error Resume Next
        resolvedMode = mp_GetSelectedModeName(ws_Dev)
        On Error GoTo 0
    End If

    mp_GetProfilesFilePath = ex_ProfilesStore.m_GetProfilesFilePath(resolvedMode, ThisWorkbook)
End Function

Private Function mp_GetProfileNode(ByVal doc As Object, ByVal profileName As String, ByVal createIfMissing As Boolean) As Object
    Set mp_GetProfileNode = ex_ProfilesStore.m_GetProfileNode(doc, profileName, createIfMissing)
End Function

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
    profileNode.OwnerDocument.setProperty "SelectionNamespaces", "xmlns:p='" & PRESETS_NS & "'"
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
    MsgBox "Failed to create group '" & PFUI_UI_BLOCK_GROUP_NAME & "'. Group ddMode + buttons manually if needed: " & Err.Description, vbExclamation
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
    profileNode.OwnerDocument.setProperty "SelectionNamespaces", "xmlns:p='" & PRESETS_NS & "'"
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

    If StrComp(normalized, PFUI_MODE_DROPDOWN_SHAPE, vbTextCompare) = 0 Then
        pfui_IsManagedUiBlockShape = True
        Exit Function
    End If

    If pfui_IsButtonShapeName(normalized) Then
        pfui_IsManagedUiBlockShape = (StrComp(normalized, PFUI_UPDATE_BUTTON_SHAPE, vbTextCompare) <> 0)
    End If
End Function

Private Function pfui_LoadUiDefinitionDom() As Object
    Set pfui_LoadUiDefinitionDom = ex_XmlCore.m_LoadDomByRelativePath( _
        ThisWorkbook, _
        PFUI_UI_DEFINITION_REL_PATH, _
        PRESETS_NS, _
        "UI definition config file was not found: ", _
        "Failed to parse UI definition config file: ")
End Function
