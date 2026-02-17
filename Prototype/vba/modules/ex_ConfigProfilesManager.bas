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
' - бизнес-логика чтения значений конфигурации по ключу находится в ex_Config.
' =============================================================================

Private Const PRESETS_NS As String = "urn:excelprototype:presets"
Private Const PRESETS_TEMPLATE As String = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><presets xmlns=""" & PRESETS_NS & """ version=""1""/>"
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

    entries = mp_ReadProfileEntries(ws, profileNode)

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    mp_WriteEntriesToConfigTable ws, entries
    On Error Resume Next
    ex_Config.m_RefreshConfigTitle ws, profileName
    On Error GoTo 0
    ex_ProfileUI.m_ApplyProfileUI ws, profileNode, profileName
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

    Set tbl = mp_GetConfigTable(ws, True)
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

    mp_WriteSheetValuesToProfile ws, doc, profileNode

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

    profileName = mp_GetSelectedProfileNameFromDropdown(ws, profiles, True)
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

    mp_WriteSheetValuesToProfile ws, doc, profileNode
    mp_SavePresetsDom doc
    On Error Resume Next
    ex_Config.m_RefreshConfigTitle ws, profileName
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

    entries = mp_ReadConfigTableEntries(ws)
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
    mp_GetProfileNames = ex_UILoader.m_GetDropdownItemsByName("ddProfile", ThisWorkbook, modeName)
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

Private Function mp_GetSelectedProfileNameFromDropdown(ByVal ws As Worksheet, ByVal profiles As Variant, Optional ByVal syncItems As Boolean = False) As String
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
        matchedIndex = mp_FindProfileIndexByName(profiles, previousName)
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

    modeNames = ex_UILoader.m_GetDropdownItemsByName("ddMode", ThisWorkbook)
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

    Set shp = ex_ProfileUI.m_GetShapeByName(ws, "ddMode")
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

    ex_ProfileUI.m_ApplyModeVisibility ws, profileNode
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

    Set shp = ex_ProfileUI.m_GetShapeByName(ws, "ddProfile")
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

Private Function mp_LoadPresetsDom(Optional ByVal ws As Worksheet) As Object
    Dim filePath As String
    Dim doc As Object
    Dim modeName As String

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If

    modeName = mp_GetSelectedModeName(ws)
    filePath = mp_GetProfilesFilePath(modeName)
    If Len(filePath) = 0 Then Exit Function

    Set doc = CreateObject("MSXML2.DOMDocument.6.0")

    doc.async = False
    doc.validateOnParse = False

    If Len(Dir(filePath)) > 0 Then
        If Not doc.Load(filePath) Then
            doc.loadXML PRESETS_TEMPLATE
        End If
    Else
        doc.loadXML PRESETS_TEMPLATE
    End If

    doc.setProperty "SelectionNamespaces", "xmlns:p='" & PRESETS_NS & "'"

    Set mp_LoadPresetsDom = doc
End Function

Private Sub mp_SavePresetsDom(ByVal doc As Object)
    Dim filePath As String
    Dim modeName As String

    modeName = mp_GetSelectedModeName(ws_Dev)
    filePath = mp_GetProfilesFilePath(modeName)
    If Len(filePath) = 0 Then Exit Sub

    If Len(Dir(filePath)) = 0 Then
        MsgBox "Profiles config file was not found: " & filePath, vbExclamation
        Exit Sub
    End If

    On Error GoTo EH
    mp_SaveXmlPretty doc, filePath
    Exit Sub
EH:
    MsgBox "Failed to save profiles config file '" & filePath & "': " & Err.Description, vbExclamation
End Sub

Private Sub mp_SaveXmlPretty(ByVal doc As Object, ByVal filePath As String)
    Dim reader As Object
    Dim writer As Object
    Dim stream As Object
    Dim xmlText As String

    Set writer = CreateObject("MSXML2.MXXMLWriter.6.0")
    writer.omitXMLDeclaration = False
    writer.indent = True
    writer.standalone = True
    writer.encoding = "UTF-8"

    Set reader = CreateObject("MSXML2.SAXXMLReader.6.0")
    Set reader.contentHandler = writer
    Set reader.dtdHandler = writer
    Set reader.errorHandler = writer
    On Error Resume Next
    reader.putProperty "http://xml.org/sax/properties/lexical-handler", writer
    On Error GoTo 0

    reader.parse doc.XML
    xmlText = CStr(writer.output)

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText xmlText
    stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    stream.Close
End Sub

Private Sub mp_WriteSheetValuesToProfile(ByVal ws As Worksheet, ByVal doc As Object, ByVal profileNode As Object)
    Dim entries As Variant
    Dim i As Long
    Dim vNode As Object
    Dim child As Object

    For Each child In profileNode.selectNodes("p:v")
        profileNode.removeChild child
    Next child

    entries = mp_ReadConfigTableEntries(ws)
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
End Sub

' Читает профиль в универсальный массив [marker, key, value, note].
' Поддерживает 2 формата:
' - актуальный: атрибуты key/type/note (+ опциональный legacy idx, который игнорируется);
' - legacy: старый `row`-подход, который мапится на текущую структуру таблицы.
Private Function mp_ReadProfileEntries(ByVal ws As Worksheet, ByVal profileNode As Object) As Variant
    Dim nodes As Object
    Dim hasKeyFormat As Boolean
    Dim i As Long
    Dim node As Object
    Dim entries() As Variant

    Set nodes = profileNode.selectNodes("p:v")
    If nodes Is Nothing Then
        mp_ReadProfileEntries = Array()
        Exit Function
    End If
    If nodes.Length = 0 Then
        mp_ReadProfileEntries = Array()
        Exit Function
    End If

    hasKeyFormat = False
    For i = 0 To nodes.Length - 1
        If Len(mp_NodeAttrText(nodes.Item(i), "key")) > 0 _
           Or Len(mp_NodeAttrText(nodes.Item(i), "type")) > 0 _
           Or Len(mp_NodeAttrText(nodes.Item(i), "note")) > 0 Then
            hasKeyFormat = True
            Exit For
        End If
    Next i

    If hasKeyFormat Then
        ReDim entries(1 To nodes.Length, 1 To DEV_CONFIG_COL_COUNT)
        For i = 0 To nodes.Length - 1
            Set node = nodes.Item(i)
            entries(i + 1, DEV_CONFIG_MARKER_COL) = mp_NodeAttrText(node, "type")
            entries(i + 1, DEV_CONFIG_KEY_COL) = mp_NodeAttrText(node, "key")
            entries(i + 1, DEV_CONFIG_VALUE_COL) = CStr(node.Text)
            entries(i + 1, DEV_CONFIG_NOTE_COL) = mp_NodeAttrText(node, "note")
            mp_NormalizeLegacyMarkerEntry entries, i + 1
        Next i
        mp_ReadProfileEntries = entries
        Exit Function
    End If

    ' Если профиль не содержит новой разметки, пробуем legacy-формат.
    mp_ReadProfileEntries = mp_ReadLegacyProfileEntries(ws, nodes)
End Function

' Legacy-конвертер: переносит старые значения по `row` в текущую табличную модель.
' Ключи и заметки для таких строк берутся из текущей таблицы Dev.
Private Function mp_ReadLegacyProfileEntries(ByVal ws As Worksheet, ByVal nodes As Object) As Variant
    Dim tbl As ListObject
    Dim tableValues As Variant
    Dim rowCount As Long
    Dim entries() As Variant
    Dim i As Long
    Dim rowAttr As String
    Dim entryIndex As Long
    Dim node As Object
    Dim maxIndex As Long

    Set tbl = mp_GetConfigTable(ws, True)
    If tbl Is Nothing Then
        MsgBox "Config table '" & DEV_CONFIG_TABLE_NAME & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        mp_ReadLegacyProfileEntries = Array()
        Exit Function
    End If

    rowCount = mp_GetTableDataRowCount(tbl)
    maxIndex = mp_GetMaxLegacyIndex(nodes)
    If maxIndex > rowCount Then
        rowCount = maxIndex
        mp_ResizeConfigTableRows ws, tbl, rowCount
    End If
    If rowCount = 0 Then
        mp_ReadLegacyProfileEntries = Array()
        Exit Function
    End If

    tableValues = mp_ReadConfigTableValues(tbl)
    ReDim entries(1 To rowCount, 1 To DEV_CONFIG_COL_COUNT)

    For i = 1 To rowCount
        entries(i, DEV_CONFIG_MARKER_COL) = CStr(tableValues(i, DEV_CONFIG_MARKER_COL))
        entries(i, DEV_CONFIG_KEY_COL) = CStr(tableValues(i, DEV_CONFIG_KEY_COL))
        entries(i, DEV_CONFIG_VALUE_COL) = vbNullString
        entries(i, DEV_CONFIG_NOTE_COL) = CStr(tableValues(i, DEV_CONFIG_NOTE_COL))
    Next i

    For i = 0 To nodes.Length - 1
        Set node = nodes.Item(i)
        rowAttr = mp_NodeAttrText(node, "row")
        If Len(rowAttr) > 0 And IsNumeric(rowAttr) Then
            entryIndex = CLng(rowAttr) - DEV_CONFIG_HEADER_ROW
            If entryIndex >= 1 And entryIndex <= rowCount Then
                entries(entryIndex, DEV_CONFIG_VALUE_COL) = CStr(node.Text)
            End If
        End If
    Next i

    mp_ReadLegacyProfileEntries = entries
End Function

Private Function mp_GetMaxLegacyIndex(ByVal nodes As Object) As Long
    Dim i As Long
    Dim rowAttr As String
    Dim idx As Long

    For i = 0 To nodes.Length - 1
        rowAttr = mp_NodeAttrText(nodes.Item(i), "row")
        If Len(rowAttr) > 0 And IsNumeric(rowAttr) Then
            idx = CLng(rowAttr) - DEV_CONFIG_HEADER_ROW
            If idx > mp_GetMaxLegacyIndex Then
                mp_GetMaxLegacyIndex = idx
            End If
        End If
    Next i
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
' - очищает старое содержимое/формат в прежних границах;
' - ресайзит таблицу под новый объём;
' - накладывает тему и маркерные стили.
Private Sub mp_WriteEntriesToConfigTable(ByVal ws As Worksheet, ByVal entries As Variant)
    Dim tbl As ListObject
    Dim previousRowCount As Long
    Dim rowCount As Long
    Dim values() As Variant
    Dim i As Long

    Set tbl = mp_GetConfigTable(ws, True)
    If tbl Is Nothing Then
        MsgBox "Config table '" & DEV_CONFIG_TABLE_NAME & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If

    previousRowCount = mp_GetTableDataRowCount(tbl)
    mp_ClearConfigDataArea ws, tbl

    rowCount = mp_ArrayRowCount(entries)
    mp_ResizeConfigTableRows ws, tbl, rowCount
    mp_ApplySheetThemeToFormerTableTail ws, tbl, previousRowCount, rowCount

    If rowCount = 0 Then Exit Sub

    ReDim values(1 To rowCount, 1 To DEV_CONFIG_COL_COUNT)
    For i = 1 To rowCount
        values(i, DEV_CONFIG_MARKER_COL) = CStr(entries(i, DEV_CONFIG_MARKER_COL))
        values(i, DEV_CONFIG_KEY_COL) = CStr(entries(i, DEV_CONFIG_KEY_COL))
        values(i, DEV_CONFIG_VALUE_COL) = CStr(entries(i, DEV_CONFIG_VALUE_COL))
        values(i, DEV_CONFIG_NOTE_COL) = CStr(entries(i, DEV_CONFIG_NOTE_COL))
    Next i

    tbl.DataBodyRange.Cells(1, 1).Resize(rowCount, DEV_CONFIG_COL_COUNT).Value = values
    mp_ApplyConfigTableDarkTheme tbl
    mp_ApplyConfigMarkerStyles tbl
End Sub

Private Function mp_ReadConfigTableEntries(ByVal ws As Worksheet) As Variant
    Dim tbl As ListObject
    Dim values As Variant
    Dim entries() As Variant
    Dim rowCount As Long
    Dim i As Long

    Set tbl = mp_GetConfigTable(ws, True)
    If tbl Is Nothing Then
        MsgBox "Config table '" & DEV_CONFIG_TABLE_NAME & "' was not found on sheet '" & ws.Name & "'.", vbExclamation
        mp_ReadConfigTableEntries = Array()
        Exit Function
    End If

    rowCount = mp_GetTableDataRowCount(tbl)
    If rowCount = 0 Then
        mp_ReadConfigTableEntries = Array()
        Exit Function
    End If

    values = mp_ReadConfigTableValues(tbl)
    ReDim entries(1 To rowCount, 1 To DEV_CONFIG_COL_COUNT)
    For i = 1 To rowCount
        entries(i, DEV_CONFIG_MARKER_COL) = CStr(values(i, DEV_CONFIG_MARKER_COL))
        entries(i, DEV_CONFIG_KEY_COL) = CStr(values(i, DEV_CONFIG_KEY_COL))
        entries(i, DEV_CONFIG_VALUE_COL) = CStr(values(i, DEV_CONFIG_VALUE_COL))
        entries(i, DEV_CONFIG_NOTE_COL) = CStr(values(i, DEV_CONFIG_NOTE_COL))
        mp_NormalizeLegacyMarkerEntry entries, i
    Next i

    mp_ReadConfigTableEntries = entries
End Function

Private Function mp_ReadConfigTableValues(ByVal tbl As ListObject) As Variant
    Dim rowCount As Long
    Dim rawValues As Variant
    Dim values() As Variant

    rowCount = mp_GetTableDataRowCount(tbl)
    If rowCount = 0 Then
        mp_ReadConfigTableValues = Array()
        Exit Function
    End If

    rawValues = tbl.DataBodyRange.Cells(1, 1).Resize(rowCount, DEV_CONFIG_COL_COUNT).Value
    If rowCount = 1 Then
        ReDim values(1 To 1, 1 To DEV_CONFIG_COL_COUNT)
        values(1, 1) = rawValues(1, 1)
        values(1, 2) = rawValues(1, 2)
        values(1, 3) = rawValues(1, 3)
        values(1, 4) = rawValues(1, 4)
        mp_ReadConfigTableValues = values
        Exit Function
    End If

    mp_ReadConfigTableValues = rawValues
End Function

Private Function mp_GetConfigTable(ByVal ws As Worksheet, Optional ByVal createIfMissing As Boolean = False) As ListObject
    Dim tbl As ListObject

    On Error Resume Next
    Set tbl = ws.ListObjects(DEV_CONFIG_TABLE_NAME)
    On Error GoTo 0

    If Not tbl Is Nothing Then
        mp_EnsureConfigTableLayout ws, tbl
    End If

    If tbl Is Nothing And createIfMissing Then
        Set tbl = mp_CreateConfigTable(ws)
    End If

    Set mp_GetConfigTable = tbl
End Function

Private Function mp_CreateConfigTable(ByVal ws As Worksheet) As ListObject
    Dim lastRow As Long
    Dim rangeToTable As Range

    If Trim$(CStr(ws.Cells(DEV_CONFIG_HEADER_ROW, DEV_CONFIG_MARKER_COL).Value)) <> DEV_MARKER_HEADER Then
        ws.Cells(DEV_CONFIG_HEADER_ROW, DEV_CONFIG_MARKER_COL).Value = DEV_MARKER_HEADER
    End If
    If Trim$(CStr(ws.Cells(DEV_CONFIG_HEADER_ROW, DEV_CONFIG_KEY_COL).Value)) <> "Key" Then
        ws.Cells(DEV_CONFIG_HEADER_ROW, DEV_CONFIG_KEY_COL).Value = "Key"
    End If
    If Trim$(CStr(ws.Cells(DEV_CONFIG_HEADER_ROW, DEV_CONFIG_NOTE_COL).Value)) <> "Note" Then
        ws.Cells(DEV_CONFIG_HEADER_ROW, DEV_CONFIG_NOTE_COL).Value = "Note"
    End If

    lastRow = mp_GetLastConfigRow(ws)
    If lastRow < DEV_CONFIG_HEADER_ROW Then
        lastRow = DEV_CONFIG_HEADER_ROW
    End If

    Set rangeToTable = ws.Range( _
        ws.Cells(DEV_CONFIG_HEADER_ROW, DEV_CONFIG_MARKER_COL), _
        ws.Cells(lastRow, DEV_CONFIG_NOTE_COL) _
    )

    On Error Resume Next
    Set mp_CreateConfigTable = ws.ListObjects.Add(xlSrcRange, rangeToTable, , xlYes)
    If Err.Number <> 0 Then
        MsgBox "Failed to create config table on sheet '" & ws.Name & "': " & Err.Description, vbExclamation
        Err.Clear
        On Error GoTo 0
        Set mp_CreateConfigTable = Nothing
        Exit Function
    End If
    On Error GoTo 0

    mp_EnsureConfigTableLayout ws, mp_CreateConfigTable

    On Error Resume Next
    mp_CreateConfigTable.Name = DEV_CONFIG_TABLE_NAME
    On Error GoTo 0

    On Error Resume Next
    ex_Config.m_RefreshConfigTitle ws
    On Error GoTo 0
End Function

Private Sub mp_ResizeConfigTableRows(ByVal ws As Worksheet, ByVal tbl As ListObject, ByVal rowCount As Long)
    Dim topRow As Long
    Dim leftCol As Long
    Dim bottomRow As Long
    Dim rightCol As Long
    Dim resizeRange As Range

    If rowCount < 0 Then rowCount = 0

    topRow = tbl.HeaderRowRange.Row
    leftCol = tbl.Range.Column
    rightCol = leftCol + DEV_CONFIG_COL_COUNT - 1
    bottomRow = topRow + rowCount

    Set resizeRange = ws.Range(ws.Cells(topRow, leftCol), ws.Cells(bottomRow, rightCol))
    tbl.Resize resizeRange
End Sub

Private Function mp_GetTableDataRowCount(ByVal tbl As ListObject) As Long
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    mp_GetTableDataRowCount = tbl.DataBodyRange.Rows.Count
End Function

Private Function mp_GetLastConfigRow(ByVal ws As Worksheet) As Long
    Dim lastKey As Long
    Dim lastValue As Long

    lastKey = ws.Cells(ws.Rows.Count, DEV_CONFIG_MARKER_COL).End(xlUp).Row
    lastValue = ws.Cells(ws.Rows.Count, DEV_CONFIG_VALUE_COL).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, DEV_CONFIG_NOTE_COL).End(xlUp).Row > lastValue Then
        lastValue = ws.Cells(ws.Rows.Count, DEV_CONFIG_NOTE_COL).End(xlUp).Row
    End If

    mp_GetLastConfigRow = lastKey
    If lastValue > mp_GetLastConfigRow Then
        mp_GetLastConfigRow = lastValue
    End If
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

Private Sub mp_ClearConfigDataArea(ByVal ws As Worksheet, ByVal tbl As ListObject)
    Dim rowCount As Long
    Dim topRow As Long
    Dim leftCol As Long
    Dim rightCol As Long
    Dim clearRange As Range

    rowCount = mp_GetTableDataRowCount(tbl)
    If rowCount <= 0 Then Exit Sub

    topRow = tbl.HeaderRowRange.Row + 1
    leftCol = tbl.Range.Column
    rightCol = leftCol + DEV_CONFIG_COL_COUNT - 1

    Set clearRange = ws.Range(ws.Cells(topRow, leftCol), ws.Cells(topRow + rowCount - 1, rightCol))
    clearRange.Clear
End Sub

Private Sub mp_ApplySheetThemeToFormerTableTail(ByVal ws As Worksheet, ByVal tbl As ListObject, ByVal previousRowCount As Long, ByVal newRowCount As Long)
    Dim topRow As Long
    Dim leftCol As Long
    Dim rightCol As Long
    Dim previousBottom As Long
    Dim newBottom As Long
    Dim tailRange As Range

    If previousRowCount <= newRowCount Then Exit Sub

    topRow = tbl.HeaderRowRange.Row
    leftCol = tbl.Range.Column
    rightCol = leftCol + DEV_CONFIG_COL_COUNT - 1
    previousBottom = topRow + previousRowCount
    newBottom = topRow + newRowCount

    If previousBottom <= newBottom Then Exit Sub

    Set tailRange = ws.Range(ws.Cells(newBottom + 1, leftCol), ws.Cells(previousBottom, rightCol))
    With tailRange
        .Interior.Pattern = xlSolid
        .Interior.Color = THEME_BG
        .Font.Color = THEME_TEXT
        .Borders.LineStyle = xlContinuous
        .Borders.Color = THEME_BORDER
        .Borders.Weight = xlThin
    End With
End Sub

Private Sub mp_ApplyConfigMarkerStyles(ByVal tbl As ListObject)
    Dim rowCount As Long
    Dim i As Long
    Dim keyText As String
    Dim markerKind As String
    Dim rowRange As Range

    rowCount = mp_GetTableDataRowCount(tbl)
    If rowCount <= 0 Then Exit Sub

    For i = 1 To rowCount
        Set rowRange = tbl.DataBodyRange.Cells(i, 1).Resize(1, DEV_CONFIG_COL_COUNT)
        keyText = Trim$(CStr(rowRange.Cells(1, DEV_CONFIG_KEY_COL).Value))
        markerKind = Trim$(CStr(rowRange.Cells(1, DEV_CONFIG_MARKER_COL).Value))
        If StrComp(markerKind, DEV_MARKER_SYMBOL, vbTextCompare) <> 0 And Not mp_IsMarkerKey(keyText) Then
            GoTo NextRow
        End If

        rowRange.Cells(1, DEV_CONFIG_MARKER_COL).Value = DEV_MARKER_SYMBOL

        rowRange.Interior.Pattern = xlSolid
        rowRange.Interior.Color = RGB(45, 45, 45)
        rowRange.Font.Color = DEV_COLOR_TEXT
        rowRange.Font.Bold = False

        rowRange.Cells(1, DEV_CONFIG_MARKER_COL).Font.Color = DEV_COLOR_TEXT

        If Len(Trim$(CStr(rowRange.Cells(1, DEV_CONFIG_KEY_COL).Value))) > 0 Then
            rowRange.Cells(1, DEV_CONFIG_KEY_COL).Font.Bold = True
            rowRange.Cells(1, DEV_CONFIG_KEY_COL).Font.Color = RGB(245, 245, 245)
        Else
            rowRange.Cells(1, DEV_CONFIG_VALUE_COL).Value = vbNullString
            rowRange.Cells(1, DEV_CONFIG_NOTE_COL).Value = vbNullString
        End If
NextRow:
    Next i
End Sub

Private Function mp_IsMarkerKey(ByVal keyText As String) As Boolean
    keyText = Trim$(keyText)
    If Len(keyText) < Len(DEV_MARKER_PREFIX) Then Exit Function
    mp_IsMarkerKey = (StrComp(Left$(keyText, Len(DEV_MARKER_PREFIX)), DEV_MARKER_PREFIX, vbTextCompare) = 0)
End Function

Private Sub mp_NormalizeLegacyMarkerEntry(ByRef entries As Variant, ByVal rowIndex As Long)
    Dim markerText As String
    Dim keyText As String
    Dim valueText As String

    markerText = Trim$(CStr(entries(rowIndex, DEV_CONFIG_MARKER_COL)))
    keyText = Trim$(CStr(entries(rowIndex, DEV_CONFIG_KEY_COL)))
    valueText = CStr(entries(rowIndex, DEV_CONFIG_VALUE_COL))

    If StrComp(markerText, DEV_MARKER_SYMBOL, vbTextCompare) = 0 Then
        Exit Sub
    End If

    If StrComp(Left$(keyText, Len(DEV_MARKER_SECTION)), DEV_MARKER_SECTION, vbTextCompare) = 0 Then
        entries(rowIndex, DEV_CONFIG_MARKER_COL) = DEV_MARKER_SYMBOL
        entries(rowIndex, DEV_CONFIG_KEY_COL) = valueText
        entries(rowIndex, DEV_CONFIG_VALUE_COL) = vbNullString
        entries(rowIndex, DEV_CONFIG_NOTE_COL) = vbNullString
        Exit Sub
    End If

    If StrComp(Left$(keyText, Len(DEV_MARKER_SPACER)), DEV_MARKER_SPACER, vbTextCompare) = 0 Then
        entries(rowIndex, DEV_CONFIG_MARKER_COL) = DEV_MARKER_SYMBOL
        entries(rowIndex, DEV_CONFIG_KEY_COL) = vbNullString
        entries(rowIndex, DEV_CONFIG_VALUE_COL) = vbNullString
        entries(rowIndex, DEV_CONFIG_NOTE_COL) = vbNullString
    End If
End Sub

' Приводит таблицу к схеме из 4 колонок (`.. | Key | Value | Note`).
' Если найдена старая 2-колоночная таблица, выполняется миграция данных.
Private Sub mp_EnsureConfigTableLayout(ByVal ws As Worksheet, ByVal tbl As ListObject)
    Dim rowCount As Long
    Dim i As Long
    Dim oldData As Variant
    Dim migrated() As Variant
    Dim noteColIndex As Long

    rowCount = mp_GetTableDataRowCount(tbl)
    If tbl.ListColumns.Count = DEV_CONFIG_COL_COUNT Then
        tbl.HeaderRowRange.Cells(1, DEV_CONFIG_MARKER_COL).Value = DEV_MARKER_HEADER
        tbl.HeaderRowRange.Cells(1, DEV_CONFIG_KEY_COL).Value = "Key"
        tbl.HeaderRowRange.Cells(1, DEV_CONFIG_NOTE_COL).Value = "Note"
        On Error Resume Next
        ex_Config.m_RefreshConfigTitle ws
        On Error GoTo 0
        Exit Sub
    End If

    If tbl.ListColumns.Count = 2 Then
        If rowCount > 0 Then
            oldData = tbl.DataBodyRange.Cells(1, 1).Resize(rowCount, 2).Value
            ReDim migrated(1 To rowCount, 1 To DEV_CONFIG_COL_COUNT)

            noteColIndex = tbl.Range.Column + 2
            For i = 1 To rowCount
                migrated(i, DEV_CONFIG_MARKER_COL) = vbNullString
                migrated(i, DEV_CONFIG_KEY_COL) = CStr(oldData(i, 1))
                migrated(i, DEV_CONFIG_VALUE_COL) = CStr(oldData(i, 2))
                migrated(i, DEV_CONFIG_NOTE_COL) = CStr(ws.Cells(tbl.HeaderRowRange.Row + i, noteColIndex).Value)
                mp_NormalizeLegacyMarkerEntry migrated, i
            Next i
        End If

        tbl.Resize ws.Range( _
            ws.Cells(tbl.HeaderRowRange.Row, tbl.Range.Column), _
            ws.Cells(tbl.HeaderRowRange.Row + rowCount, tbl.Range.Column + DEV_CONFIG_COL_COUNT - 1) _
        )

        tbl.HeaderRowRange.Cells(1, DEV_CONFIG_MARKER_COL).Value = DEV_MARKER_HEADER
        tbl.HeaderRowRange.Cells(1, DEV_CONFIG_KEY_COL).Value = "Key"
        tbl.HeaderRowRange.Cells(1, DEV_CONFIG_NOTE_COL).Value = "Note"
        On Error Resume Next
        ex_Config.m_RefreshConfigTitle ws
        On Error GoTo 0

        If rowCount > 0 Then
            tbl.DataBodyRange.Cells(1, 1).Resize(rowCount, DEV_CONFIG_COL_COUNT).Value = migrated
        End If
        Exit Sub
    End If

    MsgBox "Unsupported config table layout in '" & DEV_CONFIG_TABLE_NAME & "' (columns: " & CStr(tbl.ListColumns.Count) & ").", vbExclamation
End Sub

Private Sub mp_ApplyConfigTableDarkTheme(ByVal tbl As ListObject)
    Dim targetRange As Range
    Dim bodyRange As Range

    On Error Resume Next
    tbl.TableStyle = vbNullString
    tbl.ShowTableStyleColumnStripes = False
    tbl.ShowTableStyleRowStripes = False
    tbl.ShowTableStyleFirstColumn = False
    tbl.ShowTableStyleLastColumn = False
    On Error GoTo 0

    Set targetRange = tbl.Range
    With targetRange
        .Interior.Pattern = xlSolid
        .Interior.Color = DEV_COLOR_BG
        .Font.Color = DEV_COLOR_TEXT
        .Font.Bold = False
        .Borders.LineStyle = xlContinuous
        .Borders.Color = DEV_COLOR_BORDER
        .Borders.Weight = xlThin
    End With

    Set bodyRange = tbl.DataBodyRange
    If Not bodyRange Is Nothing Then
        bodyRange.Font.Bold = False
        bodyRange.Columns(DEV_CONFIG_NOTE_COL).Font.Color = DEV_COLOR_NOTE_TEXT
    End If

    tbl.HeaderRowRange.Font.Bold = True
    tbl.HeaderRowRange.Cells(1, DEV_CONFIG_NOTE_COL).Font.Color = DEV_COLOR_TEXT
    tbl.Range.EntireColumn.AutoFit
    If tbl.ListColumns(DEV_CONFIG_MARKER_COL).Range.ColumnWidth < 4 Then
        tbl.ListColumns(DEV_CONFIG_MARKER_COL).Range.ColumnWidth = 4
    End If
End Sub

Private Function mp_GetProfilesFilePath(Optional ByVal modeName As String = vbNullString) As String
    Dim resolvedMode As String

    resolvedMode = Trim$(modeName)
    If Len(resolvedMode) = 0 Then
        resolvedMode = MODE_PERSONAL_CARD
        On Error Resume Next
        resolvedMode = mp_GetSelectedModeName(ws_Dev)
        On Error GoTo 0
    End If

    mp_GetProfilesFilePath = ex_UILoader.m_GetProfilesFilePathByMode(resolvedMode, ThisWorkbook, "profilesByMode")
End Function

Private Function mp_GetProfileNode(ByVal doc As Object, ByVal profileName As String, ByVal createIfMissing As Boolean) As Object
    Dim node As Object
    Dim root As Object

    Set node = doc.selectSingleNode("/p:presets/p:profile[@name=" & mp_XPathLiteral(profileName) & "]")
    If node Is Nothing And createIfMissing Then
        Set root = doc.selectSingleNode("/p:presets")
        If root Is Nothing Then Exit Function
        Set node = doc.createNode(1, "profile", PRESETS_NS)
        node.setAttribute "name", profileName
        root.appendChild node
    End If

    Set mp_GetProfileNode = node
End Function

Private Function mp_XPathLiteral(ByVal value As String) As String
    If InStr(value, "'") = 0 Then
        mp_XPathLiteral = "'" & value & "'"
        Exit Function
    End If

    If InStr(value, """") = 0 Then
        mp_XPathLiteral = """" & value & """"
        Exit Function
    End If

    Dim parts() As String
    Dim i As Long

    parts = Split(value, "'")
    mp_XPathLiteral = "concat("
    For i = 0 To UBound(parts)
        If i > 0 Then
            mp_XPathLiteral = mp_XPathLiteral & ", ""'"" , "
        End If
        mp_XPathLiteral = mp_XPathLiteral & "'" & parts(i) & "'"
    Next i
    mp_XPathLiteral = mp_XPathLiteral & ")"
End Function
