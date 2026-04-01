Attribute VB_Name = "ex_ConfigProvider"
Option Explicit

' =============================================================================
' ex_ConfigProvider
' =============================================================================
' Назначение модуля:
' - поддерживать рабочую таблицу конфигурации на листе Dev (`tblDevConfig`);
' - отдавать значения конфига другим модулям через `m_GetConfigValue`;
' - обновлять служебный заголовок в третьей колонке таблицы
'   (`Config [profile = ...]`) на основе текущего активного профиля (state).
'
' Важно:
' - модуль НЕ хранит профили; он работает только с текущим отображением таблицы;
' - профильный XML читает/пишет соседний модуль ex_ConfigProfilesManager.
' =============================================================================

Private Const DEV_SHEET_NAME As String = "Dev"
Private Const DEV_CONFIG_TABLE_NAME As String = "tblDevConfig"
Private Const DEV_MARKER_SYMBOL As String = "#"
Private Const DEV_MARKER_HEADER As String = ".."
Private Const DEV_COL_MARKER As Long = 1
Private Const DEV_COL_KEY As Long = 2
Private Const DEV_COL_VALUE As Long = 3
Private Const DEV_COL_STYLES As Long = 4
Private Const DEV_HEADER_STYLES As String = "Styles"
Private Const CONFIG_TITLE_TEMPLATE As String = "Config [profile = <CURRENT_PROFILE>]"
Private Const FETCH_ADDITIONAL_DATA_NODE As String = "fetchAdditionalData"
Private Const FETCH_ADDITIONAL_DATA_ATTR_TABLE_REF As String = "tableRef"

Private Const CONFIG_TOP As Long = 1
Private Const CONFIG_LEFT As Long = 1
Private Const CONFIG_COLS As Long = 4

' =============================================================================
' Public API
' =============================================================================

Public Sub m_OpenConfigOnDev()
    Dim wsDev As Worksheet
    Dim tbl As ListObject

    Set wsDev = mp_EnsureDevSheet()
    ex_OutputFormattingPipeline.m_ApplySheetPipeline wsDev
    mp_EnsureConfigArea wsDev
    Set tbl = ex_ConfigTableStore.m_GetConfigTable(wsDev, False)
    If Not tbl Is Nothing Then
        ex_ConfigTableStore.m_ApplyConfigMarkerStyles tbl
    End If

    wsDev.Activate
    wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT + 1).Select
End Sub

' Читает значение конфигурации по ключу из таблицы `tblDevConfig`.
' Поиск идёт только по обычным строкам (marker-строки с символом `#` пропускаются).
' Если ключ не найден или значение пустое, возвращается `defaultValue`.
Public Function m_GetConfigValue( _
    ByVal keyName As String, _
    Optional ByVal defaultValue As String = vbNullString _
) As String

    Dim wsDev As Worksheet
    Dim cfgTable As ListObject
    Dim dataRange As Range
    Dim r As Long
    Dim markerText As String
    Dim keyText As String
    Dim valueText As String
    Dim keyCol As Long
    Dim valueCol As Long
    Dim markerCol As Long

    Set wsDev = mp_EnsureDevSheet()

    ' Гарантируем базовую область "Config" и таблицу Key/Value.
    mp_EnsureConfigArea wsDev

    Set cfgTable = mp_GetConfigTable(wsDev, True)
    If cfgTable Is Nothing Then
        m_GetConfigValue = defaultValue
        Exit Function
    End If

    If cfgTable.DataBodyRange Is Nothing Then
        m_GetConfigValue = defaultValue
        Exit Function
    End If

    Set dataRange = cfgTable.DataBodyRange
    keyCol = DEV_COL_KEY
    valueCol = DEV_COL_VALUE
    markerCol = DEV_COL_MARKER
    If cfgTable.ListColumns.Count < DEV_COL_VALUE Then
        keyCol = 1
        valueCol = 2
        markerCol = 0
    End If

    For r = 1 To dataRange.Rows.Count
        markerText = vbNullString
        If markerCol > 0 Then
            markerText = Trim$(CStr(dataRange.Cells(r, markerCol).Value))
        End If
        keyText = Trim$(CStr(dataRange.Cells(r, keyCol).Value))

        If StrComp(markerText, DEV_MARKER_SYMBOL, vbTextCompare) <> 0 Then
            If StrComp(keyText, keyName, vbTextCompare) = 0 Then
                valueText = CStr(dataRange.Cells(r, valueCol).Value)
                If Len(valueText) = 0 Then
                    m_GetConfigValue = defaultValue
                Else
                    m_GetConfigValue = valueText
                End If
                Exit Function
            End If
        End If
    Next r

    m_GetConfigValue = defaultValue
End Function

' Возвращает значение атрибута `type` (первая колонка таблицы Dev config)
' для строки с указанным ключом.
' Пример: если у `CommonKey` в колонке marker стоит `rx`,
' функция вернет `rx`.
Public Function m_GetConfigEntryType( _
    ByVal keyName As String, _
    Optional ByVal defaultValue As String = vbNullString _
) As String
    Dim wsDev As Worksheet
    Dim cfgTable As ListObject
    Dim dataRange As Range
    Dim r As Long
    Dim markerText As String
    Dim keyText As String
    Dim keyCol As Long
    Dim markerCol As Long

    keyName = Trim$(CStr(keyName))
    If Len(keyName) = 0 Then
        m_GetConfigEntryType = defaultValue
        Exit Function
    End If

    Set wsDev = mp_EnsureDevSheet()
    mp_EnsureConfigArea wsDev

    Set cfgTable = mp_GetConfigTable(wsDev, True)
    If cfgTable Is Nothing Then
        m_GetConfigEntryType = defaultValue
        Exit Function
    End If
    If cfgTable.DataBodyRange Is Nothing Then
        m_GetConfigEntryType = defaultValue
        Exit Function
    End If

    Set dataRange = cfgTable.DataBodyRange
    keyCol = DEV_COL_KEY
    markerCol = DEV_COL_MARKER
    If cfgTable.ListColumns.Count < DEV_COL_VALUE Then
        keyCol = 1
        markerCol = 0
    End If

    For r = 1 To dataRange.Rows.Count
        markerText = vbNullString
        If markerCol > 0 Then
            markerText = Trim$(CStr(dataRange.Cells(r, markerCol).Value))
        End If
        keyText = Trim$(CStr(dataRange.Cells(r, keyCol).Value))

        If StrComp(markerText, DEV_MARKER_SYMBOL, vbTextCompare) <> 0 Then
            If StrComp(keyText, keyName, vbTextCompare) = 0 Then
                If Len(markerText) = 0 Then
                    m_GetConfigEntryType = defaultValue
                Else
                    m_GetConfigEntryType = markerText
                End If
                Exit Function
            End If
        End If
    Next r

    m_GetConfigEntryType = defaultValue
End Function

Public Function m_LoadConfigDictionary( _
    Optional ByVal errSource As String = "ex_ConfigProvider", _
    Optional ByVal errNoTableCode As Long = 1330, _
    Optional ByVal errNoRowsCode As Long = 1331 _
) As Object
    Dim wsDev As Worksheet
    Dim cfgTable As ListObject
    Dim dict As Object
    Dim dataRange As Range
    Dim r As Long
    Dim markerText As String
    Dim keyText As String

    Set wsDev = mp_EnsureDevSheet()

    On Error Resume Next
    Set cfgTable = wsDev.ListObjects(DEV_CONFIG_TABLE_NAME)
    On Error GoTo 0

    If cfgTable Is Nothing Then
        Err.Raise vbObjectError + errNoTableCode, CStr(errSource), _
            "Config table '" & DEV_CONFIG_TABLE_NAME & "' was not found on sheet '" & wsDev.Name & "'."
    End If
    If cfgTable.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + errNoRowsCode, CStr(errSource), _
            "Config table '" & DEV_CONFIG_TABLE_NAME & "' has no data rows."
    End If

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    Set dataRange = cfgTable.DataBodyRange

    For r = 1 To dataRange.Rows.Count
        markerText = Trim$(CStr(dataRange.Cells(r, DEV_COL_MARKER).Value))
        If StrComp(markerText, DEV_MARKER_SYMBOL, vbTextCompare) = 0 Then GoTo ContinueRow

        keyText = Trim$(CStr(dataRange.Cells(r, DEV_COL_KEY).Value))
        If Len(keyText) = 0 Then GoTo ContinueRow

        dict(keyText) = CStr(dataRange.Cells(r, DEV_COL_VALUE).Value)
ContinueRow:
    Next r

    mp_AppendFetchAdditionalDataFromActiveProfile dict, wsDev, CStr(errSource)

    Set m_LoadConfigDictionary = dict
End Function

Public Function m_GetResolvedSheetName( _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    Optional ByVal cfg As Object = Nothing, _
    Optional ByVal required As Boolean = True, _
    Optional ByVal errSource As String = "ex_ConfigProvider", _
    Optional ByVal errMissingKeyCode As Long = 1900, _
    Optional ByVal errEmptyValueCode As Long = 1901, _
    Optional ByVal errMissingResolverCode As Long = 1902, _
    Optional ByVal errEmptyResolvedValueCode As Long = 1903, _
    Optional ByVal errResolverFailedCode As Long = 1904 _
) As String
    Dim sheetKey As String
    Dim resolverKey As String
    Dim resolverArgsKey As String
    Dim rawSheetName As String
    Dim resolverName As String
    Dim resolverArgs As String
    Dim resolverCallName As String
    Dim resolvedValue As Variant
    Dim resolvedSheetName As String

    sourceAlias = Trim$(CStr(sourceAlias))
    tableAlias = Trim$(CStr(tableAlias))
    If Len(sourceAlias) = 0 Or Len(tableAlias) = 0 Then
        Err.Raise vbObjectError + errMissingKeyCode, CStr(errSource), "Source alias and table alias are required for SheetName resolver."
    End If

    sheetKey = sourceAlias & ".Sheet[" & tableAlias & "].SheetName"
    resolverKey = sourceAlias & ".Sheet[" & tableAlias & "].SheetNameResolver"
    resolverArgsKey = sourceAlias & ".Sheet[" & tableAlias & "].SheetNameResolverArgs"

    If required Then
        rawSheetName = mp_GetConfigValueRequired(cfg, sheetKey, CStr(errSource), errMissingKeyCode, errEmptyValueCode)
    Else
        rawSheetName = mp_GetConfigValueOptional(cfg, sheetKey, vbNullString)
        If Len(rawSheetName) = 0 Then Exit Function
    End If

    resolverName = mp_GetConfigValueOptional(cfg, resolverKey, vbNullString)
    resolverArgs = mp_GetConfigValueOptional(cfg, resolverArgsKey, vbNullString)

    If Len(resolverName) = 0 Then
        If mp_HasPlaceholderTokens(rawSheetName) Then
            Err.Raise vbObjectError + errMissingResolverCode, CStr(errSource), _
                "SheetName contains placeholders but no resolver is configured for key '" & sheetKey & "'. " & _
                "Set '" & resolverKey & "' (for example: ex_SourceResolvers.m_ResolveWithDateFormatter)."
        End If
        m_GetResolvedSheetName = rawSheetName
        Exit Function
    End If

    If InStr(1, resolverName, "!", vbBinaryCompare) > 0 Then
        resolverCallName = resolverName
    Else
        resolverCallName = "'" & ThisWorkbook.Name & "'!" & resolverName
    End If

    On Error GoTo ResolverEH
    resolvedValue = Application.Run(resolverCallName, rawSheetName, resolverArgs)
    On Error GoTo 0

    resolvedSheetName = Trim$(CStr(resolvedValue))
    If Len(resolvedSheetName) = 0 Then
        Err.Raise vbObjectError + errEmptyResolvedValueCode, CStr(errSource), _
            "SheetName resolver '" & resolverName & "' returned an empty value for key '" & sheetKey & "'."
    End If

    m_GetResolvedSheetName = resolvedSheetName
    Exit Function

ResolverEH:
    Err.Raise vbObjectError + errResolverFailedCode, CStr(errSource), _
        "SheetName resolver failed for key '" & sheetKey & "' (resolver='" & resolverName & "'): " & Err.Description
End Function

Public Function m_GetConfigStyle( _
    ByVal keyName As String, _
    Optional ByVal defaultValue As String = vbNullString _
) As String
    ' Inline styles from config table are deprecated.
    m_GetConfigStyle = defaultValue
End Function

Public Function m_GetConfigStylesDictionary() As Object
    Dim stylesDict As Object

    Set stylesDict = CreateObject("Scripting.Dictionary")
    stylesDict.CompareMode = 1
    ' Inline styles from config table are deprecated.
    Set m_GetConfigStylesDictionary = stylesDict
End Function

Public Sub m_SetConfigValue( _
    ByVal keyName As String, _
    ByVal valueText As String, _
    Optional ByVal createIfMissing As Boolean = True _
)
    Dim wsDev As Worksheet
    Dim cfgTable As ListObject
    Dim dataRange As Range
    Dim keyCol As Long
    Dim valueCol As Long
    Dim markerCol As Long
    Dim r As Long
    Dim markerText As String
    Dim keyText As String
    Dim newRow As ListRow

    keyName = Trim$(keyName)
    If Len(keyName) = 0 Then
        MsgBox "Config key name must not be empty.", vbExclamation
        Exit Sub
    End If

    Set wsDev = mp_EnsureDevSheet()
    mp_EnsureConfigArea wsDev

    Set cfgTable = mp_GetConfigTable(wsDev, True)
    If cfgTable Is Nothing Then
        MsgBox "Config table '" & DEV_CONFIG_TABLE_NAME & "' was not found on sheet '" & wsDev.Name & "'.", vbExclamation
        Exit Sub
    End If
    ex_ConfigTableStore.m_EnsureConfigTableTextFormat cfgTable

    keyCol = DEV_COL_KEY
    valueCol = DEV_COL_VALUE
    markerCol = DEV_COL_MARKER
    If cfgTable.ListColumns.Count < DEV_COL_VALUE Then
        keyCol = 1
        valueCol = 2
        markerCol = 0
    End If

    If Not cfgTable.DataBodyRange Is Nothing Then
        Set dataRange = cfgTable.DataBodyRange
        For r = 1 To dataRange.Rows.Count
            markerText = vbNullString
            If markerCol > 0 Then
                markerText = Trim$(CStr(dataRange.Cells(r, markerCol).Value))
            End If
            keyText = Trim$(CStr(dataRange.Cells(r, keyCol).Value))
            If StrComp(markerText, DEV_MARKER_SYMBOL, vbTextCompare) <> 0 Then
                If StrComp(keyText, keyName, vbTextCompare) = 0 Then
                    dataRange.Cells(r, valueCol).NumberFormat = "@"
                    dataRange.Cells(r, valueCol).Value2 = CStr(valueText)
                    Exit Sub
                End If
            End If
        Next r
    End If

    If Not createIfMissing Then
        MsgBox "Config key '" & keyName & "' was not found in '" & DEV_CONFIG_TABLE_NAME & "'.", vbExclamation
        Exit Sub
    End If

    Set newRow = cfgTable.ListRows.Add
    If newRow Is Nothing Then
        MsgBox "Failed to append a new row into config table '" & DEV_CONFIG_TABLE_NAME & "'.", vbExclamation
        Exit Sub
    End If

    If markerCol > 0 Then
        newRow.Range.Cells(1, markerCol).NumberFormat = "@"
        newRow.Range.Cells(1, markerCol).Value2 = vbNullString
    End If
    newRow.Range.Cells(1, keyCol).NumberFormat = "@"
    newRow.Range.Cells(1, keyCol).Value2 = CStr(keyName)
    newRow.Range.Cells(1, valueCol).NumberFormat = "@"
    newRow.Range.Cells(1, valueCol).Value2 = CStr(valueText)
End Sub

Public Sub m_SetConfigStyle( _
    ByVal keyName As String, _
    ByVal styleText As String, _
    Optional ByVal createIfMissing As Boolean = False _
)
    ' Inline styles from config table are deprecated.
End Sub

' Обновляет служебный текст в заголовке 3-й колонки таблицы.
' Текст формируется по шаблону CONFIG_TITLE_TEMPLATE и подставляет:
' - явный `profileName`, если передан;
' - иначе активный профиль из state;
' - иначе `<none>`.
Public Sub m_RefreshConfigTitle( _
    Optional ByVal wsDev As Worksheet, _
    Optional ByVal profileName As String = vbNullString _
)
    Dim titleCell As Range
    Dim titleText As String
    Dim resolvedProfile As String
    Dim cfgTable As ListObject

    If wsDev Is Nothing Then
        Set wsDev = mp_EnsureDevSheet()
    End If

    Set cfgTable = mp_GetConfigTable(wsDev, True)
    If cfgTable Is Nothing Then Exit Sub
    If cfgTable.ListColumns.Count < DEV_COL_VALUE Then Exit Sub

    resolvedProfile = Trim$(profileName)
    If Len(resolvedProfile) = 0 Then
        resolvedProfile = Trim$(ex_ConfigProfilesManager.m_GetActiveProfileName(wsDev))
    End If
    If Len(resolvedProfile) = 0 Then
        resolvedProfile = "<none>"
    End If

    Set titleCell = cfgTable.HeaderRowRange.Cells(1, DEV_COL_VALUE)
    titleText = Replace(CONFIG_TITLE_TEMPLATE, "<CURRENT_PROFILE>", resolvedProfile)

    If StrComp(CStr(titleCell.Value2), titleText, vbBinaryCompare) = 0 Then Exit Sub
    On Error Resume Next
    titleCell.Value2 = titleText
    On Error GoTo 0
End Sub

' =============================================================================
' Internal
' =============================================================================

Private Function mp_EnsureDevSheet() As Worksheet
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, DEV_SHEET_NAME, vbTextCompare) = 0 Then
            Set mp_EnsureDevSheet = ws
            Exit Function
        End If
    Next ws

    Err.Raise vbObjectError + 1000, "ex_Config", _
        "Лист '" & DEV_SHEET_NAME & "' не найден."
End Function

Private Sub mp_EnsureConfigArea(ByVal wsDev As Worksheet)
    Dim cfgTable As ListObject

    Set cfgTable = ex_ConfigTableStore.m_GetConfigTable(wsDev, False)
    If Not cfgTable Is Nothing Then
        m_RefreshConfigTitle wsDev
        Exit Sub
    End If
	
    If Trim$(CStr(wsDev.Cells(CONFIG_TOP, CONFIG_LEFT).Value)) = "Key" And _
       Trim$(CStr(wsDev.Cells(CONFIG_TOP, CONFIG_LEFT + 1).Value)) = "Value" Then
        mp_EnsureConfigTable wsDev
        Exit Sub
    End If

    If Trim$(CStr(wsDev.Cells(CONFIG_TOP, CONFIG_LEFT + DEV_COL_MARKER - 1).Value)) = DEV_MARKER_HEADER And _
       Trim$(CStr(wsDev.Cells(CONFIG_TOP, CONFIG_LEFT + DEV_COL_KEY - 1).Value)) = "Key" And _
       mp_IsConfigTitleHeader(CStr(wsDev.Cells(CONFIG_TOP, CONFIG_LEFT + DEV_COL_VALUE - 1).Value)) And _
       Trim$(CStr(wsDev.Cells(CONFIG_TOP, CONFIG_LEFT + DEV_COL_STYLES - 1).Value)) = DEV_HEADER_STYLES Then
        mp_EnsureConfigTable wsDev
        Exit Sub
    End If

    mp_RenderConfigArea wsDev
    mp_EnsureConfigTable wsDev
End Sub

Private Function mp_IsConfigTitleHeader(ByVal cellText As String) As Boolean
    cellText = Trim$(cellText)
    If Len(cellText) = 0 Then Exit Function
    mp_IsConfigTitleHeader = (InStr(1, cellText, "Config [profile =", vbTextCompare) = 1)
End Function

Private Sub mp_EnsureConfigTable(ByVal wsDev As Worksheet)
    Dim cfgTable As ListObject
    Dim lastRow As Long
    Dim tableRange As Range
    Dim isLegacyHeaders As Boolean

    Set cfgTable = mp_GetConfigTable(wsDev, False)
    If Not cfgTable Is Nothing Then Exit Sub

    isLegacyHeaders = (Trim$(CStr(wsDev.Cells(CONFIG_TOP, CONFIG_LEFT).Value)) = "Key" And _
                       Trim$(CStr(wsDev.Cells(CONFIG_TOP, CONFIG_LEFT + 1).Value)) = "Value")

    lastRow = wsDev.Cells(wsDev.Rows.Count, CONFIG_LEFT).End(xlUp).Row
    If wsDev.Cells(wsDev.Rows.Count, CONFIG_LEFT + 1).End(xlUp).Row > lastRow Then
        lastRow = wsDev.Cells(wsDev.Rows.Count, CONFIG_LEFT + 1).End(xlUp).Row
    End If
    If Not isLegacyHeaders And wsDev.Cells(wsDev.Rows.Count, CONFIG_LEFT + DEV_COL_STYLES - 1).End(xlUp).Row > lastRow Then
        lastRow = wsDev.Cells(wsDev.Rows.Count, CONFIG_LEFT + DEV_COL_STYLES - 1).End(xlUp).Row
    End If
    If lastRow < CONFIG_TOP Then
        lastRow = CONFIG_TOP
    End If

    If isLegacyHeaders Then
        Set tableRange = wsDev.Range( _
            wsDev.Cells(CONFIG_TOP, CONFIG_LEFT), _
            wsDev.Cells(lastRow, CONFIG_LEFT + 1) _
        )
    Else
        Set tableRange = wsDev.Range( _
            wsDev.Cells(CONFIG_TOP, CONFIG_LEFT), _
            wsDev.Cells(lastRow, CONFIG_LEFT + DEV_COL_STYLES - 1) _
        )
    End If

    On Error Resume Next
    Set cfgTable = wsDev.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    On Error Resume Next
    cfgTable.Name = DEV_CONFIG_TABLE_NAME
    On Error GoTo 0
End Sub

Private Function mp_GetConfigTable(ByVal wsDev As Worksheet, Optional ByVal createIfMissing As Boolean = False) As ListObject
    Dim cfgTable As ListObject

    On Error Resume Next
    Set cfgTable = wsDev.ListObjects(DEV_CONFIG_TABLE_NAME)
    On Error GoTo 0

    If cfgTable Is Nothing And createIfMissing Then
        mp_EnsureConfigTable wsDev
        On Error Resume Next
        Set cfgTable = wsDev.ListObjects(DEV_CONFIG_TABLE_NAME)
        On Error GoTo 0
    End If

    Set mp_GetConfigTable = cfgTable
End Function

Private Function mp_GetConfigValueOptional( _
    ByVal cfg As Object, _
    ByVal keyName As String, _
    ByVal defaultValue As String _
) As String
    keyName = Trim$(CStr(keyName))
    If Len(keyName) = 0 Then
        mp_GetConfigValueOptional = CStr(defaultValue)
        Exit Function
    End If

    If Not cfg Is Nothing Then
        If cfg.Exists(keyName) Then
            mp_GetConfigValueOptional = Trim$(CStr(cfg(keyName)))
            Exit Function
        End If
    End If

    mp_GetConfigValueOptional = Trim$(CStr(m_GetConfigValue(keyName, defaultValue)))
End Function

Private Function mp_GetConfigValueRequired( _
    ByVal cfg As Object, _
    ByVal keyName As String, _
    ByVal errSource As String, _
    ByVal errMissingKeyCode As Long, _
    ByVal errEmptyValueCode As Long _
) As String
    keyName = Trim$(CStr(keyName))
    If Len(keyName) = 0 Then
        Err.Raise vbObjectError + errMissingKeyCode, CStr(errSource), "Config key name is empty."
    End If

    If Not cfg Is Nothing Then
        If Not cfg.Exists(keyName) Then
            Err.Raise vbObjectError + errMissingKeyCode, CStr(errSource), "Missing required config key '" & keyName & "'."
        End If

        mp_GetConfigValueRequired = Trim$(CStr(cfg(keyName)))
        If Len(mp_GetConfigValueRequired) = 0 Then
            Err.Raise vbObjectError + errEmptyValueCode, CStr(errSource), "Config key '" & keyName & "' is empty."
        End If
        Exit Function
    End If

    mp_GetConfigValueRequired = Trim$(CStr(m_GetConfigValue(keyName, vbNullString)))
    If Len(mp_GetConfigValueRequired) = 0 Then
        Err.Raise vbObjectError + errMissingKeyCode, CStr(errSource), "Missing required config key '" & keyName & "'."
    End If
End Function

Private Function mp_HasPlaceholderTokens(ByVal valueText As String) As Boolean
    Dim normalized As String

    normalized = Trim$(valueText)
    If Len(normalized) = 0 Then Exit Function

    mp_HasPlaceholderTokens = (InStr(1, normalized, "{", vbBinaryCompare) > 0) _
                              And (InStr(1, normalized, "}", vbBinaryCompare) > 0)
End Function

Private Sub mp_AppendFetchAdditionalDataFromActiveProfile( _
    ByVal cfg As Object, _
    Optional ByVal ws As Worksheet = Nothing, _
    Optional ByVal errSource As String = "ex_ConfigProvider" _
)
    Dim modeKey As String
    Dim profileName As String
    Dim filePath As String
    Dim doc As Object
    Dim profileNode As Object
    Dim fetchNodes As Object
    Dim fetchNode As Object
    Dim tableRef As String
    Dim dslText As String
    Dim cfgKey As String

    If cfg Is Nothing Then Exit Sub

    If ws Is Nothing Then
        Set ws = ws_Dev
    End If

    modeKey = Trim$(ex_ConfigProfilesManager.m_GetActiveModeKey(ws))
    profileName = Trim$(ex_ConfigProfilesManager.m_GetActiveProfileName(ws))
    If Len(modeKey) = 0 Or Len(profileName) = 0 Then Exit Sub

    filePath = Trim$(ex_ProfilesStore.m_GetProfilesFilePath(modeKey, ThisWorkbook))
    If Len(filePath) = 0 Then Exit Sub

    Set doc = ex_ProfilesStore.m_LoadProfilesDom(filePath)
    If doc Is Nothing Then Exit Sub

    Set profileNode = ex_ProfilesStore.m_GetProfileNode(doc, profileName, False)
    If profileNode Is Nothing Then Exit Sub

    Set fetchNodes = profileNode.selectNodes("p:" & FETCH_ADDITIONAL_DATA_NODE)
    If fetchNodes Is Nothing Then Exit Sub
    If fetchNodes.Length = 0 Then Exit Sub

    For Each fetchNode In fetchNodes
        tableRef = Trim$(mp_NodeAttrText(fetchNode, FETCH_ADDITIONAL_DATA_ATTR_TABLE_REF))
        If Len(tableRef) = 0 Then
            Err.Raise vbObjectError + 1837, CStr(errSource), _
                "Attribute '" & FETCH_ADDITIONAL_DATA_ATTR_TABLE_REF & "' is required for <" & FETCH_ADDITIONAL_DATA_NODE & "> in profile '" & profileName & "'."
        End If

        dslText = Trim$(CStr(fetchNode.Text))
        If Len(dslText) = 0 Then
            Err.Raise vbObjectError + 1838, CStr(errSource), _
                "<" & FETCH_ADDITIONAL_DATA_NODE & "> for tableRef '" & tableRef & "' is empty in profile '" & profileName & "'."
        End If

        cfgKey = tableRef & ".Fetch.Dsl"
        cfg(cfgKey) = dslText
    Next fetchNode
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

Private Sub mp_RenderConfigArea(ByVal wsDev As Worksheet)
    Dim clearRange As Range
    Dim lastRow As Long

    lastRow = wsDev.Cells(wsDev.Rows.Count, CONFIG_LEFT).End(xlUp).Row
    If wsDev.Cells(wsDev.Rows.Count, CONFIG_LEFT + 1).End(xlUp).Row > lastRow Then
        lastRow = wsDev.Cells(wsDev.Rows.Count, CONFIG_LEFT + 1).End(xlUp).Row
    End If
    If wsDev.Cells(wsDev.Rows.Count, CONFIG_LEFT + DEV_COL_STYLES - 1).End(xlUp).Row > lastRow Then
        lastRow = wsDev.Cells(wsDev.Rows.Count, CONFIG_LEFT + DEV_COL_STYLES - 1).End(xlUp).Row
    End If
    If lastRow < CONFIG_TOP + 6 Then
        lastRow = CONFIG_TOP + 6
    End If
    Set clearRange = wsDev.Range( _
        wsDev.Cells(CONFIG_TOP, CONFIG_LEFT), _
        wsDev.Cells(lastRow, CONFIG_LEFT + DEV_COL_STYLES - 1) _
    )
    clearRange.Clear

    ' Header
    wsDev.Cells(CONFIG_TOP, CONFIG_LEFT + DEV_COL_MARKER - 1).Value = DEV_MARKER_HEADER
    wsDev.Cells(CONFIG_TOP, CONFIG_LEFT + DEV_COL_KEY - 1).Value = "Key"
    wsDev.Cells(CONFIG_TOP, CONFIG_LEFT + DEV_COL_STYLES - 1).Value = DEV_HEADER_STYLES
    m_RefreshConfigTitle wsDev

    ' Keys (стартовые)
    wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT + DEV_COL_KEY - 1).Value = "StateFilePath"
    wsDev.Cells(CONFIG_TOP + 2, CONFIG_LEFT + DEV_COL_KEY - 1).Value = "StateTableName"
    wsDev.Cells(CONFIG_TOP + 3, CONFIG_LEFT + DEV_COL_KEY - 1).Value = "EventsFilePath"
    wsDev.Cells(CONFIG_TOP + 4, CONFIG_LEFT + DEV_COL_KEY - 1).Value = "EventsTableName"
    wsDev.Cells(CONFIG_TOP + 5, CONFIG_LEFT + DEV_COL_KEY - 1).Value = "KeyColumnName"
    wsDev.Cells(CONFIG_TOP + 5, CONFIG_LEFT + DEV_COL_VALUE - 1).Value = "Id"
    wsDev.Cells(CONFIG_TOP + 6, CONFIG_LEFT + DEV_COL_KEY - 1).Value = "PersonFIO"

    ex_ConfigTableStore.m_AutoFitConfigColumnsWithinStableZone wsDev, CONFIG_LEFT, CONFIG_COLS, DEV_COL_MARKER
End Sub

Private Function mp_NormalizePath(ByVal inputPath As String) As String
    Dim basePath As String

    inputPath = Trim$(inputPath)
    If Len(inputPath) = 0 Then
        mp_NormalizePath = vbNullString
        Exit Function
    End If

    ' Абсолютный или UNC
    If Left$(inputPath, 2) = "\\" Or InStr(1, inputPath, ":\", vbTextCompare) > 0 Then
        mp_NormalizePath = inputPath
        Exit Function
    End If

    basePath = ThisWorkbook.Path
    If Len(basePath) = 0 Then
        mp_NormalizePath = inputPath
        Exit Function
    End If

    If Right$(basePath, 1) <> "\" Then
        basePath = basePath & "\"
    End If

    mp_NormalizePath = basePath & inputPath
End Function
