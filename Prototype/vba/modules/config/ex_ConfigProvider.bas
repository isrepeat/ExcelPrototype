Attribute VB_Name = "ex_ConfigProvider"
Option Explicit

' =============================================================================
' ex_ConfigProvider
' =============================================================================
' Назначение модуля:
' - поддерживать рабочую таблицу конфигурации на листе Dev (`tblDevConfig`);
' - отдавать значения конфига другим модулям через `m_GetConfigValue`;
' - обновлять служебный заголовок в третьей колонке таблицы
'   (`Config [profile = ...]`) на основе текущего профиля из `ddProfile`.
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
Private Const DEV_COL_NOTE As Long = 4
Private Const CONFIG_TITLE_TEMPLATE As String = "Config [profile = <CURRENT_PROFILE>]"

Private Const CONFIG_TOP As Long = 1
Private Const CONFIG_LEFT As Long = 1
Private Const CONFIG_ROWS As Long = 8
Private Const CONFIG_COLS As Long = 4

' Цвета (согласованы с темой результатов)
Private Const COLOR_BG As Long = &H1E1E1E
Private Const COLOR_TEXT As Long = &HEBEBEB
Private Const COLOR_BORDER As Long = &H505050

' =============================================================================
' Public API
' =============================================================================

Public Sub m_OpenConfigOnDev()
    Dim wsDev As Worksheet

    Set wsDev = mp_EnsureDevSheet()
    mp_EnsureConfigArea wsDev

    wsDev.Activate
    wsDev.Cells(CONFIG_TOP + 2, CONFIG_LEFT + 1).Select
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

' Обновляет служебный текст в заголовке 3-й колонки таблицы.
' Текст формируется по шаблону CONFIG_TITLE_TEMPLATE и подставляет:
' - явный `profileName`, если передан;
' - иначе активный элемент из dropdown `ddProfile`;
' - иначе `<none>`.
Public Sub m_RefreshConfigTitle( _
    Optional ByVal wsDev As Worksheet, _
    Optional ByVal profileName As String = vbNullString _
)
    Dim titleCell As Range
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
        resolvedProfile = mp_GetProfileNameFromDropdown(wsDev)
    End If
    If Len(resolvedProfile) = 0 Then
        resolvedProfile = "<none>"
    End If

    Set titleCell = cfgTable.HeaderRowRange.Cells(1, DEV_COL_VALUE)
    titleCell.Value = Replace(CONFIG_TITLE_TEMPLATE, "<CURRENT_PROFILE>", resolvedProfile)
    titleCell.Font.Bold = True
    titleCell.HorizontalAlignment = xlCenter
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

    Set cfgTable = mp_GetConfigTable(wsDev, False)
    If Not cfgTable Is Nothing Then
        m_RefreshConfigTitle wsDev
        Exit Sub
    End If
	
    If Trim$(CStr(wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT).Value)) = "Key" And _
       Trim$(CStr(wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT + 1).Value)) = "Value" Then
        mp_EnsureConfigTable wsDev
        Exit Sub
    End If

    If Trim$(CStr(wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT + DEV_COL_MARKER - 1).Value)) = DEV_MARKER_HEADER And _
       Trim$(CStr(wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT + DEV_COL_KEY - 1).Value)) = "Key" And _
       mp_IsConfigTitleHeader(CStr(wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT + DEV_COL_VALUE - 1).Value)) And _
       Trim$(CStr(wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT + DEV_COL_NOTE - 1).Value)) = "Note" Then
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

Private Function mp_GetProfileNameFromDropdown(ByVal wsDev As Worksheet) As String
    Dim shp As Shape
    Dim cf As Object
    Dim idx As Long

    Set shp = ex_ConfigProfilesManager.m_GetShapeByName(wsDev, "ddProfile")
    If shp Is Nothing Then Exit Function

    On Error Resume Next
    Set cf = shp.ControlFormat
    On Error GoTo 0
    If cf Is Nothing Then Exit Function

    On Error Resume Next
    idx = CLng(cf.Value)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If idx < 1 Then Exit Function
    On Error Resume Next
    mp_GetProfileNameFromDropdown = CStr(cf.List(idx))
    On Error GoTo 0
End Function

Private Sub mp_EnsureConfigTable(ByVal wsDev As Worksheet)
    Dim cfgTable As ListObject
    Dim lastRow As Long
    Dim tableRange As Range
    Dim isLegacyHeaders As Boolean

    Set cfgTable = mp_GetConfigTable(wsDev, False)
    If Not cfgTable Is Nothing Then Exit Sub

    isLegacyHeaders = (Trim$(CStr(wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT).Value)) = "Key" And _
                       Trim$(CStr(wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT + 1).Value)) = "Value")

    lastRow = wsDev.Cells(wsDev.Rows.Count, CONFIG_LEFT).End(xlUp).Row
    If wsDev.Cells(wsDev.Rows.Count, CONFIG_LEFT + 1).End(xlUp).Row > lastRow Then
        lastRow = wsDev.Cells(wsDev.Rows.Count, CONFIG_LEFT + 1).End(xlUp).Row
    End If
    If Not isLegacyHeaders And wsDev.Cells(wsDev.Rows.Count, CONFIG_LEFT + DEV_COL_NOTE - 1).End(xlUp).Row > lastRow Then
        lastRow = wsDev.Cells(wsDev.Rows.Count, CONFIG_LEFT + DEV_COL_NOTE - 1).End(xlUp).Row
    End If
    If lastRow < CONFIG_TOP + 1 Then
        lastRow = CONFIG_TOP + 1
    End If

    If isLegacyHeaders Then
        Set tableRange = wsDev.Range( _
            wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT), _
            wsDev.Cells(lastRow, CONFIG_LEFT + 1) _
        )
    Else
        Set tableRange = wsDev.Range( _
            wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT), _
            wsDev.Cells(lastRow, CONFIG_LEFT + DEV_COL_NOTE - 1) _
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

Private Sub mp_RenderConfigArea(ByVal wsDev As Worksheet)
    Dim rng As Range
    Dim headerRange As Range

    Set rng = mp_GetConfigRange(wsDev)
    rng.Clear

    ' Header
    Set headerRange = wsDev.Range(wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT), _
                                  wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT + DEV_COL_NOTE - 1))
    wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT + DEV_COL_MARKER - 1).Value = DEV_MARKER_HEADER
    wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT + DEV_COL_KEY - 1).Value = "Key"
    wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT + DEV_COL_NOTE - 1).Value = "Note"
    m_RefreshConfigTitle wsDev
    headerRange.Font.Bold = True

    ' Keys (стартовые)
    wsDev.Cells(CONFIG_TOP + 2, CONFIG_LEFT + DEV_COL_KEY - 1).Value = "StateFilePath"
    wsDev.Cells(CONFIG_TOP + 3, CONFIG_LEFT + DEV_COL_KEY - 1).Value = "StateTableName"
    wsDev.Cells(CONFIG_TOP + 4, CONFIG_LEFT + DEV_COL_KEY - 1).Value = "EventsFilePath"
    wsDev.Cells(CONFIG_TOP + 5, CONFIG_LEFT + DEV_COL_KEY - 1).Value = "EventsTableName"
    wsDev.Cells(CONFIG_TOP + 6, CONFIG_LEFT + DEV_COL_KEY - 1).Value = "KeyColumnName"
    wsDev.Cells(CONFIG_TOP + 6, CONFIG_LEFT + DEV_COL_VALUE - 1).Value = "Id"
    wsDev.Cells(CONFIG_TOP + 7, CONFIG_LEFT + DEV_COL_KEY - 1).Value = "PersonFIO"

    rng.EntireColumn.AutoFit
    If rng.Columns(DEV_COL_MARKER).ColumnWidth < 4 Then
        rng.Columns(DEV_COL_MARKER).ColumnWidth = 4
    End If

    mp_ApplyDarkThemeToRange rng
End Sub

Private Function mp_GetConfigRange(ByVal wsDev As Worksheet) As Range
    Set mp_GetConfigRange = wsDev.Range( _
        wsDev.Cells(CONFIG_TOP, CONFIG_LEFT), _
        wsDev.Cells(CONFIG_TOP + CONFIG_ROWS - 1, CONFIG_LEFT + CONFIG_COLS - 1) _
    )
End Function

Private Sub mp_ApplyDarkThemeToRange(ByVal target As Range)
    With target
        .Interior.Pattern = xlSolid
        .Interior.Color = COLOR_BG

        .Font.Color = COLOR_TEXT

        .Borders.LineStyle = xlContinuous
        .Borders.Color = COLOR_BORDER
        .Borders.Weight = xlThin
    End With
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
