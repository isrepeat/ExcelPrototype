Attribute VB_Name = "ex_Config"
Option Explicit

' =============================================================================
' ex_Config
' =============================================================================
' Конфигурация хранится на листе Dev в виде пары Key / Value.
' =============================================================================

Private Const DEV_SHEET_NAME As String = "Dev"

Private Const CONFIG_TOP As Long = 1
Private Const CONFIG_LEFT As Long = 1
Private Const CONFIG_ROWS As Long = 8
Private Const CONFIG_COLS As Long = 2

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

Public Function m_GetConfigValue( _
    ByVal keyName As String, _
    Optional ByVal defaultValue As String = vbNullString _
) As String

    Dim wsDev As Worksheet
    Dim lastRow As Long
    Dim searchRange As Range
    Dim foundCell As Range
    Dim valueText As String

    Set wsDev = mp_EnsureDevSheet()

    ' Гарантируем базовую область "Config" (шапка Key/Value)
    mp_EnsureConfigArea wsDev

    ' Ищем по всей колонке Key (A) до последней заполненной строки,
    ' чтобы поддержать большие конфиги (Layout.*, Map.*, Label.* и т.д.)
    lastRow = wsDev.Cells(wsDev.Rows.Count, CONFIG_LEFT).End(xlUp).Row
    If lastRow < 1 Then
        m_GetConfigValue = defaultValue
        Exit Function
    End If

    Set searchRange = wsDev.Range( _
        wsDev.Cells(1, CONFIG_LEFT), _
        wsDev.Cells(lastRow, CONFIG_LEFT + 1) _
    )

    Set foundCell = searchRange.Columns(1).Find( _
        What:=keyName, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        MatchCase:=False _
    )

    If foundCell Is Nothing Then
        m_GetConfigValue = defaultValue
        Exit Function
    End If

    valueText = CStr(wsDev.Cells(foundCell.Row, foundCell.Column + 1).Value)
    If Len(valueText) = 0 Then
        m_GetConfigValue = defaultValue
    Else
        m_GetConfigValue = valueText
    End If

End Function

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

    If Trim$(CStr(wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT).Value)) = "Key" And _
       Trim$(CStr(wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT + 1).Value)) = "Value" Then
        Exit Sub
    End If

    mp_RenderConfigArea wsDev

End Sub

Private Sub mp_RenderConfigArea(ByVal wsDev As Worksheet)

    Dim rng As Range
    Dim titleRange As Range
    Dim headerRange As Range

    Set rng = mp_GetConfigRange(wsDev)
    rng.Clear

    ' Title
    Set titleRange = wsDev.Range(wsDev.Cells(CONFIG_TOP, CONFIG_LEFT), _
                                 wsDev.Cells(CONFIG_TOP, CONFIG_LEFT + 1))
    titleRange.Merge
    titleRange.Value = "Config"
    titleRange.Font.Bold = True
    titleRange.HorizontalAlignment = xlCenter

    ' Header
    Set headerRange = wsDev.Range(wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT), _
                                  wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT + 1))
    wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT).Value = "Key"
    wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT + 1).Value = "Value"
    headerRange.Font.Bold = True

    ' Keys (стартовые)
    wsDev.Cells(CONFIG_TOP + 2, CONFIG_LEFT).Value = "StateFilePath"
    wsDev.Cells(CONFIG_TOP + 3, CONFIG_LEFT).Value = "StateTableName"
    wsDev.Cells(CONFIG_TOP + 4, CONFIG_LEFT).Value = "EventsFilePath"
    wsDev.Cells(CONFIG_TOP + 5, CONFIG_LEFT).Value = "EventsTableName"
    wsDev.Cells(CONFIG_TOP + 6, CONFIG_LEFT).Value = "KeyColumnName"
    wsDev.Cells(CONFIG_TOP + 6, CONFIG_LEFT + 1).Value = "Id"
    wsDev.Cells(CONFIG_TOP + 7, CONFIG_LEFT).Value = "PersonFIO"

    rng.Columns(1).ColumnWidth = 18
    rng.Columns(2).ColumnWidth = 50

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
