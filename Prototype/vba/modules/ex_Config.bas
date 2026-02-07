Attribute VB_Name = "ex_Config"
Option Explicit

' =============================================================================
' ex_Config
' =============================================================================
' Конфигурация хранится на листе Dev в виде пары Key / Value.
'
' Актуальные ключи:
'   OldFilePath
'   OldTableName
'   NewFilePath
'   NewTableName
'   KeyColumnName
'
' Конфиг является UI + API-обёрткой над чтением значений из Dev-листа.
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
    EnsureConfigArea wsDev

    wsDev.Activate
    wsDev.Cells(CONFIG_TOP + 2, CONFIG_LEFT + 1).Select
End Sub

'' Helper wrappers removed. Use m_GetConfigValue and mp_NormalizePath directly.

Public Function m_GetConfigValue( _
    ByVal keyName As String, _
    Optional ByVal defaultValue As String = vbNullString _
) As String
    Dim wsDev As Worksheet
    Dim cfgRange As Range
    Dim foundCell As Range
    Dim valueText As String

        Set wsDev = mp_EnsureDevSheet()
    mp_EnsureConfigArea wsDev
        Set cfgRange = mp_GetConfigRange(wsDev)

    Set foundCell = cfgRange.Columns(1).Find( _
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

    ' Keys (АКТУАЛЬНЫЕ)
    wsDev.Cells(CONFIG_TOP + 2, CONFIG_LEFT).Value = "OldFilePath"
    wsDev.Cells(CONFIG_TOP + 3, CONFIG_LEFT).Value = "OldTableName"
    wsDev.Cells(CONFIG_TOP + 4, CONFIG_LEFT).Value = "NewFilePath"
    wsDev.Cells(CONFIG_TOP + 5, CONFIG_LEFT).Value = "NewTableName"
    wsDev.Cells(CONFIG_TOP + 6, CONFIG_LEFT).Value = "KeyColumnName"
    wsDev.Cells(CONFIG_TOP + 6, CONFIG_LEFT + 1).Value = "Id"

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
