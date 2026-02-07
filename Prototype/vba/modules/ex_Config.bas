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

Public Sub OpenConfigOnDev()
    Dim wsDev As Worksheet

    Set wsDev = EnsureDevSheet()
    EnsureConfigArea wsDev

    wsDev.Activate
    wsDev.Cells(CONFIG_TOP + 2, CONFIG_LEFT + 1).Select
End Sub

Public Function GetOldFilePath() As String
    GetOldFilePath = NormalizePath(GetConfigValue("OldFilePath"))
End Function

Public Function GetOldTableName() As String
    GetOldTableName = GetConfigValue("OldTableName")
End Function

Public Function GetNewFilePath() As String
    GetNewFilePath = NormalizePath(GetConfigValue("NewFilePath"))
End Function

Public Function GetNewTableName() As String
    GetNewTableName = GetConfigValue("NewTableName")
End Function

Public Function GetKeyColumnName() As String
    GetKeyColumnName = GetConfigValue("KeyColumnName", "Id")
End Function

Public Function GetConfigValue( _
    ByVal keyName As String, _
    Optional ByVal defaultValue As String = vbNullString _
) As String
    Dim wsDev As Worksheet
    Dim cfgRange As Range
    Dim foundCell As Range
    Dim valueText As String

    Set wsDev = EnsureDevSheet()
    EnsureConfigArea wsDev
    Set cfgRange = GetConfigRange(wsDev)

    Set foundCell = cfgRange.Columns(1).Find( _
        What:=keyName, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        MatchCase:=False _
    )

    If foundCell Is Nothing Then
        GetConfigValue = defaultValue
        Exit Function
    End If

    valueText = CStr(wsDev.Cells(foundCell.Row, foundCell.Column + 1).Value)
    If Len(valueText) = 0 Then
        GetConfigValue = defaultValue
    Else
        GetConfigValue = valueText
    End If
End Function

' =============================================================================
' Internal
' =============================================================================

Private Function EnsureDevSheet() As Worksheet
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, DEV_SHEET_NAME, vbTextCompare) = 0 Then
            Set EnsureDevSheet = ws
            Exit Function
        End If
    Next ws

    Err.Raise vbObjectError + 1000, "ex_Config", _
        "Лист '" & DEV_SHEET_NAME & "' не найден."
End Function

Private Sub EnsureConfigArea(ByVal wsDev As Worksheet)
    If Trim$(CStr(wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT).Value)) = "Key" And _
       Trim$(CStr(wsDev.Cells(CONFIG_TOP + 1, CONFIG_LEFT + 1).Value)) = "Value" Then
        Exit Sub
    End If

    RenderConfigArea wsDev
End Sub

Private Sub RenderConfigArea(ByVal wsDev As Worksheet)
    Dim rng As Range
    Dim titleRange As Range
    Dim headerRange As Range

    Set rng = GetConfigRange(wsDev)
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

    ApplyDarkThemeToRange rng
End Sub

Private Function GetConfigRange(ByVal wsDev As Worksheet) As Range
    Set GetConfigRange = wsDev.Range( _
        wsDev.Cells(CONFIG_TOP, CONFIG_LEFT), _
        wsDev.Cells(CONFIG_TOP + CONFIG_ROWS - 1, CONFIG_LEFT + CONFIG_COLS - 1) _
    )
End Function

Private Sub ApplyDarkThemeToRange(ByVal target As Range)
    With target
        .Interior.Pattern = xlSolid
        .Interior.Color = COLOR_BG

        .Font.Color = COLOR_TEXT

        .Borders.LineStyle = xlContinuous
        .Borders.Color = COLOR_BORDER
        .Borders.Weight = xlThin
    End With
End Sub

Private Function NormalizePath(ByVal inputPath As String) As String
    Dim basePath As String

    inputPath = Trim$(inputPath)
    If Len(inputPath) = 0 Then
        NormalizePath = vbNullString
        Exit Function
    End If

    ' Абсолютный или UNC
    If Left$(inputPath, 2) = "\\" Or InStr(1, inputPath, ":\", vbTextCompare) > 0 Then
        NormalizePath = inputPath
        Exit Function
    End If

    basePath = ThisWorkbook.Path
    If Len(basePath) = 0 Then
        NormalizePath = inputPath
        Exit Function
    End If

    If Right$(basePath, 1) <> "\" Then
        basePath = basePath & "\"
    End If

    NormalizePath = basePath & inputPath
End Function
