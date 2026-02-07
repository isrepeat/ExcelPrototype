Attribute VB_Name = "ex_SourceLoader"
Option Explicit

Public Sub LoadOldNewFromConfigToInternalSheets()
    Dim oldPath As String
    Dim oldTableName As String
    Dim newPath As String
    Dim newTableName As String

    oldPath = ResolvePath(GetConfigValueSafe("OldFilePath"))
    oldTableName = Trim$(GetConfigValueSafe("OldTableName"))

    newPath = ResolvePath(GetConfigValueSafe("NewFilePath"))
    newTableName = Trim$(GetConfigValueSafe("NewTableName"))

    If Len(oldPath) = 0 Or Len(oldTableName) = 0 Or Len(newPath) = 0 Or Len(newTableName) = 0 Then
        Err.Raise vbObjectError + 500, "ex_SourceLoader", _
            "Конфиг не заполнен. Нужны OldFilePath/OldTableName/NewFilePath/NewTableName."
    End If

    If Dir(oldPath) = vbNullString Then
        Err.Raise vbObjectError + 501, "ex_SourceLoader", "OldFilePath не найден: " & oldPath
    End If

    If Dir(newPath) = vbNullString Then
        Err.Raise vbObjectError + 502, "ex_SourceLoader", "NewFilePath не найден: " & newPath
    End If

    ImportTableToInternal oldPath, oldTableName, "Old"
    ImportTableToInternal newPath, newTableName, "New"
End Sub

Public Sub LoadStateEventsFromConfigToInternalSheets()

    Dim statePath As String
    Dim stateTableName As String
    Dim eventsPath As String
    Dim eventsTableName As String

    statePath = ResolvePath(GetConfigValueSafe("StateFilePath"))
    stateTableName = Trim$(GetConfigValueSafe("StateTableName"))

    eventsPath = ResolvePath(GetConfigValueSafe("EventsFilePath"))
    eventsTableName = Trim$(GetConfigValueSafe("EventsTableName"))

    If Len(statePath) = 0 Or Len(stateTableName) = 0 Or Len(eventsPath) = 0 Or Len(eventsTableName) = 0 Then
        Err.Raise vbObjectError + 530, "ex_SourceLoader", _
            "Конфиг не заполнен. Нужны StateFilePath/StateTableName/EventsFilePath/EventsTableName."
    End If

    If Dir(statePath) = vbNullString Then
        Err.Raise vbObjectError + 531, "ex_SourceLoader", "StateFilePath не найден: " & statePath
    End If

    If Dir(eventsPath) = vbNullString Then
        Err.Raise vbObjectError + 532, "ex_SourceLoader", "EventsFilePath не найден: " & eventsPath
    End If

    ImportTableToInternal statePath, stateTableName, "State"
    ImportTableToInternal eventsPath, eventsTableName, "Events"

End Sub


' =============================================================================
' Internal
' =============================================================================

Private Sub ImportTableToInternal( _
    ByVal sourceWorkbookPath As String, _
    ByVal sourceTableName As String, _
    ByVal targetBaseName As String _
)
    Dim wbSrc As Workbook
    Dim wsDst As Worksheet
    Dim srcListObject As ListObject
    Dim srcRange As Range
    Dim dstRange As Range
    Dim fullDstName As String

    fullDstName = "g_" & targetBaseName

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo EH

    Set wbSrc = Workbooks.Open( _
        Filename:=sourceWorkbookPath, _
        ReadOnly:=True, _
        UpdateLinks:=0 _
    )

    Set srcListObject = FindListObjectByName(wbSrc, sourceTableName)
    If srcListObject Is Nothing Then
        Err.Raise vbObjectError + 510, "ex_SourceLoader", _
            "Таблица '" & sourceTableName & "' не найдена в файле: " & sourceWorkbookPath
    End If

    ' Range таблицы с заголовками
    Set srcRange = srcListObject.Range

    Set wsDst = GetOrCreateWorksheetByFullName(fullDstName)
    wsDst.Cells.Clear

    Set dstRange = wsDst.Range( _
        wsDst.Cells(1, 1), _
        wsDst.Cells(srcRange.Rows.Count, srcRange.Columns.Count) _
    )

    dstRange.Value = srcRange.Value
    wsDst.Columns.AutoFit

    ex_SheetTheme.ApplyDarkThemeToSheet wsDst

Cleanup:
    On Error Resume Next
    If Not wbSrc Is Nothing Then
        wbSrc.Close SaveChanges:=False
    End If
    On Error GoTo 0

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

EH:
    Dim errText As String
    errText = "Ошибка импорта (" & targetBaseName & "): " & Err.Description

    On Error Resume Next
    If Not wbSrc Is Nothing Then
        wbSrc.Close SaveChanges:=False
    End If
    On Error GoTo 0

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Err.Raise vbObjectError + 520, "ex_SourceLoader", errText
End Sub

Private Function FindListObjectByName( _
    ByVal wbSrc As Workbook, _
    ByVal tableName As String _
) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In wbSrc.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set FindListObjectByName = lo
                Exit Function
            End If
        Next lo
    Next ws

    Set FindListObjectByName = Nothing
End Function

Private Function GetOrCreateWorksheetByFullName(ByVal fullName As String) As Worksheet
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, fullName, vbTextCompare) = 0 Then
            Set GetOrCreateWorksheetByFullName = ws
            Exit Function
        End If
    Next ws

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = fullName
    Call ex_ApplyDefaultSheetView(ws)

    Set GetOrCreateWorksheetByFullName = ws
End Function

Private Function ResolvePath(ByVal inputPath As String) As String
    Dim basePath As String

    inputPath = Trim$(inputPath)

    If Len(inputPath) = 0 Then
        ResolvePath = vbNullString
        Exit Function
    End If

    If Left$(inputPath, 2) = "\\" Or InStr(1, inputPath, ":\", vbTextCompare) > 0 Then
        ResolvePath = inputPath
        Exit Function
    End If

    basePath = ThisWorkbook.Path
    If Len(basePath) = 0 Then
        ResolvePath = inputPath
        Exit Function
    End If

    If Right$(basePath, 1) <> "\" Then
        basePath = basePath & "\"
    End If

    ResolvePath = basePath & inputPath
End Function

Private Function GetConfigValueSafe(ByVal keyName As String) As String
    On Error GoTo Fallback

    GetConfigValueSafe = ex_Config.GetConfigValue(keyName, vbNullString)
    Exit Function

Fallback:
    GetConfigValueSafe = GetConfigValueFromDevFallback(keyName)
End Function

Private Function GetConfigValueFromDevFallback(ByVal keyName As String) As String
    Dim wsDev As Worksheet
    Dim rng As Range
    Dim foundCell As Range

    Set wsDev = ThisWorkbook.Worksheets("Dev")
    Set rng = wsDev.Range("A1:B200")

    Set foundCell = rng.Columns(1).Find( _
        What:=keyName, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        MatchCase:=False _
    )

    If foundCell Is Nothing Then
        GetConfigValueFromDevFallback = vbNullString
        Exit Function
    End If

    GetConfigValueFromDevFallback = CStr(wsDev.Cells(foundCell.Row, foundCell.Column + 1).Value)
End Function
