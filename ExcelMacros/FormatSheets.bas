Option Explicit

Private mStatusClearAt As Date
Private mStatusClearPending As Boolean
Private Const SHEET_NAME_SPS As String = "ШПС"
Private Const SHEET_BACKUP_SUFFIX As String = " (orig)"

Public Sub dev_FormatKnownSupportedSheets()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim startSheet As Worksheet: Set startSheet = ws
    Const HEADER_SCAN_LIMIT As Long = 30
    Dim isTargetSheet As Boolean
    isTargetSheet = IsSpsSheet(ws.Name)

    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation
    Dim errText As String
    Dim resultMsg As String

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    If isTargetSheet Then
        Dim headerRow As Long
        headerRow = DetectHeaderRow(ws, Array("#", "№ з/п"), HEADER_SCAN_LIMIT)
        If headerRow = 0 Then
            MsgBox "Header row was not found in the first " & HEADER_SCAN_LIMIT & " rows.", vbExclamation
            GoTo Cleanup
        End If

        SaveStyleBackup ws

        Dim keepHeaders As Variant
        keepHeaders = Array( _
            "#", _
            "Код посади", _
            "Військове звання", _
            "Прізвище, ім’я, по батькові", _
            "Повна назва посади", _
            "ІПН", _
            "Вид військової служби", _
            "Дата підписання контракту", _
            "Дата завершення контракту", _
            "Дата та № наказу про присвоэння звання", _
            "Дата та № наказу призначення на посаду", _
            "Дата та № наказу про зарахування", _
            "Дата та № наказу доступу до ""Таємно""", _
            "Прибув з:", _
            "Місцезнаходження", _
            "Дата та № наказу місцезнаходження", _
            "Х1" _
        )

        Dim anchorHeaders As Variant
        anchorHeaders = Array("Прізвище, ім’я, по батькові", "Прізвище, ім'я, по батькові")

        Dim keep As Object
        Set keep = CreateObject("Scripting.Dictionary")
        keep.CompareMode = vbTextCompare
        BuildKeepMap keepHeaders, keep

        Dim lastCol As Long
        Dim lastRow As Long
        Dim moveInfo As String

        lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
        lastRow = LastUsedRow(ws)
        If lastRow < 1 Then lastRow = 1

        If Not MoveColumnAfterHeader(ws, headerRow, lastRow, "Повна назва посади", anchorHeaders, moveInfo) Then
            MsgBox moveInfo, vbExclamation
            GoTo Cleanup
        End If

        lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
        ApplySheetTheme ws
        HideNotDesiredColumns ws, headerRow, lastCol, keep
        ApplyFinalFormatting ws, headerRow, lastRow
        resultMsg = "Layout applied."
    Else
        ApplySheetTheme ws
        resultMsg = "Theme applied only. Non-target sheet: '" & ws.Name & "'."
    End If

Cleanup:
    Application.Calculation = prevCalculation
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    On Error Resume Next
    startSheet.Activate
    On Error GoTo 0

    If Len(errText) > 0 Then
        MsgBox "Layout failed: " & errText, vbCritical
        Exit Sub
    End If

    If Len(resultMsg) > 0 Then
        ShowStatusForSeconds resultMsg, 3
    End If
    Exit Sub

Fail:
    errText = Err.Description
    Resume Cleanup
End Sub

Private Function IsSpsSheet(sheetName As String) As Boolean
    Dim normalized As String
    normalized = Replace(sheetName, ChrW(160), " ")
    normalized = Trim$(normalized)

    If StrComp(normalized, SHEET_NAME_SPS, vbTextCompare) = 0 Then
        IsSpsSheet = True
    ElseIf Len(normalized) > Len(SHEET_NAME_SPS) + 2 Then
        IsSpsSheet = (StrComp(Left$(normalized, Len(SHEET_NAME_SPS) + 2), SHEET_NAME_SPS & " (", vbTextCompare) = 0)
    Else
        IsSpsSheet = False
    End If
End Function

Private Function DetectHeaderRow(ws As Worksheet, keys As Variant, maxRowsToCheck As Long) As Long
    Dim r As Long, c As Long, lastCol As Long
    Dim rowData As Variant
    Dim keyMap As Object
    Set keyMap = CreateObject("Scripting.Dictionary")
    keyMap.CompareMode = vbTextCompare

    Dim i As Long, key As String
    For i = LBound(keys) To UBound(keys)
        key = NormalizeHeader(CStr(keys(i)))
        If Len(key) > 0 Then keyMap(key) = True
    Next i

    For r = 1 To maxRowsToCheck
        lastCol = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        If lastCol > 0 Then
            rowData = ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Value2
            For c = 1 To lastCol
                key = NormalizeHeader(CStr(rowData(1, c)))
                If keyMap.Exists(key) Then
                    DetectHeaderRow = r
                    Exit Function
                End If
            Next c
        End If
    Next r

    DetectHeaderRow = 0
End Function

Private Sub BuildKeepMap(headers As Variant, keep As Object)
    Dim i As Long, key As String
    For i = LBound(headers) To UBound(headers)
        key = NormalizeHeader(CStr(headers(i)))
        If Len(key) > 0 Then keep(key) = True
    Next i
End Sub

Private Sub SaveStyleBackup(ws As Worksheet)
    On Error GoTo BackupFail

    Dim wb As Workbook: Set wb = ws.Parent
    Dim backupSheetName As String
    backupSheetName = GetBackupSheetName()
    Dim prevDisplayAlerts As Boolean
    prevDisplayAlerts = Application.DisplayAlerts
    Dim hadOldBackup As Boolean

    If wb.ProtectStructure Then
        Err.Raise vbObjectError + 1201, "SaveStyleBackup", _
                  "Workbook structure is protected. Cannot create backup sheet '" & backupSheetName & "'."
    End If

    hadOldBackup = WorksheetExists(wb, backupSheetName)

    Application.DisplayAlerts = False
    If hadOldBackup Then wb.Worksheets(backupSheetName).Delete
    Application.DisplayAlerts = prevDisplayAlerts

    If WorksheetExists(wb, backupSheetName) Then
        Err.Raise vbObjectError + 1202, "SaveStyleBackup", _
                  "Could not replace backup sheet '" & backupSheetName & "'."
    End If

    ws.Copy After:=wb.Worksheets(wb.Worksheets.Count)

    Dim backupWs As Worksheet
    Set backupWs = wb.Worksheets(wb.Worksheets.Count)

    ' TODO: Backup sheet rename is unstable in current environment/workbook state.
    ' TODO: Sometimes Excel leaves copied sheet with auto-generated name and rename fails.
    ' TODO: Need deeper investigation (workbook structure state, name collisions, Personal.xlsb context).
    Dim renameErrText As String
    If Not TryRenameSheet(backupWs, backupSheetName, 3, renameErrText) Then
        Err.Raise vbObjectError + 1203, "SaveStyleBackup", _
                  "Could not rename backup sheet to '" & backupSheetName & "': " & renameErrText
    End If

    backupWs.Visible = xlSheetVeryHidden
    Exit Sub

BackupFail:
    On Error Resume Next
    Application.DisplayAlerts = prevDisplayAlerts
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function GetBackupSheetName() As String
    GetBackupSheetName = SHEET_NAME_SPS & SHEET_BACKUP_SUFFIX
End Function

Private Function WorksheetExists(wb As Workbook, sheetName As String) As Boolean
    Dim tmp As Worksheet
    On Error Resume Next
    Set tmp = wb.Worksheets(sheetName)
    On Error GoTo 0
    WorksheetExists = Not tmp Is Nothing
End Function

Private Function TryRenameSheet(sheetToRename As Worksheet, targetName As String, retries As Long, _
                                ByRef lastErrText As String) As Boolean
    Dim i As Long
    Dim errNum As Long

    ' TODO: Temporary retry workaround; root cause of rename failure is still unknown.
    ' TODO: Replace with deterministic rename strategy after root-cause analysis.
    If retries < 1 Then retries = 1

    For i = 1 To retries
        On Error Resume Next
        sheetToRename.Name = targetName
        errNum = Err.Number
        lastErrText = Err.Description
        On Error GoTo 0

        If errNum = 0 Then
            TryRenameSheet = True
            Exit Function
        End If

        DoEvents
    Next i

    TryRenameSheet = False
End Function

Private Sub ShowStatusForSeconds(messageText As String, secondsCount As Long)
    If secondsCount < 1 Then secondsCount = 1

    On Error Resume Next
    If mStatusClearPending Then
        Application.OnTime EarliestTime:=mStatusClearAt, Procedure:=StatusClearProcedureName(), Schedule:=False
        mStatusClearPending = False
    End If
    On Error GoTo 0

    Application.StatusBar = messageText
    mStatusClearAt = Now + (secondsCount / 86400#)
    mStatusClearPending = True

    On Error Resume Next
    Application.OnTime EarliestTime:=mStatusClearAt, Procedure:=StatusClearProcedureName(), Schedule:=True
    On Error GoTo 0
End Sub

Public Sub dev_ClearStatusBarMessage()
    Application.StatusBar = False
    mStatusClearPending = False
End Sub

Private Function StatusClearProcedureName() As String
    StatusClearProcedureName = "'" & ThisWorkbook.Name & "'!dev_ClearStatusBarMessage"
End Function

Private Function MoveColumnAfterHeader( _
    ws As Worksheet, _
    headerRow As Long, _
    lastRow As Long, _
    movingHeader As String, _
    anchorHeaders As Variant, _
    ByRef info As String) As Boolean

    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    If lastCol = 0 Then
        info = "Header row is empty."
        Exit Function
    End If

    Dim headers As Variant
    headers = ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, lastCol)).Value2

    Dim srcCol As Long
    srcCol = FindColInArrayPreferVisible(ws, headers, NormalizeHeader(movingHeader))
    If srcCol = 0 Then
        info = "Column '" & movingHeader & "' was not found."
        Exit Function
    End If

    Dim dstCol As Long
    dstCol = FindAnyColInArrayPreferVisible(ws, headers, anchorHeaders)
    If dstCol = 0 Then
        info = "Target column 'Прізвище, ім’я, по батькові' was not found."
        Exit Function
    End If

    If srcCol = dstCol + 1 Then
        MoveColumnAfterHeader = True
        Exit Function
    End If

    Dim insertCol As Long
    insertCol = dstCol + 1

    ws.Columns(srcCol).Cut
    ws.Columns(insertCol).Insert Shift:=xlToRight
    Application.CutCopyMode = False

    MoveColumnAfterHeader = True
End Function

Private Function FindColInArrayPreferVisible(ws As Worksheet, headers As Variant, wantedKey As String) As Long
    Dim c As Long
    Dim firstMatch As Long

    For c = 1 To UBound(headers, 2)
        If NormalizeHeader(CStr(headers(1, c))) = wantedKey Then
            If firstMatch = 0 Then firstMatch = c
            If Not ws.Columns(c).Hidden Then
                FindColInArrayPreferVisible = c
                Exit Function
            End If
        End If
    Next c

    FindColInArrayPreferVisible = firstMatch
End Function

Private Function FindAnyColInArrayPreferVisible(ws As Worksheet, headers As Variant, variants As Variant) As Long
    Dim i As Long, key As String, c As Long
    Dim firstMatch As Long

    For i = LBound(variants) To UBound(variants)
        key = NormalizeHeader(CStr(variants(i)))
        For c = 1 To UBound(headers, 2)
            If NormalizeHeader(CStr(headers(1, c))) = key Then
                If firstMatch = 0 Then firstMatch = c
                If Not ws.Columns(c).Hidden Then
                    FindAnyColInArrayPreferVisible = c
                    Exit Function
                End If
            End If
        Next c
    Next i

    FindAnyColInArrayPreferVisible = firstMatch
End Function

Private Function FindColInArray(headers As Variant, wantedKey As String) As Long
    Dim c As Long
    For c = 1 To UBound(headers, 2)
        If NormalizeHeader(CStr(headers(1, c))) = wantedKey Then
            FindColInArray = c
            Exit Function
        End If
    Next c
    FindColInArray = 0
End Function

Private Function FindAnyColInArray(headers As Variant, variants As Variant) As Long
    Dim i As Long, key As String, c As Long
    For i = LBound(variants) To UBound(variants)
        key = NormalizeHeader(CStr(variants(i)))
        For c = 1 To UBound(headers, 2)
            If NormalizeHeader(CStr(headers(1, c))) = key Then
                FindAnyColInArray = c
                Exit Function
            End If
        Next c
    Next i
    FindAnyColInArray = 0
End Function

Private Sub HideNotDesiredColumns(ws As Worksheet, headerRow As Long, lastCol As Long, keep As Object)
    Dim headers As Variant
    headers = ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, lastCol)).Value2

    Dim c As Long, key As String, hideCol As Boolean
    For c = 1 To lastCol
        key = NormalizeHeader(CStr(headers(1, c)))
        hideCol = (Len(key) = 0) Or (Not keep.Exists(key))
        If ws.Columns(c).Hidden <> hideCol Then
            ws.Columns(c).Hidden = hideCol
        End If
    Next c
End Sub

Private Sub ApplyFinalFormatting(ws As Worksheet, headerRow As Long, lastRow As Long)
    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    If lastCol = 0 Then Exit Sub

    Dim headers As Variant
    headers = ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, lastCol)).Value2

    SetColumnWidthByHeader ws, headers, "Код посади", 30
    SetColumnWidthByHeader ws, headers, "Військове звання", 30
    SetColumnWidthByHeader ws, headers, "Прізвище, ім’я, по батькові", 54
    SetColumnWidthByHeader ws, headers, "Повна назва посади", 50
    SetColumnWidthByHeader ws, headers, "ІПН", 28

    ApplyHashColumnFormatting ws, headers, headerRow, lastRow
End Sub

Private Sub SetColumnWidthByHeader(ws As Worksheet, headers As Variant, headerName As String, widthValue As Double)
    Dim colIdx As Long
    colIdx = FindColInArray(headers, NormalizeHeader(headerName))
    If colIdx > 0 Then
        ws.Columns(colIdx).ColumnWidth = widthValue
    End If
End Sub

Private Sub ApplyHashColumnFormatting(ws As Worksheet, headers As Variant, headerRow As Long, lastRow As Long)
    Dim hashCol As Long
    hashCol = FindColInArray(headers, NormalizeHeader("#"))
    If hashCol = 0 Then Exit Sub

    Dim firstDataRow As Long
    firstDataRow = headerRow + 1
    If firstDataRow > lastRow Then Exit Sub

    Dim rngData As Range
    Set rngData = ws.Range(ws.Cells(firstDataRow, hashCol), ws.Cells(lastRow, hashCol))
    rngData.Font.Size = 22
    rngData.Font.Color = RGB(226, 107, 10)

    rngData.FormatConditions.Delete

    Dim colLetter As String
    colLetter = Split(ws.Cells(1, hashCol).Address(False, False), "1")(0)

    Dim sep As String
    sep = Application.International(xlListSeparator)

    Dim normalizedExpr As String
    normalizedExpr = "TRIM(SUBSTITUTE($" & colLetter & firstDataRow & sep & "CHAR(160)" & sep & """ ""))"

    Dim formulaText As String
    formulaText = "=OR(" & normalizedExpr & "=""" & "РОЗП" & """" & sep & _
                        normalizedExpr & "=""" & "СПИС" & """)"

    With rngData.FormatConditions.Add(Type:=xlExpression, Formula1:=formulaText)
        .Interior.Color = RGB(255, 0, 0)
        .Font.Color = RGB(226, 107, 10)
    End With
End Sub

Private Sub ApplySheetTheme(ws As Worksheet)
    On Error GoTo Fallback

    With ws.Columns
        .Interior.Color = RGB(38, 38, 38)    ' #262626
        .Font.Color = RGB(118, 147, 60)      ' #76933C
    End With
    Exit Sub

Fallback:
    Err.Clear
    With ws.UsedRange
        .Interior.Color = RGB(38, 38, 38)
        .Font.Color = RGB(118, 147, 60)
    End With
End Sub

Private Function LastUsedRow(ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If f Is Nothing Then
        LastUsedRow = 0
    Else
        LastUsedRow = f.Row
    End If
End Function

Private Function NormalizeHeader(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, ChrW(160), " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Trim$(t)

    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop

    NormalizeHeader = t
End Function
