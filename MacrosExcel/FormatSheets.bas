Option Explicit

Private mStatusClearAt As Date
Private mStatusClearPending As Boolean

Private Const SHEET_BACKUP_SUFFIX As String = " (orig)"
Private Const PROFILE_DEFAULT As String = "default"

Private Const SHEET_NAME_SPS As String = "ШПС"
Private Const PROFILE_SPS As String = "sps"
Private Const SPS_HEADER_SCAN_LIMIT As Long = 30

Private Const SHEET_NAME_TEMP_ARRIVED As String = "4. Тимчасово прибулі"
Private Const PROFILE_TEMP_ARRIVED As String = "temp_arrived"

Private Const SHEET_NAME_TEMP_ABSENT As String = "5. Тимчасово відсутні"
Private Const PROFILE_TEMP_ABSENT As String = "temp_absent"

Public Sub dev_FormatKnownSupportedSheets()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim startSheet As Worksheet: Set startSheet = ws
    Dim sheetProfile As String
    sheetProfile = ResolveSheetProfile(ws.Name)

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

    resultMsg = ApplySheetProfile(ws, sheetProfile)

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

Private Function ApplySheetProfile(ws As Worksheet, sheetProfile As String) As String
    Select Case sheetProfile
        Case PROFILE_SPS
            ApplySheetProfile = ApplyStyleProfileSps(ws, SPS_HEADER_SCAN_LIMIT)
        Case PROFILE_TEMP_ARRIVED
            ApplySheetProfile = ApplyStyleProfileTemporaryArrived(ws)
        Case PROFILE_TEMP_ABSENT
            ApplySheetProfile = ApplyStyleProfileTemporaryAbsent(ws)
        Case Else
            ApplySheetProfile = ApplyStyleProfileDefault(ws)
    End Select
End Function

Private Function ResolveSheetProfile(sheetName As String) As String
    Dim normalized As String
    normalized = NormalizeSheetName(sheetName)

    If IsSpsSheet(normalized) Then
        ResolveSheetProfile = PROFILE_SPS
    ElseIf IsSheetNameOrCopy(normalized, SHEET_NAME_TEMP_ARRIVED) Then
        ResolveSheetProfile = PROFILE_TEMP_ARRIVED
    ElseIf IsSheetNameOrCopy(normalized, SHEET_NAME_TEMP_ABSENT) Then
        ResolveSheetProfile = PROFILE_TEMP_ABSENT
    Else
        ResolveSheetProfile = PROFILE_DEFAULT
    End If
End Function

Private Function ApplyStyleProfileSps(ws As Worksheet, headerScanLimit As Long) As String
    Dim headerRow As Long
    headerRow = DetectHeaderRow(ws, Array("#", "№ з/п"), headerScanLimit)
    If headerRow = 0 Then
        Err.Raise vbObjectError + 1301, "ApplyStyleProfileSps", _
                  "Header row was not found in the first " & headerScanLimit & " rows."
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
        Err.Raise vbObjectError + 1302, "ApplyStyleProfileSps", moveInfo
    End If

    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    ApplySheetTheme ws
    HideNotDesiredColumns ws, headerRow, lastCol, keep
    ApplyFinalFormatting ws, headerRow, lastRow

    ApplyStyleProfileSps = "Layout applied."
End Function

Private Function ApplyStyleProfileTemporaryArrived(ws As Worksheet) As String
    ApplyTemporaryPersonnelBaseStyle ws
    ApplyTemporaryArrivedColumnWidths ws
    ApplyStyleProfileTemporaryArrived = "Layout applied for '" & SHEET_NAME_TEMP_ARRIVED & "'."
End Function

Private Function ApplyStyleProfileTemporaryAbsent(ws As Worksheet) As String
    ApplyTemporaryPersonnelBaseStyle ws
    ApplyTemporaryAbsentColumnWidths ws
    ApplyStyleProfileTemporaryAbsent = "Layout applied for '" & SHEET_NAME_TEMP_ABSENT & "'."
End Function

Private Function ApplyStyleProfileDefault(ws As Worksheet) As String
    ApplySheetTheme ws
    ApplyStyleProfileDefault = "Theme applied only. Non-target sheet: '" & ws.Name & "'."
End Function

Private Function IsSpsSheet(sheetName As String) As Boolean
    IsSpsSheet = IsSheetNameOrCopy(NormalizeSheetName(sheetName), SHEET_NAME_SPS)
End Function

Private Function IsSheetNameOrCopy(normalizedSheetName As String, baseSheetName As String) As Boolean
    If StrComp(normalizedSheetName, baseSheetName, vbTextCompare) = 0 Then
        IsSheetNameOrCopy = True
    ElseIf Len(normalizedSheetName) > Len(baseSheetName) + 2 Then
        IsSheetNameOrCopy = (StrComp(Left$(normalizedSheetName, Len(baseSheetName) + 2), _
                             baseSheetName & " (", vbTextCompare) = 0)
    Else
        IsSheetNameOrCopy = False
    End If
End Function

Private Function NormalizeSheetName(sheetName As String) As String
    NormalizeSheetName = NormalizeHeader(sheetName)
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
    Dim sourceWs As Worksheet: Set sourceWs = ws
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
    Set backupWs = ActiveSheet
    backupWs.Name = backupSheetName

    backupWs.Visible = xlSheetVisible
    sourceWs.Activate
    Exit Sub

BackupFail:
    On Error Resume Next
    Application.DisplayAlerts = prevDisplayAlerts
    sourceWs.Activate
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
    dstCol = FindAnyColInArrayPreferVisibleRightmost(ws, headers, anchorHeaders)
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

Private Function FindAnyColInArrayPreferVisibleRightmost(ws As Worksheet, headers As Variant, variants As Variant) As Long
    Dim i As Long, key As String, c As Long
    Dim bestVisible As Long
    Dim bestAny As Long

    For i = LBound(variants) To UBound(variants)
        key = NormalizeHeader(CStr(variants(i)))
        For c = 1 To UBound(headers, 2)
            If NormalizeHeader(CStr(headers(1, c))) = key Then
                If c > bestAny Then bestAny = c
                If Not ws.Columns(c).Hidden Then
                    If c > bestVisible Then bestVisible = c
                End If
            End If
        Next c
    Next i

    If bestVisible > 0 Then
        FindAnyColInArrayPreferVisibleRightmost = bestVisible
    Else
        FindAnyColInArrayPreferVisibleRightmost = bestAny
    End If
End Function

Private Sub ApplyTemporaryPersonnelBaseStyle(ws As Worksheet)
    ApplySheetTheme ws
    ApplyTemporaryAbsentKeywordColors ws
End Sub

Private Sub ApplyTemporaryAbsentKeywordColors(ws As Worksheet)
    Dim usedRng As Range
    Set usedRng = ws.UsedRange
    If usedRng Is Nothing Then Exit Sub

    ColorKeywordsInRange usedRng, TemporaryAbsentRedKeywords(), RGB(255, 0, 0)
    ColorKeywordsInRange usedRng, TemporaryAbsentBlueKeywords(), RGB(0, 176, 240)
End Sub

Private Sub ApplyTemporaryAbsentColumnWidths(ws As Worksheet)
    ws.Columns("F:M").AutoFit
    ws.Columns("O:O").AutoFit
End Sub

Private Sub ApplyTemporaryArrivedColumnWidths(ws As Worksheet)
    ws.Columns("F:K").AutoFit
End Sub

Private Sub ColorKeywordsInRange(targetRange As Range, keywords As Variant, fontColor As Long)
    Dim cell As Range
    For Each cell In targetRange.Cells
        Dim v As Variant
        v = cell.Value2
        If Not IsError(v) Then
            Dim txt As String
            txt = CStr(v)
            If Len(txt) > 0 Then
                Dim i As Long
                For i = LBound(keywords) To UBound(keywords)
                    ApplyKeywordColorInCell cell, txt, CStr(keywords(i)), fontColor
                Next i
            End If
        End If
    Next cell
End Sub

Private Sub ApplyKeywordColorInCell(targetCell As Range, ByVal cellText As String, ByVal keyword As String, ByVal fontColor As Long)
    If Len(keyword) = 0 Then Exit Sub

    Dim pos As Long
    Dim searchFrom As Long
    searchFrom = 1

    Do
        pos = InStr(searchFrom, cellText, keyword, vbBinaryCompare) ' case-sensitive
        If pos = 0 Then Exit Do
        targetCell.Characters(pos, Len(keyword)).Font.Color = fontColor
        searchFrom = pos + Len(keyword)
    Loop
End Sub

Private Function TemporaryAbsentRedKeywords() As Variant
    TemporaryAbsentRedKeywords = Array( _
        "ПЕРЕБУВАВ у відпустці для лікуванні", _
        "ПЕРЕБУВАВ у відпустці для лікування", _
        "СЗЧ", _
        "ВИБУВАЄ у відпустку для лікування", _
        "ПЕРЕБУВАВ на лікуванні", _
        "ПРИБУВ" _
    )
End Function

Private Function TemporaryAbsentBlueKeywords() As Variant
    TemporaryAbsentBlueKeywords = Array("ТВО")
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
