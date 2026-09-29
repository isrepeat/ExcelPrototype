Option Explicit

Private m_LastFilterDateText As String
Private m_StatusClearAt As Date
Private m_StatusClearScheduled As Boolean
Private Const HELPER_HEADER As String = "__tmp_date_match__"
Private Const HELPER_COLUMN_INDEX As Long = 16384 ' XFD

' ============================================================
' Macro #1: Filter rows (via native AutoFilter)
' - Prompts for a date (dd.mm.yyyy)
' - Searches all data rows from HEADER_ROW+1 to last used row
' - Keeps separator rows visible (multiple marker texts supported)
' - Checks column E (DATE_COLUMN)
' - Uses helper column + AutoFilter (no manual row hiding)
' ============================================================
Public Sub FilterHospByDate_Regex_ToMarker()

    Const DATE_COLUMN As Long = 5
    Const HEADER_ROW As Long = 3

    Dim markerTexts As Variant
    markerTexts = Array("ВЛК Амбулаторно", "Виписані", "Виписані з ВЛК амбулаторно")

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim inputValue As String
    inputValue = InputBox( _
        Prompt:="Enter date in format dd.mm.yyyy (e.g. 23.02.2025):", _
        Title:="Filter by Date (AutoFilter)", _
        Default:=m_LastFilterDateText)

    Dim dateText As String
    dateText = Trim$(inputValue)
    If Len(dateText) = 0 Then Exit Sub

    On Error Resume Next
    Dim tmpDate As Date
    tmpDate = CDate(dateText)
    If Err.Number = 0 Then
        dateText = Format$(tmpDate, "dd.mm.yyyy")
    End If
    Err.Clear
    On Error GoTo 0

    m_LastFilterDateText = dateText

    Dim expectedDateSerial As Long
    If Not TryParseDateToken(dateText, expectedDateSerial) Then
        ShowStatusMessage "Invalid date. Use dd.mm.yyyy"
        Exit Sub
    End If

    Dim firstDataRow As Long
    firstDataRow = HEADER_ROW + 1

    Dim lastDataRow As Long
    Dim lastUsedCell As Range
    Set lastUsedCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)

    If lastUsedCell Is Nothing Then
        ShowStatusMessage "Nothing to filter (sheet is empty)."
        Exit Sub
    End If

    lastDataRow = lastUsedCell.Row
    If lastDataRow < firstDataRow Then
        ShowStatusMessage "Nothing to filter (no rows below header)."
        Exit Sub
    End If

    Dim reDateToken As Object
    Set reDateToken = CreateObject("VBScript.RegExp")
    reDateToken.Pattern = "([0-9]{1,2}\.[0-9]{1,2}\.[0-9]{4})"
    reDateToken.IgnoreCase = True
    reDateToken.Global = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim markerRows As Object
    Set markerRows = CreateObject("Scripting.Dictionary")

    Dim markerText As Variant
    Dim foundCell As Range
    Dim firstFoundAddress As String

    For Each markerText In markerTexts
        Set foundCell = ws.Cells.Find(What:=CStr(markerText), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        If Not foundCell Is Nothing Then
            firstFoundAddress = foundCell.Address
            Do
                If foundCell.Row >= firstDataRow And foundCell.Row <= lastDataRow Then
                    markerRows(CStr(foundCell.Row)) = True
                End If
                Set foundCell = ws.Cells.FindNext(After:=foundCell)
            Loop While Not foundCell Is Nothing And foundCell.Address <> firstFoundAddress
        End If
    Next markerText

    Dim helperColumn As Long
    helperColumn = HELPER_COLUMN_INDEX

    Dim legacyHelperColumn As Long
    legacyHelperColumn = FindHelperColumn(ws, HEADER_ROW)
    If legacyHelperColumn > 0 And legacyHelperColumn <> helperColumn Then
        ws.Columns(legacyHelperColumn).Hidden = False
        ws.Range(ws.Cells(HEADER_ROW, legacyHelperColumn), ws.Cells(lastDataRow, legacyHelperColumn)).ClearContents
    End If

    ws.Cells(HEADER_ROW, helperColumn).Value = HELPER_HEADER

    Dim r As Long
    Dim matchedCount As Long
    matchedCount = 0

    For r = firstDataRow To lastDataRow
        If RowHasDateMatch(ws, r, DATE_COLUMN, expectedDateSerial, markerRows, reDateToken) Then
            ws.Cells(r, helperColumn).Value = True
            matchedCount = matchedCount + 1
        Else
            ws.Cells(r, helperColumn).Value = False
        End If
    Next r

    If ws.AutoFilterMode Then
        On Error Resume Next
        If ws.FilterMode Then ws.ShowAllData
        On Error GoTo 0
        ws.AutoFilterMode = False
    End If

    ws.Range(ws.Cells(HEADER_ROW, helperColumn), ws.Cells(lastDataRow, helperColumn)).AutoFilter Field:=1, Criteria1:=True
    ws.Columns(helperColumn).Hidden = True

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ShowStatusMessage "Done. Matching rows: " & matchedCount, 4
End Sub

' ============================================================
' Macro #2: Reset (clear filter and show all rows back)
' - Removes AutoFilter created by this macro
' - Clears helper column if present
' ============================================================
Public Sub ResetHospFilter_ToMarker()

    Const HEADER_ROW As Long = 3

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim firstDataRow As Long
    firstDataRow = HEADER_ROW + 1

    Dim lastDataRow As Long
    Dim lastUsedCell As Range
    Set lastUsedCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)

    If lastUsedCell Is Nothing Then
        ShowStatusMessage "Nothing to reset (sheet is empty)."
        Exit Sub
    End If

    lastDataRow = lastUsedCell.Row
    If lastDataRow < firstDataRow Then
        ShowStatusMessage "Nothing to reset (no rows below header)."
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    On Error GoTo 0

    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If

    Dim helperColumn As Long
    helperColumn = HELPER_COLUMN_INDEX
    ws.Columns(helperColumn).Hidden = False
    ws.Range(ws.Cells(HEADER_ROW, helperColumn), ws.Cells(lastDataRow, helperColumn)).ClearContents

    Dim legacyHelperColumn As Long
    legacyHelperColumn = FindHelperColumn(ws, HEADER_ROW)
    If legacyHelperColumn > 0 And legacyHelperColumn <> helperColumn Then
        ws.Columns(legacyHelperColumn).Hidden = False
        ws.Range(ws.Cells(HEADER_ROW, legacyHelperColumn), ws.Cells(lastDataRow, legacyHelperColumn)).ClearContents
    End If

    helperColumn = FindHelperColumn(ws, HEADER_ROW)
    If helperColumn > 0 Then
        ws.Columns(helperColumn).Hidden = False
        ws.Range(ws.Cells(HEADER_ROW, helperColumn), ws.Cells(lastDataRow, helperColumn)).ClearContents
    End If

    Dim r As Long
    For r = firstDataRow To lastDataRow
        ws.Rows(r).Hidden = False
    Next r

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ShowStatusMessage "Reset complete (all rows shown).", 2
End Sub

Private Sub ShowStatusMessage(ByVal message As String, Optional ByVal secondsToShow As Long = 2)
    If secondsToShow < 1 Then secondsToShow = 1

    If m_StatusClearScheduled Then
        On Error Resume Next
        Application.OnTime EarliestTime:=m_StatusClearAt, Procedure:="ClearStatusMessage", Schedule:=False
        On Error GoTo 0
        m_StatusClearScheduled = False
    End If

    Application.StatusBar = message

    m_StatusClearAt = Now + TimeSerial(0, 0, secondsToShow)
    Application.OnTime EarliestTime:=m_StatusClearAt, Procedure:="ClearStatusMessage", Schedule:=True
    m_StatusClearScheduled = True
End Sub

Public Sub ClearStatusMessage()
    Application.StatusBar = False
    m_StatusClearScheduled = False
End Sub

Private Function RowHasDateMatch(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal dateColumn As Long, ByVal expectedDateSerial As Long, ByVal markerRows As Object, ByVal reDateToken As Object) As Boolean
    If markerRows.Exists(CStr(rowIndex)) Then
        RowHasDateMatch = True
        Exit Function
    End If

    Dim cellValue As Variant
    cellValue = ws.Cells(rowIndex, dateColumn).Value

    If IsDate(cellValue) Then
        RowHasDateMatch = (CLng(CDate(cellValue)) = expectedDateSerial)
        Exit Function
    End If

    Dim cellText As String
    cellText = CStr(ws.Cells(rowIndex, dateColumn).Value2)
    cellText = Replace(cellText, vbCrLf, vbLf)
    cellText = Replace(cellText, vbCr, vbLf)

    If Not reDateToken.Test(cellText) Then
        RowHasDateMatch = False
        Exit Function
    End If

    Dim matches As Object
    Set matches = reDateToken.Execute(cellText)

    Dim m As Object
    Dim tokenSerial As Long
    For Each m In matches
        If TryParseDateToken(CStr(m.Value), tokenSerial) Then
            If tokenSerial = expectedDateSerial Then
                RowHasDateMatch = True
                Exit Function
            End If
        End If
    Next m

    RowHasDateMatch = False
End Function

Private Function TryParseDateToken(ByVal token As String, ByRef outDateSerial As Long) As Boolean
    Dim parts() As String
    parts = Split(Trim$(token), ".")

    If UBound(parts) <> 2 Then
        TryParseDateToken = False
        Exit Function
    End If

    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Or Not IsNumeric(parts(2)) Then
        TryParseDateToken = False
        Exit Function
    End If

    Dim d As Long
    Dim m As Long
    Dim y As Long
    d = CLng(parts(0))
    m = CLng(parts(1))
    y = CLng(parts(2))

    If d < 1 Or d > 31 Or m < 1 Or m > 12 Or y < 1900 Or y > 9999 Then
        TryParseDateToken = False
        Exit Function
    End If

    On Error GoTo ParseFail
    Dim dt As Date
    dt = DateSerial(y, m, d)
    If Year(dt) <> y Or Month(dt) <> m Or Day(dt) <> d Then
        GoTo ParseFail
    End If

    outDateSerial = CLng(dt)
    TryParseDateToken = True
    Exit Function

ParseFail:
    TryParseDateToken = False
End Function

Private Function FindHelperColumn(ByVal ws As Worksheet, ByVal headerRow As Long) As Long
    Dim foundCell As Range
    Set foundCell = ws.Rows(headerRow).Find(What:=HELPER_HEADER, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    If foundCell Is Nothing Then
        FindHelperColumn = 0
    Else
        FindHelperColumn = foundCell.Column
    End If
End Function
