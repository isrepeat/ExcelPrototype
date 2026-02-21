Option Explicit

Private m_LastFilterDateText As String
Private m_StatusClearAt As Date
Private m_StatusClearScheduled As Boolean

' ============================================================
' Macro #1: Filter rows (by hiding non-matching rows)
' - Prompts for a date (dd.mm.yyyy)
' - Searches all data rows from HEADER_ROW+1 to last used row
' - Keeps separator rows visible (multiple marker texts supported)
' - Checks column E (DATE_COLUMN) with RegExp
' - Hides rows that do NOT match
' ============================================================
Public Sub FilterHospByDate_Regex_ToMarker()

    ' === CONFIG ===
    Const DATE_COLUMN As Long = 5          ' Column E: "Äàòà ãîñï³òàë³çàö³¿ / ïåðåâåäåííÿ"
    Const HEADER_ROW As Long = 3           ' Header row (E3 is the header cell)
    ' ==============

    Dim MARKER_TEXTS As Variant
    MARKER_TEXTS = Array("ВЛК Амбулаторно", "Виписані", "Виписані з ВЛК амбулаторно")

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Ask user for date (text)
    Dim inputValue As String
    inputValue = InputBox( _
        Prompt:="Enter date in format dd.mm.yyyy (e.g. 23.02.2025):", _
        Title:="Filter by Date (RegExp)", _
        Default:=m_LastFilterDateText)

    Dim dateText As String
    dateText = Trim$(inputValue)
    If Len(dateText) = 0 Then Exit Sub

    ' Normalize date if Excel can parse it
    On Error Resume Next
    Dim tmpDate As Date
    tmpDate = CDate(dateText)
    If Err.Number = 0 Then
        dateText = Format$(tmpDate, "dd.mm.yyyy")
    End If
    Err.Clear
    On Error GoTo 0

    m_LastFilterDateText = dateText

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
    reDateToken.pattern = "([0-9]{1,2}\.[0-9]{1,2}\.[0-9]{4})"
    reDateToken.IgnoreCase = True
    reDateToken.Global = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim markerRows As Object
    Set markerRows = CreateObject("Scripting.Dictionary")

    Dim markerText As Variant
    Dim foundCell As Range
    Dim firstFoundAddress As String

    For Each markerText In MARKER_TEXTS
        Set foundCell = ws.Cells.Find(What:=CStr(markerText), LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
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

    ' First, show all rows in our target block (reset previous hide-based filtering)
    Dim r As Long
    For r = firstDataRow To lastDataRow
        ws.Rows(r).Hidden = False
    Next r

    ' Apply "filter" by hiding non-matching rows
    Dim cellText As String
    Dim cellValue As Variant
    Dim rowMatched As Boolean
    Dim matches As Object
    Dim m As Object
    Dim matchedCount As Long
    matchedCount = 0

    For r = firstDataRow To lastDataRow
        If markerRows.Exists(CStr(r)) Then
            ws.Rows(r).Hidden = False
            GoTo ContinueRow
        End If

        rowMatched = False
        cellValue = ws.Cells(r, DATE_COLUMN).Value

        If IsDate(cellValue) Then
            rowMatched = (Format$(CDate(cellValue), "dd.mm.yyyy") = dateText)
        Else
            cellText = CStr(ws.Cells(r, DATE_COLUMN).Value2)

            ' Normalize line breaks (Alt+Enter)
            cellText = Replace(cellText, vbCrLf, vbLf)
            cellText = Replace(cellText, vbCr, vbLf)

            If reDateToken.Test(cellText) Then
                Set matches = reDateToken.Execute(cellText)
                For Each m In matches
                    On Error Resume Next
                    If Format$(CDate(CStr(m.Value)), "dd.mm.yyyy") = dateText Then
                        rowMatched = True
                    End If
                    Err.Clear
                    On Error GoTo 0

                    If rowMatched Then Exit For
                Next m
            End If
        End If

        If rowMatched Then
            ws.Rows(r).Hidden = False
            matchedCount = matchedCount + 1
        Else
            ws.Rows(r).Hidden = True
        End If

ContinueRow:
    Next r

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ShowStatusMessage "Done. Matching rows: " & matchedCount, 4

End Sub


' ============================================================
' Macro #2: Reset (show all rows back)
' - Unhides all rows from HEADER_ROW+1 up to last used row
' ============================================================
Public Sub ResetHospFilter_ToMarker()

    ' === CONFIG ===
    Const HEADER_ROW As Long = 3
    ' ==============

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