Option Explicit

Private Const SHEET_NAV_STEP As Long = 10

Public Sub m_GoToFirstSheet()
    Dim wb As Workbook

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Active workbook was not found.", vbExclamation, "Navigation"
        Exit Sub
    End If

    If wb.Worksheets.Count = 0 Then
        MsgBox "There are no worksheets to navigate to.", vbExclamation, "Navigation"
        Exit Sub
    End If

    wb.Worksheets(1).Activate
End Sub

Public Sub m_GoToLastSheet()
    Dim wb As Workbook

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Active workbook was not found.", vbExclamation, "Navigation"
        Exit Sub
    End If

    If wb.Worksheets.Count = 0 Then
        MsgBox "There are no worksheets to navigate to.", vbExclamation, "Navigation"
        Exit Sub
    End If

    wb.Worksheets(wb.Worksheets.Count).Activate
End Sub

Public Sub m_GoForwardBy10Sheets()
    m_MoveBySheetsOffset SHEET_NAV_STEP
End Sub

Public Sub m_GoBackwardBy10Sheets()
    m_MoveBySheetsOffset -SHEET_NAV_STEP
End Sub

Public Sub m_MoveBySheetsOffset(ByVal offset As Long)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim currentIndex As Long
    Dim targetIndex As Long
    Dim maxIndex As Long

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Active workbook was not found.", vbExclamation, "Navigation"
        Exit Sub
    End If

    If wb.Worksheets.Count = 0 Then
        MsgBox "There are no worksheets to navigate to.", vbExclamation, "Navigation"
        Exit Sub
    End If

    If offset = 0 Then Exit Sub

    On Error Resume Next
    Set ws = ActiveSheet
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "The active sheet is not a worksheet.", vbExclamation, "Navigation"
        Exit Sub
    End If

    currentIndex = ws.Index
    maxIndex = wb.Worksheets.Count
    targetIndex = currentIndex + offset

    If targetIndex < 1 Then targetIndex = 1
    If targetIndex > maxIndex Then targetIndex = maxIndex

    wb.Worksheets(targetIndex).Activate
End Sub

Public Sub m_DeleteCurrentSheetWithConfirm()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim shouldDelete As VbMsgBoxResult
    Dim originalDisplayAlerts As Boolean

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Active workbook was not found.", vbExclamation, "Navigation"
        Exit Sub
    End If

    If wb.Worksheets.Count <= 1 Then
        MsgBox "You cannot delete the last worksheet in the workbook.", vbExclamation, "Navigation"
        Exit Sub
    End If

    On Error Resume Next
    Set ws = ActiveSheet
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "The active sheet is not a worksheet.", vbExclamation, "Navigation"
        Exit Sub
    End If

    shouldDelete = MsgBox( _
        "Delete the current worksheet '" & ws.Name & "'?", _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "Navigation" _
    )
    If shouldDelete <> vbYes Then Exit Sub

    On Error GoTo EH
    originalDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = originalDisplayAlerts
    Exit Sub
EH:
    Application.DisplayAlerts = originalDisplayAlerts
    MsgBox "Failed to delete worksheet: " & Err.Description, vbCritical, "Navigation"
End Sub

Public Sub m_CopyCurrentSheetToEndWithName()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim copiedWs As Worksheet
    Dim defaultName As String
    Dim newName As String
    Dim validationError As String
    Dim originalDisplayAlerts As Boolean

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Active workbook was not found.", vbExclamation, "Navigation"
        Exit Sub
    End If

    On Error Resume Next
    Set ws = ActiveSheet
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "The active sheet is not a worksheet.", vbExclamation, "Navigation"
        Exit Sub
    End If

    defaultName = mp_BuildUniqueSheetName(wb, ws.Name & "_copy")

    Do
        newName = InputBox("Enter the name of the new worksheet:", "Navigation", defaultName)
        If Len(newName) = 0 Then Exit Sub
        newName = Trim$(newName)

        If Not mp_IsValidWorksheetName(newName, validationError) Then
            MsgBox validationError, vbExclamation, "Navigation"
        ElseIf mp_WorksheetNameExists(wb, newName) Then
            MsgBox "A worksheet named '" & newName & "' already exists.", vbExclamation, "Navigation"
        Else
            Exit Do
        End If
    Loop

    On Error GoTo EH
    ws.Copy After:=wb.Worksheets(wb.Worksheets.Count)
    Set copiedWs = wb.Worksheets(wb.Worksheets.Count)
    copiedWs.Name = newName
    Exit Sub
EH:
    If Not copiedWs Is Nothing Then
        originalDisplayAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        copiedWs.Delete
        Application.DisplayAlerts = originalDisplayAlerts
    End If
    MsgBox "Failed to copy worksheet: " & Err.Description, vbCritical, "Navigation"
End Sub

Private Function mp_WorksheetNameExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    Dim targetName As String

    targetName = LCase$(Trim$(sheetName))
    For Each ws In wb.Worksheets
        If LCase$(ws.Name) = targetName Then
            mp_WorksheetNameExists = True
            Exit Function
        End If
    Next ws
End Function

Private Function mp_IsValidWorksheetName(ByVal sheetName As String, ByRef outError As String) As Boolean
    Dim invalidChars As Variant
    Dim i As Long
    Dim value As String

    value = Trim$(sheetName)
    If Len(value) = 0 Then
        outError = "Worksheet name cannot be empty."
        Exit Function
    End If

    If Len(value) > 31 Then
        outError = "Worksheet name cannot be longer than 31 characters."
        Exit Function
    End If

    invalidChars = Array(":", "\", "/", "?", "*", "[", "]")
    For i = LBound(invalidChars) To UBound(invalidChars)
        If InStr(1, value, CStr(invalidChars(i)), vbBinaryCompare) > 0 Then
            outError = "Worksheet name contains an invalid character: " & CStr(invalidChars(i))
            Exit Function
        End If
    Next i

    mp_IsValidWorksheetName = True
End Function

Private Function mp_BuildUniqueSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim candidate As String
    Dim suffix As String
    Dim i As Long
    Dim maxBaseLength As Long

    candidate = Trim$(baseName)
    If Len(candidate) = 0 Then candidate = "Sheet_copy"
    If Len(candidate) > 31 Then candidate = Left$(candidate, 31)

    If Not mp_WorksheetNameExists(wb, candidate) Then
        mp_BuildUniqueSheetName = candidate
        Exit Function
    End If

    i = 2
    Do
        suffix = " (" & CStr(i) & ")"
        maxBaseLength = 31 - Len(suffix)
        If maxBaseLength < 1 Then maxBaseLength = 1
        candidate = Left$(Trim$(baseName), maxBaseLength) & suffix
        If Not mp_WorksheetNameExists(wb, candidate) Then
            mp_BuildUniqueSheetName = candidate
            Exit Function
        End If
        i = i + 1
    Loop
End Function
