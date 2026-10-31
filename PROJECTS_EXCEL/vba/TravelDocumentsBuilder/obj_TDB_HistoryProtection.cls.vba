VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_TDB_HistoryProtection"
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private Const SHEET_PASSWORD As String = "TDB.Password"

Private m_sheets As Collection
Private m_failureTemplate As String

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize( _
    ByVal tickets As ListObject, _
    ByVal revisions As ListObject _
) As Boolean
    If m_isInitialized Or m_isDisposed Then Exit Function
    If tickets Is Nothing Or revisions Is Nothing Then Exit Function
    m_failureTemplate = ex_TDB_History.fn_Config("text.ProtectionFailed")
    Set m_sheets = New Collection
    m_sheets.Add tickets.Parent
    If Not tickets.Parent Is revisions.Parent Then m_sheets.Add revisions.Parent
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    Dim diagnostic As String

    If m_isDisposed Then Exit Sub
    If m_isInitialized Then
        If Not Me.TryProtect(diagnostic) Then ex_Core.fn_Diagnostic_WriteLog "TDB_PROTECTION_RESTORE_FAILED | " & diagnostic
    End If
    m_isInitialized = False
    m_isDisposed = True
    Set m_sheets = Nothing
End Sub

Public Sub Unprotect()
    Dim sheet As Worksheet
    Dim failures As String
    Dim diagnostic As String

    If Not m_isInitialized Or m_isDisposed Then ex_TDB_History.fn_Fail "ProtectionUnavailable"
    ' Attempt every sheet and immediately restore protection after any unlock failure.
    On Error Resume Next
    For Each sheet In m_sheets
        VBA.Err.Clear
        sheet.Unprotect Password:=SHEET_PASSWORD
        If VBA.Err.Number <> 0 Then failures = failures & sheet.Name & ": " & VBA.Err.Description & VBA.vbLf
    Next sheet
    On Error GoTo 0
    If VBA.Len(failures) > 0 Then
        If Not Me.TryProtect(diagnostic) Then failures = failures & diagnostic
        VBA.Err.Raise VBA.vbObjectError + 2470, "TicketsHistory.Unprotect", VBA.Replace$(m_failureTemplate, "{details}", failures)
    End If
End Sub

Public Function TryProtect(ByRef diagnostic As String) As Boolean
    Dim sheet As Worksheet
    Dim failure As String
    Dim failures As String

    diagnostic = VBA.vbNullString
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    ' A failure on one sheet must never skip protection of the other sheet.
    For Each sheet In m_sheets
        failure = private_ProtectSheet(sheet)
        If VBA.Len(failure) > 0 Then failures = failures & sheet.Name & ": " & failure & VBA.vbLf
    Next sheet
    If VBA.Len(failures) > 0 Then
        diagnostic = VBA.Replace$(m_failureTemplate, "{details}", failures)
        Exit Function
    End If
    TryProtect = True
End Function

' //
' // Private
' //
Private Function private_ProtectSheet(ByVal sheet As Worksheet) As String
    On Error GoTo Failed
    If Not sheet.ProtectContents Then
        sheet.Cells.Locked = True
        sheet.Protect Password:=SHEET_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
            UserInterfaceOnly:=False, AllowFiltering:=True
    End If
    sheet.EnableSelection = xlNoRestrictions
    If Not sheet.ProtectContents Then ex_TDB_History.fn_Fail "ProtectionNotApplied"
    Exit Function
Failed:
    private_ProtectSheet = VBA.Err.Description
End Function