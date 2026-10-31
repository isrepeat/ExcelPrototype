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
Private m_restoreRequired As Boolean

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
    m_restoreRequired = True
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
    m_restoreRequired = True
    ' Attempt every sheet and immediately restore protection after any unlock failure.
    On Error Resume Next
    For Each sheet In m_sheets
        private_Trace "Unprotect", sheet
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

Public Sub PrepareMetadataWrite()
    Dim diagnostic As String

    If Not m_isInitialized Or m_isDisposed Then ex_TDB_History.fn_Fail "ProtectionUnavailable"
    ' Metadata changes do not resize tables and can remain protected from the UI throughout.
    If Not Me.TryProtect(diagnostic) Then VBA.Err.Raise VBA.vbObjectError + 2470, "TicketsHistory.Protection", diagnostic
End Sub

Public Function TryProtect(ByRef diagnostic As String) As Boolean
    Dim sheet As Worksheet
    Dim failure As String
    Dim failures As String

    diagnostic = VBA.vbNullString
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not m_restoreRequired Then
        TryProtect = True
        Exit Function
    End If
    ' A failure on one sheet must never skip protection of the other sheet.
    For Each sheet In m_sheets
        failure = private_ProtectSheet(sheet)
        If VBA.Len(failure) > 0 Then failures = failures & sheet.Name & ": " & failure & VBA.vbLf
    Next sheet
    If VBA.Len(failures) > 0 Then
        diagnostic = VBA.Replace$(m_failureTemplate, "{details}", failures)
        Exit Function
    End If
    m_restoreRequired = False
    TryProtect = True
End Function

' //
' // Private
' //
Private Function private_ProtectSheet(ByVal sheet As Worksheet) As String
    On Error GoTo Failed
    If Not sheet.ProtectContents Then
        sheet.Cells.Locked = True
    End If
    If Not sheet.ProtectContents Or Not sheet.ProtectionMode Then
        private_Trace "Protect", sheet
        sheet.Protect Password:=SHEET_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
            UserInterfaceOnly:=True, AllowFiltering:=True
    End If
    ' Avoid resetting selection policy on an inactive workbook when it is already correct.
    If sheet.EnableSelection <> xlNoRestrictions Then
        private_Trace "EnableSelection", sheet
        sheet.EnableSelection = xlNoRestrictions
    End If
    If Not sheet.ProtectContents Or Not sheet.ProtectionMode Then ex_TDB_History.fn_Fail "ProtectionNotApplied"
    Exit Function
Failed:
    private_ProtectSheet = VBA.Err.Description
End Function

Private Sub private_Trace( _
    ByVal action As String, _
    ByVal sheet As Worksheet _
)
    ex_Core.fn_Diagnostic_WriteLog "TDB_PROTECTION_CALL | Action=" & action & " | Sheet=" & sheet.Name & _
        " | Timer=" & VBA.Format$(VBA.Timer, "0.000") & " | ScreenUpdating=" & VBA.CStr(Application.ScreenUpdating)
End Sub