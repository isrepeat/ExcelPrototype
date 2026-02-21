Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    Application.Run "ex_Startup.Startup_Open"
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error GoTo EH
    Application.Run "ex_ConfigProfilesManager.m_SaveCurrentProfileToConfig", ws_Dev
    Exit Sub
EH:
    MsgBox "Failed to save profiles config during Workbook_BeforeSave: " & Err.Description, vbExclamation
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
End Sub

Private Sub Workbook_WindowResize(ByVal Wn As Window)
End Sub

Private Sub Workbook_WindowScroll(ByVal Wn As Window)
End Sub
