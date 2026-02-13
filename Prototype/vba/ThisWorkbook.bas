Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    Application.Run "ex_Bootstrap.Bootstrap_Open"
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error GoTo EH
    Application.Run "ex_Presets.m_SaveCurrentProfileToConfig", ws_Dev
    Exit Sub
EH:
    MsgBox "Failed to save profiles config during Workbook_BeforeSave: " & Err.Description, vbExclamation
End Sub
