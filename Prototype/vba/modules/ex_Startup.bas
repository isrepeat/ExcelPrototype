Attribute VB_Name = "ex_Startup"
Option Explicit

' Startup entry point invoked from ThisWorkbook.Workbook_Open.
Public Sub Startup_Open()
    On Error Resume Next
    Application.Run "ex_ConfigProfilesManager.m_EnsureProfileDropdown", ws_Dev
    On Error GoTo 0
End Sub
