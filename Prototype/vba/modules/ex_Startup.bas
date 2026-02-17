Attribute VB_Name = "ex_Startup"
Option Explicit

' Startup entry point invoked from ThisWorkbook.Workbook_Open.
Public Sub Startup_Open()
    On Error GoTo EH
    ex_UILoader.m_LoadUiFromConfig ThisWorkbook
    Application.Run "ex_ConfigProfilesManager.m_OnModeChanged"
    Exit Sub
EH:
    MsgBox "Startup initialization failed: " & Err.Description, vbExclamation
End Sub

Public Sub m_HelloWorld()
    MsgBox "HelloWorld macro executed successfully.", vbInformation
End Sub
