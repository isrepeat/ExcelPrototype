Attribute VB_Name = "ex_Startup"
Option Explicit

' Startup entry point invoked from ThisWorkbook.Workbook_Open.
Public Sub Startup_Open()
    On Error GoTo EH
    If Not ex_SheetStylesXmlProvider.m_InitializeStyles(ThisWorkbook) Then
        MsgBox "Startup initialization failed: styles were not loaded.", vbExclamation
        Exit Sub
    End If
    ex_UILoader.m_LoadUiFromConfig ThisWorkbook
    Application.Run "ex_ConfigProfilesManager.m_RestoreSelectionState"
    Exit Sub
EH:
    MsgBox "Startup initialization failed: " & Err.Description, vbExclamation
End Sub

Public Sub m_HelloWorld()
    MsgBox "HelloWorld macro executed successfully.", vbInformation
End Sub
