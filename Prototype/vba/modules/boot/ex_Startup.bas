Attribute VB_Name = "ex_Startup"
Option Explicit

' Startup entry point invoked from ThisWorkbook.Workbook_Open.
Public Sub Startup_Open()
    On Error GoTo EH
    If Not ex_SheetStylesXmlProvider.m_InitializeStyles(ThisWorkbook) Then
        MsgBox "Startup initialization failed: styles were not loaded.", vbExclamation
        Exit Sub
    End If
    ex_OutputFormattingPipeline.m_ApplySheetPipeline ws_Dev
    ex_UILoader.m_LoadUiFromConfig ThisWorkbook
    Application.Run "ex_ConfigProfilesManager.m_RestoreSelectionState"
    ex_CustomDropdown.m_InitDevTestDropdown ThisWorkbook
    Exit Sub
EH:
    MsgBox "Startup initialization failed: [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description, vbExclamation
End Sub

Public Sub m_HelloWorld()
    MsgBox "HelloWorld macro executed successfully.", vbInformation
End Sub
