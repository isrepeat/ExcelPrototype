Attribute VB_Name = "ex_Startup"
Option Explicit

' Startup entry point invoked from ThisWorkbook.Workbook_Open.
Public Sub Startup_Open()
    Dim stepName As String
    Dim ws As Worksheet

    On Error GoTo EH
    stepName = "initialize-styles"
    If Not ex_SheetStylesXmlProvider.m_InitializeStyles(ThisWorkbook) Then
        MsgBox "Startup initialization failed: styles were not loaded.", vbExclamation
        Exit Sub
    End If

    stepName = "prepare-dev-sheet"
    Set ws = ws_Dev
    If Not ws Is Nothing Then
        If ws.ProtectContents Then
            On Error Resume Next
            ws.Unprotect
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo EH
        End If

        If Not ws.ProtectContents Then
            stepName = "apply-dev-sheet-pipeline"
            ex_OutputFormattingPipeline.m_ApplySheetPipeline ws
        End If
    End If

    stepName = "load-ui"
    ex_UILoader.m_LoadUiFromConfig ThisWorkbook
    stepName = "restore-selection-state"
    Application.Run "ex_ConfigProfilesManager.m_RestoreSelectionState"
    stepName = "init-custom-dropdowns"
    ex_CustomDropdown.m_InitDevTestDropdown ThisWorkbook
    Exit Sub
EH:
    MsgBox "Startup initialization failed at step '" & stepName & "': [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description, vbExclamation
End Sub

Public Sub m_HelloWorld()
    MsgBox "HelloWorld macro executed successfully.", vbInformation
End Sub
