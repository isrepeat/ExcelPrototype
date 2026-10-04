Option Explicit

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Sub Workbook_Open()
    Dim runtimeContext As Object

    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then
        Exit Sub
    End If
    On Error GoTo Failed
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_OPEN | Workbook=" & ThisWorkbook.Name
    If Not ex_PADC.fn_Initialize() Then
        GoTo Cleanup
    End If
    ex_AppHotkeys.fn_Activate
Cleanup:
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
Failed:
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_OPEN_ERROR | Description=" & VBA.Err.Description
    ex_Core.fn_Diagnostic_Flush
    Resume Cleanup
End Sub

Private Sub Workbook_SheetChange(ByVal sheet As Object, ByVal target As Range)
    Dim runtimeContext As Object

    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then
        Exit Sub
    End If
    On Error GoTo Failed
    ex_UiPageManager.fn_HandleCellChange target
Cleanup:
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
Failed:
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_SHEET_CHANGE_ERROR | Sheet=" & _
        sheet.Name & " | Description=" & VBA.Err.Description
    Resume Cleanup
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal sheet As Object, ByVal target As Range)
    Dim runtimeContext As Object

    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then
        Exit Sub
    End If
    On Error GoTo Cleanup
    If TypeOf sheet Is Worksheet Then
        ex_UiBindings.fn_HandleSelection target
    End If
Cleanup:
    ex_RuntimeLifecycle.fn_Leave runtimeContext
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_CLOSE | Workbook=" & ThisWorkbook.Name
    ex_AppHotkeys.fn_Deactivate
    ex_Core.fn_Diagnostic_Flush
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------