Option Explicit

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Sub Workbook_Open()
    Dim runtimeContext As Object

    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    On Error GoTo Failed
    If Not ex_PRM.fn_Initialize() Then GoTo Cleanup
Cleanup:
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
Failed:
    ex_WindowsUi.fn_ShowMessage Err.Description, vbExclamation
    Resume Cleanup
End Sub

Private Sub Workbook_Activate()
    Dim runtimeContext As Object

    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    On Error GoTo Failed
    ex_AppCommands.fn_RegisterHotkeys
    ex_AppHotkeys.fn_Activate
Cleanup:
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
Failed:
    ex_WindowsUi.fn_ShowMessage Err.Description, vbExclamation
    Resume Cleanup
End Sub

Private Sub Workbook_Deactivate()
    ex_AppHotkeys.fn_Deactivate
End Sub

Private Sub Workbook_SheetChange( _
    ByVal sheet As Object, _
    ByVal target As Range _
)
    Dim runtimeContext As Object

    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    On Error GoTo Failed
    ex_UiPageManager.fn_HandleCellChange target
Cleanup:
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
Failed:
    ex_WindowsUi.fn_ShowMessage Err.Description, vbExclamation
    Resume Cleanup
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ex_PRM.fn_Module_Dispose
    ex_Core.fn_Diagnostic_Flush
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------