Option Explicit

Private Sub Workbook_Open()
    ex_PersonalEventBuilder.fn_Initialize
End Sub

Private Sub Workbook_SheetChange(ByVal sheet As Object, ByVal target As Range)
    ex_UiPageManager.fn_HandleCellChange target
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal sheet As Object, ByVal target As Range)
    ex_UiBindings.fn_CollapseSelectControls
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If ex_Core.fn_Diagnostic_Flush() Then Exit Sub
    Cancel = True
    VBA.MsgBox "The diagnostic log buffer could not be written. The workbook will remain open.", _
        VBA.vbExclamation, "PersonalEventBuilder"
End Sub