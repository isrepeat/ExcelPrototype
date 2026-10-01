Option Explicit

Private Sub Workbook_Open()
    ex_PersonalEventBuilder.fn_Initialize
End Sub

Private Sub Workbook_SheetChange(ByVal sheet As Object, ByVal target As Range)
    ex_UiPageManager.fn_HandleCellChange target
End Sub