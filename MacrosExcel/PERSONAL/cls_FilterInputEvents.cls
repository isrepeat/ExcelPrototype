Option Explicit

Public WithEvents ExcelApplication As Application

Private Sub ExcelApplication_SheetChange( _
    ByVal sheetObject As Object, _
    ByVal target As Range _
)
    module_ShortcutsHandlers.fn_HandleFilterInputCellChange sheetObject, target
End Sub

Private Sub ExcelApplication_SheetSelectionChange( _
    ByVal sheetObject As Object, _
    ByVal target As Range _
)
    module_ShortcutsHandlers.fn_HandleFilterInputSelectionChange sheetObject, target
End Sub