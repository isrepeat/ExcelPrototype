Option Explicit

' --------------------------------------
' namespace WorkbookEvents {
' --------------------------------------
Private Sub Workbook_Open()
    ex_DocumentGenerationBootstrap.fn_Initialize
End Sub

' Main entry point for cell changes in this workbook.
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo EH
    ex_CellChangeRouter.fn_OnSheetChange Sh, Target
    Exit Sub
EH:
    VBA.MsgBox "Cell-change dispatch failed: [" & _
        VBA.CStr(VBA.Err.Number) & "] " & VBA.Err.Description, _
        VBA.vbExclamation, "Document Generation"
End Sub

' Main entry point for cell selection in this workbook.
Private Sub Workbook_SheetSelectionChange( _
    ByVal Sh As Object, _
    ByVal Target As Range _
)
    On Error GoTo EH
    ex_CellChangeRouter.fn_OnSheetSelectionChange Sh, Target
    Exit Sub
EH:
    VBA.MsgBox "Selection-change dispatch failed: [" & _
        VBA.CStr(VBA.Err.Number) & "] " & VBA.Err.Description, _
        VBA.vbExclamation, "Document Generation"
End Sub
' --------------------------------------
' } // namespace WorkbookEvents
' --------------------------------------