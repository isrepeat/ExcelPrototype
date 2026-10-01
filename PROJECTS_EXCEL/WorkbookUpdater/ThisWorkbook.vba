Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Cancel = Not ex_WorkbookUpdater.fn_CanUnload()
    If Cancel Then Application.StatusBar = "Workbook updater has a pending operation. Cancel it or wait for completion before unloading."
End Sub