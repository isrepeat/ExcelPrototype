Option Explicit

Private Sub Workbook_Open()
    private_Initialize
End Sub

Private Sub Workbook_Activate()
    ex_AppHotkeys.fn_Activate
End Sub

Private Sub Workbook_Deactivate()
    ex_AppHotkeys.fn_Deactivate
End Sub

Private Sub Workbook_SheetChange(ByVal sheet As Object, ByVal target As Range)
    Dim runtimeContext As Object

    On Error GoTo EH
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    ex_UiPageManager.fn_HandleCellChange target
Cleanup:
    If Not runtimeContext Is Nothing Then ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH:
    ex_UiBusy.fn_Reset
    ex_WindowsUi.fn_ShowMessage VBA.Err.Description, VBA.vbExclamation, "TravelDocumentsBuilder"
    Resume Cleanup
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal sheet As Object, ByVal target As Range)
    Dim runtimeContext As Object

    On Error GoTo EH
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    If TypeOf sheet Is Worksheet Then ex_UiBindings.fn_HandleSelection target
Cleanup:
    If Not runtimeContext Is Nothing Then ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH:
    ex_UiBusy.fn_Reset
    ex_WindowsUi.fn_ShowMessage VBA.Err.Description, VBA.vbExclamation, "TravelDocumentsBuilder"
    Resume Cleanup
End Sub

Private Sub Workbook_SheetBeforeDoubleClick( _
    ByVal sheet As Object, _
    ByVal target As Range, _
    Cancel As Boolean _
)
    Dim runtimeContext As Object

    On Error GoTo EH
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    If TypeOf sheet Is Worksheet Then Cancel = ex_UiBindings.fn_HandleSelection(target)
Cleanup:
    If Not runtimeContext Is Nothing Then ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH:
    ex_UiBusy.fn_Reset
    ex_WindowsUi.fn_ShowMessage VBA.Err.Description, VBA.vbExclamation, "TravelDocumentsBuilder"
    Resume Cleanup
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ex_UiBusy.fn_Reset
    ex_AppHotkeys.fn_Deactivate
End Sub

Private Sub private_Initialize()
    Dim runtimeContext As Object

    On Error GoTo EH
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    ex_AppHotkeys.fn_Activate
    If Not ex_TravelDocumentsBuilder.fn_Initialize Then _
        VBA.Err.Raise VBA.vbObjectError + 2231, "TravelDocumentsBuilder", "Page initialization failed."
Cleanup:
    If Not runtimeContext Is Nothing Then ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH:
    ex_UiBusy.fn_Reset
    ex_WindowsUi.fn_ShowMessage VBA.Err.Description, VBA.vbExclamation, "TravelDocumentsBuilder"
    Resume Cleanup
End Sub