Option Explicit

Private m_runtimeReleased As Boolean

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Sub Workbook_Open()
    private_Initialize
End Sub

Private Sub Workbook_Activate()
    If m_runtimeReleased Then
        private_Initialize
    Else
        ex_AppCommands.fn_RegisterHotkeys
        ex_AppHotkeys.fn_Activate
    End If
End Sub

Private Sub Workbook_Deactivate()
    If m_runtimeReleased Then Exit Sub
    ex_AppHotkeys.fn_Deactivate
End Sub

Private Sub Workbook_SheetChange( _
    ByVal sheet As Object, _
    ByVal target As Range _
)
    Dim runtimeContext As Object

    If m_runtimeReleased Then
        private_Initialize
        Exit Sub
    End If
    On Error GoTo EH
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    ex_UiPageManager.fn_HandleCellChange target
Cleanup:
    If Not runtimeContext Is Nothing Then ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH:
    private_ReportError "Workbook_SheetChange"
    Resume Cleanup
End Sub

Private Sub Workbook_SheetSelectionChange( _
    ByVal sheet As Object, _
    ByVal target As Range _
)
    Dim runtimeContext As Object

    If m_runtimeReleased Then
        private_Initialize
        Exit Sub
    End If
    On Error GoTo EH
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    If TypeOf sheet Is Worksheet Then ex_UiBindings.fn_HandleSelection target
Cleanup:
    If Not runtimeContext Is Nothing Then ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH:
    private_ReportError "Workbook_SheetSelectionChange"
    Resume Cleanup
End Sub

Private Sub Workbook_SheetBeforeDoubleClick( _
    ByVal sheet As Object, _
    ByVal target As Range, _
    Cancel As Boolean _
)
    Dim runtimeContext As Object

    If m_runtimeReleased Then
        private_Initialize
        Exit Sub
    End If
    On Error GoTo EH
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    If TypeOf sheet Is Worksheet Then Cancel = ex_UiBindings.fn_HandleSelection(target)
Cleanup:
    If Not runtimeContext Is Nothing Then ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH:
    private_ReportError "Workbook_SheetBeforeDoubleClick"
    Resume Cleanup
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim runtimeContext As Object
    Dim updater As Workbook
    Dim errorDescription As String

    If Cancel Or m_runtimeReleased Then Exit Sub
    On Error GoTo Failed
    Set runtimeContext = ex_RuntimeLifecycle.fn_Context()
    If runtimeContext("Phase") = "Requested" Then
        Application.Run "'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_CancelPending", ThisWorkbook
    End If

    ' Preserve runtime state if closing is cancelled.
    If Not ex_Core.fn_Diagnostic_Flush() Then
        errorDescription = ex_TravelDocumentsBuilder.fn_Config("text.CloseFlushFailed")
        GoTo CancelClose
    End If
    ex_AppHotkeys.fn_Deactivate
    If Not ex_Core.fn_Diagnostic_Flush() Then
        errorDescription = ex_TravelDocumentsBuilder.fn_Config("text.CloseFlushFailed")
        GoTo CancelClose
    End If
    ex_TDB_WordSession.fn_Module_Dispose
    ex_AppCommands.fn_Module_Dispose
    On Error Resume Next
    Set updater = Application.Workbooks("WorkbookUpdater.xlam")
    On Error GoTo Failed
    If Not updater Is Nothing Then
        Application.Run "'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_ForgetContext", ThisWorkbook
    End If
    m_runtimeReleased = True
    Exit Sub
Failed:
    errorDescription = VBA.Err.Description
CancelClose:
    Cancel = True
    On Error Resume Next
    ex_AppCommands.fn_RegisterHotkeys
    ex_AppHotkeys.fn_Activate
    ex_WindowsUi.fn_ShowMessage errorDescription, VBA.vbExclamation
    On Error GoTo 0
End Sub

Private Sub private_Initialize()
    Dim runtimeContext As Object

    On Error GoTo EH
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    m_runtimeReleased = False
    ex_AppCommands.fn_RegisterHotkeys
    ex_AppHotkeys.fn_Activate
    If Not ex_TravelDocumentsBuilder.fn_Initialize Then _
        VBA.Err.Raise VBA.vbObjectError + 2231, "TravelDocumentsBuilder", "Page initialization failed."
Cleanup:
    If Not runtimeContext Is Nothing Then ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH:
    private_ReportError "Workbook_Open.Initialize"
    Resume Cleanup
End Sub

Private Sub private_ReportError(ByVal stageName As String)
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    errorNumber = VBA.Err.Number
    errorSource = VBA.Err.Source
    errorDescription = VBA.Err.Description
    ex_Core.fn_Diagnostic_WriteError stageName, errorNumber, errorSource, errorDescription
    ex_UiBusy.fn_Reset
    ex_WindowsUi.fn_ShowMessage errorDescription, VBA.vbExclamation, "TravelDocumentsBuilder"
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------