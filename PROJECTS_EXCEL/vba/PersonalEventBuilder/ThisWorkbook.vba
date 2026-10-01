Option Explicit

Private Sub Workbook_Open()
    Dim runtimeContext As Object
    Dim runtimeEntered As Boolean
    Dim startedAt As Double
    Dim errorNumber As Long
    Dim errorDescription As String

    startedAt = VBA.Timer
    On Error GoTo EH
    runtimeEntered = ex_RuntimeLifecycle.fn_TryEnter(runtimeContext)
    If Not runtimeEntered Then Exit Sub
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_OPEN_STARTED | Workbook=" & _
        ThisWorkbook.Name & " | ExcelVersion=" & Application.Version & _
        " | Workbooks=" & VBA.CStr(Application.Workbooks.Count)
    ex_Core.fn_Diagnostic_Flush
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_OPEN_STAGE | Name=ActivateHotkeys"
    ex_Core.fn_Diagnostic_Flush
    ex_AppHotkeys.fn_Activate
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_OPEN_STAGE_COMPLETED | Name=ActivateHotkeys"
    ex_Core.fn_Diagnostic_Flush
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_OPEN_STAGE | Name=InitializeApplication"
    ex_Core.fn_Diagnostic_Flush
    ex_PersonalEventBuilder.fn_Initialize
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_OPEN_STAGE_COMPLETED | Name=InitializeApplication"
    ex_Core.fn_Diagnostic_Flush
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_OPEN_COMPLETED | ElapsedMs=" & _
        private_FormatElapsedMilliseconds(startedAt)
    ex_Core.fn_Diagnostic_Flush
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    If runtimeEntered Then ex_RuntimeLifecycle.fn_Leave runtimeContext
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_OPEN_ERROR | Number=" & _
        VBA.CStr(errorNumber) & " | Description=" & errorDescription
    ex_Core.fn_Diagnostic_Flush
    Err.Raise errorNumber, "PersonalEventBuilder.Workbook_Open", _
        errorDescription
End Sub

Private Sub Workbook_Activate()
    Dim runtimeContext As Object
    Dim runtimeEntered As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo EH
    runtimeEntered = ex_RuntimeLifecycle.fn_TryEnter(runtimeContext)
    If Not runtimeEntered Then Exit Sub
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_ACTIVATE_STARTED | Workbook=" & _
        ThisWorkbook.Name
    ex_AppHotkeys.fn_Activate
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_ACTIVATE_COMPLETED"
    ex_Core.fn_Diagnostic_Flush
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    If runtimeEntered Then ex_RuntimeLifecycle.fn_Leave runtimeContext
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_ACTIVATE_ERROR | Number=" & _
        VBA.CStr(errorNumber) & " | Description=" & errorDescription
    ex_Core.fn_Diagnostic_Flush
    Err.Raise errorNumber, "PersonalEventBuilder.Workbook_Activate", _
        errorDescription
End Sub

Private Sub Workbook_Deactivate()
    Dim runtimeContext As Object
    Dim runtimeEntered As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo EH
    runtimeEntered = ex_RuntimeLifecycle.fn_TryEnter(runtimeContext)
    If Not runtimeEntered Then Exit Sub
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_DEACTIVATE_STARTED | Workbook=" & _
        ThisWorkbook.Name
    ex_AppHotkeys.fn_Deactivate
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_DEACTIVATE_COMPLETED"
    ex_Core.fn_Diagnostic_Flush
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    If runtimeEntered Then ex_RuntimeLifecycle.fn_Leave runtimeContext
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_DEACTIVATE_ERROR | Number=" & _
        VBA.CStr(errorNumber) & " | Description=" & errorDescription
    ex_Core.fn_Diagnostic_Flush
    Err.Raise errorNumber, "PersonalEventBuilder.Workbook_Deactivate", _
        errorDescription
End Sub

Private Sub Workbook_SheetChange(ByVal sheet As Object, ByVal target As Range)
    Dim runtimeContext As Object
    Dim runtimeEntered As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo EH
    runtimeEntered = ex_RuntimeLifecycle.fn_TryEnter(runtimeContext)
    If Not runtimeEntered Then Exit Sub
    ex_UiPageManager.fn_HandleCellChange target
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    If runtimeEntered Then ex_RuntimeLifecycle.fn_Leave runtimeContext
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_SHEET_CHANGE_ERROR | Sheet=" & _
        sheet.Name & " | Range=" & target.Address(False, False) & _
        " | Number=" & VBA.CStr(errorNumber) & _
        " | Description=" & errorDescription
    ex_Core.fn_Diagnostic_Flush
    Err.Raise errorNumber, "PersonalEventBuilder.Workbook_SheetChange", _
        errorDescription
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal sheet As Object, ByVal target As Range)
    Dim runtimeContext As Object
    Dim runtimeEntered As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo EH
    runtimeEntered = ex_RuntimeLifecycle.fn_TryEnter(runtimeContext)
    If Not runtimeEntered Then Exit Sub
    ex_UiBindings.fn_CollapseSelectControls
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    If runtimeEntered Then ex_RuntimeLifecycle.fn_Leave runtimeContext
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_SELECTION_CHANGE_ERROR | Sheet=" & _
        sheet.Name & " | Range=" & target.Address(False, False) & _
        " | Number=" & VBA.CStr(errorNumber) & _
        " | Description=" & errorDescription
    ex_Core.fn_Diagnostic_Flush
    Err.Raise errorNumber, _
        "PersonalEventBuilder.Workbook_SheetSelectionChange", _
        errorDescription
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim runtimeContext As Object
    Dim startedAt As Double
    Dim flushSucceeded As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    startedAt = VBA.Timer
    On Error GoTo EH
    Set runtimeContext = ex_RuntimeLifecycle.fn_Context()
    If runtimeContext("Phase") = "Requested" Then
        Application.Run "'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_CancelPending", ThisWorkbook
    End If
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_BEFORE_CLOSE_STARTED | Workbook=" & _
        ThisWorkbook.Name
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_CLOSE_STAGE | Name=FlushExistingLogBuffer"
    flushSucceeded = ex_Core.fn_Diagnostic_Flush()
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_CLOSE_STAGE_COMPLETED | Name=FlushExistingLogBuffer" & _
        " | Succeeded=" & VBA.CStr(flushSucceeded)
    If Not flushSucceeded Then GoTo CancelClose

    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_CLOSE_STAGE | Name=DeactivateHotkeys"
    ex_AppHotkeys.fn_Deactivate
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_CLOSE_STAGE_COMPLETED | Name=DeactivateHotkeys"
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_BEFORE_CLOSE_COMPLETED | ElapsedMs=" & _
        private_FormatElapsedMilliseconds(startedAt)
    flushSucceeded = ex_Core.fn_Diagnostic_Flush()
    If flushSucceeded Then
        private_ForgetReloadContext
        Exit Sub
    End If

    ex_AppHotkeys.fn_Activate
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_CLOSE_CANCELLED | Reason=FinalLogFlushFailed"
    ex_Core.fn_Diagnostic_Flush
    Cancel = True
    VBA.MsgBox "The diagnostic log buffer could not be written. The workbook will remain open.", _
        VBA.vbExclamation, "PersonalEventBuilder"
    Exit Sub

CancelClose:
    Cancel = True
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_CLOSE_CANCELLED | Reason=InitialLogFlushFailed"
    ex_Core.fn_Diagnostic_Flush
    VBA.MsgBox "The diagnostic log buffer could not be written. The workbook will remain open.", _
        VBA.vbExclamation, "PersonalEventBuilder"
    Exit Sub
EH:
    Cancel = True
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    On Error Resume Next
    ex_AppHotkeys.fn_Activate
    ex_Core.fn_Diagnostic_WriteLog "WORKBOOK_BEFORE_CLOSE_ERROR | Number=" & _
        VBA.CStr(errorNumber) & " | Description=" & errorDescription
    ex_Core.fn_Diagnostic_Flush
    VBA.MsgBox "Workbook close diagnostics failed. The workbook will remain open.", _
        VBA.vbExclamation, "PersonalEventBuilder"
End Sub

Private Sub private_ForgetReloadContext()
    Dim updater As Workbook
    On Error Resume Next
    Set updater = Application.Workbooks("WorkbookUpdater.xlam")
    On Error GoTo 0
    If Not updater Is Nothing Then
        Application.Run "'WorkbookUpdater.xlam'!ex_WorkbookUpdater.fn_ForgetContext", ThisWorkbook
    End If
End Sub

Private Function private_FormatElapsedMilliseconds(ByVal startedAt As Double) As String
    Dim elapsedSeconds As Double

    elapsedSeconds = VBA.Timer - startedAt
    If elapsedSeconds < 0 Then elapsedSeconds = elapsedSeconds + 86400#
    private_FormatElapsedMilliseconds = VBA.Format$(elapsedSeconds * 1000#, "0.0")
End Function