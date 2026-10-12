Option Explicit

#Const ENABLE_LOGGING_VERBOSE = False

Private m_runtimeReleased As Boolean

Private Sub Workbook_Open()
    private_Initialize
End Sub

Private Sub Workbook_Activate()
    If m_runtimeReleased Then
        private_Initialize
    Else
        ex_AppHotkeys.fn_Activate
    End If
End Sub

Private Sub Workbook_Deactivate()
    If m_runtimeReleased Then Exit Sub
    ex_AppHotkeys.fn_Deactivate
End Sub

Private Sub Workbook_SheetChange(ByVal sheet As Object, ByVal target As Range)
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

Private Sub Workbook_SheetSelectionChange(ByVal sheet As Object, ByVal target As Range)
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
    Dim previousEvents As Boolean
    Dim stage As Variant
    Dim stageName As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    Dim cleanupFailed As Boolean

    If Cancel Or m_runtimeReleased Then Exit Sub
    previousEvents = Application.EnableEvents
    On Error GoTo EH_CLOSE
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "TDB_CLOSE_STARTED"
#End If
    m_runtimeReleased = True
    Application.EnableEvents = False
    For Each stage In VBA.Array("Hotkeys", "Pages", "TableData", "Bindings", _
            "ElementFactory", "ControlFactory", "Busy", "Paths", "Flush")
        stageName = VBA.CStr(stage)
#If ENABLE_LOGGING_VERBOSE Then
        ex_Core.fn_Diagnostic_WriteLog "TDB_CLOSE_STAGE_STARTED | Name=" & stageName
#End If
        Select Case stageName
            Case "Hotkeys"
                ex_AppHotkeys.fn_Module_Dispose
            Case "Pages"
                ex_TravelDocumentsBuilder.fn_Module_Dispose
            Case "TableData"
                ex_TableDataRuntime.fn_Module_Dispose
            Case "Bindings"
                ex_UiBindings.fn_Module_Dispose
            Case "ElementFactory"
                ex_UiElementFactory.fn_Module_Dispose
            Case "ControlFactory"
                ex_UiControlFactory.fn_Module_Dispose
            Case "Busy"
                ex_UiBusy.fn_Module_Dispose
            Case "Paths"
                ex_RuntimePaths.fn_Module_Dispose
            Case "Flush"
                If Not ex_Core.fn_Diagnostic_Flush() Then _
                    VBA.Err.Raise VBA.vbObjectError + 2232, "Workbook_BeforeClose", "Diagnostic log flush failed."
        End Select
#If ENABLE_LOGGING_VERBOSE Then
        ex_Core.fn_Diagnostic_WriteLog "TDB_CLOSE_STAGE_COMPLETED | Name=" & stageName
#End If
NextStage:
    Next stage
    stageName = VBA.vbNullString
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "TDB_CLOSE_COMPLETED | Success=" & VBA.CStr(Not cleanupFailed)
#End If
    If Not ex_Core.fn_Diagnostic_Flush() Then _
        ex_Core.fn_Diagnostic_WriteLog "TDB_CLOSE_FINAL_FLUSH_FAILED"
CleanClose:
    Application.EnableEvents = previousEvents
    Exit Sub
EH_CLOSE:
    cleanupFailed = True
    errorNumber = VBA.Err.Number
    errorSource = VBA.Err.Source
    errorDescription = VBA.Err.Description
    ex_Core.fn_Diagnostic_WriteError "Workbook_BeforeClose." & stageName, _
        errorNumber, errorSource, errorDescription
    If VBA.Len(stageName) > 0 Then Resume NextStage
    Resume CleanClose
End Sub

Private Sub private_Initialize()
    Dim runtimeContext As Object

    On Error GoTo EH
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    m_runtimeReleased = False
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