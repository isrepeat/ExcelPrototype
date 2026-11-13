Option Explicit

Private m_dispatching As Boolean
Private m_afterDispatch As String

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_OnShapeClick()
    private_DispatchShapeClick True
End Sub

Public Sub fn_OnShapeClickWithoutBusy()
    private_DispatchShapeClick False
End Sub

Public Function fn_TryAfterShapeDispatch(ByVal callback As String) As Boolean
    If Not m_dispatching Then Exit Function
    If Len(m_afterDispatch) > 0 And m_afterDispatch <> callback Then Exit Function
    m_afterDispatch = callback
    fn_TryAfterShapeDispatch = True
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Sub private_DispatchShapeClick(ByVal showBusy As Boolean)
    Dim runtimeContext As Object
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim context As obj_UiRenderContext
    Dim targetWorksheet As Worksheet
    Dim shapeName As String
    Dim diagnostic As String
    Dim previousScreenUpdating As Boolean
    Dim busyEntered As Boolean
    Dim layoutFlushed As Boolean
    Dim startedAt As Double
    Dim stageStartedAt As Double
    Dim previousDispatching As Boolean
    Dim previousCallback As String
    Dim callback As String

    startedAt = VBA.Timer
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    previousScreenUpdating = Application.ScreenUpdating
    previousDispatching = m_dispatching
    previousCallback = m_afterDispatch
    On Error GoTo EH
    If showBusy Then busyEntered = ex_UiBusy.fn_Begin()
    ' Keep long-running commands and their progress visible.
    shapeName = VBA.CStr(Application.Caller)
    If TypeOf Application.ActiveSheet Is Worksheet Then
        Set targetWorksheet = Application.ActiveSheet
        If ex_UiRuntime.fn_TryGetContext(targetWorksheet, context) Then
            stageStartedAt = VBA.Timer
            m_dispatching = True
            m_afterDispatch = VBA.vbNullString
            context.Router.DispatchShape shapeName
            callback = m_afterDispatch
            m_dispatching = False
            m_afterDispatch = VBA.vbNullString
            If Len(callback) > 0 Then Application.Run callback
            ex_Core.fn_Diagnostic_WritePerf "Ui.Click.Dispatch | Shape=" & shapeName, stageStartedAt
            ' The command can replace and dispose the page context during rendering.
            Set context = Nothing
            If ex_UiRuntime.fn_TryGetContext(targetWorksheet, context) Then
                stageStartedAt = VBA.Timer
                ' Suppress intermediate layout changes only during the flush.
                Application.ScreenUpdating = False
                ex_UiBusy.fn_Refresh
                If Not context.FlushLayout(diagnostic) Then
                    VBA.Err.Raise VBA.vbObjectError + 2231, , diagnostic
                End If
                layoutFlushed = True
                ex_Core.fn_Diagnostic_WritePerf "Ui.Click.Flush | Shape=" & shapeName, stageStartedAt
            End If
        End If
    End If
CleanExit:
    m_dispatching = previousDispatching
    m_afterDispatch = previousCallback
    Application.ScreenUpdating = previousScreenUpdating
    ex_UiBusy.fn_Refresh
    Set context = Nothing
    Set targetWorksheet = Nothing
    ex_Core.fn_Diagnostic_WritePerf "Ui.Click.Total | Shape=" & shapeName, startedAt
    If busyEntered Then ex_UiBusy.fn_End layoutFlushed
    busyEntered = False
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    m_dispatching = previousDispatching
    m_afterDispatch = previousCallback
    Application.ScreenUpdating = previousScreenUpdating
    If busyEntered Then ex_UiBusy.fn_End
    Set context = Nothing
    Set targetWorksheet = Nothing
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    VBA.Err.Raise errorNumber, "ex_UiBridge.fn_OnShapeClick", errorDescription
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------