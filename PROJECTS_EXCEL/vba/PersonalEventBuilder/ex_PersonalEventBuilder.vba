Option Explicit

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    ex_UiPageManager.fn_Module_Dispose
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Initialize() As Boolean
    Dim runtimeContext As Object
    Dim profileId As String
    Dim startedAt As Double
    Dim errorNumber As Long
    Dim errorDescription As String

    Set runtimeContext = ex_RuntimeLifecycle.fn_Context()
    If runtimeContext("StopRequested") And runtimeContext("Phase") <> "Initializing" Then Exit Function
    startedAt = VBA.Timer
    On Error GoTo EH
    ex_Core.fn_Diagnostic_WriteLog "INITIALIZE_STARTED | Workbook=" & ThisWorkbook.Name
    ex_Core.fn_Diagnostic_Flush
    ex_Core.fn_Diagnostic_WriteLog "INITIALIZE_STAGE | Name=ReadWorkbookProfile"
    ex_Core.fn_Diagnostic_Flush
    If Not ex_Core.fn_TryGetWorkbookProfileId(profileId) Then
        ex_Core.fn_Diagnostic_WriteLog "INITIALIZE_ABORTED | Stage=ReadWorkbookProfile"
        ex_Core.fn_Diagnostic_Flush
        Exit Function
    End If
    ex_Core.fn_Diagnostic_WriteLog "INITIALIZE_STAGE_COMPLETED | Name=ReadWorkbookProfile" & _
        " | Profile=" & profileId
    ex_Core.fn_Diagnostic_WriteLog "INITIALIZE_STAGE | Name=RenderInitialPage"
    ex_Core.fn_Diagnostic_Flush
    If Not ex_UiPageManager.fn_ShowPage("PersonalEventBuilder", profileId) Then
        ex_Core.fn_Diagnostic_WriteLog "INITIALIZE_ABORTED | Stage=RenderInitialPage" & _
            " | Profile=" & profileId
        ex_Core.fn_Diagnostic_Flush
        Exit Function
    End If
    ex_Core.fn_Diagnostic_WriteLog "INITIALIZE_STAGE_COMPLETED | Name=RenderInitialPage"
    ex_Core.fn_Diagnostic_WriteLog "INITIALIZE_COMPLETED | Workbook=" & _
        ThisWorkbook.Name & " | ElapsedMs=" & _
        private_FormatElapsedMilliseconds(startedAt)
    ex_Core.fn_Diagnostic_Flush
    fn_Initialize = True
    Exit Function
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    ex_Core.fn_Diagnostic_WriteLog "INITIALIZE_ERROR | Number=" & _
        VBA.CStr(errorNumber) & " | Description=" & errorDescription
    ex_Core.fn_Diagnostic_Flush
    Err.Raise errorNumber, "ex_PersonalEventBuilder.fn_Initialize", _
        errorDescription
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_FormatElapsedMilliseconds(ByVal startedAt As Double) As String
    Dim elapsedSeconds As Double

    elapsedSeconds = VBA.Timer - startedAt
    If elapsedSeconds < 0 Then elapsedSeconds = elapsedSeconds + 86400#
    private_FormatElapsedMilliseconds = VBA.Format$(elapsedSeconds * 1000#, "0.0")
End Function