Option Explicit

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    ex_TableSnapshots.fn_Module_Dispose
    ex_PEB_Testing.fn_Module_Dispose
    ex_PEB_EventRoutes.fn_Module_Dispose
    ex_UiPageManager.fn_Module_Dispose
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_CreatePage(ByVal pageId As String) As obj_IPage
    If VBA.StrComp(VBA.Trim$(pageId), "PersonalEventBuilder", VBA.vbTextCompare) <> 0 Then
        VBA.Err.Raise VBA.vbObjectError + 2211, "ex_PersonalEventBuilder.fn_CreatePage", _
            "The page is not registered in this workbook: " & pageId
    End If
    Select Case ex_PEB_Export.fn_Config("ui.pageClass")
        Case "obj_PEB_PgMain"
            Set fn_CreatePage = New obj_PEB_PgMain
        Case "obj_PEB_PgTests"
            Set fn_CreatePage = New obj_PEB_PgTests
        Case Else
            Err.Raise 5, , "Unknown configured PEB page class: " & ex_PEB_Export.fn_Config("ui.pageClass")
    End Select
End Function

Public Sub fn_RegisterHotkeys()
    Dim keySequence As String

    If Not ex_Core.fn_TryGetWorkbookConfigValue("PersonalEventBuilder::hotkeys.preview", keySequence) Then
        VBA.Err.Raise 5, , "Required configuration key not found: PersonalEventBuilder::hotkeys.preview"
    End If
    If VBA.Len(VBA.Trim$(keySequence)) = 0 Then
        VBA.Err.Raise 5, , "Empty configuration key: PersonalEventBuilder::hotkeys.preview"
    End If
    If Not ex_AppHotkeys.fn_Register(keySequence, "ex_PersonalEventBuilder.fn_PreviewWord") Then
        VBA.Err.Raise 5, , "Hotkey registration failed: " & keySequence
    End If
    If Not ex_Core.fn_TryGetWorkbookConfigValue("PersonalEventBuilder::hotkeys.saveHistory", keySequence) Then
        VBA.Err.Raise 5, , "Required configuration key not found: PersonalEventBuilder::hotkeys.saveHistory"
    End If
    If VBA.Len(VBA.Trim$(keySequence)) = 0 Then
        VBA.Err.Raise 5, , "Empty configuration key: PersonalEventBuilder::hotkeys.saveHistory"
    End If
    If Not ex_AppHotkeys.fn_Register(keySequence, "ex_PersonalEventBuilder.fn_SaveHistory") Then
        VBA.Err.Raise 5, , "Hotkey registration failed: " & keySequence
    End If
End Sub

Public Sub fn_PreviewWord()
    Dim runtimeContext As Object
    Dim uiRenderContext As obj_UiRenderContext
    Dim uiCommand As obj_UiCommand
    Dim diagnostic As String
    Dim previousScreenUpdating As Boolean
    Dim busyEntered As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    If Not Application.ActiveWorkbook Is ThisWorkbook Then Exit Sub
    If Not TypeOf Application.ActiveSheet Is Worksheet Then Exit Sub
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    previousScreenUpdating = Application.ScreenUpdating
    On Error GoTo EH_PREVIEW
    If Not ex_UiRuntime.fn_TryGetContext(Application.ActiveSheet, uiRenderContext) Then GoTo CleanPreview
    If Not ex_UiBindingRuntime.fn_TryResolveCommand( _
            "{Binding Path=Commands.PreviewWord}", uiRenderContext.BindingContext, uiCommand) Then
        GoTo CleanPreview
    End If
    busyEntered = ex_UiBusy.fn_Begin()
    Application.ScreenUpdating = False
    uiCommand.Execute
    Set uiRenderContext = Nothing
    If ex_UiRuntime.fn_TryGetContext(Application.ActiveSheet, uiRenderContext) Then
        If Not uiRenderContext.FlushLayout(diagnostic) Then
            VBA.Err.Raise 5, , diagnostic
        End If
    End If
CleanPreview:
    Application.ScreenUpdating = previousScreenUpdating
    If busyEntered Then ex_UiBusy.fn_End
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH_PREVIEW:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    Application.ScreenUpdating = previousScreenUpdating
    If busyEntered Then ex_UiBusy.fn_End
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    VBA.Err.Raise errorNumber, "ex_PersonalEventBuilder.fn_PreviewWord", errorDescription
End Sub

Public Sub fn_SaveHistory()
    Dim runtimeContext As Object
    Dim uiRenderContext As obj_UiRenderContext
    Dim uiCommand As obj_UiCommand
    Dim diagnostic As String
    Dim previousScreenUpdating As Boolean
    Dim busyEntered As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    If Not Application.ActiveWorkbook Is ThisWorkbook Then Exit Sub
    If Not TypeOf Application.ActiveSheet Is Worksheet Then Exit Sub
    If Not ex_RuntimeLifecycle.fn_TryEnter(runtimeContext) Then Exit Sub
    previousScreenUpdating = Application.ScreenUpdating
    On Error GoTo EH_SAVE
    If Not ex_UiRuntime.fn_TryGetContext(Application.ActiveSheet, uiRenderContext) Then GoTo CleanSave
    If Not ex_UiBindingRuntime.fn_TryResolveCommand( _
            "{Binding Path=Commands.SaveHistory}", uiRenderContext.BindingContext, uiCommand) Then
        GoTo CleanSave
    End If
    busyEntered = ex_UiBusy.fn_Begin()
    Application.ScreenUpdating = False
    uiCommand.Execute
    Set uiRenderContext = Nothing
    If ex_UiRuntime.fn_TryGetContext(Application.ActiveSheet, uiRenderContext) Then
        If Not uiRenderContext.FlushLayout(diagnostic) Then
            VBA.Err.Raise 5, , diagnostic
        End If
    End If
CleanSave:
    Application.ScreenUpdating = previousScreenUpdating
    If busyEntered Then ex_UiBusy.fn_End
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    Exit Sub
EH_SAVE:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    Application.ScreenUpdating = previousScreenUpdating
    If busyEntered Then ex_UiBusy.fn_End
    ex_RuntimeLifecycle.fn_Leave runtimeContext
    VBA.Err.Raise errorNumber, "ex_PersonalEventBuilder.fn_SaveHistory", errorDescription
End Sub

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
    ex_PEB_EventRoutes.fn_Module_Dispose
    If Not ex_UiPageManager.fn_ShowPage("PersonalEventBuilder", profileId) Then
        ex_Core.fn_Diagnostic_WriteLog "INITIALIZE_ABORTED | Stage=RenderInitialPage" & _
            " | Profile=" & profileId
        ex_Core.fn_Diagnostic_Flush
        Exit Function
    End If
    fn_RegisterHotkeys
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
    VBA.Err.Raise errorNumber, "ex_PersonalEventBuilder.fn_Initialize", _
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