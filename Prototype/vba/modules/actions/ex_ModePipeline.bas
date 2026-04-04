Attribute VB_Name = "ex_ModePipeline"
Option Explicit

Private Const CONTEXT_FIELD_PREPROCESS_OUTPUT As String = "PreProcessOutput"
Private Const CONTEXT_FIELD_PREPROCESS_CONTEXT As String = "PreProcessContext"
Private Const CONTEXT_FIELD_MODE_OUTPUT As String = "ModeOutput"
Private Const CONTEXT_FIELD_RESULTLAYOUT_EXECUTED As String = "ResultLayoutExecuted"
Private Const CONTEXT_FIELD_POST_EXECUTED As String = "PostExecuted"

Private Const MODE_RESULT_FIELD_OUTPUT As String = "Output"
Private Const MODE_RESULT_FIELD_WORKSHEET As String = "Worksheet"
Private Const MODE_RESULT_FIELD_RESULT_TABLES As String = "ResultTables"
Private Const INPUT_KEY_LAYOUT_ROWKINDS As String = "__ResultLayoutRowKinds"
Private Const INPUT_KEY_LAYOUT_FIELDRANGES As String = "__ResultLayoutFieldRanges"

Private Const AUTO_POSTPROCESS_SCRIPT_KEY As String = "PostProcess.Script.Implicit"
Private Const DEBUG_LOG_PATH As String = "Logs\personalcard_pipeline.log"
Private Const DEBUG_LOG_ENABLED As Boolean = True

Private g_PipelineBusyDepth As Long
Private g_PipelinePrevCursor As XlMousePointer
Private g_PipelinePrevScreenUpdating As Boolean

Public Function m_RunModePipeline( _
    ByVal cfg As Object, _
    ByVal modeExecutorMacro As String, _
    Optional ByVal pipelineInput As Object = Nothing, _
    Optional ByVal requirePreScript As Boolean = False _
) As Object
    Dim ctx As Object
    Dim stageName As String
    Dim runStartStamp As Double
    Dim stageStartStamp As Double
    Dim preProcessContext As Object
    Dim modeInput As Object
    Dim modeResult As Object
    Dim modeSheet As Worksheet
    Dim modeTables As Collection
    Dim modeOutput As Object
    Dim resultLayoutExecuted As Boolean
    Dim layoutRefreshError As String
    Dim busyScopeActive As Boolean

    On Error GoTo EH

    runStartStamp = Timer
    stageName = "init"
    mp_BeginPipelineBusy
    busyScopeActive = True

    mp_ResetDebugLog
    Set ctx = CreateObject("Scripting.Dictionary")
    ctx.CompareMode = 1
    ctx(CONTEXT_FIELD_POST_EXECUTED) = "false"
    ' mp_DebugLog "RUN START modeExecutor='" & CStr(modeExecutorMacro) & "'"

    stageName = "prepare-input"
    If pipelineInput Is Nothing Then
        Set pipelineInput = New obj_ScriptIOPayload
        ' mp_DebugLog "pipelineInput auto-created"
    End If

    stageName = "run-preprocess"
    stageStartStamp = Timer
    Set preProcessContext = ex_PreProcessPipeline.m_Run(cfg, pipelineInput, requirePreScript)
    mp_DebugLog "stage-done stage='run-preprocess' duration=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(stageStartStamp))
    Set modeInput = pipelineInput
    If Not preProcessContext Is Nothing Then
        If preProcessContext.Exists("Output") Then
            Set modeInput = preProcessContext("Output")
            Set ctx(CONTEXT_FIELD_PREPROCESS_OUTPUT) = modeInput
        End If
        Set ctx(CONTEXT_FIELD_PREPROCESS_CONTEXT) = preProcessContext
    End If

    stageName = "check-mode-executor"
    If Len(Trim$(modeExecutorMacro)) = 0 Then
        Set ctx(CONTEXT_FIELD_MODE_OUTPUT) = modeInput
        ' mp_DebugLog "SKIP mode executor is empty"
        Set m_RunModePipeline = ctx
        Exit Function
    End If

    stageName = "run-mode-executor"
    stageStartStamp = Timer
    Set modeResult = mp_RunModeExecutor(modeExecutorMacro, cfg, modeInput, preProcessContext)
    mp_DebugLog "stage-done stage='run-mode-executor' duration=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(stageStartStamp))
    Set ctx(CONTEXT_FIELD_MODE_OUTPUT) = modeResult

    If modeResult Is Nothing Then
        Err.Raise vbObjectError + 6113, "ex_ModePipeline", "Mode executor returned Nothing. Expected dictionary result object."
    End If

    stageName = "resolve-mode-output"
    If Not mp_TryGetModeResultObject(modeResult, MODE_RESULT_FIELD_OUTPUT, modeOutput) Then
        Err.Raise vbObjectError + 6117, "ex_ModePipeline", "Mode result must provide object field '" & MODE_RESULT_FIELD_OUTPUT & "'."
    End If

    stageName = "resolve-post-context"
    If Not mp_TryGetModeResultWorksheet(modeResult, modeSheet) Then
        Err.Raise vbObjectError + 6114, "ex_ModePipeline", "Mode result must provide Worksheet for post-process execution."
    End If
    If Not mp_TryGetModeResultTables(modeResult, modeTables) Then
        Err.Raise vbObjectError + 6115, "ex_ModePipeline", "Mode result must provide ResultTables for post-process execution."
    End If

    stageName = "run-result-layout"
    stageStartStamp = Timer
    resultLayoutExecuted = ex_ResultLayoutPipeline.m_Run( _
        cfg, _
        modeSheet, _
        modeTables, _
        modeOutput)
    mp_DebugLog "stage-done stage='run-result-layout' duration=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(stageStartStamp))
    If resultLayoutExecuted Then
        ctx(CONTEXT_FIELD_RESULTLAYOUT_EXECUTED) = "true"
    Else
        ctx(CONTEXT_FIELD_RESULTLAYOUT_EXECUTED) = "false"
    End If

    stageName = "run-postprocess"
    stageStartStamp = Timer
    If ex_PostProcessPipeline.m_Run(modeSheet, cfg, modeTables, modeOutput, AUTO_POSTPROCESS_SCRIPT_KEY, False, False) Then
        mp_DebugLog "stage-done stage='run-postprocess' duration=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(stageStartStamp))
        ctx(CONTEXT_FIELD_POST_EXECUTED) = "true"
        ' mp_DebugLog "post process executed scriptKey='" & AUTO_POSTPROCESS_SCRIPT_KEY & "'"
    Else
        mp_DebugLog "stage-done stage='run-postprocess' duration=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(stageStartStamp))
        ' mp_DebugLog "post process skipped scriptKey='" & AUTO_POSTPROCESS_SCRIPT_KEY & "'"
    End If

    If resultLayoutExecuted Then
        If LCase$(CStr(ctx(CONTEXT_FIELD_POST_EXECUTED))) = "true" Then
            stageName = "refresh-layout-after-postprocess"
            stageStartStamp = Timer
            layoutRefreshError = vbNullString
            If Not ex_ResultLayoutItemsRt.m_Refresh(modeSheet, vbNullString, layoutRefreshError) Then
                If Len(layoutRefreshError) = 0 Then layoutRefreshError = "Unknown layout refresh error after implicit post-process."
                Err.Raise vbObjectError + 6119, "ex_ModePipeline", layoutRefreshError
            End If
            mp_DebugLog "stage-done stage='refresh-layout-after-postprocess' duration=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(stageStartStamp))
        Else
            stageName = "apply-result-layout-styles"
            stageStartStamp = Timer
            mp_ApplyLayoutSheetStyling modeSheet, modeOutput
            mp_DebugLog "stage-done stage='apply-result-layout-styles' duration=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(stageStartStamp))
        End If
    End If

    stageName = "done"
    mp_DebugLog "RUN DURATION total=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(runStartStamp))
    ' mp_DebugLog "RUN END modeExecutor='" & CStr(modeExecutorMacro) & "'"
    If busyScopeActive Then
        mp_EndPipelineBusy
        busyScopeActive = False
    End If
    Set m_RunModePipeline = ctx
    Exit Function

EH:
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    mp_DebugLog "FAIL stage='" & stageName & "' err=[" & errSource & " #" & CStr(errNumber) & "] " & errDescription & " | elapsed=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(runStartStamp))
    If errNumber = 0 Then errNumber = vbObjectError + 6116
    If Len(errSource) = 0 Then errSource = "ex_ModePipeline"
    If Len(errDescription) = 0 Then errDescription = "Unknown pipeline failure."
    If busyScopeActive Then
        mp_EndPipelineBusy
        busyScopeActive = False
    End If
    Err.Raise errNumber, errSource, errDescription
End Function

Private Function mp_StageElapsedSeconds(ByVal startStamp As Double) As Double
    mp_StageElapsedSeconds = Timer - startStamp
    If mp_StageElapsedSeconds < 0# Then mp_StageElapsedSeconds = mp_StageElapsedSeconds + 86400#
End Function

Private Function mp_FormatElapsedSeconds(ByVal elapsedSeconds As Double) As String
    mp_FormatElapsedSeconds = Format$(elapsedSeconds, "0.000") & "s"
End Function

Private Sub mp_BeginPipelineBusy()
    On Error Resume Next

    If g_PipelineBusyDepth <= 0 Then
        g_PipelinePrevCursor = Application.Cursor
        g_PipelinePrevScreenUpdating = Application.ScreenUpdating
        Application.Cursor = xlWait
        Application.ScreenUpdating = False
        g_PipelineBusyDepth = 0
    End If

    g_PipelineBusyDepth = g_PipelineBusyDepth + 1
    On Error GoTo 0
End Sub

Private Sub mp_EndPipelineBusy()
    On Error Resume Next

    If g_PipelineBusyDepth <= 0 Then
        Application.ScreenUpdating = g_PipelinePrevScreenUpdating
        Application.Cursor = xlDefault
        g_PipelineBusyDepth = 0
        On Error GoTo 0
        Exit Sub
    End If

    g_PipelineBusyDepth = g_PipelineBusyDepth - 1
    If g_PipelineBusyDepth <= 0 Then
        Application.ScreenUpdating = g_PipelinePrevScreenUpdating
        Application.Cursor = g_PipelinePrevCursor
        g_PipelineBusyDepth = 0
    End If

    On Error GoTo 0
End Sub

Private Function mp_RunModeExecutor( _
    ByVal modeExecutorMacro As String, _
    ByVal cfg As Object, _
    ByVal modeInput As Object, _
    ByVal preProcessContext As Object _
) As Object
    Dim result As Object

    Set result = Application.Run(modeExecutorMacro, cfg, modeInput, preProcessContext)
    If result Is Nothing Then
        Err.Raise vbObjectError + 6118, "ex_ModePipeline", "Mode executor '" & modeExecutorMacro & "' must return object result."
    End If

    Set mp_RunModeExecutor = result
End Function

Private Function mp_TryGetModeResultObject( _
    ByVal modeResult As Object, _
    ByVal fieldName As String, _
    ByRef outObject As Object _
) As Boolean
    If modeResult Is Nothing Then Exit Function
    If Not modeResult.Exists(fieldName) Then Exit Function
    If Not IsObject(modeResult(fieldName)) Then Exit Function

    Set outObject = modeResult(fieldName)
    mp_TryGetModeResultObject = True
End Function

Private Function mp_TryGetModeResultWorksheet(ByVal modeResult As Object, ByRef outSheet As Worksheet) As Boolean
    Dim valueObject As Object

    If Not mp_TryGetModeResultObject(modeResult, MODE_RESULT_FIELD_WORKSHEET, valueObject) Then Exit Function
    If valueObject Is Nothing Then Exit Function
    If Not TypeOf valueObject Is Worksheet Then Exit Function

    Set outSheet = valueObject
    mp_TryGetModeResultWorksheet = True
End Function

Private Function mp_TryGetModeResultTables(ByVal modeResult As Object, ByRef outTables As Collection) As Boolean
    Dim valueObject As Object

    If Not mp_TryGetModeResultObject(modeResult, MODE_RESULT_FIELD_RESULT_TABLES, valueObject) Then Exit Function
    If valueObject Is Nothing Then Exit Function
    If TypeName(valueObject) <> "Collection" Then Exit Function

    Set outTables = valueObject
    mp_TryGetModeResultTables = True
End Function

Private Sub mp_DebugLog(ByVal messageText As String)
    If Not DEBUG_LOG_ENABLED Then Exit Sub
    On Error Resume Next
    ex_Messaging.m_LogToFile "[ex_ModePipeline] " & CStr(messageText), DEBUG_LOG_PATH
    On Error GoTo 0
End Sub

Private Sub mp_ApplyLayoutSheetStyling(ByVal ws As Worksheet, ByVal modeOutput As Object)
    Dim objectValue As Object
    Dim rowKindRanges As Object
    Dim resultFieldRanges As Collection

    If ws Is Nothing Then Exit Sub

    Set resultFieldRanges = Nothing
    If Not modeOutput Is Nothing Then
        Set objectValue = Nothing
        If ex_ScriptIO.m_TryGetObject(modeOutput, INPUT_KEY_LAYOUT_FIELDRANGES, objectValue) Then
            If Not objectValue Is Nothing Then
                If TypeName(objectValue) = "Collection" Then
                    Set resultFieldRanges = objectValue
                End If
            End If
        End If

        Set objectValue = Nothing
        If ex_ScriptIO.m_TryGetObject(modeOutput, INPUT_KEY_LAYOUT_ROWKINDS, objectValue) Then
            If Not objectValue Is Nothing Then
                Set rowKindRanges = objectValue
            End If
        End If
    End If

    ex_OutputFormattingPipeline.m_ApplySheetPipeline ws, resultFieldRanges, Nothing, rowKindRanges

End Sub

Private Sub mp_ResetDebugLog()
    If Not DEBUG_LOG_ENABLED Then Exit Sub
    If Not mp_IsFileLogEnabled() Then Exit Sub
    On Error Resume Next
    ex_Messaging.m_ClearLogFile DEBUG_LOG_PATH
    On Error GoTo 0
End Sub

Private Function mp_IsFileLogEnabled() As Boolean
    Dim rawValue As String
    Dim parsedValue As Boolean

    rawValue = ex_XmlCore.m_GetSettingsValue("st_FileLogEnabled", "false")
    If ex_XmlCore.m_TryParseBoolean(rawValue, parsedValue) Then
        mp_IsFileLogEnabled = parsedValue
    Else
        mp_IsFileLogEnabled = False
    End If
End Function
