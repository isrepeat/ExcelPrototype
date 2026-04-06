Attribute VB_Name = "ex_PostProcessPipeline"
Option Explicit

Private Const DEBUG_LOG_PATH As String = "Logs\personalcard_pipeline.log"
Private Const DEBUG_LOG_ENABLED As Boolean = True

Public Function m_Run( _
    ByVal ws As Worksheet, _
    ByVal cfg As Object, _
    ByVal resultTables As Collection, _
    Optional ByVal inputObject As Object = Nothing, _
    Optional ByVal scriptConfigKey As String = "PostProcess.Script.Implicit", _
    Optional ByVal requireScript As Boolean = False, _
    Optional ByVal layoutBatchApplyRefresh As Boolean = True _
) As Boolean
    Dim scriptKey As String
    Dim scriptText As String
    Dim scriptLoadError As String
    Dim stageName As String
    Dim runStartStamp As Double
    Dim stageStartStamp As Double

    On Error GoTo EH
    runStartStamp = Timer
    stageName = "validate-input"
    ' mp_DebugLog "RUN START"

    If ws Is Nothing Then
        Err.Raise vbObjectError + 6220, "ex_PostProcessPipeline", "Worksheet is required for post-process execution."
    End If
    If cfg Is Nothing Then
        Err.Raise vbObjectError + 6221, "ex_PostProcessPipeline", "Config object is required for post-process execution."
    End If
    If resultTables Is Nothing Then
        Err.Raise vbObjectError + 6222, "ex_PostProcessPipeline", "ResultTables are required for post-process execution."
    End If

    stageName = "resolve-script-key"
    scriptKey = Trim$(CStr(scriptConfigKey))
    If Len(scriptKey) = 0 Then scriptKey = "PostProcess.Script.Implicit"
    ' mp_DebugLog "scriptKey='" & scriptKey & "' requireScript=" & LCase$(CStr(requireScript))

    stageName = "load-script"
    stageStartStamp = Timer
    If Not ex_ScriptSourceLoader.m_TryGetScriptText(cfg, scriptKey, scriptText, scriptLoadError) Then
        Err.Raise vbObjectError + 6223, "ex_PostProcessPipeline", scriptLoadError
    End If
    mp_DebugLog "stage-done stage='load-script' duration=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(stageStartStamp))

    If Len(scriptText) = 0 Then
        stageName = "skip-no-script"
        If requireScript Then
            Err.Raise vbObjectError + 6224, "ex_PostProcessPipeline", "Missing required post-process script for key '" & scriptKey & "'."
        End If
        mp_DebugLog "RUN DURATION total=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(runStartStamp))
        ' mp_DebugLog "RUN END SKIP scriptKey='" & scriptKey & "'"
        Exit Function
    End If

    stageName = "run-postprocess-script"
    stageStartStamp = Timer
    ex_ScriptIO.m_SetInput inputObject
    ex_ScriptDSL.m_ApplyScriptToSheet ws, cfg, resultTables, scriptKey, layoutBatchApplyRefresh
    mp_DebugLog "stage-done stage='run-postprocess-script' duration=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(stageStartStamp))
    m_Run = True
    mp_DebugLog "RUN DURATION total=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(runStartStamp))
    ' mp_DebugLog "RUN END OK scriptKey='" & scriptKey & "' sheet='" & ws.Name & "'"
    Exit Function

EH:
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    mp_DebugLog "FAIL stage='" & stageName & "' scriptKey='" & scriptKey & "' err=[" & errSource & " #" & CStr(errNumber) & "] " & errDescription & " | elapsed=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(runStartStamp))
    If errNumber = 0 Then errNumber = vbObjectError + 6225
    If Len(errSource) = 0 Then errSource = "ex_PostProcessPipeline"
    If Len(errDescription) = 0 Then errDescription = "Unknown post-process pipeline failure."
    Err.Raise errNumber, errSource, errDescription
End Function

Private Function mp_StageElapsedSeconds(ByVal startStamp As Double) As Double
    mp_StageElapsedSeconds = Timer - startStamp
    If mp_StageElapsedSeconds < 0# Then mp_StageElapsedSeconds = mp_StageElapsedSeconds + 86400#
End Function

Private Function mp_FormatElapsedSeconds(ByVal elapsedSeconds As Double) As String
    mp_FormatElapsedSeconds = Format$(elapsedSeconds, "0.000") & "s"
End Function

Private Sub mp_DebugLog(ByVal messageText As String)
    If Not DEBUG_LOG_ENABLED Then Exit Sub
    On Error Resume Next
    ex_Messaging.m_LogToFile "[ex_PostProcessPipeline] " & CStr(messageText), DEBUG_LOG_PATH
    On Error GoTo 0
End Sub
