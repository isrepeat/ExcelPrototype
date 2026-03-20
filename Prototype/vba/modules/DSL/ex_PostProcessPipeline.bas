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
    Optional ByVal injectedRuntimeVars As Object = Nothing, _
    Optional ByVal injectedRuntimeVarTypes As Object = Nothing, _
    Optional ByVal requireScript As Boolean = False _
) As Boolean
    Dim scriptKey As String
    Dim scriptText As String
    Dim scriptLoadError As String
    Dim stageName As String

    On Error GoTo EH
    stageName = "validate-input"
    mp_DebugLog "RUN START"

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
    mp_DebugLog "scriptKey='" & scriptKey & "' requireScript=" & LCase$(CStr(requireScript))

    stageName = "load-script"
    If Not ex_ScriptSourceLoader.m_TryGetScriptText(cfg, scriptKey, scriptText, scriptLoadError) Then
        Err.Raise vbObjectError + 6223, "ex_PostProcessPipeline", scriptLoadError
    End If

    If Len(scriptText) = 0 Then
        stageName = "skip-no-script"
        If requireScript Then
            Err.Raise vbObjectError + 6224, "ex_PostProcessPipeline", "Missing required post-process script for key '" & scriptKey & "'."
        End If
        mp_DebugLog "RUN END SKIP scriptKey='" & scriptKey & "'"
        Exit Function
    End If

    stageName = "run-postprocess-script"
    ex_ScriptIO.m_SetInput inputObject
    ex_ScriptDSL.m_ApplyScriptToSheet ws, cfg, resultTables, scriptKey, injectedRuntimeVars, injectedRuntimeVarTypes
    m_Run = True
    mp_DebugLog "RUN END OK scriptKey='" & scriptKey & "' sheet='" & ws.Name & "'"
    Exit Function

EH:
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    mp_DebugLog "FAIL stage='" & stageName & "' scriptKey='" & scriptKey & "' err=[" & errSource & " #" & CStr(errNumber) & "] " & errDescription
    If errNumber = 0 Then errNumber = vbObjectError + 6225
    If Len(errSource) = 0 Then errSource = "ex_PostProcessPipeline"
    If Len(errDescription) = 0 Then errDescription = "Unknown post-process pipeline failure."
    Err.Raise errNumber, errSource, errDescription
End Function

Private Sub mp_DebugLog(ByVal messageText As String)
    If Not DEBUG_LOG_ENABLED Then Exit Sub
    On Error Resume Next
    ex_Messaging.m_LogToFile "[ex_PostProcessPipeline] " & CStr(messageText), DEBUG_LOG_PATH
    On Error GoTo 0
End Sub
