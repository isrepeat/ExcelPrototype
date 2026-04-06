Attribute VB_Name = "ex_PreProcessPipeline"
Option Explicit

Private Const PREPROCESS_SCRIPT_KEY As String = "Input.PreProcessScript"
Private Const CONTEXT_FIELD_OUTPUT As String = "Output"
Private Const CONTEXT_FIELD_HAS_SCRIPT As String = "HasScript"
Private Const DEBUG_LOG_PATH As String = "Logs\personalcard_pipeline.log"
Private Const DEBUG_LOG_ENABLED As Boolean = True

Public Function m_Run( _
    ByVal cfg As Object, _
    Optional ByVal inputObject As Object = Nothing, _
    Optional ByVal requireScript As Boolean = False _
) As Object
    Dim scriptRef As String
    Dim scriptLoadError As String
    Dim outputObject As Object
    Dim context As Object
    Dim resultTables As Collection
    Dim stageName As String
    Dim runStartStamp As Double
    Dim stageStartStamp As Double

    On Error GoTo EH

    runStartStamp = Timer
    stageName = "init"
    Set context = CreateObject("Scripting.Dictionary")
    context.CompareMode = 1
    ' mp_DebugLog "RUN START"

    stageName = "read-script-ref"
    stageStartStamp = Timer
    scriptLoadError = vbNullString
    If Not ex_ScriptSourceLoader.m_TryGetScriptText(cfg, PREPROCESS_SCRIPT_KEY, scriptRef, scriptLoadError) Then
        Err.Raise vbObjectError + 6156, "ex_PreProcessPipeline", scriptLoadError
    End If
    scriptRef = Trim$(scriptRef)
    If Len(scriptLoadError) > 0 Then
        Err.Raise vbObjectError + 6156, "ex_PreProcessPipeline", scriptLoadError
    End If
    mp_DebugLog "stage-done stage='read-script-ref' duration=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(stageStartStamp))
    ' mp_DebugLog "scriptLoaded=" & LCase$(CStr(Len(scriptRef) > 0)) & " scriptLength=" & CStr(Len(scriptRef)) & " requireScript=" & LCase$(CStr(requireScript))

    stageName = "reset-scriptio-context"
    ex_ScriptIO.m_ResetContext inputObject

    If Len(scriptRef) = 0 Then
        stageName = "fallback-no-script"
        If requireScript Then
            Err.Raise vbObjectError + 6150, "ex_PreProcessPipeline", "Missing required config key '" & PREPROCESS_SCRIPT_KEY & "'."
        End If

        Set outputObject = mp_CreateFallbackOutput(ex_ScriptIO.m_GetInput())

        context(CONTEXT_FIELD_HAS_SCRIPT) = "false"
        Set context(CONTEXT_FIELD_OUTPUT) = outputObject
        mp_DebugLog "RUN DURATION total=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(runStartStamp))
        ' mp_DebugLog "RUN END fallback"
        Set m_Run = context
        Exit Function
    End If

    stageName = "run-preprocess-script"
    stageStartStamp = Timer
    Set resultTables = New Collection
    ex_ScriptDSL.m_ApplyScriptToSheet mp_GetExecutionSheet(), cfg, resultTables, PREPROCESS_SCRIPT_KEY
    mp_DebugLog "stage-done stage='run-preprocess-script' duration=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(stageStartStamp))

    stageName = "read-script-output"
    Set outputObject = ex_ScriptIO.m_GetLastOutput()
    If outputObject Is Nothing Then
        Err.Raise vbObjectError + 6151, "ex_PreProcessPipeline", "PreProcess script must call ex_ScriptIO.m_CreateOutput() and populate output."
    End If

    context(CONTEXT_FIELD_HAS_SCRIPT) = "true"
    Set context(CONTEXT_FIELD_OUTPUT) = outputObject
    mp_DebugLog "RUN DURATION total=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(runStartStamp))
    ' mp_DebugLog "RUN END script"
    Set m_Run = context
    Exit Function

EH:
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    mp_DebugLog "FAIL stage='" & stageName & "' err=[" & errSource & " #" & CStr(errNumber) & "] " & errDescription & " | elapsed=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(runStartStamp))
    If errNumber = 0 Then errNumber = vbObjectError + 6155
    If Len(errSource) = 0 Then errSource = "ex_PreProcessPipeline"
    If Len(errDescription) = 0 Then errDescription = "Unknown pre-process pipeline failure."
    Err.Raise errNumber, errSource, errDescription
End Function

Private Function mp_StageElapsedSeconds(ByVal startStamp As Double) As Double
    mp_StageElapsedSeconds = Timer - startStamp
    If mp_StageElapsedSeconds < 0# Then mp_StageElapsedSeconds = mp_StageElapsedSeconds + 86400#
End Function

Private Function mp_FormatElapsedSeconds(ByVal elapsedSeconds As Double) As String
    mp_FormatElapsedSeconds = Format$(elapsedSeconds, "0.000") & "s"
End Function

Private Function mp_CreateFallbackOutput(ByVal inputObject As Object) As Object
    If inputObject Is Nothing Then
        Set mp_CreateFallbackOutput = ex_ScriptIO.m_CreateOutput()
        Exit Function
    End If

    Set mp_CreateFallbackOutput = inputObject
End Function

Private Function mp_GetExecutionSheet() As Worksheet
    On Error Resume Next
    Set mp_GetExecutionSheet = ws_Dev
    On Error GoTo 0

    If mp_GetExecutionSheet Is Nothing Then
        On Error Resume Next
        Set mp_GetExecutionSheet = ActiveSheet
        On Error GoTo 0
    End If

    If mp_GetExecutionSheet Is Nothing Then
        Set mp_GetExecutionSheet = ThisWorkbook.Worksheets(1)
    End If
End Function

Private Sub mp_DebugLog(ByVal messageText As String)
    If Not DEBUG_LOG_ENABLED Then Exit Sub
    On Error Resume Next
    ' Comment next line to disable file logger quickly.
    ex_Messaging.m_LogToFile "[ex_PreProcessPipeline] " & CStr(messageText), DEBUG_LOG_PATH
    On Error GoTo 0
End Sub
