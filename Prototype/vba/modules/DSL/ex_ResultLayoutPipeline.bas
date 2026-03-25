Attribute VB_Name = "ex_ResultLayoutPipeline"
Option Explicit

Private Const RESULT_LAYOUT_SCRIPT_KEY As String = "ResultLayout.Script"
Private Const INPUT_KEY_RESULT_TABLES As String = "__ResultTables"
Private Const INPUT_KEY_LAYOUT_WORKSHEET As String = "__ResultLayoutWorksheet"
Private Const INPUT_KEY_LAYOUT_SHEET_NAME As String = "__ResultLayoutSheetName"
Private Const DEBUG_LOG_PATH As String = "Logs\personalcard_pipeline.log"
Private Const DEBUG_LOG_ENABLED As Boolean = True

Public Function m_Run( _
    ByVal cfg As Object, _
    ByVal ws As Worksheet, _
    ByVal resultTables As Collection, _
    Optional ByVal inputObject As Object = Nothing, _
    Optional ByVal requireScript As Boolean = False _
) As Boolean
    Dim scriptRef As String
    Dim scriptLoadError As String
    Dim stageName As String

    On Error GoTo EH

    stageName = "validate-input"
    If cfg Is Nothing Then
        Err.Raise vbObjectError + 6240, "ex_ResultLayoutPipeline", "Config object is required for result-layout execution."
    End If
    If ws Is Nothing Then
        Err.Raise vbObjectError + 6244, "ex_ResultLayoutPipeline", "Worksheet is required for result-layout execution."
    End If
    If resultTables Is Nothing Then
        Err.Raise vbObjectError + 6245, "ex_ResultLayoutPipeline", "ResultTables are required for result-layout execution."
    End If

    If inputObject Is Nothing Then
        Set inputObject = New obj_ScriptIOPayload
    End If

    stageName = "prepare-script-input"
    ex_ScriptIO.m_SetObject inputObject, INPUT_KEY_RESULT_TABLES, resultTables
    ex_ScriptIO.m_SetObject inputObject, INPUT_KEY_LAYOUT_WORKSHEET, ws
    ex_ScriptIO.m_SetString inputObject, INPUT_KEY_LAYOUT_SHEET_NAME, ws.Name

    stageName = "load-script"
    scriptLoadError = vbNullString
    If Not ex_ScriptSourceLoader.m_TryGetScriptText(cfg, RESULT_LAYOUT_SCRIPT_KEY, scriptRef, scriptLoadError) Then
        Err.Raise vbObjectError + 6241, "ex_ResultLayoutPipeline", scriptLoadError
    End If

    scriptRef = Trim$(scriptRef)
    If Len(scriptLoadError) > 0 Then
        Err.Raise vbObjectError + 6241, "ex_ResultLayoutPipeline", scriptLoadError
    End If

    If Len(scriptRef) = 0 Then
        If requireScript Then
            Err.Raise vbObjectError + 6242, "ex_ResultLayoutPipeline", "Missing required resultLayoutScript for key '" & RESULT_LAYOUT_SCRIPT_KEY & "'."
        End If
        Exit Function
    End If

    stageName = "execute-script"
    ex_ScriptIO.m_SetInput inputObject
    ex_ScriptDSL.m_ApplyScriptToSheet ws, cfg, resultTables, RESULT_LAYOUT_SCRIPT_KEY
    m_Run = True
    Exit Function

EH:
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    mp_DebugLog "FAIL stage='" & stageName & "' err=[" & errSource & " #" & CStr(errNumber) & "] " & errDescription
    If errNumber = 0 Then errNumber = vbObjectError + 6243
    If Len(errSource) = 0 Then errSource = "ex_ResultLayoutPipeline"
    If Len(errDescription) = 0 Then errDescription = "Unknown result-layout pipeline failure."
    Err.Raise errNumber, errSource, errDescription
End Function

Private Sub mp_DebugLog(ByVal messageText As String)
    If Not DEBUG_LOG_ENABLED Then Exit Sub
    On Error Resume Next
    ex_Messaging.m_LogToFile "[ex_ResultLayoutPipeline] " & CStr(messageText), DEBUG_LOG_PATH
    On Error GoTo 0
End Sub
