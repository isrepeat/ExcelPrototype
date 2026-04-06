Attribute VB_Name = "ex_ResultLayoutPipeline"
Option Explicit

Private Const DEBUG_LOG_PATH As String = "Logs\personalcard_pipeline.log"
Private Const DEBUG_LOG_ENABLED As Boolean = True

Public Function m_Run( _
    ByVal cfg As Object, _
    ByVal ws As Worksheet, _
    ByVal resultTables As Collection, _
    Optional ByVal inputObject As Object = Nothing _
) As Boolean
    Dim stageName As String
    Dim runStartStamp As Double
    Dim stageStartStamp As Double
    Dim xmlErrorText As String
    Dim xmlLayoutDoc As Object
    Dim xmlLayoutPath As String
    Dim hasXmlLayout As Boolean
    Dim xmlApplied As Boolean

    On Error GoTo EH

    runStartStamp = Timer
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

    If inputObject Is Nothing Then Set inputObject = New obj_ScriptIOPayload

    stageName = "load-xml-layout-dom"
    stageStartStamp = Timer
    xmlErrorText = vbNullString
    If Not mp_TryLoadXmlLayoutDom(xmlLayoutDoc, xmlLayoutPath, hasXmlLayout, xmlErrorText) Then
        Err.Raise vbObjectError + 6246, "ex_ResultLayoutPipeline", xmlErrorText
    End If
    mp_DebugLog "stage-done stage='load-xml-layout-dom' duration=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(stageStartStamp))

    If Not hasXmlLayout Then
        Err.Raise vbObjectError + 6247, "ex_ResultLayoutPipeline", _
            "Mode UI must define XML layout grid (/uiDefinition/layout/grid). Legacy result layout rendering is no longer supported."
    End If

    stageName = "execute-xml-layout"
    stageStartStamp = Timer
    If Not ex_ResultLayoutXmlEngine.m_ApplyResultLayoutFromDom(xmlLayoutDoc, ws, resultTables, inputObject, xmlErrorText) Then
        If Len(xmlErrorText) = 0 Then
            Err.Raise vbObjectError + 6246, "ex_ResultLayoutPipeline", "Result XML layout execution failed for mode UI file '" & xmlLayoutPath & "'."
        Else
            Err.Raise vbObjectError + 6246, "ex_ResultLayoutPipeline", "Result XML layout execution failed for mode UI file '" & xmlLayoutPath & "': " & xmlErrorText
        End If
    End If
    mp_DebugLog "stage-done stage='execute-xml-layout' duration=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(stageStartStamp))
    xmlApplied = True
    stageStartStamp = Timer
    ex_ResultLayoutItemsRt.m_RegisterSession ws, xmlLayoutDoc, resultTables, inputObject
    mp_DebugLog "stage-done stage='register-layout-session' duration=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(stageStartStamp))

    mp_DebugLog "RUN DURATION total=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(runStartStamp))
    m_Run = xmlApplied
    Exit Function

EH:
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    mp_DebugLog "FAIL stage='" & stageName & "' err=[" & errSource & " #" & CStr(errNumber) & "] " & errDescription & " | elapsed=" & mp_FormatElapsedSeconds(mp_StageElapsedSeconds(runStartStamp))
    If errNumber = 0 Then errNumber = vbObjectError + 6243
    If Len(errSource) = 0 Then errSource = "ex_ResultLayoutPipeline"
    If Len(errDescription) = 0 Then errDescription = "Unknown result-layout pipeline failure."
    On Error Resume Next
    ex_ResultLayoutItemsRt.m_ClearSession ws
    On Error GoTo 0
    Err.Raise errNumber, errSource, errDescription
End Function

Private Function mp_StageElapsedSeconds(ByVal startStamp As Double) As Double
    mp_StageElapsedSeconds = Timer - startStamp
    If mp_StageElapsedSeconds < 0# Then mp_StageElapsedSeconds = mp_StageElapsedSeconds + 86400#
End Function

Private Function mp_FormatElapsedSeconds(ByVal elapsedSeconds As Double) As String
    mp_FormatElapsedSeconds = Format$(elapsedSeconds, "0.000") & "s"
End Function

Private Function mp_TryLoadXmlLayoutDom( _
    ByRef outModeUiDoc As Object, _
    ByRef outModeUiFilePath As String, _
    ByRef outHasXmlLayout As Boolean, _
    ByRef outErrorText As String _
) As Boolean
    outHasXmlLayout = False
    outErrorText = vbNullString
    outModeUiFilePath = vbNullString
    Set outModeUiDoc = Nothing

    If Not ex_ResultLayoutXmlProvider.m_TryLoadActiveModeUiDom(ThisWorkbook, outModeUiDoc, outModeUiFilePath, outErrorText) Then
        Exit Function
    End If

    outHasXmlLayout = ex_ResultLayoutXmlProvider.m_HasResultLayoutGrid(outModeUiDoc)
    mp_TryLoadXmlLayoutDom = True
End Function

Private Sub mp_DebugLog(ByVal messageText As String)
    If Not DEBUG_LOG_ENABLED Then Exit Sub
    On Error Resume Next
    ex_Messaging.m_LogToFile "[ex_ResultLayoutPipeline] " & CStr(messageText), DEBUG_LOG_PATH
    On Error GoTo 0
End Sub
