Attribute VB_Name = "ex_PostProcessScriptSource"
Option Explicit

Private Const DEFAULT_SCRIPT_KEY As String = "PostProcess.Script"

Public Function m_TryGetScriptText( _
    ByVal cfg As Object, _
    ByVal scriptConfigKey As String, _
    ByRef outScriptText As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim scriptFromProfile As String

    outScriptText = vbNullString
    outErrorText = vbNullString

    If Not mp_TryGetScriptTextFromActiveProfile(scriptFromProfile, outErrorText) Then Exit Function
    If Len(scriptFromProfile) > 0 Then
        outScriptText = scriptFromProfile
        m_TryGetScriptText = True
        Exit Function
    End If

    On Error GoTo EH

    scriptConfigKey = Trim$(scriptConfigKey)
    If Len(scriptConfigKey) = 0 Then scriptConfigKey = DEFAULT_SCRIPT_KEY

    If Not cfg Is Nothing Then
        If cfg.Exists(scriptConfigKey) Then
            outScriptText = Trim$(CStr(cfg(scriptConfigKey)))
            If Len(outScriptText) > 0 Then outScriptText = Replace(outScriptText, "\n", vbLf)
        End If
    End If

    m_TryGetScriptText = True
    Exit Function

EH:
    outErrorText = "PostProcess script load failed: [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description
End Function

Private Function mp_TryGetScriptTextFromActiveProfile(ByRef outScriptText As String, ByRef outErrorText As String) As Boolean
    Dim modeKey As String
    Dim profileName As String
    Dim filePath As String
    Dim doc As Object
    Dim profileNode As Object
    Dim scriptNode As Object
    Dim stepName As String

    On Error GoTo EH

    outScriptText = vbNullString
    outErrorText = vbNullString

    stepName = "read-active-mode-profile"
    modeKey = Trim$(ex_ConfigProfilesManager.m_GetActiveModeKey(ws_Dev))
    profileName = Trim$(ex_ConfigProfilesManager.m_GetActiveProfileName(ws_Dev))
    If Len(modeKey) = 0 Or Len(profileName) = 0 Then
        mp_TryGetScriptTextFromActiveProfile = True
        Exit Function
    End If

    stepName = "resolve-profiles-path"
    filePath = ex_ProfilesStore.m_GetProfilesFilePath(modeKey, ThisWorkbook)
    If Len(Trim$(filePath)) = 0 Then
        mp_TryGetScriptTextFromActiveProfile = True
        Exit Function
    End If

    stepName = "load-profiles-dom"
    Set doc = ex_ProfilesStore.m_LoadProfilesDom(filePath)
    If doc Is Nothing Then
        mp_TryGetScriptTextFromActiveProfile = True
        Exit Function
    End If

    stepName = "find-profile-node"
    Set profileNode = ex_ProfilesStore.m_GetProfileNode(doc, profileName, False)
    If profileNode Is Nothing Then
        mp_TryGetScriptTextFromActiveProfile = True
        Exit Function
    End If

    stepName = "read-postprocess-node"
    Set scriptNode = profileNode.selectSingleNode("p:postProcessScript")
    If scriptNode Is Nothing Then
        mp_TryGetScriptTextFromActiveProfile = True
        Exit Function
    End If

    outScriptText = Trim$(CStr(scriptNode.Text))
    If Len(outScriptText) > 0 Then outScriptText = Replace(outScriptText, "\n", vbLf)
    mp_TryGetScriptTextFromActiveProfile = True
    Exit Function

EH:
    outErrorText = "PostProcess script load failed at step '" & stepName & "' [modeKey=" & modeKey & "] [profile=" & profileName & "] [file=" & filePath & "]: [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description
End Function
