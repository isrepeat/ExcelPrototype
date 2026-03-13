Attribute VB_Name = "ex_PostProcessScriptSource"
Option Explicit

Private Const DEFAULT_SCRIPT_KEY As String = "PostProcess.Script"
Private Const IMPLICIT_SCRIPT_KEY As String = "PostProcess.Script.Implicit"
Private Const EXPLICIT_SCRIPT_KEY As String = "PostProcess.Script.Explicit"
Private Const EXECUTION_MODE_IMPLICIT As String = "implicit"
Private Const EXECUTION_MODE_EXPLICIT As String = "explicit"
Private Const ASCII_UPPER As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Private Const ASCII_LOWER As String = "abcdefghijklmnopqrstuvwxyz"

Public Function m_TryGetScriptText( _
    ByVal cfg As Object, _
    ByVal scriptConfigKey As String, _
    ByRef outScriptText As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim scriptFromProfile As String
    Dim normalizedScriptKey As String

    outScriptText = vbNullString
    outErrorText = vbNullString
    normalizedScriptKey = Trim$(scriptConfigKey)
    If Len(normalizedScriptKey) = 0 Then normalizedScriptKey = DEFAULT_SCRIPT_KEY

    If Not mp_TryGetScriptTextFromActiveProfile(normalizedScriptKey, scriptFromProfile, outErrorText) Then Exit Function
    If Len(scriptFromProfile) > 0 Then
        outScriptText = scriptFromProfile
        m_TryGetScriptText = True
        Exit Function
    End If

    m_TryGetScriptText = True
    Exit Function
End Function

Private Function mp_TryGetScriptTextFromActiveProfile( _
    ByVal scriptConfigKey As String, _
    ByRef outScriptText As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim modeKey As String
    Dim profileName As String
    Dim filePath As String
    Dim doc As Object
    Dim profileNode As Object
    Dim scriptNodes As Object
    Dim scriptNode As Object
    Dim normalizedScriptKey As String
    Dim executionMode As String
    Dim xpath As String
    Dim nodeText As String
    Dim executionModesSeen As Object
    Dim executionAttrRaw As Variant
    Dim executionAttrValue As String
    Dim stepName As String

    On Error GoTo EH

    outScriptText = vbNullString
    outErrorText = vbNullString
    normalizedScriptKey = Trim$(scriptConfigKey)
    If Len(normalizedScriptKey) = 0 Then normalizedScriptKey = DEFAULT_SCRIPT_KEY
    executionMode = mp_GetExecutionModeByScriptKey(normalizedScriptKey)

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

    stepName = "validate-postprocess-execution"
    Set scriptNodes = profileNode.selectNodes("p:postProcessScript")
    If Not scriptNodes Is Nothing Then
        If scriptNodes.Length > 2 Then
            Err.Raise vbObjectError + 1767, "ex_PostProcessScriptSource", _
                "Only up to two postProcessScript nodes are allowed per profile: execution='Implicit' and execution='Explicit'."
        End If

        Set executionModesSeen = CreateObject("Scripting.Dictionary")
        executionModesSeen.CompareMode = 1 ' vbTextCompare

        For Each scriptNode In scriptNodes
            executionAttrRaw = scriptNode.getAttribute("execution")
            If IsNull(executionAttrRaw) Then
                executionAttrValue = vbNullString
            Else
                executionAttrValue = LCase$(Trim$(CStr(executionAttrRaw)))
            End If
            If Len(executionAttrValue) = 0 Then
                Err.Raise vbObjectError + 1768, "ex_PostProcessScriptSource", _
                    "Attribute execution is required for postProcessScript. Allowed values: Implicit, Explicit."
            End If

            If StrComp(executionAttrValue, EXECUTION_MODE_IMPLICIT, vbTextCompare) <> 0 And _
               StrComp(executionAttrValue, EXECUTION_MODE_EXPLICIT, vbTextCompare) <> 0 Then
                Err.Raise vbObjectError + 1766, "ex_PostProcessScriptSource", _
                    "Unsupported postProcessScript execution='" & executionAttrValue & "'. Allowed values: Implicit, Explicit."
            End If

            If executionModesSeen.Exists(executionAttrValue) Then
                Err.Raise vbObjectError + 1769, "ex_PostProcessScriptSource", _
                    "Duplicate postProcessScript execution='" & executionAttrValue & "'. Only one script per execution mode is allowed."
            End If
            executionModesSeen(executionAttrValue) = True
        Next scriptNode
    End If

    stepName = "read-postprocess-node"
    If StrComp(executionMode, EXECUTION_MODE_IMPLICIT, vbTextCompare) = 0 Then
        xpath = "p:postProcessScript[translate(normalize-space(@execution), '" & ASCII_UPPER & "', '" & ASCII_LOWER & "')='implicit']"
    Else
        xpath = "p:postProcessScript[translate(normalize-space(@execution), '" & ASCII_UPPER & "', '" & ASCII_LOWER & "')='explicit']"
    End If

    Set scriptNodes = profileNode.selectNodes(xpath)
    If scriptNodes Is Nothing Then
        mp_TryGetScriptTextFromActiveProfile = True
        Exit Function
    End If

    For Each scriptNode In scriptNodes
        nodeText = Trim$(CStr(scriptNode.Text))
        If Len(nodeText) = 0 Then GoTo ContinueNode
        nodeText = Replace(nodeText, "\n", vbLf)
        If Len(outScriptText) > 0 Then outScriptText = outScriptText & vbLf & vbLf
        outScriptText = outScriptText & nodeText
ContinueNode:
    Next scriptNode

    mp_TryGetScriptTextFromActiveProfile = True
    Exit Function

EH:
    outErrorText = "PostProcess script load failed at step '" & stepName & "' [modeKey=" & modeKey & "] [profile=" & profileName & "] [file=" & filePath & "]: [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description
End Function

Private Function mp_GetExecutionModeByScriptKey(ByVal scriptConfigKey As String) As String
    If StrComp(scriptConfigKey, IMPLICIT_SCRIPT_KEY, vbTextCompare) = 0 Then
        mp_GetExecutionModeByScriptKey = EXECUTION_MODE_IMPLICIT
        Exit Function
    End If

    If StrComp(scriptConfigKey, EXPLICIT_SCRIPT_KEY, vbTextCompare) = 0 Then
        mp_GetExecutionModeByScriptKey = EXECUTION_MODE_EXPLICIT
        Exit Function
    End If

    If StrComp(scriptConfigKey, DEFAULT_SCRIPT_KEY, vbTextCompare) = 0 Then
        mp_GetExecutionModeByScriptKey = EXECUTION_MODE_EXPLICIT
        Exit Function
    End If

    mp_GetExecutionModeByScriptKey = EXECUTION_MODE_EXPLICIT
End Function
