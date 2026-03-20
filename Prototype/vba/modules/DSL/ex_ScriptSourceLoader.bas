Attribute VB_Name = "ex_ScriptSourceLoader"
Option Explicit

Private Const DEFAULT_SCRIPT_KEY As String = "PostProcess.Script"
Private Const IMPLICIT_SCRIPT_KEY As String = "PostProcess.Script.Implicit"
Private Const EXPLICIT_SCRIPT_KEY As String = "PostProcess.Script.Explicit"
Private Const PREPROCESS_SCRIPT_KEY As String = "Input.PreProcessScript"
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
    Dim normalizedScriptKey As String

    outScriptText = vbNullString
    outErrorText = vbNullString
    normalizedScriptKey = Trim$(scriptConfigKey)
    If Len(normalizedScriptKey) = 0 Then normalizedScriptKey = DEFAULT_SCRIPT_KEY

    If Not mp_TryGetScriptTextFromActiveProfile(normalizedScriptKey, outScriptText, outErrorText) Then
        Exit Function
    End If

    If Len(outScriptText) = 0 Then
        outErrorText = "Script is not configured for key '" & normalizedScriptKey & "' in the active profile."
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
    Dim preScriptNodes As Object
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

    If StrComp(normalizedScriptKey, PREPROCESS_SCRIPT_KEY, vbTextCompare) = 0 Then
        stepName = "read-preprocess-script-node"
        Set preScriptNodes = profileNode.selectNodes("p:preProcessScript")
        If preScriptNodes Is Nothing Then
            outErrorText = "preProcessScript node is required for key '" & PREPROCESS_SCRIPT_KEY & "'."
            Exit Function
        End If

        If preScriptNodes.Length = 0 Then
            outErrorText = "preProcessScript node is required for key '" & PREPROCESS_SCRIPT_KEY & "'."
            Exit Function
        End If

        If preScriptNodes.Length > 1 Then
            Err.Raise vbObjectError + 1776, "ex_ScriptSourceLoader", _
                "Only one preProcessScript node is allowed per profile."
        End If

        For Each scriptNode In preScriptNodes
            outScriptText = mp_GetScriptNodeText(scriptNode, filePath, "preProcessScript")
            Exit For
        Next scriptNode

        mp_TryGetScriptTextFromActiveProfile = True
        Exit Function
    End If

    stepName = "validate-script-execution"
    Set scriptNodes = profileNode.selectNodes("p:postProcessScript")
    If Not scriptNodes Is Nothing Then
        If scriptNodes.Length > 2 Then
            Err.Raise vbObjectError + 1767, "ex_ScriptSourceLoader", _
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
                Err.Raise vbObjectError + 1768, "ex_ScriptSourceLoader", _
                    "Attribute execution is required for postProcessScript. Allowed values: Implicit, Explicit."
            End If

            If StrComp(executionAttrValue, EXECUTION_MODE_IMPLICIT, vbTextCompare) <> 0 And _
               StrComp(executionAttrValue, EXECUTION_MODE_EXPLICIT, vbTextCompare) <> 0 Then
                Err.Raise vbObjectError + 1766, "ex_ScriptSourceLoader", _
                    "Unsupported postProcessScript execution='" & executionAttrValue & "'. Allowed values: Implicit, Explicit."
            End If

            If executionModesSeen.Exists(executionAttrValue) Then
                Err.Raise vbObjectError + 1769, "ex_ScriptSourceLoader", _
                    "Duplicate postProcessScript execution='" & executionAttrValue & "'. Only one script per execution mode is allowed."
            End If
            executionModesSeen(executionAttrValue) = True
        Next scriptNode
    End If

    stepName = "read-script-node"
    If StrComp(executionMode, EXECUTION_MODE_IMPLICIT, vbTextCompare) = 0 Then
        xpath = "p:postProcessScript[translate(normalize-space(@execution), '" & ASCII_UPPER & "', '" & ASCII_LOWER & "')='implicit']"
    Else
        xpath = "p:postProcessScript[translate(normalize-space(@execution), '" & ASCII_UPPER & "', '" & ASCII_LOWER & "')='explicit']"
    End If

    Set scriptNodes = profileNode.selectNodes(xpath)
    If scriptNodes Is Nothing Then
        outErrorText = "postProcessScript execution='" & executionMode & "' is required for key '" & normalizedScriptKey & "'."
        Exit Function
    End If

    If scriptNodes.Length = 0 Then
        outErrorText = "postProcessScript execution='" & executionMode & "' is required for key '" & normalizedScriptKey & "'."
        Exit Function
    End If

    For Each scriptNode In scriptNodes
        nodeText = mp_GetScriptNodeText(scriptNode, filePath, "postProcessScript")
        If Len(nodeText) = 0 Then GoTo ContinueNode
        If Len(outScriptText) > 0 Then outScriptText = outScriptText & vbLf & vbLf
        outScriptText = outScriptText & nodeText
ContinueNode:
    Next scriptNode

    mp_TryGetScriptTextFromActiveProfile = True
    Exit Function

EH:
    outErrorText = "Script load failed at step '" & stepName & "' [modeKey=" & modeKey & "] [profile=" & profileName & "] [file=" & filePath & "]: [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description
End Function

Private Function mp_GetScriptNodeText( _
    ByVal scriptNode As Object, _
    ByVal ownerFilePath As String, _
    ByVal scriptNodeName As String _
) As String
    Dim includePath As String
    Dim includeFullPath As String
    Dim inlineScriptText As String

    includePath = Trim$(ex_XmlCore.m_NodeAttrText(scriptNode, "include"))
    inlineScriptText = CStr(scriptNode.Text)

    If Len(includePath) > 0 Then
        If Len(Trim$(inlineScriptText)) > 0 Then
            Err.Raise vbObjectError + 1773, "ex_ScriptSourceLoader", _
                scriptNodeName & " cannot define both inline body and include attribute in the same node."
        End If

        includeFullPath = mp_ResolveIncludeFilePath(ownerFilePath, includePath)
        If Len(includeFullPath) = 0 Then
            Err.Raise vbObjectError + 1774, "ex_ScriptSourceLoader", _
                scriptNodeName & " include path could not be resolved: '" & includePath & "'."
        End If

        mp_GetScriptNodeText = mp_ReadTextFileUtf8(includeFullPath, scriptNodeName)
        Exit Function
    End If

    mp_GetScriptNodeText = Trim$(inlineScriptText)
    If Len(mp_GetScriptNodeText) > 0 Then
        mp_GetScriptNodeText = Replace(mp_GetScriptNodeText, "\n", vbLf)
    End If
End Function

Private Function mp_ReadTextFileUtf8(ByVal filePath As String, Optional ByVal scriptNodeName As String = "script") As String
    Dim normalizedPath As String
    Dim stream As Object

    normalizedPath = mp_NormalizeFilePath(filePath)
    If Len(normalizedPath) = 0 Then Exit Function
    If Len(Dir$(normalizedPath)) = 0 Then
        Err.Raise vbObjectError + 1775, "ex_ScriptSourceLoader", _
            scriptNodeName & " include file was not found: '" & normalizedPath & "'."
    End If

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile normalizedPath
    mp_ReadTextFileUtf8 = CStr(stream.ReadText(-1))
    stream.Close
End Function

Private Function mp_ResolveIncludeFilePath(ByVal ownerFilePath As String, ByVal includePath As String) As String
    Dim normalizedIncludePath As String
    Dim ownerDir As String
    Dim combinedPath As String

    normalizedIncludePath = mp_NormalizeFilePath(includePath)
    If Len(normalizedIncludePath) = 0 Then Exit Function

    If mp_IsAbsolutePath(normalizedIncludePath) Then
        mp_ResolveIncludeFilePath = normalizedIncludePath
        Exit Function
    End If

    ownerDir = mp_GetParentDirectory(ownerFilePath)
    If Len(ownerDir) = 0 Then Exit Function

    combinedPath = ownerDir & "\" & normalizedIncludePath
    mp_ResolveIncludeFilePath = mp_NormalizeFilePath(combinedPath)
End Function

Private Function mp_GetParentDirectory(ByVal filePath As String) As String
    Dim slashPos As Long
    Dim normalized As String

    normalized = mp_NormalizeFilePath(filePath)
    If Len(normalized) = 0 Then Exit Function

    slashPos = InStrRev(normalized, "\", -1, vbBinaryCompare)
    If slashPos <= 0 Then Exit Function
    If slashPos = 1 Then
        mp_GetParentDirectory = "\"
    Else
        mp_GetParentDirectory = Left$(normalized, slashPos - 1)
    End If
End Function

Private Function mp_IsAbsolutePath(ByVal filePath As String) As Boolean
    Dim normalized As String

    normalized = mp_NormalizeFilePath(filePath)
    If Len(normalized) = 0 Then Exit Function

    If Left$(normalized, 2) = "\\" Then
        mp_IsAbsolutePath = True
        Exit Function
    End If

    If Len(normalized) >= 3 Then
        If Mid$(normalized, 2, 1) = ":" And Mid$(normalized, 3, 1) = "\" Then
            mp_IsAbsolutePath = True
            Exit Function
        End If
    End If
End Function

Private Function mp_NormalizeFilePath(ByVal filePath As String) As String
    filePath = Trim$(filePath)
    If Len(filePath) = 0 Then Exit Function
    filePath = Replace$(filePath, "/", "\")
    mp_NormalizeFilePath = filePath
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

