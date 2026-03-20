Attribute VB_Name = "ex_ScriptActionInvoker"
Option Explicit

Private Const ERR_SOURCE As String = "ex_ScriptDsl"
Private Const DEBUG_LOG_PATH As String = "Logs\personalcard_pipeline.log"
Private Const DEBUG_LOG_ENABLED As Boolean = True

Public Sub m_RunMacroWithArgs(ByVal macroName As String, ByVal args As Collection)
    m_RunMacroWithArgsReturn macroName, args
End Sub

Public Function m_RunObjectMacroWithArgsReturn(ByVal macroName As String, ByVal args As Collection) As Object
    Dim argsDump As String
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim objectResult As Object

    macroName = Trim$(macroName)
    If Len(macroName) = 0 Then
        Err.Raise vbObjectError + 1599, ERR_SOURCE, "Macro name is empty."
    End If

    argsDump = mp_BuildArgsDebugText(args)
    mp_DebugLog "RUN object macro='" & macroName & "' args=" & argsDump

    On Error GoTo RunErr
    Set objectResult = mp_RunObjectMacroWithArgs(macroName, args)
    On Error GoTo 0

    If objectResult Is Nothing Then
        Err.Raise vbObjectError + 1607, ERR_SOURCE, "callMacroObject returned Nothing for '" & macroName & "'."
    End If

    Set m_RunObjectMacroWithArgsReturn = objectResult
    mp_DebugLog "OK object macro='" & macroName & "' result=<object:" & TypeName(objectResult) & ">"
    Exit Function

RunErr:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    mp_DebugLog "FAIL object macro='" & macroName & "' args=" & argsDump & " err=[" & errSource & " #" & CStr(errNumber) & "] " & errDescription
    On Error GoTo 0
    Err.Raise vbObjectError + 1608, ERR_SOURCE, _
        "Application.Run failed for object macro '" & macroName & "' with args " & argsDump & _
        ": [" & errSource & " #" & CStr(errNumber) & "] " & errDescription
End Function

Public Function m_RunMacroWithArgsReturn(ByVal macroName As String, ByVal args As Collection) As Variant
    Dim argCount As Long
    Dim argsDump As String
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim runResult As Variant
    Dim objectResult As Object

    macroName = Trim$(macroName)
    If Len(macroName) = 0 Then
        Err.Raise vbObjectError + 1599, ERR_SOURCE, "Macro name is empty."
    End If

    argsDump = mp_BuildArgsDebugText(args)
    mp_DebugLog "RUN macro='" & macroName & "' args=" & argsDump

    If args Is Nothing Then
        On Error GoTo RunErr
        runResult = Application.Run(macroName)
        On Error GoTo 0
        If IsObject(runResult) Then
            Set m_RunMacroWithArgsReturn = runResult
        Else
            m_RunMacroWithArgsReturn = runResult
        End If
        mp_DebugLog "OK macro='" & macroName & "' result=<" & mp_DescribeVariant(runResult) & ">"
        Exit Function
    End If

    argCount = args.Count

    Select Case argCount
        Case 0
            On Error GoTo RunErr
            runResult = Application.Run(macroName)
            On Error GoTo 0
        Case 1
            On Error GoTo RunErr
            runResult = Application.Run(macroName, args(1))
            On Error GoTo 0
        Case 2
            On Error GoTo RunErr
            runResult = Application.Run(macroName, args(1), args(2))
            On Error GoTo 0
        Case 3
            On Error GoTo RunErr
            runResult = Application.Run(macroName, args(1), args(2), args(3))
            On Error GoTo 0
        Case 4
            On Error GoTo RunErr
            runResult = Application.Run(macroName, args(1), args(2), args(3), args(4))
            On Error GoTo 0
        Case 5
            On Error GoTo RunErr
            runResult = Application.Run(macroName, args(1), args(2), args(3), args(4), args(5))
            On Error GoTo 0
        Case 6
            On Error GoTo RunErr
            runResult = Application.Run(macroName, args(1), args(2), args(3), args(4), args(5), args(6))
            On Error GoTo 0
        Case 7
            On Error GoTo RunErr
            runResult = Application.Run(macroName, args(1), args(2), args(3), args(4), args(5), args(6), args(7))
            On Error GoTo 0
        Case Else
            Err.Raise vbObjectError + 1600, ERR_SOURCE, "Too many callMacro arguments (max 7)."
    End Select

    If IsObject(runResult) Then
        Set m_RunMacroWithArgsReturn = runResult
    Else
        m_RunMacroWithArgsReturn = runResult
    End If

    mp_DebugLog "OK macro='" & macroName & "' result=<" & mp_DescribeVariant(runResult) & ">"

    Exit Function

RunErr:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    mp_DebugLog "FAIL macro='" & macroName & "' args=" & argsDump & " err=[" & errSource & " #" & CStr(errNumber) & "] " & errDescription
    On Error GoTo 0
    Err.Raise vbObjectError + 1606, ERR_SOURCE, _
        "Application.Run failed for '" & macroName & "' with args " & argsDump & _
        ": [" & errSource & " #" & CStr(errNumber) & "] " & errDescription
End Function

Private Function mp_RunObjectMacroWithArgs(ByVal macroName As String, ByVal args As Collection) As Object
    Dim argCount As Long

    If args Is Nothing Then
        Set mp_RunObjectMacroWithArgs = Application.Run(macroName)
        Exit Function
    End If

    argCount = args.Count
    Select Case argCount
        Case 0
            Set mp_RunObjectMacroWithArgs = Application.Run(macroName)
        Case 1
            Set mp_RunObjectMacroWithArgs = Application.Run(macroName, args(1))
        Case 2
            Set mp_RunObjectMacroWithArgs = Application.Run(macroName, args(1), args(2))
        Case 3
            Set mp_RunObjectMacroWithArgs = Application.Run(macroName, args(1), args(2), args(3))
        Case 4
            Set mp_RunObjectMacroWithArgs = Application.Run(macroName, args(1), args(2), args(3), args(4))
        Case 5
            Set mp_RunObjectMacroWithArgs = Application.Run(macroName, args(1), args(2), args(3), args(4), args(5))
        Case 6
            Set mp_RunObjectMacroWithArgs = Application.Run(macroName, args(1), args(2), args(3), args(4), args(5), args(6))
        Case 7
            Set mp_RunObjectMacroWithArgs = Application.Run(macroName, args(1), args(2), args(3), args(4), args(5), args(6), args(7))
        Case Else
            Err.Raise vbObjectError + 1600, ERR_SOURCE, "Too many callMacro arguments (max 7)."
    End Select
End Function

Private Sub mp_DebugLog(ByVal messageText As String)
    If Not DEBUG_LOG_ENABLED Then Exit Sub
    On Error Resume Next
    ' Comment next line to disable file logger quickly.
    ex_Messaging.m_LogToFile "[ex_ScriptActionInvoker] " & CStr(messageText), DEBUG_LOG_PATH
    On Error GoTo 0
End Sub

Private Function mp_DescribeVariant(ByVal valueRef As Variant) As String
    If IsObject(valueRef) Then
        If valueRef Is Nothing Then
            mp_DescribeVariant = "object:Nothing"
        Else
            mp_DescribeVariant = "object:" & TypeName(valueRef)
        End If
        Exit Function
    End If
    If IsNull(valueRef) Then
        mp_DescribeVariant = "null"
        Exit Function
    End If
    If IsError(valueRef) Then
        mp_DescribeVariant = "error"
        Exit Function
    End If
    mp_DescribeVariant = "scalar:" & TypeName(valueRef)
End Function

Private Function mp_BuildArgsDebugText(ByVal args As Collection) As String
    Dim i As Long
    Dim partText As String
    Dim valueRef As Variant
    Dim objRef As Object

    If args Is Nothing Then
        mp_BuildArgsDebugText = "[]"
        Exit Function
    End If

    partText = "["
    For i = 1 To args.Count
        If i > 1 Then partText = partText & ", "
        If IsObject(args(i)) Then
            Set objRef = args(i)
            partText = partText & "<object:" & TypeName(objRef) & ">"
        Else
            valueRef = args(i)
            If IsNull(valueRef) Then
                partText = partText & "<null>"
            ElseIf IsError(valueRef) Then
                partText = partText & "<error>"
            Else
                partText = partText & """" & Replace(CStr(valueRef), """", """""") & """"
            End If
        End If
    Next i
    partText = partText & "]"

    mp_BuildArgsDebugText = partText
End Function
