Attribute VB_Name = "ex_PostProcessActionInvoker"
Option Explicit

Private Const ERR_SOURCE As String = "ex_PostProcessDsl"

Public Sub m_RunMacroWithArgs(ByVal macroName As String, ByVal args As Collection)
    m_RunMacroWithArgsReturn macroName, args
End Sub

Public Function m_RunMacroWithArgsReturn(ByVal macroName As String, ByVal args As Collection) As Variant
    Dim argCount As Long
    Dim argsDump As String
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    macroName = Trim$(macroName)
    If Len(macroName) = 0 Then
        Err.Raise vbObjectError + 1599, ERR_SOURCE, "Macro name is empty."
    End If

    If mp_ExpectsObjectResult(macroName) Then
        Set m_RunMacroWithArgsReturn = mp_RunObjectMacroWithArgsReturn(macroName, args)
        Exit Function
    End If

    argsDump = mp_BuildArgsDebugText(args)

    If args Is Nothing Then
        On Error GoTo RunErr
        m_RunMacroWithArgsReturn = Application.Run(macroName)
        On Error GoTo 0
        Exit Function
    End If

    argCount = args.Count

    Select Case argCount
        Case 0
            On Error GoTo RunErr
            m_RunMacroWithArgsReturn = Application.Run(macroName)
            On Error GoTo 0
        Case 1
            On Error GoTo RunErr
            m_RunMacroWithArgsReturn = Application.Run(macroName, args(1))
            On Error GoTo 0
        Case 2
            On Error GoTo RunErr
            m_RunMacroWithArgsReturn = Application.Run(macroName, args(1), args(2))
            On Error GoTo 0
        Case 3
            On Error GoTo RunErr
            m_RunMacroWithArgsReturn = Application.Run(macroName, args(1), args(2), args(3))
            On Error GoTo 0
        Case 4
            On Error GoTo RunErr
            m_RunMacroWithArgsReturn = Application.Run(macroName, args(1), args(2), args(3), args(4))
            On Error GoTo 0
        Case 5
            On Error GoTo RunErr
            m_RunMacroWithArgsReturn = Application.Run(macroName, args(1), args(2), args(3), args(4), args(5))
            On Error GoTo 0
        Case 6
            On Error GoTo RunErr
            m_RunMacroWithArgsReturn = Application.Run(macroName, args(1), args(2), args(3), args(4), args(5), args(6))
            On Error GoTo 0
        Case 7
            On Error GoTo RunErr
            m_RunMacroWithArgsReturn = Application.Run(macroName, args(1), args(2), args(3), args(4), args(5), args(6), args(7))
            On Error GoTo 0
        Case Else
            Err.Raise vbObjectError + 1600, ERR_SOURCE, "Too many callMacro arguments (max 7)."
    End Select
    Exit Function

RunErr:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    On Error GoTo 0
    Err.Raise vbObjectError + 1606, ERR_SOURCE, _
        "Application.Run failed for '" & macroName & "' with args " & argsDump & _
        ": [" & errSource & " #" & CStr(errNumber) & "] " & errDescription
End Function

Private Function mp_ExpectsObjectResult(ByVal macroName As String) As Boolean
    macroName = LCase$(Trim$(macroName))
    If Right$(macroName, Len(".m_getrelativerow")) = ".m_getrelativerow" Then
        mp_ExpectsObjectResult = True
    End If
End Function

Private Function mp_RunObjectMacroWithArgsReturn(ByVal macroName As String, ByVal args As Collection) As Object
    Dim normalized As String
    Dim rowRef As obj_ResultRow
    Dim rowOffsetText As String

    normalized = LCase$(Trim$(macroName))

    If Right$(normalized, Len(".m_getrelativerow")) = ".m_getrelativerow" Then
        If args Is Nothing Or args.Count <> 2 Then
            Err.Raise vbObjectError + 1602, ERR_SOURCE, "m_GetRelativeRow expects exactly 2 arguments: rowRef, rowOffsetText."
        End If
        If Not IsObject(args(1)) Then
            Err.Raise vbObjectError + 1603, ERR_SOURCE, "m_GetRelativeRow expects rowRef object as first argument."
        End If
        If Not TypeOf args(1) Is obj_ResultRow Then
            Err.Raise vbObjectError + 1604, ERR_SOURCE, "m_GetRelativeRow first argument must be obj_ResultRow."
        End If

        Set rowRef = args(1)
        rowOffsetText = CStr(args(2))
        Set mp_RunObjectMacroWithArgsReturn = ex_PostProcessActions.m_GetRelativeRow(rowRef, rowOffsetText)
        Exit Function
    End If

    Err.Raise vbObjectError + 1605, ERR_SOURCE, "Unsupported object-return macro: " & macroName
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
