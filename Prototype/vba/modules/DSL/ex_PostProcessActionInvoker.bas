Attribute VB_Name = "ex_PostProcessActionInvoker"
Option Explicit

Private Const ERR_SOURCE As String = "ex_PostProcessDsl"

Public Sub m_RunMacroWithArgs(ByVal macroName As String, ByVal args As Collection)
    Dim argCount As Long

    macroName = Trim$(macroName)
    If Len(macroName) = 0 Then
        Err.Raise vbObjectError + 1599, ERR_SOURCE, "Macro name is empty."
    End If

    If args Is Nothing Then
        Application.Run macroName
        Exit Sub
    End If

    argCount = args.Count

    Select Case argCount
        Case 0
            Application.Run macroName
        Case 1
            Application.Run macroName, args(1)
        Case 2
            Application.Run macroName, args(1), args(2)
        Case 3
            Application.Run macroName, args(1), args(2), args(3)
        Case 4
            Application.Run macroName, args(1), args(2), args(3), args(4)
        Case 5
            Application.Run macroName, args(1), args(2), args(3), args(4), args(5)
        Case Else
            Err.Raise vbObjectError + 1600, ERR_SOURCE, "Too many callMacro arguments (max 5)."
    End Select
End Sub
