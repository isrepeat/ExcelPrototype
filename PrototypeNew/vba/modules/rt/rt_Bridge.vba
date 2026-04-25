Attribute VB_Name = "rt_Bridge"
Option Explicit

' Стабильный мост между Shape.OnAction и rt_PageManager.
' Runtime-состояние контролов теперь принадлежит конкретной странице.

Private g_IsDispatchingClick As Boolean

' //
' // API
' //
Public Sub m_OnShapeClick()
    Dim callerShapeName As String
    Dim activeSheetObj As Object
    Dim ws As Worksheet
    Dim page As obj_IPage
    Dim dispatchOk As Boolean
    Dim wsName As String
    Dim wsCodeName As String

    On Error GoTo EH_CLICK
    g_IsDispatchingClick = True

    On Error Resume Next
    callerShapeName = VBA.CStr(Application.Caller)
    On Error GoTo EH_CLICK

    callerShapeName = VBA.Trim$(callerShapeName)
    If VBA.Len(callerShapeName) = 0 Then
        private_LogBridgeError "click-skip reason='caller-empty'"
        GoTo CleanExit
    End If

    Set activeSheetObj = Application.ActiveSheet
    If Not TypeOf activeSheetObj Is Worksheet Then
        private_LogBridgeError "click-skip reason='active-sheet-not-worksheet' shape='" & private_EscapeForLog(callerShapeName) & "'"
        GoTo CleanExit
    End If

    Set ws = activeSheetObj
    wsName = VBA.Trim$(VBA.CStr(ws.Name))
    wsCodeName = VBA.Trim$(VBA.CStr(ws.CodeName))
    private_LogBridgeInfo "click-start shape='" & private_EscapeForLog(callerShapeName) & "' sheet='" & private_EscapeForLog(wsName) & "' codeName='" & private_EscapeForLog(wsCodeName) & "'"

    If Not rt_PageManager.m_TryGetPageByWorksheet(ws, page) Then
        private_LogBridgeError "click-skip reason='page-not-found' shape='" & private_EscapeForLog(callerShapeName) & "' sheet='" & private_EscapeForLog(wsName) & "' codeName='" & private_EscapeForLog(wsCodeName) & "'"
        GoTo CleanExit
    End If

    dispatchOk = page.DispatchShapeClick(callerShapeName)
    If Not dispatchOk Then
        private_LogBridgeError "click-dispatch-failed shape='" & private_EscapeForLog(callerShapeName) & "' sheet='" & private_EscapeForLog(wsName) & "'"
        GoTo CleanExit
    End If

    private_LogBridgeInfo "click-done shape='" & private_EscapeForLog(callerShapeName) & "' sheet='" & private_EscapeForLog(wsName) & "'"

CleanExit:
    g_IsDispatchingClick = False
    Exit Sub

EH_CLICK:
    g_IsDispatchingClick = False
    private_LogBridgeError "click-exception err='" & private_EscapeForLog(Err.Description) & "'"
    VBA.MsgBox "rt_Bridge: shape click dispatch failed: " & Err.Description, VBA.vbExclamation
End Sub


Public Function m_IsDispatchingClick() As Boolean
    m_IsDispatchingClick = g_IsDispatchingClick
End Function


Public Function m_RunMacro(ByVal macroRef As String) As Boolean
    macroRef = VBA.Trim$(macroRef)
    If VBA.Len(macroRef) = 0 Then
        m_RunMacro = True
        Exit Function
    End If

    On Error GoTo EH_RUN
    Application.Run macroRef
    m_RunMacro = True
    Exit Function

EH_RUN:
    VBA.MsgBox "rt_Bridge: failed to execute macro '" & macroRef & "': " & Err.Description, VBA.vbExclamation
End Function

Private Function private_EscapeForLog(ByVal valueText As String) As String
    private_EscapeForLog = VBA.Replace$(VBA.CStr(valueText), "'", "''")
End Function

Private Sub private_LogBridgeInfo(ByVal messageText As String)
    On Error Resume Next
    ex_Core.m_Diagnostic_LogInfo "bridge:" & VBA.Trim$(messageText)
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub private_LogBridgeError(ByVal messageText As String)
    On Error Resume Next
    ex_Core.m_Diagnostic_LogError "bridge:" & VBA.Trim$(messageText)
    Err.Clear
    On Error GoTo 0
End Sub
