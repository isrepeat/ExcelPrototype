Attribute VB_Name = "rt_UndoManager"
Option Explicit

Private Const DEFAULT_MAX_DEPTH As Long = 200

Private g_UndoStack As Collection
Private g_RedoStack As Collection
Private g_MaxDepth As Long
Private g_IsReplaying As Boolean
Private g_IsRecordingSuspended As Boolean

Public Sub fn_Module_Dispose()
    fn_CancelGlobalCallbacks
    Set g_UndoStack = Nothing
    Set g_RedoStack = Nothing
    g_MaxDepth = 0
    g_IsReplaying = False
    g_IsRecordingSuspended = False
End Sub


Public Sub fn_CancelGlobalCallbacks()
    private_ClearExcelUndoHotkeys
End Sub

' //
' // API
' //
Public Sub fn_SetMaxDepth(ByVal maxDepth As Long)
    private_EnsureStorage
    If maxDepth < 1 Then
        g_MaxDepth = DEFAULT_MAX_DEPTH
    Else
        g_MaxDepth = maxDepth
    End If
    private_TrimStack g_UndoStack
    private_TrimStack g_RedoStack
End Sub

Public Function fn_GetMaxDepth() As Long
    private_EnsureStorage
    fn_GetMaxDepth = g_MaxDepth
End Function

Public Function fn_GetUndoCount() As Long
    private_EnsureStorage
    fn_GetUndoCount = g_UndoStack.Count
End Function

Public Function fn_GetRedoCount() As Long
    private_EnsureStorage
    fn_GetRedoCount = g_RedoStack.Count
End Function

Public Function fn_CanUndo() As Boolean
    private_EnsureStorage
    fn_CanUndo = (g_UndoStack.Count > 0)
End Function

Public Function fn_CanRedo() As Boolean
    private_EnsureStorage
    fn_CanRedo = (g_RedoStack.Count > 0)
End Function

Public Function fn_TryUndoLastByScopePrefix( _
    ByVal scopePrefix As String, _
    Optional ByRef outActionMatched As Boolean _
) As Boolean
    Dim action As obj_IUndoAction
    Dim normalizedPrefix As String

    On Error GoTo EH
    outActionMatched = False
    private_EnsureStorage
    normalizedPrefix = VBA.Trim$(scopePrefix)
    If VBA.Len(normalizedPrefix) = 0 Then Exit Function
    If g_IsReplaying Or g_UndoStack.Count = 0 Then Exit Function

    Set action = g_UndoStack.Item(g_UndoStack.Count)
    If action Is Nothing Then Exit Function

    ' Scoped-команда не ищет подходящее действие глубже в истории: undo должен
    ' сохранять строгий LIFO-порядок и никогда не перепрыгивать более новое действие.
    If VBA.StrComp( _
        VBA.Left$(action.GetScopeKey(), VBA.Len(normalizedPrefix)), _
        normalizedPrefix, _
        VBA.vbTextCompare) <> 0 Then Exit Function

    outActionMatched = True
    fn_TryUndoLastByScopePrefix = fn_UndoLast()
EH:
End Function

Public Sub fn_ClearAll()
    private_EnsureStorage
    Set g_UndoStack = New Collection
    Set g_RedoStack = New Collection
    private_RegisterExcelUndoRedo
End Sub

Public Sub fn_SuspendRecording()
    g_IsRecordingSuspended = True
End Sub

Public Sub fn_ResumeRecording()
    g_IsRecordingSuspended = False
End Sub

Public Function fn_ExecuteAction(ByVal action As obj_IUndoAction) As Boolean
    Dim errorText As String
    Dim actionLabel As String

    On Error GoTo EH_EXECUTE
    actionLabel = private_GetActionDebugLabel(action)

    If action Is Nothing Then Exit Function
    If g_IsReplaying Then Exit Function
    If g_IsRecordingSuspended Then Exit Function


    If Not action.IsValid() Then
        VBA.MsgBox "PrototypeNew: undo action target is invalid.", VBA.vbExclamation, "PrototypeNew / Undo"
        Exit Function
    End If

    If Not action.Execute(errorText) Then
        If VBA.Len(VBA.Trim$(errorText)) = 0 Then errorText = "Unknown execute error."
        VBA.MsgBox "PrototypeNew: failed to execute action. " & errorText, VBA.vbExclamation, "PrototypeNew / Undo"
        Exit Function
    End If

    If Not fn_PushExecutedAction(action) Then Exit Function
    fn_ExecuteAction = True
    Exit Function

EH_EXECUTE:
End Function

Public Function fn_PushExecutedAction(ByVal action As obj_IUndoAction) As Boolean
    Dim actionLabel As String

    On Error GoTo EH_PUSH
    private_EnsureStorage
    actionLabel = private_GetActionDebugLabel(action)
    If action Is Nothing Then Exit Function
    If g_IsRecordingSuspended Then
        fn_PushExecutedAction = True
        Exit Function
    End If
    If g_IsReplaying Then
        fn_PushExecutedAction = True
        Exit Function
    End If

    g_UndoStack.Add action
    Set g_RedoStack = New Collection

    private_TrimStack g_UndoStack
    private_RegisterExcelUndoRedo
    fn_PushExecutedAction = True
    Exit Function

EH_PUSH:
End Function

Public Function fn_UndoLast() As Boolean
    Dim action As obj_IUndoAction
    Dim errorText As String
    Dim actionLabel As String

    On Error GoTo EH_UNDO

    private_EnsureStorage
    If g_IsReplaying Then Exit Function
    If g_UndoStack.Count <= 0 Then Exit Function


    Set action = private_PopLast(g_UndoStack)
    If action Is Nothing Then Exit Function
    actionLabel = private_GetActionDebugLabel(action)

    g_IsReplaying = True
    If action.Undo(errorText) Then
        g_RedoStack.Add action
        private_TrimStack g_RedoStack
        fn_UndoLast = True
    Else
        g_UndoStack.Add action
        If VBA.Len(VBA.Trim$(errorText)) = 0 Then errorText = "Unknown undo error."
        VBA.MsgBox "PrototypeNew: failed to undo action. " & errorText, VBA.vbExclamation, "PrototypeNew / Undo"
    End If
    g_IsReplaying = False
    private_RegisterExcelUndoRedo
    Exit Function

EH_UNDO:
    g_IsReplaying = False
    If Not action Is Nothing Then
        On Error Resume Next
        g_UndoStack.Add action
        On Error GoTo 0
    End If
    private_RegisterExcelUndoRedo
End Function

Public Function fn_RedoLast() As Boolean
    Dim action As obj_IUndoAction
    Dim errorText As String
    Dim actionLabel As String

    On Error GoTo EH_REDO

    private_EnsureStorage
    If g_IsReplaying Then Exit Function
    If g_RedoStack.Count <= 0 Then Exit Function


    Set action = private_PopLast(g_RedoStack)
    If action Is Nothing Then Exit Function
    actionLabel = private_GetActionDebugLabel(action)

    g_IsReplaying = True
    If action.Redo(errorText) Then
        g_UndoStack.Add action
        private_TrimStack g_UndoStack
        fn_RedoLast = True
    Else
        g_RedoStack.Add action
        If VBA.Len(VBA.Trim$(errorText)) = 0 Then errorText = "Unknown redo error."
        VBA.MsgBox "PrototypeNew: failed to redo action. " & errorText, VBA.vbExclamation, "PrototypeNew / Undo"
    End If
    g_IsReplaying = False
    private_RegisterExcelUndoRedo
    Exit Function

EH_REDO:
    g_IsReplaying = False
    If Not action Is Nothing Then
        On Error Resume Next
        g_RedoStack.Add action
        On Error GoTo 0
    End If
    private_RegisterExcelUndoRedo
End Function

' //
' // Internal
' //
Private Sub private_EnsureStorage()
    If g_UndoStack Is Nothing Then Set g_UndoStack = New Collection
    If g_RedoStack Is Nothing Then Set g_RedoStack = New Collection
    If g_MaxDepth <= 0 Then g_MaxDepth = DEFAULT_MAX_DEPTH
End Sub

Private Sub private_TrimStack(ByRef stackRef As Collection)
    private_EnsureStorage
    If stackRef Is Nothing Then Exit Sub

    Do While stackRef.Count > g_MaxDepth
        stackRef.Remove 1
    Loop
End Sub

Private Function private_PopLast(ByRef stackRef As Collection) As obj_IUndoAction
    Dim lastIndex As Long

    If stackRef Is Nothing Then Exit Function
    If stackRef.Count <= 0 Then Exit Function

    lastIndex = stackRef.Count
    Set private_PopLast = stackRef.Item(lastIndex)
    stackRef.Remove lastIndex
End Function

Private Sub private_RegisterExcelUndoRedo()
    ' История PrototypeNew доступна только через явные UI-команды.
    ' Не регистрируем Application.OnUndo/OnRepeat и не перехватываем Ctrl+Z,
    ' Ctrl+Y или Ctrl+Shift+Z: эти клавиши остаются в распоряжении Excel.
    private_EnsureStorage
    private_ClearExcelUndoHotkeys
End Sub

Private Sub private_ClearExcelUndoHotkeys()
    Application.OnKey "^z"
    Application.OnKey "^y"
    Application.OnKey "^+z"
End Sub

Private Function private_GetActionDebugLabel(ByVal action As obj_IUndoAction) As String
    Dim actionId As String
    Dim caption As String

    If action Is Nothing Then
        private_GetActionDebugLabel = "<nothing>"
        Exit Function
    End If

    On Error Resume Next
    actionId = VBA.Trim$(action.GetActionId())
    caption = VBA.Trim$(action.GetCaption())
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        private_GetActionDebugLabel = "<unavailable>"
        Exit Function
    End If
    On Error GoTo 0

    If VBA.Len(caption) = 0 Then caption = "<no-caption>"
    If VBA.Len(actionId) = 0 Then actionId = "<no-id>"
    private_GetActionDebugLabel = "id='" & VBA.Replace$(actionId, "'", "''") & "' caption='" & VBA.Replace$(caption, "'", "''") & "'"
End Function
