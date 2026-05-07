Attribute VB_Name = "rt_Bridge"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

' Стабильный мост между Shape.OnAction и rt_PageManager.
' Runtime-состояние контролов теперь принадлежит конкретной странице.

Private g_IsDispatchingClick As Boolean

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:rt_Bridge.fn_Module_Dispose"
#End If
End Sub

' //
' // API
' //
Public Sub fn_OnShapeClick()
    Dim callerShapeName As String
    Dim activeSheetObj As Object
    Dim ws As Worksheet
    Dim page As obj_IPage
    Dim dispatchOk As Boolean
    Dim wsName As String
    Dim wsCodeName As String

    ' Почему нужен bridge:
    ' - Excel Shape.OnAction принимает только имя макроса (строку),
    '   но не умеет вызывать method конкретного class instance.
    ' - Поэтому все shape клики сходятся в один модульный entrypoint,
    '   а дальше мы сами маршрутизируем к нужной странице/контролу.
    On Error GoTo EH_CLICK
    g_IsDispatchingClick = True

    ' 1) Получаем имя shape, по которому кликнули (Application.Caller).
    On Error Resume Next
    callerShapeName = VBA.CStr(Application.Caller)
    On Error GoTo EH_CLICK

    callerShapeName = VBA.Trim$(callerShapeName)
    If VBA.Len(callerShapeName) = 0 Then
        private_LogBridgeError "click-skip reason='caller-empty'"
        GoTo CleanExit
    End If

    ' 2) Определяем активный лист и находим page instance для этого листа.
    Set activeSheetObj = Application.ActiveSheet
    If Not TypeOf activeSheetObj Is Worksheet Then
        private_LogBridgeError "click-skip reason='active-sheet-not-worksheet' shape='" & private_EscapeForLog(callerShapeName) & "'"
        GoTo CleanExit
    End If

    Set ws = activeSheetObj
    wsName = VBA.Trim$(VBA.CStr(ws.Name))
    wsCodeName = VBA.Trim$(VBA.CStr(ws.CodeName))
    private_LogBridgeInfo "click-start shape='" & private_EscapeForLog(callerShapeName) & "' sheet='" & private_EscapeForLog(wsName) & "' codeName='" & private_EscapeForLog(wsCodeName) & "'"

    If Not rt_PageManager.fn_TryGetPageByWorksheet(ws, page) Then
        private_LogBridgeError "click-skip reason='page-not-found' shape='" & private_EscapeForLog(callerShapeName) & "' sheet='" & private_EscapeForLog(wsName) & "' codeName='" & private_EscapeForLog(wsCodeName) & "'"
        GoTo CleanExit
    End If

    ' 3) Передаем shapeName в page-level dispatcher.
    ' Дальше PageBase уже мапит shape -> (controlKey, methodName, arg)
    ' и вызывает method у конкретного VM объекта.
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
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "rt_Bridge: shape click dispatch failed: " & Err.Description
#End If
End Sub


Public Function fn_IsDispatchingClick() As Boolean
    fn_IsDispatchingClick = g_IsDispatchingClick
End Function


Public Function fn_RunCallback( _
    ByVal callbackRef As String, _
    Optional ByVal callbackContext As Object _
) As Boolean
    Dim callbackResult As Variant
    Dim qualifiedMacroRef As String
    Dim wbName As String

    callbackRef = VBA.Trim$(callbackRef)
    If VBA.Len(callbackRef) = 0 Then
        fn_RunCallback = True
        Exit Function
    End If

    On Error GoTo EH_RUN

    ' Если callback выглядит как имя метода без module/workbook,
    ' и есть object-context (например PageMainController) — вызываем method напрямую.
    If Not callbackContext Is Nothing Then
        If VBA.InStr(1, callbackRef, ".", VBA.vbBinaryCompare) = 0 And _
           VBA.InStr(1, callbackRef, "!", VBA.vbBinaryCompare) = 0 Then
            callbackResult = VBA.CallByName(callbackContext, callbackRef, VbMethod)
            If VBA.VarType(callbackResult) = vbBoolean Then
                fn_RunCallback = VBA.CBool(callbackResult)
            Else
                fn_RunCallback = True
            End If
            Exit Function
        End If
    End If

    qualifiedMacroRef = callbackRef
    If VBA.InStr(1, qualifiedMacroRef, "!", VBA.vbBinaryCompare) = 0 Then
        wbName = VBA.Replace$(ThisWorkbook.Name, "'", "''")
        qualifiedMacroRef = "'" & wbName & "'!" & qualifiedMacroRef
    End If

    Application.Run qualifiedMacroRef
    fn_RunCallback = True
    Exit Function

EH_RUN:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "rt_Bridge: failed to execute callback '" & callbackRef & "' (context='" & VBA.TypeName(callbackContext) & "'): " & Err.Description
#End If
End Function


Private Function private_EscapeForLog(ByVal valueText As String) As String
    private_EscapeForLog = VBA.Replace$(VBA.CStr(valueText), "'", "''")
End Function


Private Sub private_LogBridgeInfo(ByVal messageText As String)
    On Error Resume Next
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "bridge:" & VBA.Trim$(messageText)
#End If
    Err.Clear
    On Error GoTo 0
End Sub


Private Sub private_LogBridgeError(ByVal messageText As String)
    On Error Resume Next
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "bridge:" & VBA.Trim$(messageText)
#End If
    Err.Clear
    On Error GoTo 0
End Sub
