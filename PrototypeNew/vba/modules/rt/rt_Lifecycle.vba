Attribute VB_Name = "rt_Lifecycle"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True

Public Sub fn_Module_Dispose()
    ' Модуль не хранит состояние и только координирует lifecycle других
    ' runtime-компонентов. Вызывать fn_DisposeRuntime отсюда нельзя, иначе
    ' общий dispose-проход стал бы рекурсивным.
End Sub


Public Function fn_InitializeRuntime( _
    Optional ByVal reasonText As String = "runtime-initialize", _
    Optional ByVal showErrorUi As Boolean = True _
) As Boolean
    Dim activeCallee As String
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim eventsWereEnabled As Boolean

    reasonText = VBA.Trim$(reasonText)
    If VBA.Len(reasonText) = 0 Then reasonText = "runtime-initialize"
    On Error GoTo EH_INITIALIZE

#If LOGGING_DEBUG_ENABLED Then
    private_LogMethodEnter "rt_Lifecycle.fn_InitializeRuntime", reasonText
#End If

    eventsWereEnabled = Application.EnableEvents
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogEventInfo _
        "lifecycle:checkpoint method='rt_Lifecycle.fn_InitializeRuntime' " & _
        "step='capture-enable-events' value='" & _
        VBA.LCase$(VBA.CStr(eventsWereEnabled)) & "'"
#End If

    activeCallee = "rt_HotkeyRuntime.fn_BeginSession"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter "rt_Lifecycle.fn_InitializeRuntime", activeCallee
#End If
    rt_HotkeyRuntime.fn_BeginSession
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit "rt_Lifecycle.fn_InitializeRuntime", activeCallee
#End If

    activeCallee = "ThisWorkbook.m_ResetWorkbookAndCreateMainPage"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter "rt_Lifecycle.fn_InitializeRuntime", activeCallee
#End If
    If Not ThisWorkbook.m_ResetWorkbookAndCreateMainPage( _
        reasonText, showErrorUi) Then
#If LOGGING_DEBUG_ENABLED Then
        private_LogCallFailed _
            "rt_Lifecycle.fn_InitializeRuntime", activeCallee, _
            "result=false"
        private_LogMethodExit _
            "rt_Lifecycle.fn_InitializeRuntime", False
#End If
        Exit Function
    End If
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit "rt_Lifecycle.fn_InitializeRuntime", activeCallee
#End If

    ' Main может уже быть активным, поэтому SheetActivate не гарантирован.
    activeCallee = "ex_Core.fn_Diagnostic_ApplyLoggingPagePolicy"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter "rt_Lifecycle.fn_InitializeRuntime", activeCallee
#End If
    ex_Core.fn_Diagnostic_ApplyLoggingPagePolicy "Main"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit "rt_Lifecycle.fn_InitializeRuntime", activeCallee
#End If

    ' Успешно созданный runtime обязан принимать Workbook-события. Нельзя
    ' наследовать False от hot-import или от ранее аварийно прерванного кода:
    ' Shape.OnAction при этом работает, а SheetChange/SelectionChange — нет.
    activeCallee = "Application.EnableEvents"
    Application.EnableEvents = True
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogEventInfo _
        "lifecycle:checkpoint method='rt_Lifecycle.fn_InitializeRuntime' " & _
        "step='enable-workbook-events' previous='" & _
        VBA.LCase$(VBA.CStr(eventsWereEnabled)) & "' current='" & _
        VBA.LCase$(VBA.CStr(Application.EnableEvents)) & "'"
#End If

#If LOGGING_DEBUG_ENABLED Then
    private_LogMethodExit "rt_Lifecycle.fn_InitializeRuntime", True
#End If
    fn_InitializeRuntime = True
    Exit Function

EH_INITIALIZE:
    errorNumber = Err.Number
    errorDescription = Err.Description
#If LOGGING_DEBUG_ENABLED Then
    private_LogMethodError _
        "rt_Lifecycle.fn_InitializeRuntime", activeCallee, _
        errorNumber, errorDescription
#End If
    Err.Raise errorNumber, "rt_Lifecycle.fn_InitializeRuntime", _
        errorDescription
End Function


Public Sub fn_DisposeRuntime( _
    ByVal isWorkbookClosing As Boolean, _
    Optional ByVal reasonText As String = "runtime-dispose" _
)
    Dim activeCallee As String
    Dim errorNumber As Long
    Dim errorDescription As String

    reasonText = VBA.Trim$(reasonText)
    If VBA.Len(reasonText) = 0 Then reasonText = "runtime-dispose"
    On Error GoTo EH_DISPOSE_RUNTIME
#If LOGGING_DEBUG_ENABLED Then
    private_LogMethodEnter "rt_Lifecycle.fn_DisposeRuntime", reasonText
#End If

    activeCallee = "rt_Lifecycle.private_PrepareRuntimeDispose"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter "rt_Lifecycle.fn_DisposeRuntime", activeCallee
#End If
    private_PrepareRuntimeDispose isWorkbookClosing, reasonText
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit "rt_Lifecycle.fn_DisposeRuntime", activeCallee
#End If

    activeCallee = "rt_Lifecycle.private_DisposePreparedRuntime"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter "rt_Lifecycle.fn_DisposeRuntime", activeCallee
#End If
    private_DisposePreparedRuntime isWorkbookClosing, reasonText
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit "rt_Lifecycle.fn_DisposeRuntime", activeCallee
    private_LogMethodExit "rt_Lifecycle.fn_DisposeRuntime", True
#End If
    Exit Sub

EH_DISPOSE_RUNTIME:
    errorNumber = Err.Number
    errorDescription = Err.Description
#If LOGGING_DEBUG_ENABLED Then
    private_LogMethodError _
        "rt_Lifecycle.fn_DisposeRuntime", activeCallee, _
        errorNumber, errorDescription
#End If
    Err.Raise errorNumber, "rt_Lifecycle.fn_DisposeRuntime", _
        errorDescription
End Sub


Private Sub private_PrepareRuntimeDispose( _
    ByVal isWorkbookClosing As Boolean, _
    Optional ByVal reasonText As String = "runtime-dispose" _
)
    Dim activeCallee As String
    Dim errorNumber As Long
    Dim errorDescription As String

    reasonText = VBA.Trim$(reasonText)
    If VBA.Len(reasonText) = 0 Then reasonText = "runtime-dispose"
    On Error GoTo EH_PREPARE

#If LOGGING_DEBUG_ENABLED Then
    private_LogMethodEnter _
        "rt_Lifecycle.private_PrepareRuntimeDispose", reasonText
#End If

    ' Сначала снимаем все OnTime-задачи, пока runtime ещё полностью работоспособен.
    ' Ошибка на этом этапе прекращает lifecycle до разрыва графа объектов.
    activeCallee = "rt_CoreActions.fn_DisposeForLifecycle"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter _
        "rt_Lifecycle.private_PrepareRuntimeDispose", activeCallee
#End If
    rt_CoreActions.fn_DisposeForLifecycle isWorkbookClosing
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit _
        "rt_Lifecycle.private_PrepareRuntimeDispose", activeCallee
#End If

    activeCallee = "ex_Core.fn_CancelDeferredTasks"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter _
        "rt_Lifecycle.private_PrepareRuntimeDispose", activeCallee
#End If
    ex_Core.fn_CancelDeferredTasks
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit _
        "rt_Lifecycle.private_PrepareRuntimeDispose", activeCallee
#End If

    activeCallee = "rt_Messaging.fn_CancelDeferredTasks"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter _
        "rt_Lifecycle.private_PrepareRuntimeDispose", activeCallee
#End If
    rt_Messaging.fn_CancelDeferredTasks
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit _
        "rt_Lifecycle.private_PrepareRuntimeDispose", activeCallee
#End If

    activeCallee = "rt_HotkeyRuntime.fn_BeginShutdown"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter _
        "rt_Lifecycle.private_PrepareRuntimeDispose", activeCallee
#End If
    rt_HotkeyRuntime.fn_BeginShutdown
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit _
        "rt_Lifecycle.private_PrepareRuntimeDispose", activeCallee
#End If

    activeCallee = "rt_UndoManager.fn_CancelGlobalCallbacks"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter _
        "rt_Lifecycle.private_PrepareRuntimeDispose", activeCallee
#End If
    rt_UndoManager.fn_CancelGlobalCallbacks
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit _
        "rt_Lifecycle.private_PrepareRuntimeDispose", activeCallee
#End If

#If LOGGING_DEBUG_ENABLED Then
    private_LogMethodExit _
        "rt_Lifecycle.private_PrepareRuntimeDispose", True
#End If
    Exit Sub

EH_PREPARE:
    errorNumber = Err.Number
    errorDescription = Err.Description
#If LOGGING_DEBUG_ENABLED Then
    private_LogMethodError _
        "rt_Lifecycle.private_PrepareRuntimeDispose", activeCallee, _
        errorNumber, errorDescription
#End If
    Err.Raise errorNumber, _
        "rt_Lifecycle.private_PrepareRuntimeDispose", errorDescription
End Sub


Private Sub private_DisposePreparedRuntime( _
    ByVal isWorkbookClosing As Boolean, _
    Optional ByVal reasonText As String = "runtime-dispose" _
)
    Dim activeCallee As String
    Dim errorNumber As Long
    Dim errorDescription As String

    reasonText = VBA.Trim$(reasonText)
    If VBA.Len(reasonText) = 0 Then reasonText = "runtime-dispose"
    On Error GoTo EH_DISPOSE_PREPARED

#If LOGGING_DEBUG_ENABLED Then
    private_LogMethodEnter _
        "rt_Lifecycle.private_DisposePreparedRuntime", reasonText
#End If

    ' Все внешние callbacks уже сняты фазой prepare. Здесь начинается только
    ' освобождение внутренних runtime-объектов.
    activeCallee = "rt_Messaging.fn_Module_Dispose"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If
    rt_Messaging.fn_Module_Dispose
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If

    ' Внешние COM/ADO-ресурсы освобождаем до разрыва графа страниц.
    activeCallee = "rt_WordExportRuntime.fn_Module_Dispose"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If
    rt_WordExportRuntime.fn_Module_Dispose
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If

    activeCallee = "ex_ExternalExcelSqlEngine.fn_Module_Dispose"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If
    ex_ExternalExcelSqlEngine.fn_Module_Dispose
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If

    ' PageManager владеет корневыми ссылками на page/controller graph.
    activeCallee = "rt_PageManager.fn_Module_Dispose"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If
    rt_PageManager.fn_Module_Dispose
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If

    activeCallee = "rt_UndoManager.fn_Module_Dispose"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If
    rt_UndoManager.fn_Module_Dispose
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If

    ' Оставшиеся module-level registries могут удерживать Range, Shape и
    ' mode-controller даже после Page.Dispose.
    activeCallee = "ex_ControlPartsRuntime.fn_Module_Dispose"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If
    ex_ControlPartsRuntime.fn_Module_Dispose
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If

    activeCallee = "ex_ControlRefreshRuntime.fn_Module_Dispose"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If
    ex_ControlRefreshRuntime.fn_Module_Dispose
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If

    activeCallee = "ex_StylePipelineEngine.fn_Module_Dispose"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If
    ex_StylePipelineEngine.fn_Module_Dispose
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If

    activeCallee = "ex_LayoutControlFallbackRndr.fn_Module_Dispose"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If
    ex_LayoutControlFallbackRndr.fn_Module_Dispose
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If

    activeCallee = "ex_ShapeMetaRuntime.fn_Module_Dispose"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If
    ex_ShapeMetaRuntime.fn_Module_Dispose
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If

    activeCallee = "ex_SelectItemsSourceProviders.fn_Module_Dispose"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If
    ex_SelectItemsSourceProviders.fn_Module_Dispose
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If

    activeCallee = "ex_CacheRuntime.fn_Module_Dispose"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If
    ex_CacheRuntime.fn_Module_Dispose
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If

    activeCallee = "rt_Bridge.fn_Module_Dispose"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If
    rt_Bridge.fn_Module_Dispose
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "lifecycle:dispose-before-core reason='" & _
        VBA.Replace$(reasonText, "'", "''") & "'"
#End If
    ' ex_Core всегда последний: он владеет собственными очередями обновления и
    ' глобальными runtime sources, нужными предыдущим disposer-ам.
    activeCallee = "ex_Core.fn_Module_Dispose"
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallEnter _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
#End If
    ex_Core.fn_Module_Dispose
#If LOGGING_DEBUG_ENABLED Then
    private_LogCallExit _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee
    private_LogMethodExit _
        "rt_Lifecycle.private_DisposePreparedRuntime", True
#End If
    Exit Sub

EH_DISPOSE_PREPARED:
    errorNumber = Err.Number
    errorDescription = Err.Description
#If LOGGING_DEBUG_ENABLED Then
    private_LogMethodError _
        "rt_Lifecycle.private_DisposePreparedRuntime", activeCallee, _
        errorNumber, errorDescription
#End If
    Err.Raise errorNumber, _
        "rt_Lifecycle.private_DisposePreparedRuntime", errorDescription
End Sub


Private Sub private_LogMethodEnter( _
    ByVal methodName As String, _
    Optional ByVal reasonText As String = VBA.vbNullString _
)
    ex_Core.fn_Diagnostic_LogInfo _
        "lifecycle:method-enter method='" & _
        private_EscapeLogValue(methodName) & "' reason='" & _
        private_EscapeLogValue(reasonText) & "'"
End Sub


Private Sub private_LogMethodExit( _
    ByVal methodName As String, _
    ByVal resultValue As Boolean _
)
    ex_Core.fn_Diagnostic_LogInfo _
        "lifecycle:method-exit method='" & _
        private_EscapeLogValue(methodName) & "' result='" & _
        VBA.LCase$(VBA.CStr(resultValue)) & "'"
End Sub


Private Sub private_LogCallEnter( _
    ByVal callerName As String, _
    ByVal calleeName As String _
)
    ex_Core.fn_Diagnostic_LogInfo _
        "lifecycle:call-enter caller='" & _
        private_EscapeLogValue(callerName) & "' callee='" & _
        private_EscapeLogValue(calleeName) & "'"
End Sub


Private Sub private_LogCallExit( _
    ByVal callerName As String, _
    ByVal calleeName As String _
)
    ex_Core.fn_Diagnostic_LogInfo _
        "lifecycle:call-exit caller='" & _
        private_EscapeLogValue(callerName) & "' callee='" & _
        private_EscapeLogValue(calleeName) & "'"
End Sub


Private Sub private_LogCallFailed( _
    ByVal callerName As String, _
    ByVal calleeName As String, _
    ByVal failureText As String _
)
    ex_Core.fn_Diagnostic_LogError _
        "lifecycle:call-failed caller='" & _
        private_EscapeLogValue(callerName) & "' callee='" & _
        private_EscapeLogValue(calleeName) & "' reason='" & _
        private_EscapeLogValue(failureText) & "'"
End Sub


Private Sub private_LogMethodError( _
    ByVal methodName As String, _
    ByVal activeCallee As String, _
    ByVal errorNumber As Long, _
    ByVal errorDescription As String _
)
    ex_Core.fn_Diagnostic_LogError _
        "lifecycle:method-error method='" & _
        private_EscapeLogValue(methodName) & "' callee='" & _
        private_EscapeLogValue(activeCallee) & "' errNumber='" & _
        VBA.CStr(errorNumber) & "' err='" & _
        private_EscapeLogValue(errorDescription) & "'"
End Sub


Private Function private_EscapeLogValue(ByVal valueText As String) As String
    private_EscapeLogValue = VBA.Replace$(valueText, "'", "''")
End Function
