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
    reasonText = VBA.Trim$(reasonText)
    If VBA.Len(reasonText) = 0 Then reasonText = "runtime-initialize"

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "lifecycle:initialize-start reason='" & _
        VBA.Replace$(reasonText, "'", "''") & "'"
#End If

    rt_HotkeyRuntime.fn_BeginSession
    If Not ThisWorkbook.m_ResetWorkbookAndCreateMainPage( _
        reasonText, showErrorUi) Then Exit Function

    ' Main может уже быть активным, поэтому SheetActivate не гарантирован.
    ex_Core.fn_Diagnostic_ApplyLoggingPagePolicy "Main"

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "lifecycle:initialize-done reason='" & _
        VBA.Replace$(reasonText, "'", "''") & "'"
#End If
    fn_InitializeRuntime = True
End Function


Public Sub fn_DisposeRuntime( _
    ByVal isWorkbookClosing As Boolean, _
    Optional ByVal reasonText As String = "runtime-dispose" _
)
    fn_PrepareRuntimeDispose isWorkbookClosing, reasonText
    fn_DisposePreparedRuntime isWorkbookClosing, reasonText
End Sub


Public Sub fn_PrepareRuntimeDispose( _
    ByVal isWorkbookClosing As Boolean, _
    Optional ByVal reasonText As String = "runtime-dispose" _
)
    reasonText = VBA.Trim$(reasonText)
    If VBA.Len(reasonText) = 0 Then reasonText = "runtime-dispose"

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "lifecycle:prepare-start reason='" & _
        VBA.Replace$(reasonText, "'", "''") & "'"
#End If

    ' Сначала снимаем все OnTime-задачи, пока runtime ещё полностью работоспособен.
    ' Ошибка на этом этапе прекращает lifecycle до разрыва графа объектов.
    rt_CoreActions.fn_DisposeForLifecycle isWorkbookClosing
    ex_Core.fn_CancelDeferredTasks
    rt_Messaging.fn_CancelDeferredTasks
    rt_HotkeyRuntime.fn_BeginShutdown
    rt_UndoManager.fn_CancelGlobalCallbacks

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "lifecycle:prepare-done reason='" & _
        VBA.Replace$(reasonText, "'", "''") & "'"
#End If
End Sub


Public Sub fn_DisposePreparedRuntime( _
    ByVal isWorkbookClosing As Boolean, _
    Optional ByVal reasonText As String = "runtime-dispose" _
)
    reasonText = VBA.Trim$(reasonText)
    If VBA.Len(reasonText) = 0 Then reasonText = "runtime-dispose"

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "lifecycle:dispose-start reason='" & _
        VBA.Replace$(reasonText, "'", "''") & "'"
#End If

    ' Все внешние callbacks уже сняты фазой prepare. Здесь начинается только
    ' освобождение внутренних runtime-объектов.
    rt_Messaging.fn_Module_Dispose

    ' Внешние COM/ADO-ресурсы освобождаем до разрыва графа страниц.
    rt_WordExportRuntime.fn_Module_Dispose
    ex_ExternalExcelSqlEngine.fn_Module_Dispose

    ' PageManager владеет корневыми ссылками на page/controller graph.
    rt_PageManager.fn_Module_Dispose
    rt_UndoManager.fn_Module_Dispose

    ' Оставшиеся module-level registries могут удерживать Range, Shape и
    ' mode-controller даже после Page.Dispose.
    ex_ControlPartsRuntime.fn_Module_Dispose
    ex_ControlRefreshRuntime.fn_Module_Dispose
    ex_StylePipelineEngine.fn_Module_Dispose
    ex_LayoutControlFallbackRndr.fn_Module_Dispose
    ex_ShapeMetaRuntime.fn_Module_Dispose
    ex_SelectItemsSourceProviders.fn_Module_Dispose
    ex_CacheRuntime.fn_Module_Dispose
    rt_Bridge.fn_Module_Dispose

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "lifecycle:dispose-before-core reason='" & _
        VBA.Replace$(reasonText, "'", "''") & "'"
#End If
    ' ex_Core всегда последний: он владеет собственными очередями обновления и
    ' глобальными runtime sources, нужными предыдущим disposer-ам.
    ex_Core.fn_Module_Dispose
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "lifecycle:dispose-done reason='" & _
        VBA.Replace$(reasonText, "'", "''") & "'"
#End If
End Sub
