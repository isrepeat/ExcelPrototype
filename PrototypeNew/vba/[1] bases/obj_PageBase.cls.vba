VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PageBase"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
#Const LOGGING_ROUTE_VERBOSE_ENABLED = False

Private m_Worksheet As Worksheet
Private m_Page As obj_IPage
Private m_UiPath As String
Private m_LastRenderedUiPath As String
Private m_PageId As String
Private m_UiDom As Object
Private m_IsDisposed As Boolean
Private m_IsRendering As Boolean
Private m_ControlByKey As Object
Private m_LayoutContainerByName As Object
Private m_LayoutTagEntriesByTag As Object
Private m_RouteByShape As Object
Private m_RouteByCell As Object
Private m_RouteByHotkey As Object
Private m_SelectionHandlerContext As Object
Private m_SelectionHandlerMethod As String
Private m_PageRuntimeSources As obj_PageRuntimeSources
Private m_InlineRunEntries As Collection
' Кэш inline-профилей на уровне страницы: ключ = partName (banner/button/...).
' Почему в PageBase:
' 1) не создаем десятки одинаковых профилей в каждом VM/ViewItem;
' 2) все участники страницы используют один и тот же объект правил для partName;
' 3) lifecycle привязан к странице (очищается в Dispose вместе с runtime-реестрами).
Private m_InlineProfileByPart As Object

Private Const ROUTE_TYPE_CONTROL As String = "control"
Private Const UI_NS As String = "urn:excelprototype:profiles"
Private Const SHEET_UI_BASE_REL_PATH As String = "ui\"
Private Const SHEET_UI_FILE_SUFFIX As String = "UI.xml"
Private Const CONTROL_SNAPSHOT_ENTRY_ROOT As String = "controlSnapshot"
Private Const CONTROL_SNAPSHOT_ENTRY_NS As String = "urn:excelprototype:runtime-control-snapshot-entry:v1"
Private Const INLINE_TARGET_RANGE As String = "range"
Private Const INLINE_TARGET_SHAPE As String = "shape"

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Dispose False
    On Error GoTo 0
End Sub

' //
' // Properties
' //
Public Property Get Worksheet() As Worksheet
    Set Worksheet = m_Worksheet
End Property

Public Property Get Page() As obj_IPage
    Set Page = m_Page
End Property

Public Property Get UiPath() As String
    UiPath = m_UiPath
End Property

Public Property Get PageId() As String
    PageId = m_PageId
End Property

Public Property Get XmlDom() As Object
    Set XmlDom = m_UiDom
End Property

Public Property Get RuntimeSources() As obj_PageRuntimeSources
    If Not private_EnsureNotDisposed("RuntimeSources") Then Exit Property
    Set RuntimeSources = m_PageRuntimeSources
End Property

Public Property Get IsDisposed() As Boolean
    IsDisposed = m_IsDisposed
End Property

' //
' // API
' //
Public Function Initialize( _
    ByVal ws As Worksheet, _
    ByVal page As obj_IPage, _
    Optional ByVal uiPath As String = VBA.vbNullString, _
    Optional ByVal pageId As String = VBA.vbNullString _
) As Boolean
    Dim normalizedPageId As String
    Dim clearRange As Range

#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "startup:method-enter method='obj_PageBase.Initialize' pageId='" & _
        VBA.Replace$(pageId, "'", "''") & "'"
#End If
    If Not private_EnsureNotDisposed("Initialize") Then Exit Function

    normalizedPageId = VBA.LCase$(VBA.Trim$(pageId))
    If VBA.Len(normalizedPageId) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: page id is empty during Initialize."
#End If
        Exit Function
    End If

    Set m_Worksheet = ws
    Set m_Page = page
    m_UiPath = VBA.Trim$(uiPath)
    m_LastRenderedUiPath = VBA.vbNullString
    m_PageId = normalizedPageId
    Set m_UiDom = Nothing
    m_IsRendering = False
    Set m_PageRuntimeSources = New obj_PageRuntimeSources
    If Not m_PageRuntimeSources.Initialize(m_Page) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError _
            "startup:method-exit method='obj_PageBase.Initialize' " & _
            "result='false' step='obj_PageRuntimeSources.Initialize'"
#End If
        Exit Function
    End If

    ' Ранний дефолт: при инициализации страницы фиксируем текстовый формат текущего used-range.
    ' Это защитный baseline; основной повтор формата выполняется в runtime-clear перед каждым Render.
    On Error Resume Next
    Set clearRange = m_Worksheet.UsedRange
    If Not clearRange Is Nothing Then clearRange.NumberFormat = "@"
    Set clearRange = Nothing
    On Error GoTo 0

    Set m_InlineRunEntries = Nothing
    Set m_SelectionHandlerContext = Nothing
    m_SelectionHandlerMethod = VBA.vbNullString
    ' Реестр профилей стартует пустым и заполняется лениво по мере рендера.
    Set m_InlineProfileByPart = Nothing
    Call Me.ResetControlActions
    Initialize = Me.IsReady()
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "startup:method-exit method='obj_PageBase.Initialize' result='" & _
        VBA.LCase$(VBA.CStr(Initialize)) & "'"
#End If
End Function

Public Sub Dispose(Optional ByVal deleteWorksheet As Boolean = True)
    Dim ws As Worksheet
    Dim worksheetName As String
    Dim previousDisplayAlerts As Boolean
    Dim displayAlertsCaptured As Boolean
    Dim deleteErrorNumber As Long
    Dim deleteErrorDescription As String

#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "lifecycle:method-enter method='obj_PageBase.Dispose' pageId='" & _
        VBA.Replace$(m_PageId, "'", "''") & "' deleteWorksheet='" & _
        VBA.LCase$(VBA.CStr(deleteWorksheet)) & "'"
#End If
    If m_IsDisposed Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo _
            "lifecycle:method-exit method='obj_PageBase.Dispose' " & _
            "result='true' reason='already-disposed'"
#End If
        Exit Sub
    End If

    Call Me.ResetControlActions
    Set ws = m_Worksheet
    If Not ws Is Nothing Then
        On Error Resume Next
        worksheetName = VBA.Trim$(ws.Name)
        On Error GoTo 0
        If VBA.Len(worksheetName) > 0 Then
            Call ex_ControlPartsRuntime.fn_RemoveControlPartsByWorksheetName(worksheetName)
            ' Partial-render registry также имеет worksheet lifecycle. Иначе
            ' удалённая страница оставляет retained bounds до следующего полного
            ' render или выгрузки VBA-проекта.
            Call ex_ControlRefreshRuntime.fn_ResetRegisteredControlsByWorksheet( _
                worksheetName)
        End If
    End If
    Set m_Worksheet = Nothing
    Set m_Page = Nothing
    m_UiPath = VBA.vbNullString
    m_LastRenderedUiPath = VBA.vbNullString
    m_PageId = VBA.vbNullString
    Set m_UiDom = Nothing
    Set m_PageRuntimeSources = Nothing
    Set m_InlineRunEntries = Nothing
    Set m_SelectionHandlerContext = Nothing
    m_SelectionHandlerMethod = VBA.vbNullString
    ' Сбрасываем профильный кэш вместе со страницей (единый lifecycle PageBase).
    Set m_InlineProfileByPart = Nothing
    m_IsRendering = False
    m_IsDisposed = True

    If Not deleteWorksheet Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo _
            "lifecycle:method-exit method='obj_PageBase.Dispose' result='true'"
#End If
        Exit Sub
    End If
    If ws Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo _
            "lifecycle:method-exit method='obj_PageBase.Dispose' " & _
            "result='true' reason='worksheet-missing'"
#End If
        Exit Sub
    End If

    On Error GoTo EH_DELETE
    previousDisplayAlerts = Application.DisplayAlerts
    displayAlertsCaptured = True
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = previousDisplayAlerts
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo _
        "lifecycle:method-exit method='obj_PageBase.Dispose' result='true'"
#End If
    Exit Sub

EH_DELETE:
    deleteErrorNumber = Err.Number
    deleteErrorDescription = Err.Description
    If displayAlertsCaptured Then Application.DisplayAlerts = previousDisplayAlerts
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError _
        "lifecycle:method-error method='obj_PageBase.Dispose' " & _
        "step='Worksheet.Delete' err='" & _
        VBA.Replace$(deleteErrorDescription, "'", "''") & "'"
#End If
    Err.Raise deleteErrorNumber, "obj_PageBase.Dispose", _
        deleteErrorDescription
End Sub

Public Function IsReady() As Boolean
    If m_IsDisposed Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Page was disposed"
#End If
        Exit Function
    End If
    IsReady = Not m_Worksheet Is Nothing
End Function

Public Function GetPageBase() As obj_PageBase
    If Not private_EnsureNotDisposed("GetPageBase") Then Exit Function
    Set GetPageBase = Me
End Function

' Callstack[1]: rt_RestoreManager.m_RestorePageSnapshots -> serializablePage.TryDeserializeSnapshot(obj_PageMain) -> m_Base.ReadBaseSnapshotAttributes -> obj_PageBase.SetUiPath
' Callstack[2]: obj_PageBase.ReadBaseSnapshotAttributes -> obj_PageBase.SetUiPath
Public Sub SetUiPath(ByVal uiPath As String)
    If Not private_EnsureNotDisposed("SetUiPath") Then Exit Sub
    m_UiPath = VBA.Trim$(uiPath)
    Set m_UiDom = Nothing
End Sub

' Callstack[1]: ThisWorkbook.Workbook_Open -> ThisWorkbook.m_ResetWorkbookAndCreateMainPage -> private_ResetWorkbookAndCreateMainPage -> rt_PageManager.fn_RenderPageById -> rt_PageManager.fn_RenderPage -> obj_PageMain.obj_IPage_Render -> obj_PageBase.Render
' Callstack[2]: ex_Core.private_TryRecoverUiAfterUpdate -> ThisWorkbook.m_ResetWorkbookAndCreateMainPage -> private_ResetWorkbookAndCreateMainPage -> rt_PageManager.fn_RenderPageById -> rt_PageManager.fn_RenderPage -> obj_PageMain.obj_IPage_Render -> obj_PageBase.Render
' Callstack[3]: ex_Test.private_RenderWorksheetPage -> rt_PageManager.fn_RenderPageById -> rt_PageManager.fn_RenderPage -> obj_PageMain.obj_IPage_Render -> obj_PageBase.Render
' Callstack[4]: rt_PageManager.fn_RenderActivePage -> rt_PageManager.fn_RenderPage -> obj_PageMain.obj_IPage_Render -> obj_PageBase.Render
' Callstack[5]: ex_Test.fn_TEST_SetDemoConfigVariantA -> ex_HelpersSheet.fn_TryRerenderActivePage -> rt_PageManager.fn_RenderPage -> obj_PageMain.obj_IPage_Render -> obj_PageBase.Render
' Callstack[6]: ex_Test.fn_TEST_SetDemoConfigVariantB -> ex_HelpersSheet.fn_TryRerenderActivePage -> rt_PageManager.fn_RenderPage -> obj_PageMain.obj_IPage_Render -> obj_PageBase.Render
' Callstack[7]: ex_Test.private_TrySetItemsSource -> ex_Test.private_TryRerenderPage -> rt_PageManager.fn_RenderPage -> obj_PageMain.obj_IPage_Render -> obj_PageBase.Render
' Callstack[8]: ex_Test.private_TrySetObjectSource -> ex_Test.private_TryRerenderPage -> rt_PageManager.fn_RenderPage -> obj_PageMain.obj_IPage_Render -> obj_PageBase.Render
' Callstack[9]: ex_Test.private_TryRemoveObjectSource -> ex_Test.private_TryRerenderPage -> rt_PageManager.fn_RenderPage -> obj_PageMain.obj_IPage_Render -> obj_PageBase.Render
' Callstack[10]: rt_RestoreManager.m_RestorePageSnapshots(renderRestored:=True) -> rt_PageManager.fn_RenderPage -> obj_PageMain.obj_IPage_Render -> obj_PageBase.Render
' Callstack[11]: obj_PageMain.private_TryRerenderByDataChange -> rt_PageManager.fn_RenderPage -> obj_PageMain.obj_IPage_Render -> obj_PageBase.Render
Public Function Render() As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim app As Application
    Dim previousUiPath As String
    Dim resolvedUiPath As String
    Dim retainGeneratedShapes As Boolean
    Dim pageNode As Object
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation
    Dim prevStatusBar As Variant
    Dim applicationStateCaptured As Boolean
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim layoutRenderContext As obj_LayoutRenderContext
    Dim perfTotalStartedAt As Double, perfStageStartedAt As Double
    Dim domMs As Double, resetMs As Double, clearMs As Double
    Dim layoutMs As Double, numberFormatMs As Double, stylesMs As Double
    Dim inlineMs As Double, orphanShapesMs As Double

    perfTotalStartedAt = VBA.Timer

    If Not private_EnsureNotDisposed("Render") Then Exit Function
    If Not Me.IsReady() Then Exit Function

    If m_IsRendering Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError _
            "PageBase: render skipped because another render is active " & _
            "pageId='" & private_EscapeForLog(m_PageId) & "'."
#End If
        Exit Function
    End If

    Set ws = m_Worksheet
    Set wb = ws.Parent
    If wb Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: workbook is not specified."
#End If
        Exit Function
    End If

    ' Важно: retained-mode должен сравнивать с последним УСПЕШНО отрендеренным UI,
    ' а не с m_UiPath (он может быть уже заменен через UpdateUiPath до входа в Render).
    previousUiPath = VBA.Trim$(m_LastRenderedUiPath)
    ' Вычисляем фактический путь к разметке страницы для текущего рендера.
    resolvedUiPath = private_ResolvePageUiPath(m_UiPath)
    If VBA.Len(resolvedUiPath) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to resolve page UI path."
#End If
        Exit Function
    End If

    ' Загружаем и сохраняем DOM, чтобы стили и снапшоты работали с одним деревом.
    perfStageStartedAt = VBA.Timer
    Set m_UiDom = ex_XmlCore.fn_LoadDomByRelativePath( _
        wb, _
        resolvedUiPath, _
        "PrototypeNew: page UI file was not found: ", _
        "PrototypeNew: failed to parse page UI file: ", _
        UI_NS)
    If m_UiDom Is Nothing Then Exit Function
    domMs = private_PerfElapsedMs(perfStageStartedAt)

    Set pageNode = m_UiDom.selectSingleNode("/p:page")
    If pageNode Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: page UI root node <page> is missing."
#End If
        Exit Function
    End If

    m_UiPath = resolvedUiPath
    retainGeneratedShapes = private_ShouldRetainGeneratedShapes(previousUiPath, resolvedUiPath)
    On Error GoTo EH_RENDER
    m_IsRendering = True
    Set app = Application
    private_EnterFastRenderMode app, prevScreenUpdating, prevEnableEvents, _
        prevCalculation, prevStatusBar, applicationStateCaptured

    ' Сбрасываем runtime-реестры, чтобы не тянуть старые контролы/маршруты.
    perfStageStartedAt = VBA.Timer
    ex_ControlPartsRuntime.fn_ResetControlParts
    Me.ResetInlineRuns
    ' Bounds других уже отрендеренных страниц нужны для их будущего partial
    ' reflow. Сбрасываем только записи текущего worksheet.
    ex_ControlRefreshRuntime.fn_ResetRegisteredControlsByWorksheet ws.Name
    ex_StylePipelineEngine.fn_ResetLayoutBounds
    ex_LayoutControlFallbackRndr.fn_ResetControlFallbacks
    resetMs = private_PerfElapsedMs(perfStageStartedAt)

    perfStageStartedAt = VBA.Timer
    If Not Me.ResetControlActions(True) Then GoTo Cleanup
    If Not private_TryClearPageRuntime(Not retainGeneratedShapes) Then GoTo Cleanup
    clearMs = private_PerfElapsedMs(perfStageStartedAt)
    ' Один контекст на один проход: worksheet/workbook и seed-ы runtime ключей.
    Set layoutRenderContext = New obj_LayoutRenderContext
    If Not layoutRenderContext.Initialize(m_Page) Then GoTo Cleanup
    perfStageStartedAt = VBA.Timer
    If Not ex_XmlLayoutEngine.fn_RenderNode(layoutRenderContext, pageNode) Then GoTo Cleanup
    layoutMs = private_PerfElapsedMs(perfStageStartedAt)
    ' Layout уже собрал bounds всех контролов. Применяем текстовый формат одним
    ' batch COM-вызовом до общего style pass вместо одного вызова на control.
    perfStageStartedAt = VBA.Timer
    If Not ex_StylePipelineEngine.fn_ApplyTextNumberFormatToControlBounds(ws) Then GoTo Cleanup
    numberFormatMs = private_PerfElapsedMs(perfStageStartedAt)
    perfStageStartedAt = VBA.Timer
    If Not ex_StylePipelineEngine.fn_ApplyPageStyles(ws, m_UiDom) Then GoTo Cleanup
    stylesMs = private_PerfElapsedMs(perfStageStartedAt)
    ex_LayoutControlFallbackRndr.fn_ApplyPendingControlFallbacks ws
    perfStageStartedAt = VBA.Timer
    If Not Me.ApplyInlineRuns() Then GoTo Cleanup
    inlineMs = private_PerfElapsedMs(perfStageStartedAt)

    ' В retained-режиме глобально shape не удаляем до рендера.
    ' После рендера чистим только orphan-shape (контролы, которые больше не присутствуют в текущем layout).
    If retainGeneratedShapes Then
        perfStageStartedAt = VBA.Timer
        Call private_DeleteOrphanRuntimeShapesByControlRegistry(ws)
        orphanShapesMs = private_PerfElapsedMs(perfStageStartedAt)
    End If

    private_LogRuntimeInfo "render-bindings controls=" & VBA.CStr(private_GetDictionaryCount(m_ControlByKey)) & " shapeRoutes=" & VBA.CStr(private_GetDictionaryCount(m_RouteByShape)) & " cellRoutes=" & VBA.CStr(private_GetDictionaryCount(m_RouteByCell))

    Render = True
    m_LastRenderedUiPath = resolvedUiPath
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "perf:page-base-render totalMs='" & _
        VBA.Format$(private_PerfElapsedMs(perfTotalStartedAt), "0") & _
        "' domMs='" & VBA.Format$(domMs, "0") & _
        "' resetMs='" & VBA.Format$(resetMs, "0") & _
        "' clearMs='" & VBA.Format$(clearMs, "0") & _
        "' layoutMs='" & VBA.Format$(layoutMs, "0") & _
        "' numberFormatMs='" & VBA.Format$(numberFormatMs, "0") & _
        "' stylesMs='" & VBA.Format$(stylesMs, "0") & _
        "' inlineMs='" & VBA.Format$(inlineMs, "0") & _
        "' orphanShapesMs='" & VBA.Format$(orphanShapesMs, "0") & _
        "' controls='" & VBA.CStr(private_GetDictionaryCount(m_ControlByKey)) & _
        "' shapes='" & VBA.CStr(ws.Shapes.Count) & _
        "' retained='" & VBA.LCase$(VBA.CStr(retainGeneratedShapes)) & "'"
#End If

Cleanup:
    If applicationStateCaptured Then _
        private_LeaveFastRenderMode app, prevScreenUpdating, _
            prevEnableEvents, prevCalculation, prevStatusBar
    m_IsRendering = False
    Exit Function

EH_RENDER:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description

    If applicationStateCaptured Then _
        private_LeaveFastRenderMode app, prevScreenUpdating, _
            prevEnableEvents, prevCalculation, prevStatusBar
    m_IsRendering = False
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "PrototypeNew: render failed: [" & errSource & " #" & VBA.CStr(errNumber) & "] " & errDescription
#End If
End Function

' Частичный vertical reflow одного динамического контрола.
' В отличие от Render этот путь не сбрасывает page registries и не вызывает
' Configure/Render у соседей: готовый хвост страницы переносится как subtree.
'
' Жизненный цикл операции:
' 1) повторно measure только target по актуальному runtime source;
' 2) по retained layout-дереву вычислить patches зависимых siblings/предков;
' 3) физически перенести уже готовые диапазоны и Shapes;
' 4) заново Render + style только target;
' 5) зафиксировать новые bounds target и его предков.
'
' API рассчитан на изменение высоты существующего видимого control. Если узел
' был Collapsed, descriptor/bounds для него отсутствуют — нужно обновлять его
' именованный родитель через TryReflowLayoutContainer.
Public Function TryReflowControl(ByVal controlName As String) As Boolean
    Dim ws As Worksheet
    Dim controlNode As Object
    Dim renderCtx As obj_LayoutRenderContext
    Dim oldRowStart As Long, oldColStart As Long, oldRowEnd As Long, oldColEnd As Long
    Dim newSpanRows As Long, newSpanCols As Long
    Dim newRowEnd As Long, newColEnd As Long
    Dim rowDelta As Long
    Dim reflowPatches As Collection
    Dim ancestorUpdates As Collection
    Dim app As Application
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation
    Dim prevStatusBar As Variant
    Dim applicationStateCaptured As Boolean
    Dim escapedName As String
    Dim oldVisualScope As Range
    Dim selectionAreas As Collection
    Dim perfTotalStartedAt As Double
    Dim perfStageStartedAt As Double
    Dim planMs As Double, cleanupMs As Double, patchesMs As Double
    Dim reconcileMs As Double, baseStylesMs As Double, renderMs As Double
    Dim controlStylesMs As Double, commitMs As Double
    Dim shapesBefore As Long
    Dim patchCount As Long

    perfTotalStartedAt = VBA.Timer

    If Not private_EnsureNotDisposed("TryReflowControl") Then Exit Function
    If m_IsRendering Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError _
            "PageBase: control reflow skipped because another render is active " & _
            "pageId='" & private_EscapeForLog(m_PageId) & "'."
#End If
        Exit Function
    End If
    controlName = VBA.Trim$(controlName)
    If VBA.Len(controlName) = 0 Then Exit Function
    Set ws = m_Worksheet
    If ws Is Nothing Or m_UiDom Is Nothing Then Exit Function

    If Not ex_ControlRefreshRuntime.fn_TryGetControlRenderBounds( _
        controlName, ws.Name, oldRowStart, oldColStart, oldRowEnd, oldColEnd) Then Exit Function

    escapedName = ex_XmlCore.fn_XPathLiteral(controlName)
    Set controlNode = m_UiDom.selectSingleNode( _
        "/p:page//p:control[@name=" & escapedName & "] | " & _
        "/p:uiDefinition/p:layout//p:control[@name=" & escapedName & "]")
    If controlNode Is Nothing Then Exit Function

    Set renderCtx = New obj_LayoutRenderContext
    If Not renderCtx.Initialize(m_Page) Then Exit Function
    If Not ex_XmlLayoutEngine.fn_TryGetEffectiveNodeSpan( _
        renderCtx, controlNode, newSpanRows, newSpanCols) Then Exit Function
    If newSpanRows <= 0 Then newSpanRows = 1
    If newSpanCols <= 0 Then newSpanCols = oldColEnd - oldColStart + 1
    ' Текущий reflow patch решает vertical size changes. Изменение ширины
    ' потребовало бы отдельного horizontal propagation и column patches.
    If newSpanCols <> oldColEnd - oldColStart + 1 Then Exit Function

    newRowEnd = oldRowStart + newSpanRows - 1
    newColEnd = oldColStart + newSpanCols - 1
    rowDelta = newRowEnd - oldRowEnd

    perfStageStartedAt = VBA.Timer
    If Not ex_ControlRefreshRuntime.fn_TryBuildLayoutReflowPlan( _
        ws.Name, controlName, newSpanRows, reflowPatches, ancestorUpdates) Then Exit Function
    planMs = private_PerfElapsedMs(perfStageStartedAt)
    If Not reflowPatches Is Nothing Then patchCount = reflowPatches.Count
    shapesBefore = ws.Shapes.Count
    Set selectionAreas = private_CaptureSelectionAreas(ws)
    private_TranslateSelectionAreasByPatches selectionAreas, reflowPatches

    Set app = Application
    On Error GoTo EH_REFLOW
    m_IsRendering = True
    private_EnterFastRenderMode app, prevScreenUpdating, prevEnableEvents, _
        prevCalculation, prevStatusBar, applicationStateCaptured

    ' Исправляет уже созданные старой Copy-реализацией дубликаты одиночных
    ' кнопок. Выполняем даже при rowDelta = 0, чтобы обычный локальный refresh
    ' также мог вернуть страницу в консистентное состояние.
    private_DeleteDuplicateSingleButtonRuntimeShapes ws

    ' Старый target очищается до переноса. При shrink переносимый хвост займет
    ' освободившуюся нижнюю часть и не будет случайно очищен после translation.
    perfStageStartedAt = VBA.Timer
    If Not ex_ControlPartsRuntime.fn_TryGetControlVisualScope(ws, controlName, oldVisualScope) Then GoTo Cleanup
    If oldVisualScope Is Nothing Then
        Set oldVisualScope = ws.Range(ws.Cells(oldRowStart, oldColStart), ws.Cells(oldRowEnd, oldColEnd))
    End If
    oldVisualScope.Clear
    If Not ex_ControlPartsRuntime.fn_RemoveControlPartsByControl(ws.Name, controlName) Then GoTo Cleanup
    If Not ex_StylePipelineEngine.fn_RemoveLayoutBoundsByControl(ws.Name, controlName) Then GoTo Cleanup
    cleanupMs = private_PerfElapsedMs(perfStageStartedAt)

    perfStageStartedAt = VBA.Timer
    If Not reflowPatches Is Nothing Then
        If Not private_TryApplyLayoutReflowPatches(ws, reflowPatches) Then GoTo Cleanup
    End If
    patchesMs = private_PerfElapsedMs(perfStageStartedAt)

    ' Translate переносит содержимое subtree, но итоговая геометрия retained
    ' Shape должна определяться декларативным layout, а не его текущей позицией
    ' на листе. Это также исправляет ручное перетаскивание кнопки пользователем.
    perfStageStartedAt = VBA.Timer
    If Not private_TryReconcileSingleButtonRuntimeShapes(ws) Then GoTo Cleanup
    reconcileMs = private_PerfElapsedMs(perfStageStartedAt)

    perfStageStartedAt = VBA.Timer
    If Not ex_StylePipelineEngine.fn_ApplySheetBaseStylesToRange( _
        ws, m_UiDom, _
        ws.Range(ws.Cells(oldRowStart, oldColStart), ws.Cells(newRowEnd, newColEnd))) Then GoTo Cleanup
    baseStylesMs = private_PerfElapsedMs(perfStageStartedAt)

    perfStageStartedAt = VBA.Timer
    If Not ex_XmlLayoutEngine.fn_RenderNodeInBounds( _
        renderCtx, controlNode, oldRowStart, oldColStart, newRowEnd, newColEnd) Then GoTo Cleanup
    renderMs = private_PerfElapsedMs(perfStageStartedAt)
    perfStageStartedAt = VBA.Timer
    If Not ex_StylePipelineEngine.fn_ApplyControlPartStylesForControl( _
        ws, m_UiDom, controlName) Then GoTo Cleanup
    controlStylesMs = private_PerfElapsedMs(perfStageStartedAt)
    perfStageStartedAt = VBA.Timer
    If Not ex_ControlRefreshRuntime.fn_CommitLayoutReflowPlan( _
        ws.Name, controlName, newRowEnd, newColEnd, ancestorUpdates) Then GoTo Cleanup
    If Not private_CommitRuntimeAncestorUpdates(ancestorUpdates) Then GoTo Cleanup
    If Not ex_StylePipelineEngine.fn_CommitAncestorLayoutBounds( _
        ws.Name, ancestorUpdates) Then GoTo Cleanup
    commitMs = private_PerfElapsedMs(perfStageStartedAt)

    private_RestoreSelectionAreas ws, selectionAreas
    TryReflowControl = True

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "perf:control-reflow totalMs='" & _
        VBA.Format$(private_PerfElapsedMs(perfTotalStartedAt), "0") & _
        "' planMs='" & VBA.Format$(planMs, "0") & _
        "' cleanupMs='" & VBA.Format$(cleanupMs, "0") & _
        "' patchesMs='" & VBA.Format$(patchesMs, "0") & _
        "' reconcileMs='" & VBA.Format$(reconcileMs, "0") & _
        "' baseStylesMs='" & VBA.Format$(baseStylesMs, "0") & _
        "' renderMs='" & VBA.Format$(renderMs, "0") & _
        "' controlStylesMs='" & VBA.Format$(controlStylesMs, "0") & _
        "' commitMs='" & VBA.Format$(commitMs, "0") & _
        "' control='" & VBA.Replace$(controlName, "'", "''") & _
        "' oldRows='" & VBA.CStr(oldRowEnd - oldRowStart + 1) & _
        "' newRows='" & VBA.CStr(newSpanRows) & _
        "' rowDelta='" & VBA.CStr(rowDelta) & _
        "' patches='" & VBA.CStr(patchCount) & _
        "' shapesBefore='" & VBA.CStr(shapesBefore) & _
        "' shapesAfter='" & VBA.CStr(ws.Shapes.Count) & "'"
#End If

Cleanup:
    If applicationStateCaptured Then _
        private_LeaveFastRenderMode app, prevScreenUpdating, _
            prevEnableEvents, prevCalculation, prevStatusBar
    m_IsRendering = False
    Exit Function

EH_REFLOW:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "PageBase: partial reflow failed for control '" & _
        VBA.Replace$(controlName, "'", "''") & "': " & Err.Description
#End If
    Resume Cleanup
End Function

' Частично пересчитывает именованный layout-container целиком, но сохраняет все
' ветви страницы вне него. Внутри контейнера допускается изменение visibility,
' состава и размеров нескольких controls: subtree measure/render выполняется
' заново, а расположенный после контейнера хвост только транслируется.
'
' Имя здесь является публичным адресом reflow boundary. Выбирать слишком
' крупный container невыгодно (увеличится локальный render), слишком маленький
' нельзя, если его siblings совместно меняют visibility или геометрию.
Public Function TryReflowLayoutContainer(ByVal containerName As String) As Boolean
    Dim ws As Worksheet
    Dim containerNode As Object
    Dim controlNodes As Object
    Dim controlNode As Object
    Dim renderCtx As obj_LayoutRenderContext
    Dim oldRange As Range
    Dim clearRange As Range
    Dim rowStart As Long, colStart As Long
    Dim oldRowEnd As Long, oldColEnd As Long
    Dim newSpanRows As Long, newSpanCols As Long
    Dim newRowEnd As Long, newColEnd As Long
    Dim clearRowEnd As Long, clearColEnd As Long
    Dim controlName As String
    Dim reflowPatches As Collection
    Dim ancestorUpdates As Collection
    Dim app As Application
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation
    Dim prevStatusBar As Variant
    Dim applicationStateCaptured As Boolean
    Dim escapedName As String
    Dim selectionAreas As Collection

    If Not private_EnsureNotDisposed("TryReflowLayoutContainer") Then Exit Function
    If m_IsRendering Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError _
            "PageBase: container reflow skipped because another render is active " & _
            "pageId='" & private_EscapeForLog(m_PageId) & "'."
#End If
        Exit Function
    End If
    containerName = VBA.Trim$(containerName)
    If VBA.Len(containerName) = 0 Then Exit Function
    Set ws = m_Worksheet
    If ws Is Nothing Or m_UiDom Is Nothing Then Exit Function
    If Not Me.TryGetLayoutContainerRange(containerName, oldRange) Then Exit Function
    If oldRange Is Nothing Then Exit Function

    rowStart = oldRange.Row
    colStart = oldRange.Column
    oldRowEnd = rowStart + oldRange.Rows.Count - 1
    oldColEnd = colStart + oldRange.Columns.Count - 1
    escapedName = ex_XmlCore.fn_XPathLiteral(containerName)
    Set containerNode = m_UiDom.selectSingleNode( _
        "/p:page//p:stackPanel[@name=" & escapedName & "] | " & _
        "/p:uiDefinition/p:layout//p:stackPanel[@name=" & escapedName & "]")
    If containerNode Is Nothing Then Exit Function

    Set renderCtx = New obj_LayoutRenderContext
    If Not renderCtx.Initialize(m_Page) Then Exit Function
    If Not ex_XmlLayoutEngine.fn_TryGetEffectiveNodeSpan( _
        renderCtx, containerNode, newSpanRows, newSpanCols) Then Exit Function
    If newSpanRows <= 0 Or newSpanCols <= 0 Then Exit Function

    newRowEnd = rowStart + newSpanRows - 1
    newColEnd = colStart + newSpanCols - 1
    clearRowEnd = oldRowEnd
    If newRowEnd > clearRowEnd Then clearRowEnd = newRowEnd
    clearColEnd = oldColEnd
    If newColEnd > clearColEnd Then clearColEnd = newColEnd

    If Not ex_ControlRefreshRuntime.fn_TryBuildLayoutContainerReflowPlan( _
        ws.Name, containerName, newSpanRows, reflowPatches, ancestorUpdates) Then Exit Function
    Set selectionAreas = private_CaptureSelectionAreas(ws)
    private_TranslateSelectionAreasByPatches selectionAreas, reflowPatches

    Set app = Application
    On Error GoTo EH_CONTAINER_REFLOW
    m_IsRendering = True
    private_EnterFastRenderMode app, prevScreenUpdating, prevEnableEvents, _
        prevCalculation, prevStatusBar, applicationStateCaptured

    ' Удаляем только runtime metadata дочерних контролов. Shape-кнопки не
    ' удаляются: дочерний render переиспользует их по стабильным именам.
    If Not ex_StylePipelineEngine.fn_RemoveLayoutBoundsByNode( _
        ws.Name, "stackpanel", containerName) Then GoTo CleanupContainer
    private_RemoveInlineRunEntriesInRange oldRange
    Set controlNodes = containerNode.selectNodes(".//p:control[@name]")
    If Not controlNodes Is Nothing Then
        For Each controlNode In controlNodes
            controlName = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "name")))
            If VBA.Len(controlName) = 0 Then GoTo ContinueCleanupControl
            If Not ex_ControlPartsRuntime.fn_RemoveControlPartsByControl( _
                ws.Name, controlName) Then GoTo CleanupContainer
            If Not ex_StylePipelineEngine.fn_RemoveLayoutBoundsByControl( _
                ws.Name, controlName) Then GoTo CleanupContainer
ContinueCleanupControl:
        Next controlNode
    End If

    oldRange.Clear

    If Not reflowPatches Is Nothing Then
        If Not private_TryApplyLayoutReflowPatches(ws, reflowPatches) Then GoTo CleanupContainer
    End If
    ' Range-объект нельзя держать через Cut: Excel перенаправляет его на
    ' destination и последующий Clear стирает уже перемещённый sibling.
    ' Восстанавливаем scope по сохранённым числовым координатам после patches.
    Set clearRange = ws.Range( _
        ws.Cells(rowStart, colStart), _
        ws.Cells(clearRowEnd, clearColEnd))
    clearRange.Clear
    If Not ex_StylePipelineEngine.fn_ApplySheetBaseStylesToRange( _
        ws, m_UiDom, clearRange) Then GoTo CleanupContainer
    If Not ex_XmlLayoutEngine.fn_RenderNodeInBounds( _
        renderCtx, containerNode, rowStart, colStart, newRowEnd, newColEnd) Then GoTo CleanupContainer

    If Not controlNodes Is Nothing Then
        For Each controlNode In controlNodes
            controlName = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "name")))
            If VBA.Len(controlName) = 0 Then GoTo ContinueStyleControl
            If Not ex_StylePipelineEngine.fn_ApplyControlPartStylesForControl( _
                ws, m_UiDom, controlName) Then GoTo CleanupContainer
ContinueStyleControl:
        Next controlNode
    End If
    If Not ex_StylePipelineEngine.fn_ApplyRetainedControlStyles(ws, m_UiDom) Then GoTo CleanupContainer
    If Not Me.ApplyInlineRuns() Then GoTo CleanupContainer

    If Not ex_ControlRefreshRuntime.fn_CommitLayoutContainerReflowPlan( _
        ws.Name, containerName, newRowEnd, newColEnd, ancestorUpdates) Then GoTo CleanupContainer
    If Not private_CommitRuntimeAncestorUpdates(ancestorUpdates) Then GoTo CleanupContainer
    If Not ex_StylePipelineEngine.fn_CommitAncestorLayoutBounds( _
        ws.Name, ancestorUpdates) Then GoTo CleanupContainer
    If Not private_TryReconcileSingleButtonRuntimeShapes(ws) Then GoTo CleanupContainer

    private_RestoreSelectionAreas ws, selectionAreas
    TryReflowLayoutContainer = True

CleanupContainer:
    If applicationStateCaptured Then _
        private_LeaveFastRenderMode app, prevScreenUpdating, _
            prevEnableEvents, prevCalculation, prevStatusBar
    m_IsRendering = False
    Exit Function

EH_CONTAINER_REFLOW:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "PageBase: partial container reflow failed for '" & _
        VBA.Replace$(containerName, "'", "''") & "': " & Err.Description
#End If
    Resume CleanupContainer
End Function

Private Sub private_RemoveInlineRunEntriesInRange(ByVal targetRange As Range)
    Dim idx As Long
    Dim entry As Object
    Dim entryRange As Range
    Dim overlapRange As Range

    If targetRange Is Nothing Or m_InlineRunEntries Is Nothing Then Exit Sub
    For idx = m_InlineRunEntries.Count To 1 Step -1
        Set entry = m_InlineRunEntries(idx)
        If VBA.LCase$(VBA.CStr(entry("TargetType"))) <> INLINE_TARGET_RANGE Then GoTo ContinueEntry
        Set entryRange = Nothing
        Set overlapRange = Nothing
        On Error Resume Next
        Set entryRange = targetRange.Worksheet.Range(VBA.CStr(entry("CellAddress")))
        If Not entryRange Is Nothing Then Set overlapRange = Application.Intersect(entryRange, targetRange)
        On Error GoTo 0
        If Not overlapRange Is Nothing Then m_InlineRunEntries.Remove idx
ContinueEntry:
    Next idx
End Sub

' Callstack[1]: obj_BannerViewItem.Render -> m_PageBase.RegisterInlineRuns -> obj_PageBase.RegisterInlineRuns
Public Function RegisterInlineRuns( _
    ByVal targetRange As Range, _
    ByVal runs As Collection, _
    ByVal inlineTextProfile As obj_InlineTextProfile _
) As Boolean
    Dim firstCell As Range
    Dim entry As Object
    Dim targetKey As String

    If Not private_EnsureNotDisposed("RegisterInlineRuns") Then Exit Function

    If targetRange Is Nothing Then
        RegisterInlineRuns = True
        Exit Function
    End If
    If runs Is Nothing Then
        RegisterInlineRuns = True
        Exit Function
    End If
    If inlineTextProfile Is Nothing Then
        RegisterInlineRuns = True
        Exit Function
    End If

    Set firstCell = targetRange.Cells(1, 1)
    targetKey = VBA.LCase$(firstCell.Address(False, False))
    If m_InlineRunEntries Is Nothing Then Set m_InlineRunEntries = New Collection

    private_RemoveInlineRunEntriesByTarget INLINE_TARGET_RANGE, targetKey

    Set entry = VBA.CreateObject("Scripting.Dictionary")
    entry.CompareMode = 1
    entry("TargetType") = INLINE_TARGET_RANGE
    entry("TargetKey") = targetKey
    entry("CellAddress") = firstCell.Address(False, False)
    Set entry("Runs") = runs
    Set entry("InlineProfile") = inlineTextProfile

    m_InlineRunEntries.Add entry
    RegisterInlineRuns = True
End Function

Public Function TryResolveInlineTextByPart( _
    ByVal partName As String, _
    ByVal rawText As String, _
    ByRef outText As String, _
    ByRef outRuns As Collection _
) As Boolean
    Dim inlineTextProfile As obj_InlineTextProfile

    If Not private_EnsureNotDisposed("TryResolveInlineTextByPart") Then Exit Function

    ' partName = логический ключ части UI (например banner/button),
    ' по нему выбираем профиль правил inline-текста.
    partName = VBA.Trim$(partName)
    If VBA.Len(partName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: part name is empty for inline text resolve."
#End If
        Exit Function
    End If

    If Not Me.TryGetInlineTextProfile(partName, inlineTextProfile) Then Exit Function
    If Not inlineTextProfile.TryResolveInlineText(rawText, outText, outRuns) Then Exit Function

    TryResolveInlineTextByPart = True
End Function

Public Function TryGetInlineTextProfile( _
    ByVal partName As String, _
    ByRef outInlineProfile As obj_InlineTextProfile _
) As Boolean
    If Not private_EnsureNotDisposed("TryGetInlineTextProfile") Then Exit Function

    partName = VBA.Trim$(partName)
    If VBA.Len(partName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: part name is empty for inline text profile."
#End If
        Exit Function
    End If

    ' Возвращаем профиль из кэша или создаем новый при первом обращении.
    ' Это исключает дублирование одинаковых profile-объектов по разным VM/ViewItem.
    If Not private_TryGetInlineProfileByPart(partName, outInlineProfile) Then Exit Function
    TryGetInlineTextProfile = True
End Function

Public Function RegisterInlineRunsByPart( _
    ByVal targetRange As Range, _
    ByVal runs As Collection, _
    ByVal partName As String _
) As Boolean
    Dim inlineTextProfile As obj_InlineTextProfile

    If Not private_EnsureNotDisposed("RegisterInlineRunsByPart") Then Exit Function

    partName = VBA.Trim$(partName)
    If VBA.Len(partName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: part name is empty for range inline runs registration."
#End If
        Exit Function
    End If

    If Not Me.TryGetInlineTextProfile(partName, inlineTextProfile) Then Exit Function
    RegisterInlineRunsByPart = Me.RegisterInlineRuns(targetRange, runs, inlineTextProfile)
End Function

Public Function RegisterInlineRunsForShapeByPart( _
    ByVal targetShape As Shape, _
    ByVal runs As Collection, _
    ByVal partName As String _
) As Boolean
    Dim inlineTextProfile As obj_InlineTextProfile

    If Not private_EnsureNotDisposed("RegisterInlineRunsForShapeByPart") Then Exit Function

    partName = VBA.Trim$(partName)
    If VBA.Len(partName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: part name is empty for shape inline runs registration."
#End If
        Exit Function
    End If

    If Not Me.TryGetInlineTextProfile(partName, inlineTextProfile) Then Exit Function
    RegisterInlineRunsForShapeByPart = Me.RegisterInlineRunsForShape(targetShape, runs, inlineTextProfile)
End Function

Public Function RegisterInlineRunsForShape( _
    ByVal targetShape As Shape, _
    ByVal runs As Collection, _
    ByVal inlineTextProfile As obj_InlineTextProfile _
) As Boolean
    Dim entry As Object
    Dim targetKey As String

    If Not private_EnsureNotDisposed("RegisterInlineRunsForShape") Then Exit Function

    If targetShape Is Nothing Then
        RegisterInlineRunsForShape = True
        Exit Function
    End If
    If runs Is Nothing Then
        RegisterInlineRunsForShape = True
        Exit Function
    End If
    If inlineTextProfile Is Nothing Then
        RegisterInlineRunsForShape = True
        Exit Function
    End If

    targetKey = VBA.LCase$(VBA.Trim$(targetShape.Name))
    If VBA.Len(targetKey) = 0 Then
        RegisterInlineRunsForShape = True
        Exit Function
    End If

    If m_InlineRunEntries Is Nothing Then Set m_InlineRunEntries = New Collection
    private_RemoveInlineRunEntriesByTarget INLINE_TARGET_SHAPE, targetKey

    Set entry = VBA.CreateObject("Scripting.Dictionary")
    entry.CompareMode = 1
    entry("TargetType") = INLINE_TARGET_SHAPE
    entry("TargetKey") = targetKey
    entry("ShapeName") = targetShape.Name
    Set entry("Runs") = runs
    Set entry("InlineProfile") = inlineTextProfile

    m_InlineRunEntries.Add entry
    RegisterInlineRunsForShape = True
End Function

' Callstack[1]: obj_PageBase.Render -> obj_PageBase.ApplyInlineRuns
' Callstack[2]: ex_ControlRefreshRuntime.fn_TryRefreshStaticControl -> pageBase.ApplyInlineRuns -> obj_PageBase.ApplyInlineRuns
Public Function ApplyInlineRuns() As Boolean
    Dim entry As Object
    Dim targetCell As Range
    Dim targetShape As Shape
    Dim runs As Collection
    Dim inlineTextProfile As obj_InlineTextProfile
    Dim targetType As String

    If Not private_EnsureNotDisposed("ApplyInlineRuns") Then Exit Function
    If m_Worksheet Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: worksheet is not specified for inline runs."
#End If
        Exit Function
    End If

    If m_InlineRunEntries Is Nothing Then
        ApplyInlineRuns = True
        Exit Function
    End If

    ' Post-style проход: применяем уже зарегистрированные runs
    ' после того как базовые стили страницы/контролов выставлены.
    For Each entry In m_InlineRunEntries
        Set targetCell = Nothing
        Set targetShape = Nothing
        Set runs = Nothing
        Set inlineTextProfile = Nothing
        targetType = VBA.vbNullString

        On Error Resume Next
        targetType = VBA.LCase$(VBA.Trim$(VBA.CStr(entry("TargetType"))))
        If targetType = INLINE_TARGET_RANGE Then
            Set targetCell = m_Worksheet.Range(VBA.CStr(entry("CellAddress")))
        ElseIf targetType = INLINE_TARGET_SHAPE Then
            Set targetShape = m_Worksheet.Shapes(VBA.CStr(entry("ShapeName")))
        End If
        Set runs = entry("Runs")
        Set inlineTextProfile = entry("InlineProfile")
        On Error GoTo 0

        If runs Is Nothing Then GoTo ContinueEntry
        If inlineTextProfile Is Nothing Then GoTo ContinueEntry

        If targetType = INLINE_TARGET_RANGE Then
            If targetCell Is Nothing Then GoTo ContinueEntry
            inlineTextProfile.ApplyInlineRuns targetCell, runs
        ElseIf targetType = INLINE_TARGET_SHAPE Then
            If targetShape Is Nothing Then GoTo ContinueEntry
            inlineTextProfile.ApplyInlineRunsToShape targetShape, runs
        End If

ContinueEntry:
    Next entry

    ApplyInlineRuns = True
End Function

' Callstack[1]: obj_PageBase.Render -> obj_PageBase.ResetInlineRuns
' Callstack[2]: obj_PageBase.Clear -> obj_PageBase.ResetInlineRuns
Public Sub ResetInlineRuns()
    If Not private_EnsureNotDisposed("ResetInlineRuns") Then Exit Sub
    Set m_InlineRunEntries = Nothing
End Sub

Private Function private_TryGetInlineProfileByPart( _
    ByVal partName As String, _
    ByRef outInlineProfile As obj_InlineTextProfile _
) As Boolean
    Dim partKey As String
    Dim inlineTextProfile As obj_InlineTextProfile

    partKey = VBA.LCase$(VBA.Trim$(partName))
    If VBA.Len(partKey) = 0 Then Exit Function

    private_EnsureInlineProfileStorage
    ' Если профиль уже создан для partName, переиспользуем его.
    If m_InlineProfileByPart.Exists(partKey) Then
        Set outInlineProfile = m_InlineProfileByPart(partKey)
        private_TryGetInlineProfileByPart = True
        Exit Function
    End If

    ' Ленивое создание профиля: сейчас правила одинаковые,
    ' но архитектура позволяет отличать их по partName.
    Set inlineTextProfile = New obj_InlineTextProfile
    inlineTextProfile.PartName = partKey
    inlineTextProfile.InlineMarkersEnabled = True
    Set inlineTextProfile.StyleDoc = m_UiDom
    Set m_InlineProfileByPart(partKey) = inlineTextProfile
    Set outInlineProfile = inlineTextProfile
    private_TryGetInlineProfileByPart = True
End Function

Private Sub private_EnsureInlineProfileStorage()
    If Not m_InlineProfileByPart Is Nothing Then Exit Sub

    Set m_InlineProfileByPart = VBA.CreateObject("Scripting.Dictionary")
    m_InlineProfileByPart.CompareMode = 1
End Sub

Private Sub private_RemoveInlineRunEntriesByTarget(ByVal targetType As String, ByVal targetKey As String)
    Dim idx As Long
    Dim entry As Object
    Dim entryType As String
    Dim entryKey As String

    If m_InlineRunEntries Is Nothing Then Exit Sub

    targetType = VBA.LCase$(VBA.Trim$(targetType))
    targetKey = VBA.LCase$(VBA.Trim$(targetKey))
    If VBA.Len(targetType) = 0 Or VBA.Len(targetKey) = 0 Then Exit Sub

    For idx = m_InlineRunEntries.Count To 1 Step -1
        Set entry = m_InlineRunEntries(idx)

        entryType = VBA.vbNullString
        entryKey = VBA.vbNullString
        On Error Resume Next
        entryType = VBA.LCase$(VBA.Trim$(VBA.CStr(entry("TargetType"))))
        entryKey = VBA.LCase$(VBA.Trim$(VBA.CStr(entry("TargetKey"))))
        On Error GoTo 0

        If entryType = targetType And entryKey = targetKey Then
            m_InlineRunEntries.Remove idx
        End If
    Next idx
End Sub

' Callstack[1]: obj_PageMain.Clear -> obj_PageBase.Clear
Public Sub Clear()
    If Not private_EnsureNotDisposed("Clear") Then Exit Sub
    If m_Worksheet Is Nothing Then Exit Sub
    Call Me.ResetInlineRuns
    Call Me.ResetControlActions
    Call private_TryClearPageRuntime
End Sub

' Callstack[1]: rt_PageManager.fn_RenderPage -> page.Render -> obj_PageBase.Render -> ex_XmlLayoutEngine.fn_RenderNode -> ex_LayoutControlRenderer.fn_Render -> obj_ButtonControlVM.obj_IControl_Render -> m_Page.RegisterControl -> obj_PageBase.RegisterControl
' Callstack[2]: rt_PageManager.fn_RenderPage -> page.Render -> obj_PageBase.Render -> ex_XmlLayoutEngine.fn_RenderNode -> ex_LayoutControlRenderer.fn_Render -> obj_SelectControlVM.private_TryBindRuntimeRoutes -> m_Page.RegisterControl -> obj_PageBase.RegisterControl
Public Function RegisterControl(ByVal controlKey As String, ByVal iControl As Object) As Boolean
    If Not private_EnsureNotDisposed("RegisterControl") Then Exit Function
    controlKey = VBA.LCase$(VBA.Trim$(controlKey))
    If VBA.Len(controlKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: control key is empty."
#End If
        Exit Function
    End If
    If iControl Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: control VM is not specified for key '" & controlKey & "'."
#End If
        Exit Function
    End If

    private_EnsureStorage
    Set m_ControlByKey(controlKey) = iControl
#If LOGGING_ROUTE_VERBOSE_ENABLED Then
    private_LogRuntimeInfo "register-control key='" & private_EscapeForLog(controlKey) & "' controls=" & VBA.CStr(private_GetDictionaryCount(m_ControlByKey))
#End If
    RegisterControl = True
End Function

Public Function RegisterLayoutContainer( _
    ByVal containerName As String, _
    ByVal containerType As String, _
    ByVal sheetName As String, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long _
) As Boolean
    Dim entry As Object
    Dim containerKey As String

    If Not private_EnsureNotDisposed("RegisterLayoutContainer") Then Exit Function

    containerName = VBA.Trim$(containerName)
    containerType = VBA.LCase$(VBA.Trim$(containerType))
    sheetName = VBA.Trim$(sheetName)

    If VBA.Len(containerName) = 0 Then Exit Function
    If VBA.Len(containerType) = 0 Then Exit Function
    If VBA.Len(sheetName) = 0 Then Exit Function
    If rowStart <= 0 Or colStart <= 0 Then Exit Function
    If rowEnd < rowStart Or colEnd < colStart Then Exit Function

    private_EnsureStorage
    containerKey = VBA.LCase$(containerName)

    Set entry = VBA.CreateObject("Scripting.Dictionary")
    entry.CompareMode = 1
    entry("Name") = containerName
    entry("Type") = containerType
    entry("Sheet") = sheetName
    entry("RowStart") = VBA.CLng(rowStart)
    entry("ColStart") = VBA.CLng(colStart)
    entry("RowEnd") = VBA.CLng(rowEnd)
    entry("ColEnd") = VBA.CLng(colEnd)

    If m_LayoutContainerByName.Exists(containerKey) Then m_LayoutContainerByName.Remove containerKey
    m_LayoutContainerByName.Add containerKey, entry

    RegisterLayoutContainer = True
End Function

Public Function TryGetLayoutContainerRange( _
    ByVal containerName As String, _
    ByRef outRange As Range _
) As Boolean
    Dim containerKey As String
    Dim entry As Object
    Dim ws As Worksheet

    If Not private_EnsureNotDisposed("TryGetLayoutContainerRange") Then Exit Function
    Set outRange = Nothing

    containerKey = VBA.LCase$(VBA.Trim$(containerName))
    If VBA.Len(containerKey) = 0 Then Exit Function
    If m_LayoutContainerByName Is Nothing Then Exit Function
    If Not m_LayoutContainerByName.Exists(containerKey) Then Exit Function

    Set entry = m_LayoutContainerByName(containerKey)
    If entry Is Nothing Then Exit Function

    Set ws = m_Worksheet
    If ws Is Nothing Then Exit Function
    If VBA.StrComp(VBA.Trim$(VBA.CStr(entry("Sheet"))), ws.Name, VBA.vbTextCompare) <> 0 Then Exit Function

    On Error Resume Next
    Set outRange = ws.Range( _
        ws.Cells(VBA.CLng(entry("RowStart")), VBA.CLng(entry("ColStart"))), _
        ws.Cells(VBA.CLng(entry("RowEnd")), VBA.CLng(entry("ColEnd"))))
    On Error GoTo 0

    TryGetLayoutContainerRange = Not outRange Is Nothing
End Function

Public Function RegisterLayoutTags( _
    ByVal tagsText As String, _
    ByVal elementName As String, _
    ByVal elementType As String, _
    ByVal sheetName As String, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long, _
    Optional ByVal visibilityState As String = "visible" _
) As Boolean
    Dim tags As Collection
    Dim tagObj As Variant
    Dim tagText As String
    Dim tagKey As String
    Dim entries As Collection
    Dim entry As Object

    If Not private_EnsureNotDisposed("RegisterLayoutTags") Then Exit Function

    ' Layout tags are runtime metadata produced by XML layout rendering.
    ' They bind a logical tag from XML, for example tags="FIO;Person",
    ' to the actual worksheet bounds that the renderer assigned this turn.
    tagsText = VBA.Trim$(tagsText)
    elementName = VBA.Trim$(elementName)
    elementType = VBA.LCase$(VBA.Trim$(elementType))
    sheetName = VBA.Trim$(sheetName)
    visibilityState = VBA.LCase$(VBA.Trim$(visibilityState))

    If VBA.Len(tagsText) = 0 Then
        RegisterLayoutTags = True
        Exit Function
    End If
    If VBA.Len(elementType) = 0 Then Exit Function
    If VBA.Len(sheetName) = 0 Then Exit Function
    If rowStart <= 0 Or colStart <= 0 Then Exit Function
    If rowEnd < rowStart Or colEnd < colStart Then Exit Function
    If VBA.Len(visibilityState) = 0 Then visibilityState = "visible"

    Set tags = private_SplitTags(tagsText)
    If tags Is Nothing Then
        RegisterLayoutTags = True
        Exit Function
    End If

    private_EnsureStorage
    For Each tagObj In tags
        tagText = VBA.CStr(tagObj)
        tagKey = private_NormalizeLayoutTagText(tagText)
        If VBA.Len(tagKey) = 0 Then GoTo ContinueTag
        tagText = tagKey

        If m_LayoutTagEntriesByTag.Exists(tagKey) Then
            Set entries = m_LayoutTagEntriesByTag(tagKey)
        Else
            Set entries = New Collection
            Set m_LayoutTagEntriesByTag(tagKey) = entries
        End If

        Set entry = VBA.CreateObject("Scripting.Dictionary")
        entry.CompareMode = 1
        entry("Tag") = tagText
        entry("Name") = elementName
        entry("Type") = elementType
        entry("Sheet") = sheetName
        entry("RowStart") = VBA.CLng(rowStart)
        entry("ColStart") = VBA.CLng(colStart)
        entry("RowEnd") = VBA.CLng(rowEnd)
        entry("ColEnd") = VBA.CLng(colEnd)
        entry("Visibility") = visibilityState
        entries.Add entry

ContinueTag:
    Next tagObj

    RegisterLayoutTags = True
End Function

Public Function TryGetFirstLayoutTagRange( _
    ByVal tagText As String, _
    ByRef outRange As Range, _
    Optional ByVal visibilityStateFilter As String = "visible" _
) As Boolean
    Dim tagKey As String
    Dim entries As Collection
    Dim entryObj As Variant
    Dim entry As Object
    Dim ws As Worksheet
    Dim entryVisibility As String
    Dim filterText As String

    If Not private_EnsureNotDisposed("TryGetFirstLayoutTagRange") Then Exit Function
    Set outRange = Nothing

    ' Consumers such as LookupCandidates usually need the rendered visible
    ' control for a logical tag, not the old static XML/config order.
    tagKey = private_NormalizeLayoutTagText(tagText)
    If VBA.Len(tagKey) = 0 Then Exit Function
    If m_LayoutTagEntriesByTag Is Nothing Then Exit Function
    If Not m_LayoutTagEntriesByTag.Exists(tagKey) Then Exit Function

    Set ws = m_Worksheet
    If ws Is Nothing Then Exit Function

    filterText = VBA.LCase$(VBA.Trim$(visibilityStateFilter))
    Set entries = m_LayoutTagEntriesByTag(tagKey)
    If entries Is Nothing Then Exit Function

    For Each entryObj In entries
        If Not VBA.IsObject(entryObj) Then GoTo ContinueEntry
        Set entry = entryObj
        If entry Is Nothing Then GoTo ContinueEntry

        If VBA.StrComp(VBA.Trim$(VBA.CStr(entry("Sheet"))), ws.Name, VBA.vbTextCompare) <> 0 Then GoTo ContinueEntry
        entryVisibility = VBA.LCase$(VBA.Trim$(VBA.CStr(entry("Visibility"))))
        If VBA.Len(filterText) > 0 Then
            If VBA.StrComp(entryVisibility, filterText, VBA.vbTextCompare) <> 0 Then GoTo ContinueEntry
        End If

        On Error Resume Next
        Set outRange = ws.Range( _
            ws.Cells(VBA.CLng(entry("RowStart")), VBA.CLng(entry("ColStart"))), _
            ws.Cells(VBA.CLng(entry("RowEnd")), VBA.CLng(entry("ColEnd"))))
        On Error GoTo 0

        If Not outRange Is Nothing Then
            TryGetFirstLayoutTagRange = True
            Exit Function
        End If

ContinueEntry:
    Next entryObj
End Function

Public Function TryGetLayoutTagEntries( _
    ByVal tagText As String, _
    ByRef outEntries As Collection, _
    Optional ByVal visibilityStateFilter As String = VBA.vbNullString _
) As Boolean
    Dim tagKey As String
    Dim entries As Collection
    Dim entryObj As Variant
    Dim entry As Object
    Dim entryCopy As Object
    Dim entryVisibility As String
    Dim filterText As String
    Dim ws As Worksheet
    Dim keyObj As Variant

    If Not private_EnsureNotDisposed("TryGetLayoutTagEntries") Then Exit Function
    Set outEntries = Nothing

    tagKey = private_NormalizeLayoutTagText(tagText)
    If VBA.Len(tagKey) = 0 Then Exit Function
    If m_LayoutTagEntriesByTag Is Nothing Then Exit Function
    If Not m_LayoutTagEntriesByTag.Exists(tagKey) Then Exit Function

    Set ws = m_Worksheet
    If ws Is Nothing Then Exit Function

    filterText = VBA.LCase$(VBA.Trim$(visibilityStateFilter))
    Set entries = m_LayoutTagEntriesByTag(tagKey)
    If entries Is Nothing Then Exit Function

    Set outEntries = New Collection
    For Each entryObj In entries
        If Not VBA.IsObject(entryObj) Then GoTo ContinueEntry
        Set entry = entryObj
        If entry Is Nothing Then GoTo ContinueEntry

        If VBA.StrComp(VBA.Trim$(VBA.CStr(entry("Sheet"))), ws.Name, VBA.vbTextCompare) <> 0 Then GoTo ContinueEntry
        entryVisibility = VBA.LCase$(VBA.Trim$(VBA.CStr(entry("Visibility"))))
        If VBA.Len(filterText) > 0 Then
            If VBA.StrComp(entryVisibility, filterText, VBA.vbTextCompare) <> 0 Then GoTo ContinueEntry
        End If

        Set entryCopy = VBA.CreateObject("Scripting.Dictionary")
        entryCopy.CompareMode = 1
        For Each keyObj In entry.Keys
            If VBA.IsObject(entry(keyObj)) Then
                Set entryCopy(keyObj) = entry(keyObj)
            Else
                entryCopy(keyObj) = entry(keyObj)
            End If
        Next keyObj
        outEntries.Add entryCopy

ContinueEntry:
    Next entryObj

    TryGetLayoutTagEntries = True
End Function

Public Function TryGetLayoutTagEntriesInRange( _
    ByVal scopeRange As Range, _
    ByRef outEntries As Collection, _
    Optional ByVal visibilityStateFilter As String = "visible" _
) As Boolean
    Dim tagKeyObj As Variant
    Dim entries As Collection
    Dim entryObj As Variant
    Dim entry As Object
    Dim entryCopy As Object
    Dim entryRange As Range
    Dim intersectRange As Range
    Dim entryVisibility As String
    Dim filterText As String
    Dim ws As Worksheet
    Dim keyObj As Variant

    If Not private_EnsureNotDisposed("TryGetLayoutTagEntriesInRange") Then Exit Function
    Set outEntries = Nothing

    If scopeRange Is Nothing Then Exit Function
    Set ws = m_Worksheet
    If ws Is Nothing Then Exit Function

    Set outEntries = New Collection
    If m_LayoutTagEntriesByTag Is Nothing Then
        TryGetLayoutTagEntriesInRange = True
        Exit Function
    End If

    filterText = VBA.LCase$(VBA.Trim$(visibilityStateFilter))
    For Each tagKeyObj In m_LayoutTagEntriesByTag.Keys
        Set entries = m_LayoutTagEntriesByTag(tagKeyObj)
        If entries Is Nothing Then GoTo ContinueTag

        For Each entryObj In entries
            If Not VBA.IsObject(entryObj) Then GoTo ContinueEntry
            Set entry = entryObj
            If entry Is Nothing Then GoTo ContinueEntry

            If VBA.StrComp(VBA.Trim$(VBA.CStr(entry("Sheet"))), ws.Name, VBA.vbTextCompare) <> 0 Then GoTo ContinueEntry
            entryVisibility = VBA.LCase$(VBA.Trim$(VBA.CStr(entry("Visibility"))))
            If VBA.Len(filterText) > 0 Then
                If VBA.StrComp(entryVisibility, filterText, VBA.vbTextCompare) <> 0 Then GoTo ContinueEntry
            End If

            Set entryRange = Nothing
            Set intersectRange = Nothing
            On Error Resume Next
            Set entryRange = ws.Range( _
                ws.Cells(VBA.CLng(entry("RowStart")), VBA.CLng(entry("ColStart"))), _
                ws.Cells(VBA.CLng(entry("RowEnd")), VBA.CLng(entry("ColEnd"))))
            Set intersectRange = Application.Intersect(entryRange, scopeRange)
            On Error GoTo 0
            If entryRange Is Nothing Then GoTo ContinueEntry
            If intersectRange Is Nothing Then GoTo ContinueEntry

            Set entryCopy = VBA.CreateObject("Scripting.Dictionary")
            entryCopy.CompareMode = 1
            For Each keyObj In entry.Keys
                If VBA.IsObject(entry(keyObj)) Then
                    Set entryCopy(keyObj) = entry(keyObj)
                Else
                    entryCopy(keyObj) = entry(keyObj)
                End If
            Next keyObj
            outEntries.Add entryCopy

ContinueEntry:
        Next entryObj

ContinueTag:
    Next tagKeyObj

    TryGetLayoutTagEntriesInRange = True
End Function

' Callstack[1]: rt_PageManager.fn_RenderPage -> page.Render -> obj_PageBase.Render -> ex_XmlLayoutEngine.fn_RenderNode -> ex_LayoutControlRenderer.fn_Render -> obj_ButtonControlVM.private_TryBindRuntimeRoute -> m_Page.RegisterShapeRoute -> obj_PageBase.RegisterShapeRoute
' Callstack[2]: rt_PageManager.fn_RenderPage -> page.Render -> obj_PageBase.Render -> ex_XmlLayoutEngine.fn_RenderNode -> ex_LayoutControlRenderer.fn_Render -> obj_SelectControlVM.private_TryBindRuntimeRoutes -> m_Page.RegisterShapeRoute -> obj_PageBase.RegisterShapeRoute
Public Function RegisterShapeRoute( _
    ByVal shapeName As String, _
    ByVal controlKey As String, _
    ByVal methodName As String, _
    Optional ByVal hasArg As Boolean = False, _
    Optional ByVal argValue As Variant _
) As Boolean
    Dim shapeKey As String
    Dim entry As Object

    If Not private_EnsureNotDisposed("RegisterShapeRoute") Then Exit Function
    shapeKey = VBA.LCase$(VBA.Trim$(shapeName))
    controlKey = VBA.LCase$(VBA.Trim$(controlKey))
    methodName = VBA.Trim$(methodName)

    If VBA.Len(shapeKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: shape name is empty."
#End If
        Exit Function
    End If
    If VBA.Len(controlKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: control key is empty for shape '" & shapeName & "'."
#End If
        Exit Function
    End If
    If VBA.Len(methodName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: method name is empty for shape '" & shapeName & "'."
#End If
        Exit Function
    End If

    private_EnsureStorage
    If Not m_ControlByKey.Exists(controlKey) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: control '" & controlKey & "' is not registered for shape '" & shapeName & "'."
#End If
        Exit Function
    End If

    ' Запись маршрута описывает, как клик по shape вызвать действие контрола.
    Set entry = VBA.CreateObject("Scripting.Dictionary")
    entry.CompareMode = 1
    entry("RouteType") = ROUTE_TYPE_CONTROL
    entry("ControlKey") = controlKey
    entry("MethodName") = methodName
    entry("HasArg") = VBA.CBool(hasArg)
    If hasArg Then
        entry("ArgValue") = argValue
    Else
        entry("ArgValue") = Empty
    End If

    Set m_RouteByShape(shapeKey) = entry
#If LOGGING_ROUTE_VERBOSE_ENABLED Then
    private_LogRuntimeInfo "register-route shape='" & private_EscapeForLog(shapeName) & "' control='" & private_EscapeForLog(controlKey) & "' method='" & private_EscapeForLog(methodName) & "' routes=" & VBA.CStr(private_GetDictionaryCount(m_RouteByShape))
#End If
    RegisterShapeRoute = True
End Function

Public Function RegisterCellRoute( _
    ByVal cellAddress As String, _
    ByVal controlKey As String, _
    ByVal methodName As String, _
    Optional ByVal hasArg As Boolean = False, _
    Optional ByVal argValue As Variant _
) As Boolean
    Dim cellKey As String
    Dim entry As Object

    If Not private_EnsureNotDisposed("RegisterCellRoute") Then Exit Function
    cellKey = VBA.UCase$(VBA.Trim$(cellAddress))
    controlKey = VBA.LCase$(VBA.Trim$(controlKey))
    methodName = VBA.Trim$(methodName)

    If VBA.Len(cellKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: cell address is empty."
#End If
        Exit Function
    End If
    If VBA.Len(controlKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: control key is empty for cell '" & cellAddress & "'."
#End If
        Exit Function
    End If
    If VBA.Len(methodName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: method name is empty for cell '" & cellAddress & "'."
#End If
        Exit Function
    End If

    private_EnsureStorage
    If Not m_ControlByKey.Exists(controlKey) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: control '" & controlKey & "' is not registered for cell '" & cellAddress & "'."
#End If
        Exit Function
    End If

    Set entry = VBA.CreateObject("Scripting.Dictionary")
    entry.CompareMode = 1
    entry("RouteType") = ROUTE_TYPE_CONTROL
    entry("ControlKey") = controlKey
    entry("MethodName") = methodName
    entry("HasArg") = VBA.CBool(hasArg)
    If hasArg Then
        entry("ArgValue") = argValue
    Else
        entry("ArgValue") = Empty
    End If

    Set m_RouteByCell(cellKey) = entry
#If LOGGING_ROUTE_VERBOSE_ENABLED Then
    private_LogRuntimeInfo "register-route cell='" & private_EscapeForLog(cellAddress) & "' control='" & private_EscapeForLog(controlKey) & "' method='" & private_EscapeForLog(methodName) & "' routes=" & VBA.CStr(private_GetDictionaryCount(m_RouteByCell))
#End If
    RegisterCellRoute = True
End Function

Public Function RegisterHotkeyRoute( _
    ByVal hotkeyText As String, _
    ByVal controlKey As String, _
    ByVal methodName As String, _
    Optional ByVal hasArg As Boolean = False, _
    Optional ByVal argValue As Variant _
) As Boolean
    Dim hotkeyKey As String
    Dim entry As Object

    ' Совместимый путь для вызывающего кода, который передает человекочитаемый текст
    ' вроде CTRL+ENTER. Новые контролы могут сами парсить/валидировать ввод
    ' и вызывать RegisterHotkeyRouteByKey.
    If Not private_EnsureNotDisposed("RegisterHotkeyRoute") Then Exit Function
    hotkeyText = VBA.Trim$(hotkeyText)
    controlKey = VBA.LCase$(VBA.Trim$(controlKey))
    methodName = VBA.Trim$(methodName)

    If Not rt_HotkeyRuntime.fn_TryNormalizeHotkey(hotkeyText, hotkeyKey) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: unsupported hotkey '" & private_EscapeForLog(hotkeyText) & "'."
#End If
        VBA.MsgBox "PrototypeNew: unsupported hotkey '" & hotkeyText & "'. Use combinations like CTRL+ENTER, CTRL+SHIFT+Q, ALT+R.", VBA.vbExclamation, "PrototypeNew / Hotkeys"
        Exit Function
    End If
    If VBA.Len(controlKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: control key is empty for hotkey '" & private_EscapeForLog(hotkeyText) & "'."
#End If
        Exit Function
    End If
    If VBA.Len(methodName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: method name is empty for hotkey '" & private_EscapeForLog(hotkeyText) & "'."
#End If
        Exit Function
    End If

    private_EnsureStorage
    If Not m_ControlByKey.Exists(controlKey) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: control '" & controlKey & "' is not registered for hotkey '" & private_EscapeForLog(hotkeyText) & "'."
#End If
        Exit Function
    End If

    Set entry = VBA.CreateObject("Scripting.Dictionary")
    entry.CompareMode = 1
    entry("RouteType") = ROUTE_TYPE_CONTROL
    entry("ControlKey") = controlKey
    entry("MethodName") = methodName
    entry("HasArg") = VBA.CBool(hasArg)
    If hasArg Then
        entry("ArgValue") = argValue
    Else
        entry("ArgValue") = Empty
    End If

    If Not rt_HotkeyRuntime.fn_RegisterPageHotkey(m_PageId, hotkeyKey) Then Exit Function
    Set m_RouteByHotkey(hotkeyKey) = entry
#If LOGGING_ROUTE_VERBOSE_ENABLED Then
    private_LogRuntimeInfo "register-route hotkey='" & private_EscapeForLog(hotkeyKey) & "' control='" & private_EscapeForLog(controlKey) & "' method='" & private_EscapeForLog(methodName) & "' routes=" & VBA.CStr(private_GetDictionaryCount(m_RouteByHotkey))
#End If
    RegisterHotkeyRoute = True
End Function

Public Function RegisterHotkeyRouteByKey( _
    ByVal hotkeyKey As String, _
    ByVal controlKey As String, _
    ByVal methodName As String, _
    Optional ByVal hasArg As Boolean = False, _
    Optional ByVal argValue As Variant _
) As Boolean
    Dim entry As Object

    ' PageBase хранит локальный route страницы:
    '   OnKey token -> registered control key -> methodName(optional arg)
    ' rt_HotkeyRuntime хранит только глобальную Excel-привязку для того же OnKey token.
    ' При dispatch rt_Bridge сначала выбирает активную страницу, а эта map уже решает,
    ' что хоткей значит именно на этой странице.
    If Not private_EnsureNotDisposed("RegisterHotkeyRouteByKey") Then Exit Function
    controlKey = VBA.LCase$(VBA.Trim$(controlKey))
    methodName = VBA.Trim$(methodName)

    If VBA.Len(hotkeyKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: hotkey key is empty."
#End If
        Exit Function
    End If
    If VBA.Len(controlKey) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: control key is empty for hotkey key '" & private_EscapeForLog(hotkeyKey) & "'."
#End If
        Exit Function
    End If
    If VBA.Len(methodName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: method name is empty for hotkey key '" & private_EscapeForLog(hotkeyKey) & "'."
#End If
        Exit Function
    End If

    private_EnsureStorage
    If Not m_ControlByKey.Exists(controlKey) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: control '" & controlKey & "' is not registered for hotkey key '" & private_EscapeForLog(hotkeyKey) & "'."
#End If
        Exit Function
    End If

    Set entry = VBA.CreateObject("Scripting.Dictionary")
    entry.CompareMode = 1
    entry("RouteType") = ROUTE_TYPE_CONTROL
    entry("ControlKey") = controlKey
    entry("MethodName") = methodName
    entry("HasArg") = VBA.CBool(hasArg)
    If hasArg Then
        entry("ArgValue") = argValue
    Else
        entry("ArgValue") = Empty
    End If

    If Not rt_HotkeyRuntime.fn_RegisterPageHotkey(m_PageId, hotkeyKey) Then Exit Function
    Set m_RouteByHotkey(hotkeyKey) = entry
#If LOGGING_ROUTE_VERBOSE_ENABLED Then
    private_LogRuntimeInfo "register-route hotkey='" & private_EscapeForLog(hotkeyKey) & "' control='" & private_EscapeForLog(controlKey) & "' method='" & private_EscapeForLog(methodName) & "' routes=" & VBA.CStr(private_GetDictionaryCount(m_RouteByHotkey))
#End If
    RegisterHotkeyRouteByKey = True
End Function

Public Function ResetHotkeyRoutes() As Boolean
    If Not private_EnsureNotDisposed("ResetHotkeyRoutes") Then Exit Function
    ' Повторное применение таблицы хоткеев заменяет только hotkey routes этой страницы.
    ' Shape/cell routes не трогаются. Для активной страницы физические OnKey
    ' оставляем до финальной сверки, чтобы не снимать/назначать тот же набор заново.
    Set m_RouteByHotkey = Nothing
    rt_HotkeyRuntime.fn_UnregisterPageHotkeys m_PageId, True
    ResetHotkeyRoutes = True
End Function

' Callstack[1]: obj_PageMain.UnregisterControl -> obj_PageMain.obj_IPage_UnregisterControl -> obj_PageBase.UnregisterControl
' Callstack[2]: page.UnregisterControl(obj_IPage) -> obj_PageMain.obj_IPage_UnregisterControl -> obj_PageBase.UnregisterControl
Public Function UnregisterControl(ByVal controlKey As String) As Boolean
    Dim routeKey As Variant
    Dim routeEntry As Object
    Dim controlKeyNorm As String
    Dim routeKeysToRemove As Collection
    Dim cellRouteKeysToRemove As Collection
    Dim hotkeyRouteKeysToRemove As Collection
    Dim removeKey As Variant

    If Not private_EnsureNotDisposed("UnregisterControl") Then Exit Function
    controlKeyNorm = VBA.LCase$(VBA.Trim$(controlKey))
    If VBA.Len(controlKeyNorm) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: control key is empty."
#End If
        Exit Function
    End If

    private_EnsureStorage

    If m_ControlByKey.Exists(controlKeyNorm) Then
        m_ControlByKey.Remove controlKeyNorm
    End If

    Set routeKeysToRemove = New Collection
    Set cellRouteKeysToRemove = New Collection
    Set hotkeyRouteKeysToRemove = New Collection
    For Each routeKey In m_RouteByShape.Keys
        Set routeEntry = m_RouteByShape(routeKey)
        If VBA.LCase$(VBA.Trim$(VBA.CStr(routeEntry("ControlKey")))) = controlKeyNorm Then
            routeKeysToRemove.Add VBA.CStr(routeKey)
        End If
    Next routeKey

    For Each removeKey In routeKeysToRemove
        m_RouteByShape.Remove VBA.CStr(removeKey)
    Next removeKey

    For Each routeKey In m_RouteByCell.Keys
        Set routeEntry = m_RouteByCell(routeKey)
        If VBA.LCase$(VBA.Trim$(VBA.CStr(routeEntry("ControlKey")))) = controlKeyNorm Then
            cellRouteKeysToRemove.Add VBA.CStr(routeKey)
        End If
    Next routeKey

    For Each removeKey In cellRouteKeysToRemove
        m_RouteByCell.Remove VBA.CStr(removeKey)
    Next removeKey

    For Each routeKey In m_RouteByHotkey.Keys
        Set routeEntry = m_RouteByHotkey(routeKey)
        If VBA.LCase$(VBA.Trim$(VBA.CStr(routeEntry("ControlKey")))) = controlKeyNorm Then
            hotkeyRouteKeysToRemove.Add VBA.CStr(routeKey)
        End If
    Next routeKey

    For Each removeKey In hotkeyRouteKeysToRemove
        m_RouteByHotkey.Remove VBA.CStr(removeKey)
    Next removeKey

    UnregisterControl = True
End Function

' Callstack[1]: obj_PageBase.Initialize -> ResetControlActions
' Callstack[2]: obj_PageBase.Render -> ResetControlActions
' Callstack[3]: obj_PageBase.Clear -> ResetControlActions
' Callstack[4]: obj_PageBase.Dispose -> ResetControlActions
' Callstack[5]: obj_PageMain.ResetControlActions -> obj_PageMain.obj_IPage_ResetControlActions -> obj_PageBase.ResetControlActions
Public Function ResetControlActions(Optional ByVal keepActivePhysicalHotkeys As Boolean = False) As Boolean
    Dim key As Variant

    If Not private_EnsureNotDisposed("ResetControlActions") Then Exit Function

    If Not m_ControlByKey Is Nothing Then
        For Each key In m_ControlByKey.Keys
            Set m_ControlByKey(key) = Nothing
        Next key
    End If

    Set m_ControlByKey = Nothing
    Set m_LayoutContainerByName = Nothing
    Set m_LayoutTagEntriesByTag = Nothing
    Set m_RouteByShape = Nothing
    Set m_RouteByCell = Nothing
    Set m_RouteByHotkey = Nothing
    rt_HotkeyRuntime.fn_UnregisterPageHotkeys m_PageId, keepActivePhysicalHotkeys
    private_LogRuntimeInfo "reset-control-actions"
    ResetControlActions = True
End Function

' Callstack[1]: Shape.OnAction -> rt_Bridge.fn_OnShapeClick -> rt_PageManager.fn_TryGetPageByWorksheet -> page.DispatchShapeClick -> obj_PageMain.obj_IPage_DispatchShapeClick -> obj_PageBase.DispatchShapeClick
Public Function DispatchShapeClick(ByVal shapeName As String) As Boolean
    Dim routeEntry As Object
    Dim controlKey As String
    Dim methodName As String
    Dim hasArg As Boolean
    Dim argValue As Variant
    Dim iControl As Object
    Dim actionOk As Boolean
    Dim failureReason As String
    Dim invokeErrorText As String

    ' Центральная диспетчеризация клика внутри страницы.
    ' Зачем нужна:
    ' 1) Shape.OnAction может указывать только макрос, а не method class instance.
    ' 2) Здесь мы превращаем shapeName в runtime-route:
    '    shape -> controlKey -> methodName(+arg) и вызываем method у VM.
    If Not private_EnsureNotDisposed("DispatchShapeClick") Then Exit Function
    shapeName = VBA.Trim$(shapeName)
    If VBA.Len(shapeName) = 0 Then Exit Function

    private_LogRuntimeInfo "dispatch-click start shape='" & private_EscapeForLog(shapeName) & "' routes=" & VBA.CStr(private_GetDictionaryCount(m_RouteByShape)) & " controls=" & VBA.CStr(private_GetDictionaryCount(m_ControlByKey))

    ' Шаг A: ищем зарегистрированный маршрут для shape.
    If Not private_TryGetShapeRoute(shapeName, routeEntry, failureReason) Then
        private_LogRuntimeError "dispatch-click route-miss shape='" & private_EscapeForLog(shapeName) & "' reason='" & private_EscapeForLog(failureReason) & "'"
        Exit Function
    End If

    controlKey = VBA.LCase$(VBA.Trim$(VBA.CStr(routeEntry("ControlKey"))))
    methodName = VBA.Trim$(VBA.CStr(routeEntry("MethodName")))
    hasArg = VBA.CBool(routeEntry("HasArg"))
    If hasArg Then
        argValue = routeEntry("ArgValue")
    Else
        argValue = Empty
    End If

    ' Шаг B: по controlKey получаем живой control VM объект.
    If Not private_TryGetControl(controlKey, iControl, failureReason) Then
        private_LogRuntimeError "dispatch-click control-miss shape='" & private_EscapeForLog(shapeName) & "' control='" & private_EscapeForLog(controlKey) & "' reason='" & private_EscapeForLog(failureReason) & "'"
        private_RemoveShapeRoute shapeName
        Exit Function
    End If

    ' Контракт диспетчеризации клика:
    ' 1) опциональные глобальные хуки всех контролов
    ' 2) вызов целевого действия контрола
    If Not private_TryNotifyGlobalClick(controlKey, failureReason) Then
        private_LogRuntimeError "dispatch-click global-hook-blocked shape='" & private_EscapeForLog(shapeName) & "' control='" & private_EscapeForLog(controlKey) & "' reason='" & private_EscapeForLog(failureReason) & "'"
        Exit Function
    End If

    ' Шаг C: динамический invoke через CallByName (внутри private_TryInvokeControlAction).
    ' Именно это место дает возможность вызывать методы классов,
    ' а не только модульные макросы.
    If Not private_TryInvokeControlAction(iControl, methodName, hasArg, argValue, actionOk, invokeErrorText) Then
        private_LogRuntimeError "dispatch-click invoke-failed shape='" & private_EscapeForLog(shapeName) & "' control='" & private_EscapeForLog(controlKey) & "' method='" & private_EscapeForLog(methodName) & "' err='" & private_EscapeForLog(invokeErrorText) & "'"
        Exit Function
    End If

    If Not actionOk Then
        private_LogRuntimeError "dispatch-click action-returned-false shape='" & private_EscapeForLog(shapeName) & "' control='" & private_EscapeForLog(controlKey) & "' method='" & private_EscapeForLog(methodName) & "'"
        Exit Function
    End If

    private_LogRuntimeInfo "dispatch-click done shape='" & private_EscapeForLog(shapeName) & "' control='" & private_EscapeForLog(controlKey) & "' method='" & private_EscapeForLog(methodName) & "'"
    DispatchShapeClick = True
End Function

Public Function RegisterSelectionHandler( _
    ByVal callbackContext As Object, _
    ByVal methodName As String _
) As Boolean
    methodName = VBA.Trim$(methodName)
    If callbackContext Is Nothing Then Exit Function
    If VBA.Len(methodName) = 0 Then Exit Function
    Set m_SelectionHandlerContext = callbackContext
    m_SelectionHandlerMethod = methodName
    ex_Core.fn_Diagnostic_LogEventInfo _
        "event:selection-handler-registered pageId='" & _
        private_EscapeForLog(m_PageId) & "' sheet='" & _
        private_EscapeForLog(m_Worksheet.Name) & "' method='" & _
        private_EscapeForLog(methodName) & "'"
    RegisterSelectionHandler = True
End Function

Public Sub ClearSelectionHandler(ByVal callbackContext As Object)
    If callbackContext Is Nothing Then Exit Sub
    If m_SelectionHandlerContext Is Nothing Then Exit Sub
    If Not m_SelectionHandlerContext Is callbackContext Then Exit Sub
    Set m_SelectionHandlerContext = Nothing
    m_SelectionHandlerMethod = VBA.vbNullString
End Sub

Public Function DispatchSelectionChange(ByVal target As Range) As Boolean
    If Not private_EnsureNotDisposed("DispatchSelectionChange") Then Exit Function
    If target Is Nothing Then
        DispatchSelectionChange = True
        Exit Function
    End If
    If m_SelectionHandlerContext Is Nothing Or _
        VBA.Len(m_SelectionHandlerMethod) = 0 Then
        ex_Core.fn_Diagnostic_LogEventInfo _
            "event:selection-dispatch-skip pageId='" & _
            private_EscapeForLog(m_PageId) & _
            "' reason='handler-not-registered'"
        DispatchSelectionChange = True
        Exit Function
    End If
    DispatchSelectionChange = rt_Bridge.fn_RunCallback( _
        m_SelectionHandlerMethod, m_SelectionHandlerContext, target)
    If Not DispatchSelectionChange Then
        ex_Core.fn_Diagnostic_LogEventError _
            "event:selection-dispatch-failed pageId='" & _
            private_EscapeForLog(m_PageId) & "' method='" & _
            private_EscapeForLog(m_SelectionHandlerMethod) & _
            "' reason='callback-returned-false'"
    End If
End Function

Public Function DispatchSheetChange(ByVal target As Range) As Boolean
    Dim cell As Range
    Dim routeEntry As Object
    Dim cellKey As String
    Dim controlKey As String
    Dim methodName As String
    Dim hasArg As Boolean
    Dim argValue As Variant
    Dim iControl As Object
    Dim actionOk As Boolean
    Dim failureReason As String
    Dim invokeErrorText As String

    If Not private_EnsureNotDisposed("DispatchSheetChange") Then Exit Function
    If target Is Nothing Then
        DispatchSheetChange = True
        Exit Function
    End If
    If m_RouteByCell Is Nothing Then
        DispatchSheetChange = True
        Exit Function
    End If
    If m_RouteByCell.Count = 0 Then
        DispatchSheetChange = True
        Exit Function
    End If

    private_LogRuntimeInfo "dispatch-change start cells=" & VBA.CStr(target.Cells.CountLarge) & " cellRoutes=" & VBA.CStr(private_GetDictionaryCount(m_RouteByCell))

    ' Обрабатываем каждую ячейку отдельно, чтобы корректно поддержать multi-cell change
    ' и частичный матч только по тем адресам, где реально зарегистрирован input-route.
    For Each cell In target.Cells
        If cell Is Nothing Then GoTo ContinueCell

        cellKey = VBA.UCase$(VBA.Trim$(cell.Address(False, False)))
        If VBA.Len(cellKey) = 0 Then GoTo ContinueCell

        If Not private_TryGetCellRoute(cellKey, routeEntry, failureReason) Then GoTo ContinueCell

        controlKey = VBA.LCase$(VBA.Trim$(VBA.CStr(routeEntry("ControlKey"))))
        methodName = VBA.Trim$(VBA.CStr(routeEntry("MethodName")))
        hasArg = VBA.CBool(routeEntry("HasArg"))
        If hasArg Then
            argValue = routeEntry("ArgValue")
        Else
            argValue = Empty
        End If

        If Not private_TryGetControl(controlKey, iControl, failureReason) Then
            private_LogRuntimeError "dispatch-change control-miss cell='" & private_EscapeForLog(cellKey) & "' control='" & private_EscapeForLog(controlKey) & "' reason='" & private_EscapeForLog(failureReason) & "'"
            ' Если контрол уже удален/пересоздан, route считаем устаревшим и чистим его адресно.
            private_RemoveCellRoute cellKey
            GoTo ContinueCell
        End If

        If Not private_TryInvokeControlAction(iControl, methodName, hasArg, argValue, actionOk, invokeErrorText) Then
            private_LogRuntimeError "dispatch-change invoke-failed cell='" & private_EscapeForLog(cellKey) & "' control='" & private_EscapeForLog(controlKey) & "' method='" & private_EscapeForLog(methodName) & "' err='" & private_EscapeForLog(invokeErrorText) & "'"
            GoTo ContinueCell
        End If

        If Not actionOk Then
            private_LogRuntimeError "dispatch-change action-returned-false cell='" & private_EscapeForLog(cellKey) & "' control='" & private_EscapeForLog(controlKey) & "' method='" & private_EscapeForLog(methodName) & "'"
            GoTo ContinueCell
        End If

ContinueCell:
    Next cell

    private_LogRuntimeInfo "dispatch-change done"
    DispatchSheetChange = True
End Function

Public Function DispatchHotkey(ByVal hotkeyKey As String) As Boolean
    Dim routeEntry As Object
    Dim controlKey As String
    Dim methodName As String
    Dim hasArg As Boolean
    Dim argValue As Variant
    Dim iControl As Object
    Dim actionOk As Boolean
    Dim failureReason As String
    Dim invokeErrorText As String

    If Not private_EnsureNotDisposed("DispatchHotkey") Then Exit Function
    If VBA.Len(hotkeyKey) = 0 Then Exit Function

    ' Этот метод вызывается только после того, как rt_Bridge сопоставил
    ' Application.ActiveSheet с этим PageBase. Если такой же физический hotkey есть
    ' на другом листе, там будет вызван DispatchHotkey уже другой страницы.
    private_LogRuntimeInfo "dispatch-hotkey start hotkey='" & private_EscapeForLog(hotkeyKey) & "' routes=" & VBA.CStr(private_GetDictionaryCount(m_RouteByHotkey)) & " controls=" & VBA.CStr(private_GetDictionaryCount(m_ControlByKey))

    If Not private_TryGetHotkeyRoute(hotkeyKey, routeEntry, failureReason) Then
        private_LogRuntimeError "dispatch-hotkey route-miss hotkey='" & private_EscapeForLog(hotkeyKey) & "' reason='" & private_EscapeForLog(failureReason) & "'"
        Exit Function
    End If

    controlKey = VBA.LCase$(VBA.Trim$(VBA.CStr(routeEntry("ControlKey"))))
    methodName = VBA.Trim$(VBA.CStr(routeEntry("MethodName")))
    hasArg = VBA.CBool(routeEntry("HasArg"))
    If hasArg Then
        argValue = routeEntry("ArgValue")
    Else
        argValue = Empty
    End If

    If Not private_TryGetControl(controlKey, iControl, failureReason) Then
        private_LogRuntimeError "dispatch-hotkey control-miss hotkey='" & private_EscapeForLog(hotkeyKey) & "' control='" & private_EscapeForLog(controlKey) & "' reason='" & private_EscapeForLog(failureReason) & "'"
        private_RemoveHotkeyRoute hotkeyKey
        Exit Function
    End If

    If Not private_TryInvokeControlAction(iControl, methodName, hasArg, argValue, actionOk, invokeErrorText) Then
        private_LogRuntimeError "dispatch-hotkey invoke-failed hotkey='" & private_EscapeForLog(hotkeyKey) & "' control='" & private_EscapeForLog(controlKey) & "' method='" & private_EscapeForLog(methodName) & "' err='" & private_EscapeForLog(invokeErrorText) & "'"
        Exit Function
    End If

    If Not actionOk Then
        private_LogRuntimeError "dispatch-hotkey action-returned-false hotkey='" & private_EscapeForLog(hotkeyKey) & "' control='" & private_EscapeForLog(controlKey) & "' method='" & private_EscapeForLog(methodName) & "'"
        Exit Function
    End If

    private_LogRuntimeInfo "dispatch-hotkey done hotkey='" & private_EscapeForLog(hotkeyKey) & "' control='" & private_EscapeForLog(controlKey) & "' method='" & private_EscapeForLog(methodName) & "'"
    DispatchHotkey = True
End Function

Public Function TryCollectSerializableControlSnapshots(ByRef outSnapshots As Collection) As Boolean
    Dim key As Variant
    Dim iControl As Object
    Dim iSerializable As obj_ISerializable
    Dim typeRoot As String
    Dim payloadXml As String
    Dim controlKey As String
    Dim snapshotXml As String
    Dim pageKey As String

    If Not private_EnsureNotDisposed("TryCollectSerializableControlSnapshots") Then Exit Function
    Set outSnapshots = New Collection
    If m_ControlByKey Is Nothing Then
        TryCollectSerializableControlSnapshots = True
        Exit Function
    End If

    ' Снапшот: envelope(pageKey/controlKey/type) + XML payload контрола.
    pageKey = private_BuildPageKey()
    If VBA.Len(pageKey) = 0 Then Exit Function

    For Each key In m_ControlByKey.Keys
        Set iControl = m_ControlByKey(key)
        If iControl Is Nothing Then GoTo ContinueControl
        If Not private_TryCastSerializableControl(iControl, iSerializable) Then GoTo ContinueControl

        controlKey = VBA.LCase$(VBA.Trim$(VBA.CStr(key)))
        If VBA.Len(controlKey) = 0 Then GoTo ContinueControl

        typeRoot = VBA.LCase$(VBA.Trim$(iSerializable.GetSerializableTypeRoot()))
        If VBA.Len(typeRoot) = 0 Then GoTo ContinueControl

        payloadXml = VBA.vbNullString
        If Not iSerializable.TrySerializeSnapshot(payloadXml) Then GoTo ContinueControl
        If VBA.Len(VBA.Trim$(payloadXml)) = 0 Then GoTo ContinueControl

        snapshotXml = VBA.vbNullString
        If Not private_TrySerializeControlSnapshotEnvelope(pageKey, controlKey, typeRoot, payloadXml, snapshotXml) Then GoTo ContinueControl
        outSnapshots.Add snapshotXml

ContinueControl:
    Next key

    TryCollectSerializableControlSnapshots = True
End Function

' Callstack[1]: ThisWorkbook.Workbook_Open -> rt_RestoreManager.m_RestorePageSnapshots -> serializablePage.TryDeserializeSnapshot(obj_PageMain) -> obj_PageMain.TryDeserializeSnapshot -> obj_PageMain.obj_IPage_Render -> obj_PageMain.private_TryRestorePendingControlSnapshots -> m_Base.TryRestoreSerializableControlSnapshots -> obj_PageBase.TryRestoreSerializableControlSnapshots
' Callstack[2]: rt_CoreActions.fn_RerenderLastPageAfterUpdate -> rt_RestoreManager.m_RestorePageSnapshots -> serializablePage.TryDeserializeSnapshot(obj_PageMain) -> obj_PageMain.TryDeserializeSnapshot -> obj_PageMain.obj_IPage_Render -> obj_PageMain.private_TryRestorePendingControlSnapshots -> m_Base.TryRestoreSerializableControlSnapshots -> obj_PageBase.TryRestoreSerializableControlSnapshots
' Callstack[3]: obj_PageMain.private_TryRestorePendingControlSnapshots -> m_Base.TryRestoreSerializableControlSnapshots -> obj_PageBase.TryRestoreSerializableControlSnapshots
Public Function TryRestoreSerializableControlSnapshots(ByVal snapshots As Collection) As Boolean
    Dim item As Variant
    Dim snapshotXml As String
    Dim pageKey As String
    Dim controlKey As String
    Dim typeRoot As String
    Dim payloadXml As String
    Dim registeredControl As Object
    Dim iSerializable As obj_ISerializable

    If Not private_EnsureNotDisposed("TryRestoreSerializableControlSnapshots") Then Exit Function
    If snapshots Is Nothing Then
        TryRestoreSerializableControlSnapshots = True
        Exit Function
    End If

    ' В restore-фазе страница уже отрисована из актуального UI-контракта.
    ' Поэтому snapshot не должен пересоздавать controls или удалять свежие routes:
    ' он только достраивает runtime-state уже зарегистрированных controls.

    For Each item In snapshots
        snapshotXml = VBA.Trim$(VBA.CStr(item))
        If VBA.Len(snapshotXml) = 0 Then GoTo ContinueSnapshot

        pageKey = VBA.vbNullString
        controlKey = VBA.vbNullString
        typeRoot = VBA.vbNullString
        payloadXml = VBA.vbNullString
        If Not Me.TryDeserializeControlSnapshotEnvelope(snapshotXml, pageKey, controlKey, typeRoot, payloadXml) Then GoTo ContinueSnapshot
        If VBA.Len(typeRoot) = 0 Then GoTo ContinueSnapshot
        If VBA.Len(payloadXml) = 0 Then GoTo ContinueSnapshot

        Set registeredControl = Nothing
        If Not Me.TryGetRegisteredControlByKey(controlKey, registeredControl) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogWarning _
                "page-base:restore-control-snapshot skipped control-not-rendered key='" & _
                private_EscapeForLog(controlKey) & "' type='" & private_EscapeForLog(typeRoot) & "'"
#End If
            GoTo ContinueSnapshot
        End If
        If registeredControl Is Nothing Then GoTo ContinueSnapshot
        If Not private_TryCastSerializableControl(registeredControl, iSerializable) Then GoTo ContinueSnapshot
        If VBA.StrComp(VBA.LCase$(VBA.Trim$(iSerializable.GetSerializableTypeRoot())), typeRoot, VBA.vbTextCompare) <> 0 Then GoTo ContinueSnapshot

        If Not iSerializable.TryDeserializeSnapshot(payloadXml) Then GoTo ContinueSnapshot

ContinueSnapshot:
    Next item

    TryRestoreSerializableControlSnapshots = True
End Function

' Callstack[1]: obj_PageMain.TryGetRegisteredControls -> obj_PageMain.obj_IPage_TryGetRegisteredControls -> obj_PageBase.TryGetRegisteredControls
' Callstack[2]: page.TryGetRegisteredControls(obj_IPage) -> obj_PageMain.obj_IPage_TryGetRegisteredControls -> obj_PageBase.TryGetRegisteredControls
Public Function TryGetRegisteredControls(ByRef outControlsByKey As Object) As Boolean
    Dim key As Variant

    If Not private_EnsureNotDisposed("TryGetRegisteredControls") Then Exit Function
    Set outControlsByKey = VBA.CreateObject("Scripting.Dictionary")
    outControlsByKey.CompareMode = 1

    If m_ControlByKey Is Nothing Then
        TryGetRegisteredControls = True
        Exit Function
    End If

    For Each key In m_ControlByKey.Keys
        Set outControlsByKey(VBA.CStr(key)) = m_ControlByKey(key)
    Next key

    TryGetRegisteredControls = True
End Function

Public Function TryGetRegisteredControlByKey(ByVal controlKey As String, ByRef outControl As Object) As Boolean
    Dim reason As String

    If Not private_EnsureNotDisposed("TryGetRegisteredControlByKey") Then Exit Function
    Set outControl = Nothing

    controlKey = VBA.LCase$(VBA.Trim$(controlKey))
    If VBA.Len(controlKey) = 0 Then Exit Function

    If Not private_TryGetControl(controlKey, outControl, reason) Then Exit Function
    TryGetRegisteredControlByKey = True
End Function

Public Function TryGetRegisteredControlByName(ByVal controlName As String, ByRef outControl As Object) As Boolean
    Dim key As Variant
    Dim keyText As String
    Dim keyControlName As String
    Dim normalizedControlName As String
    Dim matchCount As Long

    If Not private_EnsureNotDisposed("TryGetRegisteredControlByName") Then Exit Function
    Set outControl = Nothing

    normalizedControlName = VBA.LCase$(VBA.Trim$(controlName))
    If VBA.Len(normalizedControlName) = 0 Then Exit Function
    If m_ControlByKey Is Nothing Then Exit Function

    For Each key In m_ControlByKey.Keys
        keyText = VBA.LCase$(VBA.Trim$(VBA.CStr(key)))
        keyControlName = private_ExtractControlNameFromControlKey(keyText)
        If VBA.StrComp(keyControlName, normalizedControlName, VBA.vbTextCompare) <> 0 Then GoTo ContinueControlByName

        Set outControl = m_ControlByKey(key)
        If outControl Is Nothing Then
            Set outControl = Nothing
            Exit Function
        End If

        matchCount = matchCount + 1
        If matchCount > 1 Then
            Set outControl = Nothing
            Exit Function
        End If

ContinueControlByName:
    Next key

    TryGetRegisteredControlByName = (matchCount = 1 And Not outControl Is Nothing)
End Function

' Callstack[1]: ThisWorkbook.Workbook_Open -> rt_RestoreManager.m_RestorePageSnapshots -> serializablePage.TryDeserializeSnapshot(obj_PageMain) -> obj_PageMain.obj_IPage_Render -> obj_PageMain.private_TryRestorePendingControlSnapshots -> obj_PageBase.TryRestoreSerializableControlSnapshots -> obj_PageBase.TryDeserializeControlSnapshotEnvelope
' Callstack[2]: rt_CoreActions.fn_RerenderLastPageAfterUpdate -> rt_RestoreManager.m_RestorePageSnapshots -> serializablePage.TryDeserializeSnapshot(obj_PageMain) -> obj_PageMain.obj_IPage_Render -> obj_PageMain.private_TryRestorePendingControlSnapshots -> obj_PageBase.TryRestoreSerializableControlSnapshots -> obj_PageBase.TryDeserializeControlSnapshotEnvelope
' Callstack[3]: obj_PageMain.private_TryRestorePendingControlSnapshots -> obj_PageBase.TryRestoreSerializableControlSnapshots -> obj_PageBase.TryDeserializeControlSnapshotEnvelope
Public Function TryDeserializeControlSnapshotEnvelope( _
    ByVal snapshotXml As String, _
    ByRef outPageKey As String, _
    ByRef outControlKey As String, _
    ByRef outTypeRoot As String, _
    ByRef outPayloadXml As String _
) As Boolean
    Dim dom As Object
    Dim rootNode As Object
    Dim payloadNode As Object

    outPageKey = VBA.vbNullString
    outControlKey = VBA.vbNullString
    outTypeRoot = VBA.vbNullString
    outPayloadXml = VBA.vbNullString

    snapshotXml = VBA.Trim$(snapshotXml)
    If VBA.Len(snapshotXml) = 0 Then Exit Function

    If Not ex_Core.fn_CustomXmlPartStore_TryLoadDomFromXml(snapshotXml, dom) Then Exit Function

    Set rootNode = dom.DocumentElement
    If rootNode Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: control snapshot root node is missing."
#End If
        Exit Function
    End If
    If VBA.StrComp(VBA.LCase$(VBA.CStr(rootNode.baseName)), CONTROL_SNAPSHOT_ENTRY_ROOT, VBA.vbTextCompare) <> 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: unexpected control snapshot root '" & VBA.CStr(rootNode.baseName) & "'."
#End If
        Exit Function
    End If

    outPageKey = VBA.LCase$(VBA.Trim$(VBA.CStr(rootNode.getAttribute("pageKey"))))
    outControlKey = VBA.LCase$(VBA.Trim$(VBA.CStr(rootNode.getAttribute("key"))))
    outTypeRoot = VBA.LCase$(VBA.Trim$(VBA.CStr(rootNode.getAttribute("type"))))

    Set payloadNode = rootNode.selectSingleNode("*[local-name()='payload']")
    If Not payloadNode Is Nothing Then
        outPayloadXml = VBA.CStr(payloadNode.Text)
    End If

    If VBA.Len(outTypeRoot) = 0 Then Exit Function
    If VBA.Len(outPayloadXml) = 0 Then Exit Function

    TryDeserializeControlSnapshotEnvelope = True
End Function

Public Function TryCreateSnapshotRoot( _
    ByVal rootName As String, _
    ByRef outDom As Object, _
    ByRef outRoot As Object _
) As Boolean
    Set outDom = Nothing
    Set outRoot = Nothing

    rootName = VBA.Trim$(rootName)
    If VBA.Len(rootName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: snapshot root is empty."
#End If
        Exit Function
    End If

    If Not ex_Core.fn_CustomXmlPartStore_TryCreateEmptyDom(rootName, "urn:excelprototype:serializable:page:v1", outDom) Then Exit Function
    Set outRoot = outDom.DocumentElement
    If outRoot Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: snapshot root node is missing."
#End If
        Exit Function
    End If

    TryCreateSnapshotRoot = True
End Function

' Callstack[1]: ThisWorkbook.Workbook_Open -> rt_RestoreManager.m_RestorePageSnapshots -> serializablePage.TryDeserializeSnapshot(obj_PageMain) -> m_Base.TryLoadSnapshotRoot -> obj_PageBase.TryLoadSnapshotRoot
' Callstack[2]: rt_CoreActions.fn_RerenderLastPageAfterUpdate -> rt_RestoreManager.m_RestorePageSnapshots -> serializablePage.TryDeserializeSnapshot(obj_PageMain) -> m_Base.TryLoadSnapshotRoot -> obj_PageBase.TryLoadSnapshotRoot
Public Function TryLoadSnapshotRoot( _
    ByVal snapshotXml As String, _
    ByVal expectedRootName As String, _
    ByRef outDom As Object, _
    ByRef outRoot As Object _
) As Boolean
    Set outDom = Nothing
    Set outRoot = Nothing

    snapshotXml = VBA.Trim$(snapshotXml)
    expectedRootName = VBA.LCase$(VBA.Trim$(expectedRootName))

    If VBA.Len(snapshotXml) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: snapshot XML is empty."
#End If
        Exit Function
    End If
    If VBA.Len(expectedRootName) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: expected root name is empty."
#End If
        Exit Function
    End If

    If Not ex_Core.fn_CustomXmlPartStore_TryLoadDomFromXml(snapshotXml, outDom) Then Exit Function
    Set outRoot = outDom.DocumentElement
    If outRoot Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: snapshot root node is missing."
#End If
        Exit Function
    End If
    If VBA.StrComp(VBA.LCase$(VBA.CStr(outRoot.baseName)), expectedRootName, VBA.vbTextCompare) <> 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: unexpected snapshot root '" & VBA.CStr(outRoot.baseName) & "'."
#End If
        Exit Function
    End If

    TryLoadSnapshotRoot = True
End Function

Public Sub WriteBaseSnapshotAttributes(ByVal targetNode As Object)
    If targetNode Is Nothing Then Exit Sub
    targetNode.setAttribute "uiPath", m_UiPath
End Sub

' Callstack[1]: ThisWorkbook.Workbook_Open -> rt_RestoreManager.m_RestorePageSnapshots -> serializablePage.TryDeserializeSnapshot(obj_PageMain) -> m_Base.ReadBaseSnapshotAttributes -> obj_PageBase.ReadBaseSnapshotAttributes
' Callstack[2]: rt_CoreActions.fn_RerenderLastPageAfterUpdate -> rt_RestoreManager.m_RestorePageSnapshots -> serializablePage.TryDeserializeSnapshot(obj_PageMain) -> m_Base.ReadBaseSnapshotAttributes -> obj_PageBase.ReadBaseSnapshotAttributes
' Callstack[3]: obj_PageBase.ReadBaseSnapshotAttributes -> obj_PageBase.SetUiPath
Public Sub ReadBaseSnapshotAttributes(ByVal sourceNode As Object)
    Dim restoredUiPath As String

    If sourceNode Is Nothing Then Exit Sub

    restoredUiPath = VBA.Trim$(VBA.CStr(sourceNode.getAttribute("uiPath")))
    If VBA.Len(restoredUiPath) = 0 Then Exit Sub

    Me.SetUiPath restoredUiPath
End Sub

' //
' // Internal
' //
Private Function private_TryClearPageRuntime(Optional ByVal deleteGeneratedShapes As Boolean = True) As Boolean
    Dim ws As Worksheet
    Dim clearRange As Range
    Dim i As Long

    If Not private_EnsureNotDisposed("private_TryClearPageRuntime") Then Exit Function

    Set ws = m_Worksheet
    If ws Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: worksheet is not specified."
#End If
        Exit Function
    End If

    ' Очищаем только runtime-артефакты (ячейки листа + shape с meta pn.control).
    On Error Resume Next
    Set clearRange = ws.UsedRange
    If Not clearRange Is Nothing Then
        clearRange.Clear
        ' Формат по умолчанию назначаем рано, до рендера контролов,
        ' чтобы Excel не автоконвертировал строки вида "01.05" в число.
        clearRange.NumberFormat = "@"
    End If
    On Error GoTo 0

    If deleteGeneratedShapes Then
        On Error Resume Next
        For i = ws.Shapes.Count To 1 Step -1
            If private_IsGeneratedRuntimeShape(ws.Shapes(i)) Then
                ws.Shapes(i).Delete
            End If
        Next i
        On Error GoTo 0
    End If

    private_TryClearPageRuntime = True
End Function

Private Function private_ShouldRetainGeneratedShapes(ByVal previousUiPath As String, ByVal currentUiPath As String) As Boolean
    previousUiPath = VBA.LCase$(VBA.Trim$(previousUiPath))
    currentUiPath = VBA.LCase$(VBA.Trim$(currentUiPath))
    If VBA.Len(previousUiPath) = 0 Or VBA.Len(currentUiPath) = 0 Then Exit Function

    ' Retained-режим используем только когда рендерим ту же самую страницу,
    ' чтобы безопасно переиспользовать runtime-shape по стабильным именам.
    private_ShouldRetainGeneratedShapes = (VBA.StrComp(previousUiPath, currentUiPath, VBA.vbBinaryCompare) = 0)
End Function

Private Function private_SplitTags(ByVal tagsText As String) As Collection
    Dim result As Collection
    Dim parts As Variant
    Dim idx As Long
    Dim tagText As String

    tagsText = VBA.Trim$(tagsText)
    If VBA.Len(tagsText) = 0 Then Exit Function

    Set result = New Collection
    parts = VBA.Split(tagsText, ";")
    For idx = LBound(parts) To UBound(parts)
        tagText = private_NormalizeLayoutTagText(VBA.CStr(parts(idx)))
        If VBA.Len(tagText) > 0 Then result.Add tagText
    Next idx

    If result.Count > 0 Then Set private_SplitTags = result
End Function

Private Function private_NormalizeLayoutTagText(ByVal tagText As String) As String
    tagText = VBA.CStr(tagText)
    tagText = VBA.Replace(tagText, VBA.vbCr, " ")
    tagText = VBA.Replace(tagText, VBA.vbLf, " ")
    tagText = VBA.Replace(tagText, VBA.vbTab, " ")
    private_NormalizeLayoutTagText = VBA.LCase$(VBA.Trim$(tagText))
End Function

Private Sub private_DeleteOrphanRuntimeShapesByControlRegistry(ByVal ws As Worksheet)
    Dim i As Long
    Dim shp As Shape
    Dim controlMeta As String
    Dim activeControlNames As Object
    Dim key As Variant
    Dim controlName As String

    If ws Is Nothing Then Exit Sub
    If m_ControlByKey Is Nothing Then Exit Sub

    Set activeControlNames = VBA.CreateObject("Scripting.Dictionary")
    activeControlNames.CompareMode = 1

    For Each key In m_ControlByKey.Keys
        controlName = private_ExtractControlNameFromControlKey(VBA.CStr(key))
        controlName = VBA.LCase$(VBA.Trim$(controlName))
        If VBA.Len(controlName) = 0 Then GoTo ContinueKey
        If Not activeControlNames.Exists(controlName) Then activeControlNames.Add controlName, True
ContinueKey:
    Next key

    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If shp Is Nothing Then GoTo ContinueShape

        controlMeta = VBA.LCase$(VBA.Trim$(ex_ShapeMetaRuntime.fn_GetShapeMetaValue(shp, "pn.control", VBA.vbNullString)))
        If VBA.Len(controlMeta) = 0 Then GoTo ContinueShape
        If Not activeControlNames.Exists(controlMeta) Then
            shp.Delete
        End If
ContinueShape:
    Next i
    On Error GoTo 0
End Sub

Private Function private_IsGeneratedRuntimeShape(ByVal shp As Shape) As Boolean
    Dim controlMeta As String

    If shp Is Nothing Then Exit Function

    controlMeta = VBA.Trim$(ex_ShapeMetaRuntime.fn_GetShapeMetaValue(shp, "pn.control", VBA.vbNullString))
    private_IsGeneratedRuntimeShape = (VBA.Len(controlMeta) > 0)
End Function

Private Sub private_EnsureStorage()
    If m_ControlByKey Is Nothing Then
        Set m_ControlByKey = VBA.CreateObject("Scripting.Dictionary")
        m_ControlByKey.CompareMode = 1
    End If

    If m_LayoutContainerByName Is Nothing Then
        Set m_LayoutContainerByName = VBA.CreateObject("Scripting.Dictionary")
        m_LayoutContainerByName.CompareMode = 1
    End If

    If m_LayoutTagEntriesByTag Is Nothing Then
        Set m_LayoutTagEntriesByTag = VBA.CreateObject("Scripting.Dictionary")
        m_LayoutTagEntriesByTag.CompareMode = 1
    End If

    If m_RouteByShape Is Nothing Then
        Set m_RouteByShape = VBA.CreateObject("Scripting.Dictionary")
        m_RouteByShape.CompareMode = 1
    End If

    If m_RouteByCell Is Nothing Then
        Set m_RouteByCell = VBA.CreateObject("Scripting.Dictionary")
        m_RouteByCell.CompareMode = 1
    End If

    If m_RouteByHotkey Is Nothing Then
        Set m_RouteByHotkey = VBA.CreateObject("Scripting.Dictionary")
        m_RouteByHotkey.CompareMode = 1
    End If
End Sub

Private Function private_ResolvePageUiPath(ByVal wsUiPath As String) As String
    Dim ws As Worksheet

    wsUiPath = VBA.Trim$(wsUiPath)
    If VBA.Len(wsUiPath) > 0 Then
        private_ResolvePageUiPath = wsUiPath
        Exit Function
    End If

    Set ws = m_Worksheet
    If ws Is Nothing Then Exit Function

    ' Конвенция по умолчанию: ui/<WorksheetName>UI.xml
    private_ResolvePageUiPath = SHEET_UI_BASE_REL_PATH & ws.Name & SHEET_UI_FILE_SUFFIX
End Function

Private Sub private_EnterFastRenderMode( _
    ByVal app As Application, _
    ByRef prevScreenUpdating As Boolean, _
    ByRef prevEnableEvents As Boolean, _
    ByRef prevCalculation As XlCalculation, _
    ByRef prevStatusBar As Variant, _
    ByRef outApplicationStateCaptured As Boolean _
)
    Dim activeStep As String
    Dim errorNumber As Long
    Dim errorDescription As String

    outApplicationStateCaptured = False
    If app Is Nothing Then Exit Sub
    On Error GoTo EH_ENTER_FAST_RENDER

    activeStep = "capture-ScreenUpdating"
    prevScreenUpdating = app.ScreenUpdating
    activeStep = "capture-EnableEvents"
    prevEnableEvents = app.EnableEvents
    activeStep = "capture-Calculation"
    prevCalculation = app.Calculation
    activeStep = "capture-StatusBar"
    prevStatusBar = app.StatusBar
    outApplicationStateCaptured = True

    activeStep = "disable-ScreenUpdating"
    app.ScreenUpdating = False
    activeStep = "disable-EnableEvents"
    app.EnableEvents = False
    activeStep = "set-CalculationManual"
    app.Calculation = xlCalculationManual
    activeStep = "set-StatusBar"
    app.StatusBar = "PrototypeNew: rendering UI..."
    Exit Sub

EH_ENTER_FAST_RENDER:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Raise errorNumber, _
        "obj_PageBase.private_EnterFastRenderMode", _
        "step='" & activeStep & "': " & errorDescription
End Sub

Private Sub private_LeaveFastRenderMode( _
    ByVal app As Application, _
    ByVal prevScreenUpdating As Boolean, _
    ByVal prevEnableEvents As Boolean, _
    ByVal prevCalculation As XlCalculation, _
    ByVal prevStatusBar As Variant _
)
    Dim cleanupErrors As String

    If app Is Nothing Then Exit Sub

    On Error Resume Next
    app.ScreenUpdating = prevScreenUpdating
    private_AppendFastRenderCleanupError _
        cleanupErrors, "ScreenUpdating", Err.Number, Err.Description
    Err.Clear
    app.EnableEvents = prevEnableEvents
    private_AppendFastRenderCleanupError _
        cleanupErrors, "EnableEvents", Err.Number, Err.Description
    Err.Clear
    app.Calculation = prevCalculation
    private_AppendFastRenderCleanupError _
        cleanupErrors, "Calculation", Err.Number, Err.Description
    Err.Clear
    app.StatusBar = prevStatusBar
    private_AppendFastRenderCleanupError _
        cleanupErrors, "StatusBar", Err.Number, Err.Description
    Err.Clear
    On Error GoTo 0

#If LOGGING_DEBUG_ENABLED Then
    If VBA.Len(cleanupErrors) > 0 Then
        ex_Core.fn_Diagnostic_LogError _
            "PageBase: fast render cleanup failed: " & cleanupErrors
    End If
#End If
End Sub


Private Sub private_AppendFastRenderCleanupError( _
    ByRef cleanupErrors As String, _
    ByVal propertyName As String, _
    ByVal errorNumber As Long, _
    ByVal errorDescription As String _
)
    If errorNumber = 0 Then Exit Sub
    If VBA.Len(cleanupErrors) > 0 Then cleanupErrors = cleanupErrors & "; "
    cleanupErrors = cleanupErrors & propertyName & " [" & _
        VBA.CStr(errorNumber) & "] " & errorDescription
End Sub

Private Function private_BuildPageKey() As String
    Dim wb As Workbook
    Dim codeNameValue As String

    If m_Worksheet Is Nothing Then Exit Function
    Set wb = m_Worksheet.Parent
    If wb Is Nothing Then Exit Function

    codeNameValue = VBA.Trim$(m_Worksheet.CodeName)
    If VBA.Len(codeNameValue) = 0 Then codeNameValue = VBA.Trim$(m_Worksheet.Name)
    If VBA.Len(codeNameValue) = 0 Then Exit Function

    ' Используем workbook + sheet codename как стабильный идентификатор при переименованиях.
    private_BuildPageKey = VBA.LCase$(VBA.Trim$(wb.Name) & "|" & codeNameValue)
End Function

Private Function private_BuildPageSheetKey() As String
    Dim wb As Workbook
    Dim sheetNameValue As String

    If m_Worksheet Is Nothing Then Exit Function
    Set wb = m_Worksheet.Parent
    If wb Is Nothing Then Exit Function

    sheetNameValue = VBA.Trim$(m_Worksheet.Name)
    If VBA.Len(sheetNameValue) = 0 Then Exit Function

    ' SheetName-key помогает матчингу snapshot-ов после full reload,
    ' когда CodeName листа может измениться.
    private_BuildPageSheetKey = VBA.LCase$(VBA.Trim$(wb.Name) & "|" & sheetNameValue)
End Function

Private Function private_TrySerializeControlSnapshotEnvelope( _
    ByVal pageKey As String, _
    ByVal controlKey As String, _
    ByVal typeRoot As String, _
    ByVal payloadXml As String, _
    ByRef outSnapshotXml As String _
) As Boolean
    Dim dom As Object
    Dim rootNode As Object
    Dim payloadNode As Object

    outSnapshotXml = VBA.vbNullString

    pageKey = VBA.LCase$(VBA.Trim$(pageKey))
    controlKey = VBA.LCase$(VBA.Trim$(controlKey))
    typeRoot = VBA.LCase$(VBA.Trim$(typeRoot))
    payloadXml = VBA.CStr(payloadXml)

    If VBA.Len(pageKey) = 0 Then Exit Function
    If VBA.Len(controlKey) = 0 Then Exit Function
    If VBA.Len(typeRoot) = 0 Then Exit Function
    If VBA.Len(payloadXml) = 0 Then Exit Function

    If Not ex_Core.fn_CustomXmlPartStore_TryCreateEmptyDom(CONTROL_SNAPSHOT_ENTRY_ROOT, CONTROL_SNAPSHOT_ENTRY_NS, dom) Then Exit Function

    Set rootNode = dom.DocumentElement
    If rootNode Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: control snapshot root node is missing."
#End If
        Exit Function
    End If

    rootNode.setAttribute "pageKey", pageKey
    rootNode.setAttribute "pageSheetKey", private_BuildPageSheetKey()
    rootNode.setAttribute "key", controlKey
    rootNode.setAttribute "type", typeRoot

    ' Храним payload как escaped-текст, чтобы избежать проблем со вложенными namespace.
    Set payloadNode = dom.createElement("payload")
    payloadNode.Text = payloadXml
    rootNode.appendChild payloadNode

    outSnapshotXml = VBA.CStr(dom.XML)
    private_TrySerializeControlSnapshotEnvelope = (VBA.Len(VBA.Trim$(outSnapshotXml)) > 0)
End Function

Private Function private_TryGetShapeRoute( _
    ByVal shapeName As String, _
    ByRef outEntry As Object, _
    Optional ByRef outReason As String = VBA.vbNullString _
) As Boolean
    Dim shapeKey As String

    outReason = VBA.vbNullString
    If m_RouteByShape Is Nothing Then
        outReason = "route-storage-empty"
        Exit Function
    End If

    shapeKey = VBA.LCase$(VBA.Trim$(shapeName))
    If VBA.Len(shapeKey) = 0 Then
        outReason = "shape-name-empty"
        Exit Function
    End If
    If Not m_RouteByShape.Exists(shapeKey) Then
        outReason = "route-not-found"
        Exit Function
    End If

    Set outEntry = m_RouteByShape(shapeKey)
    If outEntry Is Nothing Then
        outReason = "route-entry-empty"
        Exit Function
    End If

    private_TryGetShapeRoute = True
End Function

Private Sub private_RemoveShapeRoute(ByVal shapeName As String)
    Dim shapeKey As String

    If m_RouteByShape Is Nothing Then Exit Sub
    shapeKey = VBA.LCase$(VBA.Trim$(shapeName))
    If VBA.Len(shapeKey) = 0 Then Exit Sub
    If m_RouteByShape.Exists(shapeKey) Then
        m_RouteByShape.Remove shapeKey
    End If
End Sub

Private Function private_TryGetCellRoute( _
    ByVal cellAddress As String, _
    ByRef outEntry As Object, _
    Optional ByRef outReason As String = VBA.vbNullString _
) As Boolean
    Dim cellKey As String

    outReason = VBA.vbNullString
    If m_RouteByCell Is Nothing Then
        outReason = "route-storage-empty"
        Exit Function
    End If

    cellKey = VBA.UCase$(VBA.Trim$(cellAddress))
    If VBA.Len(cellKey) = 0 Then
        outReason = "cell-address-empty"
        Exit Function
    End If
    If Not m_RouteByCell.Exists(cellKey) Then
        outReason = "route-not-found"
        Exit Function
    End If

    Set outEntry = m_RouteByCell(cellKey)
    If outEntry Is Nothing Then
        outReason = "route-entry-empty"
        Exit Function
    End If

    private_TryGetCellRoute = True
End Function

Private Sub private_RemoveCellRoute(ByVal cellAddress As String)
    Dim cellKey As String

    If m_RouteByCell Is Nothing Then Exit Sub
    cellKey = VBA.UCase$(VBA.Trim$(cellAddress))
    If VBA.Len(cellKey) = 0 Then Exit Sub
    If m_RouteByCell.Exists(cellKey) Then
        m_RouteByCell.Remove cellKey
    End If
End Sub

Private Function private_TryGetHotkeyRoute( _
    ByVal hotkeyKey As String, _
    ByRef outEntry As Object, _
    Optional ByRef outReason As String = VBA.vbNullString _
) As Boolean
    outReason = VBA.vbNullString
    If m_RouteByHotkey Is Nothing Then
        outReason = "route-storage-empty"
        Exit Function
    End If

    If VBA.Len(hotkeyKey) = 0 Then
        outReason = "hotkey-empty"
        Exit Function
    End If
    If Not m_RouteByHotkey.Exists(hotkeyKey) Then
        outReason = "route-not-found"
        Exit Function
    End If

    Set outEntry = m_RouteByHotkey(hotkeyKey)
    If outEntry Is Nothing Then
        outReason = "route-entry-empty"
        Exit Function
    End If

    private_TryGetHotkeyRoute = True
End Function

Private Sub private_RemoveHotkeyRoute(ByVal hotkeyKey As String)
    If m_RouteByHotkey Is Nothing Then Exit Sub
    If VBA.Len(hotkeyKey) = 0 Then Exit Sub
    If m_RouteByHotkey.Exists(hotkeyKey) Then
        m_RouteByHotkey.Remove hotkeyKey
    End If
End Sub

Private Function private_TryGetControl( _
    ByVal controlKey As String, _
    ByRef outControl As Object, _
    Optional ByRef outReason As String = VBA.vbNullString _
) As Boolean
    outReason = VBA.vbNullString

    If m_ControlByKey Is Nothing Then
        outReason = "control-storage-empty"
        Exit Function
    End If
    If Not m_ControlByKey.Exists(controlKey) Then
        outReason = "control-not-found"
        Exit Function
    End If

    Set outControl = m_ControlByKey(controlKey)
    If outControl Is Nothing Then
        outReason = "control-entry-empty"
        Exit Function
    End If

    private_TryGetControl = True
End Function

Private Function private_ExtractControlNameFromControlKey(ByVal controlKey As String) As String
    Dim delimiterPos As Long

    controlKey = VBA.Trim$(controlKey)
    If VBA.Len(controlKey) = 0 Then Exit Function

    delimiterPos = VBA.InStrRev(controlKey, "|", -1, VBA.vbBinaryCompare)
    If delimiterPos <= 0 Then
        private_ExtractControlNameFromControlKey = VBA.LCase$(controlKey)
        Exit Function
    End If

    private_ExtractControlNameFromControlKey = VBA.LCase$(VBA.Trim$(VBA.Mid$(controlKey, delimiterPos + 1)))
End Function

Private Function private_TryNotifyGlobalClick( _
    ByVal clickedControlKey As String, _
    Optional ByRef outReason As String = VBA.vbNullString _
) As Boolean
    Dim key As Variant
    Dim iControl As Object
    Dim resultValue As Variant
    Dim errNo As Long

    outReason = VBA.vbNullString
    If m_ControlByKey Is Nothing Then
        private_TryNotifyGlobalClick = True
        Exit Function
    End If

    clickedControlKey = VBA.LCase$(VBA.Trim$(clickedControlKey))

    For Each key In m_ControlByKey.Keys
        Set iControl = m_ControlByKey(key)
        If iControl Is Nothing Then GoTo ContinueControl

        On Error Resume Next
        resultValue = VBA.CallByName(iControl, "m_RuntimeOnGlobalClick", VbMethod, clickedControlKey)
        errNo = Err.Number
        Err.Clear
        On Error GoTo 0

        If errNo <> 0 Then
            If errNo <> 438 Then
                outReason = "global-hook-exception:" & VBA.TypeName(iControl)
#If LOGGING_DEBUG_ENABLED Then
                ex_Core.fn_Diagnostic_LogError "PageBase: global click hook failed on '" & VBA.TypeName(iControl) & "'."
#End If
                Exit Function
            End If
        Else
            If VBA.VarType(resultValue) = vbBoolean Then
                If Not VBA.CBool(resultValue) Then
                    outReason = "global-hook-cancelled:" & VBA.TypeName(iControl)
                    Exit Function
                End If
            End If
        End If

ContinueControl:
    Next key

    private_TryNotifyGlobalClick = True
End Function

Private Function private_TryInvokeControlAction( _
    ByVal iControl As Object, _
    ByVal methodName As String, _
    ByVal hasArg As Boolean, _
    ByVal argValue As Variant, _
    ByRef outActionOk As Boolean, _
    Optional ByRef outErrorText As String = VBA.vbNullString _
) As Boolean
    Dim resultValue As Variant

    outErrorText = VBA.vbNullString
    If iControl Is Nothing Then Exit Function
    methodName = VBA.Trim$(methodName)
    If VBA.Len(methodName) = 0 Then
        outErrorText = "method-name-empty"
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: method name is empty."
#End If
        Exit Function
    End If

    On Error GoTo EH_INVOKE
    ' Поддерживаем сигнатуры действий как с аргументом, так и без аргумента.
    If hasArg Then
        resultValue = VBA.CallByName(iControl, methodName, VbMethod, argValue)
    Else
        resultValue = VBA.CallByName(iControl, methodName, VbMethod)
    End If
    On Error GoTo 0

    If VBA.VarType(resultValue) = vbBoolean Then
        outActionOk = VBA.CBool(resultValue)
    Else
        outActionOk = True
    End If

    private_TryInvokeControlAction = True
    Exit Function

EH_INVOKE:
    outErrorText = Err.Description
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "PageBase: failed to invoke method '" & methodName & "' on '" & VBA.TypeName(iControl) & "': " & Err.Description
#End If
End Function

Private Function private_TryResetSerializableControlActions() As Boolean
    Dim key As Variant
    Dim keyText As String
    Dim iControl As Object
    Dim iSerializable As obj_ISerializable
    Dim keysToRemove As Collection
    Dim removeKey As Variant

    private_EnsureStorage
    If m_ControlByKey Is Nothing Then
        private_TryResetSerializableControlActions = True
        Exit Function
    End If

    Set keysToRemove = New Collection

    For Each key In m_ControlByKey.Keys
        Set iControl = m_ControlByKey(key)
        If iControl Is Nothing Then GoTo ContinueControl
        If Not private_TryCastSerializableControl(iControl, iSerializable) Then GoTo ContinueControl

        keyText = VBA.LCase$(VBA.Trim$(VBA.CStr(key)))
        If VBA.Len(keyText) = 0 Then GoTo ContinueControl
        keysToRemove.Add keyText
ContinueControl:
    Next key

    For Each removeKey In keysToRemove
        If Not Me.UnregisterControl(VBA.CStr(removeKey)) Then Exit Function
    Next removeKey

    private_TryResetSerializableControlActions = True
End Function

Private Function private_TryCastSerializableControl(ByVal iControl As Object, ByRef outSerializableControl As obj_ISerializable) As Boolean
    If iControl Is Nothing Then Exit Function

    Set outSerializableControl = Nothing
    On Error Resume Next
    Set outSerializableControl = iControl
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    private_TryCastSerializableControl = Not outSerializableControl Is Nothing
End Function

Private Function private_GetDictionaryCount(ByVal dictObj As Object) As Long
    On Error Resume Next
    If Not dictObj Is Nothing Then private_GetDictionaryCount = VBA.CLng(dictObj.Count)
    Err.Clear
    On Error GoTo 0
End Function

Private Function private_CaptureSelectionAreas(ByVal ws As Worksheet) As Collection
    Dim selectedRange As Range
    Dim selectedArea As Range
    Dim areaInfo As Object
    Dim result As Collection
    Dim activeCellRange As Range

    Set result = New Collection
    Set private_CaptureSelectionAreas = result
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    If VBA.TypeName(Application.Selection) = "Range" Then
        Set selectedRange = Application.Selection
    End If
    Set activeCellRange = Application.ActiveCell
    On Error GoTo 0
    If selectedRange Is Nothing Then Exit Function
    If Not (selectedRange.Worksheet Is ws) Then Exit Function

    For Each selectedArea In selectedRange.Areas
        Set areaInfo = VBA.CreateObject("Scripting.Dictionary")
        areaInfo.CompareMode = 1
        areaInfo("RowStart") = VBA.CLng(selectedArea.Row)
        areaInfo("ColStart") = VBA.CLng(selectedArea.Column)
        areaInfo("RowCount") = VBA.CLng(selectedArea.Rows.Count)
        areaInfo("ColCount") = VBA.CLng(selectedArea.Columns.Count)
        If result.Count = 0 And Not activeCellRange Is Nothing Then
            If activeCellRange.Worksheet Is ws Then
                areaInfo("ActiveRow") = VBA.CLng(activeCellRange.Row)
                areaInfo("ActiveCol") = VBA.CLng(activeCellRange.Column)
            End If
        End If
        result.Add areaInfo
    Next selectedArea
End Function

Private Sub private_TranslateSelectionAreasByPatches( _
    ByVal selectionAreas As Collection, _
    ByVal patches As Collection _
)
    Dim patchIndex As Long
    Dim patch As Object
    Dim moveDown As Boolean

    If selectionAreas Is Nothing Or patches Is Nothing Then Exit Sub
    If selectionAreas.Count = 0 Or patches.Count = 0 Then Exit Sub

    Set patch = patches.Item(1)
    moveDown = (VBA.CLng(patch("RowDelta")) > 0)
    If moveDown Then
        For patchIndex = patches.Count To 1 Step -1
            Set patch = patches.Item(patchIndex)
            private_TranslateSelectionAreasForPatch selectionAreas, patch
        Next patchIndex
    Else
        For patchIndex = 1 To patches.Count
            Set patch = patches.Item(patchIndex)
            private_TranslateSelectionAreasForPatch selectionAreas, patch
        Next patchIndex
    End If
End Sub

Private Sub private_TranslateSelectionAreasForPatch( _
    ByVal selectionAreas As Collection, _
    ByVal patch As Object _
)
    Dim areaInfo As Object
    Dim rowStart As Long
    Dim colStart As Long
    Dim rowEnd As Long
    Dim colEnd As Long
    Dim rowDelta As Long
    Dim areaRowEnd As Long
    Dim areaColEnd As Long

    If selectionAreas Is Nothing Or patch Is Nothing Then Exit Sub
    rowStart = VBA.CLng(patch("RowStart"))
    colStart = VBA.CLng(patch("ColStart"))
    rowEnd = VBA.CLng(patch("RowEnd"))
    colEnd = VBA.CLng(patch("ColEnd"))
    rowDelta = VBA.CLng(patch("RowDelta"))
    If rowDelta = 0 Then Exit Sub

    For Each areaInfo In selectionAreas
        areaRowEnd = VBA.CLng(areaInfo("RowStart")) + VBA.CLng(areaInfo("RowCount")) - 1
        areaColEnd = VBA.CLng(areaInfo("ColStart")) + VBA.CLng(areaInfo("ColCount")) - 1
        If VBA.CLng(areaInfo("RowStart")) < rowStart Or areaRowEnd > rowEnd Then GoTo ContinueArea
        If VBA.CLng(areaInfo("ColStart")) < colStart Or areaColEnd > colEnd Then GoTo ContinueArea

        areaInfo("RowStart") = VBA.CLng(areaInfo("RowStart")) + rowDelta
        If areaInfo.Exists("ActiveRow") Then
            If VBA.CLng(areaInfo("ActiveRow")) >= rowStart And _
                VBA.CLng(areaInfo("ActiveRow")) <= rowEnd And _
                VBA.CLng(areaInfo("ActiveCol")) >= colStart And _
                VBA.CLng(areaInfo("ActiveCol")) <= colEnd Then
                areaInfo("ActiveRow") = VBA.CLng(areaInfo("ActiveRow")) + rowDelta
            End If
        End If
ContinueArea:
    Next areaInfo
End Sub

Private Sub private_RestoreSelectionAreas( _
    ByVal ws As Worksheet, _
    ByVal selectionAreas As Collection _
)
    Dim areaInfo As Object
    Dim areaRange As Range
    Dim restoredRange As Range
    Dim activeRow As Long
    Dim activeCol As Long
    Dim hasActiveCell As Boolean

    If ws Is Nothing Or selectionAreas Is Nothing Then Exit Sub
    If selectionAreas.Count = 0 Then Exit Sub
    On Error GoTo RestoreFailed

    For Each areaInfo In selectionAreas
        Set areaRange = ws.Range( _
            ws.Cells(VBA.CLng(areaInfo("RowStart")), VBA.CLng(areaInfo("ColStart"))), _
            ws.Cells( _
                VBA.CLng(areaInfo("RowStart")) + VBA.CLng(areaInfo("RowCount")) - 1, _
                VBA.CLng(areaInfo("ColStart")) + VBA.CLng(areaInfo("ColCount")) - 1))
        If restoredRange Is Nothing Then
            Set restoredRange = areaRange
        Else
            Set restoredRange = Application.Union(restoredRange, areaRange)
        End If
        If areaInfo.Exists("ActiveRow") Then
            activeRow = VBA.CLng(areaInfo("ActiveRow"))
            activeCol = VBA.CLng(areaInfo("ActiveCol"))
            hasActiveCell = True
        End If
    Next areaInfo

    If restoredRange Is Nothing Then Exit Sub
    restoredRange.Select
    If hasActiveCell Then ws.Cells(activeRow, activeCol).Activate
RestoreFailed:
    Err.Clear
End Sub

Private Function private_TryApplyLayoutReflowPatches( _
    ByVal ws As Worksheet, _
    ByVal patches As Collection _
) As Boolean
    Dim patchIndex As Long
    Dim patch As Object
    Dim moveDown As Boolean

    If ws Is Nothing Or patches Is Nothing Then Exit Function
    If patches.Count = 0 Then
        private_TryApplyLayoutReflowPatches = True
        Exit Function
    End If

    Set patch = patches.Item(1)
    moveDown = (VBA.CLng(patch("RowDelta")) > 0)
    If moveDown Then
        ' При росте сначала двигаем нижние siblings: верхний translate не должен
        ' перезаписать еще не перенесенный нижний subtree.
        For patchIndex = patches.Count To 1 Step -1
            Set patch = patches.Item(patchIndex)
            If Not private_TryApplySingleLayoutReflowPatch(ws, patch) Then Exit Function
        Next patchIndex
    Else
        For patchIndex = 1 To patches.Count
            Set patch = patches.Item(patchIndex)
            If Not private_TryApplySingleLayoutReflowPatch(ws, patch) Then Exit Function
        Next patchIndex
    End If

    private_TryApplyLayoutReflowPatches = True
End Function

Private Function private_CommitRuntimeAncestorUpdates(ByVal ancestorUpdates As Collection) As Boolean
    Dim key As Variant
    Dim entry As Object
    Dim updateEntry As Object

    If ancestorUpdates Is Nothing Or m_LayoutContainerByName Is Nothing Then
        private_CommitRuntimeAncestorUpdates = True
        Exit Function
    End If
    For Each key In m_LayoutContainerByName.Keys
        Set entry = m_LayoutContainerByName(key)
        For Each updateEntry In ancestorUpdates
            If VBA.CLng(entry("RowStart")) = VBA.CLng(updateEntry("RowStart")) And _
               VBA.CLng(entry("ColStart")) = VBA.CLng(updateEntry("ColStart")) And _
               VBA.CLng(entry("RowEnd")) = VBA.CLng(updateEntry("OldRowEnd")) And _
               VBA.CLng(entry("ColEnd")) = VBA.CLng(updateEntry("ColEnd")) Then
                entry("RowEnd") = VBA.CLng(updateEntry("RowEnd"))
                Exit For
            End If
        Next updateEntry
    Next key
    private_CommitRuntimeAncestorUpdates = True
End Function

Private Function private_TryApplySingleLayoutReflowPatch( _
    ByVal ws As Worksheet, _
    ByVal patch As Object _
) As Boolean
    Dim rowStart As Long, colStart As Long, rowEnd As Long, colEnd As Long, rowDelta As Long
    Dim vacatedRange As Range

    If patch Is Nothing Then Exit Function
    rowStart = VBA.CLng(patch("RowStart")): colStart = VBA.CLng(patch("ColStart"))
    rowEnd = VBA.CLng(patch("RowEnd")): colEnd = VBA.CLng(patch("ColEnd"))
    rowDelta = VBA.CLng(patch("RowDelta"))
    If rowDelta = 0 Or rowEnd < rowStart Then
        private_TryApplySingleLayoutReflowPatch = True
        Exit Function
    End If

    If Not private_TryTranslateWorksheetSubtreeRows( _
        ws, rowStart, colStart, rowEnd, colEnd, rowDelta) Then Exit Function
    If rowDelta > 0 Then
        Set vacatedRange = ws.Range( _
            ws.Cells(rowStart, colStart), _
            ws.Cells(rowStart + rowDelta - 1, colEnd))
    Else
        Set vacatedRange = ws.Range( _
            ws.Cells(rowEnd + rowDelta + 1, colStart), _
            ws.Cells(rowEnd, colEnd))
    End If
    ' Блочный Copy намеренно оставляет исходные ячейки на месте. Очищаем
    ' только полосу, которая действительно освободилась после translate;
    ' пересекающуюся часть source/destination трогать нельзя.
    vacatedRange.Clear
    If Not ex_StylePipelineEngine.fn_ApplySheetBaseStylesToRange( _
        ws, m_UiDom, vacatedRange) Then Exit Function
    If Not private_TranslateRuntimeRegion(rowStart, colStart, rowEnd, colEnd, rowDelta) Then Exit Function
    If Not ex_ControlRefreshRuntime.fn_TranslateRegisteredControlsInRegion( _
        ws.Name, rowStart, colStart, rowEnd, colEnd, rowDelta) Then Exit Function
    If Not ex_ControlRefreshRuntime.fn_TranslateLayoutEntriesInRegion( _
        ws.Name, rowStart, colStart, rowEnd, colEnd, rowDelta) Then Exit Function
    If Not ex_ControlPartsRuntime.fn_TranslateControlPartsInRegion( _
        ws, rowStart, colStart, rowEnd, colEnd, rowDelta) Then Exit Function
    If Not ex_StylePipelineEngine.fn_TranslateLayoutBoundsInRegion( _
        ws.Name, rowStart, colStart, rowEnd, colEnd, rowDelta) Then Exit Function

    private_TryApplySingleLayoutReflowPatch = True
End Function

Private Function private_TryTranslateWorksheetSubtreeRows( _
    ByVal ws As Worksheet, _
    ByVal firstRow As Long, _
    ByVal firstCol As Long, _
    ByVal lastRow As Long, _
    ByVal lastCol As Long, _
    ByVal rowDelta As Long _
) As Boolean
    Dim sourceRange As Range
    Dim destinationCell As Range
    Dim rowHeights() As Double
    Dim rowIndex As Long
    Dim shapeInfo As Collection
    Dim knownShapeNames As Object
    Dim info As Object
    Dim shp As Shape
    Dim topCell As Range
    Dim newTopCell As Range
    Dim shapeIndex As Long
    Dim moveRow As Long
    Dim copyErrorNumber As Long
    Dim perfTotalStartedAt As Double, perfStageStartedAt As Double
    Dim captureMs As Double, copyMs As Double, duplicateCleanupMs As Double
    Dim placementRestoreMs As Double, restoreRowsAndShapesMs As Double
    Dim shapesBefore As Long

    perfTotalStartedAt = VBA.Timer

    If ws Is Nothing Then Exit Function
    If firstRow <= 0 Or firstCol <= 0 Or lastRow < firstRow Or lastCol < firstCol Then Exit Function
    If firstRow + rowDelta <= 0 Or lastRow + rowDelta > ws.Rows.Count Then Exit Function
    If rowDelta = 0 Then
        private_TryTranslateWorksheetSubtreeRows = True
        Exit Function
    End If
    ' Generated UI содержит отрисованные значения и стили, а не вычислительные
    ' формулы, поэтому перекрывающийся subtree можно перенести одним Copy.
    ' Это принципиально быстрее прежнего построчного Cut: каждый Cut заставлял
    ' Excel отдельно перестраивать зависимости и занимал до нескольких секунд.
    '
    ' При Placement = xlMoveAndSize Copy может создать дубликаты Shapes. Ниже
    ' удаляем все новые имена, а исходные Shapes выставляем по сохранённым
    ' координатам. Высоты строк также восстанавливаются отдельно, поскольку
    ' Excel не переносит RowHeight вместе с прямоугольником ячеек.
    '
    ' Мы не вставляем/удаляем строки листа: такая операция сдвинула бы также
    ' независимые контролы слева/справа. Переносится только прямоугольник,
    ' рассчитанный layout-планом, поэтому horizontal/grid siblings сохраняют
    ' декларативные позиции.
    ReDim rowHeights(firstRow To lastRow)
    For rowIndex = firstRow To lastRow
        rowHeights(rowIndex) = ws.Rows(rowIndex).RowHeight
    Next rowIndex

    Set knownShapeNames = VBA.CreateObject("Scripting.Dictionary")
    knownShapeNames.CompareMode = 1
    shapesBefore = ws.Shapes.Count
    perfStageStartedAt = VBA.Timer
    For Each shp In ws.Shapes
        knownShapeNames(shp.Name) = True
    Next shp

    Set shapeInfo = New Collection
    For Each shp In ws.Shapes
        Set topCell = Nothing
        On Error Resume Next
        Set topCell = shp.TopLeftCell
        On Error GoTo 0
        If topCell Is Nothing Then GoTo ContinueShape
        If topCell.Row < firstRow Or topCell.Row > lastRow Then GoTo ContinueShape
        If topCell.Column < firstCol Or topCell.Column > lastCol Then GoTo ContinueShape

        Set info = VBA.CreateObject("Scripting.Dictionary")
        info.CompareMode = 1
        info("Name") = shp.Name
        info("Row") = VBA.CLng(topCell.Row)
        info("Col") = VBA.CLng(topCell.Column)
        info("TopOffset") = VBA.CDbl(shp.Top - topCell.Top)
        info("LeftOffset") = VBA.CDbl(shp.Left - topCell.Left)
        info("Placement") = VBA.CLng(shp.Placement)
        info("PlacementChanged") = False
        ' Range.Copy дублирует Shapes с xlMove/xlMoveAndSize, хотя ниже runtime
        ' всё равно переносит оригиналы вручную. Временно отвязываем Shape от
        ' ячеек, чтобы repeated partial reflow не фрагментировал drawing-layer.
        On Error Resume Next
        shp.Placement = xlFreeFloating
        info("PlacementChanged") = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
        shapeInfo.Add info
ContinueShape:
    Next shp
    captureMs = private_PerfElapsedMs(perfStageStartedAt)
    Set sourceRange = ws.Range( _
        ws.Cells(firstRow, firstCol), _
        ws.Cells(lastRow, lastCol))
    Set destinationCell = ws.Cells(firstRow + rowDelta, firstCol)
    perfStageStartedAt = VBA.Timer
    On Error Resume Next
    sourceRange.Copy Destination:=destinationCell
    copyErrorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    If copyErrorNumber <> 0 Then
        ' Некоторые старые сборки Excel могут запретить блочный Copy между
        ' перекрывающимися диапазонами. Медленный путь сохраняем только как
        ' совместимый fallback, чтобы partial reflow не ломал страницу.
        If rowDelta > 0 Then
            For moveRow = lastRow To firstRow Step -1
                Set sourceRange = ws.Range( _
                    ws.Cells(moveRow, firstCol), _
                    ws.Cells(moveRow, lastCol))
                Set destinationCell = ws.Cells(moveRow + rowDelta, firstCol)
                sourceRange.Copy Destination:=destinationCell
            Next moveRow
        Else
            For moveRow = firstRow To lastRow
                Set sourceRange = ws.Range( _
                    ws.Cells(moveRow, firstCol), _
                    ws.Cells(moveRow, lastCol))
                Set destinationCell = ws.Cells(moveRow + rowDelta, firstCol)
                sourceRange.Copy Destination:=destinationCell
            Next moveRow
        End If
    End If
    Application.CutCopyMode = False
    copyMs = private_PerfElapsedMs(perfStageStartedAt)

    ' Copy завершён — сразу возвращаем исходную привязку оригинальных Shapes.
    ' Это также ограничивает время, в течение которого UI находится во
    ' временном xlFreeFloating-состоянии при последующей ошибке cleanup.
    perfStageStartedAt = VBA.Timer
    For Each info In shapeInfo
        If VBA.CBool(info("PlacementChanged")) Then
            Set shp = Nothing
            On Error Resume Next
            Set shp = ws.Shapes(VBA.CStr(info("Name")))
            If Not shp Is Nothing Then shp.Placement = VBA.CLng(info("Placement"))
            On Error GoTo 0
        End If
    Next info
    placementRestoreMs = private_PerfElapsedMs(perfStageStartedAt)

    ' Оставляем только Shapes, существовавшие до блочного Copy.
    perfStageStartedAt = VBA.Timer
    For shapeIndex = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(shapeIndex)
        If Not knownShapeNames.Exists(shp.Name) Then shp.Delete
    Next shapeIndex
    duplicateCleanupMs = private_PerfElapsedMs(perfStageStartedAt)

    perfStageStartedAt = VBA.Timer
    For rowIndex = firstRow To lastRow
        ws.Rows(rowIndex + rowDelta).RowHeight = rowHeights(rowIndex)
    Next rowIndex

    For Each info In shapeInfo
        Set shp = Nothing
        On Error Resume Next
        Set shp = ws.Shapes(VBA.CStr(info("Name")))
        On Error GoTo 0
        If shp Is Nothing Then GoTo ContinueMovedShape
        Set newTopCell = ws.Cells(VBA.CLng(info("Row")) + rowDelta, VBA.CLng(info("Col")))
        shp.Top = newTopCell.Top + VBA.CDbl(info("TopOffset"))
        shp.Left = newTopCell.Left + VBA.CDbl(info("LeftOffset"))
ContinueMovedShape:
    Next info
    restoreRowsAndShapesMs = private_PerfElapsedMs(perfStageStartedAt)

    private_TryTranslateWorksheetSubtreeRows = True
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "perf:worksheet-subtree-translate totalMs='" & _
        VBA.Format$(private_PerfElapsedMs(perfTotalStartedAt), "0") & _
        "' captureMs='" & VBA.Format$(captureMs, "0") & _
        "' copyMs='" & VBA.Format$(copyMs, "0") & _
        "' placementRestoreMs='" & VBA.Format$(placementRestoreMs, "0") & _
        "' duplicateCleanupMs='" & VBA.Format$(duplicateCleanupMs, "0") & _
        "' restoreMs='" & VBA.Format$(restoreRowsAndShapesMs, "0") & _
        "' rows='" & VBA.CStr(lastRow - firstRow + 1) & _
        "' cols='" & VBA.CStr(lastCol - firstCol + 1) & _
        "' rowDelta='" & VBA.CStr(rowDelta) & _
        "' movedShapes='" & VBA.CStr(shapeInfo.Count) & _
        "' shapesBefore='" & VBA.CStr(shapesBefore) & _
        "' shapesAfter='" & VBA.CStr(ws.Shapes.Count) & _
        "' copyFallback='" & VBA.LCase$(VBA.CStr(copyErrorNumber <> 0)) & "'"
#End If
End Function

Private Sub private_DeleteDuplicateSingleButtonRuntimeShapes(ByVal ws As Worksheet)
    Dim canonicalNamesByControl As Object
    Dim shp As Shape
    Dim controlName As String
    Dim canonicalName As String
    Dim shapeIndex As Long

    If ws Is Nothing Then Exit Sub
    Set canonicalNamesByControl = VBA.CreateObject("Scripting.Dictionary")
    canonicalNamesByControl.CompareMode = 1

    ' Сначала находим именно одиночные ButtonControlVM по стабильному имени.
    ' ButtonGroup/Select/Hotkeys используют другие схемы имен и сюда не попадут.
    For Each shp In ws.Shapes
        controlName = VBA.Trim$(ex_ShapeMetaRuntime.fn_GetShapeMetaValue( _
            shp, "pn.control", VBA.vbNullString))
        If VBA.Len(controlName) = 0 Then GoTo ContinueCanonicalShape
        canonicalName = "btn_" & controlName
        If VBA.StrComp(shp.Name, canonicalName, VBA.vbTextCompare) = 0 Then
            canonicalNamesByControl(controlName) = canonicalName
        End If
ContinueCanonicalShape:
    Next shp

    If canonicalNamesByControl.Count = 0 Then Exit Sub
    For shapeIndex = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(shapeIndex)
        controlName = VBA.Trim$(ex_ShapeMetaRuntime.fn_GetShapeMetaValue( _
            shp, "pn.control", VBA.vbNullString))
        If VBA.Len(controlName) = 0 Then GoTo ContinueDuplicateShape
        If Not canonicalNamesByControl.Exists(controlName) Then GoTo ContinueDuplicateShape
        canonicalName = VBA.CStr(canonicalNamesByControl(controlName))
        If VBA.StrComp(shp.Name, canonicalName, VBA.vbTextCompare) <> 0 Then shp.Delete
ContinueDuplicateShape:
    Next shapeIndex
End Sub

Private Function private_TryReconcileSingleButtonRuntimeShapes(ByVal ws As Worksheet) As Boolean
    Dim shp As Shape
    Dim controlName As String
    Dim canonicalName As String
    Dim rowStart As Long
    Dim colStart As Long
    Dim rowEnd As Long
    Dim colEnd As Long
    Dim targetRange As Range

    If ws Is Nothing Then Exit Function
    On Error GoTo EH_RECONCILE_BUTTONS

    For Each shp In ws.Shapes
        controlName = VBA.Trim$(ex_ShapeMetaRuntime.fn_GetShapeMetaValue( _
            shp, "pn.control", VBA.vbNullString))
        If VBA.Len(controlName) = 0 Then GoTo ContinueShape

        ' Только ButtonControlVM имеет ровно один Shape со стабильным именем
        ' btn_<control>. ButtonGroup, Select и Hotkeys имеют внутренние части,
        ' геометрию которых нельзя приравнивать к общим bounds контрола.
        ' Координаты намеренно берём из runtime registry, а не из текущего Shape.
        ' Поэтому локальный refresh восстанавливает декларативную позицию даже
        ' после ручного перетаскивания кнопки пользователем.
        canonicalName = "btn_" & controlName
        If VBA.StrComp(shp.Name, canonicalName, VBA.vbTextCompare) <> 0 Then GoTo ContinueShape
        If Not ex_ControlRefreshRuntime.fn_TryGetControlRenderBounds( _
            controlName, ws.Name, rowStart, colStart, rowEnd, colEnd) Then GoTo ContinueShape

        Set targetRange = ws.Range( _
            ws.Cells(rowStart, colStart), _
            ws.Cells(rowEnd, colEnd))
        shp.Left = targetRange.Left
        shp.Top = targetRange.Top
        shp.Width = targetRange.Width
        shp.Height = targetRange.Height
        shp.Placement = xlMoveAndSize
ContinueShape:
    Next shp

    private_TryReconcileSingleButtonRuntimeShapes = True
    Exit Function

EH_RECONCILE_BUTTONS:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError _
        "PageBase: failed to reconcile retained Button shapes: " & Err.Description
#End If
End Function

Private Function private_TranslateRuntimeRows( _
    ByVal firstRow As Long, _
    ByVal rowDelta As Long _
) As Boolean
    Dim key As Variant
    Dim entryObj As Variant
    Dim entries As Collection
    Dim translatedRoutes As Object
    Dim routeEntry As Object
    Dim cellRange As Range
    Dim newCellKey As String
    Dim inlineEntry As Object

    If firstRow <= 0 Then Exit Function
    If rowDelta = 0 Then
        private_TranslateRuntimeRows = True
        Exit Function
    End If

    ' Named containers: downstream nodes сдвигаются целиком, а ancestor,
    ' пересекающий boundary, только меняет нижнюю границу.
    If Not m_LayoutContainerByName Is Nothing Then
        For Each key In m_LayoutContainerByName.Keys
            Set entryObj = m_LayoutContainerByName(key)
            private_TranslateCoordinateEntry entryObj, firstRow, rowDelta
        Next key
    End If

    ' Layout tags должны продолжать указывать на те же логические поля после
    ' физического переноса нижней части страницы.
    If Not m_LayoutTagEntriesByTag Is Nothing Then
        For Each key In m_LayoutTagEntriesByTag.Keys
            Set entries = m_LayoutTagEntriesByTag(key)
            For Each entryObj In entries
                private_TranslateCoordinateEntry entryObj, firstRow, rowDelta
            Next entryObj
        Next key
    End If

    ' SheetChange маршрутизируется по адресу ячейки, поэтому переносим и key,
    ' и строковый callback argument, зарегистрированный InputControlVM.
    If Not m_RouteByCell Is Nothing Then
        Set translatedRoutes = VBA.CreateObject("Scripting.Dictionary")
        translatedRoutes.CompareMode = 1
        For Each key In m_RouteByCell.Keys
            Set routeEntry = m_RouteByCell(key)
            Set cellRange = Nothing
            On Error Resume Next
            Set cellRange = m_Worksheet.Range(VBA.CStr(key))
            On Error GoTo 0
            newCellKey = VBA.CStr(key)
            If Not cellRange Is Nothing Then
                If cellRange.Row >= firstRow Then
                    newCellKey = cellRange.Offset(rowDelta, 0).Address(False, False)
                    If routeEntry.Exists("HasArg") Then
                        If VBA.CBool(routeEntry("HasArg")) Then routeEntry("ArgValue") = newCellKey
                    End If
                End If
            End If
            Set translatedRoutes(VBA.UCase$(newCellKey)) = routeEntry
        Next key
        Set m_RouteByCell = translatedRoutes
    End If

    If Not m_InlineRunEntries Is Nothing Then
        For Each inlineEntry In m_InlineRunEntries
            If VBA.LCase$(VBA.CStr(inlineEntry("TargetType"))) = INLINE_TARGET_RANGE Then
                Set cellRange = Nothing
                On Error Resume Next
                Set cellRange = m_Worksheet.Range(VBA.CStr(inlineEntry("CellAddress")))
                On Error GoTo 0
                If Not cellRange Is Nothing Then
                    If cellRange.Row >= firstRow Then
                        newCellKey = cellRange.Offset(rowDelta, 0).Address(False, False)
                        inlineEntry("CellAddress") = newCellKey
                        inlineEntry("TargetKey") = VBA.LCase$(newCellKey)
                    End If
                End If
            End If
        Next inlineEntry
    End If

    private_TranslateRuntimeRows = True
End Function

Private Function private_TranslateRuntimeRegion( _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long, _
    ByVal rowDelta As Long _
) As Boolean
    Dim key As Variant
    Dim entryObj As Variant
    Dim entries As Collection
    Dim translatedRoutes As Object
    Dim routeEntry As Object
    Dim cellRange As Range
    Dim newCellKey As String
    Dim inlineEntry As Object

    If Not m_LayoutContainerByName Is Nothing Then
        For Each key In m_LayoutContainerByName.Keys
            Set entryObj = m_LayoutContainerByName(key)
            private_TranslateCoordinateEntryInRegion entryObj, rowStart, colStart, rowEnd, colEnd, rowDelta
        Next key
    End If
    If Not m_LayoutTagEntriesByTag Is Nothing Then
        For Each key In m_LayoutTagEntriesByTag.Keys
            Set entries = m_LayoutTagEntriesByTag(key)
            For Each entryObj In entries
                private_TranslateCoordinateEntryInRegion entryObj, rowStart, colStart, rowEnd, colEnd, rowDelta
            Next entryObj
        Next key
    End If

    If Not m_RouteByCell Is Nothing Then
        Set translatedRoutes = VBA.CreateObject("Scripting.Dictionary")
        translatedRoutes.CompareMode = 1
        For Each key In m_RouteByCell.Keys
            Set routeEntry = m_RouteByCell(key)
            Set cellRange = Nothing
            On Error Resume Next
            Set cellRange = m_Worksheet.Range(VBA.CStr(key))
            On Error GoTo 0
            newCellKey = VBA.CStr(key)
            If Not cellRange Is Nothing Then
                If cellRange.Row >= rowStart And cellRange.Row <= rowEnd And _
                   cellRange.Column >= colStart And cellRange.Column <= colEnd Then
                    newCellKey = cellRange.Offset(rowDelta, 0).Address(False, False)
                    If VBA.CBool(routeEntry("HasArg")) Then routeEntry("ArgValue") = newCellKey
                End If
            End If
            Set translatedRoutes(VBA.UCase$(newCellKey)) = routeEntry
        Next key
        Set m_RouteByCell = translatedRoutes
    End If

    If Not m_InlineRunEntries Is Nothing Then
        For Each inlineEntry In m_InlineRunEntries
            If VBA.LCase$(VBA.CStr(inlineEntry("TargetType"))) = INLINE_TARGET_RANGE Then
                Set cellRange = Nothing
                On Error Resume Next
                Set cellRange = m_Worksheet.Range(VBA.CStr(inlineEntry("CellAddress")))
                On Error GoTo 0
                If Not cellRange Is Nothing Then
                    If cellRange.Row >= rowStart And cellRange.Row <= rowEnd And _
                       cellRange.Column >= colStart And cellRange.Column <= colEnd Then
                        newCellKey = cellRange.Offset(rowDelta, 0).Address(False, False)
                        inlineEntry("CellAddress") = newCellKey
                        inlineEntry("TargetKey") = VBA.LCase$(newCellKey)
                    End If
                End If
            End If
        Next inlineEntry
    End If
    private_TranslateRuntimeRegion = True
End Function

Private Sub private_TranslateCoordinateEntryInRegion( _
    ByVal entry As Object, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long, _
    ByVal rowDelta As Long _
)
    If entry Is Nothing Then Exit Sub
    If VBA.CLng(entry("RowStart")) < rowStart Or VBA.CLng(entry("RowEnd")) > rowEnd Then Exit Sub
    If VBA.CLng(entry("ColStart")) < colStart Or VBA.CLng(entry("ColEnd")) > colEnd Then Exit Sub
    entry("RowStart") = VBA.CLng(entry("RowStart")) + rowDelta
    entry("RowEnd") = VBA.CLng(entry("RowEnd")) + rowDelta
End Sub

Private Sub private_TranslateCoordinateEntry( _
    ByVal entry As Object, _
    ByVal firstRow As Long, _
    ByVal rowDelta As Long _
)
    Dim rowStart As Long
    Dim rowEnd As Long

    If entry Is Nothing Then Exit Sub
    rowStart = VBA.CLng(entry("RowStart"))
    rowEnd = VBA.CLng(entry("RowEnd"))
    If rowStart >= firstRow Then
        entry("RowStart") = rowStart + rowDelta
        entry("RowEnd") = rowEnd + rowDelta
    ElseIf rowEnd >= firstRow Then
        entry("RowEnd") = rowEnd + rowDelta
    End If
End Sub

Private Function private_BuildLogContext() As String
    Dim sheetName As String
    Dim codeNameValue As String

    sheetName = VBA.vbNullString
    codeNameValue = VBA.vbNullString

    If Not m_Worksheet Is Nothing Then
        On Error Resume Next
        sheetName = VBA.Trim$(VBA.CStr(m_Worksheet.Name))
        codeNameValue = VBA.Trim$(VBA.CStr(m_Worksheet.CodeName))
        Err.Clear
        On Error GoTo 0
    End If

    private_BuildLogContext = "pageId='" & private_EscapeForLog(VBA.Trim$(m_PageId)) & "' sheet='" & private_EscapeForLog(sheetName) & "' codeName='" & private_EscapeForLog(codeNameValue) & "'"
End Function

Private Function private_EscapeForLog(ByVal valueText As String) As String
    private_EscapeForLog = VBA.Replace$(VBA.CStr(valueText), "'", "''")
End Function

Private Function private_PerfElapsedMs(ByVal startedAt As Double) As Double
    Dim finishedAt As Double
    finishedAt = VBA.Timer
    If finishedAt < startedAt Then finishedAt = finishedAt + 86400#
    private_PerfElapsedMs = (finishedAt - startedAt) * 1000#
End Function

#If LOGGING_DEBUG_ENABLED Then

#End If

Private Sub private_LogRuntimeInfo(ByVal messageText As String)
    On Error Resume Next
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "page-base:" & VBA.Trim$(messageText) & " " & private_BuildLogContext()
#End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub private_LogRuntimeError(ByVal messageText As String)
    On Error Resume Next
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "page-base:" & VBA.Trim$(messageText) & " " & private_BuildLogContext()
#End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Function private_EnsureNotDisposed(ByVal methodName As String) As Boolean
    If m_IsDisposed Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageBase: method '" & methodName & "' cannot be used after Dispose."
#End If
        Exit Function
    End If

    private_EnsureNotDisposed = True
End Function
