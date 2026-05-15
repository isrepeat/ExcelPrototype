VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PageBase"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private m_Worksheet As Worksheet
Private m_Page As obj_IPage
Private m_UiPath As String
Private m_LastRenderedUiPath As String
Private m_PageId As String
Private m_UiDom As Object
Private m_IsDisposed As Boolean
Private m_IsRendering As Boolean
Private m_ControlByKey As Object
Private m_RouteByShape As Object
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
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
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
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
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
    If Not m_PageRuntimeSources.Initialize(m_Page) Then Exit Function

    ' Ранний дефолт: при инициализации страницы фиксируем текстовый формат текущего used-range.
    ' Это защитный baseline; основной повтор формата выполняется в runtime-clear перед каждым Render.
    On Error Resume Next
    Set clearRange = m_Worksheet.UsedRange
    If Not clearRange Is Nothing Then clearRange.NumberFormat = "@"
    Set clearRange = Nothing
    On Error GoTo 0

    Set m_InlineRunEntries = Nothing
    ' Реестр профилей стартует пустым и заполняется лениво по мере рендера.
    Set m_InlineProfileByPart = Nothing
    Call Me.ResetControlActions
    Initialize = Me.IsReady()
End Function

Public Sub Dispose(Optional ByVal deleteWorksheet As Boolean = True)
    Dim ws As Worksheet

#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub

    Set ws = m_Worksheet
    Set m_Worksheet = Nothing
    Set m_Page = Nothing
    m_UiPath = VBA.vbNullString
    m_LastRenderedUiPath = VBA.vbNullString
    m_PageId = VBA.vbNullString
    Set m_UiDom = Nothing
    Set m_PageRuntimeSources = Nothing
    Set m_InlineRunEntries = Nothing
    ' Сбрасываем профильный кэш вместе со страницей (единый lifecycle PageBase).
    Set m_InlineProfileByPart = Nothing
    m_IsRendering = False
    Call Me.ResetControlActions
    m_IsDisposed = True

    If Not deleteWorksheet Then Exit Sub
    If ws Is Nothing Then Exit Sub

    On Error GoTo EH_DELETE
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    Exit Sub

EH_DELETE:
    Application.DisplayAlerts = True
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "PageBase: failed to delete worksheet during dispose: " & Err.Description
#End If
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
    Dim prevDisplayAlerts As Boolean
    Dim prevCalculation As XlCalculation
    Dim prevStatusBar As Variant
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim layoutRenderContext As obj_LayoutRenderContext

    If Not private_EnsureNotDisposed("Render") Then Exit Function
    If Not Me.IsReady() Then Exit Function

    If m_IsRendering Then Exit Function

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
    Set m_UiDom = ex_XmlCore.fn_LoadDomByRelativePath( _
        wb, _
        resolvedUiPath, _
        "PrototypeNew: page UI file was not found: ", _
        "PrototypeNew: failed to parse page UI file: ", _
        UI_NS)
    If m_UiDom Is Nothing Then Exit Function

    Set pageNode = m_UiDom.selectSingleNode("/p:page")
    If pageNode Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: page UI root node <page> is missing."
#End If
        Exit Function
    End If

    m_UiPath = resolvedUiPath
    retainGeneratedShapes = private_ShouldRetainGeneratedShapes(previousUiPath, resolvedUiPath)
    m_IsRendering = True
    Set app = Application
    private_EnterFastRenderMode app, prevScreenUpdating, prevEnableEvents, prevDisplayAlerts, prevCalculation, prevStatusBar
    On Error GoTo EH_RENDER

    ' Сбрасываем runtime-реестры, чтобы не тянуть старые контролы/маршруты.
    ex_ControlPartsRuntime.fn_ResetControlParts
    Me.ResetInlineRuns
    ex_ControlRefreshRuntime.fn_ResetRegisteredControls
    ex_StylePipelineEngine.fn_ResetLayoutBounds
    ex_LayoutControlFallbackRndr.fn_ResetControlFallbacks

    If Not Me.ResetControlActions() Then GoTo Cleanup
    If Not private_TryClearPageRuntime(Not retainGeneratedShapes) Then GoTo Cleanup
    ' Один контекст на один проход: worksheet/workbook и seed-ы runtime ключей.
    Set layoutRenderContext = New obj_LayoutRenderContext
    If Not layoutRenderContext.Initialize(m_Page) Then GoTo Cleanup
    If Not ex_XmlLayoutEngine.fn_RenderNode(layoutRenderContext, pageNode) Then GoTo Cleanup
    If Not ex_StylePipelineEngine.fn_ApplyPageStyles(ws, m_UiDom) Then GoTo Cleanup
    ex_LayoutControlFallbackRndr.fn_ApplyPendingControlFallbacks ws
    If Not Me.ApplyInlineRuns() Then GoTo Cleanup

    ' В retained-режиме глобально shape не удаляем до рендера.
    ' После рендера чистим только orphan-shape (контролы, которые больше не присутствуют в текущем layout).
    If retainGeneratedShapes Then
        Call private_DeleteOrphanRuntimeShapesByControlRegistry(ws)
    End If

    private_LogRuntimeInfo "render-bindings controls=" & VBA.CStr(private_GetDictionaryCount(m_ControlByKey)) & " routes=" & VBA.CStr(private_GetDictionaryCount(m_RouteByShape))

    Render = True
    m_LastRenderedUiPath = resolvedUiPath

Cleanup:
    private_LeaveFastRenderMode app, prevScreenUpdating, prevEnableEvents, prevDisplayAlerts, prevCalculation, prevStatusBar
    m_IsRendering = False
    Exit Function

EH_RENDER:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description

    private_LeaveFastRenderMode app, prevScreenUpdating, prevEnableEvents, prevDisplayAlerts, prevCalculation, prevStatusBar
    m_IsRendering = False
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "PrototypeNew: render failed: [" & errSource & " #" & VBA.CStr(errNumber) & "] " & errDescription
#End If
End Function

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
    private_LogRuntimeInfo "register-control key='" & private_EscapeForLog(controlKey) & "' controls=" & VBA.CStr(private_GetDictionaryCount(m_ControlByKey))
    RegisterControl = True
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
    private_LogRuntimeInfo "register-route shape='" & private_EscapeForLog(shapeName) & "' control='" & private_EscapeForLog(controlKey) & "' method='" & private_EscapeForLog(methodName) & "' routes=" & VBA.CStr(private_GetDictionaryCount(m_RouteByShape))
    RegisterShapeRoute = True
End Function

' Callstack[1]: obj_PageMain.UnregisterControl -> obj_PageMain.obj_IPage_UnregisterControl -> obj_PageBase.UnregisterControl
' Callstack[2]: page.UnregisterControl(obj_IPage) -> obj_PageMain.obj_IPage_UnregisterControl -> obj_PageBase.UnregisterControl
Public Function UnregisterControl(ByVal controlKey As String) As Boolean
    Dim routeKey As Variant
    Dim routeEntry As Object
    Dim controlKeyNorm As String
    Dim routeKeysToRemove As Collection
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
    For Each routeKey In m_RouteByShape.Keys
        Set routeEntry = m_RouteByShape(routeKey)
        If VBA.LCase$(VBA.Trim$(VBA.CStr(routeEntry("ControlKey")))) = controlKeyNorm Then
            routeKeysToRemove.Add VBA.CStr(routeKey)
        End If
    Next routeKey

    For Each removeKey In routeKeysToRemove
        m_RouteByShape.Remove VBA.CStr(removeKey)
    Next removeKey

    UnregisterControl = True
End Function

' Callstack[1]: obj_PageBase.Initialize -> ResetControlActions
' Callstack[2]: obj_PageBase.Render -> ResetControlActions
' Callstack[3]: obj_PageBase.Clear -> ResetControlActions
' Callstack[4]: obj_PageBase.Dispose -> ResetControlActions
' Callstack[5]: obj_PageMain.ResetControlActions -> obj_PageMain.obj_IPage_ResetControlActions -> obj_PageBase.ResetControlActions
Public Function ResetControlActions() As Boolean
    Dim key As Variant

    If Not private_EnsureNotDisposed("ResetControlActions") Then Exit Function

    If Not m_ControlByKey Is Nothing Then
        For Each key In m_ControlByKey.Keys
            Set m_ControlByKey(key) = Nothing
        Next key
    End If

    Set m_ControlByKey = Nothing
    Set m_RouteByShape = Nothing
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

' Callstack[1]: rt_CoreActions.fn_UpdateCodeFullAndRerender -> private_ScheduleUpdateAndRerender -> rt_RestoreManager.m_SavePageSnapshots -> serializablePage.TrySerializeSnapshot(obj_PageMain) -> obj_PageMain.TrySerializeSnapshot -> m_Base.TryCollectSerializableControlSnapshots -> obj_PageBase.TryCollectSerializableControlSnapshots
' Callstack[2]: rt_CoreActions.fn_UpdateCodeDateAndRerender -> private_ScheduleUpdateAndRerender -> rt_RestoreManager.m_SavePageSnapshots -> serializablePage.TrySerializeSnapshot(obj_PageMain) -> obj_PageMain.TrySerializeSnapshot -> m_Base.TryCollectSerializableControlSnapshots -> obj_PageBase.TryCollectSerializableControlSnapshots
' Callstack[3]: rt_CoreActions.fn_UpdateCodeSizeAndRerender -> private_ScheduleUpdateAndRerender -> rt_RestoreManager.m_SavePageSnapshots -> serializablePage.TrySerializeSnapshot(obj_PageMain) -> obj_PageMain.TrySerializeSnapshot -> m_Base.TryCollectSerializableControlSnapshots -> obj_PageBase.TryCollectSerializableControlSnapshots
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
    Dim iControl As obj_IControl
    Dim iSerializable As obj_ISerializable

    If Not private_EnsureNotDisposed("TryRestoreSerializableControlSnapshots") Then Exit Function
    If snapshots Is Nothing Then
        TryRestoreSerializableControlSnapshots = True
        Exit Function
    End If

    ' В restore-фазе страница уже отрисована.
    ' Важно: нельзя удалять все serializable-контролы, иначе новые контролы,
    ' которых еще нет в старом snapshot, выпадут из runtime-реестра/route-map.
    ' Поэтому удаляем только те controlKey, которые реально присутствуют в snapshots.
    If Not private_TryResetSnapshotControlsActions(snapshots) Then Exit Function

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

        Set iControl = ex_ControlFactory.fn_CreateControlByTypeRoot(typeRoot, m_Page)
        If iControl Is Nothing Then GoTo ContinueSnapshot
        If Not private_TryCastSerializableControl(iControl, iSerializable) Then GoTo ContinueSnapshot
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

' Callstack[1]: ThisWorkbook.Workbook_BeforeClose -> rt_RestoreManager.m_SavePageSnapshots -> serializablePage.TrySerializeSnapshot(obj_PageMain) -> m_Base.TryCreateSnapshotRoot -> obj_PageBase.TryCreateSnapshotRoot
' Callstack[2]: rt_CoreActions.fn_UpdateCodeFullAndRerender -> private_ScheduleUpdateAndRerender -> rt_RestoreManager.m_SavePageSnapshots -> serializablePage.TrySerializeSnapshot(obj_PageMain) -> m_Base.TryCreateSnapshotRoot -> obj_PageBase.TryCreateSnapshotRoot
' Callstack[3]: rt_CoreActions.fn_UpdateCodeDateAndRerender -> private_ScheduleUpdateAndRerender -> rt_RestoreManager.m_SavePageSnapshots -> serializablePage.TrySerializeSnapshot(obj_PageMain) -> m_Base.TryCreateSnapshotRoot -> obj_PageBase.TryCreateSnapshotRoot
' Callstack[4]: rt_CoreActions.fn_UpdateCodeSizeAndRerender -> private_ScheduleUpdateAndRerender -> rt_RestoreManager.m_SavePageSnapshots -> serializablePage.TrySerializeSnapshot(obj_PageMain) -> m_Base.TryCreateSnapshotRoot -> obj_PageBase.TryCreateSnapshotRoot
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

' Callstack[1]: ThisWorkbook.Workbook_BeforeClose -> rt_RestoreManager.m_SavePageSnapshots -> serializablePage.TrySerializeSnapshot(obj_PageMain) -> m_Base.WriteBaseSnapshotAttributes -> obj_PageBase.WriteBaseSnapshotAttributes
' Callstack[2]: rt_CoreActions.fn_UpdateCodeFullAndRerender -> private_ScheduleUpdateAndRerender -> rt_RestoreManager.m_SavePageSnapshots -> serializablePage.TrySerializeSnapshot(obj_PageMain) -> m_Base.WriteBaseSnapshotAttributes -> obj_PageBase.WriteBaseSnapshotAttributes
' Callstack[3]: rt_CoreActions.fn_UpdateCodeDateAndRerender -> private_ScheduleUpdateAndRerender -> rt_RestoreManager.m_SavePageSnapshots -> serializablePage.TrySerializeSnapshot(obj_PageMain) -> m_Base.WriteBaseSnapshotAttributes -> obj_PageBase.WriteBaseSnapshotAttributes
' Callstack[4]: rt_CoreActions.fn_UpdateCodeSizeAndRerender -> private_ScheduleUpdateAndRerender -> rt_RestoreManager.m_SavePageSnapshots -> serializablePage.TrySerializeSnapshot(obj_PageMain) -> m_Base.WriteBaseSnapshotAttributes -> obj_PageBase.WriteBaseSnapshotAttributes
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

    If m_RouteByShape Is Nothing Then
        Set m_RouteByShape = VBA.CreateObject("Scripting.Dictionary")
        m_RouteByShape.CompareMode = 1
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
    ByRef prevDisplayAlerts As Boolean, _
    ByRef prevCalculation As XlCalculation, _
    ByRef prevStatusBar As Variant _
)
    If app Is Nothing Then Exit Sub

    prevScreenUpdating = app.ScreenUpdating
    prevEnableEvents = app.EnableEvents
    prevDisplayAlerts = app.DisplayAlerts
    prevCalculation = app.Calculation
    prevStatusBar = app.StatusBar

    app.ScreenUpdating = False
    app.EnableEvents = False
    app.DisplayAlerts = False
    app.Calculation = xlCalculationManual
    app.StatusBar = "PrototypeNew: rendering UI..."
End Sub

Private Sub private_LeaveFastRenderMode( _
    ByVal app As Application, _
    ByVal prevScreenUpdating As Boolean, _
    ByVal prevEnableEvents As Boolean, _
    ByVal prevDisplayAlerts As Boolean, _
    ByVal prevCalculation As XlCalculation, _
    ByVal prevStatusBar As Variant _
)
    If app Is Nothing Then Exit Sub

    On Error Resume Next
    app.ScreenUpdating = prevScreenUpdating
    app.EnableEvents = prevEnableEvents
    app.DisplayAlerts = prevDisplayAlerts
    app.Calculation = prevCalculation
    app.StatusBar = prevStatusBar
    On Error GoTo 0
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

Private Function private_TryResetSnapshotControlsActions(ByVal snapshots As Collection) As Boolean
    Dim item As Variant
    Dim snapshotXml As String
    Dim pageKey As String
    Dim controlKey As String
    Dim typeRoot As String
    Dim payloadXml As String
    Dim keysToRemove As Object
    Dim removeKey As Variant

    If snapshots Is Nothing Then
        private_TryResetSnapshotControlsActions = True
        Exit Function
    End If

    Set keysToRemove = VBA.CreateObject("Scripting.Dictionary")
    keysToRemove.CompareMode = 1

    For Each item In snapshots
        snapshotXml = VBA.Trim$(VBA.CStr(item))
        If VBA.Len(snapshotXml) = 0 Then GoTo ContinueSnapshot

        pageKey = VBA.vbNullString
        controlKey = VBA.vbNullString
        typeRoot = VBA.vbNullString
        payloadXml = VBA.vbNullString
        If Not Me.TryDeserializeControlSnapshotEnvelope(snapshotXml, pageKey, controlKey, typeRoot, payloadXml) Then GoTo ContinueSnapshot

        controlKey = VBA.LCase$(VBA.Trim$(controlKey))
        If VBA.Len(controlKey) = 0 Then GoTo ContinueSnapshot
        If Not keysToRemove.Exists(controlKey) Then keysToRemove.Add controlKey, True
ContinueSnapshot:
    Next item

    For Each removeKey In keysToRemove.Keys
        If Not Me.UnregisterControl(VBA.CStr(removeKey)) Then Exit Function
    Next removeKey

    private_TryResetSnapshotControlsActions = True
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
