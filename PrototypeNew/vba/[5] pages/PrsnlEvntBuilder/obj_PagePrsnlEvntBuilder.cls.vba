VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PagePrsnlEvntBuilder"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
#Const PRSNL_EVNT_BUILDER_HOTKEYS_ENABLED = True

Implements obj_IPage
Implements obj_ISerializable
Implements obj_IPageRestoreContextProvider

Private Const SERIALIZABLE_TYPE_ROOT As String = "page.prsnlevntbuilder"
Private Const SNAPSHOT_ROOT_NODE As String = "pageState"
Private Const CONTROL_SNAPSHOT_NODE As String = "controlSnapshot"
Private Const HOTKEYS_SNAPSHOT_NODE As String = "hotkeys"
Private Const HOTKEY_ROW_SNAPSHOT_NODE As String = "row"
Private Const PARENT_PAGE_ID_ATTR As String = "parentPageId"
Private Const CONFIG_CONTEXT_NODE As String = "modeConfigContext"
Private Const CONFIG_CONTEXT_ROW_NODE As String = "row"
Private Const PAGE_RUNTIME_OBJECT_KEY As String = "RuntimeObjects.PrsnlEvntBuilder"
Private Const LOOKUP_CANDIDATES_CONTROL_NAME As String = "LookupCandidatesTable"
Private Const HOTKEYS_RUNTIME_KEY As String = "RuntimeItems.PrsnlEvntBuilder.Hotkeys"
Private Const HOTKEYS_CONTROL_NAME As String = "SheetHotkeys"
Private Const DICTIONARY_MISSING_MEMBER_AS_EMPTY_KEY As String = "__MissingMemberAsEmpty"
Private Const EVENT_DRAFT_VALUES_CONTAINER_NAME As String = "EventDraftValues"
Private Const EVENT_DRAFT_ORDER_NO_CONTAINER_NAME As String = "EventDraftOrderNoValue"

Private m_PageBase As obj_PageBase
Private m_Controller As obj_PagePrsnlEvntBuilderCtrl
Private m_PendingControlSnapshots As Collection
Private m_ParentPageId As String
Private m_ParentPage As obj_IPage
Private m_ConfigContext As obj_ModeConfigContext
Private m_AppliedConfigRevision As Long
Private m_QueryValuesByLookupKey As Object
Private m_IsDisposed As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & TypeName(Me) & ".Class_Initialize"
#End If
    Set m_PageBase = New obj_PageBase
    Set m_Controller = Nothing
    private_ResetLookupState
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    private_Dispose False
    On Error GoTo 0
End Sub

' //
' // Interface
' //
Private Function obj_IPage_Initialize( _
    ByVal ws As Worksheet, _
    Optional ByVal uiPath As String = VBA.vbNullString, _
    Optional ByVal pageId As String = VBA.vbNullString, _
    Optional ByVal Context As Object = Nothing _
) As Boolean
    Dim parentPage As obj_IPage
    Dim configContext As obj_ModeConfigContext

    m_ParentPageId = VBA.vbNullString
    Set m_ParentPage = Nothing
    Set m_ConfigContext = Nothing
    m_AppliedConfigRevision = 0

    If Not Context Is Nothing Then
        If TypeOf Context Is obj_ModeConfigContext Then
            Set configContext = Context
            Set m_ConfigContext = configContext
            Set parentPage = configContext.ParentPage
            If Not parentPage Is Nothing Then
                m_ParentPageId = VBA.LCase$(VBA.Trim$(parentPage.GetPageId()))
                Set m_ParentPage = parentPage
            End If
        End If
    End If

    If m_ConfigContext Is Nothing Then
        VBA.MsgBox "PrototypeNew: PrsnlEvntBuilder initialization requires obj_ModeConfigContext.", _
            vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If

    If Not m_PageBase.Initialize(ws, Me, uiPath, pageId) Then Exit Function
    If Not m_PageBase.RuntimeSources.SetObjectSource(PAGE_RUNTIME_OBJECT_KEY, Me) Then Exit Function

    Set m_Controller = New obj_PagePrsnlEvntBuilderCtrl
    If Not m_Controller.Initialize(Me) Then Exit Function

    obj_IPage_Initialize = True
End Function

Private Sub obj_IPage_Dispose(Optional ByVal deleteWorksheet As Boolean = True)
    private_Dispose deleteWorksheet
End Sub

Private Function obj_IPage_RunPagePipeline() As Boolean
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PagePrsnlEvntBuilder.RunPagePipeline"
#End If
    If Not m_PageBase.IsReady() Then Exit Function
    If Not m_Controller Is Nothing Then
        If Not private_TryApplyConfigContext() Then Exit Function
        If Not private_SyncLookupQueryKeysFromController() Then Exit Function
        If Not m_Controller.PrepareRuntime(False) Then Exit Function
    End If

    obj_IPage_RunPagePipeline = True
End Function

Private Function obj_IPage_Render() As Boolean
    Dim draftValuesByTag As Object
    Dim activeDraftFieldTag As String
    Dim orderNoValues As Variant
    Dim hasOrderNoValues As Boolean
    Dim app As Application
    Dim prevCursor As Variant
    Dim hasPrevCursor As Boolean
    Dim renderOk As Boolean
    Dim perfStart As Double
    Dim perfLast As Double

    On Error GoTo EH

    If Not m_PageBase.IsReady() Then GoTo Cleanup

    Set app = Application
    ' PageBase снимает fast render mode до восстановления значений и hotkeys.
    ' Держим wait-курсор до конца page-level хвоста, чтобы Excel не мигал busy/default/busy.
    prevCursor = xlDefault
    On Error Resume Next
    prevCursor = app.Cursor
    If Err.Number <> 0 Then Err.Clear
    app.Cursor = xlWait
    hasPrevCursor = (Err.Number = 0)
    Err.Clear
    On Error GoTo EH

    perfStart = VBA.Timer
    perfLast = perfStart
#If LOGGING_DEBUG_ENABLED Then
    private_LogRenderPerfStep "prsnlevnt:render:start", perfStart, perfLast
#End If

    If Not private_TryCaptureActiveDraftFieldTag(activeDraftFieldTag) Then GoTo Cleanup
    If Not private_TryCaptureTaggedDraftValues(draftValuesByTag) Then GoTo Cleanup
#If LOGGING_DEBUG_ENABLED Then
    private_LogRenderPerfStep "prsnlevnt:render:capture-tagged-draft-values", perfStart, perfLast
#End If

    hasOrderNoValues = private_TryCaptureLayoutContainerValues(EVENT_DRAFT_ORDER_NO_CONTAINER_NAME, orderNoValues)
#If LOGGING_DEBUG_ENABLED Then
    private_LogRenderPerfStep "prsnlevnt:render:capture-order-no-values", perfStart, perfLast, "hasValues=" & VBA.CStr(hasOrderNoValues)
#End If

    If Not m_PageBase.Render() Then GoTo Cleanup
#If LOGGING_DEBUG_ENABLED Then
    private_LogRenderPerfStep "prsnlevnt:render:pagebase-render-returned", perfStart, perfLast
#End If

    If Not private_TryRestoreTaggedDraftValues(draftValuesByTag) Then GoTo Cleanup
#If LOGGING_DEBUG_ENABLED Then
    private_LogRenderPerfStep "prsnlevnt:render:restore-tagged-draft-values", perfStart, perfLast
#End If

    If hasOrderNoValues Then
        If Not private_TryRestoreLayoutContainerValues(EVENT_DRAFT_ORDER_NO_CONTAINER_NAME, orderNoValues) Then GoTo Cleanup
    End If
#If LOGGING_DEBUG_ENABLED Then
    private_LogRenderPerfStep "prsnlevnt:render:restore-order-no-values", perfStart, perfLast, "hasValues=" & VBA.CStr(hasOrderNoValues)
#End If

    If Not private_TryRestorePendingControlSnapshots() Then GoTo Cleanup
#If LOGGING_DEBUG_ENABLED Then
    private_LogRenderPerfStep "prsnlevnt:render:restore-pending-control-snapshots", perfStart, perfLast
#End If

#If PRSNL_EVNT_BUILDER_HOTKEYS_ENABLED Then
    ' HotkeysControl рендерится из RuntimeItems. После render повторно применяем
    ' его текущую таблицу, чтобы restored/default строки стали активными OnKey-привязками.
    If Not private_TryRegisterRenderedHotkeys() Then GoTo Cleanup
#If LOGGING_DEBUG_ENABLED Then
    private_LogRenderPerfStep "prsnlevnt:render:register-rendered-hotkeys", perfStart, perfLast
#End If
#End If
    If Not private_TryRestoreActiveDraftField(activeDraftFieldTag) Then GoTo Cleanup
    renderOk = True

Cleanup:
#If LOGGING_DEBUG_ENABLED Then
    If perfStart > 0 Then private_LogRenderPerfStep "prsnlevnt:render:complete", perfStart, perfLast, "ok=" & VBA.CStr(renderOk)
#End If
    If hasPrevCursor Then
        On Error Resume Next
        app.Cursor = prevCursor
        On Error GoTo 0
    End If
    obj_IPage_Render = renderOk
    Exit Function

EH:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "obj_PagePrsnlEvntBuilder.Render failed: " & Err.Description
#End If
    renderOk = False
    Resume Cleanup
End Function

Private Function obj_IPage_UpdateUiPath( _
    ByVal uiPath As String, _
    Optional ByVal reason As String = VBA.vbNullString _
) As Boolean
    Dim iPage As obj_IPage
    Dim normalizedReason As String
    Dim normalizedUiPath As String

    If Not m_PageBase.IsReady() Then Exit Function

    normalizedUiPath = VBA.Trim$(uiPath)
    If VBA.Len(normalizedUiPath) = 0 Then Exit Function

    m_PageBase.SetUiPath normalizedUiPath

    normalizedReason = VBA.Trim$(reason)
    If VBA.Len(normalizedReason) = 0 Then normalizedReason = "obj_PagePrsnlEvntBuilder.UpdateUiPath"

    Set iPage = Me
    obj_IPage_UpdateUiPath = rt_PageManager.fn_RenderPage(iPage, normalizedReason)
End Function

Private Function obj_IPage_GetPageBase() As obj_PageBase
    Set obj_IPage_GetPageBase = m_PageBase.GetPageBase()
End Function

Private Function obj_IPage_GetPageId() As String
    obj_IPage_GetPageId = m_PageBase.PageId
End Function

Private Function obj_IPage_TryGetController(ByRef outController As Object) As Boolean
    Set outController = Nothing
    If m_Controller Is Nothing Then Exit Function
    Set outController = m_Controller
    obj_IPage_TryGetController = True
End Function

Private Function obj_IPage_RegisterControl(ByVal controlKey As String, ByVal controlVm As Object) As Boolean
    obj_IPage_RegisterControl = m_PageBase.RegisterControl(controlKey, controlVm)
End Function

Private Function obj_IPage_RegisterShapeRoute( _
    ByVal shapeName As String, _
    ByVal controlKey As String, _
    ByVal methodName As String, _
    Optional ByVal hasArg As Boolean = False, _
    Optional ByVal argValue As Variant _
) As Boolean
    obj_IPage_RegisterShapeRoute = m_PageBase.RegisterShapeRoute(shapeName, controlKey, methodName, hasArg, argValue)
End Function

Private Function obj_IPage_UnregisterControl(ByVal controlKey As String) As Boolean
    obj_IPage_UnregisterControl = m_PageBase.UnregisterControl(controlKey)
End Function

Private Function obj_IPage_ResetControlActions() As Boolean
    obj_IPage_ResetControlActions = m_PageBase.ResetControlActions()
End Function

Private Function obj_IPage_DispatchShapeClick(ByVal shapeName As String) As Boolean
    obj_IPage_DispatchShapeClick = m_PageBase.DispatchShapeClick(shapeName)
End Function

Private Function obj_IPage_TryCollectSerializableControlSnapshots(ByRef outSnapshots As Collection) As Boolean
    obj_IPage_TryCollectSerializableControlSnapshots = m_PageBase.TryCollectSerializableControlSnapshots(outSnapshots)
End Function

Private Function obj_IPage_TryRestoreSerializableControlSnapshots(ByVal snapshots As Collection) As Boolean
    obj_IPage_TryRestoreSerializableControlSnapshots = m_PageBase.TryRestoreSerializableControlSnapshots(snapshots)
End Function

Private Function obj_IPage_TryGetRegisteredControls(ByRef outControlsByKey As Object) As Boolean
    obj_IPage_TryGetRegisteredControls = m_PageBase.TryGetRegisteredControls(outControlsByKey)
End Function

Private Function obj_IPage_TryGetRegisteredControlByKey(ByVal controlKey As String, ByRef outControl As Object) As Boolean
    obj_IPage_TryGetRegisteredControlByKey = m_PageBase.TryGetRegisteredControlByKey(controlKey, outControl)
End Function

Private Function obj_IPage_TryGetRegisteredControlByName(ByVal controlName As String, ByRef outControl As Object) As Boolean
    obj_IPage_TryGetRegisteredControlByName = m_PageBase.TryGetRegisteredControlByName(controlName, outControl)
End Function

Private Function obj_ISerializable_GetSerializableTypeRoot() As String
    obj_ISerializable_GetSerializableTypeRoot = SERIALIZABLE_TYPE_ROOT
End Function

Private Function obj_ISerializable_TrySerializeSnapshot(ByRef outSnapshotXml As String) As Boolean
    obj_ISerializable_TrySerializeSnapshot = private_TrySerializeSnapshot(outSnapshotXml)
End Function

Private Function obj_ISerializable_TryDeserializeSnapshot(ByVal snapshotXml As String) As Boolean
    obj_ISerializable_TryDeserializeSnapshot = private_TryDeserializeSnapshot(snapshotXml)
End Function

Private Function obj_ISerializable_TryRestoreState() As Boolean
    Dim parentPage As obj_IPage
    Dim mainPage As obj_PageMain

    If Not m_PageBase.IsReady() Then Exit Function

    If VBA.Len(m_ParentPageId) > 0 Then
        If Not private_TryGetParentPage(parentPage) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PagePrsnlEvntBuilder: parent page is not found during RestoreState. parentPageId='" & VBA.Replace$(m_ParentPageId, "'", "''") & "'."
#End If
            Exit Function
        End If
        Set m_ParentPage = parentPage
    End If

    If Not m_Controller Is Nothing Then
        ' До регистрации в Main используем точный config snapshot страницы;
        ' Main-контролы ещё не отрисованы и пока не могут отдать live-значения.
        If Not private_TryApplyConfigContext() Then Exit Function
        If Not private_SyncLookupQueryKeysFromController() Then Exit Function
        If Not m_Controller.PrepareRuntime(False) Then Exit Function
    End If

    If Not m_ParentPage Is Nothing Then
        If TypeOf m_ParentPage Is obj_PageMain Then
            Set mainPage = m_ParentPage
            If Not mainPage.AttachRestoredModeConfigContext(m_ConfigContext) Then Exit Function
        End If
    End If

    obj_ISerializable_TryRestoreState = True
End Function

Private Function obj_IPageRestoreContextProvider_TryBuildRestoreContext( _
    ByVal snapshotXml As String, _
    ByRef outContext As Object _
) As Boolean
    Dim dom As Object
    Dim rootNode As Object
    Dim contextNode As Object
    Dim rowNodes As Object
    Dim rowNode As Object
    Dim configTable As obj_ConfigTable
    Dim configContext As obj_ModeConfigContext
    Dim contextId As String
    Dim modeId As String
    Dim profileId As String

    Set outContext = Nothing
    If Not ex_Core.fn_CustomXmlPartStore_TryLoadDomFromXml(snapshotXml, dom) Then Exit Function
    Set rootNode = dom.DocumentElement
    If rootNode Is Nothing Then Exit Function
    Set contextNode = rootNode.selectSingleNode("*[local-name()='" & CONFIG_CONTEXT_NODE & "']")
    If contextNode Is Nothing Then Exit Function

    contextId = VBA.Trim$(VBA.CStr(contextNode.getAttribute("contextId")))
    modeId = VBA.Trim$(VBA.CStr(contextNode.getAttribute("modeId")))
    profileId = VBA.Trim$(VBA.CStr(contextNode.getAttribute("profileId")))
    Set configTable = New obj_ConfigTable
    If Not configTable.Initialize() Then Exit Function
    Set rowNodes = contextNode.selectNodes("*[local-name()='" & CONFIG_CONTEXT_ROW_NODE & "']")
    If Not rowNodes Is Nothing Then
        For Each rowNode In rowNodes
            If Not configTable.AddRow( _
                VBA.CStr(rowNode.getAttribute("attr")), _
                VBA.CStr(rowNode.getAttribute("key")), _
                VBA.CStr(rowNode.getAttribute("value"))) Then Exit Function
        Next rowNode
    End If

    Set configContext = New obj_ModeConfigContext
    If Not configContext.Initialize(contextId, modeId, profileId, configTable) Then Exit Function
    Set outContext = configContext
    obj_IPageRestoreContextProvider_TryBuildRestoreContext = True
End Function

' //
' // API
' //
Public Function OnRenderCommand(Optional ByVal arg As Variant) As Boolean
    Dim pageRef As obj_IPage
    Dim previousEnableEvents As Boolean

    Set pageRef = Me
    previousEnableEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo EH

    OnRenderCommand = rt_PageManager.fn_RenderPage(pageRef, "prsnlevntbuilder:public-render-page")

Cleanup:
    Application.EnableEvents = previousEnableEvents
    Exit Function

EH:
    Resume Cleanup
End Function

Public Property Get LookupQueryValues() As Object
    private_EnsureLookupQueryStorage
    Set LookupQueryValues = m_QueryValuesByLookupKey
End Property

Public Function OnLookupInputCellChangedCommand(Optional ByVal arg As Variant) As Boolean
    Dim lookupKey As String
    Dim changedCellAddress As String
    Dim queryText As String

    If VBA.IsMissing(arg) Then Exit Function
    If Not private_TryReadLookupChangePayload(arg, lookupKey, changedCellAddress) Then Exit Function
    If Not private_TryReadCellText(changedCellAddress, queryText) Then Exit Function

    private_SetLookupQueryValue lookupKey, queryText
    If m_Controller Is Nothing Then Exit Function
    ' Disabled lookup still captures pasted input values, but deliberately skips
    ' dependent-field clearing, candidate clearing, SQL requests and render.
    If Not m_Controller.IsLookupEnabled Then
        OnLookupInputCellChangedCommand = True
        Exit Function
    End If
    If Not m_Controller.ClearDependentDraftFields(lookupKey) Then Exit Function
    If VBA.Len(VBA.Trim$(queryText)) = 0 Then
        OnLookupInputCellChangedCommand = m_Controller.ClearLookupCandidates(False)
        Exit Function
    End If

    OnLookupInputCellChangedCommand = private_TryRunLookupSearch(lookupKey, queryText, "prsnlevntbuilder:auto-search-" & private_NormalizeReasonToken(lookupKey))
End Function

' //
' // Internal
' //
Private Sub private_Dispose(Optional ByVal deleteWorksheet As Boolean = True)
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True

    On Error Resume Next
    If Not m_Controller Is Nothing Then m_Controller.Dispose
    Set m_Controller = Nothing
    Set m_PendingControlSnapshots = Nothing
    m_ParentPageId = VBA.vbNullString
    Set m_ParentPage = Nothing
    Set m_ConfigContext = Nothing
    m_AppliedConfigRevision = 0
    If Not m_PageBase Is Nothing Then m_PageBase.Dispose deleteWorksheet
    Set m_PageBase = Nothing
    On Error GoTo 0
End Sub

Private Function private_TrySerializeSnapshot(ByRef outSnapshotXml As String) As Boolean
    Dim dom As Object
    Dim rootNode As Object
    Dim controlSnapshots As Collection
    Dim snapshotItem As Variant
    Dim controlNode As Object
    Dim snapshotXml As String
    Dim contextNode As Object
    Dim rowNode As Object
    Dim configTable As obj_ConfigTable
    Dim configEntry As obj_ConfigEntry
    Dim entryIndex As Long

    outSnapshotXml = VBA.vbNullString

    If Not m_PageBase.TryCreateSnapshotRoot(SNAPSHOT_ROOT_NODE, dom, rootNode) Then Exit Function

    m_PageBase.WriteBaseSnapshotAttributes rootNode
    rootNode.setAttribute PARENT_PAGE_ID_ATTR, VBA.LCase$(VBA.Trim$(m_ParentPageId))
    If m_ConfigContext Is Nothing Then Exit Function
    Set configTable = m_ConfigContext.ConfigTable
    If configTable Is Nothing Then Exit Function

    ' Контекст входит в snapshot страницы, потому что он обязателен уже на этапе
    ' Initialize и не может быть восстановлен из текущего выбора на Main.
    Set contextNode = dom.createElement(CONFIG_CONTEXT_NODE)
    contextNode.setAttribute "contextId", m_ConfigContext.ContextId
    contextNode.setAttribute "modeId", m_ConfigContext.ModeId
    contextNode.setAttribute "profileId", m_ConfigContext.ProfileId
    For entryIndex = 1 To configTable.Items.Count
        Set configEntry = configTable.Items.Item(entryIndex)
        If configEntry Is Nothing Then GoTo ContinueConfigEntry
        Set rowNode = dom.createElement(CONFIG_CONTEXT_ROW_NODE)
        rowNode.setAttribute "attr", configEntry.Attr
        rowNode.setAttribute "key", configEntry.Key
        rowNode.setAttribute "value", configEntry.Value
        contextNode.appendChild rowNode
ContinueConfigEntry:
    Next entryIndex
    rootNode.appendChild contextNode
    ' Строки хоткеев — это данные страницы, а не внутреннее состояние HotkeysControl.
    ' Сохраняем bound RuntimeItems collection здесь, чтобы будущий render контрола
    ' прочитал уже пользовательские значения.
    If Not private_TryAppendHotkeyRowsSnapshot(dom, rootNode) Then Exit Function

    Set controlSnapshots = Nothing
    If Not m_PageBase.TryCollectSerializableControlSnapshots(controlSnapshots) Then Exit Function
    If Not controlSnapshots Is Nothing Then
        For Each snapshotItem In controlSnapshots
            snapshotXml = VBA.Trim$(VBA.CStr(snapshotItem))
            If VBA.Len(snapshotXml) = 0 Then GoTo ContinueSnapshot

            Set controlNode = dom.createElement(CONTROL_SNAPSHOT_NODE)
            controlNode.Text = snapshotXml
            rootNode.appendChild controlNode
ContinueSnapshot:
        Next snapshotItem
    End If

    outSnapshotXml = VBA.CStr(dom.XML)
    private_TrySerializeSnapshot = (VBA.Len(VBA.Trim$(outSnapshotXml)) > 0)
End Function

Private Function private_TryDeserializeSnapshot(ByVal snapshotXml As String) As Boolean
    Dim dom As Object
    Dim rootNode As Object
    Dim controlNodes As Object
    Dim controlNode As Object
    Dim controlSnapshots As Collection
    Dim controlSnapshotXml As String

    Set m_PendingControlSnapshots = Nothing
    snapshotXml = VBA.Trim$(snapshotXml)
    If VBA.Len(snapshotXml) = 0 Then
        private_TryDeserializeSnapshot = True
        Exit Function
    End If

    If Not m_PageBase.TryLoadSnapshotRoot(snapshotXml, SNAPSHOT_ROOT_NODE, dom, rootNode) Then Exit Function

    private_ResetLookupState
    m_PageBase.ReadBaseSnapshotAttributes rootNode
    m_ParentPageId = VBA.LCase$(VBA.Trim$(VBA.CStr(rootNode.getAttribute(PARENT_PAGE_ID_ATTR))))
    ' Восстанавливаем строки в RuntimeSources до control snapshots и render-time
    ' регистрации. Сам HotkeysControl остается обычным XML-rendered контролом.
    If Not private_TryRestoreHotkeyRowsSnapshot(rootNode) Then Exit Function

    Set controlSnapshots = New Collection
    Set controlNodes = rootNode.selectNodes("*[local-name()='" & CONTROL_SNAPSHOT_NODE & "']")
    If Not controlNodes Is Nothing Then
        For Each controlNode In controlNodes
            controlSnapshotXml = VBA.Trim$(VBA.CStr(controlNode.Text))
            If VBA.Len(controlSnapshotXml) = 0 Then GoTo ContinueControlSnapshot
            controlSnapshots.Add controlSnapshotXml
ContinueControlSnapshot:
        Next controlNode
    End If
    If controlSnapshots.Count > 0 Then Set m_PendingControlSnapshots = controlSnapshots

    private_TryDeserializeSnapshot = True
End Function

Private Function private_TryAppendHotkeyRowsSnapshot( _
    ByVal dom As Object, _
    ByVal rootNode As Object _
) As Boolean
    Dim runtimeSources As obj_PageRuntimeSources
    Dim hotkeyRows As Collection
    Dim hotkeysNode As Object
    Dim rowNode As Object
    Dim rowItem As Variant
    Dim configEntry As obj_ConfigEntry

    ' Сериализуем текущую RuntimeItems.PrsnlEvntBuilder.Hotkeys collection в компактный
    ' page-level node. Так пользовательские примененные bindings переживают reload.
    If dom Is Nothing Then Exit Function
    If rootNode Is Nothing Then Exit Function
    If m_PageBase Is Nothing Then Exit Function
    Set runtimeSources = m_PageBase.RuntimeSources
    If runtimeSources Is Nothing Then Exit Function

    If Not runtimeSources.TryGetItemsSourceByKey(VBA.LCase$(HOTKEYS_RUNTIME_KEY), hotkeyRows, True) Then Exit Function
    If hotkeyRows Is Nothing Then
        private_TryAppendHotkeyRowsSnapshot = True
        Exit Function
    End If
    If hotkeyRows.Count = 0 Then
        private_TryAppendHotkeyRowsSnapshot = True
        Exit Function
    End If

    Set hotkeysNode = dom.createElement(HOTKEYS_SNAPSHOT_NODE)
    hotkeysNode.setAttribute "version", "1"
    hotkeysNode.setAttribute "sourceKey", VBA.LCase$(VBA.Trim$(HOTKEYS_RUNTIME_KEY))

    For Each rowItem In hotkeyRows
        Set configEntry = Nothing
        If Not IsObject(rowItem) Then GoTo ContinueHotkeyRow
        Set configEntry = rowItem
        If configEntry Is Nothing Then GoTo ContinueHotkeyRow
        If VBA.Len(VBA.Trim$(configEntry.Key)) = 0 Then GoTo ContinueHotkeyRow

        Set rowNode = dom.createElement(HOTKEY_ROW_SNAPSHOT_NODE)
        rowNode.setAttribute "action", VBA.Trim$(configEntry.Key)
        rowNode.setAttribute "hotkey", VBA.Trim$(configEntry.Value)
        hotkeysNode.appendChild rowNode

ContinueHotkeyRow:
    Next rowItem

    rootNode.appendChild hotkeysNode
    private_TryAppendHotkeyRowsSnapshot = True
End Function

Private Function private_TryRestoreHotkeyRowsSnapshot(ByVal rootNode As Object) As Boolean
    Dim runtimeSources As obj_PageRuntimeSources
    Dim hotkeysNode As Object
    Dim rowNodes As Object
    Dim rowNode As Object
    Dim hotkeyRows As Collection
    Dim actionText As String
    Dim hotkeyText As String

    ' Если старый snapshot не содержит hotkeys node, ничего не делаем: контроллер
    ' посеет defaults через PrepareRuntime/private_EnsureHotkeyRows.
    If rootNode Is Nothing Then Exit Function
    If m_PageBase Is Nothing Then Exit Function
    Set runtimeSources = m_PageBase.RuntimeSources
    If runtimeSources Is Nothing Then Exit Function

    Set hotkeysNode = rootNode.selectSingleNode("*[local-name()='" & HOTKEYS_SNAPSHOT_NODE & "']")
    If hotkeysNode Is Nothing Then
        private_TryRestoreHotkeyRowsSnapshot = True
        Exit Function
    End If

    Set hotkeyRows = New Collection
    Set rowNodes = hotkeysNode.selectNodes("*[local-name()='" & HOTKEY_ROW_SNAPSHOT_NODE & "']")
    If Not rowNodes Is Nothing Then
        For Each rowNode In rowNodes
            actionText = VBA.Trim$(VBA.CStr(rowNode.getAttribute("action")))
            hotkeyText = VBA.Trim$(VBA.CStr(rowNode.getAttribute("hotkey")))
            If VBA.Len(actionText) = 0 Then GoTo ContinueHotkeyRow
            If Not private_AddHotkeyRow(hotkeyRows, actionText, hotkeyText) Then Exit Function
ContinueHotkeyRow:
        Next rowNode
    End If

    If Not runtimeSources.RemoveItemsSource(VBA.LCase$(HOTKEYS_RUNTIME_KEY)) Then Exit Function
    If Not runtimeSources.SetItemsSource(VBA.LCase$(HOTKEYS_RUNTIME_KEY), hotkeyRows, False) Then Exit Function

    private_TryRestoreHotkeyRowsSnapshot = True
End Function

Private Function private_AddHotkeyRow( _
    ByVal hotkeyRows As Collection, _
    ByVal actionText As String, _
    ByVal hotkeyText As String _
) As Boolean
    Dim configEntry As obj_ConfigEntry

    If hotkeyRows Is Nothing Then Exit Function
    actionText = VBA.Trim$(actionText)
    If VBA.Len(actionText) = 0 Then Exit Function

    Set configEntry = New obj_ConfigEntry
    configEntry.Attr = VBA.vbNullString
    configEntry.Key = actionText
    configEntry.Value = VBA.Trim$(hotkeyText)
    hotkeyRows.Add configEntry

    private_AddHotkeyRow = True
End Function

Private Function private_TryGetParentPage(ByRef outParentPage As obj_IPage) As Boolean
    Set outParentPage = Nothing
    If Not m_ParentPage Is Nothing Then
        Set outParentPage = m_ParentPage
        private_TryGetParentPage = True
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(m_ParentPageId)) = 0 Then Exit Function
    If Not rt_PageManager.fn_TryGetPageById(m_ParentPageId, outParentPage) Then Exit Function
    Set m_ParentPage = outParentPage
    private_TryGetParentPage = True
End Function

Private Function private_TryEnsureControllerData() As Boolean
    If m_Controller Is Nothing Then Exit Function
    If Not private_TryApplyConfigContext() Then Exit Function
    If Not private_SyncLookupQueryKeysFromController() Then Exit Function
    private_TryEnsureControllerData = True
End Function

Public Function EnsureModeConfigCurrent() As Boolean
    If Not private_TryApplyConfigContext() Then Exit Function
    If Not private_SyncLookupQueryKeysFromController() Then Exit Function
    EnsureModeConfigCurrent = True
End Function

Private Function private_TryApplyConfigContext() As Boolean
    Dim configTable As obj_ConfigTable

    If m_Controller Is Nothing Then Exit Function
    If m_ConfigContext Is Nothing Then
        VBA.MsgBox "PrototypeNew: PrsnlEvntBuilder mode configuration context is unavailable.", _
            vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If
    If Not m_ConfigContext.RefreshFromOwnerIfActive() Then Exit Function
    ' При чужом активном конфиге Main сохраняется последняя таблица PEB.
    ' При своём — несохранённые изменения уже попали в неё через refresh выше.
    If m_AppliedConfigRevision = m_ConfigContext.Revision Then
        private_TryApplyConfigContext = True
        Exit Function
    End If

    Set configTable = m_ConfigContext.ConfigTable
    If configTable Is Nothing Then
        VBA.MsgBox "PrototypeNew: PrsnlEvntBuilder configuration context '" & _
            m_ConfigContext.ContextId & "' has no ConfigTable.", _
            vbExclamation, "PrototypeNew / Config runtime"
        Exit Function
    End If
    If Not m_Controller.UpdateDataFromConfigTable(configTable) Then Exit Function

    m_AppliedConfigRevision = m_ConfigContext.Revision
    private_TryApplyConfigContext = True
End Function

Private Function private_SyncLookupQueryKeysFromController() As Boolean
    Dim lookupKeys As Collection
    Dim lookupKeyObj As Variant
    Dim lookupKey As String

    private_EnsureLookupQueryStorage
    If m_Controller Is Nothing Then
        private_SyncLookupQueryKeysFromController = True
        Exit Function
    End If

    If Not m_Controller.TryGetLookupKeys(lookupKeys) Then Exit Function
    If lookupKeys Is Nothing Then Exit Function

    For Each lookupKeyObj In lookupKeys
        lookupKey = VBA.Trim$(VBA.CStr(lookupKeyObj))
        If VBA.Len(lookupKey) = 0 Then GoTo ContinueLookup
        If Not m_QueryValuesByLookupKey.Exists(lookupKey) Then m_QueryValuesByLookupKey(lookupKey) = VBA.vbNullString
ContinueLookup:
    Next lookupKeyObj

    private_SyncLookupQueryKeysFromController = True
End Function

Private Function private_TryRestorePendingControlSnapshots() As Boolean
    If m_PendingControlSnapshots Is Nothing Then
        private_TryRestorePendingControlSnapshots = True
        Exit Function
    End If

    If m_PendingControlSnapshots.Count = 0 Then
        Set m_PendingControlSnapshots = Nothing
        private_TryRestorePendingControlSnapshots = True
        Exit Function
    End If

    If Not m_PageBase.TryRestoreSerializableControlSnapshots(m_PendingControlSnapshots) Then Exit Function
    Set m_PendingControlSnapshots = Nothing
    private_TryRestorePendingControlSnapshots = True
End Function

Private Function private_TryRegisterRenderedHotkeys() As Boolean
    Dim rawControl As Object
    Dim hotkeysControl As obj_HotkeysControlVM

    ' С точки зрения страницы контрол опционален: если XML-layout больше не содержит
    ' SheetHotkeys, render страницы всё равно должен пройти. Если контрол есть,
    ' применение текущих строк активирует хоткеи сразу после render/restore.
    Set rawControl = Nothing
    If Not m_PageBase.TryGetRegisteredControlByName(HOTKEYS_CONTROL_NAME, rawControl) Then
        private_TryRegisterRenderedHotkeys = True
        Exit Function
    End If
    If rawControl Is Nothing Then
        private_TryRegisterRenderedHotkeys = True
        Exit Function
    End If
    If Not TypeOf rawControl Is obj_HotkeysControlVM Then Exit Function

    Set hotkeysControl = rawControl
    ' Auto-register после render восстанавливает page-local routes из bound itemsSource.
    ' Лист здесь не читаем: persist пользовательских правок делается только по Apply.
    private_TryRegisterRenderedHotkeys = hotkeysControl.RuntimeRegisterBoundRows(False)
End Function

Private Function private_TryReadLookupChangePayload( _
    ByVal arg As Variant, _
    ByRef outLookupKey As String, _
    ByRef outChangedCellAddress As String _
) As Boolean
    Dim payload As Object

    outLookupKey = VBA.vbNullString
    outChangedCellAddress = VBA.vbNullString

    If Not IsObject(arg) Then
        VBA.MsgBox "PrototypeNew: EntityLookup input callback requires payload with lookup key.", vbExclamation, "PrototypeNew / EntityLookup runtime"
        Exit Function
    End If

    On Error Resume Next
    Set payload = arg
    On Error GoTo 0
    If payload Is Nothing Then
        VBA.MsgBox "PrototypeNew: EntityLookup input callback payload is empty.", vbExclamation, "PrototypeNew / EntityLookup runtime"
        Exit Function
    End If
    If Not payload.Exists("ChangedCellAddress") Then
        VBA.MsgBox "PrototypeNew: EntityLookup input callback payload has no ChangedCellAddress.", vbExclamation, "PrototypeNew / EntityLookup runtime"
        Exit Function
    End If
    If Not payload.Exists("Arg") Then
        VBA.MsgBox "PrototypeNew: EntityLookup input callback payload has no lookup key Arg.", vbExclamation, "PrototypeNew / EntityLookup runtime"
        Exit Function
    End If

    outChangedCellAddress = VBA.Trim$(VBA.CStr(payload("ChangedCellAddress")))
    outLookupKey = VBA.Trim$(VBA.CStr(payload("Arg")))
    If VBA.Len(outChangedCellAddress) = 0 Or VBA.Len(outLookupKey) = 0 Then
        VBA.MsgBox "PrototypeNew: EntityLookup input callback payload has empty lookup key or cell address.", vbExclamation, "PrototypeNew / EntityLookup runtime"
        Exit Function
    End If

    private_TryReadLookupChangePayload = True
End Function

Private Function private_TryReadCellText( _
    ByVal changedCellAddress As String, _
    ByRef outValue As String _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim changedCell As Range

    outValue = VBA.vbNullString
    changedCellAddress = VBA.Trim$(changedCellAddress)
    If VBA.Len(changedCellAddress) = 0 Then Exit Function

    Set pageBase = m_PageBase.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set changedCell = ws.Range(changedCellAddress)
    On Error GoTo 0
    If changedCell Is Nothing Then Exit Function

    outValue = VBA.Trim$(VBA.CStr(changedCell.Value2))
    private_TryReadCellText = True
End Function

Private Function private_TryRunLookupSearch( _
    ByVal lookupKey As String, _
    ByVal queryText As String, _
    ByVal rerenderReason As String _
) As Boolean
    Dim countFound As Long

    If m_Controller Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(queryText)) = 0 Then
        If Not m_Controller.ClearLookupCandidates(False) Then Exit Function
        private_TryRunLookupSearch = private_TryReflowLookupCandidates(rerenderReason & ":clear")
        Exit Function
    End If

    If Not private_TryEnsureControllerData() Then Exit Function
    If Not m_Controller.SearchCandidates(lookupKey, queryText, countFound, False) Then Exit Function
    private_TryRunLookupSearch = private_TryReflowLookupCandidates(rerenderReason)
End Function

Private Function private_TryReflowLookupCandidates(ByVal reason As String) As Boolean
    Dim startedAt As Single

    If m_PageBase Is Nothing Then Exit Function
    startedAt = VBA.Timer
    If Not m_PageBase.TryReflowControl(LOOKUP_CANDIDATES_CONTROL_NAME) Then
        VBA.MsgBox _
            "PrototypeNew: partial reflow of '" & LOOKUP_CANDIDATES_CONTROL_NAME & _
            "' failed. The page was not fully re-rendered; use Update Sheet to recover.", _
            VBA.vbExclamation, _
            "PrototypeNew / partial reflow"
        Exit Function
    End If

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "perf:prsnlevntbuilder:partial-reflow control='" & _
        LOOKUP_CANDIDATES_CONTROL_NAME & "' reason='" & _
        VBA.Replace$(VBA.Trim$(reason), "'", "''") & "' ms=" & _
        VBA.Format$((VBA.Timer - startedAt) * 1000!, "0")
#End If
    private_TryReflowLookupCandidates = True
End Function

Private Function private_TryCaptureLayoutContainerValues( _
    ByVal containerName As String, _
    ByRef outValues As Variant _
) As Boolean
    Dim containerRange As Range

    If m_PageBase Is Nothing Then Exit Function
    Set containerRange = Nothing
    If Not m_PageBase.TryGetLayoutContainerRange(containerName, containerRange) Then Exit Function
    If containerRange Is Nothing Then Exit Function

    outValues = containerRange.Value2
    private_TryCaptureLayoutContainerValues = True
End Function

Private Function private_TryCaptureTaggedDraftValues(ByRef outValuesByTag As Object) As Boolean
    Dim containerRange As Range
    Dim tagValues As Object
    Dim tagEntries As Collection
    Dim tagEntryObj As Variant
    Dim tagEntry As Object
    Dim tagText As String
    Dim tagRange As Range

    Set outValuesByTag = Nothing
    If m_PageBase Is Nothing Then Exit Function

    Set tagValues = ex_Helpers.fn_CreateDictionaryTextCompare()
    Set containerRange = Nothing
    If Not m_PageBase.TryGetLayoutContainerRange(EVENT_DRAFT_VALUES_CONTAINER_NAME, containerRange) Then
        Set outValuesByTag = tagValues
        private_TryCaptureTaggedDraftValues = True
        Exit Function
    End If
    If containerRange Is Nothing Then
        Set outValuesByTag = tagValues
        private_TryCaptureTaggedDraftValues = True
        Exit Function
    End If
    ' Строка черновика может менять набор и порядок видимых колонок между профилями.
    ' Поэтому сохраняем значения по логическим field-тегам, а не по номеру колонки.
    ' Пример: значение из tags="DocNo" должно вернуться в новую ячейку DocNo,
    ' даже если DocNo после пересчета видимости сдвинулся влево/вправо.
    ' Служебные profile.* теги групповые и повторяются у нескольких колонок,
    ' поэтому для восстановления значений они намеренно игнорируются.
    Set tagEntries = Nothing
    If Not m_PageBase.TryGetLayoutTagEntriesInRange(containerRange, tagEntries, "visible") Then Exit Function
    If tagEntries Is Nothing Then
        Set outValuesByTag = tagValues
        private_TryCaptureTaggedDraftValues = True
        Exit Function
    End If
    For Each tagEntryObj In tagEntries
        If Not VBA.IsObject(tagEntryObj) Then GoTo ContinueTag
        Set tagEntry = tagEntryObj
        If tagEntry Is Nothing Then GoTo ContinueTag
        If Not tagEntry.Exists("Tag") Then GoTo ContinueTag

        tagText = VBA.Trim$(VBA.CStr(tagEntry("Tag")))
        If VBA.Len(tagText) = 0 Then GoTo ContinueTag
        If private_IsProfileTag(tagText) Then GoTo ContinueTag

        Set tagRange = Nothing
        If Not private_TryGetLayoutTagEntryRange(tagEntry, tagRange) Then GoTo ContinueTag
        If tagRange Is Nothing Then GoTo ContinueTag

        ' Снимок имеет форму: логический field-тег -> текущее значение ячейки.
        ' Номера колонок тут специально не сохраняем.
        tagValues(tagText) = tagRange.Cells(1, 1).Value2

ContinueTag:
    Next tagEntryObj

    Set outValuesByTag = tagValues
    private_TryCaptureTaggedDraftValues = True
End Function

Private Function private_TryCaptureActiveDraftFieldTag(ByRef outFieldTag As String) As Boolean
    Dim containerRange As Range
    Dim activeCellRange As Range
    Dim tagEntries As Collection
    Dim tagEntryObj As Variant
    Dim tagEntry As Object
    Dim tagRange As Range
    Dim tagText As String

    outFieldTag = VBA.vbNullString
    If m_PageBase Is Nothing Then Exit Function

    Set containerRange = Nothing
    If Not m_PageBase.TryGetLayoutContainerRange(EVENT_DRAFT_VALUES_CONTAINER_NAME, containerRange) Then
        private_TryCaptureActiveDraftFieldTag = True
        Exit Function
    End If
    If containerRange Is Nothing Then
        private_TryCaptureActiveDraftFieldTag = True
        Exit Function
    End If

    On Error Resume Next
    Set activeCellRange = Application.ActiveCell
    On Error GoTo 0
    If activeCellRange Is Nothing Then
        private_TryCaptureActiveDraftFieldTag = True
        Exit Function
    End If
    If Not (activeCellRange.Worksheet Is containerRange.Worksheet) Then
        private_TryCaptureActiveDraftFieldTag = True
        Exit Function
    End If
    If Application.Intersect(activeCellRange, containerRange) Is Nothing Then
        private_TryCaptureActiveDraftFieldTag = True
        Exit Function
    End If

    Set tagEntries = Nothing
    If Not m_PageBase.TryGetLayoutTagEntriesInRange(containerRange, tagEntries, "visible") Then Exit Function
    If tagEntries Is Nothing Then
        private_TryCaptureActiveDraftFieldTag = True
        Exit Function
    End If

    ' У одной ячейки есть field-тег (_FIO, _Rank...) и несколько profile.* тегов.
    ' Для позиционирования сохраняем только логический тег поля.
    For Each tagEntryObj In tagEntries
        If Not VBA.IsObject(tagEntryObj) Then GoTo ContinueTag
        Set tagEntry = tagEntryObj
        If tagEntry Is Nothing Then GoTo ContinueTag
        If Not tagEntry.Exists("Tag") Then GoTo ContinueTag

        tagText = VBA.Trim$(VBA.CStr(tagEntry("Tag")))
        If VBA.Len(tagText) = 0 Then GoTo ContinueTag
        If private_IsProfileTag(tagText) Then GoTo ContinueTag
        If VBA.Left$(tagText, 1) <> "_" Then GoTo ContinueTag

        Set tagRange = Nothing
        If Not private_TryGetLayoutTagEntryRange(tagEntry, tagRange) Then GoTo ContinueTag
        If tagRange Is Nothing Then GoTo ContinueTag
        If Not Application.Intersect(activeCellRange, tagRange) Is Nothing Then
            outFieldTag = tagText
            Exit For
        End If
ContinueTag:
    Next tagEntryObj

    private_TryCaptureActiveDraftFieldTag = True
End Function

Private Function private_TryRestoreActiveDraftField(ByVal fieldTag As String) As Boolean
    Dim containerRange As Range
    Dim tagEntries As Collection
    Dim tagEntryObj As Variant
    Dim tagEntry As Object
    Dim tagRange As Range
    Dim tagText As String

    fieldTag = VBA.Trim$(fieldTag)
    If VBA.Len(fieldTag) = 0 Then
        private_TryRestoreActiveDraftField = True
        Exit Function
    End If
    If m_PageBase Is Nothing Then Exit Function

    Set containerRange = Nothing
    If Not m_PageBase.TryGetLayoutContainerRange(EVENT_DRAFT_VALUES_CONTAINER_NAME, containerRange) Then
        private_TryRestoreActiveDraftField = True
        Exit Function
    End If
    If containerRange Is Nothing Then
        private_TryRestoreActiveDraftField = True
        Exit Function
    End If

    Set tagEntries = Nothing
    If Not m_PageBase.TryGetLayoutTagEntriesInRange(containerRange, tagEntries, "visible") Then Exit Function
    If tagEntries Is Nothing Then
        private_TryRestoreActiveDraftField = True
        Exit Function
    End If

    For Each tagEntryObj In tagEntries
        If Not VBA.IsObject(tagEntryObj) Then GoTo ContinueTag
        Set tagEntry = tagEntryObj
        If tagEntry Is Nothing Then GoTo ContinueTag
        If Not tagEntry.Exists("Tag") Then GoTo ContinueTag

        tagText = VBA.Trim$(VBA.CStr(tagEntry("Tag")))
        If VBA.StrComp(tagText, fieldTag, VBA.vbTextCompare) <> 0 Then GoTo ContinueTag

        Set tagRange = Nothing
        If Not private_TryGetLayoutTagEntryRange(tagEntry, tagRange) Then GoTo ContinueTag
        If tagRange Is Nothing Then GoTo ContinueTag
        Application.Goto tagRange.Cells(1, 1), False
        Exit For
ContinueTag:
    Next tagEntryObj

    ' Отсутствие поля в новой секции является штатным случаем: физическую
    ' активную ячейку тогда не переносим принудительно.
    private_TryRestoreActiveDraftField = True
End Function

Private Function private_TryRestoreTaggedDraftValues(ByVal valuesByTag As Object) As Boolean
    Dim containerRange As Range
    Dim visibleRangeByTag As Object
    Dim tagObj As Variant
    Dim tagText As String
    Dim tagRange As Range
    Dim restoredValues As Variant
    Dim hasRestoredValue As Boolean

    If valuesByTag Is Nothing Then
        private_TryRestoreTaggedDraftValues = True
        Exit Function
    End If
    If valuesByTag.Count = 0 Then
        private_TryRestoreTaggedDraftValues = True
        Exit Function
    End If
    If m_PageBase Is Nothing Then Exit Function

    Set containerRange = Nothing
    If Not m_PageBase.TryGetLayoutContainerRange(EVENT_DRAFT_VALUES_CONTAINER_NAME, containerRange) Then
        private_TryRestoreTaggedDraftValues = True
        Exit Function
    End If
    If containerRange Is Nothing Then
        private_TryRestoreTaggedDraftValues = True
        Exit Function
    End If
    ' После render старые Range уже невалидны: то же логическое поле может оказаться
    ' в другой колонке листа. Строим свежую карту только для тех field-тегов,
    ' значения которых реально были сохранены перед render.
    Set visibleRangeByTag = Nothing
    If Not private_TryBuildVisibleLayoutTagRangeMapInContainer(containerRange, valuesByTag, visibleRangeByTag) Then Exit Function
    If visibleRangeByTag Is Nothing Then
        private_TryRestoreTaggedDraftValues = True
        Exit Function
    End If

    restoredValues = containerRange.Value2
    For Each tagObj In valuesByTag.Keys
        tagText = VBA.Trim$(VBA.CStr(tagObj))
        If VBA.Len(tagText) = 0 Then GoTo ContinueTag
        If private_IsProfileTag(tagText) Then GoTo ContinueTag

        Set tagRange = Nothing
        If Not visibleRangeByTag.Exists(tagText) Then
            GoTo ContinueTag
        End If
        Set tagRange = visibleRangeByTag(tagText)
        If tagRange Is Nothing Then GoTo ContinueTag

        ' Восстанавливаем по смысловому тегу поля, а не по прежней позиции на экране.
        If Not private_TrySetContainerValueByCell(containerRange, restoredValues, tagRange.Cells(1, 1), valuesByTag(tagObj)) Then Exit Function
        hasRestoredValue = True

ContinueTag:
    Next tagObj

    If Not hasRestoredValue Then
        private_TryRestoreTaggedDraftValues = True
        Exit Function
    End If

    ' Пишем всю строку формы одним вызовом в Excel. Это заметно дешевле,
    ' чем делать NumberFormat/Value2 отдельно для каждой восстановленной ячейки.
    containerRange.NumberFormat = "@"
    containerRange.Value2 = restoredValues

    private_TryRestoreTaggedDraftValues = True
End Function

Private Function private_TryRestoreLayoutContainerValues( _
    ByVal containerName As String, _
    ByRef values As Variant _
) As Boolean
    Dim containerRange As Range

    If m_PageBase Is Nothing Then Exit Function
    Set containerRange = Nothing
    If Not m_PageBase.TryGetLayoutContainerRange(containerName, containerRange) Then Exit Function
    If containerRange Is Nothing Then Exit Function

    If Not private_ContainerValueShapeMatches(containerRange, values) Then Exit Function
    containerRange.NumberFormat = "@"
    containerRange.Value2 = values

    private_TryRestoreLayoutContainerValues = True
End Function

Private Function private_ContainerValueShapeMatches( _
    ByVal containerRange As Range, _
    ByRef values As Variant _
) As Boolean
    If containerRange Is Nothing Then Exit Function

    If VBA.IsArray(values) Then
        private_ContainerValueShapeMatches = _
            (containerRange.Rows.Count = private_ArrayRowCount(values) And _
             containerRange.Columns.Count = private_ArrayColumnCount(values))
    Else
        private_ContainerValueShapeMatches = (containerRange.Rows.Count = 1 And containerRange.Columns.Count = 1)
    End If
End Function

Private Function private_TrySetContainerValueByCell( _
    ByVal containerRange As Range, _
    ByRef containerValues As Variant, _
    ByVal targetCell As Range, _
    ByVal valueToSet As Variant _
) As Boolean
    Dim rowOffset As Long
    Dim colOffset As Long

    If containerRange Is Nothing Then Exit Function
    If targetCell Is Nothing Then Exit Function

    rowOffset = targetCell.Row - containerRange.Row + 1
    colOffset = targetCell.Column - containerRange.Column + 1
    If rowOffset <= 0 Or colOffset <= 0 Then Exit Function
    If rowOffset > containerRange.Rows.Count Or colOffset > containerRange.Columns.Count Then Exit Function

    If VBA.IsArray(containerValues) Then
        containerValues(rowOffset, colOffset) = valueToSet
    Else
        If rowOffset <> 1 Or colOffset <> 1 Then Exit Function
        containerValues = valueToSet
    End If

    private_TrySetContainerValueByCell = True
End Function

Private Function private_ArrayRowCount(ByRef values As Variant) As Long
    On Error Resume Next
    private_ArrayRowCount = UBound(values, 1) - LBound(values, 1) + 1
    On Error GoTo 0
End Function

Private Function private_ArrayColumnCount(ByRef values As Variant) As Long
    On Error Resume Next
    private_ArrayColumnCount = UBound(values, 2) - LBound(values, 2) + 1
    On Error GoTo 0
End Function

Private Function private_TryBuildVisibleLayoutTagRangeMapInContainer( _
    ByVal containerRange As Range, _
    ByVal requiredTags As Object, _
    ByRef outRangeByTag As Object _
) As Boolean
    Dim tagEntries As Collection
    Dim tagEntryObj As Variant
    Dim tagEntry As Object
    Dim tagObj As Variant
    Dim tagText As String
    Dim tagRange As Range

    Set outRangeByTag = Nothing
    If m_PageBase Is Nothing Then Exit Function
    If containerRange Is Nothing Then Exit Function
    If requiredTags Is Nothing Then
        Set outRangeByTag = ex_Helpers.fn_CreateDictionaryTextCompare()
        private_TryBuildVisibleLayoutTagRangeMapInContainer = True
        Exit Function
    End If

    Set outRangeByTag = ex_Helpers.fn_CreateDictionaryTextCompare()

    ' Берем только обычные field-теги из snapshot-а. Служебные profile.* теги
    ' описывают видимость, часто повторяются у многих колонок и не являются
    ' адресом для restore. Точечный lookup по тегу дешевле, чем обход всей
    ' runtime-карты layout-тегов, особенно после добавления длинных profile.* списков.
    For Each tagObj In requiredTags.Keys
        tagText = VBA.Trim$(VBA.CStr(tagObj))
        If VBA.Len(tagText) = 0 Then GoTo ContinueTag
        If private_IsProfileTag(tagText) Then GoTo ContinueTag
        If outRangeByTag.Exists(tagText) Then GoTo ContinueTag

        Set tagEntries = Nothing
        If Not m_PageBase.TryGetLayoutTagEntries(tagText, tagEntries, "visible") Then
            ' Поле могло быть видимым в старом профиле и исчезнуть в новом.
            ' Это нормальная ситуация при смене профиля: пропускаем тег,
            ' но продолжаем восстанавливать остальные общие поля.
            GoTo ContinueTag
        End If
        If tagEntries Is Nothing Then
            GoTo ContinueTag
        End If

        For Each tagEntryObj In tagEntries
            If Not VBA.IsObject(tagEntryObj) Then GoTo ContinueEntry
            Set tagEntry = tagEntryObj
            If tagEntry Is Nothing Then GoTo ContinueEntry

            Set tagRange = Nothing
            If Not private_TryGetLayoutTagEntryRange(tagEntry, tagRange) Then Exit Function
            If tagRange Is Nothing Then GoTo ContinueEntry
            If Application.Intersect(tagRange, containerRange) Is Nothing Then GoTo ContinueEntry

            Set outRangeByTag(tagText) = tagRange
            Exit For

ContinueEntry:
        Next tagEntryObj

ContinueTag:
    Next tagObj

    private_TryBuildVisibleLayoutTagRangeMapInContainer = True
End Function

Private Function private_IsProfileTag(ByVal tagText As String) As Boolean
    tagText = VBA.LCase$(VBA.Trim$(tagText))
    private_IsProfileTag = (VBA.Left$(tagText, VBA.Len("profile.")) = "profile.")
End Function

Private Function private_TryGetLayoutTagEntryRange( _
    ByVal tagEntry As Object, _
    ByRef outRange As Range _
) As Boolean
    Dim ws As Worksheet

    Set outRange = Nothing
    If tagEntry Is Nothing Then Exit Function
    If m_PageBase Is Nothing Then Exit Function
    Set ws = m_PageBase.Worksheet
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set outRange = ws.Range( _
        ws.Cells(VBA.CLng(tagEntry("RowStart")), VBA.CLng(tagEntry("ColStart"))), _
        ws.Cells(VBA.CLng(tagEntry("RowEnd")), VBA.CLng(tagEntry("ColEnd"))))
    On Error GoTo 0

    private_TryGetLayoutTagEntryRange = Not outRange Is Nothing
End Function

Private Function private_RerenderSelf(ByVal reasonText As String) As Boolean
    Dim pageRef As obj_IPage

    Set pageRef = Me
    private_RerenderSelf = rt_PageManager.fn_RenderPage(pageRef, reasonText)
End Function

Private Sub private_ResetLookupState()
    Set m_QueryValuesByLookupKey = ex_Helpers.fn_CreateDictionaryTextCompare()
End Sub

Private Sub private_EnsureLookupQueryStorage()
    If m_QueryValuesByLookupKey Is Nothing Then Set m_QueryValuesByLookupKey = ex_Helpers.fn_CreateDictionaryTextCompare()
    If Not m_QueryValuesByLookupKey.Exists(DICTIONARY_MISSING_MEMBER_AS_EMPTY_KEY) Then m_QueryValuesByLookupKey(DICTIONARY_MISSING_MEMBER_AS_EMPTY_KEY) = True
End Sub

Private Sub private_SetLookupQueryValue(ByVal lookupKey As String, ByVal queryText As String)
    lookupKey = VBA.Trim$(lookupKey)
    If VBA.Len(lookupKey) = 0 Then Exit Sub
    private_EnsureLookupQueryStorage
    m_QueryValuesByLookupKey(lookupKey) = VBA.Trim$(queryText)
End Sub

#If LOGGING_DEBUG_ENABLED Then
Private Sub private_LogRenderPerfStep( _
    ByVal stepName As String, _
    ByVal startedAt As Double, _
    ByRef lastAt As Double, _
    Optional ByVal details As String = "" _
)
    Dim nowAt As Double
    Dim stepMs As Double
    Dim totalMs As Double
    Dim messageText As String

    nowAt = VBA.Timer
    stepMs = private_ElapsedMs(lastAt, nowAt)
    totalMs = private_ElapsedMs(startedAt, nowAt)
    lastAt = nowAt

    messageText = "perf:render:" & stepName & _
        " stepMs=" & VBA.Format$(stepMs, "0.0") & _
        " totalMs=" & VBA.Format$(totalMs, "0.0")
    If VBA.Len(VBA.Trim$(details)) > 0 Then messageText = messageText & " " & details
    ex_Core.fn_Diagnostic_LogInfo messageText
End Sub

Private Function private_ElapsedMs(ByVal startedAt As Double, ByVal endedAt As Double) As Double
    If endedAt < startedAt Then endedAt = endedAt + 86400#
    private_ElapsedMs = (endedAt - startedAt) * 1000#
End Function
#End If

Private Function private_NormalizeReasonToken(ByVal valueText As String) As String
    Dim i As Long
    Dim ch As String
    Dim result As String

    valueText = VBA.LCase$(VBA.Trim$(valueText))
    For i = 1 To VBA.Len(valueText)
        ch = VBA.Mid$(valueText, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
            result = result & ch
        ElseIf VBA.Len(result) > 0 Then
            If VBA.Right$(result, 1) <> "-" Then result = result & "-"
        End If
    Next i
    Do While VBA.Right$(result, 1) = "-"
        result = VBA.Left$(result, VBA.Len(result) - 1)
    Loop
    If VBA.Len(result) = 0 Then result = "lookup"
    private_NormalizeReasonToken = result
End Function
