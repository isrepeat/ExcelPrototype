VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PageSDB"
Option Explicit

Implements obj_IPage
Implements obj_ISerializable
Implements obj_IPageRestoreContextProvider

Private Const SERIALIZABLE_TYPE_ROOT As String = "page.supportingdocumentbuilder"
Private Const SNAPSHOT_ROOT_NODE As String = "pageState"
Private Const CONTROL_SNAPSHOT_NODE As String = "controlSnapshot"
Private Const PARENT_PAGE_ID_ATTR As String = "parentPageId"
Private Const CONFIG_CONTEXT_NODE As String = "modeConfigContext"
Private Const CONFIG_CONTEXT_ROW_NODE As String = "row"
Private Const PAGE_RUNTIME_OBJECT_KEY As String = "RuntimeObjects.SupportingDocumentBuilder"
Private Const DRAFT_VALUES_CONTAINER_NAME As String = "EventDraftValues"

Private m_PageBase As obj_PageBase
Private m_Controller As obj_PageSDBCtrl
Private m_PendingControlSnapshots As Collection
Private m_ParentPageId As String
Private m_ParentPage As obj_IPage
Private m_ConfigContext As obj_ModeConfigContext
Private m_AppliedConfigRevision As Long
Private m_HasAppliedConfig As Boolean
Private m_IsDisposed As Boolean

Private Sub Class_Initialize()
    Set m_PageBase = New obj_PageBase
End Sub

Private Sub Class_Terminate()
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    private_Dispose False
    On Error GoTo 0
End Sub

Private Function obj_IPage_Initialize( _
    ByVal ws As Worksheet, _
    Optional ByVal uiPath As String = VBA.vbNullString, _
    Optional ByVal pageId As String = VBA.vbNullString, _
    Optional ByVal Context As Object = Nothing _
) As Boolean
    Dim configContext As obj_ModeConfigContext

    If Not Context Is Nothing Then
        If TypeOf Context Is obj_ModeConfigContext Then Set configContext = Context
    End If
    If configContext Is Nothing Then
        VBA.MsgBox "SupportingDocumentBuilder requires obj_ModeConfigContext.", _
            VBA.vbExclamation, "Supporting Document Builder"
        Exit Function
    End If
    Set m_ConfigContext = configContext
    Set m_ParentPage = configContext.ParentPage
    If Not m_ParentPage Is Nothing Then _
        m_ParentPageId = VBA.LCase$(VBA.Trim$(m_ParentPage.GetPageId()))

    If Not m_PageBase.Initialize(ws, Me, uiPath, pageId) Then Exit Function
    If Not m_PageBase.RuntimeSources.SetObjectSource(PAGE_RUNTIME_OBJECT_KEY, Me) Then Exit Function
    If Not private_TryCreateConfiguredController() Then Exit Function
    If Not m_Controller.Initialize(Me) Then Exit Function
    obj_IPage_Initialize = True
End Function

Private Function obj_IPage_RunPagePipeline() As Boolean
    If Not m_PageBase.IsReady() Then Exit Function
    If Not private_TryApplyConfigContext() Then Exit Function
    If Not m_Controller.PrepareRuntime() Then Exit Function
    obj_IPage_RunPagePipeline = True
End Function

Private Function obj_IPage_Render() As Boolean
    Dim draftValues As Variant
    Dim hasDraftValues As Boolean
    Dim draftRange As Range

    If Not m_PageBase.IsReady() Then Exit Function
    ' PageBase повторно строит Input-контролы из XML и тем самым очищает их
    ' исходным value="". Ручную строку сохраняем до render и возвращаем после,
    ' чтобы Update Sheet и раскрытие preview не уничтожали введённые данные.
    If m_PageBase.TryGetLayoutContainerRange(DRAFT_VALUES_CONTAINER_NAME, draftRange) Then
        If Not draftRange Is Nothing Then
            draftValues = draftRange.Value2
            hasDraftValues = True
        End If
    End If
    If Not m_PageBase.Render() Then Exit Function
    If hasDraftValues Then
        Set draftRange = Nothing
        If Not m_PageBase.TryGetLayoutContainerRange(DRAFT_VALUES_CONTAINER_NAME, draftRange) Then Exit Function
        If draftRange Is Nothing Then Exit Function
        draftRange.Value2 = draftValues
    End If
    If Not private_TryRestorePendingControlSnapshots() Then Exit Function
    obj_IPage_Render = True
End Function

Private Sub obj_IPage_Dispose(Optional ByVal deleteWorksheet As Boolean = True)
    private_Dispose deleteWorksheet
End Sub

Private Function obj_IPage_UpdateUiPath(ByVal uiPath As String, Optional ByVal reason As String = VBA.vbNullString) As Boolean
    Dim pageRef As obj_IPage
    If Not m_PageBase.IsReady() Then Exit Function
    If VBA.Len(VBA.Trim$(uiPath)) = 0 Then Exit Function
    m_PageBase.SetUiPath VBA.Trim$(uiPath)
    If VBA.Len(VBA.Trim$(reason)) = 0 Then reason = "sdb:update-ui-path"
    Set pageRef = Me
    obj_IPage_UpdateUiPath = rt_PageManager.fn_RenderPage(pageRef, reason)
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

Private Function obj_IPage_RegisterShapeRoute(ByVal shapeName As String, ByVal controlKey As String, ByVal methodName As String, Optional ByVal hasArg As Boolean = False, Optional ByVal argValue As Variant) As Boolean
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
    Dim mainPage As obj_PageMain
    If Not private_TryApplyConfigContext() Then Exit Function
    If Not m_Controller.PrepareRuntime() Then Exit Function
    If Not m_ParentPage Is Nothing Then
        If TypeOf m_ParentPage Is obj_PageMain Then
            Set mainPage = m_ParentPage
            If Not mainPage.AttachRestoredModeConfigContext(m_ConfigContext) Then Exit Function
        End If
    End If
    obj_ISerializable_TryRestoreState = True
End Function

Private Function obj_IPageRestoreContextProvider_TryBuildRestoreContext(ByVal snapshotXml As String, ByRef outContext As Object) As Boolean
    Dim dom As Object, rootNode As Object, contextNode As Object
    Dim rowNodes As Object, rowNode As Object
    Dim configTable As obj_ConfigTable, configContext As obj_ModeConfigContext

    Set outContext = Nothing
    If Not ex_Core.fn_CustomXmlPartStore_TryLoadDomFromXml(snapshotXml, dom) Then Exit Function
    Set rootNode = dom.DocumentElement
    If rootNode Is Nothing Then Exit Function
    Set contextNode = rootNode.selectSingleNode("*[local-name()='" & CONFIG_CONTEXT_NODE & "']")
    If contextNode Is Nothing Then Exit Function
    Set configTable = New obj_ConfigTable
    If Not configTable.Initialize() Then Exit Function
    Set rowNodes = contextNode.selectNodes("*[local-name()='" & CONFIG_CONTEXT_ROW_NODE & "']")
    For Each rowNode In rowNodes
        If Not configTable.AddRow(VBA.CStr(rowNode.getAttribute("attr")), VBA.CStr(rowNode.getAttribute("key")), VBA.CStr(rowNode.getAttribute("value"))) Then Exit Function
    Next rowNode
    Set configContext = New obj_ModeConfigContext
    If Not configContext.Initialize(VBA.CStr(contextNode.getAttribute("contextId")), VBA.CStr(contextNode.getAttribute("modeId")), VBA.CStr(contextNode.getAttribute("profileId")), configTable) Then Exit Function
    Set outContext = configContext
    obj_IPageRestoreContextProvider_TryBuildRestoreContext = True
End Function

Public Function OnRenderCommand(Optional ByVal arg As Variant) As Boolean
    Dim pageRef As obj_IPage
    Set pageRef = Me
    OnRenderCommand = rt_PageManager.fn_RenderPage(pageRef, "sdb:manual-render")
End Function

Public Function EnsureModeConfigCurrent() As Boolean
    EnsureModeConfigCurrent = private_TryApplyConfigContext()
End Function

Private Function private_TryCreateConfiguredController() As Boolean
    Dim cfgParserBase As obj_CfgParserBase, entries As Collection, configMap As Object
    Dim className As String, sdbFactory As obj_SDB_Factory
    Set cfgParserBase = New obj_CfgParserBase
    If Not cfgParserBase.Initialize(m_ConfigContext.ConfigTable) Then Exit Function
    If Not cfgParserBase.TryGetConfigEntries(entries) Then Exit Function
    If Not cfgParserBase.BuildConfigDictionary(entries, configMap) Then Exit Function
    className = cfgParserBase.GetOptionalConfigValue(configMap, "SupportingDocumentBuilder.ControllerClass", VBA.vbNullString)
    If VBA.Len(VBA.Trim$(className)) = 0 Then
        VBA.MsgBox "SupportingDocumentBuilder.ControllerClass is required.", VBA.vbExclamation, "Supporting Document Builder"
        Exit Function
    End If
    Set sdbFactory = New obj_SDB_Factory
    private_TryCreateConfiguredController = sdbFactory.TryCreatePageController(className, m_Controller)
End Function

Private Function private_TryApplyConfigContext() As Boolean
    If m_Controller Is Nothing Then Exit Function
    If m_ConfigContext Is Nothing Then
        VBA.MsgBox "SupportingDocumentBuilder configuration context is unavailable.", VBA.vbExclamation, "Supporting Document Builder"
        Exit Function
    End If
    If Not m_ConfigContext.RefreshFromOwnerIfActive() Then Exit Function
    If Not m_HasAppliedConfig Or _
       m_AppliedConfigRevision <> m_ConfigContext.Revision Then
        If m_ConfigContext.ConfigTable Is Nothing Then
            VBA.MsgBox "SupportingDocumentBuilder configuration table is unavailable.", VBA.vbExclamation, "Supporting Document Builder"
            Exit Function
        End If
        If Not m_Controller.UpdateDataFromConfigTable(m_ConfigContext.ConfigTable) Then Exit Function
        m_AppliedConfigRevision = m_ConfigContext.Revision
        m_HasAppliedConfig = True
    End If
    private_TryApplyConfigContext = True
End Function

Private Function private_TrySerializeSnapshot(ByRef outSnapshotXml As String) As Boolean
    Dim dom As Object, rootNode As Object, contextNode As Object, rowNode As Object
    Dim configEntry As obj_ConfigEntry, entryIndex As Long
    Dim snapshots As Collection, item As Variant, controlNode As Object
    If Not m_PageBase.TryCreateSnapshotRoot(SNAPSHOT_ROOT_NODE, dom, rootNode) Then Exit Function
    m_PageBase.WriteBaseSnapshotAttributes rootNode
    rootNode.setAttribute PARENT_PAGE_ID_ATTR, m_ParentPageId
    Set contextNode = dom.createElement(CONFIG_CONTEXT_NODE)
    contextNode.setAttribute "contextId", m_ConfigContext.ContextId
    contextNode.setAttribute "modeId", m_ConfigContext.ModeId
    contextNode.setAttribute "profileId", m_ConfigContext.ProfileId
    For entryIndex = 1 To m_ConfigContext.ConfigTable.Items.Count
        Set configEntry = m_ConfigContext.ConfigTable.Items.Item(entryIndex)
        Set rowNode = dom.createElement(CONFIG_CONTEXT_ROW_NODE)
        rowNode.setAttribute "attr", configEntry.Attr
        rowNode.setAttribute "key", configEntry.Key
        rowNode.setAttribute "value", configEntry.Value
        contextNode.appendChild rowNode
    Next entryIndex
    rootNode.appendChild contextNode
    If Not m_PageBase.TryCollectSerializableControlSnapshots(snapshots) Then Exit Function
    If Not snapshots Is Nothing Then
        For Each item In snapshots
            Set controlNode = dom.createElement(CONTROL_SNAPSHOT_NODE)
            controlNode.Text = VBA.CStr(item)
            rootNode.appendChild controlNode
        Next item
    End If
    outSnapshotXml = VBA.CStr(dom.XML)
    private_TrySerializeSnapshot = (VBA.Len(VBA.Trim$(outSnapshotXml)) > 0)
End Function

Private Function private_TryDeserializeSnapshot(ByVal snapshotXml As String) As Boolean
    Dim dom As Object, rootNode As Object, nodes As Object, node As Object
    Dim snapshots As Collection
    If VBA.Len(VBA.Trim$(snapshotXml)) = 0 Then
        private_TryDeserializeSnapshot = True
        Exit Function
    End If
    If Not m_PageBase.TryLoadSnapshotRoot(snapshotXml, SNAPSHOT_ROOT_NODE, dom, rootNode) Then Exit Function
    m_PageBase.ReadBaseSnapshotAttributes rootNode
    m_ParentPageId = VBA.CStr(rootNode.getAttribute(PARENT_PAGE_ID_ATTR))
    Set snapshots = New Collection
    Set nodes = rootNode.selectNodes("*[local-name()='" & CONTROL_SNAPSHOT_NODE & "']")
    For Each node In nodes
        If VBA.Len(VBA.Trim$(VBA.CStr(node.Text))) > 0 Then snapshots.Add VBA.CStr(node.Text)
    Next node
    If snapshots.Count > 0 Then Set m_PendingControlSnapshots = snapshots
    private_TryDeserializeSnapshot = True
End Function

Private Function private_TryRestorePendingControlSnapshots() As Boolean
    If m_PendingControlSnapshots Is Nothing Then
        private_TryRestorePendingControlSnapshots = True
        Exit Function
    End If
    If Not m_PageBase.TryRestoreSerializableControlSnapshots(m_PendingControlSnapshots) Then Exit Function
    Set m_PendingControlSnapshots = Nothing
    private_TryRestorePendingControlSnapshots = True
End Function

Private Sub private_Dispose(Optional ByVal deleteWorksheet As Boolean = True)
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    If Not m_Controller Is Nothing Then m_Controller.Dispose
    Set m_Controller = Nothing
    Set m_PendingControlSnapshots = Nothing
    Set m_ParentPage = Nothing
    Set m_ConfigContext = Nothing
    m_HasAppliedConfig = False
    If Not m_PageBase Is Nothing Then m_PageBase.Dispose deleteWorksheet
    Set m_PageBase = Nothing
    On Error GoTo 0
End Sub
