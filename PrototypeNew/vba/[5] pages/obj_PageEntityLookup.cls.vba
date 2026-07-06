VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PageEntityLookup"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Implements obj_IPage
Implements obj_ISerializable

Private Const SERIALIZABLE_TYPE_ROOT As String = "page.entitylookup"
Private Const SNAPSHOT_ROOT_NODE As String = "pageState"
Private Const CONTROL_SNAPSHOT_NODE As String = "controlSnapshot"
Private Const PARENT_PAGE_ID_ATTR As String = "parentPageId"
Private Const PARENT_CONFIG_CONTROL_NAME As String = "DevConfig"
Private Const PAGE_RUNTIME_OBJECT_KEY As String = "RuntimeObjects.PageEntityLookup"
Private Const DICTIONARY_MISSING_MEMBER_AS_EMPTY_KEY As String = "__MissingMemberAsEmpty"

Private m_PageBase As obj_PageBase
Private m_Controller As obj_PageEntityLookupCtrl
Private m_PendingControlSnapshots As Collection
Private m_ParentPageId As String
Private m_ParentPage As obj_IPage
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

    m_ParentPageId = VBA.vbNullString
    Set m_ParentPage = Nothing

    If Not Context Is Nothing Then
        If TypeOf Context Is obj_IPage Then
            Set parentPage = Context
            m_ParentPageId = VBA.LCase$(VBA.Trim$(parentPage.GetPageId()))
            Set m_ParentPage = parentPage
        End If
    End If

    If Not m_PageBase.Initialize(ws, Me, uiPath, pageId) Then Exit Function
    If Not m_PageBase.RuntimeSources.SetObjectSource(PAGE_RUNTIME_OBJECT_KEY, Me) Then Exit Function

    Set m_Controller = New obj_PageEntityLookupCtrl
    If Not m_Controller.Initialize(Me) Then Exit Function

    obj_IPage_Initialize = True
End Function

Private Sub obj_IPage_Dispose(Optional ByVal deleteWorksheet As Boolean = True)
    private_Dispose deleteWorksheet
End Sub

Private Function obj_IPage_RunPagePipeline() As Boolean
    Dim configControl As obj_ConfigControlVM

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageEntityLookup.RunPagePipeline"
#End If
    If Not m_PageBase.IsReady() Then Exit Function
    If Not private_TryResolveParentConfigControl(configControl) Then Exit Function
    If Not m_Controller Is Nothing Then
        If Not m_Controller.UpdateData(configControl) Then Exit Function
        If Not private_SyncLookupQueryKeysFromController() Then Exit Function
        If Not m_Controller.PrepareLookupRuntime(False) Then Exit Function
    End If

    obj_IPage_RunPagePipeline = True
End Function

Private Function obj_IPage_Render() As Boolean
    If Not m_PageBase.IsReady() Then Exit Function
    If Not m_PageBase.Render() Then Exit Function
    If Not private_TryRestorePendingControlSnapshots() Then Exit Function
    obj_IPage_Render = True
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
    If VBA.Len(normalizedReason) = 0 Then normalizedReason = "obj_PageEntityLookup.UpdateUiPath"

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
    Dim configControl As obj_ConfigControlVM

    If Not m_PageBase.IsReady() Then Exit Function

    If VBA.Len(m_ParentPageId) > 0 Then
        If Not private_TryGetParentPage(parentPage) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "PageEntityLookup: parent page is not found during RestoreState. parentPageId='" & VBA.Replace$(m_ParentPageId, "'", "''") & "'."
#End If
            Exit Function
        End If
        Set m_ParentPage = parentPage
    End If

    If Not m_Controller Is Nothing Then
        If private_TryResolveParentConfigControl(configControl) Then
            If Not m_Controller.UpdateData(configControl) Then Exit Function
            If Not private_SyncLookupQueryKeysFromController() Then Exit Function
        End If
        If Not m_Controller.PrepareLookupRuntime(False) Then Exit Function
    End If

    obj_ISerializable_TryRestoreState = True
End Function

' //
' // API
' //
Public Function OnRenderCommand(Optional ByVal arg As Variant) As Boolean
    Dim pageRef As obj_IPage

    Set pageRef = Me
    OnRenderCommand = rt_PageManager.fn_RenderPage(pageRef, "entitylookup:public-render-page")
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
    If VBA.Len(VBA.Trim$(queryText)) = 0 Then
        If m_Controller Is Nothing Then Exit Function
        OnLookupInputCellChangedCommand = m_Controller.ClearLookupCandidates(False)
        Exit Function
    End If

    OnLookupInputCellChangedCommand = private_TryRunLookupSearch(lookupKey, queryText, "entitylookup:auto-search-" & private_NormalizeReasonToken(lookupKey))
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

    outSnapshotXml = VBA.vbNullString

    If Not m_PageBase.TryCreateSnapshotRoot(SNAPSHOT_ROOT_NODE, dom, rootNode) Then Exit Function

    m_PageBase.WriteBaseSnapshotAttributes rootNode
    rootNode.setAttribute PARENT_PAGE_ID_ATTR, VBA.LCase$(VBA.Trim$(m_ParentPageId))

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

Private Function private_TryResolveParentConfigControl(ByRef outConfigControl As obj_ConfigControlVM) As Boolean
    Dim parentPage As obj_IPage
    Dim rawControl As Object

    Set outConfigControl = Nothing
    If Not private_TryGetParentPage(parentPage) Then Exit Function
    If parentPage Is Nothing Then Exit Function

    Set rawControl = Nothing
    If Not parentPage.TryGetRegisteredControlByName(PARENT_CONFIG_CONTROL_NAME, rawControl) Then Exit Function
    If rawControl Is Nothing Then Exit Function
    If Not TypeOf rawControl Is obj_ConfigControlVM Then Exit Function

    Set outConfigControl = rawControl
    private_TryResolveParentConfigControl = True
End Function

Private Function private_TryEnsureControllerData() As Boolean
    Dim configControl As obj_ConfigControlVM

    If m_Controller Is Nothing Then Exit Function
    If Not private_TryResolveParentConfigControl(configControl) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageEntityLookup: failed to resolve parent DevConfig before lookup search."
#End If
        VBA.MsgBox "PrototypeNew: failed to resolve parent DevConfig before lookup search.", vbExclamation, "PrototypeNew / EntityLookup runtime"
        Exit Function
    End If

    If Not m_Controller.UpdateData(configControl) Then Exit Function
    If Not private_SyncLookupQueryKeysFromController() Then Exit Function
    private_TryEnsureControllerData = True
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
        private_TryRunLookupSearch = True
        Exit Function
    End If

    If Not private_TryEnsureControllerData() Then Exit Function
    If Not m_Controller.SearchCandidates(lookupKey, queryText, countFound, False) Then Exit Function
    private_TryRunLookupSearch = private_RerenderSelf(rerenderReason)
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
