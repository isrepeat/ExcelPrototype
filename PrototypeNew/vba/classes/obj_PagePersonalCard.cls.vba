VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PagePersonalCard"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean

Implements obj_IPage
Implements obj_ISerializable

Private Const SERIALIZABLE_TYPE_ROOT As String = "page.personalcard"
Private Const SNAPSHOT_ROOT_NODE As String = "pageState"
Private Const CONTROL_SNAPSHOT_NODE As String = "controlSnapshot"
Private Const PARENT_PAGE_ID_ATTR As String = "parentPageId"

Private m_PageBase As obj_PageBase
Private m_PendingControlSnapshots As Collection
Private m_ParentPageId As String
Private m_ParentPage As obj_IPage

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
    Set m_PageBase = New obj_PageBase
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    private_DisposeCore False
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

    If Not m_PageBase.Initialize(ws, uiPath, pageId) Then Exit Function

    obj_IPage_Initialize = True
End Function

Private Sub obj_IPage_Dispose(Optional ByVal deleteWorksheet As Boolean = True)
    private_DisposeCore deleteWorksheet
End Sub

Private Function obj_IPage_RunPagePipeline() As Boolean
    If Not m_PageBase.IsReady() Then Exit Function
    obj_IPage_RunPagePipeline = True
End Function

Private Function obj_ISerializable_TryRestoreState() As Boolean
    Dim parentPage As obj_IPage

    If Not m_PageBase.IsReady() Then Exit Function
    Set m_ParentPage = Nothing
    If VBA.Len(m_ParentPageId) = 0 Then
        obj_ISerializable_TryRestoreState = True
        Exit Function
    End If

    If Not TryGetParentPage(parentPage) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PagePersonalCard: parent page is not found during RestoreState. parentPageId='" & VBA.Replace$(m_ParentPageId, "'", "''") & "'."
#End If
        Exit Function
    End If

    Set m_ParentPage = parentPage
    obj_ISerializable_TryRestoreState = True
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
    If VBA.Len(normalizedReason) = 0 Then normalizedReason = "obj_PagePersonalCard.UpdateUiPath"

    Set iPage = Me
    obj_IPage_UpdateUiPath = rt_PageManager.m_RenderPage(iPage, normalizedReason)
End Function

Private Function obj_IPage_GetPageBase() As obj_PageBase
    Set obj_IPage_GetPageBase = m_PageBase.GetPageBase()
End Function

Private Function obj_IPage_GetPageId() As String
    obj_IPage_GetPageId = m_PageBase.PageId
End Function

Private Function obj_IPage_TryGetController(ByRef outController As Object) As Boolean
    Set outController = Nothing
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

' //
' // API
' //
Public Function TryGetParentPage(ByRef outParentPage As obj_IPage) As Boolean
    Set outParentPage = Nothing
    If Not m_ParentPage Is Nothing Then
        Set outParentPage = m_ParentPage
        TryGetParentPage = True
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(m_ParentPageId)) = 0 Then Exit Function
    If Not rt_PageManager.m_TryGetPageById(m_ParentPageId, outParentPage) Then Exit Function
    Set m_ParentPage = outParentPage
    TryGetParentPage = True
End Function

' //
' // Internal
' //
Private Sub private_DisposeCore(Optional ByVal deleteWorksheet As Boolean = True)
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True

    On Error Resume Next
    Set m_PendingControlSnapshots = Nothing
    m_ParentPageId = VBA.vbNullString
    Set m_ParentPage = Nothing
    If Not m_PageBase Is Nothing Then
        m_PageBase.Dispose deleteWorksheet
    End If
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
    Dim parentPageId As String

    Set m_PendingControlSnapshots = Nothing
    m_ParentPageId = VBA.vbNullString
    Set m_ParentPage = Nothing
    snapshotXml = VBA.Trim$(snapshotXml)
    If VBA.Len(snapshotXml) = 0 Then
        private_TryDeserializeSnapshot = True
        Exit Function
    End If

    If Not m_PageBase.TryLoadSnapshotRoot(snapshotXml, SNAPSHOT_ROOT_NODE, dom, rootNode) Then Exit Function

    m_PageBase.ReadBaseSnapshotAttributes rootNode
    parentPageId = VBA.LCase$(VBA.Trim$(VBA.CStr(rootNode.getAttribute(PARENT_PAGE_ID_ATTR))))
    m_ParentPageId = parentPageId

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
    If controlSnapshots.Count > 0 Then
        Set m_PendingControlSnapshots = controlSnapshots
    End If

    private_TryDeserializeSnapshot = True
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
