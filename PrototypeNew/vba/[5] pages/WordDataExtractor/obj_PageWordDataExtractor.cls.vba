VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PageWordDataExtractor"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Implements obj_IPage
Implements obj_ISerializable

Private Const SERIALIZABLE_TYPE_ROOT As String = "page.worddataextractor"
Private Const SNAPSHOT_ROOT_NODE As String = "pageState"
Private Const CONTROL_SNAPSHOT_NODE As String = "controlSnapshot"
Private Const PARENT_PAGE_ID_ATTR As String = "parentPageId"
Private Const PARENT_CONFIG_CONTROL_NAME As String = "DevConfig"
Private Const PAGE_RUNTIME_OBJECT_KEY As String = "RuntimeObjects.WordDataExtractor"

Private m_PageBase As obj_PageBase
Private m_Controller As obj_IPageCtrl
Private m_ControllerObject As Object
Private m_ProfileUiPath As String
Private m_ControllerClassName As String
Private m_PendingControlSnapshots As Collection
Private m_ParentPageId As String
Private m_ParentPage As obj_IPage
Private m_IsDisposed As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
    Set m_PageBase = New obj_PageBase
    Set m_Controller = Nothing
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
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

    If m_ParentPage Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PageWordDataExtractor.Initialize: родительская страница не передана в context."
#End If
        Exit Function
    End If

    If VBA.Len(VBA.Trim$(m_ControllerClassName)) = 0 Then
        VBA.MsgBox "В профиле WordDataExtractor не указан класс контроллера.", _
            VBA.vbExclamation, "PrototypeNew / WordDataExtractor"
        Exit Function
    End If
    If Not m_PageBase.Initialize(ws, Me, uiPath, pageId) Then Exit Function
    If Not m_PageBase.RuntimeSources.SetObjectSource(PAGE_RUNTIME_OBJECT_KEY, Me) Then Exit Function

    m_ProfileUiPath = VBA.Trim$(uiPath)
    If Not private_TryCreateController( _
        m_ControllerClassName, m_Controller, _
        m_ControllerObject) Then Exit Function
    If Not m_Controller.Initialize(Me) Then Exit Function

    obj_IPage_Initialize = True
End Function

Public Function ConfigureProfileComponents( _
    ByVal uiPath As String, _
    ByVal controllerClassName As String _
) As Boolean
    Dim newController As obj_IPageCtrl
    Dim newControllerObject As Object

    uiPath = VBA.Trim$(uiPath)
    controllerClassName = VBA.Trim$(controllerClassName)
    If VBA.Len(uiPath) = 0 Then
        VBA.MsgBox "В профиле WordDataExtractor не указан UI-файл.", _
            VBA.vbExclamation, "PrototypeNew / WordDataExtractor"
        Exit Function
    End If
    If VBA.Len(controllerClassName) = 0 Then
        VBA.MsgBox "В профиле WordDataExtractor не указан класс контроллера.", _
            VBA.vbExclamation, "PrototypeNew / WordDataExtractor"
        Exit Function
    End If

    If m_PageBase Is Nothing Then
        m_ProfileUiPath = uiPath
        m_ControllerClassName = controllerClassName
        ConfigureProfileComponents = True
        Exit Function
    End If
    If Not m_PageBase.IsReady() Then
        m_ProfileUiPath = uiPath
        m_ControllerClassName = controllerClassName
        ConfigureProfileComponents = True
        Exit Function
    End If
    If VBA.StrComp(m_ProfileUiPath, uiPath, VBA.vbTextCompare) = 0 And _
        VBA.StrComp(m_ControllerClassName, controllerClassName, _
            VBA.vbTextCompare) = 0 Then
        ConfigureProfileComponents = True
        Exit Function
    End If

    If Not private_TryCreateController( _
        controllerClassName, newController, _
        newControllerObject) Then Exit Function
    If Not newController.Initialize(Me) Then
        newController.Dispose
        Exit Function
    End If
    If Not m_Controller Is Nothing Then m_Controller.Dispose
    Set m_Controller = newController
    Set m_ControllerObject = newControllerObject
    m_ProfileUiPath = uiPath
    m_ControllerClassName = controllerClassName
    m_PageBase.SetUiPath uiPath
    ConfigureProfileComponents = True
End Function

Public Function UsesProfileComponents( _
    ByVal uiPath As String, _
    ByVal controllerClassName As String _
) As Boolean
    UsesProfileComponents = _
        (VBA.StrComp(VBA.Trim$(m_ProfileUiPath), VBA.Trim$(uiPath), _
            VBA.vbTextCompare) = 0) And _
        (VBA.StrComp(VBA.Trim$(m_ControllerClassName), _
            VBA.Trim$(controllerClassName), VBA.vbTextCompare) = 0)
End Function

Private Sub obj_IPage_Dispose(Optional ByVal deleteWorksheet As Boolean = True)
    private_Dispose deleteWorksheet
End Sub

Private Function obj_IPage_RunPagePipeline() As Boolean
    Dim configControl As obj_ConfigControlVM

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageWordDataExtractor.RunPagePipeline"
#End If
    If Not m_PageBase.IsReady() Then Exit Function
    If Not private_TryResolveParentConfigControl(configControl) Then Exit Function
    If Not m_Controller Is Nothing Then
        If Not m_Controller.UpdateData(configControl) Then Exit Function
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
    If VBA.Len(normalizedReason) = 0 Then normalizedReason = "obj_PageWordDataExtractor.UpdateUiPath"

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
    If m_ControllerObject Is Nothing Then Exit Function
    Set outController = m_ControllerObject
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
        If Not private_TryGetParentPage(parentPage) Then Exit Function
        Set m_ParentPage = parentPage
    End If

    If Not m_Controller Is Nothing Then
        If private_TryResolveParentConfigControl(configControl) Then
            If Not m_Controller.UpdateData(configControl) Then Exit Function
        End If
    End If

    obj_ISerializable_TryRestoreState = True
End Function

Public Function OnExtractCommand(Optional ByVal arg As Variant) As Boolean
    Dim rawController As Object

    If m_ControllerObject Is Nothing Then Exit Function
    Set rawController = m_ControllerObject
    OnExtractCommand = VBA.CallByName( _
        rawController, "ExtractAndRender", VBA.VbMethod)
End Function

Public Function OnRenderCommand(Optional ByVal arg As Variant) As Boolean
    Dim rawController As Object

    If m_ControllerObject Is Nothing Then Exit Function
    Set rawController = m_ControllerObject
    OnRenderCommand = VBA.CallByName( _
        rawController, "Rerender", VBA.VbMethod)
End Function

Private Sub private_Dispose(Optional ByVal deleteWorksheet As Boolean = True)
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True

    On Error Resume Next
    If Not m_Controller Is Nothing Then m_Controller.Dispose
    Set m_Controller = Nothing
    Set m_ControllerObject = Nothing
    Set m_PendingControlSnapshots = Nothing
    m_ParentPageId = VBA.vbNullString
    Set m_ParentPage = Nothing
    If Not m_PageBase Is Nothing Then m_PageBase.Dispose deleteWorksheet
    Set m_PageBase = Nothing
    On Error GoTo 0
End Sub

Private Function private_TryCreateController( _
    ByVal controllerClassName As String, _
    ByRef outController As obj_IPageCtrl, _
    ByRef outControllerObject As Object _
) As Boolean
    Dim extractorController As obj_PageWordDataExtractorCtrl
    Dim searchController As obj_PageWordTextSearchCtrl

    Set outController = Nothing
    Set outControllerObject = Nothing
    Select Case VBA.LCase$(VBA.Trim$(controllerClassName))
        Case VBA.LCase$("obj_PageWordDataExtractorCtrl")
            Set extractorController = New obj_PageWordDataExtractorCtrl
            Set outController = extractorController
            Set outControllerObject = extractorController

        Case VBA.LCase$("obj_PageWordTextSearchCtrl")
            Set searchController = New obj_PageWordTextSearchCtrl
            Set outController = searchController
            Set outControllerObject = searchController

        Case Else
            VBA.MsgBox "Не поддерживается класс контроллера WordDataExtractor: " & _
                controllerClassName, VBA.vbExclamation, _
                "PrototypeNew / WordDataExtractor"
            Exit Function
    End Select
    private_TryCreateController = _
        Not outController Is Nothing And _
        Not outControllerObject Is Nothing
End Function

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
