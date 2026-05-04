VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PageMain"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean

Implements obj_IPage
Implements obj_ISerializable

Private Const SERIALIZABLE_TYPE_ROOT As String = "page.main"
Private Const SNAPSHOT_ROOT_NODE As String = "pageState"
Private Const CONTROL_SNAPSHOT_NODE As String = "controlSnapshot"
Private Const SUPPORTED_UI_PATH As String = "ui\devui.xml"

Private m_PageBase As obj_PageBase
Private m_PageMainController As obj_PageMainController
Private m_IsControllerInitialized As Boolean
Private m_PendingControlSnapshots As Collection

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
    Set m_PageBase = New obj_PageBase
    Set m_PageMainController = Nothing
    m_IsControllerInitialized = False
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
' Callstack[1]: ThisWorkbook.Workbook_Open -> ThisWorkbook.m_ResetWorkbookAndCreateMainPage -> private_ResetWorkbookAndCreateMainPage -> rt_PageManager.m_CreatePage -> obj_PageMain.obj_IPage_Initialize
' Callstack[2]: ex_Core.private_TryRecoverUiAfterUpdate -> ThisWorkbook.m_ResetWorkbookAndCreateMainPage -> private_ResetWorkbookAndCreateMainPage -> rt_PageManager.m_CreatePage -> obj_PageMain.obj_IPage_Initialize
' Callstack[3]: rt_Snapshots.m_RestorePageSnapshots -> rt_PageManager.m_CreatePage -> obj_PageMain.obj_IPage_Initialize
Private Function obj_IPage_Initialize( _
    ByVal ws As Worksheet, _
    Optional ByVal uiPath As String = VBA.vbNullString, _
    Optional ByVal pageType As Long = 1, _
    Optional ByVal pageId As String = VBA.vbNullString _
) As Boolean
    If Not m_PageBase.Initialize(ws, uiPath, pageType, pageId) Then Exit Function
    If Not private_TryValidateUiPath(m_PageBase.UiPath) Then Exit Function
    If Not private_TryPrepareRuntimeByUiPath(m_PageBase.UiPath, False) Then Exit Function

    obj_IPage_Initialize = True
End Function

' Callstack[1]: rt_PageManager.m_RenderPageById -> rt_PageManager.m_RenderPage -> obj_PageMain.obj_IPage_Render
' Callstack[2]: rt_Snapshots.m_RestorePageSnapshots(renderRestored:=True) -> rt_PageManager.m_RenderPage -> obj_PageMain.obj_IPage_Render
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
    If Not private_TryValidateUiPath(normalizedUiPath) Then Exit Function

    m_PageBase.SetUiPath normalizedUiPath
    If Not private_TryPrepareRuntimeByUiPath(normalizedUiPath, False) Then Exit Function

    normalizedReason = VBA.Trim$(reason)
    If VBA.Len(normalizedReason) = 0 Then normalizedReason = "obj_PageMain.UpdateUiPath"

    Set iPage = Me
    obj_IPage_UpdateUiPath = rt_PageManager.m_RenderPage(iPage, normalizedReason)
End Function

' Callstack[1]: ThisWorkbook.Workbook_Open -> ThisWorkbook.m_ResetWorkbookAndCreateMainPage -> private_ResetWorkbookAndCreateMainPage -> rt_PageManager.m_DisposeAllPages -> page.Dispose(False) -> obj_PageMain.obj_IPage_Dispose
' Callstack[2]: ThisWorkbook.Workbook_SheetBeforeDelete -> ex_HelpersSheet.m_RemovePageByWorksheet -> rt_PageManager.m_RemovePage -> page.Dispose(False) -> obj_PageMain.obj_IPage_Dispose
' Callstack[3]: rt_PageManager.m_RemovePageById -> rt_PageManager.m_RemovePage -> page.Dispose(deleteWorksheet) -> obj_PageMain.obj_IPage_Dispose
Private Sub obj_IPage_Dispose(Optional ByVal deleteWorksheet As Boolean = True)
    private_DisposeCore deleteWorksheet
End Sub

Private Function obj_IPage_GetPageBase() As obj_PageBase
    Set obj_IPage_GetPageBase = m_PageBase.GetPageBase()
End Function

Private Function obj_IPage_TryGetController(ByRef outController As Object) As Boolean
    Set outController = Nothing
    If Not private_TryEnsureControllerInitialized() Then Exit Function
    Set outController = m_PageMainController
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

' Callstack[1]: Shape.OnAction -> rt_Bridge.m_OnShapeClick -> rt_PageManager.m_TryGetPageByWorksheet -> page.DispatchShapeClick -> obj_PageMain.obj_IPage_DispatchShapeClick
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
' // Internal
' //
Private Sub private_DisposeCore(Optional ByVal deleteWorksheet As Boolean = True)
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True

    On Error Resume Next
    If Not m_PageMainController Is Nothing Then
        m_PageMainController.Dispose
    End If
    Set m_PageMainController = Nothing
    m_IsControllerInitialized = False
    Set m_PendingControlSnapshots = Nothing
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

Private Function private_TryValidateUiPath(ByVal uiPath As String) As Boolean
    Dim normalizedUiPath As String

    normalizedUiPath = VBA.LCase$(VBA.Trim$(uiPath))
    If VBA.Len(normalizedUiPath) = 0 Then
        private_TryValidateUiPath = True
        Exit Function
    End If

    If VBA.StrComp(normalizedUiPath, SUPPORTED_UI_PATH, VBA.vbTextCompare) = 0 Then
        private_TryValidateUiPath = True
        Exit Function
    End If

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.m_Diagnostic_LogError "PageMain: unsupported uiPath '" & VBA.Replace$(normalizedUiPath, "'", "''") & "'. Supported: '" & SUPPORTED_UI_PATH & "'."
#End If
End Function

Private Function private_TryPrepareRuntimeByUiPath( _
    ByVal uiPath As String, _
    Optional ByVal notifyChange As Boolean = False _
) As Boolean
    Dim normalizedUiPath As String

    normalizedUiPath = VBA.LCase$(VBA.Trim$(uiPath))
    If VBA.Len(normalizedUiPath) = 0 Then
        private_TryPrepareRuntimeByUiPath = True
        Exit Function
    End If

    If VBA.StrComp(normalizedUiPath, SUPPORTED_UI_PATH, VBA.vbTextCompare) <> 0 Then
        private_TryPrepareRuntimeByUiPath = True
        Exit Function
    End If

    If Not private_TryEnsureControllerInitialized() Then Exit Function
    If Not m_PageMainController.OnConfigModeChanged(notifyChange, m_PageBase) Then Exit Function

    private_TryPrepareRuntimeByUiPath = True
End Function

Private Function private_TryEnsureControllerInitialized() As Boolean
    If m_IsControllerInitialized Then
        private_TryEnsureControllerInitialized = True
        Exit Function
    End If

    If m_PageBase Is Nothing Then Exit Function
    If m_PageMainController Is Nothing Then Set m_PageMainController = New obj_PageMainController
    If m_PageMainController Is Nothing Then Exit Function

    If Not m_PageMainController.Initialize(m_PageBase) Then Exit Function
    m_IsControllerInitialized = True
    private_TryEnsureControllerInitialized = True
End Function
