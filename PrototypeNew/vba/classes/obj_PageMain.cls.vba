VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PageMain"
Option Explicit

Implements obj_IPage
Implements obj_ISerializable

Private Const DEMO_CONFIG_VARIANT_A As String = "hospitalizationdate"
Private Const DEMO_CONFIG_VARIANT_B As String = "transfersheet"
Private Const SERIALIZABLE_TYPE_ROOT As String = "page.main"
Private Const SNAPSHOT_ROOT_NODE As String = "pageState"
Private Const CONTROL_SNAPSHOT_NODE As String = "controlSnapshot"

Private m_PageBase As obj_PageBase
Private m_DemoConfigVariant As String
Private m_PendingControlSnapshots As Collection

Private Sub Class_Initialize()
    Set m_PageBase = New obj_PageBase
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
    If Not private_PrepareRuntimeByUiPath(m_PageBase.UiPath, False) Then Exit Function

    obj_IPage_Initialize = True
End Function

' Callstack[1]: rt_PageManager.m_RenderPageById -> rt_PageManager.m_RenderPage -> obj_PageMain.obj_IPage_Render
' Callstack[2]: ex_Test.m_TEST_UpdateCurrentPage -> ex_HelpersSheet.m_TryRerenderActivePage -> rt_PageManager.m_RenderPage -> obj_PageMain.obj_IPage_Render
' Callstack[3]: ex_Test.m_TEST_SetDemoConfigVariantA -> ex_HelpersSheet.m_TryRerenderActivePage -> rt_PageManager.m_RenderPage -> obj_PageMain.obj_IPage_Render
' Callstack[4]: ex_Test.m_TEST_SetDemoConfigVariantB -> ex_HelpersSheet.m_TryRerenderActivePage -> rt_PageManager.m_RenderPage -> obj_PageMain.obj_IPage_Render
' Callstack[5]: rt_Snapshots.m_RestorePageSnapshots(renderRestored:=True) -> rt_PageManager.m_RenderPage -> obj_PageMain.obj_IPage_Render
' Callstack[6]: obj_PageMain.private_TryRerenderByDataChange -> rt_PageManager.m_RenderPage -> obj_PageMain.obj_IPage_Render
Private Function obj_IPage_Render() As Boolean
    If Not m_PageBase.IsReady() Then Exit Function
    If Not m_PageBase.Render() Then Exit Function
    If Not private_TryRestorePendingControlSnapshots() Then Exit Function
    obj_IPage_Render = True
End Function

' Callstack[1]: ex_Test.private_RenderWorksheetPage -> page.UpdateUiPath -> obj_PageMain.obj_IPage_UpdateUiPath
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
    If Not private_PrepareRuntimeByUiPath(normalizedUiPath, False) Then Exit Function

    normalizedReason = VBA.Trim$(reason)
    If VBA.Len(normalizedReason) = 0 Then normalizedReason = "obj_PageMain.UpdateUiPath"

    Set iPage = Me
    obj_IPage_UpdateUiPath = rt_PageManager.m_RenderPage(iPage, normalizedReason)
End Function

' Callstack[1]: ThisWorkbook.Workbook_Open -> ThisWorkbook.m_ResetWorkbookAndCreateMainPage -> private_ResetWorkbookAndCreateMainPage -> rt_PageManager.m_DisposeAllPages -> page.Dispose(False) -> obj_PageMain.obj_IPage_Dispose
' Callstack[2]: ThisWorkbook.Workbook_SheetBeforeDelete -> ex_HelpersSheet.m_RemovePageByWorksheet -> rt_PageManager.m_RemovePage -> page.Dispose(False) -> obj_PageMain.obj_IPage_Dispose
' Callstack[3]: rt_PageManager.m_RemovePageById -> rt_PageManager.m_RemovePage -> page.Dispose(deleteWorksheet) -> obj_PageMain.obj_IPage_Dispose
Private Sub obj_IPage_Dispose(Optional ByVal deleteWorksheet As Boolean = True)
    Set m_PendingControlSnapshots = Nothing
    m_PageBase.Dispose deleteWorksheet
End Sub

Private Function obj_IPage_GetPageBase() As obj_PageBase
    Set obj_IPage_GetPageBase = m_PageBase.GetPageBase()
End Function

' Callstack[1]: obj_PageMain.RegisterControl -> obj_PageMain.obj_IPage_RegisterControl
Private Function obj_IPage_RegisterControl(ByVal controlKey As String, ByVal controlVm As Object) As Boolean
    obj_IPage_RegisterControl = m_PageBase.RegisterControl(controlKey, controlVm)
End Function

' Callstack[1]: obj_PageMain.RegisterShapeRoute -> obj_PageMain.obj_IPage_RegisterShapeRoute
Private Function obj_IPage_RegisterShapeRoute( _
    ByVal shapeName As String, _
    ByVal controlKey As String, _
    ByVal methodName As String, _
    Optional ByVal hasArg As Boolean = False, _
    Optional ByVal argValue As Variant _
) As Boolean
    obj_IPage_RegisterShapeRoute = m_PageBase.RegisterShapeRoute(shapeName, controlKey, methodName, hasArg, argValue)
End Function

' Callstack[1]: obj_PageMain.UnregisterControl -> obj_PageMain.obj_IPage_UnregisterControl
Private Function obj_IPage_UnregisterControl(ByVal controlKey As String) As Boolean
    obj_IPage_UnregisterControl = m_PageBase.UnregisterControl(controlKey)
End Function

' Callstack[1]: obj_PageMain.ResetControlActions -> obj_PageMain.obj_IPage_ResetControlActions
Private Function obj_IPage_ResetControlActions() As Boolean
    obj_IPage_ResetControlActions = m_PageBase.ResetControlActions()
End Function

' Callstack[1]: Shape.OnAction -> rt_Bridge.m_OnShapeClick -> rt_PageManager.m_TryGetPageByWorksheet -> page.DispatchShapeClick -> obj_PageMain.obj_IPage_DispatchShapeClick
' Callstack[2]: obj_PageMain.DispatchShapeClick -> obj_PageMain.obj_IPage_DispatchShapeClick
Private Function obj_IPage_DispatchShapeClick(ByVal shapeName As String) As Boolean
    obj_IPage_DispatchShapeClick = m_PageBase.DispatchShapeClick(shapeName)
End Function

' Callstack[1]: obj_PageMain.TryCollectSerializableControlSnapshots -> obj_PageMain.obj_IPage_TryCollectSerializableControlSnapshots
Private Function obj_IPage_TryCollectSerializableControlSnapshots(ByRef outSnapshots As Collection) As Boolean
    obj_IPage_TryCollectSerializableControlSnapshots = m_PageBase.TryCollectSerializableControlSnapshots(outSnapshots)
End Function

' Callstack[1]: obj_PageMain.TryRestoreSerializableControlSnapshots -> obj_PageMain.obj_IPage_TryRestoreSerializableControlSnapshots
Private Function obj_IPage_TryRestoreSerializableControlSnapshots(ByVal snapshots As Collection) As Boolean
    obj_IPage_TryRestoreSerializableControlSnapshots = m_PageBase.TryRestoreSerializableControlSnapshots(snapshots)
End Function

Private Function obj_IPage_TryGetRegisteredControls(ByRef outControlsByKey As Object) As Boolean
    obj_IPage_TryGetRegisteredControls = m_PageBase.TryGetRegisteredControls(outControlsByKey)
End Function

Private Function obj_ISerializable_GetSerializableTypeRoot() As String
    obj_ISerializable_GetSerializableTypeRoot = SERIALIZABLE_TYPE_ROOT
End Function

' Callstack[1]: ThisWorkbook.Workbook_BeforeClose -> rt_Snapshots.m_SavePageSnapshots -> serializablePage.TrySerializeSnapshot(obj_PageMain) -> obj_PageMain.obj_ISerializable_TrySerializeSnapshot
' Callstack[2]: rt_CoreActions.m_UpdateCodeFullAndRerender -> private_ScheduleUpdateAndRerender -> rt_Snapshots.m_SavePageSnapshots -> serializablePage.TrySerializeSnapshot(obj_PageMain) -> obj_PageMain.obj_ISerializable_TrySerializeSnapshot
' Callstack[3]: rt_CoreActions.m_UpdateCodeDateAndRerender -> private_ScheduleUpdateAndRerender -> rt_Snapshots.m_SavePageSnapshots -> serializablePage.TrySerializeSnapshot(obj_PageMain) -> obj_PageMain.obj_ISerializable_TrySerializeSnapshot
' Callstack[4]: rt_CoreActions.m_UpdateCodeSizeAndRerender -> private_ScheduleUpdateAndRerender -> rt_Snapshots.m_SavePageSnapshots -> serializablePage.TrySerializeSnapshot(obj_PageMain) -> obj_PageMain.obj_ISerializable_TrySerializeSnapshot
Private Function obj_ISerializable_TrySerializeSnapshot(ByRef outSnapshotXml As String) As Boolean
    obj_ISerializable_TrySerializeSnapshot = Me.TrySerializeSnapshot(outSnapshotXml)
End Function

' Callstack[1]: ThisWorkbook.Workbook_Open -> rt_Snapshots.m_RestorePageSnapshots -> serializablePage.TryDeserializeSnapshot(obj_PageMain) -> obj_PageMain.obj_ISerializable_TryDeserializeSnapshot
' Callstack[2]: rt_CoreActions.m_RerenderLastPageAfterUpdate -> rt_Snapshots.m_RestorePageSnapshots -> serializablePage.TryDeserializeSnapshot(obj_PageMain) -> obj_PageMain.obj_ISerializable_TryDeserializeSnapshot
Private Function obj_ISerializable_TryDeserializeSnapshot(ByVal snapshotXml As String) As Boolean
    obj_ISerializable_TryDeserializeSnapshot = Me.TryDeserializeSnapshot(snapshotXml)
End Function

' //
' // API
' //
' Callstack[1]: VBA.ImmediateWindow -> obj_PageMain.Initialize -> obj_PageMain.obj_IPage_Initialize
Public Function Initialize( _
    ByVal ws As Worksheet, _
    Optional ByVal uiPath As String = VBA.vbNullString, _
    Optional ByVal pageType As Long = 1, _
    Optional ByVal pageId As String = VBA.vbNullString _
) As Boolean
    Initialize = Me.obj_IPage_Initialize(ws, uiPath, pageType, pageId)
End Function

' Callstack[1]: VBA.ImmediateWindow -> obj_PageMain.Render -> obj_PageMain.obj_IPage_Render
Public Function Render() As Boolean
    Render = Me.obj_IPage_Render()
End Function

' Callstack[1]: ex_Test.private_RenderWorksheetPage -> page.UpdateUiPath -> obj_PageMain.obj_IPage_UpdateUiPath
Public Function UpdateUiPath( _
    ByVal uiPath As String, _
    Optional ByVal reason As String = VBA.vbNullString _
) As Boolean
    UpdateUiPath = Me.obj_IPage_UpdateUiPath(uiPath, reason)
End Function

' Callstack[1]: VBA.ImmediateWindow -> obj_PageMain.Dispose -> obj_PageMain.obj_IPage_Dispose
Public Sub Dispose(Optional ByVal deleteWorksheet As Boolean = True)
    Me.obj_IPage_Dispose deleteWorksheet
End Sub

Public Function GetPageBase() As obj_PageBase
    Set GetPageBase = Me.obj_IPage_GetPageBase()
End Function

 ' Callstack[1]: VBA.ImmediateWindow -> obj_PageMain.RegisterControl -> obj_PageMain.obj_IPage_RegisterControl
Public Function RegisterControl(ByVal controlKey As String, ByVal controlVm As Object) As Boolean
    RegisterControl = Me.obj_IPage_RegisterControl(controlKey, controlVm)
End Function

' Callstack[1]: VBA.ImmediateWindow -> obj_PageMain.RegisterShapeRoute -> obj_PageMain.obj_IPage_RegisterShapeRoute
Public Function RegisterShapeRoute( _
    ByVal shapeName As String, _
    ByVal controlKey As String, _
    ByVal methodName As String, _
    Optional ByVal hasArg As Boolean = False, _
    Optional ByVal argValue As Variant _
) As Boolean
    RegisterShapeRoute = Me.obj_IPage_RegisterShapeRoute(shapeName, controlKey, methodName, hasArg, argValue)
End Function

' Callstack[1]: VBA.ImmediateWindow -> obj_PageMain.UnregisterControl -> obj_PageMain.obj_IPage_UnregisterControl
Public Function UnregisterControl(ByVal controlKey As String) As Boolean
    UnregisterControl = Me.obj_IPage_UnregisterControl(controlKey)
End Function

' Callstack[1]: VBA.ImmediateWindow -> obj_PageMain.ResetControlActions -> obj_PageMain.obj_IPage_ResetControlActions
Public Function ResetControlActions() As Boolean
    ResetControlActions = Me.obj_IPage_ResetControlActions()
End Function

' Callstack[1]: VBA.ImmediateWindow -> obj_PageMain.DispatchShapeClick -> obj_PageMain.obj_IPage_DispatchShapeClick
Public Function DispatchShapeClick(ByVal shapeName As String) As Boolean
    DispatchShapeClick = Me.obj_IPage_DispatchShapeClick(shapeName)
End Function

' Callstack[1]: VBA.ImmediateWindow -> obj_PageMain.TryCollectSerializableControlSnapshots -> obj_PageMain.obj_IPage_TryCollectSerializableControlSnapshots
Public Function TryCollectSerializableControlSnapshots(ByRef outSnapshots As Collection) As Boolean
    TryCollectSerializableControlSnapshots = Me.obj_IPage_TryCollectSerializableControlSnapshots(outSnapshots)
End Function

' Callstack[1]: VBA.ImmediateWindow -> obj_PageMain.TryRestoreSerializableControlSnapshots -> obj_PageMain.obj_IPage_TryRestoreSerializableControlSnapshots
Public Function TryRestoreSerializableControlSnapshots(ByVal snapshots As Collection) As Boolean
    TryRestoreSerializableControlSnapshots = Me.obj_IPage_TryRestoreSerializableControlSnapshots(snapshots)
End Function

Public Function TryGetRegisteredControls(ByRef outControlsByKey As Object) As Boolean
    TryGetRegisteredControls = Me.obj_IPage_TryGetRegisteredControls(outControlsByKey)
End Function

 ' Callstack[1]: VBA.ImmediateWindow -> obj_PageMain.Clear -> m_PageBase.Clear
Public Sub Clear()
    m_PageBase.Clear
End Sub

Public Function GetSerializableTypeRoot() As String
    GetSerializableTypeRoot = Me.obj_ISerializable_GetSerializableTypeRoot()
End Function

 ' Callstack[1]: ThisWorkbook.Workbook_BeforeClose -> rt_Snapshots.m_SavePageSnapshots -> serializablePage.TrySerializeSnapshot(obj_PageMain) -> obj_PageMain.obj_ISerializable_TrySerializeSnapshot -> obj_PageMain.TrySerializeSnapshot
 ' Callstack[2]: rt_CoreActions.m_UpdateCodeFullAndRerender -> private_ScheduleUpdateAndRerender -> rt_Snapshots.m_SavePageSnapshots -> serializablePage.TrySerializeSnapshot(obj_PageMain) -> obj_PageMain.obj_ISerializable_TrySerializeSnapshot -> obj_PageMain.TrySerializeSnapshot
 ' Callstack[3]: rt_CoreActions.m_UpdateCodeDateAndRerender -> private_ScheduleUpdateAndRerender -> rt_Snapshots.m_SavePageSnapshots -> serializablePage.TrySerializeSnapshot(obj_PageMain) -> obj_PageMain.obj_ISerializable_TrySerializeSnapshot -> obj_PageMain.TrySerializeSnapshot
 ' Callstack[4]: rt_CoreActions.m_UpdateCodeSizeAndRerender -> private_ScheduleUpdateAndRerender -> rt_Snapshots.m_SavePageSnapshots -> serializablePage.TrySerializeSnapshot(obj_PageMain) -> obj_PageMain.obj_ISerializable_TrySerializeSnapshot -> obj_PageMain.TrySerializeSnapshot
Public Function TrySerializeSnapshot(ByRef outSnapshotXml As String) As Boolean
    Dim dom As Object
    Dim rootNode As Object
    Dim controlSnapshots As Collection
    Dim snapshotItem As Variant
    Dim controlNode As Object
    Dim snapshotXml As String

    outSnapshotXml = VBA.vbNullString

    If Not m_PageBase.TryCreateSnapshotRoot(SNAPSHOT_ROOT_NODE, dom, rootNode) Then Exit Function

    m_PageBase.WriteBaseSnapshotAttributes rootNode
    rootNode.setAttribute "demoConfigVariant", private_GetDemoConfigVariantKey()

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
    TrySerializeSnapshot = (VBA.Len(VBA.Trim$(outSnapshotXml)) > 0)
End Function

' Callstack[1]: ThisWorkbook.Workbook_Open -> rt_Snapshots.m_RestorePageSnapshots -> serializablePage.TryDeserializeSnapshot(obj_PageMain) -> obj_PageMain.obj_ISerializable_TryDeserializeSnapshot -> obj_PageMain.TryDeserializeSnapshot
' Callstack[2]: rt_CoreActions.m_RerenderLastPageAfterUpdate -> rt_Snapshots.m_RestorePageSnapshots -> serializablePage.TryDeserializeSnapshot(obj_PageMain) -> obj_PageMain.obj_ISerializable_TryDeserializeSnapshot -> obj_PageMain.TryDeserializeSnapshot
Public Function TryDeserializeSnapshot(ByVal snapshotXml As String) As Boolean
    Dim dom As Object
    Dim rootNode As Object
    Dim restoredVariant As String
    Dim controlNodes As Object
    Dim controlNode As Object
    Dim controlSnapshots As Collection
    Dim controlSnapshotXml As String

    Set m_PendingControlSnapshots = Nothing
    snapshotXml = VBA.Trim$(snapshotXml)
    If VBA.Len(snapshotXml) = 0 Then
        TryDeserializeSnapshot = True
        Exit Function
    End If

    If Not m_PageBase.TryLoadSnapshotRoot(snapshotXml, SNAPSHOT_ROOT_NODE, dom, rootNode) Then Exit Function

    m_PageBase.ReadBaseSnapshotAttributes rootNode

    restoredVariant = VBA.LCase$(VBA.Trim$(VBA.CStr(rootNode.getAttribute("demoConfigVariant"))))
    Select Case restoredVariant
        Case DEMO_CONFIG_VARIANT_A, DEMO_CONFIG_VARIANT_B
            m_DemoConfigVariant = restoredVariant
    End Select

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

    TryDeserializeSnapshot = True
End Function

' //
' // Internal
' //
' Callstack[1]: obj_PageMain.obj_IPage_Render -> obj_PageMain.private_TryRestorePendingControlSnapshots
' Callstack[2]: obj_PageMain.private_TryRestorePendingControlSnapshots -> m_PageBase.TryRestoreSerializableControlSnapshots
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


Private Function private_PrepareRuntimeByUiPath(ByVal uiPath As String, Optional ByVal notifyChange As Boolean = False) As Boolean
    Dim normalizedUiPath As String

    normalizedUiPath = VBA.LCase$(VBA.Trim$(uiPath))
    If VBA.Len(normalizedUiPath) = 0 Then
        private_PrepareRuntimeByUiPath = True
        Exit Function
    End If

    Select Case normalizedUiPath
        Case "ui\devui.xml"
            private_PrepareRuntimeByUiPath = private_PrepareDemoConfigRuntime(m_PageBase.Worksheet, notifyChange)

        Case "ui\devtablelistui.xml", "ui\devprimitivetableui.xml", "ui\devlisttablesingleui.xml", "ui\devprofiletableui.xml"
            private_PrepareRuntimeByUiPath = private_RegisterDemoTableItems(notifyChange)

        Case "ui\devsingletableui.xml"
            private_PrepareRuntimeByUiPath = private_RegisterDemoSingleTableItems(notifyChange)

        Case "ui\devtablepartstylesui.xml"
            private_PrepareRuntimeByUiPath = private_RegisterDemoTablePartStylesItems(notifyChange)

        Case Else
            private_PrepareRuntimeByUiPath = True
    End Select
End Function


Private Function private_PrepareDemoConfigRuntime(ByVal ws As Worksheet, Optional ByVal notifyChange As Boolean = False) As Boolean
    If ws Is Nothing Then Exit Function

    If Not private_TryLoadDemoConfigVariantFromStore(ws) Then Exit Function
    m_PageBase.RuntimeSources.ResetItemsSources
    m_PageBase.RuntimeSources.ResetObjectSources
    If Not ex_Test.m_TEST_RegisterDemoConfigProfileItems(notifyChange, m_PageBase) Then Exit Function
    If Not private_RegisterDemoConfigItemsByCurrentVariant(notifyChange) Then Exit Function

    private_PrepareDemoConfigRuntime = True
End Function

Private Function private_RegisterDemoConfigItemsByCurrentVariant(Optional ByVal notifyChange As Boolean = False) As Boolean
    Select Case private_GetDemoConfigVariantKey()
        Case DEMO_CONFIG_VARIANT_B
            private_RegisterDemoConfigItemsByCurrentVariant = ex_Test.m_TEST_RegisterDemoConfigItemsVariantB(notifyChange, m_PageBase)

        Case Else
            private_RegisterDemoConfigItemsByCurrentVariant = ex_Test.m_TEST_RegisterDemoConfigItemsVariantA(notifyChange, m_PageBase)
    End Select
End Function

Private Function private_RegisterDemoTableItems(Optional ByVal notifyChange As Boolean = False) As Boolean
    Dim tables As Collection

    Set tables = ex_Test.m_TEST_BuildDemoTableItems()
    If tables Is Nothing Then Exit Function

    m_PageBase.RuntimeSources.ResetItemsSources
    m_PageBase.RuntimeSources.ResetObjectSources
    If Not m_PageBase.RuntimeSources.SetItemsSource("runtimeitems.test.tables", tables) Then Exit Function
    If notifyChange Then If Not private_TryRerenderByDataChange("itemsSource:runtimeitems.test.tables") Then Exit Function

    private_RegisterDemoTableItems = True
End Function

Private Function private_RegisterDemoSingleTableItems(Optional ByVal notifyChange As Boolean = False) As Boolean
    Dim tables As Collection

    Set tables = ex_Test.m_TEST_BuildDemoSingleTableItems()
    If tables Is Nothing Then Exit Function

    m_PageBase.RuntimeSources.ResetItemsSources
    m_PageBase.RuntimeSources.ResetObjectSources
    If Not m_PageBase.RuntimeSources.SetItemsSource("runtimeitems.test.tables", tables) Then Exit Function
    If notifyChange Then If Not private_TryRerenderByDataChange("itemsSource:runtimeitems.test.tables") Then Exit Function

    private_RegisterDemoSingleTableItems = True
End Function

Private Function private_RegisterDemoTablePartStylesItems(Optional ByVal notifyChange As Boolean = False) As Boolean
    Dim tableViews As Collection

    Set tableViews = ex_Test.m_TEST_BuildDemoTableViewItems(False, False)
    If tableViews Is Nothing Then Exit Function

    m_PageBase.RuntimeSources.ResetItemsSources
    m_PageBase.RuntimeSources.ResetObjectSources
    If Not m_PageBase.RuntimeSources.SetItemsSource("runtimeitems.test.tables", tableViews) Then Exit Function
    If Not ex_Test.m_TEST_RegisterDemoBannerItems(False, notifyChange, m_PageBase) Then Exit Function
    If notifyChange Then If Not private_TryRerenderByDataChange("itemsSource:runtimeitems.test.tables") Then Exit Function

    private_RegisterDemoTablePartStylesItems = True
End Function

Private Function private_TryRerenderByDataChange(ByVal reason As String) As Boolean
    Dim ws As Worksheet
    Dim iPage As obj_IPage

    Set ws = m_PageBase.Worksheet
    If ws Is Nothing Then Exit Function

    If Not rt_PageManager.m_TryGetPageByWorksheet(ws, iPage) Then Exit Function
    private_TryRerenderByDataChange = rt_PageManager.m_RenderPage(iPage, reason)
End Function

Private Function private_GetDemoConfigVariantKey() As String
    m_DemoConfigVariant = VBA.LCase$(VBA.Trim$(m_DemoConfigVariant))
    If VBA.Len(m_DemoConfigVariant) = 0 Then m_DemoConfigVariant = DEMO_CONFIG_VARIANT_A

    Select Case m_DemoConfigVariant
        Case DEMO_CONFIG_VARIANT_A, DEMO_CONFIG_VARIANT_B
            private_GetDemoConfigVariantKey = m_DemoConfigVariant

        Case Else
            m_DemoConfigVariant = DEMO_CONFIG_VARIANT_A
            private_GetDemoConfigVariantKey = m_DemoConfigVariant
    End Select
End Function

Private Function private_TryLoadDemoConfigVariantFromStore(ByVal ws As Worksheet) As Boolean
    Dim selectStateKey As String
    Dim storedSelectedId As String
    Dim selectControlVMStatic As obj_SelectControlVMStatic

    If ws Is Nothing Then
        VBA.MsgBox "PrototypeNew: worksheet is not specified for config profile state restore.", VBA.vbExclamation
        Exit Function
    End If

    selectStateKey = VBA.LCase$(ws.Name & "|ConfigProfilePicker")
    Set selectControlVMStatic = New obj_SelectControlVMStatic
    If Not selectControlVMStatic.TryGetSelectedId(selectStateKey, storedSelectedId) Then Exit Function

    storedSelectedId = VBA.LCase$(VBA.Trim$(storedSelectedId))
    Select Case storedSelectedId
        Case DEMO_CONFIG_VARIANT_A, DEMO_CONFIG_VARIANT_B
            m_DemoConfigVariant = storedSelectedId

        Case Else
            If VBA.Len(VBA.Trim$(m_DemoConfigVariant)) = 0 Then m_DemoConfigVariant = DEMO_CONFIG_VARIANT_A
    End Select

    private_TryLoadDemoConfigVariantFromStore = True
End Function
