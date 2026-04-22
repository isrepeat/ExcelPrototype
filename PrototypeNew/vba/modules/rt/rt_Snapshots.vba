Attribute VB_Name = "rt_Snapshots"
Option Explicit

Private Const SNAPSHOT_PAYLOAD_NODE As String = "payload"

Private Const PAGE_SNAPSHOT_NS As String = "urn:excelprototype:runtime-page-snapshots:v1"
Private Const PAGE_SNAPSHOT_ROOT As String = "pageSnapshots"
Private Const PAGE_SNAPSHOT_NODE As String = "page"

Private Const RUNTIME_GLOBALS_NS As String = "urn:excelprototype:runtime-globals:v1"
Private Const RUNTIME_GLOBALS_ROOT As String = "runtimeGlobals"
Private Const RUNTIME_GLOBALS_MODULE_NODE As String = "module"
Private Const RUNTIME_GLOBALS_MODULE_NAME_ATTR As String = "name"
Private Const RUNTIME_GLOBALS_MODULE_SNAPSHOT_NODE As String = "snapshot"
Private Const MODULE_NAME_PAGE_MANAGER As String = "rt_PageManager"

Private g_IsRuntimeStateRestoreRunning As Boolean

' //
' // API
' //
' Callstack[1]: ThisWorkbook.Workbook_BeforeClose -> rt_Snapshots.m_SavePageSnapshots
' Callstack[2]: rt_CoreActions.m_UpdateCodeFullAndRerender -> private_ScheduleUpdateAndRerender -> rt_Snapshots.m_SavePageSnapshots
' Callstack[3]: rt_CoreActions.m_UpdateCodeDateAndRerender -> private_ScheduleUpdateAndRerender -> rt_Snapshots.m_SavePageSnapshots
' Callstack[4]: rt_CoreActions.m_UpdateCodeSizeAndRerender -> private_ScheduleUpdateAndRerender -> rt_Snapshots.m_SavePageSnapshots
Public Function m_SavePageSnapshots() As Boolean
    Dim pages As Collection
    Dim pageItem As Variant
    Dim page As obj_IPage
    Dim pageBase As obj_PageBase
    Dim serializablePage As obj_ISerializable
    Dim typeRoot As String
    Dim payloadXml As String
    Dim snapshotXml As String
    Dim snapshots As Collection

    Set snapshots = New Collection
    If Not private_TryCollectAllPages(pages) Then Exit Function

    For Each pageItem In pages
        Set page = pageItem
        If page Is Nothing Then GoTo ContinuePage

        Set pageBase = page.GetPageBase()
        If pageBase Is Nothing Then GoTo ContinuePage

        If Not private_TryCastSerializablePage(page, serializablePage) Then
            VBA.MsgBox "Snapshots: page class '" & VBA.TypeName(page) & "' must implement obj_ISerializable.", VBA.vbExclamation
            Exit Function
        End If

        typeRoot = VBA.LCase$(VBA.Trim$(serializablePage.GetSerializableTypeRoot()))
        If VBA.Len(typeRoot) = 0 Then
            VBA.MsgBox "Snapshots: page serializable type root is empty for '" & VBA.TypeName(page) & "'.", VBA.vbExclamation
            Exit Function
        End If

        payloadXml = VBA.vbNullString
        If Not serializablePage.TrySerializeSnapshot(payloadXml) Then Exit Function

        snapshotXml = VBA.vbNullString
        If Not pageBase.TrySerializePageSnapshotEnvelope(typeRoot, payloadXml, snapshotXml) Then Exit Function
        If VBA.Len(VBA.Trim$(snapshotXml)) = 0 Then GoTo ContinuePage

        snapshots.Add snapshotXml

ContinuePage:
    Next pageItem

    If Not private_TrySaveSnapshotXmlCollection(PAGE_SNAPSHOT_NS, PAGE_SNAPSHOT_ROOT, PAGE_SNAPSHOT_NODE, snapshots) Then Exit Function
    m_SavePageSnapshots = True
End Function


' Callstack[1]: rt_CoreActions.private_ScheduleUpdateAndRerender -> rt_Snapshots.m_SaveRuntimeGlobalsSnapshot
' Callstack[2]: ThisWorkbook.Workbook_BeforeClose -> rt_Snapshots.m_SaveRuntimeGlobalsSnapshot
Public Function m_SaveRuntimeGlobalsSnapshot() As Boolean
    Dim dom As Object
    Dim rootNode As Object
    Dim partObj As Object

    If Not ex_CustomXmlPartStore.m_TryCreateEmptyDom(RUNTIME_GLOBALS_ROOT, RUNTIME_GLOBALS_NS, dom) Then Exit Function
    Set rootNode = dom.DocumentElement
    If rootNode Is Nothing Then
        VBA.MsgBox "Snapshots: runtime globals root node is missing.", VBA.vbExclamation
        Exit Function
    End If

    If Not private_TryAppendModuleSnapshot(rootNode, MODULE_NAME_PAGE_MANAGER) Then Exit Function

    If Not ex_CustomXmlPartStore.m_TryFindPartByNamespace(RUNTIME_GLOBALS_NS, partObj) Then Exit Function
    If Not ex_CustomXmlPartStore.m_TrySaveDom(dom, partObj) Then Exit Function

    m_SaveRuntimeGlobalsSnapshot = True
End Function


' Callstack[1]: rt_CoreActions.m_RerenderLastPageAfterUpdate -> rt_Snapshots.m_RestoreRuntimeGlobalsSnapshot
' Callstack[2]: ThisWorkbook.Workbook_Open -> rt_Snapshots.m_RestoreRuntimeGlobalsSnapshot
Public Function m_RestoreRuntimeGlobalsSnapshot() As Boolean
    Dim partObj As Object
    Dim dom As Object
    Dim rootNode As Object
    Dim moduleNodes As Object
    Dim moduleNode As Object
    Dim snapshotNode As Object
    Dim moduleName As String
    Dim snapshotXml As String

    If Not ex_CustomXmlPartStore.m_TryFindPartByNamespace(RUNTIME_GLOBALS_NS, partObj) Then Exit Function
    If partObj Is Nothing Then
        m_RestoreRuntimeGlobalsSnapshot = True
        Exit Function
    End If

    If Not ex_CustomXmlPartStore.m_TryLoadPartDom(partObj, dom) Then Exit Function
    Set rootNode = dom.DocumentElement
    If rootNode Is Nothing Then
        m_RestoreRuntimeGlobalsSnapshot = True
        Exit Function
    End If

    Set moduleNodes = rootNode.selectNodes("*[local-name()='" & RUNTIME_GLOBALS_MODULE_NODE & "']")
    If moduleNodes Is Nothing Then
        m_RestoreRuntimeGlobalsSnapshot = True
        Exit Function
    End If

    For Each moduleNode In moduleNodes
        moduleName = VBA.Trim$(VBA.CStr(moduleNode.getAttribute(RUNTIME_GLOBALS_MODULE_NAME_ATTR)))
        If VBA.Len(moduleName) = 0 Then GoTo ContinueModule

        snapshotXml = VBA.vbNullString
        Set snapshotNode = moduleNode.selectSingleNode("*[local-name()='" & RUNTIME_GLOBALS_MODULE_SNAPSHOT_NODE & "']")
        If Not snapshotNode Is Nothing Then
            snapshotXml = VBA.CStr(snapshotNode.Text)
        End If

        If Not private_TryDeserializeRuntimeModuleSnapshot(moduleName, snapshotXml) Then Exit Function
ContinueModule:
    Next moduleNode

    m_RestoreRuntimeGlobalsSnapshot = True
End Function


' Callstack[1]: ex_Core.private_QueueRuntimeStateRestoreAfterUpdate (Application.OnTime) -> rt_Snapshots.m_RunDeferredRuntimeStateRestore
Public Sub m_RunDeferredRuntimeStateRestore()
    Dim restoredPagesCount As Long
    ' Выполняется через OnTime после hot-update/recovery.
    ' Цель: стабилизировать runtime после перекомпиляции модулей и отложенных сбросов state.
    Call m_TryRestoreRuntimeStateFromSnapshots("deferred:on-time", restoredPagesCount)
End Sub


' Callstack[1]: explicit recovery entrypoints -> rt_Snapshots.m_TryRestoreRuntimeStateFromSnapshots
Public Function m_TryRestoreRuntimeStateFromSnapshots( _
    Optional ByVal reasonText As String = VBA.vbNullString, _
    Optional ByRef outRestoredPagesCount As Long = 0 _
) As Boolean
    reasonText = VBA.Trim$(reasonText)
    If VBA.Len(reasonText) = 0 Then reasonText = "unknown"

    outRestoredPagesCount = 0
    If g_IsRuntimeStateRestoreRunning Then Exit Function

    g_IsRuntimeStateRestoreRunning = True
    On Error GoTo EH_RESTORE

    ex_Core.m_LogInfo "snapshots:restore-runtime-state start reason='" & VBA.Replace$(reasonText, "'", "''") & "'"

    If private_HasRuntimePageForActiveWorksheet() Then
        If Not m_RestoreRuntimeGlobalsSnapshot() Then GoTo RestoreFailed
        ex_Core.m_LogInfo "snapshots:restore-runtime-state done reason='" & VBA.Replace$(reasonText, "'", "''") & "' restoredPages=0"
        m_TryRestoreRuntimeStateFromSnapshots = True
        GoTo Cleanup
    End If

    If Not m_RestorePageSnapshots(True, "snapshots:restore-runtime-state", outRestoredPagesCount) Then GoTo RestoreFailed
    If outRestoredPagesCount <= 0 Then GoTo RestoreFailed
    If Not m_RestoreRuntimeGlobalsSnapshot() Then GoTo RestoreFailed

    ex_Core.m_LogInfo "snapshots:restore-runtime-state done reason='" & VBA.Replace$(reasonText, "'", "''") & "' restoredPages=" & VBA.CStr(outRestoredPagesCount)
    m_TryRestoreRuntimeStateFromSnapshots = True

Cleanup:
    g_IsRuntimeStateRestoreRunning = False
    Exit Function

RestoreFailed:
    ' Нюанс "редкого провала":
    ' deferred restore может стартовать в момент, когда runtime-карта страниц уже пустая,
    ' а snapshots еще не содержат актуального Main после recovery.
    ' В этом случае жестко восстанавливаем Main через общий API, чтобы не оставлять
    ' приложение в состоянии page-not-found до следующего Workbook_Open.
    If private_TryFallbackRestoreByResettingMainPage(reasonText, outRestoredPagesCount) Then
        m_TryRestoreRuntimeStateFromSnapshots = True
        GoTo Cleanup
    End If
    ex_Core.m_LogError "snapshots:restore-runtime-state failed reason='" & VBA.Replace$(reasonText, "'", "''") & "' restoredPages=" & VBA.CStr(outRestoredPagesCount)
    GoTo Cleanup

EH_RESTORE:
    ex_Core.m_LogError "snapshots:restore-runtime-state exception reason='" & VBA.Replace$(reasonText, "'", "''") & "' err='" & VBA.Replace$(Err.Description, "'", "''") & "'"
    Resume Cleanup
End Function


' Callstack[1]: rt_Snapshots.m_TryRestoreRuntimeStateFromSnapshots -> private_TryFallbackRestoreByResettingMainPage
Private Function private_TryFallbackRestoreByResettingMainPage( _
    ByVal reasonText As String, _
    ByRef outRestoredPagesCount As Long _
) As Boolean
    Dim savePagesOk As Boolean
    Dim saveRuntimeOk As Boolean
    Dim errDescription As String

    reasonText = VBA.Trim$(reasonText)
    If VBA.Len(reasonText) = 0 Then reasonText = "unknown"

    ' Fallback не пытается "дочинять" частичное состояние.
    ' Мы создаем новый Main как единую точку правды, затем пересохраняем snapshots/globals.
    ' Так следующий цикл restore уже опирается на консистентный checkpoint.
    On Error Resume Next
    If Not ThisWorkbook.m_ResetWorkbookAndCreateMainPage("rt_Snapshots:restore-runtime-state:fallback-main-reset", False) Then
        If Err.Number <> 0 Then
            errDescription = Err.Description
            Err.Clear
            On Error GoTo 0
            ex_Core.m_LogError "snapshots:restore-runtime-state fallback-main-reset failed reason='" & VBA.Replace$(reasonText, "'", "''") & "' err='" & VBA.Replace$(errDescription, "'", "''") & "'"
            Exit Function
        End If
        On Error GoTo 0
        ex_Core.m_LogError "snapshots:restore-runtime-state fallback-main-reset failed reason='" & VBA.Replace$(reasonText, "'", "''") & "' err='returned-false'"
        Exit Function
    End If
    If Err.Number <> 0 Then
        errDescription = Err.Description
        Err.Clear
        On Error GoTo 0
        ex_Core.m_LogError "snapshots:restore-runtime-state fallback-main-reset failed reason='" & VBA.Replace$(reasonText, "'", "''") & "' err='" & VBA.Replace$(errDescription, "'", "''") & "'"
        Exit Function
    End If
    On Error GoTo 0

    outRestoredPagesCount = 1
    savePagesOk = m_SavePageSnapshots()
    saveRuntimeOk = m_SaveRuntimeGlobalsSnapshot()

    ' Даже если checkpoint сохранить не удалось, UI уже поднят и клики не должны "умирать".
    ' Ошибку логируем, чтобы можно было добить первопричину по core.log.
    If Not savePagesOk Or Not saveRuntimeOk Then
        ex_Core.m_LogError "snapshots:restore-runtime-state fallback-checkpoint-failed reason='" & VBA.Replace$(reasonText, "'", "''") & "' savePages=" & VBA.CStr(savePagesOk) & " saveRuntime=" & VBA.CStr(saveRuntimeOk)
    End If

    ex_Core.m_LogInfo "snapshots:restore-runtime-state fallback-main-reset done reason='" & VBA.Replace$(reasonText, "'", "''") & "' restoredPages=" & VBA.CStr(outRestoredPagesCount)
    private_TryFallbackRestoreByResettingMainPage = True
End Function


' Callstack[1]: ThisWorkbook.Workbook_Open -> rt_Snapshots.m_RestorePageSnapshots
' Callstack[2]: rt_CoreActions.m_RerenderLastPageAfterUpdate -> rt_Snapshots.m_RestorePageSnapshots
Public Function m_RestorePageSnapshots( _
    Optional ByVal renderRestored As Boolean = True, _
    Optional ByVal reasonPrefix As String = VBA.vbNullString, _
    Optional ByRef outRestoredCount As Long = 0 _
) As Boolean
    Dim snapshots As Collection
    Dim snapshotItem As Variant
    Dim snapshotXml As String
    Dim snapshotParser As obj_PageBase
    Dim codeNameValue As String
    Dim sheetName As String
    Dim typeRoot As String
    Dim uiPath As String
    Dim payloadXml As String
    Dim pageIdFromSnapshot As String
    Dim pageTypeValue As Long
    Dim createdPageId As String
    Dim page As obj_IPage
    Dim serializablePage As obj_ISerializable
    Dim pageType As PageTypeEnum
    Dim hasSnapshots As Boolean
    Dim tmpWs As Worksheet
    Dim pageBase As obj_PageBase

    outRestoredCount = 0
    If Not private_TryLoadPageSnapshots(snapshots) Then Exit Function
    If snapshots Is Nothing Then
        m_RestorePageSnapshots = True
        Exit Function
    End If

    For Each snapshotItem In snapshots
        If VBA.Len(VBA.Trim$(VBA.CStr(snapshotItem))) > 0 Then
            hasSnapshots = True
            Exit For
        End If
    Next snapshotItem

    If Not hasSnapshots Then
        m_RestorePageSnapshots = True
        Exit Function
    End If

    If Not private_TryResetWorkbookBeforeRestore(tmpWs) Then Exit Function

    Set snapshotParser = New obj_PageBase

    For Each snapshotItem In snapshots
        snapshotXml = VBA.Trim$(VBA.CStr(snapshotItem))
        If VBA.Len(snapshotXml) = 0 Then GoTo ContinueSnapshot

        codeNameValue = VBA.vbNullString
        sheetName = VBA.vbNullString
        typeRoot = VBA.vbNullString
        uiPath = VBA.vbNullString
        payloadXml = VBA.vbNullString
        pageIdFromSnapshot = VBA.vbNullString
        pageTypeValue = 0
        createdPageId = VBA.vbNullString

        If Not snapshotParser.TryDeserializePageSnapshotEnvelope( _
            snapshotXml:=snapshotXml, _
            outCodeName:=codeNameValue, _
            outSheetName:=sheetName, _
            outTypeRoot:=typeRoot, _
            outUiPath:=uiPath, _
            outPagePayloadXml:=payloadXml, _
            outPageId:=pageIdFromSnapshot, _
            outPageType:=pageTypeValue) Then Exit Function

        pageType = private_NormalizePageType(pageTypeValue, typeRoot)
        If Not rt_PageManager.m_CreatePage(uiPath, pageType, createdPageId, sheetName) Then GoTo ContinueSnapshot
        If Not rt_PageManager.m_TryGetPageById(createdPageId, page) Then GoTo ContinueSnapshot

        If VBA.Len(payloadXml) > 0 Then
            If Not private_TryCastSerializablePage(page, serializablePage) Then GoTo ContinueSnapshot
            If Not serializablePage.TryDeserializeSnapshot(payloadXml) Then GoTo ContinueSnapshot
        End If

        If renderRestored Then
            If Not rt_PageManager.m_RenderPage(page, reasonPrefix & "|restore") Then GoTo ContinueSnapshot
        End If

        outRestoredCount = outRestoredCount + 1

ContinueSnapshot:
    Next snapshotItem

    If Not tmpWs Is Nothing Then
        On Error Resume Next
        Set pageBase = Nothing
        If outRestoredCount > 0 Then
            If rt_PageManager.m_TryGetPageById(createdPageId, page) Then
                Set pageBase = page.GetPageBase()
            End If
        End If
        If pageBase Is Nothing Then
            Application.DisplayAlerts = False
            tmpWs.Delete
            Application.DisplayAlerts = True
        ElseIf Not (pageBase.Worksheet Is tmpWs) Then
            Application.DisplayAlerts = False
            tmpWs.Delete
            Application.DisplayAlerts = True
        End If
        On Error GoTo 0
    End If

    m_RestorePageSnapshots = True
End Function
' //
' // Internal
' //
' Callstack[1]: rt_Snapshots.m_TryRestoreRuntimeStateFromSnapshots -> private_HasRuntimePageForActiveWorksheet
Private Function private_HasRuntimePageForActiveWorksheet() As Boolean
    Dim ws As Worksheet
    Dim page As obj_IPage

    On Error Resume Next
    If TypeOf Application.ActiveSheet Is Worksheet Then
        Set ws = Application.ActiveSheet
    End If
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    private_HasRuntimePageForActiveWorksheet = rt_PageManager.m_TryGetPageByWorksheet(ws, page)
End Function

' Callstack[1]: rt_Snapshots.m_SaveRuntimeGlobalsSnapshot -> private_TryAppendModuleSnapshot
Private Function private_TryAppendModuleSnapshot(ByVal rootNode As Object, ByVal moduleName As String) As Boolean
    Dim snapshotXml As String
    Dim moduleNode As Object
    Dim snapshotNode As Object
    Dim dom As Object

    If rootNode Is Nothing Then Exit Function
    moduleName = VBA.Trim$(moduleName)
    If VBA.Len(moduleName) = 0 Then Exit Function

    snapshotXml = VBA.vbNullString
    If Not private_TrySerializeRuntimeModuleSnapshot(moduleName, snapshotXml) Then Exit Function

    Set dom = rootNode.OwnerDocument
    If dom Is Nothing Then
        VBA.MsgBox "Snapshots: runtime globals DOM owner is not available.", VBA.vbExclamation
        Exit Function
    End If

    Set moduleNode = dom.createNode(1, RUNTIME_GLOBALS_MODULE_NODE, RUNTIME_GLOBALS_NS)
    moduleNode.setAttribute RUNTIME_GLOBALS_MODULE_NAME_ATTR, moduleName

    Set snapshotNode = dom.createElement(RUNTIME_GLOBALS_MODULE_SNAPSHOT_NODE)
    snapshotNode.Text = VBA.CStr(snapshotXml)
    moduleNode.appendChild snapshotNode

    rootNode.appendChild moduleNode
    private_TryAppendModuleSnapshot = True
End Function


' Callstack[1]: rt_Snapshots.private_TryAppendModuleSnapshot -> private_TrySerializeRuntimeModuleSnapshot
Private Function private_TrySerializeRuntimeModuleSnapshot(ByVal moduleName As String, ByRef outSnapshotXml As String) As Boolean
    outSnapshotXml = VBA.vbNullString
    moduleName = VBA.LCase$(VBA.Trim$(moduleName))

    Select Case moduleName
        Case VBA.LCase$(MODULE_NAME_PAGE_MANAGER)
            private_TrySerializeRuntimeModuleSnapshot = rt_PageManager.m_TrySerializeModuleSnapshot(outSnapshotXml)

        Case Else
            ex_Core.m_LogInfo "runtime-globals: serialize skipped for unknown module '" & moduleName & "'"
            private_TrySerializeRuntimeModuleSnapshot = True
    End Select
End Function


' Callstack[1]: rt_Snapshots.m_RestoreRuntimeGlobalsSnapshot -> private_TryDeserializeRuntimeModuleSnapshot
Private Function private_TryDeserializeRuntimeModuleSnapshot(ByVal moduleName As String, ByVal snapshotXml As String) As Boolean
    moduleName = VBA.LCase$(VBA.Trim$(moduleName))

    Select Case moduleName
        Case VBA.LCase$(MODULE_NAME_PAGE_MANAGER)
            private_TryDeserializeRuntimeModuleSnapshot = rt_PageManager.m_TryDeserializeModuleSnapshot(snapshotXml)

        Case Else
            ex_Core.m_LogInfo "runtime-globals: deserialize skipped for unknown module '" & moduleName & "'"
            private_TryDeserializeRuntimeModuleSnapshot = True
    End Select
End Function


Private Function private_TryCollectAllPages(ByRef outPages As Collection) As Boolean
    Dim mainPages As Collection
    Dim generatedPages As Collection
    Dim pageItem As Variant

    Set outPages = New Collection

    If Not rt_PageManager.m_TryGetPagesByType(PageTypeMain, mainPages) Then Exit Function
    If Not mainPages Is Nothing Then
        For Each pageItem In mainPages
            outPages.Add pageItem
        Next pageItem
    End If

    If Not rt_PageManager.m_TryGetPagesByType(PageTypeGenerated, generatedPages) Then Exit Function
    If Not generatedPages Is Nothing Then
        For Each pageItem In generatedPages
            outPages.Add pageItem
        Next pageItem
    End If

    private_TryCollectAllPages = True
End Function


Private Function private_TryCastSerializablePage(ByVal page As obj_IPage, ByRef outSerializable As obj_ISerializable) As Boolean
    Set outSerializable = Nothing
    If page Is Nothing Then Exit Function

    On Error Resume Next
    Set outSerializable = page
    If Err.Number <> 0 Then
        Err.Clear
        Set outSerializable = Nothing
    End If
    On Error GoTo 0

    private_TryCastSerializablePage = Not outSerializable Is Nothing
End Function


Private Function private_NormalizePageType(ByVal pageTypeValue As Long, ByVal typeRoot As String) As PageTypeEnum
    Select Case VBA.CLng(pageTypeValue)
        Case PageTypeGenerated
            private_NormalizePageType = PageTypeGenerated
            Exit Function
    End Select

    typeRoot = VBA.LCase$(VBA.Trim$(typeRoot))
    If typeRoot = "page.generated" Then
        private_NormalizePageType = PageTypeGenerated
    Else
        private_NormalizePageType = PageTypeMain
    End If
End Function


Private Function private_TryResetWorkbookBeforeRestore(ByRef outTemporaryWorksheet As Worksheet) As Boolean
    Dim wb As Workbook
    Dim tmpName As String

    Set outTemporaryWorksheet = Nothing
    Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    rt_PageManager.m_DisposeAllPages

    On Error GoTo EH_RESET
    Application.DisplayAlerts = False
    Do While wb.Worksheets.Count > 1
        wb.Worksheets(1).Delete
    Loop
    Set outTemporaryWorksheet = wb.Worksheets(1)
    Application.DisplayAlerts = True

    tmpName = "__restore_tmp__"
    On Error Resume Next
    outTemporaryWorksheet.Name = tmpName
    Err.Clear
    On Error GoTo 0

    private_TryResetWorkbookBeforeRestore = True
    Exit Function

EH_RESET:
    Application.DisplayAlerts = True
    VBA.MsgBox "Snapshots: failed to reset workbook before restore: " & Err.Description, VBA.vbExclamation
End Function


Private Function private_TryLoadPageSnapshots(ByRef outPages As Collection) As Boolean
    private_TryLoadPageSnapshots = private_TryLoadSnapshotXmlCollection(PAGE_SNAPSHOT_NS, PAGE_SNAPSHOT_NODE, outPages)
End Function


Private Function private_TrySaveSnapshotXmlCollection( _
    ByVal namespaceUri As String, _
    ByVal rootName As String, _
    ByVal nodeName As String, _
    ByVal snapshots As Collection _
) As Boolean
    Dim partObj As Object
    Dim dom As Object
    Dim rootNode As Object
    Dim item As Variant
    Dim snapshotNode As Object
    Dim payloadNode As Object
    Dim snapshotXml As String

    If Not ex_CustomXmlPartStore.m_TryCreateEmptyDom(rootName, namespaceUri, dom) Then Exit Function

    Set rootNode = dom.DocumentElement
    If rootNode Is Nothing Then
        VBA.MsgBox "Snapshots: root node is missing for namespace '" & namespaceUri & "'.", VBA.vbExclamation
        Exit Function
    End If

    If Not snapshots Is Nothing Then
        For Each item In snapshots
            snapshotXml = VBA.Trim$(VBA.CStr(item))
            If VBA.Len(snapshotXml) = 0 Then GoTo ContinueSnapshot

            Set snapshotNode = dom.createNode(1, nodeName, namespaceUri)
            Set payloadNode = dom.createElement(SNAPSHOT_PAYLOAD_NODE)
            payloadNode.Text = snapshotXml
            snapshotNode.appendChild payloadNode

            rootNode.appendChild snapshotNode

ContinueSnapshot:
        Next item
    End If

    If Not ex_CustomXmlPartStore.m_TryFindPartByNamespace(namespaceUri, partObj) Then Exit Function
    If Not ex_CustomXmlPartStore.m_TrySaveDom(dom, partObj) Then Exit Function

    private_TrySaveSnapshotXmlCollection = True
End Function


Private Function private_TryLoadSnapshotXmlCollection( _
    ByVal namespaceUri As String, _
    ByVal nodeName As String, _
    ByRef outSnapshots As Collection _
) As Boolean
    Dim partObj As Object
    Dim dom As Object
    Dim rootNode As Object
    Dim child As Object
    Dim payloadNode As Object
    Dim snapshotXml As String

    Set outSnapshots = New Collection

    If Not ex_CustomXmlPartStore.m_TryFindPartByNamespace(namespaceUri, partObj) Then Exit Function
    If partObj Is Nothing Then
        private_TryLoadSnapshotXmlCollection = True
        Exit Function
    End If

    If Not ex_CustomXmlPartStore.m_TryLoadPartDom(partObj, dom) Then Exit Function

    Set rootNode = dom.DocumentElement
    If rootNode Is Nothing Then
        private_TryLoadSnapshotXmlCollection = True
        Exit Function
    End If

    For Each child In rootNode.ChildNodes
        If child.NodeType <> 1 Then GoTo ContinueNode
        If VBA.LCase$(VBA.CStr(child.baseName)) <> VBA.LCase$(nodeName) Then GoTo ContinueNode

        snapshotXml = VBA.vbNullString
        Set payloadNode = child.selectSingleNode("*[local-name()='" & SNAPSHOT_PAYLOAD_NODE & "']")
        If Not payloadNode Is Nothing Then
            snapshotXml = VBA.Trim$(VBA.CStr(payloadNode.Text))
        End If
        If VBA.Len(snapshotXml) = 0 Then GoTo ContinueNode

        outSnapshots.Add snapshotXml

ContinueNode:
    Next child

    private_TryLoadSnapshotXmlCollection = True
End Function
