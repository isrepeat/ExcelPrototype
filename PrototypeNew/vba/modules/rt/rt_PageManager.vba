Attribute VB_Name = "rt_PageManager"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private g_PageById As Object
Private g_LastRenderedPageId As String
Private g_PageIdSeed As Long

Private Const MODULE_SNAPSHOT_ROOT As String = "pageManagerState"
Private Const MODULE_SNAPSHOT_NS As String = "urn:excelprototype:runtime-module:page-manager:v1"
Private Const MODULE_SNAPSHOT_PAGE_NODE As String = "page"
Private Const MODULE_SNAPSHOT_PAYLOAD_NODE As String = "payload"

Public Sub m_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:rt_PageManager.m_Module_Dispose"
#End If
    On Error Resume Next
    m_DisposeAllPages
    Err.Clear
    Set g_PageById = Nothing
    g_LastRenderedPageId = VBA.vbNullString
    On Error GoTo 0
End Sub

' //
' // API
' //
' Callstack[1]: rt_RestoreManager.m_SaveRuntimeGlobalsSnapshot -> private_TryAppendModuleSnapshot -> private_TrySerializeRuntimeModuleSnapshot -> rt_PageManager.m_TrySerializeModuleSnapshot
Public Function m_TrySerializeModuleSnapshot(ByRef outSnapshotXml As String) As Boolean
    Dim dom As Object
    Dim rootNode As Object
    Dim pageNode As Object
    Dim payloadNode As Object
    Dim pageKey As Variant
    Dim page As obj_IPage
    Dim pageBase As obj_PageBase
    Dim serializablePage As obj_ISerializable
    Dim typeRoot As String
    Dim payloadXml As String
    Dim pageId As String
    Dim worksheetName As String

    outSnapshotXml = VBA.vbNullString

    If Not ex_Core.m_CustomXmlPartStore_TryCreateEmptyDom(MODULE_SNAPSHOT_ROOT, MODULE_SNAPSHOT_NS, dom) Then Exit Function
    Set rootNode = dom.DocumentElement
    If rootNode Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PageManager: module snapshot root node is missing."
#End If
        Exit Function
    End If

    ' module-level metadata:
    ' 1) lastRenderedSheetName — чтобы вернуть фокус на "последнюю" страницу;
    ' 2) pageIdSeed — чтобы после restore новые id продолжали последовательность.
    worksheetName = VBA.vbNullString
    If Not m_TryGetLastRenderedWorksheetName(worksheetName) Then Exit Function
    rootNode.setAttribute "lastRenderedSheetName", worksheetName
    rootNode.setAttribute "pageIdSeed", VBA.CStr(g_PageIdSeed)

    ' Snapshot каждой страницы = transport envelope (id/type/sheet/uiPath)
    ' + page payload из obj_ISerializable.TrySerializeSnapshot.
    private_EnsureStorage
    For Each pageKey In g_PageById.Keys
        Set page = Nothing
        Set pageBase = Nothing
        Set serializablePage = Nothing
        Set page = g_PageById(VBA.CStr(pageKey))
        If page Is Nothing Then GoTo ContinuePage

        Set pageBase = page.GetPageBase()
        If pageBase Is Nothing Then GoTo ContinuePage
        If pageBase.Worksheet Is Nothing Then GoTo ContinuePage

        If Not private_TryCastSerializablePage(page, serializablePage) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.m_Diagnostic_LogError "PageManager: page class '" & VBA.TypeName(page) & "' must implement obj_ISerializable."
#End If
            Exit Function
        End If

        typeRoot = VBA.LCase$(VBA.Trim$(serializablePage.GetSerializableTypeRoot()))
        If VBA.Len(typeRoot) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.m_Diagnostic_LogError "PageManager: serializable type root is empty for page class '" & VBA.TypeName(page) & "'."
#End If
            Exit Function
        End If

        payloadXml = VBA.vbNullString
        If Not serializablePage.TrySerializeSnapshot(payloadXml) Then Exit Function

        pageId = VBA.LCase$(VBA.Trim$(pageBase.PageId))
        If VBA.Len(pageId) = 0 Then pageId = VBA.LCase$(VBA.Trim$(VBA.CStr(pageKey)))

        Set pageNode = dom.createElement(MODULE_SNAPSHOT_PAGE_NODE)
        pageNode.setAttribute "pageId", pageId
        pageNode.setAttribute "sheetName", VBA.Trim$(VBA.CStr(pageBase.Worksheet.Name))
        pageNode.setAttribute "codeName", VBA.Trim$(VBA.CStr(pageBase.Worksheet.CodeName))
        pageNode.setAttribute "type", typeRoot
        pageNode.setAttribute "uiPath", pageBase.UiPath

        Set payloadNode = dom.createElement(MODULE_SNAPSHOT_PAYLOAD_NODE)
        payloadNode.Text = VBA.CStr(payloadXml)
        pageNode.appendChild payloadNode
        rootNode.appendChild pageNode

ContinuePage:
    Next pageKey

    outSnapshotXml = VBA.CStr(dom.XML)
    m_TrySerializeModuleSnapshot = (VBA.Len(VBA.Trim$(outSnapshotXml)) > 0)
End Function

' Callstack[1]: rt_RestoreManager.m_RestoreRuntimeGlobalsSnapshot -> private_TryDeserializeRuntimeModuleSnapshot -> rt_PageManager.m_TryDeserializeModuleSnapshot
Public Function m_TryDeserializeModuleSnapshot(ByVal snapshotXml As String) As Boolean
    Dim dom As Object
    Dim rootNode As Object
    Dim pageNodes As Object
    Dim pageNode As Object
    Dim payloadNode As Object
    Dim page As obj_IPage
    Dim serializablePage As obj_ISerializable
    Dim restoredPages As Collection
    Dim tmpWs As Worksheet
    Dim pageId As String
    Dim sheetName As String
    Dim codeName As String
    Dim typeRoot As String
    Dim uiPath As String
    Dim payloadXml As String
    Dim isPageCreated As Boolean
    Dim isSnapshotSucceeded As Boolean
    Dim pageIdSeedText As String
    Dim pageIdSeedValue As Double
    Dim worksheetName As String
    Dim isRestorePrepared As Boolean
    Dim finalizeOk As Boolean

    snapshotXml = VBA.Trim$(snapshotXml)
    If VBA.Len(snapshotXml) = 0 Then
        m_TryDeserializeModuleSnapshot = True
        Exit Function
    End If

    If Not ex_Core.m_CustomXmlPartStore_TryLoadDomFromXml(snapshotXml, dom) Then Exit Function
    Set rootNode = dom.DocumentElement
    If rootNode Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PageManager: module snapshot root node is missing."
#End If
        Exit Function
    End If

    If VBA.StrComp(VBA.LCase$(VBA.CStr(rootNode.baseName)), MODULE_SNAPSHOT_ROOT, VBA.vbTextCompare) <> 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PageManager: unexpected module snapshot root '" & VBA.CStr(rootNode.baseName) & "'."
#End If
        Exit Function
    End If

    ' Подготовка workbook к восстановлению:
    ' очищаем старые runtime-страницы и оставляем временный лист-заглушку.
    If Not rt_RestoreManager.m_TryPrepareWorkbookForRestore(tmpWs) Then Exit Function
    isRestorePrepared = True
    Set restoredPages = New Collection

    ' Фаза 1: recreate всех страниц и загрузка payload.
    ' На этом шаге восстанавливаем объекты/данные, но не межобъектные связи.
    Set pageNodes = rootNode.selectNodes("*[local-name()='" & MODULE_SNAPSHOT_PAGE_NODE & "']")
    If Not pageNodes Is Nothing Then
        For Each pageNode In pageNodes
            Set page = Nothing
            Set serializablePage = Nothing
            isPageCreated = False
            isSnapshotSucceeded = False

            pageId = VBA.LCase$(VBA.Trim$(VBA.CStr(pageNode.getAttribute("pageId"))))
            sheetName = VBA.Trim$(VBA.CStr(pageNode.getAttribute("sheetName")))
            codeName = VBA.Trim$(VBA.CStr(pageNode.getAttribute("codeName")))
            typeRoot = VBA.LCase$(VBA.Trim$(VBA.CStr(pageNode.getAttribute("type"))))
            uiPath = VBA.Trim$(VBA.CStr(pageNode.getAttribute("uiPath")))
            If VBA.Len(sheetName) = 0 Then sheetName = codeName

            Set payloadNode = pageNode.selectSingleNode("*[local-name()='" & MODULE_SNAPSHOT_PAYLOAD_NODE & "']")
            payloadXml = VBA.vbNullString
            If Not payloadNode Is Nothing Then payloadXml = VBA.CStr(payloadNode.Text)

            If VBA.Len(typeRoot) = 0 Then GoTo ContinuePage
            If VBA.Len(pageId) = 0 Then GoTo ContinuePage

            If Not ex_SerializableFactory.m_TryCreatePageByTypeRoot(typeRoot, page) Then GoTo ContinuePage
            If page Is Nothing Then GoTo ContinuePage

            If Not m_RestorePage(page, uiPath, sheetName, pageId) Then GoTo ContinuePage
            isPageCreated = True

            If VBA.Len(payloadXml) > 0 Then
                If Not private_TryCastSerializablePage(page, serializablePage) Then GoTo ContinuePage
                If Not serializablePage.TryDeserializeSnapshot(payloadXml) Then GoTo ContinuePage
            End If

            restoredPages.Add page
            isSnapshotSucceeded = True

ContinuePage:
            If Not isSnapshotSucceeded Then
                ' Если страница частично восстановилась и потом упала,
                ' убираем ее сразу, чтобы не оставлять "битый" runtime state.
                On Error Resume Next
                If Not page Is Nothing And isPageCreated Then
                    Call m_RemovePage(page, True)
                End If
                On Error GoTo 0
            End If
        Next pageNode
    End If

    ' Фаза 2: достройка связей/внутреннего состояния через TryRestoreState.
    If Not rt_RestoreManager.m_TryRestoreSerializableCollectionState(restoredPages, "rt_PageManager.pages") Then GoTo EH_FAIL
    ' Фаза 3: рендер после того, как все страницы уже существуют в коллекции.
    If Not private_TryRenderPagesCollection(restoredPages, "page-manager:restore") Then GoTo EH_FAIL

    ' Возвращаем логическую "последнюю страницу" и id-seed генератора.
    worksheetName = VBA.Trim$(VBA.CStr(rootNode.getAttribute("lastRenderedSheetName")))
    If Not m_TryRestoreLastRenderedWorksheetName(worksheetName) Then GoTo EH_FAIL

    pageIdSeedText = VBA.Trim$(VBA.CStr(rootNode.getAttribute("pageIdSeed")))
    If VBA.Len(pageIdSeedText) > 0 Then
        On Error Resume Next
        pageIdSeedValue = VBA.CDbl(pageIdSeedText)
        If Err.Number <> 0 Then
            Err.Clear
            pageIdSeedValue = 0
        End If
        On Error GoTo 0
        If pageIdSeedValue > 0 Then g_PageIdSeed = VBA.CLng(pageIdSeedValue)
    End If

    ' Финализация: удаляем временный лист-заглушку, если restore дошел до конца.
    finalizeOk = rt_RestoreManager.m_TryFinalizeWorkbookAfterRestore(tmpWs)
    If Not finalizeOk Then GoTo EH_FAIL

    m_TryDeserializeModuleSnapshot = True
    Exit Function

EH_FAIL:
    On Error Resume Next
    If isRestorePrepared Then
        Call rt_RestoreManager.m_TryFinalizeWorkbookAfterRestore(tmpWs)
    End If
    On Error GoTo 0
End Function

' Callstack[1]: ThisWorkbook.Workbook_Open -> ThisWorkbook.m_ResetWorkbookAndCreateMainPage -> private_ResetWorkbookAndCreateMainPage -> rt_PageManager.m_CreatePage
' Callstack[2]: ex_Core.private_TryRecoverUiAfterUpdate -> ThisWorkbook.m_ResetWorkbookAndCreateMainPage -> private_ResetWorkbookAndCreateMainPage -> rt_PageManager.m_CreatePage
' Callstack[3]: rt_RestoreManager.m_RestorePageSnapshots -> rt_PageManager.m_RestorePage
Public Function m_CreatePage( _
    ByVal page As obj_IPage, _
    ByVal uiPath As String, _
    ByVal sheetName As String, _
    Optional ByVal createContext As Object = Nothing _
) As Boolean
    Dim pageId As String

    pageId = private_BuildPageId(private_ResolvePageIdPrefix(page))
    If VBA.Len(VBA.Trim$(pageId)) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PageManager: failed to generate page id."
#End If
        Exit Function
    End If

    m_CreatePage = private_CreatePageInternal(page, uiPath, sheetName, pageId, createContext)
End Function

Public Function m_RestorePage( _
    ByVal page As obj_IPage, _
    ByVal uiPath As String, _
    ByVal sheetName As String, _
    ByVal pageId As String, _
    Optional ByVal restoreContext As Object = Nothing _
) As Boolean
    pageId = VBA.LCase$(VBA.Trim$(pageId))
    If VBA.Len(pageId) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PageManager: restore page id is empty."
#End If
        Exit Function
    End If

    m_RestorePage = private_CreatePageInternal(page, uiPath, sheetName, pageId, restoreContext)
End Function

' Callstack[1]: rt_PageManager.m_RenderPageById -> rt_PageManager.m_TryGetPageById
' Callstack[2]: rt_PageManager.m_RemovePageById -> rt_PageManager.m_TryGetPageById
' Callstack[3]: ThisWorkbook.m_ResetWorkbookAndCreateMainPage -> private_ResetWorkbookAndCreateMainPage -> rt_PageManager.m_TryGetPageById
' Callstack[4]: rt_RestoreManager.m_RestorePageSnapshots -> rt_PageManager.m_TryGetPageById
Public Function m_TryGetPageById(ByVal pageId As String, ByRef outPage As obj_IPage) As Boolean
    Set outPage = Nothing
    pageId = VBA.LCase$(VBA.Trim$(pageId))
    If VBA.Len(pageId) = 0 Then Exit Function

    private_EnsureStorage
    If Not g_PageById.Exists(pageId) Then Exit Function

    Set outPage = g_PageById(pageId)
    If outPage Is Nothing Then Exit Function

    m_TryGetPageById = True
End Function

' Callstack[1]: Shape.OnAction -> rt_Bridge.m_OnShapeClick -> rt_PageManager.m_TryGetPageByWorksheet
' Callstack[2]: ex_Test.private_TryResolvePageBase -> ex_HelpersSheet.m_TryGetActivePageBase -> ex_HelpersSheet.m_TryGetPageBaseByWorksheet -> rt_PageManager.m_TryGetPageByWorksheet
' Callstack[3]: ex_Test.m_TEST_UpdateCurrentPage -> ex_HelpersSheet.m_TryRerenderActivePage -> rt_PageManager.m_TryGetPageByWorksheet
' Callstack[4]: ex_Test.m_TEST_SetDemoConfigVariantA -> ex_HelpersSheet.m_TryRerenderActivePage -> rt_PageManager.m_TryGetPageByWorksheet
' Callstack[5]: ex_Test.m_TEST_SetDemoConfigVariantB -> ex_HelpersSheet.m_TryRerenderActivePage -> rt_PageManager.m_TryGetPageByWorksheet
' Callstack[6]: ThisWorkbook.Workbook_SheetBeforeDelete -> ex_HelpersSheet.m_RemovePageByWorksheet -> rt_PageManager.m_TryGetPageByWorksheet
' Callstack[7]: ex_ControlRefreshRuntime.m_TryRefreshStaticControl -> rt_PageManager.m_TryGetPageByWorksheet
' Callstack[8]: obj_PageMain.private_TryRerenderByDataChange -> rt_PageManager.m_TryGetPageByWorksheet
Public Function m_TryGetPageByWorksheet(ByVal ws As Worksheet, ByRef outPage As obj_IPage) As Boolean
    Dim wsName As String
    Dim resolvedPageId As String

    Set outPage = Nothing
    If ws Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "page-manager:get-by-worksheet input-invalid worksheet is not specified"
#End If
        Exit Function
    End If

    On Error Resume Next
    wsName = VBA.Trim$(VBA.CStr(ws.Name))
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "page-manager:get-by-worksheet worksheet-name-unavailable"
#End If
        Exit Function
    End If
    On Error GoTo 0

    wsName = VBA.Replace$(wsName, "'", "''")
    If Not private_TryFindPageByWorksheet(ws, outPage, resolvedPageId) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "page-manager:get-by-worksheet page-not-found sheet='" & wsName & "'"
#End If
        Exit Function
    End If

    m_TryGetPageByWorksheet = True
End Function

' Callstack[1]: ex_HelpersSheet.m_TryGetPageBaseByWorksheetName -> rt_PageManager.m_TryGetPageByWorksheetName
Public Function m_TryGetPageByWorksheetName(ByVal worksheetName As String, ByRef outPage As obj_IPage) As Boolean
    Dim ws As Worksheet

    Set outPage = Nothing
    worksheetName = VBA.Trim$(worksheetName)
    If VBA.Len(worksheetName) = 0 Then Exit Function

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(worksheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    m_TryGetPageByWorksheetName = m_TryGetPageByWorksheet(ws, outPage)
End Function

' Callstack[1]: rt_RestoreManager.private_TryCollectAllPages -> rt_PageManager.m_TryGetAllPages
Public Function m_TryGetAllPages(ByRef outPages As Collection) As Boolean
    Dim pageId As Variant
    Dim page As obj_IPage

    Set outPages = New Collection
    private_EnsureStorage

    For Each pageId In g_PageById.Keys
        Set page = g_PageById(pageId)
        If page Is Nothing Then GoTo ContinueLoop

        outPages.Add page

ContinueLoop:
    Next pageId

    m_TryGetAllPages = True
End Function

Public Function m_TryGetPagesCount(ByRef outCount As Long) As Boolean
    Dim pageId As Variant

    outCount = 0
    private_EnsureStorage

    For Each pageId In g_PageById.Keys
        If Not (g_PageById(pageId) Is Nothing) Then outCount = outCount + 1
    Next pageId

    m_TryGetPagesCount = True
End Function

' Callstack[1]: ThisWorkbook.m_ResetWorkbookAndCreateMainPage -> private_ResetWorkbookAndCreateMainPage -> rt_PageManager.m_RenderPageById
' Callstack[2]: ex_Core.private_TryRecoverUiAfterUpdate -> ThisWorkbook.m_ResetWorkbookAndCreateMainPage -> private_ResetWorkbookAndCreateMainPage -> rt_PageManager.m_RenderPageById
Public Function m_RenderPageById(ByVal pageId As String, Optional ByVal reason As String = VBA.vbNullString) As Boolean
    Dim page As obj_IPage

    If Not m_TryGetPageById(pageId, page) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "page-manager:render-by-id input-invalid page is not found"
#End If
        Exit Function
    End If

    m_RenderPageById = m_RenderPage(page, reason)
End Function

' Callstack[1]: rt_PageManager.m_RenderPageById -> rt_PageManager.m_RenderPage
' Callstack[2]: ex_Test.m_TEST_UpdateCurrentPage -> ex_HelpersSheet.m_TryRerenderActivePage -> rt_PageManager.m_RenderPage
' Callstack[3]: ex_Test.m_TEST_SetDemoConfigVariantA -> ex_HelpersSheet.m_TryRerenderActivePage -> rt_PageManager.m_RenderPage
' Callstack[4]: ex_Test.m_TEST_SetDemoConfigVariantB -> ex_HelpersSheet.m_TryRerenderActivePage -> rt_PageManager.m_RenderPage
' Callstack[5]: ex_Test.private_TrySetItemsSource -> ex_Test.private_TryRerenderPage -> rt_PageManager.m_RenderPage
' Callstack[6]: ex_Test.private_TrySetObjectSource -> ex_Test.private_TryRerenderPage -> rt_PageManager.m_RenderPage
' Callstack[7]: ex_Test.private_TryRemoveObjectSource -> ex_Test.private_TryRerenderPage -> rt_PageManager.m_RenderPage
' Callstack[8]: obj_PageMain.private_TryRerenderByDataChange -> rt_PageManager.m_RenderPage
' Callstack[9]: rt_RestoreManager.m_RestorePageSnapshots(renderRestored:=True) -> rt_PageManager.m_RenderPage
Public Function m_RenderPage(ByVal page As obj_IPage, Optional ByVal reason As String = VBA.vbNullString) As Boolean
    Dim pageBase As obj_PageBase
    Dim sheetName As String
    Dim normalizedReason As String
    Dim errDescription As String
    Dim pageId As String

    If page Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "page-manager:render input-invalid page is not specified"
#End If
        Exit Function
    End If

    Set pageBase = page.GetPageBase()
    If pageBase Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "page-manager:render input-invalid page base is not specified"
#End If
        Exit Function
    End If

    If pageBase.Worksheet Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "page-manager:render input-invalid worksheet is not specified"
#End If
        Exit Function
    End If

    sheetName = VBA.Replace$(VBA.CStr(pageBase.Worksheet.Name), "'", "''")
    normalizedReason = VBA.Trim$(VBA.CStr(reason))
    If VBA.Len(normalizedReason) = 0 Then normalizedReason = "manual"

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "page-manager:render-start sheet='" & sheetName & "' reason='" & VBA.Replace$(normalizedReason, "'", "''") & "'"
#End If

    On Error GoTo EH_RENDER
    m_RenderPage = page.Render()

    If m_RenderPage Then
        pageId = VBA.LCase$(VBA.Trim$(pageBase.PageId))
        If VBA.Len(pageId) = 0 Then pageId = private_TryResolvePageIdByObject(page)
        g_LastRenderedPageId = pageId
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogInfo "page-manager:render-done sheet='" & sheetName & "'"
#End If
    Else
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "page-manager:render-failed sheet='" & sheetName & "'"
#End If
    End If
    Exit Function

EH_RENDER:
    errDescription = Err.Description
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.m_Diagnostic_LogError "page-manager:render-exception sheet='" & sheetName & "' err='" & VBA.Replace$(errDescription, "'", "''") & "'"
#End If
End Function

' Callstack[1]: VBA.ImmediateWindow -> rt_PageManager.m_RemovePageById
Public Function m_RemovePageById(ByVal pageId As String, Optional ByVal deleteWorksheet As Boolean = False) As Boolean
    Dim page As obj_IPage

    If Not m_TryGetPageById(pageId, page) Then
        m_RemovePageById = True
        Exit Function
    End If

    m_RemovePageById = m_RemovePage(page, deleteWorksheet)
End Function

' Callstack[1]: ThisWorkbook.Workbook_SheetBeforeDelete -> ex_HelpersSheet.m_RemovePageByWorksheet -> rt_PageManager.m_RemovePage
' Callstack[2]: rt_PageManager.m_RemovePageById -> rt_PageManager.m_RemovePage
Public Function m_RemovePage(ByVal page As obj_IPage, Optional ByVal deleteWorksheet As Boolean = False) As Boolean
    Dim pageBase As obj_PageBase
    Dim pageId As String

    If page Is Nothing Then
        m_RemovePage = True
        Exit Function
    End If

    Set pageBase = page.GetPageBase()
    If pageBase Is Nothing Then Exit Function

    pageId = VBA.LCase$(VBA.Trim$(pageBase.PageId))
    If VBA.Len(pageId) = 0 Then pageId = private_TryResolvePageIdByObject(page)

    private_EnsureStorage
    If VBA.Len(pageId) > 0 Then
        If g_PageById.Exists(pageId) Then
            Set g_PageById(pageId) = Nothing
            g_PageById.Remove pageId
        End If
        If VBA.StrComp(g_LastRenderedPageId, pageId, VBA.vbTextCompare) = 0 Then
            g_LastRenderedPageId = VBA.vbNullString
        End If
    End If

    page.Dispose deleteWorksheet
    m_RemovePage = True
End Function

' Callstack[1]: ThisWorkbook.Workbook_Open -> ThisWorkbook.m_ResetWorkbookAndCreateMainPage -> private_ResetWorkbookAndCreateMainPage -> rt_PageManager.m_DisposeAllPages
' Callstack[2]: rt_RestoreManager.private_TryResetWorkbookBeforeRestore -> rt_PageManager.m_DisposeAllPages
Public Sub m_DisposeAllPages()
    Dim pageId As Variant
    Dim page As obj_IPage

    private_EnsureStorage

    For Each pageId In g_PageById.Keys
        Set page = g_PageById(pageId)
        If Not page Is Nothing Then
            page.Dispose False
        End If
        Set g_PageById(pageId) = Nothing
    Next pageId

    Set g_PageById = Nothing
    g_LastRenderedPageId = VBA.vbNullString
End Sub

' Callstack[1]: rt_PageManager.m_TrySerializeModuleSnapshot -> rt_PageManager.m_TryGetLastRenderedWorksheetName
Public Function m_TryGetLastRenderedWorksheetName(ByRef outWorksheetName As String) As Boolean
    Dim pageId As String
    Dim page As obj_IPage
    Dim pageBase As obj_PageBase

    outWorksheetName = VBA.vbNullString
    pageId = VBA.LCase$(VBA.Trim$(g_LastRenderedPageId))
    If VBA.Len(pageId) = 0 Then
        m_TryGetLastRenderedWorksheetName = True
        Exit Function
    End If

    If Not m_TryGetPageById(pageId, page) Then
        m_TryGetLastRenderedWorksheetName = True
        Exit Function
    End If

    Set pageBase = page.GetPageBase()
    If pageBase Is Nothing Then
        m_TryGetLastRenderedWorksheetName = True
        Exit Function
    End If
    If pageBase.Worksheet Is Nothing Then
        m_TryGetLastRenderedWorksheetName = True
        Exit Function
    End If

    outWorksheetName = VBA.Trim$(VBA.CStr(pageBase.Worksheet.Name))
    m_TryGetLastRenderedWorksheetName = True
End Function

' Callstack[1]: rt_PageManager.m_TryDeserializeModuleSnapshot -> rt_PageManager.m_TryRestoreLastRenderedWorksheetName
Public Function m_TryRestoreLastRenderedWorksheetName(ByVal worksheetName As String) As Boolean
    Dim page As obj_IPage
    Dim pageBase As obj_PageBase
    Dim pageId As String

    worksheetName = VBA.Trim$(worksheetName)
    If VBA.Len(worksheetName) = 0 Then
        g_LastRenderedPageId = VBA.vbNullString
        m_TryRestoreLastRenderedWorksheetName = True
        Exit Function
    End If

    If Not m_TryGetPageByWorksheetName(worksheetName, page) Then Exit Function
    If page Is Nothing Then Exit Function

    Set pageBase = page.GetPageBase()
    If pageBase Is Nothing Then Exit Function

    pageId = VBA.LCase$(VBA.Trim$(pageBase.PageId))
    If VBA.Len(pageId) = 0 Then pageId = private_TryResolvePageIdByObject(page)
    If VBA.Len(pageId) = 0 Then Exit Function

    g_LastRenderedPageId = pageId
    m_TryRestoreLastRenderedWorksheetName = True
End Function

' //
' // Internal
' //
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


Private Function private_TryRenderPagesCollection(ByVal pages As Collection, ByVal reasonPrefix As String) As Boolean
    Dim pageItem As Variant
    Dim page As obj_IPage
    Dim renderReason As String

    private_TryRenderPagesCollection = True
    If pages Is Nothing Then Exit Function

    renderReason = VBA.Trim$(reasonPrefix)
    If VBA.Len(renderReason) = 0 Then renderReason = "restore"

    For Each pageItem In pages
        Set page = Nothing
        Set page = pageItem
        If page Is Nothing Then GoTo ContinuePage
        If m_RenderPage(page, renderReason) Then GoTo ContinuePage

        private_TryRenderPagesCollection = False
        Exit Function
ContinuePage:
    Next pageItem
End Function


Private Function private_RegisterPage(ByVal pageId As String, ByVal page As obj_IPage) As Boolean
    Dim pageBase As obj_PageBase
    Dim sheetName As String

    pageId = VBA.LCase$(VBA.Trim$(pageId))
    If VBA.Len(pageId) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PageManager: page id is empty."
#End If
        Exit Function
    End If
    If page Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PageManager: page instance is not specified."
#End If
        Exit Function
    End If

    Set pageBase = page.GetPageBase()
    If pageBase Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PageManager: page base is not specified."
#End If
        Exit Function
    End If

    If pageBase.Worksheet Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PageManager: worksheet is not specified for page '" & pageId & "'."
#End If
        Exit Function
    End If

    private_EnsureStorage

    If g_PageById.Exists(pageId) Then
        Set g_PageById(pageId) = Nothing
        g_PageById.Remove pageId
    End If
    Set g_PageById(pageId) = page

    sheetName = VBA.vbNullString
    If Not pageBase.Worksheet Is Nothing Then sheetName = VBA.Replace$(VBA.Trim$(VBA.CStr(pageBase.Worksheet.Name)), "'", "''")
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "page-manager:register-page pageId='" & VBA.Replace$(pageId, "'", "''") & "' sheet='" & sheetName & "'"
#End If

    private_RegisterPage = True
End Function


Private Function private_TryFindPageByWorksheet( _
    ByVal ws As Worksheet, _
    ByRef outPage As obj_IPage, _
    ByRef outPageId As String _
) As Boolean
    Dim key As Variant
    Dim page As obj_IPage
    Dim pageBase As obj_PageBase
    Dim wsName As String
    Dim pageSheetName As String

    Set outPage = Nothing
    outPageId = VBA.vbNullString

    If ws Is Nothing Then Exit Function
    On Error Resume Next
    wsName = VBA.Trim$(VBA.CStr(ws.Name))
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    If VBA.Len(wsName) = 0 Then Exit Function

    private_EnsureStorage

    For Each key In g_PageById.Keys
        Set page = g_PageById(VBA.CStr(key))
        If page Is Nothing Then GoTo ContinuePage

        Set pageBase = page.GetPageBase()
        If pageBase Is Nothing Then GoTo ContinuePage
        If pageBase.Worksheet Is Nothing Then GoTo ContinuePage

        If pageBase.Worksheet Is ws Then
            Set outPage = page
            outPageId = VBA.LCase$(VBA.Trim$(VBA.CStr(key)))
            private_TryFindPageByWorksheet = True
            Exit Function
        End If

        On Error Resume Next
        pageSheetName = VBA.Trim$(VBA.CStr(pageBase.Worksheet.Name))
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo ContinuePage
        End If
        On Error GoTo 0
        If VBA.StrComp(pageSheetName, wsName, VBA.vbTextCompare) = 0 Then
            Set outPage = page
            outPageId = VBA.LCase$(VBA.Trim$(VBA.CStr(key)))
            private_TryFindPageByWorksheet = True
            Exit Function
        End If

ContinuePage:
    Next key
End Function


Private Function private_TryResolvePageIdByObject(ByVal page As obj_IPage) As String
    Dim key As Variant
    Dim currentPage As obj_IPage

    If page Is Nothing Then Exit Function
    private_EnsureStorage

    For Each key In g_PageById.Keys
        Set currentPage = g_PageById(key)
        If currentPage Is page Then
            private_TryResolvePageIdByObject = VBA.CStr(key)
            Exit Function
        End If
    Next key
End Function


Private Function private_ResolvePageIdPrefix(ByVal page As obj_IPage) As String
    private_ResolvePageIdPrefix = "page"
    If page Is Nothing Then Exit Function

    If TypeOf page Is obj_PageMain Then
        private_ResolvePageIdPrefix = "main"
        Exit Function
    End If

    If TypeOf page Is obj_PagePersonalCard Then
        private_ResolvePageIdPrefix = "generated"
        Exit Function
    End If
End Function

Private Function private_CreatePageInternal( _
    ByVal page As obj_IPage, _
    ByVal uiPath As String, _
    ByVal sheetName As String, _
    ByVal pageId As String, _
    Optional ByVal pageContext As Object = Nothing _
) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim pageBase As obj_PageBase
    Dim isPageInitialized As Boolean
    Dim errDescription As String

    If page Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PageManager: page instance is not specified."
#End If
        Exit Function
    End If

    Set wb = ThisWorkbook
    If wb Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PageManager: workbook is not specified."
#End If
        Exit Function
    End If

    uiPath = VBA.Trim$(uiPath)
    sheetName = VBA.Trim$(sheetName)
    pageId = VBA.LCase$(VBA.Trim$(pageId))
    If VBA.Len(pageId) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PageManager: page id is empty during create."
#End If
        Exit Function
    End If

    On Error GoTo EH_CREATE
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    If VBA.Len(sheetName) > 0 Then ws.Name = sheetName

    If Not page.Initialize(ws, uiPath, pageId, pageContext) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PageManager: page initialize failed for page id '" & VBA.Replace$(pageId, "'", "''") & "'."
#End If
        GoTo EH_FAIL
    End If
    isPageInitialized = True

    Set pageBase = page.GetPageBase()
    If pageBase Is Nothing Then GoTo EH_FAIL
    If pageBase.Worksheet Is Nothing Then GoTo EH_FAIL
    If VBA.StrComp(VBA.LCase$(VBA.Trim$(pageBase.PageId)), pageId, VBA.vbTextCompare) <> 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "PageManager: page initialize did not assign expected page id '" & VBA.Replace$(pageId, "'", "''") & "'."
#End If
        GoTo EH_FAIL
    End If

    private_CreatePageInternal = private_RegisterPage(pageId, page)
    Exit Function

EH_FAIL:
    On Error Resume Next
    If isPageInitialized Then
        page.Dispose False
    End If
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
    Exit Function

EH_CREATE:
    Application.DisplayAlerts = True
    errDescription = Err.Description
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.m_Diagnostic_LogError "PageManager: exception during page create for page id '" & VBA.Replace$(pageId, "'", "''") & "': " & VBA.Replace$(errDescription, "'", "''")
#End If
    Resume EH_FAIL
End Function

Private Function private_BuildPageId(Optional ByVal pageIdPrefix As String = "page") As String
    pageIdPrefix = VBA.LCase$(VBA.Trim$(pageIdPrefix))
    If VBA.Len(pageIdPrefix) = 0 Then pageIdPrefix = "page"

    g_PageIdSeed = g_PageIdSeed + 1
    private_BuildPageId = pageIdPrefix & "-" & VBA.Format$(VBA.Now, "yyyymmdd-hhnnss") & "-" & VBA.CStr(g_PageIdSeed)
End Function


Private Sub private_EnsureStorage()
    If g_PageById Is Nothing Then
        Set g_PageById = CreateObject("Scripting.Dictionary")
        g_PageById.CompareMode = 1
    End If
End Sub
