Attribute VB_Name = "rt_PageManager"
Option Explicit

Private g_PageById As Object
Private g_LastRenderedPageId As String
Private g_PageIdSeed As Long
Private Const MODULE_SNAPSHOT_ROOT As String = "pageManagerState"
Private Const MODULE_SNAPSHOT_NS As String = "urn:excelprototype:runtime-module:page-manager:v1"

' //
' // API
' //
' Callstack[1]: rt_Snapshots.m_SaveRuntimeGlobalsSnapshot -> private_TryAppendModuleSnapshot -> private_TrySerializeRuntimeModuleSnapshot -> rt_PageManager.m_TrySerializeModuleSnapshot
Public Function m_TrySerializeModuleSnapshot(ByRef outSnapshotXml As String) As Boolean
    Dim dom As Object
    Dim rootNode As Object
    Dim worksheetName As String

    outSnapshotXml = VBA.vbNullString

    If Not ex_Core.m_CustomXmlPartStore_TryCreateEmptyDom(MODULE_SNAPSHOT_ROOT, MODULE_SNAPSHOT_NS, dom) Then Exit Function
    Set rootNode = dom.DocumentElement
    If rootNode Is Nothing Then
        ex_Core.m_Diagnostic_LogError "PageManager: module snapshot root node is missing."
        Exit Function
    End If

    worksheetName = VBA.vbNullString
    If Not m_TryGetLastRenderedWorksheetName(worksheetName) Then Exit Function
    rootNode.setAttribute "lastRenderedSheetName", worksheetName

    outSnapshotXml = VBA.CStr(dom.XML)
    m_TrySerializeModuleSnapshot = (VBA.Len(VBA.Trim$(outSnapshotXml)) > 0)
End Function


' Callstack[1]: rt_Snapshots.m_RestoreRuntimeGlobalsSnapshot -> private_TryDeserializeRuntimeModuleSnapshot -> rt_PageManager.m_TryDeserializeModuleSnapshot
Public Function m_TryDeserializeModuleSnapshot(ByVal snapshotXml As String) As Boolean
    Dim dom As Object
    Dim rootNode As Object
    Dim worksheetName As String

    snapshotXml = VBA.Trim$(snapshotXml)
    If VBA.Len(snapshotXml) = 0 Then
        m_TryDeserializeModuleSnapshot = True
        Exit Function
    End If

    If Not ex_Core.m_CustomXmlPartStore_TryLoadDomFromXml(snapshotXml, dom) Then Exit Function
    Set rootNode = dom.DocumentElement
    If rootNode Is Nothing Then
        ex_Core.m_Diagnostic_LogError "PageManager: module snapshot root node is missing."
        Exit Function
    End If

    worksheetName = VBA.Trim$(VBA.CStr(rootNode.getAttribute("lastRenderedSheetName")))
    If Not m_TryRestoreLastRenderedWorksheetName(worksheetName) Then Exit Function

    m_TryDeserializeModuleSnapshot = True
End Function


' Callstack[1]: ThisWorkbook.Workbook_Open -> ThisWorkbook.m_ResetWorkbookAndCreateMainPage -> private_ResetWorkbookAndCreateMainPage -> rt_PageManager.m_CreatePage
' Callstack[2]: ex_Core.private_TryRecoverUiAfterUpdate -> ThisWorkbook.m_ResetWorkbookAndCreateMainPage -> private_ResetWorkbookAndCreateMainPage -> rt_PageManager.m_CreatePage
' Callstack[3]: rt_Snapshots.m_RestorePageSnapshots -> rt_PageManager.m_CreatePage
Public Function m_CreatePage( _
    ByVal xmlUiPath As String, _
    ByVal pageType As PageTypeEnum, _
    ByRef outPageId As String, _
    Optional ByVal sheetName As String = VBA.vbNullString _
) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim page As obj_IPage
    Dim pageId As String
    Dim createStep As String

    outPageId = VBA.vbNullString

    createStep = "resolve-workbook"
    Set wb = ThisWorkbook
    If wb Is Nothing Then
        ex_Core.m_Diagnostic_LogError "PageManager: workbook is not specified."
        Exit Function
    End If

    createStep = "normalize-page-type"
    pageType = private_NormalizePageType(pageType)

    On Error GoTo EH_CREATE
    createStep = "add-worksheet"
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))

    createStep = "apply-sheet-name"
    sheetName = VBA.Trim$(sheetName)
    If VBA.Len(sheetName) > 0 Then ws.Name = sheetName

    createStep = "generate-page-id"
    pageId = private_GeneratePageId(pageType)
    If VBA.Len(pageId) = 0 Then GoTo EH_ADD

    createStep = "create-page-instance"
    If Not private_TryCreatePageByPageType(pageType, page) Then GoTo EH_ADD
    If page Is Nothing Then GoTo EH_ADD
    createStep = "initialize-page-instance"
    If Not page.Initialize(ws, VBA.Trim$(xmlUiPath), VBA.CLng(pageType), pageId) Then GoTo EH_ADD
    createStep = "register-page"
    If Not private_RegisterPage(pageId, page) Then GoTo EH_ADD

    outPageId = pageId
    m_CreatePage = True
    Exit Function

EH_ADD:
    On Error Resume Next
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Exit Function

EH_CREATE:
    Application.DisplayAlerts = True
    ex_Core.m_Diagnostic_LogError "PageManager: failed to create page at step '" & createStep & "': [" & VBA.CStr(Err.Number) & "] " & Err.Description
End Function


' Callstack[1]: rt_PageManager.m_RenderPageById -> rt_PageManager.m_TryGetPageById
' Callstack[2]: rt_PageManager.m_RemovePageById -> rt_PageManager.m_TryGetPageById
' Callstack[3]: ThisWorkbook.m_ResetWorkbookAndCreateMainPage -> private_ResetWorkbookAndCreateMainPage -> rt_PageManager.m_TryGetPageById
' Callstack[4]: rt_Snapshots.m_RestorePageSnapshots -> rt_PageManager.m_TryGetPageById
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
        ex_Core.m_Diagnostic_LogError "page-manager:get-by-worksheet input-invalid worksheet is not specified"
        Exit Function
    End If

    On Error Resume Next
    wsName = VBA.Trim$(VBA.CStr(ws.Name))
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ex_Core.m_Diagnostic_LogError "page-manager:get-by-worksheet worksheet-name-unavailable"
        Exit Function
    End If
    On Error GoTo 0

    wsName = VBA.Replace$(wsName, "'", "''")
    If Not private_TryFindPageByWorksheet(ws, outPage, resolvedPageId) Then
        ex_Core.m_Diagnostic_LogError "page-manager:get-by-worksheet page-not-found sheet='" & wsName & "'"
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


' Callstack[1]: rt_Snapshots.private_TryCollectAllPages -> rt_PageManager.m_TryGetPagesByType
Public Function m_TryGetPagesByType(ByVal pageType As PageTypeEnum, ByRef outPages As Collection) As Boolean
    Dim pageId As Variant
    Dim page As obj_IPage
    Dim pageBase As obj_PageBase
    Dim normalizedType As PageTypeEnum

    Set outPages = New Collection
    normalizedType = private_NormalizePageType(pageType)

    private_EnsureStorage

    For Each pageId In g_PageById.Keys
        Set page = g_PageById(pageId)
        If page Is Nothing Then GoTo ContinueLoop

        Set pageBase = page.GetPageBase()
        If pageBase Is Nothing Then GoTo ContinueLoop

        If VBA.CLng(pageBase.PageType) = VBA.CLng(normalizedType) Then
            outPages.Add page
        End If

ContinueLoop:
    Next pageId

    m_TryGetPagesByType = True
End Function


' Callstack[1]: ThisWorkbook.m_ResetWorkbookAndCreateMainPage -> private_ResetWorkbookAndCreateMainPage -> rt_PageManager.m_RenderPageById
' Callstack[2]: ex_Core.private_TryRecoverUiAfterUpdate -> ThisWorkbook.m_ResetWorkbookAndCreateMainPage -> private_ResetWorkbookAndCreateMainPage -> rt_PageManager.m_RenderPageById
Public Function m_RenderPageById(ByVal pageId As String, Optional ByVal reason As String = VBA.vbNullString) As Boolean
    Dim page As obj_IPage

    If Not m_TryGetPageById(pageId, page) Then
        ex_Core.m_Diagnostic_LogError "page-manager:render-by-id input-invalid page is not found"
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
' Callstack[9]: rt_Snapshots.m_RestorePageSnapshots(renderRestored:=True) -> rt_PageManager.m_RenderPage
Public Function m_RenderPage(ByVal page As obj_IPage, Optional ByVal reason As String = VBA.vbNullString) As Boolean
    Dim pageBase As obj_PageBase
    Dim sheetName As String
    Dim normalizedReason As String
    Dim errDescription As String
    Dim pageId As String

    If page Is Nothing Then
        ex_Core.m_Diagnostic_LogError "page-manager:render input-invalid page is not specified"
        Exit Function
    End If

    Set pageBase = page.GetPageBase()
    If pageBase Is Nothing Then
        ex_Core.m_Diagnostic_LogError "page-manager:render input-invalid page base is not specified"
        Exit Function
    End If

    If pageBase.Worksheet Is Nothing Then
        ex_Core.m_Diagnostic_LogError "page-manager:render input-invalid worksheet is not specified"
        Exit Function
    End If

    sheetName = VBA.Replace$(VBA.CStr(pageBase.Worksheet.Name), "'", "''")
    normalizedReason = VBA.Trim$(VBA.CStr(reason))
    If VBA.Len(normalizedReason) = 0 Then normalizedReason = "manual"

    ex_Core.m_Diagnostic_LogInfo "page-manager:render-start sheet='" & sheetName & "' reason='" & VBA.Replace$(normalizedReason, "'", "''") & "'"

    On Error GoTo EH_RENDER
    m_RenderPage = page.Render()

    If m_RenderPage Then
        pageId = VBA.LCase$(VBA.Trim$(pageBase.PageId))
        If VBA.Len(pageId) = 0 Then pageId = private_TryResolvePageIdByObject(page)
        g_LastRenderedPageId = pageId
        ex_Core.m_Diagnostic_LogInfo "page-manager:render-done sheet='" & sheetName & "'"
    Else
        ex_Core.m_Diagnostic_LogError "page-manager:render-failed sheet='" & sheetName & "'"
    End If
    Exit Function

EH_RENDER:
    errDescription = Err.Description
    ex_Core.m_Diagnostic_LogError "page-manager:render-exception sheet='" & sheetName & "' err='" & VBA.Replace$(errDescription, "'", "''") & "'"
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
' Callstack[2]: rt_Snapshots.private_TryResetWorkbookBeforeRestore -> rt_PageManager.m_DisposeAllPages
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
Private Function private_RegisterPage(ByVal pageId As String, ByVal page As obj_IPage) As Boolean
    Dim pageBase As obj_PageBase
    Dim sheetName As String

    pageId = VBA.LCase$(VBA.Trim$(pageId))
    If VBA.Len(pageId) = 0 Then
        ex_Core.m_Diagnostic_LogError "PageManager: page id is empty."
        Exit Function
    End If
    If page Is Nothing Then
        ex_Core.m_Diagnostic_LogError "PageManager: page instance is not specified."
        Exit Function
    End If

    Set pageBase = page.GetPageBase()
    If pageBase Is Nothing Then
        ex_Core.m_Diagnostic_LogError "PageManager: page base is not specified."
        Exit Function
    End If

    If pageBase.Worksheet Is Nothing Then
        ex_Core.m_Diagnostic_LogError "PageManager: worksheet is not specified for page '" & pageId & "'."
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
    ex_Core.m_Diagnostic_LogInfo "page-manager:register-page pageId='" & VBA.Replace$(pageId, "'", "''") & "' sheet='" & sheetName & "'"

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


Private Function private_TryCreatePageByPageType(ByVal pageType As PageTypeEnum, ByRef outPage As obj_IPage) As Boolean
    Dim pageMain As obj_PageMain

    Set outPage = Nothing
    pageType = private_NormalizePageType(pageType)

    Select Case pageType
        Case PageTypeMain, PageTypeGenerated
            Set pageMain = New obj_PageMain
            Set outPage = pageMain
            private_TryCreatePageByPageType = True
    End Select
End Function


Private Function private_NormalizePageType(ByVal pageType As Long) As PageTypeEnum
    Select Case VBA.CLng(pageType)
        Case PageTypeGenerated
            private_NormalizePageType = PageTypeGenerated

        Case Else
            private_NormalizePageType = PageTypeMain
    End Select
End Function


Private Function private_GeneratePageId(ByVal pageType As PageTypeEnum) As String
    Dim prefix As String

    pageType = private_NormalizePageType(pageType)
    Select Case pageType
        Case PageTypeGenerated
            prefix = "generated"
        Case Else
            prefix = "main"
    End Select

    g_PageIdSeed = g_PageIdSeed + 1
    private_GeneratePageId = VBA.LCase$(prefix & "-" & VBA.Format$(VBA.Now, "yyyymmdd-hhnnss") & "-" & VBA.CStr(g_PageIdSeed))
End Function


Private Sub private_EnsureStorage()
    If g_PageById Is Nothing Then
        Set g_PageById = CreateObject("Scripting.Dictionary")
        g_PageById.CompareMode = 1
    End If
End Sub
