Option Explicit
#Const LOGGING_DEBUG_ENABLED = True

Private Sub Workbook_Open()
    Dim restoredPagesCount As Long
    Dim restoredOk As Boolean

    On Error GoTo EH

    restoredOk = rt_RestoreManager.fn_RestoreRuntimeState("Workbook_Open", restoredPagesCount)
    If restoredOk And restoredPagesCount > 0 Then
        Exit Sub
    End If

    If Not m_ResetWorkbookAndCreateMainPage("ThisWorkbook.Workbook_Open:main-create") Then Exit Sub

    Exit Sub
EH:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "PrototypeNew: Workbook_Open failed: " & Err.Description
#End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call rt_RestoreManager.fn_SaveRuntimeState
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo EH_SHEET_CHANGE
    rt_Bridge.fn_OnSheetChange Sh, Target
    Exit Sub

EH_SHEET_CHANGE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "PrototypeNew: Workbook_SheetChange failed: " & Err.Description
#End If
End Sub

Private Sub Workbook_SheetBeforeDelete(ByVal Sh As Object)
    Dim ws As Worksheet
    Dim sheetName As String

    If Not TypeOf Sh Is Worksheet Then Exit Sub
    Set ws = Sh

    On Error Resume Next
    sheetName = VBA.LCase$(VBA.Trim$(VBA.CStr(ws.Name)))
    Err.Clear
    On Error GoTo 0

    ' Временные листы участвуют только в сценариях reset/restore.
    ' Не пробрасываем их в PageManager, чтобы не трогать runtime-реестр страниц.
    If VBA.StrComp(sheetName, "__startup_tmp__", VBA.vbTextCompare) = 0 Then Exit Sub
    If VBA.StrComp(sheetName, "__restore_tmp__", VBA.vbTextCompare) = 0 Then Exit Sub

    Call ex_HelpersSheet.fn_RemovePageByWorksheet(ws)
End Sub

' //
' // API
' //
' Callstack[1]: ThisWorkbook.Workbook_Open -> ThisWorkbook.m_ResetWorkbookAndCreateMainPage
Public Function m_ResetWorkbookAndCreateMainPage( _
    Optional ByVal renderReason As String = "ThisWorkbook.m_ResetWorkbookAndCreateMainPage", _
    Optional ByVal showErrorUi As Boolean = True _
) As Boolean
    m_ResetWorkbookAndCreateMainPage = private_ResetWorkbookAndCreateMainPage(renderReason, showErrorUi)
End Function


Private Function private_ResetWorkbookAndCreateMainPage( _
    Optional ByVal renderReason As String = "ThisWorkbook.private_ResetWorkbookAndCreateMainPage", _
    Optional ByVal showErrorUi As Boolean = True _
) As Boolean
    Dim wb As Workbook
    Dim tmpWs As Worksheet
    Dim tmpSheetName As String
    Dim createdPage As obj_IPage
    Dim isPageCreated As Boolean

    Set wb = ThisWorkbook
    If wb Is Nothing Then Exit Function

    rt_PageManager.fn_DisposeAllPages

    On Error GoTo EH_RESET
    Application.DisplayAlerts = False

    Set tmpWs = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    tmpSheetName = private_BuildUniqueWorksheetName(wb, "__startup_tmp__")
    If VBA.Len(tmpSheetName) = 0 Then
        Application.DisplayAlerts = True
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to prepare temporary worksheet name."
#End If
        Exit Function
    End If
    tmpWs.Name = tmpSheetName

    Do While wb.Worksheets.Count > 1
        If wb.Worksheets(1) Is tmpWs Then
            wb.Worksheets(2).Delete
        Else
            wb.Worksheets(1).Delete
        End If
    Loop

    Application.DisplayAlerts = True
    On Error GoTo EH_CREATE

    Set createdPage = New obj_PageMain
    If createdPage Is Nothing Then Exit Function

    If Not rt_PageManager.fn_CreatePage(createdPage, "ui\DevUI.xml", "Main") Then GoTo EH_CREATE
    isPageCreated = True

    Application.DisplayAlerts = False
    tmpWs.Delete
    Application.DisplayAlerts = True
    Set tmpWs = Nothing

    If Not rt_PageManager.fn_RenderPage(createdPage, renderReason) Then GoTo EH_CREATE

    private_ResetWorkbookAndCreateMainPage = True
    Exit Function

EH_RESET:
    Application.DisplayAlerts = True
    If showErrorUi Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to reset workbook sheets: " & Err.Description
#End If
    End If
    Exit Function

EH_CREATE:
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not createdPage Is Nothing And isPageCreated Then
        Call rt_PageManager.fn_RemovePage(createdPage, True)
    End If
    If Not tmpWs Is Nothing Then
        Application.DisplayAlerts = False
        tmpWs.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
    If showErrorUi Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: failed to create default main page: " & Err.Description
#End If
    End If
End Function


Private Function private_BuildUniqueWorksheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim i As Long
    Dim suffix As String
    Dim candidate As String

    If wb Is Nothing Then Exit Function

    baseName = VBA.Trim$(baseName)
    If VBA.Len(baseName) = 0 Then baseName = "tmp_sheet"
    If VBA.Len(baseName) > 31 Then baseName = VBA.Left$(baseName, 31)

    If Not private_WorksheetNameExists(wb, baseName) Then
        private_BuildUniqueWorksheetName = baseName
        Exit Function
    End If

    For i = 1 To 9999
        suffix = "_" & VBA.CStr(i)
        candidate = VBA.Left$(baseName, 31 - VBA.Len(suffix)) & suffix
        If VBA.Len(candidate) = 0 Then candidate = "tmp" & suffix
        If Not private_WorksheetNameExists(wb, candidate) Then
            private_BuildUniqueWorksheetName = candidate
            Exit Function
        End If
    Next i
End Function


Private Function private_WorksheetNameExists(ByVal wb As Workbook, ByVal worksheetName As String) As Boolean
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function

    worksheetName = VBA.Trim$(worksheetName)
    If VBA.Len(worksheetName) = 0 Then Exit Function

    On Error Resume Next
    Set ws = wb.Worksheets(worksheetName)
    private_WorksheetNameExists = Not ws Is Nothing
    Err.Clear
    On Error GoTo 0
End Function
