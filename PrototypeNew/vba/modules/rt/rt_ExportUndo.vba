Attribute VB_Name = "rt_ExportUndo"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const UNDO_MENU_TEXT As String = "Undo DailyScope Export"

Private g_DailyScopeUndoStack As Collection

Public Sub fn_Module_Dispose()
    Set g_DailyScopeUndoStack = Nothing
End Sub

Public Function fn_PushDailyScopeExportUndo(ByVal snapshot As Object) As Boolean
    If snapshot Is Nothing Then Exit Function

    private_EnsureStack
    g_DailyScopeUndoStack.Add snapshot
    private_RegisterUndo
    fn_PushDailyScopeExportUndo = True
End Function

Public Sub fn_UndoLastDailyScopeExport()
    Dim snapshot As Object

    On Error GoTo EH_UNDO
    private_EnsureStack
    If g_DailyScopeUndoStack.Count <= 0 Then Exit Sub

    Set snapshot = g_DailyScopeUndoStack.Item(g_DailyScopeUndoStack.Count)
    g_DailyScopeUndoStack.Remove g_DailyScopeUndoStack.Count

    If Not private_TryApplySnapshot(snapshot) Then
        VBA.MsgBox "PrototypeNew: failed to undo DailyScope export.", VBA.vbExclamation, "PrototypeNew / DailyScope export"
    End If

    If g_DailyScopeUndoStack.Count > 0 Then private_RegisterUndo
    Exit Sub

EH_UNDO:
    VBA.MsgBox "PrototypeNew: DailyScope undo failed. " & Err.Description, VBA.vbExclamation, "PrototypeNew / DailyScope export"
End Sub

Private Sub private_EnsureStack()
    If g_DailyScopeUndoStack Is Nothing Then Set g_DailyScopeUndoStack = New Collection
End Sub

Private Sub private_RegisterUndo()
    Dim macroRef As String

    macroRef = "'" & ThisWorkbook.Name & "'!rt_ExportUndo.fn_UndoLastDailyScopeExport"
    Application.OnUndo UNDO_MENU_TEXT, macroRef
End Sub

Private Function private_TryApplySnapshot(ByVal snapshot As Object) As Boolean
    Dim targetWb As Workbook
    Dim targetWs As Worksheet
    Dim targetTable As ListObject
    Dim rowIndex As Long
    Dim isInsertedRow As Boolean
    Dim columnIndexes As Collection
    Dim oldValues As Collection
    Dim i As Long

    If snapshot Is Nothing Then Exit Function

    If Not private_TryOpenWorkbookBySnapshot(snapshot, targetWb) Then Exit Function
    If targetWb Is Nothing Then Exit Function

    On Error Resume Next
    Set targetWs = targetWb.Worksheets(VBA.CStr(snapshot("SheetName")))
    On Error GoTo 0
    If targetWs Is Nothing Then Exit Function

    On Error Resume Next
    Set targetTable = targetWs.ListObjects(VBA.CStr(snapshot("TableName")))
    On Error GoTo 0
    If targetTable Is Nothing Then Exit Function

    rowIndex = VBA.CLng(snapshot("RowIndex"))
    If rowIndex <= 0 Then Exit Function

    isInsertedRow = VBA.CBool(snapshot("InsertedRow"))
    If isInsertedRow Then
        If rowIndex > targetTable.ListRows.Count Then Exit Function
        targetTable.ListRows.Item(rowIndex).Delete
        private_TryApplySnapshot = True
        Exit Function
    End If

    Set columnIndexes = Nothing
    Set oldValues = Nothing
    On Error Resume Next
    Set columnIndexes = snapshot("ColumnIndexes")
    Set oldValues = snapshot("OldValues")
    On Error GoTo 0

    If columnIndexes Is Nothing Then Exit Function
    If oldValues Is Nothing Then Exit Function
    If columnIndexes.Count <> oldValues.Count Then Exit Function
    If rowIndex > targetTable.ListRows.Count Then Exit Function

    For i = 1 To columnIndexes.Count
        targetTable.ListRows.Item(rowIndex).Range.Cells(1, VBA.CLng(columnIndexes.Item(i))).Value2 = oldValues.Item(i)
    Next i

    private_TryApplySnapshot = True
End Function

Private Function private_TryOpenWorkbookBySnapshot(ByVal snapshot As Object, ByRef outWorkbook As Workbook) As Boolean
    Dim workbookPath As String
    Dim workbookName As String
    Dim wb As Workbook

    Set outWorkbook = Nothing
    If snapshot Is Nothing Then Exit Function

    workbookPath = VBA.Trim$(VBA.CStr(snapshot("WorkbookPath")))
    workbookName = VBA.Trim$(VBA.CStr(snapshot("WorkbookName")))

    If VBA.Len(workbookPath) > 0 Then
        For Each wb In Application.Workbooks
            If VBA.StrComp(VBA.LCase$(VBA.Trim$(wb.FullName)), VBA.LCase$(workbookPath), VBA.vbBinaryCompare) = 0 Then
                Set outWorkbook = wb
                private_TryOpenWorkbookBySnapshot = True
                Exit Function
            End If
        Next wb
    End If

    If VBA.Len(workbookName) > 0 Then
        For Each wb In Application.Workbooks
            If VBA.StrComp(wb.Name, workbookName, VBA.vbTextCompare) = 0 Then
                Set outWorkbook = wb
                private_TryOpenWorkbookBySnapshot = True
                Exit Function
            End If
        Next wb
    End If

    If VBA.Len(workbookPath) = 0 Then Exit Function
    If VBA.Len(VBA.Dir$(workbookPath)) = 0 Then Exit Function

    Set outWorkbook = Application.Workbooks.Open(workbookPath)
    private_TryOpenWorkbookBySnapshot = Not outWorkbook Is Nothing
End Function
