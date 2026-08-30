VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_UndoAction"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = False

Implements obj_IUndoAction

Private Const SNAPSHOT_LOCATOR_KEY As String = "Locator"
Private Const SNAPSHOT_BEFORE_KEY As String = "Before"
Private Const SNAPSHOT_AFTER_KEY As String = "After"
' One undo action snapshot is split into locator + state before change + state after change.
Private Const LOCATOR_KIND_TABLE_ROW As String = "TableRow"
Private Const LOCATOR_KIND_RANGE As String = "Range"

Private Const LOCATOR_WORKBOOK_PATH_KEY As String = "WorkbookPath"
Private Const LOCATOR_WORKBOOK_NAME_KEY As String = "WorkbookName"
Private Const LOCATOR_SHEET_NAME_KEY As String = "SheetName"
Private Const LOCATOR_TABLE_NAME_KEY As String = "TableName"
Private Const LOCATOR_RANGE_ADDRESS_KEY As String = "RangeAddress"
Private Const LOCATOR_ROW_INDEX_KEY As String = "RowIndex"
Private Const LOCATOR_INSERTED_ROW_KEY As String = "InsertedRow"

Private Const STATE_COLUMN_INDEXES_KEY As String = "ColumnIndexes"
Private Const STATE_VALUES_KEY As String = "Values"
Private Const STATE_FILL_COLORS_KEY As String = "FillColors"
Private Const STATE_NUMBER_FORMAT_KEY As String = "NumberFormat"
Private Const STATE_FORMULA_KEY As String = "Formula"

Private m_ActionId As String
Private m_Caption As String
Private m_OperationKind As String
Private m_Snapshot As obj_SnapshotBag
Private m_LocatorSnapshot As obj_SnapshotBag
Private m_BeforeSnapshot As obj_SnapshotBag
Private m_AfterSnapshot As obj_SnapshotBag

Private Sub Class_Initialize()
    m_ActionId = private_GenerateActionId()
    m_Caption = "Undo action"
    m_OperationKind = LOCATOR_KIND_TABLE_ROW
End Sub

Private Function obj_IUndoAction_GetActionId() As String
    obj_IUndoAction_GetActionId = Me.GetActionId()
End Function

Private Function obj_IUndoAction_GetCaption() As String
    obj_IUndoAction_GetCaption = Me.GetCaption()
End Function

Private Function obj_IUndoAction_GetScopeKey() As String
    obj_IUndoAction_GetScopeKey = Me.GetScopeKey()
End Function

Private Function obj_IUndoAction_IsValid() As Boolean
    obj_IUndoAction_IsValid = Me.IsValid()
End Function

Private Function obj_IUndoAction_Execute(ByRef outErrorText As String) As Boolean
    obj_IUndoAction_Execute = Me.Execute(outErrorText)
End Function

Private Function obj_IUndoAction_Undo(ByRef outErrorText As String) As Boolean
    obj_IUndoAction_Undo = Me.Undo(outErrorText)
End Function

Private Function obj_IUndoAction_Redo(ByRef outErrorText As String) As Boolean
    obj_IUndoAction_Redo = Me.Redo(outErrorText)
End Function

' //
' // API
' //
Public Function InitializeFromSnapshot( _
    ByVal snapshot As obj_SnapshotBag, _
    Optional ByVal captionText As String = "Undo action", _
    Optional ByVal operationKind As String = "TableRow" _
) As Boolean
    Dim locatorBag As obj_SnapshotBag
    Dim beforeBag As obj_SnapshotBag
    Dim afterBag As obj_SnapshotBag

    If snapshot Is Nothing Then Exit Function

    Set m_Snapshot = snapshot
    Set m_LocatorSnapshot = Nothing
    Set m_BeforeSnapshot = Nothing
    Set m_AfterSnapshot = Nothing

    If Not private_TryGetNestedBagFromSnapshot(snapshot, SNAPSHOT_LOCATOR_KEY, locatorBag) Then Exit Function
    If Not private_TryGetNestedBagFromSnapshot(snapshot, SNAPSHOT_BEFORE_KEY, beforeBag) Then Exit Function
    If Not private_TryGetNestedBagFromSnapshot(snapshot, SNAPSHOT_AFTER_KEY, afterBag) Then Exit Function

    Set m_LocatorSnapshot = locatorBag
    Set m_BeforeSnapshot = beforeBag
    Set m_AfterSnapshot = afterBag

    captionText = VBA.Trim$(captionText)
    If VBA.Len(captionText) > 0 Then m_Caption = captionText

    operationKind = VBA.LCase$(VBA.Trim$(operationKind))
    If VBA.Len(operationKind) > 0 Then
        m_OperationKind = operationKind
    Else
        m_OperationKind = LOCATOR_KIND_TABLE_ROW
    End If

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "undo-action:init id='" & VBA.Replace$(m_ActionId, "'", "''") & "' caption='" & VBA.Replace$(m_Caption, "'", "''") & "' kind='" & VBA.Replace$(m_OperationKind, "'", "''") & "'"
#End If

    InitializeFromSnapshot = True
End Function

Public Function GetActionId() As String
    GetActionId = m_ActionId
End Function

Public Function GetCaption() As String
    GetCaption = m_Caption
End Function

Public Function GetScopeKey() As String
    GetScopeKey = m_OperationKind
End Function

Public Function IsValid() As Boolean
    Dim locatorBag As obj_SnapshotBag

    If m_Snapshot Is Nothing Then Exit Function
    If Not private_TryGetNestedBag(SNAPSHOT_LOCATOR_KEY, locatorBag) Then Exit Function
    If locatorBag Is Nothing Then Exit Function

    Select Case private_GetOperationKind()
        Case LOCATOR_KIND_RANGE
            IsValid = private_TryResolveRange(locatorBag, False, Nothing, Nothing)
        Case Else
            IsValid = private_TryResolveTableRow(locatorBag, False, Nothing, Nothing, Nothing)
    End Select

#If LOGGING_DEBUG_ENABLED Then
    If Not IsValid Then
        ex_Core.fn_Diagnostic_LogError "undo-action:is-valid false id='" & VBA.Replace$(m_ActionId, "'", "''") & "' kind='" & VBA.Replace$(private_GetOperationKind(), "'", "''") & "'"
    End If
#End If
End Function

Public Function Execute(ByRef outErrorText As String) As Boolean
    On Error GoTo EH_EXECUTE
    Execute = private_ApplyState(SNAPSHOT_AFTER_KEY, outErrorText)
    Exit Function

EH_EXECUTE:
    outErrorText = "Undo action execute failed. " & Err.Description
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "undo-action:execute-exception id='" & VBA.Replace$(m_ActionId, "'", "''") & "' err='" & VBA.Replace$(Err.Description, "'", "''") & "' source='" & VBA.Replace$(Err.Source, "'", "''") & "'"
#End If
End Function

Public Function Undo(ByRef outErrorText As String) As Boolean
    On Error GoTo EH_UNDO
    Undo = private_ApplyState(SNAPSHOT_BEFORE_KEY, outErrorText)
    Exit Function

EH_UNDO:
    outErrorText = "Undo action undo failed. " & Err.Description
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "undo-action:undo-exception id='" & VBA.Replace$(m_ActionId, "'", "''") & "' err='" & VBA.Replace$(Err.Description, "'", "''") & "' source='" & VBA.Replace$(Err.Source, "'", "''") & "'"
#End If
End Function

Public Function Redo(ByRef outErrorText As String) As Boolean
    On Error GoTo EH_REDO
    Redo = private_ApplyState(SNAPSHOT_AFTER_KEY, outErrorText)
    Exit Function

EH_REDO:
    outErrorText = "Undo action redo failed. " & Err.Description
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "undo-action:redo-exception id='" & VBA.Replace$(m_ActionId, "'", "''") & "' err='" & VBA.Replace$(Err.Description, "'", "''") & "' source='" & VBA.Replace$(Err.Source, "'", "''") & "'"
#End If
End Function

' //
' // Internal
' //
Private Function private_ApplyState(ByVal stateKey As String, ByRef outErrorText As String) As Boolean
    Dim locatorBag As obj_SnapshotBag
    Dim stateBag As obj_SnapshotBag
    Dim isUndoSnapshot As Boolean
    Dim stageName As String

    On Error GoTo EH_APPLY_STATE

    outErrorText = VBA.vbNullString
    stageName = "validate-snapshot"
    If m_Snapshot Is Nothing Then
        outErrorText = "Undo action snapshot is missing."
        Exit Function
    End If

    stageName = "resolve-locator"
    Set locatorBag = m_LocatorSnapshot
    If locatorBag Is Nothing Then
        outErrorText = "Undo action locator snapshot is missing."
        Exit Function
    End If

    stageName = "resolve-state"
    If VBA.StrComp(VBA.LCase$(stateKey), VBA.LCase$(SNAPSHOT_BEFORE_KEY), VBA.vbTextCompare) = 0 Then
        Set stateBag = m_BeforeSnapshot
    Else
        Set stateBag = m_AfterSnapshot
    End If
    If stateBag Is Nothing Then
        outErrorText = "Undo action state snapshot is missing: " & stateKey
        Exit Function
    End If

    isUndoSnapshot = (VBA.StrComp(VBA.LCase$(stateKey), VBA.LCase$(SNAPSHOT_BEFORE_KEY), VBA.vbTextCompare) = 0)

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "undo-action:apply-start id='" & VBA.Replace$(m_ActionId, "'", "''") & "' caption='" & VBA.Replace$(m_Caption, "'", "''") & "' kind='" & VBA.Replace$(private_GetOperationKind(), "'", "''") & "' state='" & VBA.Replace$(stateKey, "'", "''") & "'"
#End If

    ' The same action can replay either the before-state or the after-state.
    Select Case private_GetOperationKind()
        Case LOCATOR_KIND_RANGE
            private_ApplyState = private_ApplyRangeState(locatorBag, stateBag, outErrorText)
        Case Else
            private_ApplyState = private_ApplyTableRowState(locatorBag, stateBag, isUndoSnapshot, outErrorText)
    End Select

#If LOGGING_DEBUG_ENABLED Then
    If private_ApplyState Then
        ex_Core.fn_Diagnostic_LogInfo "undo-action:apply-done id='" & VBA.Replace$(m_ActionId, "'", "''") & "' state='" & VBA.Replace$(stateKey, "'", "''") & "'"
    Else
        ex_Core.fn_Diagnostic_LogError "undo-action:apply-failed id='" & VBA.Replace$(m_ActionId, "'", "''") & "' state='" & VBA.Replace$(stateKey, "'", "''") & "' err='" & VBA.Replace$(outErrorText, "'", "''") & "'"
    End If
#End If
    Exit Function

EH_APPLY_STATE:
    outErrorText = "Undo action apply failed at stage '" & stageName & "'. " & Err.Description
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "undo-action:apply-exception id='" & VBA.Replace$(m_ActionId, "'", "''") & "' stage='" & VBA.Replace$(stageName, "'", "''") & "' state='" & VBA.Replace$(stateKey, "'", "''") & "' err='" & VBA.Replace$(Err.Description, "'", "''") & "' source='" & VBA.Replace$(Err.Source, "'", "''") & "'"
#End If
End Function

Private Function private_ApplyTableRowState( _
    ByVal locatorBag As obj_SnapshotBag, _
    ByVal stateBag As obj_SnapshotBag, _
    ByVal isUndoSnapshot As Boolean, _
    ByRef outErrorText As String _
) As Boolean
    Dim targetWb As Workbook
    Dim targetWs As Worksheet
    Dim targetTable As ListObject
    Dim rowIndex As Long
    Dim isInsertedRow As Boolean
    Dim columnIndexes As Collection
    Dim values As Collection
    Dim fillColors As Collection
    Dim targetRowRange As Range
    Dim targetRow As ListRow

    outErrorText = VBA.vbNullString
    If Not private_TryResolveTableRow(locatorBag, True, targetWb, targetWs, targetTable) Then
        outErrorText = "Undo action target table is unavailable."
        Exit Function
    End If

    If Not locatorBag.TryGetLong(LOCATOR_ROW_INDEX_KEY, rowIndex) Then
        outErrorText = "Undo action row index is invalid."
        Exit Function
    End If

    locatorBag.TryGetBoolean LOCATOR_INSERTED_ROW_KEY, isInsertedRow

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "undo-action:table-row target table='" & VBA.Replace$(targetTable.Name, "'", "''") & "' rowIndex=" & VBA.CStr(rowIndex) & " inserted=" & VBA.IIf(isInsertedRow, "true", "false") & " undoSnapshot=" & VBA.IIf(isUndoSnapshot, "true", "false")
#End If

    ' Inserted rows undo by deleting the row, redo by recreating it and restoring values.
    If isInsertedRow And isUndoSnapshot Then
        If rowIndex <= targetTable.ListRows.Count Then
            targetTable.ListRows.Item(rowIndex).Delete
            private_ApplyTableRowState = True
            Exit Function
        End If

        outErrorText = "Undo action row is out of range for inserted-row rollback."
        Exit Function
    End If

    If Not private_TryGetCollection(stateBag, STATE_COLUMN_INDEXES_KEY, columnIndexes) Then Exit Function
    If Not private_TryGetCollection(stateBag, STATE_VALUES_KEY, values) Then Exit Function
    Call private_TryGetCollectionOptional(stateBag, STATE_FILL_COLORS_KEY, fillColors)
    If columnIndexes Is Nothing Or values Is Nothing Then Exit Function
    If columnIndexes.Count <> values.Count Then
        outErrorText = "Undo action state columns and values are inconsistent."
        Exit Function
    End If

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "undo-action:table-row values columns=" & VBA.CStr(columnIndexes.Count) & " values=" & VBA.CStr(values.Count)
#End If

    If isInsertedRow Then
        Set targetRow = private_InsertRowByIndex(targetTable, rowIndex)
        If targetRow Is Nothing Then
            outErrorText = "Undo action failed to re-insert row."
            Exit Function
        End If
        Set targetRowRange = targetRow.Range
    Else
        If rowIndex > targetTable.ListRows.Count Then
            outErrorText = "Undo action row is out of range."
            Exit Function
        End If
        Set targetRowRange = targetTable.ListRows.Item(rowIndex).Range
    End If

    If Not private_TryWriteRowValuesByIndexes(targetRowRange, columnIndexes, values, outErrorText) Then Exit Function
    If Not private_TryApplyRowFillColorsByIndexes(targetRowRange, columnIndexes, fillColors, outErrorText) Then Exit Function

    private_ApplyTableRowState = True
End Function

Private Function private_ApplyRangeState( _
    ByVal locatorBag As obj_SnapshotBag, _
    ByVal stateBag As obj_SnapshotBag, _
    ByRef outErrorText As String _
) As Boolean
    Dim targetWorkbook As Workbook
    Dim targetRange As Range
    Dim formulaValue As String
    Dim numberFormatValue As String

    outErrorText = VBA.vbNullString
    If Not private_TryResolveRange(locatorBag, True, targetWorkbook, targetRange) Then
        outErrorText = "Undo action target range is unavailable."
        Exit Function
    End If

    If Not stateBag.TryGetText(STATE_FORMULA_KEY, formulaValue) Then Exit Function
    If Not stateBag.TryGetText(STATE_NUMBER_FORMAT_KEY, numberFormatValue) Then Exit Function

#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "undo-action:range target='" & VBA.Replace$(targetRange.Address(False, False), "'", "''") & "' sheet='" & VBA.Replace$(targetRange.Worksheet.Name, "'", "''") & "'"
#End If

    ' Range undo/redo is a direct restore of the captured cell state.
    On Error GoTo EH_APPLY_RANGE
    targetRange.NumberFormat = numberFormatValue
    targetRange.Formula = formulaValue
    private_ApplyRangeState = True
    Exit Function

EH_APPLY_RANGE:
    outErrorText = "Failed to apply range state. " & Err.Description
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "UndoAction: apply range state failed: " & Err.Description
#End If
End Function

Private Function private_TryResolveTableRow( _
    ByVal locatorBag As obj_SnapshotBag, _
    ByVal allowOpenWorkbook As Boolean, _
    ByRef outWorkbook As Workbook, _
    ByRef outWorksheet As Worksheet, _
    ByRef outTable As ListObject _
) As Boolean
    Dim workbookPath As String
    Dim workbookName As String
    Dim sheetName As String
    Dim tableName As String

    Set outWorkbook = Nothing
    Set outWorksheet = Nothing
    Set outTable = Nothing
    If locatorBag Is Nothing Then Exit Function

    If Not locatorBag.TryGetText(LOCATOR_WORKBOOK_PATH_KEY, workbookPath) Then workbookPath = VBA.vbNullString
    If Not locatorBag.TryGetText(LOCATOR_WORKBOOK_NAME_KEY, workbookName) Then workbookName = VBA.vbNullString
    If Not locatorBag.TryGetText(LOCATOR_SHEET_NAME_KEY, sheetName) Then Exit Function
    If Not locatorBag.TryGetText(LOCATOR_TABLE_NAME_KEY, tableName) Then Exit Function

    If Not private_TryResolveWorkbook(workbookPath, workbookName, outWorkbook, allowOpenWorkbook) Then Exit Function
    If outWorkbook Is Nothing Then Exit Function

    On Error Resume Next
    Set outWorksheet = outWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If outWorksheet Is Nothing Then Exit Function

    On Error Resume Next
    Set outTable = outWorksheet.ListObjects(tableName)
    On Error GoTo 0
    If outTable Is Nothing Then Exit Function

    private_TryResolveTableRow = True
End Function

Private Function private_TryResolveRange( _
    ByVal locatorBag As obj_SnapshotBag, _
    ByVal allowOpenWorkbook As Boolean, _
    ByRef outWorkbook As Workbook, _
    ByRef outRange As Range _
) As Boolean
    Dim workbookPath As String
    Dim workbookName As String
    Dim sheetName As String
    Dim rangeAddress As String
    Dim ws As Worksheet

    Set outWorkbook = Nothing
    Set outRange = Nothing
    If locatorBag Is Nothing Then Exit Function

    If Not locatorBag.TryGetText(LOCATOR_WORKBOOK_PATH_KEY, workbookPath) Then workbookPath = VBA.vbNullString
    If Not locatorBag.TryGetText(LOCATOR_WORKBOOK_NAME_KEY, workbookName) Then workbookName = VBA.vbNullString
    If Not locatorBag.TryGetText(LOCATOR_SHEET_NAME_KEY, sheetName) Then Exit Function
    If Not locatorBag.TryGetText(LOCATOR_RANGE_ADDRESS_KEY, rangeAddress) Then Exit Function

    If Not private_TryResolveWorkbook(workbookPath, workbookName, outWorkbook, allowOpenWorkbook) Then Exit Function
    If outWorkbook Is Nothing Then Exit Function

    On Error Resume Next
    Set ws = outWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set outRange = ws.Range(rangeAddress)
    On Error GoTo 0
    If outRange Is Nothing Then Exit Function

    private_TryResolveRange = True
End Function

Private Function private_TryResolveWorkbook( _
    ByVal workbookPath As String, _
    ByVal workbookName As String, _
    ByRef outWorkbook As Workbook, _
    ByVal allowOpenWorkbook As Boolean _
) As Boolean
    Dim wb As Workbook
    Dim normalizedPath As String

    Set outWorkbook = Nothing
    normalizedPath = VBA.LCase$(VBA.Trim$(workbookPath))
    workbookName = VBA.Trim$(workbookName)

    For Each wb In Application.Workbooks
        If VBA.Len(normalizedPath) > 0 Then
            If VBA.StrComp(VBA.LCase$(VBA.Trim$(wb.FullName)), normalizedPath, VBA.vbBinaryCompare) = 0 Then
                Set outWorkbook = wb
                private_TryResolveWorkbook = True
#If LOGGING_DEBUG_ENABLED Then
                ex_Core.fn_Diagnostic_LogInfo "undo-action:resolve-workbook hit-open-by-path path='" & VBA.Replace$(VBA.Trim$(wb.FullName), "'", "''") & "'"
#End If
                Exit Function
            End If
        End If

        If VBA.Len(workbookName) > 0 Then
            If VBA.StrComp(wb.Name, workbookName, VBA.vbTextCompare) = 0 Then
                Set outWorkbook = wb
                private_TryResolveWorkbook = True
#If LOGGING_DEBUG_ENABLED Then
                ex_Core.fn_Diagnostic_LogInfo "undo-action:resolve-workbook hit-open-by-name name='" & VBA.Replace$(wb.Name, "'", "''") & "'"
#End If
                Exit Function
            End If
        End If
    Next wb

    If Not allowOpenWorkbook Then Exit Function
    workbookPath = VBA.Trim$(workbookPath)
    If VBA.Len(workbookPath) = 0 Then Exit Function
    If VBA.Len(VBA.Dir$(workbookPath)) = 0 Then Exit Function

    On Error Resume Next
    Set outWorkbook = Application.Workbooks.Open(workbookPath)
    On Error GoTo 0
    private_TryResolveWorkbook = Not outWorkbook Is Nothing
#If LOGGING_DEBUG_ENABLED Then
    If private_TryResolveWorkbook Then
        ex_Core.fn_Diagnostic_LogInfo "undo-action:resolve-workbook opened path='" & VBA.Replace$(workbookPath, "'", "''") & "'"
    Else
        ex_Core.fn_Diagnostic_LogError "undo-action:resolve-workbook failed path='" & VBA.Replace$(workbookPath, "'", "''") & "' name='" & VBA.Replace$(workbookName, "'", "''") & "'"
    End If
#End If
End Function

Private Function private_TryGetNestedBag(ByVal keyText As String, ByRef outBag As obj_SnapshotBag) As Boolean
    private_TryGetNestedBag = private_TryGetNestedBagFromSnapshot(m_Snapshot, keyText, outBag)
End Function

Private Function private_TryGetNestedBagFromSnapshot( _
    ByVal snapshot As obj_SnapshotBag, _
    ByVal keyText As String, _
    ByRef outBag As obj_SnapshotBag _
) As Boolean
    Dim rawObject As Object

    Set outBag = Nothing
    If snapshot Is Nothing Then Exit Function
    If Not snapshot.TryGetObject(keyText, rawObject) Then Exit Function
    ' Nested undo parts are stored as obj_SnapshotBag objects under known keys.
    On Error Resume Next
    Set outBag = rawObject
    On Error GoTo 0
    private_TryGetNestedBagFromSnapshot = Not outBag Is Nothing
End Function

Private Function private_TryResolveLocatorBag(ByRef outBag As obj_SnapshotBag) As Boolean
    private_TryGetNestedBag SNAPSHOT_LOCATOR_KEY, outBag
End Function

Private Function private_TryGetCollection(ByVal bag As obj_SnapshotBag, ByVal keyText As String, ByRef outCollection As Collection) As Boolean
    Dim rawObject As Object

    Set outCollection = Nothing
    If bag Is Nothing Then Exit Function
    If Not bag.TryGetObject(keyText, rawObject) Then Exit Function
    On Error Resume Next
    Set outCollection = rawObject
    On Error GoTo 0
    private_TryGetCollection = Not outCollection Is Nothing
End Function

Private Function private_TryGetCollectionOptional(ByVal bag As obj_SnapshotBag, ByVal keyText As String, ByRef outCollection As Collection) As Boolean
    Set outCollection = Nothing
    If bag Is Nothing Then
        private_TryGetCollectionOptional = True
        Exit Function
    End If
    If Not private_TryGetCollection(bag, keyText, outCollection) Then
        Set outCollection = Nothing
    End If
    private_TryGetCollectionOptional = True
End Function

Private Function private_TryWriteRowValuesByIndexes( _
    ByVal rowRange As Range, _
    ByVal columnIndexes As Collection, _
    ByVal values As Collection, _
    ByRef outErrorText As String _
) As Boolean
    Dim i As Long
    Dim colIndex As Long

    outErrorText = VBA.vbNullString
    If rowRange Is Nothing Then
        outErrorText = "Target row range is missing."
        Exit Function
    End If
    If columnIndexes Is Nothing Then
        outErrorText = "Column index list is missing."
        Exit Function
    End If
    If values Is Nothing Then
        outErrorText = "Value list is missing."
        Exit Function
    End If
    If columnIndexes.Count <> values.Count Then
        outErrorText = "Column index list and value list have different lengths."
        Exit Function
    End If

    On Error GoTo EH_WRITE
    For i = 1 To columnIndexes.Count
        colIndex = VBA.CLng(columnIndexes.Item(i))
        If colIndex <= 0 Then GoTo ContinueCell
        rowRange.Cells(1, colIndex).Value2 = values.Item(i)
ContinueCell:
    Next i

    private_TryWriteRowValuesByIndexes = True
    Exit Function

EH_WRITE:
    outErrorText = "Failed to apply row values. " & Err.Description
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "UndoAction: write failed: " & Err.Description
#End If
End Function

Private Function private_TryApplyRowFillColorsByIndexes( _
    ByVal rowRange As Range, _
    ByVal columnIndexes As Collection, _
    ByVal fillColors As Collection, _
    ByRef outErrorText As String _
) As Boolean
    Dim i As Long
    Dim colIndex As Long

    outErrorText = VBA.vbNullString
    If fillColors Is Nothing Then
        private_TryApplyRowFillColorsByIndexes = True
        Exit Function
    End If
    If rowRange Is Nothing Then
        outErrorText = "Target row range is missing."
        Exit Function
    End If
    If columnIndexes Is Nothing Then
        outErrorText = "Column index list is missing."
        Exit Function
    End If
    If fillColors.Count <> columnIndexes.Count Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "undo-action:fill-colors-count-mismatch columns=" & VBA.CStr(columnIndexes.Count) & " fills=" & VBA.CStr(fillColors.Count)
#End If
        private_TryApplyRowFillColorsByIndexes = True
        Exit Function
    End If

    On Error GoTo EH_APPLY_FILL
    For i = 1 To columnIndexes.Count
        colIndex = VBA.CLng(columnIndexes.Item(i))
        If colIndex <= 0 Then GoTo ContinueCell
        rowRange.Cells(1, colIndex).Interior.Color = fillColors.Item(i)
ContinueCell:
    Next i

    private_TryApplyRowFillColorsByIndexes = True
    Exit Function

EH_APPLY_FILL:
    outErrorText = "Failed to apply row fill colors. " & Err.Description
End Function

Private Function private_InsertRowByIndex(ByVal tableObj As ListObject, ByVal rowIndex As Long) As ListRow
    If tableObj Is Nothing Then Exit Function
    If rowIndex <= 0 Then Exit Function

    On Error Resume Next
    If rowIndex > tableObj.ListRows.Count Then
        Set private_InsertRowByIndex = tableObj.ListRows.Add
    Else
        Set private_InsertRowByIndex = tableObj.ListRows.Add(Position:=rowIndex)
    End If
    On Error GoTo 0
End Function

Private Function private_GetOperationKind() As String
    private_GetOperationKind = VBA.LCase$(VBA.Trim$(m_OperationKind))
    If VBA.Len(private_GetOperationKind) = 0 Then private_GetOperationKind = LOCATOR_KIND_TABLE_ROW
End Function

Private Function private_GenerateActionId() As String
    Static idSeed As Long

    idSeed = idSeed + 1
    private_GenerateActionId = "undo-" & VBA.Format$(VBA.Now, "yyyymmdd-hhnnss") & "-" & VBA.CStr(idSeed)
End Function
