Attribute VB_Name = "ex_PEB_HistoryRevisions"
Option Explicit

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_ReadState( _
    ByVal table As ListObject, _
    ByVal rowIndex As Long, _
    ByVal row As Object, _
    Optional ByVal stateColumns As Object = Nothing, _
    Optional ByVal data As Variant, _
    Optional ByVal stateFields As Variant _
)
    Dim key As Variant
    Dim column As ListColumn
    Dim value As Variant

    row("Revision") = "0"
    row("Status") = "Active"
    row("LastChangeId") = VBA.vbNullString
    If IsMissing(stateFields) Then
        stateFields = VBA.Split(ex_PEB_Export.fn_Config("history.versions.stateFields"), "|")
    End If
    For Each key In stateFields
        value = Empty
        If Not stateColumns Is Nothing Then
            If stateColumns.Exists(key) Then
                value = data(rowIndex, stateColumns(key))
            End If
        Else
            Set column = Nothing
            On Error Resume Next
            Set column = table.ListColumns(VBA.CStr(key))
            On Error GoTo 0
            If Not column Is Nothing Then
                value = table.DataBodyRange.Cells(rowIndex, column.Index).Value2
            End If
        End If
        If VBA.Len(ex_PEB_Export.fn_Text(value)) > 0 Then
            row(key) = ex_PEB_Export.fn_Text(value)
        End If
    Next key
    If row("Status") <> "Active" And row("Status") <> "Cancelled" Then
        ex_PEB_MovementLinks.fn_Fail "InvalidEventStatus", row("RowId")
    End If
    If row("Revision") Like "*[!0-9]*" Then
        ex_PEB_MovementLinks.fn_Fail "InvalidEventRevision", row("RowId")
    End If
    If VBA.CDbl(row("Revision")) > 2147483646# Then
        ex_PEB_MovementLinks.fn_Fail "InvalidEventRevision", row("RowId")
    End If
    If VBA.CLng(row("Revision")) > 0 Then
        If row("LastChangeId") <> row("RowId") & "/R" & row("Revision") Then
            ex_PEB_MovementLinks.fn_Fail "InvalidEventRevision", row("RowId")
        End If
    End If
End Sub

Public Sub fn_Prepare( _
    ByVal source As ListObject, _
    ByVal rows As Collection, _
    ByVal previous As Object, _
    ByVal values As Object, _
    ByVal changed As Boolean, _
    ByRef transaction As Object, _
    Optional ByVal nextRows As Collection = Nothing _
)
    Dim table As ListObject
    Dim key As Variant
    Dim column As ListColumn
    Dim revision As Long
    Dim status As String
    Dim changeType As String
    Dim reason As String
    Dim baseline As String
    Dim snapshotIds As Variant
    Dim snapshotIndex As Long
    Dim currentFound As Boolean

    Set transaction = VBA.CreateObject("Scripting.Dictionary")
    Set table = private_Table(source.Parent.Parent, "history.versions", source)
    transaction.Add "Table", table
    transaction.Add "OriginalRows", table.ListRows.Count
    revision = 0
    status = "Active"
    If Not previous Is Nothing Then
        revision = VBA.CLng(previous("Revision"))
        status = previous("Status")
        If revision > 0 Then
            If Not table.DataBodyRange Is Nothing Then
                snapshotIds = table.ListColumns("ChangeId").DataBodyRange.Value2
                For snapshotIndex = 1 To table.ListRows.Count
                    If table.ListRows.Count = 1 Then
                        currentFound = (snapshotIds = previous("LastChangeId"))
                    ElseIf snapshotIds(snapshotIndex, 1) = previous("LastChangeId") Then
                        currentFound = True
                    End If
                    If currentFound Then
                        Exit For
                    End If
                Next snapshotIndex
            End If
            If Not currentFound Then
                ex_PEB_MovementLinks.fn_Fail "MissingEventRevision", previous("LastChangeId")
            End If
        End If
    End If
    If values.Exists("RequestedStatus") Then
        status = values("RequestedStatus")
    End If
    If status <> "Active" And status <> "Cancelled" Then
        ex_PEB_MovementLinks.fn_Fail "InvalidEventStatus", values("RowId")
    End If
    reason = VBA.vbNullString
    If values.Exists("ChangeReason") Then
        reason = values("ChangeReason")
    End If
    If Not previous Is Nothing Then
        If status <> previous("Status") Then
            If VBA.Len(VBA.Trim$(reason)) = 0 Then
                Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryReasonRequired")
            End If
            changed = True
        End If
        If revision = 0 Then
            revision = 1
            baseline = private_Append(source, table, rows, previous, revision, "Baseline", _
                ex_PEB_Export.fn_Config("text.HistoryBaselineReason"), previous("Status"))
        End If
    End If
    For Each key In VBA.Split(ex_PEB_Export.fn_Config("history.versions.stateFields"), "|")
        Set column = Nothing
        On Error Resume Next
        Set column = source.ListColumns(VBA.CStr(key))
        On Error GoTo 0
        If column Is Nothing Then
            ex_PEB_MovementLinks.fn_Fail "MissingHistoryColumn", VBA.CStr(key)
        End If
    Next key
    If previous Is Nothing Then
        changeType = "Create"
    ElseIf status <> previous("Status") Then
        changeType = "Restore"
        If status = "Cancelled" Then
            changeType = "Cancel"
        End If
    Else
        changeType = "Update"
    End If
    values("Revision") = VBA.CStr(revision)
    values("Status") = status
    If Not previous Is Nothing Then
        values("LastChangeId") = previous("LastChangeId")
    End If
    If VBA.Len(baseline) > 0 Then
        values("LastChangeId") = baseline
    End If
    If previous Is Nothing Or changed Then
        revision = revision + 1
        values("Revision") = VBA.CStr(revision)
        If nextRows Is Nothing Then Set nextRows = rows
        values("LastChangeId") = private_Append(source, table, nextRows, values, revision, changeType, reason, status, previous)
    End If
End Sub

Public Function fn_Rollback( _
    ByVal transaction As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim table As ListObject

    On Error GoTo EH_ROLLBACK
    If transaction Is Nothing Then
        fn_Rollback = True
        Exit Function
    End If
    If Not transaction.Exists("Table") Then
        fn_Rollback = True
        Exit Function
    End If
    Set table = transaction("Table")
    Do While table.ListRows.Count > transaction("OriginalRows")
        table.ListRows(table.ListRows.Count).Delete
    Loop
    fn_Rollback = True
    Exit Function
EH_ROLLBACK:
    diagnostic = VBA.Err.Description
End Function

Public Function fn_Show( _
    ByVal identifier As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim book As Workbook
    Dim table As ListObject
    Dim candidate As Workbook
    Dim path As String

    On Error GoTo EH
    path = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("history.path"))
    For Each candidate In Application.Workbooks
        If VBA.StrComp(candidate.FullName, path, VBA.vbTextCompare) = 0 Then
            Set book = candidate
        End If
    Next candidate
    If book Is Nothing Then
        Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryOpenRequired"), "{value}", path)
    End If
    Set table = book.Worksheets(ex_PEB_Export.fn_Config("history.versions.sheet")).ListObjects( _
        ex_PEB_Export.fn_Config("history.versions.table"))
    If table.AutoFilter.FilterMode Then
        table.AutoFilter.ShowAllData
    End If
    table.Range.AutoFilter Field:=table.ListColumns("EventRowId").Index, Criteria1:=identifier
    book.Activate
    table.Parent.Activate
    fn_Show = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
End Function

Public Sub fn_ReadSync( _
    ByVal source As ListObject, _
    ByVal rows As Collection _
)
    Dim table As ListObject
    Dim row As Object
    Dim latest As Object
    Dim data As Variant
    Dim exporter As Variant
    Dim entry As Variant
    Dim index As Long
    Dim key As String
    Dim text As String
    Dim sheetName As String
    Dim tableName As String
    Dim headers As Variant

    Set latest = VBA.CreateObject("Scripting.Dictionary")
    sheetName = ex_PEB_Export.fn_Config("history.exports.sheet")
    tableName = ex_PEB_Export.fn_Config("history.exports.table")
    headers = VBA.Split(ex_PEB_Export.fn_Config("history.exports.fields"), "|")
    On Error Resume Next
    Set table = source.Parent.Parent.Worksheets(sheetName).ListObjects(tableName)
    On Error GoTo 0
    If Not table Is Nothing Then
        If table.ListColumns.Count <> UBound(headers) + 1 Then
            ex_PEB_MovementLinks.fn_Fail "InvalidRevisionSchema", table.Name
        End If
        For index = 0 To UBound(headers)
            If table.ListColumns(index + 1).Name <> headers(index) Then
                ex_PEB_MovementLinks.fn_Fail "InvalidRevisionSchema", table.Name
            End If
        Next index
        If Not table.DataBodyRange Is Nothing Then
            data = table.DataBodyRange.Value2
            For index = 1 To UBound(data, 1)
                key = VBA.Len(VBA.CStr(data(index, 1))) & ":" & data(index, 1) & data(index, 5)
                latest(key) = VBA.Array(VBA.CStr(data(index, 2)), VBA.CStr(data(index, 7)))
            Next index
        End If
    End If
    For Each row In rows
        For Each exporter In VBA.Array("Movement", "Word")
            text = ex_PEB_Export.fn_Config("text.HistorySyncUnknown")
            If row("Status") = "Cancelled" Then
                text = ex_PEB_Export.fn_Config("text.HistoryCancelPending")
            End If
            key = VBA.Len(VBA.CStr(row("RowId"))) & ":" & row("RowId") & exporter
            If latest.Exists(key) Then
                entry = latest(key)
                text = ex_PEB_Export.fn_Config("text.HistorySyncPending")
                If row("Status") = "Cancelled" Then
                    text = ex_PEB_Export.fn_Config("text.HistoryCancelPending")
                End If
                If entry(0) = row("Revision") Then
                    text = ex_PEB_Export.fn_Config("text.HistorySyncConfirmed")
                End If
            End If
            row(exporter & "Sync") = text
        Next exporter
    Next row
End Sub

Public Function fn_RecordExport( _
    ByVal actions As Collection, _
    ByVal target As String, _
    ByVal exporter As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim book As Workbook
    Dim candidate As Workbook
    Dim table As ListObject
    Dim row As ListRow
    Dim action As Object
    Dim data() As Variant
    Dim count As Long
    Dim committed As Boolean
    Dim path As String
    Dim previousBook As Workbook
    Dim previousSheet As Object
    Dim fieldTable As ListObject
    Dim conflictTable As ListObject
    Dim fieldCount As Long
    Dim conflictCount As Long
    Dim pending As New Collection
    Dim startedAt As Double

    On Error GoTo EH
    startedAt = VBA.Timer
    Set previousBook = Application.ActiveWorkbook
    Set previousSheet = Application.ActiveSheet
    path = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("history.path"))
    For Each candidate In Application.Workbooks
        If VBA.StrComp(candidate.FullName, path, VBA.vbTextCompare) = 0 Then
            Set book = candidate
        End If
    Next candidate
    If book Is Nothing Then
        Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryOpenRequired"), "{value}", path)
    End If
    If book.ReadOnly Then
        Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError015"), "{value}", path)
    End If
    Set table = private_Table(book, "history.exports", Nothing)
    count = table.ListRows.Count
    If exporter = "Movement" Then
        Set fieldTable = ex_PEB_MovementConflicts.fn_Table(book, "history.fieldState")
        fieldCount = fieldTable.ListRows.Count
        Set conflictTable = ex_PEB_MovementConflicts.fn_Table(book, "history.conflicts")
        conflictCount = conflictTable.ListRows.Count
        ex_PEB_MovementConflicts.fn_Record book, actions, target
    End If
    For Each action In actions
        ReDim data(1 To 1, 1 To table.ListColumns.Count)
        data(1, 1) = action("RowId")
        data(1, 2) = action("Revision")
        data(1, 3) = VBA.Format$(VBA.Now, "yyyy-mm-dd hh:nn:ss")
        data(1, 4) = Application.UserName & " / " & VBA.Environ$("USERNAME")
        data(1, 5) = exporter
        data(1, 6) = target
        data(1, 7) = action("Status")
        If exporter = "Movement" Then
            data(1, 7) = "Applied"
            If action("Status") = "AlreadyApplied" Then
                data(1, 7) = "Confirmed"
            End If
            If action("EventStatus") = "Cancelled" Then
                data(1, 7) = "CancelConfirmed"
            End If
        End If
        data(1, 8) = action("HistoryStamp")
        pending.Add data
    Next action
    private_WriteRows table, pending
    ex_Core.fn_Diagnostic_WritePerf "Movement.Journal.PrepareAndWrite." & exporter, startedAt
    committed = True
    fn_RecordExport = True
Cleanup:
    On Error Resume Next
    If Not committed And Not fieldTable Is Nothing Then
        Do While fieldTable.ListRows.Count > fieldCount
            fieldTable.ListRows(fieldTable.ListRows.Count).Delete
        Loop
    End If
    If Not committed And Not conflictTable Is Nothing Then
        Do While conflictTable.ListRows.Count > conflictCount
            conflictTable.ListRows(conflictTable.ListRows.Count).Delete
        Loop
    End If
    If Not committed And Not table Is Nothing Then
        Do While table.ListRows.Count > count
            table.ListRows(table.ListRows.Count).Delete
        Loop
    End If
    If Not previousBook Is Nothing Then
        ex_Core.fn_Diagnostic_WriteLog "MOVEMENT_JOURNAL_FOCUS | Stage=BeforeRestore | Target=" & previousBook.Name & _
            " | Active=" & Application.ActiveWorkbook.Name & " | ScreenUpdating=" & CStr(Application.ScreenUpdating)
        If Not Application.ActiveWorkbook Is previousBook Then
            previousBook.Activate
        End If
        If Not Application.ActiveSheet Is previousSheet Then
            previousSheet.Activate
        End If
        ex_Core.fn_Diagnostic_WriteLog "MOVEMENT_JOURNAL_FOCUS | Stage=AfterRestore | Active=" & Application.ActiveWorkbook.Name
    End If
    On Error GoTo 0
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_Table( _
    ByVal book As Workbook, _
    ByVal prefix As String, _
    ByVal source As ListObject _
) As ListObject
    Dim sheet As Worksheet
    Dim table As ListObject
    Dim headers As Variant
    Dim index As Long
    Dim sheetName As String
    Dim tableName As String
    Dim field As Variant
    Dim mapping As Variant
    Dim column As ListColumn

    headers = VBA.Split(ex_PEB_Export.fn_Config(prefix & ".fields"), "|")
    sheetName = ex_PEB_Export.fn_Config(prefix & ".sheet")
    tableName = ex_PEB_Export.fn_Config(prefix & ".table")
    On Error Resume Next
    Set sheet = book.Worksheets(sheetName)
    If Not sheet Is Nothing Then
        Set table = sheet.ListObjects(tableName)
    End If
    On Error GoTo 0
    If sheet Is Nothing Then
        ex_PEB_MovementLinks.fn_Fail "MissingHistorySheet", sheetName
    End If
    If table Is Nothing Then
        ex_PEB_MovementLinks.fn_Fail "MissingHistoryTable", sheetName & "/" & tableName
    End If
    If table.Parent.ProtectContents Then
        ex_PEB_MovementLinks.fn_Fail "ProtectedRevisionTable", table.Name
    End If
    If source Is Nothing Then
        If table.ListColumns.Count <> UBound(headers) + 1 Then
            ex_PEB_MovementLinks.fn_Fail "InvalidRevisionSchema", table.Name
        End If
    End If
    For index = 0 To UBound(headers)
        If table.ListColumns(index + 1).Name <> headers(index) Then
            ex_PEB_MovementLinks.fn_Fail "InvalidRevisionSchema", table.Name
        End If
    Next index
    If Not source Is Nothing Then
        For Each field In VBA.Split(ex_PEB_Export.fn_Config("history.fields"), "|")
            mapping = VBA.Split(field, "=")
            Set column = Nothing
            On Error Resume Next
            Set column = table.ListColumns("Snapshot." & mapping(1))
            On Error GoTo 0
            If column Is Nothing And mapping(0) = "ExtensionRole" Then
                Set column = table.ListColumns.Add
                column.Name = "Snapshot." & mapping(1)
            End If
            If column Is Nothing Then
                ex_PEB_MovementLinks.fn_Fail "MissingHistoryColumn", "Snapshot." & mapping(1)
            End If
        Next field
    End If
    Set private_Table = table
End Function

Private Function private_Append( _
    ByVal source As ListObject, _
    ByVal table As ListObject, _
    ByVal rows As Collection, _
    ByVal root As Object, _
    ByVal revision As Long, _
    ByVal changeType As String, _
    ByVal reason As String, _
    ByVal status As String, _
    Optional ByVal previous As Object = Nothing _
) As String
    Dim row As Object
    Dim parent As Object
    Dim ids As Object
    Dim seen As Object
    Dim target As ListRow
    Dim data() As Variant
    Dim snapshot As Object
    Dim field As Variant
    Dim mapping As Variant
    Dim changedFields As String
    Dim identifier As String
    Dim items As Collection
    Dim identifiers As Variant
    Dim index As Long
    Dim capturedAt As String
    Dim actor As String

    identifier = root("RowId") & "/R" & VBA.CStr(revision)
    capturedAt = VBA.Format$(VBA.Now, "yyyy-mm-dd hh:nn:ss")
    actor = Application.UserName & " / " & VBA.Environ$("USERNAME")
    If Not table.DataBodyRange Is Nothing Then
        identifiers = table.ListColumns("ChangeId").DataBodyRange.Value2
        For index = 1 To table.ListRows.Count
            If table.ListRows.Count = 1 Then
                If identifiers = identifier Then
                    ex_PEB_MovementLinks.fn_Fail "DuplicateEventRevision", identifier
                End If
            ElseIf identifiers(index, 1) = identifier Then
                ex_PEB_MovementLinks.fn_Fail "DuplicateEventRevision", identifier
            End If
        Next index
    End If
    Set ids = VBA.CreateObject("Scripting.Dictionary")
    For Each row In rows
        ids.Add row("RowId"), row
    Next row
    Set items = New Collection
    items.Add root
    For Each row In rows
        If row("RowType") <> "Event" Then
            Set parent = row
            Set seen = VBA.CreateObject("Scripting.Dictionary")
            Do While parent("RowType") <> "Event"
                If seen.Exists(parent("RowId")) Or Not ids.Exists(parent("ParentRowId")) Then
                    ex_PEB_MovementLinks.fn_Fail "InvalidRevisionChildren", row("RowId")
                End If
                seen.Add parent("RowId"), True
                Set parent = ids(parent("ParentRowId"))
            Loop
            If parent("RowId") = root("RowId") Then
                items.Add row
            End If
        End If
    Next row
    If root.Exists("ChildrenChanged") Then changedFields = "Children"
    For Each field In VBA.Split(ex_PEB_Export.fn_Config("history.fields"), "|")
        mapping = VBA.Split(field, "=")
        If Not previous Is Nothing Then
            If ex_PEB_Export.fn_Text(root(mapping(0))) <> ex_PEB_Export.fn_Text(previous(mapping(0))) Then
                If VBA.Len(changedFields) > 0 Then
                    changedFields = changedFields & "|"
                End If
                changedFields = changedFields & mapping(0)
            End If
        End If
    Next field
    If Not previous Is Nothing Then
        If previous("Status") <> status Then
            If VBA.Len(changedFields) > 0 Then
                changedFields = changedFields & "|"
            End If
            changedFields = changedFields & "Status"
        End If
    End If
    For Each snapshot In items
        ReDim data(1 To 1, 1 To table.ListColumns.Count)
        data(1, 1) = identifier
        data(1, 2) = root("RowId")
        data(1, 3) = VBA.CStr(revision)
        data(1, 4) = capturedAt
        data(1, 5) = actor
        data(1, 6) = changeType
        data(1, 7) = reason
        data(1, 8) = changedFields
        data(1, 9) = status
        For Each field In VBA.Split(ex_PEB_Export.fn_Config("history.fields"), "|")
            mapping = VBA.Split(field, "=")
            data(1, table.ListColumns("Snapshot." & mapping(1)).Index) = ex_PEB_Export.fn_Text(snapshot(mapping(0)))
        Next field
        Set target = table.ListRows.Add
        target.Range.NumberFormat = "@"
        ex_PEB_History.fn_NormalizeEmptyCells data
        target.Range.Value2 = data
        ex_PEB_History.fn_ClearEmptyCells target.Range, data
    Next snapshot
    private_Append = identifier
End Function
' --------------------------------------

Private Sub private_WriteRows( _
    ByVal table As ListObject, _
    ByVal rows As Collection _
)
    Dim data() As Variant
    Dim item As Variant
    Dim area As Range
    Dim first As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim extension As Range
    Dim other As ListObject
    Dim canResize As Boolean

    If rows.Count = 0 Then
        Exit Sub
    End If
    first = table.ListRows.Count + 1
    ReDim data(1 To rows.Count, 1 To table.ListColumns.Count)
    For rowIndex = 1 To rows.Count
        item = rows(rowIndex)
        For columnIndex = 1 To table.ListColumns.Count
            data(rowIndex, columnIndex) = item(1, columnIndex)
        Next columnIndex
    Next rowIndex
    ' Resize once when the extension is empty; retain Excel insertion semantics otherwise.
    canResize = Not table.ShowTotals
    If table.Range.Row + table.Range.Rows.Count + rows.Count - 1 > table.Parent.Rows.Count Then
        canResize = False
    End If
    If canResize Then
        Set extension = table.Range.Rows(table.Range.Rows.Count).Offset(1).Resize(rows.Count)
        canResize = (Application.CountA(extension) = 0 And Not VBA.IsNull(extension.MergeCells))
        If canResize Then
            canResize = Not extension.MergeCells
        End If
        For Each other In table.Parent.ListObjects
            If Not other Is table Then
                If Not Application.Intersect(extension, other.Range) Is Nothing Then
                    canResize = False
                End If
            End If
        Next other
    End If
    If canResize Then
        table.Resize table.Range.Resize(table.Range.Rows.Count + rows.Count)
    Else
        For rowIndex = 1 To rows.Count
            table.ListRows.Add
        Next rowIndex
    End If
    Set area = table.DataBodyRange.Rows(first).Resize(rows.Count)
    area.NumberFormat = "@"
    ex_PEB_History.fn_NormalizeEmptyCells data
    area.Value2 = data
    ex_PEB_History.fn_ClearEmptyCells area, data
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------