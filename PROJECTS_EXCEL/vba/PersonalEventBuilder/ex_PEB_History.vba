Attribute VB_Name = "ex_PEB_History"
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
Public Function fn_Save( _
    ByVal values As Object, _
    ByRef diagnostic As String, _
    Optional ByVal rebindFormExtensions As Boolean = False _
) As Boolean
    Dim book As Workbook
    Dim table As ListObject
    Dim columns As Object
    Dim rows As Collection
    Dim row As Object
    Dim mapping As Variant
    Dim field As Variant
    Dim target As ListRow
    Dim backup As Variant
    Dim data() As Variant
    Dim added As Boolean
    Dim changed As Boolean
    Dim committed As Boolean
    Dim index As Long
    Dim matchIndex As Long
    Dim identifier As String
    Dim key As String
    Dim value As String
    Dim originalColumns As Long
    Dim collectionName As Variant
    Dim transaction As Object
    Dim previous As Object
    Dim current As Object
    Dim stateField As Variant
    Dim written As Boolean
    Dim backupFormats() As Variant
    Dim snapshotStarted As Boolean
    Dim rollbackDiagnostic As String
    Dim previousBook As Workbook
    Dim previousSheet As Object
    Dim needsValidation As Boolean
    Dim historyIndex As Object
    Dim recoveryToken As String
    Dim tvoService As obj_PEB_ExportService
    Dim hasExplicitTvo As Boolean
    Dim reuseTvo As Boolean
    Dim restoredTvo As Collection
    Dim childPlan As Object
    Dim revisionRows As Collection

    On Error GoTo EH
    Set previousBook = Application.ActiveWorkbook
    Set previousSheet = Application.ActiveSheet
    ex_PEB_Export.fn_BeginConfigSnapshot
    snapshotStarted = True
    private_Open book, table
    originalColumns = table.ListColumns.Count
    For Each collectionName In VBA.Array("MetaDocumentTables", "MetaTvoTables", "HistoryReports")
        If values.Exists(collectionName) Then
            If Not values.Exists("FormExtensionRows") Then
                If values(collectionName).Count > 0 Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryChildrenInTable")
            End If
        End If
    Next collectionName
    Set columns = private_Columns(table, True)
    Set historyIndex = VBA.CreateObject("Scripting.Dictionary")
    Set rows = private_Read(table, columns, historyIndex)
    ex_PEB_FormChildren.fn_Attach rows
    key = private_Identity(values)
    For index = 1 To rows.Count
        Set row = rows(index)
        If row("RowType") = "Event" Then
            If private_Identity(row) = key Then
                If matchIndex <> 0 Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError001")
                matchIndex = index
                identifier = row("RowId")
            End If
        End If
    Next index
    If values.Exists("RowId") Then
        If VBA.Len(VBA.CStr(values("RowId"))) > 0 Then
            If matchIndex > 0 And identifier <> VBA.CStr(values("RowId")) Then
                Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryIdentityConflict")
            End If
            matchIndex = 0
            identifier = VBA.CStr(values("RowId"))
            For index = 1 To rows.Count
                If rows(index)("RowId") = identifier Then
                    If rows(index)("RowType") <> "Event" Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError002")
                    matchIndex = index
                End If
            Next index
            If matchIndex = 0 Then Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError003"), "{value}", VBA.CStr(identifier))
            If values.Exists("ExpectedHistoryStamp") Then
                If fn_Stamp(rows(matchIndex)) <> VBA.CStr(values("ExpectedHistoryStamp")) Then
                    Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryStaleEdit")
                End If
            End If
        End If
    End If
    If matchIndex = 0 Then
        identifier = "E-" & VBA.Format$(VBA.Now, "yyyymmddhhnnss") & "-" & VBA.CStr(table.ListRows.Count + 1)
        For Each row In rows
            If row("RowId") = identifier Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError004")
        Next row
    End If
    values("RowId") = identifier
    values("ParentRowId") = VBA.vbNullString
    values("RowType") = "Event"
    If matchIndex = 0 Then
        values("RowOrder") = VBA.CStr(table.ListRows.Count + 1)
    Else
        values("RowOrder") = rows(matchIndex)("RowOrder")
    End If
    If Not values.Exists("RequestedStatus") Then
        values("EventLabel") = ex_PEB_EventRoutes.fn_Value(VBA.CStr(values("EventId")), "caption")
        values("WordTemplatePath") = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("export.Word.documentTemplate"))
    End If
    If values.Exists("FormExtensionRows") Then
        For Each row In values("FormExtensionRows")
            If row("RowType") = "TVO" Then hasExplicitTvo = True
        Next row
    End If
    ' Capture return duties once; explicit history children always take priority.
    If Not values.Exists("RequestedStatus") Then
        values("TvoReturnSnapshot") = VBA.vbNullString
        If ex_PEB_EventRoutes.fn_Value(values("EventId"), "direction") = "Arrival" Then
            For Each row In rows
                If row("RowType") = "TVO" And row("ParentRowId") = identifier Then hasExplicitTvo = True
            Next row
            If Not hasExplicitTvo Then
                If matchIndex > 0 Then
                    Set previous = rows(matchIndex)
                    reuseTvo = (previous("PersonId") = private_Text(values("PersonId")) And _
                        previous("EventId") = private_Text(values("EventId")) And _
                        previous("EventDate") = private_Text(values("EventDate")))
                    If reuseTvo Then
                        reuseTvo = (VBA.Len(previous("TvoReturnSnapshot")) > 0)
                        If reuseTvo Then values("TvoReturnSnapshot") = previous("TvoReturnSnapshot")
                        If reuseTvo Then
                            Set restoredTvo = ex_PEB_ReturnTvo.fn_Decode(values)
                            reuseTvo = (restoredTvo.Count > 0)
                        End If
                    End If
                End If
                If Not reuseTvo Then
                    Set tvoService = New obj_PEB_ExportService
                    If Not tvoService.Initialize() Then Err.Raise 5, , "Cannot initialize return TVO snapshot service."
                    If Not tvoService.TryCaptureReturnTvo(values, diagnostic) Then GoTo Cleanup
                    tvoService.Dispose
                    Set tvoService = Nothing
                End If
                private_AddReturnTvoChildren values, identifier
            End If
        End If
    End If
    ReDim data(1 To 1, 1 To table.ListColumns.Count)
    If matchIndex > 0 Then
        Set target = table.ListRows(rows(matchIndex)("__tableIndex"))
        backup = target.Range.Value2
        ReDim backupFormats(1 To table.ListColumns.Count)
        For index = 1 To table.ListColumns.Count
            data(1, index) = backup(1, index)
            backupFormats(index) = target.Range.Cells(1, index).NumberFormat
        Next index
    End If
    For Each field In VBA.Split(ex_PEB_Export.fn_Config("history.fields"), "|")
        mapping = VBA.Split(VBA.CStr(field), "=")
        If values.Exists(mapping(0)) Then
            If Not VBA.IsObject(values(mapping(0))) Then
                value = private_Text(values(mapping(0)))
                data(1, columns(mapping(0))) = value
                If matchIndex > 0 Then
                    If value <> private_Text(backup(1, columns(mapping(0)))) Then changed = True
                End If
            End If
        End If
    Next field
    Set current = VBA.CreateObject("Scripting.Dictionary")
    For Each field In columns.Keys
        current(field) = private_Text(data(1, columns(field)))
    Next field
    If values.Exists("RequestedStatus") Then
        current("RequestedStatus") = values("RequestedStatus")
    End If
    If values.Exists("ChangeReason") Then
        current("ChangeReason") = values("ChangeReason")
    End If
    If matchIndex > 0 Then
        Set previous = rows(matchIndex)
    End If
    current("Revision") = "0"
    current("Status") = "Active"
    current("LastChangeId") = VBA.vbNullString
    If Not previous Is Nothing Then
        For Each stateField In VBA.Split(ex_PEB_Export.fn_Config("history.versions.stateFields"), "|")
            current(stateField) = previous(stateField)
        Next stateField
    End If
    If current.Exists("RequestedStatus") Then
        current("Status") = current("RequestedStatus")
    End If
    Set revisionRows = rows
    If values.Exists("FormExtensionRows") Then
        Set childPlan = ex_PEB_FormChildren.fn_Plan(table, rows, current, values("FormExtensionRows"), rebindFormExtensions)
        If Not childPlan Is Nothing Then
            changed = True
            current("ChildrenChanged") = True
            Set revisionRows = childPlan("Rows")
        End If
    End If
    needsValidation = changed Or Not values.Exists("RequestedStatus")
    If previous Is Nothing Then
        needsValidation = True
    ElseIf current("Status") <> previous("Status") Then
        needsValidation = True
    End If
    If needsValidation Then
        If Not ex_PEB_HistoryEvntVldtn.fn_Validate(current, previous, diagnostic, historyIndex) Then
            GoTo Cleanup
        End If
    End If
    If matchIndex > 0 And Not changed And Not values.Exists("RequestedStatus") Then
        If VBA.CLng(rows(matchIndex)("Revision")) > 0 Then
            values("HistoryAction") = "SKIP"
            For Each stateField In VBA.Split(ex_PEB_Export.fn_Config("history.versions.stateFields"), "|")
                values(stateField) = rows(matchIndex)(stateField)
            Next stateField
            fn_Save = True
            GoTo Cleanup
        End If
    End If
    ex_PEB_HistoryRevisions.fn_Prepare table, rows, previous, current, changed, transaction, revisionRows
    If Not childPlan Is Nothing Then transaction.Add "Children", childPlan
    ReDim Preserve data(1 To 1, 1 To table.ListColumns.Count)
    For Each stateField In VBA.Split(ex_PEB_Export.fn_Config("history.versions.stateFields"), "|")
        data(1, table.ListColumns(VBA.CStr(stateField)).Index) = current(stateField)
        values(stateField) = current(stateField)
    Next stateField
    recoveryToken = ex_PEB_HistoryRecovery.fn_Prepare(table, backup, data, transaction)
    If matchIndex = 0 Then
        Set target = table.ListRows.Add
        added = True
    End If
    written = True
    target.Range.NumberFormat = "@"
    target.Range.Value2 = data
    ex_PEB_FormChildren.fn_Write table, childPlan
    ex_PEB_HistoryRecovery.fn_Mark book, recoveryToken, "COMMIT"
    committed = True
    values("HistoryAction") = "INSERT"
    If Not added Then values("HistoryAction") = "UPDATE"
    fn_Save = True
Cleanup:
    On Error Resume Next
    VBA.Err.Clear
    If Not committed And Not childPlan Is Nothing Then
        ex_PEB_FormChildren.fn_Rollback table, childPlan
        Set target = Nothing
    End If
    If Not committed And Not target Is Nothing Then
        If added Then
            target.Delete
        ElseIf written Then
            target.Range.Cells(1, 1).Resize(1, UBound(backup, 2)).Value2 = backup
            For index = 1 To UBound(backupFormats)
                target.Range.Cells(1, index).NumberFormat = backupFormats(index)
            Next index
        End If
    End If
    If VBA.Err.Number <> 0 Then
        diagnostic = diagnostic & VBA.vbCrLf & VBA.Replace$( _
            ex_PEB_Export.fn_Config("text.HistoryRollbackError"), "{details}", VBA.Err.Description)
    End If
    If Not committed Then
        If Not ex_PEB_HistoryRevisions.fn_Rollback(transaction, rollbackDiagnostic) Then
            diagnostic = diagnostic & VBA.vbCrLf & VBA.Replace$( _
                ex_PEB_Export.fn_Config("text.HistoryRollbackError"), "{details}", rollbackDiagnostic)
        End If
    End If
    VBA.Err.Clear
    If Not committed And originalColumns > 0 Then
        For index = table.ListColumns.Count To originalColumns + 1 Step -1
            table.ListColumns(index).Delete
        Next index
    End If
    If VBA.Err.Number <> 0 Then
        diagnostic = diagnostic & VBA.vbCrLf & VBA.Replace$( _
            ex_PEB_Export.fn_Config("text.HistoryRollbackError"), "{details}", VBA.Err.Description)
    End If
    If Not previousBook Is Nothing Then
        previousBook.Activate
        previousSheet.Activate
    End If
    If Not tvoService Is Nothing Then tvoService.Dispose
    If snapshotStarted Then
        ex_PEB_Export.fn_EndConfigSnapshot
    End If
    On Error GoTo 0
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function

Public Function fn_EnsureReturnTvo( _
    ByVal orderNumber As String, _
    ByVal orderDate As Variant, _
    ByRef diagnostic As String _
) As Boolean
    Dim events As Collection
    Dim row As Object
    Dim child As Object
    Dim hasTvo As Boolean

    If Not fn_ReadOrder(orderNumber, orderDate, events, diagnostic) Then Exit Function
    For Each row In events
        If row("Status") = "Active" And ex_PEB_EventRoutes.fn_Value(row("EventId"), "direction") = "Arrival" Then
            hasTvo = False
            For Each child In row("FormExtensionRows")
                If child("RowType") = "TVO" Then hasTvo = True
            Next child
            If Not hasTvo Then
                row("ExpectedHistoryStamp") = fn_Stamp(row)
                If Not fn_Save(row, diagnostic) Then Exit Function
            End If
        End If
    Next row
    fn_EnsureReturnTvo = True
End Function

Public Function fn_ReadOrder( _
    ByVal orderNumber As String, _
    ByVal orderDate As Variant, _
    ByRef events As Collection, _
    ByRef diagnostic As String, _
    Optional ByVal includeChildren As Boolean = True, _
    Optional ByVal people As Object = Nothing, _
    Optional ByVal readSync As Boolean = True _
) As Boolean
    Dim book As Workbook
    Dim table As ListObject
    Dim columns As Object
    Dim rows As Collection
    Dim ids As Object
    Dim identities As Object
    Dim selected As Collection
    Dim row As Object
    Dim parent As Object
    Dim root As Object
    Dim item As Object
    Dim chain As Object
    Dim index As Long
    Dim parentId As String
    Dim typeName As String
    Dim dateText As String
    Dim snapshotStarted As Boolean

    On Error GoTo EH
    ex_PEB_Export.fn_BeginConfigSnapshot
    snapshotStarted = True
    Set events = Nothing
    If people Is Nothing Then
        If VBA.Len(VBA.Trim$(orderNumber)) = 0 Or VBA.Len(ex_PEB_Export.fn_Text(orderDate)) = 0 Then
            Err.Raise 5, , ex_PEB_Export.fn_Config("text.OrderChooseRequired")
        End If
        dateText = private_Text(fn_Date(orderDate))
    End If
    private_Open book, table
    Set columns = private_Columns(table, False)
    Set rows = private_Read(table, columns, Nothing, readSync)
    Set ids = VBA.CreateObject("Scripting.Dictionary")
    Set identities = VBA.CreateObject("Scripting.Dictionary")
    Set selected = New Collection
    For Each row In rows
        If VBA.Len(row("RowId")) = 0 Or ids.Exists(row("RowId")) Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError005")
        ids.Add row("RowId"), row
        typeName = row("RowType")
        If typeName <> "Event" And typeName <> "Report" And typeName <> "Document" And typeName <> "TVO" Then
            Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError006"), "{value}", VBA.CStr(typeName))
        End If
        private_Order row
        If typeName = "Event" Then
            If VBA.Len(row("ParentRowId")) > 0 Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError007")
            If private_InScope(row, orderNumber, dateText, people) Then
                parentId = private_Identity(row)
                If identities.Exists(parentId) Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError008")
                identities.Add parentId, True
                private_AddSorted selected, row
            End If
            Set row("MetaDocumentTables") = New Collection
            Set row("MetaTvoTables") = New Collection
            Set row("HistoryReports") = New Collection
            Set row("FormExtensionRows") = New Collection
        End If
    Next row
    If Not includeChildren Then
        If selected.Count = 0 And people Is Nothing Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryEmptyOrder")
        Set events = selected
        fn_ReadOrder = True
        GoTo Cleanup
    End If
    For Each row In rows
        If row("RowType") <> "Event" Then
            Set chain = VBA.CreateObject("Scripting.Dictionary")
            Set root = row
            Do While root("RowType") <> "Event"
                If chain.Exists(root("RowId")) Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError009")
                chain.Add root("RowId"), True
                parentId = root("ParentRowId")
                If Not ids.Exists(parentId) Then Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError010"), "{value}", VBA.CStr(parentId))
                Set parent = ids(parentId)
                If parent("RowType") <> "Event" Then
                    If root("RowType") <> "TVO" Or parent("RowType") <> "TVO" Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError011")
                End If
                Set root = parent
            Loop
            If private_InScope(root, orderNumber, dateText, people) Then
                root("FormExtensionRows").Add row
                Select Case row("RowType")
                    Case "Document"
                        Set item = VBA.CreateObject("Scripting.Dictionary")
                        item("DocumentNote") = row("Document")
                        item("DocNo") = row("DocumentNumber")
                        item("DocDateShort") = row("DocumentDate")
                        item("RowOrder") = row("RowOrder")
                        private_AddSorted root("MetaDocumentTables"), item
                    Case "Report"
                        private_AddSorted root("HistoryReports"), row
                    Case "TVO"
                        Set parent = ids(row("ParentRowId"))
                        row("DutyPositionGenitive") = parent("PositionGenitive")
                        private_AddSorted root("MetaTvoTables"), row
                End Select
            End If
        End If
    Next row
    If selected.Count = 0 And people Is Nothing Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryEmptyOrder")
    For Each row In selected
        Set row("MetaTvoTables") = New Collection
        private_Tvo rows, row, row("MetaTvoTables")
        If row("MetaTvoTables").Count = 0 Then
            Set row("MetaTvoTables") = ex_PEB_ReturnTvo.fn_Decode(row)
        End If
    Next row
    Set events = selected
    fn_ReadOrder = True
Cleanup:
    On Error Resume Next
    If snapshotStarted Then
        ex_PEB_Export.fn_EndConfigSnapshot
    End If
    On Error GoTo 0
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function

Public Function fn_Stamp(ByVal values As Object) As String
    Dim field As Variant
    Dim mapping As Variant
    Dim text As String
    Dim child As Object

    For Each field In VBA.Split(ex_PEB_Export.fn_Config("history.fields"), "|")
        mapping = VBA.Split(VBA.CStr(field), "=")
        text = private_Text(ex_PEB_Export.fn_Required(values, VBA.CStr(mapping(0))))
        fn_Stamp = fn_Stamp & VBA.Len(text) & ":" & text
    Next field
    For Each field In VBA.Split(ex_PEB_Export.fn_Config("history.versions.stateFields"), "|")
        text = ex_PEB_Export.fn_Text(values(field))
        fn_Stamp = fn_Stamp & VBA.Len(text) & ":" & text
    Next field
    If values.Exists("FormExtensionRows") Then
        For Each child In values("FormExtensionRows")
            For Each field In Split(ex_PEB_Export.fn_Config("history.fields"), "|")
                mapping = Split(field, "=")
                text = private_Text(child(mapping(0)))
                fn_Stamp = fn_Stamp & Len(text) & ":" & text
            Next field
        Next child
    End If
End Function

Public Function fn_Date(ByVal value As Variant) As Date
    Dim parts As Variant
    Dim result As Date
    Dim text As String

    If VBA.VarType(value) = VBA.vbDate Then
        fn_Date = value
        Exit Function
    End If
    text = VBA.Trim$(VBA.CStr(value))
    If Not text Like "##.##.####" Then Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError012"), "{value}", VBA.CStr(text))
    parts = VBA.Split(text, ".")
    If VBA.CLng(parts(2)) < 1900 Then Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError013"), "{value}", text)
    result = VBA.DateSerial(VBA.CInt(parts(2)), VBA.CInt(parts(1)), VBA.CInt(parts(0)))
    If VBA.Format$(result, "dd.mm.yyyy") <> text Then Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError013"), "{value}", VBA.CStr(text))
    fn_Date = result
End Function
Public Function fn_SetStatus( _
    ByVal row As Object, _
    ByVal status As String, _
    ByVal reason As String, _
    ByVal expectedStamp As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim values As Object
    Dim key As Variant

    Set values = VBA.CreateObject("Scripting.Dictionary")
    For Each key In row.Keys
        If Not VBA.IsObject(row(key)) Then
            values(key) = row(key)
        End If
    Next key
    values("RequestedStatus") = status
    values("ChangeReason") = reason
    values("ExpectedHistoryStamp") = expectedStamp
    fn_SetStatus = fn_Save(values, diagnostic)
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Sub private_AddReturnTvoChildren( _
    ByVal values As Object, _
    ByVal parentId As String _
)
    Dim restored As Collection
    Dim children As Collection
    Dim source As Object
    Dim child As Object
    Dim field As Variant
    Dim key As Variant
    Dim mapping As Variant

    Set restored = ex_PEB_ReturnTvo.fn_Decode(values)
    If Not values.Exists("FormExtensionRows") Then Set values("FormExtensionRows") = New Collection
    Set children = values("FormExtensionRows")
    For Each source In restored
        Set child = VBA.CreateObject("Scripting.Dictionary")
        For Each field In VBA.Split(ex_PEB_Export.fn_Config("history.fields"), "|")
            mapping = VBA.Split(field, "=")
            child(mapping(0)) = VBA.vbNullString
        Next field
        For Each key In source.Keys
            child(key) = source(key)
        Next key
        child("RowType") = "TVO"
        child("ParentRowId") = parentId
        child("PersonPosition") = source("PositionGenitive")
        children.Add child
    Next source
    Set values("MetaTvoTables") = restored
End Sub

Private Sub private_Open( _
    ByRef book As Workbook, _
    ByRef table As ListObject _
)
    Dim candidate As Workbook
    Dim path As String

    path = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("history.path"))
    For Each candidate In Application.Workbooks
        If VBA.StrComp(candidate.FullName, path, VBA.vbTextCompare) = 0 Then Set book = candidate
    Next candidate
    If book Is Nothing Then
        If VBA.Len(VBA.Dir$(path)) = 0 Then Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError014"), "{value}", VBA.CStr(path))
        Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryOpenRequired"), "{value}", path)
    End If
    If book.ReadOnly Then Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError015"), "{value}", VBA.CStr(path))
    If book.Date1904 Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError016")
    Set table = book.Worksheets(ex_PEB_Export.fn_Config("history.sheet")).ListObjects(ex_PEB_Export.fn_Config("history.table"))
    If table.Parent.ProtectContents Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError017")
End Sub

Private Function private_Columns( _
    ByVal table As ListObject, _
    ByVal extend As Boolean _
) As Object
    Dim result As Object
    Dim field As Variant
    Dim mapping As Variant
    Dim column As ListColumn
    Dim found As Boolean

    Set result = VBA.CreateObject("Scripting.Dictionary")
    For Each field In VBA.Split(ex_PEB_Export.fn_Config("history.fields"), "|")
        mapping = VBA.Split(VBA.CStr(field), "=")
        If UBound(mapping) <> 1 Then Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError018"), "{value}", VBA.CStr(field))
        found = False
        For Each column In table.ListColumns
            If column.Name = mapping(1) Then
                result.Add mapping(0), column.Index
                found = True
                Exit For
            End If
        Next column
        If Not found And mapping(0) = "TvoReturnSnapshot" Then
            If extend Then
                Set column = table.ListColumns.Add
                column.Name = mapping(1)
                result.Add mapping(0), column.Index
            Else
                result.Add mapping(0), 0
            End If
            found = True
        End If
        If Not found Then Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError019"), "{value}", VBA.CStr(mapping(1)))
    Next field
    Set private_Columns = result
End Function

Private Function private_Read( _
    ByVal table As ListObject, _
    ByVal columns As Object, _
    Optional ByVal historyIndex As Object = Nothing, _
    Optional ByVal readSync As Boolean = True _
) As Collection
    Dim result As New Collection
    Dim row As Object
    Dim data As Variant
    Dim field As Variant
    Dim index As Long
    Dim hasData As Boolean
    Dim stateColumns As Object
    Dim column As ListColumn
    Dim key As Variant
    Dim byPerson As Object
    Dim byId As Object
    Dim bucket As Collection
    Dim stateFields As Variant
    Dim fieldKeys As Variant
    Dim startedAt As Double

    startedAt = VBA.Timer
    stateFields = VBA.Split(ex_PEB_Export.fn_Config("history.versions.stateFields"), "|")
    fieldKeys = columns.Keys
    If Not historyIndex Is Nothing Then
        Set byPerson = VBA.CreateObject("Scripting.Dictionary")
        byPerson.CompareMode = VBA.vbTextCompare
        Set byId = VBA.CreateObject("Scripting.Dictionary")
        historyIndex.Add "ByPerson", byPerson
        historyIndex.Add "ById", byId
    End If
    Set stateColumns = VBA.CreateObject("Scripting.Dictionary")
    For Each column In table.ListColumns
        For Each key In stateFields
            If column.Name = key Then
                stateColumns.Add key, column.Index
            End If
        Next key
    Next column
    If Not table.DataBodyRange Is Nothing Then
        data = table.DataBodyRange.Value2
        For index = 1 To UBound(data, 1)
            Set row = VBA.CreateObject("Scripting.Dictionary")
            hasData = False
            For Each field In fieldKeys
                If columns(field) = 0 Then
                    row.Add field, VBA.vbNullString
                Else
                    row.Add field, private_Text(data(index, columns(field)))
                End If
                If VBA.Len(row(field)) > 0 Then hasData = True
            Next field
            row.Add "__tableIndex", index
            ex_PEB_HistoryRevisions.fn_ReadState table, index, row, stateColumns, data, stateFields
            If hasData Then
                If VBA.Len(row("RowId")) = 0 Or VBA.Len(row("RowType")) = 0 Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError020")
                result.Add row
                If Not historyIndex Is Nothing Then
                    If byId.Exists(row("RowId")) Then
                        ex_PEB_MovementLinks.fn_Fail "DuplicateHistoryRowId", row("RowId")
                    End If
                    byId.Add row("RowId"), row
                    If row("RowType") = "Event" Then
                        If Not byPerson.Exists(row("PersonId")) Then
                            Set bucket = New Collection
                            byPerson.Add row("PersonId"), bucket
                        End If
                        byPerson(row("PersonId")).Add row
                    End If
                End If
            End If
        Next index
    End If
    ex_Core.fn_Diagnostic_WritePerf "History.Read.RowsAndState", startedAt
    If readSync Then
        startedAt = VBA.Timer
        ex_PEB_HistoryRevisions.fn_ReadSync table, result
        ex_Core.fn_Diagnostic_WritePerf "History.Read.Sync", startedAt
    End If
    Set private_Read = result
End Function

Private Function private_Text(ByVal value As Variant) As String
    If VBA.IsError(value) Or VBA.IsNull(value) Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError021")
    Select Case VBA.VarType(value)
        Case VBA.vbDate
            private_Text = VBA.Format$(value, "dd.mm.yyyy")
        Case VBA.vbBoolean
            If value Then
                private_Text = "true"
            Else
                private_Text = "false"
            End If
        Case Else
            private_Text = VBA.CStr(value)
    End Select
End Function

Private Function private_InScope( _
    ByVal row As Object, _
    ByVal orderNumber As String, _
    ByVal dateText As String, _
    ByVal people As Object _
) As Boolean
    If Not people Is Nothing Then
        private_InScope = people.Exists(ex_PEB_Export.fn_Text(row("PersonId")))
    ElseIf row("AcceptedOrderNumber") = orderNumber Then
        private_InScope = (private_Text(fn_Date(row("AcceptedOrderDate"))) = dateText)
    End If
End Function

Private Function private_Identity(ByVal values As Object) As String
    Dim field As Variant
    Dim value As String

    For Each field In VBA.Split(ex_PEB_Export.fn_Config("history.identityFields"), "|")
        value = private_Text(ex_PEB_Export.fn_Required(values, VBA.CStr(field)))
        If VBA.Len(value) = 0 Then Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError022"), "{value}", VBA.CStr(field))
        If field = "EventDate" Or field = "AcceptedOrderDate" Then value = private_Text(fn_Date(values(field)))
        private_Identity = private_Identity & VBA.Len(value) & ":" & value
    Next field
End Function

Private Sub private_Order(ByVal row As Object)
    Dim value As String

    value = VBA.CStr(row("RowOrder"))
    If VBA.Len(value) = 0 Or value Like "*[!0-9]*" Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError023")
    If VBA.CDbl(value) < 1 Or VBA.CDbl(value) > 2147483647# Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError024")
End Sub

Private Sub private_AddSorted( _
    ByVal rows As Collection, _
    ByVal row As Object _
)
    Dim index As Long

    For index = 1 To rows.Count
        If VBA.CLng(row("RowOrder")) < VBA.CLng(rows(index)("RowOrder")) Then
            rows.Add row, Before:=index
            Exit Sub
        End If
    Next index
    rows.Add row
End Sub

Private Sub private_Tvo( _
    ByVal rows As Collection, _
    ByVal parent As Object, _
    ByVal result As Collection _
)
    Dim children As New Collection
    Dim row As Object
    Dim field As Variant

    For Each row In rows
        If row("RowType") = "TVO" And row("ParentRowId") = parent("RowId") Then private_AddSorted children, row
    Next row
    For Each row In children
        row("DutyPositionGenitive") = VBA.vbNullString
        For Each field In VBA.Array("RankGenitive", "FIOAccusative", "PositionGenitive", "RankDative", "FIODative")
            If VBA.Len(row(field)) = 0 Then Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError025"), "{value}", VBA.CStr(field))
        Next field
        result.Add row
        private_Tvo rows, row, result
    Next row
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------