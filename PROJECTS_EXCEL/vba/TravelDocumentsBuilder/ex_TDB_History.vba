Option Explicit

Private Type TicketGuid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (ByRef value As TicketGuid) As Long
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" ( _
    ByRef value As TicketGuid, _
    ByVal buffer As LongPtr, _
    ByVal count As Long _
) As Long

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
Public Function fn_Config(ByVal suffix As String) As String
    fn_Config = ex_TravelDocumentsBuilder.fn_Config("history." & suffix)
End Function

Public Sub fn_Fail(ByVal key As String)
    VBA.Err.Raise VBA.vbObjectError + 2450, "TicketsHistory", fn_Config("text." & key)
End Sub

Public Function fn_Date(ByVal value As Variant) As Date
    Dim text As String
    Dim parts As Variant
    Dim parsed As Date

    On Error GoTo InvalidDate
    If VBA.VarType(value) = VBA.vbDate Then
        fn_Date = value
        Exit Function
    End If
    ' Existing serial dates remain readable until their event is saved as text.
    Select Case VBA.VarType(value)
        Case VBA.vbByte, VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
            If VBA.CDbl(value) < 2 Or VBA.CDbl(value) > 2958465# Then GoTo InvalidDate
            If VBA.CDbl(value) <> VBA.Fix(VBA.CDbl(value)) Then GoTo InvalidDate
            fn_Date = VBA.CDate(value)
            Exit Function
    End Select
    text = private_Text(value)
    If Not text Like "##.##.####" Then GoTo InvalidDate
    parts = VBA.Split(text, ".")
    If VBA.CLng(parts(2)) < 1900 Then GoTo InvalidDate
    parsed = VBA.DateSerial(VBA.CInt(parts(2)), VBA.CInt(parts(1)), VBA.CInt(parts(0)))
    If VBA.Format$(parsed, "dd.mm.yyyy") <> text Then GoTo InvalidDate
    fn_Date = parsed
    Exit Function
InvalidDate:
    VBA.Err.Raise VBA.vbObjectError + 2451, "TicketsHistory.Date", VBA.Replace$(fn_Config("text.InvalidDate"), "{value}", text)
End Function

Public Function fn_ReturnDate( _
    ByVal departure As Date, _
    ByVal vacationDays As Long, _
    ByVal roadDays As Long, _
    Optional ByVal donationDays As Long = 0 _
) As Date
    If vacationDays < 0 Or roadDays < 0 Or donationDays < 0 Then fn_Fail "InvalidRows"
    fn_ReturnDate = VBA.DateAdd("d", vacationDays + roadDays + donationDays, departure)
End Function

Public Sub fn_Open( _
    ByRef tickets As ListObject, _
    ByRef revisions As ListObject _
)
    Dim book As Workbook
    Dim path As String
    Dim fileSystem As Object

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    path = fn_Config("path")
    If fileSystem.GetDriveName(path) = VBA.vbNullString And VBA.Left$(path, 2) <> "\\" Then
        path = fileSystem.BuildPath(ThisWorkbook.Path, path)
    End If
    path = fileSystem.GetAbsolutePathName(path)
    For Each book In Application.Workbooks
        If VBA.StrComp(book.FullName, path, VBA.vbTextCompare) = 0 Then
            If book.ReadOnly Then fn_Fail "ReadOnly"
            Set tickets = private_GetTable(book, fn_Config("sheet"), fn_Config("table"))
            Set revisions = private_GetTable(book, fn_Config("versions.sheet"), fn_Config("versions.table"))
            If tickets.Parent.ProtectContents Or revisions.Parent.ProtectContents Then fn_Fail "ReadOnly"
            private_ValidateSchema tickets, fn_Config("columns")
            private_ValidateSchema revisions, fn_Config("versions.columns")
            Exit Sub
        End If
    Next book
    fn_Fail "OpenRequired"
End Sub

Public Function fn_NewRow() As Object
    Dim field As Variant
    Dim row As Object

    Set row = VBA.CreateObject("Scripting.Dictionary")
    For Each field In VBA.Split(fn_Config("columns"), "|")
        row.Add VBA.CStr(field), VBA.vbNullString
    Next field
    Set fn_NewRow = row
End Function

Public Function fn_NewId() As String
    Dim value As TicketGuid
    Dim identifier As String

    identifier = VBA.String$(39, VBA.vbNullChar)
    If CoCreateGuid(value) <> 0 Then fn_Fail "InvalidRows"
    If StringFromGUID2(value, VBA.StrPtr(identifier), 39) <> 39 Then fn_Fail "InvalidRows"
    fn_NewId = "T-" & VBA.Mid$(identifier, 2, 36)
End Function

Public Function fn_Actor( _
    ByVal lookup As obj_LookupService, _
    ByVal personId As String, _
    ByVal personName As String, _
    ByVal positionCode As String _
) As Object
    Dim result As obj_LookupResult
    Dim record As obj_LookupRecord
    Dim row As Object
    Dim mapping As Variant
    Dim pair As Variant
    Dim value As Variant
    Dim count As Long
    Dim diagnostic As String

    If VBA.Len(personId) = 0 Or VBA.Len(personName) = 0 Or VBA.Len(positionCode) = 0 Then fn_Fail "InvalidChain"
    If Not lookup.TrySearch(personId, result, diagnostic) Then VBA.Err.Raise 5, , diagnostic
    For Each record In result.Records
        If Not record.TryGetValue("Personnel.TaxId", value) Then fn_Fail "InvalidChain"
        If VBA.CStr(value) = personId Then
            count = count + 1
            Set row = fn_NewRow()
            For Each mapping In VBA.Split(fn_Config("actorFields"), "|")
                pair = VBA.Split(mapping, "=")
                If Not record.TryGetValue(VBA.CStr(pair(0)), value) Then fn_Fail "InvalidChain"
                row(pair(1)) = value
            Next mapping
        End If
    Next record
    If count <> 1 Or result.HasMore Then fn_Fail "InvalidChain"
    If VBA.StrComp(VBA.CStr(row(fn_Config("personNameField"))), personName, VBA.vbTextCompare) <> 0 Then fn_Fail "InvalidChain"
    If VBA.CStr(row("PositionCode")) <> positionCode Then fn_Fail "InvalidChain"
    Set fn_Actor = row
End Function

Public Function fn_Read(ByVal table As ListObject) As Object
    Dim rows As Object
    Dim row As Object
    Dim target As ListRow
    Dim field As Variant
    Dim identifier As String
    Dim hasData As Boolean

    Set rows = VBA.CreateObject("Scripting.Dictionary")
    For Each target In table.ListRows
        Set row = fn_NewRow()
        hasData = False
        For Each field In row.Keys
            row(field) = private_FieldText(VBA.CStr(field), target.Range.Cells(1, table.ListColumns(VBA.CStr(field)).Index).Value2)
            If VBA.IsError(row(field)) Then fn_Fail "InvalidRows"
            If VBA.Len(private_Text(row(field))) > 0 Then hasData = True
        Next field
        If hasData Then
            identifier = private_Text(row("RowId"))
            If VBA.Len(identifier) = 0 Or rows.Exists(identifier) Then fn_Fail "InvalidRows"
            rows.Add identifier, row
        End If
    Next target
    private_ValidateGraph rows
    Set fn_Read = rows
End Function

Public Function fn_Group( _
    ByVal rows As Object, _
    ByVal identifier As String _
) As Collection
    Dim result As New Collection
    Dim row As Object
    Dim key As Variant
    Dim parent As Object
    Dim ids As New Collection
    Dim index As Long
    Dim inserted As Boolean

    If Not rows.Exists(identifier) Then fn_Fail "NotFound"
    If rows(identifier)("RowType") <> "Event" Then fn_Fail "InvalidRows"
    result.Add rows(identifier)
    For Each key In rows.Keys
        Set row = rows(key)
        If row("RowType") <> "Event" Then
            Set parent = row
            Do While parent("RowType") <> "Event"
                Set parent = rows(parent("ParentRowId"))
            Loop
            If parent("RowId") = identifier Then
                inserted = False
                For index = 1 To ids.Count
                    If VBA.StrComp(VBA.CStr(key), VBA.CStr(ids(index)), VBA.vbBinaryCompare) < 0 Then
                        ids.Add key, Before:=index
                        inserted = True
                        Exit For
                    End If
                Next index
                If Not inserted Then ids.Add key
            End If
        End If
    Next key
    For Each key In ids
        result.Add rows(key)
    Next key
    Set fn_Group = result
End Function

Public Function fn_Stamp(ByVal rows As Collection) As String
    Dim row As Object
    Dim field As Variant
    Dim text As String
    Dim result As String

    For Each row In rows
        For Each field In VBA.Split(fn_Config("columns"), "|")
            If Not private_IsDocumentField(VBA.CStr(field)) Then
                text = private_Text(row(field))
                result = result & VBA.CStr(VBA.Len(text)) & ":" & text
            End If
        Next field
    Next row
    fn_Stamp = result
End Function

Public Function fn_Begin( _
    ByVal tickets As ListObject, _
    ByVal revisions As ListObject, _
    ByVal identifier As String, _
    ByVal expectedStamp As String _
) As Object
    Dim rows As Object
    Dim group As Collection
    Dim transaction As Object

    Set rows = fn_Read(tickets)
    private_CheckRoots revisions, rows
    If VBA.Len(identifier) > 0 Then
        Set group = fn_Group(rows, identifier)
        private_CheckRevision revisions, group
        If VBA.Len(expectedStamp) = 0 Or fn_Stamp(group) <> expectedStamp Then fn_Fail "Stale"
    Else
        Set group = New Collection
    End If
    Set transaction = VBA.CreateObject("Scripting.Dictionary")
    transaction.Add "Previous", group
    transaction.Add "TicketsCount", tickets.ListRows.Count
    transaction.Add "RevisionsCount", revisions.ListRows.Count
    transaction.Add "Stamp", private_TableStamp(tickets) & private_TableStamp(revisions)
    If tickets.ListRows.Count > 0 Then transaction.Add "TicketsData", tickets.DataBodyRange.Value2
    If revisions.ListRows.Count > 0 Then transaction.Add "RevisionsData", revisions.DataBodyRange.Value2
    transaction.Add "TicketsFormulas", private_Formulas(tickets)
    transaction.Add "RevisionsFormulas", private_Formulas(revisions)
    transaction.Add "TicketsFormats", private_Formats(tickets)
    transaction.Add "RevisionsFormats", private_Formats(revisions)
    Set fn_Begin = transaction
End Function

Public Sub fn_AssertCurrent( _
    ByVal tickets As ListObject, _
    ByVal revisions As ListObject, _
    ByVal transaction As Object _
)
    If private_TableStamp(tickets) & private_TableStamp(revisions) <> transaction("Stamp") Then fn_Fail "Stale"
    If tickets.Parent.ProtectContents Or revisions.Parent.ProtectContents Then fn_Fail "ReadOnly"
End Sub

Public Sub fn_Write( _
    ByVal table As ListObject, _
    ByVal target As ListRow, _
    ByVal row As Object _
)
    Dim field As Variant
    Dim cell As Range
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    Dim address As String

    On Error GoTo Failed
    target.Range.NumberFormat = "@"
    For Each field In row.Keys
        address = VBA.vbNullString
        Set cell = target.Range.Cells(1, table.ListColumns(VBA.CStr(field)).Index)
        address = cell.Address(External:=True)
        If VBA.Len(private_FieldText(VBA.CStr(field), row(field))) = 0 Then
            cell.ClearContents
        Else
            cell.Value2 = private_FieldText(VBA.CStr(field), row(field))
        End If
    Next field
    Exit Sub
Failed:
    errorNumber = VBA.Err.Number
    errorSource = VBA.Err.Source
    errorDescription = VBA.Err.Description
    ex_Core.fn_Diagnostic_WriteError "TDB.History.Write | Table=" & table.Name & " | Field=" & VBA.CStr(field) & " | Cell=" & address, _
        errorNumber, errorSource, errorDescription
    VBA.Err.Raise errorNumber, "TicketsHistory.Write/" & table.Name & "/" & VBA.CStr(field), errorDescription
End Sub

Public Function fn_Find( _
    ByVal table As ListObject, _
    ByVal identifier As String _
) As ListRow
    Dim row As ListRow

    For Each row In table.ListRows
        If private_Text(row.Range.Cells(1, table.ListColumns("RowId").Index).Value2) = identifier Then
            If Not fn_Find Is Nothing Then fn_Fail "InvalidRows"
            Set fn_Find = row
        End If
    Next row
End Function

Public Sub fn_Commit( _
    ByVal tickets As ListObject, _
    ByVal revisions As ListObject, _
    ByVal identifier As String, _
    ByVal transaction As Object, _
    ByVal reason As String _
)
    Dim rows As Object
    Dim group As Collection
    Dim previous As Collection
    Dim root As Object
    Dim row As Object
    Dim field As Variant
    Dim target As ListRow
    Dim revision As Long
    Dim changeType As String
    Dim changed As Boolean
    Dim changeId As String
    Dim snapshot As Object
    Dim versionRow As ListRow
    Dim capturedAt As String

    Set rows = fn_Read(tickets)
    Set group = fn_Group(rows, identifier)
    Set previous = transaction("Previous")
    Set root = group(1)
    changed = (private_ContentStamp(group) <> private_ContentStamp(previous))
    revision = 0
    changeType = "Create"
    If previous.Count > 0 Then
        revision = VBA.CLng(previous(1)("Revision"))
        changeType = "Update"
        If root("Status") <> previous(1)("Status") Then
            If VBA.Len(VBA.Trim$(reason)) = 0 Then fn_Fail "ReasonRequired"
            changeType = "Restore"
            If root("Status") = "Cancelled" Then changeType = "Cancel"
        End If
    End If
    If changed Then revision = revision + 1
    changeId = identifier & "/R" & VBA.CStr(revision)
    capturedAt = VBA.Format$(VBA.Now, "yyyy-mm-dd hh:nn:ss")
    For Each row In group
        Set target = fn_Find(tickets, VBA.CStr(row("RowId")))
        row("Revision") = revision
        row("Status") = root("Status")
        fn_Write tickets, target, row
    Next row
    If Not changed Then Exit Sub

    ' A revision owns a complete immutable snapshot, including every child.
    For Each versionRow In revisions.ListRows
        If private_Text(versionRow.Range.Cells(1, revisions.ListColumns("ChangeId").Index).Value2) = changeId Then fn_Fail "InvalidRevision"
    Next versionRow
    For Each row In group
        Set snapshot = VBA.CreateObject("Scripting.Dictionary")
        snapshot.Add "ChangeId", changeId
        snapshot.Add "EventRowId", identifier
        snapshot.Add "Revision", revision
        snapshot.Add "ChangedAt", capturedAt
        snapshot.Add "ChangedBy", Application.UserName & " / " & VBA.Environ$("USERNAME")
        snapshot.Add "ChangeType", changeType
        snapshot.Add "Reason", reason
        snapshot.Add "ChangedFields", private_ChangedFields(group, previous)
        snapshot.Add "Status", root("Status")
        For Each field In VBA.Split(fn_Config("snapshotFields"), "|")
            snapshot.Add "Snapshot." & field, row(field)
        Next field
        Set target = revisions.ListRows.Add
        fn_Write revisions, target, snapshot
    Next row
End Sub

Public Sub fn_Rollback( _
    ByVal tickets As ListObject, _
    ByVal revisions As ListObject, _
    ByVal transaction As Object _
)
    Dim errors As String

    On Error Resume Next
    private_Restore tickets, transaction, "Tickets"
    If VBA.Err.Number <> 0 Then errors = VBA.Err.Description
    VBA.Err.Clear
    private_Restore revisions, transaction, "Revisions"
    If VBA.Err.Number <> 0 Then errors = errors & VBA.vbLf & VBA.Err.Description
    On Error GoTo 0
    If VBA.Len(errors) > 0 Then VBA.Err.Raise 5, "TicketsHistory.Rollback", errors
End Sub

Public Function fn_Load( _
    ByVal identifier As String, _
    ByVal bindingContext As obj_UiBindingContext, _
    Optional ByVal expectedStamp As String = VBA.vbNullString _
) As Collection
    Dim tickets As ListObject
    Dim revisions As ListObject
    Dim rows As Object
    Dim group As Collection
    Dim root As Object
    Dim row As Object
    Dim mapping As Variant
    Dim pair As Variant
    Dim updates As Object
    Dim chain As New Collection
    Dim diagnostic As String
    Dim parentId As String
    Dim found As Boolean

    fn_Open tickets, revisions
    Set rows = fn_Read(tickets)
    Set group = fn_Group(rows, identifier)
    private_CheckRevision revisions, group
    If VBA.Len(expectedStamp) > 0 Then
        If fn_Stamp(group) <> expectedStamp Then fn_Fail "Stale"
    End If
    Set root = group(1)
    Set updates = VBA.CreateObject("Scripting.Dictionary")
    For Each mapping In VBA.Split(fn_Config("formFields"), "|")
        pair = VBA.Split(mapping, "=")
        updates.Add "Form." & pair(0), root(pair(1))
    Next mapping
    For Each mapping In VBA.Split(fn_Config("dateFormFields"), "|")
        pair = VBA.Split(mapping, "=")
        If VBA.Len(private_Text(root(pair(1)))) > 0 Then
            updates("Form." & pair(0)) = VBA.Format$(fn_Date(root(pair(1))), "dd.mm.yyyy")
        End If
    Next mapping
    updates("Form.OrderReference") = root(fn_Config("orderNumberField")) & " / " & VBA.Format$(fn_Date(root("OrderDate")), "dd.mm.yyyy")
    updates("Form.VacationStatus") = ex_TravelDocumentsBuilder.fn_Config("generation.status.active")
    If root("Status") = "Cancelled" Then updates("Form.VacationStatus") = ex_TravelDocumentsBuilder.fn_Config("generation.status.cancelled")
    updates("Form.DonationDays") = "0"
    For Each row In group
        If row("RowType") = "Donation" Then updates("Form.DonationDays") = row(fn_Config("durationField"))
    Next row
    updates("Form.HistoryRowId") = identifier
    updates("Form.HistoryStamp") = fn_Stamp(group)
    updates("Form.ChangeReason") = VBA.vbNullString
    updates("Form.TvoName") = VBA.vbNullString
    updates("Form.TvoIpn") = VBA.vbNullString
    updates("Form.TvoPosition") = VBA.vbNullString
    parentId = identifier
    Do
        found = False
        For Each row In group
            If row("RowType") = "TVO" And row("ParentRowId") = parentId Then
                chain.Add row
                parentId = row("RowId")
                found = True
                Exit For
            End If
        Next row
    Loop While found
    If Not bindingContext.TryApplyValues(updates, diagnostic) Then VBA.Err.Raise 5, , diagnostic
    If Not bindingContext.SetObject("Data", "TvoChain", chain) Then fn_Fail "InvalidRows"
    Set fn_Load = chain
End Function

Public Function fn_List( _
    ByVal orderNumber As String, _
    ByVal orderDate As Date _
) As obj_LookupResult
    Dim tickets As ListObject
    Dim revisions As ListObject
    Dim rows As Object
    Dim key As Variant
    Dim field As Variant
    Dim parts As Variant
    Dim row As Object
    Dim column As Object
    Dim columns As New Collection
    Dim records As New Collection
    Dim record As obj_LookupRecord
    Dim result As New obj_LookupResult
    Dim diagnostic As String
    Dim display As Object

    If VBA.Len(orderNumber) = 0 Or orderDate = 0 Then fn_Fail "OrderRequired"
    fn_Open tickets, revisions
    Set rows = fn_Read(tickets)
    private_CheckRoots revisions, rows
    For Each field In VBA.Split(fn_Config("listFields"), "|")
        parts = VBA.Split(field, "=")
        Set column = VBA.CreateObject("Scripting.Dictionary")
        column.Add "name", parts(0)
        column.Add "caption", fn_Config("text." & parts(1))
        columns.Add column
    Next field
    For Each key In rows.Keys
        Set row = rows(key)
        If row("RowType") = "Event" Then
            If private_Text(row(fn_Config("orderNumberField"))) = orderNumber And fn_Date(row("OrderDate")) = orderDate Then
                Set display = fn_NewRow()
                For Each field In row.Keys
                    display(field) = row(field)
                Next field
                For Each field In VBA.Split(fn_Config("dateFields"), "|")
                    If VBA.Len(private_Text(display(field))) > 0 Then display(field) = VBA.Format$(fn_Date(display(field)), "dd.mm.yyyy")
                Next field
                display.Add "HistoryStamp", fn_Stamp(fn_Group(rows, VBA.CStr(key)))
                Set record = New obj_LookupRecord
                If Not record.Initialize(VBA.CStr(key), display) Then fn_Fail "InvalidRows"
                records.Add record
            End If
        End If
    Next key
    If Not result.Initialize(records, columns, False, diagnostic) Then VBA.Err.Raise 5, , diagnostic
    Set fn_List = result
End Function

Public Function fn_OrderGroups( _
    ByVal tickets As ListObject, _
    ByVal revisions As ListObject, _
    ByVal orderNumber As String, _
    ByVal orderDate As Date _
) As Collection
    Dim rows As Object
    Dim key As Variant
    Dim row As Object
    Dim group As Collection
    Dim groups As New Collection

    If VBA.Len(orderNumber) = 0 Or orderDate = 0 Then fn_Fail "OrderRequired"
    Set rows = fn_Read(tickets)
    private_CheckRoots revisions, rows
    For Each key In rows.Keys
        Set row = rows(key)
        If row("RowType") = "Event" Then
            If private_Text(row(fn_Config("orderNumberField"))) = orderNumber And fn_Date(row("OrderDate")) = orderDate Then
                Set group = fn_Group(rows, VBA.CStr(key))
                private_CheckRevision revisions, group
                groups.Add group
            End If
        End If
    Next key
    Set fn_OrderGroups = groups
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_GetTable( _
    ByVal book As Workbook, _
    ByVal sheetName As String, _
    ByVal tableName As String _
) As ListObject
    On Error Resume Next
    Set private_GetTable = book.Worksheets(sheetName).ListObjects(tableName)
    On Error GoTo 0
    If private_GetTable Is Nothing Then fn_Fail "Schema"
End Function

Private Sub private_ValidateSchema( _
    ByVal table As ListObject, _
    ByVal fieldList As String _
)
    Dim field As Variant
    Dim column As ListColumn
    Dim fields As Variant

    fields = VBA.Split(fieldList, "|")
    If table.ListColumns.Count <> UBound(fields) + 1 Then fn_Fail "Schema"
    For Each field In fields
        Set column = table.ListColumns(VBA.CStr(field))
    Next field
End Sub

Private Function private_Text(ByVal value As Variant) As String
    If VBA.IsError(value) Or VBA.IsNull(value) Then fn_Fail "InvalidRows"
    If Not VBA.IsEmpty(value) Then private_Text = VBA.CStr(value)
End Function

Private Function private_FieldText( _
    ByVal field As String, _
    ByVal value As Variant _
) As String
    Dim text As String

    text = private_Text(value)
    If VBA.Len(text) > 0 Then
        field = VBA.Replace$(field, "Snapshot.", VBA.vbNullString)
        If VBA.InStr(1, "|" & fn_Config("dateFields") & "|", "|" & field & "|", VBA.vbBinaryCompare) > 0 Then
            text = VBA.Format$(fn_Date(value), "dd.mm.yyyy")
        End If
    End If
    private_FieldText = text
End Function

Private Sub private_ValidateGraph(ByVal rows As Object)
    Dim key As Variant
    Dim row As Object
    Dim parent As Object
    Dim seen As Object
    Dim links As Object
    Dim link As String

    Set links = VBA.CreateObject("Scripting.Dictionary")
    For Each key In rows.Keys
        Set row = rows(key)
        If row("RowType") <> "Event" And row("RowType") <> "Donation" And row("RowType") <> "TVO" Then fn_Fail "InvalidRows"
        If row("Status") <> "Active" And row("Status") <> "Cancelled" Then fn_Fail "InvalidRows"
        If Not VBA.IsNumeric(row("Revision")) Then fn_Fail "InvalidRevision"
        If VBA.CDbl(row("Revision")) < 1 Or VBA.CDbl(row("Revision")) >= 2147483647# Or _
                VBA.CDbl(row("Revision")) <> VBA.Fix(VBA.CDbl(row("Revision"))) Then fn_Fail "InvalidRevision"
        If row("RowType") = "Event" Then
            If VBA.Len(row("ParentRowId")) > 0 Then fn_Fail "InvalidRows"
        Else
            If Not rows.Exists(row("ParentRowId")) Then fn_Fail "InvalidRows"
            Set parent = rows(row("ParentRowId"))
            If row("RowType") = "Donation" And parent("RowType") <> "Event" Then fn_Fail "InvalidRows"
            If row("RowType") = "TVO" Then
                If parent("RowType") <> "Event" And parent("RowType") <> "TVO" Then fn_Fail "InvalidRows"
                If VBA.Len(row("PositionCode")) = 0 Or row("DutyPositionCode") <> parent("PositionCode") Then fn_Fail "InvalidChain"
            End If
            link = row("ParentRowId") & "/" & row("RowType")
            If links.Exists(link) Then fn_Fail "InvalidChain"
            links.Add link, True
            Set seen = VBA.CreateObject("Scripting.Dictionary")
            Set parent = row
            Do While parent("RowType") <> "Event"
                If seen.Exists(parent("RowId")) Then fn_Fail "InvalidChain"
                seen.Add parent("RowId"), True
                If Not rows.Exists(parent("ParentRowId")) Then fn_Fail "InvalidRows"
                Set parent = rows(parent("ParentRowId"))
                If row("RowType") = "TVO" Then
                    If VBA.StrComp(private_Text(row(fn_Config("personIdField"))), private_Text(parent(fn_Config("personIdField"))), VBA.vbTextCompare) = 0 Then fn_Fail "InvalidChain"
                    If VBA.StrComp(private_Text(row("PositionCode")), private_Text(parent("PositionCode")), VBA.vbTextCompare) = 0 Then fn_Fail "InvalidChain"
                End If
            Loop
            If row("Status") <> parent("Status") Then fn_Fail "InvalidRevision"
        End If
    Next key
End Sub

Private Sub private_CheckRevision( _
    ByVal versions As ListObject, _
    ByVal group As Collection _
)
    Dim root As Object
    Dim row As ListRow
    Dim snapshot As Object
    Dim snapshots As Object
    Dim live As Object
    Dim field As Variant
    Dim identifier As String
    Dim changeId As String

    Set root = group(1)
    changeId = root("RowId") & "/R" & root("Revision")
    Set snapshots = VBA.CreateObject("Scripting.Dictionary")
    For Each row In versions.ListRows
        If private_Text(row.Range.Cells(1, versions.ListColumns("ChangeId").Index).Value2) = changeId Then
            If private_Text(row.Range.Cells(1, versions.ListColumns("EventRowId").Index).Value2) <> root("RowId") Then fn_Fail "InvalidRevision"
            If private_Text(row.Range.Cells(1, versions.ListColumns("Revision").Index).Value2) <> private_Text(root("Revision")) Then fn_Fail "InvalidRevision"
            If private_Text(row.Range.Cells(1, versions.ListColumns("Status").Index).Value2) <> root("Status") Then fn_Fail "InvalidRevision"
            Set snapshot = fn_NewRow()
            For Each field In VBA.Split(fn_Config("snapshotFields"), "|")
                snapshot(field) = private_FieldText(VBA.CStr(field), row.Range.Cells(1, versions.ListColumns("Snapshot." & field).Index).Value2)
            Next field
            identifier = private_Text(snapshot("RowId"))
            If snapshots.Exists(identifier) Then fn_Fail "InvalidRevision"
            snapshots.Add identifier, snapshot
        End If
    Next row
    If snapshots.Count <> group.Count Then fn_Fail "ManualChanges"
    For Each live In group
        If Not snapshots.Exists(live("RowId")) Then fn_Fail "ManualChanges"
        Set snapshot = snapshots(live("RowId"))
        For Each field In VBA.Split(fn_Config("snapshotFields"), "|")
            If Not private_IsDocumentField(VBA.CStr(field)) Then
                If private_Text(live(field)) <> private_Text(snapshot(field)) Then fn_Fail "ManualChanges"
            End If
        Next field
        If private_Text(live("Revision")) <> private_Text(root("Revision")) Then fn_Fail "InvalidRevision"
    Next live
End Sub

Private Sub private_CheckRoots( _
    ByVal revisions As ListObject, _
    ByVal rows As Object _
)
    Dim row As ListRow
    Dim identifier As String

    For Each row In revisions.ListRows
        identifier = private_Text(row.Range.Cells(1, revisions.ListColumns("EventRowId").Index).Value2)
        If VBA.Len(identifier) > 0 Then
            If Not rows.Exists(identifier) Then fn_Fail "ManualChanges"
            If rows(identifier)("RowType") <> "Event" Then fn_Fail "InvalidRevision"
        End If
    Next row
End Sub

Private Function private_IsDocumentField(ByVal field As String) As Boolean
    private_IsDocumentField = (field = "DocumentPath" Or field = "DocumentRevision" Or field = "DocumentGeneratedAt")
End Function

Private Function private_ContentStamp(ByVal group As Collection) As String
    Dim row As Object
    Dim field As Variant
    Dim text As String
    Dim result As String

    For Each row In group
        For Each field In VBA.Split(fn_Config("snapshotFields") & "|Status", "|")
            If Not private_IsDocumentField(VBA.CStr(field)) Then
                text = private_Text(row(field))
                result = result & VBA.Len(text) & ":" & text
            End If
        Next field
    Next row
    private_ContentStamp = result
End Function

Private Function private_ChangedFields( _
    ByVal group As Collection, _
    ByVal previous As Collection _
) As String
    Dim field As Variant
    Dim result As String
    Dim different As Boolean

    For Each field In VBA.Split(fn_Config("snapshotFields") & "|Status", "|")
        different = (previous.Count = 0)
        If previous.Count > 0 Then different = (private_Text(group(1)(field)) <> private_Text(previous(1)(field)))
        If different And Not private_IsDocumentField(VBA.CStr(field)) Then result = result & "|" & field
    Next field
    If private_ChildrenStamp(group) <> private_ChildrenStamp(previous) Then result = result & "|Children"
    private_ChangedFields = VBA.Mid$(result, 2)
End Function

Private Function private_ChildrenStamp(ByVal group As Collection) As String
    Dim children As New Collection
    Dim index As Long

    For index = 2 To group.Count
        children.Add group(index)
    Next index
    private_ChildrenStamp = private_ContentStamp(children)
End Function

Private Function private_TableStamp(ByVal table As ListObject) As String
    Dim cell As Range
    Dim text As String
    Dim result As String

    result = table.Range.Address & ":"
    For Each cell In table.Range.Cells
        text = private_Text(cell.Formula)
        result = result & VBA.Len(text) & ":" & text
    Next cell
    private_TableStamp = result
End Function

Private Function private_Formats(ByVal table As ListObject) As Object
    Dim result As Object
    Dim cell As Range

    Set result = VBA.CreateObject("Scripting.Dictionary")
    If Not table.DataBodyRange Is Nothing Then
        For Each cell In table.DataBodyRange.Cells
            result.Add cell.Row & "/" & cell.Column, cell.NumberFormat
        Next cell
    End If
    Set private_Formats = result
End Function

Private Function private_Formulas(ByVal table As ListObject) As Object
    Dim result As Object
    Dim cell As Range

    Set result = VBA.CreateObject("Scripting.Dictionary")
    If Not table.DataBodyRange Is Nothing Then
        For Each cell In table.DataBodyRange.Cells
            If cell.HasFormula Then result.Add cell.Row & "/" & cell.Column, cell.Formula
        Next cell
    End If
    Set private_Formulas = result
End Function

Private Sub private_Restore( _
    ByVal table As ListObject, _
    ByVal transaction As Object, _
    ByVal prefix As String _
)
    Dim count As Long
    Dim cell As Range
    Dim formats As Object
    Dim formulas As Object
    Dim data As Variant
    Dim key As String

    count = transaction(prefix & "Count")
    Do While table.ListRows.Count > count
        table.ListRows(table.ListRows.Count).Delete
    Loop
    Do While table.ListRows.Count < count
        table.ListRows.Add
    Loop
    If count = 0 Then Exit Sub
    Set formats = transaction(prefix & "Formats")
    Set formulas = transaction(prefix & "Formulas")
    data = transaction(prefix & "Data")
    ' Restore values as values, including literal text beginning with an equals sign.
    table.DataBodyRange.NumberFormat = "@"
    table.DataBodyRange.Value2 = data
    For Each cell In table.DataBodyRange.Cells
        key = cell.Row & "/" & cell.Column
        cell.NumberFormat = formats(key)
        If formulas.Exists(key) Then cell.Formula = formulas(key)
    Next cell
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------