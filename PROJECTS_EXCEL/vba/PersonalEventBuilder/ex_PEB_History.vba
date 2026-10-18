Attribute VB_Name = "ex_PEB_History"
Option Explicit

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Save( _
    ByVal values As Object, _
    ByRef diagnostic As String _
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
    Dim owned As Boolean
    Dim added As Boolean
    Dim changed As Boolean
    Dim saved As Boolean
    Dim index As Long
    Dim matchIndex As Long
    Dim identifier As String
    Dim key As String
    Dim value As String
    Dim originalColumns As Long
    Dim collectionName As Variant

    On Error GoTo EH
    private_Open book, table, owned
    originalColumns = table.ListColumns.Count
    For Each collectionName In VBA.Array("MetaDocumentTables", "MetaTvoTables", "HistoryReports")
        If values.Exists(collectionName) Then
            If values(collectionName).Count > 0 Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryChildrenInTable")
        End If
    Next collectionName
    Set columns = private_Columns(table, True)
    Set rows = private_Read(table, columns)
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
    values("EventLabel") = ex_PEB_EventRoutes.fn_Value(VBA.CStr(values("EventId")), "caption")
    values("WordTemplatePath") = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("export.Word.documentTemplate"))
    ReDim data(1 To 1, 1 To table.ListColumns.Count)
    If matchIndex > 0 Then
        Set target = table.ListRows(rows(matchIndex)("__tableIndex"))
        backup = target.Range.Value2
        For index = 1 To table.ListColumns.Count
            data(1, index) = backup(1, index)
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
    If matchIndex > 0 And Not changed Then
        values("HistoryAction") = "SKIP"
        fn_Save = True
        GoTo Cleanup
    End If
    If matchIndex = 0 Then
        Set target = table.ListRows.Add
        added = True
    End If
    target.Range.NumberFormat = "@"
    target.Range.Value2 = data
    book.Save
    saved = True
    values("HistoryAction") = "INSERT"
    If Not added Then values("HistoryAction") = "UPDATE"
    fn_Save = True
Cleanup:
    On Error Resume Next
    If Not saved And Not target Is Nothing Then
        If added Then
            target.Delete
        ElseIf changed Then
            target.Range.Value2 = backup
        End If
    End If
    If Not saved And originalColumns > 0 Then
        Do While table.ListColumns.Count > originalColumns
            table.ListColumns(table.ListColumns.Count).Delete
        Loop
    End If
    If owned And Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function

Public Function fn_ReadOrder( _
    ByVal orderNumber As String, _
    ByVal orderDate As Variant, _
    ByRef events As Collection, _
    ByRef diagnostic As String, _
    Optional ByVal includeChildren As Boolean = True _
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
    Dim owned As Boolean
    Dim index As Long
    Dim parentId As String
    Dim typeName As String
    Dim dateText As String

    On Error GoTo EH
    Set events = Nothing
    If VBA.Len(VBA.Trim$(orderNumber)) = 0 Or VBA.Len(ex_PEB_Export.fn_Text(orderDate)) = 0 Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.OrderChooseRequired")
    End If
    dateText = private_Text(fn_Date(orderDate))
    private_Open book, table, owned
    Set columns = private_Columns(table, False)
    Set rows = private_Read(table, columns)
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
            If row("AcceptedOrderNumber") = orderNumber Then
                If private_Text(fn_Date(row("AcceptedOrderDate"))) = dateText Then
                    parentId = private_Identity(row)
                    If identities.Exists(parentId) Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError008")
                    identities.Add parentId, True
                    private_AddSorted selected, row
                End If
            End If
            Set row("MetaDocumentTables") = New Collection
            Set row("MetaTvoTables") = New Collection
            Set row("HistoryReports") = New Collection
        End If
    Next row
    If Not includeChildren Then
        If selected.Count = 0 Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryEmptyOrder")
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
            If root("AcceptedOrderNumber") = orderNumber Then
                If private_Text(fn_Date(root("AcceptedOrderDate"))) = dateText Then
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
        End If
    Next row
    If selected.Count = 0 Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryEmptyOrder")
    For Each row In selected
        Set row("MetaTvoTables") = New Collection
        private_Tvo rows, row, row("MetaTvoTables")
    Next row
    Set events = selected
    fn_ReadOrder = True
Cleanup:
    On Error Resume Next
    If owned And Not book Is Nothing Then book.Close False
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

    For Each field In VBA.Split(ex_PEB_Export.fn_Config("history.fields"), "|")
        mapping = VBA.Split(VBA.CStr(field), "=")
        text = private_Text(ex_PEB_Export.fn_Required(values, VBA.CStr(mapping(0))))
        fn_Stamp = fn_Stamp & VBA.Len(text) & ":" & text
    Next field
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
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Sub private_Open( _
    ByRef book As Workbook, _
    ByRef table As ListObject, _
    ByRef owned As Boolean _
)
    Dim candidate As Workbook
    Dim path As String

    path = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("history.path"))
    For Each candidate In Application.Workbooks
        If VBA.StrComp(candidate.FullName, path, VBA.vbTextCompare) = 0 Then Set book = candidate
    Next candidate
    If book Is Nothing Then
        If VBA.Len(VBA.Dir$(path)) = 0 Then Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError014"), "{value}", VBA.CStr(path))
        Set book = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=False)
        owned = True
    End If
    If book.ReadOnly Then Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError015"), "{value}", VBA.CStr(path))
    If Not book.Saved Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistorySaveFirst")
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
        If Not found And extend Then
            Set column = table.ListColumns.Add
            column.Name = mapping(1)
            column.Range.NumberFormat = "@"
            result.Add mapping(0), column.Index
            found = True
        End If
        If Not found Then Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.HistoryError019"), "{value}", VBA.CStr(mapping(1)))
    Next field
    Set private_Columns = result
End Function

Private Function private_Read( _
    ByVal table As ListObject, _
    ByVal columns As Object _
) As Collection
    Dim result As New Collection
    Dim row As Object
    Dim data As Variant
    Dim field As Variant
    Dim index As Long
    Dim hasData As Boolean

    If Not table.DataBodyRange Is Nothing Then
        data = table.DataBodyRange.Value2
        For index = 1 To UBound(data, 1)
            Set row = VBA.CreateObject("Scripting.Dictionary")
            hasData = False
            For Each field In columns.Keys
                row.Add field, private_Text(data(index, columns(field)))
                If VBA.Len(row(field)) > 0 Then hasData = True
            Next field
            row.Add "__tableIndex", index
            If hasData Then
                If VBA.Len(row("RowId")) = 0 Or VBA.Len(row("RowType")) = 0 Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.HistoryError020")
                result.Add row
            End If
        Next index
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