Attribute VB_Name = "ex_TDB_Places"
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
Public Function fn_Source() As obj_TableSource
    Dim tableSource As New obj_TableSource
    Dim fileSystem As Object
    Dim path As String

    If Not tableSource.Initialize() Then VBA.Err.Raise 5, , "Cannot initialize places source."
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    path = ex_TravelDocumentsBuilder.fn_Config("lookup.Places.source.path")
    If fileSystem.GetDriveName(path) = VBA.vbNullString And VBA.Left$(path, 2) <> "\\" Then
        path = fileSystem.BuildPath(ThisWorkbook.Path, path)
    End If
    tableSource.WorkbookPath = fileSystem.GetAbsolutePathName(path)
    tableSource.TableName = ex_TravelDocumentsBuilder.fn_Config("lookup.Places.source.table")
    tableSource.SheetName = ex_TravelDocumentsBuilder.fn_Config("lookup.Places.source.sheet")
    Set fn_Source = tableSource
End Function

Public Function fn_Search(ByVal searchText As String) As obj_LookupResult
    Dim tableSource As obj_TableSource
    Dim tableQuery As obj_TableQuery
    Dim snapshot As obj_TableSnapshotSet
    Dim dataTable As obj_DataTable
    Dim lookupResult As obj_LookupResult
    Dim lookupRecord As obj_LookupRecord
    Dim records As New Collection
    Dim columns As New Collection
    Dim fields As Object
    Dim column As Object
    Dim field As Variant
    Dim value As Variant
    Dim address As String
    Dim rowIndex As Long
    Dim limit As Long
    Dim hasMore As Boolean
    Dim diagnostic As String
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim spans As Variant
    Dim columnIndex As Long

    On Error GoTo Failed
    Set snapshot = ex_TableSnapshots.fn_Active()
    If snapshot Is Nothing Then VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("text.DependenciesRequired")
    spans = VBA.Split(ex_TravelDocumentsBuilder.fn_Config("lookup.Places.columnSpans"), "|")
    For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("lookup.Places.fields"), "|")
        Set column = VBA.CreateObject("Scripting.Dictionary")
        column.Add "name", VBA.CStr(field)
        column.Add "caption", ex_TravelDocumentsBuilder.fn_Config("lookup.Places.columns." & field)
        If columnIndex > UBound(spans) Then VBA.Err.Raise 5, , "Places column span count mismatch."
        column.Add "columnSpan", VBA.CLng(spans(columnIndex))
        columnIndex = columnIndex + 1
        columns.Add column
    Next field
    If columnIndex <> UBound(spans) + 1 Then VBA.Err.Raise 5, , "Places column span count mismatch."
    limit = VBA.CLng(ex_TravelDocumentsBuilder.fn_Config("lookup.Places.maxResults"))
    If VBA.Len(ex_TableQuery.fn_NormalizeText(searchText)) >= VBA.CLng(ex_TravelDocumentsBuilder.fn_Config("lookup.Places.minChars")) Then
        Set tableSource = fn_Source()
        Set tableQuery = New obj_TableQuery
        If Not tableQuery.Initialize() Then VBA.Err.Raise 5, , "Cannot initialize places query."
        For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("lookup.Places.sourceFields"), "|")
            tableQuery.AddColumn ex_TravelDocumentsBuilder.fn_Config("lookup.Places.columns." & field)
        Next field
        tableQuery.AddCondition ex_TravelDocumentsBuilder.fn_Config("lookup.Places.columns.Name"), QueryContains, searchText, QueryText, True
        tableQuery.Limit = limit + 1
        If Not snapshot.TryQuery(tableSource, tableQuery, dataTable, diagnostic) Then VBA.Err.Raise 5, , diagnostic
        hasMore = dataTable.RowCount > limit
        For rowIndex = 1 To dataTable.RowCount
            If rowIndex > limit Then Exit For
            Set fields = VBA.CreateObject("Scripting.Dictionary")
            address = ex_TravelDocumentsBuilder.fn_Config("lookup.Places.addressTemplate")
            For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("lookup.Places.sourceFields"), "|")
                value = dataTable.ValueAt(rowIndex, ex_TableQuery.fn_HeaderIndex( _
                    dataTable.Headers, ex_TravelDocumentsBuilder.fn_Config("lookup.Places.columns." & field)))
                If VBA.IsError(value) Or VBA.IsNull(value) Then VBA.Err.Raise 5, , VBA.CStr(field)
                fields.Add VBA.CStr(field), ex_TableQuery.fn_NormalizeText(VBA.CStr(value))
                address = VBA.Replace$(address, "{" & field & "}", fields(field))
            Next field
            fields.Add "Address", address
            Set lookupRecord = New obj_LookupRecord
            If Not lookupRecord.Initialize(address, fields) Then VBA.Err.Raise 5, , "Cannot initialize place candidate."
            records.Add lookupRecord
        Next rowIndex
    End If
    Set lookupResult = New obj_LookupResult
    If Not lookupResult.Initialize(records, columns, hasMore, diagnostic) Then VBA.Err.Raise 5, , diagnostic
    Set fn_Search = lookupResult
Cleanup:
    If Not dataTable Is Nothing Then dataTable.Dispose
    If Not tableQuery Is Nothing Then tableQuery.Dispose
    If Not tableSource Is Nothing Then tableSource.Dispose
    If errorNumber <> 0 Then VBA.Err.Raise errorNumber, "Tickets.Places", errorDescription
    Exit Function
Failed:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    Resume Cleanup
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------