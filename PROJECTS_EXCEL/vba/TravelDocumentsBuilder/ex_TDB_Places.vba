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
Public Function fn_Source( _
    Optional ByVal isAbroad As Boolean = False, _
    Optional ByVal isMilitary As Boolean = False _
) As obj_TableSource
    Dim tableSource As New obj_TableSource
    Dim fileSystem As Object
    Dim path As String
    Dim prefix As String

    prefix = fn_Prefix(isAbroad, isMilitary)
    If Not tableSource.Initialize() Then VBA.Err.Raise 5, , "Cannot initialize places source."
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    path = ex_TravelDocumentsBuilder.fn_Config(prefix & ".source.path")
    If fileSystem.GetDriveName(path) = VBA.vbNullString And VBA.Left$(path, 2) <> "\\" Then
        path = fileSystem.BuildPath(ThisWorkbook.Path, path)
    End If
    tableSource.WorkbookPath = fileSystem.GetAbsolutePathName(path)
    tableSource.TableName = ex_TravelDocumentsBuilder.fn_Config(prefix & ".source.table")
    tableSource.SheetName = ex_TravelDocumentsBuilder.fn_Config(prefix & ".source.sheet")
    Set fn_Source = tableSource
End Function

Public Function fn_Search( _
    ByVal searchText As String, _
    Optional ByVal isAbroad As Boolean = False, _
    Optional ByVal isMilitary As Boolean = False _
) As obj_LookupResult
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
    Dim prefix As String

    On Error GoTo Failed
    prefix = fn_Prefix(isAbroad, isMilitary)
    If isMilitary Then
        searchText = VBA.Replace$(searchText, ex_TravelDocumentsBuilder.fn_Config("placeChain.military.type"), VBA.vbNullString, 1, -1, VBA.vbTextCompare)
        searchText = VBA.Trim$(VBA.Replace$(searchText, "A", "А", 1, -1, VBA.vbTextCompare))
    End If
    spans = VBA.Split(ex_TravelDocumentsBuilder.fn_Config(prefix & ".columnSpans"), "|")
    For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config(prefix & ".fields"), "|")
        Set column = VBA.CreateObject("Scripting.Dictionary")
        column.Add "name", VBA.CStr(field)
        column.Add "caption", ex_TravelDocumentsBuilder.fn_Config(prefix & ".columns." & field)
        If columnIndex > UBound(spans) Then VBA.Err.Raise 5, , "Places column span count mismatch."
        column.Add "columnSpan", VBA.CLng(spans(columnIndex))
        columnIndex = columnIndex + 1
        columns.Add column
    Next field
    If columnIndex <> UBound(spans) + 1 Then VBA.Err.Raise 5, , "Places column span count mismatch."
    limit = VBA.CLng(ex_TravelDocumentsBuilder.fn_Config(prefix & ".maxResults"))
    If ex_TDB_Candidates.fn_IsBrowse(searchText) Then
        limit = ex_TDB_Candidates.fn_BrowseLimit(searchText)
    End If
    If ex_TDB_Candidates.fn_IsBrowse(searchText) Or VBA.Len(ex_TableQuery.fn_NormalizeText(searchText)) >= VBA.CLng(ex_TravelDocumentsBuilder.fn_Config(prefix & ".minChars")) Then
        Set snapshot = ex_TableSnapshots.fn_Active()
        If snapshot Is Nothing Then VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("text.DependenciesRequired")
        Set tableSource = fn_Source(isAbroad, isMilitary)
        Set tableQuery = New obj_TableQuery
        If Not tableQuery.Initialize() Then VBA.Err.Raise 5, , "Cannot initialize places query."
        For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config(prefix & ".sourceFields"), "|")
            tableQuery.AddColumn ex_TravelDocumentsBuilder.fn_Config(prefix & ".columns." & field)
        Next field
        If Not ex_TDB_Candidates.fn_IsBrowse(searchText) Then
            If isMilitary Then
                tableQuery.AddCondition ex_TravelDocumentsBuilder.fn_Config(prefix & ".columns.Unit"), QueryContains, searchText, QueryText, True
            Else
                tableQuery.AddCondition ex_TravelDocumentsBuilder.fn_Config(prefix & ".columns.Name"), QueryContains, searchText, QueryText, True
            End If
        End If
        tableQuery.Limit = limit + 1
        If Not snapshot.TryQuery(tableSource, tableQuery, dataTable, diagnostic) Then VBA.Err.Raise 5, , diagnostic
        hasMore = dataTable.RowCount > limit
        For rowIndex = 1 To dataTable.RowCount
            If rowIndex > limit Then Exit For
            Set fields = VBA.CreateObject("Scripting.Dictionary")
            For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config(prefix & ".sourceFields"), "|")
                value = dataTable.ValueAt(rowIndex, ex_TableQuery.fn_HeaderIndex( _
                    dataTable.Headers, ex_TravelDocumentsBuilder.fn_Config(prefix & ".columns." & field)))
                If VBA.IsError(value) Then VBA.Err.Raise 5, , VBA.CStr(field)
                If VBA.IsNull(value) Or VBA.IsEmpty(value) Then value = VBA.vbNullString
                fields.Add VBA.CStr(field), ex_TableQuery.fn_NormalizeText(VBA.CStr(value))
            Next field
            fields.Add "Abroad", ex_TravelDocumentsBuilder.fn_Config("generation.abroad.no")
            If isAbroad Then fields("Abroad") = ex_TravelDocumentsBuilder.fn_Config("generation.abroad.yes")
            If isMilitary Then
                address = VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config(prefix & ".addressTemplate"), "{Unit}", fields("Unit"))
            Else
                address = fn_Address(fields, isAbroad)
            End If
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

Public Function fn_Prefix( _
    ByVal isAbroad As Boolean, _
    Optional ByVal isMilitary As Boolean = False _
) As String
    fn_Prefix = "lookup.Places"
    If isAbroad Then fn_Prefix = "lookup.PlacesAbroad"
    If isMilitary Then fn_Prefix = "lookup.Militaries"
End Function

Public Function fn_IsAbroad(ByVal value As String) As Boolean
    If value = ex_TravelDocumentsBuilder.fn_Config("generation.abroad.yes") Then
        fn_IsAbroad = True
    ElseIf value <> ex_TravelDocumentsBuilder.fn_Config("generation.abroad.no") And VBA.Len(value) > 0 Then
        VBA.Err.Raise 5, , VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config("generation.message.Invalid"), _
            "{value}", ex_TravelDocumentsBuilder.fn_Config("text.VacationAbroad"))
    End If
End Function

Public Function fn_Address( _
    ByVal fields As Object, _
    ByVal isAbroad As Boolean _
) As String
    Dim address As String
    Dim field As Variant
    Dim ua As String
    Dim en As String
    Dim separator As String

    If Not isAbroad Then
        If VBA.Len(ex_TableQuery.fn_NormalizeText(VBA.CStr(fields("District")))) = 0 Then
            address = ex_TravelDocumentsBuilder.fn_Config("lookup.Places.addressWithoutDistrict")
        Else
            address = ex_TravelDocumentsBuilder.fn_Config("lookup.Places.addressTemplate")
        End If
        For Each field In VBA.Split(ex_TravelDocumentsBuilder.fn_Config("lookup.Places.sourceFields"), "|")
            address = VBA.Replace$(address, "{" & field & "}", VBA.CStr(fields(field)))
        Next field
        fn_Address = address
        Exit Function
    End If
    separator = ex_TravelDocumentsBuilder.fn_Config("lookup.PlacesAbroad.separator") & VBA.Space$(1)
    ua = VBA.CStr(fields("Name"))
    en = VBA.CStr(fields("NameEn"))
    If ua = VBA.CStr(fields("Country")) And en = VBA.CStr(fields("CountryEn")) Then
        ua = VBA.CStr(fields("Country"))
        en = VBA.CStr(fields("CountryEn"))
    Else
        If VBA.Len(fields("Type")) > 0 Then ua = fields("Type") & VBA.Space$(1) & ua
        If VBA.Len(fields("Region")) > 0 And fields("Region") <> fields("Name") Then ua = ua & separator & fields("Region")
        If VBA.Len(fields("RegionEn")) > 0 And fields("RegionEn") <> fields("NameEn") Then en = en & separator & fields("RegionEn")
        ua = ua & separator & fields("Country")
        en = en & separator & fields("CountryEn")
    End If
    address = ex_TravelDocumentsBuilder.fn_Config("lookup.PlacesAbroad.addressTemplate")
    address = VBA.Replace$(address, "{UA}", ua)
    fn_Address = VBA.Replace$(address, "{EN}", en)
End Function

Public Function fn_FormFields() As Variant
    fn_FormFields = VBA.Split(ex_TravelDocumentsBuilder.fn_Config("placeChain.fields"), "|")
End Function

Public Function fn_RowAddress(ByVal row As Object) As String
    Dim fields As Object
    Dim mapping As Variant
    Dim pair As Variant
    Dim isAbroad As Boolean

    isAbroad = fn_IsAbroad(VBA.CStr(row("PlaceAbroad")))
    Set fields = VBA.CreateObject("Scripting.Dictionary")
    For Each mapping In VBA.Split(ex_TravelDocumentsBuilder.fn_Config(fn_Prefix(isAbroad) & ".formFields"), "|")
        pair = VBA.Split(mapping, "=")
        fields(pair(0)) = row(pair(1))
    Next mapping
    fn_RowAddress = fn_Address(fields, isAbroad)
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------