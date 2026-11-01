Attribute VB_Name = "ex_TDB_Declensions"
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
Public Sub fn_Enrich(ByVal values As Object)
    Dim index As Long
    Dim prefix As String
    Dim output As String
    Dim detail As String

    On Error GoTo Failed
    For index = 1 To VBA.CLng(ex_TravelDocumentsBuilder.fn_Config("declensions.count"))
        prefix = "declensions.lookups." & VBA.CStr(index)
        output = ex_TravelDocumentsBuilder.fn_Config(prefix & ".output")
        values(output) = private_Lookup(prefix, values)
    Next index
    Exit Sub
Failed:
    detail = VBA.Err.Description
    VBA.Err.Raise 5, "Tickets.Declensions", VBA.Replace$(VBA.Replace$( _
        ex_TravelDocumentsBuilder.fn_Config("declensions.error"), "{field}", output), "{detail}", detail)
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_Lookup( _
    ByVal prefix As String, _
    ByVal values As Object _
) As String
    Dim snapshot As obj_TableSnapshotSet
    Dim tableSource As obj_TableSource
    Dim tableQuery As obj_TableQuery
    Dim dataTable As obj_DataTable
    Dim fileSystem As Object
    Dim sourcePrefix As String
    Dim path As String
    Dim keyHeader As String
    Dim valueHeader As String
    Dim lookupInput As String
    Dim value As Variant
    Dim diagnostic As String
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo Failed
    Set snapshot = ex_TableSnapshots.fn_Active()
    If snapshot Is Nothing Then VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("text.DependenciesRequired")
    sourcePrefix = "declensions.sources." & ex_TravelDocumentsBuilder.fn_Config(prefix & ".source")
    Set tableSource = New obj_TableSource
    Set tableQuery = New obj_TableQuery
    If Not tableSource.Initialize() Or Not tableQuery.Initialize() Then VBA.Err.Raise 5, , sourcePrefix
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    path = ex_TravelDocumentsBuilder.fn_Config(sourcePrefix & ".path")
    If fileSystem.GetDriveName(path) = VBA.vbNullString And VBA.Left$(path, 2) <> "\\" Then
        path = fileSystem.BuildPath(ThisWorkbook.Path, path)
    End If
    tableSource.WorkbookPath = fileSystem.GetAbsolutePathName(path)
    tableSource.SheetName = ex_TravelDocumentsBuilder.fn_Config(sourcePrefix & ".sheet")
    tableSource.RangeAddress = ex_TravelDocumentsBuilder.fn_Config(sourcePrefix & ".range")
    keyHeader = ex_TravelDocumentsBuilder.fn_Config(prefix & ".keyHeader")
    valueHeader = ex_TravelDocumentsBuilder.fn_Config(prefix & ".valueHeader")
    lookupInput = VBA.CStr(values(ex_TravelDocumentsBuilder.fn_Config(prefix & ".input")))
    If VBA.Len(lookupInput) = 0 Then VBA.Err.Raise 5, , keyHeader
    tableQuery.AddColumn keyHeader
    tableQuery.AddColumn valueHeader
    tableQuery.AddCondition keyHeader, QueryEquals, lookupInput, QueryText, True
    tableQuery.Limit = 2
    If Not snapshot.TryQuery(tableSource, tableQuery, dataTable, diagnostic) Then VBA.Err.Raise 5, , diagnostic
    If dataTable.RowCount <> 1 Then VBA.Err.Raise 5, , tableSource.SheetName & ": " & keyHeader & "=" & lookupInput
    value = dataTable.ValueAt(1, ex_TableQuery.fn_HeaderIndex(dataTable.Headers, valueHeader))
    If VBA.IsError(value) Or VBA.IsNull(value) Then VBA.Err.Raise 5, , valueHeader
    private_Lookup = ex_TableQuery.fn_NormalizeText(VBA.CStr(value))
    If VBA.Len(private_Lookup) = 0 Then VBA.Err.Raise 5, , valueHeader & ": " & lookupInput
    dataTable.Dispose
    tableQuery.Dispose
    tableSource.Dispose
    Exit Function
Failed:
    failureNumber = VBA.Err.Number
    failureDescription = VBA.Err.Description
    On Error Resume Next
    If Not dataTable Is Nothing Then dataTable.Dispose
    If Not tableQuery Is Nothing Then tableQuery.Dispose
    If Not tableSource Is Nothing Then tableSource.Dispose
    On Error GoTo 0
    VBA.Err.Raise failureNumber, "Tickets.Declensions.Lookup", failureDescription
End Function
' --------------------------------------
' } // namespace Private
' --------------------------------------