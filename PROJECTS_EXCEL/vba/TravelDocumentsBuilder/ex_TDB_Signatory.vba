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
Public Function fn_Resolve( _
    ByVal orderNumber As String, _
    ByVal orderDate As Date _
) As Object
    Dim snapshot As obj_TableSnapshotSet
    Dim source As obj_TableSource
    Dim query As obj_TableQuery
    Dim table As obj_DataTable
    Dim files As Object
    Dim result As Object
    Dim candidate As Object
    Dim header As Variant
    Dim field As Variant
    Dim value As Variant
    Dim number As String
    Dim signer As String
    Dim actingPrefix As String
    Dim sourceDate As Date
    Dim index As Long
    Dim diagnostic As String
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo Failed
    Set snapshot = ex_TableSnapshots.fn_Active()
    If snapshot Is Nothing Then VBA.Err.Raise 5, , ex_TravelDocumentsBuilder.fn_Config("text.DependenciesRequired")
    Set source = New obj_TableSource
    Set query = New obj_TableQuery
    If Not source.Initialize() Or Not query.Initialize() Then private_Fail "SignatorySource", "orders.source"
    Set files = VBA.CreateObject("Scripting.FileSystemObject")
    number = ex_TravelDocumentsBuilder.fn_Config("orders.source.path")
    If files.GetDriveName(number) = VBA.vbNullString And VBA.Left$(number, 2) <> "\\" Then number = files.BuildPath(ThisWorkbook.Path, number)
    source.WorkbookPath = files.GetAbsolutePathName(number)
    source.SheetName = ex_TravelDocumentsBuilder.fn_Config("orders.source.sheet")
    source.RangeAddress = ex_TravelDocumentsBuilder.fn_Config("orders.source.range")
    For Each header In VBA.Array("dateHeader", "numberHeader", "signatoryHeader", "signatoryNameHeader", "signatoryRankHeader")
        query.AddColumn ex_TravelDocumentsBuilder.fn_Config("orders." & VBA.CStr(header))
    Next header
    If Not snapshot.TryQuery(source, query, table, diagnostic) Then private_Fail "SignatorySource", diagnostic
    actingPrefix = VBA.UCase$(ex_TravelDocumentsBuilder.fn_Config("word.signatory.actingPrefix"))
    If VBA.Len(actingPrefix) = 0 Then private_Fail "SignatorySource", "word.signatory.actingPrefix"
    For index = 1 To table.RowCount
        number = private_Value(table, index, "numberHeader")
        If VBA.IsNumeric(number) Then
            If VBA.CStr(VBA.CDec(number)) = orderNumber Then
                value = table.ValueAt(index, ex_TableQuery.fn_HeaderIndex(table.Headers, ex_TravelDocumentsBuilder.fn_Config("orders.dateHeader")))
                If VBA.VarType(value) = VBA.vbDate Or VBA.IsNumeric(value) Then
                    sourceDate = VBA.CDate(VBA.Int(VBA.CDbl(VBA.CDate(value))))
                Else
                    sourceDate = ex_TDB_History.fn_Date(value)
                End If
                If sourceDate = orderDate Then
                    signer = private_Value(table, index, "signatoryHeader")
                    Set candidate = VBA.CreateObject("Scripting.Dictionary")
                    candidate.Add "SignatoryName", private_Value(table, index, "signatoryNameHeader")
                    candidate.Add "SignatoryRank", private_Value(table, index, "signatoryRankHeader")
                    candidate.Add "SignatoryPosition", ex_TravelDocumentsBuilder.fn_Config("word.signatory.position")
                    If VBA.UCase$(signer) = actingPrefix Or VBA.Left$(VBA.UCase$(signer), VBA.Len(actingPrefix) + 1) = actingPrefix & " " Then
                        candidate("SignatoryPosition") = ex_TravelDocumentsBuilder.fn_Config("word.signatory.actingPosition")
                    End If
                    If VBA.Len(signer) = 0 Or VBA.Len(candidate("SignatoryName")) = 0 Or VBA.Len(candidate("SignatoryRank")) = 0 Then
                        private_Fail "SignatoryRequired", orderNumber & " / " & VBA.Format$(orderDate, "dd.mm.yyyy")
                    End If
                    If result Is Nothing Then
                        Set result = candidate
                    Else
                        For Each field In candidate.Keys
                            If result(field) <> candidate(field) Then private_Fail "SignatoryAmbiguous", orderNumber & " / " & VBA.Format$(orderDate, "dd.mm.yyyy")
                        Next field
                    End If
                End If
            End If
        End If
    Next index
    If result Is Nothing Then private_Fail "SignatoryRequired", orderNumber & " / " & VBA.Format$(orderDate, "dd.mm.yyyy")
    Set fn_Resolve = result
    table.Dispose
    query.Dispose
    source.Dispose
    Exit Function
Failed:
    failureNumber = VBA.Err.Number
    failureDescription = VBA.Err.Description
    On Error Resume Next
    If Not table Is Nothing Then table.Dispose
    If Not query Is Nothing Then query.Dispose
    If Not source Is Nothing Then source.Dispose
    On Error GoTo 0
    VBA.Err.Raise failureNumber, "TicketsWord.Signatory", failureDescription
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Function private_Value( _
    ByVal table As obj_DataTable, _
    ByVal row As Long, _
    ByVal headerKey As String _
) As String
    Dim value As Variant

    value = table.ValueAt(row, ex_TableQuery.fn_HeaderIndex(table.Headers, ex_TravelDocumentsBuilder.fn_Config("orders." & headerKey)))
    If VBA.IsError(value) Then private_Fail "SignatorySource", ex_TravelDocumentsBuilder.fn_Config("orders." & headerKey)
    If VBA.IsNull(value) Or VBA.IsEmpty(value) Then Exit Function
    private_Value = ex_TableQuery.fn_NormalizeText(VBA.CStr(value))
End Function

Private Sub private_Fail( _
    ByVal key As String, _
    ByVal details As String _
)
    VBA.Err.Raise VBA.vbObjectError + 2480, "TicketsWord.Signatory", _
        VBA.Replace$(ex_TravelDocumentsBuilder.fn_Config("word.text." & key), "{details}", details)
End Sub
' --------------------------------------
' } // namespace Private