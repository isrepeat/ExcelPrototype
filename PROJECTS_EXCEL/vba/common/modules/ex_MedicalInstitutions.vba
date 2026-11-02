Attribute VB_Name = "ex_MedicalInstitutions"
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
Public Function fn_TryQuery( _
    ByVal source As obj_TableSource, _
    ByVal query As obj_TableQuery, _
    ByVal backend As en_QueryBackend, _
    ByRef output As obj_DataTable, _
    ByRef diagnostic As String _
) As Boolean
    Dim institutions As obj_DataTable
    Dim units As obj_DataTable
    Dim combined As obj_DataTable
    Dim headers As Variant
    Dim unitHeaders As Variant
    Dim values As Variant
    Dim rows As New Collection
    Dim record As Variant
    Dim indexes As Variant
    Dim row As Long
    Dim column As Long
    Dim flagColumn As Long
    Dim flagValue As String

    On Error GoTo Failed
    Set output = Nothing
    Set institutions = private_Read(source.WorkbookPath, private_Config("tables.institutions"), backend)
    Set units = private_Read(source.WorkbookPath, private_Config("tables.units"), backend)
    headers = VBA.Split(private_Config("headers"), "|")
    unitHeaders = VBA.Split(private_Config("unitHeaders"), "|")
    If UBound(headers) <> UBound(unitHeaders) Then
        VBA.Err.Raise 5, , "Medical catalog column mappings differ in length."
    End If
    ReDim indexes(0 To UBound(headers))
    For column = 0 To UBound(headers)
        indexes(column) = ex_TableQuery.fn_HeaderIndex(institutions.Headers, VBA.CStr(headers(column)))
    Next column
    For row = 1 To institutions.RowCount
        ReDim record(0 To UBound(headers))
        For column = 0 To UBound(headers)
            record(column) = institutions.ValueAt(row, indexes(column))
        Next column
        rows.Add record
    Next row
    For column = 0 To UBound(headers)
        indexes(column) = ex_TableQuery.fn_HeaderIndex(units.Headers, VBA.CStr(unitHeaders(column)))
    Next column
    flagColumn = ex_TableQuery.fn_HeaderIndex(units.Headers, private_Config("flag.header"))
    flagValue = private_Config("flag.value")
    For row = 1 To units.RowCount
        If VBA.StrComp(private_Text(units.ValueAt(row, flagColumn)), flagValue, VBA.vbTextCompare) = 0 Then
            ReDim record(0 To UBound(headers))
            For column = 0 To UBound(headers)
                record(column) = units.ValueAt(row, indexes(column))
            Next column
            rows.Add record
        End If
    Next row
    If rows.Count > 0 Then
        ReDim values(1 To rows.Count, 1 To UBound(headers) + 1)
        For row = 1 To rows.Count
            record = rows(row)
            For column = 0 To UBound(headers)
                values(row, column + 1) = record(column)
            Next column
        Next row
    End If
    Set combined = New obj_DataTable
    If Not combined.Initialize(headers, values, rows.Count, diagnostic) Then
        GoTo Cleanup
    End If
    fn_TryQuery = ex_TableQuery.fn_TryApply(combined, query, output, diagnostic)
Cleanup:
    If Not institutions Is Nothing Then
        institutions.Dispose
    End If
    If Not units Is Nothing Then
        units.Dispose
    End If
    If Not combined Is Nothing Then
        combined.Dispose
    End If
    Exit Function
Failed:
    diagnostic = "Medical catalog: " & VBA.Err.Description
    Resume Cleanup
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_Read( _
    ByVal path As String, _
    ByVal tableName As String, _
    ByVal backend As en_QueryBackend _
) As obj_DataTable
    Dim source As obj_TableSource
    Dim query As obj_TableQuery
    Dim service As obj_TableQueryService
    Dim result As obj_DataTable
    Dim diagnostic As String
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo Failed
    Set source = New obj_TableSource
    Set query = New obj_TableQuery
    Set service = New obj_TableQueryService
    If Not source.Initialize() Or Not query.Initialize() Or Not service.Initialize() Then
        VBA.Err.Raise 5, , "Cannot initialize medical catalog query."
    End If
    source.WorkbookPath = path
    source.TableName = tableName
    service.Backend = backend
    If Not service.TryExecute(source, query, result, diagnostic, False) Then
        VBA.Err.Raise 5, , diagnostic
    End If
    Set private_Read = result
Cleanup:
    If Not service Is Nothing Then
        service.Dispose
    End If
    If Not query Is Nothing Then
        query.Dispose
    End If
    If Not source Is Nothing Then
        source.Dispose
    End If
    If errorNumber <> 0 Then
        On Error GoTo 0
        VBA.Err.Raise errorNumber, "MedicalCatalog.Read", errorDescription
    End If
    Exit Function
Failed:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    Resume Cleanup
End Function

Private Function private_Config(ByVal key As String) As String
    If Not ex_Core.fn_TryGetWorkbookConfigValue("__cfgDependencies.common::medicalCatalog." & key, private_Config) Then
        VBA.Err.Raise 5, , "Missing configuration: common::medicalCatalog." & key
    End If
End Function

Private Function private_Text(ByVal value As Variant) As String
    If VBA.IsError(value) Or VBA.IsNull(value) Or VBA.IsEmpty(value) Then
        Exit Function
    End If
    private_Text = ex_TableQuery.fn_NormalizeText(VBA.CStr(value))
End Function