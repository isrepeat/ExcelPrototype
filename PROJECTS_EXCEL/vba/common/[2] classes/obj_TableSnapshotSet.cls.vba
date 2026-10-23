VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_TableSnapshotSet"
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_tables As Object
Private m_prepared As Object
Private m_indexes As Object

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Properties
' //
Public Property Get Prepared() As Object
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Snapshot is unavailable."
    End If
    Set Prepared = m_prepared
End Property

' //
' // API
' //
Public Function Initialize() As Boolean
    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    Set m_tables = VBA.CreateObject("Scripting.Dictionary")
    m_tables.CompareMode = VBA.vbTextCompare
    Set m_prepared = VBA.CreateObject("Scripting.Dictionary")
    Set m_indexes = VBA.CreateObject("Scripting.Dictionary")
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    Dim key As Variant
    Dim table As obj_DataTable

    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    If Not m_tables Is Nothing Then
        For Each key In m_tables.Keys
            Set table = m_tables(key)
            table.Dispose
        Next key
    End If
    Set m_tables = Nothing
    Set m_prepared = Nothing
    Set m_indexes = Nothing
End Sub

Public Function Contains(ByVal source As obj_TableSource) As Boolean
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Snapshot is unavailable."
    End If
    Contains = m_tables.Exists(ex_TableSnapshots.fn_SourceKey(source))
End Function

Public Function OwnsWorkbook(ByVal source As obj_TableSource) As Boolean
    Dim key As Variant
    Dim path As String

    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Snapshot is unavailable."
    End If
    path = VBA.LCase$(VBA.Replace$(source.WorkbookPath, "/", "\"))
    For Each key In m_tables.Keys
        If VBA.Left$(VBA.CStr(key), VBA.Len(path) + 1) = path & VBA.vbNullChar Then
            OwnsWorkbook = True
            Exit Function
        End If
    Next key
End Function

Public Function TryLoad( _
    ByVal source As obj_TableSource, _
    ByRef diagnostic As String _
) As Boolean
    Dim service As obj_TableQueryService
    Dim query As obj_TableQuery
    Dim table As obj_DataTable
    Dim key As String

    On Error GoTo EH
    key = ex_TableSnapshots.fn_SourceKey(source)
    If m_tables.Exists(key) Then
        TryLoad = True
        Exit Function
    End If
    Set service = New obj_TableQueryService
    Set query = New obj_TableQuery
    If Not service.Initialize() Or Not query.Initialize() Then
        Err.Raise 5, , "Cannot initialize snapshot query."
    End If
    ' Dependencies are one explicit snapshot of saved files; live history is never registered.
    service.Backend = QueryAdo
    If Not service.TryExecute(source, query, table, diagnostic, False) Then
        GoTo Cleanup
    End If
    m_tables.Add key, table
    TryLoad = True
Cleanup:
    If Not service Is Nothing Then
        service.Dispose
    End If
    If Not query Is Nothing Then
        query.Dispose
    End If
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function

Public Sub WarmIndex( _
    ByVal source As obj_TableSource, _
    ByVal header As String _
)
    Dim index As Object

    Set index = private_Index(source, header)
End Sub

Public Function TryQuery( _
    ByVal source As obj_TableSource, _
    ByVal query As obj_TableQuery, _
    ByRef output As obj_DataTable, _
    ByRef diagnostic As String _
) As Boolean
    Dim table As obj_DataTable
    Dim subset As obj_DataTable
    Dim condition As obj_QueryCondition
    Dim index As Object
    Dim rows As Collection
    Dim headers As Variant
    Dim values As Variant
    Dim key As String
    Dim row As Long
    Dim column As Long
    Dim count As Long

    On Error GoTo EH
    Set table = m_tables(ex_TableSnapshots.fn_SourceKey(source))
    query.Filter.Validate table.Headers
    Set condition = query.Filter.IndexCondition
    If condition Is Nothing Then
        TryQuery = ex_TableQuery.fn_TryApply(table, query, output, diagnostic)
        Exit Function
    End If
    ' A normalized, case-insensitive index is a superset; the full filter remains authoritative.
    Set index = private_Index(source, condition.ColumnName)
    key = ex_TableQuery.fn_NormalizeText(VBA.CStr(condition.Value))
    Set rows = New Collection
    If index.Exists(key) Then
        Set rows = index(key)
    End If
    headers = table.Headers
    count = rows.Count
    If count > 0 Then
        ReDim values(1 To count, 1 To table.ColumnCount)
        For row = 1 To count
            For column = 1 To table.ColumnCount
                values(row, column) = table.ValueAt(rows(row), column)
            Next column
        Next row
    End If
    Set subset = New obj_DataTable
    If Not subset.Initialize(headers, values, count, diagnostic) Then
        GoTo Cleanup
    End If
    TryQuery = ex_TableQuery.fn_TryApply(subset, query, output, diagnostic)
Cleanup:
    If Not subset Is Nothing Then
        subset.Dispose
    End If
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function

' //
' // Private
' //
Private Function private_Index( _
    ByVal source As obj_TableSource, _
    ByVal header As String _
) As Object
    Dim table As obj_DataTable
    Dim index As Object
    Dim rows As Collection
    Dim key As String
    Dim identity As String
    Dim value As Variant
    Dim row As Long
    Dim column As Long

    identity = ex_TableSnapshots.fn_SourceKey(source) & VBA.vbNullChar & VBA.LCase$(VBA.Replace$(VBA.Trim$(header), ".", "#"))
    If m_indexes.Exists(identity) Then
        Set private_Index = m_indexes(identity)
        Exit Function
    End If
    Set table = m_tables(ex_TableSnapshots.fn_SourceKey(source))
    column = ex_TableQuery.fn_HeaderIndex(table.Headers, header)
    Set index = VBA.CreateObject("Scripting.Dictionary")
    index.CompareMode = VBA.vbTextCompare
    For row = 1 To table.RowCount
        value = table.ValueAt(row, column)
        If VBA.IsError(value) Then
            Err.Raise 5, , "Excel error in snapshot index: " & header
        End If
        If Not VBA.IsNull(value) And Not VBA.IsEmpty(value) Then
            key = ex_TableQuery.fn_NormalizeText(VBA.CStr(value))
            If Not index.Exists(key) Then
                Set rows = New Collection
                index.Add key, rows
            End If
            Set rows = index(key)
            rows.Add row
        End If
    Next row
    m_indexes.Add identity, index
    Set private_Index = index
End Function