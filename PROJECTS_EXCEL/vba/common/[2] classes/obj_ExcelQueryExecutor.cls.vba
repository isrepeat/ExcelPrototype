VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_ExcelQueryExecutor"
Option Explicit

Implements obj_ITableQueryExecutor

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Interface
' //
Private Function obj_ITableQueryExecutor_TryExecute( _
    ByVal source As obj_TableSource, _
    ByVal query As obj_TableQuery, _
    ByRef output As obj_DataTable, _
    ByRef diagnostic As String _
) As Boolean
    obj_ITableQueryExecutor_TryExecute = Me.TryExecute(source, query, output, diagnostic)
End Function

' //
' // API
' //
Public Function Initialize() As Boolean
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
End Sub

Public Function TryExecute( _
    ByVal source As obj_TableSource, _
    ByVal query As obj_TableQuery, _
    ByRef output As obj_DataTable, _
    ByRef diagnostic As String _
) As Boolean
    Dim session As New obj_QueryWorkbookSession
    Dim data As New obj_DataTable
    Dim body As Range
    Dim values As Variant
    Dim scalar As Variant
    Dim rows As Long
    Dim startedAt As Double

    On Error GoTo EH
    Set output = Nothing
    If m_isDisposed Or Not m_isInitialized Then
        diagnostic = "Query executor is not initialized or is disposed."
        Exit Function
    End If
    diagnostic = VBA.vbNullString
    If source Is Nothing Or query Is Nothing Then
        Err.Raise VBA.vbObjectError + 2140, , "Source and query are required."
    End If
    If Not source.TryValidate(diagnostic) Then
        Exit Function
    End If
    If query.Limit < 0 Then
        Err.Raise VBA.vbObjectError + 2141, , "Limit cannot be negative."
    End If
    If Not session.Initialize() Then
        diagnostic = "Workbook session initialization failed."
        Exit Function
    End If
    If Not session.TryOpen(source, diagnostic) Then
        Exit Function
    End If
    Set body = session.Body
    ex_Core.fn_Diagnostic_WriteLog "QUERY_STAGE_STARTED | Name=ReadRange"
    startedAt = VBA.Timer
    If Not body Is Nothing Then
        rows = body.Rows.Count
        values = body.Value2
        If Not VBA.IsArray(values) Then
            scalar = values
            ReDim values(1 To 1, 1 To 1)
            values(1, 1) = scalar
        End If
    End If
    ex_Core.fn_Diagnostic_WritePerf "Query.ReadRange | Rows=" & VBA.CStr(rows), startedAt
    ex_Core.fn_Diagnostic_WriteLog "QUERY_STAGE_STARTED | Name=BuildDataTable"
    startedAt = VBA.Timer
    If Not data.Initialize(session.Headers, values, rows, diagnostic) Then
        GoTo Cleanup
    End If
    ex_Core.fn_Diagnostic_WritePerf "Query.BuildDataTable | Columns=" & VBA.CStr(data.ColumnCount), startedAt
    TryExecute = ex_TableQuery.fn_TryApply(data, query, output, diagnostic)
Cleanup:
    Set body = Nothing
    startedAt = VBA.Timer
    session.Dispose
    ex_Core.fn_Diagnostic_WritePerf "Query.CloseSession", startedAt
    Exit Function
EH:
    diagnostic = "Excel query: " & Err.Description
    Set output = Nothing
    If m_isDisposed Or Not m_isInitialized Then
        diagnostic = "Query executor is not initialized or is disposed."
        Exit Function
    End If
    Resume Cleanup
End Function