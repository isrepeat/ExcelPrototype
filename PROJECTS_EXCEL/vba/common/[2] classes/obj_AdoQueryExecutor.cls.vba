VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_AdoQueryExecutor"
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
    Dim data As New obj_DataTable
    Dim conn As Object
    Dim rs As Object
    Dim headers As Variant
    Dim values As Variant
    Dim raw As Variant
    Dim sqlRef As String
    Dim props As String
    Dim extension As String
    Dim rows As Long
    Dim columns As Long
    Dim i As Long
    Dim j As Long
    Dim canonical As String
    Dim readColumns As Collection
    Dim name As Variant
    Dim selectSql As String

    On Error GoTo EH
    Set output = Nothing
    If m_isDisposed Or Not m_isInitialized Then
        diagnostic = "Query executor is not initialized or is disposed."
        Exit Function
    End If
    diagnostic = VBA.vbNullString
    If source Is Nothing Or query Is Nothing Then
        Err.Raise VBA.vbObjectError + 2150, , "Source and query are required."
    End If
    If Not source.TryValidate(diagnostic) Then
        Exit Function
    End If
    If query.Limit < 0 Then
        Err.Raise VBA.vbObjectError + 2151, , "Limit cannot be negative."
    End If
    canonical = VBA.CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(source.WorkbookPath)
    extension = VBA.LCase$(VBA.Mid$(canonical, VBA.InStrRev(canonical, ".") + 1))
    Select Case extension
        Case "xls"
            props = "Excel 8.0"
        Case "xlsx"
            props = "Excel 12.0 Xml"
        Case "xlsm"
            props = "Excel 12.0 Macro"
        Case "xlsb"
            props = "Excel 12.0"
        Case Else
            Err.Raise VBA.vbObjectError + 2152, , "Unsupported workbook format."
    End Select
    ' Метаданные берутся только из сохранённого файла, даже если книга открыта пользователем.
    If Not Me.private_TrySavedReference(source, sqlRef, headers, rows, diagnostic) Then
        Exit Function
    End If
    If rows = 0 Then
        If Not data.Initialize(headers, values, 0, diagnostic) Then
            Exit Function
        End If
        TryExecute = ex_TableQuery.fn_TryApply(data, query, output, diagnostic)
        Exit Function
    End If
    Set conn = VBA.CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""" & VBA.Replace$(canonical, """", """""") & _
        """;Extended Properties=""" & props & ";HDR=YES;IMEX=1;ReadOnly=True"";"
    Set rs = VBA.CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM " & sqlRef & " WHERE 1=0", conn, 0, 1
    ReDim headers(0 To rs.Fields.Count - 1)
    For i = 0 To rs.Fields.Count - 1
        headers(i) = rs.Fields(i).Name
    Next i
    rs.Close
    Set readColumns = ex_TableQuery.fn_ReadColumns(headers, query)
    For Each name In readColumns
        If VBA.Len(selectSql) > 0 Then
            selectSql = selectSql & ", "
        End If
        selectSql = selectSql & "[" & VBA.Replace$(VBA.CStr(name), "]", "]]") & "]"
    Next name
    rs.Open "SELECT " & selectSql & " FROM " & sqlRef, conn, 0, 1
    columns = rs.Fields.Count
    ReDim headers(0 To columns - 1)
    For i = 0 To columns - 1
        headers(i) = rs.Fields(i).Name
    Next i
    rows = 0
    If Not rs.EOF Then
        raw = rs.GetRows
        rows = UBound(raw, 2) + 1
        ReDim values(1 To rows, 1 To columns)
        For i = 1 To rows
            For j = 1 To columns
                values(i, j) = raw(j - 1, i - 1)
            Next j
        Next i
    End If
    If Not data.Initialize(headers, values, rows, diagnostic) Then
        GoTo Cleanup
    End If
    ' Общий evaluator обеспечивает одинаковые правила типов, пустоты, регистра и лимита.
    TryExecute = ex_TableQuery.fn_TryApply(data, query, output, diagnostic)
Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then
            rs.Close
        End If
    End If
    If Not conn Is Nothing Then
        If conn.State <> 0 Then
            conn.Close
        End If
    End If
    Set rs = Nothing
    Set conn = Nothing
    On Error GoTo 0
    Exit Function
EH:
    diagnostic = "ADO query: " & Err.Description
    Set output = Nothing
    If m_isDisposed Or Not m_isInitialized Then
        diagnostic = "Query executor is not initialized or is disposed."
        Exit Function
    End If
    Resume Cleanup
End Function

' //
' // Private
' //
Friend Function private_TrySavedReference( _
    ByVal source As obj_TableSource, _
    ByRef sqlRef As String, _
    ByRef headers As Variant, _
    ByRef rows As Long, _
    ByRef diagnostic As String _
) As Boolean
    Dim session As New obj_QueryWorkbookSession
    Dim body As Range

    On Error GoTo EH
    If Not session.Initialize() Then
        diagnostic = "Workbook session initialization failed."
        Exit Function
    End If
    If Not session.TryOpen(source, diagnostic, True) Then
        Exit Function
    End If
    headers = session.Headers
    sqlRef = session.SqlReference
    Set body = session.Body
    rows = 0
    If Not body Is Nothing Then
        rows = body.Rows.Count
    End If
    private_TrySavedReference = True
Cleanup:
    Set body = Nothing
    session.Dispose
    Exit Function
EH:
    diagnostic = "Saved source metadata: " & Err.Description
    Resume Cleanup
End Function