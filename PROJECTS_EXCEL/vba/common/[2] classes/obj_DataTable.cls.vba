VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_DataTable"
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_headers As Variant
Private m_values As Variant
Private m_rows As Long
Private m_columns As Long

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
Public Property Get Headers() As Variant
    Headers = m_headers
End Property

Public Property Get Values() As Variant
    Values = m_values
End Property

Public Property Get RowCount() As Long
    RowCount = m_rows
End Property

Public Property Get ColumnCount() As Long
    ColumnCount = m_columns
End Property

' //
' // API
' //
Public Function Initialize( _
    ByVal headers As Variant, _
    ByVal values As Variant, _
    ByVal rows As Long, _
    ByRef diagnostic As String _
) As Boolean
    Dim i As Long
    Dim j As Long
    Dim index As Long

    On Error GoTo EH
    If m_isDisposed Or m_isInitialized Then
        diagnostic = "Object is already initialized or disposed."
        Exit Function
    End If
    If rows < 0 Or Not VBA.IsArray(headers) Then
        Err.Raise VBA.vbObjectError + 2120, , "Invalid result dimensions."
    End If
    m_columns = UBound(headers) - LBound(headers) + 1
    If m_columns <= 0 Then
        Err.Raise VBA.vbObjectError + 2121, , "At least one header is required."
    End If
    ReDim m_headers(0 To m_columns - 1)
    For i = 1 To m_columns
        m_headers(i - 1) = VBA.CStr(headers(LBound(headers) + i - 1))
    Next i
    For i = 0 To m_columns - 1
        If VBA.Len(VBA.Trim$(m_headers(i))) = 0 Then
            Err.Raise VBA.vbObjectError + 2122, , "Empty column header."
        End If
        index = ex_TableQuery.fn_HeaderIndex(m_headers, m_headers(i))
    Next i
    If rows > 0 Then
        If Not VBA.IsArray(values) Then
            Err.Raise VBA.vbObjectError + 2123, , "Values must be a two-dimensional array."
        End If
        If UBound(values, 1) - LBound(values, 1) + 1 <> rows Or UBound(values, 2) - LBound(values, 2) + 1 <> m_columns Then
            Err.Raise VBA.vbObjectError + 2124, , "Values dimensions do not match the schema."
        End If
        ReDim m_values(1 To rows, 1 To m_columns)
        For i = 1 To rows
            For j = 1 To m_columns
                m_values(i, j) = values(LBound(values, 1) + i - 1, LBound(values, 2) + j - 1)
            Next j
        Next i
    End If
    m_rows = rows
    diagnostic = VBA.vbNullString
    m_isInitialized = True
    Initialize = True
    Exit Function
EH:
    diagnostic = Err.Description
    Me.Dispose
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    m_headers = Empty
    m_values = Empty
    m_rows = 0
    m_columns = 0
End Sub

Public Function ValueAt( _
    ByVal row As Long, _
    ByVal column As Long _
) As Variant
    If m_isDisposed Or Not m_isInitialized Then
        Err.Raise VBA.vbObjectError + 2165, , "Object is not initialized or is disposed."
    End If
    If row < 1 Or row > m_rows Or column < 1 Or column > m_columns Then
        Err.Raise VBA.vbObjectError + 2125, , "Cell index out of bounds."
    End If
    ValueAt = m_values(row, column)
End Function

Public Function TryCreateUiTable( _
    ByRef output As obj_UiRawTable, _
    ByRef diagnostic As String, _
    Optional ByVal title As String _
) As Boolean
    If m_isDisposed Or Not m_isInitialized Then
        Err.Raise VBA.vbObjectError + 2165, , "Object is not initialized or is disposed."
    End If
    Set output = New obj_UiRawTable
    If m_rows = 0 Then
        TryCreateUiTable = output.InitializeEmpty(m_headers, title)
    Else
        TryCreateUiTable = output.Initialize(m_values, m_headers, title)
    End If
    If Not TryCreateUiTable Then
        diagnostic = "Cannot create UI table."
        Set output = Nothing
    End If
End Function