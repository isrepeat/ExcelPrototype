VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_UiRawTable"
Option Explicit

Implements obj_IUiTableSource

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_values As Variant
Private m_headers As Variant
Private m_title As String
Private m_rowCount As Long
Private m_columnCount As Long

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
Private Property Get obj_IUiTableSource_TableCount() As Long
    obj_IUiTableSource_TableCount = 1
End Property

Public Property Get Values() As Variant
    Values = m_values
End Property

Public Property Get Headers() As Variant
    Headers = m_headers
End Property

Public Property Get Title() As String
    Title = m_title
End Property

Public Property Get RowCount() As Long
    RowCount = m_rowCount
End Property

Public Property Get ColumnCount() As Long
    ColumnCount = m_columnCount
End Property

' //
' // Interface
' //

Private Function obj_IUiTableSource_GetTable(ByVal index As Long) As Object
    If index <> 1 Then Exit Function
    Set obj_IUiTableSource_GetTable = Me
End Function

' //
' // API
' //
Public Function Initialize(ByVal values As Variant, Optional ByVal headers As Variant, Optional ByVal title As String) As Boolean
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    If Not IsArray(values) Then Exit Function
    On Error GoTo EH
    m_rowCount = UBound(values, 1) - LBound(values, 1) + 1
    m_columnCount = UBound(values, 2) - LBound(values, 2) + 1
    If m_rowCount <= 0 Or m_columnCount <= 0 Then Exit Function
    m_values = values
    If Not IsMissing(headers) Then m_headers = headers
    m_title = VBA.Trim$(title)
    m_isInitialized = True
    Initialize = True
    Exit Function
EH:
    Me.Dispose
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    m_values = Empty
    m_headers = Empty
    m_title = VBA.vbNullString
    m_rowCount = 0
    m_columnCount = 0
End Sub

Public Function InitializeEmpty( _
    ByVal headers As Variant, _
    Optional ByVal title As String _
) As Boolean
    On Error GoTo EH
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    If Not VBA.IsArray(headers) Then Exit Function
    m_columnCount = UBound(headers) - LBound(headers) + 1
    If m_columnCount <= 0 Then Exit Function
    m_headers = headers
    m_title = title
    m_isInitialized = True
    InitializeEmpty = True
    Exit Function
EH:
    Me.Dispose
End Function

Public Function ValueAt(ByVal rowIndex As Long, ByVal columnIndex As Long) As Variant
    If rowIndex <= 0 Or rowIndex > m_rowCount Then Exit Function
    If columnIndex <= 0 Or columnIndex > m_columnCount Then Exit Function
    ValueAt = m_values(rowIndex, columnIndex)
End Function

Public Function HeaderAt(ByVal columnIndex As Long) As Variant
    If Not IsArray(m_headers) Then Exit Function
    If columnIndex <= 0 Then Exit Function
    On Error GoTo EH
    HeaderAt = m_headers(LBound(m_headers) + columnIndex - 1)
EH:
End Function