VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_TableQuery"
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_columns As Collection
Private m_orders As Collection
Private m_filter As obj_QueryGroup
Public Limit As Long

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
    Set m_columns = New Collection
    Set m_orders = New Collection
    Set m_filter = New obj_QueryGroup
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Properties
' //
Public Property Get Columns() As Collection
    Dim result As New Collection
    Dim item As Variant

    For Each item In m_columns
        result.Add item
    Next item
    Set Columns = result
End Property

Public Property Get Orders() As Collection
    Dim result As New Collection
    Dim item As obj_QueryOrder

    For Each item In m_orders
        result.Add item
    Next item
    Set Orders = result
End Property

Public Property Get Filter() As obj_QueryGroup
    Set Filter = m_filter
End Property

' //
' // API
' //
Public Function Initialize() As Boolean
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    Set m_columns = New Collection
    Set m_orders = New Collection
    Set m_filter = New obj_QueryGroup
    If Not m_filter.Initialize() Then
        Me.Dispose
        Exit Function
    End If
    Limit = 0
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    Set m_columns = Nothing
    Set m_orders = Nothing
    If Not m_filter Is Nothing Then
        m_filter.Dispose
    End If
    Set m_filter = Nothing
    Limit = 0
End Sub

Public Sub AddColumn(ByVal columnName As String)
    Dim item As Variant

    If m_isDisposed Or Not m_isInitialized Then
        Err.Raise VBA.vbObjectError + 2165, , "Object is not initialized or is disposed."
    End If
    If VBA.Len(VBA.Trim$(columnName)) = 0 Then
        Err.Raise VBA.vbObjectError + 2118, , "Column name is required."
    End If
    For Each item In m_columns
        If VBA.StrComp(VBA.Trim$(columnName), item, VBA.vbTextCompare) = 0 Then
            Err.Raise VBA.vbObjectError + 2119, , "Duplicate selected column: " & columnName
        End If
    Next item
    m_columns.Add VBA.Trim$(columnName)
End Sub

Public Sub AddCondition( _
    ByVal columnName As String, _
    ByVal operation As en_QueryOp, _
    Optional ByVal value As Variant, _
    Optional ByVal valueType As en_QueryType = QueryText, _
    Optional ByVal normalizeText As Boolean = False, _
    Optional ByVal caseSensitive As Boolean = False _
)
    Dim condition As New obj_QueryCondition

    If m_isDisposed Or Not m_isInitialized Then
        Err.Raise VBA.vbObjectError + 2165, , "Object is not initialized or is disposed."
    End If
    If Not condition.Initialize() Then
        Err.Raise VBA.vbObjectError + 2163, , "Condition initialization failed."
    End If
    condition.ColumnName = columnName
    condition.Operation = operation
    If Not VBA.IsMissing(value) Then
        condition.Value = value
    End If
    condition.ValueType = valueType
    condition.NormalizeText = normalizeText
    condition.CaseSensitive = caseSensitive
    m_filter.AddCondition condition
End Sub

Public Sub AddOrderBy( _
    ByVal columnName As String, _
    Optional ByVal descending As Boolean = False, _
    Optional ByVal valueType As en_QueryType = QueryText, _
    Optional ByVal normalizeText As Boolean = False, _
    Optional ByVal caseSensitive As Boolean = False _
)
    Dim ordering As New obj_QueryOrder

    If m_isDisposed Or Not m_isInitialized Then
        Err.Raise VBA.vbObjectError + 2165, , "Object is not initialized or is disposed."
    End If
    If Not ordering.Initialize() Then
        Err.Raise VBA.vbObjectError + 2164, , "Order initialization failed."
    End If
    ordering.ColumnName = columnName
    ordering.Descending = descending
    ordering.ValueType = valueType
    ordering.NormalizeText = normalizeText
    ordering.CaseSensitive = caseSensitive
    m_orders.Add ordering
End Sub