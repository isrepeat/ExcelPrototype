VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_UiRawTableList"
Option Explicit

Implements obj_IUiTableSource

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_tables As Collection

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
    Set m_tables = New Collection
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Interface
' //
Private Property Get obj_IUiTableSource_TableCount() As Long
    If Not m_tables Is Nothing Then obj_IUiTableSource_TableCount = m_tables.Count
End Property

Private Function obj_IUiTableSource_GetTable(ByVal index As Long) As Object
    If m_tables Is Nothing Then Exit Function
    If index <= 0 Or index > m_tables.Count Then Exit Function
    Set obj_IUiTableSource_GetTable = m_tables.Item(index)
End Function

' //
' // API
' //
Public Function Initialize() As Boolean
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    If m_tables Is Nothing Then Set m_tables = New Collection
    Initialize = True
    m_isInitialized = Initialize
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    Set m_tables = Nothing
End Sub

Public Sub Clear()
    If m_isDisposed Or Not m_isInitialized Then
        Err.Raise VBA.vbObjectError + 2165, , "Table list is not initialized or is disposed."
    End If
    Set m_tables = New Collection
End Sub

Public Function Add(ByVal table As obj_UiRawTable) As Boolean
    If m_isDisposed Or Not m_isInitialized Then Exit Function
    If table Is Nothing Then Exit Function
    If m_tables Is Nothing Then Set m_tables = New Collection
    m_tables.Add table
    Add = True
End Function