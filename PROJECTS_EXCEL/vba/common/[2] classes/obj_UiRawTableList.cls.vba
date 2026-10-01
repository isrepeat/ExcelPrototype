VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_UiRawTableList"
Option Explicit

Implements obj_IUiTableSource

Private m_tables As Collection
Private m_isDisposed As Boolean

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

Private Function obj_IUiTableSource_GetTable(ByVal index As Long) As obj_UiRawTable
    If m_tables Is Nothing Then Exit Function
    If index <= 0 Or index > m_tables.Count Then Exit Function
    Set obj_IUiTableSource_GetTable = m_tables.Item(index)
End Function

' //
' // API
' //
Public Function Initialize() As Boolean
    m_isDisposed = False
    If m_tables Is Nothing Then Set m_tables = New Collection
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    Set m_tables = Nothing
End Sub

Public Function Add(ByVal table As obj_UiRawTable) As Boolean
    If table Is Nothing Then Exit Function
    If m_tables Is Nothing Then Set m_tables = New Collection
    m_tables.Add table
    Add = True
End Function