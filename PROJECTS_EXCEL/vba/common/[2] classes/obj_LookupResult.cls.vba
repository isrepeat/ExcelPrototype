VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_LookupResult"
Option Explicit

Implements obj_IUiTableSource
Implements obj_IUiSelectionSource

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_records As Collection
Private m_table As obj_UiRawTable
Private m_hasMore As Boolean

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
Public Property Get RowCount() As Long
    If Not m_records Is Nothing Then RowCount = m_records.Count
End Property

Public Property Get HasMore() As Boolean
    HasMore = m_hasMore
End Property

Public Property Get Records() As Collection
    Dim copy As New Collection
    Dim record As obj_LookupRecord

    If Not m_records Is Nothing Then
        For Each record In m_records
            copy.Add record
        Next record
    End If
    Set Records = copy
End Property

' //
' // Interface
' //
Private Property Get obj_IUiTableSource_TableCount() As Long
    If m_isInitialized And Not m_isDisposed Then obj_IUiTableSource_TableCount = 1
End Property

Private Function obj_IUiTableSource_GetTable(ByVal index As Long) As Object
    If m_isInitialized And Not m_isDisposed And index = 1 Then Set obj_IUiTableSource_GetTable = m_table
End Function

Private Function obj_IUiSelectionSource_GetItemAt(ByVal rowIndex As Long) As Object
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If rowIndex < 1 Or rowIndex > m_records.Count Then Exit Function
    Set obj_IUiSelectionSource_GetItemAt = m_records(rowIndex)
End Function

' //
' // API
' //
Public Function Initialize( _
    ByVal records As Collection, _
    ByVal columns As Collection, _
    ByVal hasMore As Boolean, _
    ByRef diagnostic As String _
) As Boolean
    Dim headers As Variant
    Dim values As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim column As Object
    Dim record As obj_LookupRecord
    Dim value As Variant

    If m_isInitialized Or m_isDisposed Then Exit Function
    On Error GoTo EH
    If records Is Nothing Or columns Is Nothing Then Exit Function
    If columns.Count = 0 Then
        diagnostic = "Lookup requires at least one display column."
        Exit Function
    End If
    Set m_records = New Collection
    ReDim headers(0 To columns.Count - 1)
    For columnIndex = 1 To columns.Count
        Set column = columns(columnIndex)
        headers(columnIndex - 1) = column("caption")
    Next columnIndex
    If records.Count > 0 Then ReDim values(1 To records.Count, 1 To columns.Count)
    For rowIndex = 1 To records.Count
        Set record = records(rowIndex)
        m_records.Add record
        For columnIndex = 1 To columns.Count
            Set column = columns(columnIndex)
            value = VBA.vbNullString
            record.TryGetValue column("name"), value
            values(rowIndex, columnIndex) = value
        Next columnIndex
    Next rowIndex
    Set m_table = New obj_UiRawTable
    If records.Count = 0 Then
        If Not m_table.InitializeEmpty(headers) Then Exit Function
    Else
        If Not m_table.Initialize(values, headers) Then Exit Function
    End If
    m_hasMore = hasMore
    m_isInitialized = True
    diagnostic = VBA.vbNullString
    Initialize = True
    Exit Function
EH:
    diagnostic = "Candidate projection: " & VBA.Err.Description
    Me.Dispose
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
    Set m_records = Nothing
    If Not m_table Is Nothing Then m_table.Dispose
    Set m_table = Nothing
End Sub