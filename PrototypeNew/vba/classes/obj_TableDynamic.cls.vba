VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_TableDynamic"
Option Explicit

Private m_SectionTitle As String
Private m_Columns As Collection
Private m_Rows As Collection

Private Sub Class_Initialize()
    Set m_Columns = New Collection
    Set m_Rows = New Collection
End Sub

' //
' // API
' //
Public Property Get SectionTitle() As String
    SectionTitle = m_SectionTitle
End Property

Public Property Let SectionTitle(ByVal value As String)
    m_SectionTitle = VBA.CStr(value)
End Property

Public Property Get ColumnCount() As Long
    ColumnCount = m_Columns.Count
End Property

Public Property Get RowCount() As Long
    RowCount = m_Rows.Count
End Property

Public Property Get Columns() As Collection
    Set Columns = m_Columns
End Property

Public Property Get Rows() As Collection
    Set Rows = m_Rows
End Property

Public Property Get HeaderText() As String
    Dim i As Long
    Dim colObj As obj_Column
    Dim joined As String

    For i = 1 To m_Columns.Count
        Set colObj = m_Columns(i)
        If i > 1 Then joined = joined & " | "
        joined = joined & colObj.Name
    Next i

    HeaderText = joined
End Property

Public Function AddColumn(ByVal tableColumn As obj_Column) As Boolean
    Dim newColumn As obj_Column

    If tableColumn Is Nothing Then
        VBA.MsgBox "obj_TableDynamic: column is not specified.", VBA.vbExclamation
        Exit Function
    End If

    Set newColumn = New obj_Column
    newColumn.Name = tableColumn.Name
    newColumn.Position = m_Columns.Count + 1

    If VBA.Len(newColumn.Name) = 0 Then
        newColumn.Name = "Col" & VBA.CStr(newColumn.Position)
    End If

    m_Columns.Add newColumn
    AddColumn = True
End Function

Public Function AddRow(ByVal tableRow As obj_Row) As Boolean
    Dim requiredCols As Long

    If tableRow Is Nothing Then
        VBA.MsgBox "obj_TableDynamic: row is not specified.", VBA.vbExclamation
        Exit Function
    End If

    requiredCols = tableRow.CellCount
    If requiredCols > m_Columns.Count Then
        If Not private_EnsureColumns(requiredCols) Then Exit Function
    End If

    m_Rows.Add tableRow
    AddRow = True
End Function

' //
' // Internal
' //
Private Function private_EnsureColumns(ByVal requiredCount As Long) As Boolean
    Dim i As Long
    Dim autoColumn As obj_Column

    If requiredCount <= m_Columns.Count Then
        private_EnsureColumns = True
        Exit Function
    End If

    For i = m_Columns.Count + 1 To requiredCount
        Set autoColumn = New obj_Column
        autoColumn.Position = i
        autoColumn.Name = "Col" & VBA.CStr(i)
        m_Columns.Add autoColumn
    Next i

    private_EnsureColumns = True
End Function
