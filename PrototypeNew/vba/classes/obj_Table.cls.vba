VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_Table"
Option Explicit

Private m_SectionTitle As String
Private m_Columns As Collection
Private m_Rows As Collection
Private m_IsInitialized As Boolean

Private Sub Class_Initialize()
    Set m_Columns = New Collection
    Set m_Rows = New Collection
End Sub

Public Function m_Init(ByVal rowCount As Long, ByVal columnCount As Long) As Boolean
    Dim i As Long
    Dim newColumn As obj_Column
    Dim newRow As obj_Row

    If rowCount <= 0 Then
        MsgBox "obj_Table: rowCount must be greater than zero.", vbExclamation
        Exit Function
    End If
    If columnCount <= 0 Then
        MsgBox "obj_Table: columnCount must be greater than zero.", vbExclamation
        Exit Function
    End If

    Set m_Columns = New Collection
    Set m_Rows = New Collection

    For i = 1 To columnCount
        Set newColumn = New obj_Column
        newColumn.Position = i
        newColumn.Name = "Col" & CStr(i)
        m_Columns.Add newColumn
    Next i

    For i = 1 To rowCount
        Set newRow = New obj_Row
        mp_FillRowWithBlanks newRow, columnCount
        m_Rows.Add newRow
    Next i

    m_IsInitialized = True
    m_Init = True
End Function

Public Property Get IsInitialized() As Boolean
    IsInitialized = m_IsInitialized
End Property

Public Property Get SectionTitle() As String
    SectionTitle = m_SectionTitle
End Property

Public Property Let SectionTitle(ByVal value As String)
    m_SectionTitle = CStr(value)
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

Public Function m_SetColumn(ByVal columnIndex As Long, ByVal tableColumn As obj_Column) As Boolean
    Dim targetColumn As obj_Column

    If Not m_IsInitialized Then
        MsgBox "obj_Table: call m_Init before setting columns.", vbExclamation
        Exit Function
    End If
    If tableColumn Is Nothing Then
        MsgBox "obj_Table: tableColumn is not specified.", vbExclamation
        Exit Function
    End If
    If columnIndex <= 0 Or columnIndex > m_Columns.Count Then
        MsgBox "obj_Table: column index is out of range.", vbExclamation
        Exit Function
    End If

    Set targetColumn = m_Columns(columnIndex)
    targetColumn.Name = tableColumn.Name
    If Len(targetColumn.Name) = 0 Then targetColumn.Name = "Col" & CStr(columnIndex)
    targetColumn.Position = columnIndex

    m_SetColumn = True
End Function

Public Function m_SetRow(ByVal rowIndex As Long, ByVal tableRow As obj_Row) As Boolean
    Dim targetRow As obj_Row
    Dim i As Long

    If Not m_IsInitialized Then
        MsgBox "obj_Table: call m_Init before setting rows.", vbExclamation
        Exit Function
    End If
    If tableRow Is Nothing Then
        MsgBox "obj_Table: tableRow is not specified.", vbExclamation
        Exit Function
    End If
    If rowIndex <= 0 Or rowIndex > m_Rows.Count Then
        MsgBox "obj_Table: row index is out of range.", vbExclamation
        Exit Function
    End If

    Set targetRow = m_Rows(rowIndex)

    For i = 1 To m_Columns.Count
        targetRow.m_SetCell i, tableRow.m_GetCell(i)
    Next i

    m_SetRow = True
End Function

Public Function m_SetCell(ByVal rowIndex As Long, ByVal columnIndex As Long, ByVal value As Variant) As Boolean
    Dim targetRow As obj_Row

    If Not m_IsInitialized Then
        MsgBox "obj_Table: call m_Init before setting cells.", vbExclamation
        Exit Function
    End If
    If rowIndex <= 0 Or rowIndex > m_Rows.Count Then
        MsgBox "obj_Table: row index is out of range.", vbExclamation
        Exit Function
    End If
    If columnIndex <= 0 Or columnIndex > m_Columns.Count Then
        MsgBox "obj_Table: column index is out of range.", vbExclamation
        Exit Function
    End If

    Set targetRow = m_Rows(rowIndex)
    m_SetCell = targetRow.m_SetCell(columnIndex, value)
End Function

Private Sub mp_FillRowWithBlanks(ByVal tableRow As obj_Row, ByVal columnCount As Long)
    Dim i As Long

    For i = 1 To columnCount
        tableRow.m_AddCell vbNullString
    Next i
End Sub
