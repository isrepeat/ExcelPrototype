VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_Table"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean

Private m_SectionTitle As String
Private m_Columns As list__obj_Column
Private m_Rows As list__obj_Row
Private m_IsInitialized As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
    Set m_Columns = New list__obj_Column
    Set m_Rows = New list__obj_Row
End Sub
Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Dispose
    On Error GoTo 0
End Sub

' //
' // API
' //
Public Function Initialize() As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    Initialize = True
End Function
Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Err.Clear
    Err.Clear
    Set m_Columns = Nothing
    Set m_Rows = Nothing
    On Error GoTo 0
End Sub

Public Function Init(ByVal rowCount As Long, ByVal columnCount As Long) As Boolean
    Dim i As Long
    Dim newColumn As obj_Column
    Dim newRow As obj_Row

    If rowCount <= 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "obj_Table: rowCount must be greater than zero."
#End If
        Exit Function
    End If
    If columnCount <= 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "obj_Table: columnCount must be greater than zero."
#End If
        Exit Function
    End If

    Set m_Columns = New list__obj_Column
    Set m_Rows = New list__obj_Row

    For i = 1 To columnCount
        Set newColumn = New obj_Column
        newColumn.Position = i
        newColumn.Name = "Col" & VBA.CStr(i)
        m_Columns.Add newColumn
    Next i

    For i = 1 To rowCount
        Set newRow = New obj_Row
        private_FillRowWithBlanks newRow, columnCount
        m_Rows.Add newRow
    Next i

    m_IsInitialized = True
    Init = True
End Function

Public Property Get IsInitialized() As Boolean
    IsInitialized = m_IsInitialized
End Property

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

Public Property Get Columns() As list__obj_Column
    Set Columns = m_Columns
End Property

Public Property Get Rows() As list__obj_Row
    Set Rows = m_Rows
End Property

Public Property Get HeaderText() As String
    Dim i As Long
    Dim colObj As obj_Column
    Dim joined As String

    For i = 1 To m_Columns.Count
        Set colObj = m_Columns.Item(i)
        If i > 1 Then joined = joined & " | "
        joined = joined & colObj.Name
    Next i

    HeaderText = joined
End Property

Public Function SetColumn(ByVal columnIndex As Long, ByVal tableColumn As obj_Column) As Boolean
    Dim targetColumn As obj_Column

    If Not m_IsInitialized Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "obj_Table: call Init before setting columns."
#End If
        Exit Function
    End If
    If tableColumn Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "obj_Table: tableColumn is not specified."
#End If
        Exit Function
    End If
    If columnIndex <= 0 Or columnIndex > m_Columns.Count Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "obj_Table: column index is out of range."
#End If
        Exit Function
    End If

    Set targetColumn = m_Columns.Item(columnIndex)
    targetColumn.Name = tableColumn.Name
    If VBA.Len(targetColumn.Name) = 0 Then targetColumn.Name = "Col" & VBA.CStr(columnIndex)
    targetColumn.Position = columnIndex

    SetColumn = True
End Function

Public Function SetRow(ByVal rowIndex As Long, ByVal tableRow As obj_Row) As Boolean
    Dim targetRow As obj_Row
    Dim i As Long

    If Not m_IsInitialized Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "obj_Table: call Init before setting rows."
#End If
        Exit Function
    End If
    If tableRow Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "obj_Table: tableRow is not specified."
#End If
        Exit Function
    End If
    If rowIndex <= 0 Or rowIndex > m_Rows.Count Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "obj_Table: row index is out of range."
#End If
        Exit Function
    End If

    Set targetRow = m_Rows.Item(rowIndex)

    For i = 1 To m_Columns.Count
        targetRow.SetCell i, tableRow.GetCell(i)
    Next i

    SetRow = True
End Function

Public Function SetCell(ByVal rowIndex As Long, ByVal columnIndex As Long, ByVal value As Variant) As Boolean
    Dim targetRow As obj_Row

    If Not m_IsInitialized Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "obj_Table: call Init before setting cells."
#End If
        Exit Function
    End If
    If rowIndex <= 0 Or rowIndex > m_Rows.Count Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "obj_Table: row index is out of range."
#End If
        Exit Function
    End If
    If columnIndex <= 0 Or columnIndex > m_Columns.Count Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "obj_Table: column index is out of range."
#End If
        Exit Function
    End If

    Set targetRow = m_Rows.Item(rowIndex)
    SetCell = targetRow.SetCell(columnIndex, value)
End Function

' //
' // Internal
' //
Private Sub private_FillRowWithBlanks(ByVal tableRow As obj_Row, ByVal columnCount As Long)
    Dim i As Long

    For i = 1 To columnCount
        tableRow.AddCell VBA.vbNullString
    Next i
End Sub

