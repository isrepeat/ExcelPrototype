VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_TableDynamic"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean

Private m_SectionTitle As String
Private m_Columns As list__obj_Column
Private m_Rows As list__obj_Row

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
    Set m_Columns = New list__obj_Column
    Set m_Rows = New list__obj_Row
End Sub
Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
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
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    Initialize = True
End Function
Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
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

Public Function AddColumn(ByVal tableColumn As obj_Column) As Boolean
    Dim newColumn As obj_Column

    If tableColumn Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "obj_TableDynamic: column is not specified."
#End If
        Exit Function
    End If

    Set newColumn = New obj_Column
    newColumn.Name = tableColumn.Name
    newColumn.Position = m_Columns.Count + 1

    If VBA.Len(newColumn.Name) = 0 Then
        newColumn.Name = "Col" & VBA.CStr(newColumn.Position)
    End If

    AddColumn = m_Columns.Add(newColumn)
End Function

Public Function AddRow(ByVal tableRow As obj_Row) As Boolean
    Dim requiredCols As Long

    If tableRow Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.m_Diagnostic_LogError "obj_TableDynamic: row is not specified."
#End If
        Exit Function
    End If

    requiredCols = tableRow.CellCount
    If requiredCols > m_Columns.Count Then
        If Not private_EnsureColumns(requiredCols) Then Exit Function
    End If

    AddRow = m_Rows.Add(tableRow)
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

