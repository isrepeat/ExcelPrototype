VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_Row"
Option Explicit

Private m_CellValues() As String
Private m_CellCount As Long

Private Sub Class_Initialize()
    m_CellCount = 0
End Sub

' //
' // API
' //
Public Property Get CellCount() As Long
    CellCount = m_CellCount
End Property

Public Property Get Cells() As Collection
    Dim result As Collection
    Dim i As Long

    Set result = New Collection
    For i = 1 To m_CellCount
        result.Add m_CellValues(i)
    Next i

    Set Cells = result
End Property

Public Sub AddCell(ByVal value As Variant)
    If Not private_EnsureCapacity(m_CellCount + 1) Then Exit Sub

    m_CellCount = m_CellCount + 1
    m_CellValues(m_CellCount) = VBA.CStr(value)
End Sub

Public Function SetCell(ByVal oneBasedIndex As Long, ByVal value As Variant) As Boolean
    If oneBasedIndex <= 0 Then
        VBA.MsgBox "obj_Row: cell index must be greater than zero.", VBA.vbExclamation
        Exit Function
    End If

    If oneBasedIndex > m_CellCount Then
        If Not private_EnsureCapacity(oneBasedIndex) Then Exit Function
        m_CellCount = oneBasedIndex
    End If

    m_CellValues(oneBasedIndex) = VBA.CStr(value)
    SetCell = True
End Function

Public Function GetCell(ByVal oneBasedIndex As Long) As String
    If oneBasedIndex <= 0 Then Exit Function
    If oneBasedIndex > m_CellCount Then Exit Function

    GetCell = m_CellValues(oneBasedIndex)
End Function

Public Sub CopyToMatrixRow(ByRef targetMatrix As Variant, ByVal matrixRow As Long, ByVal columnCount As Long)
    Dim i As Long
    Dim maxCols As Long

    If matrixRow <= 0 Then Exit Sub
    If columnCount <= 0 Then Exit Sub

    maxCols = columnCount
    If m_CellCount < maxCols Then maxCols = m_CellCount

    For i = 1 To maxCols
        targetMatrix(matrixRow, i) = m_CellValues(i)
    Next i
End Sub

' //
' // Internal
' //
Private Function private_EnsureCapacity(ByVal requiredCount As Long) As Boolean
    Dim oldCount As Long
    Dim i As Long

    If requiredCount <= 0 Then
        private_EnsureCapacity = True
        Exit Function
    End If

    If requiredCount <= m_CellCount Then
        private_EnsureCapacity = True
        Exit Function
    End If

    oldCount = m_CellCount
    If oldCount = 0 Then
        ReDim m_CellValues(1 To requiredCount)
    Else
        ReDim Preserve m_CellValues(1 To requiredCount)
    End If

    If oldCount + 1 <= requiredCount Then
        For i = oldCount + 1 To requiredCount
            m_CellValues(i) = VBA.vbNullString
        Next i
    End If

    private_EnsureCapacity = True
End Function
