VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ExtWorkbookCondition"
Option Explicit

' Одно условие структурированного запроса. Сейчас условия одной query
' объединяются через AND. Value не требуется для IsEmpty/IsNotEmpty.
Private m_ColumnName As String
Private m_Operation As en_ExtWorkbookQueryOp
Private m_Value As String
Private m_NormalizeValue As Boolean

Private Sub Class_Initialize()
    m_Operation = en_ExtWorkbookQueryOp.ExtQueryOpEquals
    m_NormalizeValue = True
End Sub

Public Property Get ColumnName() As String
    ColumnName = m_ColumnName
End Property

Public Property Let ColumnName(ByVal value As String)
    m_ColumnName = VBA.Trim$(value)
End Property

Public Property Get Operation() As en_ExtWorkbookQueryOp
    Operation = m_Operation
End Property

Public Property Let Operation(ByVal value As en_ExtWorkbookQueryOp)
    m_Operation = value
End Property

Public Property Get Value() As String
    Value = m_Value
End Property

Public Property Let Value(ByVal value As String)
    m_Value = VBA.CStr(value)
End Property

Public Property Get NormalizeValue() As Boolean
    NormalizeValue = m_NormalizeValue
End Property

Public Property Let NormalizeValue(ByVal value As Boolean)
    m_NormalizeValue = value
End Property

Public Function TryValidate(ByRef outError As String) As Boolean
    outError = VBA.vbNullString
    If VBA.Len(m_ColumnName) = 0 Then
        outError = "Condition column name is empty."
        Exit Function
    End If

    Select Case m_Operation
        Case en_ExtWorkbookQueryOp.ExtQueryOpEquals, _
             en_ExtWorkbookQueryOp.ExtQueryOpNotEquals, _
             en_ExtWorkbookQueryOp.ExtQueryOpContains, _
             en_ExtWorkbookQueryOp.ExtQueryOpStartsWith, _
             en_ExtWorkbookQueryOp.ExtQueryOpEndsWith, _
             en_ExtWorkbookQueryOp.ExtQueryOpIsEmpty, _
             en_ExtWorkbookQueryOp.ExtQueryOpIsNotEmpty, _
             en_ExtWorkbookQueryOp.ExtQueryOpGreaterThan, _
             en_ExtWorkbookQueryOp.ExtQueryOpLessThan
            TryValidate = True
        Case Else
            outError = "Condition operator is not supported: " & VBA.CStr(m_Operation) & "."
    End Select
End Function
