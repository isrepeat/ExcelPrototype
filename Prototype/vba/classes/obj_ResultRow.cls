VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ResultRow"
Option Explicit

Private m_RowIndex As Long
Private m_ValuesByAlias As Object
Private m_ValuesByMapKey As Object

Public Sub Initialize(ByVal rowIndex As Long)
    m_RowIndex = rowIndex
    Set m_ValuesByAlias = CreateObject("Scripting.Dictionary")
    m_ValuesByAlias.CompareMode = 1
    Set m_ValuesByMapKey = CreateObject("Scripting.Dictionary")
    m_ValuesByMapKey.CompareMode = 1
End Sub

Public Property Get RowIndex() As Long
    RowIndex = m_RowIndex
End Property

Public Sub SetValue(ByVal fieldAlias As String, ByVal mapKey As String, ByVal valueText As String)
    If m_ValuesByAlias Is Nothing Then Initialize m_RowIndex
    m_ValuesByAlias(Trim$(fieldAlias)) = valueText
    m_ValuesByMapKey(Trim$(mapKey)) = valueText
End Sub

Public Function HasAlias(ByVal fieldAlias As String) As Boolean
    If m_ValuesByAlias Is Nothing Then Exit Function
    HasAlias = m_ValuesByAlias.Exists(Trim$(fieldAlias))
End Function

Public Function GetByAlias(ByVal fieldAlias As String) As String
    GetByAlias = CStr(m_ValuesByAlias(Trim$(fieldAlias)))
End Function

Public Property Get Column(ByVal fieldAlias As String) As String
    Column = GetByAlias(fieldAlias)
End Property

Public Function HasMapKey(ByVal mapKey As String) As Boolean
    If m_ValuesByMapKey Is Nothing Then Exit Function
    HasMapKey = m_ValuesByMapKey.Exists(Trim$(mapKey))
End Function

Public Function GetByMapKey(ByVal mapKey As String) As String
    GetByMapKey = CStr(m_ValuesByMapKey(Trim$(mapKey)))
End Function
