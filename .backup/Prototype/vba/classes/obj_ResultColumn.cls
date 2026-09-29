VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ResultColumn"
Option Explicit

Private m_Alias As String
Private m_MapKey As String
Private m_Value As String

Public Sub Initialize(ByVal fieldAlias As String, ByVal mapKey As String, ByVal valueText As String)
    m_Alias = Trim$(fieldAlias)
    m_MapKey = Trim$(mapKey)
    m_Value = valueText
End Sub

Public Property Get Alias() As String
    Alias = m_Alias
End Property

Public Property Get Name() As String
    Name = m_Alias
End Property

Public Property Get MapKey() As String
    MapKey = m_MapKey
End Property

Public Property Get Value() As String
    Value = m_Value
End Property
