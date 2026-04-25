VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_Config"
Option Explicit

Private m_Attr As String
Private m_Key As String
Private m_Value As String

' //
' // API
' //
Public Property Get Attr() As String
    Attr = m_Attr
End Property

Public Property Let Attr(ByVal value As String)
    m_Attr = VBA.CStr(value)
End Property

Public Property Get Key() As String
    Key = m_Key
End Property

Public Property Let Key(ByVal value As String)
    m_Key = VBA.CStr(value)
End Property

Public Property Get Value() As String
    Value = m_Value
End Property

Public Property Let Value(ByVal value As String)
    m_Value = VBA.CStr(value)
End Property
