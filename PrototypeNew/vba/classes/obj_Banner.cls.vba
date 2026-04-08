VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_Banner"
Option Explicit

Private m_Header As String
Private m_Message As String
Private m_Visible As Boolean

Public Property Get Header() As String
    Header = m_Header
End Property

Public Property Let Header(ByVal value As String)
    m_Header = CStr(value)
End Property

Public Property Get Message() As String
    Message = m_Message
End Property

Public Property Let Message(ByVal value As String)
    m_Message = CStr(value)
End Property

Public Property Get Visible() As Boolean
    Visible = m_Visible
End Property

Public Property Let Visible(ByVal value As Boolean)
    m_Visible = CBool(value)
End Property
