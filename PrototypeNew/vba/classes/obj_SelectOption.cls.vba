VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_SelectOption"
Option Explicit

Private m_Caption As String
Private m_Id As String
Private m_OnSelect As String

' //
' // API
' //
Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal value As String)
    m_Caption = VBA.CStr(value)
End Property

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(ByVal value As String)
    m_Id = VBA.CStr(value)
End Property

Public Property Get OnSelect() As String
    OnSelect = m_OnSelect
End Property

Public Property Let OnSelect(ByVal value As String)
    m_OnSelect = VBA.CStr(value)
End Property
