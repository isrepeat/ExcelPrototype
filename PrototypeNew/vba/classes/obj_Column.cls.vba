VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_Column"
Option Explicit

Private m_Name As String
Private m_Position As Long

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(ByVal value As String)
    m_Name = Trim$(value)
End Property

Public Property Get Position() As Long
    Position = m_Position
End Property

Public Property Let Position(ByVal value As Long)
    If value > 0 Then
        m_Position = value
    Else
        m_Position = 0
    End If
End Property
