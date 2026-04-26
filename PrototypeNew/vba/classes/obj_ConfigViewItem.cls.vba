VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ConfigViewItem"
Option Explicit

Private m_Model As obj_Config
Private m_Presentation As obj_ViewPresentation

Private Sub Class_Initialize()
    Set m_Model = New obj_Config
    Set m_Presentation = New obj_ViewPresentation
End Sub

' //
' // API
' //
Public Property Get Model() As obj_Config
    Set Model = m_Model
End Property

Public Property Set Model(ByVal value As obj_Config)
    If value Is Nothing Then
        Set m_Model = New obj_Config
    Else
        Set m_Model = value
    End If
End Property

Public Property Get Presentation() As obj_ViewPresentation
    Set Presentation = m_Presentation
End Property

Public Property Set Presentation(ByVal value As obj_ViewPresentation)
    If value Is Nothing Then
        Set m_Presentation = New obj_ViewPresentation
    Else
        Set m_Presentation = value
    End If
End Property

Public Property Get Attr() As String
    Attr = m_Model.Attr
End Property

Public Property Let Attr(ByVal value As String)
    m_Model.Attr = VBA.CStr(value)
End Property

Public Property Get Key() As String
    Key = m_Model.Key
End Property

Public Property Let Key(ByVal value As String)
    m_Model.Key = VBA.CStr(value)
End Property

Public Property Get Value() As String
    Value = m_Model.Value
End Property

Public Property Let Value(ByVal value As String)
    m_Model.Value = VBA.CStr(value)
End Property
