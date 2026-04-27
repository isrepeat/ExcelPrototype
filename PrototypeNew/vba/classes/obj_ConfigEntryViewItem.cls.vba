VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ConfigEntryViewItem"
Option Explicit

Private m_ConfigEntry As obj_ConfigEntry
Private m_ViewPresentation As obj_ViewPresentation

Private Sub Class_Initialize()
    Set m_ViewPresentation = New obj_ViewPresentation
    Call Me.Initialize(Nothing)
End Sub

' //
' // API
' //
Public Function Initialize(ByVal value As obj_ConfigEntry) As Boolean
    If value Is Nothing Then
        Set m_ConfigEntry = New obj_ConfigEntry
    Else
        Set m_ConfigEntry = value
    End If

    Initialize = True
End Function

Public Property Get Model() As obj_ConfigEntry
    Set Model = m_ConfigEntry
End Property

Public Property Get Presentation() As obj_ViewPresentation
    Set Presentation = m_ViewPresentation
End Property

Public Property Set Presentation(ByVal value As obj_ViewPresentation)
    If value Is Nothing Then
        Set m_ViewPresentation = New obj_ViewPresentation
    Else
        Set m_ViewPresentation = value
    End If
End Property

Public Property Get Attr() As String
    Attr = m_ConfigEntry.Attr
End Property

Public Property Let Attr(ByVal value As String)
    m_ConfigEntry.Attr = VBA.CStr(value)
End Property

Public Property Get Key() As String
    Key = m_ConfigEntry.Key
End Property

Public Property Let Key(ByVal value As String)
    m_ConfigEntry.Key = VBA.CStr(value)
End Property

Public Property Get Value() As String
    Value = m_ConfigEntry.Value
End Property

Public Property Let Value(ByVal value As String)
    m_ConfigEntry.Value = VBA.CStr(value)
End Property
