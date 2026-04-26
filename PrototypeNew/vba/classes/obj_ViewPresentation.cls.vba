VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ViewPresentation"
Option Explicit

Private m_EffectiveVisible As Boolean
Private m_StyleName As String
Private m_SpanRows As Long
Private m_SpacerRowsAfter As Long

Private Sub Class_Initialize()
    m_EffectiveVisible = True
End Sub

Public Property Get EffectiveVisible() As Boolean
    EffectiveVisible = m_EffectiveVisible
End Property

Public Property Let EffectiveVisible(ByVal value As Boolean)
    m_EffectiveVisible = VBA.CBool(value)
End Property

Public Property Get StyleName() As String
    StyleName = m_StyleName
End Property

Public Property Let StyleName(ByVal value As String)
    m_StyleName = VBA.CStr(value)
End Property

Public Property Get SpanRows() As Long
    SpanRows = m_SpanRows
End Property

Public Property Let SpanRows(ByVal value As Long)
    If value < 0 Then
        m_SpanRows = 0
    Else
        m_SpanRows = VBA.CLng(value)
    End If
End Property

Public Property Get SpacerRowsAfter() As Long
    SpacerRowsAfter = m_SpacerRowsAfter
End Property

Public Property Let SpacerRowsAfter(ByVal value As Long)
    If value < 0 Then
        m_SpacerRowsAfter = 0
    Else
        m_SpacerRowsAfter = VBA.CLng(value)
    End If
End Property
