VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ViewPresentation"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean

Private m_EffectiveVisible As Boolean
Private m_StyleName As String
Private m_SpanRows As Long
Private m_SpacerRowsAfter As Long

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
    m_EffectiveVisible = True
End Sub
Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Dispose
    On Error GoTo 0
End Sub

' //
' // API
' //
Public Function Initialize() As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.m_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    On Error GoTo 0
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


