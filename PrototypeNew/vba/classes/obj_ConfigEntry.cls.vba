VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ConfigEntry"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean

Private m_Attr As String
Private m_Key As String
Private m_Value As String

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
End Sub
Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
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
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    Initialize = True
End Function
Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    On Error GoTo 0
End Sub

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

