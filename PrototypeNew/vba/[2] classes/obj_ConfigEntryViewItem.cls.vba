VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ConfigEntryViewItem"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean

Private m_ConfigEntry As obj_ConfigEntry
Private m_ViewPresentation As obj_ViewPresentation

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
    Set m_ViewPresentation = New obj_ViewPresentation
    Call Me.Initialize(Nothing)
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
Public Function Initialize(ByVal value As obj_ConfigEntry) As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    If value Is Nothing Then
        Set m_ConfigEntry = New obj_ConfigEntry
    Else
        Set m_ConfigEntry = value
    End If

    Initialize = True
End Function
Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Err.Clear
    Err.Clear
    Set m_ConfigEntry = Nothing
    Set m_ViewPresentation = Nothing
    On Error GoTo 0
End Sub

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

