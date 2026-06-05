VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "list__obj_String"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False

Private m_IsDisposed As Boolean
Private m_Items As Collection

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
    Set m_Items = New Collection
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Me.Dispose
    On Error GoTo 0
End Sub

' //
' // API
' //
Public Function Initialize() As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    m_IsDisposed = False
    Set m_Items = New Collection
    Initialize = Not m_Items Is Nothing
End Function

Public Sub Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True

    On Error Resume Next
    Set m_Items = Nothing
    On Error GoTo 0
End Sub

Public Property Get Count() As Long
    If m_Items Is Nothing Then Exit Property
    Count = m_Items.Count
End Property

Public Property Get IsEmpty() As Boolean
    If m_Items Is Nothing Then
        IsEmpty = True
        Exit Property
    End If

    IsEmpty = (m_Items.Count = 0)
End Property

Public Function Add(ByVal valueText As String) As Boolean
    If m_Items Is Nothing Then Set m_Items = New Collection
    m_Items.Add VBA.CStr(valueText)
    Add = True
End Function

Public Property Get Item(ByVal oneBasedIndex As Long) As String
    If m_Items Is Nothing Then Exit Property
    If oneBasedIndex <= 0 Or oneBasedIndex > m_Items.Count Then Exit Property

    Item = VBA.CStr(m_Items.Item(oneBasedIndex))
End Property

Public Function RemoveAt(ByVal oneBasedIndex As Long) As Boolean
    If m_Items Is Nothing Then Exit Function
    If oneBasedIndex <= 0 Or oneBasedIndex > m_Items.Count Then Exit Function

    m_Items.Remove oneBasedIndex
    RemoveAt = True
End Function

Public Sub Clear()
    Set m_Items = New Collection
End Sub

Public Property Get AsCollection() As Collection
    If m_Items Is Nothing Then Set m_Items = New Collection
    Set AsCollection = m_Items
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    If m_Items Is Nothing Then Set m_Items = New Collection
    Set NewEnum = m_Items.[_NewEnum]
End Property
