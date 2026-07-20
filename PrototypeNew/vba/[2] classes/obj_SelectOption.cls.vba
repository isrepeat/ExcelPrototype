VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_SelectOption"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean
Implements obj_IButtonGroupItem

Private m_Caption As String
Private m_Id As String
Private m_OnSelect As String
Private m_StyleName As String
Private m_Tags As Collection
Private m_States As Collection

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
    Set m_Tags = New Collection
    Set m_States = New Collection
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
    Set m_Tags = Nothing
    Set m_States = Nothing
    On Error GoTo 0
End Sub

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

Public Property Get StyleName() As String
    StyleName = m_StyleName
End Property

Public Property Let StyleName(ByVal value As String)
    m_StyleName = VBA.CStr(value)
End Property

Public Function AddTag(ByVal tagName As String) As Boolean
    Dim tagItem As Variant

    tagName = VBA.LCase$(VBA.Trim$(tagName))
    If VBA.Len(tagName) = 0 Then Exit Function
    If m_Tags Is Nothing Then Set m_Tags = New Collection
    For Each tagItem In m_Tags
        If VBA.StrComp(VBA.CStr(tagItem), tagName, VBA.vbBinaryCompare) = 0 Then
            AddTag = True
            Exit Function
        End If
    Next tagItem
    m_Tags.Add tagName
    AddTag = True
End Function

Public Property Get Tags() As Collection
    Set Tags = m_Tags
End Property

Public Function SetState(ByVal stateName As String, ByVal isActive As Boolean) As Boolean
    Dim stateItem As Variant
    Dim nextStates As Collection

    stateName = VBA.LCase$(VBA.Trim$(stateName))
    If VBA.Len(stateName) = 0 Then Exit Function
    If m_States Is Nothing Then Set m_States = New Collection

    If isActive Then
        For Each stateItem In m_States
            If VBA.StrComp(VBA.CStr(stateItem), stateName, VBA.vbBinaryCompare) = 0 Then
                SetState = True
                Exit Function
            End If
        Next stateItem
        m_States.Add stateName
    Else
        Set nextStates = New Collection
        For Each stateItem In m_States
            If VBA.StrComp(VBA.CStr(stateItem), stateName, VBA.vbBinaryCompare) <> 0 Then nextStates.Add stateItem
        Next stateItem
        Set m_States = nextStates
    End If
    SetState = True
End Function

Public Property Get States() As Collection
    Set States = m_States
End Property

' //
' // obj_IButtonGroupItem
' //
Private Property Get obj_IButtonGroupItem_Id() As String
    obj_IButtonGroupItem_Id = m_Id
End Property

Private Property Get obj_IButtonGroupItem_Caption() As String
    obj_IButtonGroupItem_Caption = m_Caption
End Property

Private Property Get obj_IButtonGroupItem_Tags() As Collection
    Set obj_IButtonGroupItem_Tags = m_Tags
End Property

Private Property Get obj_IButtonGroupItem_States() As Collection
    Set obj_IButtonGroupItem_States = m_States
End Property
