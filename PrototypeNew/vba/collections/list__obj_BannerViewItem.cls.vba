VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "list__obj_BannerViewItem"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean

Private m_ObjectCollectionBase As obj_ObjectCollectionBase

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
    Set m_ObjectCollectionBase = New obj_ObjectCollectionBase
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
    Err.Clear
    Set m_ObjectCollectionBase = Nothing
    On Error GoTo 0
End Sub

Public Property Get Count() As Long
    Count = m_ObjectCollectionBase.Count
End Property

Public Property Get IsEmpty() As Boolean
    IsEmpty = m_ObjectCollectionBase.IsEmpty
End Property

Public Function Add(ByVal item As obj_BannerViewItem) As Boolean

    m_ObjectCollectionBase.AddObject item
    Add = True
End Function

Public Property Get Item(ByVal oneBasedIndex As Long) As obj_BannerViewItem
    Set Item = m_ObjectCollectionBase.ItemObject(oneBasedIndex)
End Property

Public Function RemoveAt(ByVal oneBasedIndex As Long) As Boolean
    m_ObjectCollectionBase.RemoveAt oneBasedIndex
    RemoveAt = True
End Function

Public Sub Clear()
    m_ObjectCollectionBase.Clear
End Sub

Public Property Get AsCollection() As Collection
    Set AsCollection = m_ObjectCollectionBase.AsCollection
End Property

