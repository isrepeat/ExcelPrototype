VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "list__obj_ConfigEntry"
Option Explicit

Private m_Base As obj_ObjectCollectionBase

Private Sub Class_Initialize()
    Set m_Base = New obj_ObjectCollectionBase
End Sub

' //
' // API
' //
Public Property Get Count() As Long
    Count = m_Base.Count
End Property

Public Property Get IsEmpty() As Boolean
    IsEmpty = m_Base.IsEmpty
End Property

Public Function Add(ByVal item As obj_ConfigEntry) As Boolean

    m_Base.AddObject item
    Add = True
End Function

Public Property Get Item(ByVal oneBasedIndex As Long) As obj_ConfigEntry
    Set Item = m_Base.ItemObject(oneBasedIndex)
End Property

Public Function RemoveAt(ByVal oneBasedIndex As Long) As Boolean
    m_Base.RemoveAt oneBasedIndex
    RemoveAt = True
End Function

Public Sub Clear()
    m_Base.Clear
End Sub

Public Property Get AsCollection() As Collection
    Set AsCollection = m_Base.AsCollection
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = m_Base.NewEnum
End Property
