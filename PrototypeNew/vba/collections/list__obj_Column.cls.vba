VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "list__obj_Column"
Option Explicit

Private m_ObjectCollectionBase As obj_ObjectCollectionBase

Private Sub Class_Initialize()
    Set m_ObjectCollectionBase = New obj_ObjectCollectionBase
End Sub

' //
' // API
' //
Public Property Get Count() As Long
    Count = m_ObjectCollectionBase.Count
End Property

Public Property Get IsEmpty() As Boolean
    IsEmpty = m_ObjectCollectionBase.IsEmpty
End Property

Public Function Add(ByVal item As obj_Column) As Boolean

    m_ObjectCollectionBase.AddObject item
    Add = True
End Function

Public Property Get Item(ByVal oneBasedIndex As Long) As obj_Column
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

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = m_ObjectCollectionBase.NewEnum
End Property
