VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "list__obj_Row"
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

Public Function Add(ByVal rowItem As obj_Row) As Boolean

    m_ObjectCollectionBase.AddObject rowItem
    Add = True
End Function

Public Property Get Item(ByVal oneBasedIndex As Long) As obj_Row
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
