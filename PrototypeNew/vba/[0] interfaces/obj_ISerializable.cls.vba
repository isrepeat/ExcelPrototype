VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ISerializable"
Option Explicit

Public Function GetSerializableTypeRoot() As String
End Function

Public Function TrySerializeSnapshot(ByRef outSnapshotXml As String) As Boolean
End Function

Public Function TryDeserializeSnapshot(ByVal snapshotXml As String) As Boolean
End Function

Public Function TryRestoreState() As Boolean
End Function
