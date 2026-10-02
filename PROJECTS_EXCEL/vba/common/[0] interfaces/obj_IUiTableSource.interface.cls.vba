VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_IUiTableSource"
Option Explicit

' //
' // Interface
' //
Public Property Get TableCount() As Long
End Property

' Return Object to avoid a circular interface dependency with its implementing class.
' Callers validate the returned object by assigning it to obj_UiRawTable.
Public Function GetTable(ByVal index As Long) As Object
End Function