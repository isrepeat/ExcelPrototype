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

Public Function GetTable(ByVal index As Long) As obj_UiRawTable
End Function