VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_IPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function Initialize(ByVal profileId As String) As Boolean
End Function

Public Function Render() As Boolean
End Function

Public Function HandleCellChange(ByVal target As Range) As Boolean
End Function

Public Sub Dispose()
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------