VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_IUiControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' //
' // Interface
' //
Public Function Initialize() As Boolean
End Function

Public Sub Dispose()
End Sub

Public Function Configure(ByVal controlNode As Object) As Boolean
End Function

Public Function Measure(ByVal uiRenderContext As obj_UiRenderContext) As Range
End Function

Public Function Render(ByVal uiRenderContext As obj_UiRenderContext) As Boolean
End Function

Public Function HandleCellChange(ByVal target As Range) As Boolean
End Function