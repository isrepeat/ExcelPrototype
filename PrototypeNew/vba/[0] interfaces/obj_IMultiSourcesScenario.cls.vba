VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_IMultiSourcesScenario"
Option Explicit

Public Function Initialize( _
    ByVal page As obj_IPage, _
    ByVal configTable As obj_ConfigTable _
) As Boolean
End Function

Public Sub Dispose()
End Sub

Public Function RunPipeline( _
    Optional ByVal notifyChange As Boolean = True _
) As Boolean
End Function

Public Property Get OrderNoText() As String
End Property
