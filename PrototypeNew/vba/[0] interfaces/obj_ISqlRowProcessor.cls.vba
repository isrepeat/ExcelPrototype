VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ISqlRowProcessor"
Option Explicit

Public Function Initialize( _
    ByVal inputTable As obj_TableDynamic, _
    ByVal sqlParams As obj_SqlParams _
) As Boolean
End Function

Public Function HandleRow( _
    ByVal row As obj_Row _
) As Boolean
End Function

Public Function BuildResult() As obj_TableDynamic
End Function

Public Sub Dispose()
End Sub
