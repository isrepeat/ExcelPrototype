VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ITableTransformer"
Option Explicit

' Получает таблицу предыдущего шага и возвращает отдельную таблицу,
' предназначенную для следующего transformer или итогового рендера.
Public Function Transform( _
    ByVal sourceTable As obj_TableDynamic, _
    ByVal configTable As obj_ConfigTable, _
    ByRef outTable As obj_TableDynamic _
) As Boolean
End Function

Public Sub Dispose()
End Sub
