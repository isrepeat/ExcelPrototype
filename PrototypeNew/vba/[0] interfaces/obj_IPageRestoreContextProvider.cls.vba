VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_IPageRestoreContextProvider"
Option Explicit

' Страница восстанавливает обязательный Initialize-context из собственного payload
' до создания Worksheet и запуска обычного snapshot pipeline.
Public Function TryBuildRestoreContext( _
    ByVal snapshotXml As String, _
    ByRef outContext As Object _
) As Boolean
End Function
