VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_IPage"
Option Explicit

Public Function Initialize( _
    ByVal ws As Worksheet, _
    Optional ByVal uiPath As String = VBA.vbNullString, _
    Optional ByVal pageId As String = VBA.vbNullString, _
    Optional ByVal Context As Object = Nothing _
) As Boolean
End Function

Public Sub Dispose(Optional ByVal deleteWorksheet As Boolean = True)
End Sub

Public Function RunPagePipeline() As Boolean
End Function

Public Function Render() As Boolean
End Function

Public Function UpdateUiPath( _
ByVal uiPath As String, _
Optional ByVal reason As String = VBA.vbNullString _
) As Boolean
End Function

Public Function GetPageBase() As obj_PageBase
End Function

Public Function GetPageId() As String
End Function

Public Function TryGetController(ByRef outController As Object) As Boolean
End Function

Public Function RegisterControl(ByVal controlKey As String, ByVal controlVm As Object) As Boolean
End Function

Public Function RegisterShapeRoute( _
  ByVal shapeName As String, _
  ByVal controlKey As String, _
  ByVal methodName As String, _
  Optional ByVal hasArg As Boolean = False, _
  Optional ByVal argValue As Variant _
) As Boolean
End Function

Public Function UnregisterControl(ByVal controlKey As String) As Boolean
End Function

Public Function ResetControlActions() As Boolean
End Function

Public Function DispatchShapeClick(ByVal shapeName As String) As Boolean
End Function

Public Function TryCollectSerializableControlSnapshots(ByRef outSnapshots As Collection) As Boolean
End Function

Public Function TryRestoreSerializableControlSnapshots(ByVal snapshots As Collection) As Boolean
End Function

Public Function TryGetRegisteredControls(ByRef outControlsByKey As Object) As Boolean
End Function

Public Function TryGetRegisteredControlByKey(ByVal controlKey As String, ByRef outControl As Object) As Boolean
End Function

Public Function TryGetRegisteredControlByName(ByVal controlName As String, ByRef outControl As Object) As Boolean
End Function
