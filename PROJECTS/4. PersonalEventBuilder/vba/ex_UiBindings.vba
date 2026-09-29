Option Explicit

Private shapeActions As Object

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_Reset()
    Set shapeActions = VBA.CreateObject("Scripting.Dictionary")
    shapeActions.CompareMode = VBA.vbTextCompare
End Sub

Public Sub fn_Register(ByVal shapeName As String, ByVal callbackName As String)
    If shapeActions Is Nothing Then fn_Reset
    shapeActions(shapeName) = callbackName
End Sub

Public Function fn_TryGetCallback(ByVal shapeName As String, ByRef outCallbackName As String) As Boolean
    outCallbackName = VBA.vbNullString
    If shapeActions Is Nothing Then Exit Function
    If Not shapeActions.Exists(shapeName) Then Exit Function
    outCallbackName = VBA.CStr(shapeActions(shapeName))
    fn_TryGetCallback = True
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------