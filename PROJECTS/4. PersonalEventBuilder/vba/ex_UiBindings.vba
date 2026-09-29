Option Explicit

Private shapeActions As Object

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    Set shapeActions = Nothing
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_Reset()
    Set shapeActions = VBA.CreateObject("Scripting.Dictionary")
    shapeActions.CompareMode = VBA.vbTextCompare
    ex_Core.fn_Diagnostic_WriteLog "UI_BINDINGS_RESET"
End Sub

Public Sub fn_Register(ByVal shapeName As String, ByVal callbackName As String)
    If shapeActions Is Nothing Then fn_Reset
    shapeActions(shapeName) = callbackName
    ex_Core.fn_Diagnostic_WriteLog "UI_BINDING_REGISTERED | Shape=" & shapeName & _
        " | Callback=" & callbackName
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