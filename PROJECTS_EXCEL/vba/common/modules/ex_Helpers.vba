Option Explicit

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace RTTI {
' --------------------------------------
Public Function fn_RTTI_IsDictionary(ByVal sourceObject As Object) As Boolean
    If sourceObject Is Nothing Then Exit Function
    fn_RTTI_IsDictionary = (TypeName(sourceObject) = "Dictionary" Or _
        TypeName(sourceObject) = "Scripting.Dictionary")
End Function
' --------------------------------------
' } // namespace RTTI
' --------------------------------------