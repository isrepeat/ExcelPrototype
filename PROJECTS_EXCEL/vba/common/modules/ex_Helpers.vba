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
Public Function fn_TryParseInvariantNumber( _
    ByVal text As String, _
    ByRef number As Double _
) As Boolean
    Dim pattern As Object

    On Error GoTo InvalidNumber
    number = 0
    text = VBA.Trim$(text)
    Set pattern = VBA.CreateObject("VBScript.RegExp")
    pattern.Pattern = "^[+-]?([0-9]+([.][0-9]*)?|[.][0-9]+)([eE][+-]?[0-9]+)?$"
    If Not pattern.Test(text) Then Exit Function
    ' Val uses a decimal point regardless of Windows regional settings.
    number = VBA.Val(text)
    fn_TryParseInvariantNumber = True
    Exit Function
InvalidNumber:
    number = 0
End Function
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