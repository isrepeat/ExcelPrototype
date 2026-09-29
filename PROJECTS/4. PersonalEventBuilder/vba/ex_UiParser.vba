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
Public Function fn_TryLoadPage(ByVal xamlPath As String, ByRef outPage As Object) As Boolean
    Dim document As Object

    Set outPage = Nothing
    Set document = VBA.CreateObject("Msxml2.DOMDocument.6.0")
    document.async = False
    If Not document.Load(xamlPath) Then
        VBA.MsgBox "Cannot load XAML: " & document.parseError.reason, _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    Set outPage = document
    fn_TryLoadPage = True
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------