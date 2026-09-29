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
    ex_Core.fn_Diagnostic_WriteLog "XAML_LOAD_STARTED | Path=" & xamlPath
    Set document = VBA.CreateObject("Msxml2.DOMDocument.6.0")
    document.async = False
    If Not document.Load(xamlPath) Then
        ex_Core.fn_Diagnostic_WriteLog "XAML_LOAD_ERROR | Path=" & xamlPath & _
            " | Description=" & document.parseError.reason
        VBA.MsgBox "Cannot load XAML: " & document.parseError.reason, _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    Set outPage = document
    fn_TryLoadPage = True
    ex_Core.fn_Diagnostic_WriteLog "XAML_LOAD_COMPLETED | Path=" & xamlPath
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------