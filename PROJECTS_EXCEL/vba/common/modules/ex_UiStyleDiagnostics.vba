Attribute VB_Name = "ex_UiStyleDiagnostics"
Option Explicit

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_Raise( _
    ByVal messageKey As String, _
    Optional ByVal detail As String _
)
    Dim profile As String
    Dim key As String
    Dim message As String

    If Not ex_Core.fn_TryGetWorkbookConfigValue("ThisWorkbook::id", profile) Then
        VBA.Err.Raise 5, , "Required configuration key not found: ThisWorkbook::id"
    End If
    key = profile & "::text.Style" & messageKey
    If Not ex_Core.fn_TryGetWorkbookConfigValue(key, message) Then
        VBA.Err.Raise 5, , "Required configuration key not found: " & key
    End If
    VBA.Err.Raise 5, "UiStylePipeline", VBA.Replace$(message, "{value}", detail)
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------