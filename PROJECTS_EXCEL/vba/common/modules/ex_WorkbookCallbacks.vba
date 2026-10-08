Attribute VB_Name = "ex_WorkbookCallbacks"
Option Explicit

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Resolve(ByVal configKey As String) As String
    Dim callbackName As String
    Dim parts As Variant
    Dim part As Variant

    If Not ex_Core.fn_TryGetWorkbookConfigValue(configKey, callbackName) Then
        GoTo Invalid
    End If
    callbackName = VBA.Trim$(callbackName)
    parts = VBA.Split(callbackName, ".")
    If UBound(parts) <> 1 Then
        GoTo Invalid
    End If
    For Each part In parts
        If VBA.Len(part) = 0 Or VBA.Len(part) > 31 Then
            GoTo Invalid
        End If
        If Not VBA.Left$(part, 1) Like "[A-Za-z]" Or part Like "*[!A-Za-z0-9_]*" Then
            GoTo Invalid
        End If
    Next part
    fn_Resolve = "'" & VBA.Replace(ThisWorkbook.Name, "'", "''") & "'!" & callbackName
    Exit Function
Invalid:
    VBA.Err.Raise VBA.vbObjectError + 2210, "ex_WorkbookCallbacks.fn_Resolve", _
        "Required configuration key is missing or invalid: " & configKey
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------