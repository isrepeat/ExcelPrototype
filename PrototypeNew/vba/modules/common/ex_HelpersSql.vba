Attribute VB_Name = "ex_HelpersSql"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:ex_HelpersSql.fn_Module_Dispose"
#End If
End Sub

' //
' // API
' //
Public Function fn_BuildWhereEqualsSql(ByVal sourceColumnHeader As String, ByVal valueText As String) As String
    sourceColumnHeader = VBA.Trim$(sourceColumnHeader)
    If VBA.Len(sourceColumnHeader) = 0 Then Exit Function

    fn_BuildWhereEqualsSql = fn_QuoteSqlIdentifier(sourceColumnHeader) & " = " & fn_QuoteSqlLiteral(valueText)
End Function

Public Function fn_QuoteSqlIdentifier(ByVal valueText As String) As String
    valueText = VBA.Trim$(valueText)
    If VBA.Len(valueText) >= 2 Then
        If VBA.Left$(valueText, 1) = "[" And VBA.Right$(valueText, 1) = "]" Then
            valueText = VBA.Mid$(valueText, 2, VBA.Len(valueText) - 2)
        End If
    End If

    fn_QuoteSqlIdentifier = "[" & VBA.Replace$(valueText, "]", "]]") & "]"
End Function

Public Function fn_QuoteSqlLiteral(ByVal valueText As String) As String
    fn_QuoteSqlLiteral = "'" & VBA.Replace$(VBA.CStr(valueText), "'", "''") & "'"
End Function
