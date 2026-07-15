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

Public Function fn_BuildWhereContainsSql(ByVal sourceColumnHeader As String, ByVal valueText As String) As String
    sourceColumnHeader = VBA.Trim$(sourceColumnHeader)
    valueText = VBA.Trim$(valueText)
    If VBA.Len(sourceColumnHeader) = 0 Then Exit Function
    If VBA.Len(valueText) = 0 Then Exit Function

    fn_BuildWhereContainsSql = fn_QuoteSqlIdentifier(sourceColumnHeader) & " LIKE " & fn_QuoteSqlLiteral("%" & valueText & "%")
End Function

Public Function fn_BuildWhereNotBlankSql(ByVal sourceColumnHeader As String) As String
    Dim quotedHeader As String

    sourceColumnHeader = VBA.Trim$(sourceColumnHeader)
    If VBA.Len(sourceColumnHeader) = 0 Then Exit Function

    quotedHeader = fn_QuoteSqlIdentifier(sourceColumnHeader)
    fn_BuildWhereNotBlankSql = quotedHeader & " IS NOT NULL AND Trim(" & quotedHeader & ") <> ''"
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

Public Function fn_TryExtractWhereEqualsValue( _
    ByVal whereText As String, _
    ByRef outValue As String _
) As Boolean
    Dim eqPos As Long
    Dim rightPart As String

    outValue = VBA.vbNullString
    whereText = VBA.Trim$(whereText)
    If VBA.Len(whereText) = 0 Then Exit Function

    eqPos = VBA.InStr(1, whereText, "=", VBA.vbBinaryCompare)
    If eqPos <= 0 Then Exit Function

    rightPart = VBA.Trim$(VBA.Mid$(whereText, eqPos + 1))
    If VBA.Len(rightPart) = 0 Then Exit Function

    If VBA.Len(rightPart) >= 2 Then
        If VBA.Left$(rightPart, 1) = "'" And VBA.Right$(rightPart, 1) = "'" Then
            rightPart = VBA.Mid$(rightPart, 2, VBA.Len(rightPart) - 2)
            rightPart = VBA.Replace$(rightPart, "''", "'")
        End If
    End If

    outValue = rightPart
    fn_TryExtractWhereEqualsValue = True
End Function
