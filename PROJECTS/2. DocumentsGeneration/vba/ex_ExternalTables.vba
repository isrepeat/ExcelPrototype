Option Explicit

Private Const AD_OPEN_STATIC As Long = 3
Private Const AD_LOCK_READ_ONLY As Long = 1

' --------------------------------------
' namespace ExternalTables {
' --------------------------------------
Public Function ex_ExternalTables_TryLookup( _
    ByVal connection As Object, _
    ByVal sourceCaption As String, _
    ByVal tableRef As String, _
    ByVal keyHeader As String, _
    ByVal keyValue As String, _
    ByVal resultHeader As String, _
    ByRef outValue As String _
) As Boolean
    Dim recordset As Object
    Dim sqlText As String

    On Error GoTo EH
    outValue = VBA.vbNullString
    sqlText = "SELECT TOP 2 [" & resultHeader & "] FROM " & tableRef & _
        " WHERE UCASE(TRIM(CSTR(IIF(ISNULL([" & keyHeader & _
        "]), '', [" & keyHeader & "])))) = '" & _
        ex_ExternalTables_EscapeSql(VBA.UCase$(ex_Helpers.private_Text_Normalize(keyValue))) & "'"
    ex_Helpers.LogDebug "Lookup SQL: " & sqlText

    Set recordset = VBA.CreateObject("ADODB.Recordset")
    recordset.Open sqlText, connection, AD_OPEN_STATIC, AD_LOCK_READ_ONLY

    If recordset.EOF Then
        ex_Helpers.LogError tableRef & ": value was not found | Key=" & keyHeader & _
            " | Value=" & keyValue
        VBA.MsgBox sourceCaption & " " & tableRef & ": value '" & keyValue & _
            "' was not found in column '" & keyHeader & "'.", _
            VBA.vbExclamation, "Document Generation"
        GoTo CleanExit
    End If
    outValue = ex_ExternalTables_ReadText(recordset, resultHeader)
    recordset.MoveNext
    If Not recordset.EOF Then
        ex_Helpers.LogError tableRef & ": multiple rows found | Key=" & keyHeader & _
            " | Value=" & keyValue
        VBA.MsgBox sourceCaption & " " & tableRef & ": multiple rows were found for '" & _
            keyValue & "'.", VBA.vbExclamation, "Document Generation"
        GoTo CleanExit
    End If
    If VBA.Len(outValue) = 0 Then
        ex_Helpers.LogError tableRef & ": result is empty | Column=" & resultHeader & _
            " | Key=" & keyValue
        VBA.MsgBox sourceCaption & " " & tableRef & ": column '" & resultHeader & _
            "' is empty for '" & keyValue & "'.", _
            VBA.vbExclamation, "Document Generation"
        GoTo CleanExit
    End If

    ex_ExternalTables_TryLookup = True

CleanExit:
    On Error Resume Next
    If Not recordset Is Nothing Then recordset.Close
    Set recordset = Nothing
    On Error GoTo 0
    Exit Function
EH:
    ex_Helpers.LogError "Failed to read " & tableRef & " | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
    VBA.MsgBox "Failed to read " & sourceCaption & " " & tableRef & ": [" & _
        VBA.CStr(Err.Number) & "] " & Err.Description, _
        VBA.vbExclamation, "Document Generation"
    Resume CleanExit
End Function

' Универсальные операции с внешними Excel-таблицами через ADO.
Public Function ex_ExternalTables_TryOpenConnection( _
    ByVal workbookPath As String, _
    ByVal sourceCaption As String, _
    ByRef outConnection As Object _
) As Boolean
    On Error GoTo EH
    Set outConnection = VBA.CreateObject("ADODB.Connection")
    outConnection.Open _
        "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & workbookPath & ";" & _
        "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
    ex_ExternalTables_TryOpenConnection = True
    Exit Function
EH:
    ex_Helpers.LogError "Failed to open " & sourceCaption & " | Number=" & _
        VBA.CStr(Err.Number) & " | Description=" & Err.Description
    VBA.MsgBox "Failed to open " & sourceCaption & ": [" & _
        VBA.CStr(Err.Number) & "] " & Err.Description, _
        VBA.vbExclamation, "Document Generation"
End Function

Public Function ex_ExternalTables_EscapeSql(ByVal valueText As String) As String
    ex_ExternalTables_EscapeSql = VBA.Replace$(valueText, "'", "''")
End Function

Public Function ex_ExternalTables_ReadText( _
    ByVal recordset As Object, _
    ByVal fieldName As String _
) As String
    If VBA.IsNull(recordset.Fields(fieldName).Value) Then Exit Function
    ex_ExternalTables_ReadText = ex_Helpers.private_Text_Normalize( _
        VBA.CStr(recordset.Fields(fieldName).Value))
End Function
' --------------------------------------
' } // namespace ExternalTables
' --------------------------------------