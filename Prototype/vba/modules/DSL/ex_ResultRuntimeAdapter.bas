Attribute VB_Name = "ex_ResultRuntimeAdapter"
Option Explicit

Private Const ERR_SOURCE As String = "ex_ScriptDSL"

Public Sub m_BuildRuntimeContext( _
    ByVal resultTables As Collection, _
    ByRef outTablesByRef As Object, _
    ByRef outFieldsByTable As Object _
)
    Dim i As Long
    Dim tableObj As obj_ResultTable
    Dim fieldMap As Object
    Dim tableFields As Object
    Dim aliasKey As Variant
    Dim mapKey As String

    Set outTablesByRef = CreateObject("Scripting.Dictionary")
    outTablesByRef.CompareMode = 1

    Set outFieldsByTable = CreateObject("Scripting.Dictionary")
    outFieldsByTable.CompareMode = 1

    If resultTables Is Nothing Then Exit Sub

    For i = 1 To resultTables.Count
        Set tableObj = resultTables(i)
        If tableObj Is Nothing Then GoTo ContinueTable

        If Not outTablesByRef.Exists(tableObj.TableRef) Then
            outTablesByRef.Add tableObj.TableRef, tableObj
        End If

        Set tableFields = CreateObject("Scripting.Dictionary")
        tableFields.CompareMode = 1

        Set fieldMap = tableObj.FieldMapByAlias
        For Each aliasKey In fieldMap.Keys
            mapKey = CStr(fieldMap(aliasKey))
            tableFields(mapKey) = True
        Next aliasKey

        If outFieldsByTable.Exists(tableObj.TableRef) Then
            outFieldsByTable.Remove tableObj.TableRef
        End If
        outFieldsByTable.Add tableObj.TableRef, tableFields
ContinueTable:
    Next i
End Sub

Public Function m_TryGetRowsForTableRef( _
    ByVal tablesByRef As Object, _
    ByVal tableRef As String, _
    ByRef outRows As Collection _
) As Boolean
    Dim tableObj As obj_ResultTable

    If tablesByRef Is Nothing Then Exit Function
    If Not tablesByRef.Exists(tableRef) Then Exit Function

    Set tableObj = tablesByRef(tableRef)
    If tableObj Is Nothing Then Exit Function

    Set outRows = tableObj.Rows
    m_TryGetRowsForTableRef = True
End Function

Public Function m_TryResolveConditionTokenForValidation( _
    ByVal tokenText As String, _
    ByVal currentTableRef As String, _
    ByVal currentRowVar As String, _
    ByVal allowedTableFields As Object, _
    ByRef outResolvedTableRef As String, _
    ByRef outResolvedMapKey As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim rowVarName As String
    Dim fieldAlias As String
    Dim rowSelector As String
    Dim sourceAlias As String
    Dim tableAlias As String
    Dim tableRef As String
    Dim tableMemberName As String
    Dim mapTableAlias As String
    Dim mapFieldAlias As String
    Dim mapSourceAlias As String
    Dim rowIndex As Long
    Dim mapKey As String

    tokenText = Trim$(tokenText)
    If Len(tokenText) = 0 Then
        outErrorText = "Field reference is empty."
        Exit Function
    End If

    If ex_obj_ResultRowDsl.m_TryParseRowColumnRef(tokenText, rowVarName, fieldAlias) Then
        If Len(currentRowVar) = 0 Then
            outErrorText = "Row variable '" & rowVarName & "' is not available in this scope."
            Exit Function
        End If
        If StrComp(rowVarName, currentRowVar, vbTextCompare) <> 0 Then
            outErrorText = "Unknown row variable '" & rowVarName & "'. Expected '" & currentRowVar & "'."
            Exit Function
        End If
        If Len(currentTableRef) = 0 Then
            outErrorText = "Current table scope is not defined for row variable '" & rowVarName & "'."
            Exit Function
        End If
        If Not mp_TryResolveMapKeyByFieldAlias(allowedTableFields, currentTableRef, fieldAlias, outResolvedMapKey, outErrorText) Then Exit Function
        outResolvedTableRef = currentTableRef
        m_TryResolveConditionTokenForValidation = True
        Exit Function
    End If

    If ex_obj_ResultTableDsl.m_TryParseSpecialRowColumnRef(tokenText, sourceAlias, tableAlias, rowSelector, fieldAlias) Then
        tableRef = sourceAlias & ".Sheet[" & tableAlias & "]"
        If allowedTableFields Is Nothing Or Not allowedTableFields.Exists(tableRef) Then
            outErrorText = "Unknown table reference in reference '" & tokenText & "'."
            Exit Function
        End If
        If Not mp_TryResolveMapKeyByFieldAlias(allowedTableFields, tableRef, fieldAlias, outResolvedMapKey, outErrorText) Then Exit Function
        outResolvedTableRef = tableRef
        m_TryResolveConditionTokenForValidation = True
        Exit Function
    End If

    If ex_obj_ResultTableDsl.m_TryParseTableScalarRef(tokenText, sourceAlias, tableAlias, tableMemberName) Then
        tableRef = sourceAlias & ".Sheet[" & tableAlias & "]"
        If allowedTableFields Is Nothing Or Not allowedTableFields.Exists(tableRef) Then
            outErrorText = "Unknown table reference in reference '" & tokenText & "'."
            Exit Function
        End If
        outResolvedMapKey = vbNullString
        outResolvedTableRef = tableRef
        m_TryResolveConditionTokenForValidation = True
        Exit Function
    End If

    If mp_TryParseMapKey(tokenText, mapTableAlias, mapFieldAlias, mapSourceAlias) Then
        tableRef = mapSourceAlias & ".Sheet[" & mapTableAlias & "]"
        If allowedTableFields Is Nothing Or Not allowedTableFields.Exists(tableRef) Then
            outErrorText = "Unknown table reference in reference '" & tokenText & "'."
            Exit Function
        End If
        If Not allowedTableFields(tableRef).Exists(tokenText) Then
            outErrorText = "Field reference '" & tokenText & "' is not configured."
            Exit Function
        End If
        outResolvedMapKey = tokenText
        outResolvedTableRef = tableRef
        m_TryResolveConditionTokenForValidation = True
        Exit Function
    End If

    If ex_obj_ResultTableDsl.m_TryParseFullRowColumnRef(tokenText, sourceAlias, tableAlias, rowIndex, fieldAlias) Then
        tableRef = sourceAlias & ".Sheet[" & tableAlias & "]"
        mapKey = m_BuildMapKey(sourceAlias, tableAlias, fieldAlias)

        If rowIndex < 0 Then
            outErrorText = "Row index must be >= 0 in reference '" & tokenText & "'."
            Exit Function
        End If
        If allowedTableFields Is Nothing Or Not allowedTableFields.Exists(tableRef) Then
            outErrorText = "Unknown table reference in reference '" & tokenText & "'."
            Exit Function
        End If
        If Not allowedTableFields(tableRef).Exists(mapKey) Then
            outErrorText = "Field reference '" & tokenText & "' is not configured."
            Exit Function
        End If

        outResolvedMapKey = mapKey
        outResolvedTableRef = tableRef
        m_TryResolveConditionTokenForValidation = True
        Exit Function
    End If

    outErrorText = "Unsupported field reference '" & tokenText & "'. Use <rowVar>.column[FieldAlias], Source.Sheet[TableAlias].row[N].column[FieldAlias], Source.Sheet[TableAlias].lastRow.column[FieldAlias], Source.Sheet[TableAlias].prevRow.column[FieldAlias], Source.Sheet[TableAlias].count, or Source.Sheet[TableAlias].rowCount."
End Function

Public Function m_TryResolveRuntimeValue( _
    ByVal tokenText As String, _
    ByVal currentTableRef As String, _
    ByVal currentRowVar As String, _
    ByVal currentRowRef As Object, _
    ByVal tablesByRef As Object, _
    ByRef outValue As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim rowVarName As String
    Dim fieldAlias As String
    Dim rowSelector As String
    Dim sourceAlias As String
    Dim tableAlias As String
    Dim tableRef As String
    Dim tableMemberName As String
    Dim rowIndex As Long
    Dim targetTable As obj_ResultTable
    Dim targetRowRef As obj_ResultRow

    tokenText = Trim$(tokenText)
    If Len(tokenText) = 0 Then
        outErrorText = "Field reference is empty."
        Exit Function
    End If

    If ex_obj_ResultRowDsl.m_TryParseRowColumnRef(tokenText, rowVarName, fieldAlias) Then
        If Len(currentRowVar) = 0 Then
            outErrorText = "Row variable '" & rowVarName & "' is not available in this scope."
            Exit Function
        End If
        If StrComp(rowVarName, currentRowVar, vbTextCompare) <> 0 Then
            outErrorText = "Unknown row variable '" & rowVarName & "'. Expected '" & currentRowVar & "'."
            Exit Function
        End If
        If currentRowRef Is Nothing Then
            outErrorText = "Current row is not available for variable '" & rowVarName & "'."
            Exit Function
        End If
        If Not currentRowRef.HasAlias(fieldAlias) Then
            outErrorText = "Unknown field alias '" & fieldAlias & "' for table '" & currentTableRef & "'."
            Exit Function
        End If

        outValue = CStr(currentRowRef.Column(fieldAlias))
        m_TryResolveRuntimeValue = True
        Exit Function
    End If

    If ex_obj_ResultTableDsl.m_TryParseSpecialRowColumnRef(tokenText, sourceAlias, tableAlias, rowSelector, fieldAlias) Then
        tableRef = sourceAlias & ".Sheet[" & tableAlias & "]"
        If tablesByRef Is Nothing Or Not tablesByRef.Exists(tableRef) Then
            outErrorText = "Table '" & tableRef & "' is not available in current result."
            Exit Function
        End If

        Set targetTable = tablesByRef(tableRef)
        If Not ex_obj_ResultTableDsl.m_TryGetRowBySelector(targetTable, rowSelector, targetRowRef) Then
            If StrComp(rowSelector, "prevrow", vbTextCompare) = 0 Then
                outErrorText = "Table '" & tableRef & "' has fewer than 2 rows; cannot resolve prevRow."
            Else
                outErrorText = "Table '" & tableRef & "' has no rows; cannot resolve lastRow."
            End If
            Exit Function
        End If
        If Not targetRowRef.HasAlias(fieldAlias) Then
            outErrorText = "Field alias '" & fieldAlias & "' is not available at " & rowSelector & " of table '" & tableRef & "'."
            Exit Function
        End If

        outValue = CStr(targetRowRef.Column(fieldAlias))
        m_TryResolveRuntimeValue = True
        Exit Function
    End If

    If ex_obj_ResultTableDsl.m_TryParseTableScalarRef(tokenText, sourceAlias, tableAlias, tableMemberName) Then
        tableRef = sourceAlias & ".Sheet[" & tableAlias & "]"
        If tablesByRef Is Nothing Or Not tablesByRef.Exists(tableRef) Then
            outErrorText = "Table '" & tableRef & "' is not available in current result."
            Exit Function
        End If

        Set targetTable = tablesByRef(tableRef)
        Select Case LCase$(tableMemberName)
            Case "count"
                outValue = CStr(targetTable.Count)
            Case "rowcount"
                outValue = CStr(targetTable.RowCount)
            Case Else
                outErrorText = "Unsupported table property '" & tableMemberName & "' in reference '" & tokenText & "'."
                Exit Function
        End Select

        m_TryResolveRuntimeValue = True
        Exit Function
    End If

    If ex_obj_ResultTableDsl.m_TryParseFullRowColumnRef(tokenText, sourceAlias, tableAlias, rowIndex, fieldAlias) Then
        tableRef = sourceAlias & ".Sheet[" & tableAlias & "]"
        If tablesByRef Is Nothing Or Not tablesByRef.Exists(tableRef) Then
            outErrorText = "Table '" & tableRef & "' is not available in current result."
            Exit Function
        End If

        Set targetTable = tablesByRef(tableRef)
        If rowIndex < 0 Then
            outErrorText = "Row index must be >= 0 in reference '" & tokenText & "'."
            Exit Function
        End If

        Set targetRowRef = targetTable.Row(rowIndex)
        If targetRowRef Is Nothing Then
            outErrorText = "Row '" & CStr(rowIndex) & "' is out of range for table '" & tableRef & "'."
            Exit Function
        End If
        If Not targetRowRef.HasAlias(fieldAlias) Then
            outErrorText = "Field alias '" & fieldAlias & "' is not available at row '" & CStr(rowIndex) & "'."
            Exit Function
        End If

        outValue = CStr(targetRowRef.Column(fieldAlias))
        m_TryResolveRuntimeValue = True
        Exit Function
    End If

    outErrorText = "Unsupported field reference '" & tokenText & "'. Use <rowVar>.column[FieldAlias], Source.Sheet[TableAlias].row[N].column[FieldAlias], Source.Sheet[TableAlias].lastRow.column[FieldAlias], Source.Sheet[TableAlias].prevRow.column[FieldAlias], Source.Sheet[TableAlias].count, or Source.Sheet[TableAlias].rowCount."
End Function

Public Function m_TryParseMacroArg(ByVal argText As String, ByRef outArgSpec As Object) As Boolean
    Dim literalText As String
    Dim integerLiteral As Long
    Dim sourceAlias As String
    Dim tableAlias As String
    Dim rowIndex As Long
    Dim rowSelector As String
    Dim tableRef As String
    Dim fieldAlias As String
    Dim rowVarName As String

    Set outArgSpec = CreateObject("Scripting.Dictionary")
    outArgSpec.CompareMode = 1

    If mp_TryParseQuotedString(argText, literalText) Then
        outArgSpec("Kind") = "string"
        outArgSpec("Value") = literalText
        m_TryParseMacroArg = True
        Exit Function
    End If

    If ex_XmlCore.m_TryParseLong(Trim$(argText), integerLiteral) Then
        outArgSpec("Kind") = "number"
        outArgSpec("Value") = CLng(integerLiteral)
        m_TryParseMacroArg = True
        Exit Function
    End If

    If ex_ScriptParserCore.m_IsIdentifier(argText) Then
        outArgSpec("Kind") = "varref"
        outArgSpec("Name") = Trim$(argText)
        m_TryParseMacroArg = True
        Exit Function
    End If

    If ex_obj_ResultTableDsl.m_TryParseTableSpecialRowRef(argText, sourceAlias, tableAlias, rowSelector) Then
        tableRef = sourceAlias & ".Sheet[" & tableAlias & "]"
        outArgSpec("Kind") = "rowref"
        outArgSpec("TableRef") = tableRef
        outArgSpec("RowSelector") = rowSelector
        m_TryParseMacroArg = True
        Exit Function
    End If

    If ex_obj_ResultTableDsl.m_TryParseTableRowRef(argText, sourceAlias, tableAlias, rowIndex) Then
        tableRef = sourceAlias & ".Sheet[" & tableAlias & "]"
        outArgSpec("Kind") = "rowref"
        outArgSpec("TableRef") = tableRef
        outArgSpec("RowIndex") = CLng(rowIndex)
        m_TryParseMacroArg = True
        Exit Function
    End If

    If ex_obj_ResultTableDsl.m_TryParseSpecialRowColumnRef(argText, sourceAlias, tableAlias, rowSelector, fieldAlias) Then
        tableRef = sourceAlias & ".Sheet[" & tableAlias & "]"
        outArgSpec("Kind") = "cellref"
        outArgSpec("TableRef") = tableRef
        outArgSpec("RowSelector") = rowSelector
        outArgSpec("FieldAlias") = fieldAlias
        m_TryParseMacroArg = True
        Exit Function
    End If

    If ex_obj_ResultTableDsl.m_TryParseFullRowColumnRef(argText, sourceAlias, tableAlias, rowIndex, fieldAlias) Then
        tableRef = sourceAlias & ".Sheet[" & tableAlias & "]"
        outArgSpec("Kind") = "cellref"
        outArgSpec("TableRef") = tableRef
        outArgSpec("RowIndex") = CLng(rowIndex)
        outArgSpec("FieldAlias") = fieldAlias
        m_TryParseMacroArg = True
        Exit Function
    End If

    If ex_obj_ResultRowDsl.m_TryParseRowColumnRef(argText, rowVarName, fieldAlias) Then
        outArgSpec("Kind") = "scopecellref"
        outArgSpec("RowVar") = rowVarName
        outArgSpec("FieldAlias") = fieldAlias
        m_TryParseMacroArg = True
        Exit Function
    End If
End Function

Public Function m_ValidateMacroArgSpec( _
    ByVal argSpec As Object, _
    ByVal scopeVarTypes As Object, _
    ByVal allowedTableFields As Object, _
    ByRef outErrorText As String _
) As Boolean
    Dim tableRef As String
    Dim resolvedMapKey As String

    If argSpec Is Nothing Then Exit Function

    Select Case LCase$(CStr(argSpec("Kind")))
        Case "number"
            ' Numeric literal does not need additional validation.

        Case "varref"
            If scopeVarTypes Is Nothing Then
                outErrorText = "callMacro variable '" & CStr(argSpec("Name")) & "' is not available in this scope."
                Exit Function
            End If
            If Not scopeVarTypes.Exists(CStr(argSpec("Name"))) Then
                outErrorText = "callMacro variable '" & CStr(argSpec("Name")) & "' is not available in this scope."
                Exit Function
            End If

        Case "rowref"
            tableRef = CStr(argSpec("TableRef"))
            If allowedTableFields Is Nothing Or Not allowedTableFields.Exists(tableRef) Then
                outErrorText = "callMacro row argument references unknown table '" & tableRef & "'."
                Exit Function
            End If
            If argSpec.Exists("RowIndex") Then
                If CLng(argSpec("RowIndex")) < 0 Then
                    outErrorText = "callMacro row argument index must be >= 0 for '" & tableRef & "'."
                    Exit Function
                End If
            ElseIf Not argSpec.Exists("RowSelector") Then
                outErrorText = "callMacro row argument is missing row selector/index for '" & tableRef & "'."
                Exit Function
            End If

        Case "cellref"
            tableRef = CStr(argSpec("TableRef"))
            If allowedTableFields Is Nothing Or Not allowedTableFields.Exists(tableRef) Then
                outErrorText = "callMacro cell argument references unknown table '" & tableRef & "'."
                Exit Function
            End If
            If argSpec.Exists("RowIndex") Then
                If CLng(argSpec("RowIndex")) < 0 Then
                    outErrorText = "callMacro cell argument index must be >= 0 for '" & tableRef & "'."
                    Exit Function
                End If
            ElseIf Not argSpec.Exists("RowSelector") Then
                outErrorText = "callMacro cell argument is missing row selector/index for '" & tableRef & "'."
                Exit Function
            End If
            If Not mp_TryResolveMapKeyByFieldAlias(allowedTableFields, tableRef, CStr(argSpec("FieldAlias")), resolvedMapKey, outErrorText) Then Exit Function

        Case "scopecellref"
            If scopeVarTypes Is Nothing Then
                outErrorText = "callMacro row-cell argument variable '" & CStr(argSpec("RowVar")) & "' is not available in this scope."
                Exit Function
            End If
            If Not scopeVarTypes.Exists(CStr(argSpec("RowVar"))) Then
                outErrorText = "callMacro row-cell argument variable '" & CStr(argSpec("RowVar")) & "' is not available in this scope."
                Exit Function
            End If
            If StrComp(LCase$(CStr(scopeVarTypes(CStr(argSpec("RowVar"))))), "row", vbBinaryCompare) <> 0 Then
                outErrorText = "callMacro row-cell argument variable '" & CStr(argSpec("RowVar")) & "' must be row type."
                Exit Function
            End If
    End Select

    m_ValidateMacroArgSpec = True
End Function

Public Function m_ResolveCellReferenceArg(ByVal argSpec As Object, ByVal tablesByRef As Object) As String
    Dim rowObj As Object
    Dim fieldAlias As String

    Set rowObj = m_ResolveRowReferenceArg(argSpec, tablesByRef)
    fieldAlias = CStr(argSpec("FieldAlias"))

    If Not rowObj.HasAlias(fieldAlias) Then
        Err.Raise vbObjectError + 1605, ERR_SOURCE, "Field alias '" & fieldAlias & "' is not available in referenced row."
    End If

    m_ResolveCellReferenceArg = CStr(rowObj.Column(fieldAlias))
End Function

Public Function m_ResolveRowReferenceArg(ByVal argSpec As Object, ByVal tablesByRef As Object) As Object
    Dim tableRef As String
    Dim rowIndex As Long
    Dim rowSelector As String
    Dim tableObj As obj_ResultTable
    Dim rowObj As obj_ResultRow

    tableRef = CStr(argSpec("TableRef"))
    If tablesByRef Is Nothing Or Not tablesByRef.Exists(tableRef) Then
        Err.Raise vbObjectError + 1603, ERR_SOURCE, "Table '" & tableRef & "' is not available for row reference."
    End If

    Set tableObj = tablesByRef(tableRef)

    If argSpec.Exists("RowSelector") Then
        rowSelector = LCase$(Trim$(CStr(argSpec("RowSelector"))))
        If Not ex_obj_ResultTableDsl.m_TryGetRowBySelector(tableObj, rowSelector, rowObj) Then
            If rowSelector = "prevrow" Then
                Err.Raise vbObjectError + 1606, ERR_SOURCE, "Table '" & tableRef & "' has fewer than 2 rows; cannot resolve prevRow."
            Else
                Err.Raise vbObjectError + 1607, ERR_SOURCE, "Table '" & tableRef & "' has no rows; cannot resolve lastRow."
            End If
        End If
        Set m_ResolveRowReferenceArg = rowObj
        Exit Function
    End If

    rowIndex = CLng(argSpec("RowIndex"))
    If rowIndex < 0 Then
        Err.Raise vbObjectError + 1602, ERR_SOURCE, "Row index must be >= 0 in '" & tableRef & ".row[" & CStr(rowIndex) & "]'."
    End If

    Set rowObj = tableObj.Row(rowIndex)
    If rowObj Is Nothing Then
        Err.Raise vbObjectError + 1604, ERR_SOURCE, "Row index " & CStr(rowIndex) & " is out of range for table '" & tableRef & "' (rows=" & CStr(tableObj.Count) & ")."
    End If

    Set m_ResolveRowReferenceArg = rowObj
End Function

Public Function m_BuildMapKey(ByVal sourceAlias As String, ByVal tableAlias As String, ByVal fieldAlias As String) As String
    m_BuildMapKey = Trim$(sourceAlias) & ".Sheet[" & Trim$(tableAlias) & "].Map[" & Trim$(fieldAlias) & "]"
End Function

Private Function mp_TryResolveMapKeyByFieldAlias( _
    ByVal fieldsByTable As Object, _
    ByVal tableRef As String, _
    ByVal fieldAlias As String, _
    ByRef outMapKey As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim mapKeys As Object
    Dim key As Variant
    Dim parsedTableAlias As String
    Dim parsedFieldAlias As String
    Dim hits As Long

    If fieldsByTable Is Nothing Or Not fieldsByTable.Exists(tableRef) Then
        outErrorText = "Unknown table reference '" & tableRef & "'."
        Exit Function
    End If
    Set mapKeys = fieldsByTable(tableRef)

    For Each key In mapKeys.Keys
        If mp_TryParseMapKey(CStr(key), parsedTableAlias, parsedFieldAlias) Then
            If StrComp(parsedFieldAlias, fieldAlias, vbTextCompare) = 0 Then
                outMapKey = CStr(key)
                hits = hits + 1
            End If
        End If
    Next key

    If hits = 0 Then
        outErrorText = "Field alias '" & fieldAlias & "' is not configured for table '" & tableRef & "'."
        Exit Function
    End If
    If hits > 1 Then
        outErrorText = "Field alias '" & fieldAlias & "' is ambiguous for table '" & tableRef & "'. Use full reference."
        Exit Function
    End If

    mp_TryResolveMapKeyByFieldAlias = True
End Function

Private Function mp_TryParseMapKey( _
    ByVal mapKey As String, _
    ByRef outTableAlias As String, _
    ByRef outFieldAlias As String, _
    Optional ByRef outSourceAlias As String _
) As Boolean
    Dim sheetStart As Long
    Dim sheetEnd As Long
    Dim mapStart As Long
    Dim mapEnd As Long

    sheetStart = InStr(1, mapKey, ".Sheet[", vbTextCompare)
    mapStart = InStr(1, mapKey, "].Map[", vbTextCompare)
    If sheetStart <= 0 Or mapStart <= sheetStart Then Exit Function

    outSourceAlias = Trim$(Left$(mapKey, sheetStart - 1))
    If Len(outSourceAlias) = 0 Then Exit Function

    sheetStart = sheetStart + Len(".Sheet[")
    sheetEnd = mapStart
    outTableAlias = Mid$(mapKey, sheetStart, sheetEnd - sheetStart)
    If Len(Trim$(outTableAlias)) = 0 Then Exit Function

    mapStart = mapStart + Len("].Map[")
    mapEnd = InStr(mapStart, mapKey, "]", vbBinaryCompare)
    If mapEnd <= mapStart Then Exit Function

    outFieldAlias = Mid$(mapKey, mapStart, mapEnd - mapStart)
    If Len(Trim$(outFieldAlias)) = 0 Then Exit Function

    mp_TryParseMapKey = True
End Function

Private Function mp_TryParseQuotedString(ByVal valueText As String, ByRef outValue As String) As Boolean
    Dim rawInner As String

    valueText = mp_TrimOuterWhitespace(valueText)
    If Len(valueText) >= 6 Then
        If Left$(valueText, 3) = "```" And Right$(valueText, 3) = "```" Then
            outValue = Mid$(valueText, 4, Len(valueText) - 6)
            mp_TryParseQuotedString = True
            Exit Function
        End If
    End If

    If Len(valueText) < 2 Then Exit Function
    If Left$(valueText, 1) <> """" Then Exit Function
    If Right$(valueText, 1) <> """" Then Exit Function

    rawInner = Mid$(valueText, 2, Len(valueText) - 2)
    outValue = mp_DecodeEscapes(rawInner)
    mp_TryParseQuotedString = True
End Function

Private Function mp_DecodeEscapes(ByVal textValue As String) As String
    Dim result As String

    result = Replace(textValue, "\""", """")
    result = Replace(result, "\n", vbLf)
    result = Replace(result, "\\", "\")

    mp_DecodeEscapes = result
End Function

Private Function mp_TrimOuterWhitespace(ByVal textValue As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim ch As String

    textValue = CStr(textValue)
    startPos = 1
    endPos = Len(textValue)

    Do While startPos <= endPos
        ch = Mid$(textValue, startPos, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Do
        startPos = startPos + 1
    Loop

    Do While endPos >= startPos
        ch = Mid$(textValue, endPos, 1)
        If ch <> " " And ch <> vbTab And ch <> vbCr And ch <> vbLf Then Exit Do
        endPos = endPos - 1
    Loop

    If endPos < startPos Then
        mp_TrimOuterWhitespace = vbNullString
    Else
        mp_TrimOuterWhitespace = Mid$(textValue, startPos, endPos - startPos + 1)
    End If
End Function
