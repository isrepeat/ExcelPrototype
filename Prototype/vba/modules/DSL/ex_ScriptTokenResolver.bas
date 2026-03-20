Attribute VB_Name = "ex_ScriptTokenResolver"
Option Explicit

Private Const VAR_TYPE_ROW As String = "row"
Private Const VAR_TYPE_COLUMN As String = "column"

Public Function m_TryResolveTokenForValidation( _
    ByVal tokenText As String, _
    ByVal currentTableRef As String, _
    ByVal currentRowVar As String, _
    ByVal scopeVarTypes As Object, _
    ByVal allowedTableFields As Object, _
    ByRef outResolvedTableRef As String, _
    ByRef outResolvedMapKey As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim variableName As String
    Dim memberName As String
    Dim variableType As String
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

    If ex_ScriptParserCore.m_IsIdentifier(tokenText) Then
        If scopeVarTypes Is Nothing Or Not scopeVarTypes.Exists(tokenText) Then
            outErrorText = "Unknown variable '" & tokenText & "' in condition."
            Exit Function
        End If
        outResolvedTableRef = vbNullString
        outResolvedMapKey = vbNullString
        m_TryResolveTokenForValidation = True
        Exit Function
    End If

    If mp_TryParseVariableMemberRef(tokenText, variableName, memberName) Then
        If scopeVarTypes Is Nothing Or Not scopeVarTypes.Exists(variableName) Then
            outErrorText = "Unknown variable '" & variableName & "' in condition."
            Exit Function
        End If

        variableType = LCase$(CStr(scopeVarTypes(variableName)))
        Select Case variableType
            Case VAR_TYPE_COLUMN
                If Not ex_ScriptDslContracts.m_IsMemberAllowed(ex_ScriptDslContracts.TYPE_COLUMN, memberName) Then
                    outErrorText = "Unsupported column member '" & memberName & "' in token '" & tokenText & "'."
                    Exit Function
                End If
            Case Else
                outErrorText = "Variable '" & variableName & "' does not support member access in token '" & tokenText & "'."
                Exit Function
        End Select

        outResolvedTableRef = vbNullString
        outResolvedMapKey = vbNullString
        m_TryResolveTokenForValidation = True
        Exit Function
    End If

    If ex_obj_ResultRowDsl.m_TryParseRowColumnRef(tokenText, rowVarName, fieldAlias) Then
        If scopeVarTypes Is Nothing Or Not scopeVarTypes.Exists(rowVarName) Then
            outErrorText = "Unknown row variable '" & rowVarName & "' in condition."
            Exit Function
        End If
        If StrComp(CStr(scopeVarTypes(rowVarName)), VAR_TYPE_ROW, vbTextCompare) <> 0 Then
            outErrorText = "Variable '" & rowVarName & "' must be a row variable in token '" & tokenText & "'."
            Exit Function
        End If
        outResolvedTableRef = vbNullString
        outResolvedMapKey = vbNullString
        m_TryResolveTokenForValidation = True
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
        m_TryResolveTokenForValidation = True
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
        m_TryResolveTokenForValidation = True
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
        m_TryResolveTokenForValidation = True
        Exit Function
    End If

    If ex_obj_ResultTableDsl.m_TryParseFullRowColumnRef(tokenText, sourceAlias, tableAlias, rowIndex, fieldAlias) Then
        tableRef = sourceAlias & ".Sheet[" & tableAlias & "]"
        mapKey = mp_BuildMapKey(sourceAlias, tableAlias, fieldAlias)

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
        m_TryResolveTokenForValidation = True
        Exit Function
    End If

    outErrorText = "Unsupported field reference '" & tokenText & "'. Use <rowVar>.column[FieldAlias], Source.Sheet[TableAlias].row[N].column[FieldAlias], Source.Sheet[TableAlias].lastRow.column[FieldAlias], Source.Sheet[TableAlias].prevRow.column[FieldAlias], Source.Sheet[TableAlias].count, or Source.Sheet[TableAlias].rowCount."
End Function

Public Function m_TryResolveTokenRuntime( _
    ByVal tokenText As String, _
    ByVal currentTableRef As String, _
    ByVal currentRowVar As String, _
    ByVal currentRowRef As obj_ResultRow, _
    ByVal tablesByRef As Object, _
    ByVal runtimeVars As Object, _
    ByRef outValue As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim variableName As String
    Dim memberName As String
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

    If ex_ScriptParserCore.m_IsIdentifier(tokenText) Then
        If runtimeVars Is Nothing Or Not runtimeVars.Exists(tokenText) Then
            outErrorText = "Unknown variable '" & tokenText & "'."
            Exit Function
        End If
        If Not mp_TryConvertScopeEntryToString(runtimeVars, tokenText, outValue, outErrorText) Then Exit Function
        m_TryResolveTokenRuntime = True
        Exit Function
    End If

    If mp_TryParseVariableMemberRef(tokenText, variableName, memberName) Then
        If Not mp_TryResolveScopeMemberValue(runtimeVars, variableName, memberName, tokenText, outValue, outErrorText) Then Exit Function
        m_TryResolveTokenRuntime = True
        Exit Function
    End If

    If ex_obj_ResultRowDsl.m_TryParseRowColumnRef(tokenText, rowVarName, fieldAlias) Then
        If Not mp_TryResolveScopedRowCellValue(runtimeVars, rowVarName, fieldAlias, tokenText, outValue, outErrorText) Then Exit Function
        m_TryResolveTokenRuntime = True
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
        m_TryResolveTokenRuntime = True
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

        m_TryResolveTokenRuntime = True
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
        m_TryResolveTokenRuntime = True
        Exit Function
    End If

    outErrorText = "Unsupported field reference '" & tokenText & "'. Use <rowVar>.column[FieldAlias], Source.Sheet[TableAlias].row[N].column[FieldAlias], Source.Sheet[TableAlias].lastRow.column[FieldAlias], Source.Sheet[TableAlias].prevRow.column[FieldAlias], Source.Sheet[TableAlias].count, or Source.Sheet[TableAlias].rowCount."
End Function

Private Function mp_BuildMapKey(ByVal sourceAlias As String, ByVal tableAlias As String, ByVal fieldAlias As String) As String
    mp_BuildMapKey = Trim$(sourceAlias) & ".Sheet[" & Trim$(tableAlias) & "].Map[" & Trim$(fieldAlias) & "]"
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
        outErrorText = "Field alias '" & fieldAlias & "' is ambiguous in table '" & tableRef & "'. Use full map key in script token."
        Exit Function
    End If

    mp_TryResolveMapKeyByFieldAlias = True
End Function

Private Function mp_TryParseMapKey( _
    ByVal mapKey As String, _
    ByRef outTableAlias As String, _
    ByRef outFieldAlias As String, _
    Optional ByRef outSourceAlias As String = "" _
) As Boolean
    Dim srcPos As Long
    Dim tblPos As Long
    Dim mapPos As Long
    Dim tableStart As Long
    Dim tableEnd As Long
    Dim fieldStart As Long
    Dim fieldEnd As Long

    mapKey = Trim$(mapKey)
    If Len(mapKey) = 0 Then Exit Function

    srcPos = InStr(1, mapKey, ".Sheet[", vbTextCompare)
    mapPos = InStr(1, mapKey, "].Map[", vbTextCompare)
    If srcPos <= 1 Or mapPos <= srcPos Then Exit Function

    outSourceAlias = Trim$(Left$(mapKey, srcPos - 1))
    If Len(outSourceAlias) = 0 Then Exit Function

    tableStart = srcPos + Len(".Sheet[")
    tableEnd = InStr(tableStart, mapKey, "]", vbBinaryCompare)
    If tableEnd <= tableStart Then Exit Function
    outTableAlias = Trim$(Mid$(mapKey, tableStart, tableEnd - tableStart))
    If Len(outTableAlias) = 0 Then Exit Function

    tblPos = InStr(tableEnd, mapKey, ".Map[", vbTextCompare)
    If tblPos <> tableEnd + 1 Then Exit Function
    fieldStart = tblPos + Len(".Map[")
    fieldEnd = InStr(fieldStart, mapKey, "]", vbBinaryCompare)
    If fieldEnd <= fieldStart Then Exit Function
    If fieldEnd <> Len(mapKey) Then Exit Function

    outFieldAlias = Trim$(Mid$(mapKey, fieldStart, fieldEnd - fieldStart))
    If Len(outFieldAlias) = 0 Then Exit Function

    mp_TryParseMapKey = True
End Function

Private Function mp_TryParseVariableMemberRef( _
    ByVal tokenText As String, _
    ByRef outVariableName As String, _
    ByRef outMemberName As String _
) As Boolean
    Dim dotPos As Long

    tokenText = Trim$(tokenText)
    If InStr(1, tokenText, "[", vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, tokenText, "]", vbBinaryCompare) > 0 Then Exit Function

    dotPos = InStr(1, tokenText, ".", vbBinaryCompare)
    If dotPos <= 1 Then Exit Function
    If InStr(dotPos + 1, tokenText, ".", vbBinaryCompare) > 0 Then Exit Function

    outVariableName = Trim$(Left$(tokenText, dotPos - 1))
    outMemberName = Trim$(Mid$(tokenText, dotPos + 1))
    If Len(outVariableName) = 0 Or Len(outMemberName) = 0 Then Exit Function
    If Not ex_ScriptParserCore.m_IsIdentifier(outVariableName) Then Exit Function
    If Not ex_ScriptParserCore.m_IsIdentifier(outMemberName) Then Exit Function

    mp_TryParseVariableMemberRef = True
End Function

Private Function mp_TryResolveVariableMemberValue( _
    ByVal variableObject As Object, _
    ByVal memberName As String, _
    ByVal tokenText As String, _
    ByRef outValue As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim columnObj As obj_ResultColumn

    If TypeOf variableObject Is obj_ResultColumn Then
        Set columnObj = variableObject
        Select Case LCase$(memberName)
            Case "alias", "name"
                outValue = columnObj.Alias
            Case "value"
                outValue = columnObj.Value
            Case "mapkey"
                outValue = columnObj.MapKey
            Case Else
                outErrorText = "Unsupported column member '" & memberName & "' in token '" & tokenText & "'."
                Exit Function
        End Select
        mp_TryResolveVariableMemberValue = True
        Exit Function
    End If

    outErrorText = "Variable in token '" & tokenText & "' does not support member access."
End Function

Private Function mp_TryConvertScopeValueToString( _
    ByVal variableObject As Object, _
    ByRef outValue As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim columnObj As obj_ResultColumn

    If TypeOf variableObject Is obj_ResultColumn Then
        Set columnObj = variableObject
        outValue = columnObj.Value
        mp_TryConvertScopeValueToString = True
        Exit Function
    End If

    outErrorText = "Variable value is object and cannot be rendered as string."
End Function

Private Function mp_TryConvertScopeEntryToString( _
    ByVal scopeRef As Object, _
    ByVal variableName As String, _
    ByRef outValue As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim scopeValue As obj_ScriptScopeValue
    Dim valueObj As Object

    If Not ex_ScriptScopeValue.m_TryGetScopeValue(scopeRef, variableName, scopeValue, outErrorText) Then Exit Function

    If scopeValue.HasObjectValue Then
        If Not ex_ScriptScopeValue.m_TryGetObjectValue(scopeValue, valueObj, outErrorText) Then Exit Function
        mp_TryConvertScopeEntryToString = mp_TryConvertScopeValueToString(valueObj, outValue, outErrorText)
        Exit Function
    End If

    mp_TryConvertScopeEntryToString = ex_ScriptScopeValue.m_TryGetStringValue(scopeValue, outValue, outErrorText)
End Function

Private Function mp_TryResolveScopeMemberValue( _
    ByVal scopeRef As Object, _
    ByVal variableName As String, _
    ByVal memberName As String, _
    ByVal tokenText As String, _
    ByRef outValue As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim scopeValue As obj_ScriptScopeValue
    Dim variableObject As Object

    If Not ex_ScriptScopeValue.m_TryGetScopeValue(scopeRef, variableName, scopeValue, outErrorText) Then
        outErrorText = "Unknown variable '" & variableName & "' in token '" & tokenText & "': " & outErrorText
        Exit Function
    End If
    If Not ex_ScriptScopeValue.m_TryGetObjectValue(scopeValue, variableObject, outErrorText) Then
        outErrorText = "Variable '" & variableName & "' does not support member access in token '" & tokenText & "'."
        Exit Function
    End If
    mp_TryResolveScopeMemberValue = mp_TryResolveVariableMemberValue(variableObject, memberName, tokenText, outValue, outErrorText)
End Function

Private Function mp_TryResolveScopedRowCellValue( _
    ByVal scopeRef As Object, _
    ByVal rowVarName As String, _
    ByVal fieldAlias As String, _
    ByVal tokenText As String, _
    ByRef outValue As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim scopeValue As obj_ScriptScopeValue
    Dim rowObject As obj_ResultRow
    Dim rowAnyObject As Object

    rowVarName = Trim$(rowVarName)
    fieldAlias = Trim$(fieldAlias)
    If Len(rowVarName) = 0 Or Len(fieldAlias) = 0 Then Exit Function

    If Not ex_ScriptScopeValue.m_TryGetScopeValue(scopeRef, rowVarName, scopeValue, outErrorText) Then
        outErrorText = "Unknown row variable '" & rowVarName & "' in token '" & tokenText & "': " & outErrorText
        Exit Function
    End If
    If Not ex_ScriptScopeValue.m_TryGetObjectValue(scopeValue, rowAnyObject, outErrorText) Then
        outErrorText = "Variable '" & rowVarName & "' is not a row object in token '" & tokenText & "'."
        Exit Function
    End If
    If Not TypeOf rowAnyObject Is obj_ResultRow Then
        outErrorText = "Variable '" & rowVarName & "' must be row object in token '" & tokenText & "'."
        Exit Function
    End If

    Set rowObject = rowAnyObject
    If Not rowObject.HasAlias(fieldAlias) Then
        outErrorText = "Unknown field alias '" & fieldAlias & "' for row variable '" & rowVarName & "'."
        Exit Function
    End If

    outValue = CStr(rowObject.Column(fieldAlias))
    mp_TryResolveScopedRowCellValue = True
End Function
