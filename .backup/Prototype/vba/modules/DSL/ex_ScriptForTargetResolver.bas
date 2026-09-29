Attribute VB_Name = "ex_ScriptForTargetResolver"
Option Explicit

Private Const LOOP_TARGET_TABLE_ROWS As String = "tablerows"
Private Const LOOP_TARGET_ROW_COLUMNS As String = "rowcolumns"
Private Const LOOP_TARGET_MEMBER_ROWS As String = "memberrows"

Private Const VAR_TYPE_ROW As String = "row"
Private Const VAR_TYPE_COLUMN As String = "column"
Private Const VAR_TYPE_OBJECT As String = "object"

Private Const BATCH_KEYRESULT_KEY_ALIAS As String = "Key"
Private Const BATCH_KEYRESULT_KEYFIELD_SUFFIX As String = ".KeyFieldAlias"

Public Function m_TryParseForTarget( _
    ByVal targetText As String, _
    ByRef outDescriptor As Object, _
    ByRef outErrorText As String _
) As Boolean
    Dim tableRef As String
    Dim sourceRowVar As String
    Dim memberVarName As String
    Dim memberName As String

    Set outDescriptor = CreateObject("Scripting.Dictionary")
    outDescriptor.CompareMode = 1

    If ex_obj_ResultTableDsl.m_TryParseTableRowsRef(targetText, tableRef) Then
        outDescriptor("LoopTarget") = LOOP_TARGET_TABLE_ROWS
        outDescriptor("TableRef") = tableRef
        m_TryParseForTarget = True
        Exit Function
    End If

    If ex_obj_ResultRowDsl.m_TryParseRowColumnsRef(targetText, sourceRowVar) Then
        outDescriptor("LoopTarget") = LOOP_TARGET_ROW_COLUMNS
        outDescriptor("SourceRowVar") = sourceRowVar
        m_TryParseForTarget = True
        Exit Function
    End If

    If mp_TryParseMemberRowsTarget(targetText, memberVarName, memberName) Then
        outDescriptor("LoopTarget") = LOOP_TARGET_MEMBER_ROWS
        outDescriptor("ScopeVar") = memberVarName
        outDescriptor("MemberName") = memberName
        m_TryParseForTarget = True
        Exit Function
    End If

    outErrorText = "Invalid for target. Use Source.Sheet[Table].rows, <scopeVar>.<member>, or <rowVar>.columns."
End Function

Public Function m_ValidateForTargetDescriptor( _
    ByVal descriptor As Object, _
    ByVal scopeVarTypes As Object, _
    ByVal allowedTableFields As Object, _
    ByRef outLoopVarType As String, _
    ByRef outCurrentTableRef As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim loopTarget As String
    Dim sourceRowVar As String
    Dim tableRef As String

    If descriptor Is Nothing Then
        outErrorText = "For-loop target descriptor is missing."
        Exit Function
    End If
    If Not descriptor.Exists("LoopTarget") Then
        outErrorText = "For-loop target descriptor has no LoopTarget."
        Exit Function
    End If

    loopTarget = LCase$(CStr(descriptor("LoopTarget")))

    Select Case loopTarget
        Case LOOP_TARGET_TABLE_ROWS
            tableRef = CStr(descriptor("TableRef"))
            If allowedTableFields Is Nothing Or Not allowedTableFields.Exists(tableRef) Then
                outErrorText = "Unknown table reference in script: '" & tableRef & "'."
                Exit Function
            End If
            outLoopVarType = VAR_TYPE_ROW
            outCurrentTableRef = tableRef

        Case LOOP_TARGET_ROW_COLUMNS
            sourceRowVar = CStr(descriptor("SourceRowVar"))
            If scopeVarTypes Is Nothing Or Not scopeVarTypes.Exists(sourceRowVar) Then
                outErrorText = "Unknown row variable '" & sourceRowVar & "' in for-loop."
                Exit Function
            End If
            If StrComp(CStr(scopeVarTypes(sourceRowVar)), VAR_TYPE_ROW, vbTextCompare) <> 0 Then
                outErrorText = "Variable '" & sourceRowVar & "' must be a row variable to use .columns iteration."
                Exit Function
            End If
            outLoopVarType = VAR_TYPE_COLUMN
            outCurrentTableRef = vbNullString

        Case LOOP_TARGET_MEMBER_ROWS
            sourceRowVar = CStr(descriptor("ScopeVar"))
            If scopeVarTypes Is Nothing Or Not scopeVarTypes.Exists(sourceRowVar) Then
                outErrorText = "Unknown scoped variable '" & sourceRowVar & "' in for-loop."
                Exit Function
            End If
            If StrComp(CStr(scopeVarTypes(sourceRowVar)), VAR_TYPE_ROW, vbTextCompare) <> 0 And _
               StrComp(CStr(scopeVarTypes(sourceRowVar)), VAR_TYPE_OBJECT, vbTextCompare) <> 0 Then
                outErrorText = "Variable '" & sourceRowVar & "' must be a row/object variable to iterate member rows."
                Exit Function
            End If
            outLoopVarType = VAR_TYPE_ROW
            outCurrentTableRef = vbNullString

        Case Else
            outErrorText = "Unsupported for-loop target '" & loopTarget & "'."
            Exit Function
    End Select

    m_ValidateForTargetDescriptor = True
End Function

Public Function m_ResolveForTargetRuntimeContext( _
    ByVal descriptor As Object, _
    ByVal runtimeVars As Object, _
    ByVal tablesByRef As Object, _
    ByRef outRows As Collection, _
    ByRef outSourceRow As obj_ResultRow, _
    ByRef outCurrentTableRef As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim loopTarget As String
    Dim sourceRowVar As String
    Dim tableRef As String
    Dim scopeValue As obj_ScriptScopeValue
    Dim scopeObject As Object
    Dim resolvedTableRef As String
    Dim memberName As String

    If descriptor Is Nothing Or Not descriptor.Exists("LoopTarget") Then
        outErrorText = "For-loop target descriptor is missing."
        Exit Function
    End If

    loopTarget = LCase$(CStr(descriptor("LoopTarget")))

    Select Case loopTarget
        Case LOOP_TARGET_TABLE_ROWS
            tableRef = CStr(descriptor("TableRef"))
            If Not ex_ResultRuntimeAdapter.m_TryGetRowsForTableRef(tablesByRef, tableRef, outRows) Then
                outErrorText = "Table rows target was not found: '" & tableRef & "'."
                Exit Function
            End If
            outCurrentTableRef = tableRef

        Case LOOP_TARGET_ROW_COLUMNS
            sourceRowVar = CStr(descriptor("SourceRowVar"))
            If Not mp_TryResolveRuntimeRowVar(runtimeVars, sourceRowVar, "Row variable", outSourceRow, outErrorText) Then Exit Function
            Set outRows = outSourceRow.Columns
            outCurrentTableRef = vbNullString

        Case LOOP_TARGET_MEMBER_ROWS
            sourceRowVar = CStr(descriptor("ScopeVar"))
            memberName = CStr(descriptor("MemberName"))
            If Not ex_ScriptScopeValue.m_TryGetScopeValue(runtimeVars, sourceRowVar, scopeValue, outErrorText) Then
                outErrorText = "Scoped variable '" & sourceRowVar & "' is not available for loop iteration: " & outErrorText
                Exit Function
            End If
            If Not ex_ScriptScopeValue.m_TryGetObjectValue(scopeValue, scopeObject, outErrorText) Then
                outErrorText = "Scoped variable '" & sourceRowVar & "' must be object for member rows iteration: " & outErrorText
                Exit Function
            End If

            If Not m_TryResolveMemberRows(scopeObject, memberName, tablesByRef, outRows, resolvedTableRef, outErrorText) Then
                Exit Function
            End If
            If TypeOf scopeObject Is obj_ResultRow Then Set outSourceRow = scopeObject
            outCurrentTableRef = resolvedTableRef

        Case Else
            outErrorText = "Unsupported for-loop target '" & loopTarget & "'."
            Exit Function
    End Select

    m_ResolveForTargetRuntimeContext = True
End Function

Public Function m_TryResolveMemberRows( _
    ByVal scopeObject As Object, _
    ByVal memberName As String, _
    ByVal tablesByRef As Object, _
    ByRef outRows As Collection, _
    ByRef outTableRef As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim resolvedTableRef As String
    Dim memberObj As Object
    Dim memberScopeValue As obj_ScriptScopeValue

    If scopeObject Is Nothing Then
        outErrorText = "Scoped object is missing for member rows resolve."
        Exit Function
    End If

    memberName = Trim$(memberName)
    If Len(memberName) = 0 Then
        outErrorText = "Member name is empty in member rows resolve."
        Exit Function
    End If

    If TypeOf scopeObject Is obj_ResultRow Then
        If Not mp_TryResolveMemberRowsFromRowAlias(scopeObject, memberName, resolvedTableRef, outErrorText) Then
            Exit Function
        End If
        If Not ex_ResultRuntimeAdapter.m_TryGetRowsForTableRef(tablesByRef, resolvedTableRef, outRows) Then
            outErrorText = "Member rows table was not found: '" & resolvedTableRef & "'."
            Exit Function
        End If
        outTableRef = resolvedTableRef
        m_TryResolveMemberRows = True
        Exit Function
    End If

    If TypeOf scopeObject Is obj_ResultTable Then
        If StrComp(memberName, "rows", vbTextCompare) <> 0 Then
            outErrorText = "Unsupported table member '" & memberName & "' for member rows resolve. Use '.rows'."
            Exit Function
        End If

        Set outRows = scopeObject.Rows
        outTableRef = scopeObject.TableRef
        m_TryResolveMemberRows = True
        Exit Function
    End If

    If TypeName(scopeObject) = "Dictionary" Then
        If Not scopeObject.Exists(memberName) Then
            outErrorText = "Member '" & memberName & "' is missing in scoped object."
            Exit Function
        End If

        If Not IsObject(scopeObject(memberName)) Then
            outErrorText = "Member '" & memberName & "' must be typed scope value (obj_ScriptScopeValue)."
            Exit Function
        End If

        Set memberObj = scopeObject(memberName)
        If Not TypeOf memberObj Is obj_ScriptScopeValue Then
            outErrorText = "Unsupported member scope container for '" & memberName & "': " & TypeName(memberObj) & ". Expected obj_ScriptScopeValue."
            Exit Function
        End If
        Set memberScopeValue = memberObj
        m_TryResolveMemberRows = mp_TryResolveRowsFromScopeValue(memberScopeValue, tablesByRef, outRows, outTableRef, outErrorText)
        Exit Function
    End If

    outErrorText = "Unsupported scoped object type for member rows resolve: " & TypeName(scopeObject)
End Function

Private Function mp_TryResolveRowsFromScopeValue( _
    ByVal scopeValue As obj_ScriptScopeValue, _
    ByVal tablesByRef As Object, _
    ByRef outRows As Collection, _
    ByRef outTableRef As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim valueObject As Object
    Dim valueText As String

    If scopeValue Is Nothing Then
        outErrorText = "Member scope value is missing."
        Exit Function
    End If

    If scopeValue.HasObjectValue Then
        Set valueObject = scopeValue.ObjectValue
        If valueObject Is Nothing Then
            outErrorText = "Member scope object payload is empty."
            Exit Function
        End If
        If TypeOf valueObject Is obj_ResultTable Then
            outTableRef = valueObject.TableRef
            Set outRows = valueObject.Rows
            mp_TryResolveRowsFromScopeValue = True
            Exit Function
        End If
        If TypeName(valueObject) = "Collection" Then
            Set outRows = valueObject
            outTableRef = vbNullString
            mp_TryResolveRowsFromScopeValue = True
            Exit Function
        End If

        outErrorText = "Unsupported scope value object kind for member rows resolve: " & TypeName(valueObject)
        Exit Function
    End If

    valueText = Trim$(scopeValue.TextValue)
    If Len(valueText) = 0 Then
        outErrorText = "Member scope value text is empty."
        Exit Function
    End If
    If Not ex_ScriptParserCore.m_IsSheetRef(valueText) Then
        outErrorText = "Member scope value text must be sheet reference."
        Exit Function
    End If
    If Not ex_ResultRuntimeAdapter.m_TryGetRowsForTableRef(tablesByRef, valueText, outRows) Then
        outErrorText = "Member rows table was not found: '" & valueText & "'."
        Exit Function
    End If

    outTableRef = valueText
    mp_TryResolveRowsFromScopeValue = True
End Function

Public Function m_TryResolveMemberTableRef( _
    ByVal scopeRow As obj_ResultRow, _
    ByVal memberName As String, _
    ByRef outTableRef As String _
) As Boolean
    ' Deprecated: row member table refs are now resolved from direct alias <member>.
    ' This helper remains only for backward compatibility with external callers.
    Dim aliasName As String

    If scopeRow Is Nothing Then Exit Function

    memberName = Trim$(memberName)
    If Len(memberName) = 0 Then Exit Function

    aliasName = memberName & "TableRef"
    If Not scopeRow.HasAlias(aliasName) Then Exit Function

    outTableRef = Trim$(CStr(scopeRow.Column(aliasName)))
    If Len(outTableRef) = 0 Then Exit Function
    If Not ex_ScriptParserCore.m_IsSheetRef(outTableRef) Then Exit Function

    m_TryResolveMemberTableRef = True
End Function

Private Function mp_TryResolveMemberRowsFromRowAlias( _
    ByVal scopeRow As obj_ResultRow, _
    ByVal memberName As String, _
    ByRef outTableRef As String, _
    ByRef outErrorText As String _
) As Boolean
    memberName = Trim$(memberName)
    If scopeRow Is Nothing Then
        outErrorText = "Scoped source row is missing for member rows resolve."
        Exit Function
    End If
    If Len(memberName) = 0 Then
        outErrorText = "Member name is empty in member rows resolve."
        Exit Function
    End If

    If Not scopeRow.HasAlias(memberName) Then
        outErrorText = "Scoped row has no alias '" & memberName & "' for member rows resolve."
        Exit Function
    End If

    outTableRef = Trim$(CStr(scopeRow.Column(memberName)))
    If Len(outTableRef) = 0 Then
        outErrorText = "Alias '" & memberName & "' is empty in scoped row member rows resolve."
        Exit Function
    End If
    If Not ex_ScriptParserCore.m_IsSheetRef(outTableRef) Then
        outErrorText = "Alias '" & memberName & "' must contain table reference in form Source.Sheet[Table]."
        Exit Function
    End If

    mp_TryResolveMemberRowsFromRowAlias = True
End Function

Public Function m_TryResolveScopedKeyContext( _
    ByVal descriptor As Object, _
    ByVal scopeRow As obj_ResultRow, _
    ByRef outKeyValue As String, _
    ByRef outKeyFieldAlias As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim sourceAlias As String
    Dim tableAlias As String
    Dim compositeAlias As String

    If descriptor Is Nothing Then
        outErrorText = "Scoped target descriptor is missing."
        Exit Function
    End If
    If scopeRow Is Nothing Then
        outErrorText = "Scoped source row is missing."
        Exit Function
    End If

    If Not scopeRow.HasAlias(BATCH_KEYRESULT_KEY_ALIAS) Then
        outErrorText = "Scope row is missing key alias '" & BATCH_KEYRESULT_KEY_ALIAS & "'."
        Exit Function
    End If
    outKeyValue = CStr(scopeRow.Column(BATCH_KEYRESULT_KEY_ALIAS))

    sourceAlias = Trim$(CStr(descriptor("SourceAlias")))
    tableAlias = Trim$(CStr(descriptor("TableAlias")))
    compositeAlias = sourceAlias & "." & tableAlias & BATCH_KEYRESULT_KEYFIELD_SUFFIX

    If scopeRow.HasAlias(compositeAlias) Then
        outKeyFieldAlias = Trim$(CStr(scopeRow.Column(compositeAlias)))
    ElseIf scopeRow.HasAlias("KeyFieldAlias") Then
        outKeyFieldAlias = Trim$(CStr(scopeRow.Column("KeyFieldAlias")))
    Else
        outKeyFieldAlias = vbNullString
    End If

    If Len(outKeyFieldAlias) = 0 Then
        outErrorText = "Scope row is missing key field alias metadata for target '" & CStr(descriptor("TableRef")) & "'."
        Exit Function
    End If

    m_TryResolveScopedKeyContext = True
End Function

Private Function mp_TryResolveRuntimeRowVar( _
    ByVal runtimeVars As Object, _
    ByVal varName As String, _
    ByVal varLabel As String, _
    ByRef outRow As obj_ResultRow, _
    ByRef outErrorText As String _
) As Boolean
    Dim scopeValue As obj_ScriptScopeValue
    Dim scopeObject As Object

    If Not ex_ScriptScopeValue.m_TryGetScopeValue(runtimeVars, varName, scopeValue, outErrorText) Then
        outErrorText = varLabel & " '" & varName & "' is not available for loop iteration: " & outErrorText
        Exit Function
    End If
    If Not ex_ScriptScopeValue.m_TryGetObjectValue(scopeValue, scopeObject, outErrorText) Then
        outErrorText = varLabel & " '" & varName & "' must be row object for loop iteration: " & outErrorText
        Exit Function
    End If
    If Not TypeOf scopeObject Is obj_ResultRow Then
        outErrorText = varLabel & " '" & varName & "' must be row object for loop iteration."
        Exit Function
    End If

    Set outRow = scopeObject
    mp_TryResolveRuntimeRowVar = True
End Function

Private Function mp_TryParseMemberRowsTarget( _
    ByVal targetText As String, _
    ByRef outScopeVarName As String, _
    ByRef outMemberName As String _
) As Boolean
    If Not mp_TryParseVariableMemberRef(targetText, outScopeVarName, outMemberName) Then Exit Function
    If StrComp(outMemberName, "columns", vbTextCompare) = 0 Then Exit Function
    mp_TryParseMemberRowsTarget = True
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
