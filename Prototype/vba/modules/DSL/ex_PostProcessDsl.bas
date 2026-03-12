Attribute VB_Name = "ex_PostProcessDsl"
Option Explicit

Private Const SCRIPT_KEY As String = "PostProcess.Script"
Private Const ACTION_CALL_MACRO As String = "callmacro"
Private Const ACTION_LET As String = "let"
Private Const ACTION_ASSIGN As String = "assign"
Private Const ACTION_BREAK As String = "break"
Private Const ACTION_CONTINUE As String = "continue"
Private Const ACTION_RETURN As String = "return"
Private Const ASSIGN_KIND_CALL_MACRO As String = "callmacro"
Private Const ASSIGN_KIND_STRING_EXPR As String = "stringexpr"
Private Const EXPR_PART_LITERAL As String = "literal"
Private Const EXPR_PART_TOKEN As String = "token"

Private Const LOOP_TARGET_TABLE_ROWS As String = "tablerows"
Private Const LOOP_TARGET_ROW_COLUMNS As String = "rowcolumns"

Private Const VAR_TYPE_ROW As String = "row"
Private Const VAR_TYPE_COLUMN As String = "column"
Private Const VAR_TYPE_STRING As String = "string"

Private Const EXEC_FLOW_NONE As String = ""
Private Const EXEC_FLOW_BREAK As String = "break"
Private Const EXEC_FLOW_CONTINUE As String = "continue"
Private Const EXEC_FLOW_RETURN As String = "return"

Private g_CachedScriptConfigKey As String
Private g_CachedScriptText As String
Private g_CachedScriptBlocks As Collection
Private g_CachedValidationSignature As String

Public Sub m_ResetScriptCache()
    g_CachedScriptConfigKey = vbNullString
    g_CachedScriptText = vbNullString
    Set g_CachedScriptBlocks = Nothing
    g_CachedValidationSignature = vbNullString
End Sub

Public Function m_ValidateScriptAgainstConfig( _
    ByVal cfg As Object, _
    ByVal allowedTableFields As Object, _
    ByRef outErrorText As String, _
    Optional ByVal scriptConfigKey As String = SCRIPT_KEY _
) As Boolean
    Dim scriptText As String
    Dim blocks As Collection
    Dim validationSignature As String
    Dim stepName As String

    On Error GoTo EH

    stepName = "load-compile-script"
    If Not mp_TryGetCompiledScriptBlocks(cfg, scriptConfigKey, scriptText, blocks, outErrorText) Then Exit Function
    If Len(scriptText) = 0 Then
        m_ValidateScriptAgainstConfig = True
        Exit Function
    End If

    validationSignature = mp_BuildValidationSignature(allowedTableFields)
    If mp_IsValidationCacheHit(scriptConfigKey, scriptText, validationSignature) Then
        m_ValidateScriptAgainstConfig = True
        Exit Function
    End If

    stepName = "validate-blocks"
    If Not mp_ValidateBlocks(blocks, allowedTableFields, outErrorText) Then Exit Function
    g_CachedValidationSignature = validationSignature

    m_ValidateScriptAgainstConfig = True
    Exit Function

EH:
    outErrorText = "PostProcess validation runtime error"
    If Len(stepName) > 0 Then
        outErrorText = outErrorText & " at step '" & stepName & "'"
    End If
    outErrorText = outErrorText & ": [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description
    m_ValidateScriptAgainstConfig = False
End Function

Public Sub m_ApplyScriptToSheet( _
    ByVal ws As Worksheet, _
    ByVal cfg As Object, _
    ByVal resultTables As Collection, _
    Optional ByVal scriptConfigKey As String = SCRIPT_KEY _
)
    Dim scriptText As String
    Dim blocks As Collection
    Dim parseOrValidationError As String
    Dim runtimeValidationSignature As String
    Dim ctxTablesByRef As Object
    Dim ctxFields As Object
    Dim postProcessFooterLines As Collection
    Dim usedCols As Long

    If ws Is Nothing Then Exit Sub
    If cfg Is Nothing Then Exit Sub
    If resultTables Is Nothing Then Exit Sub

    If Not mp_TryGetCompiledScriptBlocks(cfg, scriptConfigKey, scriptText, blocks, parseOrValidationError) Then
        Err.Raise vbObjectError + 1592, "ex_PostProcessDsl", parseOrValidationError
    End If
    If Len(scriptText) = 0 Then Exit Sub

    ex_ResultRuntimeAdapter.m_BuildRuntimeContext resultTables, ctxTablesByRef, ctxFields

    runtimeValidationSignature = mp_BuildValidationSignature(ctxFields)
    If Not mp_IsValidationCacheHit(scriptConfigKey, scriptText, runtimeValidationSignature) Then
        If Not mp_ValidateBlocks(blocks, ctxFields, parseOrValidationError) Then
            Err.Raise vbObjectError + 1591, "ex_PostProcessDsl", "PostProcess script validation failed: " & parseOrValidationError
        End If
        g_CachedValidationSignature = runtimeValidationSignature
    End If

    Set postProcessFooterLines = New Collection
    usedCols = mp_GetLastUsedColumn(ws)
    If usedCols <= 0 Then usedCols = 1

    ex_PostProcessActions.m_ResetPostProcessHeaderCursor ws
    ex_PostProcessActions.m_ResetPostProcessFooterCursor ws
    mp_ExecuteBlocks ws, blocks, ctxTablesByRef, postProcessFooterLines, usedCols
    ex_PostProcessActions.m_ScrollToPostProcessResults ws
End Sub

' Backward-compatible wrappers for existing callers.
Public Function m_ValidateTimelineScript( _
    ByVal cfg As Object, _
    ByVal allowedTableFields As Object, _
    ByRef outErrorText As String _
) As Boolean
    m_ValidateTimelineScript = m_ValidateScriptAgainstConfig(cfg, allowedTableFields, outErrorText, SCRIPT_KEY)
End Function

' Backward-compatible wrappers for existing callers.
Public Sub m_ApplyTimelineScript( _
    ByVal ws As Worksheet, _
    ByVal cfg As Object, _
    ByVal resultTables As Collection _
)
    m_ApplyScriptToSheet ws, cfg, resultTables, SCRIPT_KEY
End Sub

Private Function mp_TryGetCompiledScriptBlocks( _
    ByVal cfg As Object, _
    ByVal scriptConfigKey As String, _
    ByRef outScriptText As String, _
    ByRef outBlocks As Collection, _
    ByRef outErrorText As String _
) As Boolean
    Dim parseError As String
    Dim parsedBlocks As Collection
    Dim normalizedScriptKey As String

    outScriptText = vbNullString
    outErrorText = vbNullString
    Set outBlocks = Nothing

    normalizedScriptKey = Trim$(scriptConfigKey)
    If Len(normalizedScriptKey) = 0 Then normalizedScriptKey = SCRIPT_KEY

    If Not ex_PostProcessScriptSource.m_TryGetScriptText(cfg, normalizedScriptKey, outScriptText, outErrorText) Then Exit Function
    If Len(outScriptText) = 0 Then
        mp_TryGetCompiledScriptBlocks = True
        Exit Function
    End If

    If StrComp(g_CachedScriptConfigKey, normalizedScriptKey, vbTextCompare) = 0 Then
        If StrComp(g_CachedScriptText, outScriptText, vbBinaryCompare) = 0 Then
            If Not g_CachedScriptBlocks Is Nothing Then
                Set outBlocks = g_CachedScriptBlocks
                mp_TryGetCompiledScriptBlocks = True
                Exit Function
            End If
        End If
    End If

    If Not mp_ParseScript(outScriptText, parsedBlocks, parseError) Then
        outErrorText = "PostProcess script parse failed: " & parseError
        Exit Function
    End If

    g_CachedScriptConfigKey = normalizedScriptKey
    g_CachedScriptText = outScriptText
    Set g_CachedScriptBlocks = parsedBlocks
    g_CachedValidationSignature = vbNullString

    Set outBlocks = g_CachedScriptBlocks
    mp_TryGetCompiledScriptBlocks = True
End Function

Private Function mp_IsValidationCacheHit( _
    ByVal scriptConfigKey As String, _
    ByVal scriptText As String, _
    ByVal validationSignature As String _
) As Boolean
    Dim normalizedScriptKey As String

    normalizedScriptKey = Trim$(scriptConfigKey)
    If Len(normalizedScriptKey) = 0 Then normalizedScriptKey = SCRIPT_KEY

    If Len(validationSignature) = 0 Then Exit Function
    If StrComp(g_CachedScriptConfigKey, normalizedScriptKey, vbTextCompare) <> 0 Then Exit Function
    If StrComp(g_CachedScriptText, scriptText, vbBinaryCompare) <> 0 Then Exit Function
    If StrComp(g_CachedValidationSignature, validationSignature, vbBinaryCompare) <> 0 Then Exit Function

    mp_IsValidationCacheHit = True
End Function

Private Function mp_BuildValidationSignature(ByVal allowedTableFields As Object) As String
    Dim tableKeys As Variant
    Dim tableKey As Variant
    Dim fieldMap As Object
    Dim fieldKeys As Variant
    Dim i As Long

    If allowedTableFields Is Nothing Then
        mp_BuildValidationSignature = "none"
        Exit Function
    End If

    tableKeys = allowedTableFields.Keys
    If mp_IsEmptyArrayLocal(tableKeys) Then
        mp_BuildValidationSignature = "empty"
        Exit Function
    End If
    mp_SortVariantTextArrayLocal tableKeys

    For i = LBound(tableKeys) To UBound(tableKeys)
        tableKey = tableKeys(i)
        mp_BuildValidationSignature = mp_BuildValidationSignature & "|T:" & CStr(tableKey)

        If IsObject(allowedTableFields(CStr(tableKey))) Then
            Set fieldMap = allowedTableFields(CStr(tableKey))
            fieldKeys = fieldMap.Keys
            If Not mp_IsEmptyArrayLocal(fieldKeys) Then
                mp_SortVariantTextArrayLocal fieldKeys
                mp_BuildValidationSignature = mp_BuildValidationSignature & "|F:" & Join(fieldKeys, ";")
            End If
        End If
    Next i
End Function

Private Sub mp_SortVariantTextArrayLocal(ByRef arr As Variant)
    Dim i As Long
    Dim j As Long
    Dim tmp As Variant

    If mp_IsEmptyArrayLocal(arr) Then Exit Sub

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(CStr(arr(i)), CStr(arr(j)), vbTextCompare) > 0 Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Function mp_IsEmptyArrayLocal(ByVal arr As Variant) As Boolean
    On Error GoTo EH
    If IsArray(arr) = False Then
        mp_IsEmptyArrayLocal = True
        Exit Function
    End If
    If UBound(arr) < LBound(arr) Then
        mp_IsEmptyArrayLocal = True
        Exit Function
    End If
    mp_IsEmptyArrayLocal = False
    Exit Function
EH:
    mp_IsEmptyArrayLocal = True
End Function

Private Function mp_ParseScript(ByVal scriptText As String, ByRef outBlocks As Collection, ByRef outErrorText As String) As Boolean
    Set outBlocks = New Collection
    Dim sourceText As String
    Dim pos As Long
    Dim lineNo As Long

    sourceText = mp_NormalizeScript(scriptText)
    pos = 1
    lineNo = 1
    If Not mp_ParseStatements(sourceText, pos, lineNo, outBlocks, False, outErrorText) Then Exit Function
    mp_ParseScript = True
End Function

Private Function mp_ValidateBlocks(ByVal blocks As Collection, ByVal allowedTableFields As Object, ByRef outErrorText As String) As Boolean
    Dim rootScopeVarTypes As Object

    If blocks Is Nothing Then
        mp_ValidateBlocks = True
        Exit Function
    End If

    Set rootScopeVarTypes = mp_CreateVarScope()
    mp_ValidateBlocks = mp_ValidateStatements(blocks, allowedTableFields, vbNullString, vbNullString, rootScopeVarTypes, 0, outErrorText)
End Function

Private Sub mp_ExecuteBlocks( _
    ByVal ws As Worksheet, _
    ByVal blocks As Collection, _
    ByVal tablesByRef As Object, _
    ByVal postProcessFooterLines As Collection, _
    ByVal usedCols As Long _
)
    Dim rootRuntimeVars As Object
    Dim execFlow As String

    Set rootRuntimeVars = mp_CreateVarScope()
    execFlow = mp_ExecuteStatements(ws, blocks, tablesByRef, postProcessFooterLines, usedCols, vbNullString, vbNullString, Nothing, rootRuntimeVars)
    Select Case LCase$(execFlow)
        Case EXEC_FLOW_NONE, EXEC_FLOW_RETURN
            ' no-op
        Case EXEC_FLOW_BREAK, EXEC_FLOW_CONTINUE
            Err.Raise vbObjectError + 1618, "ex_PostProcessDsl", "'" & execFlow & "' is only allowed inside for-loop."
        Case Else
            Err.Raise vbObjectError + 1619, "ex_PostProcessDsl", "Unsupported control-flow signal: " & execFlow
    End Select
End Sub

Private Function mp_ExecuteStatements( _
    ByVal ws As Worksheet, _
    ByVal statements As Collection, _
    ByVal tablesByRef As Object, _
    ByVal postProcessFooterLines As Collection, _
    ByVal usedCols As Long, _
    ByVal currentTableRef As String, _
    ByVal currentRowVar As String, _
    ByVal currentRowRef As obj_ResultRow, _
    ByVal runtimeVars As Object _
)
    Dim i As Long
    Dim statement As Object
    Dim statementType As String
    Dim macroArgs As Collection
    Dim macroResult As Variant
    Dim rowsList As Collection
    Dim rowRef As obj_ResultRow
    Dim rowIdx As Long
    Dim childRuntimeVars As Object
    Dim loopVarName As String
    Dim sourceRowVarName As String
    Dim sourceRowRef As obj_ResultRow
    Dim rowColumns As Collection
    Dim columnObj As obj_ResultColumn
    Dim macroResultObject As Object
    Dim localLetDeclarations As Object
    Dim letVarName As String
    Dim bodyFlow As String
    Dim assignKind As String

    If statements Is Nothing Then
        mp_ExecuteStatements = EXEC_FLOW_NONE
        Exit Function
    End If
    If runtimeVars Is Nothing Then Set runtimeVars = mp_CreateVarScope()
    Set localLetDeclarations = mp_CreateVarScope()

    For i = 1 To statements.Count
        Set statement = statements(i)
        statementType = LCase$(CStr(statement("Type")))

        Select Case statementType
            Case ACTION_CALL_MACRO
                On Error GoTo CallMacroErr
                Set macroArgs = mp_BuildMacroRuntimeArgs(statement, currentTableRef, currentRowVar, currentRowRef, tablesByRef, runtimeVars)
                ex_PostProcessActionInvoker.m_RunMacroWithArgs CStr(statement("MacroName")), macroArgs
                On Error GoTo 0

            Case ACTION_LET
                letVarName = CStr(statement("VarName"))
                If localLetDeclarations.Exists(letVarName) Then
                    Err.Raise vbObjectError + 1617, "ex_PostProcessDsl", "Variable '" & letVarName & "' is already declared in this scope."
                End If
                assignKind = mp_GetStatementAssignKind(statement)
                Select Case assignKind
                    Case ASSIGN_KIND_STRING_EXPR
                        mp_SetScopeValue runtimeVars, letVarName, mp_EvaluateStringExpression( _
                            statement("ExprParts"), _
                            currentTableRef, _
                            currentRowVar, _
                            currentRowRef, _
                            tablesByRef, _
                            runtimeVars _
                        )
                    Case ASSIGN_KIND_CALL_MACRO
                        On Error GoTo CallMacroErr
                        Set macroArgs = mp_BuildMacroRuntimeArgs(statement, currentTableRef, currentRowVar, currentRowRef, tablesByRef, runtimeVars)
                        If mp_LetExpectsObjectResult(statement) Then
                            Set macroResultObject = ex_PostProcessActionInvoker.m_RunMacroWithArgsReturn(CStr(statement("MacroName")), macroArgs)
                            mp_SetScopeObject runtimeVars, letVarName, macroResultObject
                        Else
                            macroResult = ex_PostProcessActionInvoker.m_RunMacroWithArgsReturn(CStr(statement("MacroName")), macroArgs)
                            mp_SetScopeValue runtimeVars, letVarName, mp_ConvertVariantToString(macroResult)
                        End If
                        On Error GoTo 0
                    Case Else
                        Err.Raise vbObjectError + 1624, "ex_PostProcessDsl", "Unsupported assignment kind: " & assignKind
                End Select
                mp_SetScopeValue localLetDeclarations, letVarName, "1"

            Case ACTION_ASSIGN
                If runtimeVars Is Nothing Or Not runtimeVars.Exists(CStr(statement("VarName"))) Then
                    Err.Raise vbObjectError + 1614, "ex_PostProcessDsl", "Assignment to undeclared variable '" & CStr(statement("VarName")) & "'."
                End If
                assignKind = mp_GetStatementAssignKind(statement)
                Select Case assignKind
                    Case ASSIGN_KIND_STRING_EXPR
                        If IsObject(runtimeVars(CStr(statement("VarName")))) Then
                            Err.Raise vbObjectError + 1615, "ex_PostProcessDsl", "Assignment type mismatch for variable '" & CStr(statement("VarName")) & "': expected row object result."
                        End If
                        mp_SetScopeValue runtimeVars, CStr(statement("VarName")), mp_EvaluateStringExpression( _
                            statement("ExprParts"), _
                            currentTableRef, _
                            currentRowVar, _
                            currentRowRef, _
                            tablesByRef, _
                            runtimeVars _
                        )
                    Case ASSIGN_KIND_CALL_MACRO
                        On Error GoTo CallMacroErr
                        Set macroArgs = mp_BuildMacroRuntimeArgs(statement, currentTableRef, currentRowVar, currentRowRef, tablesByRef, runtimeVars)
                        If IsObject(runtimeVars(CStr(statement("VarName")))) Then
                            If Not mp_LetExpectsObjectResult(statement) Then
                                Err.Raise vbObjectError + 1615, "ex_PostProcessDsl", "Assignment type mismatch for variable '" & CStr(statement("VarName")) & "': expected row object result."
                            End If
                            Set macroResultObject = ex_PostProcessActionInvoker.m_RunMacroWithArgsReturn(CStr(statement("MacroName")), macroArgs)
                            mp_SetScopeObject runtimeVars, CStr(statement("VarName")), macroResultObject
                        Else
                            If mp_LetExpectsObjectResult(statement) Then
                                Err.Raise vbObjectError + 1616, "ex_PostProcessDsl", "Assignment type mismatch for variable '" & CStr(statement("VarName")) & "': expected string-compatible result."
                            End If
                            macroResult = ex_PostProcessActionInvoker.m_RunMacroWithArgsReturn(CStr(statement("MacroName")), macroArgs)
                            mp_SetScopeValue runtimeVars, CStr(statement("VarName")), mp_ConvertVariantToString(macroResult)
                        End If
                        On Error GoTo 0
                    Case Else
                        Err.Raise vbObjectError + 1624, "ex_PostProcessDsl", "Unsupported assignment kind: " & assignKind
                End Select

            Case "if"
                bodyFlow = EXEC_FLOW_NONE
                If mp_EvaluateCondition(CStr(statement("Condition")), currentTableRef, currentRowVar, currentRowRef, tablesByRef, runtimeVars) Then
                    bodyFlow = mp_ExecuteStatements(ws, statement("Body"), tablesByRef, postProcessFooterLines, usedCols, currentTableRef, currentRowVar, currentRowRef, runtimeVars)
                ElseIf statement.Exists("ElseBody") Then
                    bodyFlow = mp_ExecuteStatements(ws, statement("ElseBody"), tablesByRef, postProcessFooterLines, usedCols, currentTableRef, currentRowVar, currentRowRef, runtimeVars)
                End If
                If Len(bodyFlow) > 0 Then
                    mp_ExecuteStatements = bodyFlow
                    Exit Function
                End If

            Case "for"
                Select Case LCase$(CStr(statement("LoopTarget")))
                    Case LOOP_TARGET_TABLE_ROWS
                        If ex_ResultRuntimeAdapter.m_TryGetRowsForTableRef(tablesByRef, CStr(statement("TableRef")), rowsList) Then
                            loopVarName = CStr(statement("LoopVar"))
                            For rowIdx = 1 To rowsList.Count
                                Set rowRef = rowsList(rowIdx)
                                Set childRuntimeVars = mp_CloneVarScope(runtimeVars)
                                mp_SetScopeObject childRuntimeVars, loopVarName, rowRef
                                bodyFlow = mp_ExecuteStatements(ws, statement("Body"), tablesByRef, postProcessFooterLines, usedCols, CStr(statement("TableRef")), loopVarName, rowRef, childRuntimeVars)
                                mp_SyncAssignedParentScope runtimeVars, childRuntimeVars
                                Select Case LCase$(bodyFlow)
                                    Case EXEC_FLOW_NONE
                                        ' no-op
                                    Case EXEC_FLOW_CONTINUE
                                        ' next row
                                    Case EXEC_FLOW_BREAK
                                        Exit For
                                    Case EXEC_FLOW_RETURN
                                        mp_ExecuteStatements = EXEC_FLOW_RETURN
                                        Exit Function
                                    Case Else
                                        Err.Raise vbObjectError + 1620, "ex_PostProcessDsl", "Unsupported control-flow signal in for-loop: " & bodyFlow
                                End Select
                            Next rowIdx
                        End If

                    Case LOOP_TARGET_ROW_COLUMNS
                        sourceRowVarName = CStr(statement("SourceRowVar"))
                        If runtimeVars Is Nothing Or Not runtimeVars.Exists(sourceRowVarName) Then
                            Err.Raise vbObjectError + 1608, "ex_PostProcessDsl", "Row variable '" & sourceRowVarName & "' is not available for .columns iteration."
                        End If
                        If Not IsObject(runtimeVars(sourceRowVarName)) Then
                            Err.Raise vbObjectError + 1609, "ex_PostProcessDsl", "Variable '" & sourceRowVarName & "' is not an object row reference for .columns iteration."
                        End If
                        If Not TypeOf runtimeVars(sourceRowVarName) Is obj_ResultRow Then
                            Err.Raise vbObjectError + 1610, "ex_PostProcessDsl", "Variable '" & sourceRowVarName & "' must be row object for .columns iteration."
                        End If

                        Set sourceRowRef = runtimeVars(sourceRowVarName)
                        Set rowColumns = sourceRowRef.Columns
                        loopVarName = CStr(statement("LoopVar"))

                        For rowIdx = 1 To rowColumns.Count
                            Set columnObj = rowColumns(rowIdx)
                            Set childRuntimeVars = mp_CloneVarScope(runtimeVars)
                            mp_SetScopeObject childRuntimeVars, loopVarName, columnObj
                            bodyFlow = mp_ExecuteStatements(ws, statement("Body"), tablesByRef, postProcessFooterLines, usedCols, currentTableRef, sourceRowVarName, sourceRowRef, childRuntimeVars)
                            mp_SyncAssignedParentScope runtimeVars, childRuntimeVars
                            Select Case LCase$(bodyFlow)
                                Case EXEC_FLOW_NONE
                                    ' no-op
                                Case EXEC_FLOW_CONTINUE
                                    ' next column
                                Case EXEC_FLOW_BREAK
                                    Exit For
                                Case EXEC_FLOW_RETURN
                                    mp_ExecuteStatements = EXEC_FLOW_RETURN
                                    Exit Function
                                Case Else
                                    Err.Raise vbObjectError + 1621, "ex_PostProcessDsl", "Unsupported control-flow signal in for-loop: " & bodyFlow
                            End Select
                        Next rowIdx

                    Case Else
                        Err.Raise vbObjectError + 1611, "ex_PostProcessDsl", "Unsupported for-loop target: " & CStr(statement("LoopTarget"))
                End Select

            Case ACTION_BREAK
                mp_ExecuteStatements = EXEC_FLOW_BREAK
                Exit Function

            Case ACTION_CONTINUE
                mp_ExecuteStatements = EXEC_FLOW_CONTINUE
                Exit Function

            Case ACTION_RETURN
                mp_ExecuteStatements = EXEC_FLOW_RETURN
                Exit Function

            Case Else
                Err.Raise vbObjectError + 1593, "ex_PostProcessDsl", "Unsupported statement type: " & statementType
        End Select
    Next i
    mp_ExecuteStatements = EXEC_FLOW_NONE
    Exit Function

CallMacroErr:
    Err.Raise vbObjectError + 1597, "ex_PostProcessDsl", "callMacro failed for '" & CStr(statement("MacroName")) & "': " & Err.Description
End Function

Private Sub mp_SyncAssignedParentScope(ByVal parentScope As Object, ByVal childScope As Object)
    Dim scopeKey As Variant

    If parentScope Is Nothing Then Exit Sub
    If childScope Is Nothing Then Exit Sub

    For Each scopeKey In parentScope.Keys
        If childScope.Exists(CStr(scopeKey)) Then
            If IsObject(childScope(CStr(scopeKey))) Then
                Set parentScope(CStr(scopeKey)) = childScope(CStr(scopeKey))
            Else
                parentScope(CStr(scopeKey)) = childScope(CStr(scopeKey))
            End If
        End If
    Next scopeKey
End Sub

Private Function mp_ParseStatements( _
    ByVal sourceText As String, _
    ByRef pos As Long, _
    ByRef lineNo As Long, _
    ByRef outStatements As Collection, _
    ByVal stopOnCloseBrace As Boolean, _
    ByRef outErrorText As String _
) As Boolean
    Dim ch As String
    Dim statement As Object

    Do While pos <= Len(sourceText)
        mp_SkipWhitespace sourceText, pos, lineNo
        If pos > Len(sourceText) Then Exit Do

        ch = Mid$(sourceText, pos, 1)
        If ch = "}" Then
            If stopOnCloseBrace Then
                pos = pos + 1
                mp_ParseStatements = True
                Exit Function
            End If
            outErrorText = "Unexpected '}' at line " & CStr(lineNo)
            Exit Function
        End If

        If Not mp_ParseStatement(sourceText, pos, lineNo, statement, outErrorText) Then Exit Function
        outStatements.Add statement
    Loop

    If stopOnCloseBrace Then
        outErrorText = "Missing '}' for block near line " & CStr(lineNo)
        Exit Function
    End If

    mp_ParseStatements = True
End Function

Private Function mp_ParseStatement( _
    ByVal sourceText As String, _
    ByRef pos As Long, _
    ByRef lineNo As Long, _
    ByRef outStatement As Object, _
    ByRef outErrorText As String _
) As Boolean
    Dim keywordText As String
    Dim probePos As Long
    Dim probeLine As Long

    probePos = pos
    probeLine = lineNo
    mp_SkipWhitespace sourceText, probePos, probeLine
    If Not mp_ReadIdentifier(sourceText, probePos, probeLine, keywordText) Then
        outErrorText = "Expected statement at line " & CStr(lineNo)
        Exit Function
    End If

    Select Case LCase$(keywordText)
        Case "for"
            mp_ParseStatement = mp_TryParseForStatement(sourceText, pos, lineNo, outStatement, outErrorText)
        Case "if"
            mp_ParseStatement = mp_TryParseIfStatement(sourceText, pos, lineNo, outStatement, outErrorText)
        Case "else"
            outErrorText = "Unexpected 'else' without matching if at line " & CStr(lineNo)
            mp_ParseStatement = False
        Case "callmacro"
            mp_ParseStatement = mp_TryParseCallMacroStatement(sourceText, pos, lineNo, outStatement, outErrorText)
        Case "let"
            mp_ParseStatement = mp_TryParseLetStatement(sourceText, pos, lineNo, outStatement, outErrorText)
        Case ACTION_BREAK
            mp_ParseStatement = mp_TryParseKeywordNoArgStatement(sourceText, pos, lineNo, ACTION_BREAK, outStatement, outErrorText)
        Case ACTION_CONTINUE
            mp_ParseStatement = mp_TryParseKeywordNoArgStatement(sourceText, pos, lineNo, ACTION_CONTINUE, outStatement, outErrorText)
        Case ACTION_RETURN
            mp_ParseStatement = mp_TryParseKeywordNoArgStatement(sourceText, pos, lineNo, ACTION_RETURN, outStatement, outErrorText)
        Case Else
            mp_ParseStatement = mp_TryParseAssignStatement(sourceText, pos, lineNo, outStatement, outErrorText)
    End Select
End Function

Private Function mp_TryParseKeywordNoArgStatement( _
    ByVal sourceText As String, _
    ByRef pos As Long, _
    ByRef lineNo As Long, _
    ByVal expectedKeyword As String, _
    ByRef outStatement As Object, _
    ByRef outErrorText As String _
) As Boolean
    Dim statementText As String
    Dim bodyText As String
    Dim stmtLine As Long

    stmtLine = lineNo
    If Not mp_ReadStatementToSemicolon(sourceText, pos, lineNo, statementText, outErrorText) Then Exit Function

    bodyText = Trim$(statementText)
    If Right$(bodyText, 1) <> ";" Then
        outErrorText = "Expected ';' after '" & expectedKeyword & "' at line " & CStr(stmtLine)
        Exit Function
    End If
    bodyText = Trim$(Left$(bodyText, Len(bodyText) - 1))
    If StrComp(bodyText, expectedKeyword, vbTextCompare) <> 0 Then
        outErrorText = "'" & expectedKeyword & "' statement does not accept arguments at line " & CStr(stmtLine)
        Exit Function
    End If

    Set outStatement = CreateObject("Scripting.Dictionary")
    outStatement.CompareMode = 1
    outStatement("Type") = LCase$(expectedKeyword)
    outStatement("Line") = CLng(stmtLine)
    mp_TryParseKeywordNoArgStatement = True
End Function

Private Function mp_TryParseForStatement( _
    ByVal sourceText As String, _
    ByRef pos As Long, _
    ByRef lineNo As Long, _
    ByRef outStatement As Object, _
    ByRef outErrorText As String _
) As Boolean
    Dim keywordText As String
    Dim forHeaderText As String
    Dim loopVarName As String
    Dim targetText As String
    Dim tableRef As String
    Dim sourceRowVarName As String
    Dim bodyStatements As Collection
    Dim stmtLine As Long

    stmtLine = lineNo
    If Not mp_ReadIdentifier(sourceText, pos, lineNo, keywordText) Then Exit Function
    If StrComp(keywordText, "for", vbTextCompare) <> 0 Then Exit Function

    mp_SkipWhitespace sourceText, pos, lineNo
    If pos <= Len(sourceText) And Mid$(sourceText, pos, 1) = "(" Then
        If Not mp_ReadBalanced(sourceText, pos, lineNo, "(", ")", forHeaderText, outErrorText) Then
            If Len(outErrorText) = 0 Then outErrorText = "Expected '(let item in Source.Sheet[Table].rows)' after for at line " & CStr(stmtLine)
            Exit Function
        End If
        If Not mp_TryParseForHeaderText(forHeaderText, loopVarName, targetText) Then
            outErrorText = "Expected 'for (let item in Source.Sheet[Table].rows)' at line " & CStr(stmtLine)
            Exit Function
        End If
    Else
        If Not mp_TryParseForHeaderFromStream(sourceText, pos, lineNo, loopVarName, targetText) Then
            outErrorText = "Expected 'for let item in Source.Sheet[Table].rows' at line " & CStr(stmtLine)
            Exit Function
        End If
    End If
    If Not mp_TryValidateScriptVariableName(loopVarName, stmtLine, outErrorText) Then Exit Function

    mp_SkipWhitespace sourceText, pos, lineNo
    If pos > Len(sourceText) Or Mid$(sourceText, pos, 1) <> "{" Then
        outErrorText = "Expected '{' after for-statement at line " & CStr(stmtLine)
        Exit Function
    End If
    pos = pos + 1

    Set bodyStatements = New Collection
    If Not mp_ParseStatements(sourceText, pos, lineNo, bodyStatements, True, outErrorText) Then Exit Function

    Set outStatement = CreateObject("Scripting.Dictionary")
    outStatement.CompareMode = 1
    outStatement("Type") = "for"
    outStatement("LoopVar") = loopVarName

    If ex_obj_ResultTableDsl.m_TryParseTableRowsRef(targetText, tableRef) Then
        outStatement("LoopTarget") = LOOP_TARGET_TABLE_ROWS
        outStatement("TableRef") = tableRef
    ElseIf ex_obj_ResultRowDsl.m_TryParseRowColumnsRef(targetText, sourceRowVarName) Then
        outStatement("LoopTarget") = LOOP_TARGET_ROW_COLUMNS
        outStatement("SourceRowVar") = sourceRowVarName
    Else
        outErrorText = "Invalid for target at line " & CStr(stmtLine) & ". Use Source.Sheet[Table].rows or <rowVar>.columns."
        Exit Function
    End If

    outStatement.Add "Body", bodyStatements
    outStatement("Line") = CLng(stmtLine)
    mp_TryParseForStatement = True
End Function

Private Function mp_TryParseForHeaderFromStream( _
    ByVal sourceText As String, _
    ByRef pos As Long, _
    ByRef lineNo As Long, _
    ByRef outRowVarName As String, _
    ByRef outTargetText As String _
) As Boolean
    Dim loopVarToken As String
    Dim inKeyword As String

    If Not mp_ReadIdentifier(sourceText, pos, lineNo, loopVarToken) Then Exit Function
    If StrComp(loopVarToken, "let", vbTextCompare) = 0 Then
        If Not mp_ReadIdentifier(sourceText, pos, lineNo, outRowVarName) Then Exit Function
    Else
        outRowVarName = loopVarToken
    End If

    mp_SkipWhitespace sourceText, pos, lineNo
    If Not mp_ReadIdentifier(sourceText, pos, lineNo, inKeyword) Then Exit Function
    If StrComp(inKeyword, "in", vbTextCompare) <> 0 Then Exit Function
    mp_SkipWhitespace sourceText, pos, lineNo
    If Not mp_ReadToken(sourceText, pos, lineNo, outTargetText) Then Exit Function

    mp_TryParseForHeaderFromStream = True
End Function

Private Function mp_TryParseForHeaderText( _
    ByVal headerText As String, _
    ByRef outRowVarName As String, _
    ByRef outTargetText As String _
) As Boolean
    Dim bodyText As String
    Dim lowerBody As String
    Dim inPos As Long

    bodyText = Trim$(headerText)
    If Len(bodyText) = 0 Then Exit Function

    If Len(bodyText) >= 4 Then
        If LCase$(Left$(bodyText, 4)) = "let " Then
            bodyText = Trim$(Mid$(bodyText, 5))
            If Len(bodyText) = 0 Then Exit Function
        End If
    End If

    lowerBody = LCase$(bodyText)
    inPos = InStr(1, lowerBody, " in ", vbBinaryCompare)
    If inPos <= 1 Then Exit Function

    outRowVarName = Trim$(Left$(bodyText, inPos - 1))
    If Len(outRowVarName) = 0 Then Exit Function
    If Not ex_PostProcessParserCore.m_IsIdentifier(outRowVarName) Then Exit Function

    outTargetText = Trim$(Mid$(bodyText, inPos + 4))
    If Len(outTargetText) = 0 Then Exit Function

    mp_TryParseForHeaderText = True
End Function

Private Function mp_TryParseIfStatement( _
    ByVal sourceText As String, _
    ByRef pos As Long, _
    ByRef lineNo As Long, _
    ByRef outStatement As Object, _
    ByRef outErrorText As String _
) As Boolean
    Dim keywordText As String
    Dim conditionText As String
    Dim bodyStatements As Collection
    Dim elseBodyStatements As Collection
    Dim probePos As Long
    Dim probeLine As Long
    Dim nextKeyword As String
    Dim stmtLine As Long

    stmtLine = lineNo
    If Not mp_ReadIdentifier(sourceText, pos, lineNo, keywordText) Then Exit Function
    If StrComp(keywordText, "if", vbTextCompare) <> 0 Then Exit Function

    mp_SkipWhitespace sourceText, pos, lineNo
    If Not mp_ReadBalanced(sourceText, pos, lineNo, "(", ")", conditionText, outErrorText) Then
        If Len(outErrorText) = 0 Then outErrorText = "Expected condition in if-statement at line " & CStr(stmtLine)
        Exit Function
    End If
    conditionText = Trim$(conditionText)
    If Len(conditionText) = 0 Then
        outErrorText = "Condition in if-statement cannot be empty at line " & CStr(stmtLine)
        Exit Function
    End If

    mp_SkipWhitespace sourceText, pos, lineNo
    If pos > Len(sourceText) Or Mid$(sourceText, pos, 1) <> "{" Then
        outErrorText = "Expected '{' after if-statement at line " & CStr(stmtLine)
        Exit Function
    End If
    pos = pos + 1

    Set bodyStatements = New Collection
    If Not mp_ParseStatements(sourceText, pos, lineNo, bodyStatements, True, outErrorText) Then Exit Function

    probePos = pos
    probeLine = lineNo
    mp_SkipWhitespace sourceText, probePos, probeLine
    If mp_ReadIdentifier(sourceText, probePos, probeLine, nextKeyword) Then
        If StrComp(nextKeyword, "else", vbTextCompare) = 0 Then
            pos = probePos
            lineNo = probeLine
            mp_SkipWhitespace sourceText, pos, lineNo
            If pos > Len(sourceText) Or Mid$(sourceText, pos, 1) <> "{" Then
                outErrorText = "Expected '{' after else at line " & CStr(stmtLine)
                Exit Function
            End If
            pos = pos + 1
            Set elseBodyStatements = New Collection
            If Not mp_ParseStatements(sourceText, pos, lineNo, elseBodyStatements, True, outErrorText) Then Exit Function
        End If
    End If
    If elseBodyStatements Is Nothing Then Set elseBodyStatements = New Collection

    Set outStatement = CreateObject("Scripting.Dictionary")
    outStatement.CompareMode = 1
    outStatement("Type") = "if"
    outStatement("Condition") = conditionText
    outStatement.Add "Body", bodyStatements
    outStatement.Add "ElseBody", elseBodyStatements
    outStatement("Line") = CLng(stmtLine)
    mp_TryParseIfStatement = True
End Function

Private Function mp_TryParseCallMacroStatement( _
    ByVal sourceText As String, _
    ByRef pos As Long, _
    ByRef lineNo As Long, _
    ByRef outStatement As Object, _
    ByRef outErrorText As String _
) As Boolean
    Dim statementText As String
    Dim stmtLine As Long

    stmtLine = lineNo
    If Not mp_ReadStatementToSemicolon(sourceText, pos, lineNo, statementText, outErrorText) Then Exit Function
    If Not mp_TryParseAction(statementText, outStatement, outErrorText) Then
        outErrorText = outErrorText & " at line " & CStr(stmtLine)
        Exit Function
    End If
    outStatement("Line") = CLng(stmtLine)
    mp_TryParseCallMacroStatement = True
End Function

Private Function mp_TryParseLetStatement( _
    ByVal sourceText As String, _
    ByRef pos As Long, _
    ByRef lineNo As Long, _
    ByRef outStatement As Object, _
    ByRef outErrorText As String _
) As Boolean
    Dim keywordText As String
    Dim varName As String
    Dim rhsStatementText As String
    Dim rhsStatement As Object
    Dim stmtLine As Long

    stmtLine = lineNo
    If Not mp_ReadIdentifier(sourceText, pos, lineNo, keywordText) Then Exit Function
    If StrComp(keywordText, "let", vbTextCompare) <> 0 Then Exit Function

    mp_SkipWhitespace sourceText, pos, lineNo
    If Not mp_ReadIdentifier(sourceText, pos, lineNo, varName) Then
        outErrorText = "Expected variable name after let at line " & CStr(stmtLine)
        Exit Function
    End If
    If Not mp_TryValidateScriptVariableName(varName, stmtLine, outErrorText) Then Exit Function

    mp_SkipWhitespace sourceText, pos, lineNo
    If pos > Len(sourceText) Or Mid$(sourceText, pos, 1) <> "=" Then
        outErrorText = "Expected '=' in let statement at line " & CStr(stmtLine)
        Exit Function
    End If
    pos = pos + 1

    If Not mp_ReadStatementToSemicolon(sourceText, pos, lineNo, rhsStatementText, outErrorText) Then Exit Function
    If Not mp_TryParseAssignmentRhs(rhsStatementText, rhsStatement, outErrorText) Then
        outErrorText = outErrorText & " at line " & CStr(stmtLine)
        Exit Function
    End If

    Set outStatement = rhsStatement
    outStatement("Type") = ACTION_LET
    outStatement("VarName") = varName
    outStatement("Line") = CLng(stmtLine)
    mp_TryParseLetStatement = True
End Function

Private Function mp_TryParseAssignStatement( _
    ByVal sourceText As String, _
    ByRef pos As Long, _
    ByRef lineNo As Long, _
    ByRef outStatement As Object, _
    ByRef outErrorText As String _
) As Boolean
    Dim statementText As String
    Dim rhsActionText As String
    Dim varName As String
    Dim rhsStatement As Object
    Dim stmtLine As Long

    stmtLine = lineNo
    If Not mp_ReadStatementToSemicolon(sourceText, pos, lineNo, statementText, outErrorText) Then Exit Function
    If Not mp_TrySplitAssignmentStatement(statementText, varName, rhsActionText, outErrorText) Then
        If Len(outErrorText) = 0 Then
            outErrorText = "Unsupported statement '" & Trim$(statementText) & "' at line " & CStr(stmtLine)
        End If
        Exit Function
    End If

    If Not mp_TryValidateScriptVariableName(varName, stmtLine, outErrorText) Then Exit Function
    If Not mp_TryParseAssignmentRhs(rhsActionText, rhsStatement, outErrorText) Then
        outErrorText = outErrorText & " at line " & CStr(stmtLine)
        Exit Function
    End If

    Set outStatement = rhsStatement
    outStatement("Type") = ACTION_ASSIGN
    outStatement("VarName") = varName
    outStatement("Line") = CLng(stmtLine)
    mp_TryParseAssignStatement = True
End Function

Private Function mp_TryParseAssignmentRhs( _
    ByVal rhsStatementText As String, _
    ByRef outStatement As Object, _
    ByRef outErrorText As String _
) As Boolean
    Dim trimmedRhs As String
    Dim actionStatement As Object
    Dim exprParts As Collection

    trimmedRhs = Trim$(rhsStatementText)
    If Len(trimmedRhs) = 0 Then
        outErrorText = "Assignment value is empty."
        Exit Function
    End If

    If Left$(LCase$(trimmedRhs), Len("callmacro(")) = "callmacro(" Then
        If Not mp_TryParseAction(trimmedRhs, actionStatement, outErrorText) Then Exit Function
        If LCase$(CStr(actionStatement("Type"))) <> ACTION_CALL_MACRO Then
            outErrorText = "Unsupported assignment action: '" & trimmedRhs & "'."
            Exit Function
        End If
        actionStatement("AssignKind") = ASSIGN_KIND_CALL_MACRO
        Set outStatement = actionStatement
        mp_TryParseAssignmentRhs = True
        Exit Function
    End If

    If Not mp_TryParseStringExpression(trimmedRhs, exprParts, outErrorText) Then Exit Function

    Set outStatement = CreateObject("Scripting.Dictionary")
    outStatement.CompareMode = 1
    outStatement("AssignKind") = ASSIGN_KIND_STRING_EXPR
    outStatement.Add "ExprParts", exprParts
    mp_TryParseAssignmentRhs = True
End Function

Private Function mp_TryParseStringExpression( _
    ByVal expressionText As String, _
    ByRef outParts As Collection, _
    ByRef outErrorText As String _
) As Boolean
    Dim bodyText As String
    Dim terms As Collection
    Dim i As Long
    Dim termText As String
    Dim literalValue As String
    Dim partSpec As Object

    expressionText = Trim$(expressionText)
    If Len(expressionText) = 0 Then
        outErrorText = "String expression is empty."
        Exit Function
    End If
    If Right$(expressionText, 1) <> ";" Then
        outErrorText = "String expression must end with ';'."
        Exit Function
    End If

    bodyText = Trim$(Left$(expressionText, Len(expressionText) - 1))
    If Len(bodyText) = 0 Then
        outErrorText = "String expression is empty."
        Exit Function
    End If

    If Not mp_TrySplitStringExpressionTerms(bodyText, terms, outErrorText) Then Exit Function

    Set outParts = New Collection
    For i = 1 To terms.Count
        termText = Trim$(CStr(terms(i)))
        If Len(termText) = 0 Then
            outErrorText = "String expression contains empty operand."
            Exit Function
        End If

        Set partSpec = CreateObject("Scripting.Dictionary")
        partSpec.CompareMode = 1
        If mp_TryParseQuotedString(termText, literalValue) Then
            partSpec("Kind") = EXPR_PART_LITERAL
            partSpec("Value") = literalValue
        Else
            partSpec("Kind") = EXPR_PART_TOKEN
            partSpec("Value") = termText
        End If
        outParts.Add partSpec
    Next i

    mp_TryParseStringExpression = True
End Function

Private Function mp_TrySplitStringExpressionTerms( _
    ByVal expressionBody As String, _
    ByRef outTerms As Collection, _
    ByRef outErrorText As String _
) As Boolean
    Dim i As Long
    Dim ch As String
    Dim currentTerm As String
    Dim inQuotes As Boolean
    Dim parenDepth As Long
    Dim bracketDepth As Long
    Dim braceDepth As Long

    Set outTerms = New Collection
    currentTerm = vbNullString

    For i = 1 To Len(expressionBody)
        ch = Mid$(expressionBody, i, 1)
        If ch = """" And Not mp_IsEscapedQuote(expressionBody, i) Then
            inQuotes = Not inQuotes
            currentTerm = currentTerm & ch
        ElseIf Not inQuotes Then
            Select Case ch
                Case "("
                    parenDepth = parenDepth + 1
                    currentTerm = currentTerm & ch
                Case ")"
                    If parenDepth > 0 Then parenDepth = parenDepth - 1
                    currentTerm = currentTerm & ch
                Case "["
                    bracketDepth = bracketDepth + 1
                    currentTerm = currentTerm & ch
                Case "]"
                    If bracketDepth > 0 Then bracketDepth = bracketDepth - 1
                    currentTerm = currentTerm & ch
                Case "{"
                    braceDepth = braceDepth + 1
                    currentTerm = currentTerm & ch
                Case "}"
                    If braceDepth > 0 Then braceDepth = braceDepth - 1
                    currentTerm = currentTerm & ch
                Case "+"
                    If parenDepth = 0 And bracketDepth = 0 And braceDepth = 0 Then
                        If Len(Trim$(currentTerm)) = 0 Then
                            outErrorText = "String expression has empty operand before '+'."
                            Exit Function
                        End If
                        outTerms.Add Trim$(currentTerm)
                        currentTerm = vbNullString
                    Else
                        currentTerm = currentTerm & ch
                    End If
                Case Else
                    currentTerm = currentTerm & ch
            End Select
        Else
            currentTerm = currentTerm & ch
        End If
    Next i

    If inQuotes Then
        outErrorText = "Unterminated quoted string in string expression."
        Exit Function
    End If

    If Len(Trim$(currentTerm)) = 0 Then
        outErrorText = "String expression has empty operand after '+'."
        Exit Function
    End If
    outTerms.Add Trim$(currentTerm)

    mp_TrySplitStringExpressionTerms = True
End Function

Private Function mp_ValidateStatements( _
    ByVal statements As Collection, _
    ByVal allowedTableFields As Object, _
    ByVal currentTableRef As String, _
    ByVal currentRowVar As String, _
    ByVal scopeVarTypes As Object, _
    ByVal loopDepth As Long, _
    ByRef outErrorText As String _
) As Boolean
    Dim i As Long
    Dim statement As Object
    Dim statementType As String
    Dim statementLine As Long
    Dim loopTarget As String
    Dim tableRef As String
    Dim loopVarName As String
    Dim sourceRowVarName As String
    Dim childScopeVarTypes As Object
    Dim localLetDeclarations As Object
    Dim varName As String
    Dim expectedType As String
    Dim actualType As String

    If statements Is Nothing Then
        mp_ValidateStatements = True
        Exit Function
    End If

    If scopeVarTypes Is Nothing Then Set scopeVarTypes = mp_CreateVarScope()
    Set localLetDeclarations = mp_CreateVarScope()

    For i = 1 To statements.Count
        Set statement = statements(i)
        statementType = LCase$(CStr(statement("Type")))
        statementLine = 0
        If statement.Exists("Line") Then statementLine = CLng(statement("Line"))

        Select Case statementType
            Case ACTION_CALL_MACRO
                If Not mp_ValidateCallMacroArgs(statement, scopeVarTypes, allowedTableFields, outErrorText) Then Exit Function

            Case ACTION_LET
                varName = CStr(statement("VarName"))
                If Not mp_TryValidateScriptVariableName(varName, statementLine, outErrorText) Then Exit Function
                If localLetDeclarations.Exists(varName) Then
                    outErrorText = "Variable '" & varName & "' is already declared in this scope."
                    Exit Function
                End If
                Select Case mp_GetStatementAssignKind(statement)
                    Case ASSIGN_KIND_STRING_EXPR
                        If Not mp_ValidateStringExpressionParts(statement("ExprParts"), currentTableRef, currentRowVar, scopeVarTypes, allowedTableFields, outErrorText) Then Exit Function
                        actualType = VAR_TYPE_STRING
                    Case ASSIGN_KIND_CALL_MACRO
                        If Not mp_ValidateCallMacroArgs(statement, scopeVarTypes, allowedTableFields, outErrorText) Then Exit Function
                        actualType = mp_InferLetVarType(statement, scopeVarTypes)
                    Case Else
                        outErrorText = "Unsupported assignment kind in let statement: '" & mp_GetStatementAssignKind(statement) & "'."
                        Exit Function
                End Select
                mp_SetScopeValue scopeVarTypes, varName, actualType
                mp_SetScopeValue localLetDeclarations, varName, "1"

            Case ACTION_ASSIGN
                varName = CStr(statement("VarName"))
                If scopeVarTypes Is Nothing Or Not scopeVarTypes.Exists(varName) Then
                    outErrorText = "Assignment to undeclared variable '" & varName & "'. Declare it first via let."
                    Exit Function
                End If
                expectedType = LCase$(CStr(scopeVarTypes(varName)))
                Select Case mp_GetStatementAssignKind(statement)
                    Case ASSIGN_KIND_STRING_EXPR
                        If Not mp_ValidateStringExpressionParts(statement("ExprParts"), currentTableRef, currentRowVar, scopeVarTypes, allowedTableFields, outErrorText) Then Exit Function
                        actualType = VAR_TYPE_STRING
                    Case ASSIGN_KIND_CALL_MACRO
                        If Not mp_ValidateCallMacroArgs(statement, scopeVarTypes, allowedTableFields, outErrorText) Then Exit Function
                        actualType = mp_InferLetVarType(statement, scopeVarTypes)
                    Case Else
                        outErrorText = "Unsupported assignment kind in assignment statement: '" & mp_GetStatementAssignKind(statement) & "'."
                        Exit Function
                End Select
                If StrComp(expectedType, actualType, vbTextCompare) <> 0 Then
                    outErrorText = "Type mismatch in assignment to '" & varName & "': expected " & expectedType & ", got " & actualType & "."
                    Exit Function
                End If

            Case "if"
                If Not mp_ValidateConditionText(CStr(statement("Condition")), currentTableRef, currentRowVar, scopeVarTypes, allowedTableFields, outErrorText) Then Exit Function
                Set childScopeVarTypes = mp_CloneVarScope(scopeVarTypes)
                If Not mp_ValidateStatements(statement("Body"), allowedTableFields, currentTableRef, currentRowVar, childScopeVarTypes, loopDepth, outErrorText) Then Exit Function
                If statement.Exists("ElseBody") Then
                    Set childScopeVarTypes = mp_CloneVarScope(scopeVarTypes)
                    If Not mp_ValidateStatements(statement("ElseBody"), allowedTableFields, currentTableRef, currentRowVar, childScopeVarTypes, loopDepth, outErrorText) Then Exit Function
                End If

            Case "for"
                loopTarget = LCase$(CStr(statement("LoopTarget")))
                loopVarName = CStr(statement("LoopVar"))
                If Not mp_TryValidateScriptVariableName(loopVarName, statementLine, outErrorText) Then Exit Function
                Select Case loopTarget
                    Case LOOP_TARGET_TABLE_ROWS
                        tableRef = CStr(statement("TableRef"))
                        If allowedTableFields Is Nothing Or Not allowedTableFields.Exists(tableRef) Then
                            outErrorText = "Unknown table reference in script: '" & tableRef & "'."
                            Exit Function
                        End If
                        Set childScopeVarTypes = mp_CloneVarScope(scopeVarTypes)
                        mp_SetScopeValue childScopeVarTypes, loopVarName, VAR_TYPE_ROW
                        If Not mp_ValidateStatements(statement("Body"), allowedTableFields, tableRef, loopVarName, childScopeVarTypes, loopDepth + 1, outErrorText) Then Exit Function

                    Case LOOP_TARGET_ROW_COLUMNS
                        sourceRowVarName = CStr(statement("SourceRowVar"))
                        If scopeVarTypes Is Nothing Or Not scopeVarTypes.Exists(sourceRowVarName) Then
                            outErrorText = "Unknown row variable '" & sourceRowVarName & "' in for-loop."
                            Exit Function
                        End If
                        If StrComp(CStr(scopeVarTypes(sourceRowVarName)), VAR_TYPE_ROW, vbTextCompare) <> 0 Then
                            outErrorText = "Variable '" & sourceRowVarName & "' must be a row variable to use .columns iteration."
                            Exit Function
                        End If
                        Set childScopeVarTypes = mp_CloneVarScope(scopeVarTypes)
                        mp_SetScopeValue childScopeVarTypes, loopVarName, VAR_TYPE_COLUMN
                        If Not mp_ValidateStatements(statement("Body"), allowedTableFields, currentTableRef, sourceRowVarName, childScopeVarTypes, loopDepth + 1, outErrorText) Then Exit Function

                    Case Else
                        outErrorText = "Unsupported for-loop target '" & loopTarget & "'."
                        Exit Function
                End Select

            Case ACTION_BREAK
                If loopDepth <= 0 Then
                    outErrorText = "'break' is only allowed inside for-loop."
                    Exit Function
                End If

            Case ACTION_CONTINUE
                If loopDepth <= 0 Then
                    outErrorText = "'continue' is only allowed inside for-loop."
                    Exit Function
                End If

            Case ACTION_RETURN
                ' Allowed at any script nesting level.

            Case Else
                outErrorText = "Unsupported statement type '" & statementType & "'."
                Exit Function
        End Select
    Next i

    mp_ValidateStatements = True
End Function

Private Function mp_ValidateConditionText( _
    ByVal conditionText As String, _
    ByVal currentTableRef As String, _
    ByVal currentRowVar As String, _
    ByVal scopeVarTypes As Object, _
    ByVal allowedTableFields As Object, _
    ByRef outErrorText As String _
) As Boolean
    Dim condParts As Collection
    Dim condOps As Collection
    Dim i As Long
    Dim leftTokenText As String
    Dim rightTokenText As String
    Dim opText As String
    Dim rightIsToken As Boolean
    Dim resolvedTableRef As String
    Dim resolvedMapKey As String

    If Not mp_TrySplitConditionExpression(conditionText, condParts, condOps, outErrorText) Then Exit Function

    For i = 1 To condParts.Count
        If Not mp_ParseConditionPart(CStr(condParts(i)), leftTokenText, opText, rightTokenText, rightIsToken) Then
            outErrorText = "Unsupported condition token: '" & Trim$(CStr(condParts(i))) & "'."
            Exit Function
        End If
        If Not mp_TryResolveConditionTokenForValidation(leftTokenText, currentTableRef, currentRowVar, scopeVarTypes, allowedTableFields, resolvedTableRef, resolvedMapKey, outErrorText) Then
            Exit Function
        End If
        If rightIsToken Then
            If Not mp_TryResolveConditionTokenForValidation(rightTokenText, currentTableRef, currentRowVar, scopeVarTypes, allowedTableFields, resolvedTableRef, resolvedMapKey, outErrorText) Then
                Exit Function
            End If
        End If
    Next i

    mp_ValidateConditionText = True
End Function

Private Function mp_ValidateStringExpressionParts( _
    ByVal exprParts As Collection, _
    ByVal currentTableRef As String, _
    ByVal currentRowVar As String, _
    ByVal scopeVarTypes As Object, _
    ByVal allowedTableFields As Object, _
    ByRef outErrorText As String _
) As Boolean
    Dim i As Long
    Dim partSpec As Object
    Dim partKind As String
    Dim tokenText As String
    Dim resolvedTableRef As String
    Dim resolvedMapKey As String
    Dim variableName As String
    Dim memberName As String
    Dim variableType As String

    If exprParts Is Nothing Or exprParts.Count = 0 Then
        outErrorText = "String expression is empty."
        Exit Function
    End If

    For i = 1 To exprParts.Count
        Set partSpec = exprParts(i)
        If partSpec Is Nothing Then
            outErrorText = "String expression contains invalid operand."
            Exit Function
        End If
        partKind = LCase$(Trim$(CStr(partSpec("Kind"))))

        Select Case partKind
            Case EXPR_PART_LITERAL
                ' always valid

            Case EXPR_PART_TOKEN
                tokenText = Trim$(CStr(partSpec("Value")))
                If Len(tokenText) = 0 Then
                    outErrorText = "String expression contains empty token operand."
                    Exit Function
                End If

                If ex_PostProcessParserCore.m_IsIdentifier(tokenText) Then
                    If scopeVarTypes Is Nothing Or Not scopeVarTypes.Exists(tokenText) Then
                        outErrorText = "Unknown variable '" & tokenText & "' in string expression."
                        Exit Function
                    End If
                    variableType = LCase$(CStr(scopeVarTypes(tokenText)))
                    If StrComp(variableType, VAR_TYPE_STRING, vbTextCompare) <> 0 Then
                        outErrorText = "Variable '" & tokenText & "' is type '" & variableType & "' and cannot be concatenated as string."
                        Exit Function
                    End If
                    GoTo ContinueLoop
                End If

                If mp_TryParseVariableMemberRef(tokenText, variableName, memberName) Then
                    If scopeVarTypes Is Nothing Or Not scopeVarTypes.Exists(variableName) Then
                        outErrorText = "Unknown variable '" & variableName & "' in string expression."
                        Exit Function
                    End If
                    variableType = LCase$(CStr(scopeVarTypes(variableName)))
                    Select Case variableType
                        Case VAR_TYPE_COLUMN
                            If Not ex_PostProcessDslContracts.m_IsMemberAllowed(ex_PostProcessDslContracts.TYPE_COLUMN, memberName) Then
                                outErrorText = "Unsupported column member '" & memberName & "' in string expression token '" & tokenText & "'."
                                Exit Function
                            End If
                        Case Else
                            outErrorText = "Variable '" & variableName & "' does not support member access in string expression token '" & tokenText & "'."
                            Exit Function
                    End Select
                    GoTo ContinueLoop
                End If

                If Not mp_TryResolveConditionTokenForValidation(tokenText, currentTableRef, currentRowVar, scopeVarTypes, allowedTableFields, resolvedTableRef, resolvedMapKey, outErrorText) Then
                    Exit Function
                End If

            Case Else
                outErrorText = "Unsupported string expression operand kind '" & partKind & "'."
                Exit Function
        End Select
ContinueLoop:
    Next i

    mp_ValidateStringExpressionParts = True
End Function

Private Function mp_TryValidateScriptVariableName( _
    ByVal variableName As String, _
    ByVal lineNo As Long, _
    ByRef outErrorText As String _
) As Boolean
    variableName = Trim$(variableName)
    If Not ex_PostProcessParserCore.m_IsIdentifier(variableName) Then
        If lineNo > 0 Then
            outErrorText = "Invalid variable name '" & variableName & "' at line " & CStr(lineNo) & "."
        Else
            outErrorText = "Invalid variable name '" & variableName & "'."
        End If
        Exit Function
    End If

    If mp_IsReservedDslKeyword(variableName) Then
        If lineNo > 0 Then
            outErrorText = "Variable name '" & variableName & "' is reserved keyword at line " & CStr(lineNo) & "."
        Else
            outErrorText = "Variable name '" & variableName & "' is reserved keyword."
        End If
        Exit Function
    End If

    mp_TryValidateScriptVariableName = True
End Function

Private Function mp_IsReservedDslKeyword(ByVal tokenText As String) As Boolean
    Select Case LCase$(Trim$(tokenText))
        Case "if", "else", "for", "callmacro", "let", "in", "and", "or", "gt", "lt", "gte", "lte", "break", "continue", "return"
            mp_IsReservedDslKeyword = True
    End Select
End Function

Private Sub mp_SkipWhitespace(ByVal sourceText As String, ByRef pos As Long, ByRef lineNo As Long)
    Dim ch As String
    Do While pos <= Len(sourceText)
        ch = Mid$(sourceText, pos, 1)
        If ch = vbLf Then
            lineNo = lineNo + 1
            pos = pos + 1
        ElseIf ch = " " Or ch = vbTab Or ch = vbCr Then
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop
End Sub

Private Function mp_ReadIdentifier( _
    ByVal sourceText As String, _
    ByRef pos As Long, _
    ByRef lineNo As Long, _
    ByRef outIdentifier As String _
) As Boolean
    Dim startPos As Long
    Dim ch As String

    mp_SkipWhitespace sourceText, pos, lineNo
    If pos > Len(sourceText) Then Exit Function

    ch = Mid$(sourceText, pos, 1)
    If Not ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or ch = "_") Then Exit Function

    startPos = pos
    pos = pos + 1
    Do While pos <= Len(sourceText)
        ch = Mid$(sourceText, pos, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "_" Then
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop

    outIdentifier = Mid$(sourceText, startPos, pos - startPos)
    mp_ReadIdentifier = True
End Function

Private Function mp_ReadToken( _
    ByVal sourceText As String, _
    ByRef pos As Long, _
    ByRef lineNo As Long, _
    ByRef outToken As String _
) As Boolean
    Dim startPos As Long
    Dim ch As String

    mp_SkipWhitespace sourceText, pos, lineNo
    If pos > Len(sourceText) Then Exit Function

    startPos = pos
    Do While pos <= Len(sourceText)
        ch = Mid$(sourceText, pos, 1)
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Or ch = "{" Or ch = "}" Or ch = ";" Then Exit Do
        pos = pos + 1
    Loop

    outToken = Trim$(Mid$(sourceText, startPos, pos - startPos))
    mp_ReadToken = (Len(outToken) > 0)
End Function

Private Function mp_ReadBalanced( _
    ByVal sourceText As String, _
    ByRef pos As Long, _
    ByRef lineNo As Long, _
    ByVal openChar As String, _
    ByVal closeChar As String, _
    ByRef outInnerText As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim startPos As Long
    Dim depth As Long
    Dim i As Long
    Dim ch As String
    Dim inQuotes As Boolean

    mp_SkipWhitespace sourceText, pos, lineNo
    If pos > Len(sourceText) Or Mid$(sourceText, pos, 1) <> openChar Then Exit Function

    startPos = pos + 1
    depth = 1
    i = pos + 1
    Do While i <= Len(sourceText)
        ch = Mid$(sourceText, i, 1)
        If ch = vbLf Then lineNo = lineNo + 1

        If ch = """" And Not mp_IsEscapedQuote(sourceText, i) Then
            inQuotes = Not inQuotes
        ElseIf Not inQuotes Then
            If ch = openChar Then
                depth = depth + 1
            ElseIf ch = closeChar Then
                depth = depth - 1
                If depth = 0 Then
                    outInnerText = Mid$(sourceText, startPos, i - startPos)
                    pos = i + 1
                    mp_ReadBalanced = True
                    Exit Function
                End If
            End If
        End If
        i = i + 1
    Loop

    outErrorText = "Missing '" & closeChar & "' for expression started at line " & CStr(lineNo)
End Function

Private Function mp_ReadStatementToSemicolon( _
    ByVal sourceText As String, _
    ByRef pos As Long, _
    ByRef lineNo As Long, _
    ByRef outStatementText As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim startPos As Long
    Dim i As Long
    Dim ch As String
    Dim inQuotes As Boolean
    Dim depth As Long

    mp_SkipWhitespace sourceText, pos, lineNo
    If pos > Len(sourceText) Then Exit Function

    startPos = pos
    i = pos
    Do While i <= Len(sourceText)
        ch = Mid$(sourceText, i, 1)
        If ch = vbLf Then lineNo = lineNo + 1

        If ch = """" And Not mp_IsEscapedQuote(sourceText, i) Then
            inQuotes = Not inQuotes
        ElseIf Not inQuotes Then
            If ch = "(" Then
                depth = depth + 1
            ElseIf ch = ")" Then
                depth = depth - 1
                If depth < 0 Then
                    outErrorText = "Unexpected ')' at line " & CStr(lineNo)
                    Exit Function
                End If
            ElseIf ch = ";" And depth = 0 Then
                outStatementText = Trim$(Mid$(sourceText, startPos, i - startPos + 1))
                pos = i + 1
                mp_ReadStatementToSemicolon = True
                Exit Function
            End If
        End If

        i = i + 1
    Loop

    outErrorText = "Missing ';' at end of statement near line " & CStr(lineNo)
End Function

Private Function mp_EvaluateCondition( _
    ByVal conditionText As String, _
    ByVal currentTableRef As String, _
    ByVal currentRowVar As String, _
    ByVal currentRowRef As obj_ResultRow, _
    ByVal tablesByRef As Object, _
    ByVal runtimeVars As Object _
) As Boolean
    Dim condParts As Collection
    Dim condOps As Collection
    Dim i As Long
    Dim refToken As String
    Dim opText As String
    Dim boolOp As String
    Dim expectedValueRaw As String
    Dim expectedValue As String
    Dim expectedIsToken As Boolean
    Dim actualValue As String
    Dim compareResult As Long
    Dim resolveError As String
    Dim partResult As Boolean
    Dim currentTerm As Boolean
    Dim finalResult As Boolean
    Dim hasCurrentTerm As Boolean
    Dim hasFinalResult As Boolean

    If Not mp_TrySplitConditionExpression(conditionText, condParts, condOps, resolveError) Then
        Err.Raise vbObjectError + 1594, "ex_PostProcessDsl", resolveError
    End If

    For i = 1 To condParts.Count
        If Not mp_ParseConditionPart(CStr(condParts(i)), refToken, opText, expectedValueRaw, expectedIsToken) Then
            Err.Raise vbObjectError + 1594, "ex_PostProcessDsl", "Invalid condition: " & Trim$(CStr(condParts(i)))
        End If
        If Not mp_TryResolveRuntimeValue(refToken, currentTableRef, currentRowVar, currentRowRef, tablesByRef, runtimeVars, actualValue, resolveError) Then
            Err.Raise vbObjectError + 1595, "ex_PostProcessDsl", resolveError
        End If
        If expectedIsToken Then
            If Not mp_TryResolveRuntimeValue(expectedValueRaw, currentTableRef, currentRowVar, currentRowRef, tablesByRef, runtimeVars, expectedValue, resolveError) Then
                Err.Raise vbObjectError + 1595, "ex_PostProcessDsl", resolveError
            End If
        Else
            expectedValue = expectedValueRaw
        End If

        compareResult = mp_CompareConditionValues(actualValue, expectedValue)
        Select Case opText
            Case "=="
                partResult = (compareResult = 0)
            Case "!="
                partResult = (compareResult <> 0)
            Case "gt"
                partResult = (compareResult > 0)
            Case "lt"
                partResult = (compareResult < 0)
            Case "gte"
                partResult = (compareResult >= 0)
            Case "lte"
                partResult = (compareResult <= 0)
            Case Else
                Err.Raise vbObjectError + 1596, "ex_PostProcessDsl", "Unsupported operator in condition: " & opText
        End Select

        If Not hasCurrentTerm Then
            currentTerm = partResult
            hasCurrentTerm = True
        Else
            boolOp = LCase$(Trim$(CStr(condOps(i - 1))))
            Select Case boolOp
                Case "and"
                    currentTerm = (currentTerm And partResult)
                Case "or"
                    If Not hasFinalResult Then
                        finalResult = currentTerm
                        hasFinalResult = True
                    Else
                        finalResult = (finalResult Or currentTerm)
                    End If
                    currentTerm = partResult
                Case Else
                    Err.Raise vbObjectError + 1596, "ex_PostProcessDsl", "Unsupported boolean operator in condition: " & boolOp
            End Select
        End If
    Next i

    If hasFinalResult Then
        mp_EvaluateCondition = (finalResult Or currentTerm)
    Else
        mp_EvaluateCondition = currentTerm
    End If

End Function

Private Function mp_CompareConditionValues(ByVal actualValue As String, ByVal expectedValue As String) As Long
    Dim leftNumber As Double
    Dim rightNumber As Double
    Dim leftText As String
    Dim rightText As String

    leftText = Trim$(actualValue)
    rightText = Trim$(expectedValue)

    If ex_XmlCore.m_TryParseDouble(leftText, leftNumber, True) And ex_XmlCore.m_TryParseDouble(rightText, rightNumber, True) Then
        If leftNumber < rightNumber Then
            mp_CompareConditionValues = -1
        ElseIf leftNumber > rightNumber Then
            mp_CompareConditionValues = 1
        Else
            mp_CompareConditionValues = 0
        End If
        Exit Function
    End If

    mp_CompareConditionValues = StrComp(actualValue, expectedValue, vbTextCompare)
End Function

Private Function mp_TryParseAction(ByVal lineText As String, ByRef outAction As Object, ByRef outErrorText As String) As Boolean
    Dim payload As String
    Dim macroName As String
    Dim argSpecs As Collection

    Set outAction = CreateObject("Scripting.Dictionary")
    outAction.CompareMode = 1

    If Left$(LCase$(lineText), Len("callmacro(")) = "callmacro(" And Right$(lineText, 2) = ");" Then
        payload = Trim$(Mid$(lineText, Len("callMacro(") + 1, Len(lineText) - Len("callMacro(") - 2))
        If Not mp_TryParseCallMacroArgs(payload, macroName, argSpecs, outErrorText) Then
            Exit Function
        End If
        outAction("Type") = ACTION_CALL_MACRO
        outAction("MacroName") = macroName
        outAction.Add "Args", argSpecs
        mp_TryParseAction = True
        Exit Function
    End If

    outErrorText = "Unsupported action: '" & lineText & "'. Only callMacro(...) is supported."
End Function

Private Function mp_ParseConditionPart( _
    ByVal rawPart As String, _
    ByRef outFieldName As String, _
    ByRef outOp As String, _
    ByRef outValue As String, _
    ByRef outValueIsToken As Boolean _
) As Boolean
    Dim part As String
    Dim partLower As String
    Dim opPos As Long
    Dim opLen As Long
    Dim rhs As String

    part = mp_TrimDslWhitespace(rawPart)
    partLower = LCase$(part)
    opPos = InStr(1, part, "==", vbBinaryCompare)
    If opPos > 0 Then
        outOp = "=="
        opLen = 2
    Else
        opPos = InStr(1, part, "!=", vbBinaryCompare)
        If opPos > 0 Then
            outOp = "!="
            opLen = 2
        ElseIf mp_TryFindConditionWordOperator(partLower, "gte", opPos, opLen) Then
            outOp = "gte"
        ElseIf mp_TryFindConditionWordOperator(partLower, "lte", opPos, opLen) Then
            outOp = "lte"
        ElseIf mp_TryFindConditionWordOperator(partLower, "gt", opPos, opLen) Then
            outOp = "gt"
        ElseIf mp_TryFindConditionWordOperator(partLower, "lt", opPos, opLen) Then
            outOp = "lt"
        End If
    End If
    If opPos <= 1 Then Exit Function

    outFieldName = mp_TrimDslWhitespace(Left$(part, opPos - 1))
    rhs = mp_TrimDslWhitespace(Mid$(part, opPos + opLen))
    If Len(outFieldName) = 0 Then Exit Function
    If Len(rhs) = 0 Then Exit Function
    If mp_TryParseQuotedString(rhs, outValue) Then
        outValueIsToken = False
    ElseIf mp_ShouldTreatConditionRhsAsToken(rhs) Then
        outValue = rhs
        outValueIsToken = True
    Else
        outValue = rhs
        outValueIsToken = False
    End If

    mp_ParseConditionPart = True
End Function

Private Function mp_TryFindConditionWordOperator( _
    ByVal textValue As String, _
    ByVal opWord As String, _
    ByRef outPos As Long, _
    ByRef outLen As Long _
) As Boolean
    Dim i As Long
    Dim inQuotes As Boolean
    Dim ch As String
    Dim prevCh As String
    Dim nextCh As String
    Dim opText As String

    outLen = Len(opWord)
    If outLen = 0 Then Exit Function
    If Len(textValue) < outLen Then Exit Function

    For i = 1 To Len(textValue) - outLen + 1
        ch = Mid$(textValue, i, 1)
        If ch = """" And Not mp_IsEscapedQuote(textValue, i) Then
            inQuotes = Not inQuotes
            GoTo ContinueLoop
        End If

        If Not inQuotes Then
            opText = Mid$(textValue, i, outLen)
            If opText = opWord Then
                If i > 1 Then
                    prevCh = Mid$(textValue, i - 1, 1)
                    If mp_IsConditionIdentifierChar(prevCh) Then GoTo ContinueLoop
                End If

                If i + outLen <= Len(textValue) Then
                    nextCh = Mid$(textValue, i + outLen, 1)
                    If mp_IsConditionIdentifierChar(nextCh) Then GoTo ContinueLoop
                End If

                outPos = i
                mp_TryFindConditionWordOperator = True
                Exit Function
            End If
        End If
ContinueLoop:
    Next i
End Function

Private Function mp_TryExtractConditionField(ByVal rawPart As String, ByRef outFieldName As String) As Boolean
    Dim opText As String
    Dim valueText As String
    Dim valueIsToken As Boolean
    mp_TryExtractConditionField = mp_ParseConditionPart(rawPart, outFieldName, opText, valueText, valueIsToken)
End Function

Private Function mp_ShouldTreatConditionRhsAsToken(ByVal rhsText As String) As Boolean
    Dim rowVarName As String
    Dim fieldAlias As String
    Dim variableName As String
    Dim memberName As String

    rhsText = Trim$(rhsText)
    If Len(rhsText) = 0 Then Exit Function

    If ex_PostProcessParserCore.m_IsIdentifier(rhsText) Then
        mp_ShouldTreatConditionRhsAsToken = True
        Exit Function
    End If
    If ex_obj_ResultRowDsl.m_TryParseRowColumnRef(rhsText, rowVarName, fieldAlias) Then
        mp_ShouldTreatConditionRhsAsToken = True
        Exit Function
    End If
    If mp_TryParseVariableMemberRef(rhsText, variableName, memberName) Then
        mp_ShouldTreatConditionRhsAsToken = True
    End If
End Function

Private Function mp_TrySplitConditionExpression( _
    ByVal conditionText As String, _
    ByRef outParts As Collection, _
    ByRef outOps As Collection, _
    ByRef outErrorText As String _
) As Boolean
    Dim i As Long
    Dim ch As String
    Dim partText As String
    Dim inQuotes As Boolean

    Set outParts = New Collection
    Set outOps = New Collection

    conditionText = Trim$(conditionText)
    If Len(conditionText) = 0 Then
        outErrorText = "Condition cannot be empty."
        Exit Function
    End If

    i = 1
    Do While i <= Len(conditionText)
        ch = Mid$(conditionText, i, 1)
        If ch = """" And Not mp_IsEscapedQuote(conditionText, i) Then
            inQuotes = Not inQuotes
            partText = partText & ch
            i = i + 1
            GoTo ContinueLoop
        End If

        If Not inQuotes Then
            If mp_IsWordOperatorAt(conditionText, i, "and") Then
                If Not mp_TryPushConditionPart(partText, "and", outParts, outOps, outErrorText) Then Exit Function
                partText = vbNullString
                i = i + 3
                GoTo ContinueLoop
            End If

            If mp_IsWordOperatorAt(conditionText, i, "or") Then
                If Not mp_TryPushConditionPart(partText, "or", outParts, outOps, outErrorText) Then Exit Function
                partText = vbNullString
                i = i + 2
                GoTo ContinueLoop
            End If
        End If

        partText = partText & ch
        i = i + 1
ContinueLoop:
    Loop

    If inQuotes Then
        outErrorText = "Unterminated quoted string in condition."
        Exit Function
    End If

    partText = Trim$(partText)
    If Len(partText) = 0 Then
        outErrorText = "Invalid condition: missing token after boolean operator."
        Exit Function
    End If
    outParts.Add partText

    mp_TrySplitConditionExpression = True
End Function

Private Function mp_TryPushConditionPart( _
    ByVal partText As String, _
    ByVal boolOp As String, _
    ByVal outParts As Collection, _
    ByVal outOps As Collection, _
    ByRef outErrorText As String _
) As Boolean
    partText = Trim$(partText)
    If Len(partText) = 0 Then
        outErrorText = "Invalid condition: missing token before boolean operator '" & boolOp & "'."
        Exit Function
    End If
    outParts.Add partText
    outOps.Add boolOp
    mp_TryPushConditionPart = True
End Function

Private Function mp_IsWordOperatorAt(ByVal textValue As String, ByVal pos As Long, ByVal wordOp As String) As Boolean
    Dim opLen As Long
    Dim prevCh As String
    Dim nextCh As String
    Dim opText As String

    opLen = Len(wordOp)
    If pos + opLen - 1 > Len(textValue) Then Exit Function

    opText = LCase$(Mid$(textValue, pos, opLen))
    If opText <> wordOp Then Exit Function

    If pos > 1 Then
        prevCh = Mid$(textValue, pos - 1, 1)
        If mp_IsConditionIdentifierChar(prevCh) Then Exit Function
    End If

    If pos + opLen <= Len(textValue) Then
        nextCh = Mid$(textValue, pos + opLen, 1)
        If mp_IsConditionIdentifierChar(nextCh) Then Exit Function
    End If

    mp_IsWordOperatorAt = True
End Function

Private Function mp_IsConditionIdentifierChar(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    mp_IsConditionIdentifierChar = _
        ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "_")
End Function

Private Function mp_TryParseQuotedString(ByVal valueText As String, ByRef outValue As String) As Boolean
    Dim rawInner As String
    valueText = Trim$(valueText)
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

Private Function mp_RenderTemplate( _
    ByVal templateText As String, _
    ByVal currentTableRef As String, _
    ByVal currentRowVar As String, _
    ByVal currentRowRef As obj_ResultRow, _
    ByVal tablesByRef As Object, _
    ByVal runtimeVars As Object _
) As String
    Dim result As String
    Dim openPos As Long
    Dim closePos As Long
    Dim tokenText As String
    Dim tokenRef As String
    Dim formatterName As String
    Dim tokenValue As String
    Dim resolveError As String

    result = templateText
    openPos = InStr(1, result, "{", vbBinaryCompare)
    Do While openPos > 0
        closePos = InStr(openPos + 1, result, "}", vbBinaryCompare)
        If closePos <= openPos Then Exit Do

        tokenText = Mid$(result, openPos + 1, closePos - openPos - 1)
        mp_SplitTemplateTokenFormatter tokenText, tokenRef, formatterName

        If Not mp_TryResolveRuntimeValue(tokenRef, currentTableRef, currentRowVar, currentRowRef, tablesByRef, runtimeVars, tokenValue, resolveError) Then
            Err.Raise vbObjectError + 1595, "ex_PostProcessDsl", resolveError
        End If

        If Len(formatterName) > 0 Then
            tokenValue = ex_ResultTemplatesParser.m_FormatValue(tokenValue, formatterName)
        End If

        result = Left$(result, openPos - 1) & tokenValue & Mid$(result, closePos + 1)
        openPos = InStr(openPos + Len(tokenValue), result, "{", vbBinaryCompare)
    Loop

    mp_RenderTemplate = result
End Function

Private Sub mp_SplitTemplateTokenFormatter( _
    ByVal tokenText As String, _
    ByRef outTokenRef As String, _
    ByRef outFormatterName As String _
)
    Dim i As Long
    Dim ch As String
    Dim bracketDepth As Long
    Dim sepPos As Long

    tokenText = CStr(tokenText)
    outTokenRef = Trim$(tokenText)
    outFormatterName = vbNullString

    For i = 1 To Len(tokenText)
        ch = Mid$(tokenText, i, 1)
        Select Case ch
            Case "["
                bracketDepth = bracketDepth + 1
            Case "]"
                If bracketDepth > 0 Then bracketDepth = bracketDepth - 1
            Case "|"
                If bracketDepth = 0 Then
                    sepPos = i
                    Exit For
                End If
        End Select
    Next i

    If sepPos <= 0 Then Exit Sub

    outTokenRef = Trim$(Left$(tokenText, sepPos - 1))
    outFormatterName = Trim$(Mid$(tokenText, sepPos + 1))
    If Len(outTokenRef) = 0 Or Len(outFormatterName) = 0 Then
        outTokenRef = Trim$(tokenText)
        outFormatterName = vbNullString
    End If
End Sub

Private Function mp_NormalizeScript(ByVal scriptText As String) As String
    Dim lines As Variant
    Dim i As Long
    Dim rawLine As String
    Dim cleaned As String
    Dim normalized As String

    scriptText = Replace(scriptText, vbCrLf, vbLf)
    scriptText = Replace(scriptText, vbCr, vbLf)
    scriptText = mp_StripMultiLineComments(scriptText)
    lines = Split(scriptText, vbLf)

    For i = LBound(lines) To UBound(lines)
        rawLine = CStr(lines(i))
        rawLine = Replace(rawLine, vbTab, " ")
        rawLine = Replace(rawLine, ChrW$(160), " ")
        rawLine = mp_StripSingleLineComment(rawLine)
        cleaned = Trim$(rawLine)
        If Len(normalized) > 0 Then normalized = normalized & vbLf
        normalized = normalized & cleaned
    Next i

    mp_NormalizeScript = normalized
End Function

Private Function mp_StripMultiLineComments(ByVal sourceText As String) As String
    Dim i As Long
    Dim ch As String
    Dim nextCh As String
    Dim inQuotes As Boolean
    Dim inCommentBlock As Boolean
    Dim result As String

    i = 1
    Do While i <= Len(sourceText)
        ch = Mid$(sourceText, i, 1)

        If inCommentBlock Then
            If ch = "*" And i < Len(sourceText) Then
                nextCh = Mid$(sourceText, i + 1, 1)
                If nextCh = "/" Then
                    inCommentBlock = False
                    i = i + 2
                    GoTo ContinueLoop
                End If
            End If

            If ch = vbLf Then result = result & vbLf
            i = i + 1
            GoTo ContinueLoop
        End If

        If ch = """" And Not mp_IsEscapedQuote(sourceText, i) Then
            inQuotes = Not inQuotes
            result = result & ch
            i = i + 1
            GoTo ContinueLoop
        End If

        If Not inQuotes And ch = "/" And i < Len(sourceText) Then
            nextCh = Mid$(sourceText, i + 1, 1)
            If nextCh = "*" Then
                inCommentBlock = True
                result = result & " "
                i = i + 2
                GoTo ContinueLoop
            End If
        End If

        result = result & ch
        i = i + 1
ContinueLoop:
    Loop

    mp_StripMultiLineComments = result
End Function

Private Function mp_StripSingleLineComment(ByVal lineText As String) As String
    Dim i As Long
    Dim inQuotes As Boolean
    Dim ch As String
    Dim nextCh As String

    For i = 1 To Len(lineText)
        ch = Mid$(lineText, i, 1)
        If ch = """" And Not mp_IsEscapedQuote(lineText, i) Then
            inQuotes = Not inQuotes
        End If
        If Not inQuotes And ch = "/" Then
            If i < Len(lineText) Then
                nextCh = Mid$(lineText, i + 1, 1)
                If nextCh = "/" Then
                    mp_StripSingleLineComment = Left$(lineText, i - 1)
                    Exit Function
                End If
            End If
        End If
    Next i

    mp_StripSingleLineComment = lineText
End Function

Private Function mp_TryResolveConditionTokenForValidation( _
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

    tokenText = Trim$(tokenText)
    If Len(tokenText) = 0 Then
        outErrorText = "Field reference is empty."
        Exit Function
    End If

    If ex_PostProcessParserCore.m_IsIdentifier(tokenText) Then
        If scopeVarTypes Is Nothing Or Not scopeVarTypes.Exists(tokenText) Then
            outErrorText = "Unknown variable '" & tokenText & "' in condition."
            Exit Function
        End If
        outResolvedTableRef = vbNullString
        outResolvedMapKey = vbNullString
        mp_TryResolveConditionTokenForValidation = True
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
                If Not ex_PostProcessDslContracts.m_IsMemberAllowed(ex_PostProcessDslContracts.TYPE_COLUMN, memberName) Then
                    outErrorText = "Unsupported column member '" & memberName & "' in token '" & tokenText & "'."
                    Exit Function
                End If
            Case Else
                outErrorText = "Variable '" & variableName & "' does not support member access in token '" & tokenText & "'."
                Exit Function
        End Select

        outResolvedTableRef = vbNullString
        outResolvedMapKey = vbNullString
        mp_TryResolveConditionTokenForValidation = True
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
        mp_TryResolveConditionTokenForValidation = True
        Exit Function
    End If

    mp_TryResolveConditionTokenForValidation = ex_ResultRuntimeAdapter.m_TryResolveConditionTokenForValidation( _
        tokenText, _
        currentTableRef, _
        currentRowVar, _
        allowedTableFields, _
        outResolvedTableRef, _
        outResolvedMapKey, _
        outErrorText _
    )
End Function

Private Function mp_TryResolveRuntimeValue( _
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

    tokenText = Trim$(tokenText)
    If Len(tokenText) = 0 Then
        outErrorText = "Field reference is empty."
        Exit Function
    End If

    If ex_PostProcessParserCore.m_IsIdentifier(tokenText) Then
        If runtimeVars Is Nothing Or Not runtimeVars.Exists(tokenText) Then
            outErrorText = "Unknown variable '" & tokenText & "'."
            Exit Function
        End If
        If Not mp_TryConvertScopeEntryToString(runtimeVars, tokenText, outValue, outErrorText) Then Exit Function
        mp_TryResolveRuntimeValue = True
        Exit Function
    End If

    If mp_TryParseVariableMemberRef(tokenText, variableName, memberName) Then
        If Not mp_TryResolveScopeMemberValue(runtimeVars, variableName, memberName, tokenText, outValue, outErrorText) Then Exit Function
        mp_TryResolveRuntimeValue = True
        Exit Function
    End If

    If ex_obj_ResultRowDsl.m_TryParseRowColumnRef(tokenText, rowVarName, fieldAlias) Then
        If Not mp_TryResolveScopedRowCellValue(runtimeVars, rowVarName, fieldAlias, tokenText, outValue, outErrorText) Then Exit Function
        mp_TryResolveRuntimeValue = True
        Exit Function
    End If

    mp_TryResolveRuntimeValue = ex_ResultRuntimeAdapter.m_TryResolveRuntimeValue( _
        tokenText, _
        currentTableRef, _
        currentRowVar, _
        currentRowRef, _
        tablesByRef, _
        outValue, _
        outErrorText _
    )
End Function

Private Function mp_TryParseCallMacroArgs( _
    ByVal argsText As String, _
    ByRef outMacroName As String, _
    ByRef outArgSpecs As Collection, _
    ByRef outErrorText As String _
) As Boolean
    Dim parts As Collection
    Dim i As Long
    Dim partText As String
    Dim argSpec As Object

    argsText = Trim$(argsText)
    If Len(argsText) = 0 Then
        outErrorText = "callMacro requires at least macro name: callMacro(""Module.Proc"", ...)"
        Exit Function
    End If

    If Not mp_SplitArgs(argsText, parts, outErrorText) Then Exit Function
    If parts Is Nothing Or parts.Count = 0 Then
        outErrorText = "callMacro requires at least macro name: callMacro(""Module.Proc"", ...)"
        Exit Function
    End If

    If Not mp_TryParseQuotedString(CStr(parts(1)), outMacroName) Then
        outErrorText = "callMacro first argument must be quoted macro name."
        Exit Function
    End If
    outMacroName = Trim$(outMacroName)
    If Len(outMacroName) = 0 Then
        outErrorText = "callMacro macro name cannot be empty."
        Exit Function
    End If

    Set outArgSpecs = New Collection
    For i = 2 To parts.Count
        partText = Trim$(CStr(parts(i)))
        If Len(partText) = 0 Then
            outErrorText = "callMacro argument #" & CStr(i - 1) & " is empty."
            Exit Function
        End If
        If Not mp_TryParseMacroArg(partText, argSpec) Then
            outErrorText = "Unsupported callMacro argument '" & partText & "'. Use variable, quoted string, Source.Sheet[Table].row[N], Source.Sheet[Table].lastRow, Source.Sheet[Table].prevRow, or a .column[Field] variant."
            Exit Function
        End If
        outArgSpecs.Add argSpec
    Next i

    mp_TryParseCallMacroArgs = True
End Function

Private Function mp_TryParseMacroArg(ByVal argText As String, ByRef outArgSpec As Object) As Boolean
    mp_TryParseMacroArg = ex_ResultRuntimeAdapter.m_TryParseMacroArg(argText, outArgSpec)
End Function

Private Function mp_ValidateCallMacroArgs( _
    ByVal action As Object, _
    ByVal scopeVarTypes As Object, _
    ByVal allowedTableFields As Object, _
    ByRef outErrorText As String _
) As Boolean
    Dim argSpecs As Collection
    Dim i As Long
    Dim argSpec As Object
    If action Is Nothing Then Exit Function
    If Not action.Exists("Args") Then
        mp_ValidateCallMacroArgs = True
        Exit Function
    End If

    Set argSpecs = action("Args")
    For i = 1 To argSpecs.Count
        Set argSpec = argSpecs(i)
        If Not ex_ResultRuntimeAdapter.m_ValidateMacroArgSpec(argSpec, scopeVarTypes, allowedTableFields, outErrorText) Then Exit Function
    Next i

    mp_ValidateCallMacroArgs = True
End Function

Private Function mp_BuildMacroRuntimeArgs( _
    ByVal action As Object, _
    ByVal currentTableRef As String, _
    ByVal currentRowVar As String, _
    ByVal currentRowRef As obj_ResultRow, _
    ByVal tablesByRef As Object, _
    ByVal runtimeVars As Object _
) As Collection
    Dim result As Collection
    Dim argSpecs As Collection
    Dim i As Long
    Dim argSpec As Object
    Dim argKind As String
    Dim renderedText As String
    Dim argObject As Object
    Dim argValue As Variant

    Set result = New Collection
    If action Is Nothing Then
        Set mp_BuildMacroRuntimeArgs = result
        Exit Function
    End If

    If Not action.Exists("Args") Then
        Set mp_BuildMacroRuntimeArgs = result
        Exit Function
    End If

    Set argSpecs = action("Args")
    For i = 1 To argSpecs.Count
        Set argSpec = argSpecs(i)
        argKind = LCase$(CStr(argSpec("Kind")))
        Select Case argKind
            Case "varref"
                If runtimeVars Is Nothing Or Not runtimeVars.Exists(CStr(argSpec("Name"))) Then
                    Err.Raise vbObjectError + 1601, "ex_PostProcessDsl", "Variable '" & CStr(argSpec("Name")) & "' is not available for callMacro argument."
                End If
                If IsObject(runtimeVars(CStr(argSpec("Name")))) Then
                    Set argObject = runtimeVars(CStr(argSpec("Name")))
                    result.Add argObject
                Else
                    argValue = runtimeVars(CStr(argSpec("Name")))
                    result.Add argValue
                End If
            Case "rowref"
                result.Add ex_ResultRuntimeAdapter.m_ResolveRowReferenceArg(argSpec, tablesByRef)
            Case "cellref"
                result.Add ex_ResultRuntimeAdapter.m_ResolveCellReferenceArg(argSpec, tablesByRef)
            Case "string"
                renderedText = mp_RenderTemplate(CStr(argSpec("Value")), currentTableRef, currentRowVar, currentRowRef, tablesByRef, runtimeVars)
                result.Add renderedText
            Case Else
                Err.Raise vbObjectError + 1598, "ex_PostProcessDsl", "Unsupported callMacro argument kind: " & argKind
        End Select
    Next i

    Set mp_BuildMacroRuntimeArgs = result
End Function

Private Function mp_CreateVarScope() As Object
    Set mp_CreateVarScope = CreateObject("Scripting.Dictionary")
    mp_CreateVarScope.CompareMode = 1
End Function

Private Function mp_CloneVarScope(ByVal sourceScope As Object) As Object
    Dim result As Object
    Dim key As Variant

    Set result = mp_CreateVarScope()
    If sourceScope Is Nothing Then
        Set mp_CloneVarScope = result
        Exit Function
    End If

    For Each key In sourceScope.Keys
        If IsObject(sourceScope(key)) Then
            Set result(CStr(key)) = sourceScope(key)
        Else
            result(CStr(key)) = sourceScope(key)
        End If
    Next key

    Set mp_CloneVarScope = result
End Function

Private Sub mp_SetScopeValue(ByVal targetScope As Object, ByVal variableName As String, ByVal variableValue As String)
    variableName = Trim$(variableName)
    If Len(variableName) = 0 Then Exit Sub
    If targetScope Is Nothing Then Exit Sub

    targetScope(variableName) = variableValue
End Sub

Private Sub mp_SetScopeObject(ByVal targetScope As Object, ByVal variableName As String, ByVal variableObject As Object)
    variableName = Trim$(variableName)
    If Len(variableName) = 0 Then Exit Sub
    If targetScope Is Nothing Then Exit Sub
    Set targetScope(variableName) = variableObject
End Sub

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
    If Not ex_PostProcessParserCore.m_IsIdentifier(outVariableName) Then Exit Function
    If Not ex_PostProcessParserCore.m_IsIdentifier(outMemberName) Then Exit Function

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
    Dim valueObj As Object
    Dim scalarValue As Variant

    If scopeRef Is Nothing Or Not scopeRef.Exists(variableName) Then
        outErrorText = "Unknown variable '" & variableName & "'."
        Exit Function
    End If

    If IsObject(scopeRef(variableName)) Then
        Set valueObj = scopeRef(variableName)
        mp_TryConvertScopeEntryToString = mp_TryConvertScopeValueToString(valueObj, outValue, outErrorText)
        Exit Function
    End If

    scalarValue = scopeRef(variableName)
    If IsNull(scalarValue) Then
        outValue = vbNullString
    ElseIf IsError(scalarValue) Then
        outErrorText = "Variable value contains error and cannot be rendered."
        Exit Function
    Else
        outValue = CStr(scalarValue)
    End If

    mp_TryConvertScopeEntryToString = True
End Function

Private Function mp_TryResolveScopeMemberValue( _
    ByVal scopeRef As Object, _
    ByVal variableName As String, _
    ByVal memberName As String, _
    ByVal tokenText As String, _
    ByRef outValue As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim variableObject As Object

    If scopeRef Is Nothing Or Not scopeRef.Exists(variableName) Then
        outErrorText = "Unknown variable '" & variableName & "' in token '" & tokenText & "'."
        Exit Function
    End If
    If Not IsObject(scopeRef(variableName)) Then
        outErrorText = "Variable '" & variableName & "' does not support member access in token '" & tokenText & "'."
        Exit Function
    End If

    Set variableObject = scopeRef(variableName)
    mp_TryResolveScopeMemberValue = mp_TryResolveVariableMemberValue(variableObject, memberName, tokenText, outValue, outErrorText)
End Function

Private Function mp_ConvertVariantToString(ByVal valueRef As Variant) As String
    If IsNull(valueRef) Then
        mp_ConvertVariantToString = vbNullString
    ElseIf IsError(valueRef) Then
        Err.Raise vbObjectError + 1612, "ex_PostProcessDsl", "callMacro returned error value; expected string-compatible result."
    ElseIf IsObject(valueRef) Then
        Err.Raise vbObjectError + 1613, "ex_PostProcessDsl", "callMacro returned object; expected string-compatible result."
    Else
        mp_ConvertVariantToString = CStr(valueRef)
    End If
End Function

Private Function mp_GetStatementAssignKind(ByVal statement As Object) As String
    Dim assignKind As String

    mp_GetStatementAssignKind = ASSIGN_KIND_CALL_MACRO
    If statement Is Nothing Then Exit Function
    If Not statement.Exists("AssignKind") Then Exit Function

    assignKind = LCase$(Trim$(CStr(statement("AssignKind"))))
    If Len(assignKind) = 0 Then Exit Function
    mp_GetStatementAssignKind = assignKind
End Function

Private Function mp_EvaluateStringExpression( _
    ByVal exprParts As Collection, _
    ByVal currentTableRef As String, _
    ByVal currentRowVar As String, _
    ByVal currentRowRef As obj_ResultRow, _
    ByVal tablesByRef As Object, _
    ByVal runtimeVars As Object _
) As String
    Dim i As Long
    Dim partSpec As Object
    Dim partKind As String
    Dim tokenText As String
    Dim tokenValue As String
    Dim resolveError As String
    Dim resultText As String

    If exprParts Is Nothing Then Exit Function

    For i = 1 To exprParts.Count
        Set partSpec = exprParts(i)
        If partSpec Is Nothing Then
            Err.Raise vbObjectError + 1625, "ex_PostProcessDsl", "String expression contains invalid operand."
        End If
        partKind = LCase$(Trim$(CStr(partSpec("Kind"))))

        Select Case partKind
            Case EXPR_PART_LITERAL
                resultText = resultText & CStr(partSpec("Value"))
            Case EXPR_PART_TOKEN
                tokenText = CStr(partSpec("Value"))
                If Not mp_TryResolveRuntimeValue(tokenText, currentTableRef, currentRowVar, currentRowRef, tablesByRef, runtimeVars, tokenValue, resolveError) Then
                    Err.Raise vbObjectError + 1622, "ex_PostProcessDsl", "Unable to resolve string expression token '" & tokenText & "': " & resolveError
                End If
                resultText = resultText & tokenValue
            Case Else
                Err.Raise vbObjectError + 1623, "ex_PostProcessDsl", "Unsupported string expression operand kind: " & partKind
        End Select
    Next i

    mp_EvaluateStringExpression = resultText
End Function

Private Function mp_InferLetVarType( _
    ByVal statement As Object, _
    ByVal scopeVarTypes As Object _
) As String
    Dim macroName As String
    Dim args As Collection
    Dim firstArg As Object
    Dim varName As String

    mp_InferLetVarType = VAR_TYPE_STRING
    If statement Is Nothing Then Exit Function

    macroName = LCase$(Trim$(CStr(statement("MacroName"))))
    If Right$(macroName, Len(".m_getrelativerow")) <> ".m_getrelativerow" Then Exit Function
    If Not statement.Exists("Args") Then Exit Function

    Set args = statement("Args")
    If args Is Nothing Or args.Count < 1 Then Exit Function

    Set firstArg = args(1)
    If firstArg Is Nothing Then Exit Function
    If LCase$(CStr(firstArg("Kind"))) <> "varref" Then Exit Function

    varName = CStr(firstArg("Name"))
    If scopeVarTypes Is Nothing Or Not scopeVarTypes.Exists(varName) Then Exit Function
    If StrComp(CStr(scopeVarTypes(varName)), VAR_TYPE_ROW, vbTextCompare) <> 0 Then Exit Function

    mp_InferLetVarType = VAR_TYPE_ROW
End Function

Private Function mp_LetExpectsObjectResult(ByVal statement As Object) As Boolean
    Dim macroName As String

    If statement Is Nothing Then Exit Function
    macroName = LCase$(Trim$(CStr(statement("MacroName"))))
    If Right$(macroName, Len(".m_getrelativerow")) = ".m_getrelativerow" Then
        mp_LetExpectsObjectResult = True
    End If
End Function

Private Function mp_TryResolveScopedRowCellValue( _
    ByVal scopeRef As Object, _
    ByVal rowVarName As String, _
    ByVal fieldAlias As String, _
    ByVal tokenText As String, _
    ByRef outValue As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim rowObject As obj_ResultRow

    rowVarName = Trim$(rowVarName)
    fieldAlias = Trim$(fieldAlias)
    If Len(rowVarName) = 0 Or Len(fieldAlias) = 0 Then Exit Function

    If scopeRef Is Nothing Or Not scopeRef.Exists(rowVarName) Then
        outErrorText = "Unknown row variable '" & rowVarName & "' in token '" & tokenText & "'."
        Exit Function
    End If
    If Not IsObject(scopeRef(rowVarName)) Then
        outErrorText = "Variable '" & rowVarName & "' is not a row object in token '" & tokenText & "'."
        Exit Function
    End If
    If Not TypeOf scopeRef(rowVarName) Is obj_ResultRow Then
        outErrorText = "Variable '" & rowVarName & "' must be row object in token '" & tokenText & "'."
        Exit Function
    End If

    Set rowObject = scopeRef(rowVarName)
    If Not rowObject.HasAlias(fieldAlias) Then
        outErrorText = "Unknown field alias '" & fieldAlias & "' for row variable '" & rowVarName & "'."
        Exit Function
    End If

    outValue = CStr(rowObject.Column(fieldAlias))
    mp_TryResolveScopedRowCellValue = True
End Function

Private Function mp_TrySplitAssignmentStatement( _
    ByVal statementText As String, _
    ByRef outVarName As String, _
    ByRef outActionText As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim bodyText As String
    Dim i As Long
    Dim ch As String
    Dim inQuotes As Boolean
    Dim depth As Long
    Dim assignPos As Long

    statementText = Trim$(statementText)
    If Len(statementText) = 0 Then Exit Function
    If Right$(statementText, 1) <> ";" Then
        outErrorText = "Assignment statement must end with ';'."
        Exit Function
    End If

    bodyText = Trim$(Left$(statementText, Len(statementText) - 1))
    If Len(bodyText) = 0 Then Exit Function

    For i = 1 To Len(bodyText)
        ch = Mid$(bodyText, i, 1)
        If ch = """" And Not mp_IsEscapedQuote(bodyText, i) Then
            inQuotes = Not inQuotes
        ElseIf Not inQuotes Then
            If ch = "(" Then
                depth = depth + 1
            ElseIf ch = ")" Then
                If depth > 0 Then depth = depth - 1
            ElseIf ch = "=" And depth = 0 Then
                If i < Len(bodyText) And Mid$(bodyText, i + 1, 1) = "=" Then
                    outErrorText = "Invalid assignment syntax."
                    Exit Function
                End If
                assignPos = i
                Exit For
            End If
        End If
    Next i

    If assignPos <= 1 Then Exit Function

    outVarName = Trim$(Left$(bodyText, assignPos - 1))
    outActionText = Trim$(Mid$(bodyText, assignPos + 1))
    If Len(outVarName) = 0 Or Len(outActionText) = 0 Then
        outErrorText = "Invalid assignment syntax."
        Exit Function
    End If

    outActionText = outActionText & ";"
    mp_TrySplitAssignmentStatement = True
End Function

Private Function mp_SplitArgs(ByVal argsText As String, ByRef outParts As Collection, ByRef outErrorText As String) As Boolean
    Dim i As Long
    Dim ch As String
    Dim partText As String
    Dim inQuotes As Boolean

    Set outParts = New Collection
    partText = vbNullString

    For i = 1 To Len(argsText)
        ch = Mid$(argsText, i, 1)
        If ch = """" And Not mp_IsEscapedQuote(argsText, i) Then
            inQuotes = Not inQuotes
            partText = partText & ch
        ElseIf ch = "," And Not inQuotes Then
            outParts.Add Trim$(partText)
            partText = vbNullString
        Else
            partText = partText & ch
        End If
    Next i

    If inQuotes Then
        outErrorText = "Unterminated quoted string in callMacro arguments."
        Exit Function
    End If

    outParts.Add Trim$(partText)
    mp_SplitArgs = True
End Function

Private Function mp_IsEscapedQuote(ByVal textValue As String, ByVal pos As Long) As Boolean
    If pos <= 1 Then Exit Function
    mp_IsEscapedQuote = (Mid$(textValue, pos, 1) = """" And Mid$(textValue, pos - 1, 1) = "\")
End Function

Private Function mp_TrimDslWhitespace(ByVal textValue As String) As String
    textValue = CStr(textValue)
    textValue = Replace(textValue, vbCr, " ")
    textValue = Replace(textValue, vbLf, " ")
    textValue = Replace(textValue, vbTab, " ")

    Do While InStr(1, textValue, "  ", vbBinaryCompare) > 0
        textValue = Replace(textValue, "  ", " ")
    Loop

    mp_TrimDslWhitespace = Trim$(textValue)
End Function

Private Function mp_GetLastUsedColumn(ByVal ws As Worksheet) As Long
    On Error GoTo ExitFn
    mp_GetLastUsedColumn = ws.Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
ExitFn:
End Function
