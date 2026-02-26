Attribute VB_Name = "ex_PostProcessDsl"
Option Explicit

Private Const SCRIPT_KEY As String = "PostProcess.Script"
Private Const ACTION_CALL_MACRO As String = "callmacro"
Private Const ACTION_LET As String = "let"

Private Const LOOP_TARGET_TABLE_ROWS As String = "tablerows"
Private Const LOOP_TARGET_ROW_COLUMNS As String = "rowcolumns"

Private Const VAR_TYPE_ROW As String = "row"
Private Const VAR_TYPE_COLUMN As String = "column"
Private Const VAR_TYPE_STRING As String = "string"

Public Function m_ValidateScriptAgainstConfig( _
    ByVal cfg As Object, _
    ByVal allowedTableFields As Object, _
    ByRef outErrorText As String, _
    Optional ByVal scriptConfigKey As String = SCRIPT_KEY _
) As Boolean
    Dim scriptText As String
    Dim blocks As Collection
    Dim stepName As String

    On Error GoTo EH

    stepName = "load-script"
    If Not ex_PostProcessScriptSource.m_TryGetScriptText(cfg, scriptConfigKey, scriptText, outErrorText) Then Exit Function
    If Len(scriptText) = 0 Then
        m_ValidateScriptAgainstConfig = True
        Exit Function
    End If

    stepName = "parse-script"
    If Not mp_ParseScript(scriptText, blocks, outErrorText) Then Exit Function
    stepName = "validate-blocks"
    If Not mp_ValidateBlocks(blocks, allowedTableFields, outErrorText) Then Exit Function

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
    Dim scriptLoadError As String
    Dim blocks As Collection
    Dim parseError As String
    Dim ctxTablesByRef As Object
    Dim ctxFields As Object
    Dim postProcessFooterLines As Collection
    Dim usedCols As Long

    If ws Is Nothing Then Exit Sub
    If cfg Is Nothing Then Exit Sub
    If resultTables Is Nothing Then Exit Sub

    If Not ex_PostProcessScriptSource.m_TryGetScriptText(cfg, scriptConfigKey, scriptText, scriptLoadError) Then
        Err.Raise vbObjectError + 1592, "ex_PostProcessDsl", scriptLoadError
    End If
    If Len(scriptText) = 0 Then Exit Sub

    If Not mp_ParseScript(scriptText, blocks, parseError) Then
        Err.Raise vbObjectError + 1590, "ex_PostProcessDsl", "PostProcess script parse failed: " & parseError
    End If

    ex_ResultRuntimeAdapter.m_BuildRuntimeContext resultTables, ctxTablesByRef, ctxFields

    If Not mp_ValidateBlocks(blocks, ctxFields, parseError) Then
        Err.Raise vbObjectError + 1591, "ex_PostProcessDsl", "PostProcess script validation failed: " & parseError
    End If

    Set postProcessFooterLines = New Collection
    usedCols = mp_GetLastUsedColumn(ws)
    If usedCols <= 0 Then usedCols = 1

    mp_ExecuteBlocks ws, blocks, ctxTablesByRef, postProcessFooterLines, usedCols
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
    mp_ValidateBlocks = mp_ValidateStatements(blocks, allowedTableFields, vbNullString, vbNullString, rootScopeVarTypes, outErrorText)
End Function

Private Sub mp_ExecuteBlocks( _
    ByVal ws As Worksheet, _
    ByVal blocks As Collection, _
    ByVal tablesByRef As Object, _
    ByVal postProcessFooterLines As Collection, _
    ByVal usedCols As Long _
)
    Dim rootRuntimeVars As Object

    Set rootRuntimeVars = mp_CreateVarScope()
    mp_ExecuteStatements ws, blocks, tablesByRef, postProcessFooterLines, usedCols, vbNullString, vbNullString, Nothing, rootRuntimeVars
End Sub

Private Sub mp_ExecuteStatements( _
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

    If statements Is Nothing Then Exit Sub
    If runtimeVars Is Nothing Then Set runtimeVars = mp_CreateVarScope()

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
                On Error GoTo CallMacroErr
                Set macroArgs = mp_BuildMacroRuntimeArgs(statement, currentTableRef, currentRowVar, currentRowRef, tablesByRef, runtimeVars)
                macroResult = ex_PostProcessActionInvoker.m_RunMacroWithArgsReturn(CStr(statement("MacroName")), macroArgs)
                mp_SetScopeValue runtimeVars, CStr(statement("VarName")), mp_ConvertVariantToString(macroResult)
                On Error GoTo 0

            Case "if"
                If mp_EvaluateCondition(CStr(statement("Condition")), currentTableRef, currentRowVar, currentRowRef, tablesByRef, runtimeVars) Then
                    Set childRuntimeVars = mp_CloneVarScope(runtimeVars)
                    mp_ExecuteStatements ws, statement("Body"), tablesByRef, postProcessFooterLines, usedCols, currentTableRef, currentRowVar, currentRowRef, childRuntimeVars
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
                                mp_ExecuteStatements ws, statement("Body"), tablesByRef, postProcessFooterLines, usedCols, CStr(statement("TableRef")), loopVarName, rowRef, childRuntimeVars
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
                            mp_ExecuteStatements ws, statement("Body"), tablesByRef, postProcessFooterLines, usedCols, currentTableRef, sourceRowVarName, sourceRowRef, childRuntimeVars
                        Next rowIdx

                    Case Else
                        Err.Raise vbObjectError + 1611, "ex_PostProcessDsl", "Unsupported for-loop target: " & CStr(statement("LoopTarget"))
                End Select

            Case Else
                Err.Raise vbObjectError + 1593, "ex_PostProcessDsl", "Unsupported statement type: " & statementType
        End Select
    Next i
    Exit Sub

CallMacroErr:
    Err.Raise vbObjectError + 1597, "ex_PostProcessDsl", "callMacro failed for '" & CStr(statement("MacroName")) & "': " & Err.Description
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
        Case "callmacro"
            mp_ParseStatement = mp_TryParseCallMacroStatement(sourceText, pos, lineNo, outStatement, outErrorText)
        Case "let"
            mp_ParseStatement = mp_TryParseLetStatement(sourceText, pos, lineNo, outStatement, outErrorText)
        Case Else
            outErrorText = "Unsupported statement '" & keywordText & "' at line " & CStr(lineNo)
    End Select
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
            If Len(outErrorText) = 0 Then outErrorText = "Expected '(item in Source.Sheet[Table].rows)' after for at line " & CStr(stmtLine)
            Exit Function
        End If
        If Not mp_TryParseForHeaderText(forHeaderText, loopVarName, targetText) Then
            outErrorText = "Expected 'for (item in Source.Sheet[Table].rows)' at line " & CStr(stmtLine)
            Exit Function
        End If
    Else
        If Not mp_TryParseForHeaderFromStream(sourceText, pos, lineNo, loopVarName, targetText) Then
            outErrorText = "Expected 'for item in Source.Sheet[Table].rows' at line " & CStr(stmtLine)
            Exit Function
        End If
    End If

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
    Dim inKeyword As String

    If Not mp_ReadIdentifier(sourceText, pos, lineNo, outRowVarName) Then Exit Function
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

    Set outStatement = CreateObject("Scripting.Dictionary")
    outStatement.CompareMode = 1
    outStatement("Type") = "if"
    outStatement("Condition") = conditionText
    outStatement.Add "Body", bodyStatements
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
    Dim actionStatement As Object
    Dim stmtLine As Long

    stmtLine = lineNo
    If Not mp_ReadIdentifier(sourceText, pos, lineNo, keywordText) Then Exit Function
    If StrComp(keywordText, "let", vbTextCompare) <> 0 Then Exit Function

    mp_SkipWhitespace sourceText, pos, lineNo
    If Not mp_ReadIdentifier(sourceText, pos, lineNo, varName) Then
        outErrorText = "Expected variable name after let at line " & CStr(stmtLine)
        Exit Function
    End If
    If Not ex_PostProcessParserCore.m_IsIdentifier(varName) Then
        outErrorText = "Invalid variable name '" & varName & "' at line " & CStr(stmtLine)
        Exit Function
    End If

    mp_SkipWhitespace sourceText, pos, lineNo
    If pos > Len(sourceText) Or Mid$(sourceText, pos, 1) <> "=" Then
        outErrorText = "Expected '=' in let statement at line " & CStr(stmtLine)
        Exit Function
    End If
    pos = pos + 1

    If Not mp_ReadStatementToSemicolon(sourceText, pos, lineNo, rhsStatementText, outErrorText) Then Exit Function
    If Not mp_TryParseAction(rhsStatementText, actionStatement, outErrorText) Then
        outErrorText = outErrorText & " at line " & CStr(stmtLine)
        Exit Function
    End If
    If LCase$(CStr(actionStatement("Type"))) <> ACTION_CALL_MACRO Then
        outErrorText = "let supports only callMacro(...) at line " & CStr(stmtLine)
        Exit Function
    End If

    Set outStatement = actionStatement
    outStatement("Type") = ACTION_LET
    outStatement("VarName") = varName
    outStatement("Line") = CLng(stmtLine)
    mp_TryParseLetStatement = True
End Function

Private Function mp_ValidateStatements( _
    ByVal statements As Collection, _
    ByVal allowedTableFields As Object, _
    ByVal currentTableRef As String, _
    ByVal currentRowVar As String, _
    ByVal scopeVarTypes As Object, _
    ByRef outErrorText As String _
) As Boolean
    Dim i As Long
    Dim statement As Object
    Dim statementType As String
    Dim loopTarget As String
    Dim tableRef As String
    Dim loopVarName As String
    Dim sourceRowVarName As String
    Dim childScopeVarTypes As Object

    If statements Is Nothing Then
        mp_ValidateStatements = True
        Exit Function
    End If

    If scopeVarTypes Is Nothing Then Set scopeVarTypes = mp_CreateVarScope()

    For i = 1 To statements.Count
        Set statement = statements(i)
        statementType = LCase$(CStr(statement("Type")))

        Select Case statementType
            Case ACTION_CALL_MACRO
                If Not mp_ValidateCallMacroArgs(statement, scopeVarTypes, allowedTableFields, outErrorText) Then Exit Function

            Case ACTION_LET
                If Not mp_ValidateCallMacroArgs(statement, scopeVarTypes, allowedTableFields, outErrorText) Then Exit Function
                mp_SetScopeValue scopeVarTypes, CStr(statement("VarName")), VAR_TYPE_STRING

            Case "if"
                If Not mp_ValidateConditionText(CStr(statement("Condition")), currentTableRef, currentRowVar, scopeVarTypes, allowedTableFields, outErrorText) Then Exit Function
                Set childScopeVarTypes = mp_CloneVarScope(scopeVarTypes)
                If Not mp_ValidateStatements(statement("Body"), allowedTableFields, currentTableRef, currentRowVar, childScopeVarTypes, outErrorText) Then Exit Function

            Case "for"
                loopTarget = LCase$(CStr(statement("LoopTarget")))
                loopVarName = CStr(statement("LoopVar"))
                Select Case loopTarget
                    Case LOOP_TARGET_TABLE_ROWS
                        tableRef = CStr(statement("TableRef"))
                        If allowedTableFields Is Nothing Or Not allowedTableFields.Exists(tableRef) Then
                            outErrorText = "Unknown table reference in script: '" & tableRef & "'."
                            Exit Function
                        End If
                        Set childScopeVarTypes = mp_CloneVarScope(scopeVarTypes)
                        mp_SetScopeValue childScopeVarTypes, loopVarName, VAR_TYPE_ROW
                        If Not mp_ValidateStatements(statement("Body"), allowedTableFields, tableRef, loopVarName, childScopeVarTypes, outErrorText) Then Exit Function

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
                        If Not mp_ValidateStatements(statement("Body"), allowedTableFields, currentTableRef, sourceRowVarName, childScopeVarTypes, outErrorText) Then Exit Function

                    Case Else
                        outErrorText = "Unsupported for-loop target '" & loopTarget & "'."
                        Exit Function
                End Select

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
    Dim tokenText As String
    Dim resolvedTableRef As String
    Dim resolvedMapKey As String

    If Not mp_TrySplitConditionExpression(conditionText, condParts, condOps, outErrorText) Then Exit Function

    For i = 1 To condParts.Count
        If Not mp_TryExtractConditionField(CStr(condParts(i)), tokenText) Then
            outErrorText = "Unsupported condition token: '" & Trim$(CStr(condParts(i))) & "'."
            Exit Function
        End If
        If Not mp_TryResolveConditionTokenForValidation(tokenText, currentTableRef, currentRowVar, scopeVarTypes, allowedTableFields, resolvedTableRef, resolvedMapKey, outErrorText) Then
            Exit Function
        End If
    Next i

    mp_ValidateConditionText = True
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
    Dim expectedValue As String
    Dim actualValue As String
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
        If Not mp_ParseConditionPart(CStr(condParts(i)), refToken, opText, expectedValue) Then
            Err.Raise vbObjectError + 1594, "ex_PostProcessDsl", "Invalid condition: " & Trim$(CStr(condParts(i)))
        End If
        If Not mp_TryResolveRuntimeValue(refToken, currentTableRef, currentRowVar, currentRowRef, tablesByRef, runtimeVars, actualValue, resolveError) Then
            Err.Raise vbObjectError + 1595, "ex_PostProcessDsl", resolveError
        End If

        Select Case opText
            Case "=="
                partResult = (StrComp(actualValue, expectedValue, vbTextCompare) = 0)
            Case "!="
                partResult = (StrComp(actualValue, expectedValue, vbTextCompare) <> 0)
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

Private Function mp_ParseConditionPart(ByVal rawPart As String, ByRef outFieldName As String, ByRef outOp As String, ByRef outValue As String) As Boolean
    Dim part As String
    Dim opPos As Long
    Dim opLen As Long
    Dim rhs As String

    part = Trim$(rawPart)
    opPos = InStr(1, part, "==", vbBinaryCompare)
    If opPos > 0 Then
        outOp = "=="
        opLen = 2
    Else
        opPos = InStr(1, part, "!=", vbBinaryCompare)
        If opPos > 0 Then
            outOp = "!="
            opLen = 2
        End If
    End If
    If opPos <= 1 Then Exit Function

    outFieldName = Trim$(Left$(part, opPos - 1))
    rhs = Trim$(Mid$(part, opPos + opLen))
    If Len(outFieldName) = 0 Then Exit Function
    If Not mp_TryParseQuotedString(rhs, outValue) Then Exit Function

    mp_ParseConditionPart = True
End Function

Private Function mp_TryExtractConditionField(ByVal rawPart As String, ByRef outFieldName As String) As Boolean
    Dim opText As String
    Dim valueText As String
    mp_TryExtractConditionField = mp_ParseConditionPart(rawPart, outFieldName, opText, valueText)
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
            If i < Len(conditionText) And Mid$(conditionText, i, 2) = "&&" Then
                If Not mp_TryPushConditionPart(partText, "and", outParts, outOps, outErrorText) Then Exit Function
                partText = vbNullString
                i = i + 2
                GoTo ContinueLoop
            End If

            If i < Len(conditionText) And Mid$(conditionText, i, 2) = "||" Then
                If Not mp_TryPushConditionPart(partText, "or", outParts, outOps, outErrorText) Then Exit Function
                partText = vbNullString
                i = i + 2
                GoTo ContinueLoop
            End If

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
    Dim tokenValue As String
    Dim resolveError As String

    result = templateText
    openPos = InStr(1, result, "{", vbBinaryCompare)
    Do While openPos > 0
        closePos = InStr(openPos + 1, result, "}", vbBinaryCompare)
        If closePos <= openPos Then Exit Do

        tokenText = Mid$(result, openPos + 1, closePos - openPos - 1)
        If Not mp_TryResolveRuntimeValue(tokenText, currentTableRef, currentRowVar, currentRowRef, tablesByRef, runtimeVars, tokenValue, resolveError) Then
            Err.Raise vbObjectError + 1595, "ex_PostProcessDsl", resolveError
        End If

        result = Left$(result, openPos - 1) & tokenValue & Mid$(result, closePos + 1)
        openPos = InStr(openPos + Len(tokenValue), result, "{", vbBinaryCompare)
    Loop

    mp_RenderTemplate = result
End Function

Private Function mp_NormalizeScript(ByVal scriptText As String) As String
    Dim lines As Variant
    Dim i As Long
    Dim rawLine As String
    Dim cleaned As String
    Dim normalized As String

    scriptText = Replace(scriptText, vbCrLf, vbLf)
    scriptText = Replace(scriptText, vbCr, vbLf)
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

Private Function mp_GetLastUsedColumn(ByVal ws As Worksheet) As Long
    On Error GoTo ExitFn
    mp_GetLastUsedColumn = ws.Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
ExitFn:
End Function
