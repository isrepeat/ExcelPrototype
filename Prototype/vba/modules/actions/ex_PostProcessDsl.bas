Attribute VB_Name = "ex_PostProcessDsl"
Option Explicit

Private Const SCRIPT_KEY As String = "PostProcess.Script"
Private Const ACTION_CALL_MACRO As String = "callmacro"
Private Const DSL_TYPE_SHEET_REF As String = "sheetref"
Private Const DSL_TYPE_ROW As String = "row"
Private g_LastScriptLoadError As String
Private g_DslMembersByType As Object

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
    scriptText = mp_GetScriptText(cfg, scriptConfigKey)
    If Len(g_LastScriptLoadError) > 0 Then
        outErrorText = g_LastScriptLoadError
        Exit Function
    End If
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
    Dim blocks As Collection
    Dim parseError As String
    Dim ctxTablesByRef As Object
    Dim ctxFields As Object
    Dim i As Long
    Dim tableObj As pp_ResultTable
    Dim fieldMap As Object
    Dim aliasKey As Variant
    Dim mapKey As String
    Dim tableFields As Object
    Dim footerLines As Collection
    Dim usedCols As Long

    If ws Is Nothing Then Exit Sub
    If cfg Is Nothing Then Exit Sub
    If resultTables Is Nothing Then Exit Sub

    scriptText = mp_GetScriptText(cfg, scriptConfigKey)
    If Len(scriptText) = 0 Then Exit Sub

    If Not mp_ParseScript(scriptText, blocks, parseError) Then
        Err.Raise vbObjectError + 1590, "ex_PostProcessDsl", "PostProcess script parse failed: " & parseError
    End If

    Set ctxTablesByRef = CreateObject("Scripting.Dictionary")
    ctxTablesByRef.CompareMode = 1
    Set ctxFields = CreateObject("Scripting.Dictionary")
    ctxFields.CompareMode = 1

    For i = 1 To resultTables.Count
        Set tableObj = resultTables(i)
        If tableObj Is Nothing Then GoTo ContinueTable
        If Not ctxTablesByRef.Exists(tableObj.TableRef) Then
            ctxTablesByRef.Add tableObj.TableRef, tableObj
        End If

        Set tableFields = CreateObject("Scripting.Dictionary")
        tableFields.CompareMode = 1
        Set fieldMap = tableObj.FieldMapByAlias
        For Each aliasKey In fieldMap.Keys
            mapKey = CStr(fieldMap(aliasKey))
            tableFields(mapKey) = True
        Next aliasKey
        If ctxFields.Exists(tableObj.TableRef) Then
            ctxFields.Remove tableObj.TableRef
        End If
        ctxFields.Add tableObj.TableRef, tableFields
ContinueTable:
    Next i

    If Not mp_ValidateBlocks(blocks, ctxFields, parseError) Then
        Err.Raise vbObjectError + 1591, "ex_PostProcessDsl", "PostProcess script validation failed: " & parseError
    End If

    Set footerLines = New Collection
    usedCols = mp_GetLastUsedColumn(ws)
    If usedCols <= 0 Then usedCols = 1

    mp_ExecuteBlocks ws, blocks, ctxTablesByRef, footerLines, usedCols
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
    If blocks Is Nothing Then
        mp_ValidateBlocks = True
        Exit Function
    End If

    mp_ValidateBlocks = mp_ValidateStatements(blocks, allowedTableFields, vbNullString, vbNullString, outErrorText)
End Function

Private Sub mp_ExecuteBlocks( _
    ByVal ws As Worksheet, _
    ByVal blocks As Collection, _
    ByVal tablesByRef As Object, _
    ByVal footerLines As Collection, _
    ByVal usedCols As Long _
)
    mp_ExecuteStatements ws, blocks, tablesByRef, footerLines, usedCols, vbNullString, vbNullString, Nothing
End Sub

Private Sub mp_ExecuteStatements( _
    ByVal ws As Worksheet, _
    ByVal statements As Collection, _
    ByVal tablesByRef As Object, _
    ByVal footerLines As Collection, _
    ByVal usedCols As Long, _
    ByVal currentTableRef As String, _
    ByVal currentRowVar As String, _
    ByVal currentRowRef As pp_ResultRow _
)
    Dim i As Long
    Dim statement As Object
    Dim statementType As String
    Dim macroArgs As Collection
    Dim tableObj As pp_ResultTable
    Dim rowsList As Collection
    Dim rowRef As pp_ResultRow
    Dim rowIdx As Long

    If statements Is Nothing Then Exit Sub

    For i = 1 To statements.Count
        Set statement = statements(i)
        statementType = LCase$(CStr(statement("Type")))

        Select Case statementType
            Case ACTION_CALL_MACRO
                On Error GoTo CallMacroErr
                Set macroArgs = mp_BuildMacroRuntimeArgs(statement, currentTableRef, currentRowVar, currentRowRef, tablesByRef)
                mp_RunMacroWithArgs CStr(statement("MacroName")), macroArgs
                On Error GoTo 0

            Case "if"
                If mp_EvaluateCondition(CStr(statement("Condition")), currentTableRef, currentRowVar, currentRowRef, tablesByRef) Then
                    mp_ExecuteStatements ws, statement("Body"), tablesByRef, footerLines, usedCols, currentTableRef, currentRowVar, currentRowRef
                End If

            Case "for"
                If tablesByRef.Exists(CStr(statement("TableRef"))) Then
                    Set tableObj = tablesByRef(CStr(statement("TableRef")))
                    Set rowsList = tableObj.Rows
                    For rowIdx = 1 To rowsList.Count
                        Set rowRef = rowsList(rowIdx)
                        mp_ExecuteStatements ws, statement("Body"), tablesByRef, footerLines, usedCols, CStr(statement("TableRef")), CStr(statement("RowVar")), rowRef
                    Next rowIdx
                End If

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
    Dim rowVarName As String
    Dim targetText As String
    Dim tableRef As String
    Dim bodyStatements As Collection
    Dim stmtLine As Long
    Dim memberName As String

    stmtLine = lineNo
    If Not mp_ReadIdentifier(sourceText, pos, lineNo, keywordText) Then Exit Function
    If StrComp(keywordText, "for", vbTextCompare) <> 0 Then Exit Function

    mp_SkipWhitespace sourceText, pos, lineNo
    If pos <= Len(sourceText) And Mid$(sourceText, pos, 1) = "(" Then
        If Not mp_ReadBalanced(sourceText, pos, lineNo, "(", ")", forHeaderText, outErrorText) Then
            If Len(outErrorText) = 0 Then outErrorText = "Expected '(row in Source.Sheet[Table].rows)' after for at line " & CStr(stmtLine)
            Exit Function
        End If
        If Not mp_TryParseForHeaderText(forHeaderText, rowVarName, targetText) Then
            outErrorText = "Expected 'for (row in Source.Sheet[Table].rows)' at line " & CStr(stmtLine)
            Exit Function
        End If
    Else
        If Not mp_TryParseForHeaderFromStream(sourceText, pos, lineNo, rowVarName, targetText) Then
            outErrorText = "Expected 'for row in Source.Sheet[Table].rows' at line " & CStr(stmtLine)
            Exit Function
        End If
    End If

    If LCase$(Right$(targetText, 5)) <> ".rows" Then
        outErrorText = "Expected '.rows' in for target at line " & CStr(stmtLine)
        Exit Function
    End If
    tableRef = Trim$(Left$(targetText, Len(targetText) - 5))
    If Len(tableRef) = 0 Or Not mp_IsSheetRef(tableRef) Then
        outErrorText = "Invalid table reference in for-statement at line " & CStr(stmtLine)
        Exit Function
    End If
    memberName = "rows"
    If Not mp_IsMemberAllowed(DSL_TYPE_SHEET_REF, memberName) Then
        outErrorText = "Member '.rows' is not allowed for table reference at line " & CStr(stmtLine)
        Exit Function
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
    outStatement("RowVar") = rowVarName
    outStatement("TableRef") = tableRef
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

Private Function mp_ValidateStatements( _
    ByVal statements As Collection, _
    ByVal allowedTableFields As Object, _
    ByVal currentTableRef As String, _
    ByVal currentRowVar As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim i As Long
    Dim statement As Object
    Dim statementType As String
    Dim tableRef As String
    Dim rowVarName As String

    If statements Is Nothing Then
        mp_ValidateStatements = True
        Exit Function
    End If

    For i = 1 To statements.Count
        Set statement = statements(i)
        statementType = LCase$(CStr(statement("Type")))

        Select Case statementType
            Case ACTION_CALL_MACRO
                If Not mp_ValidateCallMacroArgs(statement, currentRowVar, allowedTableFields, outErrorText) Then Exit Function

            Case "if"
                If Not mp_ValidateConditionText(CStr(statement("Condition")), currentTableRef, currentRowVar, allowedTableFields, outErrorText) Then Exit Function
                If Not mp_ValidateStatements(statement("Body"), allowedTableFields, currentTableRef, currentRowVar, outErrorText) Then Exit Function

            Case "for"
                tableRef = CStr(statement("TableRef"))
                rowVarName = CStr(statement("RowVar"))
                If allowedTableFields Is Nothing Or Not allowedTableFields.Exists(tableRef) Then
                    outErrorText = "Unknown table reference in script: '" & tableRef & "'."
                    Exit Function
                End If
                If Not mp_ValidateStatements(statement("Body"), allowedTableFields, tableRef, rowVarName, outErrorText) Then Exit Function

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
    ByVal allowedTableFields As Object, _
    ByRef outErrorText As String _
) As Boolean
    Dim condParts As Variant
    Dim part As Variant
    Dim tokenText As String
    Dim resolvedTableRef As String
    Dim resolvedMapKey As String

    condParts = Split(conditionText, "&&")
    For Each part In condParts
        If Not mp_TryExtractConditionField(CStr(part), tokenText) Then
            outErrorText = "Unsupported condition token: '" & Trim$(CStr(part)) & "'."
            Exit Function
        End If
        If Not mp_TryResolveConditionTokenForValidation(tokenText, currentTableRef, currentRowVar, allowedTableFields, resolvedTableRef, resolvedMapKey, outErrorText) Then
            Exit Function
        End If
    Next part

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
    ByVal currentRowRef As pp_ResultRow, _
    ByVal tablesByRef As Object _
) As Boolean
    Dim condParts As Variant
    Dim part As Variant
    Dim refToken As String
    Dim opText As String
    Dim expectedValue As String
    Dim actualValue As String
    Dim resolveError As String

    condParts = Split(conditionText, "&&")
    For Each part In condParts
        If Not mp_ParseConditionPart(CStr(part), refToken, opText, expectedValue) Then
            Err.Raise vbObjectError + 1594, "ex_PostProcessDsl", "Invalid condition: " & Trim$(CStr(part))
        End If
        If Not mp_TryResolveRuntimeValue(refToken, currentTableRef, currentRowVar, currentRowRef, tablesByRef, actualValue, resolveError) Then
            Err.Raise vbObjectError + 1595, "ex_PostProcessDsl", resolveError
        End If

        Select Case opText
            Case "=="
                If StrComp(actualValue, expectedValue, vbTextCompare) <> 0 Then Exit Function
            Case "!="
                If StrComp(actualValue, expectedValue, vbTextCompare) = 0 Then Exit Function
            Case Else
                Err.Raise vbObjectError + 1596, "ex_PostProcessDsl", "Unsupported operator in condition: " & opText
        End Select
    Next part

    mp_EvaluateCondition = True
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

Private Function mp_TryParseIfHeader(ByVal lineText As String, ByRef outConditionText As String) As Boolean
    Dim openPos As Long
    Dim closePos As Long

    If Left$(LCase$(lineText), 2) <> "if" Then Exit Function
    If Right$(lineText, 1) <> "{" Then Exit Function

    openPos = InStr(1, lineText, "(", vbBinaryCompare)
    closePos = InStrRev(lineText, ")", -1, vbBinaryCompare)
    If openPos <= 0 Or closePos <= openPos Then Exit Function

    outConditionText = Trim$(Mid$(lineText, openPos + 1, closePos - openPos - 1))
    If Len(outConditionText) = 0 Then Exit Function
    mp_TryParseIfHeader = True
End Function

Private Function mp_TryParseForHeader(ByVal lineText As String, ByRef outRowVar As String, ByRef outTableRef As String) As Boolean
    Dim bodyText As String
    Dim lowerBody As String
    Dim inPos As Long
    Dim targetText As String
    Dim rowsPos As Long
    Dim memberName As String

    If Right$(lineText, 1) <> "{" Then Exit Function

    bodyText = Trim$(Left$(lineText, Len(lineText) - 1))
    lowerBody = LCase$(bodyText)
    If Left$(lowerBody, 4) <> "for " Then Exit Function

    bodyText = Trim$(Mid$(bodyText, 5))
    lowerBody = LCase$(bodyText)
    inPos = InStr(1, lowerBody, " in ", vbBinaryCompare)
    If inPos <= 1 Then Exit Function

    outRowVar = Trim$(Left$(bodyText, inPos - 1))
    If Len(outRowVar) = 0 Then Exit Function

    targetText = Trim$(Mid$(bodyText, inPos + 4))
    If Len(targetText) = 0 Then Exit Function

    rowsPos = InStrRev(targetText, ".rows", -1, vbTextCompare)
    If rowsPos <= 0 Or rowsPos + Len(".rows") - 1 <> Len(targetText) Then Exit Function

    outTableRef = Trim$(Left$(targetText, rowsPos - 1))
    If Len(outTableRef) = 0 Then Exit Function
    If Not mp_IsSheetRef(outTableRef) Then Exit Function
    memberName = "rows"
    If Not mp_IsMemberAllowed(DSL_TYPE_SHEET_REF, memberName) Then Exit Function
    mp_TryParseForHeader = True
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
    ByVal currentRowRef As pp_ResultRow, _
    ByVal tablesByRef As Object _
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
        If Not mp_TryResolveRuntimeValue(tokenText, currentTableRef, currentRowVar, currentRowRef, tablesByRef, tokenValue, resolveError) Then
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

Private Function mp_GetScriptText(ByVal cfg As Object, Optional ByVal scriptConfigKey As String = SCRIPT_KEY) As String
    Dim scriptFromProfile As String
    g_LastScriptLoadError = vbNullString
    scriptFromProfile = mp_GetScriptTextFromActiveProfile()
    If Len(scriptFromProfile) > 0 Then
        mp_GetScriptText = scriptFromProfile
        Exit Function
    End If

    On Error GoTo ExitFn
    If cfg Is Nothing Then Exit Function
    scriptConfigKey = Trim$(scriptConfigKey)
    If Len(scriptConfigKey) = 0 Then scriptConfigKey = SCRIPT_KEY
    If Not cfg.Exists(scriptConfigKey) Then Exit Function
    mp_GetScriptText = Trim$(CStr(cfg(scriptConfigKey)))
    If Len(mp_GetScriptText) = 0 Then Exit Function
    mp_GetScriptText = Replace(mp_GetScriptText, "\n", vbLf)
ExitFn:
End Function

Private Function mp_GetScriptTextFromActiveProfile() As String
    Dim modeName As String
    Dim profileName As String
    Dim filePath As String
    Dim doc As Object
    Dim profileNode As Object
    Dim scriptNode As Object
    Dim stepName As String

    On Error GoTo ExitFn

    stepName = "read-active-mode-profile"
    modeName = Trim$(ex_ConfigProfilesManager.m_GetActiveModeName(ws_Dev))
    profileName = Trim$(ex_ConfigProfilesManager.m_GetActiveProfileName(ws_Dev))
    If Len(modeName) = 0 Or Len(profileName) = 0 Then Exit Function

    stepName = "resolve-profiles-path"
    filePath = ex_ProfilesStore.m_GetProfilesFilePath(modeName, ThisWorkbook)
    If Len(Trim$(filePath)) = 0 Then Exit Function

    stepName = "load-profiles-dom"
    Set doc = ex_ProfilesStore.m_LoadProfilesDom(filePath)
    If doc Is Nothing Then Exit Function
    stepName = "find-profile-node"
    Set profileNode = ex_ProfilesStore.m_GetProfileNode(doc, profileName, False)
    If profileNode Is Nothing Then Exit Function

    stepName = "read-postprocess-node"
    Set scriptNode = profileNode.selectSingleNode("p:postProcessScript")
    If scriptNode Is Nothing Then Exit Function

    mp_GetScriptTextFromActiveProfile = Trim$(CStr(scriptNode.Text))
    If Len(mp_GetScriptTextFromActiveProfile) = 0 Then Exit Function
    mp_GetScriptTextFromActiveProfile = Replace(mp_GetScriptTextFromActiveProfile, "\n", vbLf)
ExitFn:
    If Err.Number <> 0 Then
        g_LastScriptLoadError = "PostProcess script load failed at step '" & stepName & "' [mode=" & modeName & "] [profile=" & profileName & "] [file=" & filePath & "]: [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description
        mp_GetScriptTextFromActiveProfile = vbNullString
    End If
End Function

Private Function mp_TryParseMapKey(ByVal mapKey As String, ByRef outTableAlias As String, ByRef outFieldAlias As String, Optional ByRef outSourceAlias As String) As Boolean
    Dim sheetStart As Long
    Dim sheetEnd As Long
    Dim mapStart As Long
    Dim mapEnd As Long

    sheetStart = InStr(1, mapKey, ".Sheet[", vbTextCompare)
    mapStart = InStr(1, mapKey, "].Map[", vbTextCompare)
    If sheetStart <= 0 Or mapStart <= sheetStart Then Exit Function

    outSourceAlias = Left$(mapKey, sheetStart - 1)
    outSourceAlias = Trim$(outSourceAlias)
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

Private Function mp_TryResolveConditionTokenForValidation( _
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
    Dim sourceAlias As String
    Dim tableAlias As String
    Dim tableRef As String
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

    If mp_TryParseRowColumnRef(tokenText, rowVarName, fieldAlias) Then
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
        mp_TryResolveConditionTokenForValidation = True
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
        mp_TryResolveConditionTokenForValidation = True
        Exit Function
    End If

    If mp_TryParseFullRowColumnRef(tokenText, sourceAlias, tableAlias, rowIndex, fieldAlias) Then
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
        mp_TryResolveConditionTokenForValidation = True
        Exit Function
    End If

    outErrorText = "Unsupported field reference '" & tokenText & "'. Use <rowVar>.column[FieldAlias] or Source.Sheet[TableAlias].row[N].column[FieldAlias]."
End Function

Private Function mp_TryResolveRuntimeValue( _
    ByVal tokenText As String, _
    ByVal currentTableRef As String, _
    ByVal currentRowVar As String, _
    ByVal currentRowRef As pp_ResultRow, _
    ByVal tablesByRef As Object, _
    ByRef outValue As String, _
    ByRef outErrorText As String _
) As Boolean
    Dim rowVarName As String
    Dim fieldAlias As String
    Dim sourceAlias As String
    Dim tableAlias As String
    Dim tableRef As String
    Dim rowIndex As Long
    Dim targetTable As pp_ResultTable
    Dim targetRowRef As pp_ResultRow

    tokenText = Trim$(tokenText)
    If Len(tokenText) = 0 Then
        outErrorText = "Field reference is empty."
        Exit Function
    End If

    If mp_TryParseRowColumnRef(tokenText, rowVarName, fieldAlias) Then
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
        outValue = currentRowRef.GetByAlias(fieldAlias)
        mp_TryResolveRuntimeValue = True
        Exit Function
    End If

    If mp_TryParseFullRowColumnRef(tokenText, sourceAlias, tableAlias, rowIndex, fieldAlias) Then
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
        If rowIndex + 1 > targetTable.Rows.Count Then
            outErrorText = "Row '" & CStr(rowIndex) & "' is out of range for table '" & tableRef & "'."
            Exit Function
        End If
        Set targetRowRef = targetTable.Rows(rowIndex + 1)
        If Not targetRowRef.HasAlias(fieldAlias) Then
            outErrorText = "Field alias '" & fieldAlias & "' is not available at row '" & CStr(rowIndex) & "'."
            Exit Function
        End If
        outValue = targetRowRef.GetByAlias(fieldAlias)
        mp_TryResolveRuntimeValue = True
        Exit Function
    End If

    outErrorText = "Unsupported field reference '" & tokenText & "'."
End Function

Private Function mp_TryParseRowColumnRef(ByVal refText As String, ByRef outRowVar As String, ByRef outFieldAlias As String) As Boolean
    Dim dotPos As Long
    Dim memberName As String
    refText = Trim$(refText)
    dotPos = InStr(1, refText, ".column[", vbTextCompare)
    If dotPos <= 1 Then Exit Function
    If Right$(refText, 1) <> "]" Then Exit Function
    memberName = Mid$(refText, dotPos + 1, Len("column"))
    If Not mp_IsMemberAllowed(DSL_TYPE_ROW, memberName) Then Exit Function
    outRowVar = Trim$(Left$(refText, dotPos - 1))
    If Len(outRowVar) = 0 Then Exit Function
    If Not mp_IsIdentifier(outRowVar) Then Exit Function
    outFieldAlias = Mid$(refText, dotPos + Len(".column["), Len(refText) - (dotPos + Len(".column[")))
    outFieldAlias = Trim$(outFieldAlias)
    If Len(outFieldAlias) = 0 Then Exit Function
    mp_TryParseRowColumnRef = True
End Function

Private Function mp_TryParseFullRowColumnRef( _
    ByVal refText As String, _
    ByRef outSourceAlias As String, _
    ByRef outTableAlias As String, _
    ByRef outRowIndex As Long, _
    ByRef outFieldAlias As String _
) As Boolean
    Dim sheetPos As Long
    Dim rowPos As Long
    Dim colPos As Long
    Dim rowText As String
    Dim tableRef As String
    Dim memberName As String

    refText = Trim$(refText)
    sheetPos = InStr(1, refText, ".Sheet[", vbTextCompare)
    rowPos = InStr(1, refText, "].row[", vbTextCompare)
    colPos = InStr(1, refText, "].column[", vbTextCompare)
    If sheetPos <= 1 Then Exit Function
    If rowPos <= sheetPos Then Exit Function
    If colPos <= rowPos Then Exit Function
    If Right$(refText, 1) <> "]" Then Exit Function

    outSourceAlias = Trim$(Left$(refText, sheetPos - 1))
    If Len(outSourceAlias) = 0 Then Exit Function

    outTableAlias = Trim$(Mid$(refText, sheetPos + Len(".Sheet["), rowPos - (sheetPos + Len(".Sheet["))))
    If Len(outTableAlias) = 0 Then Exit Function
    tableRef = outSourceAlias & ".Sheet[" & outTableAlias & "]"
    If Not mp_IsSheetRef(tableRef) Then Exit Function

    memberName = "row"
    If Not mp_IsMemberAllowed(DSL_TYPE_SHEET_REF, memberName) Then Exit Function

    rowText = Trim$(Mid$(refText, rowPos + Len("].row["), colPos - (rowPos + Len("].row["))))
    If Len(rowText) = 0 Then Exit Function
    If Not ex_XmlCore.m_TryParseLong(rowText, outRowIndex) Then Exit Function
    If outRowIndex < 0 Then Exit Function

    memberName = "column"
    If Not mp_IsMemberAllowed(DSL_TYPE_SHEET_REF, memberName) Then Exit Function

    outFieldAlias = Trim$(Mid$(refText, colPos + Len("].column["), Len(refText) - (colPos + Len("].column["))))
    If Len(outFieldAlias) = 0 Then Exit Function

    mp_TryParseFullRowColumnRef = True
End Function

Private Function mp_TryParseTableRowRef( _
    ByVal refText As String, _
    ByRef outSourceAlias As String, _
    ByRef outTableAlias As String, _
    ByRef outRowIndex As Long _
) As Boolean
    Dim sheetPos As Long
    Dim rowPos As Long
    Dim rowText As String
    Dim tableRef As String
    Dim memberName As String

    refText = Trim$(refText)
    sheetPos = InStr(1, refText, ".Sheet[", vbTextCompare)
    rowPos = InStr(1, refText, "].row[", vbTextCompare)
    If sheetPos <= 1 Then Exit Function
    If rowPos <= sheetPos Then Exit Function
    If Right$(refText, 1) <> "]" Then Exit Function

    outSourceAlias = Trim$(Left$(refText, sheetPos - 1))
    If Len(outSourceAlias) = 0 Then Exit Function

    outTableAlias = Trim$(Mid$(refText, sheetPos + Len(".Sheet["), rowPos - (sheetPos + Len(".Sheet["))))
    If Len(outTableAlias) = 0 Then Exit Function
    tableRef = outSourceAlias & ".Sheet[" & outTableAlias & "]"
    If Not mp_IsSheetRef(tableRef) Then Exit Function

    memberName = "row"
    If Not mp_IsMemberAllowed(DSL_TYPE_SHEET_REF, memberName) Then Exit Function

    rowText = Trim$(Mid$(refText, rowPos + Len("].row["), Len(refText) - (rowPos + Len("].row["))))
    If Len(rowText) = 0 Then Exit Function
    If Not ex_XmlCore.m_TryParseLong(rowText, outRowIndex) Then Exit Function
    If outRowIndex < 0 Then Exit Function

    mp_TryParseTableRowRef = True
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

Private Function mp_BuildMapKey(ByVal sourceAlias As String, ByVal tableAlias As String, ByVal fieldAlias As String) As String
    mp_BuildMapKey = Trim$(sourceAlias) & ".Sheet[" & Trim$(tableAlias) & "].Map[" & Trim$(fieldAlias) & "]"
End Function

Private Function mp_IsSheetRef(ByVal refText As String) As Boolean
    Dim sheetPos As Long
    Dim openPos As Long
    Dim closePos As Long
    Dim sourceAlias As String
    Dim tableAlias As String

    refText = Trim$(refText)
    sheetPos = InStr(1, refText, ".Sheet[", vbTextCompare)
    If sheetPos <= 1 Then Exit Function

    sourceAlias = Trim$(Left$(refText, sheetPos - 1))
    If Len(sourceAlias) = 0 Then Exit Function

    openPos = sheetPos + Len(".Sheet[")
    closePos = InStr(openPos, refText, "]", vbBinaryCompare)
    If closePos <= openPos Then Exit Function
    If closePos <> Len(refText) Then Exit Function

    tableAlias = Trim$(Mid$(refText, openPos, closePos - openPos))
    If Len(tableAlias) = 0 Then Exit Function

    mp_IsSheetRef = True
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
            outErrorText = "Unsupported callMacro argument '" & partText & "'. Use row variable, quoted string, Source.Sheet[Table].row[N], or Source.Sheet[Table].row[N].column[Field]."
            Exit Function
        End If
        outArgSpecs.Add argSpec
    Next i

    mp_TryParseCallMacroArgs = True
End Function

Private Function mp_TryParseMacroArg(ByVal argText As String, ByRef outArgSpec As Object) As Boolean
    Dim literalText As String
    Dim sourceAlias As String
    Dim tableAlias As String
    Dim rowIndex As Long
    Dim tableRef As String
    Dim fieldAlias As String
    Set outArgSpec = CreateObject("Scripting.Dictionary")
    outArgSpec.CompareMode = 1

    If mp_TryParseQuotedString(argText, literalText) Then
        outArgSpec("Kind") = "string"
        outArgSpec("Value") = literalText
        mp_TryParseMacroArg = True
        Exit Function
    End If

    If mp_IsIdentifier(argText) Then
        outArgSpec("Kind") = "rowvar"
        outArgSpec("Name") = Trim$(argText)
        mp_TryParseMacroArg = True
        Exit Function
    End If

    If mp_TryParseTableRowRef(argText, sourceAlias, tableAlias, rowIndex) Then
        tableRef = sourceAlias & ".Sheet[" & tableAlias & "]"
        outArgSpec("Kind") = "rowref"
        outArgSpec("TableRef") = tableRef
        outArgSpec("RowIndex") = CLng(rowIndex)
        mp_TryParseMacroArg = True
        Exit Function
    End If

    If mp_TryParseFullRowColumnRef(argText, sourceAlias, tableAlias, rowIndex, fieldAlias) Then
        tableRef = sourceAlias & ".Sheet[" & tableAlias & "]"
        outArgSpec("Kind") = "cellref"
        outArgSpec("TableRef") = tableRef
        outArgSpec("RowIndex") = CLng(rowIndex)
        outArgSpec("FieldAlias") = fieldAlias
        mp_TryParseMacroArg = True
    End If
End Function

Private Function mp_ValidateCallMacroArgs( _
    ByVal action As Object, _
    ByVal currentRowVar As String, _
    ByVal allowedTableFields As Object, _
    ByRef outErrorText As String _
) As Boolean
    Dim argSpecs As Collection
    Dim i As Long
    Dim argSpec As Object
    Dim tableRef As String
    Dim resolvedMapKey As String

    If action Is Nothing Then Exit Function
    If Not action.Exists("Args") Then
        mp_ValidateCallMacroArgs = True
        Exit Function
    End If

    Set argSpecs = action("Args")
    For i = 1 To argSpecs.Count
        Set argSpec = argSpecs(i)
        Select Case LCase$(CStr(argSpec("Kind")))
            Case "rowvar"
                If Len(currentRowVar) = 0 Then
                    outErrorText = "callMacro row argument '" & CStr(argSpec("Name")) & "' is not available in this scope."
                    Exit Function
                End If
                If StrComp(CStr(argSpec("Name")), currentRowVar, vbTextCompare) <> 0 Then
                    outErrorText = "callMacro row argument must be current row variable '" & currentRowVar & "'."
                    Exit Function
                End If
            Case "rowref"
                tableRef = CStr(argSpec("TableRef"))
                If allowedTableFields Is Nothing Or Not allowedTableFields.Exists(tableRef) Then
                    outErrorText = "callMacro row argument references unknown table '" & tableRef & "'."
                    Exit Function
                End If
                If CLng(argSpec("RowIndex")) < 0 Then
                    outErrorText = "callMacro row argument index must be >= 0 for '" & tableRef & "'."
                    Exit Function
                End If
            Case "cellref"
                tableRef = CStr(argSpec("TableRef"))
                If allowedTableFields Is Nothing Or Not allowedTableFields.Exists(tableRef) Then
                    outErrorText = "callMacro cell argument references unknown table '" & tableRef & "'."
                    Exit Function
                End If
                If CLng(argSpec("RowIndex")) < 0 Then
                    outErrorText = "callMacro cell argument index must be >= 0 for '" & tableRef & "'."
                    Exit Function
                End If
                If Not mp_TryResolveMapKeyByFieldAlias(allowedTableFields, tableRef, CStr(argSpec("FieldAlias")), resolvedMapKey, outErrorText) Then
                    Exit Function
                End If
        End Select
    Next i

    mp_ValidateCallMacroArgs = True
End Function

Private Function mp_BuildMacroRuntimeArgs( _
    ByVal action As Object, _
    ByVal currentTableRef As String, _
    ByVal currentRowVar As String, _
    ByVal currentRowRef As pp_ResultRow, _
    ByVal tablesByRef As Object _
) As Collection
    Dim result As Collection
    Dim argSpecs As Collection
    Dim i As Long
    Dim argSpec As Object
    Dim argKind As String
    Dim renderedText As String

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
            Case "rowvar"
                If currentRowRef Is Nothing Then
                    Err.Raise vbObjectError + 1601, "ex_PostProcessDsl", "Current row is not available for callMacro row argument."
                End If
                result.Add currentRowRef
            Case "rowref"
                result.Add mp_ResolveRowReferenceArg(argSpec, tablesByRef)
            Case "cellref"
                result.Add mp_ResolveCellReferenceArg(argSpec, tablesByRef)
            Case "string"
                renderedText = mp_RenderTemplate(CStr(argSpec("Value")), currentTableRef, currentRowVar, currentRowRef, tablesByRef)
                result.Add renderedText
            Case Else
                Err.Raise vbObjectError + 1598, "ex_PostProcessDsl", "Unsupported callMacro argument kind: " & argKind
        End Select
    Next i

    Set mp_BuildMacroRuntimeArgs = result
End Function

Private Function mp_ResolveCellReferenceArg(ByVal argSpec As Object, ByVal tablesByRef As Object) As String
    Dim rowObj As pp_ResultRow
    Dim fieldAlias As String

    Set rowObj = mp_ResolveRowReferenceArg(argSpec, tablesByRef)
    fieldAlias = CStr(argSpec("FieldAlias"))
    If Not rowObj.HasAlias(fieldAlias) Then
        Err.Raise vbObjectError + 1605, "ex_PostProcessDsl", "Field alias '" & fieldAlias & "' is not available in referenced row."
    End If
    mp_ResolveCellReferenceArg = rowObj.GetByAlias(fieldAlias)
End Function

Private Function mp_ResolveRowReferenceArg(ByVal argSpec As Object, ByVal tablesByRef As Object) As pp_ResultRow
    Dim tableRef As String
    Dim rowIndex As Long
    Dim tableObj As pp_ResultTable
    Dim rowsList As Collection
    Dim rowObj As pp_ResultRow

    tableRef = CStr(argSpec("TableRef"))
    rowIndex = CLng(argSpec("RowIndex"))
    If rowIndex < 0 Then
        Err.Raise vbObjectError + 1602, "ex_PostProcessDsl", "Row index must be >= 0 in '" & tableRef & ".row[" & CStr(rowIndex) & "]'."
    End If
    If tablesByRef Is Nothing Or Not tablesByRef.Exists(tableRef) Then
        Err.Raise vbObjectError + 1603, "ex_PostProcessDsl", "Table '" & tableRef & "' is not available for row reference."
    End If

    Set tableObj = tablesByRef(tableRef)
    Set rowsList = tableObj.Rows
    If rowIndex + 1 > rowsList.Count Then
        Err.Raise vbObjectError + 1604, "ex_PostProcessDsl", "Row index " & CStr(rowIndex) & " is out of range for table '" & tableRef & "' (rows=" & CStr(rowsList.Count) & ")."
    End If

    Set rowObj = rowsList(rowIndex + 1)
    Set mp_ResolveRowReferenceArg = rowObj
End Function

Private Sub mp_RunMacroWithArgs(ByVal macroName As String, ByVal args As Collection)
    Dim argCount As Long

    macroName = Trim$(macroName)
    If Len(macroName) = 0 Then
        Err.Raise vbObjectError + 1599, "ex_PostProcessDsl", "Macro name is empty."
    End If

    If args Is Nothing Then
        Application.Run macroName
        Exit Sub
    End If

    argCount = args.Count

    Select Case argCount
        Case 0
            Application.Run macroName
        Case 1
            Application.Run macroName, args(1)
        Case 2
            Application.Run macroName, args(1), args(2)
        Case 3
            Application.Run macroName, args(1), args(2), args(3)
        Case 4
            Application.Run macroName, args(1), args(2), args(3), args(4)
        Case 5
            Application.Run macroName, args(1), args(2), args(3), args(4), args(5)
        Case Else
            Err.Raise vbObjectError + 1600, "ex_PostProcessDsl", "Too many callMacro arguments (max 5)."
    End Select
End Sub

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

Private Function mp_IsIdentifier(ByVal valueText As String) As Boolean
    Dim i As Long
    Dim ch As String

    valueText = Trim$(valueText)
    If Len(valueText) = 0 Then Exit Function

    For i = 1 To Len(valueText)
        ch = Mid$(valueText, i, 1)
        If Not ((ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Or ch = "_") Then
            Exit Function
        End If
        If i = 1 And (ch >= "0" And ch <= "9") Then Exit Function
    Next i

    mp_IsIdentifier = True
End Function

Private Function mp_IsMemberAllowed(ByVal objectType As String, ByVal memberName As String) As Boolean
    Dim membersByType As Object
    Dim memberSet As Object

    objectType = LCase$(Trim$(objectType))
    memberName = LCase$(Trim$(memberName))
    If Len(objectType) = 0 Or Len(memberName) = 0 Then Exit Function

    Set membersByType = mp_GetDslMembersByType()
    If membersByType Is Nothing Then Exit Function
    If Not membersByType.Exists(objectType) Then Exit Function
    Set memberSet = membersByType(objectType)
    mp_IsMemberAllowed = memberSet.Exists(memberName)
End Function

Private Function mp_GetDslMembersByType() As Object
    If g_DslMembersByType Is Nothing Then
        Set g_DslMembersByType = CreateObject("Scripting.Dictionary")
        g_DslMembersByType.CompareMode = 1
        g_DslMembersByType.Add DSL_TYPE_SHEET_REF, mp_CreateMemberSet("rows,column,row")
        g_DslMembersByType.Add DSL_TYPE_ROW, mp_CreateMemberSet("column")
    End If
    Set mp_GetDslMembersByType = g_DslMembersByType
End Function

Private Function mp_CreateMemberSet(ByVal csvMembers As String) As Object
    Dim result As Object
    Dim parts As Variant
    Dim i As Long
    Dim memberName As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1
    parts = Split(csvMembers, ",")
    For i = LBound(parts) To UBound(parts)
        memberName = Trim$(CStr(parts(i)))
        If Len(memberName) > 0 Then result(memberName) = True
    Next i
    Set mp_CreateMemberSet = result
End Function

Private Function mp_GetLastUsedRow(ByVal ws As Worksheet) As Long
    On Error GoTo ExitFn
    mp_GetLastUsedRow = ws.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
ExitFn:
End Function

Private Function mp_GetLastUsedColumn(ByVal ws As Worksheet) As Long
    On Error GoTo ExitFn
    mp_GetLastUsedColumn = ws.Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
ExitFn:
End Function
