Attribute VB_Name = "ex_ScriptDSL"
Option Explicit

Private Const SCRIPT_KEY As String = "PostProcess.Script"
Private Const ACTION_CALL_MACRO As String = "callmacro"
Private Const ACTION_CALL_MACRO_OBJECT As String = "callmacroobject"
Private Const ACTION_LET As String = "let"
Private Const ACTION_ASSIGN As String = "assign"
Private Const ACTION_BREAK As String = "break"
Private Const ACTION_CONTINUE As String = "continue"
Private Const ACTION_RETURN As String = "return"
Private Const ASSIGN_KIND_CALL_MACRO As String = "callmacro"
Private Const ASSIGN_KIND_CALL_MACRO_OBJECT As String = "callmacroobject"
Private Const ASSIGN_KIND_STRING_EXPR As String = "stringexpr"
Private Const EXPR_PART_LITERAL As String = "literal"
Private Const EXPR_PART_TOKEN As String = "token"

Private Const LOOP_TARGET_TABLE_ROWS As String = "tablerows"
Private Const LOOP_TARGET_ROW_COLUMNS As String = "rowcolumns"
Private Const LOOP_TARGET_MEMBER_ROWS As String = "memberrows"

Private Const VAR_TYPE_ROW As String = "row"
Private Const VAR_TYPE_COLUMN As String = "column"
Private Const VAR_TYPE_STRING As String = "string"
Private Const VAR_TYPE_OBJECT As String = "object"

Private Const BATCH_KEYS_RESULTS_MEMBER As String = "keysresults"
Private Const BATCH_CONTEXT_KEYS_RESULTS_TABLE_REF As String = "KeysResultsTableRef"
Private Const BATCH_KEYRESULT_KEY_ALIAS As String = "Key"
Private Const BATCH_KEYRESULT_KEYFIELD_SUFFIX As String = ".KeyFieldAlias"
Private Const DEBUG_LOG_PATH As String = "Logs\personalcard_pipeline.log"
Private Const DEBUG_LOG_ENABLED As Boolean = False

Private Const EXEC_FLOW_NONE As String = ""
Private Const EXEC_FLOW_BREAK As String = "break"
Private Const EXEC_FLOW_CONTINUE As String = "continue"
Private Const EXEC_FLOW_RETURN As String = "return"

Private g_ParsedBlocksByScriptCacheKey As Object
Private g_ValidationCacheBySignatureKey As Object
Private g_ParseHasBacktickLiterals As Boolean

Public Sub m_ResetScriptCache()
    Set g_ParsedBlocksByScriptCacheKey = Nothing
    Set g_ValidationCacheBySignatureKey = Nothing
End Sub

Public Function m_ParseScriptToBlocks( _
    ByVal scriptText As String, _
    ByRef outBlocks As Collection, _
    ByRef outErrorText As String _
) As Boolean
    m_ParseScriptToBlocks = mp_ParseScript(scriptText, outBlocks, outErrorText)
End Function

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
    mp_MarkValidationCache scriptConfigKey, scriptText, validationSignature

    m_ValidateScriptAgainstConfig = True
    Exit Function

EH:
    outErrorText = "Script validation runtime error"
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
    Optional ByVal scriptConfigKey As String = SCRIPT_KEY, _
    Optional ByVal layoutBatchApplyRefresh As Boolean = True _
)
    Dim scriptText As String
    Dim blocks As Collection
    Dim parseOrValidationError As String
    Dim runtimeValidationSignature As String
    Dim ctxTablesByRef As Object
    Dim ctxFields As Object
    Dim postProcessFooterLines As Collection
    Dim usedCols As Long
    Dim startedDeferredRender As Boolean
    Dim startedLayoutBatch As Boolean
    Dim layoutBatchError As String
    Dim prevScreenUpdating As Boolean
    Dim runStamp As String
    Dim prevActiveSheet As Worksheet
    Dim prevActiveSheetName As String

    On Error GoTo EH

    If ws Is Nothing Then Exit Sub
    If cfg Is Nothing Then Exit Sub
    If resultTables Is Nothing Then Exit Sub

    runStamp = Format$(Now, "yyyy-mm-dd HH:nn:ss")
    mp_DebugLog "RUN START " & runStamp & " | scriptKey=" & CStr(scriptConfigKey) & " | sheet=" & ws.Name

    prevScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    If Not mp_TryGetCompiledScriptBlocks(cfg, scriptConfigKey, scriptText, blocks, parseOrValidationError) Then
        mp_DebugLog "compile failed: " & parseOrValidationError
        Err.Raise vbObjectError + 1592, "ex_ScriptDSL", parseOrValidationError
    End If
    If Len(scriptText) = 0 Then Exit Sub
    mp_DebugLog "compile ok | scriptLength=" & CStr(Len(scriptText))

    ex_ResultRuntimeAdapter.m_BuildRuntimeContext resultTables, ctxTablesByRef, ctxFields
    mp_DebugLog "runtime context built"

    runtimeValidationSignature = mp_BuildValidationSignature(ctxFields)
    If Not mp_IsValidationCacheHit(scriptConfigKey, scriptText, runtimeValidationSignature) Then
        mp_DebugLog "validate start"
        If Not mp_ValidateBlocks(blocks, ctxFields, parseOrValidationError) Then
            mp_DebugLog "validate failed: " & parseOrValidationError
            Err.Raise vbObjectError + 1591, "ex_ScriptDSL", "Script validation failed: " & parseOrValidationError
        End If
        mp_MarkValidationCache scriptConfigKey, scriptText, runtimeValidationSignature
        mp_DebugLog "validate ok"
    Else
        mp_DebugLog "validate cache hit"
    End If

    Set postProcessFooterLines = New Collection
    usedCols = mp_GetLastUsedColumn(ws)
    If usedCols <= 0 Then usedCols = 1

    ex_ResultLayoutItemsRt.m_BeginBatchUpdate ws
    startedLayoutBatch = True
    ex_PostProcessActions.m_ResetScriptHeaderCursor ws
    ex_PostProcessActions.m_ResetScriptFooterCursor ws
    ex_PostProcessActions.m_SetExecutionSheetContext ws
    ex_PostProcessActions.m_BeginDeferredRender ws
    startedDeferredRender = True

    On Error Resume Next
    Set prevActiveSheet = ActiveSheet
    If Not prevActiveSheet Is Nothing Then prevActiveSheetName = prevActiveSheet.Name
    On Error GoTo EH

    If Not ws Is ActiveSheet Then ws.Activate

    mp_DebugLog "execute blocks start"
    mp_ExecuteBlocks ws, blocks, ctxTablesByRef, postProcessFooterLines, usedCols
    mp_DebugLog "execute blocks ok"
    ex_PostProcessActions.m_CommitDeferredRender ws
    mp_DebugLog "deferred commit ok"
    If Not ex_ResultLayoutItemsRt.m_EndBatchUpdate(ws, layoutBatchApplyRefresh, layoutBatchError) Then
        If Len(layoutBatchError) = 0 Then layoutBatchError = "Layout batch update failed."
        Err.Raise vbObjectError + 1593, "ex_ScriptDSL", layoutBatchError
    End If
    startedLayoutBatch = False

    If Len(prevActiveSheetName) > 0 Then
        On Error Resume Next
        ThisWorkbook.Worksheets(prevActiveSheetName).Activate
        On Error GoTo EH
    End If

    ex_PostProcessActions.m_SetExecutionSheetContext Nothing
    startedDeferredRender = False
    ' Временно отключено: автопрокрутка к футеру после post-processing.
    ' ex_PostProcessActions.m_ScrollToScriptResults ws
    Application.ScreenUpdating = prevScreenUpdating
    Exit Sub

EH:
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description

    mp_DebugLog "ERROR: [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description
    If startedDeferredRender Then
        On Error Resume Next
        ex_PostProcessActions.m_EndDeferredRender ws
        On Error GoTo 0
    End If
    If startedLayoutBatch Then
        On Error Resume Next
        layoutBatchError = vbNullString
        ex_ResultLayoutItemsRt.m_EndBatchUpdate ws, False, layoutBatchError
        On Error GoTo 0
    End If
    If Len(prevActiveSheetName) > 0 Then
        On Error Resume Next
        ThisWorkbook.Worksheets(prevActiveSheetName).Activate
        On Error GoTo 0
    End If
    On Error Resume Next
    ex_PostProcessActions.m_SetExecutionSheetContext Nothing
    On Error GoTo 0
    On Error Resume Next
    Application.ScreenUpdating = prevScreenUpdating
    On Error GoTo 0
    If errNumber = 0 Then errNumber = vbObjectError + 1590
    If Len(errSource) = 0 Then errSource = "ex_ScriptDSL"
    If Len(errDescription) = 0 Then errDescription = "Unknown script execution failure."
    Err.Raise errNumber, errSource, errDescription
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
    Dim scriptCacheKey As String
    Dim stepName As String

    On Error GoTo EH

    stepName = "init"
    outScriptText = vbNullString
    outErrorText = vbNullString
    Set outBlocks = Nothing

    stepName = "normalize-script-key"
    normalizedScriptKey = Trim$(scriptConfigKey)
    If Len(normalizedScriptKey) = 0 Then normalizedScriptKey = SCRIPT_KEY

    stepName = "load-script-text"
    If Not ex_ScriptSourceLoader.m_TryGetScriptText(cfg, normalizedScriptKey, outScriptText, outErrorText) Then Exit Function
    If Len(outScriptText) = 0 Then
        mp_TryGetCompiledScriptBlocks = True
        Exit Function
    End If

    scriptCacheKey = mp_BuildScriptCacheKey(normalizedScriptKey, outScriptText)
    mp_EnsureScriptCaches

    stepName = "check-script-cache"
    If g_ParsedBlocksByScriptCacheKey.Exists(scriptCacheKey) Then
        Set outBlocks = g_ParsedBlocksByScriptCacheKey(scriptCacheKey)
        mp_TryGetCompiledScriptBlocks = True
        Exit Function
    End If

    stepName = "parse-script"
    If Not ex_ScriptDslParser.m_ParseScript(outScriptText, parsedBlocks, parseError) Then
        outErrorText = "Script parse failed: " & parseError
        Exit Function
    End If

    stepName = "update-script-cache"
    Set g_ParsedBlocksByScriptCacheKey(scriptCacheKey) = parsedBlocks

    stepName = "set-output-blocks"
    Set outBlocks = g_ParsedBlocksByScriptCacheKey(scriptCacheKey)
    mp_TryGetCompiledScriptBlocks = True
    Exit Function

EH:
    outErrorText = "Script compile runtime error at step '" & stepName & "': [" & Err.Source & " #" & CStr(Err.Number) & "] " & Err.Description
    mp_TryGetCompiledScriptBlocks = False
End Function

Private Function mp_IsValidationCacheHit( _
    ByVal scriptConfigKey As String, _
    ByVal scriptText As String, _
    ByVal validationSignature As String _
) As Boolean
    Dim validationCacheKey As String

    If Len(validationSignature) = 0 Then Exit Function

    validationCacheKey = mp_BuildValidationCacheKey(scriptConfigKey, scriptText, validationSignature)
    mp_EnsureScriptCaches
    If Not g_ValidationCacheBySignatureKey.Exists(validationCacheKey) Then Exit Function

    mp_IsValidationCacheHit = True
End Function

Private Sub mp_MarkValidationCache( _
    ByVal scriptConfigKey As String, _
    ByVal scriptText As String, _
    ByVal validationSignature As String _
)
    Dim validationCacheKey As String

    If Len(validationSignature) = 0 Then Exit Sub

    validationCacheKey = mp_BuildValidationCacheKey(scriptConfigKey, scriptText, validationSignature)
    mp_EnsureScriptCaches
    g_ValidationCacheBySignatureKey(validationCacheKey) = True
End Sub

Private Function mp_BuildValidationCacheKey( _
    ByVal scriptConfigKey As String, _
    ByVal scriptText As String, _
    ByVal validationSignature As String _
) As String
    mp_BuildValidationCacheKey = mp_BuildScriptCacheKey(scriptConfigKey, scriptText) & "|V:" & validationSignature
End Function

Private Function mp_BuildScriptCacheKey(ByVal scriptConfigKey As String, ByVal scriptText As String) As String
    Dim normalizedScriptKey As String

    normalizedScriptKey = Trim$(scriptConfigKey)
    If Len(normalizedScriptKey) = 0 Then normalizedScriptKey = SCRIPT_KEY

    mp_BuildScriptCacheKey = normalizedScriptKey & "|S:" & mp_BuildScriptContentToken(scriptText)
End Function

Private Function mp_BuildScriptContentToken(ByVal scriptText As String) As String
    ' Avoid numeric hash arithmetic here: VBA floating/integer casts can overflow on some inputs.
    mp_BuildScriptContentToken = CStr(Len(scriptText)) & ":" & scriptText
End Function

Private Sub mp_EnsureScriptCaches()
    If g_ParsedBlocksByScriptCacheKey Is Nothing Then
        Set g_ParsedBlocksByScriptCacheKey = CreateObject("Scripting.Dictionary")
        g_ParsedBlocksByScriptCacheKey.CompareMode = 0
    End If

    If g_ValidationCacheBySignatureKey Is Nothing Then
        Set g_ValidationCacheBySignatureKey = CreateObject("Scripting.Dictionary")
        g_ValidationCacheBySignatureKey.CompareMode = 0
    End If
End Sub

Private Function mp_BuildValidationSignature(ByVal allowedTableFields As Object) As String
    Dim tableKeys As Variant
    Dim tableKey As Variant
    Dim fieldMap As Object
    Dim fieldKeys As Variant
    Dim i As Long

    On Error GoTo EH

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
    Exit Function

EH:
    mp_BuildValidationSignature = "error:" & CStr(Err.Number) & ":" & Err.Description
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
    Dim prevParseHasBacktickLiterals As Boolean

    On Error GoTo EH

    sourceText = ex_ScriptDslParser.m_NormalizeScript(scriptText)
    prevParseHasBacktickLiterals = g_ParseHasBacktickLiterals
    g_ParseHasBacktickLiterals = (InStr(1, sourceText, "`", vbBinaryCompare) > 0)
    pos = 1
    lineNo = 1
    If Not mp_ParseStatements(sourceText, pos, lineNo, outBlocks, False, outErrorText) Then
        g_ParseHasBacktickLiterals = prevParseHasBacktickLiterals
        Exit Function
    End If
    g_ParseHasBacktickLiterals = prevParseHasBacktickLiterals
    mp_ParseScript = True
    Exit Function

EH:
    g_ParseHasBacktickLiterals = prevParseHasBacktickLiterals
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function mp_ValidateBlocks( _
    ByVal blocks As Collection, _
    ByVal allowedTableFields As Object, _
    ByRef outErrorText As String _
) As Boolean
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
            Err.Raise vbObjectError + 1618, "ex_ScriptDSL", "'" & execFlow & "' is only allowed inside for-loop."
        Case Else
            Err.Raise vbObjectError + 1619, "ex_ScriptDSL", "Unsupported control-flow signal: " & execFlow
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
    Dim targetDescriptor As Object
    Dim targetCurrentTableRef As String
    Dim targetResolveError As String
    Dim loopItemObject As Object
    Dim loopSourceRowVar As String
    Dim loopSourceRowRef As obj_ResultRow
    Dim targetLoopKind As String
    Dim existingRuntimeValue As obj_ScriptScopeValue

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
                macroResult = ex_ScriptActionInvoker.m_RunMacroWithArgsReturn(CStr(statement("MacroName")), macroArgs)
                Call mp_ConvertVariantToString(macroResult)
                On Error GoTo 0

            Case ACTION_CALL_MACRO_OBJECT
                On Error GoTo CallMacroErr
                Set macroArgs = mp_BuildMacroRuntimeArgs(statement, currentTableRef, currentRowVar, currentRowRef, tablesByRef, runtimeVars)
                Set macroResultObject = ex_ScriptActionInvoker.m_RunObjectMacroWithArgsReturn(CStr(statement("MacroName")), macroArgs)
                On Error GoTo 0

            Case ACTION_LET
                letVarName = CStr(statement("VarName"))
                If localLetDeclarations.Exists(letVarName) Then
                    Err.Raise vbObjectError + 1617, "ex_ScriptDSL", "Variable '" & letVarName & "' is already declared in this scope."
                End If
                assignKind = mp_GetStatementAssignKind(statement)
                Select Case assignKind
                    Case ASSIGN_KIND_STRING_EXPR
                        mp_SetRuntimeScopeString runtimeVars, letVarName, mp_EvaluateStringExpression( _
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
                        macroResult = ex_ScriptActionInvoker.m_RunMacroWithArgsReturn(CStr(statement("MacroName")), macroArgs)
                        mp_SetRuntimeScopeString runtimeVars, letVarName, mp_ConvertVariantToString(macroResult)
                        On Error GoTo 0
                    Case ASSIGN_KIND_CALL_MACRO_OBJECT
                        On Error GoTo CallMacroErr
                        Set macroArgs = mp_BuildMacroRuntimeArgs(statement, currentTableRef, currentRowVar, currentRowRef, tablesByRef, runtimeVars)
                        Set macroResultObject = ex_ScriptActionInvoker.m_RunObjectMacroWithArgsReturn(CStr(statement("MacroName")), macroArgs)
                        mp_SetRuntimeScopeObject runtimeVars, letVarName, macroResultObject
                        On Error GoTo 0
                    Case Else
                        Err.Raise vbObjectError + 1624, "ex_ScriptDSL", "Unsupported assignment kind: " & assignKind
                End Select
                mp_SetScopeValue localLetDeclarations, letVarName, "1"

            Case ACTION_ASSIGN
                If runtimeVars Is Nothing Or Not runtimeVars.Exists(CStr(statement("VarName"))) Then
                    Err.Raise vbObjectError + 1614, "ex_ScriptDSL", "Assignment to undeclared variable '" & CStr(statement("VarName")) & "'."
                End If
                If Not ex_ScriptScopeValue.m_TryGetScopeValue(runtimeVars, CStr(statement("VarName")), existingRuntimeValue, targetResolveError) Then
                    Err.Raise vbObjectError + 1614, "ex_ScriptDSL", "Assignment to variable '" & CStr(statement("VarName")) & "' failed: " & targetResolveError
                End If
                assignKind = mp_GetStatementAssignKind(statement)
                Select Case assignKind
                    Case ASSIGN_KIND_STRING_EXPR
                        If existingRuntimeValue.HasObjectValue Then
                            Err.Raise vbObjectError + 1615, "ex_ScriptDSL", "Assignment type mismatch for variable '" & CStr(statement("VarName")) & "': expected row object result."
                        End If
                        mp_SetRuntimeScopeString runtimeVars, CStr(statement("VarName")), mp_EvaluateStringExpression( _
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
                        macroResult = ex_ScriptActionInvoker.m_RunMacroWithArgsReturn(CStr(statement("MacroName")), macroArgs)
                        If existingRuntimeValue.HasObjectValue Then
                            Err.Raise vbObjectError + 1615, "ex_ScriptDSL", "Assignment type mismatch for variable '" & CStr(statement("VarName")) & "': expected row/object result."
                        End If
                        mp_SetRuntimeScopeString runtimeVars, CStr(statement("VarName")), mp_ConvertVariantToString(macroResult)
                        On Error GoTo 0
                    Case ASSIGN_KIND_CALL_MACRO_OBJECT
                        On Error GoTo CallMacroErr
                        Set macroArgs = mp_BuildMacroRuntimeArgs(statement, currentTableRef, currentRowVar, currentRowRef, tablesByRef, runtimeVars)
                        Set macroResultObject = ex_ScriptActionInvoker.m_RunObjectMacroWithArgsReturn(CStr(statement("MacroName")), macroArgs)
                        If Not existingRuntimeValue.HasObjectValue Then
                            Err.Raise vbObjectError + 1616, "ex_ScriptDSL", "Assignment type mismatch for variable '" & CStr(statement("VarName")) & "': expected string-compatible result."
                        End If
                        mp_SetRuntimeScopeObject runtimeVars, CStr(statement("VarName")), macroResultObject
                        On Error GoTo 0
                    Case Else
                        Err.Raise vbObjectError + 1624, "ex_ScriptDSL", "Unsupported assignment kind: " & assignKind
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
                Set targetDescriptor = mp_GetForTargetDescriptor(statement)
                If targetDescriptor Is Nothing Then
                    Err.Raise vbObjectError + 1621, "ex_ScriptDSL", "For-loop target descriptor is missing."
                End If
                targetResolveError = vbNullString
                targetCurrentTableRef = vbNullString
                Set rowsList = Nothing
                Set loopSourceRowRef = Nothing
                If Not ex_ScriptForTargetResolver.m_ResolveForTargetRuntimeContext( _
                    targetDescriptor, _
                    runtimeVars, _
                    tablesByRef, _
                    rowsList, _
                    loopSourceRowRef, _
                    targetCurrentTableRef, _
                    targetResolveError _
                ) Then
                    Err.Raise vbObjectError + 1628, "ex_ScriptDSL", "Unable to resolve for-loop target: " & targetResolveError
                End If
                If rowsList Is Nothing Then
                    Err.Raise vbObjectError + 1629, "ex_ScriptDSL", "Resolved for-loop target produced no rows collection."
                End If

                loopVarName = CStr(statement("LoopVar"))
                targetLoopKind = LCase$(CStr(targetDescriptor("LoopTarget")))
                For rowIdx = 1 To rowsList.Count
                    Set loopItemObject = rowsList(rowIdx)
                    Set childRuntimeVars = mp_CloneVarScope(runtimeVars)
                    mp_SetRuntimeScopeObject childRuntimeVars, loopVarName, loopItemObject

                    If TypeOf loopItemObject Is obj_ResultRow Then
                        Set rowRef = loopItemObject
                    Else
                        Set rowRef = Nothing
                    End If

                    If targetLoopKind = LOOP_TARGET_ROW_COLUMNS Then
                        loopSourceRowVar = CStr(targetDescriptor("SourceRowVar"))
                        bodyFlow = mp_ExecuteStatements(ws, statement("Body"), tablesByRef, postProcessFooterLines, usedCols, currentTableRef, loopSourceRowVar, loopSourceRowRef, childRuntimeVars)
                    ElseIf Len(targetCurrentTableRef) > 0 Then
                        bodyFlow = mp_ExecuteStatements(ws, statement("Body"), tablesByRef, postProcessFooterLines, usedCols, targetCurrentTableRef, loopVarName, rowRef, childRuntimeVars)
                    Else
                        bodyFlow = mp_ExecuteStatements(ws, statement("Body"), tablesByRef, postProcessFooterLines, usedCols, currentTableRef, loopVarName, rowRef, childRuntimeVars)
                    End If

                    mp_SyncAssignedParentScope runtimeVars, childRuntimeVars
                    Select Case LCase$(bodyFlow)
                        Case EXEC_FLOW_NONE
                            ' no-op
                        Case EXEC_FLOW_CONTINUE
                            ' next loop item
                        Case EXEC_FLOW_BREAK
                            Exit For
                        Case EXEC_FLOW_RETURN
                            mp_ExecuteStatements = EXEC_FLOW_RETURN
                            Exit Function
                        Case Else
                            Err.Raise vbObjectError + 1620, "ex_ScriptDSL", "Unsupported control-flow signal in for-loop: " & bodyFlow
                    End Select
                Next rowIdx

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
                Err.Raise vbObjectError + 1593, "ex_ScriptDSL", "Unsupported statement type: " & statementType
        End Select
    Next i
    mp_ExecuteStatements = EXEC_FLOW_NONE
    Exit Function

CallMacroErr:
    Err.Raise vbObjectError + 1597, "ex_ScriptDSL", "callMacro failed for '" & CStr(statement("MacroName")) & "': " & Err.Description
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
        Case "callmacroobject"
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
    Dim targetDescriptor As Object
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

    If Not ex_ScriptForTargetResolver.m_TryParseForTarget(targetText, targetDescriptor, outErrorText) Then
        outErrorText = outErrorText & " (line " & CStr(stmtLine) & ")"
        Exit Function
    End If
    Set outStatement("TargetDescriptor") = targetDescriptor
    mp_CopyForTargetDescriptor outStatement, targetDescriptor

    outStatement.Add "Body", bodyStatements
    outStatement("Line") = CLng(stmtLine)
    mp_TryParseForStatement = True
End Function

Private Sub mp_CopyForTargetDescriptor(ByVal targetStatement As Object, ByVal targetDescriptor As Object)
    Dim key As Variant

    If targetStatement Is Nothing Then Exit Sub
    If targetDescriptor Is Nothing Then Exit Sub

    For Each key In targetDescriptor.Keys
        targetStatement(CStr(key)) = targetDescriptor(CStr(key))
    Next key
End Sub

Private Function mp_GetForTargetDescriptor(ByVal statement As Object) As Object
    Dim descriptor As Object

    If statement Is Nothing Then Exit Function

    If statement.Exists("TargetDescriptor") Then
        If IsObject(statement("TargetDescriptor")) Then
            Set mp_GetForTargetDescriptor = statement("TargetDescriptor")
            Exit Function
        End If
    End If

    Set descriptor = CreateObject("Scripting.Dictionary")
    descriptor.CompareMode = 1
    If statement.Exists("LoopTarget") Then descriptor("LoopTarget") = statement("LoopTarget")
    If statement.Exists("TableRef") Then descriptor("TableRef") = statement("TableRef")
    If statement.Exists("SourceRowVar") Then descriptor("SourceRowVar") = statement("SourceRowVar")
    If statement.Exists("BatchVar") Then descriptor("BatchVar") = statement("BatchVar")
    If statement.Exists("ScopeVar") Then descriptor("ScopeVar") = statement("ScopeVar")
    If statement.Exists("SourceAlias") Then descriptor("SourceAlias") = statement("SourceAlias")
    If statement.Exists("TableAlias") Then descriptor("TableAlias") = statement("TableAlias")

    If descriptor.Exists("LoopTarget") Then Set mp_GetForTargetDescriptor = descriptor
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
    If Not ex_ScriptParserCore.m_IsIdentifier(outRowVarName) Then Exit Function

    outTargetText = Trim$(Mid$(bodyText, inPos + 4))
    If Len(outTargetText) = 0 Then Exit Function

    mp_TryParseForHeaderText = True
End Function

Private Function mp_TryParseBatchKeysResultsTarget( _
    ByVal targetText As String, _
    ByRef outBatchVarName As String _
) As Boolean
    Dim memberName As String

    If Not mp_TryParseVariableMemberRef(targetText, outBatchVarName, memberName) Then Exit Function
    If StrComp(memberName, BATCH_KEYS_RESULTS_MEMBER, vbTextCompare) <> 0 Then Exit Function

    mp_TryParseBatchKeysResultsTarget = True
End Function

Private Function mp_TryParseScopedTableRowsTarget( _
    ByVal targetText As String, _
    ByRef outScopeVarName As String, _
    ByRef outSourceAlias As String, _
    ByRef outTableAlias As String, _
    ByRef outTableRef As String _
) As Boolean
    Dim dotPos As Long
    Dim scopedTail As String

    targetText = Trim$(targetText)
    dotPos = InStr(1, targetText, ".", vbBinaryCompare)
    If dotPos <= 1 Then Exit Function

    outScopeVarName = Trim$(Left$(targetText, dotPos - 1))
    If Not ex_ScriptParserCore.m_IsIdentifier(outScopeVarName) Then Exit Function

    scopedTail = Trim$(Mid$(targetText, dotPos + 1))
    If Not ex_obj_ResultTableDsl.m_TryParseTableRowsRef(scopedTail, outTableRef) Then Exit Function
    If Not mp_TryParseTableRef(outTableRef, outSourceAlias, outTableAlias) Then Exit Function

    mp_TryParseScopedTableRowsTarget = True
End Function

Private Function mp_TryParseTableRef( _
    ByVal tableRef As String, _
    ByRef outSourceAlias As String, _
    ByRef outTableAlias As String _
) As Boolean
    Dim sheetPos As Long
    Dim sheetStart As Long
    Dim sheetEnd As Long

    tableRef = Trim$(tableRef)
    sheetPos = InStr(1, tableRef, ".Sheet[", vbTextCompare)
    If sheetPos <= 1 Then Exit Function

    outSourceAlias = Trim$(Left$(tableRef, sheetPos - 1))
    If Not ex_ScriptParserCore.m_TryNormalizeSourceAliasToken(outSourceAlias, outSourceAlias) Then Exit Function

    sheetStart = sheetPos + Len(".Sheet[")
    sheetEnd = InStr(sheetStart, tableRef, "]", vbBinaryCompare)
    If sheetEnd <= sheetStart Then Exit Function

    outTableAlias = Trim$(Mid$(tableRef, sheetStart, sheetEnd - sheetStart))
    If Len(outTableAlias) = 0 Then Exit Function
    If sheetEnd <> Len(tableRef) Then Exit Function

    mp_TryParseTableRef = True
End Function

Private Function mp_TryResolveBatchKeysResultsTableRef( _
    ByVal batchContextRow As obj_ResultRow, _
    ByRef outTableRef As String _
) As Boolean
    If batchContextRow Is Nothing Then Exit Function
    If Not batchContextRow.HasAlias(BATCH_CONTEXT_KEYS_RESULTS_TABLE_REF) Then Exit Function

    outTableRef = Trim$(CStr(batchContextRow.Column(BATCH_CONTEXT_KEYS_RESULTS_TABLE_REF)))
    If Len(outTableRef) = 0 Then Exit Function
    If Not ex_ScriptParserCore.m_IsSheetRef(outTableRef) Then Exit Function

    mp_TryResolveBatchKeysResultsTableRef = True
End Function

Private Function mp_ResolveScopedTableKeyFieldAlias( _
    ByVal scopeRow As obj_ResultRow, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String _
) As String
    Dim compositeAlias As String

    If scopeRow Is Nothing Then Exit Function

    compositeAlias = Trim$(sourceAlias) & "." & Trim$(tableAlias) & BATCH_KEYRESULT_KEYFIELD_SUFFIX
    If scopeRow.HasAlias(compositeAlias) Then
        mp_ResolveScopedTableKeyFieldAlias = Trim$(CStr(scopeRow.Column(compositeAlias)))
        If Len(mp_ResolveScopedTableKeyFieldAlias) > 0 Then Exit Function
    End If

    If scopeRow.HasAlias("KeyFieldAlias") Then
        mp_ResolveScopedTableKeyFieldAlias = Trim$(CStr(scopeRow.Column("KeyFieldAlias")))
    End If
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

    If Left$(LCase$(trimmedRhs), Len("callmacroobject(")) = "callmacroobject(" Then
        If Not mp_TryParseAction(trimmedRhs, actionStatement, outErrorText) Then Exit Function
        If LCase$(CStr(actionStatement("Type"))) <> ACTION_CALL_MACRO_OBJECT Then
            outErrorText = "Unsupported assignment action: '" & trimmedRhs & "'."
            Exit Function
        End If
        actionStatement("AssignKind") = ASSIGN_KIND_CALL_MACRO_OBJECT
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
    Dim inBacktickLiteral As Boolean
    Dim parenDepth As Long
    Dim bracketDepth As Long
    Dim braceDepth As Long

    Set outTerms = New Collection
    currentTerm = vbNullString

    For i = 1 To Len(expressionBody)
        ch = Mid$(expressionBody, i, 1)
        If ch = """" And Not mp_IsEscapedQuote(expressionBody, i) Then
            If Not inBacktickLiteral Then inQuotes = Not inQuotes
            currentTerm = currentTerm & ch
        ElseIf ch = "`" Then
            If Not inQuotes Then inBacktickLiteral = Not inBacktickLiteral
            currentTerm = currentTerm & ch
        ElseIf Not inQuotes And Not inBacktickLiteral Then
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
    If inBacktickLiteral Then
        outErrorText = "Unterminated backtick string in string expression."
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
    Dim targetDescriptor As Object
    Dim loopItemType As String
    Dim targetCurrentTableRef As String

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
            Case ACTION_CALL_MACRO, ACTION_CALL_MACRO_OBJECT
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
                    Case ASSIGN_KIND_CALL_MACRO, ASSIGN_KIND_CALL_MACRO_OBJECT
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
                    Case ASSIGN_KIND_CALL_MACRO, ASSIGN_KIND_CALL_MACRO_OBJECT
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
                loopVarName = CStr(statement("LoopVar"))
                If Not mp_TryValidateScriptVariableName(loopVarName, statementLine, outErrorText) Then Exit Function
                Set targetDescriptor = mp_GetForTargetDescriptor(statement)
                If targetDescriptor Is Nothing Then
                    outErrorText = "For-loop target descriptor is missing at line " & CStr(statementLine) & "."
                    Exit Function
                End If

                If Not ex_ScriptForTargetResolver.m_ValidateForTargetDescriptor( _
                    targetDescriptor, _
                    scopeVarTypes, _
                    allowedTableFields, _
                    loopItemType, _
                    targetCurrentTableRef, _
                    outErrorText _
                ) Then Exit Function

                Set childScopeVarTypes = mp_CloneVarScope(scopeVarTypes)
                mp_SetScopeValue childScopeVarTypes, loopVarName, loopItemType

                loopTarget = LCase$(CStr(targetDescriptor("LoopTarget")))
                Select Case loopTarget
                    Case LOOP_TARGET_ROW_COLUMNS
                        sourceRowVarName = CStr(targetDescriptor("SourceRowVar"))
                        If Not mp_ValidateStatements(statement("Body"), allowedTableFields, currentTableRef, sourceRowVarName, childScopeVarTypes, loopDepth + 1, outErrorText) Then Exit Function
                    Case Else
                        tableRef = targetCurrentTableRef
                        If Len(tableRef) = 0 Then tableRef = currentTableRef
                        If Not mp_ValidateStatements(statement("Body"), allowedTableFields, tableRef, loopVarName, childScopeVarTypes, loopDepth + 1, outErrorText) Then Exit Function
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
    Dim partText As String
    Dim groupedConditionText As String

    If Not mp_TrySplitConditionExpression(conditionText, condParts, condOps, outErrorText) Then Exit Function

    For i = 1 To condParts.Count
        partText = CStr(condParts(i))
        If mp_TryUnwrapConditionGroup(partText, groupedConditionText) Then
            If Not mp_ValidateConditionText(groupedConditionText, currentTableRef, currentRowVar, scopeVarTypes, allowedTableFields, outErrorText) Then
                Exit Function
            End If
        Else
            If Not mp_ParseConditionPart(partText, leftTokenText, opText, rightTokenText, rightIsToken) Then
                outErrorText = "Unsupported condition token: '" & Trim$(partText) & "'."
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

                If ex_ScriptParserCore.m_IsIdentifier(tokenText) Then
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
                            If Not ex_ScriptDslContracts.m_IsMemberAllowed(ex_ScriptDslContracts.TYPE_COLUMN, memberName) Then
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
    If Not ex_ScriptParserCore.m_IsIdentifier(variableName) Then
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
        Case "if", "else", "for", "callmacro", "callmacroobject", "let", "in", "break", "continue", "return"
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
    Dim inBacktickLiteral As Boolean

    mp_SkipWhitespace sourceText, pos, lineNo
    If pos > Len(sourceText) Or Mid$(sourceText, pos, 1) <> openChar Then Exit Function

    startPos = pos + 1
    depth = 1
    i = pos + 1
    Do While i <= Len(sourceText)
        ch = Mid$(sourceText, i, 1)

        If g_ParseHasBacktickLiterals And ch = "`" And Not inQuotes Then
            inBacktickLiteral = Not inBacktickLiteral
            i = i + 1
            GoTo ContinueLoop
        End If

        If ch = vbLf Then lineNo = lineNo + 1

        If Not inBacktickLiteral Then
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
        End If
        i = i + 1
ContinueLoop:
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
    Dim inBacktickLiteral As Boolean

    mp_SkipWhitespace sourceText, pos, lineNo
    If pos > Len(sourceText) Then Exit Function

    startPos = pos
    i = pos
    Do While i <= Len(sourceText)
        ch = Mid$(sourceText, i, 1)

        If g_ParseHasBacktickLiterals And ch = "`" And Not inQuotes Then
            inBacktickLiteral = Not inBacktickLiteral
            i = i + 1
            GoTo ContinueLoop
        End If

        If ch = vbLf Then lineNo = lineNo + 1

        If Not inBacktickLiteral Then
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
        End If

        i = i + 1
ContinueLoop:
    Loop

    If inBacktickLiteral Then
        outErrorText = "Unterminated backtick string literal (`...`) near line " & CStr(lineNo)
        Exit Function
    End If

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
    Dim partText As String
    Dim groupedConditionText As String

    If Not mp_TrySplitConditionExpression(conditionText, condParts, condOps, resolveError) Then
        Err.Raise vbObjectError + 1594, "ex_ScriptDSL", resolveError
    End If

    For i = 1 To condParts.Count
        partText = CStr(condParts(i))
        If mp_TryUnwrapConditionGroup(partText, groupedConditionText) Then
            partResult = mp_EvaluateCondition(groupedConditionText, currentTableRef, currentRowVar, currentRowRef, tablesByRef, runtimeVars)
        Else
            If Not mp_ParseConditionPart(partText, refToken, opText, expectedValueRaw, expectedIsToken) Then
                Err.Raise vbObjectError + 1594, "ex_ScriptDSL", "Invalid condition: " & Trim$(partText)
            End If
            If Not mp_TryResolveRuntimeValue(refToken, currentTableRef, currentRowVar, currentRowRef, tablesByRef, runtimeVars, actualValue, resolveError) Then
                Err.Raise vbObjectError + 1595, "ex_ScriptDSL", resolveError
            End If
            If expectedIsToken Then
                If Not mp_TryResolveRuntimeValue(expectedValueRaw, currentTableRef, currentRowVar, currentRowRef, tablesByRef, runtimeVars, expectedValue, resolveError) Then
                    Err.Raise vbObjectError + 1595, "ex_ScriptDSL", resolveError
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
                Case ">"
                    partResult = (compareResult > 0)
                Case "<"
                    partResult = (compareResult < 0)
                Case ">="
                    partResult = (compareResult >= 0)
                Case "<="
                    partResult = (compareResult <= 0)
                Case Else
                    Err.Raise vbObjectError + 1596, "ex_ScriptDSL", "Unsupported operator in condition: " & opText
            End Select
        End If

        If Not hasCurrentTerm Then
            currentTerm = partResult
            hasCurrentTerm = True
        Else
            boolOp = LCase$(Trim$(CStr(condOps(i - 1))))
            Select Case boolOp
                Case "&&"
                    currentTerm = (currentTerm And partResult)
                Case "||"
                    If Not hasFinalResult Then
                        finalResult = currentTerm
                        hasFinalResult = True
                    Else
                        finalResult = (finalResult Or currentTerm)
                    End If
                    currentTerm = partResult
                Case Else
                    Err.Raise vbObjectError + 1596, "ex_ScriptDSL", "Unsupported boolean operator in condition: " & boolOp
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

    If Left$(LCase$(lineText), Len("callmacroobject(")) = "callmacroobject(" And Right$(lineText, 2) = ");" Then
        payload = Trim$(Mid$(lineText, Len("callMacroObject(") + 1, Len(lineText) - Len("callMacroObject(") - 2))
        If Not mp_TryParseCallMacroArgs(payload, macroName, argSpecs, outErrorText) Then
            Exit Function
        End If
        outAction("Type") = ACTION_CALL_MACRO_OBJECT
        outAction("MacroName") = macroName
        outAction.Add "Args", argSpecs
        mp_TryParseAction = True
        Exit Function
    End If

    outErrorText = "Unsupported action: '" & lineText & "'. Supported: callMacro(...), callMacroObject(...)."
End Function

Private Function mp_ParseConditionPart( _
    ByVal rawPart As String, _
    ByRef outFieldName As String, _
    ByRef outOp As String, _
    ByRef outValue As String, _
    ByRef outValueIsToken As Boolean _
) As Boolean
    Dim part As String
    Dim opPos As Long
    Dim opLen As Long
    Dim rhs As String

    part = mp_TrimDslWhitespace(rawPart)
    If mp_TryFindConditionSymbolOperator(part, "==", opPos, opLen) Then
        outOp = "=="
    ElseIf mp_TryFindConditionSymbolOperator(part, "!=", opPos, opLen) Then
        outOp = "!="
    ElseIf mp_TryFindConditionSymbolOperator(part, ">=", opPos, opLen) Then
        outOp = ">="
    ElseIf mp_TryFindConditionSymbolOperator(part, "<=", opPos, opLen) Then
        outOp = "<="
    ElseIf mp_TryFindConditionSymbolOperator(part, ">", opPos, opLen) Then
        outOp = ">"
    ElseIf mp_TryFindConditionSymbolOperator(part, "<", opPos, opLen) Then
        outOp = "<"
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

Private Function mp_TryFindConditionSymbolOperator( _
    ByVal textValue As String, _
    ByVal opSymbol As String, _
    ByRef outPos As Long, _
    ByRef outLen As Long _
) As Boolean
    Dim i As Long
    Dim inQuotes As Boolean
    Dim inBacktickLiteral As Boolean
    Dim ch As String

    outLen = Len(opSymbol)
    If outLen = 0 Then Exit Function
    If Len(textValue) < outLen Then Exit Function

    For i = 1 To Len(textValue) - outLen + 1
        ch = Mid$(textValue, i, 1)
        If ch = """" And Not mp_IsEscapedQuote(textValue, i) Then
            If Not inBacktickLiteral Then inQuotes = Not inQuotes
            GoTo ContinueLoop
        End If
        If ch = "`" Then
            If Not inQuotes Then inBacktickLiteral = Not inBacktickLiteral
            GoTo ContinueLoop
        End If

        If Not inQuotes And Not inBacktickLiteral Then
            If Mid$(textValue, i, outLen) = opSymbol Then
                outPos = i
                mp_TryFindConditionSymbolOperator = True
                Exit Function
            End If
        End If
ContinueLoop:
    Next i
End Function

Private Function mp_TryFindConditionWordOperator( _
    ByVal textValue As String, _
    ByVal opWord As String, _
    ByRef outPos As Long, _
    ByRef outLen As Long _
) As Boolean
    Dim i As Long
    Dim inQuotes As Boolean
    Dim inBacktickLiteral As Boolean
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
            If Not inBacktickLiteral Then inQuotes = Not inQuotes
            GoTo ContinueLoop
        End If
        If ch = "`" Then
            If Not inQuotes Then inBacktickLiteral = Not inBacktickLiteral
            GoTo ContinueLoop
        End If

        If Not inQuotes And Not inBacktickLiteral Then
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

    If ex_ScriptParserCore.m_IsIdentifier(rhsText) Then
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
    Dim inBacktickLiteral As Boolean
    Dim parenDepth As Long

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
            If Not inBacktickLiteral Then inQuotes = Not inQuotes
            partText = partText & ch
            i = i + 1
            GoTo ContinueLoop
        End If
        If ch = "`" Then
            If Not inQuotes Then inBacktickLiteral = Not inBacktickLiteral
            partText = partText & ch
            i = i + 1
            GoTo ContinueLoop
        End If

        If Not inQuotes And Not inBacktickLiteral Then
            If ch = "(" Then
                parenDepth = parenDepth + 1
            ElseIf ch = ")" Then
                parenDepth = parenDepth - 1
                If parenDepth < 0 Then
                    outErrorText = "Unexpected ')' in condition."
                    Exit Function
                End If
            End If

            If parenDepth = 0 And Mid$(conditionText, i, 2) = "&&" Then
                If Not mp_TryPushConditionPart(partText, "&&", outParts, outOps, outErrorText) Then Exit Function
                partText = vbNullString
                i = i + 2
                GoTo ContinueLoop
            End If

            If parenDepth = 0 And Mid$(conditionText, i, 2) = "||" Then
                If Not mp_TryPushConditionPart(partText, "||", outParts, outOps, outErrorText) Then Exit Function
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
    If inBacktickLiteral Then
        outErrorText = "Unterminated backtick string in condition."
        Exit Function
    End If
    If parenDepth <> 0 Then
        outErrorText = "Unbalanced parentheses in condition."
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

Private Function mp_TryUnwrapConditionGroup( _
    ByVal sourceText As String, _
    ByRef outInnerText As String _
) As Boolean
    Dim normalized As String

    normalized = mp_TrimDslWhitespace(sourceText)
    If Len(normalized) < 2 Then Exit Function
    If Left$(normalized, 1) <> "(" Or Right$(normalized, 1) <> ")" Then Exit Function
    If Not mp_IsConditionWrappedByOuterParens(normalized) Then Exit Function

    outInnerText = mp_TrimDslWhitespace(Mid$(normalized, 2, Len(normalized) - 2))
    If Len(outInnerText) = 0 Then Exit Function
    mp_TryUnwrapConditionGroup = True
End Function

Private Function mp_IsConditionWrappedByOuterParens(ByVal textValue As String) As Boolean
    Dim i As Long
    Dim depth As Long
    Dim ch As String
    Dim inQuotes As Boolean
    Dim inBacktickLiteral As Boolean

    textValue = mp_TrimDslWhitespace(textValue)
    If Len(textValue) < 2 Then Exit Function
    If Left$(textValue, 1) <> "(" Or Right$(textValue, 1) <> ")" Then Exit Function

    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)

        If ch = "`" And Not inQuotes Then
            inBacktickLiteral = Not inBacktickLiteral
            GoTo ContinueLoop
        End If

        If Not inBacktickLiteral Then
            If ch = """" And Not mp_IsEscapedQuote(textValue, i) Then
                inQuotes = Not inQuotes
                GoTo ContinueLoop
            End If
            If Not inQuotes Then
                If ch = "(" Then
                    depth = depth + 1
                ElseIf ch = ")" Then
                    depth = depth - 1
                    If depth < 0 Then Exit Function
                    If depth = 0 And i < Len(textValue) Then Exit Function
                End If
            End If
        End If
ContinueLoop:
    Next i

    If inQuotes Or inBacktickLiteral Then Exit Function
    mp_IsConditionWrappedByOuterParens = (depth = 0)
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
    valueText = mp_TrimOuterWhitespace(valueText)

    If mp_TryParseBacktickString(valueText, outValue) Then
        mp_TryParseQuotedString = True
        Exit Function
    End If

    If Len(valueText) < 2 Then Exit Function
    If Left$(valueText, 1) <> """" Then Exit Function
    If Right$(valueText, 1) <> """" Then Exit Function

    rawInner = Mid$(valueText, 2, Len(valueText) - 2)
    outValue = mp_DecodeEscapes(rawInner)
    mp_TryParseQuotedString = True
End Function

Private Function mp_TryParseBacktickString(ByVal valueText As String, ByRef outValue As String) As Boolean
    valueText = mp_TrimOuterWhitespace(valueText)
    If Len(valueText) < 2 Then Exit Function
    If Left$(valueText, 1) <> "`" Then Exit Function
    If Right$(valueText, 1) <> "`" Then Exit Function

    outValue = Mid$(valueText, 2, Len(valueText) - 2)
    mp_TryParseBacktickString = True
End Function

Private Function mp_TryParseDoubleQuotedString(ByVal valueText As String, ByRef outValue As String) As Boolean
    Dim rawInner As String

    valueText = mp_TrimOuterWhitespace(valueText)
    If Len(valueText) < 2 Then Exit Function
    If Left$(valueText, 1) <> """" Then Exit Function
    If Right$(valueText, 1) <> """" Then Exit Function

    rawInner = Mid$(valueText, 2, Len(valueText) - 2)
    outValue = mp_DecodeEscapes(rawInner)
    mp_TryParseDoubleQuotedString = True
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
            Err.Raise vbObjectError + 1595, "ex_ScriptDSL", resolveError
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
    Dim inBacktickLiteral As Boolean

    scriptText = Replace(scriptText, vbCrLf, vbLf)
    scriptText = Replace(scriptText, vbCr, vbLf)
    scriptText = mp_StripMultiLineComments(scriptText)
    lines = Split(scriptText, vbLf)

    For i = LBound(lines) To UBound(lines)
        rawLine = CStr(lines(i))
        rawLine = Replace(rawLine, vbTab, " ")
        rawLine = Replace(rawLine, ChrW$(160), " ")
        rawLine = mp_StripSingleLineComment(rawLine, inBacktickLiteral)
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
    Dim inBacktickLiteral As Boolean
    Dim result As String

    i = 1
    Do While i <= Len(sourceText)
        If g_ParseHasBacktickLiterals Then
            If mp_IsBacktickAt(sourceText, i) And Not inQuotes Then
                inBacktickLiteral = Not inBacktickLiteral
                result = result & "`"
                i = i + 1
                GoTo ContinueLoop
            End If
        End If

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

        If Not inBacktickLiteral Then
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
        End If

        result = result & ch
        i = i + 1
ContinueLoop:
    Loop

    mp_StripMultiLineComments = result
End Function

Private Function mp_StripSingleLineComment( _
    ByVal lineText As String, _
    ByRef inBacktickLiteral As Boolean _
) As String
    Dim i As Long
    Dim inQuotes As Boolean
    Dim ch As String
    Dim nextCh As String

    For i = 1 To Len(lineText)
        If g_ParseHasBacktickLiterals Then
            If mp_IsBacktickAt(lineText, i) And Not inQuotes Then
                inBacktickLiteral = Not inBacktickLiteral
                GoTo ContinueLoop
            End If
        End If

        ch = Mid$(lineText, i, 1)
        If Not inBacktickLiteral Then
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
        End If
ContinueLoop:
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
    mp_TryResolveConditionTokenForValidation = ex_ScriptTokenResolver.m_TryResolveTokenForValidation( _
        tokenText, _
        currentTableRef, _
        currentRowVar, _
        scopeVarTypes, _
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
    mp_TryResolveRuntimeValue = ex_ScriptTokenResolver.m_TryResolveTokenRuntime( _
        tokenText, _
        currentTableRef, _
        currentRowVar, _
        currentRowRef, _
        tablesByRef, _
        runtimeVars, _
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

    argsText = mp_TrimOuterWhitespace(argsText)
    If Len(argsText) = 0 Then
        outErrorText = "callMacro/callMacroObject requires at least macro name: callMacro(""Module.Proc"", ...)"
        Exit Function
    End If

    If Not mp_SplitArgs(argsText, parts, outErrorText) Then Exit Function
    If parts Is Nothing Or parts.Count = 0 Then
        outErrorText = "callMacro/callMacroObject requires at least macro name: callMacro(""Module.Proc"", ...)"
        Exit Function
    End If

    If Not mp_TryParseQuotedString(CStr(parts(1)), outMacroName) Then
        outErrorText = "callMacro/callMacroObject first argument must be quoted macro name."
        Exit Function
    End If
    outMacroName = Trim$(outMacroName)
    If Len(outMacroName) = 0 Then
        outErrorText = "callMacro/callMacroObject macro name cannot be empty."
        Exit Function
    End If

    Set outArgSpecs = New Collection
    For i = 2 To parts.Count
        partText = mp_TrimOuterWhitespace(CStr(parts(i)))
        If Len(partText) = 0 Then
            outErrorText = "callMacro/callMacroObject argument #" & CStr(i - 1) & " is empty."
            Exit Function
        End If
        If Not mp_TryParseMacroArg(partText, argSpec) Then
            outErrorText = "Unsupported macro argument '" & partText & "'. Use variable, numeric literal, string literal (""..."" or `...`), template string ($`...`), Source.Sheet[Table].row[N], Source.Sheet[Table].lastRow, Source.Sheet[Table].prevRow, or a .column[Field] variant."
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
    Dim rowCellResolveError As String
    Dim scopeValue As obj_ScriptScopeValue
    Dim scopeResolveError As String

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
            Case "number"
                result.Add CLng(argSpec("Value"))

            Case "varref"
                scopeResolveError = vbNullString
                If Not ex_ScriptScopeValue.m_TryGetScopeValue(runtimeVars, CStr(argSpec("Name")), scopeValue, scopeResolveError) Then
                    Err.Raise vbObjectError + 1601, "ex_ScriptDSL", "Variable '" & CStr(argSpec("Name")) & "' is not available for callMacro argument: " & scopeResolveError
                End If
                If scopeValue.HasObjectValue Then
                    Set argObject = scopeValue.ObjectValue
                    result.Add argObject
                Else
                    argValue = scopeValue.TextValue
                    result.Add argValue
                End If
            Case "rowref"
                result.Add ex_ResultRuntimeAdapter.m_ResolveRowReferenceArg(argSpec, tablesByRef)
            Case "cellref"
                result.Add ex_ResultRuntimeAdapter.m_ResolveCellReferenceArg(argSpec, tablesByRef)
            Case "scopecellref"
                If Not mp_TryResolveScopedRowCellValue( _
                    runtimeVars, _
                    CStr(argSpec("RowVar")), _
                    CStr(argSpec("FieldAlias")), _
                    CStr(argSpec("RowVar")) & ".column[" & CStr(argSpec("FieldAlias")) & "]", _
                    renderedText, _
                    rowCellResolveError _
                ) Then
                    Err.Raise vbObjectError + 1626, "ex_ScriptDSL", _
                        "Unable to resolve callMacro row-cell argument '" & _
                        CStr(argSpec("RowVar")) & ".column[" & CStr(argSpec("FieldAlias")) & "]': " & _
                        rowCellResolveError
                End If
                result.Add renderedText
            Case "string"
                result.Add CStr(argSpec("Value"))
            Case "template"
                renderedText = mp_RenderTemplate(CStr(argSpec("Value")), currentTableRef, currentRowVar, currentRowRef, tablesByRef, runtimeVars)
                result.Add renderedText
            Case Else
                Err.Raise vbObjectError + 1598, "ex_ScriptDSL", "Unsupported callMacro argument kind: " & argKind
        End Select
    Next i

    Set mp_BuildMacroRuntimeArgs = result
End Function

Private Function mp_CreateVarScope() As Object
    Set mp_CreateVarScope = CreateObject("Scripting.Dictionary")
    mp_CreateVarScope.CompareMode = 1
End Function

Private Sub mp_DebugLog(ByVal messageText As String)
    If Not DEBUG_LOG_ENABLED Then Exit Sub
    On Error Resume Next
    ' Comment next line to disable file logger quickly.
    ex_Messaging.m_LogToFile "[ex_ScriptDSL] " & CStr(messageText), DEBUG_LOG_PATH
    On Error GoTo 0
End Sub

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

Private Sub mp_SetRuntimeScopeValue(ByVal targetScope As Object, ByVal variableName As String, ByVal scopeValue As obj_ScriptScopeValue)
    variableName = Trim$(variableName)
    If Len(variableName) = 0 Then Exit Sub
    If targetScope Is Nothing Then Exit Sub
    If scopeValue Is Nothing Then Exit Sub

    Set targetScope(variableName) = scopeValue
End Sub

Private Sub mp_SetRuntimeScopeString(ByVal targetScope As Object, ByVal variableName As String, ByVal variableValue As String)
    Dim scopeValue As obj_ScriptScopeValue

    Set scopeValue = ex_ScriptScopeValue.m_CreateStringValue(variableValue)
    mp_SetRuntimeScopeValue targetScope, variableName, scopeValue
End Sub

Private Sub mp_SetRuntimeScopeObject(ByVal targetScope As Object, ByVal variableName As String, ByVal variableObject As Object)
    Dim scopeValue As obj_ScriptScopeValue

    Set scopeValue = ex_ScriptScopeValue.m_CreateObjectValue(variableObject)
    mp_SetRuntimeScopeValue targetScope, variableName, scopeValue
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
    Dim rowObj As obj_ResultRow

    If TypeOf variableObject Is obj_ResultRow Then
        Set rowObj = variableObject
        If Not rowObj.HasAlias(memberName) Then
            outErrorText = "Unknown row field alias '" & memberName & "' in token '" & tokenText & "'."
            Exit Function
        End If

        outValue = rowObj.Column(memberName)
        mp_TryResolveVariableMemberValue = True
        Exit Function
    End If

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

Private Function mp_ConvertVariantToString(ByVal valueRef As Variant) As String
    If IsNull(valueRef) Then
        mp_ConvertVariantToString = vbNullString
    ElseIf IsError(valueRef) Then
        Err.Raise vbObjectError + 1612, "ex_ScriptDSL", "callMacro returned error value; expected string-compatible result."
    ElseIf IsObject(valueRef) Then
        Err.Raise vbObjectError + 1613, "ex_ScriptDSL", "callMacro returned object; expected string-compatible result."
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
            Err.Raise vbObjectError + 1625, "ex_ScriptDSL", "String expression contains invalid operand."
        End If
        partKind = LCase$(Trim$(CStr(partSpec("Kind"))))

        Select Case partKind
            Case EXPR_PART_LITERAL
                resultText = resultText & CStr(partSpec("Value"))
            Case EXPR_PART_TOKEN
                tokenText = CStr(partSpec("Value"))
                If Not mp_TryResolveRuntimeValue(tokenText, currentTableRef, currentRowVar, currentRowRef, tablesByRef, runtimeVars, tokenValue, resolveError) Then
                    Err.Raise vbObjectError + 1622, "ex_ScriptDSL", "Unable to resolve string expression token '" & tokenText & "': " & resolveError
                End If
                resultText = resultText & tokenValue
            Case Else
                Err.Raise vbObjectError + 1623, "ex_ScriptDSL", "Unsupported string expression operand kind: " & partKind
        End Select
    Next i

    mp_EvaluateStringExpression = resultText
End Function

Private Function mp_InferLetVarType( _
    ByVal statement As Object, _
    ByVal scopeVarTypes As Object _
) As String
    Dim macroName As String
    Dim assignKind As String
    Dim args As Collection
    Dim firstArg As Object
    Dim varName As String

    mp_InferLetVarType = VAR_TYPE_STRING
    If statement Is Nothing Then Exit Function

    assignKind = mp_GetStatementAssignKind(statement)
    If assignKind = ASSIGN_KIND_CALL_MACRO Then
        mp_InferLetVarType = VAR_TYPE_STRING
        Exit Function
    End If

    If assignKind = ASSIGN_KIND_CALL_MACRO_OBJECT Then
        mp_InferLetVarType = VAR_TYPE_OBJECT
    Else
        Exit Function
    End If

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
    Dim inBacktickLiteral As Boolean

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
        If g_ParseHasBacktickLiterals And ch = "`" And Not inQuotes Then
            inBacktickLiteral = Not inBacktickLiteral
            GoTo ContinueLoop
        End If

        If Not inBacktickLiteral Then
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
        End If
ContinueLoop:
    Next i

    If inBacktickLiteral Then
        outErrorText = "Unterminated backtick string literal (`...`) in assignment statement."
        Exit Function
    End If

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
    Dim inBacktickLiteral As Boolean

    Set outParts = New Collection
    partText = vbNullString

    For i = 1 To Len(argsText)
        ch = Mid$(argsText, i, 1)
        If g_ParseHasBacktickLiterals And ch = "`" And Not inQuotes Then
            inBacktickLiteral = Not inBacktickLiteral
            partText = partText & ch
            GoTo ContinueLoop
        End If

        If Not inBacktickLiteral Then
            If ch = """" And Not mp_IsEscapedQuote(argsText, i) Then
                inQuotes = Not inQuotes
                partText = partText & ch
                GoTo ContinueLoop
            End If

            If ch = "," And Not inQuotes Then
                outParts.Add mp_TrimOuterWhitespace(partText)
                partText = vbNullString
                GoTo ContinueLoop
            End If
        End If

        partText = partText & ch
ContinueLoop:
    Next i

    If inQuotes Then
        outErrorText = "Unterminated quoted string in callMacro arguments."
        Exit Function
    End If
    If inBacktickLiteral Then
        outErrorText = "Unterminated backtick string literal (`...`) in callMacro arguments."
        Exit Function
    End If

    outParts.Add mp_TrimOuterWhitespace(partText)
    mp_SplitArgs = True
End Function

Private Function mp_IsEscapedQuote(ByVal textValue As String, ByVal pos As Long) As Boolean
    If pos <= 1 Then Exit Function
    mp_IsEscapedQuote = (Mid$(textValue, pos, 1) = """" And Mid$(textValue, pos - 1, 1) = "\")
End Function

Private Function mp_IsBacktickAt(ByVal textValue As String, ByVal pos As Long) As Boolean
    If pos < 1 Then Exit Function
    If pos > Len(textValue) Then Exit Function
    mp_IsBacktickAt = (Mid$(textValue, pos, 1) = "`")
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
