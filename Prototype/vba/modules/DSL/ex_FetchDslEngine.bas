Attribute VB_Name = "ex_FetchDslEngine"
Option Explicit

Private Const DSL_VIRTUAL_KIND_MARKER As String = "virtual"

' Supported Fetch DSL keywords/tokens
Private Const DSL_KW_DEFINE As String = "DEFINE"
Private Const DSL_KW_AS As String = "AS"
Private Const DSL_KW_DATABASE As String = "DATABASE"
Private Const DSL_KW_FIND As String = "FIND"
Private Const DSL_KW_SEARCH_ABOVE As String = "SEARCH_ABOVE"
Private Const DSL_KW_SEARCH_BELOW As String = "SEARCH_BELOW"
Private Const DSL_KW_IF As String = "@IF"
Private Const DSL_KW_PUSH As String = "@PUSH"
Private Const DSL_KW_KEEP As String = "@KEEP"
Private Const DSL_KW_GENERATE As String = "GENERATE"
Private Const DSL_KW_COLUMNS As String = "COLUMNS"
Private Const DSL_KW_ROW As String = "ROW"
Private Const DSL_KW_FOREACH As String = "@FOREACH"
Private Const DSL_KW_LAST As String = "@LAST"
Private Const DSL_KW_AND As String = "AND"
Private Const DSL_KW_OR As String = "OR"
Private Const DSL_KW_NOT As String = "NOT"
Private Const DSL_FN_IS_EMPTY As String = "@ISEMPTY"
Private Const DSL_FN_IS_NUMBER As String = "@ISNUMBER"
Private Const DSL_OP_EQ As String = "=="
Private Const DSL_KEEP_MODE_ALL As String = "ALL"
Private Const DSL_KEEP_MODE_LAST As String = "LAST"
Private Const DSL_KEEP_MODE_FIRST As String = "FIRST"
Private Const DSL_KEEP_MODE_CONTIGUOUS As String = "CONTIGUOUS"
Private Const DSL_VAR_CURRENT_TABLE As String = "$CurrentTable"
Private Const DSL_VAR_KEY As String = "$Key"
Private Const DSL_COMMENT_LINE As String = "//"
Private Const DSL_COMMENT_BLOCK_REGEX As String = "/\*[\s\S]*?\*/"

Private g_FetchDslPlanByTable As Object

Public Sub m_ResetPlanCache()
    Set g_FetchDslPlanByTable = Nothing
End Sub

Private Sub mp_EnsureFetchDslPlanCacheContainer()
    If g_FetchDslPlanByTable Is Nothing Then
        Set g_FetchDslPlanByTable = CreateObject("Scripting.Dictionary")
        g_FetchDslPlanByTable.CompareMode = 1
    End If
End Sub

Public Function m_GetDslText(ByVal cfg As Object, ByVal sourceAlias As String, ByVal tableAlias As String) As String
    m_GetDslText = mp_GetFetchDslText(cfg, sourceAlias, tableAlias)
End Function

Public Function m_HasDslConfig(ByVal cfg As Object, ByVal sourceAlias As String, ByVal tableAlias As String) As Boolean
    m_HasDslConfig = (Len(mp_GetFetchDslText(cfg, sourceAlias, tableAlias)) > 0)
End Function

Public Function m_GetGeneratedKindValue() As String
    m_GetGeneratedKindValue = DSL_VIRTUAL_KIND_MARKER
End Function

Public Function m_TryGetPlan(ByVal cfg As Object, ByVal sourceAlias As String, ByVal tableAlias As String, ByRef outPlan As Object) As Boolean
    m_TryGetPlan = mp_TryGetFetchDslPlan(cfg, sourceAlias, tableAlias, outPlan)
End Function

Public Function m_GetVirtualColumns(ByVal cfg As Object, ByVal sourceAlias As String, ByVal tableAlias As String) As Variant
    Dim plan As Object
    Dim columns As Collection
    Dim arr() As String
    Dim i As Long

    If Not mp_TryGetFetchDslPlan(cfg, sourceAlias, tableAlias, plan) Then
        m_GetVirtualColumns = Array()
        Exit Function
    End If

    Set columns = plan("GenerateColumns")
    If columns Is Nothing Then
        m_GetVirtualColumns = Array()
        Exit Function
    End If
    If columns.Count = 0 Then
        m_GetVirtualColumns = Array()
        Exit Function
    End If

    ReDim arr(0 To columns.Count - 1)
    For i = 1 To columns.Count
        arr(i - 1) = CStr(columns(i))
    Next i
    m_GetVirtualColumns = arr
End Function

Public Function m_IsVirtualFieldAlias(ByVal cfg As Object, ByVal sourceAlias As String, ByVal tableAlias As String, ByVal fieldAlias As String) As Boolean
    Dim dslVirtuals As Variant

    fieldAlias = Trim$(fieldAlias)
    dslVirtuals = m_GetVirtualColumns(cfg, sourceAlias, tableAlias)
    If mp_ArrayContainsText(dslVirtuals, fieldAlias) Then
        m_IsVirtualFieldAlias = True
    End If
End Function

Public Function m_ApplyFetchRowsFromSource( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal adoConn As Object, _
    ByVal tableRef As String, _
    ByVal keyHeader As String, _
    ByVal keyValue As String, _
    ByVal fields As Variant, _
    ByRef ioOutValues() As Variant, _
    ByRef ioRowCount As Long, _
    ByVal fieldCount As Long, _
    Optional ByRef outKindsByOutRow As Object = Nothing _
) As Boolean
    Dim dslPlan As Object

    If Not mp_TryGetFetchDslPlan(cfg, sourceAlias, tableAlias, dslPlan) Then Exit Function
    m_ApplyFetchRowsFromSource = mp_ApplyFetchDslRowsFromSource(cfg, sourceAlias, tableAlias, adoConn, tableRef, keyHeader, keyValue, fields, ioOutValues, ioRowCount, fieldCount, dslPlan, outKindsByOutRow)
End Function
Private Sub mp_AppendRowsToOutputMatrix( _
    ByRef ioBaseValues() As Variant, _
    ByRef ioBaseRowCount As Long, _
    ByVal appendValues As Variant, _
    ByVal appendRowCount As Long, _
    ByVal fieldCount As Long _
)
    Dim merged() As Variant
    Dim r As Long
    Dim c As Long
    Dim totalRows As Long

    If appendRowCount <= 0 Then Exit Sub
    If fieldCount <= 0 Then Exit Sub

    totalRows = ioBaseRowCount + appendRowCount
    ReDim merged(1 To totalRows, 1 To fieldCount)

    For r = 1 To ioBaseRowCount
        For c = 1 To fieldCount
            merged(r, c) = ioBaseValues(r, c)
        Next c
    Next r

    For r = 1 To appendRowCount
        For c = 1 To fieldCount
            merged(ioBaseRowCount + r, c) = appendValues(r, c)
        Next c
    Next r

    ReDim ioBaseValues(1 To totalRows, 1 To fieldCount)
    For r = 1 To totalRows
        For c = 1 To fieldCount
            ioBaseValues(r, c) = merged(r, c)
        Next c
    Next r
    ioBaseRowCount = totalRows
End Sub

Private Function mp_TryGetFetchDslPlan(ByVal cfg As Object, ByVal sourceAlias As String, ByVal tableAlias As String, ByRef outPlan As Object) As Boolean
    Dim dslText As String
    Dim cacheKey As String
    Dim cached As Object
    Dim parsed As Object

    dslText = mp_GetFetchDslText(cfg, sourceAlias, tableAlias)
    If Len(dslText) = 0 Then Exit Function

    mp_EnsureFetchDslPlanCacheContainer
    cacheKey = LCase$(Trim$(sourceAlias) & "|" & Trim$(tableAlias))

    If Not g_FetchDslPlanByTable Is Nothing Then
        If g_FetchDslPlanByTable.Exists(cacheKey) Then
            Set cached = g_FetchDslPlanByTable(cacheKey)
            If Not cached Is Nothing Then
                If StrComp(CStr(cached("__dslText")), dslText, vbBinaryCompare) = 0 Then
                    Set outPlan = cached
                    mp_TryGetFetchDslPlan = True
                    Exit Function
                End If
            End If
        End If
    End If

    Set parsed = mp_FetchDslParsePlan(dslText)
    parsed("__dslText") = dslText

    If g_FetchDslPlanByTable Is Nothing Then
        mp_EnsureFetchDslPlanCacheContainer
    End If
    Set g_FetchDslPlanByTable(cacheKey) = parsed
    Set outPlan = parsed
    mp_TryGetFetchDslPlan = True
End Function

Private Function mp_ApplyFetchDslRowsFromSource( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal adoConn As Object, _
    ByVal tableRef As String, _
    ByVal keyHeader As String, _
    ByVal keyValue As String, _
    ByVal fields As Variant, _
    ByRef ioOutValues() As Variant, _
    ByRef ioRowCount As Long, _
    ByVal fieldCount As Long, _
    ByVal dslPlan As Object, _
    Optional ByRef outKindsByOutRow As Object = Nothing _
) As Boolean
    Dim rs As Object
    Dim sql As String
    Dim rowsData As Variant
    Dim rowLower As Long
    Dim rowUpper As Long
    Dim keyRows As Collection
    Dim keyRow As Variant
    Dim ownerSeq As Long
    Dim runtimeCtx As Object
    Dim nextKeyRow As Long
    Dim fieldOrdByToken As Object
    Dim ownerValues As Object
    Dim metaRows As Collection
    Dim rowValues As Object
    Dim generatedRows As Collection
    Dim fieldIndexMap As Object
    Dim fieldTypeByToken As Object
    Dim appendValues() As Variant
    Dim appendCount As Long
    Dim baseRowCount As Long
    Dim r As Long
    Dim c As Long
    Dim fieldAlias As Variant
    Dim rowIndex As Long
    Dim assignValue As String
    Dim kindsByOutRow As Object

    On Error GoTo EH

    If Len(Trim$(tableRef)) = 0 Then Exit Function
    If ioRowCount <= 0 Then Exit Function
    If fieldCount <= 0 Then Exit Function
    If mp_IsEmptyVariantArray(fields) Then Exit Function
    If dslPlan Is Nothing Then Exit Function

    ' Use full projection to avoid provider-specific parser issues on localized
    ' headers with punctuation (e.g. "Вх. Дата"/"Вх. №").
    sql = "SELECT * FROM " & tableRef
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, adoConn, 0, 1
    If rs.EOF Then
        rs.Close
        Exit Function
    End If

    Set fieldOrdByToken = mp_FetchDslBuildSourceFieldOrdinals(cfg, sourceAlias, tableAlias, dslPlan, rs)
    Set fieldTypeByToken = mp_FetchDslBuildSourceFieldTypes(fieldOrdByToken, rs)
    Set dslPlan("__FieldOrdByToken") = fieldOrdByToken
    Set dslPlan("__FieldTypeByToken") = fieldTypeByToken
    rowsData = rs.GetRows
    rs.Close
    Set rs = Nothing

    rowLower = LBound(rowsData, 2)
    rowUpper = UBound(rowsData, 2)
    Set keyRows = mp_FetchDslFindKeyRows(rowsData, dslPlan, keyValue, fieldOrdByToken)
    If keyRows Is Nothing Then Exit Function
    If keyRows.Count = 0 Then Exit Function

    Set generatedRows = New Collection
    Set fieldIndexMap = mp_FetchDslBuildFieldIndexMap(fields)
    Set kindsByOutRow = mp_FetchDslCreateDictionary()
    ownerSeq = 0

    For Each keyRow In keyRows
        ownerSeq = ownerSeq + 1

        Set runtimeCtx = mp_FetchDslInitRuntimeContext(dslPlan)
        mp_FetchDslRunSearchRule dslPlan("AboveRule"), rowsData, rowLower, CLng(keyRow) - 1, runtimeCtx, dslPlan, keyValue, fieldOrdByToken, CLng(keyRow), Nothing

        nextKeyRow = mp_FetchDslFindNextKeyRow(rowsData, dslPlan, keyValue, fieldOrdByToken, CLng(keyRow), rowUpper)
        mp_FetchDslRunSearchRule dslPlan("BelowRule"), rowsData, CLng(keyRow) + 1, nextKeyRow - 1, runtimeCtx, dslPlan, keyValue, fieldOrdByToken, CLng(keyRow), Nothing

        Set ownerValues = mp_FetchDslEvaluateAssignments(dslPlan("GenerateKeyAssignments"), dslPlan, rowsData, CLng(keyRow), keyValue, runtimeCtx, Nothing)
        If ownerSeq <= ioRowCount Then
            For Each fieldAlias In ownerValues.Keys
                If fieldIndexMap.Exists(LCase$(Trim$(CStr(fieldAlias)))) Then
                    rowIndex = CLng(fieldIndexMap(LCase$(Trim$(CStr(fieldAlias)))))
                    ioOutValues(ownerSeq, rowIndex) = ownerValues(fieldAlias)
                End If
            Next fieldAlias
        End If

        Set metaRows = mp_FetchDslBuildGeneratedMetaRows(dslPlan, rowsData, CLng(keyRow), keyValue, runtimeCtx)
        If Not metaRows Is Nothing Then
            For Each rowValues In metaRows
                generatedRows.Add rowValues
            Next rowValues
        End If
    Next keyRow

    appendCount = generatedRows.Count
    If appendCount <= 0 Then
        If kindsByOutRow.Count > 0 Then Set outKindsByOutRow = kindsByOutRow
        mp_ApplyFetchDslRowsFromSource = True
        Exit Function
    End If

    baseRowCount = ioRowCount
    ReDim appendValues(1 To appendCount, 1 To fieldCount)
    For r = 1 To appendCount
        Set rowValues = generatedRows(r)
        kindsByOutRow(CStr(baseRowCount + r)) = DSL_VIRTUAL_KIND_MARKER
        For c = 1 To fieldCount
            fieldAlias = Trim$(CStr(fields(LBound(fields) + c - 1)))
            If fieldIndexMap.Exists(LCase$(fieldAlias)) Then
                If rowValues.Exists(fieldAlias) Then
                    assignValue = CStr(rowValues(fieldAlias))
                ElseIf rowValues.Exists(LCase$(fieldAlias)) Then
                    assignValue = CStr(rowValues(LCase$(fieldAlias)))
                Else
                    assignValue = vbNullString
                End If
            Else
                assignValue = vbNullString
            End If
            appendValues(r, c) = assignValue
        Next c
    Next r

    mp_AppendRowsToOutputMatrix ioOutValues, ioRowCount, appendValues, appendCount, fieldCount
    If kindsByOutRow.Count > 0 Then Set outKindsByOutRow = kindsByOutRow
    mp_ApplyFetchDslRowsFromSource = True
    Exit Function

EH:
    Dim innerErrDescription As String
    innerErrDescription = Err.Description
    innerErrDescription = mp_FetchDslLocalizeInnerErrorRu(innerErrDescription)

    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    On Error GoTo 0

    Err.Raise vbObjectError + 1768, "ex_FetchDslEngine", _
        "Ошибка Fetch DSL для '" & sourceAlias & ".Sheet[" & tableAlias & "]'. SQL=[" & sql & "]: " & innerErrDescription
End Function

Private Function mp_FetchDslParsePlan(ByVal dslTextRaw As String) As Object
    Dim dslText As String
    Dim plan As Object
    Dim contexts As Object
    Dim contextAlias As String
    Dim sourceAlias As String
    Dim sourceInputVar As String
    Dim keyField As String
    Dim keyVar As String
    Dim findBody As String
    Dim generateBody As String
    Dim foreachVar As String
    Dim foreachCtx As String
    Dim keyAssignmentsText As String
    Dim foreachAssignmentsText As String

    dslText = mp_FetchDslStripComments(dslTextRaw)
    Set plan = mp_FetchDslCreateDictionary()

    mp_FetchDslParseDefineSource dslText, sourceInputVar, sourceAlias
    plan("SourceInputVar") = sourceInputVar
    plan("SourceAlias") = sourceAlias

    Set contexts = mp_FetchDslParseContexts(dslText, contextAlias)
    plan("ContextAlias") = contextAlias
    Set plan("Contexts") = contexts

    mp_FetchDslParseFindHeader dslText, sourceAlias, keyField, keyVar, findBody
    plan("KeyField") = keyField
    plan("KeyVar") = keyVar

    Set plan("AboveRule") = mp_FetchDslParseSearchRule(findBody, DSL_KW_SEARCH_ABOVE, sourceAlias, contextAlias, contexts)
    Set plan("BelowRule") = mp_FetchDslParseSearchRule(findBody, DSL_KW_SEARCH_BELOW, sourceAlias, contextAlias, contexts)

    generateBody = mp_FetchDslExtractGenerateBody(findBody)
    Set plan("GenerateColumns") = mp_FetchDslParseColumns(generateBody)

    keyAssignmentsText = mp_FetchDslExtractFirstRowAssignments(generateBody)
    Set plan("GenerateKeyAssignments") = mp_FetchDslParseRowAssignments(keyAssignmentsText, sourceAlias, vbNullString, contextAlias)

    mp_FetchDslExtractForeach generateBody, contextAlias, foreachVar, foreachCtx, foreachAssignmentsText
    plan("GenerateForeachVar") = foreachVar
    plan("GenerateForeachCtx") = foreachCtx
    Set plan("GenerateForeachAssignments") = mp_FetchDslParseRowAssignments(foreachAssignmentsText, sourceAlias, foreachVar, contextAlias)

    If Not contexts.Exists(foreachCtx) Then
        Err.Raise vbObjectError + 1769, "ex_FetchDslEngine", "Fetch DSL: unknown FOREACH context Ctx." & foreachCtx
    End If

    Set mp_FetchDslParsePlan = plan
End Function

Private Function mp_FetchDslStripComments(ByVal textIn As String) As String
    Dim normalized As String
    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim p As Long
    Dim resultText As String
    Dim rgxBlock As Object

    normalized = Replace(textIn, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)

    Set rgxBlock = CreateObject("VBScript.RegExp")
    rgxBlock.Pattern = DSL_COMMENT_BLOCK_REGEX
    rgxBlock.Global = True
    rgxBlock.MultiLine = True
    rgxBlock.IgnoreCase = False
    normalized = rgxBlock.Replace(normalized, vbLf)

    lines = Split(normalized, vbLf)

    For i = LBound(lines) To UBound(lines)
        lineText = CStr(lines(i))
        p = InStr(1, lineText, DSL_COMMENT_LINE, vbBinaryCompare)
        If p > 0 Then lineText = Left$(lineText, p - 1)
        lineText = Trim$(lineText)
        If Len(lineText) > 0 Then
            If Len(resultText) > 0 Then resultText = resultText & vbLf
            resultText = resultText & lineText
        End If
    Next i

    mp_FetchDslStripComments = resultText
End Function

Private Function mp_FetchDslCreateRegex(ByVal pattern As String, Optional ByVal isGlobal As Boolean = False) As Object
    Dim rgx As Object
    Set rgx = CreateObject("VBScript.RegExp")
    rgx.Pattern = pattern
    rgx.Global = isGlobal
    rgx.MultiLine = True
    rgx.IgnoreCase = True
    Set mp_FetchDslCreateRegex = rgx
End Function

Private Function mp_FetchDslMatchOne(ByVal textIn As String, ByVal pattern As String) As Object
    Dim rgx As Object
    Dim matches As Object

    Set rgx = mp_FetchDslCreateRegex(pattern, False)
    Set matches = rgx.Execute(textIn)
    If matches.Count > 0 Then Set mp_FetchDslMatchOne = matches(0)
End Function

Private Function mp_FetchDslMatchAll(ByVal textIn As String, ByVal pattern As String) As Object
    Dim rgx As Object
    Set rgx = mp_FetchDslCreateRegex(pattern, True)
    Set mp_FetchDslMatchAll = rgx.Execute(textIn)
End Function

Private Function mp_FetchDslCreateDictionary() As Object
    Set mp_FetchDslCreateDictionary = CreateObject("Scripting.Dictionary")
    mp_FetchDslCreateDictionary.CompareMode = 1
End Function

Private Sub mp_FetchDslParseDefineSource(ByVal dslText As String, ByRef outInputVar As String, ByRef outSourceAlias As String)
    Dim m As Object
    Set m = mp_FetchDslMatchOne(dslText, DSL_KW_DEFINE & "\s+(\$[A-Za-z_][A-Za-z0-9_]*)\s+" & DSL_KW_AS & "\s+([A-Za-z_][A-Za-z0-9_]*)\s*;")
    If m Is Nothing Then
        Err.Raise vbObjectError + 1770, "ex_FetchDslEngine", "Fetch DSL: expected '" & DSL_KW_DEFINE & " " & DSL_VAR_CURRENT_TABLE & " " & DSL_KW_AS & " Src;'."
    End If

    outInputVar = Trim$(CStr(m.SubMatches(0)))
    outSourceAlias = Trim$(CStr(m.SubMatches(1)))
End Sub

Private Function mp_FetchDslParseContexts(ByVal dslText As String, ByRef outContextAlias As String) As Object
    Dim m As Object
    Dim body As String
    Dim entries As Object
    Dim entry As Object
    Dim contexts As Object
    Dim ctxDef As Object
    Dim fieldsSet As Object
    Dim fieldList As Variant
    Dim i As Long
    Dim fieldName As String

    Set m = mp_FetchDslMatchOne(dslText, DSL_KW_DEFINE & "\s+" & DSL_KW_DATABASE & "\s+([A-Za-z_][A-Za-z0-9_]*)\s*\{([\s\S]*?)\}\s*;")
    If m Is Nothing Then
        Err.Raise vbObjectError + 1771, "ex_FetchDslEngine", "Fetch DSL: expected '" & DSL_KW_DEFINE & " " & DSL_KW_DATABASE & " ... { ... };'."
    End If

    outContextAlias = Trim$(CStr(m.SubMatches(0)))
    body = CStr(m.SubMatches(1))

    Set contexts = mp_FetchDslCreateDictionary()
    Set entries = mp_FetchDslMatchAll(body, "([A-Za-z_][A-Za-z0-9_]*)\s*:\s*\[([\s\S]*?)\]\s*,?")
    If entries Is Nothing Or entries.Count = 0 Then
        Err.Raise vbObjectError + 1772, "ex_FetchDslEngine", "Fetch DSL: no context collections declared in " & DSL_KW_DATABASE & " block."
    End If

    For Each entry In entries
        Set ctxDef = mp_FetchDslCreateDictionary()
        Set fieldsSet = mp_FetchDslCreateDictionary()
        fieldList = mp_FetchDslSplitCsv(CStr(entry.SubMatches(1)))
        If mp_IsEmptyVariantArray(fieldList) Then
            Err.Raise vbObjectError + 1773, "ex_FetchDslEngine", "Fetch DSL: context '" & CStr(entry.SubMatches(0)) & "' has empty field list."
        End If
        For i = LBound(fieldList) To UBound(fieldList)
            fieldName = mp_FetchDslTrimToken(CStr(fieldList(i)))
            If Len(fieldName) > 0 Then fieldsSet(fieldName) = True
        Next i
        Set ctxDef("Fields") = fieldsSet
        Set contexts(Trim$(CStr(entry.SubMatches(0)))) = ctxDef
    Next entry

    Set mp_FetchDslParseContexts = contexts
End Function

Private Sub mp_FetchDslParseFindHeader(ByVal dslText As String, ByVal sourceAlias As String, ByRef outKeyField As String, ByRef outKeyVar As String, ByRef outFindBody As String)
    Dim m As Object
    Dim lhsAlias As String

    Set m = mp_FetchDslMatchOne(dslText, DSL_KW_FIND & "\s*\[\s*([A-Za-z_][A-Za-z0-9_]*)\.([A-Za-z_][A-Za-z0-9_]*)\s*" & DSL_OP_EQ & "\s*(\$[A-Za-z_][A-Za-z0-9_]*)\s*\]\s*\{([\s\S]*)\}\s*$")
    If m Is Nothing Then
        Err.Raise vbObjectError + 1774, "ex_FetchDslEngine", "Fetch DSL: expected '" & DSL_KW_FIND & "[Src.FIO " & DSL_OP_EQ & " " & DSL_VAR_KEY & "] { ... }'."
    End If

    lhsAlias = Trim$(CStr(m.SubMatches(0)))
    If StrComp(lhsAlias, sourceAlias, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 1775, "ex_FetchDslEngine", "Fetch DSL: " & DSL_KW_FIND & " condition must use source alias '" & sourceAlias & "'."
    End If

    outKeyField = Trim$(CStr(m.SubMatches(1)))
    outKeyVar = Trim$(CStr(m.SubMatches(2)))
    outFindBody = CStr(m.SubMatches(3))
End Sub

Private Function mp_FetchDslParseSearchRule( _
    ByVal findBody As String, _
    ByVal blockName As String, _
    ByVal sourceAlias As String, _
    ByVal contextAlias As String, _
    ByVal contexts As Object _
) As Object
    Dim pattern As String
    Dim m As Object
    Dim rule As Object
    Dim condText As String
    Dim pushText As String
    Dim keepMode As String
    Dim pushCtxAlias As String
    Dim pushCtxName As String
    Dim assigns As Collection

    pattern = blockName & "\s*" & DSL_KW_IF & "\(\s*([\s\S]*?)\s*\)\s*" & DSL_KW_PUSH & "\(\s*([A-Za-z_][A-Za-z0-9_]*)\.([A-Za-z_][A-Za-z0-9_]*)\s*\)\s*\{([\s\S]*?)\}\s*" & DSL_KW_KEEP & "\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*\)\s*;"
    Set m = mp_FetchDslMatchOne(findBody, pattern)
    If m Is Nothing Then
        Err.Raise vbObjectError + 1776, "ex_FetchDslEngine", "Fetch DSL: block '" & blockName & "' is invalid or missing."
    End If

    condText = Trim$(CStr(m.SubMatches(0)))
    pushCtxAlias = Trim$(CStr(m.SubMatches(1)))
    pushCtxName = Trim$(CStr(m.SubMatches(2)))
    pushText = CStr(m.SubMatches(3))
    keepMode = UCase$(Trim$(CStr(m.SubMatches(4))))

    If StrComp(pushCtxAlias, contextAlias, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 1777, "ex_FetchDslEngine", "Fetch DSL: " & blockName & " must push into " & contextAlias & ".<name>."
    End If
    If Not contexts.Exists(pushCtxName) Then
        Err.Raise vbObjectError + 1778, "ex_FetchDslEngine", "Fetch DSL: unknown context '" & contextAlias & "." & pushCtxName & "' in " & blockName & "."
    End If

    Set assigns = mp_FetchDslParsePushAssignments(pushText, sourceAlias, contextAlias, contexts(pushCtxName)("Fields"), pushCtxName)

    Set rule = mp_FetchDslCreateDictionary()
    Set rule("Condition") = mp_FetchDslParseCondition(condText, sourceAlias, contextAlias)
    rule("CtxName") = pushCtxName
    rule("KeepMode") = keepMode
    Set rule("PushAssignments") = assigns

    If keepMode <> DSL_KEEP_MODE_ALL And keepMode <> DSL_KEEP_MODE_LAST And keepMode <> DSL_KEEP_MODE_FIRST And keepMode <> DSL_KEEP_MODE_CONTIGUOUS Then
        Err.Raise vbObjectError + 1779, "ex_FetchDslEngine", "Fetch DSL: unsupported " & DSL_KW_KEEP & " mode '" & keepMode & "' in " & blockName & "."
    End If

    Set mp_FetchDslParseSearchRule = rule
End Function

Private Function mp_FetchDslExtractGenerateBody(ByVal findBody As String) As String
    Dim m As Object
    Set m = mp_FetchDslMatchOne(findBody, DSL_KW_GENERATE & "\s*\{([\s\S]*)\}\s*$")
    If m Is Nothing Then
        Err.Raise vbObjectError + 1780, "ex_FetchDslEngine", "Fetch DSL: missing " & DSL_KW_GENERATE & " block."
    End If
    mp_FetchDslExtractGenerateBody = CStr(m.SubMatches(0))
End Function

Private Function mp_FetchDslParseColumns(ByVal generateBody As String) As Collection
    Dim m As Object
    Dim listValues As Variant
    Dim result As New Collection
    Dim i As Long
    Dim token As String

    Set m = mp_FetchDslMatchOne(generateBody, DSL_KW_DEFINE & "\s+" & DSL_KW_COLUMNS & "\s*\[([\s\S]*?)\]\s*;")
    If m Is Nothing Then
        Err.Raise vbObjectError + 1781, "ex_FetchDslEngine", "Fetch DSL: missing " & DSL_KW_DEFINE & " " & DSL_KW_COLUMNS & " in " & DSL_KW_GENERATE & " block."
    End If

    listValues = mp_FetchDslSplitCsv(CStr(m.SubMatches(0)))
    If mp_IsEmptyVariantArray(listValues) Then
        Err.Raise vbObjectError + 1782, "ex_FetchDslEngine", "Fetch DSL: " & DSL_KW_DEFINE & " " & DSL_KW_COLUMNS & " list is empty."
    End If

    For i = LBound(listValues) To UBound(listValues)
        token = Trim$(CStr(listValues(i)))
        If Len(token) > 0 Then result.Add token
    Next i

    Set mp_FetchDslParseColumns = result
End Function

Private Function mp_FetchDslExtractFirstRowAssignments(ByVal generateBody As String) As String
    Dim m As Object
    Set m = mp_FetchDslMatchOne(generateBody, DSL_KW_DEFINE & "\s+" & DSL_KW_ROW & "\s*\{([\s\S]*?)\}\s*;")
    If m Is Nothing Then
        Err.Raise vbObjectError + 1783, "ex_FetchDslEngine", "Fetch DSL: missing owner " & DSL_KW_DEFINE & " " & DSL_KW_ROW & " block."
    End If
    mp_FetchDslExtractFirstRowAssignments = CStr(m.SubMatches(0))
End Function

Private Sub mp_FetchDslExtractForeach(ByVal generateBody As String, ByVal contextAlias As String, ByRef outForeachVar As String, ByRef outForeachCtx As String, ByRef outAssignments As String)
    Dim m As Object
    Dim ctxAliasInForeach As String

    Set m = mp_FetchDslMatchOne(generateBody, DSL_KW_FOREACH & "\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*:\s*([A-Za-z_][A-Za-z0-9_]*)\.([A-Za-z_][A-Za-z0-9_]*)\s*\)\s*\{\s*" & DSL_KW_DEFINE & "\s+" & DSL_KW_ROW & "\s*\{([\s\S]*?)\}\s*;\s*\}\s*;")
    If m Is Nothing Then
        Err.Raise vbObjectError + 1784, "ex_FetchDslEngine", "Fetch DSL: missing " & DSL_KW_FOREACH & "(...) " & DSL_KW_DEFINE & " " & DSL_KW_ROW & " block."
    End If

    outForeachVar = Trim$(CStr(m.SubMatches(0)))
    ctxAliasInForeach = Trim$(CStr(m.SubMatches(1)))
    outForeachCtx = Trim$(CStr(m.SubMatches(2)))
    outAssignments = CStr(m.SubMatches(3))

    If StrComp(ctxAliasInForeach, contextAlias, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 1785, "ex_FetchDslEngine", "Fetch DSL: FOREACH must use context alias '" & contextAlias & "'."
    End If
End Sub

Private Function mp_FetchDslParseCondition(ByVal conditionText As String, ByVal sourceAlias As String, ByVal contextAlias As String) As Object
    conditionText = mp_FetchDslTrimToken(conditionText)
    If Len(conditionText) = 0 Then
        Err.Raise vbObjectError + 1786, "ex_FetchDslEngine", "Fetch DSL: empty " & DSL_KW_IF & " condition."
    End If

    Set mp_FetchDslParseCondition = mp_FetchDslParseCondOr(conditionText, sourceAlias, contextAlias)
End Function

Private Function mp_FetchDslParseCondOr(ByVal conditionText As String, ByVal sourceAlias As String, ByVal contextAlias As String) As Object
    Dim parts As Variant
    Dim i As Long
    Dim node As Object
    Dim items As New Collection
    Dim child As Object

    parts = mp_FetchDslSplitTopLevelByKeyword(conditionText, DSL_KW_OR)
    If Not mp_IsEmptyVariantArray(parts) Then
        If UBound(parts) > LBound(parts) Then
            Set node = mp_FetchDslCreateDictionary()
            node("Kind") = "Or"
            For i = LBound(parts) To UBound(parts)
                Set child = mp_FetchDslParseCondAnd(CStr(parts(i)), sourceAlias, contextAlias)
                items.Add child
            Next i
            Set node("Items") = items
            Set mp_FetchDslParseCondOr = node
            Exit Function
        End If
    End If

    Set mp_FetchDslParseCondOr = mp_FetchDslParseCondAnd(conditionText, sourceAlias, contextAlias)
End Function

Private Function mp_FetchDslParseCondAnd(ByVal conditionText As String, ByVal sourceAlias As String, ByVal contextAlias As String) As Object
    Dim parts As Variant
    Dim i As Long
    Dim node As Object
    Dim items As New Collection
    Dim child As Object

    parts = mp_FetchDslSplitTopLevelByKeyword(conditionText, DSL_KW_AND)
    If Not mp_IsEmptyVariantArray(parts) Then
        If UBound(parts) > LBound(parts) Then
            Set node = mp_FetchDslCreateDictionary()
            node("Kind") = "And"
            For i = LBound(parts) To UBound(parts)
                Set child = mp_FetchDslParseCondUnary(CStr(parts(i)), sourceAlias, contextAlias)
                items.Add child
            Next i
            Set node("Items") = items
            Set mp_FetchDslParseCondAnd = node
            Exit Function
        End If
    End If

    Set mp_FetchDslParseCondAnd = mp_FetchDslParseCondUnary(conditionText, sourceAlias, contextAlias)
End Function

Private Function mp_FetchDslParseCondUnary(ByVal conditionText As String, ByVal sourceAlias As String, ByVal contextAlias As String) As Object
    Dim textTrim As String
    Dim node As Object
    Dim restText As String
    Dim nextChar As String

    textTrim = mp_FetchDslTrimToken(conditionText)
    If Len(textTrim) = 0 Then
        Err.Raise vbObjectError + 1787, "ex_FetchDslEngine", "Fetch DSL: malformed " & DSL_KW_IF & " condition."
    End If

    If mp_FetchDslIsWrappedByOuterParens(textTrim) Then
        Set mp_FetchDslParseCondUnary = mp_FetchDslParseCondOr(Mid$(textTrim, 2, Len(textTrim) - 2), sourceAlias, contextAlias)
        Exit Function
    End If

    If Len(textTrim) >= Len(DSL_KW_NOT) Then
        If StrComp(UCase$(Left$(textTrim, Len(DSL_KW_NOT))), DSL_KW_NOT, vbBinaryCompare) = 0 Then
            If Len(textTrim) = Len(DSL_KW_NOT) Then
                nextChar = vbNullString
            Else
                nextChar = Mid$(textTrim, Len(DSL_KW_NOT) + 1, 1)
            End If
            If Len(nextChar) = 0 Or nextChar = " " Or nextChar = "(" Then
                restText = mp_FetchDslTrimToken(Mid$(textTrim, Len(DSL_KW_NOT) + 1))
                If Len(restText) = 0 Then
                    Err.Raise vbObjectError + 1788, "ex_FetchDslEngine", "Fetch DSL: " & DSL_KW_NOT & " requires condition operand."
                End If
                Set node = mp_FetchDslCreateDictionary()
                node("Kind") = "Not"
                Set node("Node") = mp_FetchDslParseCondUnary(restText, sourceAlias, contextAlias)
                Set mp_FetchDslParseCondUnary = node
                Exit Function
            End If
        End If
    End If

    Set mp_FetchDslParseCondUnary = mp_FetchDslParseCondPredicate(textTrim, sourceAlias, contextAlias)
End Function

Private Function mp_FetchDslParseCondPredicate(ByVal conditionText As String, ByVal sourceAlias As String, ByVal contextAlias As String) As Object
    Dim node As Object
    Dim m As Object
    Dim lhsAlias As String
    Dim rhsText As String

    Set m = mp_FetchDslMatchOne(conditionText, "^" & DSL_FN_IS_EMPTY & "\(\s*([\s\S]*)\s*\)$")
    If Not m Is Nothing Then
        Set node = mp_FetchDslCreateDictionary()
        node("Kind") = "Fn"
        node("Fn") = UCase$(DSL_FN_IS_EMPTY)
        Set node("ArgExpr") = mp_FetchDslParseExpr(CStr(m.SubMatches(0)), sourceAlias, vbNullString, contextAlias)
        Set mp_FetchDslParseCondPredicate = node
        Exit Function
    End If

    Set m = mp_FetchDslMatchOne(conditionText, "^" & DSL_FN_IS_NUMBER & "\(\s*([\s\S]*)\s*\)$")
    If Not m Is Nothing Then
        Set node = mp_FetchDslCreateDictionary()
        node("Kind") = "Fn"
        node("Fn") = UCase$(DSL_FN_IS_NUMBER)
        Set node("ArgExpr") = mp_FetchDslParseExpr(CStr(m.SubMatches(0)), sourceAlias, vbNullString, contextAlias)
        Set mp_FetchDslParseCondPredicate = node
        Exit Function
    End If

    Set m = mp_FetchDslMatchOne(conditionText, "^([A-Za-z_][A-Za-z0-9_]*)\.([A-Za-z_][A-Za-z0-9_]*)\s*" & DSL_OP_EQ & "\s*([\s\S]+)$")
    If Not m Is Nothing Then
        lhsAlias = mp_FetchDslTrimToken(CStr(m.SubMatches(0)))
        If StrComp(lhsAlias, sourceAlias, vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 1790, "ex_FetchDslEngine", "Fetch DSL: " & DSL_KW_IF & " " & DSL_OP_EQ & " must use source alias '" & sourceAlias & "'."
        End If
        rhsText = mp_FetchDslTrimToken(CStr(m.SubMatches(2)))

        Set node = mp_FetchDslCreateDictionary()
        node("Kind") = "Eq"
        node("Field") = mp_FetchDslTrimToken(CStr(m.SubMatches(1)))
        Set node("RhsExpr") = mp_FetchDslParseExpr(rhsText, sourceAlias, vbNullString, contextAlias)
        Set mp_FetchDslParseCondPredicate = node
        Exit Function
    End If

    Err.Raise vbObjectError + 1791, "ex_FetchDslEngine", _
        "Fetch DSL: unsupported " & DSL_KW_IF & " condition '" & conditionText & "'. Supported: " & _
        DSL_OP_EQ & ", " & DSL_KW_AND & ", " & DSL_KW_OR & ", " & DSL_KW_NOT & ", " & _
        DSL_FN_IS_EMPTY & "(...), " & DSL_FN_IS_NUMBER & "(...)."
End Function

Private Function mp_FetchDslSplitTopLevelByKeyword(ByVal textIn As String, ByVal keywordText As String) As Variant
    Dim src As String
    Dim parts() As String
    Dim startPos As Long
    Dim depth As Long
    Dim inQuotes As Boolean
    Dim i As Long
    Dim kwLen As Long
    Dim ch As String
    Dim prevChar As String
    Dim nextChar As String
    Dim chunk As String
    Dim partCount As Long

    src = mp_FetchDslTrimToken(textIn)
    If Len(src) = 0 Then
        mp_FetchDslSplitTopLevelByKeyword = Array()
        Exit Function
    End If

    kwLen = Len(keywordText)
    startPos = 1

    For i = 1 To Len(src)
        ch = Mid$(src, i, 1)
        If ch = """" Then
            inQuotes = Not inQuotes
        ElseIf Not inQuotes Then
            If ch = "(" Then
                depth = depth + 1
            ElseIf ch = ")" Then
                If depth > 0 Then depth = depth - 1
            End If

            If depth = 0 Then
                If (i + kwLen - 1) <= Len(src) Then
                    If StrComp(UCase$(Mid$(src, i, kwLen)), UCase$(keywordText), vbBinaryCompare) = 0 Then
                        If i = 1 Then
                            prevChar = vbNullString
                        Else
                            prevChar = Mid$(src, i - 1, 1)
                        End If

                        If (i + kwLen) > Len(src) Then
                            nextChar = vbNullString
                        Else
                            nextChar = Mid$(src, i + kwLen, 1)
                        End If

                        If mp_FetchDslIsBoundaryChar(prevChar) And mp_FetchDslIsBoundaryChar(nextChar) Then
                            chunk = mp_FetchDslTrimToken(Mid$(src, startPos, i - startPos))
                            ReDim Preserve parts(0 To partCount)
                            parts(partCount) = chunk
                            partCount = partCount + 1
                            startPos = i + kwLen
                            i = startPos - 1
                        End If
                    End If
                End If
            End If
        End If
    Next i

    If partCount = 0 Then
        mp_FetchDslSplitTopLevelByKeyword = Array(src)
        Exit Function
    End If

    chunk = mp_FetchDslTrimToken(Mid$(src, startPos))
    ReDim Preserve parts(0 To partCount)
    parts(partCount) = chunk
    mp_FetchDslSplitTopLevelByKeyword = parts
End Function

Private Function mp_FetchDslIsBoundaryChar(ByVal ch As String) As Boolean
    Dim code As Long

    If Len(ch) = 0 Then
        mp_FetchDslIsBoundaryChar = True
        Exit Function
    End If

    code = AscW(ch)
    If (code >= 48 And code <= 57) Or _
       (code >= 65 And code <= 90) Or _
       (code >= 97 And code <= 122) Or _
       code = 95 Then
        mp_FetchDslIsBoundaryChar = False
        Exit Function
    End If

    mp_FetchDslIsBoundaryChar = True
End Function

Private Function mp_FetchDslIsWrappedByOuterParens(ByVal textIn As String) As Boolean
    Dim src As String
    Dim depth As Long
    Dim i As Long
    Dim ch As String
    Dim inQuotes As Boolean

    src = mp_FetchDslTrimToken(textIn)
    If Len(src) < 2 Then Exit Function
    If Left$(src, 1) <> "(" Or Right$(src, 1) <> ")" Then Exit Function

    For i = 1 To Len(src)
        ch = Mid$(src, i, 1)
        If ch = """" Then
            inQuotes = Not inQuotes
        ElseIf Not inQuotes Then
            If ch = "(" Then
                depth = depth + 1
            ElseIf ch = ")" Then
                depth = depth - 1
                If depth = 0 And i < Len(src) Then Exit Function
                If depth < 0 Then Exit Function
            End If
        End If
    Next i

    mp_FetchDslIsWrappedByOuterParens = (depth = 0)
End Function

Private Function mp_FetchDslParsePushAssignments( _
    ByVal blockText As String, _
    ByVal sourceAlias As String, _
    ByVal contextAlias As String, _
    ByVal allowedFields As Object, _
    ByVal ctxName As String _
) As Collection
    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim p As Long
    Dim targetField As String
    Dim exprText As String
    Dim assignment As Object
    Dim result As New Collection

    lines = Split(Replace(Replace(blockText, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For i = LBound(lines) To UBound(lines)
        lineText = mp_FetchDslTrimToken(CStr(lines(i)))
        If Len(lineText) = 0 Then GoTo ContinueLine
        If Right$(lineText, 1) = "," Then lineText = Trim$(Left$(lineText, Len(lineText) - 1))
        If Left$(lineText, 1) = "." Then lineText = Mid$(lineText, 2)
        p = InStr(1, lineText, "=", vbBinaryCompare)
        If p <= 1 Then GoTo ContinueLine

        targetField = mp_FetchDslTrimToken(Left$(lineText, p - 1))
        exprText = Trim$(Mid$(lineText, p + 1))
        If Len(targetField) = 0 Then GoTo ContinueLine

        If Not allowedFields.Exists(targetField) Then
            Err.Raise vbObjectError + 1789, "ex_FetchDslEngine", _
                "Fetch DSL: field '" & targetField & "' is not declared in Ctx." & ctxName & "."
        End If

        Set assignment = mp_FetchDslCreateDictionary()
        assignment("Target") = targetField
        Set assignment("Expr") = mp_FetchDslParseExpr(exprText, sourceAlias, vbNullString, contextAlias)
        result.Add assignment
ContinueLine:
    Next i

    If result.Count = 0 Then
        Err.Raise vbObjectError + 1790, "ex_FetchDslEngine", "Fetch DSL: " & DSL_KW_PUSH & " block has no assignments."
    End If

    Set mp_FetchDslParsePushAssignments = result
End Function

Private Function mp_FetchDslParseRowAssignments(ByVal blockText As String, ByVal sourceAlias As String, ByVal foreachVar As String, ByVal contextAlias As String) As Collection
    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim p As Long
    Dim targetField As String
    Dim exprText As String
    Dim assignment As Object
    Dim result As New Collection

    lines = Split(Replace(Replace(blockText, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For i = LBound(lines) To UBound(lines)
        lineText = mp_FetchDslTrimToken(CStr(lines(i)))
        If Len(lineText) = 0 Then GoTo ContinueLine
        If Right$(lineText, 1) = "," Then lineText = Trim$(Left$(lineText, Len(lineText) - 1))
        p = InStr(1, lineText, "=", vbBinaryCompare)
        If p <= 1 Then GoTo ContinueLine

        targetField = mp_FetchDslTrimToken(Left$(lineText, p - 1))
        exprText = Trim$(Mid$(lineText, p + 1))
        If Len(targetField) = 0 Then GoTo ContinueLine

        Set assignment = mp_FetchDslCreateDictionary()
        assignment("Target") = targetField
        Set assignment("Expr") = mp_FetchDslParseExpr(exprText, sourceAlias, foreachVar, contextAlias)
        result.Add assignment
ContinueLine:
    Next i

    If result.Count = 0 Then
        Err.Raise vbObjectError + 1791, "ex_FetchDslEngine", "Fetch DSL: " & DSL_KW_DEFINE & " " & DSL_KW_ROW & " block has no assignments."
    End If

    Set mp_FetchDslParseRowAssignments = result
End Function

Private Function mp_FetchDslParseExpr(ByVal exprText As String, ByVal sourceAlias As String, ByVal foreachVar As String, ByVal contextAlias As String) As Object
    Dim expr As Object
    Dim m As Object
    Dim tokenAlias As String
    Dim tokenField As String

    exprText = Trim$(exprText)
    Set expr = mp_FetchDslCreateDictionary()

    If Len(exprText) = 0 Then
        expr("Kind") = "Literal"
        expr("Value") = vbNullString
        Set mp_FetchDslParseExpr = expr
        Exit Function
    End If

    If Left$(exprText, 1) = """" And Right$(exprText, 1) = """" And Len(exprText) >= 2 Then
        expr("Kind") = "Literal"
        expr("Value") = Mid$(exprText, 2, Len(exprText) - 2)
        Set mp_FetchDslParseExpr = expr
        Exit Function
    End If

    If Left$(exprText, 1) = "$" Then
        expr("Kind") = "Variable"
        expr("Value") = exprText
        Set mp_FetchDslParseExpr = expr
        Exit Function
    End If

    Set m = mp_FetchDslMatchOne(exprText, "^" & DSL_KW_LAST & "\(\s*([A-Za-z_][A-Za-z0-9_]*)\.([A-Za-z_][A-Za-z0-9_]*)\s*\)\.([A-Za-z_][A-Za-z0-9_]*)$")
    If Not m Is Nothing Then
        If StrComp(Trim$(CStr(m.SubMatches(0))), contextAlias, vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 1792, "ex_FetchDslEngine", "Fetch DSL: " & DSL_KW_LAST & " must use context alias '" & contextAlias & "'."
        End If
        expr("Kind") = "LastCtxField"
        expr("CtxName") = Trim$(CStr(m.SubMatches(1)))
        expr("Field") = Trim$(CStr(m.SubMatches(2)))
        Set mp_FetchDslParseExpr = expr
        Exit Function
    End If

    Set m = mp_FetchDslMatchOne(exprText, "^([A-Za-z_][A-Za-z0-9_]*)\.([A-Za-z_][A-Za-z0-9_]*)$")
    If Not m Is Nothing Then
        tokenAlias = Trim$(CStr(m.SubMatches(0)))
        tokenField = Trim$(CStr(m.SubMatches(1)))
        If StrComp(tokenAlias, sourceAlias, vbTextCompare) <> 0 And _
           (Len(foreachVar) = 0 Or StrComp(tokenAlias, foreachVar, vbTextCompare) <> 0) Then
            Err.Raise vbObjectError + 1793, "ex_FetchDslEngine", "Fetch DSL: unknown expression alias '" & tokenAlias & "' in '" & exprText & "'."
        End If

        expr("Kind") = "FieldRef"
        expr("Alias") = tokenAlias
        expr("Field") = tokenField
        Set mp_FetchDslParseExpr = expr
        Exit Function
    End If

    expr("Kind") = "Literal"
    expr("Value") = exprText
    Set mp_FetchDslParseExpr = expr
End Function

Private Function mp_FetchDslSplitCsv(ByVal csvText As String) As Variant
    Dim parts() As String
    Dim out() As String
    Dim i As Long
    Dim token As String
    Dim count As Long

    csvText = Replace(Replace(csvText, vbCrLf, ","), vbCr, ",")
    csvText = Replace(csvText, vbLf, ",")
    parts = Split(csvText, ",")

    For i = LBound(parts) To UBound(parts)
        token = mp_FetchDslTrimToken(CStr(parts(i)))
        If Len(token) > 0 Then
            If Left$(token, 1) = """" And Right$(token, 1) = """" And Len(token) >= 2 Then
                token = Mid$(token, 2, Len(token) - 2)
            End If
            ReDim Preserve out(0 To count)
            out(count) = token
            count = count + 1
        End If
    Next i

    If count = 0 Then
        mp_FetchDslSplitCsv = Array()
    Else
        mp_FetchDslSplitCsv = out
    End If
End Function

Private Function mp_FetchDslTrimToken(ByVal textValue As String) As String
    textValue = Replace(textValue, vbCrLf, " ")
    textValue = Replace(textValue, vbCr, " ")
    textValue = Replace(textValue, vbLf, " ")
    textValue = Replace(textValue, vbTab, " ")
    textValue = Replace(textValue, ChrW$(160), " ")
    textValue = Replace(textValue, "  ", " ")
    textValue = Replace(textValue, "  ", " ")
    mp_FetchDslTrimToken = Trim$(textValue)
End Function

Private Function mp_FetchDslBuildSourceFieldOrdinals(ByVal cfg As Object, ByVal sourceAlias As String, ByVal tableAlias As String, ByVal plan As Object, ByVal rs As Object) As Object
    Dim tokenSet As Object
    Dim token As Variant
    Dim ordinals As Object
    Dim recordsetOrdinals As Object
    Dim srcHeader As String
    Dim normHeader As String

    Set tokenSet = mp_FetchDslCreateDictionary()
    mp_FetchDslCollectSourceFieldTokens plan, tokenSet

    Set recordsetOrdinals = mp_FetchDslBuildRecordsetFieldOrdinals(rs)
    Set ordinals = mp_FetchDslCreateDictionary()

    For Each token In tokenSet.Keys
        srcHeader = mp_FetchDslResolveSourceHeader(cfg, sourceAlias, tableAlias, CStr(token))
        normHeader = mp_FetchDslNormalizeToken(srcHeader)
        If Not recordsetOrdinals.Exists(normHeader) Then
            Err.Raise vbObjectError + 1794, "ex_FetchDslEngine", _
                "Fetch DSL: source field '" & CStr(token) & "' (header '" & srcHeader & "') was not found in ADO recordset."
        End If
        ordinals(mp_FetchDslNormalizeToken(CStr(token))) = CLng(recordsetOrdinals(normHeader))
    Next token

    Set mp_FetchDslBuildSourceFieldOrdinals = ordinals
End Function

Private Function mp_FetchDslBuildSourceProjectionSql( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal plan As Object, _
    ByVal tableRef As String _
) As String
    Dim tokenSet As Object
    Dim headers As Object
    Dim token As Variant
    Dim sourceHeader As String
    Dim normalizedHeader As String
    Dim selectItems As Collection
    Dim item As Variant
    Dim selectList As String

    If Len(Trim$(tableRef)) = 0 Then Exit Function

    Set tokenSet = mp_FetchDslCreateDictionary()
    mp_FetchDslCollectSourceFieldTokens plan, tokenSet

    If tokenSet.Count = 0 Then
        mp_FetchDslBuildSourceProjectionSql = "SELECT * FROM " & tableRef
        Exit Function
    End If

    Set headers = mp_FetchDslCreateDictionary()
    Set selectItems = New Collection

    For Each token In tokenSet.Keys
        sourceHeader = mp_FetchDslResolveSourceHeader(cfg, sourceAlias, tableAlias, CStr(token))
        normalizedHeader = mp_FetchDslNormalizeToken(sourceHeader)
        If Len(normalizedHeader) = 0 Then GoTo ContinueToken
        If headers.Exists(normalizedHeader) Then GoTo ContinueToken

        headers(normalizedHeader) = True
        selectItems.Add mp_FetchDslQuoteSqlIdentifier(sourceHeader)
ContinueToken:
    Next token

    For Each item In selectItems
        If Len(selectList) > 0 Then selectList = selectList & ", "
        selectList = selectList & CStr(item)
    Next item

    If Len(selectList) = 0 Then
        mp_FetchDslBuildSourceProjectionSql = "SELECT * FROM " & tableRef
    Else
        mp_FetchDslBuildSourceProjectionSql = "SELECT " & selectList & " FROM " & tableRef
    End If
End Function

Private Function mp_FetchDslQuoteSqlIdentifier(ByVal valueText As String) As String
    valueText = Trim$(valueText)
    If Len(valueText) >= 2 Then
        If Left$(valueText, 1) = "[" And Right$(valueText, 1) = "]" Then
            valueText = Mid$(valueText, 2, Len(valueText) - 2)
        End If
    End If
    valueText = Replace$(valueText, "]", "]]")
    mp_FetchDslQuoteSqlIdentifier = "[" & valueText & "]"
End Function

Private Function mp_FetchDslLocalizeInnerErrorRu(ByVal errorText As String) As String
    errorText = Replace$(errorText, "Неприпустиме використання дужок для імен", "Недопустимое использование скобок для имен")
    errorText = Replace$(errorText, "Ім'я", "Имя")
    errorText = Replace$(errorText, "імен", "имен")
    mp_FetchDslLocalizeInnerErrorRu = errorText
End Function

Private Sub mp_FetchDslCollectSourceFieldTokens(ByVal plan As Object, ByVal ioTokenSet As Object)
    Dim rule As Object
    Dim assignment As Variant
    Dim expr As Object
    Dim sourceAlias As String

    If ioTokenSet Is Nothing Then Exit Sub
    sourceAlias = CStr(plan("SourceAlias"))

    ioTokenSet(mp_FetchDslNormalizeToken(CStr(plan("KeyField")))) = True

    Set rule = plan("AboveRule")
    mp_FetchDslCollectFieldRefsFromCondition rule("Condition"), sourceAlias, ioTokenSet
    For Each assignment In rule("PushAssignments")
        Set expr = assignment("Expr")
        mp_FetchDslCollectFieldRefFromExpr expr, sourceAlias, ioTokenSet
    Next assignment

    Set rule = plan("BelowRule")
    mp_FetchDslCollectFieldRefsFromCondition rule("Condition"), sourceAlias, ioTokenSet
    For Each assignment In rule("PushAssignments")
        Set expr = assignment("Expr")
        mp_FetchDslCollectFieldRefFromExpr expr, sourceAlias, ioTokenSet
    Next assignment

    For Each assignment In plan("GenerateKeyAssignments")
        Set expr = assignment("Expr")
        mp_FetchDslCollectFieldRefFromExpr expr, sourceAlias, ioTokenSet
    Next assignment
    For Each assignment In plan("GenerateForeachAssignments")
        Set expr = assignment("Expr")
        mp_FetchDslCollectFieldRefFromExpr expr, sourceAlias, ioTokenSet
    Next assignment
End Sub

Private Sub mp_FetchDslCollectFieldRefsFromCondition(ByVal condNode As Object, ByVal sourceAlias As String, ByVal ioTokenSet As Object)
    Dim kindName As String
    Dim expr As Object
    Dim item As Variant
    Dim nestedNode As Object

    If condNode Is Nothing Then Exit Sub
    kindName = UCase$(CStr(condNode("Kind")))

    Select Case kindName
        Case "EQ"
            ioTokenSet(mp_FetchDslNormalizeToken(CStr(condNode("Field")))) = True
            Set expr = condNode("RhsExpr")
            mp_FetchDslCollectFieldRefFromExpr expr, sourceAlias, ioTokenSet

        Case "FN"
            Set expr = condNode("ArgExpr")
            mp_FetchDslCollectFieldRefFromExpr expr, sourceAlias, ioTokenSet

        Case "NOT"
            Set nestedNode = condNode("Node")
            mp_FetchDslCollectFieldRefsFromCondition nestedNode, sourceAlias, ioTokenSet

        Case "AND", "OR"
            For Each item In condNode("Items")
                Set nestedNode = item
                mp_FetchDslCollectFieldRefsFromCondition nestedNode, sourceAlias, ioTokenSet
            Next item
    End Select
End Sub

Private Sub mp_FetchDslCollectFieldRefFromExpr(ByVal expr As Object, ByVal sourceAlias As String, ByVal ioTokenSet As Object)
    If expr Is Nothing Then Exit Sub
    If StrComp(CStr(expr("Kind")), "FieldRef", vbTextCompare) <> 0 Then Exit Sub
    If StrComp(CStr(expr("Alias")), sourceAlias, vbTextCompare) <> 0 Then Exit Sub
    ioTokenSet(mp_FetchDslNormalizeToken(CStr(expr("Field")))) = True
End Sub

Private Function mp_FetchDslBuildRecordsetFieldOrdinals(ByVal rs As Object) As Object
    Dim result As Object
    Dim i As Long
    Dim fieldName As String

    Set result = mp_FetchDslCreateDictionary()
    For i = 0 To rs.Fields.Count - 1
        fieldName = CStr(rs.Fields(i).Name)
        result(mp_FetchDslNormalizeToken(fieldName)) = i
    Next i
    Set mp_FetchDslBuildRecordsetFieldOrdinals = result
End Function

Private Function mp_FetchDslBuildSourceFieldTypes(ByVal fieldOrdByToken As Object, ByVal rs As Object) As Object
    Dim result As Object
    Dim token As Variant
    Dim ordinal As Long

    Set result = mp_FetchDslCreateDictionary()
    If fieldOrdByToken Is Nothing Then
        Set mp_FetchDslBuildSourceFieldTypes = result
        Exit Function
    End If
    If rs Is Nothing Then
        Set mp_FetchDslBuildSourceFieldTypes = result
        Exit Function
    End If

    For Each token In fieldOrdByToken.Keys
        ordinal = CLng(fieldOrdByToken(CStr(token)))
        If ordinal >= 0 And ordinal < rs.Fields.Count Then
            result(CStr(token)) = CLng(rs.Fields(ordinal).Type)
        End If
    Next token

    Set mp_FetchDslBuildSourceFieldTypes = result
End Function

Private Function mp_FetchDslResolveSourceHeader(ByVal cfg As Object, ByVal sourceAlias As String, ByVal tableAlias As String, ByVal fieldToken As String) As String
    Dim mapKey As String
    mapKey = sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldToken & "]"

    If cfg.Exists(mapKey) Then
        mp_FetchDslResolveSourceHeader = mp_GetMappedSourceHeader(cfg, sourceAlias, tableAlias, fieldToken)
    Else
        mp_FetchDslResolveSourceHeader = fieldToken
    End If
End Function

Private Function mp_FetchDslFindKeyRows(ByVal rowsData As Variant, ByVal plan As Object, ByVal keyValue As String, ByVal fieldOrdByToken As Object) As Collection
    Dim result As New Collection
    Dim keyToken As String
    Dim keyOrdinal As Long
    Dim r As Long
    Dim rowLower As Long
    Dim rowUpper As Long
    Dim cellText As String

    keyToken = mp_FetchDslNormalizeToken(CStr(plan("KeyField")))
    If Not fieldOrdByToken.Exists(keyToken) Then
        Set mp_FetchDslFindKeyRows = result
        Exit Function
    End If
    keyOrdinal = CLng(fieldOrdByToken(keyToken))

    rowLower = LBound(rowsData, 2)
    rowUpper = UBound(rowsData, 2)
    For r = rowLower To rowUpper
        cellText = Trim$(mp_ToSafeText(rowsData(keyOrdinal, r)))
        If StrComp(cellText, Trim$(keyValue), vbTextCompare) = 0 Then result.Add r
    Next r

    Set mp_FetchDslFindKeyRows = result
End Function

Private Function mp_FetchDslFindNextKeyRow(ByVal rowsData As Variant, ByVal plan As Object, ByVal keyValue As String, ByVal fieldOrdByToken As Object, ByVal currentKeyRow As Long, ByVal rowUpper As Long) As Long
    Dim keyToken As String
    Dim keyOrdinal As Long
    Dim r As Long
    Dim cellText As String

    keyToken = mp_FetchDslNormalizeToken(CStr(plan("KeyField")))
    If Not fieldOrdByToken.Exists(keyToken) Then
        mp_FetchDslFindNextKeyRow = rowUpper + 1
        Exit Function
    End If
    keyOrdinal = CLng(fieldOrdByToken(keyToken))

    For r = currentKeyRow + 1 To rowUpper
        cellText = Trim$(mp_ToSafeText(rowsData(keyOrdinal, r)))
        If StrComp(cellText, Trim$(keyValue), vbTextCompare) = 0 Then
            mp_FetchDslFindNextKeyRow = r
            Exit Function
        End If
    Next r

    mp_FetchDslFindNextKeyRow = rowUpper + 1
End Function

Private Function mp_FetchDslInitRuntimeContext(ByVal plan As Object) As Object
    Dim result As Object
    Dim ctxName As Variant

    Set result = mp_FetchDslCreateDictionary()
    For Each ctxName In plan("Contexts").Keys
        Set result(CStr(ctxName)) = New Collection
    Next ctxName
    Set mp_FetchDslInitRuntimeContext = result
End Function

Private Sub mp_FetchDslRunSearchRule( _
    ByVal rule As Object, _
    ByVal rowsData As Variant, _
    ByVal rowStart As Long, _
    ByVal rowEnd As Long, _
    ByVal runtimeCtx As Object, _
    ByVal plan As Object, _
    ByVal keyValue As String, _
    ByVal fieldOrdByToken As Object, _
    ByVal keyRowIndex As Long, _
    ByVal foreachItem As Object _
)
    Dim r As Long
    Dim bucket As Collection
    Dim item As Object
    Dim assignment As Variant
    Dim targetName As String
    Dim valueText As String
    Dim keepMode As String
    Dim isMatch As Boolean

    If rule Is Nothing Then Exit Sub
    If rowEnd < rowStart Then Exit Sub

    Set bucket = runtimeCtx(CStr(rule("CtxName")))
    keepMode = UCase$(CStr(rule("KeepMode")))
    For r = rowStart To rowEnd
        isMatch = mp_FetchDslEvalCondition(rule("Condition"), rowsData, r, plan, keyValue, fieldOrdByToken, runtimeCtx, keyRowIndex, foreachItem)
        If isMatch Then
            Set item = mp_FetchDslCreateDictionary()
            For Each assignment In rule("PushAssignments")
                targetName = CStr(assignment("Target"))
                valueText = mp_FetchDslEvalExpr(assignment("Expr"), rowsData, r, plan, keyValue, fieldOrdByToken, runtimeCtx, keyRowIndex, foreachItem)
                item(targetName) = valueText
            Next assignment
            bucket.Add item
        ElseIf StrComp(keepMode, DSL_KEEP_MODE_CONTIGUOUS, vbTextCompare) = 0 Then
            Exit For
        End If
    Next r

    mp_FetchDslApplyKeep bucket, keepMode
End Sub

Private Sub mp_FetchDslApplyKeep(ByVal bucket As Collection, ByVal keepMode As String)
    Dim keepItem As Object

    If bucket Is Nothing Then Exit Sub
    If bucket.Count = 0 Then Exit Sub
    If StrComp(keepMode, DSL_KEEP_MODE_ALL, vbTextCompare) = 0 Then Exit Sub
    If StrComp(keepMode, DSL_KEEP_MODE_CONTIGUOUS, vbTextCompare) = 0 Then Exit Sub

    If StrComp(keepMode, DSL_KEEP_MODE_LAST, vbTextCompare) = 0 Then
        Set keepItem = bucket(bucket.Count)
    ElseIf StrComp(keepMode, DSL_KEEP_MODE_FIRST, vbTextCompare) = 0 Then
        Set keepItem = bucket(1)
    Else
        Exit Sub
    End If

    Do While bucket.Count > 0
        bucket.Remove bucket.Count
    Loop
    bucket.Add keepItem
End Sub

Private Function mp_FetchDslEvalCondition( _
    ByVal cond As Object, _
    ByVal rowsData As Variant, _
    ByVal sourceRowIndex As Long, _
    ByVal plan As Object, _
    ByVal keyValue As String, _
    ByVal fieldOrdByToken As Object, _
    ByVal runtimeCtx As Object, _
    ByVal keyRowIndex As Long, _
    ByVal foreachItem As Object _
) As Boolean
    mp_FetchDslEvalCondition = mp_FetchDslEvalConditionNode(cond, rowsData, sourceRowIndex, plan, keyValue, fieldOrdByToken, runtimeCtx, keyRowIndex, foreachItem)
End Function

Private Function mp_FetchDslEvalConditionNode( _
    ByVal condNode As Object, _
    ByVal rowsData As Variant, _
    ByVal sourceRowIndex As Long, _
    ByVal plan As Object, _
    ByVal keyValue As String, _
    ByVal fieldOrdByToken As Object, _
    ByVal runtimeCtx As Object, _
    ByVal keyRowIndex As Long, _
    ByVal foreachItem As Object _
) As Boolean
    Dim kindName As String
    Dim lhsText As String
    Dim rhsText As String
    Dim argText As String
    Dim item As Variant
    Dim nestedNode As Object
    Dim fnName As String
    Dim fieldTypeByToken As Object

    If condNode Is Nothing Then Exit Function
    kindName = UCase$(CStr(condNode("Kind")))
    If plan.Exists("__FieldTypeByToken") Then Set fieldTypeByToken = plan("__FieldTypeByToken")

    Select Case kindName
        Case "EQ"
            lhsText = mp_FetchDslGetSourceValue(rowsData, sourceRowIndex, CStr(condNode("Field")), fieldOrdByToken, fieldTypeByToken)
            rhsText = mp_FetchDslEvalExpr(condNode("RhsExpr"), rowsData, sourceRowIndex, plan, keyValue, fieldOrdByToken, runtimeCtx, keyRowIndex, foreachItem)
            mp_FetchDslEvalConditionNode = (StrComp(lhsText, rhsText, vbTextCompare) = 0)

        Case "FN"
            fnName = UCase$(CStr(condNode("Fn")))
            argText = mp_FetchDslEvalExpr(condNode("ArgExpr"), rowsData, sourceRowIndex, plan, keyValue, fieldOrdByToken, runtimeCtx, keyRowIndex, foreachItem)
            argText = mp_FetchDslTrimToken(argText)
            Select Case fnName
                Case UCase$(DSL_FN_IS_EMPTY)
                    mp_FetchDslEvalConditionNode = (Len(argText) = 0)
                Case UCase$(DSL_FN_IS_NUMBER)
                    If Len(argText) > 0 Then
                        mp_FetchDslEvalConditionNode = IsNumeric(argText)
                    End If
                Case Else
                    Err.Raise vbObjectError + 1792, "ex_FetchDslEngine", "Fetch DSL: unsupported predicate function '" & fnName & "'."
            End Select

        Case "NOT"
            Set nestedNode = condNode("Node")
            mp_FetchDslEvalConditionNode = Not mp_FetchDslEvalConditionNode(nestedNode, rowsData, sourceRowIndex, plan, keyValue, fieldOrdByToken, runtimeCtx, keyRowIndex, foreachItem)

        Case "AND"
            mp_FetchDslEvalConditionNode = True
            For Each item In condNode("Items")
                Set nestedNode = item
                If Not mp_FetchDslEvalConditionNode(nestedNode, rowsData, sourceRowIndex, plan, keyValue, fieldOrdByToken, runtimeCtx, keyRowIndex, foreachItem) Then
                    mp_FetchDslEvalConditionNode = False
                    Exit Function
                End If
            Next item

        Case "OR"
            For Each item In condNode("Items")
                Set nestedNode = item
                If mp_FetchDslEvalConditionNode(nestedNode, rowsData, sourceRowIndex, plan, keyValue, fieldOrdByToken, runtimeCtx, keyRowIndex, foreachItem) Then
                    mp_FetchDslEvalConditionNode = True
                    Exit Function
                End If
            Next item

        Case Else
            Err.Raise vbObjectError + 1793, "ex_FetchDslEngine", "Fetch DSL: unsupported condition node kind '" & kindName & "'."
    End Select
End Function

Private Function mp_FetchDslEvalExpr( _
    ByVal expr As Object, _
    ByVal rowsData As Variant, _
    ByVal sourceRowIndex As Long, _
    ByVal plan As Object, _
    ByVal keyValue As String, _
    ByVal fieldOrdByToken As Object, _
    ByVal runtimeCtx As Object, _
    ByVal keyRowIndex As Long, _
    ByVal foreachItem As Object _
) As String
    Dim kindName As String
    Dim varName As String
    Dim refAlias As String
    Dim refField As String
    Dim ctxName As String
    Dim ctxRows As Collection
    Dim lastItem As Object
    Dim fieldTypeByToken As Object

    If expr Is Nothing Then Exit Function
    kindName = CStr(expr("Kind"))
    If plan.Exists("__FieldTypeByToken") Then Set fieldTypeByToken = plan("__FieldTypeByToken")

    Select Case UCase$(kindName)
        Case "LITERAL"
            mp_FetchDslEvalExpr = CStr(expr("Value"))
        Case "VARIABLE"
            varName = CStr(expr("Value"))
            If StrComp(varName, CStr(plan("KeyVar")), vbTextCompare) = 0 Then
                mp_FetchDslEvalExpr = keyValue
            Else
                mp_FetchDslEvalExpr = vbNullString
            End If
        Case "FIELDREF"
            refAlias = CStr(expr("Alias"))
            refField = CStr(expr("Field"))
            If StrComp(refAlias, CStr(plan("SourceAlias")), vbTextCompare) = 0 Then
                If sourceRowIndex >= 0 Then
                    mp_FetchDslEvalExpr = mp_FetchDslGetSourceValue(rowsData, sourceRowIndex, refField, fieldOrdByToken, fieldTypeByToken)
                Else
                    mp_FetchDslEvalExpr = mp_FetchDslGetSourceValue(rowsData, keyRowIndex, refField, fieldOrdByToken, fieldTypeByToken)
                End If
            ElseIf Not foreachItem Is Nothing And foreachItem.Exists(refField) Then
                mp_FetchDslEvalExpr = CStr(foreachItem(refField))
            ElseIf Not foreachItem Is Nothing And foreachItem.Exists(LCase$(refField)) Then
                mp_FetchDslEvalExpr = CStr(foreachItem(LCase$(refField)))
            End If
        Case "LASTCTXFIELD"
            ctxName = CStr(expr("CtxName"))
            If Not runtimeCtx Is Nothing Then
                If runtimeCtx.Exists(ctxName) Then
                    Set ctxRows = runtimeCtx(ctxName)
                    If Not ctxRows Is Nothing Then
                        If ctxRows.Count > 0 Then
                            Set lastItem = ctxRows(ctxRows.Count)
                            refField = CStr(expr("Field"))
                            If lastItem.Exists(refField) Then
                                mp_FetchDslEvalExpr = CStr(lastItem(refField))
                            ElseIf lastItem.Exists(LCase$(refField)) Then
                                mp_FetchDslEvalExpr = CStr(lastItem(LCase$(refField)))
                            End If
                        End If
                    End If
                End If
            End If
    End Select
End Function

Private Function mp_FetchDslGetSourceValue( _
    ByVal rowsData As Variant, _
    ByVal rowIndex As Long, _
    ByVal fieldToken As String, _
    ByVal fieldOrdByToken As Object, _
    Optional ByVal fieldTypeByToken As Object = Nothing _
) As String
    Dim tokenKey As String
    Dim ordinal As Long
    Dim adoFieldType As Long

    tokenKey = mp_FetchDslNormalizeToken(fieldToken)
    If Not fieldOrdByToken.Exists(tokenKey) Then Exit Function
    ordinal = CLng(fieldOrdByToken(tokenKey))
    adoFieldType = -1
    If Not fieldTypeByToken Is Nothing Then
        If fieldTypeByToken.Exists(tokenKey) Then
            adoFieldType = CLng(fieldTypeByToken(tokenKey))
        End If
    End If
    mp_FetchDslGetSourceValue = Trim$(ex_SqlAdoHelpers.m_ToNormalizedText(rowsData(ordinal, rowIndex), adoFieldType))
End Function

Private Function mp_FetchDslEvaluateAssignments( _
    ByVal assignments As Collection, _
    ByVal plan As Object, _
    ByVal rowsData As Variant, _
    ByVal keyRowIndex As Long, _
    ByVal keyValue As String, _
    ByVal runtimeCtx As Object, _
    ByVal foreachItem As Object _
) As Object
    Dim result As Object
    Dim assignment As Variant
    Dim targetName As String
    Dim valueText As String
    Dim fieldOrdByToken As Object

    Set result = mp_FetchDslCreateDictionary()
    Set fieldOrdByToken = plan("__FieldOrdByToken")

    For Each assignment In assignments
        targetName = CStr(assignment("Target"))
        valueText = mp_FetchDslEvalExpr(assignment("Expr"), rowsData, -1, plan, keyValue, fieldOrdByToken, runtimeCtx, keyRowIndex, foreachItem)
        result(targetName) = valueText
        result(LCase$(targetName)) = valueText
    Next assignment

    Set mp_FetchDslEvaluateAssignments = result
End Function

Private Function mp_FetchDslBuildGeneratedMetaRows(ByVal plan As Object, ByVal rowsData As Variant, ByVal keyRowIndex As Long, ByVal keyValue As String, ByVal runtimeCtx As Object) As Collection
    Dim result As New Collection
    Dim foreachCtxName As String
    Dim entries As Collection
    Dim entry As Object
    Dim rowValues As Object
    Dim keyName As Variant
    Dim assignValues As Object

    foreachCtxName = CStr(plan("GenerateForeachCtx"))
    If Not runtimeCtx.Exists(foreachCtxName) Then
        Set mp_FetchDslBuildGeneratedMetaRows = result
        Exit Function
    End If
    Set entries = runtimeCtx(foreachCtxName)
    If entries Is Nothing Then
        Set mp_FetchDslBuildGeneratedMetaRows = result
        Exit Function
    End If

    For Each entry In entries
        Set rowValues = mp_FetchDslCreateDictionary()
        For Each keyName In entry.Keys
            rowValues(CStr(keyName)) = CStr(entry(keyName))
            rowValues(LCase$(CStr(keyName))) = CStr(entry(keyName))
        Next keyName

        Set assignValues = mp_FetchDslEvaluateAssignments(plan("GenerateForeachAssignments"), plan, rowsData, keyRowIndex, keyValue, runtimeCtx, entry)
        For Each keyName In assignValues.Keys
            rowValues(CStr(keyName)) = CStr(assignValues(keyName))
        Next keyName

        result.Add rowValues
    Next entry

    Set mp_FetchDslBuildGeneratedMetaRows = result
End Function

Private Function mp_FetchDslBuildFieldIndexMap(ByVal fields As Variant) As Object
    Dim result As Object
    Dim i As Long
    Dim aliasName As String

    Set result = mp_FetchDslCreateDictionary()
    If mp_IsEmptyVariantArray(fields) Then
        Set mp_FetchDslBuildFieldIndexMap = result
        Exit Function
    End If

    For i = LBound(fields) To UBound(fields)
        aliasName = Trim$(CStr(fields(i)))
        If Len(aliasName) > 0 Then
            result(LCase$(aliasName)) = 1 + (i - LBound(fields))
        End If
    Next i

    Set mp_FetchDslBuildFieldIndexMap = result
End Function

Private Function mp_FetchDslNormalizeToken(ByVal tokenText As String) As String
    tokenText = Replace$(tokenText, vbCr, " ")
    tokenText = Replace$(tokenText, vbLf, " ")
    tokenText = Replace$(tokenText, vbTab, " ")
    tokenText = Replace$(tokenText, ChrW$(160), " ")
    ' ACE/OLEDB can expose dots in Excel headers as '#', e.g. "Вх. №" -> "Вх# №".
    tokenText = Replace$(tokenText, "#", ".")
    tokenText = Replace$(tokenText, ChrW$(&H2019), "'")
    tokenText = Replace$(tokenText, ChrW$(&H2BC), "'")
    tokenText = Replace$(tokenText, ChrW$(&H60), "'")
    tokenText = Replace$(tokenText, ChrW$(&HB4), "'")
    tokenText = Replace$(tokenText, "  ", " ")
    tokenText = Replace$(tokenText, "  ", " ")
    mp_FetchDslNormalizeToken = LCase$(Trim$(tokenText))
End Function

Private Function mp_GetFetchDslText(ByVal cfg As Object, ByVal sourceAlias As String, ByVal tableAlias As String) As String
    Dim prefix As String
    prefix = sourceAlias & ".Sheet[" & tableAlias & "].Fetch."
    mp_GetFetchDslText = Trim$(mp_GetCfgOptional(cfg, prefix & "Dsl", vbNullString))
End Function

Private Function mp_ArrayContainsText(ByVal values As Variant, ByVal needle As String) As Boolean
    Dim i As Long

    If mp_IsEmptyVariantArray(values) Then Exit Function

    For i = LBound(values) To UBound(values)
        If StrComp(Trim$(CStr(values(i))), Trim$(needle), vbTextCompare) = 0 Then
            mp_ArrayContainsText = True
            Exit Function
        End If
    Next i
End Function

Private Function mp_GetMappedSourceHeader( _
    ByVal cfg As Object, _
    ByVal sourceAlias As String, _
    ByVal tableAlias As String, _
    ByVal fieldAlias As String _
) As String
    Dim raw As String
    Dim p As Long

    raw = mp_GetCfgRequired(cfg, sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]")
    p = InStr(1, raw, "|", vbBinaryCompare)

    If p > 0 Then
        mp_GetMappedSourceHeader = Trim$(Left$(raw, p - 1))
    Else
        mp_GetMappedSourceHeader = Trim$(raw)
    End If

    If Len(mp_GetMappedSourceHeader) >= 2 Then
        If Left$(mp_GetMappedSourceHeader, 1) = "[" And Right$(mp_GetMappedSourceHeader, 1) = "]" Then
            mp_GetMappedSourceHeader = Trim$(Mid$(mp_GetMappedSourceHeader, 2, Len(mp_GetMappedSourceHeader) - 2))
        End If
    End If

    If Len(mp_GetMappedSourceHeader) = 0 Then
        Err.Raise vbObjectError + 1390, "ex_FetchDslEngine", _
            "Mapped source header is empty for " & sourceAlias & ".Sheet[" & tableAlias & "].Map[" & fieldAlias & "]"
    End If
End Function

Private Function mp_GetCfgRequired(ByVal cfg As Object, ByVal keyName As String) As String
    Dim valueText As String

    If Not cfg.Exists(keyName) Then
        Err.Raise vbObjectError + 1370, "ex_FetchDslEngine", "Missing config key: " & keyName
    End If

    valueText = Trim$(CStr(cfg(keyName)))
    If Len(valueText) = 0 Then
        Err.Raise vbObjectError + 1371, "ex_FetchDslEngine", "Empty config value: " & keyName
    End If

    mp_GetCfgRequired = valueText
End Function

Private Function mp_GetCfgOptional(ByVal cfg As Object, ByVal keyName As String, ByVal defaultValue As String) As String
    If cfg.Exists(keyName) Then
        mp_GetCfgOptional = Trim$(CStr(cfg(keyName)))
    Else
        mp_GetCfgOptional = defaultValue
    End If
End Function

Private Function mp_ToSafeText(ByVal valueIn As Variant) As String
    If IsError(valueIn) Then Exit Function
    If IsNull(valueIn) Then Exit Function
    If IsEmpty(valueIn) Then Exit Function

    mp_ToSafeText = Trim$(CStr(valueIn))
End Function

Private Function mp_IsEmptyVariantArray(ByVal v As Variant) As Boolean
    On Error GoTo EH

    If IsArray(v) = False Then
        mp_IsEmptyVariantArray = True
        Exit Function
    End If

    If UBound(v) < LBound(v) Then
        mp_IsEmptyVariantArray = True
        Exit Function
    End If

    mp_IsEmptyVariantArray = False
    Exit Function
EH:
    mp_IsEmptyVariantArray = True
End Function
