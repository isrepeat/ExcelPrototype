Attribute VB_Name = "ex_TableLayoutActions"
Option Explicit

Private Const INPUT_KEY_RESULT_TABLES As String = "__ResultTables"
Private Const PLAN_TYPE_TABLE_LAYOUT As String = "TableLayout"

Public Function m_CreatePlan(ByVal inputObject As Object) As Object
    Dim plan As Object
    Dim resultTablesObj As Object
    Dim mergeSpecs As Collection

    If inputObject Is Nothing Then
        Err.Raise vbObjectError + 6480, "ex_TableLayoutActions", "Input object is required for m_CreatePlan."
    End If

    If Not ex_ScriptIO.m_TryGetObject(inputObject, INPUT_KEY_RESULT_TABLES, resultTablesObj) Then
        Err.Raise vbObjectError + 6481, "ex_TableLayoutActions", "Input object does not contain '__ResultTables'."
    End If
    If TypeName(resultTablesObj) <> "Collection" Then
        Err.Raise vbObjectError + 6482, "ex_TableLayoutActions", "'__ResultTables' must be Collection."
    End If

    Set plan = CreateObject("Scripting.Dictionary")
    plan.CompareMode = 1
    plan("PlanType") = PLAN_TYPE_TABLE_LAYOUT
    Set plan("Input") = inputObject
    Set plan("ResultTables") = resultTablesObj

    Set mergeSpecs = New Collection
    Set plan("MergeSpecs") = mergeSpecs

    Set m_CreatePlan = plan
End Function

Public Function m_UpsertCombinedTable(ByVal plan As Object, ByVal mergeSpec As Object) As String
    Dim normalizedSpec As Object

    mp_ValidatePlan plan
    Set normalizedSpec = mp_NormalizeMergeSpec(mergeSpec)
    mp_GetPlanMergeSpecs(plan).Add normalizedSpec

    m_UpsertCombinedTable = CStr(normalizedSpec("TargetTableRef"))
End Function

Public Function m_Apply(ByVal plan As Object) As String
    Dim resultTables As Collection
    Dim inputObject As Object
    Dim mergeSpecs As Collection
    Dim i As Long

    mp_ValidatePlan plan
    Set resultTables = mp_GetPlanResultTables(plan)
    Set inputObject = mp_GetPlanInput(plan)
    Set mergeSpecs = mp_GetPlanMergeSpecs(plan)

    For i = 1 To mergeSpecs.Count
        mp_ApplyMergeSpec resultTables, mergeSpecs(i)
    Next i

    ex_ScriptIO.m_SetObject inputObject, INPUT_KEY_RESULT_TABLES, resultTables
    m_Apply = CStr(resultTables.Count)
End Function

Private Function mp_GetPlanInput(ByVal plan As Object) As Object
    mp_ValidatePlan plan
    Set mp_GetPlanInput = ex_ScriptIO.m_DictionaryGetObject(plan, "Input")
End Function

Private Function mp_GetPlanResultTables(ByVal plan As Object) As Collection
    Dim objValue As Object

    mp_ValidatePlan plan
    Set objValue = ex_ScriptIO.m_DictionaryGetObject(plan, "ResultTables")
    If TypeName(objValue) <> "Collection" Then
        Err.Raise vbObjectError + 6483, "ex_TableLayoutActions", "Plan.ResultTables must be Collection."
    End If

    Set mp_GetPlanResultTables = objValue
End Function

Private Function mp_GetPlanMergeSpecs(ByVal plan As Object) As Collection
    Dim objValue As Object

    mp_ValidatePlan plan
    Set objValue = ex_ScriptIO.m_DictionaryGetObject(plan, "MergeSpecs")
    If TypeName(objValue) <> "Collection" Then
        Err.Raise vbObjectError + 6484, "ex_TableLayoutActions", "Plan.MergeSpecs must be Collection."
    End If

    Set mp_GetPlanMergeSpecs = objValue
End Function

Private Sub mp_ValidatePlan(ByVal plan As Object)
    If plan Is Nothing Then
        Err.Raise vbObjectError + 6485, "ex_TableLayoutActions", "Plan is required."
    End If
    If TypeName(plan) <> "Dictionary" And TypeName(plan) <> "Scripting.Dictionary" Then
        Err.Raise vbObjectError + 6486, "ex_TableLayoutActions", "Plan must be Dictionary."
    End If
    If Not plan.Exists("PlanType") Then
        Err.Raise vbObjectError + 6487, "ex_TableLayoutActions", "Invalid plan object: missing PlanType."
    End If
    If StrComp(CStr(plan("PlanType")), PLAN_TYPE_TABLE_LAYOUT, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 6488, "ex_TableLayoutActions", "Invalid plan type: expected '" & PLAN_TYPE_TABLE_LAYOUT & "'."
    End If
End Sub

Private Function mp_NormalizeMergeSpec(ByVal mergeSpec As Object) As Object
    Dim spec As Object
    Dim targetTableRef As String
    Dim sourcesRaw As String
    Dim columnsRaw As String
    Dim sortBy As String
    Dim sortOrder As String

    If mergeSpec Is Nothing Then
        Err.Raise vbObjectError + 6489, "ex_TableLayoutActions", "Merge spec is required."
    End If
    If TypeName(mergeSpec) <> "Dictionary" And TypeName(mergeSpec) <> "Scripting.Dictionary" Then
        Err.Raise vbObjectError + 6490, "ex_TableLayoutActions", "Merge spec must be Dictionary."
    End If

    targetTableRef = Trim$(ex_ScriptIO.m_GetStringOrDefault(mergeSpec, "TargetTableRef", vbNullString))
    If Len(targetTableRef) = 0 Then
        Err.Raise vbObjectError + 6491, "ex_TableLayoutActions", "Merge spec key 'TargetTableRef' is required."
    End If

    sourcesRaw = Trim$(ex_ScriptIO.m_GetStringOrDefault(mergeSpec, "Sources", vbNullString))
    If Len(sourcesRaw) = 0 Then
        Err.Raise vbObjectError + 6492, "ex_TableLayoutActions", "Merge spec key 'Sources' is required."
    End If

    columnsRaw = Trim$(ex_ScriptIO.m_GetStringOrDefault(mergeSpec, "Columns", vbNullString))
    sortBy = Trim$(ex_ScriptIO.m_GetStringOrDefault(mergeSpec, "SortBy", vbNullString))
    sortOrder = LCase$(Trim$(ex_ScriptIO.m_GetStringOrDefault(mergeSpec, "SortOrder", "asc")))

    If Len(sortOrder) = 0 Then sortOrder = "asc"
    If sortOrder <> "asc" And sortOrder <> "desc" Then
        Err.Raise vbObjectError + 6493, "ex_TableLayoutActions", "Merge spec key 'SortOrder' must be 'asc' or 'desc'."
    End If

    Set spec = CreateObject("Scripting.Dictionary")
    spec.CompareMode = 1
    spec("TargetTableRef") = targetTableRef
    spec("Sources") = sourcesRaw
    spec("Columns") = columnsRaw
    spec("SortBy") = sortBy
    spec("SortOrder") = sortOrder

    Set mp_NormalizeMergeSpec = spec
End Function

Private Sub mp_ApplyMergeSpec(ByVal resultTables As Collection, ByVal spec As Object)
    Dim targetTableRef As String
    Dim sources As Collection
    Dim columns As Collection
    Dim sortBy As String
    Dim sortDesc As Boolean
    Dim sourceTable As obj_ResultTable
    Dim sourceRef As Variant
    Dim rowObj As obj_ResultRow
    Dim rowData As Object
    Dim mergedRows As Collection
    Dim sortedRows As Collection
    Dim targetTable As obj_ResultTable
    Dim colAlias As Variant
    Dim rowIndex As Long
    Dim mapKey As String
    Dim valueText As String

    targetTableRef = CStr(spec("TargetTableRef"))
    sortBy = Trim$(CStr(spec("SortBy")))
    sortDesc = (StrComp(CStr(spec("SortOrder")), "desc", vbTextCompare) = 0)

    Set sources = mp_ParseSemicolonList(CStr(spec("Sources")))
    Set columns = mp_ParseSemicolonList(CStr(spec("Columns")))
    Set mergedRows = New Collection

    For Each sourceRef In sources
        If mp_TryFindTableByRef(resultTables, CStr(sourceRef), sourceTable) Then
            If columns.Count = 0 Then
                mp_EnsureColumnsFromTable columns, sourceTable
            End If

            For Each rowObj In sourceTable.Rows
                Set rowData = CreateObject("Scripting.Dictionary")
                rowData.CompareMode = 1

                For Each colAlias In columns
                    valueText = vbNullString
                    If rowObj.HasAlias(CStr(colAlias)) Then
                        valueText = rowObj.Column(CStr(colAlias))
                    End If
                    rowData(CStr(colAlias)) = valueText
                Next colAlias

                mergedRows.Add rowData
            Next rowObj
        End If
    Next sourceRef

    Set sortedRows = mp_SortRows(mergedRows, sortBy, sortDesc)

    Set targetTable = New obj_ResultTable
    targetTable.Initialize targetTableRef

    For Each colAlias In columns
        mapKey = CStr(colAlias)
        targetTable.AddFieldMap CStr(colAlias), mapKey
    Next colAlias

    rowIndex = 0
    For Each rowData In sortedRows
        For Each colAlias In columns
            valueText = ex_ScriptIO.m_DictionaryGetStringOrDefault(rowData, CStr(colAlias), vbNullString)
            targetTable.SetRowValue rowIndex, CStr(colAlias), CStr(colAlias), valueText
        Next colAlias
        rowIndex = rowIndex + 1
    Next rowData

    mp_UpsertTable resultTables, targetTable
End Sub

Private Sub mp_EnsureColumnsFromTable(ByVal columns As Collection, ByVal sourceTable As obj_ResultTable)
    Dim key As Variant
    Dim firstRow As obj_ResultRow
    Dim columnObj As obj_ResultColumn

    If sourceTable Is Nothing Then Exit Sub

    If Not sourceTable.FieldMapByAlias Is Nothing Then
        For Each key In sourceTable.FieldMapByAlias.Keys
            columns.Add CStr(key)
        Next key
    End If

    If columns.Count > 0 Then Exit Sub
    If sourceTable.Count <= 0 Then Exit Sub

    Set firstRow = sourceTable.Rows(1)
    If firstRow Is Nothing Then Exit Sub

    For Each columnObj In firstRow.Columns
        columns.Add columnObj.Alias
    Next columnObj
End Sub

Private Function mp_ParseSemicolonList(ByVal rawText As String) As Collection
    Dim result As Collection
    Dim parts As Variant
    Dim i As Long
    Dim token As String

    Set result = New Collection
    rawText = Trim$(rawText)
    If Len(rawText) = 0 Then
        Set mp_ParseSemicolonList = result
        Exit Function
    End If

    parts = Split(rawText, ";")
    For i = LBound(parts) To UBound(parts)
        token = Trim$(CStr(parts(i)))
        If Len(token) > 0 Then result.Add token
    Next i

    Set mp_ParseSemicolonList = result
End Function

Private Function mp_SortRows(ByVal rows As Collection, ByVal sortBy As String, ByVal sortDesc As Boolean) As Collection
    Dim result As Collection
    Dim arr() As Variant
    Dim i As Long
    Dim j As Long
    Dim tmp As Variant

    Set result = New Collection
    If rows Is Nothing Then
        Set mp_SortRows = result
        Exit Function
    End If

    If rows.Count = 0 Then
        Set mp_SortRows = result
        Exit Function
    End If

    If Len(Trim$(sortBy)) = 0 Then
        For i = 1 To rows.Count
            result.Add rows(i)
        Next i
        Set mp_SortRows = result
        Exit Function
    End If

    ReDim arr(1 To rows.Count)
    For i = 1 To rows.Count
        Set arr(i) = rows(i)
    Next i

    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If mp_ShouldSwapRows(arr(i), arr(j), sortBy, sortDesc) Then
                Set tmp = arr(i)
                Set arr(i) = arr(j)
                Set arr(j) = tmp
            End If
        Next j
    Next i

    For i = 1 To UBound(arr)
        result.Add arr(i)
    Next i

    Set mp_SortRows = result
End Function

Private Function mp_ShouldSwapRows(ByVal leftRow As Object, ByVal rightRow As Object, ByVal sortBy As String, ByVal sortDesc As Boolean) As Boolean
    Dim leftValue As String
    Dim rightValue As String
    Dim cmp As Long

    leftValue = ex_ScriptIO.m_DictionaryGetStringOrDefault(leftRow, sortBy, vbNullString)
    rightValue = ex_ScriptIO.m_DictionaryGetStringOrDefault(rightRow, sortBy, vbNullString)

    cmp = mp_CompareSortValues(leftValue, rightValue)
    If sortDesc Then
        mp_ShouldSwapRows = (cmp < 0)
    Else
        mp_ShouldSwapRows = (cmp > 0)
    End If
End Function

Private Function mp_CompareSortValues(ByVal leftValue As String, ByVal rightValue As String) As Long
    Dim leftNum As Double
    Dim rightNum As Double

    If IsDate(leftValue) And IsDate(rightValue) Then
        If CDate(leftValue) < CDate(rightValue) Then
            mp_CompareSortValues = -1
        ElseIf CDate(leftValue) > CDate(rightValue) Then
            mp_CompareSortValues = 1
        Else
            mp_CompareSortValues = 0
        End If
        Exit Function
    End If

    If mp_TryParseDouble(leftValue, leftNum) And mp_TryParseDouble(rightValue, rightNum) Then
        If leftNum < rightNum Then
            mp_CompareSortValues = -1
        ElseIf leftNum > rightNum Then
            mp_CompareSortValues = 1
        Else
            mp_CompareSortValues = 0
        End If
        Exit Function
    End If

    mp_CompareSortValues = StrComp(leftValue, rightValue, vbTextCompare)
End Function

Private Function mp_TryParseDouble(ByVal rawText As String, ByRef outValue As Double) As Boolean
    On Error GoTo ParseFail
    If Len(Trim$(rawText)) = 0 Then Exit Function
    outValue = CDbl(rawText)
    mp_TryParseDouble = True
    Exit Function

ParseFail:
    mp_TryParseDouble = False
End Function

Private Sub mp_UpsertTable(ByVal resultTables As Collection, ByVal tableObj As obj_ResultTable)
    Dim i As Long
    Dim existing As obj_ResultTable

    If resultTables Is Nothing Then Exit Sub
    If tableObj Is Nothing Then Exit Sub

    For i = 1 To resultTables.Count
        Set existing = resultTables(i)
        If Not existing Is Nothing Then
            If StrComp(existing.TableRef, tableObj.TableRef, vbTextCompare) = 0 Then
                resultTables.Remove i
                If i > resultTables.Count Then
                    resultTables.Add tableObj
                Else
                    resultTables.Add tableObj, Before:=i
                End If
                Exit Sub
            End If
        End If
    Next i

    resultTables.Add tableObj
End Sub

Private Function mp_TryFindTableByRef( _
    ByVal resultTables As Collection, _
    ByVal tableRef As String, _
    ByRef outTable As obj_ResultTable _
) As Boolean
    Dim i As Long
    Dim candidate As obj_ResultTable

    tableRef = Trim$(tableRef)
    If Len(tableRef) = 0 Then Exit Function
    If resultTables Is Nothing Then Exit Function

    For i = 1 To resultTables.Count
        Set candidate = resultTables(i)
        If Not candidate Is Nothing Then
            If StrComp(candidate.TableRef, tableRef, vbTextCompare) = 0 Then
                Set outTable = candidate
                mp_TryFindTableByRef = True
                Exit Function
            End If
        End If
    Next i
End Function
