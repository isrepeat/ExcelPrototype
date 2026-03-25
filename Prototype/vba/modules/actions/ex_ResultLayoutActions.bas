Attribute VB_Name = "ex_ResultLayoutActions"
Option Explicit

Private Const INPUT_KEY_RESULT_TABLES As String = "__ResultTables"
Private Const INPUT_KEY_LAYOUT_SHEET_NAME As String = "__ResultLayoutSheetName"
Private Const INPUT_KEY_LAYOUT_WORKSHEET As String = "__ResultLayoutWorksheet"
Private Const INPUT_KEY_LAYOUT_ROWKINDS As String = "__ResultLayoutRowKinds"
Private Const INPUT_KEY_LAYOUT_FIELDRANGES As String = "__ResultLayoutFieldRanges"
Private Const PLAN_TYPE_RESULT_LAYOUT As String = "ResultLayout"

Private g_HeaderCaptionByMapKey As Object

Public Function m_CreatePlan(ByVal inputObject As Object) As Object
    Dim plan As Object
    Dim resultTablesObj As Object
    Dim items As Collection

    If inputObject Is Nothing Then
        Err.Raise vbObjectError + 6460, "ex_ResultLayoutActions", "Input object is required for m_CreatePlan."
    End If

    If Not ex_ScriptIO.m_TryGetObject(inputObject, INPUT_KEY_RESULT_TABLES, resultTablesObj) Then
        Err.Raise vbObjectError + 6461, "ex_ResultLayoutActions", "Input object does not contain '__ResultTables'."
    End If
    If TypeName(resultTablesObj) <> "Collection" Then
        Err.Raise vbObjectError + 6462, "ex_ResultLayoutActions", "'__ResultTables' must be Collection."
    End If

    Set plan = CreateObject("Scripting.Dictionary")
    plan.CompareMode = 1
    plan("PlanType") = PLAN_TYPE_RESULT_LAYOUT
    Set plan("Input") = inputObject
    Set plan("ResultTables") = resultTablesObj

    Set items = New Collection
    Set plan("Items") = items

    Set m_CreatePlan = plan
End Function

Public Function m_PushTableIfNotEmpty(ByVal plan As Object, ByVal tableRef As String) As String
    Dim resultTable As obj_ResultTable

    tableRef = Trim$(tableRef)
    If Len(tableRef) = 0 Then Exit Function

    If mp_TryFindTableByRef(mp_GetPlanResultTables(plan), tableRef, resultTable) Then
        If resultTable.Count > 0 Then
            mp_AddTableItem mp_GetPlanItems(plan), tableRef
        End If
    End If

    m_PushTableIfNotEmpty = tableRef
End Function

Public Function m_PushTable(ByVal plan As Object, ByVal tableRef As String) As String
    Dim resultTable As obj_ResultTable

    tableRef = Trim$(tableRef)
    If Len(tableRef) = 0 Then
        Err.Raise vbObjectError + 6463, "ex_ResultLayoutActions", "TableRef is required for m_PushTable."
    End If

    If Not mp_TryFindTableByRef(mp_GetPlanResultTables(plan), tableRef, resultTable) Then
        Err.Raise vbObjectError + 6464, "ex_ResultLayoutActions", "Table was not found in result set: " & tableRef
    End If

    mp_AddTableItem mp_GetPlanItems(plan), tableRef
    m_PushTable = tableRef
End Function

Public Function m_PushAllTablesIfNotEmpty(ByVal plan As Object) As String
    Dim resultTables As Collection
    Dim i As Long
    Dim resultTable As obj_ResultTable
    Dim pushedCount As Long

    Set resultTables = mp_GetPlanResultTables(plan)

    For i = 1 To resultTables.Count
        Set resultTable = resultTables(i)
        If Not resultTable Is Nothing Then
            If resultTable.Count > 0 Then
                mp_AddTableItem mp_GetPlanItems(plan), resultTable.TableRef
                pushedCount = pushedCount + 1
            End If
        End If
    Next i

    m_PushAllTablesIfNotEmpty = CStr(pushedCount)
End Function

Public Function m_PushSpacer(ByVal plan As Object, ByVal spacerRows As Variant) As String
    Dim rowsCount As Long

    rowsCount = mp_ParseNonNegativeLongOrRaise(spacerRows, "m_PushSpacer")
    If rowsCount <= 0 Then
        m_PushSpacer = "0"
        Exit Function
    End If

    mp_AddSpacerItem mp_GetPlanItems(plan), rowsCount
    m_PushSpacer = CStr(rowsCount)
End Function

Public Function m_Apply(ByVal plan As Object) As String
    Dim inputObject As Object
    Dim resultTables As Collection
    Dim items As Collection
    Dim ws As Worksheet
    Dim sheetName As String
    Dim startRow As Long
    Dim rowIndex As Long
    Dim pendingSpacerRows As Long
    Dim lastRenderedTable As Boolean
    Dim i As Long
    Dim item As Object
    Dim itemKind As String
    Dim tableRef As String
    Dim spacerRows As Long
    Dim resultTable As obj_ResultTable
    Dim tablesByRef As Object
    Dim resultFieldRanges As Collection
    Dim headerRows As Collection
    Dim sectionRows As Collection
    Dim contentRows As Collection
    Dim rowKinds As Object

    Set inputObject = mp_GetPlanInput(plan)
    Set resultTables = mp_GetPlanResultTables(plan)
    Set items = mp_GetPlanItems(plan)

    sheetName = ex_ScriptIO.m_GetStringOrDefault(inputObject, INPUT_KEY_LAYOUT_SHEET_NAME, vbNullString)
    If Len(sheetName) = 0 Then sheetName = "g_ResultLayout"

    Set ws = mp_CreateOrClearSheet(sheetName)
    ex_Messaging.m_ClearBannerAnchors ws
    ex_Messaging.m_ClearResultTableAnchors ws
    ex_Messaging.m_ClearResultRowAnchors ws
    ws.Cells.NumberFormat = "@"

    startRow = mp_GetOutputStartRow()
    If startRow < 1 Then startRow = 1
    rowIndex = startRow

    Set tablesByRef = mp_BuildTablesByRef(resultTables)
    Set resultFieldRanges = New Collection

    Set headerRows = New Collection
    Set sectionRows = New Collection
    Set contentRows = New Collection

    pendingSpacerRows = 0
    lastRenderedTable = False

    For i = 1 To items.Count
        Set item = items(i)
        itemKind = LCase$(ex_ScriptIO.m_DictionaryGetStringOrDefault(item, "Kind", vbNullString))

        Select Case itemKind
            Case "spacer"
                spacerRows = CLng(Val(ex_ScriptIO.m_DictionaryGetStringOrDefault(item, "Rows", "0")))
                If spacerRows > 0 And lastRenderedTable Then
                    pendingSpacerRows = pendingSpacerRows + spacerRows
                End If

            Case "table"
                tableRef = ex_ScriptIO.m_DictionaryGetStringOrDefault(item, "TableRef", vbNullString)
                If Len(tableRef) = 0 Then GoTo ContinueItem
                If Not mp_TryGetTableByRef(tablesByRef, tableRef, resultTable) Then GoTo ContinueItem
                If resultTable Is Nothing Then GoTo ContinueItem
                If resultTable.Count <= 0 Then GoTo ContinueItem

                If pendingSpacerRows > 0 And lastRenderedTable Then
                    rowIndex = rowIndex + pendingSpacerRows
                End If
                pendingSpacerRows = 0

                rowIndex = mp_RenderTableBlock(ws, resultTable, rowIndex, headerRows, sectionRows, contentRows, resultFieldRanges)
                lastRenderedTable = True
        End Select
ContinueItem:
    Next i

    Set rowKinds = CreateObject("Scripting.Dictionary")
    rowKinds.CompareMode = 1
    Set rowKinds("header") = headerRows
    Set rowKinds("section") = sectionRows
    Set rowKinds("content") = contentRows

    ex_ScriptIO.m_SetObject inputObject, INPUT_KEY_RESULT_TABLES, resultTables
    ex_ScriptIO.m_SetObject inputObject, INPUT_KEY_LAYOUT_WORKSHEET, ws
    ex_ScriptIO.m_SetObject inputObject, INPUT_KEY_LAYOUT_ROWKINDS, rowKinds
    ex_ScriptIO.m_SetObject inputObject, INPUT_KEY_LAYOUT_FIELDRANGES, resultFieldRanges

    ws.Activate
    m_Apply = ws.Name
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
        Err.Raise vbObjectError + 6465, "ex_ResultLayoutActions", "Plan.ResultTables must be Collection."
    End If

    Set mp_GetPlanResultTables = objValue
End Function

Private Function mp_GetPlanItems(ByVal plan As Object) As Collection
    Dim objValue As Object

    mp_ValidatePlan plan
    Set objValue = ex_ScriptIO.m_DictionaryGetObject(plan, "Items")
    If TypeName(objValue) <> "Collection" Then
        Err.Raise vbObjectError + 6466, "ex_ResultLayoutActions", "Plan.Items must be Collection."
    End If

    Set mp_GetPlanItems = objValue
End Function

Private Sub mp_ValidatePlan(ByVal plan As Object)
    If plan Is Nothing Then
        Err.Raise vbObjectError + 6467, "ex_ResultLayoutActions", "Plan is required."
    End If
    If TypeName(plan) <> "Dictionary" And TypeName(plan) <> "Scripting.Dictionary" Then
        Err.Raise vbObjectError + 6468, "ex_ResultLayoutActions", "Plan must be Dictionary."
    End If
    If Not plan.Exists("PlanType") Then
        Err.Raise vbObjectError + 6469, "ex_ResultLayoutActions", "Invalid plan object: missing PlanType."
    End If
    If StrComp(CStr(plan("PlanType")), PLAN_TYPE_RESULT_LAYOUT, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 6470, "ex_ResultLayoutActions", "Invalid plan type: expected '" & PLAN_TYPE_RESULT_LAYOUT & "'."
    End If
End Sub

Private Sub mp_AddTableItem(ByVal items As Collection, ByVal tableRef As String)
    Dim item As Object

    Set item = CreateObject("Scripting.Dictionary")
    item.CompareMode = 1
    item("Kind") = "table"
    item("TableRef") = Trim$(tableRef)
    items.Add item
End Sub

Private Sub mp_AddSpacerItem(ByVal items As Collection, ByVal rowsCount As Long)
    Dim item As Object

    Set item = CreateObject("Scripting.Dictionary")
    item.CompareMode = 1
    item("Kind") = "spacer"
    item("Rows") = CStr(rowsCount)
    items.Add item
End Sub

Private Function mp_RenderTableBlock( _
    ByVal ws As Worksheet, _
    ByVal resultTable As obj_ResultTable, _
    ByVal startRow As Long, _
    ByVal headerRows As Collection, _
    ByVal sectionRows As Collection, _
    ByVal contentRows As Collection, _
    ByVal resultFieldRanges As Collection _
) As Long
    Dim rowIndex As Long
    Dim sectionRow As Long
    Dim tableEndRow As Long
    Dim fieldAliases As Collection
    Dim rowObj As obj_ResultRow
    Dim fieldAlias As String
    Dim valueText As String
    Dim i As Long
    Dim colIndex As Long
    Dim rowAnchorName As String

    sectionRow = startRow
    ws.Cells(sectionRow, 1).Value = resultTable.TableRef
    ws.Cells(sectionRow, 1).Font.Bold = True
    sectionRows.Add sectionRow

    rowIndex = sectionRow + 1
    Set fieldAliases = mp_GetFieldAliasOrder(resultTable)

    If fieldAliases.Count > 0 Then
        For colIndex = 1 To fieldAliases.Count
            fieldAlias = CStr(fieldAliases(colIndex))
            ws.Cells(rowIndex, colIndex).Value = mp_GetFieldHeaderCaption(resultTable, fieldAlias)
        Next colIndex
        headerRows.Add rowIndex
        rowIndex = rowIndex + 1
    End If

    For i = 1 To resultTable.Rows.Count
        Set rowObj = resultTable.Rows(i)
        For colIndex = 1 To fieldAliases.Count
            fieldAlias = CStr(fieldAliases(colIndex))
            valueText = vbNullString
            If rowObj.HasAlias(fieldAlias) Then
                valueText = rowObj.Column(fieldAlias)
            End If
            ws.Cells(rowIndex, colIndex).Value = valueText
        Next colIndex

        contentRows.Add rowIndex

        rowAnchorName = ex_Messaging.m_BuildResultRowAnchorName(resultTable.TableRef, i)
        If Len(rowAnchorName) > 0 Then
            rowObj.RowAnchorName = rowAnchorName
            ex_Messaging.m_RegisterResultRowAnchor ws, rowAnchorName, rowIndex
        End If

        rowIndex = rowIndex + 1
    Next i

    tableEndRow = rowIndex - 1
    If tableEndRow < sectionRow Then tableEndRow = sectionRow
    ex_Messaging.m_RegisterResultTableAnchor ws, resultTable.TableRef, sectionRow, tableEndRow
    mp_AddRenderedFieldRanges resultFieldRanges, resultTable, fieldAliases, sectionRow + 1, tableEndRow

    mp_RenderTableBlock = rowIndex
End Function

Private Sub mp_AddRenderedFieldRanges( _
    ByVal resultFieldRanges As Collection, _
    ByVal resultTable As obj_ResultTable, _
    ByVal fieldAliases As Collection, _
    ByVal rowStart As Long, _
    ByVal rowEnd As Long _
)
    Dim i As Long
    Dim fieldAlias As String
    Dim mapKey As String
    Dim target As Object

    If resultFieldRanges Is Nothing Then Exit Sub
    If resultTable Is Nothing Then Exit Sub
    If fieldAliases Is Nothing Then Exit Sub
    If fieldAliases.Count = 0 Then Exit Sub
    If rowStart <= 0 Then Exit Sub
    If rowEnd < rowStart Then rowEnd = rowStart

    For i = 1 To fieldAliases.Count
        fieldAlias = Trim$(CStr(fieldAliases(i)))
        If Len(fieldAlias) = 0 Then GoTo ContinueField

        mapKey = fieldAlias
        On Error Resume Next
        If resultTable.HasFieldAlias(fieldAlias) Then
            mapKey = resultTable.MapKeyByAlias(fieldAlias)
        End If
        On Error GoTo 0

        Set target = CreateObject("Scripting.Dictionary")
        target.CompareMode = 1
        target("MapKey") = CStr(mapKey)
        target("ColumnIndex") = CLng(i)
        target("RowStart") = CLng(rowStart)
        target("RowEnd") = CLng(rowEnd)

        resultFieldRanges.Add target
ContinueField:
    Next i
End Sub

Private Function mp_BuildTablesByRef(ByVal resultTables As Collection) As Object
    Dim result As Object
    Dim i As Long
    Dim tableObj As obj_ResultTable

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1

    If Not resultTables Is Nothing Then
        For i = 1 To resultTables.Count
            Set tableObj = resultTables(i)
            If Not tableObj Is Nothing Then
                If Len(Trim$(tableObj.TableRef)) > 0 Then
                    Set result(tableObj.TableRef) = tableObj
                End If
            End If
        Next i
    End If

    Set mp_BuildTablesByRef = result
End Function

Private Function mp_TryGetTableByRef( _
    ByVal tablesByRef As Object, _
    ByVal tableRef As String, _
    ByRef outTable As obj_ResultTable _
) As Boolean
    tableRef = Trim$(tableRef)
    If Len(tableRef) = 0 Then Exit Function
    If tablesByRef Is Nothing Then Exit Function
    If Not tablesByRef.Exists(tableRef) Then Exit Function

    Set outTable = tablesByRef(tableRef)
    mp_TryGetTableByRef = Not (outTable Is Nothing)
End Function

Private Function mp_GetFieldHeaderCaption(ByVal resultTable As obj_ResultTable, ByVal fieldAlias As String) As String
    Dim mapKey As String
    Dim caption As String

    fieldAlias = Trim$(fieldAlias)
    If Len(fieldAlias) = 0 Then Exit Function

    On Error Resume Next
    If Not resultTable Is Nothing Then
        If resultTable.HasFieldAlias(fieldAlias) Then
            mapKey = Trim$(resultTable.MapKeyByAlias(fieldAlias))
        End If
    End If
    On Error GoTo 0

    If Len(mapKey) > 0 Then
        caption = mp_GetCaptionByMapKey(mapKey)
        If Len(caption) > 0 Then
            mp_GetFieldHeaderCaption = caption
            Exit Function
        End If
    End If

    mp_GetFieldHeaderCaption = fieldAlias
End Function

Private Function mp_GetCaptionByMapKey(ByVal mapKey As String) As String
    Dim rawValue As String

    mapKey = Trim$(mapKey)
    If Len(mapKey) = 0 Then Exit Function

    mp_EnsureHeaderCaptionCache
    If g_HeaderCaptionByMapKey.Exists(mapKey) Then
        mp_GetCaptionByMapKey = CStr(g_HeaderCaptionByMapKey(mapKey))
        Exit Function
    End If

    rawValue = Trim$(ex_ConfigProvider.m_GetConfigValue(mapKey, vbNullString))
    mp_GetCaptionByMapKey = mp_ExtractDisplayCaption(rawValue)

    g_HeaderCaptionByMapKey(mapKey) = mp_GetCaptionByMapKey
End Function

Private Function mp_ExtractDisplayCaption(ByVal rawValue As String) As String
    Dim parts As Variant
    Dim i As Long
    Dim token As String

    rawValue = Trim$(rawValue)
    If Len(rawValue) = 0 Then Exit Function

    parts = Split(rawValue, "|")
    For i = UBound(parts) To LBound(parts) Step -1
        token = Trim$(CStr(parts(i)))
        If Len(token) > 0 Then
            mp_ExtractDisplayCaption = token
            Exit Function
        End If
    Next i

    mp_ExtractDisplayCaption = rawValue
End Function

Private Sub mp_EnsureHeaderCaptionCache()
    If g_HeaderCaptionByMapKey Is Nothing Then
        Set g_HeaderCaptionByMapKey = CreateObject("Scripting.Dictionary")
        g_HeaderCaptionByMapKey.CompareMode = 1
    End If
End Sub

Private Function mp_GetFieldAliasOrder(ByVal resultTable As obj_ResultTable) As Collection
    Dim result As Collection
    Dim fieldMap As Object
    Dim key As Variant
    Dim firstRow As obj_ResultRow
    Dim columns As Collection
    Dim columnObj As obj_ResultColumn

    Set result = New Collection
    If resultTable Is Nothing Then
        Set mp_GetFieldAliasOrder = result
        Exit Function
    End If

    Set fieldMap = resultTable.FieldMapByAlias
    If Not fieldMap Is Nothing Then
        For Each key In fieldMap.Keys
            result.Add CStr(key)
        Next key
    End If

    If result.Count = 0 And resultTable.Count > 0 Then
        Set firstRow = resultTable.Rows(1)
        Set columns = firstRow.Columns
        If Not columns Is Nothing Then
            For Each columnObj In columns
                result.Add columnObj.Alias
            Next columnObj
        End If
    End If

    Set mp_GetFieldAliasOrder = result
End Function

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

Private Function mp_CreateOrClearSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    sheetName = Trim$(sheetName)
    If Len(sheetName) = 0 Then
        Err.Raise vbObjectError + 6471, "ex_ResultLayoutActions", "Sheet name is required for result layout apply."
    End If

    If mp_WorksheetExists(sheetName) Then
        Set ws = ThisWorkbook.Worksheets(sheetName)
        ws.Cells.Clear
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    Set mp_CreateOrClearSheet = ws
End Function

Private Function mp_WorksheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    mp_WorksheetExists = Not ws Is Nothing
    Set ws = Nothing
    On Error GoTo 0
End Function

Private Function mp_GetOutputStartRow() As Long
    Dim outputStyle As ex_SheetStylesXmlProvider.t_OutputSheetStyle

    If ex_SheetStylesXmlProvider.m_GetOutputSheetStyle(outputStyle, ThisWorkbook) Then
        mp_GetOutputStartRow = ex_SheetStylesXmlProvider.m_GetOutputViewStartRow(ThisWorkbook)
    Else
        mp_GetOutputStartRow = 1
    End If
End Function

Private Function mp_ParseNonNegativeLongOrRaise(ByVal rawValue As Variant, ByVal methodName As String) As Long
    If Not mp_TryParseNonNegativeLong(rawValue, mp_ParseNonNegativeLongOrRaise) Then
        Err.Raise vbObjectError + 6472, "ex_ResultLayoutActions", _
            "Invalid non-negative integer for " & methodName & ": '" & CStr(rawValue) & "'."
    End If
End Function

Private Function mp_TryParseNonNegativeLong(ByVal rawValue As Variant, ByRef outValue As Long) As Boolean
    Dim textValue As String
    Dim i As Long
    Dim ch As String

    textValue = Trim$(CStr(rawValue))
    If Len(textValue) = 0 Then Exit Function

    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i

    On Error GoTo ParseFail
    outValue = CLng(textValue)
    If outValue < 0 Then Exit Function

    mp_TryParseNonNegativeLong = True
    Exit Function

ParseFail:
    mp_TryParseNonNegativeLong = False
End Function
