Attribute VB_Name = "ex_Test"
Option Explicit

Public Sub m_TEST_RenderDevUI()
    Dim ws As Worksheet

    Set ws = mp_GetActiveWorksheet()
    If ws Is Nothing Then Exit Sub

    ex_SheetRenderer.m_RenderWorksheet ws, "ui\DevUI.xml"
End Sub

Public Sub m_TEST_RenderDevTableListUI()
    Dim ws As Worksheet

    Set ws = mp_GetActiveWorksheet()
    If ws Is Nothing Then Exit Sub

    If Not m_TEST_RegisterDemoTableItems() Then Exit Sub
    ex_SheetRenderer.m_RenderWorksheet ws, "ui\DevTableListUI.xml"
End Sub

Public Sub m_TEST_RenderDevPrimitiveTableUI()
    Dim ws As Worksheet

    Set ws = mp_GetActiveWorksheet()
    If ws Is Nothing Then Exit Sub

    If Not m_TEST_RegisterDemoTableItems() Then Exit Sub
    ex_SheetRenderer.m_RenderWorksheet ws, "ui\DevPrimitiveTableUI.xml"
End Sub

Public Sub m_TEST_RenderDevListTableSingleUI()
    Dim ws As Worksheet

    Set ws = mp_GetActiveWorksheet()
    If ws Is Nothing Then Exit Sub

    If Not m_TEST_RegisterDemoTableItems() Then Exit Sub
    ex_SheetRenderer.m_RenderWorksheet ws, "ui\DevListTableSingleUI.xml"
End Sub

Public Sub m_TEST_RenderDevTablePartStylesUI()
    Dim ws As Worksheet
    Dim tables As Collection

    Set ws = mp_GetActiveWorksheet()
    If ws Is Nothing Then Exit Sub

    Set tables = m_TEST_BuildDemoTableItems()
    If tables Is Nothing Then Exit Sub

    ex_ListItemsSourceRuntime.m_ResetItemsSources
    ex_ObjectSourceRuntime.m_ResetObjectSources
    If Not ex_ListItemsSourceRuntime.m_SetItemsSource("RuntimeItems.Test.Tables", tables, False) Then Exit Sub
    If Not m_TEST_RegisterDemoBannerItems(False, False) Then Exit Sub

    ex_SheetRenderer.m_RenderWorksheet ws, "ui\DevTablePartStylesUI.xml"
End Sub

Public Sub m_TEST_SetDemoTableItemsMany()
    Dim tables As Collection

    Set tables = m_TEST_BuildDemoTableItems()
    If tables Is Nothing Then Exit Sub

    If Not ex_ListItemsSourceRuntime.m_SetItemsSource("RuntimeItems.Test.Tables", tables, True) Then Exit Sub
End Sub

Public Sub m_TEST_SetDemoTableItemsSingle()
    Dim tables As Collection

    Set tables = m_TEST_BuildDemoSingleTableItems()
    If tables Is Nothing Then Exit Sub

    If Not ex_ListItemsSourceRuntime.m_SetItemsSource("RuntimeItems.Test.Tables", tables, True) Then Exit Sub
End Sub

Public Sub m_TEST_InsertDemoBanner()
    If Not m_TEST_RegisterDemoBannerItems(True, True) Then Exit Sub
End Sub

Public Sub m_TEST_ProfileDevTableListUI()
    Dim ws As Worksheet
    Dim tables As Collection
    Dim t0 As Double
    Dim t1 As Double
    Dim t2 As Double
    Dim t3 As Double

    Set ws = mp_GetActiveWorksheet()
    If ws Is Nothing Then Exit Sub

    t0 = Timer
    Set tables = m_TEST_BuildDemoTableItems()
    t1 = Timer

    If tables Is Nothing Then Exit Sub

    ex_ListItemsSourceRuntime.m_ResetItemsSources
    If Not ex_ListItemsSourceRuntime.m_SetItemsSource("RuntimeItems.Test.Tables", tables, False) Then Exit Sub
    t2 = Timer

    ex_SheetRenderer.m_RenderWorksheet ws, "ui\DevProfileTableUI.xml"
    t3 = Timer

    MsgBox "Profile (ms):" & vbCrLf & _
           "Build data: " & Format$((t1 - t0) * 1000#, "0") & vbCrLf & _
           "Register source: " & Format$((t2 - t1) * 1000#, "0") & vbCrLf & _
           "Render UI: " & Format$((t3 - t2) * 1000#, "0") & vbCrLf & _
           "Total: " & Format$((t3 - t0) * 1000#, "0"), vbInformation
End Sub

Public Sub m_TEST_FillNumbersRangeSimple()
    Dim ws As Worksheet
    Dim targetRange As Range
    Dim values() As Variant
    Dim r As Long
    Dim c As Long
    Dim n As Long

    Set ws = mp_GetActiveWorksheet()
    If ws Is Nothing Then Exit Sub

    Set targetRange = ws.Range("L5:V28")
    ReDim values(1 To targetRange.Rows.Count, 1 To targetRange.Columns.Count)

    n = 1
    For r = 1 To targetRange.Rows.Count
        For c = 1 To targetRange.Columns.Count
            values(r, c) = n
            n = n + 1
        Next c
    Next r

    targetRange.Value2 = values
End Sub

Public Sub m_TEST_RenderDevSingleTableUI()
    Dim ws As Worksheet

    Set ws = mp_GetActiveWorksheet()
    If ws Is Nothing Then Exit Sub

    If Not m_TEST_RegisterDemoSingleTableItems() Then Exit Sub
    ex_SheetRenderer.m_RenderWorksheet ws, "ui\DevSingleTableUI.xml"
End Sub

Public Function m_TEST_RegisterDemoListItems(Optional ByVal notifyChange As Boolean = False) As Boolean
    Dim items As Collection

    Set items = m_TEST_BuildDemoListItems()
    If items Is Nothing Then Exit Function

    ex_ListItemsSourceRuntime.m_ResetItemsSources
    ex_ObjectSourceRuntime.m_ResetObjectSources
    If Not ex_ListItemsSourceRuntime.m_SetItemsSource("RuntimeItems.Test.People", items, notifyChange) Then Exit Function

    m_TEST_RegisterDemoListItems = True
End Function

Public Function m_TEST_RegisterDemoTableItems(Optional ByVal notifyChange As Boolean = False) As Boolean
    Dim tables As Collection

    Set tables = m_TEST_BuildDemoTableItems()
    If tables Is Nothing Then Exit Function

    ex_ListItemsSourceRuntime.m_ResetItemsSources
    ex_ObjectSourceRuntime.m_ResetObjectSources
    If Not ex_ListItemsSourceRuntime.m_SetItemsSource("RuntimeItems.Test.Tables", tables, notifyChange) Then Exit Function

    m_TEST_RegisterDemoTableItems = True
End Function

Public Function m_TEST_RegisterDemoSingleTableItems(Optional ByVal notifyChange As Boolean = False) As Boolean
    Dim tables As Collection

    Set tables = m_TEST_BuildDemoSingleTableItems()
    If tables Is Nothing Then Exit Function

    ex_ListItemsSourceRuntime.m_ResetItemsSources
    ex_ObjectSourceRuntime.m_ResetObjectSources
    If Not ex_ListItemsSourceRuntime.m_SetItemsSource("RuntimeItems.Test.Tables", tables, notifyChange) Then Exit Function

    m_TEST_RegisterDemoSingleTableItems = True
End Function

Public Function m_TEST_RegisterDemoBannerItems( _
    Optional ByVal isVisible As Boolean = False, _
    Optional ByVal notifyChange As Boolean = False _
) As Boolean
    Dim bannerObj As obj_Banner
    Dim headerText As String
    Dim messageText As String

    If isVisible Then
        headerText = "Data Source Updated"
        messageText = "Banner was inserted before table list. Current layout was fully rerendered after objectSource update."
        Set bannerObj = mp_CreateDemoBannerModel(headerText, messageText, isVisible)
        If bannerObj Is Nothing Then Exit Function
        If Not ex_ObjectSourceRuntime.m_SetObjectSource("RuntimeObjects.Test.Banner", bannerObj, notifyChange) Then Exit Function
    Else
        If Not ex_ObjectSourceRuntime.m_RemoveObjectSource("RuntimeObjects.Test.Banner", notifyChange) Then Exit Function
    End If

    m_TEST_RegisterDemoBannerItems = True
End Function

Public Function m_TEST_BuildDemoListItems() As Collection
    Dim result As Collection

    Set result = New Collection
    result.Add mp_CreateDemoPerson("Ivan Petrov", "Team Lead")
    result.Add mp_CreateDemoPerson("Anna Sidorova", "Analyst")
    result.Add mp_CreateDemoPerson("Maksym Kovalenko", "Developer")

    Set m_TEST_BuildDemoListItems = result
End Function

Public Function m_TEST_BuildDemoTableItems() As Collection
    Dim result As Collection
    Dim tableIndex As Long
    Dim rowsCount As Long
    Dim teamName As String

    Const HEADER_TEXT As String = "Name | Role | Country | Team | Level | Status | Since"

    Set result = New Collection

    For tableIndex = 1 To 20
        rowsCount = 3 + ((tableIndex - 1) Mod 3)
        teamName = "Team " & Format$(tableIndex, "00")

        result.Add mp_CreateDemoTable( _
            "People / " & teamName, _
            HEADER_TEXT, _
            mp_CreateGeneratedRowsForTeam(tableIndex, rowsCount, teamName))
    Next tableIndex

    Set m_TEST_BuildDemoTableItems = result
End Function

Private Function mp_CreateGeneratedRowsForTeam( _
    ByVal tableIndex As Long, _
    ByVal rowsCount As Long, _
    ByVal teamName As String _
) As Collection
    Dim result As Collection
    Dim rowIndex As Long
    Dim personName As String
    Dim roleName As String
    Dim countryName As String
    Dim levelName As String
    Dim statusName As String
    Dim sinceYear As String

    Set result = New Collection

    For rowIndex = 1 To rowsCount
        personName = "Person " & Format$(tableIndex, "00") & "-" & CStr(rowIndex)
        roleName = mp_GetRoleByIndex(tableIndex + rowIndex)
        countryName = mp_GetCountryByIndex(tableIndex + rowIndex)
        levelName = "L" & CStr(((tableIndex + rowIndex) Mod 4) + 1)

        If (tableIndex + rowIndex) Mod 5 = 0 Then
            statusName = "On Hold"
        Else
            statusName = "Active"
        End If

        sinceYear = CStr(2014 + ((tableIndex + rowIndex) Mod 11))

        result.Add mp_CreateDemoRowModel( _
            personName, _
            roleName, _
            countryName, _
            teamName, _
            levelName, _
            statusName, _
            sinceYear)
    Next rowIndex

    Set mp_CreateGeneratedRowsForTeam = result
End Function

Private Function mp_GetRoleByIndex(ByVal idx As Long) As String
    Select Case ((idx - 1) Mod 7) + 1
        Case 1: mp_GetRoleByIndex = "Team Lead"
        Case 2: mp_GetRoleByIndex = "Analyst"
        Case 3: mp_GetRoleByIndex = "Developer"
        Case 4: mp_GetRoleByIndex = "QA"
        Case 5: mp_GetRoleByIndex = "DevOps"
        Case 6: mp_GetRoleByIndex = "Support"
        Case Else: mp_GetRoleByIndex = "Manager"
    End Select
End Function

Private Function mp_GetCountryByIndex(ByVal idx As Long) As String
    Select Case ((idx - 1) Mod 6) + 1
        Case 1: mp_GetCountryByIndex = "Ukraine"
        Case 2: mp_GetCountryByIndex = "Poland"
        Case 3: mp_GetCountryByIndex = "Romania"
        Case 4: mp_GetCountryByIndex = "Germany"
        Case 5: mp_GetCountryByIndex = "Czechia"
        Case Else: mp_GetCountryByIndex = "Slovakia"
    End Select
End Function

Public Function m_TEST_BuildDemoSingleTableItems() As Collection
    Dim sourceTables As Collection
    Dim result As Collection
    Dim mergedTable As obj_TableDynamic
    Dim tableObj As Variant
    Dim sourceTable As obj_TableDynamic
    Dim sourceRow As Variant
    Dim targetRow As obj_Row
    Dim sourceCol As obj_Column
    Dim i As Long

    Set sourceTables = m_TEST_BuildDemoTableItems()
    If sourceTables Is Nothing Then Exit Function

    Set mergedTable = New obj_TableDynamic
    mergedTable.SectionTitle = "People / All Teams (Merged)"

    For Each tableObj In sourceTables
        If Not mp_TryResolveDemoTableDynamic(tableObj, sourceTable) Then Exit Function

        If mergedTable.ColumnCount = 0 Then
            For Each sourceCol In sourceTable.Columns
                If Not mergedTable.m_AddColumn(sourceCol) Then Exit Function
            Next sourceCol
        End If

        For Each sourceRow In sourceTable.Rows
            If TypeName(sourceRow) <> "obj_Row" Then
                MsgBox "PrototypeNew: expected obj_Row in demo table rows.", vbExclamation
                Exit Function
            End If

            Set targetRow = New obj_Row
            For i = 1 To mergedTable.ColumnCount
                targetRow.m_AddCell sourceRow.m_GetCell(i)
            Next i

            If Not mergedTable.m_AddRow(targetRow) Then Exit Function
        Next sourceRow
    Next tableObj

    Set result = New Collection
    result.Add mergedTable

    Set m_TEST_BuildDemoSingleTableItems = result
End Function

Public Sub m_TEST_NoOp()
End Sub

Private Function mp_CreateDemoPerson(ByVal displayName As String, ByVal roleName As String) As Object
    Dim rowObj As Object

    Set rowObj = CreateObject("Scripting.Dictionary")
    rowObj.CompareMode = 1
    rowObj("Display") = CStr(displayName)
    rowObj("Role") = CStr(roleName)

    Set mp_CreateDemoPerson = rowObj
End Function

Private Function mp_CreateDemoBannerModel( _
    ByVal headerText As String, _
    ByVal messageText As String, _
    ByVal isVisible As Boolean _
) As obj_Banner
    Dim bannerObj As obj_Banner

    Set bannerObj = New obj_Banner
    bannerObj.Header = CStr(headerText)
    bannerObj.Message = CStr(messageText)
    bannerObj.Visible = CBool(isVisible)

    Set mp_CreateDemoBannerModel = bannerObj
End Function

Private Function mp_TryResolveDemoTableDynamic(ByVal tableObj As Variant, ByRef outTable As obj_TableDynamic) As Boolean
    Dim fixedTable As obj_Table
    Dim sourceCol As obj_Column
    Dim sourceRow As obj_Row
    Dim dynamicTable As obj_TableDynamic
    Dim targetCol As obj_Column
    Dim targetRow As obj_Row
    Dim i As Long

    If Not IsObject(tableObj) Then
        MsgBox "PrototypeNew: demo table item is not object.", vbExclamation
        Exit Function
    End If

    Select Case LCase$(TypeName(tableObj))
        Case "obj_tabledynamic"
            Set outTable = tableObj
            mp_TryResolveDemoTableDynamic = True

        Case "obj_table"
            Set fixedTable = tableObj
            Set dynamicTable = New obj_TableDynamic
            dynamicTable.SectionTitle = fixedTable.SectionTitle

            For Each sourceCol In fixedTable.Columns
                Set targetCol = New obj_Column
                targetCol.Name = sourceCol.Name
                targetCol.Position = sourceCol.Position
                If Not dynamicTable.m_AddColumn(targetCol) Then Exit Function
            Next sourceCol

            For Each sourceRow In fixedTable.Rows
                Set targetRow = New obj_Row
                For i = 1 To dynamicTable.ColumnCount
                    targetRow.m_AddCell sourceRow.m_GetCell(i)
                Next i
                If Not dynamicTable.m_AddRow(targetRow) Then Exit Function
            Next sourceRow

            Set outTable = dynamicTable
            mp_TryResolveDemoTableDynamic = True

        Case Else
            MsgBox "PrototypeNew: unsupported demo table type '" & TypeName(tableObj) & "'.", vbExclamation
    End Select
End Function

Private Function mp_CreateDemoTable( _
    ByVal sectionTitle As String, _
    ByVal headerText As String, _
    ByVal rows As Collection _
) As Object
    Dim tableObj As obj_TableDynamic
    Dim rowObj As obj_Row
    Dim colObj As obj_Column
    Dim headerTokens As Variant
    Dim colIndex As Long

    Set tableObj = New obj_TableDynamic
    tableObj.SectionTitle = CStr(sectionTitle)

    headerTokens = Split(CStr(headerText), "|")
    For colIndex = LBound(headerTokens) To UBound(headerTokens)
        Set colObj = New obj_Column
        colObj.Position = colIndex + 1
        colObj.Name = Trim$(CStr(headerTokens(colIndex)))
        If Len(colObj.Name) = 0 Then colObj.Name = "Col" & CStr(colObj.Position)
        If Not tableObj.m_AddColumn(colObj) Then Exit Function
    Next colIndex

    If rows Is Nothing Then
        Set mp_CreateDemoTable = tableObj
        Exit Function
    End If

    For Each rowObj In rows
        If rowObj Is Nothing Then
            MsgBox "PrototypeNew: table row is not specified.", vbExclamation
            Exit Function
        End If
        If rowObj.CellCount < tableObj.ColumnCount Then
            MsgBox "PrototypeNew: table row has fewer cells than table columns.", vbExclamation
            Exit Function
        End If

        If Not tableObj.m_AddRow(rowObj) Then Exit Function
    Next rowObj

    Set mp_CreateDemoTable = tableObj
End Function

Private Function mp_CreateDemoTableRows(ParamArray values() As Variant) As Collection
    Dim result As Collection
    Dim i As Long

    Set result = New Collection
    If (UBound(values) - LBound(values) + 1) Mod 7 <> 0 Then
        MsgBox "PrototypeNew: mp_CreateDemoTableRows expects values in septets (c1..c7).", vbExclamation
        Set mp_CreateDemoTableRows = result
        Exit Function
    End If

    For i = LBound(values) To UBound(values) Step 7
        result.Add mp_CreateDemoRowModel( _
            CStr(values(i)), _
            CStr(values(i + 1)), _
            CStr(values(i + 2)), _
            CStr(values(i + 3)), _
            CStr(values(i + 4)), _
            CStr(values(i + 5)), _
            CStr(values(i + 6)))
    Next i

    Set mp_CreateDemoTableRows = result
End Function

Private Function mp_CreateDemoRowModel( _
    ByVal c1 As String, _
    ByVal c2 As String, _
    ByVal c3 As String, _
    ByVal c4 As String, _
    ByVal c5 As String, _
    ByVal c6 As String, _
    ByVal c7 As String _
) As obj_Row
    Dim rowObj As obj_Row

    Set rowObj = New obj_Row
    rowObj.m_AddCell c1
    rowObj.m_AddCell c2
    rowObj.m_AddCell c3
    rowObj.m_AddCell c4
    rowObj.m_AddCell c5
    rowObj.m_AddCell c6
    rowObj.m_AddCell c7

    Set mp_CreateDemoRowModel = rowObj
End Function

Private Function mp_GetActiveWorksheet() As Worksheet
    Dim wb As Workbook
    Dim activeSheetObj As Object

    Set wb = ThisWorkbook
    If wb Is Nothing Then
        MsgBox "PrototypeNew: workbook is not specified.", vbExclamation
        Exit Function
    End If

    Set activeSheetObj = wb.ActiveSheet
    If activeSheetObj Is Nothing Then
        MsgBox "PrototypeNew: active sheet is not specified.", vbExclamation
        Exit Function
    End If

    If Not TypeOf activeSheetObj Is Worksheet Then
        MsgBox "PrototypeNew: active sheet is not a worksheet.", vbExclamation
        Exit Function
    End If

    Set mp_GetActiveWorksheet = activeSheetObj
End Function
