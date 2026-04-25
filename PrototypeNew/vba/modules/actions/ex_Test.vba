Attribute VB_Name = "ex_Test"
Option Explicit

Private Const DEMO_CONFIG_VARIANT_A As String = "hospitalizationdate"
Private Const DEMO_CONFIG_VARIANT_B As String = "transfersheet"
Private g_DemoConfigVariant As String

' //
' // API
' //
Public Sub m_TEST_RenderDevUI()
    Dim ws As Worksheet

    Set ws = private_GetActiveWorksheet()
    If ws Is Nothing Then Exit Sub

    private_RenderWorksheetPage ws, "ui\DevUI.xml"
End Sub


Public Sub m_TEST_UpdateCurrentPage()
    If Not ex_HelpersSheet.m_TryRerenderActivePage("manual:update-sheet") Then
        rt_Messaging.m_ShowStatusBarWarning "No rendered page context is available for update.", 5
    End If
End Sub


Public Sub m_TEST_RenderDevTableListUI()
    Dim ws As Worksheet

    Set ws = private_GetActiveWorksheet()
    If ws Is Nothing Then Exit Sub

    private_RenderWorksheetPage ws, "ui\DevTableListUI.xml"
End Sub


Public Sub m_TEST_RenderDevPrimitiveTableUI()
    Dim ws As Worksheet

    Set ws = private_GetActiveWorksheet()
    If ws Is Nothing Then Exit Sub

    private_RenderWorksheetPage ws, "ui\DevPrimitiveTableUI.xml"
End Sub


Public Sub m_TEST_RenderDevListTableSingleUI()
    Dim ws As Worksheet

    Set ws = private_GetActiveWorksheet()
    If ws Is Nothing Then Exit Sub

    private_RenderWorksheetPage ws, "ui\DevListTableSingleUI.xml"
End Sub


Public Sub m_TEST_RenderDevTablePartStylesUI()
    Dim ws As Worksheet

    Set ws = private_GetActiveWorksheet()
    If ws Is Nothing Then Exit Sub

    private_RenderWorksheetPage ws, "ui\DevTablePartStylesUI.xml"
End Sub


Public Sub m_TEST_SetDemoTableItemsMany()
    Dim tableViews As Collection

    Set tableViews = m_TEST_BuildDemoTableViewItems(True, True)
    If tableViews Is Nothing Then Exit Sub

    If Not private_TrySetItemsSource("RuntimeItems.Test.Tables", tableViews, True) Then Exit Sub
End Sub


Public Sub m_TEST_SetDemoTableItemsSingle()
    Dim tableViews As Collection

    Set tableViews = m_TEST_BuildDemoSingleTableViewItems(True, True)
    If tableViews Is Nothing Then Exit Sub

    If Not private_TrySetItemsSource("RuntimeItems.Test.Tables", tableViews, True) Then Exit Sub
End Sub


Public Sub m_TEST_InsertDemoBanner()
    If Not m_TEST_RegisterDemoBannerItems(True, True) Then Exit Sub
End Sub


Public Sub m_TEST_SetDemoConfigVariantA()
    g_DemoConfigVariant = DEMO_CONFIG_VARIANT_A
    If Not m_TEST_RegisterDemoConfigProfileItems(False) Then Exit Sub
    If Not m_TEST_RegisterDemoConfigItemsVariantA(False) Then Exit Sub
    If Not private_TrySaveDemoConfigVariantToStoreForActiveSheet(g_DemoConfigVariant) Then Exit Sub
    Call ex_HelpersSheet.m_TryRerenderActivePage("configVariant:" & g_DemoConfigVariant)
End Sub


Public Sub m_TEST_SetDemoConfigVariantB()
    g_DemoConfigVariant = DEMO_CONFIG_VARIANT_B
    If Not m_TEST_RegisterDemoConfigProfileItems(False) Then Exit Sub
    If Not m_TEST_RegisterDemoConfigItemsVariantB(False) Then Exit Sub
    If Not private_TrySaveDemoConfigVariantToStoreForActiveSheet(g_DemoConfigVariant) Then Exit Sub
    Call ex_HelpersSheet.m_TryRerenderActivePage("configVariant:" & g_DemoConfigVariant)
End Sub


Public Sub m_TEST_ProfileDevTableListUI()
    Dim ws As Worksheet
    Dim tables As Collection
    Dim t0 As Double
    Dim t1 As Double
    Dim t2 As Double
    Dim t3 As Double

    Set ws = private_GetActiveWorksheet()
    If ws Is Nothing Then Exit Sub

    t0 = VBA.Timer
    Set tables = m_TEST_BuildDemoTableItems()
    t1 = VBA.Timer

    If tables Is Nothing Then Exit Sub

    If Not private_ResetItemsSources() Then Exit Sub
    If Not private_TrySetItemsSource("RuntimeItems.Test.Tables", tables, False) Then Exit Sub
    t2 = VBA.Timer

    private_RenderWorksheetPage ws, "ui\DevProfileTableUI.xml"
    t3 = VBA.Timer

    VBA.MsgBox "Profile (ms):" & VBA.vbCrLf & _
           "Build data: " & VBA.Format$((t1 - t0) * 1000#, "0") & VBA.vbCrLf & _
           "Register source: " & VBA.Format$((t2 - t1) * 1000#, "0") & VBA.vbCrLf & _
           "Render UI: " & VBA.Format$((t3 - t2) * 1000#, "0") & VBA.vbCrLf & _
           "Total: " & VBA.Format$((t3 - t0) * 1000#, "0"), VBA.vbInformation
End Sub


Public Sub m_TEST_FillNumbersRangeSimple()
    Dim ws As Worksheet
    Dim targetRange As Range
    Dim values() As Variant
    Dim r As Long
    Dim c As Long
    Dim n As Long

    Set ws = private_GetActiveWorksheet()
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

    Set ws = private_GetActiveWorksheet()
    If ws Is Nothing Then Exit Sub

    private_RenderWorksheetPage ws, "ui\DevSingleTableUI.xml"
End Sub


Public Function m_TEST_RegisterDemoListItems( _
    Optional ByVal notifyChange As Boolean = False, _
    Optional ByVal preferredPageBase As Variant _
) As Boolean
    Dim items As Collection

    Set items = m_TEST_BuildDemoListItems()
    If items Is Nothing Then Exit Function

    If Not private_ResetItemsSources(preferredPageBase) Then Exit Function
    If Not private_ResetObjectSources(preferredPageBase) Then Exit Function
    If Not private_TrySetItemsSource("RuntimeItems.Test.People", items, notifyChange, preferredPageBase) Then Exit Function

    m_TEST_RegisterDemoListItems = True
End Function


Public Function m_TEST_RegisterDemoTableItems( _
    Optional ByVal notifyChange As Boolean = False, _
    Optional ByVal preferredPageBase As Variant _
) As Boolean
    Dim tables As Collection

    Set tables = m_TEST_BuildDemoTableItems()
    If tables Is Nothing Then Exit Function

    If Not private_ResetItemsSources(preferredPageBase) Then Exit Function
    If Not private_ResetObjectSources(preferredPageBase) Then Exit Function
    If Not private_TrySetItemsSource("RuntimeItems.Test.Tables", tables, notifyChange, preferredPageBase) Then Exit Function

    m_TEST_RegisterDemoTableItems = True
End Function


Public Function m_TEST_RegisterDemoSingleTableItems( _
    Optional ByVal notifyChange As Boolean = False, _
    Optional ByVal preferredPageBase As Variant _
) As Boolean
    Dim tables As Collection

    Set tables = m_TEST_BuildDemoSingleTableItems()
    If tables Is Nothing Then Exit Function

    If Not private_ResetItemsSources(preferredPageBase) Then Exit Function
    If Not private_ResetObjectSources(preferredPageBase) Then Exit Function
    If Not private_TrySetItemsSource("RuntimeItems.Test.Tables", tables, notifyChange, preferredPageBase) Then Exit Function

    m_TEST_RegisterDemoSingleTableItems = True
End Function


Public Function m_TEST_RegisterDemoBannerItems( _
    Optional ByVal isVisible As Boolean = False, _
    Optional ByVal notifyChange As Boolean = False, _
    Optional ByVal preferredPageBase As Variant _
) As Boolean
    Dim bannerObj As obj_Banner
    Dim headerText As String
    Dim messageText As String

    If isVisible Then
        headerText = "Data Source [[accent]]Updated[[/accent]]"
        messageText = "Rows: [[ok]]20 tables[[/ok]]. State: [[warn]]runtime refresh[[/warn]]."
        Set bannerObj = private_CreateDemoBannerModel(headerText, messageText, isVisible)
        If bannerObj Is Nothing Then Exit Function
        If Not private_TrySetObjectSource("RuntimeObjects.Test.Banner", bannerObj, notifyChange, preferredPageBase) Then Exit Function
    Else
        If Not private_TryRemoveObjectSource("RuntimeObjects.Test.Banner", notifyChange, preferredPageBase) Then Exit Function
    End If

    m_TEST_RegisterDemoBannerItems = True
End Function


Public Function m_TEST_RegisterDemoConfigItemsVariantA( _
    Optional ByVal notifyChange As Boolean = False, _
    Optional ByVal preferredPageBase As Variant _
) As Boolean
    Dim items As Collection

    Set items = m_TEST_BuildDemoConfigItemsVariantA()
    If items Is Nothing Then Exit Function

    If Not private_TrySetItemsSource("RuntimeItems.Test.Config", items, notifyChange, preferredPageBase) Then Exit Function
    m_TEST_RegisterDemoConfigItemsVariantA = True
End Function


Public Function m_TEST_RegisterDemoConfigItemsVariantB( _
    Optional ByVal notifyChange As Boolean = False, _
    Optional ByVal preferredPageBase As Variant _
) As Boolean
    Dim items As Collection

    Set items = m_TEST_BuildDemoConfigItemsVariantB()
    If items Is Nothing Then Exit Function

    If Not private_TrySetItemsSource("RuntimeItems.Test.Config", items, notifyChange, preferredPageBase) Then Exit Function
    m_TEST_RegisterDemoConfigItemsVariantB = True
End Function


Public Function m_TEST_RegisterDemoConfigItemsByCurrentVariant( _
    Optional ByVal notifyChange As Boolean = False, _
    Optional ByVal preferredPageBase As Variant _
) As Boolean
    Select Case private_GetDemoConfigVariantKey()
        Case DEMO_CONFIG_VARIANT_B
            m_TEST_RegisterDemoConfigItemsByCurrentVariant = m_TEST_RegisterDemoConfigItemsVariantB(notifyChange, preferredPageBase)

        Case Else
            m_TEST_RegisterDemoConfigItemsByCurrentVariant = m_TEST_RegisterDemoConfigItemsVariantA(notifyChange, preferredPageBase)
    End Select
End Function


Public Function m_TEST_RegisterDemoConfigProfileItems( _
    Optional ByVal notifyChange As Boolean = False, _
    Optional ByVal preferredPageBase As Variant _
) As Boolean
    Dim options As Collection

    Set options = m_TEST_BuildDemoConfigProfileItems()
    If options Is Nothing Then Exit Function

    If Not private_TrySetItemsSource("RuntimeItems.Test.ConfigProfiles", options, notifyChange, preferredPageBase) Then Exit Function
    m_TEST_RegisterDemoConfigProfileItems = True
End Function


Public Function m_TEST_BuildDemoListItems() As Collection
    Dim result As Collection

    Set result = New Collection
    result.Add private_CreateDemoPerson("Ivan Petrov", "Team Lead")
    result.Add private_CreateDemoPerson("Anna Sidorova", "Analyst")
    result.Add private_CreateDemoPerson("Maksym Kovalenko", "Developer")

    Set m_TEST_BuildDemoListItems = result
End Function


Public Function m_TEST_BuildDemoConfigItemsVariantA() As Collection
    Dim result As Collection

    Set result = New Collection
    result.Add private_CreateConfigViewItem("#", "Profile.Name", "HospitalizationDate")
    result.Add private_CreateConfigViewItem("rx", "Source.Main.FilePattern", "{Main-{dd}.{mm}.{yyyy}}")
    result.Add private_CreateConfigViewItem(VBA.vbNullString, "Sheet.StateMain.Key.HospitalizationDate", "No; Unit; Rank; FIO; HospitalizationDate")
    result.Add private_CreateConfigViewItem(VBA.vbNullString, "Sheet.StateMain.Map.1", "No з/п")
    result.Add private_CreateConfigViewItem("#", "Sheet.StateMain.Map.2", "В/ч")
    result.Add private_CreateConfigViewItem("rx", "Sheet.StateMain.Map.3", "П.І.Б.")

    Set m_TEST_BuildDemoConfigItemsVariantA = result
End Function


Public Function m_TEST_BuildDemoConfigItemsVariantB() As Collection
    Dim result As Collection

    Set result = New Collection
    result.Add private_CreateConfigViewItem("#", "Profile.Name", "TransferSheet")
    result.Add private_CreateConfigViewItem("rx", "Source.Main.FileResolver", "ex_SourceResolvers.m_ResolveAllByPattern")
    result.Add private_CreateConfigViewItem(VBA.vbNullString, "Source.Main.SortOrder", "order=asc")
    result.Add private_CreateConfigViewItem(VBA.vbNullString, "Sheet.Aliases.StateMain", "StateMain")
    result.Add private_CreateConfigViewItem("rx", "Sheet.StateMain.Key.TransferDate", "{Main-{dd}.{mm}.{yyyy}}.DateTransfer")
    result.Add private_CreateConfigViewItem("#", "Sheet.StateMain.Key.DocName", "{Main-{dd}.{mm}.{yyyy}}.DocName")

    Set m_TEST_BuildDemoConfigItemsVariantB = result
End Function


Public Function m_TEST_BuildDemoConfigProfileItems() As Collection
    Dim result As Collection

    Set result = New Collection
    result.Add private_CreateSelectOption( _
        "HospitalizationDate", _
        DEMO_CONFIG_VARIANT_A, _
        "ex_Test.m_TEST_SetDemoConfigVariantA")
    result.Add private_CreateSelectOption( _
        "TransferSheet", _
        DEMO_CONFIG_VARIANT_B, _
        "ex_Test.m_TEST_SetDemoConfigVariantB")

    Set m_TEST_BuildDemoConfigProfileItems = result
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
        teamName = "Team " & VBA.Format$(tableIndex, "00")

        result.Add private_CreateDemoTable( _
            "People / " & teamName, _
            HEADER_TEXT, _
            private_CreateGeneratedRowsForTeam(tableIndex, rowsCount, teamName))
    Next tableIndex

    Set m_TEST_BuildDemoTableItems = result
End Function


Public Function m_TEST_BuildDemoTableViewItems( _
    Optional ByVal includeTableBanners As Boolean = False, _
    Optional ByVal includeRowBanners As Boolean = False _
) As Collection
    Dim sourceTables As Collection
    Dim result As Collection
    Dim tableObj As Variant
    Dim tableModel As obj_TableDynamic
    Dim tableView As obj_TableViewItem
    Dim rowObj As Variant
    Dim rowView As obj_RowViewItem
    Dim tableIndex As Long
    Dim rowIndex As Long
    Dim rowBannerTargetIndex As Long
    Dim rowBannerPositionName As String

    Set sourceTables = m_TEST_BuildDemoTableItems()
    If sourceTables Is Nothing Then Exit Function

    Set result = New Collection
    Randomize

    tableIndex = 0
    For Each tableObj In sourceTables
        tableIndex = tableIndex + 1

        If Not private_TryResolveDemoTableDynamic(tableObj, tableModel) Then Exit Function
        Set tableView = private_CreateTableViewItemFromTable(tableModel)
        If tableView Is Nothing Then Exit Function

        If includeTableBanners Then
            If (tableIndex Mod 5) = 0 Then
                Set tableView.Banner = private_CreateBannerViewItem( _
                    "Team note / " & VBA.Format$(tableIndex, "00"), _
                    "This banner is attached to table item " & VBA.CStr(tableIndex) & ".", _
                    True, _
                    2)
            End If
        End If

        rowBannerTargetIndex = 0
        rowBannerPositionName = VBA.vbNullString
        If includeRowBanners Then
            If (tableIndex Mod 4) = 0 Then
                rowBannerTargetIndex = private_GetRandomRowBannerTargetIndex(tableModel.RowCount, rowBannerPositionName)
            End If
        End If

        rowIndex = 0
        For Each rowObj In tableModel.Rows
            rowIndex = rowIndex + 1

            If VBA.TypeName(rowObj) <> "obj_Row" Then
                VBA.MsgBox "PrototypeNew: expected obj_Row in table rows for table view.", VBA.vbExclamation
                Exit Function
            End If

            Set rowView = private_CreateRowViewItemFromRow(rowObj)
            If rowView Is Nothing Then Exit Function

            If includeRowBanners Then
                If rowBannerTargetIndex > 0 And rowIndex = rowBannerTargetIndex Then
                    Set rowView.Banner = private_CreateBannerViewItem( _
                        "Row banner", _
                        "Attached to " & rowBannerPositionName & " row of Team " & VBA.Format$(tableIndex, "00") & ".", _
                        True, _
                        2)
                End If

                If rowIndex = tableModel.RowCount And (tableIndex Mod 3) = 0 Then
                    rowView.SpacerRowsAfter = 1
                End If
            End If

            tableView.RowItems.Add rowView
        Next rowObj

        result.Add tableView
    Next tableObj

    Set m_TEST_BuildDemoTableViewItems = result
End Function


Public Function m_TEST_BuildDemoSingleTableViewItems( _
    Optional ByVal includeTableBanners As Boolean = False, _
    Optional ByVal includeRowBanners As Boolean = False _
) As Collection
    Dim sourceTables As Collection
    Dim result As Collection
    Dim tableObj As Variant
    Dim tableModel As obj_TableDynamic
    Dim tableView As obj_TableViewItem
    Dim rowObj As Variant
    Dim rowView As obj_RowViewItem
    Dim rowIndex As Long

    Set sourceTables = m_TEST_BuildDemoSingleTableItems()
    If sourceTables Is Nothing Then Exit Function

    Set result = New Collection

    For Each tableObj In sourceTables
        If Not private_TryResolveDemoTableDynamic(tableObj, tableModel) Then Exit Function
        Set tableView = private_CreateTableViewItemFromTable(tableModel)
        If tableView Is Nothing Then Exit Function

        If includeTableBanners Then
            Set tableView.Banner = private_CreateBannerViewItem( _
                "Merged table note", _
                "This banner is attached to merged single table view.", _
                True, _
                2)
        End If

        rowIndex = 0
        For Each rowObj In tableModel.Rows
            rowIndex = rowIndex + 1

            If VBA.TypeName(rowObj) <> "obj_Row" Then
                VBA.MsgBox "PrototypeNew: expected obj_Row in single table rows for table view.", VBA.vbExclamation
                Exit Function
            End If

            Set rowView = private_CreateRowViewItemFromRow(rowObj)
            If rowView Is Nothing Then Exit Function

            If includeRowBanners Then
                If rowIndex = 1 Then
                    Set rowView.Banner = private_CreateBannerViewItem( _
                        "First row", _
                        "This row-level banner is attached before the first row.", _
                        True, _
                        2)
                End If
            End If

            tableView.RowItems.Add rowView
        Next rowObj

        result.Add tableView
    Next tableObj

    Set m_TEST_BuildDemoSingleTableViewItems = result
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
        If Not private_TryResolveDemoTableDynamic(tableObj, sourceTable) Then Exit Function

        If mergedTable.ColumnCount = 0 Then
            For Each sourceCol In sourceTable.Columns
                If Not mergedTable.AddColumn(sourceCol) Then Exit Function
            Next sourceCol
        End If

        For Each sourceRow In sourceTable.Rows
            If VBA.TypeName(sourceRow) <> "obj_Row" Then
                VBA.MsgBox "PrototypeNew: expected obj_Row in demo table rows.", VBA.vbExclamation
                Exit Function
            End If

            Set targetRow = New obj_Row
            For i = 1 To mergedTable.ColumnCount
                targetRow.AddCell sourceRow.GetCell(i)
            Next i

            If Not mergedTable.AddRow(targetRow) Then Exit Function
        Next sourceRow
    Next tableObj

    Set result = New Collection
    result.Add mergedTable

    Set m_TEST_BuildDemoSingleTableItems = result
End Function


Public Sub m_TEST_NoOp()
End Sub

' //
' // Internal
' //

Private Sub private_RenderWorksheetPage(ByVal ws As Worksheet, ByVal uiPath As String)
    Dim mainPage As obj_IPage
    Dim normalizedUiPath As String

    If ws Is Nothing Then Exit Sub
    normalizedUiPath = VBA.Trim$(uiPath)
    If VBA.Len(normalizedUiPath) = 0 Then Exit Sub

    If Not private_TryResolveMainPage(mainPage) Then Exit Sub
    If mainPage Is Nothing Then Exit Sub

    Call mainPage.UpdateUiPath(normalizedUiPath, "ex_Test:private_RenderWorksheetPage")
End Sub


Private Function private_CreateGeneratedRowsForTeam( _
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
        personName = "Person " & VBA.Format$(tableIndex, "00") & "-" & VBA.CStr(rowIndex)
        roleName = private_GetRoleByIndex(tableIndex + rowIndex)
        countryName = private_GetCountryByIndex(tableIndex + rowIndex)
        levelName = "L" & VBA.CStr(((tableIndex + rowIndex) Mod 4) + 1)

        If (tableIndex + rowIndex) Mod 5 = 0 Then
            statusName = "On Hold"
        Else
            statusName = "Active"
        End If

        sinceYear = VBA.CStr(2014 + ((tableIndex + rowIndex) Mod 11))

        result.Add private_CreateDemoRowModel( _
            personName, _
            roleName, _
            countryName, _
            teamName, _
            levelName, _
            statusName, _
            sinceYear)
    Next rowIndex

    Set private_CreateGeneratedRowsForTeam = result
End Function


Private Function private_GetRoleByIndex(ByVal idx As Long) As String
    Select Case ((idx - 1) Mod 7) + 1
        Case 1: private_GetRoleByIndex = "Team Lead"
        Case 2: private_GetRoleByIndex = "Analyst"
        Case 3: private_GetRoleByIndex = "Developer"
        Case 4: private_GetRoleByIndex = "QA"
        Case 5: private_GetRoleByIndex = "DevOps"
        Case 6: private_GetRoleByIndex = "Support"
        Case Else: private_GetRoleByIndex = "Manager"
    End Select
End Function


Private Function private_GetCountryByIndex(ByVal idx As Long) As String
    Select Case ((idx - 1) Mod 6) + 1
        Case 1: private_GetCountryByIndex = "Ukraine"
        Case 2: private_GetCountryByIndex = "Poland"
        Case 3: private_GetCountryByIndex = "Romania"
        Case 4: private_GetCountryByIndex = "Germany"
        Case 5: private_GetCountryByIndex = "Czechia"
        Case Else: private_GetCountryByIndex = "Slovakia"
    End Select
End Function


Private Function private_CreateDemoPerson(ByVal displayName As String, ByVal roleName As String) As Object
    Dim rowObj As Object

    Set rowObj = VBA.CreateObject("Scripting.Dictionary")
    rowObj.CompareMode = 1
    rowObj("Display") = VBA.CStr(displayName)
    rowObj("Role") = VBA.CStr(roleName)

    Set private_CreateDemoPerson = rowObj
End Function


Private Function private_CreateConfigViewItem( _
    ByVal attrText As String, _
    ByVal keyText As String, _
    ByVal valueText As String _
) As obj_ConfigViewItem
    Dim cfgModel As obj_Config
    Dim cfgView As obj_ConfigViewItem

    Set cfgModel = New obj_Config
    cfgModel.Attr = VBA.CStr(attrText)
    cfgModel.Key = VBA.CStr(keyText)
    cfgModel.Value = VBA.CStr(valueText)

    Set cfgView = New obj_ConfigViewItem
    Set cfgView.Model = cfgModel

    Set private_CreateConfigViewItem = cfgView
End Function


Private Function private_CreateSelectOption( _
    ByVal captionText As String, _
    ByVal idText As String, _
    ByVal onSelectMacro As String _
) As obj_SelectOption
    Dim selectOption As obj_SelectOption

    Set selectOption = New obj_SelectOption
    selectOption.Caption = VBA.CStr(captionText)
    selectOption.Id = VBA.CStr(idText)
    selectOption.OnSelect = VBA.CStr(onSelectMacro)

    Set private_CreateSelectOption = selectOption
End Function


Private Function private_CreateDemoBannerModel( _
    ByVal headerText As String, _
    ByVal messageText As String, _
    ByVal isVisible As Boolean _
) As obj_Banner
    Dim bannerObj As obj_Banner

    Set bannerObj = New obj_Banner
    bannerObj.Header = VBA.CStr(headerText)
    bannerObj.Message = VBA.CStr(messageText)
    bannerObj.Visible = VBA.CBool(isVisible)

    Set private_CreateDemoBannerModel = bannerObj
End Function


Private Function private_CreateBannerViewItem( _
    ByVal headerText As String, _
    ByVal messageText As String, _
    ByVal isVisible As Boolean, _
    Optional ByVal spanRows As Long = 2 _
) As obj_BannerViewItem
    Dim bannerView As obj_BannerViewItem

    Set bannerView = New obj_BannerViewItem
    bannerView.Model.Header = VBA.CStr(headerText)
    bannerView.Model.Message = VBA.CStr(messageText)
    bannerView.Model.Visible = VBA.CBool(isVisible)
    bannerView.Presentation.EffectiveVisible = VBA.CBool(isVisible)
    bannerView.Presentation.SpanRows = spanRows

    Set private_CreateBannerViewItem = bannerView
End Function


Private Function private_CreateTableViewItemFromTable(ByVal tableModel As obj_TableDynamic) As obj_TableViewItem
    Dim tableView As obj_TableViewItem

    If tableModel Is Nothing Then
        VBA.MsgBox "PrototypeNew: table model is not specified for table view.", VBA.vbExclamation
        Exit Function
    End If

    Set tableView = New obj_TableViewItem
    Set tableView.Model = tableModel
    tableView.ItemVisible = True

    Set private_CreateTableViewItemFromTable = tableView
End Function


Private Function private_CreateRowViewItemFromRow(ByVal rowModel As obj_Row) As obj_RowViewItem
    Dim rowView As obj_RowViewItem

    If rowModel Is Nothing Then
        VBA.MsgBox "PrototypeNew: row model is not specified for row view.", VBA.vbExclamation
        Exit Function
    End If

    Set rowView = New obj_RowViewItem
    Set rowView.Row = rowModel
    rowView.RowVisible = True

    Set private_CreateRowViewItemFromRow = rowView
End Function


Private Function private_GetRandomRowBannerTargetIndex( _
    ByVal rowCount As Long, _
    ByRef outPositionName As String _
) As Long
    Dim slotRoll As Long

    outPositionName = "first"
    If rowCount <= 0 Then Exit Function
    If rowCount = 1 Then
        private_GetRandomRowBannerTargetIndex = 1
        Exit Function
    End If

    slotRoll = VBA.Int(Rnd * 3) + 1

    Select Case slotRoll
        Case 1
            private_GetRandomRowBannerTargetIndex = 1
            outPositionName = "first"

        Case 2
            private_GetRandomRowBannerTargetIndex = ((rowCount - 1) \ 2) + 1
            outPositionName = "middle"

        Case Else
            private_GetRandomRowBannerTargetIndex = rowCount
            outPositionName = "last"
    End Select
End Function


Private Function private_TryResolveDemoTableDynamic(ByVal tableObj As Variant, ByRef outTable As obj_TableDynamic) As Boolean
    Dim fixedTable As obj_Table
    Dim sourceCol As obj_Column
    Dim sourceRow As obj_Row
    Dim dynamicTable As obj_TableDynamic
    Dim targetCol As obj_Column
    Dim targetRow As obj_Row
    Dim i As Long

    If Not VBA.IsObject(tableObj) Then
        VBA.MsgBox "PrototypeNew: demo table item is not object.", VBA.vbExclamation
        Exit Function
    End If

    Select Case VBA.LCase$(VBA.TypeName(tableObj))
        Case "obj_tabledynamic"
            Set outTable = tableObj
            private_TryResolveDemoTableDynamic = True

        Case "obj_table"
            Set fixedTable = tableObj
            Set dynamicTable = New obj_TableDynamic
            dynamicTable.SectionTitle = fixedTable.SectionTitle

            For Each sourceCol In fixedTable.Columns
                Set targetCol = New obj_Column
                targetCol.Name = sourceCol.Name
                targetCol.Position = sourceCol.Position
                If Not dynamicTable.AddColumn(targetCol) Then Exit Function
            Next sourceCol

            For Each sourceRow In fixedTable.Rows
                Set targetRow = New obj_Row
                For i = 1 To dynamicTable.ColumnCount
                    targetRow.AddCell sourceRow.GetCell(i)
                Next i
                If Not dynamicTable.AddRow(targetRow) Then Exit Function
            Next sourceRow

            Set outTable = dynamicTable
            private_TryResolveDemoTableDynamic = True

        Case Else
            VBA.MsgBox "PrototypeNew: unsupported demo table type '" & VBA.TypeName(tableObj) & "'.", VBA.vbExclamation
    End Select
End Function


Private Function private_CreateDemoTable( _
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
    tableObj.SectionTitle = VBA.CStr(sectionTitle)

    headerTokens = VBA.Split(VBA.CStr(headerText), "|")
    For colIndex = LBound(headerTokens) To UBound(headerTokens)
        Set colObj = New obj_Column
        colObj.Position = colIndex + 1
        colObj.Name = VBA.Trim$(VBA.CStr(headerTokens(colIndex)))
        If VBA.Len(colObj.Name) = 0 Then colObj.Name = "Col" & VBA.CStr(colObj.Position)
        If Not tableObj.AddColumn(colObj) Then Exit Function
    Next colIndex

    If rows Is Nothing Then
        Set private_CreateDemoTable = tableObj
        Exit Function
    End If

    For Each rowObj In rows
        If rowObj Is Nothing Then
            VBA.MsgBox "PrototypeNew: table row is not specified.", VBA.vbExclamation
            Exit Function
        End If
        If rowObj.CellCount < tableObj.ColumnCount Then
            VBA.MsgBox "PrototypeNew: table row has fewer cells than table columns.", VBA.vbExclamation
            Exit Function
        End If

        If Not tableObj.AddRow(rowObj) Then Exit Function
    Next rowObj

    Set private_CreateDemoTable = tableObj
End Function


Private Function private_CreateDemoTableRows(ParamArray values() As Variant) As Collection
    Dim result As Collection
    Dim i As Long

    Set result = New Collection
    If (UBound(values) - LBound(values) + 1) Mod 7 <> 0 Then
        VBA.MsgBox "PrototypeNew: private_CreateDemoTableRows expects values in septets (c1..c7).", VBA.vbExclamation
        Set private_CreateDemoTableRows = result
        Exit Function
    End If

    For i = LBound(values) To UBound(values) Step 7
        result.Add private_CreateDemoRowModel( _
            VBA.CStr(values(i)), _
            VBA.CStr(values(i + 1)), _
            VBA.CStr(values(i + 2)), _
            VBA.CStr(values(i + 3)), _
            VBA.CStr(values(i + 4)), _
            VBA.CStr(values(i + 5)), _
            VBA.CStr(values(i + 6)))
    Next i

    Set private_CreateDemoTableRows = result
End Function


Private Function private_CreateDemoRowModel( _
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
    rowObj.AddCell c1
    rowObj.AddCell c2
    rowObj.AddCell c3
    rowObj.AddCell c4
    rowObj.AddCell c5
    rowObj.AddCell c6
    rowObj.AddCell c7

    Set private_CreateDemoRowModel = rowObj
End Function


Private Function private_TryResolveMainPage(ByRef outPage As obj_IPage) As Boolean
    Dim mainWs As Worksheet
    Dim pagesByType As Collection
    Dim pageCandidate As Variant

    Set outPage = Nothing

    Set mainWs = ex_HelpersSheet.m_GetRuntimeWorksheetByName("Main")
    If Not mainWs Is Nothing Then
        If rt_PageManager.m_TryGetPageByWorksheet(mainWs, outPage) Then
            private_TryResolveMainPage = True
            Exit Function
        End If
    End If

    If rt_PageManager.m_TryGetPagesByType(PageTypeMain, pagesByType) Then
        If Not pagesByType Is Nothing Then
            For Each pageCandidate In pagesByType
                If VBA.IsObject(pageCandidate) Then
                    Set outPage = pageCandidate
                    If Not outPage Is Nothing Then
                        private_TryResolveMainPage = True
                        Exit Function
                    End If
                End If
            Next pageCandidate
        End If
    End If

    VBA.MsgBox "PrototypeNew: main page is not resolved for UI switch.", VBA.vbExclamation
End Function


Private Function private_GetActiveWorksheet() As Worksheet
    Dim wb As Workbook
    Dim activeSheetObj As Object

    Set wb = ThisWorkbook
    If wb Is Nothing Then
        VBA.MsgBox "PrototypeNew: workbook is not specified.", VBA.vbExclamation
        Exit Function
    End If

    Set activeSheetObj = wb.ActiveSheet
    If activeSheetObj Is Nothing Then
        VBA.MsgBox "PrototypeNew: active sheet is not specified.", VBA.vbExclamation
        Exit Function
    End If

    If Not TypeOf activeSheetObj Is Worksheet Then
        VBA.MsgBox "PrototypeNew: active sheet is not a worksheet.", VBA.vbExclamation
        Exit Function
    End If

    Set private_GetActiveWorksheet = activeSheetObj
End Function


Private Function private_GetDemoConfigVariantKey() As String
    g_DemoConfigVariant = VBA.LCase$(VBA.Trim$(g_DemoConfigVariant))
    If VBA.Len(g_DemoConfigVariant) = 0 Then g_DemoConfigVariant = DEMO_CONFIG_VARIANT_A

    Select Case g_DemoConfigVariant
        Case DEMO_CONFIG_VARIANT_A, DEMO_CONFIG_VARIANT_B
            private_GetDemoConfigVariantKey = g_DemoConfigVariant

        Case Else
            g_DemoConfigVariant = DEMO_CONFIG_VARIANT_A
            private_GetDemoConfigVariantKey = g_DemoConfigVariant
    End Select
End Function


Private Function private_TryResolvePageBase( _
    ByRef outPageBase As obj_PageBase, _
    Optional ByVal preferredPageBase As Variant _
) As Boolean
    Set outPageBase = Nothing
    If Not IsMissing(preferredPageBase) Then
        If VBA.IsObject(preferredPageBase) Then
            If Not preferredPageBase Is Nothing Then
                If TypeOf preferredPageBase Is obj_PageBase Then
                    Set outPageBase = preferredPageBase
                    private_TryResolvePageBase = True
                    Exit Function
                End If

                If ex_HelpersSheet.m_TryCastPageBase(preferredPageBase, outPageBase) Then
                    private_TryResolvePageBase = True
                    Exit Function
                End If

                VBA.MsgBox "PrototypeNew: preferred page runtime context has unsupported type '" & VBA.TypeName(preferredPageBase) & "'.", VBA.vbExclamation
                Exit Function
            End If
        End If
    End If

    If Not ex_HelpersSheet.m_TryGetActivePageBase(outPageBase) Then
        VBA.MsgBox "PrototypeNew: page runtime context is not resolved for active worksheet.", VBA.vbExclamation
        Exit Function
    End If
    If outPageBase Is Nothing Then Exit Function

    private_TryResolvePageBase = True
End Function


Private Function private_ResetItemsSources(Optional ByVal preferredPageBase As Variant) As Boolean
    Dim pageBase As obj_PageBase

    If Not private_TryResolvePageBase(pageBase, preferredPageBase) Then Exit Function
    pageBase.RuntimeSources.ResetItemsSources
    private_ResetItemsSources = True
End Function


Private Function private_ResetObjectSources(Optional ByVal preferredPageBase As Variant) As Boolean
    Dim pageBase As obj_PageBase

    If Not private_TryResolvePageBase(pageBase, preferredPageBase) Then Exit Function
    pageBase.RuntimeSources.ResetObjectSources
    private_ResetObjectSources = True
End Function


Private Function private_TrySetItemsSource( _
    ByVal sourceKey As String, _
    ByVal items As Collection, _
    ByVal notifyChange As Boolean, _
    Optional ByVal preferredPageBase As Variant _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim normalizedKey As String

    If Not private_TryResolvePageBase(pageBase, preferredPageBase) Then Exit Function
    normalizedKey = VBA.LCase$(VBA.Trim$(sourceKey))

    If Not pageBase.RuntimeSources.SetItemsSource(normalizedKey, items) Then Exit Function
    If notifyChange Then
        If Not private_TryRerenderPage(pageBase, "itemsSource:" & normalizedKey) Then Exit Function
    End If

    private_TrySetItemsSource = True
End Function


Private Function private_TrySetObjectSource( _
    ByVal sourceKey As String, _
    ByVal sourceObject As Object, _
    ByVal notifyChange As Boolean, _
    Optional ByVal preferredPageBase As Variant _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim normalizedKey As String

    If Not private_TryResolvePageBase(pageBase, preferredPageBase) Then Exit Function
    normalizedKey = VBA.LCase$(VBA.Trim$(sourceKey))

    If Not pageBase.RuntimeSources.SetObjectSource(normalizedKey, sourceObject) Then Exit Function
    If notifyChange Then
        If Not private_TryRerenderPage(pageBase, "objectSource:" & normalizedKey) Then Exit Function
    End If

    private_TrySetObjectSource = True
End Function


Private Function private_TryRemoveObjectSource( _
    ByVal sourceKey As String, _
    ByVal notifyChange As Boolean, _
    Optional ByVal preferredPageBase As Variant _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim normalizedKey As String

    If Not private_TryResolvePageBase(pageBase, preferredPageBase) Then Exit Function
    normalizedKey = VBA.LCase$(VBA.Trim$(sourceKey))

    If Not pageBase.RuntimeSources.RemoveObjectSource(normalizedKey) Then Exit Function
    If notifyChange Then
        If Not private_TryRerenderPage(pageBase, "objectSource:" & normalizedKey) Then Exit Function
    End If

    private_TryRemoveObjectSource = True
End Function


Private Function private_TryRerenderPage(ByVal pageBase As obj_PageBase, ByVal reason As String) As Boolean
    Dim pageRef As obj_IPage
    Dim ws As Worksheet

    If pageBase Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Function

    If Not rt_PageManager.m_TryGetPageByWorksheet(ws, pageRef) Then Exit Function
    private_TryRerenderPage = rt_PageManager.m_RenderPage(pageRef, reason)
End Function


Private Function private_TryLoadDemoConfigVariantFromStore(ByVal ws As Worksheet) As Boolean
    Dim selectStateKey As String
    Dim storedSelectedId As String
    Dim selectStatic As obj_SelectControlVMStatic

    If ws Is Nothing Then
        VBA.MsgBox "PrototypeNew: worksheet is not specified for config profile state restore.", VBA.vbExclamation
        Exit Function
    End If

    selectStateKey = VBA.LCase$(ws.Name & "|ConfigProfilePicker")
    Set selectStatic = New obj_SelectControlVMStatic
    If Not selectStatic.TryGetSelectedId(selectStateKey, storedSelectedId) Then Exit Function

    storedSelectedId = VBA.LCase$(VBA.Trim$(storedSelectedId))
    Select Case storedSelectedId
        Case DEMO_CONFIG_VARIANT_A, DEMO_CONFIG_VARIANT_B
            g_DemoConfigVariant = storedSelectedId

        Case Else
            If VBA.Len(VBA.Trim$(g_DemoConfigVariant)) = 0 Then g_DemoConfigVariant = DEMO_CONFIG_VARIANT_A
    End Select

    private_TryLoadDemoConfigVariantFromStore = True
End Function


Private Function private_TrySaveDemoConfigVariantToStoreForActiveSheet(ByVal configVariant As String) As Boolean
    Dim ws As Worksheet
    Dim selectStateKey As String
    Dim selectStatic As obj_SelectControlVMStatic

    Set ws = private_GetActiveWorksheet()
    If ws Is Nothing Then Exit Function

    selectStateKey = VBA.LCase$(ws.Name & "|ConfigProfilePicker")
    Set selectStatic = New obj_SelectControlVMStatic
    If Not selectStatic.SetSelectedId(selectStateKey, VBA.LCase$(VBA.Trim$(configVariant))) Then Exit Function

    private_TrySaveDemoConfigVariantToStoreForActiveSheet = True
End Function
