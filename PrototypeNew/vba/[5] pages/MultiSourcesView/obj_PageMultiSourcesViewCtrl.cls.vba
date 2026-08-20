VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PageMultiSourcesViewCtrl"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True

Private Const CONTROLLER_RUNTIME_OBJECT_KEY As String = "RuntimeObjects.MultiSourcesView.Controller"
Private Const TABLES_RUNTIME_KEY As String = "RuntimeItems.MultiSourcesView.Tables"
Private Const FILTERS_RUNTIME_KEY As String = "RuntimeItems.MultiSourcesView.Filters"
' Координаты строки фильтров задаются layout-контрактом MultiSourcesViewUI:
' вертикальный stack начинается с B2, после двух строк панели команд и
' одной строки постоянных заголовков фильтров.
' Это позволяет считывать значения пакетно по Search и не назначать
' обработчик изменения каждой ячейке, который вызывал бы дорогой рендер
' после каждого введенного символа.
Private Const FILTER_ROW_INDEX As Long = 5
Private Const FILTER_FIRST_COLUMN_INDEX As Long = 2
Private Const FILTER_EXPRESSION_EMPTY As String = "empty()"
Private Const FILTER_EXPRESSION_NOT_EMPTY As String = "notempty()"
Private Const ADO_LONG_VALUE_CANDIDATE_TAG As String = "ado-long-value-candidate"
Private Const ADO_TEXT_LIMIT As Long = 255

Private m_Page As obj_IPage
Private m_CfgParser As obj_MultiSourcesViewCfgParser
Private m_ConfigTable As obj_ConfigTable
Private m_TableRefs As Collection
Private m_Columns As Collection
Private m_SourceData As Collection
Private m_SourceTitles As Collection
Private m_SourceAliases As Collection
Private m_SourceAliasTemplates As Collection
Private m_FilterValues As Object
Private m_MaxRows As Long
Private m_Scenario As obj_IMultiSourcesScenario
Private m_IsConfigReady As Boolean
Private m_IsDisposed As Boolean

Public Property Get RuntimeObjectSourceKey() As String
    RuntimeObjectSourceKey = CONTROLLER_RUNTIME_OBJECT_KEY
End Property

Public Property Get OrderNoText() As String
    If Not m_Scenario Is Nothing Then OrderNoText = m_Scenario.OrderNoText
End Property

Public Function Initialize(ByVal page As obj_IPage) As Boolean
    Dim pageBase As obj_PageBase
    Dim emptyItems As Collection

    If page Is Nothing Then Exit Function
    m_IsDisposed = False
    Set m_Page = page
    Set m_FilterValues = ex_Helpers.fn_CreateDictionaryTextCompare()
    Set pageBase = page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Not pageBase.RuntimeSources.SetObjectSource(CONTROLLER_RUNTIME_OBJECT_KEY, Me) Then Exit Function
    Set emptyItems = New Collection
    If Not pageBase.RuntimeSources.SetItemsSource(TABLES_RUNTIME_KEY, emptyItems, False) Then Exit Function
    Set emptyItems = New Collection
    If Not pageBase.RuntimeSources.SetItemsSource(FILTERS_RUNTIME_KEY, emptyItems, False) Then Exit Function
    Initialize = True
End Function

Public Sub Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    If Not m_CfgParser Is Nothing Then m_CfgParser.Dispose
    Set m_CfgParser = Nothing
    Set m_ConfigTable = Nothing
    Set m_TableRefs = Nothing
    Set m_Columns = Nothing
    Set m_SourceData = Nothing
    Set m_SourceTitles = Nothing
    Set m_SourceAliases = Nothing
    Set m_SourceAliasTemplates = Nothing
    Set m_FilterValues = Nothing
    If Not m_Scenario Is Nothing Then m_Scenario.Dispose
    Set m_Scenario = Nothing
    Set m_Page = Nothing
    On Error GoTo 0
End Sub

Public Function UpdateData(ByVal configControl As obj_ConfigControlVM) As Boolean
    Dim configTable As obj_ConfigTable
    Dim cfgParser As obj_MultiSourcesViewCfgParser
    Dim scenarioClassName As String
    Dim hasScenario As Boolean

    m_IsConfigReady = False
    If configControl Is Nothing Then Exit Function
    If Not configControl.TryBuildConfigTableFromRendered(configTable) Then Exit Function
    Set cfgParser = New obj_MultiSourcesViewCfgParser
    If Not cfgParser.Initialize(configTable) Then Exit Function
    If Not cfgParser.TryGetScenarioClassName( _
        scenarioClassName, hasScenario) Then Exit Function
    If hasScenario Then
        If Not m_Scenario Is Nothing Then m_Scenario.Dispose
        Set m_Scenario = Nothing
        If Not private_TryCreateScenario( _
            scenarioClassName, m_Scenario) Then Exit Function
        If Not m_Scenario.Initialize(m_Page, configTable) Then Exit Function
    Else
        If Not m_Scenario Is Nothing Then m_Scenario.Dispose
        Set m_Scenario = Nothing
        If Not cfgParser.TryGetViewSettings(m_TableRefs, m_Columns, m_MaxRows) Then Exit Function
    End If

    If Not m_CfgParser Is Nothing Then m_CfgParser.Dispose
    Set m_CfgParser = cfgParser
    Set m_ConfigTable = configTable
    If m_Scenario Is Nothing Then
        If Not private_PublishFilterItems() Then Exit Function
    End If
    m_IsConfigReady = True
    UpdateData = True
End Function

Public Function RunPipeline(Optional ByVal notifyChange As Boolean = True) As Boolean
    Dim tableRef As Variant
    Dim sqlParams As obj_SqlParams
    Dim sqlParamsList As Collection
    Dim sqlParamsItem As Variant
    Dim tableData As obj_TableData
    Dim sourceIndex As Long
    Dim pipelineStage As String

    On Error GoTo EH

    If Not m_Scenario Is Nothing Then
        RunPipeline = m_Scenario.RunPipeline(notifyChange)
        Exit Function
    End If

    If Not m_IsConfigReady Or m_CfgParser Is Nothing Then
        VBA.MsgBox "MultiSourcesView config is not ready.", VBA.vbExclamation, "PrototypeNew / MultiSourcesView"
        Exit Function
    End If

    ' Строка фильтров доступна уже на первом рендере. Перед чтением
    ' источников забираем введенные значения прямо с листа, чтобы сценарий
    ' "настроить фильтры -> Load sources" не требовал отдельного Apply.
    pipelineStage = "read-filters"
    If Not private_TryReadFilterValuesFromSheet() Then Exit Function

    Set m_SourceData = New Collection
    Set m_SourceTitles = New Collection
    Set m_SourceAliases = New Collection
    Set m_SourceAliasTemplates = New Collection
    ex_Core.fn_Diagnostic_LogInfo "MultiSourcesView:pipeline-start tableRefs=" & _
        VBA.CStr(m_TableRefs.Count) & " columns=" & VBA.CStr(m_Columns.Count) & _
        " maxRows=" & VBA.CStr(m_MaxRows)
    For Each tableRef In m_TableRefs
        pipelineStage = "resolve:" & VBA.CStr(tableRef)
        If Not m_CfgParser.TryBuildTableSqlParamsList( _
            VBA.CStr(tableRef), m_Columns, sqlParamsList) Then Exit Function
        ex_Core.fn_Diagnostic_LogInfo "MultiSourcesView:resolved tableRef='" & _
            VBA.CStr(tableRef) & "' sources=" & VBA.CStr(sqlParamsList.Count)
        For Each sqlParamsItem In sqlParamsList
            Set sqlParams = sqlParamsItem
            sqlParams.MaxRows = m_MaxRows
            sourceIndex = sourceIndex + 1
            pipelineStage = "query:" & VBA.CStr(sourceIndex)
            ex_Core.fn_Diagnostic_LogInfo "MultiSourcesView:query-start index=" & _
                VBA.CStr(sourceIndex) & " path='" & sqlParams.SourcePath & "'"
            If Not ex_ExternalExcelSqlEngine.fn_TrySqlRequestData(sqlParams, tableData) Then Exit Function
            ex_Core.fn_Diagnostic_LogInfo "MultiSourcesView:query-done index=" & _
                VBA.CStr(sourceIndex) & " rows=" & VBA.CStr(tableData.RowCount) & _
                " columns=" & VBA.CStr(tableData.ColumnCount)
            pipelineStage = "collect:" & VBA.CStr(sourceIndex)
            m_SourceData.Add tableData
            m_SourceTitles.Add private_GetFileName(sqlParams.SourcePath)
            m_SourceAliases.Add sqlParams.SourceAlias
            m_SourceAliasTemplates.Add sqlParams.SourceAliasTemplate
        Next sqlParamsItem
    Next tableRef

    pipelineStage = "publish"
    ex_Core.fn_Diagnostic_LogInfo "MultiSourcesView:publish-request sources=" & _
        VBA.CStr(m_SourceData.Count) & " titles=" & VBA.CStr(m_SourceTitles.Count)
    If Not private_PublishVisibleTables() Then Exit Function
    If Not private_PublishFilterItems() Then Exit Function
    If notifyChange Then
        pipelineStage = "render"
        ex_Core.fn_Diagnostic_LogInfo "MultiSourcesView:render-request"
        If Not rt_PageManager.fn_RenderPage(m_Page, "multisourcesview:load") Then Exit Function
        ex_Core.fn_Diagnostic_LogInfo "MultiSourcesView:render-done"
    End If
    rt_Messaging.fn_ShowStatusBarSuccess "MultiSourcesView: loaded " & VBA.CStr(m_SourceData.Count) & " source(s).", 4
    RunPipeline = True
    Exit Function

EH:
    ex_Core.fn_Diagnostic_LogError "MultiSourcesView:pipeline-error stage='" & _
        pipelineStage & "' number=" & VBA.CStr(Err.Number) & _
        " source='" & Err.Source & "' description='" & Err.Description & "'"
    VBA.MsgBox "MultiSourcesView failed at stage '" & pipelineStage & "': " & _
        Err.Description & " (" & VBA.CStr(Err.Number) & ")", _
        VBA.vbExclamation, "PrototypeNew / MultiSourcesView"
    Err.Clear
End Function

Public Function RuntimeApplyFilters(Optional ByVal arg As Variant) As Boolean
    If m_Page Is Nothing Or m_Columns Is Nothing Then Exit Function
    If Not private_TryReadFilterValuesFromSheet() Then Exit Function

    ' SQL повторно не выполняется: фильтры применяются к сохраненным
    ' obj_TableData. Благодаря этому изменение нескольких условий требует
    ' одного прохода по памяти вместо повторного открытия всех Excel-файлов.
    If Not m_SourceData Is Nothing Then
        If Not private_PublishVisibleTables() Then Exit Function
    End If
    If Not private_PublishFilterItems() Then Exit Function
    RuntimeApplyFilters = rt_PageManager.fn_RenderPage(m_Page, "multisourcesview:apply-filters")
End Function

Private Function private_TryReadFilterValuesFromSheet() As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim columnAlias As String
    Dim filterValue As String
    Dim columnIndex As Long
    Dim cellValue As Variant

    If m_Page Is Nothing Or m_Columns Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Or pageBase.Worksheet Is Nothing Then Exit Function
    Set ws = pageBase.Worksheet
    If m_FilterValues Is Nothing Then Set m_FilterValues = ex_Helpers.fn_CreateDictionaryTextCompare()

    ' Физическая позиция каждой ячейки соответствует позиции алиаса в
    ' MultiSourcesView.Columns. Отображаемые заголовки не используются:
    ' ключом состояния всегда остается конфигурационный алиас.
    For columnIndex = 1 To m_Columns.Count
        columnAlias = VBA.Trim$(VBA.CStr(m_Columns.Item(columnIndex)))
        cellValue = ws.Cells(FILTER_ROW_INDEX, FILTER_FIRST_COLUMN_INDEX + columnIndex - 1).Value2
        filterValue = VBA.vbNullString
        If Not VBA.IsError(cellValue) And Not VBA.IsNull(cellValue) And Not VBA.IsEmpty(cellValue) Then
            filterValue = VBA.Trim$(VBA.CStr(cellValue))
        End If
        m_FilterValues(columnAlias) = filterValue
    Next columnIndex

    private_TryReadFilterValuesFromSheet = True
End Function

Public Function RerenderPage(Optional ByVal notifyChange As Boolean = True) As Boolean
    If m_Page Is Nothing Then Exit Function
    If Not private_PublishFilterItems() Then Exit Function
    RerenderPage = rt_PageManager.fn_RenderPage(m_Page, "multisourcesview:rerender")
End Function

Private Function private_TryCreateScenario( _
    ByVal className As String, _
    ByRef outScenario As obj_IMultiSourcesScenario _
) As Boolean
    Set outScenario = Nothing
    className = VBA.Trim$(className)
    If VBA.Len(className) = 0 Then
        VBA.MsgBox "MultiSourcesView.ScenarioClass is empty.", _
            VBA.vbExclamation, "PrototypeNew / MultiSourcesView"
        Exit Function
    End If

    ' VBA не умеет создавать классы текущего проекта через CreateObject.
    ' Эта фабрика является единственной точкой регистрации реализаций;
    ' конкретный класс выбирается только по имени из XML-профиля.
    Select Case VBA.LCase$(className)
        Case VBA.LCase$("obj_PEB_MovementVldtnScen")
            Set outScenario = New obj_PEB_MovementVldtnScen
        Case Else
            VBA.MsgBox "Unsupported MultiSourcesView scenario class: " & _
                className, VBA.vbExclamation, _
                "PrototypeNew / MultiSourcesView"
            Exit Function
    End Select
    private_TryCreateScenario = Not outScenario Is Nothing
End Function

Private Function private_PublishFilterItems() As Boolean
    Dim items As Collection
    Dim itemValue As Variant
    Dim entry As obj_ConfigEntry
    Dim pageBase As obj_PageBase

    Set items = New Collection
    If m_FilterValues Is Nothing Then Set m_FilterValues = ex_Helpers.fn_CreateDictionaryTextCompare()
    For Each itemValue In m_Columns
        Set entry = New obj_ConfigEntry
        entry.Key = VBA.Trim$(VBA.CStr(itemValue))
        ' На первом рендере Value остается пустым. После Apply/Load значение
        ' берется из m_FilterValues, поэтому повторный рендер страницы может
        ' очистить физические ячейки, не теряя введенные пользователем данные.
        entry.Value = VBA.vbNullString
        If m_FilterValues.Exists(entry.Key) Then
            entry.Value = VBA.CStr(m_FilterValues(entry.Key))
        End If
        items.Add entry
    Next itemValue
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    private_PublishFilterItems = pageBase.RuntimeSources.SetItemsSource(FILTERS_RUNTIME_KEY, items, False)
End Function

Private Function private_PublishVisibleTables() As Boolean
    Dim tables As Collection
    Dim tableObj As obj_TableDynamic
    Dim tableData As obj_TableData
    Dim pageBase As obj_PageBase
    Dim i As Long

    On Error GoTo EH

    Set tables = New Collection
    ex_Core.fn_Diagnostic_LogInfo "MultiSourcesView:publish-start sources=" & _
        VBA.CStr(m_SourceData.Count)
    For i = 1 To m_SourceData.Count
        Set tableData = m_SourceData.Item(i)
        ex_Core.fn_Diagnostic_LogInfo "MultiSourcesView:table-build-start index=" & _
            VBA.CStr(i) & " title='" & VBA.CStr(m_SourceTitles.Item(i)) & _
            "' sourceAlias='" & VBA.CStr(m_SourceAliases.Item(i)) & _
            "' sourceAliasTemplate='" & VBA.CStr(m_SourceAliasTemplates.Item(i)) & _
            "' rows=" & VBA.CStr(tableData.RowCount) & _
            " columns=" & VBA.CStr(tableData.ColumnCount)
        If Not private_TryBuildVisibleTable( _
            VBA.CStr(m_SourceTitles.Item(i)), _
            VBA.CStr(m_SourceAliases.Item(i)), _
            VBA.CStr(m_SourceAliasTemplates.Item(i)), _
            tableData, tableObj) Then Exit Function
        tables.Add tableObj
        ex_Core.fn_Diagnostic_LogInfo "MultiSourcesView:table-build-done index=" & _
            VBA.CStr(i) & " visibleRows=" & VBA.CStr(tableObj.RowCount) & _
            " columns=" & VBA.CStr(tableObj.ColumnCount)
    Next i
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    ex_Core.fn_Diagnostic_LogInfo "MultiSourcesView:set-items-source-start tables=" & _
        VBA.CStr(tables.Count)
    private_PublishVisibleTables = pageBase.RuntimeSources.SetItemsSource(TABLES_RUNTIME_KEY, tables, False)
    ex_Core.fn_Diagnostic_LogInfo "MultiSourcesView:set-items-source-done ok=" & _
        VBA.CStr(private_PublishVisibleTables)
    Exit Function

EH:
    ex_Core.fn_Diagnostic_LogError "MultiSourcesView:publish-error index=" & _
        VBA.CStr(i) & " number=" & VBA.CStr(Err.Number) & _
        " source='" & Err.Source & "' description='" & Err.Description & "'"
    Err.Clear
End Function

Private Function private_GetFileName(ByVal filePath As String) As String
    Dim slashPos As Long

    filePath = VBA.Replace(VBA.Trim$(filePath), "/", "\")
    slashPos = VBA.InStrRev(filePath, "\")
    If slashPos > 0 Then
        private_GetFileName = VBA.Mid$(filePath, slashPos + 1)
    Else
        private_GetFileName = filePath
    End If
End Function

Private Function private_TryBuildVisibleTable( _
    ByVal tableRef As String, _
    ByVal sourceAlias As String, _
    ByVal sourceAliasTemplate As String, _
    ByVal tableData As obj_TableData, _
    ByRef outTable As obj_TableDynamic _
) As Boolean
    Dim tableObj As obj_TableDynamic
    Dim colObj As obj_Column
    Dim rowObj As obj_Row
    Dim cellObj As obj_Cell
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim aliasText As String
    Dim buildStage As String

    On Error GoTo EH

    Set outTable = Nothing
    buildStage = "create-table"
    Set tableObj = New obj_TableDynamic
    tableObj.SectionTitle = tableRef
    tableObj.SourceAlias = sourceAlias
    tableObj.SourceAliasTemplate = sourceAliasTemplate
    buildStage = "columns"
    For colIndex = 1 To m_Columns.Count
        aliasText = VBA.Trim$(VBA.CStr(m_Columns.Item(colIndex)))
        Set colObj = New obj_Column
        colObj.Name = aliasText
        colObj.Position = colIndex
        Call colObj.AddAlias(aliasText)
        If Not tableObj.PushColumn(colObj) Then Exit Function
    Next colIndex

    For rowIndex = 1 To tableData.RowCount
        buildStage = "row=" & VBA.CStr(rowIndex)
        If private_RowMatches(tableData, rowIndex) Then
            Set rowObj = New obj_Row
            For colIndex = 1 To m_Columns.Count
                buildStage = "row=" & VBA.CStr(rowIndex) & ";col=" & VBA.CStr(colIndex)
                Set cellObj = New obj_Cell
                cellObj.Value = tableData.ValueAt(rowIndex, colIndex)
                If VBA.Len(VBA.CStr(cellObj.Value)) = ADO_TEXT_LIMIT Then
                    If Not cellObj.AddTag(ADO_LONG_VALUE_CANDIDATE_TAG) Then Exit Function
                End If
                If Not rowObj.PushCell(cellObj) Then Exit Function
            Next colIndex
            If Not tableObj.PushRow(rowObj) Then Exit Function
        End If
    Next rowIndex
    Set outTable = tableObj
    ex_Core.fn_Diagnostic_LogInfo "MultiSourcesView:table-materialized title='" & _
        tableRef & "' sourceRows=" & VBA.CStr(tableData.RowCount) & _
        " visibleRows=" & VBA.CStr(tableObj.RowCount)
    private_TryBuildVisibleTable = True
    Exit Function

EH:
    ex_Core.fn_Diagnostic_LogError "MultiSourcesView: failed to build table '" & _
        tableRef & "' at " & buildStage & " number=" & VBA.CStr(Err.Number) & _
        " source='" & Err.Source & "' description='" & Err.Description & "'"
    Err.Clear
    On Error GoTo 0
End Function

Private Function private_RowMatches(ByVal tableData As obj_TableData, ByVal rowIndex As Long) As Boolean
    Dim colIndex As Long
    Dim aliasText As String
    Dim queryText As String
    Dim cellText As String

    For colIndex = 1 To m_Columns.Count
        aliasText = VBA.Trim$(VBA.CStr(m_Columns.Item(colIndex)))
        queryText = VBA.vbNullString
        If m_FilterValues.Exists(aliasText) Then queryText = VBA.Trim$(VBA.CStr(m_FilterValues(aliasText)))
        If VBA.Len(queryText) > 0 Then
            cellText = tableData.ValueAt(rowIndex, colIndex)
            If Not private_FilterMatches(cellText, queryText) Then Exit Function
        End If
    Next colIndex
    private_RowMatches = True
End Function

Private Function private_FilterMatches( _
    ByVal cellText As String, _
    ByVal filterText As String _
) As Boolean
    Dim likeExpression As String

    filterText = VBA.Trim$(filterText)
    ' Проверки пустоты являются самостоятельными выражениями всей ячейки.
    ' Их намеренно не вкладываем в rx(): rx() описывает только LIKE-шаблон,
    ' тогда как empty()/notempty() проверяют состояние значения после Trim.
    If VBA.StrComp(filterText, FILTER_EXPRESSION_EMPTY, VBA.vbTextCompare) = 0 Then
        private_FilterMatches = (VBA.Len(VBA.Trim$(cellText)) = 0)
        Exit Function
    End If
    If VBA.StrComp(filterText, FILTER_EXPRESSION_NOT_EMPTY, VBA.vbTextCompare) = 0 Then
        private_FilterMatches = (VBA.Len(VBA.Trim$(cellText)) > 0)
        Exit Function
    End If

    ' Название rx сохранено как пользовательский маркер режима, но внутри
    ' находится не регулярное выражение, а LIKE-шаблон: поддерживаются только
    ' многосимвольные wildcard * и %. Сопоставляется вся строка целиком.
    If private_TryExtractLikeExpression(filterText, likeExpression) Then
        private_FilterMatches = private_MatchesLikeExpression(cellText, likeExpression)
        Exit Function
    End If

    ' Без обертки rx(...) сохраняется прежняя семантика: подстрока
    ' ищется без учета регистра, а * и % считаются обычным текстом.
    private_FilterMatches = (VBA.InStr(1, cellText, filterText, VBA.vbTextCompare) > 0)
End Function

Private Function private_TryExtractLikeExpression( _
    ByVal filterText As String, _
    ByRef outExpression As String _
) As Boolean
    outExpression = VBA.vbNullString
    filterText = VBA.Trim$(filterText)
    If VBA.Len(filterText) < 4 Then Exit Function
    If VBA.StrComp(VBA.Left$(filterText, 3), "rx(", VBA.vbTextCompare) <> 0 Then Exit Function
    If VBA.Right$(filterText, 1) <> ")" Then Exit Function

    outExpression = VBA.Mid$(filterText, 4, VBA.Len(filterText) - 4)
    private_TryExtractLikeExpression = True
End Function

Private Function private_MatchesLikeExpression( _
    ByVal cellText As String, _
    ByVal likeExpression As String _
) As Boolean
    Dim normalizedPattern As String

    On Error GoTo InvalidPattern
    normalizedPattern = private_NormalizeLikeExpression(likeExpression)
    private_MatchesLikeExpression = _
        (VBA.LCase$(cellText) Like VBA.LCase$(normalizedPattern))
    Exit Function

InvalidPattern:
    ex_Core.fn_Diagnostic_LogError "MultiSourcesView: invalid rx LIKE expression='" & _
        likeExpression & "' number=" & VBA.CStr(Err.Number) & _
        " description='" & Err.Description & "'"
    Err.Clear
End Function

Private Function private_NormalizeLikeExpression(ByVal expressionText As String) As String
    Dim resultText As String
    Dim currentChar As String
    Dim charIndex As Long

    ' В старом SQL-движке * и % преобразовывались в wildcard текущего
    ' диалекта. Здесь фильтр выполняется по уже загруженным данным, поэтому
    ' оба варианта приводятся к wildcard оператора VBA Like.
    '
    ' Остальные специальные символы VBA Like экранируются: пользователь
    ' пока не получает неявную поддержку ?, # и диапазонов [...]. Это важно,
    ' чтобы заявленный ограниченный синтаксис вел себя одинаково для любых
    ' текстовых данных.
    For charIndex = 1 To VBA.Len(expressionText)
        currentChar = VBA.Mid$(expressionText, charIndex, 1)
        Select Case currentChar
            Case "*", "%"
                resultText = resultText & "*"
            Case "?"
                resultText = resultText & "[?]"
            Case "#"
                resultText = resultText & "[#]"
            Case "["
                resultText = resultText & "[[]"
            Case "]"
                resultText = resultText & "[]]"
            Case Else
                resultText = resultText & currentChar
        End Select
    Next charIndex

    private_NormalizeLikeExpression = resultText
End Function
