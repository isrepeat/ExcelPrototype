VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_EntityLookupFeature"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const DEFAULT_CANDIDATES_SECTION_TEXT As String = "Candidates"
Private Const AUXILIARY_COLUMN_STYLE_ALIAS As String = "__lookup_auxiliary__"
Private Const ADO_LONG_VALUE_CANDIDATE_TAG As String = "ado-long-value-candidate"
Private Const ADO_TEXT_LIMIT As Long = 255

Private m_Page As obj_IPage
Private m_ConfigTable As obj_ConfigTable
' Исходные и расширенные данные без проекции. Позиционирование применяется только к m_CandidateTable.
Private m_CandidateDataTable As obj_TableDynamic
Private m_CandidateTable As obj_TableDynamic
Private m_ActiveFormColumnAliases As Collection
Private m_EntityLookupCfgParser As obj_EntityLookupCfgParser
' ItemsSource key для таблицы кандидатов последнего активного поиска.
' UI TableList читает именно этот source и поэтому host может выбрать свой namespace.
Private m_CandidateTablesRuntimeKey As String
Private m_RenderReasonPrefix As String
Private m_ActiveLookupKey As String
Private m_ActiveSearchColumnAlias As String
Private m_IsDisposed As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogVerbose "lifecycle:" & TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Me.Dispose
    On Error GoTo 0
End Sub

' //
' // API
' //
Public Function Initialize( _
    ByVal page As obj_IPage, _
    ByVal candidateTablesRuntimeKey As String, _
    Optional ByVal renderReasonPrefix As String = "entitylookup" _
) As Boolean
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_EntityLookupFeature.Initialize"
#End If

    If page Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: EntityLookupFeature initialization failed because page is not specified."
#End If
        Exit Function
    End If
    candidateTablesRuntimeKey = VBA.Trim$(candidateTablesRuntimeKey)
    renderReasonPrefix = VBA.Trim$(renderReasonPrefix)
    If VBA.Len(candidateTablesRuntimeKey) = 0 Then
        VBA.MsgBox "PrototypeNew: EntityLookupFeature candidate tables runtime key is empty.", vbExclamation, "PrototypeNew / EntityLookup runtime"
        Exit Function
    End If
    If VBA.Len(renderReasonPrefix) = 0 Then renderReasonPrefix = "entitylookup"

    m_IsDisposed = False
    Set m_Page = page
    ' Host-страница передает namespace runtime items. Благодаря этому один и тот же
    ' lookup feature можно безопасно использовать на разных страницах/режимах.
    m_CandidateTablesRuntimeKey = candidateTablesRuntimeKey
    m_RenderReasonPrefix = renderReasonPrefix
    Set m_ConfigTable = Nothing
    Set m_EntityLookupCfgParser = Nothing
    Set m_CandidateTable = Nothing
    Set m_CandidateDataTable = Nothing
    Set m_ActiveFormColumnAliases = Nothing
    m_ActiveLookupKey = VBA.vbNullString
    m_ActiveSearchColumnAlias = VBA.vbNullString

    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_EntityLookupFeature.Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True

    On Error Resume Next
    If Not m_EntityLookupCfgParser Is Nothing Then m_EntityLookupCfgParser.Dispose
    Set m_EntityLookupCfgParser = Nothing
    Set m_ConfigTable = Nothing
    Set m_CandidateTable = Nothing
    Set m_CandidateDataTable = Nothing
    Set m_ActiveFormColumnAliases = Nothing
    m_ActiveLookupKey = VBA.vbNullString
    m_ActiveSearchColumnAlias = VBA.vbNullString
    m_CandidateTablesRuntimeKey = VBA.vbNullString
    m_RenderReasonPrefix = VBA.vbNullString
    Set m_Page = Nothing
    On Error GoTo 0
End Sub

Public Function UpdateData(ByVal configControl As obj_ConfigControlVM) As Boolean
    Dim configTable As obj_ConfigTable

    If configControl Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: EntityLookupFeature.UpdateData failed because config control is not specified."
#End If
        Exit Function
    End If

    If Not configControl.TryBuildConfigTableFromRendered(configTable) Then Exit Function
    If configTable Is Nothing Then Exit Function

    UpdateData = UpdateDataFromConfigTable(configTable)
End Function

Public Function UpdateDataFromConfigTable(ByVal configTable As obj_ConfigTable) As Boolean
    Dim cfgParser As obj_EntityLookupCfgParser

    If configTable Is Nothing Then Exit Function

    Set cfgParser = New obj_EntityLookupCfgParser
    If Not cfgParser.Initialize(configTable) Then Exit Function

    On Error Resume Next
    If Not m_EntityLookupCfgParser Is Nothing Then m_EntityLookupCfgParser.Dispose
    On Error GoTo 0

    Set m_ConfigTable = configTable
    Set m_EntityLookupCfgParser = cfgParser
    UpdateDataFromConfigTable = True
End Function

Public Function PrepareLookupRuntime(Optional ByVal notifyChange As Boolean = False) As Boolean
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_EntityLookupFeature.PrepareLookupRuntime"
#End If
    Set m_CandidateTable = Nothing
    Set m_CandidateDataTable = Nothing
    m_ActiveLookupKey = VBA.vbNullString
    m_ActiveSearchColumnAlias = VBA.vbNullString
    If Not private_RegisterCandidateTables(False) Then Exit Function
    PrepareLookupRuntime = True
End Function

Public Function ClearLookupCandidates(Optional ByVal renderNow As Boolean = True) As Boolean
    Set m_CandidateTable = Nothing
    Set m_CandidateDataTable = Nothing
    m_ActiveLookupKey = VBA.vbNullString
    m_ActiveSearchColumnAlias = VBA.vbNullString
    If Not private_RegisterCandidateTables(False) Then Exit Function
    If renderNow Then
        If Not rt_PageManager.fn_RenderPage(m_Page, private_BuildRenderReason("clear-candidates")) Then Exit Function
    End If
    ClearLookupCandidates = True
End Function

Public Function SearchCandidates( _
    ByVal lookupKey As String, _
    ByVal queryText As String, _
    ByRef outCandidateCount As Long, _
    Optional ByVal notifyChange As Boolean = True _
) As Boolean
    SearchCandidates = private_SearchCandidates(lookupKey, queryText, outCandidateCount, notifyChange)
End Function

' Показывает уже подготовленные host-страницей данные через тот же layout,
' который используется обычным EntityLookup. Источник должен содержать алиасы
' полей активной формы; остальные колонки будут показаны как справочные.
Public Function ShowPreparedCandidates( _
    ByVal lookupKey As String, _
    ByVal searchColumnAlias As String, _
    ByVal sectionTitle As String, _
    ByVal candidateDataTable As obj_TableDynamic, _
    Optional ByVal notifyChange As Boolean = True _
) As Boolean
    If candidateDataTable Is Nothing Then Exit Function
    If m_ActiveFormColumnAliases Is Nothing Then Exit Function

    Set m_CandidateTable = Nothing
    Set m_CandidateDataTable = candidateDataTable
    m_ActiveLookupKey = VBA.Trim$(lookupKey)
    m_ActiveSearchColumnAlias = VBA.Trim$(searchColumnAlias)
    If VBA.Len(m_ActiveLookupKey) = 0 Or _
       VBA.Len(m_ActiveSearchColumnAlias) = 0 Then Exit Function

    If Not private_ProjectCandidateTable( _
        m_CandidateDataTable, m_CandidateTable) Then Exit Function
    m_CandidateTable.SectionTitle = VBA.Trim$(sectionTitle)
    If Not private_RegisterCandidateTables(False) Then Exit Function
    If notifyChange Then
        If Not rt_PageManager.fn_RenderPage( _
            m_Page, private_BuildRenderReason("prepared-candidates")) Then Exit Function
    End If
    ShowPreparedCandidates = True
End Function

Public Function TryGetFormColumnKeys(ByRef outColumnKeys As Collection) As Boolean
    Set outColumnKeys = Nothing
    If m_EntityLookupCfgParser Is Nothing Then Exit Function
    TryGetFormColumnKeys = m_EntityLookupCfgParser.TryGetLookupColumnKeys(outColumnKeys)
End Function

Public Function TryGetPrimaryCandidateColumnCount(ByRef outColumnCount As Long) As Boolean
    outColumnCount = 0
    If m_ActiveFormColumnAliases Is Nothing Then Exit Function
    outColumnCount = m_ActiveFormColumnAliases.Count
    If outColumnCount <= 0 Then Exit Function
    TryGetPrimaryCandidateColumnCount = True
End Function

' Заменяет текущую runtime-схему формы. Хост вызывает метод после отрисовки
' видимости режима/профиля и перед запуском Lookup-запроса.
Public Function UpdateActiveFormColumns(ByVal columnAliases As Collection) As Boolean
    Dim normalized As Collection
    Dim seen As Object
    Dim aliasObj As Variant
    Dim aliasText As String

    If columnAliases Is Nothing Then
        VBA.MsgBox "PrototypeNew: active EntityLookup form column list is not specified.", _
            vbExclamation, "PrototypeNew / EntityLookup runtime"
        Exit Function
    End If

    Set normalized = New Collection
    Set seen = ex_Helpers.fn_CreateDictionaryTextCompare()
    For Each aliasObj In columnAliases
        aliasText = VBA.Trim$(VBA.CStr(aliasObj))
        If VBA.Len(aliasText) = 0 Then GoTo ContinueAlias
        If seen.Exists(aliasText) Then
            VBA.MsgBox "PrototypeNew: active EntityLookup form contains duplicate column alias '" & aliasText & "'.", _
                vbExclamation, "PrototypeNew / EntityLookup runtime"
            Exit Function
        End If
        seen(aliasText) = True
        normalized.Add aliasText
ContinueAlias:
    Next aliasObj

    If normalized.Count <= 0 Then
        VBA.MsgBox "PrototypeNew: active EntityLookup form has no visible columns.", _
            vbExclamation, "PrototypeNew / EntityLookup runtime"
        Exit Function
    End If

    Set m_ActiveFormColumnAliases = normalized
    UpdateActiveFormColumns = True
End Function

' Добавляет результат другого Lookup к непроецированным данным кандидатов.
' После объединения проекция один раз перестраивается по порядку активной формы.
Public Function ExtendCandidates( _
    ByVal extensionLookupKey As String, _
    ByVal joinColumnAlias As String, _
    ByVal queryText As String, _
    ByVal candidateSelector As obj_ILookupCandidateSelector, _
    Optional ByVal notifyChange As Boolean = True _
) As Boolean
    Dim sqlParams As obj_SqlParams
    Dim extensionTable As obj_TableDynamic
    Dim ignoredSearchAlias As String
    Dim ignoredResultAliases As Collection

    If m_CandidateDataTable Is Nothing Then
        If notifyChange Then
            If Not rt_PageManager.fn_RenderPage(m_Page, private_BuildRenderReason("extend-candidates-empty")) Then Exit Function
        End If
        ExtendCandidates = True
        Exit Function
    End If
    If m_EntityLookupCfgParser Is Nothing Then Exit Function

    If Not m_EntityLookupCfgParser.TryBuildLookupSqlParams( _
        extensionLookupKey, queryText, sqlParams, ignoredSearchAlias, ignoredResultAliases) Then Exit Function
    If Not ex_ExternalExcelSqlEngine.fn_TrySqlRequest(sqlParams, extensionTable) Then Exit Function

    ' Пустой результат дополнительного источника не является ошибкой поиска.
    ' В этом случае показываем кандидатов основного запроса без расширенных данных.
    If Not extensionTable Is Nothing Then
        If extensionTable.RowCount > 0 Then
            If Not private_MergeCandidateExtension( _
                extensionTable, joinColumnAlias, candidateSelector) Then Exit Function
        End If
    End If
    If Not private_ProjectCandidateTable(m_CandidateDataTable, m_CandidateTable) Then
        Set m_CandidateTable = Nothing
        If Not private_RegisterCandidateTables(False) Then Exit Function
        If notifyChange Then
            If Not rt_PageManager.fn_RenderPage(m_Page, private_BuildRenderReason("invalid-candidate-columns")) Then Exit Function
        End If
        Exit Function
    End If
    m_CandidateTable.SectionTitle = private_GetLookupSectionCaption(m_ActiveLookupKey)
    If Not private_RegisterCandidateTables(False) Then Exit Function
    If notifyChange Then
        If Not rt_PageManager.fn_RenderPage(m_Page, private_BuildRenderReason("extend-candidates")) Then Exit Function
    End If
    ExtendCandidates = True
End Function

Public Function TryGetActiveCandidatesContext( _
    ByRef outCfgParser As obj_EntityLookupCfgParser, _
    ByRef outLookupKey As String, _
    ByRef outCandidateTable As obj_TableDynamic, _
    ByRef outSearchColumnAlias As String _
) As Boolean
    Set outCfgParser = Nothing
    outLookupKey = VBA.vbNullString
    Set outCandidateTable = Nothing
    outSearchColumnAlias = VBA.vbNullString

    If m_EntityLookupCfgParser Is Nothing Then Exit Function
    If m_CandidateTable Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(m_ActiveLookupKey)) = 0 Then Exit Function
    If VBA.Len(VBA.Trim$(m_ActiveSearchColumnAlias)) = 0 Then Exit Function

    Set outCfgParser = m_EntityLookupCfgParser
    outLookupKey = m_ActiveLookupKey
    Set outCandidateTable = m_CandidateTable
    outSearchColumnAlias = m_ActiveSearchColumnAlias
    TryGetActiveCandidatesContext = True
End Function

Public Function TryGetLookupKeys(ByRef outLookupKeys As Collection) As Boolean
    Set outLookupKeys = Nothing
    If m_EntityLookupCfgParser Is Nothing Then Exit Function
    TryGetLookupKeys = m_EntityLookupCfgParser.TryGetLookupKeys(outLookupKeys)
End Function

' //
' // Internal
' //
Private Function private_SearchCandidates( _
    ByVal lookupKey As String, _
    ByVal queryText As String, _
    ByRef outCandidateCount As Long, _
    ByVal notifyChange As Boolean _
) As Boolean
    Dim sqlParams As obj_SqlParams
    Dim sqlTable As obj_TableDynamic
    Dim searchColumnAlias As String
    Dim resultColumnAliases As Collection

    outCandidateCount = 0
    lookupKey = VBA.Trim$(lookupKey)
    queryText = VBA.Trim$(queryText)
    If VBA.Len(lookupKey) = 0 Then
        VBA.MsgBox "PrototypeNew: lookup key is empty.", vbExclamation, "PrototypeNew / EntityLookup runtime"
        Exit Function
    End If
    If m_EntityLookupCfgParser Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "EntityLookup: search failed because config parser is not initialized."
#End If
        VBA.MsgBox "PrototypeNew: EntityLookup config is not initialized. Reopen the page from Main.", vbExclamation, "PrototypeNew / EntityLookup runtime"
        Exit Function
    End If
    If m_ActiveFormColumnAliases Is Nothing Then
        VBA.MsgBox "PrototypeNew: active EntityLookup form columns were not provided before lookup '" & lookupKey & "'.", _
            vbExclamation, "PrototypeNew / EntityLookup runtime"
        Exit Function
    End If

    If Not m_EntityLookupCfgParser.TryBuildLookupSqlParams(lookupKey, queryText, sqlParams, searchColumnAlias, resultColumnAliases) Then Exit Function
    If sqlParams Is Nothing Then Exit Function

    Set sqlTable = Nothing
    If Not ex_ExternalExcelSqlEngine.fn_TrySqlRequest(sqlParams, sqlTable) Then Exit Function
    outCandidateCount = private_GetTableRowCount(sqlTable)

    Set m_CandidateTable = Nothing
    Set m_CandidateDataTable = Nothing
    m_ActiveLookupKey = VBA.vbNullString
    m_ActiveSearchColumnAlias = VBA.vbNullString
    If private_ShouldShowCandidateTable(sqlTable) Then
        Set m_CandidateDataTable = sqlTable
        m_ActiveLookupKey = lookupKey
        m_ActiveSearchColumnAlias = searchColumnAlias
        If Not private_ProjectCandidateTable(m_CandidateDataTable, m_CandidateTable) Then
            Set m_CandidateTable = Nothing
            If Not private_RegisterCandidateTables(False) Then Exit Function
            If notifyChange Then
                If Not rt_PageManager.fn_RenderPage(m_Page, private_BuildRenderReason("invalid-candidate-columns")) Then Exit Function
            End If
            Exit Function
        End If
    End If
    If Not m_CandidateTable Is Nothing Then m_CandidateTable.SectionTitle = private_GetLookupSectionCaption(lookupKey)
    If Not private_RegisterCandidateTables(False) Then Exit Function

    If notifyChange Then
        If Not rt_PageManager.fn_RenderPage(m_Page, private_BuildRenderReason("search-candidates")) Then Exit Function
    End If

    private_SearchCandidates = True
End Function

Private Function private_GetTableRowCount(ByVal tableObj As obj_TableDynamic) As Long
    If tableObj Is Nothing Then Exit Function
    private_GetTableRowCount = tableObj.RowCount
End Function

Private Function private_MergeCandidateExtension( _
    ByVal extensionTable As obj_TableDynamic, _
    ByVal joinColumnAlias As String, _
    ByVal candidateSelector As obj_ILookupCandidateSelector _
) As Boolean
    Dim i As Long
    Dim aliasText As String
    Dim existingIndex As Long
    Dim extensionIndex As Long
    Dim columnObj As obj_Column
    Dim candidateRow As obj_Row
    Dim extensionRow As obj_Row
    Dim candidateJoinIndex As Long
    Dim extensionJoinIndex As Long
    Dim candidateKey As String
    Dim extensionKey As String
    Dim destinationIndex As Long
    Dim matchingRowIndexes As Collection
    Dim selectedExtensionRowIndex As Long

    If extensionTable Is Nothing Then Exit Function
    If m_CandidateDataTable Is Nothing Then Exit Function
    If Not private_TryGetColumnIndex(m_CandidateDataTable, joinColumnAlias, candidateJoinIndex) Then Exit Function
    If Not private_TryGetColumnIndex(extensionTable, joinColumnAlias, extensionJoinIndex) Then Exit Function

    ' Добавляем колонки расширения по алиасам. Физический порядок источника не важен:
    ' private_ProjectCandidateTable затем расставит их по активной форме.
    For extensionIndex = 1 To extensionTable.ColumnCount
        Set columnObj = extensionTable.Columns.Item(extensionIndex)
        aliasText = private_GetColumnPrimaryAlias(columnObj)
        If VBA.Len(aliasText) = 0 Then aliasText = columnObj.Name
        If VBA.StrComp(aliasText, joinColumnAlias, VBA.vbTextCompare) <> 0 Then
            If Not private_TryGetColumnIndex(m_CandidateDataTable, aliasText, existingIndex) Then
                If Not m_CandidateDataTable.PushColumn(columnObj) Then Exit Function
            End If
        End If
    Next extensionIndex

    If Not private_TryGetColumnIndex(m_CandidateDataTable, joinColumnAlias, candidateJoinIndex) Then Exit Function
    For i = 1 To m_CandidateDataTable.RowCount
        Set candidateRow = m_CandidateDataTable.Rows.Item(i)
        candidateKey = private_NormalizeLookupText(candidateRow.GetCellValue(candidateJoinIndex))
        If VBA.Len(candidateKey) = 0 Then GoTo ContinueCandidate

        Set matchingRowIndexes = New Collection
        For existingIndex = 1 To extensionTable.RowCount
            Set extensionRow = extensionTable.Rows.Item(existingIndex)
            extensionKey = private_NormalizeLookupText(extensionRow.GetCellValue(extensionJoinIndex))
            If VBA.StrComp(candidateKey, extensionKey, VBA.vbTextCompare) <> 0 Then GoTo ContinueExtension
            matchingRowIndexes.Add existingIndex
ContinueExtension:
        Next existingIndex

        selectedExtensionRowIndex = 0
        If matchingRowIndexes.Count > 0 Then
            If candidateSelector Is Nothing Then
                selectedExtensionRowIndex = VBA.CLng(matchingRowIndexes.Item(1))
            ElseIf Not candidateSelector.TrySelectCandidateRow( _
                m_CandidateDataTable, candidateRow, extensionTable, matchingRowIndexes, selectedExtensionRowIndex) Then
                Exit Function
            End If
        End If
        If selectedExtensionRowIndex <= 0 Then GoTo ContinueCandidate
        If selectedExtensionRowIndex > extensionTable.RowCount Then Exit Function
        Set extensionRow = extensionTable.Rows.Item(selectedExtensionRowIndex)

        For extensionIndex = 1 To extensionTable.ColumnCount
            Set columnObj = extensionTable.Columns.Item(extensionIndex)
            aliasText = private_GetColumnPrimaryAlias(columnObj)
            If VBA.Len(aliasText) = 0 Then aliasText = columnObj.Name
            If VBA.StrComp(aliasText, joinColumnAlias, VBA.vbTextCompare) <> 0 Then
                If Not private_TryGetColumnIndex(m_CandidateDataTable, aliasText, destinationIndex) Then Exit Function
                If Not candidateRow.SetCellRaw(destinationIndex, extensionRow.GetCellValue(extensionIndex)) Then Exit Function
            End If
        Next extensionIndex
ContinueCandidate:
    Next i

    private_MergeCandidateExtension = True
End Function

Private Function private_ProjectCandidateTable( _
    ByVal sourceTable As obj_TableDynamic, _
    ByRef outProjectedTable As obj_TableDynamic _
) As Boolean
    Dim activeIndexByAlias As Object
    Dim sourceIndexByAlias As Object
    Dim sourceColumn As obj_Column
    Dim projectedColumn As obj_Column
    Dim sourceRow As obj_Row
    Dim projectedRow As obj_Row
    Dim aliasText As String
    Dim aliasObj As Variant
    Dim i As Long
    Dim sourceIndex As Long
    Dim auxiliaryAliases As Collection
    Dim projectedAuxColumn As obj_Column

    Set outProjectedTable = Nothing
    If sourceTable Is Nothing Then Exit Function
    If m_ActiveFormColumnAliases Is Nothing Then Exit Function

    Set activeIndexByAlias = ex_Helpers.fn_CreateDictionaryTextCompare()
    For i = 1 To m_ActiveFormColumnAliases.Count
        aliasText = VBA.Trim$(VBA.CStr(m_ActiveFormColumnAliases.Item(i)))
        activeIndexByAlias(aliasText) = i
    Next i

    Set sourceIndexByAlias = ex_Helpers.fn_CreateDictionaryTextCompare()
    Set auxiliaryAliases = New Collection
    For i = 1 To sourceTable.ColumnCount
        Set sourceColumn = sourceTable.Columns.Item(i)
        aliasText = private_GetColumnPrimaryAlias(sourceColumn)
        If VBA.Len(aliasText) = 0 Then aliasText = VBA.Trim$(sourceColumn.Name)
        If VBA.Len(aliasText) = 0 Then
            VBA.MsgBox "PrototypeNew: lookup '" & m_ActiveLookupKey & "' returned a column without an alias.", _
                vbExclamation, "PrototypeNew / EntityLookup runtime"
            Exit Function
        End If
        ' Lookup может вернуть поля другого режима или профиля. Они остаются
        ' в исходных данных, но не выводятся в текущем представлении.
        If Not activeIndexByAlias.Exists(aliasText) Then
            ' Алиасы полей формы имеют префикс "_". Скрытое в активном профиле поле
            ' пропускаем. Обычные алиасы источника считаются служебными колонками
            ' и добавляются после полного диапазона формы.
            If VBA.Left$(aliasText, 1) <> "_" Then auxiliaryAliases.Add aliasText
            GoTo ContinueSourceColumn
        End If
        If sourceIndexByAlias.Exists(aliasText) Then
            VBA.MsgBox "PrototypeNew: lookup '" & m_ActiveLookupKey & "' returned duplicate target column '" & aliasText & "'.", _
                vbExclamation, "PrototypeNew / EntityLookup runtime"
            Exit Function
        End If
        sourceIndexByAlias(aliasText) = i
ContinueSourceColumn:
    Next i
    If sourceIndexByAlias.Count <= 0 Then
        VBA.MsgBox "PrototypeNew: lookup '" & m_ActiveLookupKey & _
            "' returned no columns that can be mapped to the active form.", _
            vbExclamation, "PrototypeNew / EntityLookup runtime"
        Exit Function
    End If

    Set outProjectedTable = New obj_TableDynamic
    If Not outProjectedTable.Initialize() Then Exit Function
    ' Основной диапазон кандидата всегда повторяет полную схему активной формы.
    ' Отсутствующие значения становятся пустыми ячейками, а служебные колонки
    ' источника добавляются только после этого фиксированного диапазона.
    For i = 1 To m_ActiveFormColumnAliases.Count
        aliasText = VBA.Trim$(VBA.CStr(m_ActiveFormColumnAliases.Item(i)))
        If sourceIndexByAlias.Exists(aliasText) Then
            Set sourceColumn = sourceTable.Columns.Item(VBA.CLng(sourceIndexByAlias(aliasText)))
            If Not outProjectedTable.PushColumn(sourceColumn) Then Exit Function
        Else
            Set projectedColumn = New obj_Column
            ' Пробел нулевой ширины сохраняет заголовок визуально пустым и не даёт
            ' DynamicTable заменить его автоматически созданным названием ColN.
            projectedColumn.Name = VBA.ChrW$(8203)
            If Not projectedColumn.AddAlias(aliasText) Then Exit Function
            If Not outProjectedTable.PushColumn(projectedColumn) Then Exit Function
        End If
    Next i
    For Each aliasObj In auxiliaryAliases
        aliasText = VBA.Trim$(VBA.CStr(aliasObj))
        If Not private_TryGetColumnIndex(sourceTable, aliasText, sourceIndex) Then Exit Function
        Set sourceColumn = sourceTable.Columns.Item(sourceIndex)
        If Not outProjectedTable.PushColumn(sourceColumn) Then Exit Function
        Set projectedAuxColumn = outProjectedTable.Columns.Item(outProjectedTable.ColumnCount)
        If projectedAuxColumn Is Nothing Then Exit Function
        If Not projectedAuxColumn.AddAlias(AUXILIARY_COLUMN_STYLE_ALIAS) Then Exit Function
    Next aliasObj

    For i = 1 To sourceTable.RowCount
        Set sourceRow = sourceTable.Rows.Item(i)
        Set projectedRow = New obj_Row
        For Each aliasObj In m_ActiveFormColumnAliases
            aliasText = VBA.Trim$(VBA.CStr(aliasObj))
            If sourceIndexByAlias.Exists(aliasText) Then
                sourceIndex = VBA.CLng(sourceIndexByAlias(aliasText))
                projectedRow.PushCellRaw sourceRow.GetCellValue(sourceIndex)
                If VBA.Len(VBA.CStr(sourceRow.GetCellValue(sourceIndex))) = ADO_TEXT_LIMIT Then
                    If Not projectedRow.AddCellTag(projectedRow.CellCount, ADO_LONG_VALUE_CANDIDATE_TAG) Then Exit Function
                End If
            Else
                projectedRow.PushCellRaw VBA.vbNullString
            End If
        Next aliasObj
        For Each aliasObj In auxiliaryAliases
            aliasText = VBA.Trim$(VBA.CStr(aliasObj))
            If Not private_TryGetColumnIndex(sourceTable, aliasText, sourceIndex) Then Exit Function
            projectedRow.PushCellRaw sourceRow.GetCellValue(sourceIndex)
            If VBA.Len(VBA.CStr(sourceRow.GetCellValue(sourceIndex))) = ADO_TEXT_LIMIT Then
                If Not projectedRow.AddCellTag(projectedRow.CellCount, ADO_LONG_VALUE_CANDIDATE_TAG) Then Exit Function
            End If
        Next aliasObj
        If Not outProjectedTable.PushRow(projectedRow) Then Exit Function
    Next i

    private_ProjectCandidateTable = True
End Function

Private Function private_TryGetColumnIndex( _
    ByVal tableObj As obj_TableDynamic, _
    ByVal aliasText As String, _
    ByRef outIndex As Long _
) As Boolean
    outIndex = 0
    If tableObj Is Nothing Then Exit Function
    If tableObj.TryGetColumnIndexByAlias(aliasText, outIndex) Then
        private_TryGetColumnIndex = True
    ElseIf tableObj.TryGetColumnIndexByName(aliasText, outIndex) Then
        private_TryGetColumnIndex = True
    End If
End Function

Private Function private_GetColumnPrimaryAlias(ByVal columnObj As obj_Column) As String
    Dim aliases As Collection
    If columnObj Is Nothing Then Exit Function
    Set aliases = columnObj.Aliases
    If Not aliases Is Nothing Then
        If aliases.Count > 0 Then private_GetColumnPrimaryAlias = VBA.Trim$(VBA.CStr(aliases.Item(1)))
    End If
End Function

Private Function private_ShouldShowCandidateTable(ByVal tableObj As obj_TableDynamic) As Boolean
    ' Любой непустой результат показывается пользователю, даже если найдена
    ' единственная строка и ключ совпал полностью.
    private_ShouldShowCandidateTable = (private_GetTableRowCount(tableObj) > 0)
End Function

Private Function private_NormalizeLookupText(ByVal valueText As String) As String
    valueText = VBA.CStr(valueText)
    valueText = VBA.Replace$(valueText, VBA.vbCr, " ")
    valueText = VBA.Replace$(valueText, VBA.vbLf, " ")
    valueText = VBA.Replace$(valueText, VBA.vbTab, " ")
    valueText = VBA.Replace$(valueText, VBA.ChrW$(160), " ")
    Do While VBA.InStr(1, valueText, "  ", VBA.vbBinaryCompare) > 0
        valueText = VBA.Replace$(valueText, "  ", " ")
    Loop
    private_NormalizeLookupText = VBA.LCase$(VBA.Trim$(valueText))
End Function

Private Function private_RegisterCandidateTables(ByVal notifyChange As Boolean) As Boolean
    Dim pageBase As obj_PageBase
    Dim runtimeSources As obj_PageRuntimeSources
    Dim candidateTables As Collection

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set runtimeSources = pageBase.RuntimeSources
    If runtimeSources Is Nothing Then Exit Function

    Set candidateTables = New Collection
    If Not m_CandidateTable Is Nothing Then candidateTables.Add m_CandidateTable

    If Not runtimeSources.RemoveItemsSource(VBA.LCase$(m_CandidateTablesRuntimeKey)) Then Exit Function
    If Not runtimeSources.SetItemsSource(VBA.LCase$(m_CandidateTablesRuntimeKey), candidateTables, notifyChange) Then Exit Function

    private_RegisterCandidateTables = True
End Function

Private Function private_GetLookupSectionCaption(ByVal lookupKey As String) As String
    Dim minCount As Long
    Dim sectionCaption As String

    private_GetLookupSectionCaption = DEFAULT_CANDIDATES_SECTION_TEXT
    If m_EntityLookupCfgParser Is Nothing Then Exit Function

    sectionCaption = DEFAULT_CANDIDATES_SECTION_TEXT
    If Not m_EntityLookupCfgParser.TryGetLookupCandidatesConfig(lookupKey, minCount, sectionCaption) Then Exit Function
    sectionCaption = VBA.Trim$(sectionCaption)
    If VBA.Len(sectionCaption) > 0 Then private_GetLookupSectionCaption = sectionCaption
End Function

Private Function private_BuildRenderReason(ByVal reasonSuffix As String) As String
    reasonSuffix = VBA.Trim$(reasonSuffix)
    If VBA.Len(reasonSuffix) = 0 Then
        private_BuildRenderReason = m_RenderReasonPrefix
    Else
        private_BuildRenderReason = m_RenderReasonPrefix & ":" & reasonSuffix
    End If
End Function
