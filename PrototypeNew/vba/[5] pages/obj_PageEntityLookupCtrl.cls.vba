VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PageEntityLookupCtrl"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const CONTROLLER_RUNTIME_OBJECT_KEY As String = "RuntimeObjects.PageEntityLookup.Controller"
Private Const CURRENT_ROW_RUNTIME_KEY As String = "RuntimeItems.EntityLookup.CurrentRow"
Private Const CANDIDATE_TABLES_RUNTIME_KEY As String = "RuntimeItems.EntityLookup.CandidateTables"
Private Const DEFAULT_CANDIDATES_SECTION_TEXT As String = "Candidates"
Private Const CANDIDATE_TABLE_SPAN_COLS As Long = 50

Private m_Page As obj_IPage
Private m_ConfigTable As obj_ConfigTable
Private m_CfgParser As obj_EntityLookupCfgParser
Private m_CandidateTable As obj_TableDynamic
Private m_ActiveLookupKey As String
Private m_ActiveCandidatesAt As String
Private m_SelectedValuesByColumnKey As Object
Private m_IsDisposed As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Me.Dispose
    On Error GoTo 0
End Sub

' //
' // Properties
' //
Public Property Get RuntimeObjectSourceKey() As String
    RuntimeObjectSourceKey = CONTROLLER_RUNTIME_OBJECT_KEY
End Property

Public Property Get ActiveCandidatesAt() As String
    ActiveCandidatesAt = m_ActiveCandidatesAt
End Property

' //
' // API
' //
Public Function Initialize(ByVal page As obj_IPage) As Boolean
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageEntityLookupCtrl.Initialize"
#End If
    Dim pageBase As obj_PageBase

    If page Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: PageEntityLookupCtrl initialization failed because page is not specified."
#End If
        Exit Function
    End If

    m_IsDisposed = False
    Set m_Page = page
    Set m_ConfigTable = Nothing
    Set m_CfgParser = Nothing
    Set m_CandidateTable = Nothing
    m_ActiveLookupKey = VBA.vbNullString
    private_SetFallbackCandidateLayout
    private_ResetSelectedValues

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If Not pageBase.RuntimeSources.SetObjectSource(CONTROLLER_RUNTIME_OBJECT_KEY, Me) Then Exit Function
    Initialize = True
End Function

Public Sub Dispose()
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageEntityLookupCtrl.Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True

    On Error Resume Next
    If Not m_CfgParser Is Nothing Then m_CfgParser.Dispose
    Set m_CfgParser = Nothing
    Set m_ConfigTable = Nothing
    Set m_CandidateTable = Nothing
    m_ActiveLookupKey = VBA.vbNullString
    private_SetFallbackCandidateLayout
    Set m_SelectedValuesByColumnKey = Nothing
    Set m_Page = Nothing
    On Error GoTo 0
End Sub

Public Function UpdateData(ByVal configControl As obj_ConfigControlVM) As Boolean
    Dim configTable As obj_ConfigTable
    Dim cfgParser As obj_EntityLookupCfgParser

    If configControl Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: PageEntityLookupCtrl.UpdateData failed because config control is not specified."
#End If
        Exit Function
    End If

    If Not configControl.TryBuildConfigTableFromRendered(configTable) Then Exit Function
    If configTable Is Nothing Then Exit Function

    Set cfgParser = New obj_EntityLookupCfgParser
    If Not cfgParser.Initialize(configTable) Then Exit Function

    On Error Resume Next
    If Not m_CfgParser Is Nothing Then m_CfgParser.Dispose
    On Error GoTo 0

    Set m_ConfigTable = configTable
    Set m_CfgParser = cfgParser
    UpdateData = True
End Function

Public Function PrepareLookupRuntime(Optional ByVal notifyChange As Boolean = False) As Boolean
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_PageEntityLookupCtrl.PrepareLookupRuntime"
#End If
    Set m_CandidateTable = Nothing
    m_ActiveLookupKey = VBA.vbNullString
    If Not private_ResetCandidateLayoutFromConfig() Then Exit Function
    If Not private_RegisterCandidateTables(False) Then Exit Function
    If Not private_RegisterCurrentRowRuntime(notifyChange) Then Exit Function
    PrepareLookupRuntime = True
End Function

Public Function ClearLookupCandidates(Optional ByVal renderNow As Boolean = True) As Boolean
    Set m_CandidateTable = Nothing
    m_ActiveLookupKey = VBA.vbNullString
    If Not private_ResetCandidateLayoutFromConfig() Then Exit Function
    If Not private_RegisterCandidateTables(False) Then Exit Function
    If renderNow Then
        If Not rt_PageManager.fn_RenderPage(m_Page, "entitylookup:clear-candidates") Then Exit Function
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

Public Function SetSelectedValue( _
    ByVal columnKey As String, _
    ByVal columnValue As String, _
    Optional ByVal notifyChange As Boolean = False _
) As Boolean
    columnKey = VBA.Trim$(columnKey)
    If VBA.Len(columnKey) = 0 Then Exit Function

    private_EnsureSelectedValues
    m_SelectedValuesByColumnKey(columnKey) = VBA.Trim$(columnValue)
    If Not private_RegisterCurrentRowRuntime(notifyChange) Then Exit Function
    SetSelectedValue = True
End Function

Public Function GetSelectedValue(ByVal columnKey As String) As String
    columnKey = VBA.Trim$(columnKey)
    If VBA.Len(columnKey) = 0 Then Exit Function
    If m_SelectedValuesByColumnKey Is Nothing Then Exit Function
    If Not m_SelectedValuesByColumnKey.Exists(columnKey) Then Exit Function

    GetSelectedValue = VBA.CStr(m_SelectedValuesByColumnKey(columnKey))
End Function

Public Function TryGetLookupKeys(ByRef outLookupKeys As Collection) As Boolean
    Set outLookupKeys = Nothing
    If m_CfgParser Is Nothing Then Exit Function
    TryGetLookupKeys = m_CfgParser.TryGetLookupKeys(outLookupKeys)
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
        MsgBox "PrototypeNew: lookup key is empty.", vbExclamation, "PrototypeNew / EntityLookup runtime"
        Exit Function
    End If
    If m_CfgParser Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "EntityLookup: search failed because config parser is not initialized."
#End If
        MsgBox "PrototypeNew: EntityLookup config is not initialized. Reopen the page from Main.", vbExclamation, "PrototypeNew / EntityLookup runtime"
        Exit Function
    End If

    If Not m_CfgParser.TryBuildLookupSqlParams(lookupKey, queryText, sqlParams, searchColumnAlias, resultColumnAliases) Then Exit Function
    If sqlParams Is Nothing Then Exit Function

    Set sqlTable = Nothing
    If Not ex_ExternalExcelSqlEngine.fn_TrySqlRequest(sqlParams, sqlTable) Then Exit Function
    outCandidateCount = private_GetTableRowCount(sqlTable)

    Set m_CandidateTable = Nothing
    m_ActiveLookupKey = VBA.vbNullString
    If private_ShouldShowCandidateTable(sqlTable, queryText, searchColumnAlias) Then
        If Not private_TrySetActiveCandidateLayout(lookupKey, sqlTable, searchColumnAlias) Then
            outCandidateCount = 0
            If Not private_RegisterCandidateTables(False) Then Exit Function
            If notifyChange Then
                If Not rt_PageManager.fn_RenderPage(m_Page, "entitylookup:search-candidates-layout-error") Then Exit Function
            End If
            private_SearchCandidates = True
            Exit Function
        End If
        Set m_CandidateTable = sqlTable
        m_ActiveLookupKey = lookupKey
    End If
    If Not m_CandidateTable Is Nothing Then m_CandidateTable.SectionTitle = private_GetLookupSectionCaption(lookupKey)
    If Not private_RegisterCandidateTables(False) Then Exit Function

    If notifyChange Then
        If Not rt_PageManager.fn_RenderPage(m_Page, "entitylookup:search-candidates") Then Exit Function
    End If

    private_SearchCandidates = True
End Function

Private Function private_GetTableRowCount(ByVal tableObj As obj_TableDynamic) As Long
    If tableObj Is Nothing Then Exit Function
    private_GetTableRowCount = tableObj.RowCount
End Function

Private Function private_ShouldShowCandidateTable( _
    ByVal tableObj As obj_TableDynamic, _
    ByVal queryText As String, _
    ByVal searchColumnAlias As String _
) As Boolean
    Dim candidateCount As Long

    candidateCount = private_GetTableRowCount(tableObj)
    If candidateCount <= 0 Then Exit Function
    If candidateCount > 1 Then
        private_ShouldShowCandidateTable = True
        Exit Function
    End If

    private_ShouldShowCandidateTable = Not private_HasSingleExactCandidate(tableObj, queryText, searchColumnAlias)
End Function

Private Function private_HasSingleExactCandidate( _
    ByVal tableObj As obj_TableDynamic, _
    ByVal queryText As String, _
    ByVal searchColumnAlias As String _
) As Boolean
    Dim candidateValue As String

    If private_GetTableRowCount(tableObj) <> 1 Then Exit Function
    If Not private_TryGetFirstCandidateSearchValue(tableObj, searchColumnAlias, candidateValue) Then Exit Function

    private_HasSingleExactCandidate = ( _
        VBA.StrComp( _
            private_NormalizeLookupText(candidateValue), _
            private_NormalizeLookupText(queryText), _
            VBA.vbTextCompare) = 0)
End Function

Private Function private_TryGetFirstCandidateSearchValue( _
    ByVal tableObj As obj_TableDynamic, _
    ByVal searchColumnAlias As String, _
    ByRef outValue As String _
) As Boolean
    Dim searchColumnIndex As Long
    Dim firstRow As obj_Row

    outValue = VBA.vbNullString
    If tableObj Is Nothing Then Exit Function
    If tableObj.RowCount <= 0 Then Exit Function

    searchColumnAlias = VBA.Trim$(searchColumnAlias)
    If VBA.Len(searchColumnAlias) = 0 Then Exit Function
    If Not tableObj.TryGetColumnIndexByAlias(searchColumnAlias, searchColumnIndex) Then
        If Not tableObj.TryGetColumnIndexByName(searchColumnAlias, searchColumnIndex) Then Exit Function
    End If
    If searchColumnIndex <= 0 Then Exit Function

    Set firstRow = tableObj.Rows.Item(1)
    If firstRow Is Nothing Then Exit Function

    outValue = firstRow.GetCellValue(searchColumnIndex)
    private_TryGetFirstCandidateSearchValue = True
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

Private Sub private_SetFallbackCandidateLayout()
    m_ActiveCandidatesAt = "r1c1"
End Sub

Private Function private_ResetCandidateLayoutFromConfig() As Boolean
    private_SetFallbackCandidateLayout
    private_ResetCandidateLayoutFromConfig = True
End Function

Private Function private_TrySetActiveCandidateLayout( _
    ByVal lookupKey As String, _
    ByVal tableObj As obj_TableDynamic, _
    ByVal searchColumnAlias As String _
) As Boolean
    Dim inputGridCol As Long
    Dim searchColumnIndex As Long
    Dim startGridCol As Long
    Dim atText As String

    If tableObj Is Nothing Then Exit Function
    If m_CfgParser Is Nothing Then Exit Function

    If tableObj.ColumnCount > CANDIDATE_TABLE_SPAN_COLS Then
        private_ShowCandidateLayoutError "Candidate table for lookup '" & lookupKey & "' has " & VBA.CStr(tableObj.ColumnCount) & _
            " columns, but current UI slot supports only " & VBA.CStr(CANDIDATE_TABLE_SPAN_COLS) & _
            ". Reduce ResultColumnsAliases or increase TableList spanColls."
        Exit Function
    End If

    If Not m_CfgParser.TryGetLookupInputGridColumn(lookupKey, inputGridCol) Then Exit Function
    If Not private_TryGetCandidateSearchColumnIndex(tableObj, searchColumnAlias, searchColumnIndex) Then
        private_ShowCandidateLayoutError "Failed to find search column alias '" & searchColumnAlias & "' in candidate columns for lookup '" & lookupKey & "'."
        Exit Function
    End If

    startGridCol = inputGridCol - searchColumnIndex + 1
    If startGridCol <= 0 Then
        private_ShowCandidateLayoutError "Candidate table for lookup '" & lookupKey & "' cannot be aligned: search column index is " & _
            VBA.CStr(searchColumnIndex) & ", but input grid column is " & VBA.CStr(inputGridCol) & _
            ". Move the candidate area left or place SearchColumnAlias earlier in ResultColumnsAliases."
        Exit Function
    End If

    atText = private_BuildCandidateAtText(startGridCol)
    m_ActiveCandidatesAt = atText

    private_TrySetActiveCandidateLayout = True
End Function

Private Function private_TryGetCandidateSearchColumnIndex( _
    ByVal tableObj As obj_TableDynamic, _
    ByVal searchColumnAlias As String, _
    ByRef outColumnIndex As Long _
) As Boolean
    outColumnIndex = 0
    If tableObj Is Nothing Then Exit Function

    searchColumnAlias = VBA.Trim$(searchColumnAlias)
    If VBA.Len(searchColumnAlias) = 0 Then Exit Function
    If Not tableObj.TryGetColumnIndexByAlias(searchColumnAlias, outColumnIndex) Then
        If Not tableObj.TryGetColumnIndexByName(searchColumnAlias, outColumnIndex) Then Exit Function
    End If
    If outColumnIndex <= 0 Then Exit Function

    private_TryGetCandidateSearchColumnIndex = True
End Function

Private Function private_BuildCandidateAtText(ByVal gridColumn As Long) As String
    If gridColumn <= 0 Then gridColumn = 1
    private_BuildCandidateAtText = "r1c" & VBA.CStr(gridColumn)
End Function

Private Sub private_ShowCandidateLayoutError(ByVal messageText As String)
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "EntityLookup layout: " & VBA.CStr(messageText)
#End If
    MsgBox "PrototypeNew: " & VBA.CStr(messageText), vbExclamation, "PrototypeNew / EntityLookup layout"
End Sub

Private Sub private_ResetSelectedValues()
    Set m_SelectedValuesByColumnKey = ex_Helpers.fn_CreateDictionaryTextCompare()
End Sub

Private Sub private_EnsureSelectedValues()
    If m_SelectedValuesByColumnKey Is Nothing Then private_ResetSelectedValues
End Sub

Private Function private_RegisterCurrentRowRuntime(ByVal notifyChange As Boolean) As Boolean
    Dim pageBase As obj_PageBase
    Dim runtimeSources As obj_PageRuntimeSources
    Dim currentRowTables As Collection

    If m_Page Is Nothing Then Exit Function
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set runtimeSources = pageBase.RuntimeSources
    If runtimeSources Is Nothing Then Exit Function

    Set currentRowTables = New Collection
    currentRowTables.Add private_BuildCurrentRowTable()

    If Not runtimeSources.RemoveItemsSource(VBA.LCase$(CURRENT_ROW_RUNTIME_KEY)) Then Exit Function
    If Not runtimeSources.SetItemsSource(VBA.LCase$(CURRENT_ROW_RUNTIME_KEY), currentRowTables, notifyChange) Then Exit Function

    private_RegisterCurrentRowRuntime = True
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

    If Not runtimeSources.RemoveItemsSource(VBA.LCase$(CANDIDATE_TABLES_RUNTIME_KEY)) Then Exit Function
    If Not runtimeSources.SetItemsSource(VBA.LCase$(CANDIDATE_TABLES_RUNTIME_KEY), candidateTables, notifyChange) Then Exit Function

    private_RegisterCandidateTables = True
End Function

Private Function private_BuildCurrentRowTable() As obj_TableDynamic
    Dim tableObj As obj_TableDynamic
    Dim rowObj As obj_Row
    Dim columnKeys As Collection
    Dim columnKeyObj As Variant
    Dim columnKey As String

    Set tableObj = New obj_TableDynamic
    tableObj.SectionTitle = "Draft row"

    Set columnKeys = private_GetCurrentRowColumnKeys()
    If columnKeys Is Nothing Then
        Set private_BuildCurrentRowTable = tableObj
        Exit Function
    End If
    If columnKeys.Count <= 0 Then
        Set private_BuildCurrentRowTable = tableObj
        Exit Function
    End If

    For Each columnKeyObj In columnKeys
        columnKey = VBA.Trim$(VBA.CStr(columnKeyObj))
        If VBA.Len(columnKey) = 0 Then GoTo ContinueColumnHeader
        If Not private_AddColumn(tableObj, private_GetColumnCaption(columnKey, columnKey)) Then Exit Function
ContinueColumnHeader:
    Next columnKeyObj

    Set rowObj = New obj_Row
    For Each columnKeyObj In columnKeys
        columnKey = VBA.Trim$(VBA.CStr(columnKeyObj))
        If VBA.Len(columnKey) = 0 Then GoTo ContinueColumnValue
        rowObj.PushCellRaw private_GetSelectedValue(columnKey)
ContinueColumnValue:
    Next columnKeyObj
    If Not tableObj.PushRow(rowObj) Then Exit Function

    Set private_BuildCurrentRowTable = tableObj
End Function

Private Function private_GetCurrentRowColumnKeys() As Collection
    Dim result As Collection

    If Not m_CfgParser Is Nothing Then
        If m_CfgParser.TryGetLookupColumnKeys(result) Then
            If Not result Is Nothing Then
                If result.Count > 0 Then
                    Set private_GetCurrentRowColumnKeys = result
                    Exit Function
                End If
            End If
        End If
    End If

    Set result = New Collection
    Set private_GetCurrentRowColumnKeys = result
End Function

Private Function private_GetSelectedValue(ByVal columnKey As String) As String
    private_GetSelectedValue = GetSelectedValue(columnKey)
End Function

Private Function private_AddColumn( _
    ByVal tableObj As obj_TableDynamic, _
    ByVal columnName As String _
) As Boolean
    Dim colObj As obj_Column

    If tableObj Is Nothing Then Exit Function
    Set colObj = New obj_Column
    colObj.Name = VBA.Trim$(columnName)
    If VBA.Len(colObj.Name) = 0 Then colObj.Name = "Column " & VBA.CStr(tableObj.ColumnCount + 1)
    colObj.Position = tableObj.ColumnCount + 1
    private_AddColumn = tableObj.PushColumn(colObj)
End Function

Private Function private_GetColumnCaption( _
    ByVal columnKey As String, _
    ByVal fallbackCaption As String _
) As String
    Dim captionText As String

    private_GetColumnCaption = fallbackCaption
    If m_CfgParser Is Nothing Then Exit Function
    If Not m_CfgParser.TryGetColumnCaption(columnKey, captionText) Then Exit Function
    captionText = VBA.Trim$(captionText)
    If VBA.Len(captionText) > 0 Then private_GetColumnCaption = captionText
End Function

Private Function private_GetLookupSectionCaption(ByVal lookupKey As String) As String
    Dim minCount As Long
    Dim sectionCaption As String

    private_GetLookupSectionCaption = DEFAULT_CANDIDATES_SECTION_TEXT
    If m_CfgParser Is Nothing Then Exit Function

    sectionCaption = DEFAULT_CANDIDATES_SECTION_TEXT
    If Not m_CfgParser.TryGetLookupCandidatesConfig(lookupKey, minCount, sectionCaption) Then Exit Function
    sectionCaption = VBA.Trim$(sectionCaption)
    If VBA.Len(sectionCaption) > 0 Then private_GetLookupSectionCaption = sectionCaption
End Function
