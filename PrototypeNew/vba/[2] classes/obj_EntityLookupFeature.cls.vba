VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_EntityLookupFeature"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const DEFAULT_CANDIDATES_SECTION_TEXT As String = "Candidates"

Private m_Page As obj_IPage
Private m_ConfigTable As obj_ConfigTable
Private m_CandidateTable As obj_TableDynamic
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
        MsgBox "PrototypeNew: EntityLookupFeature candidate tables runtime key is empty.", vbExclamation, "PrototypeNew / EntityLookup runtime"
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
    m_ActiveLookupKey = VBA.vbNullString
    m_ActiveSearchColumnAlias = VBA.vbNullString
    m_CandidateTablesRuntimeKey = VBA.vbNullString
    m_RenderReasonPrefix = VBA.vbNullString
    Set m_Page = Nothing
    On Error GoTo 0
End Sub

Public Function UpdateData(ByVal configControl As obj_ConfigControlVM) As Boolean
    Dim configTable As obj_ConfigTable
    Dim cfgParser As obj_EntityLookupCfgParser

    If configControl Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "PrototypeNew: EntityLookupFeature.UpdateData failed because config control is not specified."
#End If
        Exit Function
    End If

    If Not configControl.TryBuildConfigTableFromRendered(configTable) Then Exit Function
    If configTable Is Nothing Then Exit Function

    Set cfgParser = New obj_EntityLookupCfgParser
    If Not cfgParser.Initialize(configTable) Then Exit Function

    On Error Resume Next
    If Not m_EntityLookupCfgParser Is Nothing Then m_EntityLookupCfgParser.Dispose
    On Error GoTo 0

    Set m_ConfigTable = configTable
    Set m_EntityLookupCfgParser = cfgParser
    UpdateData = True
End Function

Public Function PrepareLookupRuntime(Optional ByVal notifyChange As Boolean = False) As Boolean
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "enter:obj_EntityLookupFeature.PrepareLookupRuntime"
#End If
    Set m_CandidateTable = Nothing
    m_ActiveLookupKey = VBA.vbNullString
    m_ActiveSearchColumnAlias = VBA.vbNullString
    If Not private_RegisterCandidateTables(False) Then Exit Function
    PrepareLookupRuntime = True
End Function

Public Function ClearLookupCandidates(Optional ByVal renderNow As Boolean = True) As Boolean
    Set m_CandidateTable = Nothing
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
        MsgBox "PrototypeNew: lookup key is empty.", vbExclamation, "PrototypeNew / EntityLookup runtime"
        Exit Function
    End If
    If m_EntityLookupCfgParser Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "EntityLookup: search failed because config parser is not initialized."
#End If
        MsgBox "PrototypeNew: EntityLookup config is not initialized. Reopen the page from Main.", vbExclamation, "PrototypeNew / EntityLookup runtime"
        Exit Function
    End If

    If Not m_EntityLookupCfgParser.TryBuildLookupSqlParams(lookupKey, queryText, sqlParams, searchColumnAlias, resultColumnAliases) Then Exit Function
    If sqlParams Is Nothing Then Exit Function

    Set sqlTable = Nothing
    If Not ex_ExternalExcelSqlEngine.fn_TrySqlRequest(sqlParams, sqlTable) Then Exit Function
    outCandidateCount = private_GetTableRowCount(sqlTable)

    Set m_CandidateTable = Nothing
    m_ActiveLookupKey = VBA.vbNullString
    m_ActiveSearchColumnAlias = VBA.vbNullString
    If private_ShouldShowCandidateTable(sqlTable, queryText, searchColumnAlias) Then
        Set m_CandidateTable = sqlTable
        m_ActiveLookupKey = lookupKey
        m_ActiveSearchColumnAlias = searchColumnAlias
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
