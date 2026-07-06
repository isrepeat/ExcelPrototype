VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_EntityLookupLayoutState"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Private Const DEFAULT_CANDIDATES_AT As String = "r1c1"
Private Const RUNTIME_ERROR_TITLE As String = "PrototypeNew / EntityLookup layout"

Private m_ActiveCandidatesAt As String
Private m_CandidateTableSpanCols As Long
Private m_IsDisposed As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & TypeName(Me) & ".Class_Initialize"
#End If
    m_ActiveCandidatesAt = DEFAULT_CANDIDATES_AT
    m_CandidateTableSpanCols = 50
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Me.Dispose
    On Error GoTo 0
End Sub

' //
' // Properties
' //
Public Property Get ActiveCandidatesAt() As String
    ActiveCandidatesAt = m_ActiveCandidatesAt
End Property

' //
' // API
' //
Public Function Initialize(Optional ByVal candidateTableSpanCols As Long = 50) As Boolean
    m_IsDisposed = False
    If candidateTableSpanCols <= 0 Then candidateTableSpanCols = 50
    m_CandidateTableSpanCols = candidateTableSpanCols
    m_ActiveCandidatesAt = DEFAULT_CANDIDATES_AT
    Initialize = True
End Function

Public Sub Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    m_ActiveCandidatesAt = DEFAULT_CANDIDATES_AT
    m_CandidateTableSpanCols = 50
End Sub

Public Function ResetCandidateLayoutFromConfig() As Boolean
    m_ActiveCandidatesAt = DEFAULT_CANDIDATES_AT
    ResetCandidateLayoutFromConfig = True
End Function

Public Function TrySetActiveCandidateLayout( _
    ByVal cfgParser As obj_EntityLookupCfgParser, _
    ByVal lookupKey As String, _
    ByVal tableObj As obj_TableDynamic, _
    ByVal searchColumnAlias As String, _
    Optional ByVal inputGridColumnOverride As Long = 0 _
) As Boolean
    Dim inputGridCol As Long
    Dim searchColumnIndex As Long
    Dim startGridCol As Long

    If cfgParser Is Nothing Then Exit Function
    If tableObj Is Nothing Then Exit Function

    If tableObj.ColumnCount > m_CandidateTableSpanCols Then
        private_ShowLayoutError "Candidate table for lookup '" & lookupKey & "' has " & VBA.CStr(tableObj.ColumnCount) & _
            " columns, but current UI slot supports only " & VBA.CStr(m_CandidateTableSpanCols) & _
            ". Reduce ResultColumnsAliases or increase TableList spanColls."
        Exit Function
    End If

    ' inputGridColumnOverride приходит от LookupCandidatesControlVM, когда UI
    ' смог найти реальный visible input по layout tag. Если override не задан,
    ' сохраняем прежнее поведение: колонка input-а считается по порядку
    ' EntityLookup.Table.Columns в конфиге.
    If inputGridColumnOverride > 0 Then
        inputGridCol = inputGridColumnOverride
    ElseIf Not cfgParser.TryGetLookupInputGridColumn(lookupKey, inputGridCol) Then
        Exit Function
    End If
    If Not private_TryGetCandidateSearchColumnIndex(tableObj, searchColumnAlias, searchColumnIndex) Then
        private_ShowLayoutError "Failed to find search column alias '" & searchColumnAlias & "' in candidate columns for lookup '" & lookupKey & "'."
        Exit Function
    End If

    startGridCol = inputGridCol - searchColumnIndex + 1
    If startGridCol <= 0 Then
        private_ShowLayoutError "Candidate table for lookup '" & lookupKey & "' cannot be aligned: search column index is " & _
            VBA.CStr(searchColumnIndex) & ", but input grid column is " & VBA.CStr(inputGridCol) & _
            ". Move the candidate area left or place SearchColumnAlias earlier in ResultColumnsAliases."
        Exit Function
    End If

    m_ActiveCandidatesAt = private_BuildCandidateAtText(startGridCol)
    TrySetActiveCandidateLayout = True
End Function

' //
' // Internal
' //
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

Private Sub private_ShowLayoutError(ByVal messageText As String)
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "EntityLookup layout: " & VBA.CStr(messageText)
#End If
    VBA.MsgBox "PrototypeNew: " & VBA.CStr(messageText), vbExclamation, RUNTIME_ERROR_TITLE
End Sub
