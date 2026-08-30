VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PEB_MedicalReportsScen"
Option Explicit

Private Const ERROR_TITLE As String = "PrsnlEventBuilder / Стройові записки"
Private Const ADO_TEXT_LIMIT As Long = 255
Private Const ADO_LONG_VALUE_CANDIDATE_TAG As String = "ado-long-value-candidate"
Private Const FILTER_EXPRESSION_EMPTY As String = "empty()"
Private Const FILTER_EXPRESSION_NOT_EMPTY As String = "notempty()"

Private m_IsDisposed As Boolean
Private m_CfgParser As obj_MultiSourcesViewCfgParser
Private m_TableRefs As Collection
Private m_Columns As Collection
Private m_MaxRows As Long

Public Function Initialize(ByVal configTable As obj_ConfigTable) As Boolean
    If configTable Is Nothing Then Exit Function
    m_IsDisposed = False
    Set m_CfgParser = New obj_MultiSourcesViewCfgParser
    If Not m_CfgParser.Initialize(configTable) Then Exit Function
    If Not m_CfgParser.TryGetViewSettings( _
        m_TableRefs, m_Columns, m_MaxRows) Then Exit Function
    Initialize = True
End Function

Public Function TryGetFilterItems(ByRef outItems As Collection) As Boolean
    Dim columnAliasObj As Variant
    Dim entry As obj_ConfigEntry

    Set outItems = Nothing
    If m_IsDisposed Or m_Columns Is Nothing Then Exit Function
    Set outItems = New Collection
    For Each columnAliasObj In m_Columns
        Set entry = New obj_ConfigEntry
        entry.Key = VBA.Trim$(VBA.CStr(columnAliasObj))
        entry.Value = VBA.vbNullString
        outItems.Add entry
    Next columnAliasObj
    TryGetFilterItems = True
End Function

Public Function TryGetFilterColumns(ByRef outColumns As Collection) As Boolean
    Dim columnAliasObj As Variant

    Set outColumns = Nothing
    If m_IsDisposed Or m_Columns Is Nothing Then Exit Function
    Set outColumns = New Collection
    For Each columnAliasObj In m_Columns
        outColumns.Add VBA.CStr(columnAliasObj)
    Next columnAliasObj
    TryGetFilterColumns = True
End Function

Public Function TryLoadTables( _
    ByVal filterValues As Object, _
    ByRef outTables As Collection _
) As Boolean
    Dim tableRef As Variant
    Dim sqlParamsList As Collection
    Dim sqlParamsItem As Variant
    Dim sqlParams As obj_SqlParams
    Dim tableData As obj_TableData
    Dim tableObj As obj_TableDynamic

    Set outTables = Nothing
    If m_IsDisposed Or m_CfgParser Is Nothing Then
        VBA.MsgBox "Підрежим стройових записок не ініціалізований.", _
            VBA.vbExclamation, ERROR_TITLE
        Exit Function
    End If

    Set outTables = New Collection
    For Each tableRef In m_TableRefs
        If Not m_CfgParser.TryBuildTableSqlParamsList( _
            VBA.CStr(tableRef), m_Columns, sqlParamsList) Then Exit Function
        For Each sqlParamsItem In sqlParamsList
            Set sqlParams = sqlParamsItem
            sqlParams.MaxRows = m_MaxRows
            If Not ex_ExternalExcelSqlEngine.fn_TrySqlRequestData( _
                sqlParams, tableData) Then Exit Function
            If Not private_TryBuildTable( _
                private_GetFileName(sqlParams.SourcePath), _
                sqlParams.SourceAlias, sqlParams.SourceAliasTemplate, _
                tableData, filterValues, tableObj) Then Exit Function
            ' Пустые после фильтрации источники не занимают место на листе.
            If tableObj.RowCount > 0 Then outTables.Add tableObj
        Next sqlParamsItem
    Next tableRef

    TryLoadTables = True
End Function

Public Sub Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    If Not m_CfgParser Is Nothing Then m_CfgParser.Dispose
    Set m_CfgParser = Nothing
    Set m_TableRefs = Nothing
    Set m_Columns = Nothing
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    Me.Dispose
    On Error GoTo 0
End Sub

Private Function private_TryBuildTable( _
    ByVal sectionTitle As String, _
    ByVal sourceAlias As String, _
    ByVal sourceAliasTemplate As String, _
    ByVal tableData As obj_TableData, _
    ByVal filterValues As Object, _
    ByRef outTable As obj_TableDynamic _
) As Boolean
    Dim tableObj As obj_TableDynamic
    Dim columnObj As obj_Column
    Dim visibleRowIndexes As Collection
    Dim columnIndex As Long
    Dim rowIndex As Long
    Dim aliasText As String
    Dim sectionDate As String

    Set outTable = Nothing
    If tableData Is Nothing Or m_Columns Is Nothing Then Exit Function
    Set tableObj = New obj_TableDynamic
    Set visibleRowIndexes = New Collection
    If Not private_TryResolveSectionDate( _
        sectionTitle, sourceAlias, sectionDate) Then Exit Function
    tableObj.SectionTitle = sectionDate
    tableObj.SourceAlias = sourceAlias
    tableObj.SourceAliasTemplate = sourceAliasTemplate

    For columnIndex = 1 To m_Columns.Count
        aliasText = VBA.Trim$(VBA.CStr(m_Columns.Item(columnIndex)))
        Set columnObj = New obj_Column
        columnObj.Name = aliasText
        columnObj.Position = columnIndex
        If Not columnObj.AddAlias(aliasText) Then Exit Function
        If Not tableObj.PushColumn(columnObj) Then Exit Function
    Next columnIndex

    For rowIndex = 1 To tableData.RowCount
        If Not private_RowMatches( _
            tableData, rowIndex, filterValues) Then GoTo ContinueRow
        visibleRowIndexes.Add rowIndex
ContinueRow:
    Next rowIndex

    If Not tableObj.SetCompactData(tableData, visibleRowIndexes) Then Exit Function
    If Not tableObj.AddCompactTextLengthTag( _
        ADO_TEXT_LIMIT, ADO_LONG_VALUE_CANDIDATE_TAG) Then Exit Function

    Set outTable = tableObj
    private_TryBuildTable = True
End Function

Private Function private_TryResolveSectionDate( _
    ByVal sourceFileName As String, _
    ByVal sourceAlias As String, _
    ByRef outDateText As String _
) As Boolean
    Dim regExp As Object
    Dim matches As Object
    Dim searchText As String

    outDateText = VBA.vbNullString
    searchText = VBA.Trim$(sourceAlias) & " " & VBA.Trim$(sourceFileName)
    Set regExp = VBA.CreateObject("VBScript.RegExp")
    regExp.Global = True
    regExp.IgnoreCase = True
    regExp.Pattern = "([0-3][0-9]\.[01][0-9]\.[12][0-9][0-9][0-9])"
    Set matches = regExp.Execute(searchText)
    If matches.Count = 0 Then
        VBA.MsgBox _
            "Не вдалося визначити дату стройової записки з alias або назви файлу: '" & _
            sourceAlias & "' / '" & sourceFileName & "'.", _
            VBA.vbExclamation, ERROR_TITLE
        Exit Function
    End If

    outDateText = VBA.CStr(matches.Item(matches.Count - 1).SubMatches(0))
    private_TryResolveSectionDate = True
End Function

Private Function private_RowMatches( _
    ByVal tableData As obj_TableData, _
    ByVal rowIndex As Long, _
    ByVal filterValues As Object _
) As Boolean
    Dim columnIndex As Long
    Dim columnAlias As String
    Dim filterText As String

    If filterValues Is Nothing Then
        private_RowMatches = True
        Exit Function
    End If
    For columnIndex = 1 To m_Columns.Count
        columnAlias = VBA.Trim$(VBA.CStr(m_Columns.Item(columnIndex)))
        filterText = VBA.vbNullString
        If filterValues.Exists(columnAlias) Then _
            filterText = VBA.Trim$(VBA.CStr(filterValues(columnAlias)))
        If VBA.Len(filterText) > 0 Then
            If Not private_FilterMatches( _
                tableData.ValueAt(rowIndex, columnIndex), _
                filterText) Then Exit Function
        End If
    Next columnIndex
    private_RowMatches = True
End Function

Private Function private_FilterMatches( _
    ByVal cellText As String, _
    ByVal filterText As String _
) As Boolean
    Dim likeExpression As String

    filterText = VBA.Trim$(filterText)
    If VBA.StrComp(filterText, FILTER_EXPRESSION_EMPTY, _
        VBA.vbTextCompare) = 0 Then
        private_FilterMatches = (VBA.Len(VBA.Trim$(cellText)) = 0)
        Exit Function
    End If
    If VBA.StrComp(filterText, FILTER_EXPRESSION_NOT_EMPTY, _
        VBA.vbTextCompare) = 0 Then
        private_FilterMatches = (VBA.Len(VBA.Trim$(cellText)) > 0)
        Exit Function
    End If
    If private_TryExtractLikeExpression(filterText, likeExpression) Then
        private_FilterMatches = private_MatchesLikeExpression( _
            cellText, likeExpression)
        Exit Function
    End If
    private_FilterMatches = (VBA.InStr( _
        1, cellText, filterText, VBA.vbTextCompare) > 0)
End Function

Private Function private_TryExtractLikeExpression( _
    ByVal filterText As String, _
    ByRef outExpression As String _
) As Boolean
    outExpression = VBA.vbNullString
    filterText = VBA.Trim$(filterText)
    If VBA.Len(filterText) < 4 Then Exit Function
    If VBA.StrComp(VBA.Left$(filterText, 3), "rx(", _
        VBA.vbTextCompare) <> 0 Then Exit Function
    If VBA.Right$(filterText, 1) <> ")" Then Exit Function
    outExpression = VBA.Mid$(filterText, 4, VBA.Len(filterText) - 4)
    private_TryExtractLikeExpression = True
End Function

Private Function private_MatchesLikeExpression( _
    ByVal cellText As String, _
    ByVal likeExpression As String _
) As Boolean
    On Error GoTo InvalidPattern
    private_MatchesLikeExpression = ( _
        VBA.LCase$(cellText) Like _
        VBA.LCase$(private_NormalizeLikeExpression(likeExpression)))
    Exit Function
InvalidPattern:
    VBA.MsgBox "Некоректний вираз фільтра rx(" & likeExpression & ").", _
        VBA.vbExclamation, ERROR_TITLE
    Err.Clear
End Function

Private Function private_NormalizeLikeExpression( _
    ByVal expressionText As String _
) As String
    Dim resultText As String
    Dim currentChar As String
    Dim charIndex As Long

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

Private Function private_GetFileName(ByVal filePath As String) As String
    Dim slashPosition As Long

    filePath = VBA.Replace$(VBA.Trim$(filePath), "/", "\")
    slashPosition = VBA.InStrRev(filePath, "\")
    If slashPosition > 0 Then
        private_GetFileName = VBA.Mid$(filePath, slashPosition + 1)
    Else
        private_GetFileName = filePath
    End If
End Function
