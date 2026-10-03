VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_LookupService"
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_profile As obj_LookupProfile
Private m_cache As obj_DataTable

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize(ByVal profile As obj_LookupProfile) As Boolean
    If m_isInitialized Or m_isDisposed Then Exit Function
    If profile Is Nothing Then Exit Function
    If profile.Columns Is Nothing Then Exit Function
    Set m_profile = profile
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
    Me.InvalidateCache
    Set m_profile = Nothing
End Sub

Public Sub InvalidateCache()
    If Not m_cache Is Nothing Then m_cache.Dispose
    Set m_cache = Nothing
End Sub

Public Function TrySearch( _
    ByVal text As String, _
    ByRef result As obj_LookupResult, _
    ByRef diagnostic As String _
) As Boolean
    Dim query As obj_TableQuery
    Dim group As obj_QueryGroup
    Dim condition As obj_QueryCondition
    Dim search As Object
    Dim data As obj_DataTable
    Dim records As Collection
    Dim hasMore As Boolean
    Dim startedAt As Double
    Dim stageStartedAt As Double

    On Error GoTo EH
    startedAt = VBA.Timer
    ex_Core.fn_Diagnostic_WriteLog "LOOKUP_SEARCH_STARTED | InputLength=" & VBA.CStr(VBA.Len(text))
    Set result = Nothing
    diagnostic = VBA.vbNullString
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Lookup service is not initialized."
    text = ex_TableQuery.fn_NormalizeText(text)
    Set records = New Collection
    If VBA.Len(text) >= m_profile.MinChars Then
        stageStartedAt = VBA.Timer
        If Not private_LoadCache(diagnostic) Then GoTo Cleanup
        ex_Core.fn_Diagnostic_WritePerf "Lookup.LoadCache", stageStartedAt
        Set query = New obj_TableQuery
        If Not query.Initialize() Then Err.Raise 5, , "Cannot initialize lookup query."
        Set group = New obj_QueryGroup
        If Not group.Initialize() Then Err.Raise 5, , "Cannot initialize lookup filter."
        group.MatchAny = True
        For Each search In m_profile.SearchFields
            Set condition = New obj_QueryCondition
            If Not condition.Initialize() Then Err.Raise 5, , "Cannot initialize lookup condition."
            condition.ColumnName = m_profile.HeaderFor(search("name"))
            condition.Operation = QueryContains
            If search("match") = "startsWith" Then condition.Operation = QueryStartsWith
            condition.Value = text
            condition.ValueType = QueryText
            condition.NormalizeText = True
            group.AddCondition condition
        Next search
        query.Filter.AddGroup group
        query.AddCondition m_profile.HeaderFor(m_profile.KeyField), QueryIsNotEmpty, Empty, QueryText, True
        Set search = m_profile.SearchFields(1)
        query.AddOrderBy m_profile.HeaderFor(search("name")), False, QueryText, True
        query.AddOrderBy m_profile.HeaderFor(m_profile.KeyField), False, QueryText, True
        query.Limit = m_profile.MaxResults + 1
        stageStartedAt = VBA.Timer
        If Not ex_TableQuery.fn_TryApply(m_cache, query, data, diagnostic) Then GoTo Cleanup
        ex_Core.fn_Diagnostic_WritePerf "Lookup.ApplySearch", stageStartedAt
        hasMore = (data.RowCount > m_profile.MaxResults)
        stageStartedAt = VBA.Timer
        If Not private_ReadRecords(data, records, diagnostic) Then GoTo Cleanup
        ex_Core.fn_Diagnostic_WritePerf "Lookup.BuildRecords | Rows=" & VBA.CStr(records.Count), stageStartedAt
    End If
    stageStartedAt = VBA.Timer
    Set result = New obj_LookupResult
    TrySearch = result.Initialize(records, m_profile.DisplayColumns, hasMore, diagnostic)
    ex_Core.fn_Diagnostic_WritePerf "Lookup.BuildResult", stageStartedAt
Cleanup:
    If Not query Is Nothing Then query.Dispose
    If Not data Is Nothing Then data.Dispose
    ex_Core.fn_Diagnostic_WritePerf "Lookup.Search | Success=" & VBA.CStr(TrySearch), startedAt
    Exit Function
EH:
    diagnostic = "Lookup search: " & VBA.Err.Description
    ex_Core.fn_Diagnostic_WriteLog "LOOKUP_SEARCH_FAILED | Number=" & VBA.CStr(VBA.Err.Number)
    Set result = Nothing
    Resume Cleanup
End Function

Public Function TryProject( _
    ByVal records As Collection, _
    ByRef result As obj_LookupResult, _
    ByRef diagnostic As String _
) As Boolean
    If Not m_isInitialized Or m_isDisposed Then
        diagnostic = "Lookup service is not initialized."
        Exit Function
    End If
    Set result = New obj_LookupResult
    TryProject = result.Initialize(records, m_profile.DisplayColumns, False, diagnostic)
End Function

Public Function TryExtend( _
    ByVal candidates As Collection, _
    ByVal extension As Collection, _
    ByVal joinField As String, _
    ByVal duplicatePolicy As String, _
    ByVal overwrite As Boolean, _
    ByRef output As Collection, _
    ByRef diagnostic As String, _
    Optional ByVal selector As obj_ILookupRowSelector _
) As Boolean
    Dim index As Object
    Dim record As obj_LookupRecord
    Dim copy As obj_LookupRecord
    Dim matches As Collection
    Dim key As String
    Dim selected As Long

    On Error GoTo EH
    Set output = Nothing
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Lookup service is not initialized."
    If candidates Is Nothing Or extension Is Nothing Then Err.Raise 5, , "Candidate collections are required."
    If duplicatePolicy <> "error" And duplicatePolicy <> "first" And duplicatePolicy <> "last" And duplicatePolicy <> "custom" Then Err.Raise 5, , "Unknown duplicate policy: " & duplicatePolicy
    If duplicatePolicy = "custom" And selector Is Nothing Then Err.Raise 5, , "Custom extension requires a row selector."
    If Not private_Index(extension, joinField, index, diagnostic) Then Exit Function
    Set output = New Collection
    For Each record In candidates
        Set copy = private_Clone(record)
        key = private_JoinKey(record, joinField)
        If VBA.Len(key) > 0 And index.Exists(key) Then
            Set matches = index(key)
            If matches.Count > 1 And duplicatePolicy = "error" Then Err.Raise 5, , "Ambiguous candidate extension for key: " & key
            selected = 1
            If duplicatePolicy = "last" Then selected = matches.Count
            If duplicatePolicy = "custom" Then
                selected = 0
                If Not selector.TrySelect(record, matches, selected, diagnostic) Then
                    Set output = Nothing
                    Exit Function
                End If
            End If
            If selected < 0 Or selected > matches.Count Then Err.Raise 5, , "Extension selector returned an invalid row index."
            If selected > 0 Then
                If Not copy.TryMerge(matches(selected), overwrite) Then Err.Raise 5, , "Cannot merge candidate extension."
            End If
        End If
        output.Add copy
    Next record
    TryExtend = True
    Exit Function
EH:
    Set output = Nothing
    diagnostic = "Extend candidates: " & VBA.Err.Description
End Function

Public Function TryAppend( _
    ByVal candidates As Collection, _
    ByVal incoming As Collection, _
    ByVal duplicatePolicy As String, _
    ByRef output As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim byKey As Object
    Dim record As obj_LookupRecord
    Dim existing As obj_LookupRecord
    Dim all As New Collection
    Dim order As New Collection
    Dim key As Variant

    On Error GoTo EH
    Set output = Nothing
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Lookup service is not initialized."
    If candidates Is Nothing Or incoming Is Nothing Then Err.Raise 5, , "Candidate collections are required."
    If duplicatePolicy <> "error" And duplicatePolicy <> "keep" And duplicatePolicy <> "replace" Then Err.Raise 5, , "Unknown append policy: " & duplicatePolicy
    Set byKey = VBA.CreateObject("Scripting.Dictionary")
    byKey.CompareMode = VBA.vbTextCompare
    For Each record In candidates
        all.Add record
    Next record
    For Each record In incoming
        all.Add record
    Next record
    For Each record In all
        key = ex_TableQuery.fn_NormalizeText(record.Key)
        If VBA.Len(key) = 0 Then Err.Raise 5, , "Cannot append a candidate without a key."
        If byKey.Exists(key) Then
            If duplicatePolicy = "error" Then Err.Raise 5, , "Duplicate appended candidate key: " & key
            If duplicatePolicy = "replace" Then Set byKey(key) = private_Clone(record)
        Else
            Set byKey(key) = private_Clone(record)
            order.Add key
        End If
    Next record
    Set output = New Collection
    For Each key In order
        Set existing = byKey(key)
        output.Add existing
    Next key
    TryAppend = True
    Exit Function
EH:
    Set output = Nothing
    diagnostic = "Append candidates: " & VBA.Err.Description
End Function

Public Function TryAttachDetails( _
    ByVal candidates As Collection, _
    ByVal details As Collection, _
    ByVal joinField As String, _
    ByVal detailName As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim index As Object
    Dim record As obj_LookupRecord
    Dim groups As New Collection
    Dim matches As Collection
    Dim key As String
    Dim i As Long

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Lookup service is not initialized."
    If candidates Is Nothing Or details Is Nothing Then Err.Raise 5, , "Candidate collections are required."
    If VBA.Len(detailName) = 0 Then Err.Raise 5, , "Detail name is required."
    If Not private_Index(details, joinField, index, diagnostic) Then Exit Function
    For Each record In candidates
        key = private_JoinKey(record, joinField)
        Set matches = New Collection
        If VBA.Len(key) > 0 And index.Exists(key) Then Set matches = index(key)
        groups.Add matches
    Next record
    For i = 1 To candidates.Count
        Set record = candidates(i)
        If Not record.SetDetail(detailName, groups(i)) Then Err.Raise 5, , "Cannot attach candidate details."
    Next i
    TryAttachDetails = True
    Exit Function
EH:
    diagnostic = "Attach details: " & VBA.Err.Description
End Function

' //
' // Private
' //
Private Function private_LoadCache(ByRef diagnostic As String) As Boolean
    Dim service As obj_TableQueryService
    Dim query As obj_TableQuery
    Dim column As Object
    Dim keyIndex As Long
    Dim rowIndex As Long
    Dim keys As Object
    Dim key As String
    Dim startedAt As Double

    On Error GoTo EH
    If Not m_cache Is Nothing Then
        ex_Core.fn_Diagnostic_WriteLog "LOOKUP_CACHE_HIT | Rows=" & VBA.CStr(m_cache.RowCount)
        private_LoadCache = True
        Exit Function
    End If
    ex_Core.fn_Diagnostic_WriteLog "LOOKUP_CACHE_LOAD_STARTED"
    Set service = New obj_TableQueryService
    Set query = New obj_TableQuery
    If Not service.Initialize() Or Not query.Initialize() Then Err.Raise 5, , "Cannot initialize source query."
    service.Backend = QueryExcel
    For Each column In m_profile.Columns
        query.AddColumn column("header")
    Next column
    If Not service.TryExecute(m_profile.Source, query, m_cache, diagnostic) Then GoTo Cleanup
    startedAt = VBA.Timer
    keyIndex = ex_TableQuery.fn_HeaderIndex(m_cache.Headers, m_profile.HeaderFor(m_profile.KeyField))
    Set keys = VBA.CreateObject("Scripting.Dictionary")
    keys.CompareMode = VBA.vbTextCompare
    For rowIndex = 1 To m_cache.RowCount
        key = private_Text(m_cache.ValueAt(rowIndex, keyIndex))
        If VBA.Len(key) > 0 Then
            If keys.Exists(key) Then Err.Raise 5, , "Duplicate lookup entity key in source: " & key
            keys.Add key, True
        End If
    Next rowIndex
    ex_Core.fn_Diagnostic_WritePerf "Lookup.ValidateKeys | Rows=" & VBA.CStr(m_cache.RowCount), startedAt
    private_LoadCache = True
Cleanup:
    If Not service Is Nothing Then service.Dispose
    If Not query Is Nothing Then query.Dispose
    If Not private_LoadCache Then Me.InvalidateCache
    Exit Function
EH:
    diagnostic = "Load lookup source: " & VBA.Err.Description
    Resume Cleanup
End Function

Private Function private_ReadRecords( _
    ByVal data As obj_DataTable, _
    ByRef records As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim fields As Object
    Dim column As Object
    Dim record As obj_LookupRecord
    Dim value As Variant
    Dim rowIndex As Long
    Dim count As Long

    On Error GoTo EH
    count = data.RowCount
    If count > m_profile.MaxResults Then count = m_profile.MaxResults
    Set records = New Collection
    For rowIndex = 1 To count
        Set fields = VBA.CreateObject("Scripting.Dictionary")
        fields.CompareMode = VBA.vbTextCompare
        For Each column In m_profile.Columns
            value = data.ValueAt(rowIndex, ex_TableQuery.fn_HeaderIndex(data.Headers, column("header")))
            If VBA.IsError(value) Then Err.Raise 5, , "Excel error in candidate field: " & column("header")
            If VBA.IsNull(value) Or VBA.IsEmpty(value) Then value = VBA.vbNullString
            fields(column("name")) = value
        Next column
        Set record = New obj_LookupRecord
        If Not record.Initialize(private_Text(fields(m_profile.KeyField)), fields) Then Err.Raise 5, , "Cannot initialize candidate record."
        records.Add record
    Next rowIndex
    private_ReadRecords = True
    Exit Function
EH:
    diagnostic = "Read candidates: " & VBA.Err.Description
End Function

Private Function private_Index( _
    ByVal records As Collection, _
    ByVal joinField As String, _
    ByRef index As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim record As obj_LookupRecord
    Dim matches As Collection
    Dim key As String

    On Error GoTo EH
    Set index = VBA.CreateObject("Scripting.Dictionary")
    index.CompareMode = VBA.vbTextCompare
    For Each record In records
        key = private_JoinKey(record, joinField)
        If VBA.Len(key) > 0 Then
            If Not index.Exists(key) Then
                Set matches = New Collection
                Set index(key) = matches
            End If
            Set matches = index(key)
            matches.Add record
        End If
    Next record
    private_Index = True
    Exit Function
EH:
    diagnostic = "Candidate index: " & VBA.Err.Description
End Function

Private Function private_JoinKey( _
    ByVal record As obj_LookupRecord, _
    ByVal field As String _
) As String
    Dim value As Variant

    If Not record.TryGetValue(field, value) Then Err.Raise 5, , "Join field not found: " & field
    private_JoinKey = private_Text(value)
End Function

Private Function private_Clone(ByVal record As obj_LookupRecord) As obj_LookupRecord
    Dim copy As New obj_LookupRecord
    Dim name As Variant

    If Not copy.Initialize(record.Key, record.Fields) Then Err.Raise 5, , "Cannot clone candidate."
    For Each name In record.DetailNames
        If Not copy.SetDetail(VBA.CStr(name), record.GetDetail(VBA.CStr(name))) Then Err.Raise 5, , "Cannot copy candidate details."
    Next name
    Set private_Clone = copy
End Function

Private Function private_Text(ByVal value As Variant) As String
    If VBA.IsError(value) Then Err.Raise 5, , "Excel error in candidate key."
    If VBA.IsNull(value) Or VBA.IsEmpty(value) Then Exit Function
    private_Text = ex_TableQuery.fn_NormalizeText(VBA.CStr(value))
End Function