VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PEB_DependencyLoader"
Option Explicit

Implements obj_ITableSnapshotLoader

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

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
Public Function Initialize() As Boolean
    If m_isInitialized Or m_isDisposed Then
        Exit Function
    End If
    m_isInitialized = True
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
End Sub

Public Function TryBuild( _
    ByVal snapshot As obj_TableSnapshotSet, _
    ByRef diagnostic As String _
) As Boolean
    Dim prefix As Variant
    Dim mapping As Variant
    Dim pair As Variant
    Dim sources As Object
    Dim source As obj_TableSource
    Dim templates As obj_PEB_WordTemplates
    Dim startedAt As Double
    Dim busyEntered As Boolean
    Dim candidates As obj_PEB_CandidateEnricher
    Dim tickets As obj_PEB_ReturnTickets
    Dim profile As obj_LookupProfile
    Dim lookup As obj_LookupService
    Dim orders As obj_OrderResolver
    Dim snapshotStarted As Boolean
    Dim signatures As Object
    Dim path As Variant
    Dim sourceDefinitions As Variant
    Dim indexDefinitions As Variant
    Dim profileDefinitions As Variant
    Dim progressStep As Long
    Dim progressTotal As Long

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then
        Err.Raise 5, , "Dependency loader is unavailable."
    End If
    busyEntered = ex_UiBusy.fn_Begin(ex_PEB_Export.fn_Config("text.DependenciesLoading"))
    ex_PEB_Export.fn_BeginConfigSnapshot
    snapshotStarted = True
    sourceDefinitions = VBA.Split(ex_PEB_Export.fn_Config("snapshots.sources"), "|")
    indexDefinitions = VBA.Split(ex_PEB_Export.fn_Config("snapshots.indexes"), "|")
    profileDefinitions = VBA.Split(ex_PEB_Export.fn_Config("snapshots.profiles"), "|")
    progressTotal = (UBound(sourceDefinitions) + 1) + (UBound(indexDefinitions) + 1) + _
        (UBound(profileDefinitions) + 1) + 7
    Set sources = VBA.CreateObject("Scripting.Dictionary")
    Set signatures = VBA.CreateObject("Scripting.Dictionary")
    signatures.CompareMode = VBA.vbTextCompare
    For Each prefix In sourceDefinitions
        startedAt = VBA.Timer
        Set source = New obj_TableSource
        If Not source.Initialize() Then
            Err.Raise 5, , "Cannot initialize dependency source."
        End If
        source.WorkbookPath = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config(prefix & ".path"))
        source.SheetName = ex_PEB_Export.fn_Config(prefix & ".sheet")
        source.RangeAddress = ex_PEB_Export.fn_Config(prefix & ".range")
        private_Progress progressStep, progressTotal, "Table", source.SheetName
        If Not signatures.Exists(source.WorkbookPath) Then
            signatures.Add source.WorkbookPath, private_FileStamp(source.WorkbookPath)
        End If
        If Not snapshot.TryLoad(source, diagnostic) Then
            GoTo Cleanup
        End If
        sources.Add prefix, source
        ex_Core.fn_Diagnostic_WritePerf "Dependencies.Table | Source=" & prefix, startedAt
    Next prefix
    For Each mapping In indexDefinitions
        pair = VBA.Split(mapping, "=")
        Set source = sources(pair(0))
        private_Progress progressStep, progressTotal, "Index", ex_PEB_Export.fn_Config(pair(1))
        snapshot.WarmIndex source, ex_PEB_Export.fn_Config(pair(1))
    Next mapping
    For Each prefix In profileDefinitions
        Set profile = New obj_LookupProfile
        Set lookup = New obj_LookupService
        If Not profile.Initialize("PersonalEventBuilder::" & prefix, diagnostic) Then
            GoTo Cleanup
        End If
        If Not lookup.Initialize(profile) Then
            Err.Raise 5, , "Cannot initialize dependency lookup."
        End If
        private_Progress progressStep, progressTotal, "Profile", profile.Source.SheetName
        If Not lookup.PrepareSnapshot(snapshot, diagnostic) Then
            GoTo Cleanup
        End If
        lookup.Dispose
        profile.Dispose
        Set lookup = Nothing
        Set profile = Nothing
    Next prefix
    private_Progress progressStep, progressTotal, "Orders"
    Set orders = New obj_OrderResolver
    If Not orders.Initialize("PersonalEventBuilder::orders", "PersonalEventBuilder::text", diagnostic) Then
        GoTo Cleanup
    End If
    If Not orders.PrepareSnapshot(snapshot, diagnostic) Then
        GoTo Cleanup
    End If
    private_Progress progressStep, progressTotal, "Tvo"
    If Not private_PrepareTvo(snapshot, sources("lookup.Personnel.source"), diagnostic) Then
        GoTo Cleanup
    End If
    private_Progress progressStep, progressTotal, "Absence"
    Set candidates = New obj_PEB_CandidateEnricher
    If Not candidates.Initialize(diagnostic) Then
        GoTo Cleanup
    End If
    If Not candidates.PrepareSnapshot(snapshot, diagnostic) Then
        GoTo Cleanup
    End If
    private_Progress progressStep, progressTotal, "Tickets"
    Set tickets = New obj_PEB_ReturnTickets
    If Not tickets.Initialize() Then
        Err.Raise 5, , "Cannot initialize ticket preparation."
    End If
    If Not tickets.PrepareSnapshot(snapshot, diagnostic) Then
        GoTo Cleanup
    End If
    snapshot.Prepared.Add "ReturnTickets", tickets
    Set templates = New obj_PEB_WordTemplates
    For Each prefix In VBA.Array("export.Word.templates.path", "export.Movement.templates.path")
        path = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config(VBA.CStr(prefix)))
        If Not signatures.Exists(path) Then
            signatures.Add path, private_FileStamp(VBA.CStr(path))
        End If
    Next prefix
    private_Progress progressStep, progressTotal, "Word"
    If Not templates.Initialize(diagnostic) Then
        GoTo Cleanup
    End If
    snapshot.Prepared.Add "WordTemplates", templates
    private_Progress progressStep, progressTotal, "Movement"
    Set snapshot.Prepared("MovementTemplates") = ex_PEB_MovementTemplates.fn_BuildSnapshot()
    private_Progress progressStep, progressTotal, "Verify"
    For Each path In signatures.Keys
        If signatures(path) <> private_FileStamp(VBA.CStr(path)) Then
            Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.DependenciesSourceChanged"), "{path}", VBA.CStr(path))
        End If
    Next path
    TryBuild = True
Cleanup:
    If Not lookup Is Nothing Then
        lookup.Dispose
    End If
    If Not profile Is Nothing Then
        profile.Dispose
    End If
    If Not orders Is Nothing Then
        orders.Dispose
    End If
    If Not candidates Is Nothing Then
        candidates.Dispose
    End If
    If snapshotStarted Then
        ex_PEB_Export.fn_EndConfigSnapshot
    End If
    If busyEntered Then
        ex_UiBusy.fn_End
    End If
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function

' //
' // Interface
' //
Private Function obj_ITableSnapshotLoader_TryBuild( _
    ByVal snapshot As obj_TableSnapshotSet, _
    ByRef diagnostic As String _
) As Boolean
    obj_ITableSnapshotLoader_TryBuild = Me.TryBuild(snapshot, diagnostic)
End Function

' //
' // Private
' //
Private Function private_PrepareTvo( _
    ByVal snapshot As obj_TableSnapshotSet, _
    ByVal source As obj_TableSource, _
    ByRef diagnostic As String _
) As Boolean
    Dim query As obj_TableQuery
    Dim table As obj_DataTable
    Dim names As Object
    Dim reporters As Object
    Dim chains As Object
    Dim errors As Object
    Dim chainDiagnostic As String
    Dim matches As Collection
    Dim chain As Collection
    Dim header As Variant
    Dim key As Variant
    Dim name As String
    Dim actor As String
    Dim row As Long
    Dim nameColumn As Long
    Dim actorColumn As Long
    Dim ipnColumn As Long
    Dim codeColumn As Long

    On Error GoTo EH
    Set query = New obj_TableQuery
    If Not query.Initialize() Then
        Err.Raise 5, , "Cannot initialize TVO preparation."
    End If
    For Each header In VBA.Array("lookup.Personnel.columns.1.header", "lookup.Personnel.columns.2.header", _
            "lookup.Personnel.columns.6.header", "lookup.Personnel.columns.7.header")
        query.AddColumn ex_PEB_Export.fn_Config(VBA.CStr(header))
    Next header
    If Not snapshot.TryQuery(source, query, table, diagnostic) Then
        GoTo Cleanup
    End If
    nameColumn = 1
    ipnColumn = 2
    codeColumn = 3
    actorColumn = 4
    Set names = VBA.CreateObject("Scripting.Dictionary")
    names.CompareMode = VBA.vbTextCompare
    Set reporters = VBA.CreateObject("Scripting.Dictionary")
    reporters.CompareMode = VBA.vbTextCompare
    Set chains = VBA.CreateObject("Scripting.Dictionary")
    chains.CompareMode = VBA.vbTextCompare
    Set errors = VBA.CreateObject("Scripting.Dictionary")
    errors.CompareMode = VBA.vbTextCompare
    For row = 1 To table.RowCount
        name = private_Text(table.ValueAt(row, nameColumn))
        If VBA.Len(name) > 0 Then
            If Not names.Exists(name) Then
                Set matches = New Collection
                names.Add name, matches
            End If
            Set matches = names(name)
            matches.Add row
        End If
        actor = private_Text(table.ValueAt(row, actorColumn))
        If VBA.Len(actor) > 0 Then
            If Not reporters.Exists(actor) Then
                Set matches = New Collection
                reporters.Add actor, matches
            End If
            Set matches = reporters(actor)
            matches.Add private_Text(table.ValueAt(row, codeColumn))
        End If
    Next row
    For row = 1 To table.RowCount
        actor = private_Text(table.ValueAt(row, actorColumn))
        If VBA.Len(actor) > 0 Then
            ' Vacant posts can have an acting person without a root name or tax ID.
            key = private_Text(table.ValueAt(row, codeColumn))
            If VBA.Len(key) = 0 Then
                key = "@row:" & VBA.CStr(row)
                errors(key) = VBA.Replace$(VBA.Replace$(ex_PEB_Export.fn_Config("text.DependenciesTvoPosition"), _
                    "{row}", VBA.CStr(row)), "{code}", VBA.vbNullString)
            ElseIf chains.Exists(key) Or errors.Exists(key) Then
                If chains.Exists(key) Then chains.Remove key
                errors(key) = VBA.Replace$(VBA.Replace$(ex_PEB_Export.fn_Config("text.DependenciesTvoPosition"), _
                    "{row}", VBA.CStr(row)), "{code}", VBA.CStr(key))
            ElseIf private_TryChain(table, row, names, nameColumn, actorColumn, ipnColumn, codeColumn, chain, chainDiagnostic) Then
                chains.Add key, chain
            Else
                errors(key) = VBA.Replace$(VBA.Replace$(VBA.Replace$(ex_PEB_Export.fn_Config("text.DependenciesTvoIssue"), _
                    "{row}", VBA.CStr(row)), "{code}", VBA.CStr(key)), "{details}", chainDiagnostic)
            End If
        End If
    Next row
    snapshot.Prepared.Add "Tvo.ReporterPositions", reporters
    snapshot.Prepared.Add "Tvo.Chains", chains
    snapshot.Prepared.Add "Tvo.ChainErrors", errors
    For Each key In errors.Keys
        ex_Core.fn_Diagnostic_WriteLog "DEPENDENCY_TVO_ISSUE | Position=" & VBA.CStr(key) & " | Details=" & errors(key)
    Next key
    private_PrepareTvo = True
Cleanup:
    If Not query Is Nothing Then
        query.Dispose
    End If
    If Not table Is Nothing Then
        table.Dispose
    End If
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function

Private Function private_TryChain( _
    ByVal table As obj_DataTable, _
    ByVal rootRow As Long, _
    ByVal names As Object, _
    ByVal nameColumn As Long, _
    ByVal actorColumn As Long, _
    ByVal ipnColumn As Long, _
    ByVal codeColumn As Long, _
    ByRef chain As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim seen As Object
    Dim matches As Collection
    Dim entry As Object
    Dim current As Long
    Dim name As String
    Dim actor As String

    On Error GoTo EH
    Set chain = New Collection
    Set seen = VBA.CreateObject("Scripting.Dictionary")
    current = rootRow
    Do
        name = private_Text(table.ValueAt(current, nameColumn))
        If seen.Exists(VBA.CStr(current)) Then
            Err.Raise 5, , ex_PEB_Export.fn_Config("text.DependenciesTvoCycle") & " " & name
        End If
        seen.Add VBA.CStr(current), True
        actor = private_Text(table.ValueAt(current, actorColumn))
        If VBA.Len(actor) = 0 Then
            Exit Do
        End If
        If Not names.Exists(actor) Then
            Err.Raise 5, , ex_PEB_Export.fn_Config("text.DependenciesTvoMissing") & " " & actor
        End If
        Set matches = names(actor)
        If matches.Count <> 1 Then
            Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("reporter.tvo.ambiguous"), "{name}", actor)
        End If
        Set entry = VBA.CreateObject("Scripting.Dictionary")
        entry("PersonId") = private_Text(table.ValueAt(matches(1), ipnColumn))
        entry("PersonName") = actor
        entry("PositionCode") = private_Text(table.ValueAt(current, codeColumn))
        If VBA.Len(entry("PersonId")) = 0 Or VBA.Len(entry("PositionCode")) = 0 Then
            Err.Raise 5, , ex_PEB_Export.fn_Config("text.DependenciesTvoIncomplete")
        End If
        chain.Add entry
        current = matches(1)
    Loop
    private_TryChain = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Set chain = Nothing
End Function

Private Function private_Text(ByVal value As Variant) As String
    If VBA.IsError(value) Then
        Err.Raise 5, , "Excel error in TVO source."
    End If
    If Not VBA.IsNull(value) And Not VBA.IsEmpty(value) Then
        private_Text = ex_TableQuery.fn_NormalizeText(VBA.CStr(value))
    End If
End Function

Private Function private_FileStamp(ByVal path As String) As String
    Dim file As Object

    Set file = VBA.CreateObject("Scripting.FileSystemObject").GetFile(path)
    private_FileStamp = VBA.CStr(file.DateLastModified) & "|" & VBA.CStr(file.Size)
End Function

Private Sub private_Progress( _
    ByRef currentStep As Long, _
    ByVal total As Long, _
    ByVal stage As String, _
    Optional ByVal source As String = VBA.vbNullString _
)
    Dim text As String
    Dim stageText As String

    currentStep = currentStep + 1
    stageText = VBA.Replace$(ex_PEB_Export.fn_Config("snapshots.progress." & stage), "{source}", source)
    text = ex_PEB_Export.fn_Config("text.DependenciesProgress")
    text = VBA.Replace$(text, "{step}", VBA.CStr(currentStep))
    text = VBA.Replace$(text, "{total}", VBA.CStr(total))
    text = VBA.Replace$(text, "{stage}", stageText)
    ex_Core.fn_StatusBar_Show text
End Sub