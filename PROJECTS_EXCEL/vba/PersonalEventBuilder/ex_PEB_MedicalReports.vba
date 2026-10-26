Attribute VB_Name = "ex_PEB_MedicalReports"
Option Explicit

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_TryLoad( _
    ByVal orderNumber As String, _
    ByVal orderDate As Date, _
    ByRef result As obj_LookupResult, _
    ByRef status As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim files As Object
    Dim rows As Object
    Dim fileSystem As Object
    Dim path As Variant
    Dim fields As Object
    Dim records As New Collection
    Dim columns As New Collection
    Dim column As Object
    Dim record As obj_LookupRecord
    Dim field As Variant
    Dim root As String
    Dim patterns As Variant
    Dim pattern As Variant
    Dim reportFields As Variant
    Dim folder As String
    Dim day As Date
    Dim offset As Long
    Dim windowDays As Long
    Dim count As Long
    Dim warnings As Long
    Dim queryService As obj_TableQueryService
    Dim dateRx As Object
    Dim startedAt As Double
    Dim stageAt As Double
    Dim configSnapshotActive As Boolean

    On Error GoTo EH
    startedAt = VBA.Timer
    Set result = Nothing
    diagnostic = VBA.vbNullString
    stageAt = VBA.Timer
    ex_PEB_Export.fn_BeginConfigSnapshot
    configSnapshotActive = True
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.Config", stageAt
    stageAt = VBA.Timer
    orderDate = VBA.DateValue(orderDate)
    root = ex_PEB_Export.fn_Path(private_Config("root"))
    patterns = VBA.Split(private_Config("patterns"), "|")
    For Each pattern In patterns
        If VBA.InStr(VBA.CStr(pattern), "{dd.mm.yyyy}") = 0 And VBA.InStr(VBA.CStr(pattern), "{yyyymmdd}") = 0 Then
            private_Fail "MissingDateToken", VBA.CStr(pattern)
        End If
    Next pattern
    windowDays = ex_PEB_Export.fn_Count("validation.MedicalReports.windowDays", 366)
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FolderExists(root) Then
        Err.Raise 5, , VBA.Replace$(private_Text("MissingFolder"), "{path}", root)
    End If
    Set files = VBA.CreateObject("Scripting.Dictionary")
    files.CompareMode = VBA.vbTextCompare
    Set rows = VBA.CreateObject("Scripting.Dictionary")
    rows.CompareMode = VBA.vbTextCompare

    ' Discover reports independently of their modification timestamps.
    For offset = -windowDays To windowDays
        day = VBA.DateAdd("d", offset, orderDate)
        folder = fileSystem.BuildPath(root, private_Expand(private_Config("folder"), day))
        If fileSystem.FolderExists(folder) Then
            private_CollectFiles fileSystem.GetFolder(folder), patterns, files, day
        End If
    Next offset
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.Discover | Files=" & VBA.CStr(files.Count), stageAt
    Set queryService = New obj_TableQueryService
    If Not queryService.Initialize() Then private_Fail "InitFailed", "obj_TableQueryService"
    queryService.Backend = QueryAdo
    reportFields = VBA.Split(private_Config("fields"), "|")
    Set dateRx = VBA.CreateObject("VBScript.RegExp")
    dateRx.Pattern = "^\s*([0-9]{2})\.([0-9]{2})\.([0-9]{4})([^0-9]|$)"
    stageAt = VBA.Timer
    For Each path In files.Keys
        If Not private_ReadFile(VBA.CStr(path), orderDate, reportFields, rows, queryService, dateRx, diagnostic) Then
            Err.Raise 5, , diagnostic
        End If
    Next path
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.Reports | Files=" & VBA.CStr(files.Count) & " | Rows=" & VBA.CStr(rows.Count), stageAt

    If rows.Count > 0 Then
        stageAt = VBA.Timer
        If Not private_CheckPeople(rows, queryService, warnings, diagnostic) Then Err.Raise 5, , diagnostic
        ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.Identity", stageAt
        stageAt = VBA.Timer
        If Not ex_PEB_MedicalState.fn_TryApply(rows, orderNumber, orderDate, queryService, warnings, diagnostic) Then Err.Raise 5, , diagnostic
        ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.State", stageAt
    End If

    ' Keep differing report versions and merge only identical business values.
    stageAt = VBA.Timer
    For Each path In rows.Keys
        Set fields = rows(path)
        count = count + 1
        Set record = New obj_LookupRecord
        If Not record.Initialize(VBA.CStr(count), fields) Then private_Fail "InitFailed", "obj_LookupRecord"
        records.Add record
    Next path
    For Each field In VBA.Split(private_Config("display"), "|")
        Set column = VBA.CreateObject("Scripting.Dictionary")
        column("name") = field
        column("caption") = private_Text("Columns." & VBA.CStr(field))
        columns.Add column
    Next field
    Set result = New obj_LookupResult
    If Not result.Initialize(records, columns, False, diagnostic, "RowTags") Then Err.Raise 5, , diagnostic
    status = private_Text("Loaded")
    If files.Count = 0 Then status = private_Text("NoFiles")
    status = VBA.Replace$(status, "{date}", VBA.Format$(orderDate, "dd.mm.yyyy"))
    status = VBA.Replace$(status, "{files}", VBA.CStr(files.Count))
    status = VBA.Replace$(status, "{rows}", VBA.CStr(records.Count))
    status = VBA.Replace$(status, "{warnings}", VBA.CStr(warnings))
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.Result | Rows=" & VBA.CStr(records.Count), stageAt
    fn_TryLoad = True
Cleanup:
    If Not queryService Is Nothing Then queryService.Dispose
    If configSnapshotActive Then ex_PEB_Export.fn_EndConfigSnapshot
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.Load | Success=" & VBA.CStr(fn_TryLoad), startedAt
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Set result = Nothing
    Resume Cleanup
End Function

Public Function fn_TrySearch( _
    ByVal dateFrom As Date, _
    ByVal dateTo As Date, _
    ByVal personName As String, _
    ByVal eventDate As Date, _
    ByRef result As obj_UiRawTableList, _
    ByRef status As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim fileSystem As Object
    Dim files As Object
    Dim rows As Object
    Dim patterns As Variant
    Dim pattern As Variant
    Dim reportFields As Variant
    Dim displayFields As Variant
    Dim root As String
    Dim folder As String
    Dim day As Date
    Dim path As Variant
    Dim key As Variant
    Dim fields As Object
    Dim queryService As obj_TableQueryService
    Dim dateRx As Object
    Dim table As obj_UiRawTable
    Dim headers As Variant
    Dim values As Variant
    Dim row As Long
    Dim column As Long
    Dim found As Long
    Dim startedAt As Double
    Dim stageAt As Double
    Dim configSnapshotActive As Boolean

    On Error GoTo EH
    startedAt = VBA.Timer
    Set result = Nothing
    diagnostic = VBA.vbNullString
    ex_PEB_Export.fn_BeginConfigSnapshot
    configSnapshotActive = True
    dateFrom = VBA.DateValue(dateFrom)
    dateTo = VBA.DateValue(dateTo)
    If dateTo < dateFrom Then Err.Raise 5, , private_Config("search.text.InvalidRange")
    If VBA.DateDiff("d", dateFrom, dateTo) > ex_PEB_Export.fn_Count("validation.MedicalReports.search.maxDays", 3660) Then
        Err.Raise 5, , private_Config("search.text.RangeTooWide")
    End If
    root = ex_PEB_Export.fn_Path(private_Config("root"))
    patterns = VBA.Split(private_Config("patterns"), "|")
    For Each pattern In patterns
        If VBA.InStr(VBA.CStr(pattern), "{dd.mm.yyyy}") = 0 And VBA.InStr(VBA.CStr(pattern), "{yyyymmdd}") = 0 Then
            private_Fail "MissingDateToken", VBA.CStr(pattern)
        End If
    Next pattern
    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FolderExists(root) Then Err.Raise 5, , VBA.Replace$(private_Text("MissingFolder"), "{path}", root)
    Set files = VBA.CreateObject("Scripting.Dictionary")
    files.CompareMode = VBA.vbTextCompare
    stageAt = VBA.Timer
    day = dateFrom
    Do
        folder = fileSystem.BuildPath(root, private_Expand(private_Config("folder"), day))
        If fileSystem.FolderExists(folder) Then private_CollectFiles fileSystem.GetFolder(folder), patterns, files, day
        If day = dateTo Then Exit Do
        day = VBA.DateAdd("d", 1, day)
    Loop
    ex_Core.fn_Diagnostic_WritePerf "MedicalSearch.Discover | Files=" & VBA.CStr(files.Count), stageAt
    Set queryService = New obj_TableQueryService
    If Not queryService.Initialize() Then private_Fail "InitFailed", "obj_TableQueryService"
    queryService.Backend = QueryAdo
    reportFields = VBA.Split(private_Config("fields"), "|")
    displayFields = VBA.Split(private_Config("search.display"), "|")
    ReDim headers(0 To UBound(displayFields))
    For column = 0 To UBound(displayFields)
        headers(column) = private_Text("Columns." & displayFields(column))
    Next column
    Set dateRx = VBA.CreateObject("VBScript.RegExp")
    If eventDate <> 0 Then
        dateRx.Pattern = "(^|[^0-9])" & VBA.Replace$(VBA.Format$(eventDate, "dd.mm.yyyy"), ".", "\.") & "([^0-9]|$)"
    End If
    Set result = New obj_UiRawTableList
    If Not result.Initialize() Then private_Fail "InitFailed", "obj_UiRawTableList"
    stageAt = VBA.Timer
    For Each path In files.Keys
        Set rows = VBA.CreateObject("Scripting.Dictionary")
        rows.CompareMode = VBA.vbTextCompare
        If Not private_ReadFile(VBA.CStr(path), eventDate, reportFields, rows, queryService, dateRx, diagnostic, True, personName) Then
            Err.Raise 5, , diagnostic
        End If
        If rows.Count > 0 Then
            ReDim values(1 To rows.Count, 1 To UBound(displayFields) + 1)
            row = 0
            For Each key In rows.Keys
                Set fields = rows(key)
                row = row + 1
                For column = 0 To UBound(displayFields)
                    values(row, column + 1) = fields(displayFields(column))
                Next column
            Next key
            Set table = New obj_UiRawTable
            If Not table.Initialize(values, headers, VBA.Mid$(VBA.CStr(path), VBA.InStrRev(VBA.CStr(path), "\") + 1)) Then
                private_Fail "InitFailed", "obj_UiRawTable"
            End If
            If Not result.Add(table) Then private_Fail "InitFailed", "obj_UiRawTableList.Add"
            found = found + rows.Count
        End If
    Next path
    ex_Core.fn_Diagnostic_WritePerf "MedicalSearch.Read | Rows=" & VBA.CStr(found), stageAt
    status = VBA.Replace$(private_Config("search.text.Loaded"), "{files}", VBA.CStr(files.Count))
    status = VBA.Replace$(status, "{rows}", VBA.CStr(found))
    fn_TrySearch = True
Cleanup:
    If Not queryService Is Nothing Then queryService.Dispose
    If configSnapshotActive Then ex_PEB_Export.fn_EndConfigSnapshot
    ex_Core.fn_Diagnostic_WritePerf "MedicalSearch.Total | Success=" & VBA.CStr(fn_TrySearch), startedAt
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Set result = Nothing
    Resume Cleanup
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

' --------------------------------------
' namespace Private {
' --------------------------------------
Private Sub private_CollectFiles( _
    ByVal folder As Object, _
    ByVal patterns As Variant, _
    ByVal files As Object, _
    ByVal day As Date _
)
    Dim file As Object
    Dim pattern As Variant

    For Each file In folder.Files
        If VBA.Left$(file.Name, 2) <> "~$" Then
            For Each pattern In patterns
                If VBA.Len(VBA.Trim$(pattern)) = 0 Then private_Fail "EmptyPattern", "validation.MedicalReports.patterns"
                If VBA.LCase$(file.Name) Like VBA.LCase$(private_Expand(VBA.CStr(pattern), day)) Then
                    files(file.Path) = True
                    Exit For
                End If
            Next pattern
        End If
    Next file
End Sub

Private Function private_Expand( _
    ByVal pattern As String, _
    ByVal day As Date _
) As String
    Dim token As Variant
    Dim replacement As String

    For Each token In VBA.Array("yyyy.mm", "dd.mm.yyyy", "yyyymmdd", "yyyy")
        replacement = VBA.Format$(day, VBA.CStr(token))
        pattern = VBA.Replace$(pattern, "{" & token & "}", replacement)
    Next token
    If VBA.InStr(pattern, "{") > 0 Or VBA.InStr(pattern, "}") > 0 Then
        private_Fail "UnknownToken", pattern
    End If
    private_Expand = pattern
End Function

Private Function private_ReadFile( _
    ByVal path As String, _
    ByVal orderDate As Date, _
    ByVal reportFields As Variant, _
    ByVal rows As Object, _
    ByVal queryService As obj_TableQueryService, _
    ByVal dateRx As Object, _
    ByRef diagnostic As String, _
    Optional ByVal searchMode As Boolean = False, _
    Optional ByVal personName As String = "" _
) As Boolean
    Dim source As obj_TableSource
    Dim query As obj_TableQuery
    Dim data As obj_DataTable
    Dim indexes As Object
    Dim fields As Object
    Dim existing As Object
    Dim field As Variant
    Dim row As Long
    Dim key As String
    Dim value As Variant
    Dim failure As String
    Dim fileName As String
    Dim startedAt As Double
    Dim stageAt As Double
    Dim matches As Boolean
    Dim sourceColumns As Variant
    Dim columnIndex As Long
    Dim dataRows As Long
    Dim matchedRows() As Boolean
    Dim matchedCount As Long
    Dim resultStartedAt As Double
    Dim orderDateText As String

    On Error GoTo EH
    startedAt = VBA.Timer
    stageAt = VBA.Timer
    fileName = VBA.Mid$(path, VBA.InStrRev(path, "\") + 1)
    Set source = New obj_TableSource
    Set query = New obj_TableQuery
    If Not source.Initialize() Or Not query.Initialize() Then private_Fail "InitFailed", "obj_TableSource / obj_TableQuery"
    source.WorkbookPath = path
    source.SheetName = private_Config("source.sheet")
    source.RangeAddress = private_Config("source.range")
    ' Reports are selected and read afresh on validation; they are not session dependencies.
    If Not queryService.TryExecute(source, query, data, diagnostic, False) Then Err.Raise 5, , diagnostic
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.File.Query | File=" & VBA.Mid$(path, VBA.InStrRev(path, "\") + 1) & " | Rows=" & VBA.CStr(data.RowCount), stageAt
    stageAt = VBA.Timer
    Set indexes = VBA.CreateObject("Scripting.Dictionary")
    sourceColumns = VBA.Split(private_Config("source.columns"), "|")
    If data.ColumnCount <> UBound(sourceColumns) + 1 Then
        private_Fail "MissingHeader", "validation.MedicalReports.source.columns / " & fileName
    End If
    For columnIndex = LBound(sourceColumns) To UBound(sourceColumns)
        indexes.Add VBA.CStr(sourceColumns(columnIndex)), columnIndex + 1
    Next columnIndex
    For Each field In reportFields
        If Not indexes.Exists(VBA.CStr(field)) Then
            private_Fail "MissingHeader", "validation.MedicalReports.source.columns / " & VBA.CStr(field)
        End If
    Next field
    If Not indexes.Exists("RowNumber") Then
        private_Fail "MissingHeader", "validation.MedicalReports.source.columns / RowNumber"
    End If
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.File.Columns | File=" & fileName, stageAt
    stageAt = VBA.Timer
    ' The first empty number cell separates the current patients from discharged people.
    For row = 1 To data.RowCount
        value = data.ValueAt(row, indexes("RowNumber"))
        If VBA.IsError(value) Then private_Fail "InvalidCell", "RowNumber / " & VBA.CStr(row)
        If VBA.IsNull(value) Or VBA.IsEmpty(value) Then Exit For
        If VBA.Len(private_Normalize(VBA.CStr(value))) = 0 Then Exit For
        dataRows = row
    Next row
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.File.Bounds | File=" & fileName & " | Rows=" & VBA.CStr(dataRows), stageAt
    stageAt = VBA.Timer
    ReDim matchedRows(0 To dataRows)
    orderDateText = VBA.Format$(orderDate, "dd.mm.yyyy")
    For row = 1 To dataRows
        value = data.ValueAt(row, indexes("Dates"))
        If searchMode Then
            matches = private_SearchDateMatches(value, orderDate, dateRx)
            If matches Then
                value = data.ValueAt(row, indexes("PersonName"))
                If VBA.IsError(value) Then private_Fail "InvalidCell", "PersonName / " & VBA.CStr(row)
                If VBA.IsNull(value) Or VBA.IsEmpty(value) Then value = VBA.vbNullString
                matches = VBA.Len(private_Normalize(VBA.CStr(value))) > 0 And _
                    VBA.InStr(1, private_Normalize(VBA.CStr(value)), private_Normalize(personName), VBA.vbTextCompare) > 0
            End If
        Else
            matches = private_FirstDateMatches(value, orderDate, orderDateText)
        End If
        matchedRows(row) = matches
        If matches Then
            matchedCount = matchedCount + 1
        End If
    Next row
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.File.DateFilter | File=" & fileName & _
        " | Rows=" & VBA.CStr(dataRows) & " | Matches=" & VBA.CStr(matchedCount) & _
        " | Search=" & VBA.CStr(searchMode), stageAt
    resultStartedAt = VBA.Timer
    For row = 1 To dataRows
        If matchedRows(row) Then
            Set fields = VBA.CreateObject("Scripting.Dictionary")
            key = VBA.vbNullString
            For Each field In reportFields
                value = data.ValueAt(row, indexes(field))
                If VBA.IsError(value) Then private_Fail "InvalidCell", VBA.CStr(field) & " / " & VBA.CStr(row)
                If VBA.IsNull(value) Or VBA.IsEmpty(value) Then value = VBA.vbNullString
                If VBA.VarType(value) = VBA.vbDate Then value = VBA.Format$(value, "dd.mm.yyyy")
                fields.Add VBA.CStr(field), VBA.CStr(value)
                key = key & VBA.Len(private_Normalize(VBA.CStr(value))) & ":" & private_Normalize(VBA.CStr(value))
            Next field
            If rows.Exists(key) Then
                Set existing = rows(key)
                existing("Sources") = existing("Sources") & VBA.vbLf & fileName
            Else
                fields.Add "Sources", fileName
                rows.Add key, fields
            End If
        End If
    Next row
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.File.BuildResults | File=" & fileName & _
        " | Matches=" & VBA.CStr(matchedCount), resultStartedAt
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.File.Filter | File=" & VBA.Mid$(path, VBA.InStrRev(path, "\") + 1), stageAt
    private_ReadFile = True
Cleanup:
    If Not data Is Nothing Then data.Dispose
    If Not query Is Nothing Then query.Dispose
    If Not source Is Nothing Then source.Dispose
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.File.Total | Success=" & VBA.CStr(private_ReadFile), startedAt
    Exit Function
EH:
    failure = VBA.Err.Description
    diagnostic = VBA.Replace$(private_Text("ReadFailed"), "{path}", path)
    diagnostic = VBA.Replace$(diagnostic, "{error}", failure)
    Resume Cleanup
End Function

Private Function private_CheckPeople( _
    ByVal rows As Object, _
    ByVal queryService As obj_TableQueryService, _
    ByRef warnings As Long, _
    ByRef diagnostic As String _
) As Boolean
    Dim source As obj_TableSource
    Dim query As obj_TableQuery
    Dim data As obj_DataTable
    Dim people As Object
    Dim names As Object
    Dim fields As Object
    Dim key As Variant
    Dim name As Variant
    Dim row As Long
    Dim idColumn As Long
    Dim nameColumn As Long
    Dim personId As String
    Dim fullName As String
    Dim referenceNames As String
    Dim check As String
    Dim neededIds As Object
    Dim referenceValues As Variant
    Dim stageAt As Double

    On Error GoTo EH
    stageAt = VBA.Timer
    Set neededIds = VBA.CreateObject("Scripting.Dictionary")
    For Each key In rows.Keys
        Set fields = rows(key)
        personId = private_PersonId(fields("PersonId"))
        If VBA.Len(personId) > 0 Then neededIds(personId) = True
    Next key
    Set source = New obj_TableSource
    Set query = New obj_TableQuery
    If Not source.Initialize() Or Not query.Initialize() Then private_Fail "InitFailed", "obj_TableSource / obj_TableQuery"
    source.WorkbookPath = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("lookup.Personnel.source.path"))
    source.SheetName = ex_PEB_Export.fn_Config("lookup.Personnel.source.sheet")
    source.RangeAddress = ex_PEB_Export.fn_Config("lookup.Personnel.source.range")
    query.AddColumn ex_PEB_Export.fn_Config("lookup.Personnel.columns.1.header")
    query.AddColumn ex_PEB_Export.fn_Config("lookup.Personnel.columns.2.header")
    ' Use the same loaded personnel generation as the rest of the page.
    If Not queryService.TryExecute(source, query, data, diagnostic) Then Err.Raise 5, , diagnostic
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.Personnel.Query | Rows=" & VBA.CStr(data.RowCount) & " | Needed=" & VBA.CStr(neededIds.Count), stageAt
    stageAt = VBA.Timer
    nameColumn = private_HeaderIndex(data.Headers, ex_PEB_Export.fn_Config("lookup.Personnel.columns.1.header"))
    idColumn = private_HeaderIndex(data.Headers, ex_PEB_Export.fn_Config("lookup.Personnel.columns.2.header"))
    Set people = VBA.CreateObject("Scripting.Dictionary")
    referenceValues = data.Values
    For row = 1 To data.RowCount
        personId = private_PersonId(referenceValues(row, idColumn))
        ' Build name sets only for people actually present in the reports.
        If neededIds.Exists(personId) Then
            fullName = private_CellText(referenceValues(row, nameColumn))
            If VBA.Len(private_Name(fullName)) > 0 Then
                If Not people.Exists(personId) Then
                    Set names = VBA.CreateObject("Scripting.Dictionary")
                    people.Add personId, names
                End If
                Set names = people(personId)
                names(private_Name(fullName)) = fullName
            End If
        End If
    Next row
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.Personnel.Index | People=" & VBA.CStr(people.Count), stageAt
    stageAt = VBA.Timer
    For Each key In rows.Keys
        Set fields = rows(key)
        personId = private_PersonId(fields("PersonId"))
        referenceNames = VBA.vbNullString
        check = "Matched"
        If VBA.Len(private_Normalize(fields("PersonId"))) = 0 Then
            check = "MissingId"
        ElseIf VBA.Len(personId) = 0 Then
            check = "InvalidId"
        ElseIf Not people.Exists(personId) Then
            check = "UnknownId"
        Else
            Set names = people(personId)
            For Each name In names.Keys
                If VBA.Len(referenceNames) > 0 Then referenceNames = referenceNames & VBA.vbLf
                referenceNames = referenceNames & names(name)
            Next name
            If names.Count > 1 Then
                check = "AmbiguousId"
            ElseIf Not names.Exists(private_Name(fields("PersonName"))) Then
                check = "NameMismatch"
            End If
        End If
        fields.Add "ReferenceName", referenceNames
        fields.Add "IdentityStatus", private_Text("Identity." & check)
        fields.Add "RowTags", VBA.vbNullString
        If check <> "Matched" Then
            fields("RowTags") = "medicalReview"
            warnings = warnings + 1
        End If
    Next key
    ex_Core.fn_Diagnostic_WritePerf "MedicalValidation.Personnel.Compare | Rows=" & VBA.CStr(rows.Count), stageAt
    private_CheckPeople = True
Cleanup:
    If Not data Is Nothing Then data.Dispose
    If Not query Is Nothing Then query.Dispose
    If Not source Is Nothing Then source.Dispose
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function

Private Function private_CellText(ByVal value As Variant) As String
    If VBA.IsError(value) Then private_Fail "InvalidCell", "Personnel"
    If VBA.IsNull(value) Or VBA.IsEmpty(value) Then Exit Function
    private_CellText = VBA.CStr(value)
End Function

Private Function private_PersonId(ByVal value As Variant) As String
    Dim text As String

    text = private_CellText(value)
    text = VBA.Replace$(text, " ", "")
    text = VBA.Replace$(text, VBA.ChrW(160), "")
    text = VBA.Replace$(text, VBA.vbTab, "")
    text = VBA.Replace$(text, VBA.vbCr, "")
    text = VBA.Replace$(text, VBA.vbLf, "")
    If VBA.Left$(text, 1) = "'" Then text = VBA.Mid$(text, 2)
    If text Like "##########" Then private_PersonId = text
End Function

Private Function private_Name(ByVal value As String) As String
    value = VBA.Replace$(value, VBA.ChrW(&H2019), "'")
    value = VBA.Replace$(value, VBA.ChrW(&H2018), "'")
    value = VBA.Replace$(value, VBA.ChrW(&H2BC), "'")
    private_Name = private_Normalize(value)
End Function

Private Function private_SearchDateMatches( _
    ByVal value As Variant, _
    ByVal eventDate As Date, _
    ByVal dateRx As Object _
) As Boolean
    If VBA.IsError(value) Then private_Fail "InvalidCell", "Dates"
    If eventDate = 0 Then
        private_SearchDateMatches = True
    ElseIf Not VBA.IsNull(value) And Not VBA.IsEmpty(value) Then
        If VBA.VarType(value) = VBA.vbDate Or VBA.IsNumeric(value) Then
            private_SearchDateMatches = (VBA.Fix(VBA.CDbl(value)) = VBA.CDbl(eventDate))
        Else
            private_SearchDateMatches = dateRx.Test(VBA.CStr(value))
        End If
    End If
End Function

Private Function private_FirstDateMatches( _
    ByVal value As Variant, _
    ByVal orderDate As Date, _
    ByVal orderDateText As String _
) As Boolean
    Dim text As String
    Dim index As Long
    Dim character As Long

    If VBA.IsNull(value) Or VBA.IsEmpty(value) Then Exit Function
    If VBA.IsError(value) Then private_Fail "InvalidCell", "Dates"
    If VBA.VarType(value) = VBA.vbDate Then
        private_FirstDateMatches = (VBA.DateValue(value) = orderDate)
        Exit Function
    End If
    If VBA.IsNumeric(value) Then
        private_FirstDateMatches = (VBA.Fix(VBA.CDbl(value)) = VBA.CDbl(orderDate))
        Exit Function
    End If
    text = VBA.CStr(value)
    index = 1
    ' Match only the first date, preserving whitespace and digit-boundary rules.
    Do While index <= VBA.Len(text)
        character = VBA.AscW(VBA.Mid$(text, index, 1))
        If character <> 32 And character <> 160 And (character < 9 Or character > 13) Then
            Exit Do
        End If
        index = index + 1
    Loop
    If VBA.Mid$(text, index, 10) <> orderDateText Then Exit Function
    If VBA.Len(text) >= index + 10 Then
        character = VBA.AscW(VBA.Mid$(text, index + 10, 1))
        If character >= 48 And character <= 57 Then Exit Function
    End If
    private_FirstDateMatches = True
End Function

Private Function private_HeaderIndex( _
    ByVal headers As Variant, _
    ByVal expected As String _
) As Long
    Dim index As Long

    For index = LBound(headers) To UBound(headers)
        If private_Normalize(VBA.Replace$(VBA.Replace$(VBA.CStr(headers(index)), ".", "#"), "_", " ")) = _
                private_Normalize(VBA.Replace$(VBA.Replace$(expected, ".", "#"), "_", " ")) Then
            If private_HeaderIndex > 0 Then private_Fail "AmbiguousHeader", expected
            private_HeaderIndex = index - LBound(headers) + 1
        End If
    Next index
    If private_HeaderIndex = 0 Then private_Fail "MissingHeader", expected
End Function

Private Function private_Normalize(ByVal value As String) As String
    value = VBA.Replace$(value, VBA.vbCr, " ")
    value = VBA.Replace$(value, VBA.vbLf, " ")
    value = VBA.Replace$(value, VBA.vbTab, " ")
    value = VBA.Replace$(value, VBA.ChrW(160), " ")
    Do While VBA.InStr(value, "  ") > 0
        value = VBA.Replace$(value, "  ", " ")
    Loop
    private_Normalize = VBA.LCase$(VBA.Trim$(value))
End Function

Private Function private_Config(ByVal key As String) As String
    private_Config = ex_PEB_Export.fn_Config("validation.MedicalReports." & key)
End Function

Private Function private_Text(ByVal key As String) As String
    private_Text = private_Config("text." & key)
End Function

Private Sub private_Fail( _
    ByVal key As String, _
    ByVal value As String _
)
    Err.Raise 5, , VBA.Replace$(private_Text(key), "{value}", value)
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------