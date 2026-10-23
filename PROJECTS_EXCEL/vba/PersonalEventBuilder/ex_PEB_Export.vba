Attribute VB_Name = "ex_PEB_Export"
Option Explicit

Private m_configSnapshot As Object
Private m_configSnapshotDepth As Long

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
    Set m_configSnapshot = Nothing
    m_configSnapshotDepth = 0
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_IsTicketDocumentForm(ByVal formKind As String) As Boolean
    fn_IsTicketDocumentForm = InStr(1, "|" & ex_PEB_Export.fn_Config("export.ticket.form") & "|", "|" & formKind & "|", vbBinaryCompare) > 0
End Function

Public Function fn_FixedDocument(ByVal modeId As String) As String
    If InStr(1, "|" & ex_PEB_Export.fn_Config("document.fixed.modes") & "|", "|" & modeId & "|", vbBinaryCompare) > 0 Then
        fn_FixedDocument = ex_PEB_Export.fn_Config("document.fixed." & modeId)
    ElseIf ex_PEB_Export.fn_IsTicketDocumentForm(ex_PEB_EventRoutes.fn_Value(modeId, "form")) Then
        fn_FixedDocument = ex_PEB_Export.fn_Config("lookup.ReturnTickets.document")
    End If
End Function

Public Function fn_IsTicketDocumentMode(ByVal modeId As String) As Boolean
    fn_IsTicketDocumentMode = (ex_PEB_Export.fn_FixedDocument(modeId) = ex_PEB_Export.fn_Config("lookup.ReturnTickets.document"))
End Function

Public Sub fn_TicketDocumentValues(ByVal values As Object, Optional ByVal stored As Boolean = False)
    Dim document As String
    Dim legacyNumber As Variant
    Dim legacyDate As Variant

    legacyNumber = vbNullString
    legacyDate = vbNullString
    If values.Exists("TicketNumber") Then legacyNumber = values("TicketNumber")
    If values.Exists("TicketDate") Then legacyDate = values("TicketDate")
    values("TicketNumber") = vbNullString
    values("TicketDate") = vbNullString
    document = ex_PEB_Export.fn_FixedDocument(CStr(values("ModeId")))
    If Len(document) > 0 And Not ex_PEB_Export.fn_IsTicketDocumentMode(CStr(values("ModeId"))) Then
        If Not stored Then
            values("Document") = document
        End If
        ' Preserve the meaning of ticket fields in previously saved event snapshots.
        If Not stored Or CStr(values("Document")) <> ex_PEB_Export.fn_Config("lookup.ReturnTickets.document") Then
            Exit Sub
        End If
    End If
    If Len(document) = 0 Then Exit Sub
    document = ex_PEB_Export.fn_Config("lookup.ReturnTickets.document")
    If stored Then
        If CStr(values("Document")) <> document Then
            ' A saved medical document must not become a vacation ticket just because the form policy changed.
            If Len(ex_PEB_Export.fn_Text(legacyNumber)) = 0 And Len(ex_PEB_Export.fn_Text(legacyDate)) = 0 Then
                Exit Sub
            End If
            If Len(ex_PEB_Export.fn_Text(legacyNumber)) > 0 Then values("DocumentNumber") = legacyNumber
            If Len(ex_PEB_Export.fn_Text(legacyDate)) > 0 Then values("DocumentDate") = legacyDate
        End If
    End If
    values("Document") = document
    ' These aliases are derived for legacy templates; history stores only the document fields.
    values("TicketNumber") = values("DocumentNumber")
    values("TicketDate") = values("DocumentDate")
End Sub

Public Function fn_Config(ByVal key As String) As String
    Dim value As String

    If Not m_configSnapshot Is Nothing Then
        If Not m_configSnapshot.Exists("PersonalEventBuilder::" & key) Then
            Err.Raise 5, , "Required configuration key not found: PersonalEventBuilder::" & key
        End If
        value = m_configSnapshot("PersonalEventBuilder::" & key)
    Else
        If Not ex_Core.fn_TryGetWorkbookConfigValue("PersonalEventBuilder::" & key, value) Then
            Err.Raise 5, , "Required configuration key not found: PersonalEventBuilder::" & key
        End If
    End If
    If VBA.Len(VBA.Trim$(value)) = 0 Then
        Err.Raise 5, , "Empty configuration key: PersonalEventBuilder::" & key
    End If
    fn_Config = value
End Function

Public Function fn_WordTemplates() As obj_PEB_WordTemplates
    Dim snapshot As obj_TableSnapshotSet
    Dim templates As obj_PEB_WordTemplates
    Dim diagnostic As String

    Set snapshot = ex_TableSnapshots.fn_Current()
    If Not snapshot Is Nothing Then
        Set fn_WordTemplates = snapshot.Prepared("WordTemplates")
        Exit Function
    End If
    Set templates = New obj_PEB_WordTemplates
    If Not templates.Initialize(diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    Set fn_WordTemplates = templates
End Function

Public Function fn_TvoChain(ByVal positionCode As String) As Collection
    Dim snapshot As obj_TableSnapshotSet
    Dim chains As Object
    Dim errors As Object
    Dim result As Collection
    Dim entry As Object
    Dim copy As Object
    Dim field As Variant

    Set snapshot = ex_TableSnapshots.fn_Current()
    If snapshot Is Nothing Then
        Err.Raise 5, , "Dependency snapshot is unavailable."
    End If
    positionCode = ex_TableQuery.fn_NormalizeText(positionCode)
    Set chains = snapshot.Prepared("Tvo.Chains")
    Set errors = snapshot.Prepared("Tvo.ChainErrors")
    If errors.Exists(positionCode) Then
        Err.Raise 5, , errors(positionCode)
    End If
    Set result = New Collection
    If chains.Exists(positionCode) Then
        For Each entry In chains(positionCode)
            Set copy = VBA.CreateObject("Scripting.Dictionary")
            For Each field In entry.Keys
                copy.Add field, entry(field)
            Next field
            result.Add copy
        Next entry
    End If
    Set fn_TvoChain = result
End Function

Public Function fn_BasisDocuments(ByVal values As Object) As Collection
    Dim documents As New Collection
    Dim document As Object
    Dim item As Object

    Set document = VBA.CreateObject("Scripting.Dictionary")
    document.Add "DocumentNote", ex_PEB_Export.fn_Required(values, "DocumentNote")
    document.Add "DocNo", ex_PEB_Export.fn_Required(values, "DocNo")
    document.Add "DocDateShort", ex_PEB_Export.fn_Required(values, "DocDateShort")
    documents.Add document
    For Each item In values("MetaDocumentTables")
        documents.Add item
    Next item
    If values.Exists("VlkNo") And values.Exists("VlkDateShort") Then
        If Len(ex_PEB_Export.fn_Text(values("VlkNo"))) > 0 Or Len(ex_PEB_Export.fn_Text(values("VlkDateShort"))) > 0 Then
            Set document = VBA.CreateObject("Scripting.Dictionary")
            document.Add "DocumentNote", ex_PEB_Export.fn_Config("extensions.role.Vlk.document")
            document.Add "DocNo", values("VlkNo")
            document.Add "DocDateShort", values("VlkDateShort")
            documents.Add document
        End If
    End If
    Set fn_BasisDocuments = documents
End Function

Public Function fn_TryApplyReporterPositions( _
    ByVal records As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim record As obj_LookupRecord
    Dim supplement As obj_LookupRecord
    Dim names As New Collection
    Dim positions As Object
    Dim fields As Object
    Dim name As Variant
    Dim key As String

    On Error GoTo EH
    For Each record In records
        If Not record.TryGetValue("Personnel.FullName", name) Then Err.Raise 5, , "Reporter name not found."
        names.Add VBA.CStr(name)
    Next record
    If Not fn_TryReporterTvoPositions(names, positions, diagnostic) Then Exit Function
    For Each record In records
        If Not record.TryGetValue("Personnel.FullName", name) Then Err.Raise 5, , "Reporter name not found."
        key = ex_TableQuery.fn_NormalizeText(VBA.CStr(name))
        If positions.Exists(key) Then
            Set fields = VBA.CreateObject("Scripting.Dictionary")
            fields("Personnel.PositionCode") = positions(key)
            Set supplement = New obj_LookupRecord
            If Not supplement.Initialize(record.Key, fields) Then Err.Raise 5, , "Cannot initialize reporter TVO position."
            If Not record.TryMerge(supplement, True) Then Err.Raise 5, , "Cannot merge reporter TVO position."
        End If
    Next record
    fn_TryApplyReporterPositions = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
End Function

Public Function fn_TryReporterTvoPosition( _
    ByVal reporterName As String, _
    ByRef positionCode As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim names As New Collection
    Dim positions As Object
    Dim key As String

    positionCode = VBA.vbNullString
    names.Add reporterName
    If Not fn_TryReporterTvoPositions(names, positions, diagnostic) Then Exit Function
    key = ex_TableQuery.fn_NormalizeText(reporterName)
    If positions.Exists(key) Then positionCode = positions(key)
    fn_TryReporterTvoPosition = True
End Function

Public Function fn_TryReporterTvoPositions( _
    ByVal reporterNames As Collection, _
    ByRef positions As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim source As obj_TableSource
    Dim query As obj_TableQuery
    Dim service As obj_TableQueryService
    Dim result As obj_DataTable
    Dim requested As Object
    Dim available As Boolean
    Dim nameHeader As String
    Dim codeHeader As String
    Dim reporterName As Variant
    Dim row As Long
    Dim key As String
    Dim positionCode As String
    Dim snapshot As obj_TableSnapshotSet
    Dim prepared As Object
    Dim matches As Collection

    On Error GoTo EH
    Set positions = VBA.CreateObject("Scripting.Dictionary")
    positions.CompareMode = VBA.vbTextCompare
    Set requested = VBA.CreateObject("Scripting.Dictionary")
    requested.CompareMode = VBA.vbTextCompare
    If reporterNames.Count = 0 Then
        fn_TryReporterTvoPositions = True
        Exit Function
    End If
    Set snapshot = ex_TableSnapshots.fn_Current()
    If Not snapshot Is Nothing Then
        Set prepared = snapshot.Prepared("Tvo.ReporterPositions")
        For Each reporterName In reporterNames
            key = ex_TableQuery.fn_NormalizeText(VBA.CStr(reporterName))
            If prepared.Exists(key) Then
                Set matches = prepared(key)
                If matches.Count <> 1 Then
                    Err.Raise 5, , VBA.Replace$(fn_Config("reporter.tvo.ambiguous"), "{name}", VBA.CStr(reporterName))
                End If
                positionCode = matches(1)
                If VBA.Len(positionCode) = 0 Then
                    Err.Raise 5, , VBA.Replace$(fn_Config("reporter.tvo.missingCode"), "{name}", VBA.CStr(reporterName))
                End If
                positions(key) = positionCode
            End If
        Next reporterName
        fn_TryReporterTvoPositions = True
        Exit Function
    End If
    Set source = New obj_TableSource
    Set query = New obj_TableQuery
    Set service = New obj_TableQueryService
    If Not source.Initialize() Or Not query.Initialize() Or Not service.Initialize() Then
        Err.Raise 5, , "Cannot initialize reporter TVO query."
    End If
    source.WorkbookPath = fn_Path(fn_Config("lookup.Personnel.source.path"))
    source.SheetName = fn_Config("lookup.Personnel.source.sheet")
    source.RangeAddress = fn_Config("lookup.Personnel.source.range")
    nameHeader = fn_Config("reporter.tvo.nameHeader")
    codeHeader = fn_Config("reporter.tvo.codeHeader")
    query.AddColumn nameHeader
    query.AddColumn codeHeader
    For Each reporterName In reporterNames
        requested(ex_TableQuery.fn_NormalizeText(VBA.CStr(reporterName))) = True
    Next reporterName
    query.AddCondition nameHeader, QueryIsNotEmpty, Empty, QueryText, True
    If Not ex_TableDataRuntime.fn_TryQueryLoaded(source, query, available, result, diagnostic) Then GoTo Cleanup
    If Not available Then
        service.Backend = QueryAuto
        If Not service.TryExecute(source, query, result, diagnostic) Then GoTo Cleanup
    End If
    For row = 1 To result.RowCount
        reporterName = fn_Text(result.ValueAt(row, ex_TableQuery.fn_HeaderIndex(result.Headers, nameHeader)))
        key = ex_TableQuery.fn_NormalizeText(VBA.CStr(reporterName))
        If requested.Exists(key) Then
            If positions.Exists(key) Then
                Err.Raise 5, , VBA.Replace$(fn_Config("reporter.tvo.ambiguous"), "{name}", VBA.CStr(reporterName))
            End If
            positionCode = fn_Text(result.ValueAt(row, ex_TableQuery.fn_HeaderIndex(result.Headers, codeHeader)))
            If VBA.Len(positionCode) = 0 Then
                Err.Raise 5, , VBA.Replace$(fn_Config("reporter.tvo.missingCode"), "{name}", VBA.CStr(reporterName))
            End If
            positions.Add key, positionCode
        End If
    Next row
    fn_TryReporterTvoPositions = True
Cleanup:
    If Not result Is Nothing Then result.Dispose
    If Not service Is Nothing Then service.Dispose
    If Not query Is Nothing Then query.Dispose
    If Not source Is Nothing Then source.Dispose
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function

Public Sub fn_BeginConfigSnapshot()
    Dim sheet As Worksheet
    Dim table As ListObject
    Dim data As Variant
    Dim keyColumn As Long
    Dim row As Long
    Dim key As String
    Dim snapshot As Object

    If m_configSnapshotDepth > 0 Then
        m_configSnapshotDepth = m_configSnapshotDepth + 1
        Exit Sub
    End If
    For Each sheet In ThisWorkbook.Worksheets
        For Each table In sheet.ListObjects
            If VBA.StrComp(table.Name, "tbConfig", VBA.vbTextCompare) = 0 Then
                keyColumn = table.ListColumns("Key").Index
                If table.DataBodyRange Is Nothing Or keyColumn >= table.ListColumns.Count Then
                    Err.Raise 5, , "Export configuration table is empty or invalid."
                End If
                data = table.DataBodyRange.Value2
                Set snapshot = VBA.CreateObject("Scripting.Dictionary")
                snapshot.CompareMode = VBA.vbTextCompare
                For row = 1 To UBound(data, 1)
                    key = VBA.Trim$(VBA.CStr(data(row, keyColumn)))
                    If snapshot.Exists(key) Then
                        Err.Raise 5, , "Duplicate configuration key: " & key
                    End If
                    snapshot.Add key, VBA.Trim$(VBA.CStr(data(row, keyColumn + 1)))
                Next row
                Set m_configSnapshot = snapshot
                m_configSnapshotDepth = 1
                Exit Sub
            End If
        Next table
    Next sheet
    Err.Raise 5, , "Export configuration table not found: tbConfig"
End Sub

Public Sub fn_EndConfigSnapshot()
    If m_configSnapshotDepth > 0 Then m_configSnapshotDepth = m_configSnapshotDepth - 1
    If m_configSnapshotDepth = 0 Then Set m_configSnapshot = Nothing
End Sub

Public Function fn_Count( _
    ByVal key As String, _
    Optional ByVal maximum As Long = 1000 _
) As Long
    Dim value As String
    Dim number As Double

    value = fn_Config(key)
    If Not VBA.IsNumeric(value) Then
        Err.Raise 5, , "Invalid configuration count: " & key
    End If
    number = VBA.CDbl(value)
    If number < 0 Or number > maximum Or number <> VBA.Fix(number) Then
        Err.Raise 5, , "Invalid configuration count: " & key
    End If
    fn_Count = VBA.CLng(number)
End Function

Public Function fn_Path(ByVal value As String) As String
    Dim fileSystem As Object

    Set fileSystem = VBA.CreateObject("Scripting.FileSystemObject")
    value = VBA.Replace$(value, "/", "\")
    If VBA.Left$(value, 2) <> "\\" And VBA.Mid$(value, 2, 1) <> ":" Then
        value = fileSystem.BuildPath(ThisWorkbook.Path, value)
    End If
    fn_Path = fileSystem.GetAbsolutePathName(value)
End Function

Public Function fn_Text(ByVal value As Variant) As String
    If VBA.IsError(value) Then
        Err.Raise 5, , "An export value contains an Excel error."
    End If
    If VBA.IsNull(value) Or VBA.IsEmpty(value) Then
        Exit Function
    End If
    fn_Text = VBA.Trim$(VBA.CStr(value))
End Function

Public Function fn_Required( _
    ByVal values As Object, _
    ByVal key As String _
) As Variant
    If Not values.Exists(key) Then
        Err.Raise 5, , "Unknown export field: " & key
    End If
    If VBA.IsObject(values(key)) Then
        Err.Raise 5, , "Scalar export field required: " & key
    End If
    fn_Required = values(key)
End Function

Public Function fn_ParseDate( _
    ByVal value As Variant, _
    ByVal key As String _
) As Date
    Dim parts As Variant
    Dim parsed As Date
    Dim text As String
    Dim rx As Object
    Dim today As Date
    Dim dateDay As Long
    Dim dateMonth As Long
    Dim dateYear As Long

    If VBA.VarType(value) = VBA.vbDate Then
        fn_ParseDate = VBA.DateValue(value)
        Exit Function
    End If
    If VBA.IsNumeric(value) And VBA.VarType(value) <> VBA.vbString Then
        If VBA.CDbl(value) > 31 Then
            If VBA.CDbl(value) > 2958465 Then
                Err.Raise 5, , "Invalid Excel date: " & key
            End If
            fn_ParseDate = VBA.DateValue(VBA.CDate(value))
            Exit Function
        End If
    End If
    text = ex_PEB_Export.fn_Text(value)
    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.Pattern = "^\d{1,2}(\.\d{1,2}(\.\d{4})?)?$"
    If Not rx.Test(text) Then
        Err.Raise 5, , "Date D, D.M or D.M.YYYY required: " & key
    End If
    parts = VBA.Split(text, ".")
    today = VBA.DateTime.Date
    dateDay = VBA.CLng(parts(0))
    dateMonth = VBA.DateTime.Month(today)
    dateYear = VBA.DateTime.Year(today)
    If UBound(parts) >= 1 Then dateMonth = VBA.CLng(parts(1))
    If UBound(parts) = 2 Then dateYear = VBA.CLng(parts(2))
    If dateYear < 1900 Then
        Err.Raise 5, , "Invalid date year: " & key
    End If
    parsed = VBA.DateSerial(dateYear, dateMonth, dateDay)
    If VBA.DateTime.Day(parsed) <> dateDay Or VBA.DateTime.Month(parsed) <> dateMonth Or VBA.DateTime.Year(parsed) <> dateYear Then
        Err.Raise 5, , "Invalid calendar date: " & key
    End If
    fn_ParseDate = parsed
End Function

Public Function fn_ExtensionValue(ByVal value As Variant, ByVal fieldType As String, ByVal field As String) As Variant
    Select Case fieldType
        Case "text", "number", "date"
        Case Else
            Err.Raise 5, , "Unsupported extension field type: " & fieldType & " (" & field & ")"
    End Select
    If Len(ex_PEB_Export.fn_Text(value)) = 0 Then
        fn_ExtensionValue = vbNullString
        Exit Function
    End If
    Select Case fieldType
        Case "text"
            fn_ExtensionValue = CStr(value)
        Case "number"
            If Not IsNumeric(value) Then Err.Raise 5, , Replace$(ex_PEB_Export.fn_Config("extensions.invalid"), "{field}", field)
            fn_ExtensionValue = CDbl(value)
        Case "date"
            fn_ExtensionValue = ex_PEB_Export.fn_ParseDate(value, field)
    End Select
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------