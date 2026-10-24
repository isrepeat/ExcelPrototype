Attribute VB_Name = "ex_PEB_EjosValidation"
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
    Dim service As obj_TableQueryService
    Dim data As obj_DataTable
    Dim rows As New Collection
    Dim records As New Collection
    Dim columns As New Collection
    Dim indexes As Object
    Dim byName As Object
    Dim byCode As Object
    Dim people As Object
    Dim duplicates As Object
    Dim fields As Object
    Dim column As Object
    Dim item As Variant
    Dim headers As Variant
    Dim record As obj_LookupRecord
    Dim events As Collection
    Dim index As Long
    Dim row As Long
    Dim day As Date
    Dim dateHeader As String
    Dim duplicateKey As String
    Dim warnings As Long
    Dim cancelled As Long
    Dim recorded As Long
    Dim startedAt As Double
    Dim stageAt As Double
    Dim configSnapshotActive As Boolean

    On Error GoTo EH
    startedAt = VBA.Timer
    Set result = Nothing
    diagnostic = VBA.vbNullString
    orderDate = VBA.DateValue(orderDate)
    ex_PEB_Export.fn_BeginConfigSnapshot
    configSnapshotActive = True
    Set service = New obj_TableQueryService
    If Not service.Initialize() Then private_Fail "InitFailed"
    headers = VBA.Split(private_Config("columns"), "|")
    stageAt = VBA.Timer
    If Not private_Query("lookup.Absence.source", headers, service, data, diagnostic) Then Err.Raise 5, , diagnostic
    ex_Core.fn_Diagnostic_WritePerf "EjosValidation.Source | Rows=" & VBA.CStr(data.RowCount), stageAt
    stageAt = VBA.Timer
    Set indexes = VBA.CreateObject("Scripting.Dictionary")
    Set column = VBA.CreateObject("Scripting.Dictionary")
    column("name") = "ValidationNote"
    column("caption") = private_Message("NoteHeader")
    columns.Add column
    For Each item In headers
        indexes.Add VBA.CStr(item), ex_TableQuery.fn_HeaderIndex(data.Headers, VBA.CStr(item))
        Set column = VBA.CreateObject("Scripting.Dictionary")
        column("name") = item
        column("caption") = item
        columns.Add column
    Next item
    dateHeader = private_Config("fields.orderDate")
    Set duplicates = VBA.CreateObject("Scripting.Dictionary")
    For row = 1 To data.RowCount
        If private_TryDate(data.ValueAt(row, indexes(dateHeader)), day) Then
            If day = VBA.DateValue(orderDate) Then
                Set fields = VBA.CreateObject("Scripting.Dictionary")
                For Each item In headers
                    fields.Add VBA.CStr(item), private_Text(data.ValueAt(row, indexes(item)))
                Next item
                fields.Add "ValidationNote", VBA.vbNullString
                fields.Add "RowTags", VBA.vbNullString
                duplicateKey = private_Normalize(private_Value(fields, "name")) & "|" & _
                    private_Normalize(private_Value(fields, "code")) & "|" & _
                    private_Normalize(private_Value(fields, "kind")) & "|" & _
                    private_Normalize(private_Value(fields, "departure"))
                fields.Add "DuplicateKey", duplicateKey
                If Not duplicates.Exists(duplicateKey) Then duplicates.Add duplicateKey, 0
                duplicates(duplicateKey) = duplicates(duplicateKey) + 1
                rows.Add fields
            End If
        End If
    Next row
    ex_Core.fn_Diagnostic_WritePerf "EjosValidation.Filter | Rows=" & VBA.CStr(rows.Count), stageAt
    If rows.Count > 0 Then
        stageAt = VBA.Timer
        If Not private_ReadPeople(rows, service, byName, byCode, diagnostic) Then Err.Raise 5, , diagnostic
        Set people = VBA.CreateObject("Scripting.Dictionary")
        For Each item In rows
            Set fields = item
            fields.Add "PersonId", private_ResolvePerson(fields, byName, byCode)
            If VBA.Len(fields("PersonId")) > 0 Then people(fields("PersonId")) = True
        Next item
        ex_Core.fn_Diagnostic_WritePerf "EjosValidation.Identity | People=" & VBA.CStr(people.Count), stageAt
        stageAt = VBA.Timer
        If people.Count > 0 Then
            If Not ex_PEB_History.fn_ReadOrder(orderNumber, orderDate, events, diagnostic, False, people, False) Then Err.Raise 5, , diagnostic
        Else
            Set events = New Collection
        End If
        ex_Core.fn_Diagnostic_WritePerf "EjosValidation.History | Events=" & VBA.CStr(events.Count), stageAt
        stageAt = VBA.Timer
        For Each item In rows
            Set fields = item
            private_CheckRow fields, events, orderNumber, orderDate, duplicates, recorded, warnings, cancelled
        Next item
        ex_Core.fn_Diagnostic_WritePerf "EjosValidation.Compare | Rows=" & VBA.CStr(rows.Count), stageAt
    End If
    stageAt = VBA.Timer
    For Each item In rows
        Set fields = item
        index = index + 1
        Set record = New obj_LookupRecord
        If Not record.Initialize(VBA.CStr(index), fields) Then private_Fail "InitFailed"
        records.Add record
    Next item
    Set result = New obj_LookupResult
    If Not result.Initialize(records, columns, False, diagnostic, "RowTags") Then Err.Raise 5, , diagnostic
    status = private_Message("Loaded")
    status = VBA.Replace$(status, "{rows}", VBA.CStr(rows.Count))
    status = VBA.Replace$(status, "{recorded}", VBA.CStr(recorded))
    status = VBA.Replace$(status, "{warnings}", VBA.CStr(warnings))
    status = VBA.Replace$(status, "{cancelled}", VBA.CStr(cancelled))
    ex_Core.fn_Diagnostic_WritePerf "EjosValidation.Result", stageAt
    fn_TryLoad = True
Cleanup:
    If Not data Is Nothing Then data.Dispose
    If Not service Is Nothing Then service.Dispose
    If configSnapshotActive Then ex_PEB_Export.fn_EndConfigSnapshot
    ex_Core.fn_Diagnostic_WritePerf "EjosValidation.Load | Success=" & VBA.CStr(fn_TryLoad), startedAt
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
Private Function private_Query( _
    ByVal prefix As String, _
    ByVal headers As Variant, _
    ByVal service As obj_TableQueryService, _
    ByRef data As obj_DataTable, _
    ByRef diagnostic As String _
) As Boolean
    Dim source As obj_TableSource
    Dim query As obj_TableQuery
    Dim header As Variant

    On Error GoTo EH
    Set source = New obj_TableSource
    Set query = New obj_TableQuery
    If Not source.Initialize() Or Not query.Initialize() Then private_Fail "InitFailed"
    source.WorkbookPath = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config(prefix & ".path"))
    source.SheetName = ex_PEB_Export.fn_Config(prefix & ".sheet")
    source.RangeAddress = ex_PEB_Export.fn_Config(prefix & ".range")
    For Each header In headers
        query.AddColumn VBA.CStr(header)
    Next header
    ' Query the registered dependency snapshot, including for an open source file.
    private_Query = service.TryExecute(source, query, data, diagnostic)
Cleanup:
    If Not query Is Nothing Then query.Dispose
    If Not source Is Nothing Then source.Dispose
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function

Private Function private_ReadPeople( _
    ByVal rows As Collection, _
    ByVal service As obj_TableQueryService, _
    ByRef byName As Object, _
    ByRef byCode As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim neededNames As Object
    Dim neededCodes As Object
    Dim fields As Object
    Dim ids As Object
    Dim data As obj_DataTable
    Dim item As Variant
    Dim headers As Variant
    Dim values As Variant
    Dim nameColumn As Long
    Dim idColumn As Long
    Dim codeColumn As Long
    Dim row As Long
    Dim name As String
    Dim code As String
    Dim personId As String

    On Error GoTo EH
    Set byName = VBA.CreateObject("Scripting.Dictionary")
    Set byCode = VBA.CreateObject("Scripting.Dictionary")
    Set neededNames = VBA.CreateObject("Scripting.Dictionary")
    Set neededCodes = VBA.CreateObject("Scripting.Dictionary")
    For Each item In rows
        Set fields = item
        neededNames(private_Normalize(private_Value(fields, "name"))) = True
        neededCodes(private_Normalize(private_Value(fields, "code"))) = True
    Next item
    headers = VBA.Array( _
        ex_PEB_Export.fn_Config("lookup.Personnel.columns.1.header"), _
        ex_PEB_Export.fn_Config("lookup.Personnel.columns.2.header"), _
        ex_PEB_Export.fn_Config("lookup.Personnel.columns.6.header"))
    If Not private_Query("lookup.Personnel.source", headers, service, data, diagnostic) Then GoTo Cleanup
    nameColumn = ex_TableQuery.fn_HeaderIndex(data.Headers, headers(0))
    idColumn = ex_TableQuery.fn_HeaderIndex(data.Headers, headers(1))
    codeColumn = ex_TableQuery.fn_HeaderIndex(data.Headers, headers(2))
    values = data.Values
    For row = 1 To data.RowCount
        name = private_Normalize(private_Text(values(row, nameColumn)))
        code = private_Normalize(private_Text(values(row, codeColumn)))
        If (VBA.Len(name) > 0 And neededNames.Exists(name)) Or (VBA.Len(code) > 0 And neededCodes.Exists(code)) Then
            personId = private_Id(private_Text(values(row, idColumn)))
            If VBA.Len(personId) > 0 Then
                If VBA.Len(name) > 0 Then
                    If Not byName.Exists(name) Then
                        Set ids = VBA.CreateObject("Scripting.Dictionary")
                        byName.Add name, ids
                    End If
                    Set ids = byName(name)
                    If Not ids.Exists(personId) Then ids.Add personId, "|"
                    ids(personId) = ids(personId) & code & "|"
                End If
                If VBA.Len(code) > 0 Then
                    If Not byCode.Exists(code) Then
                        Set ids = VBA.CreateObject("Scripting.Dictionary")
                        byCode.Add code, ids
                    End If
                    Set ids = byCode(code)
                    ids(personId) = name
                End If
            End If
        End If
    Next row
    private_ReadPeople = True
Cleanup:
    If Not data Is Nothing Then data.Dispose
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function

Private Function private_ResolvePerson( _
    ByVal fields As Object, _
    ByVal byName As Object, _
    ByVal byCode As Object _
) As String
    Dim name As String
    Dim code As String
    Dim ids As Object
    Dim keys As Variant
    Dim personId As Variant
    Dim matchedId As String
    Dim matches As Long

    name = private_Normalize(private_Value(fields, "name"))
    code = private_Normalize(private_Value(fields, "code"))
    If VBA.Len(name) = 0 Then
        private_AddNote fields, private_Message("MissingName")
    ElseIf byName.Exists(name) Then
        Set ids = byName(name)
        If ids.Count = 1 Then
            keys = ids.Keys
            private_ResolvePerson = VBA.CStr(keys(0))
            If VBA.Len(code) = 0 Or VBA.InStr(1, ids(keys(0)), "|" & code & "|", VBA.vbBinaryCompare) = 0 Then
                private_AddNote fields, private_Message("CodeMismatch")
            End If
        Else
            If VBA.Len(code) > 0 Then
                For Each personId In ids.Keys
                    If VBA.InStr(1, ids(personId), "|" & code & "|", VBA.vbBinaryCompare) > 0 Then
                        matches = matches + 1
                        matchedId = VBA.CStr(personId)
                    End If
                Next personId
            End If
            If matches = 1 Then
                private_ResolvePerson = matchedId
            Else
                private_AddNote fields, private_Message("AmbiguousName")
            End If
        End If
    ElseIf VBA.Len(code) > 0 And byCode.Exists(code) Then
        private_AddNote fields, private_Message("NameMismatch")
    Else
        private_AddNote fields, private_Message("UnknownPerson")
    End If
End Function

Private Sub private_CheckRow( _
    ByVal fields As Object, _
    ByVal events As Collection, _
    ByVal orderNumber As String, _
    ByVal orderDate As Date, _
    ByVal duplicates As Object, _
    ByRef recorded As Long, _
    ByRef warnings As Long, _
    ByRef cancelled As Long _
)
    Dim departure As Date
    Dim arrival As Date
    Dim expected As String
    Dim storedEvent As Object
    Dim plan As Object
    Dim matched As Boolean
    Dim otherEvent As Boolean
    Dim isCancelled As Boolean
    Dim hasDeparture As Boolean
    Dim kind As String
    Dim value As String
    Dim index As Long
    Dim kindAlias As Variant

    isCancelled = (VBA.InStr(1, private_Normalize(private_Value(fields, "rank")), private_Normalize(private_Config("cancelledToken")), VBA.vbBinaryCompare) > 0 Or _
        VBA.InStr(1, private_Normalize(private_Value(fields, "destination")), private_Normalize(private_Config("cancelledToken")), VBA.vbBinaryCompare) > 0)
    kind = private_Normalize(private_Value(fields, "kind"))
    For index = 1 To ex_PEB_Export.fn_Count("validation.Ejos.kinds.count")
        For Each kindAlias In VBA.Split(private_Config("kinds." & VBA.CStr(index) & ".aliases"), "|")
            If kind = private_Normalize(VBA.CStr(kindAlias)) Then
                expected = private_Config("kinds." & VBA.CStr(index) & ".status")
                Exit For
            End If
        Next kindAlias
        If VBA.Len(expected) > 0 Then Exit For
    Next index
    If VBA.Len(expected) = 0 Then private_AddNote fields, private_Message("UnknownKind")
    hasDeparture = private_TryDate(private_Value(fields, "departure"), departure)
    If Not hasDeparture Then private_AddNote fields, private_Message("InvalidDeparture")
    value = private_Value(fields, "return")
    If VBA.Len(VBA.Trim$(value)) > 0 Then
        If Not private_TryDate(value, arrival) Then
            private_AddNote fields, private_Message("InvalidReturn")
        ElseIf hasDeparture Then
            If arrival < departure Then private_AddNote fields, private_Message("ReturnBeforeDeparture")
        End If
    End If
    value = private_Value(fields, "orderNumber")
    If VBA.Len(VBA.Trim$(value)) > 0 Then
        If private_Order(value) <> private_Order(orderNumber) Then private_AddNote fields, private_Message("OrderMismatch")
    End If
    If VBA.Len(VBA.Trim$(private_Value(fields, "destination"))) = 0 Then private_AddNote fields, private_Message("MissingDestination")
    If duplicates(fields("DuplicateKey")) > 1 Then private_AddNote fields, private_Message("Duplicate")
    For Each storedEvent In events
        If storedEvent("Status") = "Active" And storedEvent("RowType") = "Event" And VBA.Len(fields("PersonId")) > 0 Then
            If private_Id(ex_PEB_Export.fn_Text(storedEvent("PersonId"))) = fields("PersonId") And _
                    private_Order(storedEvent("AcceptedOrderNumber")) = private_Order(orderNumber) And _
                    ex_PEB_History.fn_Date(storedEvent("AcceptedOrderDate")) = orderDate Then
                If ex_PEB_MovementRegistry.fn_TryEvent(storedEvent("EventId"), plan) Then
                    If VBA.Len(plan("targetStatus")) > 0 And VBA.Len(expected) > 0 Then
                        If plan("targetStatus") = expected And hasDeparture Then
                            If ex_PEB_History.fn_Date(storedEvent("EventDate")) = departure Then
                                matched = True
                            Else
                                otherEvent = True
                            End If
                        Else
                            otherEvent = True
                        End If
                    End If
                End If
            End If
        End If
    Next storedEvent
    If isCancelled Then
        fields("RowTags") = "ejosCancelled"
        private_AddNote fields, private_Message("Cancelled")
        If matched Then private_AddNote fields, private_Message("CancelledRecorded")
        cancelled = cancelled + 1
    Else
        If Not matched And otherEvent Then private_AddNote fields, private_Message("HistoryMismatch")
        If VBA.Len(fields("ValidationNote")) > 0 Then
            fields("RowTags") = "ejosReview"
            warnings = warnings + 1
        ElseIf matched Then
            fields("RowTags") = "ejosRecorded"
        End If
    End If
    If matched Then
        recorded = recorded + 1
        private_AddNote fields, private_Message("Recorded")
    ElseIf Not isCancelled Then
        private_AddNote fields, private_Message("NotRecorded")
    End If
End Sub

Private Function private_TryDate( _
    ByVal value As Variant, _
    ByRef day As Date _
) As Boolean
    Dim text As String
    Dim dateKey As Double
    Dim d As Long
    Dim m As Long
    Dim y As Long

    If VBA.IsError(value) Or VBA.IsNull(value) Or VBA.IsEmpty(value) Then Exit Function
    If VBA.VarType(value) = VBA.vbDate Or VBA.IsNumeric(value) Then
        dateKey = VBA.Fix(VBA.CDbl(value))
        If dateKey < 1 Or dateKey > 2958465 Then Exit Function
        day = VBA.CDate(dateKey)
    Else
        text = VBA.Trim$(VBA.Replace$(VBA.CStr(value), VBA.ChrW(160), " "))
        If Not text Like "##.##.####*" Then Exit Function
        If VBA.Len(text) > 10 Then
            If VBA.Mid$(text, 11, 1) Like "#" Then Exit Function
        End If
        d = VBA.CLng(VBA.Left$(text, 2))
        m = VBA.CLng(VBA.Mid$(text, 4, 2))
        y = VBA.CLng(VBA.Mid$(text, 7, 4))
        If d < 1 Or d > 31 Or m < 1 Or m > 12 Or y < 1900 Or y > 9999 Then Exit Function
        day = VBA.DateSerial(y, m, d)
        If VBA.Day(day) <> d Or VBA.Month(day) <> m Then Exit Function
    End If
    private_TryDate = True
End Function

Private Function private_Value( _
    ByVal fields As Object, _
    ByVal key As String _
) As String
    private_Value = fields(private_Config("fields." & key))
End Function

Private Function private_Text(ByVal value As Variant) As String
    If VBA.IsError(value) Then private_Fail "InvalidCell"
    If VBA.IsNull(value) Or VBA.IsEmpty(value) Then Exit Function
    If VBA.VarType(value) = VBA.vbDate Then
        private_Text = VBA.Format$(value, "dd.mm.yyyy")
    Else
        private_Text = VBA.CStr(value)
    End If
End Function

Private Function private_Normalize(ByVal text As String) As String
    text = VBA.Replace$(text, VBA.ChrW(160), " ")
    text = VBA.Replace$(text, VBA.vbCr, " ")
    text = VBA.Replace$(text, VBA.vbLf, " ")
    text = VBA.Replace$(text, VBA.vbTab, " ")
    text = VBA.Replace$(text, VBA.ChrW(&H2019), "'")
    text = VBA.Replace$(text, VBA.ChrW(&H2018), "'")
    text = VBA.Replace$(text, VBA.ChrW(&H2BC), "'")
    Do While VBA.InStr(text, "  ") > 0
        text = VBA.Replace$(text, "  ", " ")
    Loop
    private_Normalize = VBA.LCase$(VBA.Trim$(text))
End Function

Private Function private_Id(ByVal text As String) As String
    text = VBA.Replace$(private_Normalize(text), " ", "")
    If VBA.Left$(text, 1) = "'" Then text = VBA.Mid$(text, 2)
    If text Like "##########" Then private_Id = text
End Function

Private Function private_Order(ByVal text As String) As String
    private_Order = VBA.Replace$(private_Normalize(VBA.Replace$(text, "№", "")), " ", "")
End Function

Private Sub private_AddNote( _
    ByVal fields As Object, _
    ByVal note As String _
)
    If VBA.Len(fields("ValidationNote")) > 0 Then fields("ValidationNote") = fields("ValidationNote") & VBA.vbLf
    fields("ValidationNote") = fields("ValidationNote") & note
End Sub

Private Function private_Config(ByVal key As String) As String
    private_Config = ex_PEB_Export.fn_Config("validation.Ejos." & key)
End Function

Private Function private_Message(ByVal key As String) As String
    private_Message = private_Config("text." & key)
End Function

Private Sub private_Fail(ByVal key As String)
    Err.Raise 5, , private_Message(key)
End Sub
' --------------------------------------
' } // namespace Private
' --------------------------------------