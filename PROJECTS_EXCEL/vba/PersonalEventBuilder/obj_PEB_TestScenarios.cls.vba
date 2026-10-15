VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_PEB_TestScenarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_excel As Object
Private m_book As Object
Private m_seed As Boolean
Private m_prepared As Boolean

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Properties
' //
Public Property Get TargetWorkbook() As Object
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Scenario session is not initialized."
    Set TargetWorkbook = m_book
End Property

' //
' // API
' //
Public Function Initialize(ByRef diagnostic As String) As Boolean
    Dim path As String
    Dim productionPath As String
    Dim testRoot As String
    Dim openBook As Workbook

    If m_isInitialized Or m_isDisposed Then Exit Function
    On Error GoTo EH
    If ex_PEB_Export.fn_Config("environment.name") <> "TEST" Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.TestOnly")
    End If
    path = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("test.scenarios.path"))
    productionPath = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("test.scenarios.productionPath"))
    testRoot = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("test.scenarios.root"))
    If VBA.Right$(testRoot, 1) <> "\" Then testRoot = testRoot & "\"
    If VBA.StrComp(VBA.Left$(path, VBA.Len(testRoot)), testRoot, VBA.vbTextCompare) <> 0 Then
        Err.Raise 5, , "Scenario workbook must be inside the configured test data directory."
    End If
    If VBA.StrComp(path, productionPath, VBA.vbTextCompare) = 0 Then Err.Raise 5, , "Scenario source resolves to production."
    If VBA.StrComp(path, ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("export.Movement.path")), VBA.vbTextCompare) <> 0 Then
        Err.Raise 5, , "Scenario source must equal the TEST Movement target."
    End If
    If Not VBA.CreateObject("Scripting.FileSystemObject").FileExists(path) Then Err.Raise 53, , path
    For Each openBook In Application.Workbooks
        If VBA.StrComp(openBook.FullName, path, VBA.vbTextCompare) = 0 Then
            Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.TestCloseBook"), "{path}", path)
        End If
    Next openBook

    ' Work on a private, unsaved copy until the form has been updated successfully.
    Set m_excel = VBA.CreateObject("Excel.Application")
    m_excel.Visible = False
    m_excel.DisplayAlerts = False
    m_excel.EnableEvents = False
    m_excel.AutomationSecurity = 3
    Set m_book = m_excel.Workbooks.Open(path, 0, False)
    If m_book.ReadOnly Then Err.Raise 70, , "Scenario workbook is locked or read-only: " & path
    If m_book.Date1904 Then Err.Raise 5, , "Scenario workbook must use the 1900 date system."
    m_excel.Calculation = xlCalculationManual
    m_isInitialized = True
    Initialize = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Me.Dispose
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    m_isInitialized = False
    On Error Resume Next
    If Not m_book Is Nothing Then m_book.Close False
    Set m_book = Nothing
    If Not m_excel Is Nothing Then m_excel.Quit
    Set m_excel = Nothing
    On Error GoTo 0
End Sub

Public Function TryPrepare( _
    ByVal scenarioId As String, _
    ByVal stepId As Long, _
    ByVal seedHistory As Boolean, _
    ByRef modeId As String, _
    ByRef updates As Object, _
    ByRef status As String, _
    ByRef diagnostic As String _
) As Boolean
    Dim scenarios As Object
    Dim steps As Object
    Dim history As Object
    Dim movement As Object
    Dim scenarioRow As Long
    Dim stepRow As Long
    Dim mapping As Variant
    Dim parts As Variant
    Dim field As Variant
    Dim header As String
    Dim value As Variant
    Dim defaults As Object
    Dim stepColumns As Object
    Dim column As Object
    Dim prefix As String
    Dim personId As String

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Or m_prepared Then Err.Raise 5, , "Scenario loader is not ready."
    If VBA.Len(scenarioId) = 0 Or stepId < 1 Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.TestChooseScenario")
    Set scenarios = private_Table("cases")
    Set steps = private_Table("steps")
    Set history = private_Table("history")
    Set movement = m_book.Worksheets(ex_PEB_Export.fn_Config("export.Movement.sheet")).ListObjects( _
        ex_PEB_Export.fn_Config("export.Movement.table"))
    scenarioRow = private_Find(scenarios, scenarioId, 0)
    stepRow = private_Find(steps, scenarioId, stepId)
    personId = ex_PEB_Export.fn_Text(private_Value(scenarios, scenarioRow, "personId"))
    If VBA.Len(personId) = 0 Then Err.Raise 5, , "Scenario person identity is empty."
    If ex_PEB_Export.fn_Text(private_Value(scenarios, scenarioRow, "target")) <> movement.Name Then
        Err.Raise 5, , "Scenario target differs from the configured Movement table."
    End If
    If ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "personId")) <> personId Then
        Err.Raise 5, , "Scenario and step identities differ."
    End If
    modeId = ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "mode"))
    If ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "event")) <> ex_PEB_Export.fn_Config("export.events." & modeId) Then
        Err.Raise 5, , "Step event name differs from the configured Movement event."
    End If
    prefix = ex_PEB_Export.fn_Config("modes." & modeId & ".form")
    prefix = ex_PEB_Export.fn_Config("modes." & modeId & ".direction")

    Set updates = VBA.CreateObject("Scripting.Dictionary")
    Set defaults = VBA.CreateObject("Scripting.Dictionary")
    Set stepColumns = VBA.CreateObject("Scripting.Dictionary")
    For Each column In steps.ListColumns
        stepColumns.Add column.Name, column.Index
    Next column
    For Each mapping In VBA.Split(ex_PEB_Export.fn_Config("test.scenarios.defaults"), "|")
        parts = VBA.Split(mapping, "=", 2)
        If UBound(parts) <> 1 Then Err.Raise 5, , "Invalid scenario default: " & mapping
        defaults.Add VBA.CStr(parts(0)), parts(1)
    Next mapping
    ' Clear every export input, then apply the configured step projection and explicit defaults.
    For Each field In VBA.Split(ex_PEB_Export.fn_Config("export.fields"), "|")
        If field <> "ModeId" And field <> "ModeDirection" And field <> "ModeForm" And field <> "EventType" Then
            value = VBA.vbNullString
            If defaults.Exists(field) Then value = defaults(field)
            updates("Form." & field) = private_FormText(value, VBA.CStr(field))
        End If
    Next field
    ' Optional columns allow richer fixtures without inventing values in the loader.
    For Each field In VBA.Split(ex_PEB_Export.fn_Config("test.scenarios.optionalFields"), "|")
        If Not updates.Exists("Form." & field) Then Err.Raise 5, , "Unknown optional scenario field: " & field
        If stepColumns.Exists(field) Then
            value = steps.DataBodyRange.Cells(stepRow, stepColumns(field)).Value2
            updates("Form." & field) = private_FormText(value, VBA.CStr(field))
        End If
    Next field
    For Each mapping In VBA.Split(ex_PEB_Export.fn_Config("test.scenarios.fields"), "|")
        parts = VBA.Split(mapping, "=", 2)
        If UBound(parts) <> 1 Then Err.Raise 5, , "Invalid scenario field mapping: " & mapping
        field = parts(0)
        header = parts(1)
        If Not updates.Exists("Form." & field) Then Err.Raise 5, , "Unknown scenario form field: " & field
        value = steps.DataBodyRange.Cells(stepRow, steps.ListColumns(header).Index).Value2
        updates("Form." & field) = private_FormText(value, VBA.CStr(field))
    Next mapping
    If VBA.Len(updates("Form.PersonName")) = 0 Or updates("Form.PersonId") <> personId Then
        Err.Raise 5, , "Scenario form identity is missing or inconsistent."
    End If
    For Each field In VBA.Split(ex_PEB_Export.fn_Config("export.Movement." & modeId & ".required"), "|")
        If Not updates.Exists("Form." & field) Then Err.Raise 5, , "Unknown required scenario field: " & field
        If VBA.Len(updates("Form." & field)) = 0 Then Err.Raise 5, , "Required scenario field is empty: " & field
    Next field
    updates("Form.AutoFillPersonnel") = False
    updates("Form.WordPreviewText") = VBA.vbNullString
    updates("Form.OrderQuery") = updates("Form.AcceptedOrderNumber") & " / " & updates("Form.AcceptedOrderDate")
    updates("Form.OrderYear") = VBA.Right$(updates("Form.AcceptedOrderDate"), 4)
    updates("Form.OrderChoice") = VBA.vbNullString
    status = ex_PEB_Export.fn_Config("text.TestLoaded")
    status = VBA.Replace$(status, "{scenario}", scenarioId)
    status = VBA.Replace$(status, "{step}", VBA.CStr(stepId))
    status = VBA.Replace$(status, "{expected}", ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "expected")))
    status = VBA.Replace$(status, "{code}", ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "code")))
    updates("Form.ActionStatus") = status
    updates("Form.ActionStatusStyle") = "actionSuccess"
    updates("Form.AcceptedOrderText") = updates("Form.OrderQuery")

    If seedHistory Then private_Seed history, movement, scenarioId, personId
    m_seed = seedHistory
    m_prepared = True
    TryPrepare = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
End Function

Public Function TryCommit( _
    ByRef diagnostic As String, _
    Optional ByVal persist As Boolean = True _
) As Boolean
    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Or Not m_prepared Then Err.Raise 5, , "Scenario transaction is not prepared."
    If m_seed Then
        m_excel.Calculate
        If persist Then m_book.Save
    End If
    m_prepared = False
    TryCommit = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
End Function

Public Function TryCatalog( _
    ByRef catalog As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim cases As Object
    Dim steps As Object
    Dim ids As Object
    Dim numbered As Object
    Dim item As Object
    Dim testStep As Object
    Dim ordered As Collection
    Dim rowIndex As Long
    Dim stepRow As Long
    Dim index As Long
    Dim id As String
    Dim number As String
    Dim fixtureInput As Object
    Dim inputColumn As Object

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Scenario session is not initialized."
    Set cases = private_Table("cases")
    Set steps = private_Table("steps")
    Set catalog = New Collection
    Set ids = VBA.CreateObject("Scripting.Dictionary")
    For rowIndex = 1 To cases.ListRows.Count
        id = ex_PEB_Export.fn_Text(private_Value(cases, rowIndex, "id"))
        If VBA.Len(id) = 0 Or ids.Exists(id) Then Err.Raise 5, , "Missing or duplicate scenario identity: " & id
        ids.Add id, True
        Set item = VBA.CreateObject("Scripting.Dictionary")
        item("Id") = id
        item("Name") = ex_PEB_Export.fn_Text(private_Value(cases, rowIndex, "name"))
        item("RuleStatus") = ex_PEB_Export.fn_Text(private_Value(cases, rowIndex, "ruleStatus"))
        If item("RuleStatus") <> "Defined" And item("RuleStatus") <> "PendingRule" Then
            Err.Raise 5, , "Unknown scenario rule status: " & id
        End If
        Set numbered = VBA.CreateObject("Scripting.Dictionary")
        For stepRow = 1 To steps.ListRows.Count
            If ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "id")) = id Then
                number = ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "step"))
                If Not VBA.IsNumeric(number) Then Err.Raise 5, , "Invalid step number: " & id
                If VBA.CDbl(number) < 1 Or VBA.CDbl(number) <> VBA.Fix(VBA.CDbl(number)) Then Err.Raise 5, , "Invalid step number: " & id
                number = VBA.CStr(VBA.CLng(number))
                If numbered.Exists(number) Then Err.Raise 5, , "Duplicate scenario step: " & id & "/" & number
                Set testStep = VBA.CreateObject("Scripting.Dictionary")
                testStep("Id") = VBA.CLng(number)
                testStep("ModeId") = ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "mode"))
                testStep("EventName") = ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "event"))
                testStep("PersonId") = ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "personId"))
                Set fixtureInput = VBA.CreateObject("Scripting.Dictionary")
                For Each inputColumn In steps.ListColumns
                    fixtureInput(inputColumn.Name) = steps.DataBodyRange.Cells(stepRow, inputColumn.Index).Value2
                Next inputColumn
                Set testStep("Input") = fixtureInput
                testStep("Expected") = ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "expected"))
                testStep("Code") = ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "code"))
                testStep("OpenCount") = ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "openCount"))
                Select Case testStep("Expected")
                    Case "Allow", "Block", "NoChange", "Pending"
                    Case Else
                        Err.Raise 5, , "Unknown expected test decision: " & id & "/" & number
                End Select
                If testStep("Expected") <> "Pending" Then
                    If Not VBA.IsNumeric(testStep("OpenCount")) Then Err.Raise 5, , "ExpectedOpenCount must be numeric: " & id
                    If VBA.CDbl(testStep("OpenCount")) < 0 Or VBA.CDbl(testStep("OpenCount")) <> VBA.Fix(VBA.CDbl(testStep("OpenCount"))) Then
                        Err.Raise 5, , "ExpectedOpenCount must be a nonnegative integer: " & id
                    End If
                End If
                numbered.Add number, testStep
            End If
        Next stepRow
        If numbered.Count = 0 Then Err.Raise 5, , "Scenario has no steps: " & id
        Set ordered = New Collection
        For index = 1 To numbered.Count
            If Not numbered.Exists(VBA.CStr(index)) Then Err.Raise 5, , "Scenario step sequence has a gap: " & id
            ordered.Add numbered(VBA.CStr(index))
        Next index
        Set item("Steps") = ordered
        catalog.Add item
    Next rowIndex
    If catalog.Count = 0 Then Err.Raise 5, , "Scenario catalog is empty."
    TryCatalog = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
End Function

Public Function TrySnapshot( _
    ByVal personId As String, _
    ByRef state As Object, _
    ByRef diagnostic As String, _
    Optional ByVal recalculate As Boolean = False _
) As Boolean
    Dim table As Object
    Dim column As Object
    Dim row As Object
    Dim records As New Collection
    Dim index As Long
    Dim idColumn As Long
    Dim arrivalColumn As Long
    Dim skipColumns As String
    Dim text As String
    Dim signature As String
    Dim openCount As Long
    Dim hasData As Boolean
    Dim integrity As String
    Dim rowIntegrity As String
    Dim raw As Variant
    Dim token As String
    Dim cells As Object

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Scenario session is not initialized."
    If recalculate Then m_excel.Calculate
    Set table = m_book.Worksheets(ex_PEB_Export.fn_Config("export.Movement.sheet")).ListObjects( _
        ex_PEB_Export.fn_Config("export.Movement.table"))
    idColumn = table.ListColumns(ex_PEB_Export.fn_Config("test.scenarios.headers.personId")).Index
    arrivalColumn = table.ListColumns(ex_PEB_Export.fn_Config("test.scenarios.headers.arrival")).Index
    skipColumns = "|" & ex_PEB_Export.fn_Config("test.scenarios.formulaColumns") & "|"
    For index = 1 To table.ListRows.Count
        Set row = VBA.CreateObject("Scripting.Dictionary")
        Set cells = VBA.CreateObject("Scripting.Dictionary")
        signature = VBA.vbNullString
        rowIntegrity = VBA.vbNullString
        hasData = False
        For Each column In table.ListColumns
            cells(column.Name) = table.DataBodyRange.Cells(index, column.Index).Value2
            If table.DataBodyRange.Cells(index, column.Index).HasFormula Then
                text = VBA.CStr(table.DataBodyRange.Cells(index, column.Index).FormulaR1C1)
                rowIntegrity = rowIntegrity & "F" & VBA.CStr(VBA.Len(text)) & ":" & text & "|"
            ElseIf VBA.InStr(1, skipColumns, "|" & column.Name & "|", VBA.vbBinaryCompare) > 0 Then
                raw = table.DataBodyRange.Cells(index, column.Index).Value2
                token = VBA.CStr(VBA.VarType(raw)) & ":" & private_RawText(raw)
                rowIntegrity = rowIntegrity & "P" & VBA.CStr(VBA.Len(token)) & ":" & token & "|"
            Else
                raw = table.DataBodyRange.Cells(index, column.Index).Value2
                text = private_RawText(raw)
                row(column.Name) = text
                token = VBA.CStr(VBA.VarType(raw)) & ":" & text
                signature = signature & VBA.CStr(VBA.Len(token)) & ":" & token & "|"
                If VBA.Len(text) > 0 Then hasData = True
            End If
        Next column
        integrity = integrity & VBA.CStr(VBA.Len(rowIntegrity)) & ":" & rowIntegrity & "|"
        If hasData Then
            Set row("_cells") = cells
            row("_integrity") = rowIntegrity
            row("_signature") = signature & rowIntegrity
            records.Add row
            If ex_PEB_Export.fn_Text(table.DataBodyRange.Cells(index, idColumn).Value2) = personId Then
                If VBA.Len(ex_PEB_Export.fn_Text(table.DataBodyRange.Cells(index, arrivalColumn).Value2)) = 0 Then openCount = openCount + 1
            End If
        End If
    Next index
    Set state = VBA.CreateObject("Scripting.Dictionary")
    Set state("Rows") = records
    state("OpenCount") = openCount
    state("Integrity") = integrity
    TrySnapshot = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
End Function

Public Function TryArchive( _
    ByVal report As Object, _
    ByVal before As Object, _
    ByVal after As Object, _
    ByVal fixture As Object, _
    ByVal prepared As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim results As Object
    Dim records As Object
    Dim inputs As Object
    Dim movement As Object
    Dim column As Object
    Dim row As Object
    Dim source As Object
    Dim fields As String
    Dim stage As Variant
    Dim key As Variant

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Scenario session is not initialized."
    If ex_PEB_Export.fn_Config("environment.name") <> "TEST" Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.TestOnly")
    Set results = private_ArchiveTable("results", VBA.Split(ex_PEB_Export.fn_Config("test.audit.resultFields"), "|"))
    fields = ex_PEB_Export.fn_Config("test.audit.recordFields")
    Set movement = m_book.Worksheets(ex_PEB_Export.fn_Config("export.Movement.sheet")).ListObjects( _
        ex_PEB_Export.fn_Config("export.Movement.table"))
    For Each column In movement.ListColumns
        fields = fields & "|" & column.Name
    Next column
    Set records = private_ArchiveTable("records", VBA.Split(fields, "|"))
    Set inputs = private_ArchiveTable("inputs", VBA.Split(ex_PEB_Export.fn_Config("test.audit.inputFields"), "|"))
    private_ArchiveRow results, report
    private_ArchiveState records, report, before, after, "Before"
    private_ArchiveState records, report, after, before, "After"
    For Each stage In VBA.Array("Fixture", "Prepared")
        Set source = fixture
        If stage = "Prepared" Then Set source = prepared
        If Not source Is Nothing Then
            For Each key In source.Keys
                Set row = private_ArchiveMetadata(report)
                row("InputStage") = stage
                row("Field") = key
                row("Value") = source(key)
                row("ValueType") = VBA.VarType(source(key))
                private_ArchiveRow inputs, row
            Next key
        End If
    Next stage
    m_book.Save
    TryArchive = True
    Exit Function
EH:
    diagnostic = VBA.Replace$(ex_PEB_Export.fn_Config("text.TestsArchiveFailed"), "{error}", VBA.Err.Description)
End Function

' //
' // Private
' //
Private Function private_ArchiveTable( _
    ByVal kind As String, _
    ByVal headers As Variant _
) As Object
    Dim sheet As Object
    Dim table As Object
    Dim candidate As Object
    Dim sheetName As String
    Dim tableName As String
    Dim index As Long

    sheetName = ex_PEB_Export.fn_Config("test.audit." & kind & ".sheet")
    tableName = ex_PEB_Export.fn_Config("test.audit." & kind & ".table")
    For Each candidate In m_book.Worksheets
        If candidate.Name = sheetName Then Set sheet = candidate
    Next candidate
    If sheet Is Nothing Then
        Set sheet = m_book.Worksheets.Add(, m_book.Worksheets(m_book.Worksheets.Count))
        sheet.Name = sheetName
    End If
    For Each candidate In sheet.ListObjects
        If candidate.Name = tableName Then Set table = candidate
    Next candidate
    If table Is Nothing Then
        If m_excel.WorksheetFunction.CountA(sheet.UsedRange) > 0 Then Err.Raise 5, , "Archive worksheet contains unrelated data: " & sheetName
        For index = 0 To UBound(headers)
            sheet.Cells(1, index + 1).NumberFormat = "@"
            sheet.Cells(1, index + 1).Value2 = headers(index)
        Next index
        Set table = sheet.ListObjects.Add(xlSrcRange, sheet.Range(sheet.Cells(1, 1), sheet.Cells(2, UBound(headers) + 1)), , xlYes)
        table.Name = tableName
        table.TableStyle = ex_PEB_Export.fn_Config("test.audit.tableStyle")
        table.ListRows(1).Delete
        sheet.UsedRange.Columns.ColumnWidth = ex_PEB_Export.fn_Count("test.audit.columnWidth")
    End If
    If table.ListColumns.Count <> UBound(headers) + 1 Then Err.Raise 5, , "Archive schema differs: " & tableName
    For index = 0 To UBound(headers)
        If table.ListColumns(index + 1).Name <> headers(index) Then Err.Raise 5, , "Archive column differs: " & tableName & "/" & headers(index)
    Next index
    Set private_ArchiveTable = table
End Function

Private Sub private_ArchiveRow( _
    ByVal table As Object, _
    ByVal values As Object _
)
    Dim row As Object
    Dim data() As Variant
    Dim column As Object
    Dim index As Long

    If table.ListRows.Count >= ex_PEB_Export.fn_Count("test.audit.maxRows") Then Err.Raise 5, , "Archive row limit reached: " & table.Name
    ReDim data(1 To 1, 1 To table.ListColumns.Count)
    For Each column In table.ListColumns
        index = column.Index
        If values.Exists(column.Name) Then data(1, index) = values(column.Name)
    Next column
    Set row = table.ListRows.Add
    row.Range.NumberFormat = "@"
    row.Range.Value2 = data
    row.Range.WrapText = True
    row.Range.RowHeight = ex_PEB_Export.fn_Count("test.audit.rowHeight")
    For Each column In table.ListColumns
        If VBA.InStr(1, "|" & ex_PEB_Export.fn_Config("test.audit.dateHeaders") & "|", "|" & column.Name & "|", VBA.vbBinaryCompare) > 0 Then
            row.Range.Cells(1, column.Index).NumberFormat = "dd.mm.yyyy hh:mm:ss"
        End If
    Next column
End Sub

Private Sub private_ArchiveState( _
    ByVal table As Object, _
    ByVal report As Object, _
    ByVal state As Object, _
    ByVal other As Object, _
    ByVal phase As String _
)
    Dim records As Collection
    Dim otherRows As Collection
    Dim row As Object
    Dim cells As Object
    Dim key As Variant
    Dim index As Long
    Dim count As Long
    Dim changed As Boolean

    If Not state Is Nothing Then Set records = state("Rows")
    If Not other Is Nothing Then Set otherRows = other("Rows")
    If Not records Is Nothing Then count = records.Count
    If count = 0 Then
        Set row = private_ArchiveMetadata(report)
        row("Phase") = phase
        row("StateAvailable") = Not state Is Nothing
        row("RecordIndex") = 0
        row("ChangeKind") = "Empty"
        If state Is Nothing Then row("ChangeKind") = "Unavailable"
        private_ArchiveRow table, row
        Exit Sub
    End If
    For index = 1 To count
        Set row = private_ArchiveMetadata(report)
        row("Phase") = phase
        row("StateAvailable") = True
        row("RecordIndex") = index
        changed = True
        If Not otherRows Is Nothing Then
            If otherRows.Count >= index Then changed = (records(index)("_signature") <> otherRows(index)("_signature"))
        End If
        row("ChangeKind") = "Unchanged"
        If changed Then
            row("ChangeKind") = "Modified"
            If report("Operation") = "close" Then row("ChangeKind") = "Closed"
            If otherRows Is Nothing Then
                If phase = "After" Then row("ChangeKind") = "Created"
            ElseIf otherRows.Count < index Then
                row("ChangeKind") = "Deleted"
                If phase = "After" Then row("ChangeKind") = "Created"
            End If
        End If
        Set cells = records(index)("_cells")
        For Each key In cells.Keys
            row(key) = cells(key)
        Next key
        private_ArchiveRow table, row
    Next index
End Sub

Private Function private_ArchiveMetadata(ByVal report As Object) As Object
    Dim values As Object
    Dim key As Variant

    Set values = VBA.CreateObject("Scripting.Dictionary")
    For Each key In report.Keys
        values(key) = report(key)
    Next key
    Set private_ArchiveMetadata = values
End Function

Private Function private_RawText(ByVal value As Variant) As String
    If VBA.IsError(value) Then Err.Raise 5, , "Excel error in Movement snapshot."
    If VBA.IsNull(value) Or VBA.IsEmpty(value) Then Exit Function
    private_RawText = VBA.CStr(value)
End Function

Private Function private_FormText( _
    ByVal value As Variant, _
    ByVal field As String _
) As String
    Dim text As String
    Dim parts As Variant
    Dim parsed As Date
    Dim rx As Object

    text = ex_PEB_Export.fn_Text(value)
    If VBA.Len(text) > 0 Then
        If VBA.InStr(1, "|" & ex_PEB_Export.fn_Config("test.scenarios.dateFields") & "|", "|" & field & "|", VBA.vbBinaryCompare) > 0 Then
            If VBA.IsNumeric(value) And VBA.VarType(value) <> VBA.vbString Then
                If VBA.CDbl(value) < 61 Or VBA.CDbl(value) > 2958465 Then Err.Raise 5, , "Invalid Excel date: " & field
                text = VBA.Format$(VBA.CDate(VBA.CDbl(value)), "dd.mm.yyyy")
            Else
                Set rx = VBA.CreateObject("VBScript.RegExp")
                rx.Pattern = "^\d{2}\.\d{2}\.\d{4}$"
                If Not rx.Test(text) Then Err.Raise 5, , "Full date DD.MM.YYYY required: " & field
                parts = VBA.Split(text, ".")
                If VBA.CLng(parts(2)) < 1900 Then Err.Raise 5, , "Invalid date year: " & field
                parsed = VBA.DateSerial(VBA.CInt(parts(2)), VBA.CInt(parts(1)), VBA.CInt(parts(0)))
                If VBA.Format$(parsed, "dd.mm.yyyy") <> text Then Err.Raise 5, , "Invalid calendar date: " & field
            End If
        End If
    End If
    private_FormText = text
End Function

Private Function private_Table(ByVal kind As String) As Object
    Set private_Table = m_book.Worksheets(ex_PEB_Export.fn_Config("test.scenarios." & kind & ".sheet")).ListObjects( _
        ex_PEB_Export.fn_Config("test.scenarios." & kind & ".table"))
End Function

Private Function private_Value( _
    ByVal table As Object, _
    ByVal rowIndex As Long, _
    ByVal key As String _
) As Variant
    private_Value = table.DataBodyRange.Cells(rowIndex, table.ListColumns(ex_PEB_Export.fn_Config( _
        "test.scenarios.headers." & key)).Index).Value2
    If VBA.IsError(private_Value) Then Err.Raise 5, , "Excel error in scenario column: " & key
End Function

Private Function private_Find( _
    ByVal table As Object, _
    ByVal scenarioId As String, _
    ByVal stepId As Long _
) As Long
    Dim rowIndex As Long
    Dim matches As Long
    Dim selected As Boolean

    For rowIndex = 1 To table.ListRows.Count
        selected = (ex_PEB_Export.fn_Text(private_Value(table, rowIndex, "id")) = scenarioId)
        If selected And stepId > 0 Then
            selected = (ex_PEB_Export.fn_Text(private_Value(table, rowIndex, "step")) = VBA.CStr(stepId))
        End If
        If selected Then
            matches = matches + 1
            private_Find = rowIndex
        End If
    Next rowIndex
    If matches <> 1 Then
        Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.TestStepMissing"), "{step}", scenarioId & "/" & VBA.CStr(stepId))
    End If
End Function

Private Sub private_Seed( _
    ByVal history As Object, _
    ByVal movement As Object, _
    ByVal scenarioId As String, _
    ByVal personId As String _
)
    Dim records As New Collection
    Dim column As Object
    Dim formulas As Object
    Dim formula As String
    Dim key As Variant
    Dim rowIndex As Long
    Dim outputIndex As Long
    Dim requiredRows As Long
    Dim sourceColumn As Long
    Dim value As Variant
    Dim skipColumns As String

    If movement.Parent.ProtectContents Then Err.Raise 5, , "Movement worksheet is protected."
    skipColumns = "|" & ex_PEB_Export.fn_Config("test.scenarios.formulaColumns") & "|"
    Set formulas = VBA.CreateObject("Scripting.Dictionary")
    For Each column In movement.ListColumns
        sourceColumn = history.ListColumns(column.Name).Index
        For rowIndex = 1 To movement.ListRows.Count
            If column.DataBodyRange.Cells(rowIndex, 1).HasFormula Then
                formula = column.DataBodyRange.Cells(rowIndex, 1).FormulaR1C1
                If formulas.Exists(column.Name) Then
                    If formulas(column.Name) <> formula Then Err.Raise 5, , "Mixed formulas in Movement column: " & column.Name
                Else
                    formulas.Add column.Name, formula
                End If
            End If
        Next rowIndex
        If VBA.InStr(1, skipColumns, "|" & column.Name & "|", VBA.vbBinaryCompare) > 0 Then
            If Not formulas.Exists(column.Name) Then formulas.Add column.Name, VBA.vbNullString
        End If
    Next column
    For rowIndex = 1 To history.ListRows.Count
        If ex_PEB_Export.fn_Text(private_Value(history, rowIndex, "id")) = scenarioId Then
            If ex_PEB_Export.fn_Text(private_Value(history, rowIndex, "historyTarget")) <> movement.Name Then
                Err.Raise 5, , "History target differs from Movement table."
            End If
            If ex_PEB_Export.fn_Text(private_Value(history, rowIndex, "personId")) <> personId Then
                Err.Raise 5, , "History person differs from scenario person."
            End If
            records.Add rowIndex
        End If
    Next rowIndex

    ' Keep the native table, styles and formula templates, including for an empty history.
    requiredRows = records.Count
    If requiredRows < 1 Then requiredRows = 1
    Do While movement.ListRows.Count < requiredRows
        movement.ListRows.Add
    Loop
    Do While movement.ListRows.Count > requiredRows
        movement.ListRows(movement.ListRows.Count).Delete
    Loop
    For Each column In movement.ListColumns
        If Not formulas.Exists(column.Name) Then column.DataBodyRange.ClearContents
    Next column
    For outputIndex = 1 To records.Count
        rowIndex = VBA.CLng(records(outputIndex))
        For Each column In movement.ListColumns
            If Not formulas.Exists(column.Name) Then
                value = history.DataBodyRange.Cells(rowIndex, history.ListColumns(column.Name).Index).Value2
                If VBA.IsError(value) Then Err.Raise 5, , "Excel error in history column: " & column.Name
                If VBA.VarType(value) = VBA.vbString Then
                    column.DataBodyRange.Cells(outputIndex, 1).NumberFormat = "@"
                End If
                column.DataBodyRange.Cells(outputIndex, 1).Value2 = value
            End If
        Next column
    Next outputIndex
    For Each key In formulas.Keys
        If VBA.Len(formulas(key)) > 0 Then
            movement.ListColumns(key).DataBodyRange.FormulaR1C1 = formulas(key)
        Else
            movement.ListColumns(key).DataBodyRange.ClearContents
        End If
    Next key
End Sub