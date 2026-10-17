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
Private m_scenarioBook As Object
Private m_absenceBook As Object
Private m_seed As Boolean
Private m_prepared As Boolean
Private m_orderResolver As obj_OrderResolver
Private m_fixtureTables As Object
Private m_archiveBatches As Object
Private m_fixtureRows As Object
Private m_fixtureTable As Object
Private m_fixtureRecords As Collection

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
    Set m_fixtureTables = VBA.CreateObject("Scripting.Dictionary")
    Set m_fixtureRows = VBA.CreateObject("Scripting.Dictionary")
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
    Dim started As Double
    Dim bootstrap As Object

    If m_isInitialized Or m_isDisposed Then Exit Function
    On Error GoTo EH
    If ex_PEB_Export.fn_Config("environment.name") <> "TEST" Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.TestOnly")
    End If
    path = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("export.Movement.path"))
    productionPath = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("test.scenarios.productionPath"))
    testRoot = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("test.scenarios.root"))
    If VBA.Right$(testRoot, 1) <> "\" Then testRoot = testRoot & "\"
    If VBA.StrComp(VBA.Left$(path, VBA.Len(testRoot)), testRoot, VBA.vbTextCompare) <> 0 Then
        Err.Raise 5, , "Scenario workbook must be inside the configured test data directory."
    End If
    If VBA.StrComp(path, productionPath, VBA.vbTextCompare) = 0 Then Err.Raise 5, , "Scenario source resolves to production."
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
    m_excel.ScreenUpdating = False
    m_excel.AutomationSecurity = 3
    ' Excel requires an open workbook before accepting calculation settings.
    Set bootstrap = m_excel.Workbooks.Add(xlWBATWorksheet)
    m_excel.Calculation = xlCalculationManual
    m_excel.CalculateBeforeSave = False
    ex_PEB_Testing.fn_ReportPhase "OpenMovement"
    started = VBA.Timer
    Set m_book = m_excel.Workbooks.Open(path, 0, False)
    ex_Core.fn_Diagnostic_WritePerf "Test.OpenMovement", started
    bootstrap.Close False
    Set bootstrap = Nothing
    m_excel.Calculation = xlCalculationManual
    m_excel.CalculateBeforeSave = False
    If m_book.ReadOnly Then Err.Raise 70, , "Scenario workbook is locked or read-only: " & path
    If m_book.Date1904 Then Err.Raise 5, , "Scenario workbook must use the 1900 date system."
    ex_PEB_Testing.fn_PumpUi
    path = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("test.scenarios.path"))
    If VBA.StrComp(VBA.Left$(path, VBA.Len(testRoot)), testRoot, VBA.vbTextCompare) <> 0 Then
        Err.Raise 5, , "Scenario workbook must be inside the configured test data directory."
    End If
    If VBA.StrComp(path, m_book.FullName, VBA.vbTextCompare) = 0 Then Err.Raise 5, , "Scenario and Movement workbooks must differ."
    For Each openBook In Application.Workbooks
        If VBA.StrComp(openBook.FullName, path, VBA.vbTextCompare) = 0 Then
            Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.TestCloseBook"), "{path}", path)
        End If
    Next openBook
    ex_PEB_Testing.fn_ReportPhase "OpenScenarios"
    started = VBA.Timer
    Set m_scenarioBook = m_excel.Workbooks.Open(path, 0, False)
    ex_Core.fn_Diagnostic_WritePerf "Test.OpenScenarios", started
    If m_scenarioBook.ReadOnly Then Err.Raise 70, , "Scenario workbook is locked or read-only: " & path
    If m_scenarioBook.Date1904 Then Err.Raise 5, , "Scenario workbook must use the 1900 date system."
    ex_PEB_Testing.fn_PumpUi
    path = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("lookup.Absence.source.path"))
    If VBA.StrComp(VBA.Left$(path, VBA.Len(testRoot)), testRoot, VBA.vbTextCompare) <> 0 Then
        Err.Raise 5, , "Absence workbook must be inside the configured test data directory."
    End If
    If VBA.StrComp(path, m_book.FullName, VBA.vbTextCompare) = 0 Then Err.Raise 5, , "Absence and Movement workbooks must differ."
    If VBA.StrComp(path, m_scenarioBook.FullName, VBA.vbTextCompare) = 0 Then Err.Raise 5, , "Absence and Scenario workbooks must differ."
    For Each openBook In Application.Workbooks
        If VBA.StrComp(openBook.FullName, path, VBA.vbTextCompare) = 0 Then
            Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.TestCloseBook"), "{path}", path)
        End If
    Next openBook
    ex_PEB_Testing.fn_ReportPhase "OpenAbsence"
    started = VBA.Timer
    Set m_absenceBook = m_excel.Workbooks.Open(path, 0, False)
    ex_Core.fn_Diagnostic_WritePerf "Test.OpenAbsence", started
    If m_absenceBook.ReadOnly Then Err.Raise 70, , "Absence workbook is locked or read-only: " & path
    If m_absenceBook.Date1904 Then Err.Raise 5, , "Absence workbook must use the 1900 date system."
    ex_PEB_Testing.fn_PumpUi
    m_excel.Calculation = xlCalculationManual
    started = VBA.Timer
    ex_PEB_Testing.fn_ReportPhase "Registry"
    If Not private_RegistryCurrent() Then private_SyncRegistry
    ex_Core.fn_Diagnostic_WritePerf "Test.SyncRegistry", started
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
    If Not m_orderResolver Is Nothing Then m_orderResolver.Dispose
    Set m_orderResolver = Nothing
    Set m_fixtureTables = Nothing
    Set m_fixtureRows = Nothing
    Set m_fixtureTable = Nothing
    Set m_fixtureRecords = Nothing
    Set m_archiveBatches = Nothing
    If Not m_scenarioBook Is Nothing Then m_scenarioBook.Close False
    Set m_scenarioBook = Nothing
    If Not m_absenceBook Is Nothing Then m_absenceBook.Close False
    Set m_absenceBook = Nothing
    If Not m_book Is Nothing Then m_book.Close False
    Set m_book = Nothing
    If Not m_excel Is Nothing Then m_excel.Quit
    Set m_excel = Nothing
    On Error GoTo 0
End Sub

Public Function TryCheckpoint(ByRef diagnostic As String) As Boolean
    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Scenario session is not initialized."
    If Not m_scenarioBook.Saved Then m_scenarioBook.Save
    TryCheckpoint = True
    Exit Function
EH:
    diagnostic = VBA.Replace$(ex_PEB_Export.fn_Config("text.TestsArchiveFailed"), "{error}", VBA.Err.Description)
End Function

Public Function TryPrepare( _
    ByVal scenarioId As String, _
    ByVal stepOrder As Long, _
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
    Dim prefix As String
    Dim personId As String
    Dim requiredKey As String
    Dim action As String
    Dim started As Double

    On Error GoTo EH
    started = VBA.Timer
    If Not m_isInitialized Or m_isDisposed Or m_prepared Then Err.Raise 5, , "Scenario loader is not ready."
    If VBA.Len(scenarioId) = 0 Or stepOrder < 1 Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.TestChooseScenario")
    Set scenarios = private_Table("cases")
    Set steps = private_Table("steps")
    Set history = private_Table("history")
    private_ValidateExpectationLayout
    Set movement = m_book.Worksheets(ex_PEB_Export.fn_Config("export.Movement.sheet")).ListObjects( _
        ex_PEB_Export.fn_Config("export.Movement.table"))
    scenarioRow = private_Find(scenarios, scenarioId, 0)
    stepRow = private_Find(steps, scenarioId, stepOrder)
    action = ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "action"))
    If Not ex_PEB_Testing.fn_RegistryContains("Action", action) Then Err.Raise 5, , "Unknown test action: " & action
    Select Case action
        Case "ExportMovement", "PreviewWord", "ExportEvent"
        Case Else
            Err.Raise 5, , "Test action has no executor: " & action
    End Select
    personId = ex_PEB_Export.fn_Text(private_Value(scenarios, scenarioRow, "personId"))
    If VBA.Len(personId) = 0 Then Err.Raise 5, , "Scenario person identity is empty."
    If ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "stepPersonId")) <> personId Then
        Err.Raise 5, , "Scenario and step identities differ."
    End If
    modeId = ex_PEB_Testing.fn_ResolveTestMode(ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "mode")))
    prefix = ex_PEB_EventRoutes.fn_Value(modeId, "form")
    prefix = ex_PEB_EventRoutes.fn_Value(modeId, "direction")

    Set updates = VBA.CreateObject("Scripting.Dictionary")
    ' Clear every export input, then read all scenario values from the fixture tables.
    For Each field In VBA.Split(ex_PEB_Export.fn_Config("export.fields"), "|")
        If field <> "ModeId" And field <> "ModeDirection" And field <> "ModeForm" And field <> "EventType" Then
            value = VBA.vbNullString
            updates("Form." & field) = private_FormText(value, VBA.CStr(field))
        End If
    Next field
    For Each mapping In VBA.Split(ex_PEB_Export.fn_Config("test.scenarios.fields"), "|")
        parts = VBA.Split(mapping, "=", 2)
        If UBound(parts) <> 1 Then Err.Raise 5, , "Invalid scenario field mapping: " & mapping
        field = parts(0)
        header = parts(1)
        If Not updates.Exists("Form." & field) Then Err.Raise 5, , "Unknown scenario form field: " & field
        value = private_FixtureValue(steps, stepRow, VBA.CStr(header))
        updates("Form." & field) = private_FormText(value, VBA.CStr(field))
    Next mapping
    If VBA.Len(updates("Form.PersonName")) = 0 Or updates("Form.PersonId") <> personId Then
        Err.Raise 5, , "Scenario form identity is missing or inconsistent."
    End If
    requiredKey = ex_PEB_EventRoutes.fn_Value(modeId, "movement.required")
    If action = "PreviewWord" Then
        requiredKey = ex_PEB_Export.fn_Config("export.Word." & ex_PEB_EventRoutes.fn_Value(modeId, "word.profile") & ".previewRequired")
    End If
    For Each field In VBA.Split(requiredKey, "|")
        If Not updates.Exists("Form." & field) Then Err.Raise 5, , "Unknown required scenario field: " & field
        If VBA.Len(updates("Form." & field)) = 0 Then Err.Raise 5, , "Required scenario field is empty: " & field
    Next field
    updates("Form.AutoFillPersonnel") = False
    updates("Form.WordPreviewText") = VBA.vbNullString
    private_ResolveOrder updates
    updates("Form.OrderChoice") = VBA.vbNullString
    status = ex_PEB_Export.fn_Config("text.TestLoaded")
    status = VBA.Replace$(status, "{scenario}", scenarioId)
    status = VBA.Replace$(status, "{step}", VBA.CStr(stepOrder))
    status = VBA.Replace$(status, "{expected}", ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "action")))
    status = VBA.Replace$(status, "{code}", ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "stepId")))
    updates("Form.ActionStatus") = status
    updates("Form.ActionStatusStyle") = "actionSuccess"

    If seedHistory Then
        private_SeedAbsence scenarioId, personId, VBA.CStr(updates("Form.PersonName"))
        private_Seed history, movement, scenarioId, personId
    End If
    m_seed = seedHistory
    m_prepared = True
    TryPrepare = True
    ex_Core.fn_Diagnostic_WritePerf "Test.Prepare", started
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
        private_Calculate
        If persist Then
            m_absenceBook.Save
            m_book.Save
        End If
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
    Dim inputHeaders As Variant
    Dim inputIndex As Long
    Dim allSteps As Object
    Dim stepIdentity As String
    Dim action As String
    Dim archive As Object
    Dim movement As Object
    Dim column As Object
    Dim archiveFields As String
    Dim started As Double
    Dim scenarioSteps As Object
    Dim stepRows As Collection
    Dim stepIndex As Variant

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Scenario session is not initialized."
    started = VBA.Timer
    Set cases = private_Table("cases")
    Set steps = private_Table("steps")
    inputHeaders = steps.HeaderRowRange.Value2
    Set catalog = New Collection
    Set ids = VBA.CreateObject("Scripting.Dictionary")
    Set allSteps = VBA.CreateObject("Scripting.Dictionary")
    Set scenarioSteps = VBA.CreateObject("Scripting.Dictionary")
    ' Validate every step, including rows that do not belong to a declared scenario.
    For stepRow = 1 To steps.ListRows.Count
        stepIdentity = ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "stepId"))
        If VBA.Len(stepIdentity) = 0 Or allSteps.Exists(stepIdentity) Then Err.Raise 5, , "Missing or duplicate StepId: " & stepIdentity
        allSteps.Add stepIdentity, Nothing
        id = ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "id"))
        If Not scenarioSteps.Exists(id) Then
            Set stepRows = New Collection
            scenarioSteps.Add id, stepRows
        End If
        Set stepRows = scenarioSteps(id)
        stepRows.Add stepRow
    Next stepRow
    For rowIndex = 1 To cases.ListRows.Count
        id = ex_PEB_Export.fn_Text(private_Value(cases, rowIndex, "id"))
        If VBA.Len(id) = 0 Or ids.Exists(id) Then Err.Raise 5, , "Missing or duplicate scenario identity: " & id
        ids.Add id, True
        Set item = VBA.CreateObject("Scripting.Dictionary")
        item("Id") = id
        item("Name") = ex_PEB_Export.fn_Text(private_Value(cases, rowIndex, "name"))
        item("RuleStatus") = ex_PEB_Export.fn_Text(private_Value(cases, rowIndex, "ruleStatus"))
        If Not ex_PEB_Testing.fn_RegistryContains("RuleStatus", item("RuleStatus")) Then
            Err.Raise 5, , "Unknown scenario rule status: " & id
        End If
        Set numbered = VBA.CreateObject("Scripting.Dictionary")
        If Not scenarioSteps.Exists(id) Then Err.Raise 5, , "Scenario has no steps: " & id
        Set stepRows = scenarioSteps(id)
        For Each stepIndex In stepRows
            stepRow = VBA.CLng(stepIndex)
            If ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "id")) = id Then
                number = ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "step"))
                If Not VBA.IsNumeric(number) Then Err.Raise 5, , "Invalid step number: " & id
                If VBA.CDbl(number) < 1 Or VBA.CDbl(number) <> VBA.Fix(VBA.CDbl(number)) Then Err.Raise 5, , "Invalid step number: " & id
                number = VBA.CStr(VBA.CLng(number))
                If numbered.Exists(number) Then Err.Raise 5, , "Duplicate scenario step: " & id & "/" & number
                Set testStep = VBA.CreateObject("Scripting.Dictionary")
                stepIdentity = ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "stepId"))
                testStep("Id") = stepIdentity
                testStep("Order") = VBA.CLng(number)
                testStep("ScenarioId") = id
                action = ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "action"))
                If Not ex_PEB_Testing.fn_RegistryContains("Action", action) Then Err.Raise 5, , "Unknown test action: " & stepIdentity
                Select Case action
                    Case "ExportMovement", "PreviewWord", "ExportEvent"
                    Case Else
                        Err.Raise 5, , "Test action has no executor: " & stepIdentity
                End Select
                testStep("Action") = action
                testStep("FixtureModeId") = ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "mode"))
                testStep("ModeId") = ex_PEB_Testing.fn_ResolveTestMode(testStep("FixtureModeId"))
                testStep("EventName") = ex_PEB_EventRoutes.fn_Value(testStep("ModeId"), "label")
                testStep("PersonId") = ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "stepPersonId"))
                If testStep("PersonId") <> ex_PEB_Export.fn_Text(private_Value(cases, rowIndex, "personId")) Then
                    Err.Raise 5, , "Scenario and step identities differ: " & stepIdentity
                End If
                Set fixtureInput = VBA.CreateObject("Scripting.Dictionary")
                For inputIndex = 1 To UBound(inputHeaders, 2)
                    fixtureInput(inputHeaders(1, inputIndex)) = private_FixtureValue(steps, stepRow, VBA.CStr(inputHeaders(1, inputIndex)))
                Next inputIndex
                Set testStep("Input") = fixtureInput
                testStep("Expected") = VBA.vbNullString
                testStep("Code") = VBA.vbNullString
                testStep("OpenCount") = VBA.vbNullString
                Set allSteps(stepIdentity) = testStep
                numbered.Add number, testStep
            End If
        Next stepIndex
        If numbered.Count = 0 Then Err.Raise 5, , "Scenario has no steps: " & id
        Set ordered = New Collection
        For index = 1 To numbered.Count
            If Not numbered.Exists(VBA.CStr(index)) Then Err.Raise 5, , "Scenario step sequence has a gap: " & id
            ordered.Add numbered(VBA.CStr(index))
        Next index
        Set item("Steps") = ordered
        catalog.Add item
    Next rowIndex
    For stepRow = 1 To steps.ListRows.Count
        If Not ids.Exists(ex_PEB_Export.fn_Text(private_Value(steps, stepRow, "id"))) Then
            Err.Raise 5, , "Step references an unknown scenario: " & private_Value(steps, stepRow, "stepId")
        End If
    Next stepRow
    private_LoadExpectations allSteps
    ' Validate archive schemas before any scenario mutates the working sources.
    Set archive = private_ArchiveTable("results", VBA.Split(ex_PEB_Export.fn_Config("test.audit.resultFields"), "|"))
    Set archive = private_ArchiveTable("inputs", VBA.Split(ex_PEB_Export.fn_Config("test.audit.inputFields"), "|"))
    archiveFields = ex_PEB_Export.fn_Config("test.audit.recordFields")
    Set movement = m_book.Worksheets(ex_PEB_Export.fn_Config("export.Movement.sheet")).ListObjects( _
        ex_PEB_Export.fn_Config("export.Movement.table"))
    For Each column In movement.ListColumns
        archiveFields = archiveFields & "|" & column.Name
    Next column
    Set archive = private_ArchiveTable("records", VBA.Split(archiveFields, "|"))
    If catalog.Count = 0 Then Err.Raise 5, , "Scenario catalog is empty."
    TryCatalog = True
    ex_Core.fn_Diagnostic_WritePerf "Test.Catalog", started
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
    Dim data As Variant
    Dim formulas As Variant
    Dim formulaMask As Variant
    Dim headers() As String
    Dim columnIndex As Long
    Dim started As Double

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Scenario session is not initialized."
    started = VBA.Timer
    If recalculate Then private_Calculate
    Set table = m_book.Worksheets(ex_PEB_Export.fn_Config("export.Movement.sheet")).ListObjects( _
        ex_PEB_Export.fn_Config("export.Movement.table"))
    idColumn = table.ListColumns(ex_PEB_Export.fn_Config("test.scenarios.headers.personId")).Index
    arrivalColumn = table.ListColumns(ex_PEB_Export.fn_Config("test.scenarios.headers.arrival")).Index
    skipColumns = "|" & ex_PEB_Export.fn_Config("test.scenarios.formulaColumns") & "|"
    If table.ListRows.Count > 0 Then
        data = table.DataBodyRange.Value2
        formulas = table.DataBodyRange.FormulaR1C1
        formulaMask = private_FormulaMask(table.DataBodyRange)
    End If
    ReDim headers(1 To table.ListColumns.Count)
    For Each column In table.ListColumns
        headers(column.Index) = column.Name
    Next column
    For index = 1 To table.ListRows.Count
        Set row = VBA.CreateObject("Scripting.Dictionary")
        Set cells = VBA.CreateObject("Scripting.Dictionary")
        signature = VBA.vbNullString
        rowIntegrity = VBA.vbNullString
        hasData = False
        For columnIndex = 1 To UBound(headers)
            cells(headers(columnIndex)) = data(index, columnIndex)
            If formulaMask(index, columnIndex) Then
                text = VBA.CStr(formulas(index, columnIndex))
                rowIntegrity = rowIntegrity & "F" & VBA.CStr(VBA.Len(text)) & ":" & text & "|"
            ElseIf VBA.InStr(1, skipColumns, "|" & headers(columnIndex) & "|", VBA.vbBinaryCompare) > 0 Then
                raw = data(index, columnIndex)
                token = VBA.CStr(VBA.VarType(raw)) & ":" & private_RawText(raw)
                rowIntegrity = rowIntegrity & "P" & VBA.CStr(VBA.Len(token)) & ":" & token & "|"
            Else
                raw = data(index, columnIndex)
                text = private_RawText(raw)
                row(headers(columnIndex)) = text
                token = VBA.CStr(VBA.VarType(raw)) & ":" & text
                signature = signature & VBA.CStr(VBA.Len(token)) & ":" & token & "|"
                If VBA.Len(text) > 0 Then hasData = True
            End If
        Next columnIndex
        integrity = integrity & VBA.CStr(VBA.Len(rowIntegrity)) & ":" & rowIntegrity & "|"
        If hasData Then
            Set row("_cells") = cells
            row("_integrity") = rowIntegrity
            row("_signature") = signature & rowIntegrity
            records.Add row
            If ex_PEB_Export.fn_Text(data(index, idColumn)) = personId Then
                If VBA.Len(ex_PEB_Export.fn_Text(data(index, arrivalColumn))) = 0 Then openCount = openCount + 1
            End If
        End If
    Next index
    Set state = VBA.CreateObject("Scripting.Dictionary")
    Set state("Rows") = records
    state("OpenCount") = openCount
    state("Integrity") = integrity
    TrySnapshot = True
    ex_Core.fn_Diagnostic_WritePerf "Test.Snapshot", started
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
    ByRef diagnostic As String, _
    Optional ByVal checkpoint As Boolean = True _
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
    Dim operation As String
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim errorSource As String
    Dim objectValue As Object
    Dim started As Double
    Dim saveStarted As Double

    On Error GoTo EH
    started = VBA.Timer
    operation = "ValidateSession"
    If Not m_isInitialized Or m_isDisposed Then Err.Raise 5, , "Scenario session is not initialized."
    If ex_PEB_Export.fn_Config("environment.name") <> "TEST" Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.TestOnly")
    operation = "PrepareResultsTable"
    Set results = private_ArchiveTable("results", VBA.Split(ex_PEB_Export.fn_Config("test.audit.resultFields"), "|"))
    fields = ex_PEB_Export.fn_Config("test.audit.recordFields")
    Set movement = m_book.Worksheets(ex_PEB_Export.fn_Config("export.Movement.sheet")).ListObjects( _
        ex_PEB_Export.fn_Config("export.Movement.table"))
    For Each column In movement.ListColumns
        fields = fields & "|" & column.Name
    Next column
    operation = "PrepareRecordsTable"
    Set records = private_ArchiveTable("records", VBA.Split(fields, "|"))
    operation = "PrepareInputsTable"
    Set inputs = private_ArchiveTable("inputs", VBA.Split(ex_PEB_Export.fn_Config("test.audit.inputFields"), "|"))
    operation = "WriteResults"
    Set m_archiveBatches = VBA.CreateObject("Scripting.Dictionary")
    private_ArchiveRow results, report
    operation = "WriteSnapshots"
    private_ArchiveState records, report, before, after, "Before"
    private_ArchiveState records, report, after, before, "After"
    operation = "WriteInputs"
    For Each stage In VBA.Array("Fixture", "Prepared")
        Set source = fixture
        If stage = "Prepared" Then Set source = prepared
        If Not source Is Nothing Then
            For Each key In source.Keys
                Set row = private_ArchiveMetadata(report)
                row("InputStage") = stage
                row("Field") = key
                If VBA.IsObject(source(key)) Then
                    Set objectValue = source(key)
                    If VBA.TypeName(objectValue) <> "Collection" Then Err.Raise 5, , "Unsupported archive object: " & key
                    If objectValue.Count <> 0 Then Err.Raise 5, , "Archive requires explicit serialization of nonempty collections: " & key
                    row("Value") = "Collection(0)"
                Else
                    row("Value") = source(key)
                End If
                row("ValueType") = VBA.VarType(source(key))
                private_ArchiveRow inputs, row
            Next key
        End If
    Next stage
    operation = "FlushArchive"
    private_FlushArchive results
    private_FlushArchive records
    private_FlushArchive inputs
    Set m_archiveBatches = Nothing
    saveStarted = VBA.Timer
    operation = "SaveAbsence"
    If Not m_absenceBook.Saved Then m_absenceBook.Save
    operation = "SaveMovement"
    If Not m_book.Saved Then m_book.Save
    operation = "SaveScenarios"
    If checkpoint Then
        If Not m_scenarioBook.Saved Then m_scenarioBook.Save
    End If
    ex_Core.fn_Diagnostic_WritePerf "Test.Save", saveStarted
    TryArchive = True
    ex_Core.fn_Diagnostic_WritePerf "Test.Archive", started
    Exit Function
EH:
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    errorSource = VBA.Err.Source
    ex_Core.fn_Diagnostic_WriteLog "TEST_ARCHIVE_ERROR | Stage=" & operation & " | Number=" & VBA.CStr(errorNumber) & _
        " | Source=" & errorSource & " | Description=" & errorDescription
    ex_Core.fn_Diagnostic_Flush
    diagnostic = VBA.Replace$(ex_PEB_Export.fn_Config("text.TestsArchiveFailed"), "{error}", _
        operation & " | " & VBA.CStr(errorNumber) & " | " & errorDescription)
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
    For Each candidate In m_scenarioBook.Worksheets
        If candidate.Name = sheetName Then Set sheet = candidate
    Next candidate
    If sheet Is Nothing Then
        Err.Raise 5, , "Required archive worksheet is missing: " & sheetName
    End If
    For Each candidate In sheet.ListObjects
        If candidate.Name = tableName Then Set table = candidate
    Next candidate
    If table Is Nothing Then
        Err.Raise 5, , "Required archive table is missing: " & tableName
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
    Dim batch As Collection

    If Not m_archiveBatches.Exists(table.Name) Then
        Set batch = New Collection
        m_archiveBatches.Add table.Name, batch
    End If
    Set batch = m_archiveBatches(table.Name)
    batch.Add values
End Sub

Private Sub private_FlushArchive(ByVal table As Object)
    Dim batch As Collection
    Dim values As Object
    Dim data() As Variant
    Dim headers As Variant
    Dim index As Long
    Dim columnIndex As Long
    Dim firstRow As Long
    Dim totalRows As Long
    Dim started As Double

    If Not m_archiveBatches.Exists(table.Name) Then Exit Sub
    started = VBA.Timer
    Set batch = m_archiveBatches(table.Name)
    firstRow = table.ListRows.Count + 1
    If table.ListRows.Count > 0 Then
        If m_excel.WorksheetFunction.CountA(table.ListRows(table.ListRows.Count).Range) = 0 Then firstRow = firstRow - 1
    End If
    totalRows = firstRow + batch.Count - 1
    If totalRows > ex_PEB_Export.fn_Count("test.audit.maxRows", table.Parent.Rows.Count - table.HeaderRowRange.Row) Then
        Err.Raise 5, , "Archive row limit reached: " & table.Name
    End If
    headers = table.HeaderRowRange.Value2
    ReDim data(1 To batch.Count, 1 To table.ListColumns.Count)
    For index = 1 To batch.Count
        Set values = batch(index)
        For columnIndex = 1 To UBound(data, 2)
            If values.Exists(headers(1, columnIndex)) Then data(index, columnIndex) = values(headers(1, columnIndex))
        Next columnIndex
    Next index
    If totalRows > table.ListRows.Count Then
        table.Resize table.HeaderRowRange.Resize(totalRows + 1, table.ListColumns.Count)
    End If
    table.DataBodyRange.Rows(firstRow).Resize(batch.Count, table.ListColumns.Count).Value2 = data
    ex_Core.fn_Diagnostic_WritePerf "Test.ArchiveBatch", started
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

Private Function private_RegistryCurrent() As Boolean
    Dim table As Object
    Dim target As Object
    Dim column As Object
    Dim sheet As Object
    Dim candidate As Object
    Dim kind As Variant
    Dim values As Variant
    Dim actual As Variant
    Dim binding As Variant
    Dim parts As Variant
    Dim index As Long
    Dim maximum As Long
    Dim rangeName As String
    Dim reference As String
    Dim runtimeHeader As String

    ' Any stale or missing registry component is repaired by the regular synchronizer.
    On Error GoTo Stale
    Set table = m_scenarioBook.Worksheets(ex_PEB_Export.fn_Config("test.registry.sheet")).ListObjects( _
        ex_PEB_Export.fn_Config("test.registry.table"))
    runtimeHeader = ex_PEB_Export.fn_Config("test.registry.runtimeHeader")
    For Each kind In VBA.Split(ex_PEB_Export.fn_Config("test.registry.kinds"), "|")
        values = ex_PEB_Testing.fn_RegistryValues(VBA.CStr(kind))
        If UBound(values) + 1 > maximum Then maximum = UBound(values) + 1
        Set column = table.ListColumns(VBA.CStr(kind))
        actual = column.DataBodyRange.Value2
        For index = LBound(values) To UBound(values)
            If actual(index + 1, 1) <> values(index) Then Exit Function
            If kind = "ModeId" Then
                If table.ListColumns(runtimeHeader).DataBodyRange.Cells(index + 1, 1).Value2 <> _
                    ex_PEB_Testing.fn_ResolveTestMode(values(index)) Then Exit Function
            End If
        Next index
        For index = UBound(values) + 2 To table.ListRows.Count
            If VBA.Len(ex_PEB_Export.fn_Text(actual(index, 1))) > 0 Then Exit Function
        Next index
        rangeName = ex_PEB_Export.fn_Config("test.registry." & kind & ".rangeName")
        reference = "='" & VBA.Replace$(table.Parent.Name, "'", "''") & "'!" & _
            column.DataBodyRange.Cells(1, 1).Resize(UBound(values) + 1, 1).Address
        If m_scenarioBook.Names(rangeName).RefersToRange.Address(True, True, xlA1, True) <> _
            column.DataBodyRange.Cells(1, 1).Resize(UBound(values) + 1, 1).Address(True, True, xlA1, True) Then Exit Function
    Next kind
    If table.ListRows.Count <> maximum Then Exit Function
    For Each binding In VBA.Split(ex_PEB_Export.fn_Config("test.registry.bindings"), "|")
        parts = VBA.Split(binding, ":")
        Set target = Nothing
        For Each sheet In m_scenarioBook.Worksheets
            For Each candidate In sheet.ListObjects
                If candidate.Name = parts(0) Then Set target = candidate
            Next candidate
        Next sheet
        If target Is Nothing Then Exit Function
        Set column = target.ListColumns(parts(1))
        rangeName = ex_PEB_Export.fn_Config("test.registry." & parts(2) & ".rangeName")
        If Not column.DataBodyRange Is Nothing Then
            With column.DataBodyRange.Validation
                If VBA.IsNull(.Type) Or VBA.IsNull(.Formula1) Then Exit Function
                If .Type <> xlValidateList Or .Formula1 <> "=" & rangeName Then Exit Function
                If .AlertStyle <> xlValidAlertStop Then Exit Function
                If Not .IgnoreBlank Or Not .InCellDropdown Or Not .ShowError Then Exit Function
                If .ErrorTitle <> ex_PEB_Export.fn_Config("test.registry.errorTitle") Then Exit Function
                If .ErrorMessage <> ex_PEB_Export.fn_Config("test.registry.errorMessage") Then Exit Function
            End With
        End If
    Next binding
    private_RegistryCurrent = True
Stale:
End Function

Private Sub private_SyncRegistry()
    Dim sheet As Object
    Dim table As Object
    Dim candidateSheet As Object
    Dim candidateTable As Object
    Dim column As Object
    Dim targetTable As Object
    Dim values As Variant
    Dim kind As Variant
    Dim binding As Variant
    Dim parts As Variant
    Dim index As Long
    Dim maximum As Long
    Dim rangeName As String
    Dim reference As String
    Dim runtimeHeader As String
    Dim definedName As Object
    Dim candidateName As Object

    Set sheet = m_scenarioBook.Worksheets(ex_PEB_Export.fn_Config("test.registry.sheet"))
    Set table = sheet.ListObjects(ex_PEB_Export.fn_Config("test.registry.table"))
    runtimeHeader = ex_PEB_Export.fn_Config("test.registry.runtimeHeader")
    For Each kind In VBA.Split(ex_PEB_Export.fn_Config("test.registry.kinds"), "|")
        values = ex_PEB_Testing.fn_RegistryValues(VBA.CStr(kind))
        If UBound(values) + 1 > maximum Then maximum = UBound(values) + 1
    Next kind
    Do While table.ListRows.Count < maximum
        table.ListRows.Add
    Loop
    Do While table.ListRows.Count > maximum
        table.ListRows(table.ListRows.Count).Delete
    Loop
    table.DataBodyRange.ClearContents
    For Each kind In VBA.Split(ex_PEB_Export.fn_Config("test.registry.kinds"), "|")
        values = ex_PEB_Testing.fn_RegistryValues(VBA.CStr(kind))
        Set column = table.ListColumns(VBA.CStr(kind))
        For index = LBound(values) To UBound(values)
            column.DataBodyRange.Cells(index + 1, 1).Value2 = values(index)
            If kind = "ModeId" Then
                table.ListColumns(runtimeHeader).DataBodyRange.Cells(index + 1, 1).Value2 = ex_PEB_Testing.fn_ResolveTestMode(values(index))
            End If
        Next index
        rangeName = ex_PEB_Export.fn_Config("test.registry." & kind & ".rangeName")
        reference = "='" & VBA.Replace$(sheet.Name, "'", "''") & "'!" & column.DataBodyRange.Cells(1, 1).Resize(UBound(values) + 1, 1).Address
        Set definedName = Nothing
        For Each candidateName In m_scenarioBook.Names
            If VBA.StrComp(candidateName.Name, rangeName, VBA.vbTextCompare) = 0 Then Set definedName = candidateName
        Next candidateName
        If definedName Is Nothing Then
            m_scenarioBook.Names.Add Name:=rangeName, RefersTo:=reference
        Else
            definedName.RefersTo = reference
        End If
    Next kind

    sheet.Cells(table.HeaderRowRange.Row + maximum + 2, table.HeaderRowRange.Column).Value2 = ex_PEB_Export.fn_Config("test.registry.sourceCaption")

    ' Rebind dropdowns to configuration-backed named ranges, never literal lists.
    For Each binding In VBA.Split(ex_PEB_Export.fn_Config("test.registry.bindings"), "|")
        parts = VBA.Split(binding, ":")
        If UBound(parts) <> 2 Then Err.Raise 5, , "Invalid registry binding: " & binding
        Set targetTable = Nothing
        For Each candidateSheet In m_scenarioBook.Worksheets
            For Each candidateTable In candidateSheet.ListObjects
                If candidateTable.Name = parts(0) Then Set targetTable = candidateTable
            Next candidateTable
        Next candidateSheet
        If targetTable Is Nothing Then Err.Raise 5, , "Registry target table is missing: " & parts(0)
        rangeName = ex_PEB_Export.fn_Config("test.registry." & parts(2) & ".rangeName")
        Set column = targetTable.ListColumns(parts(1))
        If Not column.DataBodyRange Is Nothing Then
            With column.DataBodyRange.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & rangeName
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowError = True
                .ErrorTitle = ex_PEB_Export.fn_Config("test.registry.errorTitle")
                .ErrorMessage = ex_PEB_Export.fn_Config("test.registry.errorMessage")
            End With
        End If
    Next binding
End Sub

Private Sub private_ResolveOrder(ByVal updates As Object)
    Dim matches As Collection
    Dim record As obj_LookupRecord
    Dim number As Variant
    Dim orderDate As Variant
    Dim query As String
    Dim year As String
    Dim diagnostic As String
    Dim caption As String

    If m_orderResolver Is Nothing Then
        Set m_orderResolver = New obj_OrderResolver
        If Not m_orderResolver.Initialize("PersonalEventBuilder::orders", "PersonalEventBuilder::text", diagnostic) Then
            Set m_orderResolver = Nothing
            Err.Raise 5, , diagnostic
        End If
    End If
    query = updates("Form.AcceptedOrderNumber") & " / " & updates("Form.AcceptedOrderDate")
    year = VBA.Right$(updates("Form.AcceptedOrderDate"), 4)
    If Not m_orderResolver.TrySearch(query, year, matches, diagnostic) Then Err.Raise 5, , diagnostic
    If matches.Count = 0 Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.OrderEmpty")
    If matches.Count <> 1 Then
        diagnostic = VBA.Replace$(ex_PEB_Export.fn_Config("text.OrderFoundMany"), "{count}", VBA.CStr(matches.Count))
        Err.Raise 5, , diagnostic
    End If
    Set record = matches(1)
    If Not record.TryGetValue("OrderNumber", number) Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.OrderSourceError")
    If Not record.TryGetValue("OrderDate", orderDate) Then Err.Raise 5, , ex_PEB_Export.fn_Config("text.OrderSourceError")
    updates("Form.AcceptedOrderNumber") = VBA.CStr(number)
    updates("Form.AcceptedOrderDate") = VBA.Format$(VBA.CDate(orderDate), "dd.mm.yyyy")
    updates("Form.OrderYear") = VBA.CStr(VBA.Year(VBA.CDate(orderDate)))
    updates("Form.OrderQuery") = updates("Form.AcceptedOrderNumber") & " / " & updates("Form.AcceptedOrderDate")
    caption = ex_PEB_Export.fn_Config("text.OrderAccepted")
    caption = VBA.Replace$(caption, "{number}", VBA.CStr(number))
    updates("Form.AcceptedOrderText") = VBA.Replace$(caption, "{date}", updates("Form.AcceptedOrderDate"))
End Sub

Private Sub private_LoadExpectations(ByVal allSteps As Object)
    Dim table As Object
    Dim movement As Object
    Dim column As Object
    Dim testStep As Object
    Dim wordExpectation As Object
    Dim expectedRow As Object
    Dim rows As Collection
    Dim seen As Object
    Dim records As Object
    Dim key As Variant
    Dim header As Variant
    Dim rowIndex As Long
    Dim stepId As String
    Dim recordId As String
    Dim outcome As String
    Dim check As String
    Dim count As Variant
    Dim skipColumns As String

    private_ValidateExpectationLayout
    Set movement = m_book.Worksheets(ex_PEB_Export.fn_Config("export.Movement.sheet")).ListObjects( _
        ex_PEB_Export.fn_Config("export.Movement.table"))
    Set seen = VBA.CreateObject("Scripting.Dictionary")
    Set table = private_Table("expectedMovement")
    private_RequireHeaders table, "test.expected.movementHeaders"
    For rowIndex = 1 To table.ListRows.Count
        stepId = ex_PEB_Export.fn_Text(private_Value(table, rowIndex, "stepId"))
        If Not allSteps.Exists(stepId) Then Err.Raise 5, , "Movement expectation references an unknown StepId: " & stepId
        If seen.Exists(stepId) Then Err.Raise 5, , "Duplicate Movement expectation: " & stepId
        seen.Add stepId, True
        Set testStep = allSteps(stepId)
        If testStep("Action") <> "ExportMovement" And testStep("Action") <> "ExportEvent" Then
            Err.Raise 5, , "Movement expectation has an incompatible action: " & stepId
        End If
        outcome = ex_PEB_Export.fn_Text(private_Value(table, rowIndex, "expected"))
        If Not ex_PEB_Testing.fn_RegistryContains("ExpectedDecision", outcome) Then Err.Raise 5, , "Unknown expected Movement decision: " & stepId
        check = ex_PEB_Export.fn_Text(private_Value(table, rowIndex, "stateCheck"))
        If Not ex_PEB_Testing.fn_RegistryContains("StateCheck", check) Then Err.Raise 5, , "Unknown state check: " & stepId
        If outcome = "Allow" Then
            If check <> "Snapshot" Then Err.Raise 5, , "Allowed export requires Snapshot: " & stepId
        ElseIf outcome = "Pending" Then
            If check <> "Pending" Then Err.Raise 5, , "Pending export requires Pending state check: " & stepId
        Else
            If check <> "Unchanged" Then Err.Raise 5, , "Rejected export requires Unchanged: " & stepId
        End If
        count = private_Value(table, rowIndex, "openCount")
        If outcome <> "Pending" Then
            If Not VBA.IsNumeric(count) Then Err.Raise 5, , "ExpectedOpenCount must be numeric: " & stepId
            If VBA.CDbl(count) < 0 Or VBA.CDbl(count) <> VBA.Fix(VBA.CDbl(count)) Then
                Err.Raise 5, , "ExpectedOpenCount must be a nonnegative integer: " & stepId
            End If
        End If
        testStep("Expected") = outcome
        testStep("Code") = ex_PEB_Export.fn_Text(private_Value(table, rowIndex, "code"))
        If Not ex_PEB_Testing.fn_RegistryContains("ExpectedCode", testStep("Code")) Then Err.Raise 5, , "Unknown expected Movement code: " & stepId
        testStep("OpenCount") = count
        testStep("StateCheck") = check
        Set rows = New Collection
        Set testStep("ExpectedRows") = rows
    Next rowIndex

    ' Expected rows contain all writable native columns, with explicit blank cells.
    Set records = VBA.CreateObject("Scripting.Dictionary")
    Set table = private_Table("expectedRows")
    table.Parent.Calculate
    skipColumns = "|" & ex_PEB_Export.fn_Config("test.scenarios.formulaColumns") & "|"
    rowIndex = table.ListColumns(ex_PEB_Export.fn_Config("test.scenarios.headers.stepId")).Index
    rowIndex = table.ListColumns(ex_PEB_Export.fn_Config("test.scenarios.headers.recordId")).Index
    For Each column In movement.ListColumns
        If VBA.InStr(1, skipColumns, "|" & column.Name & "|", VBA.vbBinaryCompare) = 0 Then
            rowIndex = table.ListColumns(column.Name).Index
        End If
    Next column
    For Each header In VBA.Split(ex_PEB_Export.fn_Config("test.expected.rowKeys"), "|")
        rowIndex = table.ListColumns(header).Index
    Next header
    For rowIndex = 1 To table.ListRows.Count
        If m_excel.WorksheetFunction.CountA(table.ListRows(rowIndex).Range) > 0 Then
            stepId = ex_PEB_Export.fn_Text(private_Value(table, rowIndex, "stepId"))
            If Not seen.Exists(stepId) Then Err.Raise 5, , "Expected row references an unknown Movement StepId: " & stepId
            Set testStep = allSteps(stepId)
            If testStep("StateCheck") <> "Snapshot" Then Err.Raise 5, , "Expected rows require Snapshot: " & stepId
            recordId = ex_PEB_Export.fn_Text(private_Value(table, rowIndex, "recordId"))
            If VBA.Len(recordId) = 0 Then Err.Raise 5, , "Expected RecordId is missing: " & stepId
            key = stepId & VBA.vbNullChar & "RecordId:" & recordId
            If records.Exists(key) Then Err.Raise 5, , "Duplicate expected RecordId: " & stepId & "/" & recordId
            records.Add key, True
            Set expectedRow = VBA.CreateObject("Scripting.Dictionary")
            expectedRow("RecordId") = recordId
            For Each column In movement.ListColumns
                If VBA.InStr(1, skipColumns, "|" & column.Name & "|", VBA.vbBinaryCompare) = 0 Then
                    expectedRow(column.Name) = private_FixtureValue(table, rowIndex, column.Name)
                    If VBA.IsError(expectedRow(column.Name)) Then Err.Raise 5, , "Excel error in expected row: " & stepId & "/" & column.Name
                End If
            Next column
            If ex_PEB_Export.fn_Text(expectedRow(ex_PEB_Export.fn_Config("test.scenarios.headers.personId"))) <> testStep("PersonId") Then
                Err.Raise 5, , "Expected row person differs from scenario person: " & stepId
            End If
            key = stepId & VBA.vbNullChar & "Identity:" & private_ExpectedIdentity(expectedRow)
            If records.Exists(key) Then Err.Raise 5, , "Duplicate expected row identity: " & stepId & "/" & recordId
            records.Add key, True
            Set rows = testStep("ExpectedRows")
            rows.Add expectedRow
        End If
    Next rowIndex
    For Each key In allSteps.Keys
        Set testStep = allSteps(key)
        If testStep("Action") = "ExportMovement" Or testStep("Action") = "ExportEvent" Then
            If Not seen.Exists(key) Then Err.Raise 5, , "Movement expectation is missing: " & key
            If testStep("StateCheck") = "Snapshot" Then
                Set rows = testStep("ExpectedRows")
                If rows.Count = 0 Then Err.Raise 5, , "Expected snapshot is empty: " & key
            End If
        End If
    Next key

    Set seen = VBA.CreateObject("Scripting.Dictionary")
    Set table = private_Table("expectedWord")
    private_RequireHeaders table, "test.expected.wordHeaders"
    For rowIndex = 1 To table.ListRows.Count
        stepId = ex_PEB_Export.fn_Text(private_Value(table, rowIndex, "stepId"))
        If Not allSteps.Exists(stepId) Then Err.Raise 5, , "Word expectation references an unknown StepId: " & stepId
        If seen.Exists(stepId) Then Err.Raise 5, , "Duplicate Word expectation: " & stepId
        seen.Add stepId, True
        Set testStep = allSteps(stepId)
        If testStep("Action") <> "PreviewWord" And testStep("Action") <> "ExportEvent" Then
            Err.Raise 5, , "Word expectation has an incompatible action: " & stepId
        End If
        outcome = ex_PEB_Export.fn_Text(private_Value(table, rowIndex, "wordOutcome"))
        If Not ex_PEB_Testing.fn_RegistryContains("Outcome", outcome) Then Err.Raise 5, , "Unknown Word outcome: " & stepId
        Set wordExpectation = testStep
        If testStep("Action") = "ExportEvent" Then
            Set wordExpectation = VBA.CreateObject("Scripting.Dictionary")
            wordExpectation("Id") = stepId
            testStep.Add "WordExpectation", wordExpectation
        End If
        wordExpectation("Expected") = outcome
        If outcome = "Error" And testStep("Action") = "PreviewWord" Then
            testStep("Code") = "PreviewError"
        End If
        wordExpectation("ExpectedText") = private_Value(table, rowIndex, "wordText")
        wordExpectation("ExpectedError") = private_Value(table, rowIndex, "wordError")
        If outcome = "Success" Then
            If VBA.Len(wordExpectation("ExpectedText")) = 0 Or VBA.Len(wordExpectation("ExpectedError")) > 0 Then
                Err.Raise 5, , "Successful Word expectation requires text and no error: " & stepId
            End If
        ElseIf outcome = "Error" Then
            If VBA.Len(wordExpectation("ExpectedError")) = 0 Or VBA.Len(wordExpectation("ExpectedText")) > 0 Then
                Err.Raise 5, , "Word error expectation requires an exact diagnostic and no text: " & stepId
            End If
        ElseIf outcome = "Omitted" Then
            If VBA.Len(wordExpectation("ExpectedText")) > 0 Or VBA.Len(wordExpectation("ExpectedError")) > 0 Then
                Err.Raise 5, , "Omitted Word expectation must contain neither text nor error: " & stepId
            End If
        End If
    Next rowIndex
    For Each key In allSteps.Keys
        Set testStep = allSteps(key)
        If (testStep("Action") = "PreviewWord" Or testStep("Action") = "ExportEvent") And Not seen.Exists(key) Then
            Err.Raise 5, , "Word expectation is missing: " & key
        End If
    Next key
End Sub

Private Sub private_ValidateExpectationLayout()
    Dim upper As Object
    Dim lower As Object
    Dim minimumGap As Long

    Set upper = private_Table("expectedMovement")
    Set lower = private_Table("expectedRows")
    minimumGap = ex_PEB_Export.fn_Count("test.expected.minimumGapRows")
    If minimumGap < 2 Then
        Err.Raise 5, , "Expected-row caption requires at least two reserved rows."
    End If
    If upper.Parent.Name = lower.Parent.Name Then
        If lower.Range.Row - upper.Range.Row - upper.Range.Rows.Count < minimumGap Then
            Err.Raise 5, , "Expected tables overlap or leave no caption space: " & upper.Name & "/" & lower.Name
        End If
    End If
End Sub

Private Sub private_RequireHeaders( _
    ByVal table As Object, _
    ByVal configKey As String _
)
    Dim header As Variant
    Dim index As Long

    For Each header In VBA.Split(ex_PEB_Export.fn_Config(configKey), "|")
        index = table.ListColumns(header).Index
    Next header
End Sub

Private Function private_ExpectedIdentity(ByVal row As Object) As String
    Dim header As Variant
    Dim text As String

    For Each header In VBA.Split(ex_PEB_Export.fn_Config("test.expected.rowKeys"), "|")
        text = private_RawText(row(header))
        private_ExpectedIdentity = private_ExpectedIdentity & VBA.CStr(VBA.Len(text)) & ":" & text & "|"
    Next header
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
    Set private_Table = m_scenarioBook.Worksheets(ex_PEB_Export.fn_Config("test.scenarios." & kind & ".sheet")).ListObjects( _
        ex_PEB_Export.fn_Config("test.scenarios." & kind & ".table"))
End Function

Private Function private_Value( _
    ByVal table As Object, _
    ByVal rowIndex As Long, _
    ByVal key As String _
) As Variant
    private_Value = private_FixtureValue(table, rowIndex, ex_PEB_Export.fn_Config("test.scenarios.headers." & key))
    If VBA.IsError(private_Value) Then Err.Raise 5, , "Excel error in scenario column: " & key
End Function

Private Function private_FixtureValue( _
    ByVal table As Object, _
    ByVal rowIndex As Long, _
    ByVal header As String _
) As Variant
    Dim data As Variant
    Dim headers As Variant
    Dim rows As Collection
    Dim record As Object
    Dim index As Long
    Dim columnIndex As Long
    Dim tableName As String

    ' Reuse the current table cache without cross-process Name calls per field.
    If table Is m_fixtureTable Then
        Set record = m_fixtureRecords(rowIndex)
        If Not record.Exists(header) Then Err.Raise 5, , "Missing fixture column: " & table.Name & "/" & header
        private_FixtureValue = record(header)
        Exit Function
    End If
    tableName = table.Name
    If Not m_fixtureTables.Exists(tableName) Then
        data = table.DataBodyRange.Value2
        headers = table.HeaderRowRange.Value2
        Set rows = New Collection
        For index = 1 To UBound(data, 1)
            Set record = VBA.CreateObject("Scripting.Dictionary")
            record.CompareMode = VBA.vbTextCompare
            For columnIndex = 1 To UBound(data, 2)
                record.Add headers(1, columnIndex), data(index, columnIndex)
            Next columnIndex
            rows.Add record
        Next index
        m_fixtureTables.Add tableName, rows
    End If
    Set rows = m_fixtureTables(tableName)
    Set m_fixtureTable = table
    Set m_fixtureRecords = rows
    Set record = rows(rowIndex)
    If Not record.Exists(header) Then Err.Raise 5, , "Missing fixture column: " & table.Name & "/" & header
    private_FixtureValue = record(header)
End Function

Private Sub private_Calculate()
    Dim started As Double

    started = VBA.Timer
    ' Preserve cross-sheet dependencies in the isolated Excel instance.
    m_excel.Calculate
    ex_Core.fn_Diagnostic_WritePerf "Test.Calculate", started
End Sub

Private Function private_FormulaMask(ByVal source As Object) As Variant
    Dim mask() As Boolean
    Dim formulas As Object
    Dim area As Object
    Dim firstRow As Long
    Dim firstColumn As Long
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim errorNumber As Long
    Dim errorDescription As String

    ReDim mask(1 To source.Rows.Count, 1 To source.Columns.Count)
    ' SpecialCells distinguishes formulas from literal strings beginning with "=".
    On Error Resume Next
    Set formulas = source.SpecialCells(xlCellTypeFormulas)
    errorNumber = VBA.Err.Number
    errorDescription = VBA.Err.Description
    On Error GoTo 0
    If errorNumber <> 0 And errorNumber <> 1004 Then Err.Raise errorNumber, , errorDescription
    If Not formulas Is Nothing Then
        For Each area In formulas.Areas
            firstRow = area.Row - source.Row + 1
            firstColumn = area.Column - source.Column + 1
            lastRow = firstRow + area.Rows.Count - 1
            lastColumn = firstColumn + area.Columns.Count - 1
            For rowIndex = firstRow To lastRow
                For columnIndex = firstColumn To lastColumn
                    mask(rowIndex, columnIndex) = True
                Next columnIndex
            Next rowIndex
        Next area
    End If
    private_FormulaMask = mask
End Function

Private Function private_Find( _
    ByVal table As Object, _
    ByVal scenarioId As String, _
    ByVal stepOrder As Long _
) As Long
    Dim rowIndex As Long
    Dim matches As Long
    Dim selected As Boolean
    Dim identity As String

    identity = table.Name & "/" & VBA.CStr(VBA.Len(scenarioId)) & ":" & scenarioId & "/" & VBA.CStr(stepOrder)
    If m_fixtureRows.Exists(identity) Then
        private_Find = m_fixtureRows(identity)
        Exit Function
    End If

    For rowIndex = 1 To table.ListRows.Count
        selected = (ex_PEB_Export.fn_Text(private_Value(table, rowIndex, "id")) = scenarioId)
        If selected And stepOrder > 0 Then
            selected = (ex_PEB_Export.fn_Text(private_Value(table, rowIndex, "step")) = VBA.CStr(stepOrder))
        End If
        If selected Then
            matches = matches + 1
            private_Find = rowIndex
        End If
    Next rowIndex
    If matches <> 1 Then
        Err.Raise 5, , VBA.Replace$(ex_PEB_Export.fn_Config("text.TestStepMissing"), "{step}", scenarioId & "/" & VBA.CStr(stepOrder))
    End If
    m_fixtureRows.Add identity, private_Find
End Function

Private Sub private_SeedAbsence( _
    ByVal scenarioId As String, _
    ByVal personId As String, _
    ByVal personName As String _
)
    Dim source As Object
    Dim target As Object
    Dim records As New Collection
    Dim record As Object
    Dim output As Object
    Dim mapping As Variant
    Dim parts As Variant
    Dim field As Variant
    Dim rowIndex As Long
    Dim matched As Boolean
    Dim nameHeader As String
    Dim hasEvent As String
    Dim value As Variant

    Set source = private_Table("absence")
    Set target = m_absenceBook.Worksheets(ex_PEB_Export.fn_Config("lookup.Absence.source.sheet")).ListObjects( _
        ex_PEB_Export.fn_Config("test.absence.table"))
    nameHeader = ex_PEB_Export.fn_Config("test.absence.nameHeader")
    If target.Parent.ProtectContents Then Err.Raise 5, , "Absence worksheet is protected."

    ' Validate and stage the entire scenario before changing the target table.
    For rowIndex = 1 To source.ListRows.Count
        If ex_PEB_Export.fn_Text(private_Value(source, rowIndex, "id")) = scenarioId Then
            matched = True
            If ex_PEB_Export.fn_Text(private_Value(source, rowIndex, "personId")) <> personId Then
                Err.Raise 5, , "Absence fixture person differs from scenario person."
            End If
            If ex_PEB_Export.fn_Text(source.DataBodyRange.Cells(rowIndex, source.ListColumns(nameHeader).Index).Value2) <> personName Then
                Err.Raise 5, , "Absence fixture name differs from scenario name."
            End If
            hasEvent = ex_PEB_Export.fn_Text(source.DataBodyRange.Cells(rowIndex, source.ListColumns( _
                ex_PEB_Export.fn_Config("test.absence.eventFlagHeader")).Index).Value2)
            If hasEvent <> "0" And hasEvent <> "1" Then Err.Raise 5, , "Absence fixture event flag must be 0 or 1."
            If hasEvent = "1" Then
                Set record = VBA.CreateObject("Scripting.Dictionary")
                For Each mapping In VBA.Split(ex_PEB_Export.fn_Config("test.absence.fields"), "|")
                    parts = VBA.Split(mapping, "=", 2)
                    If UBound(parts) <> 1 Then Err.Raise 5, , "Invalid Absence fixture mapping: " & mapping
                    value = source.DataBodyRange.Cells(rowIndex, source.ListColumns(parts(1)).Index).Value2
                    If VBA.IsError(value) Then Err.Raise 5, , "Excel error in Absence fixture: " & parts(1)
                    record.Add VBA.CStr(target.ListColumns(parts(0)).Index), value
                Next mapping
                records.Add record
            End If
        End If
    Next rowIndex
    If Not matched Then Err.Raise 5, , "Absence fixture is missing for scenario: " & scenarioId

    ' The source has no IPN column; replace only rows for the validated scenario name.
    For rowIndex = target.ListRows.Count To 1 Step -1
        If ex_PEB_Export.fn_Text(target.DataBodyRange.Cells(rowIndex, target.ListColumns(nameHeader).Index).Value2) = personName Then
            target.ListRows(rowIndex).Delete
        End If
    Next rowIndex
    For Each record In records
        Set output = target.ListRows.Add
        For Each field In record.Keys
            value = record(field)
            If VBA.VarType(value) = VBA.vbString Then output.Range.Cells(1, VBA.CLng(field)).NumberFormat = "@"
            output.Range.Cells(1, VBA.CLng(field)).Value2 = value
        Next field
    Next record
End Sub

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
    Dim mask As Variant
    Dim existingFormulas As Variant
    Dim data() As Variant
    Dim started As Double

    started = VBA.Timer
    If movement.Parent.ProtectContents Then Err.Raise 5, , "Movement worksheet is protected."
    skipColumns = "|" & ex_PEB_Export.fn_Config("test.scenarios.formulaColumns") & "|"
    Set formulas = VBA.CreateObject("Scripting.Dictionary")
    If movement.ListRows.Count > 0 Then
        mask = private_FormulaMask(movement.DataBodyRange)
        existingFormulas = movement.DataBodyRange.FormulaR1C1
    End If
    For Each column In movement.ListColumns
        sourceColumn = history.ListColumns(column.Name).Index
        For rowIndex = 1 To movement.ListRows.Count
            If mask(rowIndex, column.Index) Then
                formula = existingFormulas(rowIndex, column.Index)
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
    For Each column In movement.ListColumns
        If Not formulas.Exists(column.Name) And records.Count > 0 Then
            ReDim data(1 To records.Count, 1 To 1)
            For outputIndex = 1 To records.Count
                rowIndex = VBA.CLng(records(outputIndex))
                value = private_FixtureValue(history, rowIndex, column.Name)
                If VBA.IsError(value) Then Err.Raise 5, , "Excel error in history column: " & column.Name
                If VBA.VarType(value) = VBA.vbString Then
                    column.DataBodyRange.Cells(outputIndex, 1).NumberFormat = "@"
                End If
                data(outputIndex, 1) = value
            Next outputIndex
            column.DataBodyRange.Resize(records.Count, 1).Value2 = data
        End If
    Next column
    For Each key In formulas.Keys
        If VBA.Len(formulas(key)) > 0 Then
            movement.ListColumns(key).DataBodyRange.FormulaR1C1 = formulas(key)
        Else
            movement.ListColumns(key).DataBodyRange.ClearContents
        End If
    Next key
    ex_Core.fn_Diagnostic_WritePerf "Test.SeedMovement", started
End Sub