Attribute VB_Name = "ex_PEB_GroupingTests"
Option Explicit

' --------------------------------------
' namespace API {
' --------------------------------------
Public Sub fn_Run()
    Dim diagnostic As String
    Dim count As Long
    Dim text As String

    If private_Run(count, diagnostic) Then
        text = VBA.Replace$(ex_PEB_Export.fn_Config("text.TestsGroupingPassed"), "{count}", VBA.CStr(count))
        ex_WindowsUi.fn_ShowMessage text, VBA.vbInformation, ex_PEB_Export.fn_Config("text.TestTitle")
    Else
        text = VBA.Replace$(ex_PEB_Export.fn_Config("text.TestsGroupingFailed"), "{error}", diagnostic)
        ex_WindowsUi.fn_ShowMessage text, VBA.vbExclamation, ex_PEB_Export.fn_Config("text.TestTitle")
    End If
End Sub

Public Function fn_ReadCatalog( _
    ByRef catalog As Collection, _
    ByRef diagnostic As String _
) As Boolean
    fn_ReadCatalog = ex_PEB_Testing.fn_ReadCatalog(catalog, diagnostic, "Grouping")
End Function

Public Function fn_TryRunCase( _
    ByVal scenarioId As String, _
    ByVal runId As String, _
    ByRef preview As String, _
    ByRef diagnostic As String, _
    Optional ByVal saveReport As Boolean = True, _
    Optional ByRef reportPath As String = "" _
) As Boolean
    Dim count As Long
    Dim reportId As String

    On Error GoTo EH
    reportPath = VBA.vbNullString
    preview = VBA.vbNullString
    If saveReport Then
        reportId = "Grouping_" & runId & "_" & scenarioId
    End If
    If Not private_Run(count, diagnostic, scenarioId, reportId, preview, reportPath) Then
        Exit Function
    End If
    fn_TryRunCase = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_Run( _
    ByRef count As Long, _
    ByRef diagnostic As String, _
    Optional ByVal selectedId As String = "", _
    Optional ByVal reportId As String = "", _
    Optional ByRef preview As String = "", _
    Optional ByRef reportPath As String = "" _
) As Boolean
    Dim fixture As Collection
    Dim scenario As Object
    Dim step As Object
    Dim field As Variant
    Dim expected As Object
    Dim word As Object
    Dim document As Object
    Dim values As Object
    Dim anchors As obj_PEB_WordAnchors
    Dim templates As obj_PEB_WordTemplates
    Dim target As Object
    Dim state As Object
    Dim parts As Object
    Dim files As Object
    Dim path As String
    Dim rendered As String
    Dim failure As String
    Dim before As String
    Dim beforeState As String
    Dim identifier As String
    Dim body As String
    Dim precedingText As String
    Dim followingText As String
    Dim success As Boolean
    Dim caseId As String
    Dim reportSession As obj_PEB_TestWordSession
    Dim report As Object
    Dim personNames As Object
    Dim stepIndex As Long
    Dim caption As String
    Dim templatePath As String
    Dim expectedFont As String
    Dim expectedSize As Single

    On Error GoTo EH
    If ex_PEB_Export.fn_Config("environment.name") <> "TEST" Then
        Err.Raise 5, , ex_PEB_Export.fn_Config("text.TestOnly")
    End If
    If Not private_ReadScenarios(fixture, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    Set templates = New obj_PEB_WordTemplates
    If Not templates.Initialize(diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    Set files = VBA.CreateObject("Scripting.FileSystemObject")
    Set word = VBA.CreateObject("Word.Application")
    word.Visible = False
    word.ScreenUpdating = True
    word.DisplayAlerts = 0
    word.AutomationSecurity = 3
    identifier = ex_PEB_Export.fn_Config("export.Word.textAnchors.ToBusinessTrip")
    templatePath = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("test.word.template"))
    If Not files.FileExists(templatePath) Then
        Err.Raise 5, , "Grouping Word template not found: " & templatePath
    End If
    For Each scenario In fixture
        caseId = scenario("Id")
        If VBA.Len(selectedId) > 0 And selectedId <> caseId Then
            GoTo NextScenario
        End If
        Set document = word.Documents.Add(templatePath, False, 0, False)
        private_CheckTemplate document
        path = files.BuildPath(files.GetSpecialFolder(2), "PEB_Grouping_" & files.GetTempName & ".docx")
        document.SaveAs2 path, 16
        document.Close 0
        Set document = word.Documents.Open(path, False, False, False)
        body = private_DocumentText(document)
        precedingText = private_OutsideText(body, "{\export:" & identifier & "_Begin}", True)
        followingText = private_OutsideText(body, "{\export:" & identifier & "_End}", False)
        Set personNames = VBA.CreateObject("Scripting.Dictionary")
        Set reportSession = Nothing
        If VBA.Len(reportId) > 0 Then
            Set reportSession = New obj_PEB_TestWordSession
            If Not reportSession.Initialize(reportId, diagnostic) Then
                Err.Raise 5, , diagnostic
            End If
        End If
        stepIndex = 0
        Set target = private_FindMarker(document, "{\export:" & identifier & "_Begin}")
        expectedFont = target.Characters(1).Font.Name
        expectedSize = target.Characters(1).Font.Size
        For Each step In scenario("Steps")
            stepIndex = stepIndex + 1
            If step("Action") = "GroupEdit" Then
                Set anchors = New obj_PEB_WordAnchors
                If Not anchors.Initialize(document, diagnostic) Then
                    Err.Raise 5, , diagnostic
                End If
                If Not anchors.TryGetInsertedRange(document, values, target, diagnostic) Then
                    Err.Raise 5, , diagnostic
                End If
                target.Text = target.Text & "MANUAL"
                anchors.Dispose
            Else
                Set values = VBA.CreateObject("Scripting.Dictionary")
                For Each field In step("Values").Keys
                    values.Add field, step("Values")(field)
                Next field
                values("IPN") = values("PersonId")
                If Not templates.TryRender("ToBusinessTrip", values, rendered, diagnostic) Then
                    If VBA.Len(step("Error")) = 0 Then
                        Err.Raise 5, , diagnostic
                    End If
                End If
                values("WordText") = rendered
                values("WordPolicy") = "Render"
                values("WordGrouped") = True
                values("WordOutput") = path
                before = private_DocumentText(document)
                Set parts = document.CustomXMLParts.SelectByNamespace("urn:excelprototype:word-events:v1")
                beforeState = VBA.vbNullString
                If parts.Count = 1 Then
                    beforeState = parts(1).XML
                End If
                document.Close 0
                Set document = Nothing
                success = private_Apply(values, failure)
                Set document = word.Documents.Open(path, False, False, False)
                If VBA.Len(step("Error")) > 0 Then
                    If success Or VBA.InStr(1, failure, step("Error"), VBA.vbBinaryCompare) = 0 Then
                        Err.Raise 5, , "Expected grouping failure: " & step("Error") & " / " & failure
                    End If
                    If private_DocumentText(document) <> before Then
                        Err.Raise 5, , "Rejected export modified document text."
                    End If
                    Set parts = document.CustomXMLParts.SelectByNamespace("urn:excelprototype:word-events:v1")
                    If parts.Count <> 1 Then
                        Err.Raise 5, , "Rejected export modified the store count."
                    End If
                    If parts(1).XML <> beforeState Then
                        Err.Raise 5, , "Rejected export modified stored events."
                    End If
                ElseIf Not success Then
                    Err.Raise 5, , failure
                Else
                    personNames(VBA.CStr(values("FIOAccusative"))) = True
                End If
                If success And VBA.Len(beforeState) = 0 Then
                    Set anchors = New obj_PEB_WordAnchors
                    If Not anchors.Initialize(document, diagnostic) Then
                        Err.Raise 5, , diagnostic
                    End If
                    If Not anchors.TryGetInsertedRange(document, values, target, diagnostic) Then
                        Err.Raise 5, , diagnostic
                    End If
                    If VBA.Replace$(target.Text, VBA.vbCr, VBA.vbLf) <> VBA.Replace$(rendered, VBA.vbCrLf, VBA.vbLf) Then
                        Err.Raise 5, , "Single-event preview differs from persisted Word rendering."
                    End If
                    anchors.Dispose
                End If
            End If
            ' Every step crosses an actual save/close/reopen boundary.
            document.Save
            document.Close 0
            Set document = Nothing
            If Not reportSession Is Nothing Then
                caption = ex_PEB_Export.fn_Config("text.TestPreviewHeading")
                caption = VBA.Replace$(caption, "{scenario}", caseId)
                caption = VBA.Replace$(caption, "{step}", VBA.CStr(stepIndex))
                caption = VBA.Replace$(caption, "{event}", VBA.CStr(scenario("Name")))
                If Not reportSession.TryCollectSnapshot(path, values, caption, personNames, diagnostic) Then
                    Err.Raise 5, , diagnostic
                End If
                Set report = VBA.CreateObject("Scripting.Dictionary")
                report("WordResult") = "PASS"
                report("Details") = VBA.vbNullString
                If Not reportSession.TryAppend(report, diagnostic) Then
                    Err.Raise 5, , diagnostic
                End If
            End If
            Set document = word.Documents.Open(path, False, False, False)
        Next step
        Set parts = document.CustomXMLParts.SelectByNamespace("urn:excelprototype:word-events:v1")
        If parts.Count <> 1 Then
            Err.Raise 5, , "Expected exactly one embedded event store."
        End If
        Set state = VBA.CreateObject("MSXML2.DOMDocument.6.0")
        state.preserveWhiteSpace = True
        If Not state.LoadXML(parts(1).XML) Then
            Err.Raise 5, , "Cannot reopen embedded event data."
        End If
        state.setProperty "SelectionNamespaces", "xmlns:s='urn:excelprototype:word-events:v1'"
        If state.selectNodes("/s:state/s:section/s:event").Length <> VBA.CLng(scenario("Events")) Then
            Err.Raise 5, , "Unexpected persisted grouping event count."
        End If
        If state.documentElement.getAttribute("version") <> "2" Then
            Err.Raise 5, , "Unexpected Word store schema version."
        End If
        If state.selectNodes("/s:state/s:section/s:event/s:field[@name='WordText']").Length <> 0 Then
            Err.Raise 5, , "Word event snapshots must not store rendered text."
        End If
        body = private_DocumentText(document)
        Set anchors = New obj_PEB_WordAnchors
        If Not anchors.Initialize(document, diagnostic) Then
            Err.Raise 5, , diagnostic
        End If
        If Not anchors.TryGetInsertedRange(document, values, target, diagnostic) Then
            Err.Raise 5, , diagnostic
        End If
        If target.Font.Hidden <> 0 Then
            Err.Raise 5, , "Grouped Word content is hidden."
        End If
        If target.Font.Name <> expectedFont Or target.Font.Size <> expectedSize Then
            Err.Raise 5, , "Grouping changed the template font."
        End If
        anchors.Dispose
        If private_OutsideText(body, "{\export:" & identifier & "_Begin}", True) <> precedingText Then
            Err.Raise 5, , "Grouping changed preceding content."
        End If
        If private_OutsideText(body, "{\export:" & identifier & "_End}", False) <> followingText Then
            Err.Raise 5, , "Grouping changed following content."
        End If
        preview = target.Text
        For Each expected In scenario("Expected")
            If private_Count(preview, expected("Text")) <> VBA.CLng(expected("Count")) Then
                Err.Raise 5, , "Unexpected grouped text count: " & expected("Text")
            End If
        Next expected
        document.Close 0
        Set document = Nothing
        If Not reportSession Is Nothing Then
            If Not reportSession.TryFinish("Completed", diagnostic) Then
                Err.Raise 5, , diagnostic
            End If
            reportPath = reportSession.OutputPath
            reportSession.Dispose
            Set reportSession = Nothing
        End If
        files.DeleteFile path
        path = VBA.vbNullString
        count = count + 1
NextScenario:
    Next scenario
    If count = 0 Then
        Err.Raise 5, , "Grouping suite has no scenarios."
    End If
    private_Run = True
    GoTo Cleanup
EH:
    diagnostic = caseId & " | " & VBA.Err.Description
Cleanup:
    On Error Resume Next
    If Not reportSession Is Nothing Then
        reportSession.Dispose
    End If
    If Not document Is Nothing Then
        document.Close 0
    End If
    If Not word Is Nothing Then
        word.Quit 0
    End If
    If VBA.Len(path) > 0 Then
        If Not files Is Nothing Then
            If files.FileExists(path) Then
                files.DeleteFile path
            End If
        End If
    End If
    If Not templates Is Nothing Then
        templates.Dispose
    End If
    On Error GoTo 0
End Function

Private Function private_ReadScenarios( _
    ByRef scenarios As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim cases As Object
    Dim steps As Object
    Dim scenario As Object
    Dim item As Object
    Dim values As Object
    Dim expected As Collection
    Dim ordered As Collection
    Dim rows As Collection
    Dim record As Object
    Dim field As Variant
    Dim stepId As Variant
    Dim id As String
    Dim events As Long

    On Error GoTo EH
    If Not fn_ReadCatalog(scenarios, diagnostic) Then
        Exit Function
    End If
    Set cases = VBA.CreateObject("Scripting.Dictionary")
    Set steps = VBA.CreateObject("Scripting.Dictionary")
    For Each scenario In scenarios
        cases.Add scenario("Id"), scenario
        Set expected = New Collection
        Set scenario("Expected") = expected
    Next scenario
    If Not private_ReadRows("groupingSteps", rows, diagnostic) Then
        Exit Function
    End If
    For Each record In rows
        id = record("ScenarioId")
        If Not cases.Exists(id) Then
            Err.Raise 5, , "Grouping step references an unknown scenario: " & id
        End If
        Set item = VBA.CreateObject("Scripting.Dictionary")
        item("Id") = record("StepId")
        item("Action") = record("Action")
        item("Error") = record("ExpectedError")
        Set values = VBA.CreateObject("Scripting.Dictionary")
        If item("Action") = "GroupAdd" Then
            For Each field In VBA.Split(ex_PEB_Export.fn_Config("export.Word.grouping.stateFields"), "|")
                values.Add field, record(field)
            Next field
        ElseIf item("Action") <> "GroupEdit" Then
            Err.Raise 5, , "Unknown grouping action: " & item("Action")
        End If
        Set item("Values") = values
        steps.Add item("Id"), item
    Next record
    For Each scenario In scenarios
        Set ordered = New Collection
        For Each stepId In scenario("Steps")
            If Not steps.Exists(stepId) Then
                Err.Raise 5, , "Grouping step data is missing: " & stepId
            End If
            ordered.Add steps(stepId)
        Next stepId
        If ordered(1)("Action") <> "GroupAdd" Then
            Err.Raise 5, , "Grouping scenario must start with an add: " & scenario("Id")
        End If
        Set scenario("Steps") = ordered
    Next scenario
    If Not private_ReadRows("groupingExpected", rows, diagnostic) Then
        Exit Function
    End If
    For Each record In rows
        id = record("ScenarioId")
        If Not cases.Exists(id) Then
            Err.Raise 5, , "Grouping expectation references an unknown scenario: " & id
        End If
        Set scenario = cases(id)
        events = private_CountValue(record("ExpectedEvents"))
        If scenario.Exists("Events") Then
            If scenario("Events") <> events Then
                Err.Raise 5, , "Conflicting expected event counts: " & id
            End If
        Else
            scenario.Add "Events", events
        End If
        Set item = VBA.CreateObject("Scripting.Dictionary")
        item("Text") = record("Text")
        item("Count") = private_CountValue(record("Count"))
        If VBA.Len(item("Text")) = 0 Then
            Err.Raise 5, , "Grouping expected text is empty: " & id
        End If
        scenario("Expected").Add item
    Next record
    For Each scenario In scenarios
        If Not scenario.Exists("Events") Or scenario("Expected").Count = 0 Then
            Err.Raise 5, , "Grouping expectations are missing: " & scenario("Id")
        End If
    Next scenario
    private_ReadScenarios = True
    Exit Function
EH:
    diagnostic = VBA.Err.Description
End Function

Private Function private_ReadRows( _
    ByVal kind As String, _
    ByRef rows As Collection, _
    ByRef diagnostic As String _
) As Boolean
    Dim executor As obj_AdoQueryExecutor
    Dim source As obj_TableSource
    Dim query As obj_TableQuery
    Dim data As obj_DataTable
    Dim fields As Variant
    Dim field As Variant
    Dim record As Object
    Dim row As Long
    Dim column As Long

    On Error GoTo EH
    Set executor = New obj_AdoQueryExecutor
    Set source = New obj_TableSource
    Set query = New obj_TableQuery
    If Not executor.Initialize() Or Not source.Initialize() Or Not query.Initialize() Then
        Err.Raise 5, , "Cannot initialize grouping fixture reader."
    End If
    source.WorkbookPath = ex_PEB_Export.fn_Path(ex_PEB_Export.fn_Config("test.scenarios.path"))
    source.SheetName = ex_PEB_Export.fn_Config("test.scenarios." & kind & ".sheet")
    source.RangeAddress = ex_PEB_Export.fn_Config("test.scenarios." & kind & ".readRange")
    fields = VBA.Split(ex_PEB_Export.fn_Config("test.scenarios." & kind & ".fields"), "|")
    For Each field In fields
        query.AddColumn VBA.CStr(field)
    Next field
    If Not executor.TryExecute(source, query, data, diagnostic) Then
        GoTo Cleanup
    End If
    Set rows = New Collection
    For row = 1 To data.RowCount
        Set record = VBA.CreateObject("Scripting.Dictionary")
        For column = 0 To UBound(fields)
            record.Add fields(column), ex_PEB_Export.fn_Text(data.ValueAt(row, column + 1))
        Next column
        If VBA.Len(record(fields(0))) > 0 Then
            rows.Add record
        End If
    Next row
    private_ReadRows = True
Cleanup:
    If Not data Is Nothing Then
        data.Dispose
    End If
    If Not query Is Nothing Then
        query.Dispose
    End If
    If Not source Is Nothing Then
        source.Dispose
    End If
    If Not executor Is Nothing Then
        executor.Dispose
    End If
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function

Private Function private_CountValue(ByVal value As String) As Long
    If Not VBA.IsNumeric(value) Then
        Err.Raise 5, , "Expected grouping count must be numeric."
    End If
    If VBA.CDbl(value) < 0 Or VBA.CDbl(value) <> VBA.Fix(VBA.CDbl(value)) Then
        Err.Raise 5, , "Expected grouping count must be a nonnegative integer."
    End If
    private_CountValue = VBA.CLng(value)
End Function

Private Function private_Apply( _
    ByVal values As Object, _
    ByRef diagnostic As String _
) As Boolean
    Dim exporter As obj_PEB_ExptrWord

    diagnostic = VBA.vbNullString
    On Error GoTo EH
    Set exporter = New obj_PEB_ExptrWord
    If Not exporter.Initialize() Then
        Err.Raise 5, , "Cannot initialize grouped Word exporter."
    End If
    private_Apply = exporter.TryExportExistingTest(values, diagnostic)
Cleanup:
    If Not exporter Is Nothing Then
        exporter.Dispose
    End If
    Exit Function
EH:
    diagnostic = VBA.Err.Description
    Resume Cleanup
End Function

Private Sub private_CheckTemplate(ByVal document As Object)
    Dim anchors As obj_PEB_WordAnchors
    Dim values As Object
    Dim target As Object
    Dim diagnostic As String

    Set values = VBA.CreateObject("Scripting.Dictionary")
    values("EventId") = "ToBusinessTrip"
    Set anchors = New obj_PEB_WordAnchors
    If Not anchors.Initialize(document, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    If Not anchors.TryGetInsertedRange(document, values, target, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    If VBA.Len(VBA.Trim$(VBA.Replace$(target.Text, VBA.vbCr, " "))) > 0 Then
        Err.Raise 5, , "Grouping template must contain an empty ToBusinessTrip section."
    End If
    anchors.Dispose
End Sub

Private Function private_FindMarker( _
    ByVal document As Object, _
    ByVal marker As String _
) As Object
    Dim target As Object

    Set target = document.Content.Duplicate
    target.Find.ClearFormatting
    If Not target.Find.Execute(marker, True, False, False, False, False, True, 0, False) Then
        Err.Raise 5, , "Grouping template marker not found: " & marker
    End If
    Set private_FindMarker = target
End Function

Private Function private_DocumentText(ByVal document As Object) As String
    Dim target As Object

    Set target = document.Content.Duplicate
    target.TextRetrievalMode.IncludeHiddenText = True
    private_DocumentText = target.Text
End Function

Private Function private_OutsideText( _
    ByVal text As String, _
    ByVal marker As String, _
    ByVal preceding As Boolean _
) As String
    Dim offset As Long

    offset = VBA.InStr(1, text, marker, VBA.vbBinaryCompare)
    If offset = 0 Then
        Err.Raise 5, , "Grouping test marker is missing: " & marker
    End If
    If VBA.InStr(offset + VBA.Len(marker), text, marker, VBA.vbBinaryCompare) > 0 Then
        Err.Raise 5, , "Grouping test marker is duplicated: " & marker
    End If
    If preceding Then
        private_OutsideText = VBA.Left$(text, offset - 1)
    Else
        private_OutsideText = VBA.Mid$(text, offset + VBA.Len(marker))
    End If
End Function

Private Function private_Count( _
    ByVal text As String, _
    ByVal fragment As String _
) As Long
    Dim offset As Long
    Dim count As Long

    If VBA.Len(fragment) = 0 Then
        Err.Raise 5, , "Empty grouping assertion."
    End If
    offset = VBA.InStr(1, text, fragment, VBA.vbBinaryCompare)
    Do While offset > 0
        count = count + 1
        offset = VBA.InStr(offset + VBA.Len(fragment), text, fragment, VBA.vbBinaryCompare)
    Loop
    private_Count = count
End Function