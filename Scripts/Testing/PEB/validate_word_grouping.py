"""Validate grouping XML, persisted scalar contracts and runtime scenario fixtures without Office."""

from pathlib import Path
from xml.etree import ElementTree as ET
import re

from event_routes import ROOT, WORKBOOK, NS, parse_config, load_registry, validate_word_template
from word_grouping_fixtures import load_grouping_cases
from validate_word_grammar import validate_cases


def validate_plan(template, fields):
    validate_word_template(template)
    derived = {node.get("name") for node in template.findall("p:derived", NS)}
    assert len(derived) == len(template.findall("p:derived", NS))
    assert derived.isdisjoint(fields)
    known = fields | derived
    for attribute in ("identityBy", "orderBy"):
        keys = template.get(attribute, "").split(";")
        assert all(key.strip() in fields for key in keys), attribute
    assert len(template.findall("p:group", NS)) == 1
    for group in template.findall(".//p:group", NS):
        assert all(key.strip() in known for key in group.get("by", "").split(";"))
        assert len(group.findall("p:single/p:text", NS)) == 1
        assert len(group.findall("p:multiple", NS)) == 1
        assert len(group.findall("p:multiple/p:item", NS)) + len(group.findall("p:multiple/p:group", NS)) == 1
        aggregates = set()
        for aggregate in group.findall("p:aggregate", NS):
            assert aggregate.get("name") not in aggregates | known
            aggregates.add(aggregate.get("name"))
            assert aggregate.get("field") in fields
            assert aggregate.get("operation") == "numberRanges"
            pattern = re.compile(aggregate.get("pattern"))
            assert pattern.groups == 3
            assert aggregate.get("separator") and aggregate.get("rangeSeparator")
        for tag in ("header", "footer"):
            for node in group.findall(f"p:multiple/p:{tag}", NS):
                common = {field.strip() for field in node.get("common", "").split(";")}
                assert common <= known
                used = set(re.findall(r"\[([A-Za-z][A-Za-z0-9_]*)\]", "".join(node.itertext())))
                assert used - aggregates <= common


def main():
    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        settings = parse_config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder" / filename)
        registry = load_registry(settings)
        assert registry["events"]["ToBusinessTrip"].find("p:word", NS).get("template") == "ToBusinessTrip"
        root = ET.parse(WORKBOOK / settings["export.Word.templates.path"]).getroot()
        template = root.find("p:template[@name='ToBusinessTrip']", NS)
        fields = set(settings["export.Word.grouping.stateFields"].split("|"))
        validate_plan(template, fields)
        certificate_pattern = template.find("p:group/p:aggregate", NS).get("pattern")
        assert not re.fullmatch(certificate_pattern, "2026/298/INVALID")
        assert re.fullmatch(certificate_pattern, "2026/298/101")
        assert "PersonId" not in settings["export.Word.grouping.output"]
        assert "AcceptedOrderNumber" in settings["export.Word.grouping.output"]
        assert settings["export.Word.textAnchors.ToBusinessTrip"] == "ToBusinessTrip"
        fixture = load_grouping_cases(settings)
        defaults = {node.get("name"): node.text or "" for node in fixture.find("defaults")}
        assert set(defaults) == set(settings["test.scenarios.groupingSteps.fields"].split("|")[5:])
        assert "WordText" not in fields
        assert template.find("p:text", NS) is None
        assert "{[GroupBasis]}" in template.find("p:group/p:single/p:text", NS).text
        single = template.find("p:group/p:single/p:text", NS).text
        assert single.index("Відрядити у") < single.index("{#newline}{#newline}") < single.index("{#include block.PersonLineLower}")
        assert single.count("{[Destination]|nbspAfterSettlementAbbr}") == 1
        assert single.count("{[DateFromShort]|") == 1
        footer = template.find("p:group/p:multiple/p:footer/p:text", NS).text
        assert "№№\u00a0{[CertificateNumbers]}" in footer
        certificate = root.find("p:sharedTemplates/p:sharedTemplate[@id='block.OptionalBusinessTripCertificate']/p:text", NS)
        assert certificate is not None and "№\u00a0{[VacationTicketNo]}" in certificate.text
        assert validate_cases() == 12
        ids = set()
        failure_cases = set()
        for scenario in fixture.findall("scenario"):
            identifier = scenario.get("id")
            assert identifier not in ids
            ids.add(identifier)
            assert scenario.findall("add") and scenario.findall("expect")
            assert int(scenario.get("events")) > 0
            for step in scenario.findall("add"):
                values = defaults | {node.get("name"): node.text or "" for node in step}
                assert set(values) == set(defaults)
                assert values["ModeId"] == "ToBusinessTrip"
                assert re.fullmatch(r"ipn_\d{4}", values["PersonId"])
                if step.get("error"):
                    failure_cases.add(step.get("error"))
                else:
                    assert re.fullmatch(template.find("p:group/p:aggregate", NS).get("pattern"), values["TicketNumber"])
            for assertion in scenario.findall("expect"):
                assert assertion.text and int(assertion.get("count")) >= 0
        assert len(ids) == 16
        assert "test.word.grouping.scenarios" not in settings
        assert "test.word.grouping.output" not in settings
        if settings["environment.name"] == "TEST":
            assert "{runId}" in settings["test.word.output"]
            progress = [settings["test.progress.Grouping" + phase] for phase in ("Prepare", "Inputs", "Export", "Reopen", "Verify", "Report", "Done")]
            assert list(map(int, progress)) == sorted(map(int, progress))
            assert int(progress[0]) == 0 and progress[-1] == settings["test.progress.stageCount"]
            for phase in ("Prepare", "Inputs", "Export", "Reopen", "Verify", "Report", "Done"):
                assert settings["text.TestPhase.Grouping" + phase]
        for key in ("TestsBasic", "TestsGrouping", "TestsGroupingDocuments"):
            assert settings["text." + key]
        assert failure_cases == {"different data", "edited manually", "Duplicate grouped document number", "different order", "declared range pattern"}
    ui = ET.parse(ROOT / "PROJECTS_EXCEL/ui/PersonalEventBuilder/TestsPage.xaml").getroot()
    controls = {node.get("name"): node for node in ui.iter() if node.get("name")}
    assert int(controls["TestStatus"].get("row")) < int(controls["TestsBasic"].get("row")) < int(controls["TestSelectAll"].get("row"))
    assert controls["TestsBasic"].get("row") == controls["TestsGrouping"].get("row")
    for name in ("TestsBasic", "TestsGrouping"):
        assert controls[name].get("command") == "{Binding Path=Commands." + name + "}"
        assert len("btn_" + name) <= 31
    folder = ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder"
    state = (folder / "ex_PEB_WordState.vba").read_text(encoding="utf-8-sig")
    groups_source = (folder / "ex_PEB_WordGroups.vba").read_text(encoding="utf-8-sig")
    assert "private_Attribute = VBA.CStr(value.nodeValue)" in groups_source
    exporter = (folder / "obj_PEB_ExptrWord.cls.vba").read_text(encoding="utf-8-sig")
    suite = (folder / "Test/ex_PEB_GroupingTests.vba").read_text(encoding="utf-8-sig")
    controller = (folder / "Test/obj_PEB_PgTestsController.cls.vba").read_text(encoding="utf-8-sig")
    page = (folder / "Test/obj_PEB_PgTests.cls.vba").read_text(encoding="utf-8-sig")
    for name in ("TestsBasic", "TestsGrouping"):
        assert '"' + name + '"' in page and "Public Function " + name + "Handler()" in controller
    assert "ex_PEB_GroupingTests.fn_ReadCatalog(catalog, diagnostic)" in controller
    assert "ex_PEB_GroupingTests.fn_TryRunCase" in controller
    assert "If m_running Or m_environmentSwitchPending Then" in controller
    assert 'For Each step In scenario("Steps")' in suite
    assert 'fn_ReadCatalog(catalog, diagnostic, "Grouping")' in suite
    assert 'private_ReadRows("groupingSteps", rows, diagnostic)' in suite
    assert 'private_ReadRows("groupingExpected", rows, diagnostic)' in suite
    assert "fn_CheckGrammar" not in suite and "private_CheckGrammar" not in suite
    assert "selectedId <> caseId" in suite and "files.DeleteFile path" in suite
    assert "files.MoveFile path" not in suite
    assert 'reportId = "Grouping_" & runId & "_" & scenarioId' in suite
    assert 'report("GroupingDocument")' not in controller
    assert "m_createWordReport, reportPath)" in controller
    assert 'Set document = word.Documents.Add(templatePath, False, 0, False)' in suite
    assert "private_CheckTemplate document" in suite
    assert "test.word.grouping.templateAnchor" not in suite
    assert "exporter.TryExportExistingTest(values, diagnostic)" in suite
    export_call = suite.index("success = private_Apply(values, failure)")
    assert suite.rfind("document.Close 0", 0, export_call) < suite.rfind("Set document = Nothing", 0, export_call)
    assert 'fn_ReportPhase "GroupingExport", stepIndex, progressUnits' in suite
    assert "reportSession.TryCollectSnapshot" in suite
    assert "Grouping changed the template font" in suite
    assert "PEB_Grouping_" in exporter and 'environment.name") <> "TEST"' in exporter
    report_source = (folder / "Test/obj_PEB_TestWordSession.cls.vba").read_text(encoding="utf-8-sig")
    assert "Public Function TryCollectSnapshot(" in report_source
    assert 'fn_Config("text.TestsGroupingHeading")' in suite
    assert 'fn_Config("text.TestsGroupingStep")' in suite
    assert 'caption = VBA.Replace$(caption, "{event}"' not in suite
    assert "detailRange.Font.Color = m_headingDetailColor" in report_source
    assert 'report("StepId") = step("Id")' in suite
    assert 'fn_Config("text.TestStepDescription." & VBA.CStr(report("StepId")))' in report_source
    assert "m_eventDescription & VBA.ChrW$(11) & stepDescription" in report_source
    assert "descriptionRange.Font.Color = m_headingDetailColor" in report_source
    assert "pointRange.Font.Bold = False" in report_source
    assert 'isBasis = (VBA.Left$(' in report_source
    assert "m_insertedRange.FormattedText" in report_source
    from zipfile import ZipFile
    with ZipFile(WORKBOOK / settings["test.word.template"]) as package:
        document = ET.fromstring(package.read("word/document.xml"))
        text = "".join(document.itertext())
        for suffix in ("Begin", "End"):
            assert text.count("{\\export:ToBusinessTrip_" + suffix + "}") == 1
    assert state.index("edited manually") < state.index("anchors.TryInsert")
    assert state.index("ex_PEB_WordGroups.fn_Render") < state.index("anchors.TryInsert")
    assert state.index("CustomXMLParts.Add(state.XML)") < state.index("part.Delete")
    assert "ReplaceFileW" in state
    assert exporter.index("fileSystem.CopyFile output") < exporter.index("wordApp.Documents.Open(temporary")
    assert exporter.index("document.Close 0") < exporter.index("ex_PEB_WordState.fn_Publish")
    assert "document.SaveAs2 path, 16" in suite and "word.Documents.Open(path" in suite
    assert "Rejected export modified document text" in suite and "Rejected export modified stored events" in suite
    assert 'Private Const STATE_VERSION As String = "2"' in state
    assert state.index("state.preserveWhiteSpace = True") < state.index("state.LoadXML(part.XML)")
    assert 'text = private_Normalize(VBA.CStr(sharedTemplates(key)))' in state
    assert 'exportService.TryPrepare(binding, "Word", values, diagnostic, False)' in suite
    assert 'lookupService.TrySearch(personId, result, diagnostic)' in suite
    assert 'If matches <> 1 Then' in suite
    assert 'values("IPN") = values("PersonId")' not in suite
    assert 'step("Values")' not in suite
    assert "Single-event preview differs from persisted Word rendering" in suite
    assert "precedingText = private_OutsideText(body," in suite
    assert "followingText = private_OutsideText(body," in suite
    assert 'False) <> followingText Then' in suite
    assert 'True) <> precedingText Then' in suite
    assert "target.TextRetrievalMode.IncludeHiddenText = True" in suite
    assert not re.search(r"(?im)^\s*(?:ByVal|ByRef|Dim)\s+(?:shared|attribute)\s+As\b", suite)
    for path in (folder / "ex_PEB_WordState.vba", folder / "ex_PEB_WordGroups.vba", folder / "Test/ex_PEB_GroupingTests.vba"):
        source = path.read_text(encoding="utf-8-sig")
        module = re.search(r'Attribute VB_Name = "([^"]+)"', source).group(1)
        assert len(module) <= 31
        procedures = re.findall(r"(?im)^(?:Public|Private) (?:Function|Sub) (\w+)", source)
        assert len(procedures) == len(set(name.lower() for name in procedures))
        assert len(procedures) == len(re.findall(r"(?im)^End (?:Function|Sub)$", source))
    print("PASS: TEST/PROD exclusive template bodies, 12 grammar cases and 16 save/reopen scenario definitions. VBA/Word runtime suite was not executed.")


if __name__ == "__main__":
    main()