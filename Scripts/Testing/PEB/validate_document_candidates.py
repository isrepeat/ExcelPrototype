"""Check document lookup wiring and context-sensitive XML replacements without Office."""

import re
from xml.etree import ElementTree as ET

from event_routes import ROOT, WORKBOOK, NS, parse_config, load_registry
from validate_movement_planner import check_declarations, procedure
from validate_movement_batch import check_blocks


def main():
    folder = ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder"
    controller = (folder / "obj_PEB_PgMainController.cls.vba").read_text(encoding="utf-8-sig")
    commands = (folder / "obj_PEB_PgMain.cls.vba").read_text(encoding="utf-8-sig")
    check_declarations(controller)
    check_blocks(controller)
    handler = procedure(controller, "DocumentChangedHandler")
    assert 'private_SearchSecondary(True, False, True, True)' in handler
    assert 'If Not VBA.CBool(autoFill) Then' in handler
    assert '"DocumentChanged"' in commands
    search = procedure(controller, "private_SearchSecondary")
    assert 'm_documentProfile.Initialize("PersonalEventBuilder::lookup.Document", diagnostic)' in search
    assert 'm_reporterActive = Not hospital And Not document' in search
    assert 'm_documentActive = document' in search
    apply = procedure(controller, "private_ApplyModeRecord")
    assert 'Set fields = m_documentFields' in apply
    load = procedure(controller, "private_LoadMode")
    assert 'secondaryCandidates = m_reporterActive Or m_hospitalActive Or m_documentActive' in load
    assert 'Not m_reporterActive And Not m_hospitalActive And Not m_documentActive' in load
    assert 'If Not secondaryCandidates Then' in load
    page = ET.parse(ROOT / "PROJECTS_EXCEL/ui/PersonalEventBuilder/MainPage.xaml")
    field = page.find('.//p:field[@name="Document"]', NS)
    assert field.get("onChange") == "{Binding Path=Commands.DocumentChanged}"

    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        settings = parse_config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder" / filename)
        load_registry(settings)
        assert settings["lookup.Document.source.path"] == "../Dependencies/Документи.xlsx"
        assert settings["lookup.Document.source.sheet"] == "Документи"
        source_range = re.fullmatch(r"A2:A([0-9]+)", settings["lookup.Document.source.range"])
        assert source_range and int(source_range[1]) > 2
        assert settings["lookup.Document.columns.1.header"] == "Документи"
        assert settings["lookup.Document.search.1.name"] == settings["lookup.Document.key"] == "Document.Name"
        for key, value in settings.items():
            if key.startswith("lookup.") and key.endswith(".minChars"):
                assert value.isdigit() and int(value) >= 1 and str(int(value)) == value, key
        assert settings["document.fields.count"] == "1"
        assert settings["document.fields.1.formField"] == "Document"
        assert settings["document.fields.1.record"] == "Document.Name"
        assert "Document" in settings["forms.lookupFields"].split("|")
        assert settings["document.fields.1.sourceAlias"] == "Document"
        assert "forms.sourceStyles.Document" in settings

    routes = (folder / "ex_PEB_EventRoutes.vba").read_text(encoding="utf-8-sig")
    apply_route = procedure(routes, "fn_Apply")
    for key, plan_key in (("SourceStatus", "sourceStatus"), ("TargetStatus", "targetStatus"), ("Operation", "operation")):
        assert f'values("{key}") = fn_Value(eventId, "movement.{plan_key}")' in apply_route
    assert 'values("Operation") = "Transition"' in apply_route
    assert 'eventContext.' in apply_route
    renderer = (folder / "obj_PEB_DslRenderer.cls.vba").read_text(encoding="utf-8-sig")
    let = procedure(renderer, "private_Let")
    assert 'values(name) = private_Render(private_Unquote(expression), values, depth + 1)' in let
    assert 'Invalid #let string literal.' in let
    assert let.index('Invalid #let string literal.') < let.index('opening = VBA.InStr(expression, "(")')
    check_declarations(renderer)
    check_blocks(renderer)
    settings = parse_config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder/wsConfig.txt")
    registry = load_registry(settings)
    for filename, discharge, card in (
        ("PersonalEventBuilderWordResultTemplates.xml", "виписка із медичної карти стаціонарного хворого", "медична карта стаціонарного хворого"),
        ("PersonalEventBuilderMovementTemplates.xml", "ВМКСХ", "МКСХ"),
    ):
        tree = ET.parse(WORKBOOK / filename)
        ids = [node.get("id") for node in tree.findall("p:sharedTemplates/p:sharedTemplate", NS)]
        assert ids.count("block.DocumentName") == ids.count("block.OptionalBasisDocuments") == 1
        assert not any(name.startswith("block.OptionalBasisDocumentsHospital") for name in ids)
        name = tree.find('.//p:sharedTemplate[@id="block.DocumentName"]/p:text', NS).text
        basis = tree.find('.//p:sharedTemplate[@id="block.OptionalBasisDocuments"]/p:text', NS).text
        assert '#let __DocumentName = "{[DocumentNote]}";' in basis
        assert '#let __DocumentName = "{item.[DocumentNote]}";' in basis
        assert basis.count("{#include block.DocumentName}") == 2
        assert '[DocNo]' in basis and 'item.[DocNo]' in basis
        assert '[DocDateShort]' in basis and 'item.[DocDateShort]' in basis
        assert "regexreplace:" not in basis
        assert '"{[SourceStatus]}", "^HospitalTreatment$"' in name
        patterns = re.findall(r'regexreplace:"([^"]+)","([^"]+)"', name)
        assert [replacement for _, replacement in patterns] == [discharge, card]
        for event_id, plan in registry["movement"]["events"].items():
            pattern, replacement = patterns[0 if plan["sourceStatus"] == "HospitalTreatment" else 1]
            for value in ("м/к", "М/К", " м / к ", "М/к"):
                assert re.sub(pattern, replacement, value) == replacement, event_id
            for value in ("направлення на ВЛК", "первинна медична карта форма 300", "м/к додаток", ""):
                assert re.sub(pattern, replacement, value) == value, event_id
    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        settings = parse_config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder" / filename)
        assert {"SourceStatus", "TargetStatus", "Operation"} <= set(settings["export.Word.context.xmlKeys"].split("|"))
        assert settings["eventContext.ToBusinessTrip.targetStatus"] == "BusinessTrip"
    print("PASS: document candidates, generic event context, shared document-name blocks, string let and medical-card mappings for all Movement plans. Office/VBA not executed.")


if __name__ == "__main__":
    main()