"""Check rejection schemas and exporter integration without executing Office/VBA."""

from copy import deepcopy
from xml.etree import ElementTree as ET
import re

from event_routes import ROOT, WORKBOOK, NS, load_registry, parse_config
from template_rejections import validate_rejections


def main():
    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        settings = parse_config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder" / filename)
        registry = load_registry(settings)
        word = ET.parse(WORKBOOK / settings["export.Word.templates.path"]).getroot()
        rules = word.findall("p:template/p:reject", NS)
        assert len(rules) == 1
        annual = word.find("p:template[@name='ToAnnualMainVacation']", NS)
        rule = annual.find("p:reject", NS)
        assert rule.get("field") == "Destination" and rule.get("messageKey") == "text.ExportRejected"
        assert "equals" not in rule.attrib
        pattern = re.compile(rule.get("matches"))
        permission = word.find("p:sharedTemplates/p:sharedTemplate[@id='block.OptionalForeignVacationPermission']/p:text", NS).text
        permission_pattern = re.search(r'm_RegexIsMatch\("\{\[Destination\]\}", "([^"]+)"\)', permission)
        assert permission_pattern and permission_pattern.group(1) == rule.get("matches")
        assert "{#if __ForeignVacation}{#newline}Дозволено виїзд за кордон.{#endif}" in permission
        vacation_templates = set()
        for template in word.findall("p:template", NS):
            body = "".join(template.itertext())
            include = "{#include block.OptionalForeignVacationPermission}"
            if "{#include block.VacationDestinationPeriod}" in body:
                vacation_templates.add(template.get("name"))
                assert body.count(include) == 1, template.get("name")
                assert body.index(include) < body.index("{#include block.IncomingBasis}"), template.get("name")
            else:
                assert include not in body, template.get("name")
        assert len(vacation_templates) == 18
        assert {"ToAnnualMainVacation", "ToVacationForFamily", "ToMaternityLeave",
                "ToVacationForCuring", "ToMedicalCompanyVacationForCuring",
                "MoveFromAnnualVacationToFamilyVacation", "MoveFromFamilyVacationToAnnualVacation"} <= vacation_templates
        for destination in (
            "Сіпоо, Фінляндська Республіка (Sipoo, Republic of Finland)",
            "м. Цюрих, Швейцарія (Zurich, Switzerland)",
            "м. Цюрих, Швейцарія (Zurich, Schweiz)",
            "м. Мюнхен (München, Germany)",
            "м. Бад Кройцнах (Bad\nKreuznach, Federal Republic of Germany)",
        ):
            assert pattern.search(destination), destination
        for destination in ("", "м. Київ", "м. Київ (Україна)", "(дорога)", "(4)",
                            "Zurich, Switzerland", "(Zurich", "(Zurich, Швейцарія)"):
            assert not pattern.search(destination), destination
        assert not registry["movement"]["root"].findall("p:template/p:reject", NS)
        movement = deepcopy(registry["movement"]["root"])
        movement.find("p:template", NS).insert(0, deepcopy(rule))
        load_registry(settings, movement_text=ET.tostring(movement, encoding="unicode"))
        annual.remove(rule)
        load_registry(settings, word_text=ET.tostring(word, encoding="unicode"))
        for key, value in (("field", "UnknownField"), ("messageKey", "text.MissingReason"),
                           ("matches", ""), ("matches", "("), ("equals", "4")):
            invalid = deepcopy(rule)
            invalid.set(key, value)
            template = ET.Element("template")
            template.append(invalid)
            try:
                validate_rejections(template, settings, {"Destination"})
            except ValueError:
                pass
            else:
                raise AssertionError("Malformed rejection accepted: " + key)
    source = ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder"
    service = (source / "obj_PEB_ExportService.cls.vba").read_text(encoding="utf-8-sig")
    assert 'values("VacationAdditionalDays") = 0' in service
    assert 'values("VacationAdditionalDays") = values("VacationAdditionalDays") + additional' in service
    history = service.split("Public Function TryPrepareHistory(", 1)[1].split("End Function", 1)[0]
    assert history.index("fn_ApplyWordRejection") < history.index('values("WordOutput") =')
    assert 'TryPrepareHistory = True\n        Exit Function' in history
    movement = service.split("Private Sub private_PrepareMovementTexts(", 1)[1].split("End Sub", 1)[0]
    assert movement.index("fn_CheckRejections") < movement.index("For Each column")
    renderer = (source / "obj_PEB_WordTemplates.cls.vba").read_text(encoding="utf-8-sig")
    render = renderer.split("Public Function TryRender(", 1)[1]
    assert "fn_CheckRejections" not in render
    export = (source / "ex_PEB_Export.vba").read_text(encoding="utf-8-sig")
    skip = export.split("Public Function fn_ApplyWordRejection(", 1)[1].split("End Function", 1)[0]
    assert 'values("WordPolicy") = "Omit"' in skip
    assert 'values("WordRejectionReason") = reason' in skip
    assert "Err.Raise" not in skip
    reason = export.split("Public Function fn_RejectionReason(", 1)[1].split("End Function", 1)[0]
    assert "Err.Raise" not in reason
    batch = (source / "obj_PEB_ExptrWord.cls.vba").read_text(encoding="utf-8-sig")
    generate = batch.split("Public Function TryGenerateHistory(", 1)[1].split("End Function", 1)[0]
    assert 'If item("WordPolicy") <> "Omit" Then' in generate
    assert 'If first Is Nothing Then\n        output = vbNullString\n        TryGenerateHistory = True' in generate
    controller = (source / "obj_PEB_PgMainController.cls.vba").read_text(encoding="utf-8-sig")
    assert 'exportItem("Status") = "RejectedOmitted"' in controller
    print("PASS: PROD/TEST annual Word rule, default allow, independent Movement rules, invalid definitions and preparation guards. Office/VBA not executed.")


if __name__ == "__main__":
    main()