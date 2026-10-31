"""Check independent per-column text templates and native Movement plans."""

from copy import deepcopy
from xml.etree import ElementTree as ET
import re

from event_routes import ROOT, WORKBOOK, NS, load_registry, parse_config
from movement_templates import load_movement
from movement_plans import load_native_plans


def main():
    config = ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder"
    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        settings = parse_config(config / filename)
        configured = dict(settings)
        configured["export.Movement.headers.PersonId"] = "Custom person identifier"
        configured["export.Movement.destinations.MedicalCompany"] = "Custom medical company"
        plans = load_native_plans(configured)
        assert plans["values"]["profiles.Departure.fields.1.header"] == "Custom person identifier"
        assert plans["events"]["ToMedicalCompanyForTreatment"]["destinationValue"] == "Custom medical company"
        configured["export.Movement.headers.ArrivalBasis"] = "Custom arrival basis"
        configured["export.Movement.headers.DepartureBasis"] = "Custom departure basis"
        renamed = load_registry(configured)["movement"]["columns"]
        assert set(renamed["MoveFromVacationForCuringToHospital"]) == {"Custom arrival basis", "Custom departure basis"}
        missing_header = dict(settings)
        del missing_header["export.Movement.headers.ArrivalBasis"]
        try:
            load_movement(missing_header, WORKBOOK)
        except ValueError as error:
            assert "export.Movement.headers.ArrivalBasis" in str(error)
        else:
            raise AssertionError("Missing XML column configuration accepted")
        for key in ("export.Movement.headers.PersonId", "export.Movement.destinations.MedicalCompany"):
            missing = dict(settings)
            del missing[key]
            try:
                load_native_plans(missing)
            except ValueError as error:
                assert key in str(error)
            else:
                raise AssertionError("Missing Movement setting accepted: " + key)
        registry = load_registry(settings)
        movement = registry["movement"]
        assert len(movement["templates"]) == len(movement["events"]) == 39
        assert len(registry["events"]) == 40
        for event_id, source, target in (
            ("MoveFromAwolToBusinessTrip", "Awol", "BusinessTrip"),
            ("MoveFromAnnualVacationToVacationForCuring", "AnnualVacation", "MedicalLeave"),
            ("MoveFromFamilyVacationToVacationForCuring", "FamilyVacation", "MedicalLeave"),
            ("MoveFromVacationForCuringToAnnualVacation", "MedicalLeave", "AnnualVacation"),
        ):
            plan = movement["events"][event_id]
            assert plan["operation"] == plan["profile"] == "Transition"
            assert plan["sourceStatus"] == source and plan["targetStatus"] == target
            assert plan["destinationField"] == "Destination"
            assert "Destination" in plan["required"].split("|")
        assert movement["statuses"]["Awol"] == "СЗЧ"
        outpatient = movement["events"]["ToOutpatientExaminationOrVlk"]
        assert outpatient["destinationField"] == "Destination"
        assert "Destination" not in outpatient["required"].split("|")
        assert "HospitalShortName" not in outpatient["required"].split("|")
        assert "Hospital" not in outpatient["required"].split("|")
        assert "HospitalShortName" in movement["events"]["ToHospital"]["required"].split("|")
        assert len(movement["profiles"]) == 5 and len(movement["values"]) == 167
        root = movement["root"]
        assert root.tag == f"{{{NS['p']}}}resultTemplates"
        assert "Ref." not in ET.tostring(root, encoding="unicode")
        for name, node in movement["templates"].items():
            assert set(node.attrib) == {"name"}
            operation = movement["events"][name]["operation"]
            expected = {"Прибуття.Підстава"} if operation == "Arrival" else {"Вибуття.Підстава"}
            if operation == "Transition":
                expected.add("Прибуття.Підстава")
            assert set(movement["columns"][name]) == expected
        assert sum(len(columns) for columns in movement["columns"].values()) == 63
        for name, columns in movement["columns"].items():
            for text in columns.values():
                assert "block.OptionalVacationTicket" not in text, name
                assert "block.OptionalPreviousVacationTicket" not in text, name

    rejected = 0
    def reject(change, routed=False):
        nonlocal rejected
        altered = deepcopy(root)
        change(altered)
        try:
            text = ET.tostring(altered, encoding="unicode")
            if routed:
                load_registry(settings, movement_text=text)
            else:
                load_movement(settings, WORKBOOK, text)
        except ValueError:
            rejected += 1
        else:
            raise AssertionError("Malformed column text definition accepted")
    def template(tree):
        return tree.find("p:template", NS)
    def column(tree):
        return template(tree).find("p:column", NS)
    def text(tree):
        return column(tree).find("p:text", NS)
    reject(lambda tree: text(tree).__setattr__("text", ""))
    reject(lambda tree: text(tree).__setattr__("text", "{#include Missing}"))
    reject(lambda tree: text(tree).__setattr__("text", "{[MissingField]}"))
    reject(lambda tree: column(tree).attrib.pop("configKey"))
    reject(lambda tree: column(tree).set("configKey", " "))
    reject(lambda tree: column(tree).set("name", "LegacyHeader"))
    reject(lambda tree: template(tree).append(deepcopy(column(tree))))
    reject(lambda tree: template(tree).remove(column(tree)))
    reject(lambda tree: tree.find("p:sharedTemplates/p:sharedTemplate/p:text", NS).__setattr__("text", "{#include block.IncomingBasis}"))
    reject(lambda tree: tree.append(deepcopy(template(tree))))
    reject(lambda tree: template(tree).set("operation", "close"))
    reject(lambda tree: template(tree).set("column", "OldFormat"))
    reject(lambda tree: tree.append(ET.Element(f"{{{NS['p']}}}profiles")))
    reject(lambda tree: column(tree).set("configKey", "export.Movement.headers.UnknownHeader"), routed=True)
    reject(lambda tree: tree.find("p:template[@name='MoveFromVacationForCuringToHospital']", NS).remove(
        tree.find("p:template[@name='MoveFromVacationForCuringToHospital']/p:column", NS)), routed=True)

    custom = deepcopy(root)
    event = custom.find("p:template[@name='MoveFromVacationForCuringToHospital']", NS)
    event.find("p:column[@configKey='export.Movement.headers.ArrivalBasis']/p:text", NS).text = "Close: {[IncomingNo]}."
    event.find("p:column[@configKey='export.Movement.headers.DepartureBasis']/p:text", NS).text = "Open: {[IncomingNo]}."
    loaded = load_registry(settings, movement_text=ET.tostring(custom, encoding="unicode"))
    texts = loaded["movement"]["columns"][event.get("name")]
    assert texts["Прибуття.Підстава"] == "Close: {[IncomingNo]}."
    assert texts["Вибуття.Підстава"] == "Open: {[IncomingNo]}."

    folder = ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder"
    service = (folder / "obj_PEB_ExportService.cls.vba").read_text(encoding="utf-8-sig")
    basis = service.split("Private Sub private_PrepareMovementTexts(", 1)[1].split("\nEnd Sub", 1)[0]
    assert "m_pebWordTemplates" not in basis and "basis.prefix" not in basis
    assert "ex_PEB_MovementTemplates.fn_Texts(templateId)" in basis
    assert 'Set values("MovementColumnTexts") = results' in basis
    assert 'values("MovementBasis")' not in basis
    exporter = (folder / "obj_PEB_ExptrMovement.cls.vba").read_text(encoding="utf-8-sig")
    assert "fn_WriteValue(values, field, header)" in exporter
    assert 'ex_PEB_MovementRegistry.fn_Value(prefix & ".operation")' in exporter
    for filename in ("ex_PEB_MovementTemplates.vba", "ex_PEB_MovementRegistry.vba", "ex_PEB_EventRoutes.vba",
                     "obj_PEB_ExportService.cls.vba", "obj_PEB_ExptrMovement.cls.vba", "Test/obj_PEB_TestRunner.cls.vba"):
        text = (folder / filename).read_text(encoding="utf-8-sig")
        module = re.search(r'Attribute VB_Name = "([^"]+)"', text).group(1)
        assert len(module) <= 31
        procedures = re.findall(r"(?im)^(?:Public|Private|Friend) (?:Function|Sub) (\w+)", text)
        assert len(procedures) == len(set(name.lower() for name in procedures))
        assert len(procedures) == len(re.findall(r"(?im)^End (?:Function|Sub)$", text))
        for name, body in re.findall(r"(?ms)^(?:Public|Private|Friend) (?:Function|Sub) (\w+)(.*?)(?=^End (?:Function|Sub)$)", text):
            declarations = re.findall(r"(?im)^\s*Dim (\w+)\b", body)
            assert len(declarations) == len(set(name.lower() for name in declarations)), (filename, name)
            declared = {name.lower()} | {value.lower() for value in re.findall(r"(?im)^Private (\w+) As ", text)}
            declared.update(value.lower() for value in re.findall(r"(?:Dim |ByVal |ByRef )(\w+) As ", body))
            assigned = {value.lower() for value in re.findall(r"(?m)^\s*(?:Set )?(\w+)\s*=", body)}
            assert assigned <= declared, (filename, name, assigned - declared)

    assert "Ref.Form." not in service
    native = (folder / "ex_PEB_MovementRegistry.vba").read_text(encoding="utf-8-sig")
    assert "DOMDocument" not in native and "selectNodes" not in native
    assert 'MovementBasis' in native and 'SourceDestination' in native

    print(f"PASS: TEST/PROD, 39 templates and 63 column texts; AWOL/business-trip and medical-leave transitions; distinct transition texts accepted; {rejected} invalid definitions rejected. Office/VBA not executed.")


if __name__ == "__main__":
    main()