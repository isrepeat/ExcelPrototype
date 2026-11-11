"""Validate independent optional exporter routes without Office."""

from copy import deepcopy
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from event_routes import ROOT, WORKBOOK, NS, load_registry, parse_config
from xml_imports import import_path


def main():
    folder = ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder"
    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        settings = parse_config(folder / filename)
        registry = load_registry(settings)
        medical_leave_id = "MoveFromHospitalToMedicalCompanyVacationForCuring"
        medical_leave = registry["events"][medical_leave_id]
        assert medical_leave.get("form") == "HospitalLeave"
        assert medical_leave.get("direction") == "StatusChange"
        assert medical_leave.get("legacy") is None and "→" in medical_leave.get("caption")
        assert medical_leave.find("p:word", NS).get("template") == medical_leave_id
        assert settings["export.Word.textAnchors." + medical_leave_id] == "MoveFromHospitalToVacationForCuring"
        movement_leave = registry["movement"]["events"][medical_leave_id]
        assert movement_leave["operation"] == movement_leave["profile"] == "Transition"
        assert movement_leave["sourceStatus"] == "HospitalTreatment"
        assert movement_leave["targetStatus"] == "MedicalLeave"
        assert movement_leave["destinationValue"] == settings["export.Movement.destinations.MedicalCompany"]
        assert registry["events"]["ToBusinessTrip"].find("p:movement", NS) is not None
        movement_only = {name for name, event in registry["events"].items() if event.find("p:word", NS) is None}
        assert {"FromOutpatientExaminationOrVlk", "ToOutpatientExaminationOrVlk", "ToBusinessTrip",
                "MoveFromMissingToBusinessTrip", "MoveFromMissingToOutpatientVlk"} <= movement_only
        assert registry["events"]["ToBusinessTrip"].get("caption") == "У відрядження"
        for event_id in ("MoveFromAnnualVacationToVacationForCuring", "MoveFromFamilyVacationToVacationForCuring",
                         "MoveFromVacationForCuringToAnnualVacation"):
            event = registry["events"][event_id]
            assert event.get("direction") == "StatusChange"
            assert event.find("p:word", NS).get("template") == event_id
            assert event.find("p:movement", NS).get("template") == "StatusChangeWithVlkBasis"
            assert settings["export.Word.textAnchors." + event_id] == event_id
        templates = [path for path in WORKBOOK.glob("*.docx") if path.name.startswith("{yyyy}")]
        assert len(templates) == 1
        with ZipFile(templates[0]) as package:
            document = ET.fromstring(package.read("word/document.xml"))
            text = "".join(document.itertext())
            for event_id, event in registry["events"].items():
                if event.find("p:word", NS) is not None:
                    generated = registry["movement"]["generated"].get(event_id)
                    profile_id = generated[0] if generated else event_id
                    anchor = event.get("wordSectionAnchor", settings["export.Word.textAnchors." + profile_id])
                    for kind in ("export", "block"):
                        for boundary in ("Begin", "End"):
                            assert text.count("{\\" + kind + ":" + anchor + "_" + boundary + "}") == 1, event_id
            for event_id in ("MoveFromAnnualVacationToVacationForCuring", "MoveFromFamilyVacationToVacationForCuring",
                             "MoveFromVacationForCuringToAnnualVacation"):
                assert "{\\export:" + event_id + "_Begin}" not in text
            for event_id in ("MoveFromAnnualVacationToHospital", "MoveFromMissingToHospital"):
                assert registry["events"][event_id].get("wordSectionAnchor") == "StatusChanges"
            for event_id, anchor in (
                    ("MoveFromVacationForCuringToVacationForCuring", "StatusChanges"),
                    ("MoveFromHospitalToMedicalLeaveContinuation", "StatusChanges"),
                    ("MoveFromHospitalToAnnualVacationContinuation", "StatusChanges")):
                assert registry["events"][event_id].get("wordSectionAnchor") == anchor
                assert "{\\export:" + event_id + "_Begin}" not in text
        root = ET.parse(WORKBOOK / settings["events.path"]).getroot()
        assert [node.tag for node in root] == [f"{{{NS['p']}}}{name}" for name in ("imports", "events", "transitions")]
        for event in root.find("p:events", NS):
            for route in event:
                assert set(route.attrib) == {"template"}
        special = registry["events"]["MoveFromOutpatientVlkToVacationForCuring"].find("p:word", NS)
        assert special.get("template") == "ToVacationForCuring"

    rejected = 0
    def reject_route(change):
        nonlocal rejected
        altered = deepcopy(root)
        change(altered.find("p:events/p:event", NS))
        try:
            load_registry(settings, xml_text=ET.tostring(altered, encoding="unicode"))
        except ValueError:
            rejected += 1
        else:
            raise AssertionError("Malformed exporter route accepted")

    for key in ("policy", "profile", "anchor"):
        reject_route(lambda event, key=key: event.find("p:word", NS).set(key, "Removed"))
    reject_route(lambda event: event.find("p:word", NS).set("template", "word:MissingTemplate"))
    reject_route(lambda event: event.find("p:word", NS).set("template", "movement:FromBusinessTrip"))
    reject_route(lambda event: event.find("p:word", NS).attrib.pop("template"))
    reject_route(lambda event: event.find("p:movement", NS).set("template", "movement:MissingTemplate"))
    reject_route(lambda event: event.find("p:movement", NS).set("operation", "MissingOperation"))
    reject_route(lambda event: event.find("p:movement", NS).attrib.pop("template"))
    reject_route(lambda event: event.append(deepcopy(event.find("p:word", NS))))
    reject_route(lambda event: event.append(deepcopy(event.find("p:movement", NS))))
    reject_route(lambda event: event.append(ET.Element(f"{{{NS['p']}}}unknown")))
    reject_route(lambda event: [event.remove(child) for child in list(event)])

    for keep, remove in (("word", "movement"), ("movement", "word")):
        altered = deepcopy(root)
        altered.remove(altered.find("p:transitions", NS))
        events = altered.find("p:events", NS)
        first = events[0]
        for event in list(events)[1:]:
            events.remove(event)
        first.remove(first.find("p:" + remove, NS))
        altered.find(f"p:imports/p:import[@alias='{remove}']", NS).set("path", "THIS_FILE_MUST_NOT_BE_READ.xml")
        loaded = load_registry(settings, xml_text=ET.tostring(altered, encoding="unicode"))
        assert loaded["events"][first.get("id")].find("p:" + keep, NS) is not None
        assert loaded["events"][first.get("id")].find("p:" + remove, NS) is None

    word_root = ET.parse(import_path(WORKBOOK / settings["events.path"], "word")).getroot()
    assert word_root.find("p:omittedTemplates", NS) is None
    altered = deepcopy(word_root)
    altered.append(ET.Element(f"{{{NS['p']}}}omittedTemplates"))
    try:
        load_registry(settings, word_text=ET.tostring(altered, encoding="unicode"))
    except ValueError:
        rejected += 1
    else:
        raise AssertionError("Retired omittedTemplates section accepted")
    print(f"PASS: TEST/PROD Word and Movement routes, shared DOCX anchors and independent exporters; {rejected} invalid definitions rejected. Office/VBA not executed.")


if __name__ == "__main__":
    main()