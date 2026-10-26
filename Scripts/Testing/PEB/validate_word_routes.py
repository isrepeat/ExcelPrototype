"""Validate independent optional exporter routes without Office."""

from copy import deepcopy
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from event_routes import ROOT, WORKBOOK, NS, load_registry, parse_config


def main():
    folder = ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder"
    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        settings = parse_config(folder / filename)
        registry = load_registry(settings)
        assert len(registry["events"]) == 45
        assert sum(event.find("p:word", NS) is not None for event in registry["events"].values()) == 41
        assert registry["events"]["ToBusinessTrip"].find("p:movement", NS) is not None
        movement_only = {name for name, event in registry["events"].items() if event.find("p:word", NS) is None}
        assert movement_only == {"FromOutpatientExaminationOrVlk", "ToOutpatientExaminationOrVlk", "MoveFromAwolToBusinessTrip", "ToBusinessTrip"}
        assert registry["events"]["ToBusinessTrip"].get("caption") == "У відрядження"
        for event_id in ("MoveFromAnnualVacationToVacationForCuring", "MoveFromFamilyVacationToVacationForCuring",
                         "MoveFromVacationForCuringToAnnualVacation"):
            event = registry["events"][event_id]
            assert event.get("direction") == "StatusChange" and event.get("visible") == "true"
            assert event.find("p:word", NS).get("template") == event_id
            assert event.find("p:movement", NS).get("template") == event_id
            assert settings["export.Word.textAnchors." + event_id] == event_id
        templates = [path for path in WORKBOOK.glob("*.docx") if path.name.startswith("{yyyy}")]
        assert len(templates) == 1
        with ZipFile(templates[0]) as package:
            document = ET.fromstring(package.read("word/document.xml"))
            text = "".join(document.itertext())
            for event_id, event in registry["events"].items():
                if event.find("p:word", NS) is not None:
                    anchor = settings["export.Word.textAnchors." + event_id]
                    for boundary in ("Begin", "End"):
                        assert text.count("{\\export:" + anchor + "_" + boundary + "}") == 1, event_id
            for event_id in ("MoveFromAnnualVacationToVacationForCuring", "MoveFromFamilyVacationToVacationForCuring",
                             "MoveFromVacationForCuringToAnnualVacation"):
                for boundary in ("Begin", "End"):
                    assert text.count("{\\export:" + event_id + "_" + boundary + "}") == 1
        root = ET.parse(WORKBOOK / settings["events.path"]).getroot()
        assert [node.tag for node in root] == [f"{{{NS['p']}}}events"]
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
    reject_route(lambda event: event.find("p:word", NS).set("template", "MissingTemplate"))
    reject_route(lambda event: event.find("p:word", NS).attrib.pop("template"))
    reject_route(lambda event: event.find("p:movement", NS).set("template", "MissingTemplate"))
    reject_route(lambda event: event.find("p:movement", NS).set("operation", "MissingOperation"))
    reject_route(lambda event: event.find("p:movement", NS).attrib.pop("template"))
    reject_route(lambda event: event.append(deepcopy(event.find("p:word", NS))))
    reject_route(lambda event: event.append(deepcopy(event.find("p:movement", NS))))
    reject_route(lambda event: event.append(ET.Element(f"{{{NS['p']}}}unknown")))
    reject_route(lambda event: [event.remove(child) for child in list(event)])

    for keep, remove, settings_key in (("word", "movement", "export.Movement.templates.path"),
                                       ("movement", "word", "export.Word.templates.path")):
        altered = deepcopy(root)
        events = altered.find("p:events", NS)
        first = events[0]
        for event in list(events)[1:]:
            events.remove(event)
        first.remove(first.find("p:" + remove, NS))
        local = dict(settings)
        local[settings_key] = "./THIS_FILE_MUST_NOT_BE_READ.xml"
        loaded = load_registry(local, xml_text=ET.tostring(altered, encoding="unicode"))
        assert loaded["events"][first.get("id")].find("p:" + keep, NS) is not None
        assert loaded["events"][first.get("id")].find("p:" + remove, NS) is None

    word_root = ET.parse(WORKBOOK / settings["export.Word.templates.path"]).getroot()
    assert word_root.find("p:omittedTemplates", NS) is None
    altered = deepcopy(word_root)
    altered.append(ET.Element(f"{{{NS['p']}}}omittedTemplates"))
    try:
        load_registry(settings, word_text=ET.tostring(altered, encoding="unicode"))
    except ValueError:
        rejected += 1
    else:
        raise AssertionError("Retired omittedTemplates section accepted")
    print(f"PASS: TEST/PROD, 41 Word routes and 4 Movement-only events; medical-leave DOCX anchors verified; Word-only and Movement-only registries work independently; {rejected} invalid definitions rejected. Office/VBA not executed.")


if __name__ == "__main__":
    main()