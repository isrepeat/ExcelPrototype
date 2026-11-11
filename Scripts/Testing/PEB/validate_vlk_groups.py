"""Validate status selectors and independent VLK exporter routes without Office."""

from event_routes import ROOT, NS, load_registry, parse_config


def main():
    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        settings = parse_config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder" / filename)
        events = load_registry(settings)["events"]
        extended = settings["ui.extendedModes"].split("|")
        assert len(extended) == len(set(extended))
        assert all("→" not in events[identifier].get("caption") for identifier in extended)
        pairs = {}
        legacy = set()
        for identifier, event in events.items():
            caption = event.get("caption")
            if event.get("legacy"):
                legacy.add(identifier)
            elif "→" in caption:
                source, target = (status.strip().casefold() for status in caption.split("→"))
                assert source and target and (source, target) not in pairs, caption
                pairs[source, target] = identifier
            elif event.get("visible") == "false":
                assert identifier in extended
        assert legacy == {"MoveFromVlkToHospital", "MoveFromVlkToVacationForCuring"}
        assert len(pairs) == 28
        groups = (
            ("ExternalVlk", "ВЛК за межами → ", (
                "MoveFromExternalVlkToHospital", "MoveFromExternalVlkToVacationForCuring", "MoveFromExternalVlkToAnnualVacation")),
            ("OutpatientVlk", "Амбулаторне ВЛК → ", (
                "MoveFromOutpatientVlkToHospital", "MoveFromOutpatientVlkToVacationForCuring", "MoveFromOutpatientVlkToAnnualVacation")),
        )
        destinations = []
        for source, prefix, identifiers in groups:
            targets = set()
            for identifier in identifiers:
                event = events[identifier]
                movement = event.find("p:movement", NS)
                assert movement.get("operation") == "Transition"
                assert movement.get("sourceStatus") == source
                targets.add(movement.get("targetStatus"))
                assert event.get("caption") == event.get("label")
                assert event.get("caption").startswith(prefix)
                assert movement.get("template") == identifier
                if source == "OutpatientVlk":
                    target = {"HospitalTreatment": "ToHospital", "MedicalLeave": "ToVacationForCuring", "AnnualVacation": "ToAnnualMainVacation"}
                    assert event.find("p:word", NS).get("template") == target[movement.get("targetStatus")]
                    assert settings["export.Word.textAnchors." + identifier]
                assert settings["forms.readOnly.modes." + identifier]
                previous, target = (status.strip().casefold() for status in event.get("caption").split("→"))
                assert pairs[previous, target] == identifier
            destinations.append(targets)
        assert destinations[0] == destinations[1] == {"HospitalTreatment", "MedicalLeave", "AnnualVacation"}
    print("PASS: unique selector pairs, equal VLK destinations, preserved legacy IDs, "
          "Movement transitions and outpatient Word departures. Office/VBA not executed.")


if __name__ == "__main__":
    main()