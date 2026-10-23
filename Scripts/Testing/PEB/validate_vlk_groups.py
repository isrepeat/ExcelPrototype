"""Validate VLK menu groups and independent exporter routes without Office."""

from event_routes import ROOT, NS, load_registry, parse_config


def main():
    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        settings = parse_config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder" / filename)
        events = load_registry(settings)["events"]
        blocks = [group.split("|") for group in settings["ui.extendedModes"].split("||")]
        assert [len(group) for group in blocks] == [2, 1, 3, 3, 3, 1, 3, 3]
        menu = [identifier for group in blocks for identifier in group]
        assert len(menu) == len(set(menu))
        assert all(identifier in events for identifier in menu)
        hidden_aliases = {"MoveFromVlkToHospital", "MoveFromVlkToVacationForCuring"}
        extended = {identifier for identifier, event in events.items() if event.get("visible") == "false"}
        assert extended - hidden_aliases <= set(menu)
        groups = (
            ("ExternalVlk", "ВЛК за межами → ", (
                "MoveFromExternalVlkToHospital", "MoveFromExternalVlkToVacationForCuring", "MoveFromExternalVlkToAnnualVacation")),
            ("OutpatientVlk", "Амбулаторне ВЛК → ", (
                "MoveFromOutpatientVlkToHospital", "MoveFromOutpatientVlkToVacationForCuring", "MoveFromOutpatientVlkToAnnualVacation")),
        )
        expected = []
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
                expected.append(identifier)
            destinations.append(targets)
        assert destinations[0] == destinations[1] == {"HospitalTreatment", "MedicalLeave", "AnnualVacation"}
        start = menu.index(expected[0])
        assert menu[start:start + len(expected)] == expected
        assert not set(menu) & {"MoveFromVlkToHospital", "MoveFromVlkToVacationForCuring"}
        captions = [events[identifier].get("caption") for identifier in menu]
        assert len(captions) == len(set(captions))
        assert all(caption[0].isupper() and "->" not in caption for caption in captions)
        for group in (blocks[1], *blocks[3:]):
            sources = {events[identifier].find("p:movement", NS).get("sourceStatus") for identifier in group}
            assert len(sources) == 1
    print("PASS: equal VLK destinations, contiguous menu groups, preserved legacy IDs, "
          "Movement transitions and outpatient Word departures. Office/VBA not executed.")


if __name__ == "__main__":
    main()