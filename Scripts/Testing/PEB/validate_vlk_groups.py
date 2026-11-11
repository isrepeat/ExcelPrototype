"""Validate status selectors and independent VLK exporter routes without Office."""

from event_routes import ROOT, NS, load_registry, parse_config


def main():
    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        settings = parse_config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder" / filename)
        registry = load_registry(settings)
        events = registry["events"]
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
        generated = registry["movement"]["generated"]
        aliases = {identifier for identifier, (_, _, _, canonical) in generated.items() if identifier != canonical}
        assert legacy == {"MoveFromVlkToHospital", "MoveFromVlkToVacationForCuring"} | aliases
        assert len(pairs) == 37
        departures = {
            identifier: event for identifier, event in events.items()
            if event.get("direction") == "Departure" and "→" not in event.get("caption")
        }
        missing = {event_id for (source, _), event_id in pairs.items() if source == "безвісти зниклий"}
        assert missing == set(generated) - aliases
        assert {generated[identifier][0] for identifier in missing} == set(departures)
        assert "MoveFromMissingToAnnualVacation" in missing
        for transition_id in missing:
            base, _, template, _ = generated[transition_id]
            departure = departures[base]
            transition = events[transition_id]
            plan = transition.find("p:movement", NS)
            departure_plan = departure.find("p:movement", NS)
            assert transition.get("form") == departure.get("form")
            expected = "StatusChangeWithoutDepartureDocuments" if base in ("ToVacationForFamily", "ToMaternityLeave") else "StatusChange"
            assert plan.get("template") == expected
            if transition.find("p:word", NS) is not None:
                assert transition.find("p:word", NS).get("template") == template
            assert plan.get("operation") == "Transition" and plan.get("sourceStatus") == "Missing"
            for field in ("targetStatus", "destinationField", "destinationValue", "required"):
                assert plan.get(field) == departure_plan.get(field), (transition_id, field)
            assert (transition.find("p:word", NS) is None) == (departure.find("p:word", NS) is None)
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
                expected = "StatusChangeWithVlkBasis" if source == "ExternalVlk" and movement.get("targetStatus") != "AnnualVacation" else "StatusChange"
                assert movement.get("template") == expected
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