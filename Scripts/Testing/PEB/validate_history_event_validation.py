"""Check the three-rule save gate independently from full export, without VBA."""

import re
import xml.etree.ElementTree as ET

from event_routes import ROOT, parse_config
from movement_plans import load_native_plans
from validate_movement_batch import check_blocks, check_module_contract, check_module_sections
from validate_movement_planner import check_declarations, procedure


def main():
    folder = ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder"
    read = lambda name: (folder / name).read_text(encoding="utf-8-sig")
    validator = read("ex_PEB_HistoryEvntVldtn.vba")
    source = read("ex_PEB_MovementSource.vba")
    registry = read("ex_PEB_MovementRegistry.vba")
    history = read("ex_PEB_History.vba")
    preview = read("ex_PEB_MovementPlanner.vba")
    full_rules = read("ex_PEB_MovementValidation.vba")
    for name, text in (("ex_PEB_HistoryEvntVldtn", validator), ("ex_PEB_MovementValidation", full_rules)):
        check_module_contract(text, name)
    check_module_sections(source)
    assert not re.search(r"(?m)^Private m_", source)
    check_module_sections(history)
    check_module_sections(registry)
    for text in (validator, source, registry):
        check_declarations(text)
        check_blocks(text)
    assert not re.search(r"(?m)^Private m_", validator)
    save = procedure(history, "fn_Save")
    gate = save.index("ex_PEB_HistoryEvntVldtn.fn_Validate(current, previous, diagnostic, historyIndex)")
    for mutation in ("ex_PEB_HistoryRevisions.fn_Prepare", "Set target = table.ListRows.Add", "written = True"):
        assert gate < save.index(mutation)
    validate = procedure(validator, "fn_Validate")
    assert 'personId = current("PersonId")' in validate
    assert 'previous("PersonId")' not in validate
    assert 'fn_ReadLatest(personId)' in validate
    assert 'stored("RowId") <> current("RowId")' in validate
    assert 'stored("Status") <> "Cancelled"' in validate
    assert 'current("Status") = "Cancelled"' in validate
    assert 'private_Before(stored, current)' in validate
    assert validate.index("For Each stored In ordered") < validate.index("Set stored = current")
    assert 'anchor("RowId") = current("RowId")' in validate
    assert 'anchor("Status") = "Cancelled"' in validate
    for forbidden in ("fn_StateRules", "ex_PEB_MovementValidation", "fn_FieldsHash", "OpeningId", "ClosingId",
                      "fn_CheckPeriods", "TryPrepareHistoryMovement", "sourceDestination",
                      "CancellationOpeningNotSupported", "private_Unapply"):
        assert forbidden not in validator, forbidden
    apply = procedure(validator, "private_Apply")
    for key in ("SimpleDepartureWhileOpen", "SimpleArrivalWithoutDeparture", "SimpleArrivalTypeMismatch"):
        assert key in apply
    assert apply.count("Err.Raise") == 3
    assert 'state("IsOpen") = (eventPlan("operation") <> "Arrival")' in apply
    assert 'statuses(eventPlan("sourceStatus"))' in apply
    assert 'statuses(eventPlan("targetStatus"))' in apply
    assert "Date" not in apply and "Hash" not in apply
    ordering = procedure(validator, "private_Before")
    assert 'firstDate < secondDate' in ordering
    assert 'VBA.Val(first("RowOrder")) < VBA.Val(second("RowOrder"))' in ordering
    assert "Err.Raise" not in ordering
    query = procedure(source, "fn_ReadLatest")
    for forbidden in ("Workbooks.Open", 'CreateObject("Excel.Application")',
                      "OpeningId", "ClosingId", "OrderNumber", "Hash", "fn_ReadPeople"):
        assert forbidden not in source
    assert 'VBA.Array("EventName", "DepartureDate", "ArrivalDate")' in query
    assert 'query.AddCondition ex_PEB_Export.fn_Config("export.Movement.headers.PersonId"), QueryEquals' in query
    assert 'service.TryExecute(source, query, table, diagnostic)' in query
    assert 'dateKey > latestKey' in query and 'arrivalKey <= latestArrivalKey' in query
    cleanup = query.split("Cleanup:", 1)[1].split("EH:", 1)[0]
    assert cleanup.index("On Error GoTo 0") < cleanup.index("Err.Raise 5, , failure")
    groups = procedure(registry, "fn_EventsByOperation")
    assert 'operation <> "Departure"' in groups and 'operation <> "Arrival"' in groups
    assert 'operation <> "Transition"' in groups
    assert 'm_events(eventId)("operation") = operation' in groups
    assert "ex_PEB_MovementValidation.fn_Apply" in procedure(preview, "private_Apply")
    assert "fn_FieldsHash" in procedure(preview, "private_PlanBasisSide")
    assert "OverlappingPeriods" in full_rules
    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        config = parse_config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder" / filename)
        assert config["validation.Movement.enabled"] == "true"
        for key in ("SimpleDepartureWhileOpen", "SimpleArrivalWithoutDeparture", "SimpleArrivalTypeMismatch"):
            assert config["text." + key]
        plans = load_native_plans(config)
        events = plans["events"]
        assert events["FromAnnualMainVacation"]["sourceStatus"] != events["FromVacationForCuring"]["sourceStatus"]
        for event in events.values():
            if event["operation"] == "Arrival":
                assert event["sourceStatus"] and not event["targetStatus"]
            elif event["operation"] == "Departure":
                assert event["targetStatus"] and not event["sourceStatus"]
            else:
                assert event["operation"] == "Transition" and event["sourceStatus"] and event["targetStatus"]
    ET.parse(ROOT / "PROJECTS_EXCEL/ui/PersonalEventBuilder/MainPage.xaml")
    print("PASS: three save rules only, latest-period query, date/RowOrder ordering, pending closure, edit/cancel boundary, one registry and full-export separation. Office/VBA not executed.")


if __name__ == "__main__":
    main()