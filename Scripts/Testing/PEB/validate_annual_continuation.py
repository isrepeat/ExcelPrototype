"""Validate hospital discharge to annual-vacation continuation without Office/VBA."""

from xml.etree import ElementTree as ET
import re

from event_routes import ROOT, WORKBOOK, NS, parse_config, load_registry

EVENT = "MoveFromHospitalToAnnualVacationContinuation"


def main():
    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        settings = parse_config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder" / filename)
        registry = load_registry(settings)
        event = registry["events"][EVENT]
        known_readonly = set(settings["forms.readOnly.fields"].split("|"))
        for event_id in registry["events"]:
            readonly = settings[f"forms.readOnly.modes.{event_id}"]
            assert set(readonly.split("|")) - {"-"} <= known_readonly, event_id
        assert "DurationDays" not in settings[f"forms.readOnly.modes.{EVENT}"].split("|")
        assert EVENT in settings["lookup.AnnualContinuation.modes"].split("|")
        continuation_fields = dict(item.split("=", 1) for item in settings["lookup.AnnualContinuation.fields"].split("|"))
        assert continuation_fields["Hospital"] == "Movement.Hospital"
        assert continuation_fields["HospitalShortName"] == "Movement.HospitalShortName"
        assert continuation_fields["DocumentNumber"] == "Absence.TicketNumber"
        assert continuation_fields["DocumentDate"] == "Absence.TicketDate"
        assert continuation_fields["EventDate"] == "Absence.DepartureDate"
        card = re.compile(settings["lookup.AnnualContinuation.cardPattern"], re.IGNORECASE)
        for basis, expected in (
            ("рапорт; ВМКСХ № 5241 від 22.10.2026.", ("5241", "22.10.2026")),
            ("МКСХ № 5241 від 07.10.2026", ("5241", "07.10.2026")),
            ("м/к № 52/41 від 7.10.2026", ("52/41", "7.10.2026")),
            ("МКСХ № 5241", ("5241", None)),
        ):
            matches = list(card.finditer(basis))
            assert len(matches) == 1 and matches[0].groups() == expected, basis
        assert not list(card.finditer("рапорт № 5241 від 22.10.2026"))
        form = event.get("form")
        count = int(settings[f"forms.{form}.fields.count"])
        fields = {settings[f"forms.{form}.fields.{i}.name"] for i in range(1, count + 1)}
        assert {"Hospital", "Destination", "DurationDays", "EventDate"} <= fields
        mappings = {settings[f"forms.{form}.fields.{i}.name"]: settings.get(f"forms.{form}.fields.{i}.record") for i in range(1, count + 1)}
        assert mappings["DurationDays"] == "Absence.DurationDays"
        assert mappings["Destination"] == "Absence.Destination"
        assert event.get("direction") == "StatusChange" and event.get("form") == "HospitalLeave"
        assert event.get("visible") == "true" and EVENT in settings["ui.extendedModes"].split("|")
        plan = registry["movement"]["events"][EVENT]
        assert plan["operation"] == "Transition"
        assert plan["sourceStatus"] == "HospitalTreatment" and plan["targetStatus"] == "AnnualVacation"
        assert {"Destination", "DurationDays"} <= set(plan["required"].split("|"))
        for key in ("required", "previewRequired"):
            assert "DurationDays" in settings[f"export.Word.{EVENT}.{key}"].split("|")
        assert settings[f"export.Word.textAnchors.{EVENT}"] == EVENT
        assert settings[f"extensions.mode.{EVENT}.default.count"] == "1"
        assert settings[f"extensions.mode.{EVENT}.default.1.role"] == "MedicalCard"
        word = ET.parse(WORKBOOK / settings["export.Word.templates.path"]).getroot()
        template = word.find(f"p:template[@name='{EVENT}']", NS)
        body = "".join(template.itertext())
        assert "виписаний із лікувального закладу" in body
        assert "продовження невикористаної частини щорічної основної відпустки" in body
        assert "{#include block.StatusChangeResidenceVerb}" in body
        assert "{#include block.VacationDestinationPeriod}" in body
        assert "{#include block.OptionalForeignVacationPermission}" in body
        assert template.find("p:reject", NS) is None
        verb = word.find("p:sharedTemplates/p:sharedTemplate[@id='block.StatusChangeResidenceVerb']/p:text", NS).text
        assert 'fn_IsDateOnOrAfter("{[DateFromShort]}", "{[OrderDateShort]}")' in verb
        columns = registry["movement"]["tvo_columns"][EVENT]
        assert columns[settings["returnTvo.table.headers.Duration"]] == "Інше"
    source = (ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder/ex_PEB_AnnualContinuation.vba").read_text(encoding="utf-8-sig")
    assert "snapshot.TryQuery(source, query, table, diagnostic)" in source
    assert "service.TryExecute" not in source and "Workbooks.Open" not in source
    assert 'lookup.ReturnTickets.arrivalHeader' in source
    assert 'columns("ArrivalDate"))) = latest' in source
    assert 'DateDiff("d", CDate(latest), CDate(planned))' in source
    assert 'snapshot.Prepared.Add cacheKey, fields' in source
    print("PASS: PROD/TEST continuation route, candidate duration/destination mapping, snapshot-only linked remainder including travel, discharge document, shared tense rule and TVO columns. Office/VBA not executed.")


if __name__ == "__main__":
    main()