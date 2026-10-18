"""Read grouping scenarios from the common workbook for offline validation."""

from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.utils.cell import range_boundaries

from event_routes import ROOT, WORKBOOK, parse_config


def load_personnel(settings):
    """Check fixture referential integrity against configured sources, without emulating VBA rendering."""
    prefix = "lookup.Personnel"
    source = openpyxl.load_workbook((WORKBOOK / settings[prefix + ".source.path"]).resolve(), data_only=True)
    sheet = source[settings[prefix + ".source.sheet"]]
    left, top, right, bottom = range_boundaries(settings[prefix + ".source.range"])
    headers = {sheet.cell(top, column).value: column for column in range(left, right + 1)}
    fields = {}
    for index in range(1, int(settings[prefix + ".columns.count"]) + 1):
        field_prefix = prefix + f".columns.{index}"
        fields[settings[field_prefix + ".name"]] = headers[settings[field_prefix + ".header"]]
    people = {}
    for values in sheet.iter_rows(min_row=top + 1, max_row=bottom, max_col=right, values_only=True):
        person_id = values[fields[settings[prefix + ".key"]] - 1]
        if not person_id:
            continue
        assert person_id not in people, "Duplicate personnel ID: " + str(person_id)
        person = {name: values[column - 1] for name, column in fields.items()}
        for index in range(1, int(settings["forms.Leave.fields.count"]) + 1):
            field_prefix = f"forms.Leave.fields.{index}"
            if settings.get(field_prefix + ".sourceAlias") == "Personnel":
                person[settings[field_prefix + ".name"]] = person[settings[field_prefix + ".record"]]
        people[person_id] = person
    source.close()
    cache = {}
    for index in range(1, int(settings["export.shared.lookups.count"]) + 1):
        prefix = f"export.shared.lookups.{index}"
        path = (WORKBOOK / settings[prefix + ".path"]).resolve()
        if path not in cache:
            cache[path] = openpyxl.load_workbook(path, data_only=True)
        sheet = cache[path][settings[prefix + ".sheet"]]
        left, top, right, bottom = range_boundaries(settings[prefix + ".range"])
        headers = {sheet.cell(top, column).value: column for column in range(left, right + 1)}
        key_column = headers[settings[prefix + ".keyHeader"]]
        value_column = headers[settings[prefix + ".valueHeader"]]
        rows = {}
        for values in sheet.iter_rows(min_row=top + 1, max_row=bottom, max_col=right, values_only=True):
            key = values[key_column - 1]
            if key:
                key = str(key).strip().casefold()
                assert key not in rows, "Duplicate lookup key: " + str(key)
                rows[key] = values[value_column - 1]
        for person in people.values():
            input_field = settings[prefix + ".input"]
            if input_field in person:
                key = str(person[input_field]).strip().casefold()
                if key in rows and rows[key]:
                    person[settings[prefix + ".output"]] = rows[key]
    for book in cache.values():
        book.close()
    return people


def table_rows(book, name):
    matches = [(sheet, sheet.tables[name]) for sheet in book if name in sheet.tables]
    assert len(matches) == 1, "Missing or duplicate fixture table: " + name
    sheet, table = matches[0]
    left, top, right, bottom = range_boundaries(table.ref)
    headers = [sheet.cell(top, column).value for column in range(left, right + 1)]
    assert headers == [column.name for column in table.tableColumns]
    assert len({column.id for column in table.tableColumns}) == len(table.tableColumns)
    assert len(headers) == len(set(headers))
    rows = []
    for index in range(top + 1, bottom + 1):
        values = [sheet.cell(index, column).value for column in range(left, right + 1)]
        assert not any(sheet.cell(index, column).data_type == "e" for column in range(left, right + 1))
        if any(value is not None for value in values):
            rows.append(dict(zip(headers, values)))
    return sheet.title, headers, rows


def load_grouping_cases(settings):
    path = (WORKBOOK / settings["test.scenarios.path"]).resolve()
    book = openpyxl.load_workbook(path, data_only=False)
    _, _, all_cases = table_rows(book, settings["test.scenarios.cases.table"])
    identifiers = [row["ScenarioId"] for row in all_cases]
    assert len(identifiers) == len(set(identifiers))
    assert all(row["Suite"] in ("Basic", "Grouping") for row in all_cases)
    cases = {row["ScenarioId"]: row for row in all_cases if row["Suite"] == "Grouping"}
    tables = {}
    for kind in ("groupingSteps", "groupingExpected"):
        prefix = "test.scenarios." + kind
        sheet, headers, rows = table_rows(book, settings[prefix + ".table"])
        assert sheet == settings[prefix + ".sheet"]
        assert headers == settings[prefix + ".fields"].split("|")
        assert all(row["ScenarioId"] in cases for row in rows)
        tables[kind] = rows
    fields = settings["test.scenarios.groupingSteps.fields"].split("|")[5:]
    forbidden = {"FIOAccusative", "RankGenitive", "PositionGenitive", "FIOInitialsGenitive", "ReporterGenitive", "IPN", "PersonName", "PersonRank", "PositionCode"}
    assert forbidden.isdisjoint(fields), "Grouping fixtures must contain raw form inputs only"
    test_settings = parse_config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder/wsConfig.test.txt")
    people = load_personnel(test_settings)
    step_ids = [row["StepId"] for row in tables["groupingSteps"]]
    assert len(step_ids) == len(set(step_ids))
    _, _, basic_steps = table_rows(book, settings["test.scenarios.steps.table"])
    assert set(step_ids).isdisjoint(row["StepId"] for row in basic_steps)
    description_prefix = "text.TestStepDescription."
    all_step_ids = set(step_ids) | {row["StepId"] for row in basic_steps}
    descriptions = {key[len(description_prefix):]: value for key, value in settings.items() if key.startswith(description_prefix)}
    if settings["environment.name"] == "TEST":
        assert set(descriptions) == all_step_ids, "Step description registry does not match the scenario workbook"
    else:
        assert not descriptions, "Step descriptions belong to the TEST configuration only"
    assert all(value.strip() and not any(character in value for character in "\t\r\n") for value in descriptions.values())
    root = ET.Element("scenarios")
    defaults = ET.SubElement(root, "defaults")
    for field in fields:
        ET.SubElement(defaults, "field", name=field).text = ""
    for identifier, case in cases.items():
        steps = sorted((row for row in tables["groupingSteps"] if row["ScenarioId"] == identifier), key=lambda row: row["StepOrder"])
        expected = [row for row in tables["groupingExpected"] if row["ScenarioId"] == identifier]
        assert steps and steps[0]["Action"] == "GroupAdd"
        assert [row["StepOrder"] for row in steps] == list(range(1, len(steps) + 1))
        assert expected and len({row["ExpectedEvents"] for row in expected}) == 1
        events = expected[0]["ExpectedEvents"]
        assert isinstance(events, int) and events > 0
        scenario = ET.SubElement(root, "scenario", id=identifier, events=str(events), name=case["Назва"])
        for row in steps:
            assert row["Action"] in ("GroupAdd", "GroupEdit")
            node = ET.SubElement(scenario, "add" if row["Action"] == "GroupAdd" else "edit", id=row["StepId"])
            if row["ExpectedError"]:
                node.set("error", row["ExpectedError"])
            if row["Action"] == "GroupAdd":
                assert row["PersonId"] in people, "Unknown grouping person: " + str(row["PersonId"])
                assert not row["ReporterId"] or row["ReporterId"] in people, "Unknown grouping reporter"
                for person_id in filter(None, (row["PersonId"], row["ReporterId"])):
                    assert all(people[person_id].get(field) for field in ("PersonName", "PersonRank", "PositionCode", "PersonAccusative", "RankGenitive", "PositionGenitive", "FIOInitialsGenitive")), person_id
                for field in fields:
                    ET.SubElement(node, "field", name=field).text = str(row[field]) if row[field] is not None else ""
        for row in expected:
            assert row["Text"] and isinstance(row["Count"], int) and row["Count"] >= 0
            ET.SubElement(scenario, "expect", count=str(row["Count"])).text = row["Text"]
    book.close()
    return root