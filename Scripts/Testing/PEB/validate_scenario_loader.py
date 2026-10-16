"""Проверка схем сценариев и контрактов исходного кода без запуска Office."""

import copy
import re
import sys
from pathlib import Path

import openpyxl
from openpyxl.utils.cell import range_boundaries

ROOT = Path(__file__).resolve().parents[3]
FIXTURE = ROOT / "PROJECTS_EXCEL/tests/PersonalEventBuilder/TestScenarios.xlsx"


def config(filename):
    settings = {}
    seen = set()
    for number, line in enumerate(filename.read_text(encoding="utf-8-sig").splitlines(), 1):
        if not line.strip():
            continue
        parts = line.split("\t")
        assert len(parts) == 3, f"Invalid config row {filename}:{number}"
        assert parts[1] not in seen, f"Duplicate config key: {parts[1]}"
        seen.add(parts[1])
        settings[parts[1].split("::", 1)[-1]] = parts[2]
    return settings


def load_tables(book):
    result = {}
    for sheet in book:
        for table in sheet.tables.values():
            left, top, right, bottom = range_boundaries(table.ref)
            headers = [sheet.cell(top, column).value for column in range(left, right + 1)]
            assert headers == [column.name for column in table.tableColumns], f"Header schema differs: {table.name}"
            assert len(set(headers)) == len(headers), f"Duplicate header: {table.name}"
            rows = []
            for index in range(top + 1, bottom + 1):
                cells = [sheet.cell(index, column) for column in range(left, right + 1)]
                assert not any(cell.data_type == "e" for cell in cells), f"Excel error: {sheet.title}/{index}"
                row = {header: cell.value for header, cell in zip(headers, cells)}
                if any(value is not None for value in row.values()):
                    rows.append(row)
            result[table.name] = {"sheet": sheet.title, "headers": headers, "rows": rows}
    return result


def unique(rows, field):
    result = {}
    for row in rows:
        value = row[field]
        assert isinstance(value, str) and value.strip(), f"Missing identity: {field}"
        assert value not in result, f"Duplicate {field}: {value}"
        result[value] = row
    return result


def validate(tables, settings):
    cases = unique(tables["tbTestScenarios"]["rows"], "ScenarioId")
    steps = unique(tables["tbTestSteps"]["rows"], "StepId")
    movement = unique(tables["tbExpectedMovement"]["rows"], "StepId")
    word = unique(tables["tbExpectedWord"]["rows"], "StepId")
    snapshots = tables["tbExpectedMovementRows"]["rows"]
    assert not {"ExpectedDecision", "ExpectedCode", "ExpectedOpenCount"}.intersection(tables["tbTestSteps"]["headers"])
    for kind in ("cases", "steps", "history", "absence", "expectedMovement", "expectedRows", "expectedWord"):
        table = tables[settings[f"test.scenarios.{kind}.table"]]
        assert table["sheet"] == settings[f"test.scenarios.{kind}.sheet"], f"Wrong sheet for {kind}"
    for kind, key in (("movement", "tbExpectedMovement"), ("word", "tbExpectedWord")):
        assert tables[key]["headers"] == settings[f"test.expected.{kind}Headers"].split("|")
    native = tables["tbTestHistory"]["headers"][3:]
    helpers = settings["test.scenarios.formulaColumns"].split("|")
    assert tables["tbExpectedMovementRows"]["headers"] == ["StepId", "RecordId"] + [header for header in native if header not in helpers]
    for kind in ("results", "inputs", "records"):
        expected = settings[f"test.audit.{kind[:-1] if kind != 'records' else 'record'}Fields"].split("|")
        if kind == "records":
            expected += native
        assert tables[settings[f"test.audit.{kind}.table"]]["headers"] == expected, f"Archive schema differs: {kind}"
    for scenario_id, case in cases.items():
        assert case["RuleStatus"] in ("Defined", "PendingRule")
        orders = [row["StepOrder"] for row in steps.values() if row["ScenarioId"] == scenario_id]
        assert orders and sorted(orders) == list(range(1, len(orders) + 1)), f"Invalid StepOrder: {scenario_id}"
        assert any(row["ScenarioId"] == scenario_id for row in tables["tbTestAbsence"]["rows"]), f"Missing Absence fixture: {scenario_id}"
    for step_id, step in steps.items():
        assert step["ScenarioId"] in cases, f"Unknown scenario: {step_id}"
        assert step["ІПН"] == cases[step["ScenarioId"]]["ІПН"], f"Wrong participant: {step_id}"
        assert step["Action"] in ("ExportMovement", "PreviewWord"), f"Unknown action: {step_id}"
        expected = movement if step["Action"] == "ExportMovement" else word
        assert step_id in expected, f"Missing expectation: {step_id}"
        assert step["Подія"] == settings[f"export.events.{step['ModeId']}"]
    for step_id, expectation in movement.items():
        assert step_id in steps and steps[step_id]["Action"] == "ExportMovement", f"Orphan Movement expectation: {step_id}"
        decision = expectation["ExpectedDecision"]
        assert decision in ("Allow", "Block", "NoChange", "Pending")
        assert expectation["ExpectedCode"], f"Missing code: {step_id}"
        state_check = "Snapshot" if decision == "Allow" else "Pending" if decision == "Pending" else "Unchanged"
        assert expectation["StateCheck"] == state_check, f"Incompatible state check: {step_id}"
        rows = [row for row in snapshots if row["StepId"] == step_id]
        if state_check == "Snapshot":
            assert rows, f"Missing snapshot: {step_id}"
            assert len({row["RecordId"] for row in rows}) == len(rows)
            keys = settings["test.expected.rowKeys"].split("|")
            assert len({tuple(row[field] for field in keys) for row in rows}) == len(rows), f"Duplicate snapshot identity: {step_id}"
            assert all(row["ІПН"] == steps[step_id]["ІПН"] for row in rows)
            assert sum(row["Прибуття"] is None for row in rows) == expectation["ExpectedOpenCount"], f"Wrong expected open count: {step_id}"
        else:
            assert not rows, f"Unexpected snapshot: {step_id}"
        if decision != "Pending":
            count = expectation["ExpectedOpenCount"]
            assert isinstance(count, (int, float)) and count >= 0 and count == int(count), f"Invalid open count: {step_id}"
    assert all(row["StepId"] in movement for row in snapshots), "Orphan expected row"
    for step_id, expectation in word.items():
        assert step_id in steps and steps[step_id]["Action"] == "PreviewWord", f"Orphan Word expectation: {step_id}"
        assert expectation["Outcome"] in ("Success", "Error", "Pending")
        if expectation["Outcome"] == "Success":
            text = expectation["ExpectedText"]
            assert isinstance(text, str) and 0 < len(text) <= 32767 and not expectation["ExpectedError"], f"Missing Word text: {step_id}"
        elif expectation["Outcome"] == "Error":
            assert expectation["ExpectedError"] and not expectation["ExpectedText"], f"Missing Word diagnostic: {step_id}"
    return len(cases), len(steps), len(snapshots), len(word)


def negative_checks(tables, settings):
    def reject(change):
        altered = copy.deepcopy(tables)
        change(altered)
        try:
            validate(altered, settings)
        except AssertionError:
            return
        raise AssertionError("Invalid fixture was accepted")

    reject(lambda data: data["tbTestSteps"]["rows"][1].update(StepId=data["tbTestSteps"]["rows"][0]["StepId"]))
    reject(lambda data: data["tbExpectedMovement"]["rows"][0].update(StepId="UNKNOWN"))
    reject(lambda data: data["tbExpectedMovement"]["rows"].pop(0))
    reject(lambda data: data["tbTestSteps"]["rows"][0].update(StepOrder=5))
    reject(lambda data: data["tbExpectedWord"]["rows"][0].update(ExpectedText=None))
    reject(lambda data: data["tbExpectedMovementRows"]["rows"].append(copy.deepcopy(data["tbExpectedMovementRows"]["rows"][0])))
    reject(lambda data: data["tbExpectedMovementRows"]["rows"][0].update(StepId="UNKNOWN"))
    reject(lambda data: data["tbExpectedMovementRows"]["rows"][0].update(ІПН="UNKNOWN"))
    reordered = copy.deepcopy(tables)
    for table in reordered.values():
        table["rows"].reverse()
    assert validate(reordered, settings) == validate(tables, settings), "Row sorting changed stable links"
    print("PASS: 8 invalid fixture cases rejected; sorting preserves links.")


def source_checks(settings):
    folder = ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder"
    loader = (folder / "obj_PEB_TestScenarios.cls.vba").read_text(encoding="utf-8-sig")
    runner = (folder / "obj_PEB_TestRunner.cls.vba").read_text(encoding="utf-8-sig")
    catalog = (folder / "ex_PEB_Testing.vba").read_text(encoding="utf-8-sig")
    assert 'private_LoadExpectations allSteps' in loader
    assert 'private_Value(steps, stepRow, "expected")' not in loader
    assert 'private_Value(steps, stepRow, "code")' not in loader
    assert 'private_Value(steps, stepRow, "openCount")' not in loader
    assert 'private_ValidateRows(after, testStep, checkDiagnostic)' in runner
    assert 'private_ValidateWord(binding, testStep, report, values)' in runner
    assert 'VBA.StrComp(expected, actual, VBA.vbBinaryCompare)' in runner
    assert 'testStep("Order") = 1' in runner and 'testStep("Id") = 1' not in runner
    assert 'test.scenarios.headers.stepId' in catalog and 'TestPreviewOperation' in catalog
    for filename in ("obj_PEB_TestScenarios.cls.vba", "obj_PEB_TestRunner.cls.vba", "ex_PEB_Testing.vba", "obj_PEB_TestUi.cls.vba"):
        text = (folder / filename).read_text(encoding="utf-8-sig")
        procedures = re.findall(r"(?im)^(?:Public|Private|Friend) (?:Function|Sub) (\w+)", text)
        assert len(procedures) == len(set(procedures)), f"Duplicate procedure: {filename}"
        assert len(procedures) == len(re.findall(r"(?im)^End (?:Function|Sub)$", text)), f"Unbalanced procedures: {filename}"
        for name, body in re.findall(r"(?ms)^(?:Public|Private|Friend) (?:Function|Sub) (\w+)(.*?)(?=^End (?:Function|Sub)$)", text):
            declarations = re.findall(r"(?im)^\s*Dim (\w+)\b", body)
            assert len(declarations) == len(set(declarations)), f"Duplicate variable: {filename}/{name}"
        for key in re.findall(r'fn_Config\("([^"]+)"\)', text):
            assert key in settings, f"Missing configuration: {filename}/{key}"
    print("PASS: VBA source/configuration contracts and procedure structure.")


def directory_checks(tables, settings):
    workbook_root = ROOT / "PROJECTS_EXCEL/workbook/PersonalEventBuilder"
    movement = openpyxl.load_workbook(workbook_root / settings["export.Movement.path"])
    target = movement[settings["export.Movement.sheet"]].tables[settings["export.Movement.table"]]
    assert [column.name for column in target.tableColumns] == tables["tbTestHistory"]["headers"][3:]
    assert settings["environment.name"] == "TEST"
    test_root = (workbook_root / settings["test.scenarios.root"]).resolve()
    for key in ("export.Movement.path", "test.scenarios.path", "lookup.Absence.source.path"):
        assert (workbook_root / settings[key]).resolve().is_relative_to(test_root), f"Source outside TEST: {key}"
    assert (workbook_root / settings["export.Movement.path"]).resolve() != (workbook_root / settings["test.scenarios.productionPath"]).resolve()
    books = {}
    expected = unique(tables["tbExpectedWord"]["rows"], "StepId")
    for step in tables["tbTestSteps"]["rows"]:
        if step["Action"] != "PreviewWord":
            continue
        for prefix in [f"export.Word.lookups.{index}" for index in range(1, int(settings["export.Word.lookups.count"]) + 1)] + [
            f"export.Word.{step['ModeId']}.lookups.{index}"
            for index in range(1, int(settings[f"export.Word.{step['ModeId']}.lookups.count"]) + 1)
        ]:
            input_field = settings[f"{prefix}.input"]
            if input_field.startswith("Report"):
                continue
            header = {"PersonId": "ІПН"}.get(input_field, input_field)
            lookup_path = (workbook_root / settings[f"{prefix}.path"]).resolve()
            if lookup_path not in books:
                books[lookup_path] = openpyxl.load_workbook(lookup_path, data_only=True)
            sheet = books[lookup_path][settings[f"{prefix}.sheet"]]
            left, top, right, bottom = range_boundaries(settings[f"{prefix}.range"])
            headers = [sheet.cell(top, column).value for column in range(left, right + 1)]
            key_column = left + headers.index(settings[f"{prefix}.keyHeader"])
            value_column = left + headers.index(settings[f"{prefix}.valueHeader"])
            matches = [sheet.cell(row, value_column).value for row in range(top + 1, min(bottom, sheet.max_row) + 1)
                       if sheet.cell(row, key_column).value == step[header]]
            if expected[step["StepId"]]["Outcome"] == "Success":
                assert len(matches) == 1 and matches[0], f"Missing preview lookup: {step['StepId']}/{prefix}"
            elif "Treatment.lookups" in prefix:
                assert not matches, f"Negative hospital fixture unexpectedly resolves: {step['StepId']}"
    print("PASS: native Movement schema, TEST routing and all Word fixture directory dependencies.")


def preservation_check(path):
    before = openpyxl.load_workbook(path, data_only=False)
    after = openpyxl.load_workbook(FIXTURE, data_only=False)
    for name in ("Тест_Історія",):
        assert list(before[name].values) == list(after[name].values), f"History changed: {name}"
    for name in ("Тест_Сценарії", "Тест_ЄЖОС"):
        for old_row, new_row in zip(before[name].values, after[name].values):
            assert old_row == new_row, f"Original inputs changed: {name}"
    for name in ("Тест_Результати", "Тест_Записи", "Тест_ВхідніДані"):
        old_sheet = before[name]
        new_sheet = after[name]
        headers = [cell.value for cell in old_sheet[1]]
        scenario_column = headers.index("ScenarioId") + 1
        step_column = headers.index("StepId") + 1
        for row in old_sheet:
            for cell in row:
                expected = cell.value
                scenario = old_sheet.cell(cell.row, scenario_column).value
                if cell.column == step_column and scenario and isinstance(expected, int):
                    expected = f"{scenario}-M{expected:02}"
                assert new_sheet.cell(cell.row, cell.column).value == expected, f"Archive changed: {name}/{cell.coordinate}"
    print("PASS: original inputs, history and archive values preserved; historical StepId migrated.")


if __name__ == "__main__":
    settings = config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder/wsConfig.test.txt")
    production = config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder/wsConfig.txt")
    for key in settings:
        if key.startswith(("test.expected.", "test.scenarios.", "test.audit.")):
            assert settings[key] == production[key], f"PROD/TEST schema differs: {key}"
    tables = load_tables(openpyxl.load_workbook(FIXTURE))
    counts = validate(tables, settings)
    negative_checks(tables, settings)
    source_checks(settings)
    directory_checks(tables, settings)
    if len(sys.argv) > 1:
        preservation_check(Path(sys.argv[1]))
    print(f"PASS: {counts[0]} scenarios, {counts[1]} stable steps, {counts[2]} expected Movement rows, {counts[3]} Word expectations. Office was not started.")