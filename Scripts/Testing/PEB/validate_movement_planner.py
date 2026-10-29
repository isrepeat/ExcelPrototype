"""Check the preview's source and snapshot contracts without opening Office."""

import re

from event_routes import ROOT, parse_config
from movement_plans import load_native_plans


def procedure(source, name):
    match = re.search(
        rf"(?ims)^(?:Public|Private) (?:Function|Sub) {name}\b.*?^End (?:Function|Sub)$",
        source,
    )
    assert match, name
    return match.group()


def check_declarations(source):
    fields = set(re.findall(r"(?im)^Private (\w+) As ", source))
    for match in re.finditer(
        r"(?ims)^(?:Public|Private) (Function|Sub) (\w+)\b.*?^End \1$", source
    ):
        body = match.group()
        declared = fields | {match[2]}
        declared.update(re.findall(r"(?:Dim |ByVal |ByRef )(\w+) As ", body))
        assigned = set(re.findall(r"(?m)^\s*(?:Set )?(\w+)\s*=", body))
        loops = set(re.findall(r"(?im)^\s*For (?:Each )?(\w+)\b", body))
        assert assigned | loops <= declared, (match[2], (assigned | loops) - declared)


def main():
    folder = ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder"
    preview = (folder / "ex_PEB_MovementPlanner.vba").read_text(encoding="utf-8-sig")
    service = (folder / "obj_PEB_ExportService.cls.vba").read_text(encoding="utf-8-sig")
    prepare = procedure(service, "TryPrepareHistoryMovement")
    check_declarations(preview)
    apply_event = procedure(preview, "private_Apply")
    awol_food = apply_event.split('If values("EventId") = "MoveFromAwolToBusinessTrip" Then', 1)[1].split('End If', 1)[0]
    assert 'closingWrites(ex_PEB_Export.fn_Config("export.Movement.headers.ArrivalFoodSupport")) = eventDate' in awol_food
    assert 'DepartureFoodSupport' not in awol_food
    check_declarations(prepare)

    # These are source-level guards, not evidence of VBA runtime correctness.
    executable = "\n".join(line for line in preview.splitlines() if not line.lstrip().startswith("'"))
    assert not re.search(r"(?i)\.(?:Save|SaveAs|SaveCopyAs|Delete|Clear|Close)\b", executable)
    assert not re.search(r"(?i)\.(?:Value2?|Formula|NumberFormat)\s*=", executable)
    assert not re.search(r"(?i)\.(?:ListRows|Workbooks|Worksheets)\.Add\b", executable)
    assert "private_Enrich" not in prepare
    assert "private_LookupTable" not in prepare
    assert "private_PrepareFoodSupport" not in prepare
    assert 'For Each document In stored("HistoryReports")' in prepare
    assert 'MovementHistoryReportsUnsupported' not in prepare
    assert 'private_PrepareMovementTexts values, Nothing, True' in prepare

    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        config_path = ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder" / filename
        settings = parse_config(config_path)
        assert "{code}" in settings["text.MovementPlanError"]
        assert "{detail}" in settings["text.MovementPlanError"]
        assert settings["text.MovementHistoryReportsUnsupported"]
        saved = {part.split("=", 1)[0] for part in settings["history.fields"].split("|")}
        restored = saved | {"ModeId", "PersonAccusative", "TicketDocument"}
        for index in range(1, int(settings["export.shared.constants.count"]) + 1):
            restored.add(settings[f"export.shared.constants.{index}.name"])
        for index in range(1, int(settings["export.shared.aliases.count"]) + 1):
            prefix = f"export.shared.aliases.{index}"
            required = set(settings[prefix + ".source"].split("|"))
            assert required <= restored, (filename, prefix, required - restored)
            restored.add(settings[prefix + ".name"])
        for key in ("export.dateFields", "export.shared.inputs.dateScalars"):
            assert set(settings[key].split("|")) <= saved, (filename, key)
        plans = load_native_plans(settings)
        assert plans["events"]
        restored.update({"EventName", "SourceEventName", "SourceDestination", "MovementBasis"})
        for key, field in plans["values"].items():
            if key.endswith(".field"):
                assert field in restored, (filename, key, field)
        for line in config_path.read_text(encoding="utf-8-sig").splitlines():
            if line.strip():
                assert len(line.split("\t")) == 3, (filename, line)

    print("PASS: preview declarations, no target writes, saved alias/date providers and TEST/PROD configuration. Office/VBA not executed.")


if __name__ == "__main__":
    main()