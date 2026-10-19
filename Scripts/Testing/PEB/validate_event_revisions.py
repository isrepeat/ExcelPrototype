"""Check event revision, cancellation and export audit contracts without Office."""

import re
from xml.etree import ElementTree as ET

from event_routes import ROOT, parse_config
from validate_movement_batch import check_module_contract
from validate_movement_planner import check_declarations, procedure


def main():
    folder = ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder"
    revisions = (folder / "ex_PEB_HistoryRevisions.vba").read_text(encoding="utf-8-sig")
    history = (folder / "ex_PEB_History.vba").read_text(encoding="utf-8-sig")
    preview = (folder / "ex_PEB_MovementPlanner.vba").read_text(encoding="utf-8-sig")
    sparse = (folder / "ex_PEB_MovementWriter.vba").read_text(encoding="utf-8-sig")
    batch = (folder / "ex_PEB_MovementBatch.vba").read_text(encoding="utf-8-sig")
    controller = (folder / "obj_PEB_PgMainController.cls.vba").read_text(encoding="utf-8-sig")
    page = (folder / "obj_PEB_PgMain.cls.vba").read_text(encoding="utf-8-sig")
    text_keys = set(re.findall(r'"([^"]+)"', re.search(
        r"For Each textKey In VBA.Array\((.*?)\)\s*\n", page, re.S
    ).group(1)))
    command_keys = set(re.findall(r'"([^"]+)"', re.search(
        r"For Each orderCommand In VBA.Array\((.*?)\)\s*\n", page, re.S
    ).group(1)))
    check_module_contract(revisions, "ex_PEB_HistoryRevisions")
    conflicts = (folder / "ex_PEB_MovementConflicts.vba").read_text(encoding="utf-8-sig")
    check_declarations(conflicts)
    begin = procedure(conflicts, "fn_Begin")
    assert begin.index("If m_overwrite Then") < begin.index("Set book = private_Book()")
    record = procedure(revisions, "fn_RecordExport")
    assert record.index("private_WriteRows table, pending") < record.index("committed = True")
    for source in (history, revisions, conflicts):
        assert "book.Save" not in source
        assert "book.Saved" not in source
        assert "book.Close" not in source
        assert "Workbooks.Open" not in source
        assert "text.HistoryOpenRequired" in source
    for source in (revisions, conflicts):
        write_rows = procedure(source, "private_WriteRows")
        assert "Application.CountA(extension)" in write_rows
        assert "Application.Intersect(extension, other.Range)" in write_rows
        assert "area.Value2 = data" in write_rows
    assert "True, Nothing, False" in procedure(preview, "fn_BuildOrder")
    assert "If readSync Then" in procedure(history, "private_Read")
    check_declarations(history)
    check_declarations(controller)
    for source in (revisions, history, batch):
        assert ".ListColumns.Add" not in source
        assert ".Worksheets.Add" not in source
        assert ".ListObjects.Add" not in source

    save = procedure(history, "fn_Save")
    assert save.index("ExpectedHistoryStamp") < save.index("ex_PEB_HistoryRevisions.fn_Prepare")
    assert save.index("ex_PEB_HistoryRevisions.fn_Prepare") < save.index("target.Range.Value2 = data")
    assert save.index("target.Range.Value2 = data") < save.index("committed = True")
    assert "If Not committed Then" in save
    assert "ex_PEB_HistoryRevisions.fn_Rollback(transaction, rollbackDiagnostic)" in save
    assert "HistoryRollbackError" in save
    assert "ElseIf written Then" in save
    assert "backupFormats(index)" in save
    assert 'history.versions.stateFields' in procedure(history, "fn_Stamp")
    assert "ex_PEB_HistoryRevisions.fn_ReadState" in procedure(history, "private_Read")
    prepare = procedure(revisions, "fn_Prepare")
    assert "ByRef transaction As Object" in prepare
    assert 'revision = revision + 1' in prepare
    for kind in ("Create", "Update", "Baseline", "Cancel", "Restore"):
        assert f'"{kind}"' in prepare
    assert "HistoryReasonRequired" in prepare and "MissingEventRevision" in prepare
    snapshot = procedure(revisions, "private_Append")
    assert "DuplicateEventRevision" in snapshot
    assert 'ids(parent("ParentRowId"))' in snapshot
    assert 'If parent("RowId") = root("RowId") Then' in snapshot
    assert '"Snapshot." & mapping(1)' in snapshot
    assert 'changedFields = changedFields & "Status"' in snapshot

    cancel = procedure(preview, "private_Cancel")
    assert cancel.index("fn_FieldsHash(fields, cells)") < cancel.index('cells(entry("Header")) =')
    assert "AppliedMovementFormulaChanged" in cancel and "private_Matches(cells, joins)" in cancel
    assert '"ClosingId", "ClosingHash"' in cancel
    assert 'CancellationOpeningNotSupported' in cancel
    assert 'selected("ClosingRowId") = VBA.vbNullString' in cancel
    build = procedure(preview, "fn_Build")
    assert build.index("private_Apply(rows") < build.index("private_TryValidatePeriods(rows")
    for source in (sparse,):
        apply = procedure(source, "fn_Apply")
        assert apply.index("fn_RecordExport") < apply.rindex("fn_Apply = True")
        assert "fn_RecordExport" in apply and "private_Restore(target" in apply
    export_word = procedure(controller, "ExportWordHandler")
    assert 'historyItem("Status") = "Active"' in export_word
    assert export_word.index("Set events = activeEvents") < export_word.index("TryGenerateHistory")
    assert export_word.index("TryGenerateHistory") < export_word.index("fn_RecordExport")
    assert "CancelledOmitted" in export_word and "HistoryJournalFailed" in export_word
    assert "ExpectedHistoryStamp" in procedure(history, "fn_SetStatus")

    ui = ET.parse(ROOT / "PROJECTS_EXCEL/ui/PersonalEventBuilder/MainPage.xaml")
    names = {node.get("name") for node in ui.iter() if node.get("name")}
    for command in ("EventVersions", "CancelEvent", "RestoreEvent"):
        assert command in names
        assert command in text_keys, ("Missing Text binding initialization", command)
        assert command in command_keys, ("Missing Commands binding initialization", command)
        assert f"Public Function {command}Handler()" in controller
        assert len("btn_" + command) <= 31
    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        settings = parse_config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder" / filename)
        assert settings["history.versions.stateFields"] == "Revision|Status|LastChangeId"
        assert settings["history.versions.fields"].split("|") == [
            "ChangeId", "EventRowId", "Revision", "ChangedAt", "ChangedBy",
            "ChangeType", "Reason", "ChangedFields", "Status",
        ]
        assert len(settings["history.exports.fields"].split("|")) == 8
        for key in re.findall(r'fn_Config\("((?:history\.(?:versions|exports)|text\.History)[^"]*)"\)', revisions):
            assert key in settings, (filename, key)
        assert len(settings["history.versions.sheet"]) <= 31
        assert settings["history.list.columns.9.field"] == "Status"
        assert settings["history.list.columns.10.field"] == "Revision"
        for command in ("EventVersions", "CancelEvent", "RestoreEvent"):
            assert settings["text." + command]
    print("PASS: version snapshots, stale edits, status reasons, child snapshots, cancellation guards, export journals and UI/config contracts. Office/VBA not executed.")


if __name__ == "__main__":
    main()