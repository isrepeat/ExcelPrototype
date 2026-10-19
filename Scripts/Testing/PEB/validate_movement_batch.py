"""Check batch source safeguards and configuration without executing Office/VBA."""

import re

from event_routes import ROOT, parse_config
from validate_movement_planner import check_declarations, procedure


def check_module_sections(source):
    lines = source.splitlines()
    namespace = None
    lifecycle = 0
    for index, line in enumerate(lines):
        opening = re.fullmatch(r"' namespace (\w+) \{", line)
        closing = re.fullmatch(r"' } // namespace (\w+)", line)
        if opening or closing:
            assert index > 0 and index + 1 < len(lines), (index + 1, line)
            assert re.fullmatch(r"' -+", lines[index - 1]), (index + 1, "preceding separator")
            assert re.fullmatch(r"' -+", lines[index + 1]), (index + 1, "following separator")
            if opening:
                assert namespace is None, (index + 1, "nested namespace")
                namespace = opening[1]
            else:
                assert namespace == closing[1], (index + 1, "unmatched namespace")
                namespace = None
        procedure = re.match(r"(Public|Private) (?:Sub|Function) (\w+)", line)
        if procedure:
            assert namespace is not None, (procedure[2], "outside namespace")
            if procedure[2] == "fn_Module_Dispose":
                assert namespace == "Lifecycle"
                lifecycle += 1
            elif procedure[1] == "Private":
                assert namespace == "Private", procedure[2]
            else:
                assert namespace not in ("Private", "Lifecycle"), procedure[2]
    assert namespace is None
    assert lifecycle == 1, "Expected one Lifecycle/fn_Module_Dispose contract"


def check_blocks(source):
    code = re.sub(r'"(?:[^"]|"")*"', '""', source)
    code = "\n".join(line.split("'", 1)[0] for line in code.splitlines())
    code = re.sub(r" _\s*\n\s*", " ", code)
    stack = []
    for line in code.splitlines():
        statement = line.strip().casefold()
        if re.match(r"if .+ then$", statement):
            stack.append("if")
        elif statement.startswith("for "):
            stack.append("for")
        elif statement.startswith("select case "):
            stack.append("select")
        elif statement == "end if":
            assert stack and stack.pop() == "if", line
        elif statement.startswith("next ") or statement == "next":
            assert stack and stack.pop() == "for", line
        elif statement == "end select":
            assert stack and stack.pop() == "select", line
        elif statement in ("end function", "end sub"):
            assert not stack, (line, stack)
    assert not stack


def check_module_contract(source, name):
    check_module_sections(source)
    assert source.startswith(f'Attribute VB_Name = "{name}"\nOption Explicit\n')
    assert len(name) <= 40
    assert source.index("' namespace Lifecycle {") < source.index("' namespace API {")
    lifecycle = source.split("' namespace Lifecycle {", 1)[1].split("' } // namespace Lifecycle", 1)[0]
    assert re.search(r"(?m)^Public Sub fn_Module_Dispose\(\)\s*\nEnd Sub$", lifecycle)
    assert "' namespace API {" in source and "' } // namespace API" in source
    assert not re.search(r"(?m)^\s*(?:If|ElseIf) .+ Then .+$", source)
    check_declarations(source)
    check_blocks(source)
    names = re.findall(r"(?im)^(?:Public|Private) (?:Function|Sub) (\w+)", source)
    assert len(names) == len(set(names))
    assert len(names) == len(re.findall(r"(?im)^End (?:Function|Sub)$", source))
    for match in re.finditer(r"(?ims)^Public (?:Function|Sub) (\w+)\b.*?^End (?:Function|Sub)$", source):
        assert match[1].startswith("fn_"), match[1]
        if match[1] == "fn_Module_Dispose":
            assert match.group() in lifecycle
            continue
        assert "' namespace API {" in source[:match.start()]
        assert "' } // namespace API" not in source[:match.start()]
    for match in re.finditer(r"(?ims)^Private (?:Function|Sub) (\w+)\b.*?^End (?:Function|Sub)$", source):
        assert match[1].startswith("private_"), match[1]
        assert "' namespace Private {" in source[:match.start()]


def main():
    folder = ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder"
    preview = (folder / "ex_PEB_MovementPlanner.vba").read_text(encoding="utf-8-sig")
    links = (folder / "ex_PEB_MovementLinks.vba").read_text(encoding="utf-8-sig")
    batch = (folder / "ex_PEB_MovementBatch.vba").read_text(encoding="utf-8-sig")
    closing = (folder / "ex_PEB_MovementWriter.vba").read_text(encoding="utf-8-sig")
    controller = (folder / "obj_PEB_PgMainController.cls.vba").read_text(encoding="utf-8-sig")
    for name, source in (("ex_PEB_MovementPlanner", preview), ("ex_PEB_MovementLinks", links), ("ex_PEB_MovementBatch", batch)):
        check_module_contract(source, name)
    check_module_contract(closing, "ex_PEB_MovementWriter")

    export = procedure(batch, "fn_ExportOrder")
    assert export.count("ex_PEB_MovementPlanner.fn_BuildOrder(") == 1
    assert "ex_PEB_MovementWriter.fn_Apply(plan, target, diagnostic)" in export
    assert "private_UseSparseWriter" not in batch
    assert "Movement.Export.TotalIncludingCalculation" in export
    apply = procedure(closing, "fn_Apply")
    assert "fn_BuildOrder" not in closing and "fn_TargetStamp" not in closing
    assert "fn_StructureStamp(target)" in apply and "fn_PlanStamp(plan)" in apply
    assert apply.index("private_Capture target") < apply.index("changed = True")
    assert apply.index("settingsCaptured = True") < apply.index("Application.EnableEvents = False")
    assert apply.index("Application.Calculation = xlCalculationManual") < apply.index("changed = True")
    assert apply.index("Write.Verify") > apply.index("cell.Value2 =")
    assert apply.index("Write.RestoreSettings") < apply.index("Application.Calculation = previousCalculation")
    assert apply.index("Application.Calculation = previousCalculation") < apply.index("Application.EnableEvents = previousEvents")
    assert apply.index("Application.EnableEvents = previousEvents") < apply.index("Application.ScreenUpdating = previousScreen")
    assert 'If action("Operation") <> "append" Then' in apply
    assert 'If action("Operation") <> "close" Then' in apply
    assert 'For Each update In action("UpdateCells")' in apply
    assert 'True, update("Before")' in apply
    assert 'target.Resize target.Range.Resize(target.Range.Rows.Count + addedRows)' in apply
    assert 'fn_ReadRange(extension)' in apply
    assert 'fn_ReadRange(target.Range)' not in closing
    capture = procedure(closing, "private_Capture")
    for required in ("ConflictingPlannedWrite", "cell.HasFormula", "fn_ValueToken(original)", "TargetChangedBeforeWrite"):
        assert required in capture
    assert 'entry("Expected") = expected' in capture
    assert 'seen.Add cell.Address, entry' in capture
    assert 'entry.Add "Format", cell.NumberFormat' in capture
    verify = procedure(closing, "private_Verify")
    assert "fn_ValueToken(expected)" in verify and "fn_FieldsHash" not in verify
    restore = procedure(closing, "private_Restore")
    for required in ("target.Resize", 'cell.Value2 = entry("Value")', 'cell.NumberFormat = entry("Format")',
                     "private_Verify target, entry, True", "extensionSnapshot", "area.ClearContents"):
        assert required in restore
    extension = procedure(closing, "private_Extension")
    for required in ("AppendWithTotalsUnsupported", "OccupiedExpansionCell", "ExpansionIntersectsTable", "MergedExpansionCell"):
        assert required in extension
    assert not re.search(r"(?i)\\.(?:Save|SaveAs|SaveCopyAs|Saved|AutoSaveOn)\\b", batch + closing)
    assert "Public Function fn_Apply" not in batch
    assert "If changed Then" in apply and "Resume Cleanup" in apply
    assert 'If entries.Count = 0 Then' in apply

    existing = procedure(preview, "private_Existing")
    assert "IncompleteTransitionLinks" in existing
    side = procedure(preview, "private_PlanBasisSide")
    assert "AppliedEventChanged" in side and "AppliedMovementChanged" in side
    assert 'action("Status") = "AlreadyApplied"' in existing
    assert 'links("OpeningHash")' in existing and 'links("ClosingHash")' in existing
    assert "AppliedMovementFormulaChanged" in side
    assert side.index("fn_FieldsHash(fields, cells)") < side.index('private_UpdateCell(row("TargetRow"), VBA.CStr(key)')
    assert 'If Not ex_PEB_MovementConflicts.fn_Overwrite() Then' in side
    assert "basisHeader, VBA.vbTextCompare" in side
    assert 'private_UpdateCell(row("TargetRow"), hashHeader' in side
    assert 'action("IsUpdate") = (updates.Count > 0)' in existing
    assert existing.index("private_PlanBasisSide closing") < existing.index('row("Cells")(entry("Header")) =')
    assert "export.Movement.headers.DepartureBasis" in existing
    assert "export.Movement.headers.ArrivalBasis" in existing
    build = procedure(preview, "fn_Build")
    assert build.index("private_Apply(rows") < build.index("private_TryValidatePeriods(rows")
    assert build.index("private_TryValidatePeriods(rows") < build.index('plan.Add "IsReady"')
    simulate = procedure(preview, "private_Apply")
    assert simulate.index("private_Existing(") < simulate.index("ExistingIdentityRequiresLink")
    read = procedure(preview, "private_Read")
    assert "IncompleteLinkSchema" in read and "FormulaInLinkColumn" in read
    assert "DuplicateEventLink" in procedure(preview, "private_ValidateLink")
    assert 'stamp = "2|"' in links and "MetaDocumentTables" in links
    assert "EventStampTooLong" not in links
    assert "fn_EventStamp = fn_Hash(stamp)" in links
    assert 'private_PlanBasisSide opening, writes' in existing
    assert 'private_PlanBasisSide closing, closingWrites' in existing
    assert '"RowIndex", "Header", "Before", "After"' in procedure(links, "fn_PlanStamp")
    assert 'Case VBA.vbDate, VBA.vbByte' in procedure(links, "fn_ValueToken")
    assert '"H1-"' in procedure(links, "fn_Hash")
    assert '"ClosingIndex", "TargetIndex"' in links
    assert '"Writes", "ClosingWrites", "Joins"' in links
    assert "ex_PEB_EventRoutes.fn_Module_Dispose" in preview

    handler = procedure(controller, "ExportMovementHandler")
    assert "ex_PEB_MovementBatch.fn_ExportOrder" in handler
    assert 'private_Export("Movement")' not in handler
    assert '"AcceptedOrderNumber"' in handler and '"AcceptedOrderDate"' in handler
    assert "ex_PEB_Testing.fn_IsRunning()" in handler
    assert 'fn_Config("export.Movement.path")' not in batch
    target = procedure(batch, "private_Target")
    assert "Application.Workbooks" in target and "book.Worksheets" in target
    assert "sheet.ListObjects" in target
    assert "sheet.Name, sheetName" in target and "table.Name, tableName" in target
    assert "MovementBatchTargetMissing" in target and "MovementBatchTargetAmbiguous" in target
    validate_target = procedure(batch, "private_ValidateTarget")
    assert "Set selected = private_Target()" in validate_target
    assert "If Not target Is selected Then" in validate_target
    assert "OccupiedExpansionCell" in closing and "ExpansionIntersectsTable" in closing
    assert "MOVEMENT_TARGET | FullName=" in batch
    assert "MOVEMENT_STAGE_STARTED" in batch and "fn_Diagnostic_WriteError" in batch
    assert "fn_Diagnostic_Flush(True)" in batch
    assert "private_DiskStamp" not in batch and '"DiskStamp"' not in batch
    assert "TargetRequiresLocalPath" not in batch and "TargetFileNotFound" not in batch
    assert "FileLen" not in batch and "FileDateTime" not in batch
    assert "book.AutoSaveOn" not in validate_target
    assert "book.MultiUserEditing" in validate_target
    assert "cell.Value2" not in read
    assert 'values = snapshot("Values")' in read
    assert "Set rows = private_Read(target, headers, people, eventPeople)" in preview
    assert read.index('private_Fail "AppliedPersonChanged"') < read.index("If Not people.Exists")
    assert 'personValues = target.ListColumns(personHeader).DataBodyRange.Value2' in read
    assert read.index('If Not people.Exists') < read.index('Set row = VBA.CreateObject')
    assert 'fn_ReadRange(target.DataBodyRange.Rows(index))' in read
    assert 'row.Add "SourceRow", index' in read
    assert 'row.Add "TargetRow", index' in read
    calculate = procedure(preview, "private_Apply")
    assert 'action("ClosingIndex") = row("TargetRow")' in calculate
    assert 'row.Add "TargetRow", nextTargetIndex + 1' in calculate
    assert 'action("TargetIndex") = nextTargetIndex' in calculate
    assert 'action("TargetIndex") = candidate.Count' not in calculate
    assert read.index('private_ValidateLink linkCells') < read.index('If Not people.Exists')
    bulk = procedure(links, "fn_ReadRange")
    assert "values = area.Value2" in bulk and "formulas = area.Formula" in bulk
    assert "xlCellTypeFormulas" in bulk and "cell.Value2" not in bulk
    assert "area.NumberFormat" in procedure(links, "private_ReadFormats")
    assert "Public Function fn_RangeStamp" not in links
    assert "Public Function fn_TargetStamp" not in links
    core = (ROOT / "PROJECTS_EXCEL/vba/common/modules/ex_Core.vba").read_text(encoding="utf-8-sig")
    assert "forceBufferedLines As Boolean = False" in procedure(core, "fn_Diagnostic_Flush")

    profiles = []
    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        settings = parse_config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder" / filename)
        headers = [settings["export.Movement.links." + name] for name in ("OpeningId", "OpeningHash", "ClosingId", "ClosingHash")]
        assert len({header.casefold() for header in headers}) == 4
        business = {value.casefold() for key, value in settings.items() if key.startswith("export.Movement.headers.")}
        assert not {header.casefold() for header in headers} & business
        assert all(0 < len(header) <= 255 for header in headers)
        assert "export.Movement.backup.path" not in settings
        assert settings["export.Movement.dateFormat"]
        for key in re.findall(r'fn_Config\("((?:text\.MovementBatch|export\.Movement\.(?:backup|dateFormat))[^"\r\n]*)"\)', batch):
            assert key in settings, (filename, key)
        profiles.append(headers)
    assert profiles[0] == profiles[1]
    print("PASS: batch source contracts, link schema, direct-table writes without save/backup, rollback guards and accepted-order command wiring. Office/VBA not executed.")


if __name__ == "__main__":
    main()