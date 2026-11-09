"""Verify conflict integration contracts without starting Excel or running VBA."""

from xml.etree import ElementTree as ET

from event_routes import ROOT, parse_config
from validate_movement_batch import check_blocks, check_module_sections
from validate_movement_planner import check_declarations, procedure


def main():
    folder = ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder"
    service = (folder / "ex_PEB_MovementConflicts.vba").read_text(encoding="utf-8-sig")
    check_module_sections(service)
    check_declarations(service)
    check_blocks(service)
    resolve = procedure(service, "fn_Resolve")
    same = procedure(service, "private_Same")
    for key in ("Revision", "HistoryStamp", "Target", "GeneratedValue", "MovementValue", "LinkHash"):
        assert f'"{key}"' in same
    assert 'field <> basisHeader' not in resolve
    assert 'actual <> state(2) And actual <> generated' in resolve
    assert 'manual = (actual <> generated And VBA.Len(ex_PEB_Export.fn_Text(cells(field))) > 0)' in resolve
    assert resolve.rstrip().endswith('Next field\nEnd Function')
    assert 'AppliedMovementChanged' not in resolve and 'AppliedEventChanged' not in resolve
    assert 'state(1) = generated' in resolve
    assert 'fields(field) = cells(field)' in resolve
    assert resolve.index('If m_overwrite Then') < resolve.index('validHash =')
    assert 'fn_Resolve = True\n        Exit Function' in resolve
    assert 'fn_Save' not in service and 'fn_SetStatus' not in service
    planner = (folder / "ex_PEB_MovementPlanner.vba").read_text(encoding="utf-8-sig")
    existing = procedure(planner, "private_Existing")
    assert existing.index('fn_Pending(identifier)') < existing.index('row("Cells")(entry("Header")) =')
    side = procedure(planner, "private_PlanBasisSide")
    assert 'fn_Resolve action, side, fields, cells' in side
    assert 'AppliedEventChanged' not in side
    assert 'fn_CellValueToken(cells(key))' in side
    tvo = (folder / "ex_PEB_MovementTvo.vba").read_text(encoding="utf-8-sig")
    check_declarations(tvo)
    check_blocks(tvo)
    prepare = procedure(tvo, "fn_Prepare")
    assert prepare.index('private_ResolveFields row, original, action') < prepare.index('private_ValidateResolved model')
    assert prepare.index('fn_Pending(CStr(action("RowId")))') < prepare.index('private_ValidateResolved model')
    assert 'TvoFieldGroups' in procedure(service, 'fn_Record')
    assert 'TVO.Structure' in procedure(tvo, 'private_ResolveStructure')
    revisions = (folder / "ex_PEB_HistoryRevisions.vba").read_text(encoding="utf-8-sig")
    record = procedure(revisions, "fn_RecordExport")
    assert record.index('fn_Record book, actions, target') < record.index('committed = True')
    for table in ('fieldTable', 'conflictTable'):
        assert f'{table}.ListRows({table}.ListRows.Count).Delete' in record
    ui = ET.parse(ROOT / "PROJECTS_EXCEL/ui/PersonalEventBuilder/MainPage.xaml")
    conflict = next(node for node in ui.iter() if node.get('name') == 'MovementConflicts')
    assert conflict.get('selectionMode') == 'Confirm'
    assert conflict.get('selectedColumn') == '{Binding Path=Form.ConflictColumn}'
    element = (ROOT / 'PROJECTS_EXCEL/vba/common/[4] elements/obj_UiControlElement.cls.vba').read_text(encoding='utf-8-sig')
    assert 'attributeNode.nodeName <> "selectedItem" And attributeNode.nodeName <> "selectedColumn"' in element
    for filename in ('wsConfig.txt', 'wsConfig.test.txt'):
        config = parse_config(ROOT / 'PROJECTS_EXCEL/config/PersonalEventBuilder' / filename)
        assert config['export.Movement.manualChanges.policy'] == 'Resolve'
        columns = [value.split("=") for value in config["conflicts.columns"].split("|")]
        assert [name for name, span in columns] == [
            "PersonName", "EventId", "EventRowId", "Side", "Field", "MovementValue", "GeneratedValue"
        ]
        assert sum(int(span) for name, span in columns) == 18
        for name, span in columns:
            assert config[f"text.ConflictColumn.{name}"]
        for date_field in config["conflicts.dateFields"].split("|"):
            assert config[date_field]
        for prefix, count in (('history.fieldState', 8), ('history.conflicts', 12)):
            assert len(config[f'{prefix}.fields'].split('|')) == count
    controller = (folder / "obj_PEB_PgMainController.cls.vba").read_text(encoding="utf-8-sig")
    publish = procedure(controller, "private_PublishConflicts")
    select = procedure(controller, "SelectConflictHandler")
    assert 'fields("PersonName") = item("PersonName")' in publish
    assert 'fields("Field") = item("Field")' in publish
    assert 'conflicts.columns' in publish and 'conflicts.columns' in select
    assert 'choice = "GeneratedValue"' in select
    assert 'private_PublishConflicts()' not in select
    assert 'private_PaintConflictChoices(CLng(record.Key))' in select
    assert 'ConflictSelectedPrefix' not in publish
    assert publish.index('Application.ScreenUpdating = False') < publish.index('m_pageBase.FlushUpdates()')
    assert publish.index('private_PaintConflictChoices()') < publish.index('Cleanup:')
    assert publish.count('Application.ScreenUpdating = previousScreenUpdating') == 2
    assert publish.index('ClearOverlaysByPrefix "ConflictChoice"') < publish.index('SetObject("Data", "MovementConflicts"')
    assert 'ClearOverlaysByPrefix "MovementConflicts"' in publish
    styles = (ROOT / 'PROJECTS_EXCEL/vba/common/[2] classes/obj_UiStyleCatalog.cls.vba').read_text(encoding='utf-8-sig')
    clear = procedure(styles, 'ClearOverlaysByPrefix')
    assert clear.index('m_overlays.Remove key') < clear.index('Me.RestoreControlArea area')
    paint = procedure(controller, "private_PaintConflictChoices")
    assert 'selectedIndex = 0 Or selectedIndex = rowIndex' in paint
    assert 'Styles.ApplyOverlay cells' in paint
    assert '.Value2' not in paint and 'SetObject' not in paint
    assert conflict.get('selectedStyle') == 'conflictChosen'
    formatter = procedure(controller, "private_ConflictValue")
    assert 'conflicts.dateFields' in formatter
    assert 'fn_OrderDate(value)' in formatter and 'Format$(CDate(dateKey)' in formatter
    assert 'private_ConflictValue(item("MovementValue")' in publish
    assert 'private_ConflictValue(item("GeneratedValue")' in publish
    print('PASS: all-field conflict choices, pending-plan barrier, metadata rollback, XML and schemas. Excel/VBA not executed.')


if __name__ == '__main__':
    main()