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
    assert 'field <> basisHeader' in resolve
    assert 'fn_FieldsHash(previousFields)' in resolve
    assert 'state(1) = generated' in resolve
    assert 'fields(field) = cells(field)' in resolve
    assert resolve.index('If m_overwrite Then') < resolve.index('validHash =')
    assert 'fn_Resolve = True\n        Exit Function' in resolve
    assert 'fn_Save' not in service and 'fn_SetStatus' not in service
    planner = (folder / "ex_PEB_MovementPlanner.vba").read_text(encoding="utf-8-sig")
    existing = procedure(planner, "private_Existing")
    assert existing.index('fn_Pending(identifier)') < existing.index('row("Cells")(entry("Header")) =')
    revisions = (folder / "ex_PEB_HistoryRevisions.vba").read_text(encoding="utf-8-sig")
    record = procedure(revisions, "fn_RecordExport")
    assert record.index('fn_Record book, actions, target') < record.index('\n    book.Save\n')
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
        assert config['export.Movement.manualChanges.policy'] == 'Overwrite'
        for prefix, count in (('history.fieldState', 8), ('history.conflicts', 12)):
            assert len(config[f'{prefix}.fields'].split('|')) == count
    print('PASS: field conflict fingerprints, basis-only guards, pending-plan barrier, metadata rollback, XML and schemas. Excel/VBA not executed.')


if __name__ == '__main__':
    main()