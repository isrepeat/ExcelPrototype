"""Validate Movement diff capture, publication, UI and configuration without Office."""
from pathlib import Path
import re
import xml.etree.ElementTree as ET

from event_routes import ROOT, parse_config
from validate_movement_batch import check_blocks, check_module_contract
from validate_movement_planner import check_declarations, procedure


def main():
    folder = ROOT / 'PROJECTS_EXCEL/vba/PersonalEventBuilder'
    files = [
        'ex_PEB_MovementDiff.vba', 'ex_PEB_MovementWriter.vba',
        'ex_PEB_MovementBatch.vba', 'ex_PEB_MovementPlanner.vba',
        'obj_PEB_PgMainController.cls.vba', 'obj_PEB_PgMain.cls.vba',
        'Test/ex_PEB_MovementDiffTests.vba',
    ]
    sources = {name: (folder / name).read_text(encoding='utf-8-sig') for name in files}
    for name, source in sources.items():
        check_declarations(source)
        check_blocks(source)
        assert len(re.search(r'Attribute VB_Name = "([^"]+)"', source)[1]) <= 31
    check_module_contract(sources['ex_PEB_MovementDiff.vba'], 'ex_PEB_MovementDiff')
    check_module_contract(sources['Test/ex_PEB_MovementDiffTests.vba'], 'ex_PEB_MovementDiffTests')

    writer = sources['ex_PEB_MovementWriter.vba']
    apply = procedure(writer, 'fn_Apply')
    capture = procedure(writer, 'private_Capture')
    assert 'entry.Add "Value", cell.Value2' in capture
    assert 'entry.Add "Expected", expected' in capture
    assert 'entry.Add "Action", action' in capture
    assert 'entry.Add "Header", header' in capture
    assert apply.index('MovementDiff.fn_Build') < apply.index('changed = True')
    assert apply.index('private_Verify target, entry, False') < apply.index('If fn_Apply Then')
    assert 'If fn_Apply Then\n        Set exportDiff = pendingDiff' in apply
    assert apply.index('fn_RecordExport') < apply.index('fn_Apply = True')
    assert 'private_Restore(target, entries' in apply

    report = procedure(sources['ex_PEB_MovementDiff.vba'], 'fn_Build')
    assert 'private_Text(entry("Value"),' in report
    assert 'private_Text(entry("Expected"))' in report
    assert 'Not internalHeaders.Exists(entry("Header"))' in report
    assert 'fn_ValueToken(entry("Value")) <>' in report
    assert 'changedEvents(action("RowId")) = True' in report
    assert 'If Not changedEvents.Exists(action("RowId")) Then' in report
    assert 'originalRows = target.ListRows.Count' in report

    controller = sources['obj_PEB_PgMainController.cls.vba']
    export = procedure(controller, 'ExportMovementHandler')
    assert re.search(r'If exported Then\s+Set m_conflicts = Nothing\s+Set m_movementDiff = exportDiff', export)
    assert export.index('If Not exported Then') < export.index('If Not ShowMovementDiffHandler()')
    assert 'Set m_movementDiff = Nothing' not in export
    show = procedure(controller, 'ShowMovementDiffHandler')
    close = procedure(controller, 'CloseMovementDiffHandler')
    assert 'ex_PEB_MovementDiff.fn_Result' in show
    assert '"CandidatesVisibility", "EjosVisibility", "MedicalVisibility", "HistoryVisibility", "ConflictsVisibility"' in show
    assert '"MovementDiffVisibility", "Collapsed"' in close
    assert '"CandidatesVisibility"' not in close
    for command in ['ShowMovementDiff', 'CloseMovementDiff']:
        assert '"' + command + '"' in sources['obj_PEB_PgMain.cls.vba']

    ui = ET.parse(ROOT / 'PROJECTS_EXCEL/ui/PersonalEventBuilder/MainPage.xaml').getroot()
    nodes = {node.get('name'): node for node in ui.iter() if node.get('name')}
    assert nodes['MovementDiff'].get('itemsSource') == '{Binding Path=Data.MovementDiff}'
    assert nodes['MovementDiff'].get('closeCommand') == '{Binding Path=Commands.CloseMovementDiff}'
    assert nodes['MovementDiff'].get('closeCaption') == '{Binding Path=Text.CloseTable}'
    assert nodes['ShowMovementDiff'].get('visibility') == '{Binding Path=Form.MovementDiffAvailableVisibility}'
    for name in ['MovementDiff', 'MovementDiffSummary', 'ShowMovementDiff']:
        assert len('btn_' + name) <= 31

    settings = []
    for filename in ['wsConfig.txt', 'wsConfig.test.txt']:
        path = ROOT / 'PROJECTS_EXCEL/config/PersonalEventBuilder' / filename
        config = parse_config(path)
        labels = {key: value for key, value in config.items() if key.startswith('movementDiff.')}
        assert all(labels.values())
        for field in config['movementDiff.fields'].split('|'):
            assert config['movementDiff.columns.' + field]
        for placeholder in ['time', 'order', 'date', 'added', 'updated', 'unchanged']:
            assert '{' + placeholder + '}' in config['movementDiff.summary']
        settings.append(labels)
        for line in path.read_text(encoding='utf-8-sig').splitlines():
            assert not line.strip() or len(line.split('\t')) == 3
    assert settings[0] == settings[1]
    print('PASS: actual pre-write capture, unchanged/internal-field suppression, success-only publication, previous-report retention, navigation, PROD/TEST and VBA source contracts. Office/VBA not executed.')


if __name__ == '__main__':
    main()