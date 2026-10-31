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
    assert 'entry.Add "RowHeight", cell.EntireRow.RowHeight' in capture
    assert apply.index('private_AutoFitRows target, entries') > apply.index('private_FormatEntries target, entries, False')
    assert apply.index('private_AutoFitRows target, entries') < apply.rindex('fn_RecordExport')
    assert 'target.Parent.Rows(CLng(key)).AutoFit' in procedure(writer, 'private_AutoFitRows')
    assert 'cell.EntireRow.RowHeight = entry("RowHeight")' in procedure(writer, 'private_Restore')
    assert apply.index('Set previousWindow = Application.ActiveWindow') < apply.index('Application.ScreenUpdating = False')
    assert apply.index('previousWindow.Activate') < apply.index('Application.EnableEvents = previousEvents')
    assert apply.index('previousSheet.Activate') < apply.index('Application.ScreenUpdating = previousScreen')
    assert 'entry.Add "Value", cell.Value2' in capture
    assert 'entry.Add "Expected", expected' in capture
    assert 'entry.Add "Action", action' in capture
    assert 'entry.Add "Header", header' in capture
    assert apply.index('MovementDiff.fn_Build') < apply.index('changed = True')
    assert apply.index('private_Verify target, entry, False') < apply.index('If fn_Apply Then')
    assert 'If fn_Apply Then\n        Set exportDiff = pendingDiff' in apply
    assert apply.index('fn_RecordExport') < apply.index('fn_Apply = True')
    assert 'private_Restore(target, entries' in apply
    assert apply.index('MovementDiff.fn_Seal pendingDiff') < apply.index('fn_Apply = True')

    diff_source = sources['ex_PEB_MovementDiff.vba']
    undo = procedure(diff_source, 'fn_ApplyUndo')
    assert undo.index('private_CheckJournal journal, undoCount') < undo.index('changed = True')
    assert undo.index('row("UndoStamp") <> private_RowStamp') < undo.index('changed = True')
    assert undo.index('cell.HasFormula') < undo.index('changed = True')
    assert undo.index('private_Write cell, row("UndoBefore")') < undo.index('fn_ApplyUndo = True')
    assert undo.index('journalRow.Range.Cells(1, column).Value2') < undo.index('fn_ApplyUndo = True')
    assert 'private_Write entry("Cell"), entry("Value"), entry("Format")' in undo
    assert 'journal.ListRows(journal.ListRows.Count).Delete' in undo
    assert 'history.conflicts' in undo and 'DiffRevert' in undo
    assert 'history.fieldState' not in undo
    assert 'History.fn_' not in undo and '.Save' not in undo and '.Close' not in undo
    assert 'Application.EnableEvents = previousEvents' in undo
    assert 'Application.ScreenUpdating = previousScreen' in undo
    toggle = procedure(diff_source, 'fn_Toggle')
    assert 'movementDiff.columns.DiffId' in toggle
    assert 'If row("UndoApplied") Then' in toggle
    assert 'row("UndoSelected") = Not row("UndoSelected")' in toggle
    assert '"DiffTags"' in procedure(diff_source, 'fn_Result')

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
    for command in ['ShowMovementDiff', 'CloseMovementDiff', 'MovementDiffAbsence', 'MovementDiffTvo', 'ToggleDiff', 'ApplyDiffEdits']:
        assert '"' + command + '"' in sources['obj_PEB_PgMain.cls.vba']

    ui = ET.parse(ROOT / 'PROJECTS_EXCEL/ui/PersonalEventBuilder/MainPage.xaml').getroot()
    nodes = {node.get('name'): node for node in ui.iter() if node.get('name')}
    assert nodes['MovementDiff'].get('itemsSource') == '{Binding Path=Data.MovementDiff}'
    assert nodes['MovementDiff'].get('closeCommand') == '{Binding Path=Commands.CloseMovementDiff}'
    assert nodes['MovementDiff'].get('closeCaption') == '{Binding Path=Text.CloseTable}'
    assert nodes['ShowMovementDiff'].get('visibility') == '{Binding Path=Form.MovementDiffAvailableVisibility}'
    for name in ('MovementDiffAbsence', 'MovementDiffTvo'):
        assert nodes[name].get('command') == '{Binding Path=Commands.' + name + '}'
        assert nodes[name].get('caption') == '{Binding Path=Form.' + name + 'Name}'
        assert nodes[name].get('visibility') is None
        assert len('btn_' + name) <= 31
    assert nodes['MovementDiffTabs'].get('visibility') == '{Binding Path=Form.MovementDiffVisibility}'
    assert nodes['ApplyDiffEdits'].get('visibility') == '{Binding Path=Form.DiffEditsVisibility}'
    assert nodes['ApplyDiffEdits'].get('command') == '{Binding Path=Commands.ApplyDiffEdits}'
    assert len('btn_ApplyDiffEdits') <= 31
    assert any('tags=diffUndoPending' in node.get('selector', '') for node in ui.iter())
    assert 'row("Worksheet") <> worksheetName' in procedure(sources['ex_PEB_MovementDiff.vba'], 'fn_Result')
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
        assert config['hotkeys.diffUndo'] == '^{DELETE}'
        assert config['movementDiff.fields'].split('|')[0] == 'DiffId'
        for placeholder in ['time', 'order', 'date', 'added', 'updated', 'unchanged']:
            assert '{' + placeholder + '}' in config['movementDiff.summary']
        settings.append(labels)
        for line in path.read_text(encoding='utf-8-sig').splitlines():
            assert not line.strip() or len(line.split('\t')) == 3
    assert settings[0] == settings[1]
    print('PASS: actual pre-write capture, unchanged/internal-field suppression, success-only publication, previous-report retention, navigation, PROD/TEST and VBA source contracts. Office/VBA not executed.')


if __name__ == '__main__':
    main()