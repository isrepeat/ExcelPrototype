"""Check shared table filter suspension and cross-mode transaction integration without Office."""

from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).parent / 'PEB'))
from event_routes import ROOT
from validate_movement_planner import procedure, check_declarations
from validate_movement_batch import check_blocks


def read(path):
    source = (ROOT / 'PROJECTS_EXCEL/vba' / path).read_text(encoding='utf-8-sig')
    check_declarations(source)
    check_blocks(source)
    return source


def main():
    shared = read('common/modules/ex_TableFilters.vba')
    suspend = procedure(shared, 'fn_Suspend')
    assert suspend.index('fn_Capture(table)') < suspend.index('fn_Clear table')
    assert 'restoreError = fn_Restore(states)' in suspend
    assert 'seen.Exists(key)' in suspend
    restore = procedure(shared, 'fn_Restore')
    for text in ('Criteria1:=entry("Criteria1")', 'Criteria2:=entry("Criteria2")',
                 'Operator:=entry("Operator")', 'If states Is Nothing Then'):
        assert text in restore
    assert 'ex_PEB_' not in shared and 'ex_TDB_' not in shared
    assert 'Public Function fn_Suspend( _' in shared
    for name in ('fn_Capture', 'fn_Clear'):
        proc = procedure(shared, name)
        assert 'Set autoFilter = table.AutoFilter' in proc
        assert proc.index('If autoFilter Is Nothing Then') < proc.index('autoFilter.Filter')
    assert '"TableFilters." & context' in suspend
    writer = read('PersonalEventBuilder/ex_PEB_MovementWriter.vba')
    apply = procedure(writer, 'fn_Apply')
    assert apply.index('private_SuspendFilters plan') < apply.index('changed = True')
    assert apply.index('ex_TableFilters.fn_Restore(filterStates)') < apply.index('Application.EnableEvents = previousEvents')
    assert 'ex_PEB_MovementTvo.fn_Restore(plan)' in apply
    scope = procedure(writer, 'private_SuspendFilters')
    assert 'tables.Add target' in scope and 'tables.Add plan("TvoPlan")("Table")' in scope
    assert 'history.path' in scope
    history = read('PersonalEventBuilder/ex_PEB_History.vba')
    save = procedure(history, 'fn_Save')
    assert save.index('ex_TableFilters.fn_Suspend') < save.index('target.Range.Value2 = data')
    assert 'ex_TableFilters.fn_Restore(filterStates)' in save
    assert 'private_CaptureFilters' not in history
    undo = procedure(read('PersonalEventBuilder/ex_PEB_MovementDiff.vba'), 'fn_ApplyUndo')
    assert undo.index('ex_TableFilters.fn_Suspend') < undo.index('changed = True')
    assert undo.index('ex_TableFilters.fn_Restore(filterStates)') < undo.index('Application.EnableEvents = previousEvents')
    tdb = procedure(read('TravelDocumentsBuilder/obj_TDB_VacationGenerator.cls.vba'), 'TrySave')
    assert tdb.index('ex_TableFilters.fn_Suspend') < tdb.index('tableChanged = True')
    assert tdb.index('ex_TableFilters.fn_ClearCaptured') < tdb.index('ex_TDB_History.fn_Rollback')
    assert tdb.count('ex_TableFilters.fn_Restore(filterStates)') == 2
    word = procedure(read('TravelDocumentsBuilder/obj_TDB_OrderWordExport.cls.vba'), 'TryExport')
    assert word.index('ex_TableFilters.fn_Suspend') < word.index('historyChanged = True')
    assert word.index('ex_TableFilters.fn_ClearCaptured') < word.index('ex_TDB_History.fn_Rollback')
    assert 'ex_TableFilters.fn_Restore(filterStates)' in word
    print('PASS: shared capture/clear/restore, all-before-clear validation, duplicate handling, '
          'PEB writes/deletion/reversal and TDB save/Word rollback integration. Office/VBA not executed.')


if __name__ == '__main__':
    main()