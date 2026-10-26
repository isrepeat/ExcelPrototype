"""Check the TVO table migration and legacy read compatibility without Office/VBA."""

from event_routes import ROOT, parse_config
from movement_plans import load_native_plans
from validate_movement_batch import check_blocks, check_module_sections
from validate_movement_planner import check_declarations, procedure


def preprocess(source, enabled):
    result = []
    active = True
    for line in source.splitlines():
        if line == '#If TVO_DEPRECATED Then':
            active = enabled
        elif line == '#Else':
            active = not enabled
        elif line == '#End If':
            active = True
        elif active:
            result.append(line)
    return '\n'.join(result)


def main():
    folder = ROOT / 'PROJECTS_EXCEL/vba/PersonalEventBuilder'
    names = ('ex_PEB_MovementRegistry.vba', 'obj_PEB_ExportService.cls.vba',
             'ex_PEB_MovementConflicts.vba', 'ex_PEB_MovementPlanner.vba')
    sources = {name: (folder / name).read_text(encoding='utf-8-sig') for name in names}
    for name, source in sources.items():
        assert '#Const TVO_DEPRECATED = True' in source
        for enabled in (True, False):
            selected = preprocess(source, enabled)
            check_declarations(selected)
            check_blocks(selected)
            if not name.startswith('obj_'):
                check_module_sections(selected)
        disabled = preprocess(source, False)
        assert 'fn_IsTvoHeader(' not in disabled
        assert 'private_TvoValue(' not in disabled
        assert 'private_AddTvoMappings ' not in disabled
    registry = sources[names[0]]
    assert 'private_TvoValue' not in registry
    assert 'private_AddTvoMappings' not in registry
    prepare = procedure(sources[names[1]], 'TryPrepareHistoryMovement')
    assert 'Set values("MetaTvoTables") = stored("MetaTvoTables")' in prepare
    resolve = procedure(sources[names[2]], 'fn_Resolve')
    assert 'If Not validHash And side = "Opening" Then' in resolve
    assert 'Or m_states.Exists(fieldKey) Or _' in resolve
    assert 'VBA.Len(ex_PEB_Export.fn_Text(cells(field))) > 0 Then' in resolve
    assert 'fn_FieldsHash(legacyFields, cells)' in resolve
    assert 'fn_FieldsHash(previousFields)' in resolve
    assert 'manual = (actual <> state(2))' in resolve
    assert 'If field <> basisHeader Then' in resolve
    for filename in ('wsConfig.txt', 'wsConfig.test.txt'):
        settings = parse_config(ROOT / 'PROJECTS_EXCEL/config/PersonalEventBuilder' / filename)
        for enabled in (True, False):
            source = registry.replace('#Const TVO_DEPRECATED = True',
                                      f'#Const TVO_DEPRECATED = {enabled}')
            plans = load_native_plans(settings, source)
            for profile in plans['profiles']:
                prefix = f'profiles.{profile}.fields'
                mapping = {plans['values'][f'{prefix}.{i}.field']: plans['values'][f'{prefix}.{i}.header']
                           for i in range(1, int(plans['values'][prefix + '.count']) + 1)}
                for field in ('TvoName', 'TvoPersonId', 'TvoPositionCode'):
                    assert field not in mapping
        assert settings['returnTvo.table.name'] == 'ТВО'
        assert settings['returnTvo.table.headers.PositionCode'] == settings['reporter.tvo.codeHeader']
        assert settings['returnTvo.table.headers.StartDate'] == settings['reporter.tvo.startHeader']
        assert settings['returnTvo.table.headers.EndDate'] == settings['reporter.tvo.endHeader']
    table = (folder / 'ex_PEB_MovementTvo.vba').read_text(encoding='utf-8-sig')
    check_declarations(table)
    check_blocks(table)
    check_module_sections(table)
    predicate = procedure(table, 'fn_IsStarted')
    assert predicate.index('fn_Text(startOrder)') < predicate.index('fn_ParseDate(startDate')
    assert 'private_AdmissionOnly(model, item, eventDate)' in procedure(table, 'fn_Prepare')
    reporter = (folder / 'ex_PEB_Export.vba').read_text(encoding='utf-8-sig')
    assert 'fn_Config("returnTvo.table.headers.StartOrder")' in procedure(reporter, 'fn_TryReporterTvoPositions')
    assert 'ex_PEB_MovementTvo.fn_IsStarted(' in procedure(reporter, 'fn_TryReporterTvoPositions')
    fixture = (folder / 'Test/ex_PEB_TvoAdmissionTests.vba').read_text(encoding='utf-8-sig')
    check_declarations(fixture)
    check_blocks(fixture)
    check_module_sections(fixture)
    batch = (folder / 'ex_PEB_MovementWriter.vba').read_text(encoding='utf-8-sig')
    apply = procedure(batch, 'fn_Apply')
    assert apply.index('MovementTvo.fn_AddDiff') < apply.index('changed = True')
    assert apply.index('MovementTvo.fn_Apply plan') < apply.index('private_Stage "Write.Journal"')
    assert 'MovementTvo.fn_Restore(plan)' in apply
    assert 'TvoTargetChangedSincePreview' in procedure(table, 'fn_Apply')
    print('PASS: new exports omit legacy TVO columns; dedicated table headers, rollback, '
          'pre-write diff and target revalidation are wired. Office/VBA not executed.')


if __name__ == '__main__':
    main()