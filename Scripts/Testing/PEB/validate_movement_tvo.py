"""Check deprecated TVO export mappings and upgrade guards without Office/VBA."""

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
    writer = procedure(registry, 'private_TvoValue')
    assert 'For Each row In values("MetaTvoTables")' in writer
    assert 'VBA.Array("PersonName", "PersonId", "PositionCode")' in writer
    for guard in ('VBA.vbCr', 'VBA.vbLf', 'seen.Exists(value)', 'text.ReturnTvoInvalidChain'):
        assert guard in writer
    assert 'text = text & VBA.vbLf' in writer
    assert 'text = text & ex_PEB_Export.fn_Text(row(field))' in writer
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
                    if enabled and plans['profiles'][profile]['operation'] != 'close':
                        assert mapping[field] == settings['returnTvo.headers.' + field]
                    else:
                        assert field not in mapping
    print('PASS: TVO mappings in PROD/TEST, chain guards, legacy empty-column migration, '
          'manual-change guards and both TVO_DEPRECATED branches. Office/VBA not executed.')


if __name__ == '__main__':
    main()