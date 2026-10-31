"""Check the TVO table migration and legacy read compatibility without Office/VBA."""

from event_routes import ROOT, parse_config, load_registry, NS
from movement_plans import load_native_plans
from validate_movement_batch import check_blocks, check_module_sections
from validate_movement_planner import check_declarations, procedure
from copy import deepcopy
from xml.etree import ElementTree as ET


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
    normalize = procedure(table, 'private_HistoryDuties')
    assert 'If item.Exists("DutyPositionCode") Then' in normalize
    assert 'fn_Required(item, "DutyPositionCode")' in normalize
    assert 'Set row = CreateObject("Scripting.Dictionary")' in normalize
    history = (folder / 'ex_PEB_History.vba').read_text(encoding='utf-8-sig')
    children = procedure(history, 'private_Tvo')
    assert 'row("DutyPositionCode") = parent("PositionCode")' in children
    assert 'row("DutyPositionGenitive") = parent("PositionGenitive")' in children
    assert 'If result.Count > 0 Then Set parent = result(result.Count)' in children
    service = (folder / 'obj_PEB_ExportService.cls.vba').read_text(encoding='utf-8-sig')
    extensions = procedure(service, 'TryPrepareExtensions')
    assert 'Set parent = values' in extensions
    assert 'Set parent = row' in extensions
    assert 'row("DutyPositionCode") = ex_PEB_Export.fn_Required(parent, "PositionCode")' in extensions
    form_children = (folder / 'ex_PEB_FormChildren.vba').read_text(encoding='utf-8-sig')
    assert 'fields("ParentRowId") = tvoParentId' in form_children
    assert 'tvoParentId = identifier' in form_children
    assert 'row("DutyPersonName") = parent("PersonName")' in children
    assert 'private_Tvo rows, row, result' in children
    prepare_tvo = procedure(table, 'fn_Prepare')
    # Изменения ТВО проверяются до пропуска уже выгруженных данных отсутствующего.
    revise_call = prepare_tvo.index('private_ReviseDeparture model, chain, requested')
    assert revise_call < prepare_tvo.index('If action("Status") = "AlreadyApplied" Then')
    assert 'row("Action")("Status") = "Planned"' in prepare_tvo
    assert 'row("Action")("IsUpdate") = True' in prepare_tvo
    revise = procedure(table, 'private_ReviseDeparture')
    assert 'chain.Count <> 1 Or requested.Count <> 1' in revise
    assert 'fn_OrderDate(row("StartDate")) <> CDbl(eventDate)' in revise
    for field in ('StartOrder', 'EndDate', 'EndOrder'):
        assert f'row("{field}")' in revise
    # Текст основания без даты и приказа не подтверждает сдачу; сохраняем его без изменений.
    assert 'Len(ex_PEB_Export.fn_Text(row("EndBasis")))' not in revise
    assert 'desired(field) = row(field)' in revise
    assert 'desired("EndBasis") =' not in revise
    assert revise.index('If Not changed Then') < revise.index('fn_OrderDate(row("StartDate"))')
    assert revise.index('"TvoRevisionRequiresReview"') < revise.index('row(field) = desired(field)')
    assert 'private_RenderTexts desired, values' in revise
    assert 'ListRows.Add' not in revise
    assert 'row("PlannedEndDate") = values("DateToShort")' not in prepare_tvo
    assert 'row("PlannedEndDate") = DateAdd(' in prepare_tvo
    assert 'ex_PEB_Export.fn_ParseDate(values("DateToShort"), "DateToShort")' in prepare_tvo
    apply_tvo = procedure(table, 'fn_Apply')
    assert 'table.ListRows.Add(AlwaysInsert:=False)' in apply_tvo
    assert 'private_InheritRowFormats _' in apply_tvo
    assert 'PasteSpecial' not in table and '.Range.Copy' not in table
    assert apply_tvo.index('private_InheritRowFormats') < apply_tvo.index('cell.Value2 = expected')
    inherit = procedure(table, 'private_InheritRowFormats')
    assert '.Font.Name = sourceCell.Font.Name' in inherit
    assert '.Font.Size = sourceCell.Font.Size' in inherit
    assert '.NumberFormat = sourceCell.NumberFormat' in inherit
    assert 'table.Resize' not in apply_tvo
    assert apply_tvo.index('table.ListRows.Add') < apply_tvo.index('currentFormat = cell.NumberFormat')
    assert 'cell.NumberFormat = currentFormat' in apply_tvo
    assert prepare_tvo.index('Set requested = private_HistoryDuties') < prepare_tvo.index('private_Matches(chain, requested')
    assert 'Set chain = requested' in prepare_tvo
    assert 'private_RenderTexts row, values, replacedPerson' in prepare_tvo
    render_tvo = procedure(table, 'private_RenderTexts')
    assert 'fn_Texts(CStr(values("MovementTemplateId")), "TVO")' in render_tvo
    assert 'renderer.TryRender(' in render_tvo
    assert 'returnTvo.table.values.' not in table
    assert 'returnTvo.reasons.' not in table
    fields = procedure(table, 'private_Fields')
    assert '|Duration|Information' in fields
    for filename in ('wsConfig.txt', 'wsConfig.test.txt'):
        settings = parse_config(ROOT / 'PROJECTS_EXCEL/config/PersonalEventBuilder' / filename)
        assert settings['returnTvo.table.headers.Duration'] == 'Термін'
        assert settings['returnTvo.table.headers.Information'] == 'Додаткова інформація'
        assert not any(key.startswith(('returnTvo.table.values.', 'returnTvo.reasons.')) for key in settings)
        registry = load_registry(settings)
        movement = registry['movement']
        for mutation in ('unknown_target', 'duplicate_tvo', 'unknown_field'):
            altered = deepcopy(movement['root'])
            template = altered.find("p:template[@name='ToAnnualMainVacation']", NS)
            column = template.find("p:column[@target='TVO']", NS)
            if mutation == 'unknown_target':
                column.set('target', 'UnknownTable')
            elif mutation == 'duplicate_tvo':
                template.append(deepcopy(column))
            else:
                column.find('p:text', NS).text = '{[UnknownTvoField]}'
            try:
                load_registry(settings, movement_text=ET.tostring(altered, encoding='unicode'))
            except ValueError:
                pass
            else:
                raise AssertionError('Invalid TVO column accepted: ' + mutation)
        reasons = movement['root'].find("p:resources/p:dictionary[@name='TvoDepartureReasons']", NS)
        reason_keys = {entry.get('key') for entry in reasons.findall('p:entry', NS)}
        assert reasons.find("p:entry[@key='AnnualVacation']/p:variant", NS).get('value') == 'у щорічну відпустку'
        native = load_native_plans(settings)
        for name, event in movement['events'].items():
            texts = movement['tvo_columns'][name]
            if event['operation'] != 'Arrival':
                assert event['targetStatus'] in reason_keys
                assert set(texts) == {'Термін', 'Додаткова інформація'}
                assert texts['Термін'] == 'Інше'
                assert texts['Додаткова інформація'] == '{#include block.TvoInformation}'
            else:
                assert not texts
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