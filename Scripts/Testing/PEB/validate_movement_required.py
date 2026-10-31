"""Validate required Movement inputs and hidden destination isolation without Office."""

from event_routes import ROOT, parse_config, load_registry
from validate_movement_planner import procedure, check_declarations
from validate_movement_batch import check_blocks


def main():
    for filename in ('wsConfig.txt', 'wsConfig.test.txt'):
        settings = parse_config(ROOT / 'PROJECTS_EXCEL/config/PersonalEventBuilder' / filename)
        registry = load_registry(settings)
        events = registry['movement']['events']
        outpatient = events['ToOutpatientExaminationOrVlk']
        assert outpatient['destinationField'] == 'Destination'
        assert not {'Destination', 'Hospital', 'HospitalShortName'} & set(outpatient['required'].split('|'))
        assert 'HospitalShortName' in events['ToHospital']['required'].split('|')
        node = registry['events']['ToOutpatientExaminationOrVlk']
        form = node.get('form')
        fields = {settings[f'forms.{form}.fields.{i}.name']
                  for i in range(1, int(settings[f'forms.{form}.fields.count']) + 1)}
        assert 'Destination' not in fields
        assert '{field}' in settings['text.MovementRequiredMissing']

    folder = ROOT / 'PROJECTS_EXCEL/vba/PersonalEventBuilder'
    sources = {}
    for name in ('ex_PEB_History.vba', 'ex_PEB_Export.vba', 'ex_PEB_MovementRegistry.vba',
                 'obj_PEB_ExportService.cls.vba', 'obj_PEB_PgMainController.cls.vba'):
        source = (folder / name).read_text(encoding='utf-8-sig')
        check_declarations(source)
        check_blocks(source)
        sources[name] = source
    save = procedure(sources['ex_PEB_History.vba'], 'fn_Save')
    assert save.index('fn_CheckMovementRequired current') < save.index('HistoryRevisions.fn_Prepare')
    assert save.index('fn_CheckMovementRequired current') < save.index('target.Range.Value2 = data')
    assert 'If current("Status") <> "Cancelled" Then' in save
    assert 'If Not values.Exists("RequestedStatus") Then' in save
    assert 'fn_ClearHiddenDestination values' in save
    export = procedure(sources['obj_PEB_ExportService.cls.vba'], 'TryPrepareHistoryMovement')
    assert 'fn_CheckMovementRequired values' in export
    assert 'HistoryError030' not in export
    assert export.index('fn_ClearHiddenDestination values') < export.index('fn_Apply(values, "Movement")')
    prepare = procedure(sources['obj_PEB_ExportService.cls.vba'], 'TryPrepare')
    assert prepare.index('fn_ClearHiddenDestination values') < prepare.index('fn_Apply(values, exporter)')
    switch = procedure(sources['obj_PEB_PgMainController.cls.vba'], 'private_SwitchMode')
    assert switch.index('fn_ClearHiddenDestination updates') < switch.index('private_LoadMode(modeId')
    clear = procedure(sources['ex_PEB_Export.vba'], 'fn_ClearHiddenDestination')
    assert 'movement.destinationValue' in clear
    assert 'values("Destination") = fixed' in clear
    assert '"forms." & form & ".fields."' in clear
    assert 'ToOutpatient' not in clear
    print('PASS: optional outpatient inputs, hospital requirements, shared save/export validation, '
          'hidden form destination isolation and configured fixed destinations. Office/VBA not executed.')


if __name__ == '__main__':
    main()