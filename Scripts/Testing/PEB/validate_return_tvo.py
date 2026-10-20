"""Check arrival-only TVO snapshot contracts without executing Excel or VBA."""

from event_routes import ROOT, parse_config, load_registry
from validate_movement_planner import check_declarations, procedure
from validate_movement_batch import check_blocks, check_module_sections


def main():
    folder = ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder"
    read = lambda name: (folder / name).read_text(encoding="utf-8-sig")
    source = read("ex_PEB_MovementSource.vba")
    restore = read("ex_PEB_ReturnTvo.vba")
    history = read("ex_PEB_History.vba")
    service = read("obj_PEB_ExportService.cls.vba")
    for text in (source, restore, history, service):
        check_declarations(text)
        check_blocks(text)
    check_module_sections(restore)
    latest = procedure(source, "fn_ReadLatest")
    assert 'Optional ByVal includeTvo As Boolean = False' in latest
    assert 'If includeTvo Then' in latest
    assert 'table.DataBodyRange.Cells(index, names(field)).Value2' in latest
    assert 'records.Fields(column).Value' in latest
    capture = procedure(restore, "fn_Read")
    assert capture.index('"direction") <> "Arrival"') < capture.index('fn_ReadLatest(')
    assert 'plan("operation") <> "Arrival"' in capture
    assert 'fn_ReadLatest(values("PersonId"), True)' in capture
    assert 'statuses(plan("sourceStatus"))' in capture
    assert 'If arrivalDate <> eventDate Or _' in capture
    assert 'period("ArrivalOrder")' in capture and 'values("AcceptedOrderNumber")' in capture
    assert 'names.Count <> people.Count Or names.Count <> positions.Count' in capture
    assert 'seen.Exists(people(index))' in capture
    assert 'row("PositionCode") = positions(index)' in capture
    decode = procedure(restore, "fn_Decode")
    assert 'root.getAttribute("person") <> values("PersonId")' in decode
    assert 'root.getAttribute("event") <> values("EventId")' in decode
    assert 'root.getAttribute("date")' in decode
    assert 'seen.Exists(row("PersonId"))' in decode
    assert 'fn_ReadLatest' not in decode
    assert 'resolveExternals = False' in restore and '"ProhibitDTD", True' in restore
    save = procedure(history, "fn_Save")
    assert 'If Not values.Exists("RequestedStatus") Then' in save
    assert 'If Not hasExplicitTvo Then' in save
    assert 'If Not reuseTvo Then' in save
    assert save.index('TryCaptureReturnTvo') < save.index('ex_PEB_HistoryRecovery.fn_Prepare')
    assert 'previous("TvoReturnSnapshot")' in save
    order = procedure(history, "fn_ReadOrder")
    assert order.index('private_Tvo rows, row') < order.index('ex_PEB_ReturnTvo.fn_Decode(row)')
    assert 'If row("MetaTvoTables").Count = 0 Then' in order
    assert 'fn_ReadLatest' not in order
    prepare = procedure(service, "TryPrepare")
    assert 'If Not history And values("ModeDirection") = "Arrival" Then' in prepare
    assert 'If values("MetaTvoTables").Count = 0 Then' in prepare
    assert 'fn_ReadLatest' not in procedure(service, "TryPrepareHistory")
    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        settings = parse_config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder" / filename)
        registry = load_registry(settings)
        assert 'TvoReturnSnapshot=TvoReturnSnapshot' in settings['history.fields'].split('|')
        assert settings['returnTvo.headers.TvoName'] == 'ТВО.ПІБ'
        assert settings['returnTvo.headers.TvoPersonId'] == 'ТВО.ІПН'
        assert settings['returnTvo.headers.TvoPositionCode'] == 'ТВО.Посада'
        fields = set(settings['returnTvo.snapshot.fields'].split('|'))
        outputs = {settings[f'export.Word.meta.lookups.{i}.output']
                   for i in range(1, int(settings['export.Word.meta.lookups.count']) + 1)}
        assert outputs | {'PersonId', 'PersonName', 'PersonRank', 'PositionCode'} == fields
        assert settings['returnTvo.rank.output'] == 'PersonRank'
        assert settings['returnTvo.rank.input'] == 'PersonId'
        assert settings['returnTvo.rank.valueHeader'] == settings['lookup.Personnel.columns.3.header']
        assert any(event.get('direction') == 'Arrival' for event in registry['events'].values())
    assert 'TvoReturnSnapshot' not in read('ex_PEB_MovementWriter.vba')
    print('PASS: arrival-only TVO restoration, manual precedence, frozen snapshots, source/closed-period checks and TEST/PROD mappings. Office/VBA not executed.')


if __name__ == '__main__':
    main()