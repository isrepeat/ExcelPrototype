"""Check the AWOL transition placeholder contract without executing Office/VBA."""

from xml.etree import ElementTree as ET
from zipfile import ZipFile

from event_routes import ROOT, WORKBOOK, parse_config
from movement_plans import load_native_plans
from validate_movement_planner import procedure, check_declarations
from validate_movement_batch import check_blocks, check_module_sections
from validate_movement_tvo import preprocess


def main():
    folder = ROOT / 'PROJECTS_EXCEL/vba/PersonalEventBuilder'
    sources = {}
    for name in ('Planner', 'Writer', 'Links'):
        source = (folder / f'ex_PEB_Movement{name}.vba').read_text(encoding='utf-8-sig')
        sources[name] = source
        for enabled in (True, False):
            selected = preprocess(source, enabled)
            check_declarations(selected)
            check_blocks(selected)
            check_module_sections(selected)
    placeholder = procedure(sources['Planner'], 'private_PlanReturnPlaceholder')
    assert 'identifier = action("RowId") & "/Return"' in placeholder
    assert placeholder.index('fn_Resolve action, "Return"') < placeholder.index('Exit Sub')
    assert 'fn_Pending(CStr(action("RowId")))' in placeholder
    assert 'group.Add "Side", "Return"' in placeholder
    assert 'action.Add "PlaceholderBefore", before' in placeholder
    assert 'cells(personHeader) <> action("PersonId") Or cells(eventHeader) <> statusName' in placeholder
    assert 'writes(ex_PEB_Export.fn_Config("export.Movement.headers.TicketDocument")) = VBA.vbNullString' in placeholder
    assert 'For Each key In VBA.Array("DepartureDate", "DepartureFoodSupport", "ArrivalDate", "ArrivalFoodSupport")' in placeholder
    assert 'writes(ex_PEB_Export.fn_Config("export.Movement.headers." & key)) = action("EventDate")' in placeholder
    assert 'rows.Add row' in placeholder and 'action.Add "PlaceholderWrites", writes' in placeholder
    apply = procedure(sources['Planner'], 'private_Apply')
    assert 'closingWrites(ex_PEB_Export.fn_Config("export.Movement.headers.ArrivalFoodSupport")) = eventDate' in apply
    assert apply.index('private_PlanReturnPlaceholder rows') < apply.index('private_Existing(rows')
    assert 'values("EventId") = "MoveFromAwolToBusinessTrip"' in apply
    assert 'action.Exists("PlaceholderWrites") And action("Status") = "AlreadyApplied"' in apply
    writer = procedure(sources['Writer'], 'fn_Apply')
    assert writer.index('action.Exists("PlaceholderWrites")') < writer.index('private_Extension(target, addedRows)')
    assert 'private_Capture target, action("PlaceholderIndex")' in writer
    assert 'True, action("PlaceholderBefore")(key)' in writer
    assert 'If action("PlaceholderIndex") > target.ListRows.Count Then' in writer
    stamp = procedure(sources['Links'], 'fn_PlanStamp')
    assert '"PlaceholderWrites"' in stamp and '"PlaceholderIndex"' in stamp
    assert '"PlaceholderBefore"' in stamp
    for filename in ('wsConfig.txt', 'wsConfig.test.txt'):
        settings = parse_config(ROOT / 'PROJECTS_EXCEL/config/PersonalEventBuilder' / filename)
        plans = load_native_plans(settings)
        plan = plans['events']['MoveFromAwolToBusinessTrip']
        assert plan['operation'] == 'Transition'
        assert plans['statuses'][plan['sourceStatus']] == 'Самовільне залишення частини'
        assert plans['statuses'][plan['targetStatus']] == 'Відрядження'
        assert plans['statuses']['ReturnToOtherUnit'] == 'Повернення до іншої частини'
    routes = ET.parse(ROOT / 'PROJECTS_EXCEL/workbook/PersonalEventBuilder/PersonalEventBuilderEventRoutes.xml')
    event = next(node for node in routes.iter() if node.get('id') == 'MoveFromAwolToBusinessTrip')
    assert event.get('caption') == event.get('label') == 'СЗЧ → Відрядження'
    assert event.find('{*}word').get('template') == 'MoveFromAwolToBusinessTrip'
    word = ET.parse(ROOT / 'PROJECTS_EXCEL/workbook/PersonalEventBuilder/PersonalEventBuilderWordResultTemplates.xml')
    template = next(node for node in word.findall('{*}template') if node.get('name') == 'MoveFromAwolToBusinessTrip')
    rejection = template.find('{*}reject')
    assert rejection.get('field') == 'PersonId' and rejection.get('matches') == '^'
    assert rejection.get('messageKey') == 'text.AwolBusinessTripWordOmitted'
    assert template.find('{*}text') is not None
    assert settings['export.Word.textAnchors.MoveFromAwolToBusinessTrip'] == 'MoveFromAwolToBusinessTrip'
    with ZipFile(next(WORKBOOK.glob('{yyyy}*.docx'))) as archive:
        document = ET.fromstring(archive.read('word/document.xml'))
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    paragraphs = [''.join(node.itertext()) for node in document.findall('w:body/w:p', ns)]
    for kind in ('block', 'export'):
        for suffix in ('Begin', 'End'):
            assert paragraphs.count('{\\' + kind + ':MoveFromAwolToBusinessTrip_' + suffix + '}') == 1
    service = (folder / 'obj_PEB_ExportService.cls.vba').read_text(encoding='utf-8-sig')
    prepare = procedure(service, 'TryPrepare')
    no_word = prepare.index('If Not ex_PEB_EventRoutes.fn_HasRoute(CStr(values("EventId")), "Word") Then')
    assert no_word < prepare.index('requiredKey = "export."')
    assert 'exporter = "Movement"' in prepare[no_word:prepare.index('requiredKey = "export."')]
    check_declarations(service)
    check_blocks(service)
    print('PASS: AWOL source status, closed placeholder dates, stable link, repeated-export guard, '
          'writer allocation, plan fingerprint and caption. Office/VBA not executed.')


if __name__ == '__main__':
    main()