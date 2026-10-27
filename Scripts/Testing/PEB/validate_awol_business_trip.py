"""Check the AWOL transition placeholder contract without executing Office/VBA."""

from xml.etree import ElementTree as ET

from event_routes import ROOT, parse_config
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
    assert placeholder.index('Exit Sub') < placeholder.index('Set writes = private_CopyFields')
    assert 'cells(personHeader) <> action("PersonId") Or cells(eventHeader) <> statusName' in placeholder
    assert 'writes(ex_PEB_Export.fn_Config("export.Movement.headers.TicketDocument")) = VBA.vbNullString' in placeholder
    assert 'For Each key In VBA.Array("DepartureDate", "DepartureFoodSupport", "ArrivalDate", "ArrivalFoodSupport")' in placeholder
    assert 'writes(ex_PEB_Export.fn_Config("export.Movement.headers." & key)) = action("EventDate")' in placeholder
    assert 'rows.Add row' in placeholder and 'action.Add "PlaceholderWrites", writes' in placeholder
    apply = procedure(sources['Planner'], 'private_Apply')
    assert 'closingWrites(ex_PEB_Export.fn_Config("export.Movement.headers.DepartureFoodSupport")) = eventDate' in apply
    assert apply.index('private_PlanReturnPlaceholder rows') < apply.index('private_Existing(rows')
    assert 'values("EventId") = "MoveFromAwolToBusinessTrip"' in apply
    assert 'action.Exists("PlaceholderWrites") And action("Status") = "AlreadyApplied"' in apply
    writer = procedure(sources['Writer'], 'fn_Apply')
    assert writer.index('action.Exists("PlaceholderWrites")') < writer.index('private_Extension(target, addedRows)')
    assert 'private_Capture target, action("PlaceholderIndex")' in writer
    stamp = procedure(sources['Links'], 'fn_PlanStamp')
    assert '"PlaceholderWrites"' in stamp and '"PlaceholderIndex"' in stamp
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
    print('PASS: AWOL source status, closed placeholder dates, stable link, repeated-export guard, '
          'writer allocation, plan fingerprint and caption. Office/VBA not executed.')


if __name__ == '__main__':
    main()