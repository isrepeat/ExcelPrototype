"""Check fixed document policies without executing Office/VBA."""

from xml.etree import ElementTree as ET

from event_routes import ROOT, parse_config
from validate_movement_planner import procedure, check_declarations
from validate_movement_batch import check_blocks
from validate_movement_tvo import preprocess


def main():
    folder = ROOT / 'PROJECTS_EXCEL/vba/PersonalEventBuilder'
    sources = {}
    for name in ('ex_PEB_Export.vba', 'obj_PEB_PgMainController.cls.vba',
                 'obj_PEB_FormExtensions.cls.vba', 'obj_PEB_ExportService.cls.vba'):
        source = (folder / name).read_text(encoding='utf-8-sig')
        check_declarations(source)
        check_blocks(preprocess(source, True))
        sources[name] = source
    export = sources['ex_PEB_Export.vba']
    policy = procedure(export, 'fn_FixedDocument')
    assert policy.index('document.fixed.modes') < policy.index('fn_IsTicketDocumentForm')
    values = procedure(export, 'fn_TicketDocumentValues')
    assert 'If Not stored Then' in values
    assert values.index('fn_IsTicketDocumentMode') < values.index('values("TicketNumber") = values("DocumentNumber")')
    assert 'If Not stored Or CStr(values("Document")) <>' in values
    controller = sources['obj_PEB_PgMainController.cls.vba']
    readonly = procedure(controller, 'private_UpdateReadOnly')
    assert 'fn_FixedDocument(m_modeId) <> ex_PEB_Export.fn_FixedDocument(modeId)' in readonly
    assert 'updates("Form.DocumentReadOnly") = True' in readonly
    assert 'updates("Form.Document") = ex_PEB_Export.fn_FixedDocument(modeId)' in readonly
    assert 'column("formField") = "DocumentNumber"' in procedure(controller, 'private_SetCandidateSources')
    extensions = sources['obj_PEB_FormExtensions.cls.vba']
    for name in ('Collect', 'HandleAction'):
        assert 'private_FixedDocument(row)' in procedure(extensions, name)
    assert 'node.setAttribute "readOnly", "true"' in extensions
    assert 'private_ReadConfig("extensions.role." & row("Role") & ".document")' in extensions
    ticket = procedure(sources['obj_PEB_ExportService.cls.vba'], 'private_PrepareTicket')
    assert 'If CStr(values("Document")) <> ex_PEB_Export.fn_Config("lookup.ReturnTickets.document") Then Exit Sub' in ticket
    routes = ET.parse(ROOT / 'PROJECTS_EXCEL/workbook/PersonalEventBuilder/PersonalEventBuilderEventRoutes.xml')
    forms = {node.get('id'): node.get('form') for node in routes.iter() if node.tag.endswith('}event')}
    for filename in ('wsConfig.txt', 'wsConfig.test.txt'):
        path = ROOT / 'PROJECTS_EXCEL/config/PersonalEventBuilder' / filename
        for line in path.read_text(encoding='utf-8-sig').splitlines():
            if line.strip():
                assert len(line.split('\t')) == 3
        settings = parse_config(path)
        fixed_modes = settings['document.fixed.modes'].split('|')
        def document(mode):
            if mode in fixed_modes:
                return settings['document.fixed.' + mode]
            return settings['lookup.ReturnTickets.document'] if forms[mode] in settings['export.ticket.form'].split('|') else ''
        assert document('ToVacationForCuring') == 'В/к'
        assert document('MoveFromHospitalToVacationForCuring') == 'В/к'
        assert document('MoveFromVacationForCuringToVacationForCuring') == 'В/к'
        assert document('ToAnnualMainVacation') == document('FromAnnualMainVacation') == 'В/к'
        assert document('FromHospital') == ''
        assert settings['extensions.role.Vlk.document'] == 'Довідка ВЛК'
        assert settings['extensions.role.MedicalCard.document'] == 'м/к'
        for mode, expected in (
            ('ToVacationForCuring', {'Vlk'}),
            ('MoveFromVacationForCuringToVacationForCuring', {'Vlk'}),
            ('MoveFromHospitalToVacationForCuring', {'Vlk', 'MedicalCard'}),
        ):
            assert mode in settings['extensions.modes'].split('|')
            prefix = 'extensions.mode.' + mode + '.default'
            roles = set()
            for index in range(1, int(settings[prefix + '.count']) + 1):
                item = f'{prefix}.{index}'
                assert settings[item + '.type'] == 'Document'
                assert settings[item + '.removable'] == 'false'
                role = settings[item + '.role']
                assert role not in roles
                roles.add(role)
                assert settings['extensions.role.' + role + '.required'] == 'Document|DocumentNumber|DocumentDate'
            assert roles == expected
    print('PASS: fixed documents in PROD/TEST, editable number/date, ticket separation, '
          'stored snapshot guards and mandatory VLK row. Office/VBA not executed.')


if __name__ == '__main__':
    main()