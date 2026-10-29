"""Check local new-event reset contracts without executing Office/VBA."""

from event_routes import ROOT
from validate_movement_planner import procedure, check_declarations
from validate_movement_batch import check_blocks
from validate_movement_tvo import preprocess


def main():
    folder = ROOT / 'PROJECTS_EXCEL/vba/PersonalEventBuilder'
    controller = (folder / 'obj_PEB_PgMainController.cls.vba').read_text(encoding='utf-8-sig')
    extensions = (folder / 'obj_PEB_FormExtensions.cls.vba').read_text(encoding='utf-8-sig')
    reset = procedure(controller, 'ResetFormHandler')
    clear = procedure(controller, 'private_ClearKeyFields')
    assert 'TrySearch' not in clear
    assert 'BindingContext.TryApplyValues(updates, diagnostic)' in clear
    assert 'private_PublishCandidates(result)' in clear
    assert 'm_extensions.Reset' not in clear
    assert 'Application.ScreenUpdating = previousUpdating' in clear
    for handler, group in (('PersonNameChangedHandler', 'personnel'),
                           ('HospitalChangedHandler', 'hospital'),
                           ('DocumentChangedHandler', 'document')):
        change = procedure(controller, handler)
        assert f'private_ClearKeyFields("{group}")' in change
    hotkeys = (folder / 'ex_PersonalEventBuilder.vba').read_text(encoding='utf-8-sig')
    shortcut = procedure(hotkeys, 'fn_ClearForm')
    assert 'm_clearArmed And elapsed <=' in shortcut
    assert 'elapsed = elapsed + 86400#' in shortcut
    assert 'commandName = "ResetForm"' in shortcut
    assert 'commandName = "ClearPrimary"' in shortcut
    for full_render in ('UpdatePage', 'RenderActivePage', 'RenderPages', 'm_extensions.Build'):
        assert full_render not in reset
    for cleared in ('Form.WordPreviewText', 'Form.ActionStatus',
                    'WordContext', 'SelectedPersonnel'):
        assert cleared in reset
    assert 'updates("Text.PersonnelStatus")' not in reset
    assert 'SetValue("Text", "PersonnelStatus", VBA.vbNullString)' in reset
    page = (folder / 'obj_PEB_PgMain.cls.vba').read_text(encoding='utf-8-sig')
    assert 'SetValue("Text", "PersonnelStatus", VBA.vbNullString)' in page
    assert 'forms.styledFields' in reset
    assert 'private_EndHistoryEdit()' in reset
    assert 'private_PublishCandidates(result)' in reset
    assert 'private_PublishHistory(records, VBA.vbNullString, diagnostic)' in reset
    assert 'm_historyOrderNumber = VBA.vbNullString' in reset
    assert 'm_historyOrderDate = VBA.vbNullString' in reset
    assert 'm_acceptedOrderNumber =' not in reset and 'm_acceptedOrderDate =' not in reset
    assert 'ex_PEB_Export.fn_FixedDocument(m_modeId)' in reset
    assert reset.index('m_extensions.Reset(diagnostic)') < reset.index('m_extensions.TryRefresh(diagnostic)')
    assert reset.index('previousScreenUpdating = Application.ScreenUpdating') < reset.index('TryApplyValues')
    assert 'Application.ScreenUpdating = previousScreenUpdating' in reset
    assert 'ex_Core.fn_Diagnostic_WritePerf "Form.Reset", startedAt' in reset
    assert reset.count('Exit Function') == 2
    assert reset.rsplit('Failed:', 1)[1].strip().endswith('GoTo Cleanup\nEnd Function')
    refresh = procedure(extensions, 'TryRefresh')
    assert 'Me.Build(m_form, diagnostic)' in refresh
    assert 'RefreshContent("ExtensionMenus", diagnostic, True)' in refresh
    assert 'RefreshContent("ExtensionRows", diagnostic, True)' in refresh
    assert 'm_page.FlushUpdates()' in refresh
    extension_reset = procedure(extensions, 'Reset')
    assert extension_reset.index('fn_Cancel Me') < extension_reset.index('Set m_rows = New Collection')
    for source in (controller, extensions):
        check_declarations(source)
        check_blocks(preprocess(source, True))
    print('PASS: binding reset, local extension refresh, pending refresh cancellation, '
          'history exit, order preservation, screen restoration and timing. Office/VBA not executed.')


if __name__ == '__main__':
    main()