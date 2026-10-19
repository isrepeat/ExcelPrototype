"""Check history-to-Word preparation and measurement contracts without Office."""

from event_routes import ROOT, parse_config
from validate_movement_batch import check_blocks
from validate_movement_planner import check_declarations, procedure


def main():
    folder = ROOT / 'PROJECTS_EXCEL/vba/PersonalEventBuilder'
    sources = {name: (folder / name).read_text(encoding='utf-8-sig') for name in (
        'obj_PEB_ExptrWord.cls.vba', 'obj_PEB_ExportService.cls.vba',
        'obj_PEB_PgMainController.cls.vba')}
    for source in sources.values():
        check_declarations(source)
        check_blocks(source)
    prepare = procedure(sources['obj_PEB_ExportService.cls.vba'], 'TryPrepareHistory')
    generate = procedure(sources['obj_PEB_ExptrWord.cls.vba'], 'TryGenerateHistory')
    command = procedure(sources['obj_PEB_PgMainController.cls.vba'], 'ExportWordHandler')
    assert 'TryRender(template, values, body' not in prepare
    assert 'service.TryPrepareHistory(item, diagnostic, first, templates)' in generate
    assert 'If Not documentContext Is Nothing Then' in prepare
    assert prepare.index('fn_ValueToken(documentContext(key))') < prepare.index('Set values("WordDocumentReplacements") = documentContext')
    assert 'm_ownsWordTemplates = False' in prepare
    assert 'VBA.Join(parts, VBA.vbCrLf & VBA.vbCrLf)' in generate
    assert 'body = body & text' not in generate
    assert 'For Each eventId In rendered.Keys' in generate
    assert 'ex_PEB_EventRoutes.fn_Ids()' not in generate
    assert generate.index('anchors.TryRemoveMarkers') < generate.index('document.SaveAs2')
    assert 'anchors.HideMarkers' not in generate
    assert generate.index('Prepare.RenderSections') < generate.index('AcquireWord')
    for stage in ('AcquireWord', 'CreateDocument', 'ReplaceDocumentFields', 'Section.IndexAnchors',
                  'Section.Insert', 'Section.Verify', 'SaveDocument', 'CloseDocument', 'PublishDocument'):
        assert f'private_HistoryStage "{stage}"' in generate
    assert 'Word.Export.GenerateTotal' in generate
    for stage in ('ReadEvents', 'RecordJournal', 'RefreshHistory', 'TotalIncludingJournalAndRefresh'):
        assert f'Word.Export.{stage}' in command
    assert 'ex_PEB_Export.fn_EndConfigSnapshot' in generate
    for filename in ('wsConfig.txt', 'wsConfig.test.txt'):
        config = parse_config(ROOT / 'PROJECTS_EXCEL/config/PersonalEventBuilder' / filename)
        common = set(config['export.Word.documentContext.fields'].split('|'))
        assert common == {'AcceptedOrderNumber', 'AcceptedOrderDate', 'Signatory'}
    print('PASS: single body rendering, validated shared document context, template ownership, joined sections and export stage timers. Excel/Word not executed.')


if __name__ == '__main__':
    main()