"""Check history-to-Word preparation and measurement contracts without Office."""

from event_routes import ROOT, parse_config
from validate_movement_batch import check_blocks
from validate_movement_planner import check_declarations, procedure


def main():
    folder = ROOT / 'PROJECTS_EXCEL/vba/PersonalEventBuilder'
    sources = {name: (folder / name).read_text(encoding='utf-8-sig') for name in (
        'obj_PEB_ExptrWord.cls.vba', 'obj_PEB_ExportService.cls.vba',
        'obj_PEB_PgMainController.cls.vba', 'obj_PEB_WordAnchors.cls.vba')}
    for source in sources.values():
        check_declarations(source)
        check_blocks(source)
    prepare = procedure(sources['obj_PEB_ExportService.cls.vba'], 'TryPrepareHistory')
    generate = procedure(sources['obj_PEB_ExptrWord.cls.vba'], 'TryGenerateHistory')
    command = procedure(sources['obj_PEB_PgMainController.cls.vba'], 'ExportWordHandler')
    assert command.index('busyEntered = ex_UiBusy.fn_Begin()') < command.index('ex_PEB_History.fn_ReadOrder(')
    assert 'CleanupHistoryWord:\n    If busyEntered Then\n        ex_UiBusy.fn_End' in command
    assert 'fn_ReadOrder(orderNumber, orderDate, events, diagnostic)' in command
    assert 'usedSavedFile' not in command
    history = (folder / 'ex_PEB_History.vba').read_text(encoding='utf-8-sig')
    assert 'savedTable' not in history and 'allowSavedFile' not in history
    assert 'data = table.DataBodyRange.Value2' in history
    assert not (folder / 'ex_PEB_HistorySaved.vba').exists()
    anchors = sources['obj_PEB_WordAnchors.cls.vba']
    index_blocks = procedure(anchors, 'private_IndexBlocks')
    cleanup = procedure(anchors, 'private_ApplyCleanup')
    plan_cleanup = procedure(anchors, 'private_PlanCleanup')
    assert 'whitespace.Pattern = "[\\s"' in plan_cleanup
    assert 'whitespace.Pattern = "[\\\\s"' not in plan_cleanup
    assert 'private_MarkerRange(' not in index_blocks, 'Block indexing must not search Word for every boundary'
    assert generate.count('anchors.Initialize(document, diagnostic)') == 1
    assert generate.index('anchors.Initialize(document, diagnostic)') < generate.index('For Each eventId In rendered.Keys')
    assert 'private_FindMarker(' not in plan_cleanup, 'Cleanup planning must not search Word'
    assert 'For index = ranges.Count To 1 Step -1' in cleanup
    assert cleanup.index('ranges.Add target') < cleanup.index('target.Delete')
    assert 'VBA.StrComp(target.Text, removed, VBA.vbBinaryCompare)' in cleanup
    assert 'private_DeleteRange' not in anchors
    assert 'WORD_CLEANUP_MISMATCH' in anchors
    assert 'TryRender(template, values, body' not in prepare
    assert 'service.TryPrepareHistory(item, diagnostic, first, templates)' in generate
    assert 'If Not documentContext Is Nothing Then' in prepare
    assert prepare.index('fn_ValueToken(documentContext(key))') < prepare.index('Set values("WordDocumentReplacements") = documentContext')
    assert 'm_ownsWordTemplates = False' in prepare
    # Snapshot synchronization can supply templates while clearing the renderer.
    renderer_guard = prepare.index('If m_pebDslRenderer Is Nothing Then')
    template_guard = prepare.index('If m_pebWordTemplates Is Nothing Then')
    depth = 0
    for line in prepare[template_guard:renderer_guard].splitlines():
        stripped = line.strip()
        if stripped.startswith('If ') and stripped.endswith(' Then'):
            depth += 1
        elif stripped == 'End If':
            depth -= 1
    assert depth == 0, 'Renderer initialization must not depend on missing templates'
    assert renderer_guard < prepare.index('Set m_pebDslRenderer = New obj_PEB_DslRenderer')
    assert renderer_guard < prepare.index('m_pebDslRenderer.TryRender(')
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