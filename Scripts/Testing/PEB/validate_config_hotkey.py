"""Check the external configuration hotkey without executing Excel or VBA."""

from config_dependencies import ROOT, rows, shared_rows


def main():
    refresh = (ROOT / "PROJECTS_EXCEL/WorkbookUpdater/ex_ConfigRefresh.vba").read_text(encoding="utf-8-sig")
    proxy = (ROOT / "PROJECTS_EXCEL/vba/common/modules/ex_ConfigRefresh.vba").read_text(encoding="utf-8-sig")
    personal = (ROOT / "MacrosExcel/PERSONAL/ex_Core.vba").read_text(encoding="utf-8-sig")
    bindings = (ROOT / "MacrosExcel/PERSONAL/ThisWorkbook.vba").read_text(encoding="utf-8-sig")
    build = (ROOT / "PROJECTS_EXCEL/WorkbookUpdater/Build-Updater.ps1").read_text(encoding="utf-8-sig")
    assert 'private_TryBindKey("^%u", "ex_Core.fn_UpdateActiveWorkbookConfig")' in bindings
    assert 'private_TryUnbindKey "^%u"' in bindings
    handler = personal.split("Public Sub fn_UpdateActiveWorkbookConfig()", 1)[1].split("End Sub", 1)[0]
    assert 'Set targetWorkbook = Application.ActiveWorkbook' in handler
    assert 'ex_ConfigRefresh.fn_UpdateWorkbook", targetWorkbook' in handler
    assert '.VBProject' not in handler and 'cfgDependencies' not in handler
    assert "'WorkbookUpdater.xlam'!ex_ConfigRefresh.fn_TryUpdateWorkbook" in proxy
    assert 'ThisWorkbook, diagnostic, sourcePath' in proxy
    assert 'tbConfig' not in proxy and 'ADODB.Stream' not in proxy
    service = refresh.split("Public Function fn_TryUpdateWorkbook(", 1)[1].split("End Function", 1)[0]
    assert 'ThisWorkbook.Path' not in service and 'ThisWorkbook.Worksheets' not in service
    assert 'targetWorkbook.Path' in service and 'targetWorkbook.Worksheets' in service
    assert service.index('If requiresImports Then') < service.index('ex_ConfigImports.fn_TryExpand')
    assert service.index('Configuration profile mismatch') < service.index('changed = True')
    assert 'If changed Then' in service and 'originalRange.Value2 = originalValues' in service
    assert 'Application.EnableEvents = previousEvents' in service
    assert 'Application.ScreenUpdating = previousUpdating' in service
    assert '#Const CONFIG_QUERY_AVAILABLE = False' in build
    assert "../vba/common/modules/ex_ConfigImports.vba" in build
    import json
    manifest = json.loads((ROOT / 'PROJECTS_EXCEL/WorkbookUpdater/modules.json').read_text(encoding='utf-8-sig'))
    assert 'ex_ConfigRefresh.vba' in manifest['WorkbookUpdater']
    updater = (ROOT / 'PROJECTS_EXCEL/WorkbookUpdater/ex_WorkbookUpdater.vba').read_text(encoding='utf-8-sig')
    first_refresh = updater.index('ex_ConfigRefresh.fn_TryUpdateWorkbook(target, configDiagnostic)')
    assert first_refresh < updater.index('Set m_plan = private_ReadPlan(target, m_uiFolder)')
    assert first_refresh < updater.index('m_context("StopRequested") = True')
    assert 'ex_ConfigRefresh.fn_TryUpdateWorkbook(m_target, configDiagnostic)' in updater
    assert 'private_Macro("ex_ConfigRefresh.' not in updater
    launcher = (ROOT / 'Scripts/Update-WorkbookUpdater.cmd').read_text(encoding='utf-8-sig')
    assert 'Build-Updater.ps1" -UpdateExisting' in launcher
    assert 'if ($UpdateExisting)' in build
    imports = (ROOT / 'PROJECTS_EXCEL/vba/common/modules/ex_ConfigImports.vba').read_text(encoding='utf-8-sig')
    enabled = True
    included = []
    for line in imports.splitlines():
        if line == '#If CONFIG_QUERY_AVAILABLE Then':
            enabled = False
        elif line == '#If Not CONFIG_QUERY_AVAILABLE Then':
            enabled = True
        elif line == '#Else':
            enabled = True
        elif line == '#End If':
            enabled = True
        elif enabled:
            included.append(line)
    compiled = '\n'.join(included)
    assert 'obj_Table' not in compiled and 'QueryAuto' not in compiled and 'ex_TableQuery.' not in compiled
    assert 'ex_Core.' not in refresh
    assert 'If opened Then book.Close SaveChanges:=False' in compiled
    assert 'Application.EnableEvents = previousEvents' in compiled
    source = ROOT / 'PROJECTS_EXCEL/config/cfgDependencies/wsConfig.txt'
    book = ROOT / 'PROJECTS_EXCEL/workbook/cfgDependencies.xlsx'
    settings = {key: value for kind, key, value in rows(source)}
    assert settings['ThisWorkbook::id'] == 'cfgDependencies'
    assert (book.parent / settings['ThisWorkbook::config.source']).resolve() == source
    assert (book.parent / settings['ThisWorkbook::vbaPath']).resolve() == ROOT / 'PROJECTS_EXCEL/vba'
    assert settings['cfgDependencies::configRefresh.text.Success']
    assert settings['cfgDependencies::configRefresh.text.Failed']
    assert shared_rows(book) == rows(source)
    print('PASS: global hotkey binding, active workbook capture, shared refresh source, add-in independence, metadata paths and rollback guards. VBA not executed.')


if __name__ == '__main__':
    main()