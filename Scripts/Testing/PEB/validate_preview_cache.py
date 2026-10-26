"""Check shared dependency snapshots, prepared ownership and live history without Office."""

from pathlib import Path
import re
import subprocess

from config_dependencies import load_config
from validate_movement_planner import check_declarations, procedure
from validate_movement_batch import check_blocks, check_module_sections

ROOT = Path(__file__).resolve().parents[3]
VBA = ROOT / "PROJECTS_EXCEL/vba"


def read(path):
    return (VBA / path).read_text(encoding="utf-8-sig")


def main():
    provider = read("common/modules/ex_TableSnapshots.vba")
    snapshot = read("common/[2] classes/obj_TableSnapshotSet.cls.vba")
    loader = read("PersonalEventBuilder/obj_PEB_DependencyLoader.cls.vba")
    service = read("common/[2] classes/obj_TableQueryService.cls.vba")
    controller = read("PersonalEventBuilder/obj_PEB_PgMainController.cls.vba")
    history = read("PersonalEventBuilder/ex_PEB_History.vba")
    movement = read("PersonalEventBuilder/ex_PEB_MovementSource.vba")
    export = read("PersonalEventBuilder/obj_PEB_ExportService.cls.vba")
    lookup = read("common/[2] classes/obj_LookupService.cls.vba")
    orders = read("common/[2] classes/obj_OrderResolver.cls.vba")
    tickets = read("PersonalEventBuilder/obj_PEB_ReturnTickets.cls.vba")
    refresh = procedure(provider, "fn_Refresh")
    current = procedure(provider, "fn_Current")
    assert "fn_Refresh" not in current and "Err.Raise 5, , m_notLoadedMessage" in current
    initialize = procedure(controller, "Initialize")
    assert "SearchOrdersHandler" not in initialize and "fn_Current" not in initialize
    assert 'text.DependenciesRequired' in initialize
    assert refresh.index("m_loader.TryBuild(staged, diagnostic)") < refresh.index("Set m_active = staged")
    failed = refresh.split("EH:", 1)[1]
    assert "Set m_active" not in failed and "fn_Module_Dispose" not in failed
    assert "useSnapshot As Boolean = True" in service
    assert service.index("ex_TableSnapshots.fn_TryQuery") < service.index("selected = Backend")
    assert "False)" in procedure(snapshot, "TryLoad")
    assert "service.Backend = QueryAdo" in snapshot
    query = procedure(snapshot, "TryQuery")
    assert "query.Filter.IndexCondition" in query and "query.Filter.Validate table.Headers" in query
    assert "ex_TableQuery.fn_TryApply(subset, query" in query
    index = procedure(snapshot, "private_Index")
    assert "rows.Add row" in index and "index.CompareMode = VBA.vbTextCompare" in index
    assert "m_indexes.Add identity, index" in index
    assert "m_generation = m_generation + 1" in refresh
    for key in ("WordTemplates", "MovementTemplates", "Tvo.Chains", "ReturnTickets"):
        assert '"' + key + '"' in loader
    assert "candidates.PrepareSnapshot(snapshot" in loader
    assert "orders.PrepareSnapshot(snapshot" in loader
    assert "lookup.PrepareSnapshot(snapshot" in loader
    tvo = procedure(loader, "private_PrepareTvo")
    assert 'key = private_Text(table.ValueAt(row, codeColumn))' in tvo
    assert 'key = private_Text(table.ValueAt(row, ipnColumn))' not in tvo
    chain = procedure(loader, "private_TryChain")
    assert 'seen.Exists(VBA.CStr(current))' in chain
    assert 'names.Exists(actor)' in chain and 'matches.Count <> 1' in chain
    assert 'snapshot.Prepared.Add "Tvo.ChainErrors", errors' in tvo
    assert 'text.ReturnTvoInvalidChain' not in tvo
    helpers = read("PersonalEventBuilder/ex_PEB_Export.vba")
    assert '"Tvo.ReporterPositions"' not in helpers
    reporter = procedure(helpers, "fn_TryReporterTvoPositions")
    assert 'ex_TableDataRuntime.fn_TryQueryLoaded(source, query' in reporter
    chain_read = procedure(helpers, "fn_TvoChain")
    assert chain_read.index('errors.Exists(positionCode)') < chain_read.index('chains.Exists(positionCode)')
    assert 'Err.Raise 5, , errors(positionCode)' in chain_read
    movement_templates = read("PersonalEventBuilder/ex_PEB_MovementTemplates.vba")
    template_load = procedure(movement_templates, "private_EnsureLoaded")
    assert 'ex_TableSnapshots.fn_Active()' in template_load
    assert 'ex_TableSnapshots.fn_Current()' not in template_load
    assert "snapshot.Prepared" in lookup and "snapshot.Prepared" in orders
    assert "snapshot.Prepared" in export
    assert "private_LookupTable" not in export and "private_TemplateStamp" not in export
    assert not any(token in movement for token in ("ADODB", "m_periods", "GetFile", "Application.Workbooks"))
    assert "service.TryExecute(source, query, table, diagnostic, Not includeTvo)" in movement
    assert not any(token in tickets for token in ("GetFile", "CopyFile", "DeleteFile"))
    assert "m_loadingSnapshot.TryQuery" in tickets
    assert "ex_TableSnapshots" not in history
    save = procedure(history, "fn_Save")
    assert "private_Open book, table" in save and "private_Read(table, columns" in save
    button = procedure(controller, "RefreshPersonnelHandler")
    assert "ex_TableSnapshots.fn_Refresh(diagnostic)" in button
    assert "Me.SearchOrdersHandler()" in button
    assert "fn_ResetCaches" not in button
    assert "ex_PEB_Testing.fn_IsRunning()" in button
    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        config = load_config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder" / filename)
        get = lambda key: config["PersonalEventBuilder::" + key]
        sources = get("snapshots.sources").split("|")
        assert not any("history" in prefix.lower() for prefix in sources)
        for prefix in sources:
            for suffix in ("path", "sheet", "range"):
                assert get(prefix + "." + suffix), (filename, prefix, suffix)
        assert get("lookup.ReturnTickets.source.range") == get("validation.Movement.source.range")
        for mapping in get("snapshots.indexes").split("|"):
            source, header = mapping.split("=")
            assert source in sources and get(header)
        assert "reporter.tvo.source" in sources
        assert get("reporter.tvo.source.path") == get("validation.Movement.source.path")
        assert get("text.DependenciesReady") and get("text.DependenciesLoading")
        assert get("text.DependenciesRequired")
        for stage in re.findall(r'private_Progress progressStep, progressTotal, "([^"]+)"', loader):
            assert get("snapshots.progress." + stage)
        assert all(token in get("text.DependenciesProgress") for token in ("{step}", "{total}", "{stage}"))
        rows = (ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder" / filename).read_text(encoding="utf-8-sig").splitlines()
        assert all(len(row.split("\t")) == 3 for row in rows if row.strip())
    paths = set(subprocess.check_output(["git", "diff", "--name-only"], cwd=ROOT, text=True).splitlines())
    paths.update(subprocess.check_output(["git", "ls-files", "--others", "--exclude-standard"], cwd=ROOT, text=True).splitlines())
    for path in sorted(paths):
        if path.endswith(".vba"):
            text = (ROOT / path).read_text(encoding="utf-8-sig")
            # Existing class data members can be public; include them in declaration checks.
            declarations = re.sub(r"(?m)^Public (\w+) As ", r"Private \1 As ", text)
            check_declarations(declarations)
            check_blocks(text)
    check_module_sections(provider)
    print("PASS: atomic publication, shared tables/derived state, indexed duplicate-preserving queries, source coverage, live EventsHistory and changed VBA structure. VBA not executed.")


if __name__ == "__main__":
    main()