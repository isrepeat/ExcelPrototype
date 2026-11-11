"""Check configured export requirements and snapshot-preserving refresh without Office."""

from copy import deepcopy
from xml.etree import ElementTree as ET

from event_routes import ROOT, parse_config
from movement_plans import load_native_plans
from validate_movement_planner import check_declarations, procedure
from validate_movement_batch import check_blocks


def main():
    folder = ROOT / "PROJECTS_EXCEL"
    read = lambda path: (folder / path).read_text(encoding="utf-8-sig")
    registry = read("vba/PersonalEventBuilder/ex_PEB_MovementRegistry.vba")
    for name in ("wsConfig.txt", "wsConfig.test.txt"):
        settings = parse_config(folder / "config/PersonalEventBuilder" / name)
        plans = load_native_plans(settings)
        for identifier, plan in plans["events"].items():
            definition = plans["generated"].get(identifier)
            profile_id = definition[0] if definition else identifier
            key = f"export.Movement.{profile_id}.required"
            assert plan["required"] == settings[key]
            changed = deepcopy(settings)
            changed[key] = "PersonId|PersonName"
            assert load_native_plans(changed)["events"][identifier]["required"] == changed[key]
        changed = deepcopy(settings)
        changed["export.Movement.headers.PersonName"] = "Custom person column"
        assert "Custom person column" in load_native_plans(changed)["values"].values()
        mode = "MoveFromVacationForCuringToVlk"
        assert "HospitalShortName" not in plans["events"][mode]["required"].split("|")
        assert settings[f"export.Word.{mode}.lookups.1.when.field"] == "HospitalShortName"
        for key in ("RefreshConfig", "RefreshConfigReady", "RefreshConfigPending", "ConfigSourceChanged"):
            assert settings["text." + key]
    paths = (
        "vba/PersonalEventBuilder/obj_PEB_PgMainController.cls.vba",
        "vba/PersonalEventBuilder/obj_PEB_PgMain.cls.vba",
        "vba/PersonalEventBuilder/obj_PEB_DependencyLoader.cls.vba",
        "vba/common/modules/ex_AppCommands.vba",
        "vba/common/[2] classes/obj_TableSnapshotSet.cls.vba",
    )
    for path in paths:
        text = read(path)
        check_declarations(text)
        check_blocks(text)
    check_declarations(registry)
    check_blocks(registry)
    controller = read(paths[0])
    refresh = procedure(controller, "RefreshConfigHandler")
    assert 'loader.TryBuild(snapshot, diagnostic, True)' in refresh
    assert 'If Not snapshot Is Nothing Then' in refresh
    assert 'requirementsOnly' not in refresh
    assert 'ex_TableSnapshots.fn_Register loader, ex_PEB_Export.fn_Config("text.DependenciesRequired"), True' in refresh
    assert refresh.index('If exportOnly Then') < refresh.index('fn_RequestReinitialize(True)')
    live = procedure(controller, "private_ExportConfigOnly")
    assert 'previous.Keys' in live and 'current.Keys' in live
    assert '"export."' in live
    assert '"text."' in live
    assert 'private_TemplatesOnly(previousConfig, m_modeConfig)' in refresh
    template_branch = refresh.split('If templatesOnly Then', 1)[1].split('Else', 1)[0]
    assert 'loader.TryRefreshTemplates(snapshot, diagnostic)' in template_branch
    assert 'TryBuild' not in template_branch and 'fn_Refresh' not in template_branch
    assert 'private_RefreshTexts(previousConfig, diagnostic)' in refresh
    assert 'ex_AppCommands.fn_PreserveDependencies()' in controller
    loader = procedure(read(paths[2]), "TryBuild")
    templates = procedure(read(paths[2]), "TryRefreshTemplates")
    assert templates.index('templates.Initialize(diagnostic)') < templates.index('snapshot.ReplacePrepared prepared')
    assert templates.index('fn_BuildSnapshot()') < templates.index('snapshot.ReplacePrepared prepared')
    assert 'prepared.Add key, snapshot.Prepared(key)' in templates
    assert all(token not in templates for token in ('TryLoad', 'WarmIndex', 'PrepareSnapshot', 'fn_Refresh'))
    assert 'If reuseLoaded Then' in loader
    assert loader.index('snapshot.Contains(source)') < loader.index('snapshot.TryLoad(source, diagnostic)')
    assert 'snapshot.ReplacePrepared previousPrepared' in loader
    callback = procedure(read(paths[3]), "fn_ReinitializePage")
    assert 'If Not m_preserveDependencies Then ex_TableDataRuntime.fn_ResetCaches' in callback
    assert 'm_preserveDependencies = False' in callback
    assert 'Set preserved = ex_TableSnapshots.fn_Active()' in callback
    assert 'ex_TableSnapshots.fn_RestoreActive preserved' in callback
    page = read(paths[1])
    assert '"RefreshConfig"' in page
    markup = ET.fromstring(read("ui/PersonalEventBuilder/MainPage.xaml"))
    buttons = {node.get("name"): node for node in markup.iter() if node.tag.endswith("}button")}
    button = buttons["RefreshConfig"]
    assert 'ReinitializePage' not in buttons
    assert button.get("command") == "{Binding Path=Commands.RefreshConfig}"
    for name in ("RefreshConfig", "DeleteEvent", "DeleteLastEvent"):
        assert len("btn_" + name) <= 31
    assert len({buttons[name].get("row") for name in ("RefreshConfig", "DeleteEvent", "DeleteLastEvent")}) == 3
    print("PASS: configured requirements, column aliases, live export refresh, source reuse and page wiring. Office/VBA not executed.")


if __name__ == "__main__":
    main()