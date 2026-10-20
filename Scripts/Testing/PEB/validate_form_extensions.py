"""Check configurable extension and history transaction contracts without running Office."""

import xml.etree.ElementTree as ET
import re

from event_routes import ROOT, parse_config
from validate_movement_planner import check_declarations, procedure
from validate_movement_batch import check_blocks, check_module_sections


def main():
    folder = ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder"
    manager = (folder / "obj_PEB_FormExtensions.cls.vba").read_text(encoding="utf-8-sig")
    children = (folder / "ex_PEB_FormChildren.vba").read_text(encoding="utf-8-sig")
    history = (folder / "ex_PEB_History.vba").read_text(encoding="utf-8-sig")
    service = (folder / "obj_PEB_ExportService.cls.vba").read_text(encoding="utf-8-sig")
    controller = (folder / "obj_PEB_PgMainController.cls.vba").read_text(encoding="utf-8-sig")
    revisions = (folder / "ex_PEB_HistoryRevisions.vba").read_text(encoding="utf-8-sig")
    recovery = (folder / "ex_PEB_HistoryRecovery.vba").read_text(encoding="utf-8-sig")
    for source in (manager, children, history, service, controller, revisions, recovery):
        check_declarations(source)
        check_blocks(source)
        assert not re.search(r"\b(?:ByVal|ByRef|Dim)\s+input\b", source, re.IGNORECASE), "Input is a reserved VBA statement keyword"
    prepare_extensions = procedure(service, "TryPrepareExtensions")
    assert "ByVal extensionRows As Collection" in prepare_extensions
    assert "For Each source In extensionRows" in prepare_extensions
    assert 'row("ModeForm") = values("ModeForm")' in prepare_extensions
    assert 'If row.Exists(lookupOutput) Then row.Remove lookupOutput' in prepare_extensions
    assert prepare_extensions.index('row.Remove lookupOutput') < prepare_extensions.index('private_Enrich row, lookupPrefix')
    check_module_sections(children)
    refresh = (folder / "ex_PEB_ExtRefresh.vba").read_text(encoding="utf-8-sig")
    check_declarations(refresh)
    check_blocks(refresh)
    check_module_sections(refresh)
    action = procedure(manager, "HandleAction")
    assert "ex_PEB_ExtRefresh.fn_Request(Me)" in action
    assert "Me.Build(" not in action and "m_page.UpdatePage()" not in action
    assert "Application.OnTime" in refresh and 'Schedule:=False' in refresh
    assert 'node.setAttribute "panelStyle", "formSelectPanel"' in manager
    assert 'node.setAttribute "itemStyle", "formSelectItem"' in manager
    refresh_content = procedure(manager, "TryRefresh")
    assert 'm_page.RefreshContent("ExtensionRows", diagnostic)' in refresh_content
    assert "UpdatePage" not in refresh_content and "RenderActivePage" not in refresh_content
    commands = procedure(manager, "private_Command")
    assert "If m_commands.Exists(name) Then Exit Sub" in commands
    assert ".Dispose" not in procedure(manager, "Build")
    assert '"{Binding Path=Form." & path & "Style}"' in manager
    assert 'fn_Config("forms.lookupDefaultStyle")' in manager
    assert 'style = style & "Lookup"' in manager
    assert 'node.setAttribute "partStylesEnabled", "true"' in manager
    button = (ROOT / "PROJECTS_EXCEL/vba/common/[3] controls/obj_UiButtonControl.cls.vba").read_text(encoding="utf-8-sig")
    check_declarations(button)
    check_blocks(button)
    dispose_button = procedure(button, "Dispose")
    assert "buttonShape.Delete" in dispose_button
    assert "Router.UnregisterShape shapeName" in dispose_button
    assert "UnregisterShapeLayout shapeName" in dispose_button
    assert dispose_button.index("buttonShape.Delete") < dispose_button.index("Set m_buttonShape = Nothing")
    switch_mode = procedure(controller, "private_SwitchMode")
    assert "m_pageBase.UpdatePage()" not in switch_mode
    assert "private_SwitchMode = m_pageBase.FlushUpdates()" in switch_mode
    assert 'RefreshContent("ExtensionMenus", diagnostic, True)' in switch_mode
    assert 'RefreshContent("ExtensionRows", diagnostic, True)' in switch_mode
    load_mode = procedure(controller, "private_LoadMode")
    assert 'If CStr(currentForm) <> formKind Then' in load_mode
    assert switch_mode.index('updates("Form.EventDate") = VBA.vbNullString') < switch_mode.index('m_pageBase.FlushUpdates()')
    render_context = (ROOT / "PROJECTS_EXCEL/vba/common/[2] classes/obj_UiRenderContext.cls.vba").read_text(encoding="utf-8-sig")
    flush = procedure(render_context, "FlushUpdates")
    assert "contentPanel.RebuildContent(diagnostic)" in flush
    assert "If Not registered Is oldElement Then pending(key) = True" in flush
    assert "If covered.CountLarge = area.CountLarge Then pending.Remove childKey" in flush
    assert flush.index("contentPanel.RebuildContent") < flush.index("m_root.Measure")
    factory = (ROOT / "PROJECTS_EXCEL/vba/common/[2] classes/obj_UiElementFactory.cls.vba").read_text(encoding="utf-8-sig")
    check_declarations(factory)
    schema = procedure(factory, "private_CreateSchema")
    stack_schema = schema.split('Case "stackpanel"', 1)[1].split("Case Else", 1)[0]
    assert 'schema.AddAttribute "contentSource"' in stack_schema
    assert "schema.AddAttribute" not in procedure(factory, "Create")
    profiles = []
    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        config = parse_config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder" / filename)
        for form, index in (("Leave", 12), ("Treatment", 16)):
            date_prefix = f"forms.{form}.fields.{index}"
            assert config[date_prefix + ".name"] == "EventDate"
            assert date_prefix + ".record" not in config
            assert date_prefix + ".sourceAlias" not in config
        extension = {k: v for k, v in config.items() if k.startswith("extensions.")}
        profiles.append(extension)
        assert set(config["extensions.types"].split("|")) == {"TVO", "Report", "Document"}
        assert 1 <= int(config["extensions.limit"]) <= 100
        saved = {field.split("=")[0] for field in config["history.fields"].split("|")}
        for kind in config["extensions.types"].split("|"):
            prefix = "extensions.type." + kind + "."
            fields = config[prefix + "fields"].split("|")
            assert len(fields) == len(set(fields))
            assert set(fields) <= saved
            assert set(config[prefix + "required"].split("|")) <= set(fields)
            assert config[prefix + "key"] in fields
            mappings = dict(field.split("=") for field in config[prefix + "map"].split("|"))
            assert set(mappings) <= set(fields)
            profile = "lookup." + config[prefix + "profile"]
            names = {config[profile + ".columns." + str(i) + ".name"]
                     for i in range(1, int(config[profile + ".columns.count"]) + 1)}
            assert set(mappings.values()) <= names
            for field in fields:
                assert "text.Mode" + field + "Label" in config
                # Existing field/shape prefixes must fit Excel's 31-character limit.
                assert len("fld_Ex100_" + field) <= 31
            for form in ("Leave", "Treatment"):
                count = int(config["forms." + form + ".fields.count"])
                columns = {config["forms." + form + ".fields." + str(i) + ".name"]: i
                           for i in range(1, count + 1)}
                column = columns.get(config[prefix + "start"], int(config[prefix + "fallbackColumn"]))
                assert 1 <= column and column + len(fields) - 1 <= 16
        assert "HistoryReports" in config["export.shared.inputs.collections"].split("|")
    assert profiles[0] == profiles[1], "PROD/TEST extension contracts differ"

    page = ET.parse(ROOT / "PROJECTS_EXCEL/ui/PersonalEventBuilder/MainPage.xaml").getroot()
    ns = {"p": "urn:excelprototype:profiles", "c": "urn:excelprototype:controls"}
    workspace = page.find(".//p:stackPanel[@name='EventWorkspace']", ns)
    assert page.find(".//c:label[@name='PersonnelSection']", ns) is None
    order = [node.get("name") for node in workspace]
    assert order.index("ExtensionMenus") < order.index("PersonnelForm") < order.index("ExtensionRows") < order.index("RecordPanels")
    assert workspace.find("p:stackPanel[@name='ExtensionRows']", ns).get("contentSource") == "ExtensionBody"
    assert 'payload("Types")' in manager  # One menu per anchor supports multiple configured choices.
    assert 'private_Search(row, diagnostic)' in manager
    # Each precreated table keeps its own projection; searches never rebuild live forms.
    search = procedure(manager, "private_Search")
    publish = procedure(manager, "private_PublishCandidates")
    build = procedure(manager, "Build")
    assert '"ExCandidates" & row("Id"), projected' in publish
    assert '"ExSelected" & row("Id"), vbNullString' in publish
    assert '"{Binding Path=Data.ExCandidates" & row("Id") & "}"' in build
    assert '"{Binding Path=Data.ExSelected" & row("Id") & "}"' in build
    assert 'node.setAttribute "name", "ExCandidates" & row("Id")' in build
    assert 'If m_active("Id") <> row("Id") Then' in action
    search_action = action.split('Case "search"', 1)[1].split('Case "select"', 1)[0]
    assert 'Exit Function' in search_action.split('private_Search(row, diagnostic)', 1)[1]
    assert 'fn_Request' not in search_action
    assert 'private_PublishCandidates(row, Nothing, diagnostic)' in procedure(manager, "private_Add")
    assert '"ExVisible" & row("Id"), "Visible"' in search
    assert '"Data", "ExtCandidates"' not in manager
    assert 'ExecuteWithPayload' in (ROOT / "PROJECTS_EXCEL/vba/common/[2] classes/obj_UiCommand.cls.vba").read_text(encoding="utf-8-sig")
    assert 'm_extensions.TryLoad(row("FormExtensionRows"), diagnostic)' in controller
    save = procedure(history, "fn_Save")
    assert save.index("fn_Plan(") < save.index("ex_PEB_HistoryRevisions.fn_Prepare")
    assert save.index("ex_PEB_HistoryRecovery.fn_Prepare") < save.index("target.Range.Value2 = data")
    assert save.index("ex_PEB_FormChildren.fn_Write") < save.index('recoveryToken, "COMMIT"')
    assert "ex_PEB_FormChildren.fn_Rollback table, childPlan" in save
    assert 'transaction.Add "Children", childPlan' in save
    assert 'children("Old")' in recovery and 'children("New")' in recovery
    assert 'private_Append(source, table, nextRows, values' in revisions
    assert 'root("FormExtensionRows").Add row' in history
    assert 'If people.Exists(row("PersonId"))' in children
    assert 'If Not changed Then Exit Function' in children
    assert 'Set values("FormExtensionRows") = children' in service
    assert "MovementHistoryReportsUnsupported" not in procedure(service, "TryPrepareHistoryMovement")
    print("PASS: extension configuration, layout, isolated lookup routing, frozen child fields, revisions, recovery and rollback contracts. Office/VBA not executed.")


if __name__ == "__main__":
    main()