"""Validate page mappings, declarative controls and page ownership without Office."""

from pathlib import Path
from xml.etree import ElementTree as ET
import re


ROOT = Path(__file__).resolve().parents[3]
UI = ROOT / "PROJECTS_EXCEL/ui/PersonalEventBuilder"
VBA = ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder"


def read_config(path):
    result = {}
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        if not line:
            continue
        kind, key, value = line.split("\t")
        assert key not in result, key
        result[key] = value
    return result


def resolve_mappings(value, sheets, folder):
    if value is None:
        return {sheet: folder / f"{sheet}.xaml" for sheet in sheets}
    assert value.strip(), "empty mapping"
    result = {}
    for entry in value.split("|"):
        parts = entry.split("=", 1)
        assert len(parts) == 2, "malformed mapping"
        sheet, relative = (part.strip() for part in parts)
        assert sheet and relative, "incomplete mapping"
        assert sheet.casefold() not in result, "duplicate sheet"
        assert sheet.casefold() in {name.casefold() for name in sheets}, "missing sheet"
        assert ":" not in relative and not relative.startswith(("/", "\\")), "absolute path"
        path = (folder / relative.replace("\\", "/")).resolve()
        assert path.is_relative_to(folder.resolve()), "outside folder"
        assert path.suffix.lower() == ".xaml", "wrong extension"
        assert path.is_file(), "missing XAML"
        result[sheet.casefold()] = path
    return result


def main():
    for filename, page_class, page_file in (
        ("wsConfig.txt", "obj_PEB_PgMain", "MainPage.xaml"),
        ("wsConfig.test.txt", "obj_PEB_PgTests", "TestsPage.xaml"),
    ):
        cfg = read_config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder" / filename)
        assert cfg["PersonalEventBuilder::ui.folder"] == "PersonalEventBuilder"
        assert cfg["PersonalEventBuilder::ui.pageClass"] == page_class
        paths = resolve_mappings(cfg["ThisWorkbook::ui.pageMappings"], {"MainPage", "wsConfig"}, UI)
        assert paths == {"mainpage": (UI / page_file).resolve()}
        source = (VBA / f"{page_class}.cls.vba").read_text(encoding="utf-8-sig")
        assert "Implements obj_IPage" in source and "m_pageBase.Render()" in source
        assert "m_controller.Initialize(m_pageBase)" in source

    multiple = resolve_mappings("MainPage=MainPage.xaml|Tests=TestsPage.xaml", {"MainPage", "Tests"}, UI)
    assert len(multiple) == 2
    assert resolve_mappings(None, {"MainPage"}, UI)["MainPage"] == UI / "MainPage.xaml"
    rejected = 0
    for value in (
        "", "MainPage", "MainPage=", "=TestsPage.xaml",
        "MainPage=MainPage.xaml|mainpage=TestsPage.xaml",
        "Missing=TestsPage.xaml", "MainPage=missing.xaml",
        "MainPage=../PersonalEventBuilder/../../outside.xaml",
        "MainPage=C:\\outside.xaml", "MainPage=\\outside.xaml",
        "MainPage=TestsPage.xml", "MainPage=TestsPage.xaml|",
    ):
        try:
            resolve_mappings(value, {"MainPage"}, UI)
        except AssertionError:
            rejected += 1
        else:
            raise AssertionError(f"Invalid mapping accepted: {value}")

    page = ET.parse(UI / "TestsPage.xaml").getroot()
    commands = set()
    for node in page.iter():
        for attribute in ("command", "onChange"):
            value = node.attrib.get(attribute, "")
            if value.startswith("{Binding Path=Commands."):
                commands.add(value.removeprefix("{Binding Path=Commands.").removesuffix("}"))
    assert commands == {"TestRun", "CancelTests", "EnvironmentSwitch", "SelectionChanged", "TestPreview"}
    controller = (VBA / "obj_PEB_PgTestsController.cls.vba").read_text(encoding="utf-8-sig")
    for command in commands:
        assert f"Public Function {command}Handler(" in controller, command
    for forbidden in ("createElement", "Shapes.Add", "control.Arrange", "ex_UiControlFactory", "OpenPage"):
        assert forbidden not in controller, forbidden
    regular = (VBA / "obj_PEB_PgMain.cls.vba").read_text(encoding="utf-8-sig")
    assert "ex_PEB_Testing.fn_Open" not in regular
    assert '"Data", "TestCatalog"' not in regular
    scheduler = (VBA / "ex_PEB_Testing.vba").read_text(encoding="utf-8-sig")
    assert "New obj_PEB_PgTestsController" not in scheduler
    assert "Public Function fn_RunSelected" not in scheduler
    assert "Public Function fn_NavigatePreview" not in scheduler
    assert "m_testController.Dispose" not in scheduler
    assert "ex_PEB_Testing.fn_Detach Me" in controller
    selection = controller.split("Public Function SelectionChangedHandler()", 1)[1].split("End Function", 1)[0]
    assert "private_RefreshSelection" in selection and "private_BindCatalog" not in selection
    refresh = controller.split("Private Sub private_RefreshSelection", 1)[1].split("End Sub", 1)[0]
    assert 'private_MarkDirty "TestCatalog", False' in refresh
    assert "m_catalogTable.SetValueAt" in refresh
    engine_root = ROOT / "PROJECTS_EXCEL/vba/common"
    templates = (engine_root / "[3] controls/obj_UiTableCellTemplates.cls.vba").read_text(encoding="utf-8-sig")
    update = templates.split("Public Function TryRefreshBindings", 1)[1].split("End Function", 1)[0]
    assert "target.RefreshBindings(m_context)" in update
    for forbidden in ("Me.Clear", "TryRender", "child.Render", "Shapes.Add", "RegisterShape"):
        assert forbidden not in update, forbidden
    table = (engine_root / "[3] controls/obj_UiTableControl.cls.vba").read_text(encoding="utf-8-sig")
    update = table.split("Public Function TryRefreshTemplateBindings", 1)[1].split("End Function", 1)[0]
    assert re.search(r"If Not table Is m_table Then\s+Exit Function\s+End If", update)
    assert "private_SourceSchema(m_source, table) <> m_renderedSchema" in update
    assert "m_cellTemplates.CanRefreshBindings(table, m_source)" in update
    assert "m_cellTemplates.TryRefreshBindings(table, diagnostic)" in update
    context = (engine_root / "[2] classes/obj_UiRenderContext.cls.vba").read_text(encoding="utf-8-sig")
    assert "If pending.Count > 0 Then" in context
    assert not (VBA / "obj_PEB_TestUi.cls.vba").exists()
    assert not list((ROOT / "PROJECTS_EXCEL/ui/PersonalEventBuilderTest").glob("*"))

    runtime = (ROOT / "PROJECTS_EXCEL/vba/common/modules/ex_UiRuntime.vba").read_text(encoding="utf-8-sig")
    assert '"ThisWorkbook::ui.pageMappings"' in runtime
    assert "uiFolderPath = fileSystem.GetParentFolderName(xamlPath)" in runtime
    assert "PersonalEventBuilder" not in runtime.split("Private Function private_TryResolveXaml", 1)[1].split("End Function", 1)[0]
    for path in (VBA / "obj_PEB_PgTests.cls.vba", VBA / "obj_PEB_PgTestsController.cls.vba"):
        text = path.read_text(encoding="utf-8-sig")
        procedures = re.findall(r"(?im)^(?:Public|Private) (?:Function|Sub) (\w+)", text)
        assert len(procedures) == len(set(procedures)), path
        assert len(procedures) == len(re.findall(r"(?im)^End (?:Function|Sub)$", text)), path
    print(f"PASS: PROD/TEST mappings, multiple and legacy mappings, {rejected} invalid maps, command routes and declarative test page. Office/VBA not executed.")


if __name__ == "__main__":
    main()