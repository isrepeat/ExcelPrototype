"""Validate container visibility schemas and fixture contracts without Office/VBA."""

import re
from xml.etree import ElementTree as ET

from event_routes import ROOT
from validate_movement_batch import check_blocks, check_module_contract
from validate_movement_planner import check_declarations


def main():
    common = ROOT / "PROJECTS_EXCEL/vba/common"
    factory = (common / "[2] classes/obj_UiElementFactory.cls.vba").read_text(encoding="utf-8-sig")
    schema = factory.split("Private Function private_CreateSchema()", 1)[1]
    schema = schema.split("Private Function private_TrackSchema", 1)[0]
    allowed = {}
    for kind in ("grid", "stackpanel"):
        branch = re.search(rf'Case "{kind}"\s*\n(.*?)(?=\n\s*Case |\n\s*End Select)', schema, re.S)[1]
        allowed[kind] = set(re.findall(r'schema.AddAttribute "([^"]+)"', branch))
        assert 'schema.AddAttribute "visibility", "enum", False, "Visible|Collapsed", True, 0' in branch
    count = 0
    for path in (ROOT / "PROJECTS_EXCEL/ui").rglob("*.xaml"):
        for node in ET.parse(path).getroot().iter():
            kind = node.tag.removeprefix("{urn:excelprototype:profiles}").lower()
            if kind not in allowed:
                continue
            assert set(node.attrib) <= allowed[kind], (path, node.attrib.keys() - allowed[kind])
            visibility = node.get("visibility", "Visible")
            assert visibility in ("Visible", "Collapsed") or visibility.startswith("{Binding "), (path, visibility)
            count += 1
    helper = common / "[2] classes/obj_UiPanelVisibility.cls.vba"
    panels = [common / f"[4] elements/obj_Ui{kind}Element.cls.vba" for kind in ("Grid", "StackPanel")]
    for path in [helper, *panels]:
        text = path.read_text(encoding="utf-8-sig")
        check_declarations(text)
        check_blocks(text)
        assert len(re.search(r'Attribute VB_Name = "([^"]+)"', text)[1]) <= 31
    fixture = ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder/Test/ex_UiPanelVisibilityTests.vba"
    text = fixture.read_text(encoding="utf-8-sig")
    check_module_contract(text, "ex_UiPanelVisibilityTests")
    print(f"PASS: grid/stackPanel schemas, {count} real container definitions and visibility fixture source. Office/VBA were not executed.")


if __name__ == "__main__":
    main()