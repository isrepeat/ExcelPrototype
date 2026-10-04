"""Validate style declarations against the VBA compiler without opening Excel."""

from pathlib import Path
import re
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
CLASSES = ROOT / "vba/common/[2] classes"
pipeline = (CLASSES / "obj_UiStylePipeline.cls.vba").read_text(encoding="utf-8-sig")
catalog = (CLASSES / "obj_UiStyleCatalog.cls.vba").read_text(encoding="utf-8-sig")
assert 'Set selected("shape") = Nothing' in pipeline
assert 'Set selected("area") = scope' in pipeline
assert 'Set region("area") = area' in pipeline
assert 'Set region("shape") = shape' in pipeline
assert "For Each region In m_regions.Items" not in pipeline
assert "For Each region In m_overlays.Items" not in catalog


def case_keys(text, anchor):
    start = text.index(anchor)
    match = re.search(r'Case ("[^\n]+)', text[start:])
    return set(re.findall(r'"([a-z]+)"', match[1]))


targets = case_keys(pipeline, 'target = VBA.LCase$')
factory = (CLASSES / "obj_UiElementFactory.cls.vba").read_text(encoding="utf-8-sig")
rule_schema = factory.split("Private Function private_StyleRuleSchema()", 1)[1].split("End Function", 1)[0]
schema_targets = re.search(r'AddAttribute "target", "enum", True, "([^"]+)"', rule_schema)[1]
assert {value.lower() for value in schema_targets.split("|")} == targets, "Rule schema targets differ from compiler"
rule_attributes = set(re.findall(r'AddAttribute "([^"]+)"', rule_schema))
assert {"target", "selector", "style", "styles", "enabled"} <= rule_attributes
assert 'AddAttribute "styles", "styleblock", False' in rule_schema
assert 'RequireAnyAttribute "style|styles"' in rule_schema
for method in ("private_StylePipelineStageSchema", "private_StyleLayerSchema"):
    schema = factory.split("Private Function " + method + "()", 1)[1].split("End Function", 1)[0]
    assert 'AddAttribute "enabled", "boolean", False' in schema, method
selectors = case_keys(pipeline, 'key = VBA.LCase$')
properties_function = catalog.split('Private Function private_IsVisualProperty', 1)[1].split('End Function', 1)[0]
properties = set(re.findall(r'"([a-z]+)"', properties_function))
config = {}
for line in (ROOT / "config/PersonalEventBuilder/wsConfig.txt").read_text(encoding="utf-8").splitlines():
    kind, key, value = line.split("\t")
    config[key] = value
for key in re.findall(r'fn_Raise "([^"]+)"', pipeline + catalog):
    assert config.get("PersonalEventBuilder::text.Style" + key), ("Missing configured style diagnostic", key)
count = 0
for folder in (ROOT / "ui").iterdir():
    if not folder.is_dir():
        continue
    common = folder / "CommonControlStyles.xaml"
    common_names = set()
    if common.exists():
        common_names = {n.attrib["name"] for n in ET.parse(common).getroot().iter()
                        if n.tag.rsplit("}", 1)[-1] == "controlStyle"}
    for path in folder.glob("*.xaml"):
        root = ET.parse(path).getroot()
        style_names = common_names | {n.attrib["name"] for n in root.iter()
                                      if n.tag.rsplit("}", 1)[-1] == "controlStyle"}
        stages = set()
        for stage in root.iter():
            if stage.tag.rsplit("}", 1)[-1] != "stylePipelineStage":
                continue
            name = stage.attrib.get("name", "")
            assert name and name not in stages, (path, "Missing or duplicated stage", name)
            stages.add(name)
            for node in stage.iter():
                assert node.attrib.get("enabled", "true").lower() in {"true", "false"}, (path, "Invalid enabled")
                if node.tag.rsplit("}", 1)[-1] != "rule":
                    continue
                count += 1
                assert node.attrib["target"].lower() in targets, (path, "Unknown target", node.attrib)
                selected = {}
                assert set(node.attrib) <= rule_attributes, (path, "Unsupported rule attribute")
                assert node.attrib.get("style", "").strip() or node.attrib.get("styles", "").strip(), (path, "Rule lacks style/styles")
                for pair in node.attrib.get("selector", "").split(";"):
                    if not pair.strip():
                        continue
                    key, value = pair.split("=", 1)
                    key, value = key.strip().lower(), value.strip()
                    assert key in selectors and value and key not in selected, (path, "Invalid selector", pair)
                    selected[key] = value
                if "elementdepth" in selected:
                    depth = [int(value) for value in selected["elementdepth"].split(":")]
                    assert 1 <= len(depth) <= 2 and 0 <= depth[0] <= depth[-1], (path, "Invalid element depth")
                if node.attrib["target"].lower() in {"controlpart", "inlinepart"}:
                    assert {"type", "part"} <= selected.keys(), (path, "Part selector lacks type/part")
                if "style" in node.attrib:
                    assert node.attrib["style"] in style_names, (path, "Unknown named rule style")
                for pair in node.attrib.get("styles", "").strip().strip("{}").split(";"):
                    if pair.strip():
                        key, value = pair.split(":", 1)
                        assert key.strip().lower() in properties and value.strip(), (path, "Invalid property", pair)
print(f"PASS: {count} pipeline rules validated against VBA targets, selectors and properties; no Excel or VBA execution.")