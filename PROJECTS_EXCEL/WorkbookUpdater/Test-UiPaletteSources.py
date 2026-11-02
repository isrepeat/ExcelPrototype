"""Check palette/page integration without opening Excel or executing VBA."""

from pathlib import Path
import re
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
NS = "{urn:excelprototype:profiles}"
RESOURCE = re.compile(r"\{ThemeResource ([A-Za-z0-9_.-]+)\}")
COLOR_PROPERTIES = {"backcolor", "textcolor", "fontcolor", "bordercolor", "shadowcolor"}
page_count = 0
reference_count = 0

for path in (ROOT / "ui").rglob("*.xaml"):
    page = ET.parse(path).getroot()
    if page.tag != NS + "page" or "paletteSources" not in page.attrib:
        continue
    palettes = {}
    for source in page.attrib["paletteSources"].split(";"):
        assert source.strip(), (path, "Empty palette source")
        palette_path = path.parent / source
        palette = ET.parse(palette_path).getroot()
        assert palette.tag == NS + "palette", palette_path
        assert palette.attrib["version"] == "1", palette_path
        name = palette.attrib["name"]
        assert name and name not in palettes, (palette_path, "Duplicate/empty palette")
        colors = {}
        for color in palette:
            assert color.tag == NS + "color" and len(color) == 0, palette_path
            assert set(color.attrib) == {"name", "value"}, palette_path
            key, value = color.attrib["name"], color.attrib["value"]
            assert re.fullmatch(r"[A-Za-z0-9_.-]+", key) and key not in colors, palette_path
            assert re.fullmatch(r"#[0-9A-Fa-f]{6}(?:[0-9A-Fa-f]{2})?", value), palette_path
            colors[key] = value
        assert colors, palette_path
        palettes[name] = colors
    contract = set(next(iter(palettes.values())))
    assert all(set(colors) == contract for colors in palettes.values()), path
    documents = [page]
    for source in page.attrib.get("styleSources", "").split(";"):
        if source.strip():
            documents.append(ET.parse(path.parent / source).getroot())
    for document in documents:
        for node in document.iter():
            declarations = [(key.lower(), value) for key, value in node.attrib.items()]
            block = node.attrib.get("styles", "").strip()
            if block:
                assert block.startswith("{") and block.endswith("}"), (path, block)
                declarations += [tuple(pair.split(":", 1)) for pair in block[1:-1].split(";") if pair.strip()]
            for key, value in declarations:
                if key.strip().lower() not in COLOR_PROPERTIES:
                    continue
                match = RESOURCE.fullmatch(value)
                assert match, (path, "Color must reference a palette", key, value)
                resource = match[1]
                assert resource in contract, (path, "Undefined color", resource)
                if key.strip().lower() != "backcolor":
                    assert all(len(colors[resource]) == 7 for colors in palettes.values()), resource
                reference_count += 1
    page_count += 1

tdb = ROOT / "ui/TravelDocumentsBuilder/MainPage.xaml"
page = ET.parse(tdb).getroot()
config = {}
for line in (ROOT / "config/TravelDocumentsBuilder/wsConfig.txt").read_text(encoding="utf-8-sig").splitlines():
    kind, key, value = line.split("\t")
    assert key not in config, ("Duplicate config key", key)
    config[key] = value
assert config["TravelDocumentsBuilder::ui.theme"] in {"Dark", "Light"}
assert config["TravelDocumentsBuilder::text.ToggleTheme"]
assert "{detail}" in config["TravelDocumentsBuilder::text.ThemeFailed"]
assert page.attrib["palette"] == "{Binding Path=Settings.Theme}"
button = next(node for node in page.iter() if node.attrib.get("name") == "ToggleTheme")
assert button.attrib["command"] == "{Binding Path=Commands.ToggleTheme}"
assert button.attrib["caption"] == "{Binding Path=Text.ToggleTheme}"
assert (button.attrib["row"], button.attrib["column"], button.attrib["columnSpan"]) == ("1", "13", "2")
assert not any(node.attrib.get("name") == "ActionsSection" for node in page.iter())
assert len("btn_" + button.attrib["name"]) <= 31
assert button.attrib["style"] == "themeToggleButton"
toggle_style = next(node for node in page.iter() if node.attrib.get("name") == "themeToggleButton")
assert toggle_style.attrib["fontBold"] == "true"
assert all(toggle_style.attrib[key] == "0" for key in ("marginLeft", "marginRight", "marginTop", "marginBottom", "paddingLeft", "paddingRight", "paddingTop", "paddingBottom"))
caption = config["TravelDocumentsBuilder::text.ToggleTheme"]
assert set(RESOURCE.findall(caption)) == {"ThemeLightText", "ThemeDarkText"}
for name in ("Dark", "Light"):
    colors = {node.attrib["name"]: node.attrib["value"] for node in ET.parse(tdb.parent / (name + ".xaml")).getroot()}
    resolved = RESOURCE.sub(lambda match: colors[match[1]], caption)
    runs = re.findall(r"\[\[color=(#[0-9A-Fa-f]{6})\]\]([^\[]+)\[\[/color\]\]", resolved)
    assert len(runs) == 2 and "{ThemeResource" not in resolved
    light_level = sum(bytes.fromhex(colors["ThemeLightText"][1:]))
    dark_level = sum(bytes.fromhex(colors["ThemeDarkText"][1:]))
    assert light_level < dark_level
palette_source = (ROOT / "vba/common/[2] classes/obj_UiPalette.cls.vba").read_text(encoding="utf-8-sig")
for diagnostic in re.findall(r'fn_Raise "([^"]+)"', palette_source):
    assert config.get("TravelDocumentsBuilder::text.Style" + diagnostic), diagnostic
assert config["TravelDocumentsBuilder::text.StyleThemeContext"]
assert "{name}" in config["TravelDocumentsBuilder::text.ThemeUnknown"]
text_keys = [key.split("::text.", 1)[1] for key in config if key.startswith("TravelDocumentsBuilder::text.")]
assert text_keys == sorted(text_keys, key=lambda key: tuple((len(part), part) for part in key.split(".")))
assert page_count > 0 and reference_count > 0
print(f"PASS: {page_count} palette pages, {reference_count} color references, matching palette contracts and TDB configuration. Excel/VBA were not executed.")