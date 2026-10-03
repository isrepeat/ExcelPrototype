"""Read-only checks for lookup sources; never opens Excel or invokes VBA."""

from collections import Counter
from pathlib import Path
import re
import subprocess
import xml.etree.ElementTree as ET
import zipfile


ROOT = Path(__file__).resolve().parents[1]
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def check(condition, message):
    if not condition:
        raise AssertionError(message)


def declarations(text):
    return re.findall(
        r"(?im)^\s*(?:Public|Private|Friend)?\s*(?:Static\s+)?"
        r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)", text
    )


def validate_sources():
    vba = ROOT / "vba"
    modules = {}
    for path in vba.rglob("*.vba"):
        if path.name == "ThisWorkbook.vba":
            continue
        text = path.read_text(encoding="utf-8-sig")
        name = re.search(r'^Attribute VB_Name = "([^"]+)"', text, re.M)
        module = name[1] if name else path.stem.removesuffix(".cls")
        check(len(module) <= 31, f"VBA module name too long: {module}")
        check(module not in modules, f"Duplicate module: {module}")
        modules[module] = (path, text)
    changed = subprocess.check_output(
        ["git", "diff", "--name-only"], cwd=ROOT, text=True, encoding="utf-8"
    ).splitlines()
    added = subprocess.check_output(
        ["git", "ls-files", "--others", "--exclude-standard"],
        cwd=ROOT, text=True, encoding="utf-8"
    ).splitlines()
    repository = Path(subprocess.check_output(
        ["git", "rev-parse", "--show-toplevel"], cwd=ROOT, text=True
    ).strip())
    text_suffixes = {".vba", ".py", ".ps1", ".md", ".xml", ".xaml", ".json", ".txt"}
    for name in set(changed + added):
        path = repository / name
        if not path.is_file() or path.suffix.lower() not in text_suffixes:
            continue
        content = path.read_bytes()
        check(bool(content), f"Empty changed text file: {name}")
        last_byte = content[-1]
        check(last_byte != 9 and last_byte != 10 and last_byte != 13 and last_byte != 32,
              f"Forbidden last byte {last_byte}: {name}")
        text = content.decode("utf-8-sig")
        check(all(line == line.rstrip() for line in text.splitlines()),
              f"Trailing whitespace: {name}")
        if path.suffix != ".vba":
            continue
        check(not re.search(r"(?im)\b(?:Dim|ByVal|ByRef)\s+attribute\s+As\b", text),
              f"Reserved VBA metadata keyword used as identifier: {name}")
        procedures = declarations(text)
        check(len(procedures) == len(set(procedures)), f"Duplicate procedure: {name}")
        for kind in ("Sub", "Function", "Property"):
            starts = sum(k.startswith(kind) for k, _ in procedures)
            ends = len(re.findall(rf"(?im)^\s*End {kind}\s*$", text))
            check(starts == ends, f"Unbalanced {kind}: {name}")
        private = re.findall(r"(?im)^Private\s+(?:Sub|Function)\s+(\w+)", text)
        for method in private:
            check(not re.search(rf"\bMe\.{re.escape(method)}\b", text, re.I),
                  f"Private method called through Me: {name}: {method}")
        for created in re.findall(r"\bNew\s+(obj_\w+)\b", text, re.I):
            check(created in modules, f"Unknown class: {name}: {created}")
        for contract in re.findall(r"(?im)^Implements\s+(\w+)", text):
            check(contract in modules, f"Unknown interface: {name}: {contract}")
            for _, method in declarations(modules[contract][1]):
                check(re.search(rf"\b{contract}_{method}\b", text, re.I),
                      f"Interface method missing: {name}: {contract}.{method}")
    for name in ("obj_LookupRecord", "obj_LookupProfile", "obj_LookupResult", "obj_LookupService"):
        text = modules[name][1]
        for member in ("Class_Initialize", "Class_Terminate", "Initialize", "Dispose"):
            check(any(method == member for _, method in declarations(text)), f"Missing lifecycle: {name}.{member}")
        check("Friend Function private_" not in text, f"Helper uses inaccessible Friend interface: {name}")
    return modules


def validate_profile():
    path = ROOT / "config/PersonalEventBuilder/wsConfig.txt"
    config = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        parts = line.split("\t")
        check(len(parts) == 3, "Config requires actual tab separators")
        check(parts[1] not in config, "Duplicate configuration key")
        config[parts[1]] = parts[2]
    prefix = "PersonalEventBuilder::lookup.Personnel."
    profile = ET.Element("lookup", {k: config[prefix + k] for k in ("key", "minChars", "maxResults")})
    source = ET.SubElement(profile, "source")
    for k in ("path", "sheet", "range", "table"):
        if prefix + "source." + k in config:
            source.set(k, config[prefix + "source." + k])
    for section, tag in (("columns", "column"), ("display", "column"), ("search", "field"), ("apply", "map")):
        container = ET.SubElement(profile, section)
        for i in range(1, int(config[prefix + section + ".count"]) + 1):
            item_prefix = prefix + section + "." + str(i) + "."
            ET.SubElement(container, tag, {k.removeprefix(item_prefix): v for k, v in config.items() if k.startswith(item_prefix)})
    columns = profile.findall("columns/column")
    names = [c.attrib["name"] for c in columns]
    check(len(names) == len(set(names)), "Duplicate profile field")
    check(profile.attrib["key"] in names, "Key must be a source field")
    check(int(profile.attrib["minChars"]) >= 1, "Invalid minimum search length")
    check(1 <= int(profile.attrib["maxResults"]) <= 1000, "Invalid result limit")
    source = profile.find("source")
    check(source.attrib["sheet"] == "ОС", "Personnel source must use ОС")
    check(source.attrib["range"] == "auto", "Personnel source must grow automatically")
    workbook = (ROOT / "workbook/PersonalEventBuilder" / source.attrib["path"]).resolve()
    check(workbook.is_file(), "Source workbook not found")
    with zipfile.ZipFile(workbook) as archive:
        shared = ["".join(n.itertext()) for n in ET.fromstring(archive.read("xl/sharedStrings.xml"))]
        wb = ET.fromstring(archive.read("xl/workbook.xml"))
        rel_id = next(s.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
                      for s in wb.find("s:sheets", NS) if s.attrib["name"] == source.attrib["sheet"])
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        target = next(r.attrib["Target"] for r in rels if r.attrib["Id"] == rel_id)
        headers = {}
        keys = []
        missing_keys = 0
        people = 0
        full_name_header = next(c.attrib["header"] for c in columns if c.attrib["name"] == "Personnel.FullName")
        key_header = next(c.attrib["header"] for c in columns if c.attrib["name"] == profile.attrib["key"])
        def value(cell):
            node = cell.find("s:v", NS)
            if node is None:
                return ""
            return shared[int(node.text)] if cell.attrib.get("t") == "s" else node.text or ""
        with archive.open("xl/" + target) as stream:
            for event, row in ET.iterparse(stream, events=("end",)):
                if row.tag != "{" + NS["s"] + "}row":
                    continue
                cells = {re.sub(r"\d+$", "", c.attrib["r"]): value(c) for c in row}
                if row.attrib["r"] == "1":
                    headers = {text: address for address, text in cells.items()}
                    for column in columns:
                        check(column.attrib["header"] in headers, f"Source header missing: {column.attrib['header']}")
                elif cells.get(headers[full_name_header], "").strip():
                    people += 1
                    key = cells.get(headers[key_header], "").strip()
                    missing_keys += not bool(key)
                    keys.append(key)
                row.clear()
    check(missing_keys == 0, "Personnel rows have empty keys")
    check(all(count == 1 for count in Counter(keys).values()), "Duplicate personnel keys")
    for search in profile.findall("search/field"):
        check(search.attrib["name"] in names, "Unknown search field")
        check(search.attrib["match"] in {"contains", "startsWith"}, "Unsupported search match")
    mappings = profile.findall("apply/map")
    targets = [m.attrib["target"] for m in mappings]
    check(len(targets) == len(set(targets)), "Duplicate apply target")
    for mapping in mappings:
        check(mapping.attrib["field"] in names, "Unknown mapped field")
        check(mapping.attrib["policy"] in {"overwrite", "fillIfEmpty", "ignoreEmpty"}, "Unknown apply policy")
    config_lines = (ROOT / "config/PersonalEventBuilder/wsConfig.txt").read_text(encoding="utf-8").splitlines()
    config = {}
    for line in config_lines:
        parts = re.split(r"\t|\\t", line)
        if len(parts) == 3 and parts[0] == "value":
            check(parts[1] not in config, "Duplicate configuration text key")
            config[parts[1]] = parts[2]
    for column in columns:
        if "captionKey" in column.attrib:
            check(bool(config.get(column.attrib["captionKey"])), "Caption configuration missing")
    for path in (ROOT / "vba/PersonalEventBuilder").glob("*.vba"):
        text = path.read_text(encoding="utf-8")
        check(not re.search(r'"[^"\n]*[А-Яа-яІіЇїЄє][^"\n]*"', text), "Localized message hardcoded in VBA")
    for name in ("PersonnelCount", "PersonnelMore"):
        check("{count}" in config["PersonalEventBuilder::text." + name], "Count placeholder missing")
    check("{minChars}" in config["PersonalEventBuilder::text.PersonnelMinimum"], "Minimum-length placeholder missing")
    return profile, people


def validate_page(profile, modules):
    page = ET.parse(ROOT / "ui/PersonalEventBuilder/MainPage.xaml").getroot()
    names = []
    bindings = []
    for node in page.iter():
        name = node.attrib.get("name")
        if name and node.tag.startswith("{urn:excelprototype:controls}"):
            names.append(name)
            check(len(name) <= 25, f"Control name exceeds field/Shape limits: {name}")
            if node.tag.endswith("button"):
                check(len("btn_" + name) <= 31, f"Shape name too long: {name}")
        if node.tag == "{urn:excelprototype:profiles}field":
            name = node.attrib["name"]
            check(len(name) <= 25 and len("sel_" + name + "_input") <= 31, f"Generated name too long: {name}")
            bindings.append(node.attrib["value"])
    check(len(names) == len(set(names)), "Duplicate named controls")
    for mapping in profile.findall("apply/map"):
        target = mapping.attrib["target"].removeprefix("Form.")
        check("{Binding Path=" + target + "}" in bindings, f"Mapped form field missing: {target}")
    candidates = next(n for n in page.iter() if n.attrib.get("name") == "Candidates")
    for member in ("itemsSource", "selectedItem", "onSelect", "selectedStyle"):
        check(member in candidates.attrib, f"Selection attribute missing: {member}")
        check(f'schema.AddAttribute "{member}"' in modules["obj_UiControlFactory"][1], f"Schema attribute missing: {member}")
    page_source = (ROOT / "vba/PersonalEventBuilder/obj_PEB_PgMain.cls.vba").read_text(encoding="utf-8")
    all_sources = "\n".join(t for _, t in modules.values())
    for node in page.iter():
        for attribute in ("command", "onChange", "onSelect"):
            raw = node.attrib.get(attribute, "")
            match = re.search(r"Path=Commands\.(\w+)", raw)
            if match:
                check(f'SetObject("Commands", "{match[1]}"' in page_source, f"Command not registered: {match[1]}")
    for handler in re.findall(r'lookupCommand.Initialize\(m_controller, "(\w+)"\)', page_source):
        check(re.search(rf"Public Function {handler}\b", all_sources), f"Handler missing: {handler}")


if __name__ == "__main__":
    modules = validate_sources()
    profile, count = validate_profile()
    validate_page(profile, modules)
    print(f"PASS: source declarations, interface members, helper visibility, markup, profile, mapped fields, names and EOF bytes; personnel sheet has {count} rows with unique nonempty keys.")
    print("VBA compilation and execution were not performed.")