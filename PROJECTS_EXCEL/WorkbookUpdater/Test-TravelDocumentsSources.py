"""Validate form, configuration, registry and DOCX contracts without running Excel or VBA."""

from pathlib import Path
import re
import xml.etree.ElementTree as ET
import zipfile


ROOT = Path(__file__).resolve().parents[1]
PREFIX = "TravelDocumentsBuilder::"
config = {}
for line in (ROOT / "config/TravelDocumentsBuilder/wsConfig.txt").read_text(encoding="utf-8-sig").splitlines():
    if not line.strip():
        continue
    fields = line.split("\t")
    assert len(fields) == 3, f"Invalid config row: {line}"
    assert fields[1] not in config, f"Duplicate config key: {fields[1]}"
    config[fields[1]] = fields[2].strip()


def cfg(key):
    assert PREFIX + key in config, f"Missing config: {key}"
    return config[PREFIX + key]


sources = ROOT / "vba/TravelDocumentsBuilder"
page = (sources / "obj_TDB_PgMain.cls.vba").read_text(encoding="utf-8-sig")
controller = (sources / "obj_TDB_PgMainController.cls.vba").read_text(encoding="utf-8-sig")
generator = (sources / "obj_TDB_VacationGenerator.cls.vba").read_text(encoding="utf-8-sig")
all_sources = page + controller + generator

# Check required keys and option lists against configuration.
key_assignment = re.search(r"    keys = (.*?)\n    For Each key", generator, re.S)[1]
required = "".join(re.findall(r'"([^"]*)"', key_assignment)).split(",")
for key in required:
    cfg("generation." + key)
for kind in cfg("generation.kind.ids").split(","):
    for part in ("input", "word", "event"):
        cfg(f"generation.kind.{kind}.{part}")
kinds = [cfg(f"generation.kind.{kind}.input") for kind in cfg("generation.kind.ids").split(",")]
assert cfg("generation.items.VacationKinds").split("|") == kinds
assert cfg("generation.defaults.VacationKind") in kinds
assert len(cfg("generation.months").split("|")) == 12
assert cfg("generation.defaults.VacationAbroad") in cfg("generation.items.AbroadOptions").split(",")
assert cfg("generation.defaults.VacationStatus") in cfg("generation.items.StatusOptions").split(",")

# Check command routes, unique names and initial form values.
markup = ET.parse(ROOT / "ui/TravelDocumentsBuilder/MainPage.xaml")
registered = dict(re.findall(r'bindingContext, "(\w+)", "(\w+)"', page))
names = []
for node in markup.iter():
    name = node.attrib.get("name")
    if name and node.tag.startswith("{urn:excelprototype:controls}"):
        assert len(name) <= 25, f"Control name too long: {name}"
        names.append(name)
        if node.tag.endswith("button"):
            assert len("btn_" + name) <= 31
    for attribute, raw in node.attrib.items():
        binding = re.search(r"Path=(\w+)\.(\w+)", raw)
        if not binding:
            continue
        category, key = binding.groups()
        if category == "Commands":
            assert key in registered, f"Unregistered command: {key}"
            assert f"Public Function {registered[key]}" in controller, f"Missing handler: {key}"
        elif category == "Form":
            assert f'"{key}"' in page, f"Form binding not initialized: {key}"
        elif category == "Text":
            if key not in ("PersonnelStatus", "CandidateCount"):
                assert f'"{key}"' in all_sources, f"Unknown text binding: {key}"
        elif category == "Data":
            assert f'"{key}"' in page + controller, f"Unknown data binding: {key}"
assert len(names) == len(set(names)), "Duplicate control names"
assert "TvoLookup" not in names
for name in ("TvoName", "TvoIpn", "TvoPosition", "LookupTarget"):
    assert name in names, f"Missing TVO control: {name}"
assert int(cfg("lookup.Tvo.apply.count")) == 3
assert {
    cfg(f"lookup.Tvo.apply.{index}.target") for index in range(1, 4)
} == {"Form.TvoName", "Form.TvoIpn", "Form.TvoPosition"}
for state in ("Personnel", "Tvo"):
    caption = cfg(f"text.LookupTarget.{state}")
    assert caption.count("[[color=#") == caption.count("[[/color]]") == 1
    assert re.search(r"\[\[color=#[0-9A-Fa-f]{6}\]\].+?\[\[/color\]\]", caption)

# Validate contracts against actual files rather than a copied test schema.
workbook_folder = ROOT / "workbook/TravelDocumentsBuilder"
with zipfile.ZipFile(workbook_folder / cfg("generation.registry.workbook")) as archive:
    namespace = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    tables = [
        ET.fromstring(archive.read(name))
        for name in archive.namelist()
        if re.fullmatch(r"xl/tables/table\d+\.xml", name)
    ]
    table = next(table for table in tables if table.attrib["name"] == cfg("generation.registry.table"))
    headers = [column.attrib["name"] for column in table.find("s:tableColumns", namespace)]
    assert len(headers) == len(set(headers)) == 21
    assert set(headers) == set(cfg("generation.registry.columns").split("|")), "Registry schema mismatch"
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    assert any(sheet.attrib["name"] == cfg("generation.registry.sheet")
               for sheet in workbook.find("s:sheets", namespace)), "Registry sheet missing"

template = workbook_folder / cfg("generation.templates.VacationTicket").replace("\\", "/")
with zipfile.ZipFile(template) as archive:
    document = ET.fromstring(archive.read("word/document.xml"))
    text = "".join(document.itertext())
    placeholders = re.search(r"    names = VBA.Array\((.*?)\)\n    values", generator, re.S)[1]
    markers = re.findall(r'"(\w+)"', placeholders)
    assert len(markers) == 12
    for marker in markers:
        assert f"<{marker}></{marker}>" in text, f"Template placeholder missing: {marker}"

# Validate explicit dependency ranges and source headers against saved XLSX files.
def source_headers(prefix):
    path = (workbook_folder / cfg(prefix + ".path").replace("\\", "/")).resolve()
    match = re.fullmatch(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", cfg(prefix + ".range"))
    assert match, f"Explicit range required: {prefix}"
    with zipfile.ZipFile(path) as archive:
        shared = archive.read("xl/sharedStrings.xml") if "xl/sharedStrings.xml" in archive.namelist() else None
        strings = ["".join(item.itertext()) for item in ET.fromstring(shared)] if shared else []
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        sheet = next(item for item in workbook.find("s:sheets", namespace)
                     if item.attrib["name"] == cfg(prefix + ".sheet"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relationship_id = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
        target = next(item.attrib["Target"] for item in relationships if item.attrib["Id"] == relationship_id)
        target = target.lstrip("/") if target.startswith("/") else "xl/" + target
        data = ET.fromstring(archive.read(target))
        def column_number(text):
            result = 0
            for char in text:
                result = result * 26 + ord(char) - ord("A") + 1
            return result
        first, last = column_number(match[1]), column_number(match[3])
        actual = []
        for cell in data.findall(".//s:c", namespace):
            address = re.fullmatch(r"([A-Z]+)(\d+)", cell.attrib["r"])
            if address[2] != match[2] or not first <= column_number(address[1]) <= last:
                continue
            value = cell.find("s:v", namespace)
            if cell.attrib.get("t") == "s" and value is not None:
                actual.append(strings[int(value.text)])
            elif cell.attrib.get("t") == "inlineStr":
                actual.append("".join(cell.find("s:is", namespace).itertext()))
            elif value is not None:
                actual.append(value.text)
        return actual

assert cfg("snapshots.sources").split("|") == ["lookup.Personnel.source", "orders.source"]
personnel_headers = source_headers("lookup.Personnel.source")
for index in range(1, int(cfg("lookup.Personnel.columns.count")) + 1):
    header = cfg(f"lookup.Personnel.columns.{index}.header")
    assert personnel_headers.count(header) == 1, f"Missing or ambiguous personnel header: {header}"
order_headers = source_headers("orders.source")
for header in (cfg("orders.dateHeader"), cfg("orders.numberHeader")):
    assert order_headers.count(header) == 1, f"Missing or ambiguous order header: {header}"

# Guard the migration boundaries: prepared reads, explicit loading and live registry writes.
loader = (sources / "obj_TDB_DependencyLoader.cls.vba").read_text(encoding="utf-8-sig")
entry = (sources / "ex_TravelDocumentsBuilder.vba").read_text(encoding="utf-8-sig")
workbook_source = (sources / "ThisWorkbook.vba").read_text(encoding="utf-8-sig")
assert "Implements obj_ITableSnapshotLoader" in loader
assert "snapshot.TryLoad(source, diagnostic)" in loader
assert "lookup.PrepareSnapshot(snapshot, diagnostic)" in loader
assert "orders.PrepareSnapshot(snapshot, diagnostic)" in loader
assert loader.count("private_FileStamp(") == 3
assert "ex_TableSnapshots.fn_Register dependencyLoader" in controller
assert "ex_TableSnapshots.fn_Refresh(diagnostic)" in controller
assert "ex_TableDataRuntime.fn_ResetCaches" not in controller
assert "ex_TableSnapshots.fn_Module_Dispose" in entry
assert "obj_OrderResolver" in generator and "m_orderResolver.TrySearch" in generator
assert "ADODB.Connection" not in generator and "generation.orders." not in generator
assert "matches.Count = 0 And Not allowMissing" in generator
assert "If matches.Count = 0 Then Exit Sub" in generator
assert "Set ticketsWorkbook = private_OpenWorkbook()" in generator
assert "Set ticketsTable = private_Table(ticketsWorkbook)" in generator
assert "table.ListRows.Add" in generator
assert cfg("candidates.selectionMode") in ("Focus", "Confirm")
candidate_table = next(node for node in markup.iter() if node.attrib.get("name") == "Candidates")
assert candidate_table.attrib["selectionMode"] == "{Binding Path=Settings.CandidateSelectionMode}"
assert "ThisWorkbook::ui.confirmSelectionKey" in config
assert "fn_CancelPending" in workbook_source and "fn_ForgetContext" in workbook_source
assert "Cancel = True" in workbook_source
for key in re.findall(r'fn_Config\("([^"]+)"\)', page + controller + generator + loader + entry + workbook_source):
    cfg(key)
for key in ("OrderSourceError", "OrderInvalidYear", "OrderInvalidInput", "OrderInvalidDate",
            "OrderYearConflict", "OrderDataInvalid", "OrderDuplicate"):
    cfg("text." + key)
assert config["ThisWorkbook::ui.pageMappings"] == "MainPage=MainPage.xaml"
assert (ROOT / "ui" / cfg("ui.folder") / "MainPage.xaml").is_file()
keys = list(config)
assert keys == sorted(keys, key=lambda key: tuple((len(part), part) for part in key.replace("::", ".").split("."))), "Configuration staircase order"
print(f"PASS: TravelDocumentsBuilder config, {len(names)} controls, routes, registry schema, DOCX markers, dependency headers and snapshot contracts.")
print("Excel, Word and VBA were not executed.")