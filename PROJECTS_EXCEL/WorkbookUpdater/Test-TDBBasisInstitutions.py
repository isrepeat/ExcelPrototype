"""Validate institution lookup configuration, source schemas and form bindings without VBA."""

from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile


ROOT = Path(__file__).resolve().parents[1]
PREFIX = "TravelDocumentsBuilder::"
config = {}
for line in (ROOT / "config/TravelDocumentsBuilder/wsConfig.txt").read_text(encoding="utf-8-sig").splitlines():
    if not line.strip():
        continue
    fields = line.split("\t")
    assert len(fields) == 3, f"Invalid configuration row: {line}"
    assert fields[1] not in config, f"Duplicate key: {fields[1]}"
    config[fields[1]] = fields[2]


def cfg(key):
    return config[PREFIX + key]


kinds = dict(entry.split("~") for entry in cfg("basisChain.institutions").split("|"))
assert kinds == {"Unit": "Військова частина", "Hospital": "Госпіталь", "Command": "Орган військового управління", "Court": "Суд"}
assert cfg("basisChain.lookup.Unit.nameTemplate").replace("{Key}", "А2573") == "військова частина А2573"
assert cfg("basisChain.lookup.Unit.genitiveTemplate").replace("{Key}", "А2573") == "військової частини А2573"
assert "{key}" in cfg("basisChain.lookup.text.MissingGenitive")
for source in cfg("snapshots.sources").split("|"):
    selectors = [config.get(PREFIX + source + suffix, "") for suffix in (".table", ".range")]
    assert sum(bool(value) for value in selectors) == 1, f"Specify exactly one table or range: {source}"
loader = (ROOT / "vba/TravelDocumentsBuilder/obj_TDB_DependencyLoader.cls.vba").read_text(encoding="utf-8-sig")
assert 'VBA.CStr(prefix) & ".table"' in loader, "Dependency loader must recognize table sources"
assert "source.TableName = tableName" in loader
assert 'source.RangeAddress = ex_TravelDocumentsBuilder.fn_Config(VBA.CStr(prefix) & ".range")' in loader
for kind in kinds:
    prefix = f"basisChain.lookup.{kind}"
    if kind in {"Command", "Court"}:
        assert cfg(prefix + ".columns.Key") == cfg(prefix + ".columns.Name") == "Назва"
    assert prefix + ".source" in cfg("snapshots.sources").split("|")
    path = (ROOT / "workbook/TravelDocumentsBuilder" / cfg(prefix + ".source.path")).resolve()
    with zipfile.ZipFile(path) as book:
        workbook = ET.fromstring(book.read("xl/workbook.xml"))
        assert any(node.attrib.get("name") == cfg(prefix + ".source.sheet") for node in workbook.iter())
        tables = [ET.fromstring(book.read(name)) for name in book.namelist() if name.startswith("xl/tables/") and name.endswith(".xml")]
        table = next(table for table in tables if table.attrib.get("name") == cfg(prefix + ".source.table"))
        headers = {node.attrib["name"] for node in table.iter() if node.tag.endswith("}tableColumn")}
        for field in cfg("basisChain.lookup.sourceFields").split("|"):
            assert cfg(prefix + ".columns." + field) in headers
    assert cfg(prefix + ".searchField") in cfg("basisChain.lookup.sourceFields").split("|")

fields = cfg("basisChain.fields").split("|")
assert "BasisInstitutionName" in fields
assert set(fields) <= set(cfg("history.columns").split("|"))
assert set(fields) <= set(cfg("history.snapshotFields").split("|"))
assert {"Snapshot." + field for field in fields} <= set(cfg("history.versions.columns").split("|"))
assert "{BasisInstitutionName}" in cfg("basisChain.documentTemplate")
assert "{BasisInstitution}" not in cfg("basisChain.documentTemplate")
assert "PlaceKind" in cfg("history.columns").split("|")
assert "PlaceKind" in cfg("history.snapshotFields").split("|")
assert "Snapshot.PlaceKind" in cfg("history.versions.columns").split("|")
assert cfg("lookup.Militaries.addressTemplate") == "{Unit}"
assert len(cfg("basisChain.lookup.fields").split("|")) == len(cfg("basisChain.lookup.columnSpans").split("|"))

page = ET.parse(ROOT / "ui/TravelDocumentsBuilder/MainPage.xaml")
nodes = list(page.getroot().iter())
by_name = {node.attrib["name"]: node for node in nodes if "name" in node.attrib}
assert nodes.index(by_name["BasisInstitution"]) < nodes.index(by_name["BasisInstitutionQuery"]) < nodes.index(by_name["BasisType"])
assert by_name["BasisInstitutionQuery"].attrib["value"] == "{Binding Path=Form.BasisInstitutionQuery}"
assert by_name["BasisInstitutionQuery"].attrib["onChange"] == "{Binding Path=Commands.BasisInstitutionQueryChanged}"
for name in ("BasisInstitution", "BasisType", "BasisIssuer"):
    assert by_name[name].attrib.get("editable", "false") == "false"
    assert "placeholder" in by_name[name].attrib
for name in ("BasisList", "TvoList", "PlaceList"):
    assert by_name[name].attrib["readOnly"] == "true"
    assert not any(node.tag.endswith("}cellTemplate") for node in by_name[name].iter())

delete_button = by_name["DeleteEvent"]
assert delete_button.attrib["visibility"] == "{Binding Path=Form.DeleteEventVisibility}"
assert delete_button.attrib["command"] == "{Binding Path=Commands.DeleteEvent}"
assert (delete_button.attrib["row"], delete_button.attrib["column"]) == ("17", "13")
assert "{id}" in cfg("history.text.DeleteConfirm")
history = (ROOT / "vba/TravelDocumentsBuilder/ex_TDB_History.vba").read_text(encoding="utf-8-sig")
delete = history.split("Public Sub fn_DeleteEvent(", 1)[1].split("End Sub", 1)[0]
assert "fn_Begin(tickets, revisions, identifier, expectedStamp)" in delete
assert "fn_AssertCurrent tickets, revisions, transaction" in delete
assert 'Set group = transaction("Previous")' in delete
assert 'identifiers.Add VBA.CStr(row("RowId")), True' in delete
assert 'revisions.ListColumns("EventRowId").Index' in delete
assert "For rowIndex = revisions.ListRows.Count To 1 Step -1" in delete
assert "For rowIndex = tickets.ListRows.Count To 1 Step -1" in delete
assert "fn_Rollback tickets, revisions, transaction" in delete
assert "ex_TableFilters.fn_Restore(filterStates)" in delete
assert "Application.EnableEvents = previousEvents" in delete
assert "Application.ScreenUpdating = previousScreen" in delete
assert ".Save" not in delete and "Kill " not in delete

print("Institution categories, source schemas, snapshot registration, persistence and UI order: OK")