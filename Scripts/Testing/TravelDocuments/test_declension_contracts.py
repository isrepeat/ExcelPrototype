"""Check declension source, snapshot and Word contracts without Excel or VBA."""

from pathlib import Path
import re


ROOT = Path(__file__).resolve().parents[3] / "PROJECTS_EXCEL"
config = {}
for line in (ROOT / "config/TravelDocumentsBuilder/wsConfig.txt").read_text(encoding="utf-8-sig").splitlines():
    kind, key, value = line.split("\t")
    assert key not in config, key
    config[key] = (kind.strip(), value)


def cfg(key):
    return config["TravelDocumentsBuilder::" + key][1]


keys = list(config)
assert keys == sorted(keys, key=lambda key: tuple((len(part), part) for part in key.replace("::", ".").split(".")))
import_kind, import_path = config["Import::workbooks.__cfgDependencies"]
assert import_kind == "importCfg", "External references require an explicit workbook import"
assert (ROOT / "workbook/TravelDocumentsBuilder" / import_path).resolve().is_file()
fields = set()
dependencies = dict(
    (parts[1], parts[2])
    for line in (ROOT / "config/cfgDependencies/wsConfig.txt").read_text(encoding="utf-8-sig").splitlines()
    if len(parts := line.split("\t")) == 3
)
for index in range(1, int(cfg("declensions.count")) + 1):
    prefix = f"declensions.lookups.{index}."
    field = cfg(prefix + "output")
    assert field not in fields
    fields.add(field)
    source = "declensions.sources." + cfg(prefix + "source")
    assert source in cfg("snapshots.sources").split("|")
    for suffix in ("path", "sheet", "range"):
        kind, reference = config["TravelDocumentsBuilder::" + source + "." + suffix]
        assert kind == "ref"
        assert reference.removeprefix("__cfgDependencies.") in dependencies, reference
    assert cfg(prefix + "keyHeader") and cfg(prefix + "valueHeader")

assert fields == {"RankGenitive", "RankDative", "FIOAccusative", "FIODative", "FIOInitialsGenitive", "PositionGenitive"}
assert fields <= set(cfg("history.columns").split("|"))
assert fields <= set(cfg("history.snapshotFields").split("|"))
assert {"Snapshot." + field for field in fields} <= set(cfg("history.versions.columns").split("|"))
generator = (ROOT / "vba/TravelDocumentsBuilder/obj_TDB_VacationGenerator.cls.vba").read_text(encoding="utf-8-sig")
saved = set(re.findall(r'metadata.Add "(\w+)", m_form\(', generator))
assert fields <= saved
export = generator[generator.index("    m_operationId = \"WORD/"):generator.index("Private Sub private_ReadForm")]
assert "fn_Enrich" not in export, "Export must use the saved snapshot rather than current dictionaries"
assert 'm_form("RankDative") & " " & m_form("FIODative")' in generator
assert "root(key)" in export
print("Declension configuration, snapshots and Word contracts: OK")