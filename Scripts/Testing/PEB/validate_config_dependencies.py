"""Check shared imports, overrides and migration parity without running VBA."""

from pathlib import Path
from subprocess import check_output
from zipfile import ZipFile
from xml.etree import ElementTree as ET

from config_dependencies import ROOT, WORKBOOK, expand, load_config, rows, shared_rows


def reject(local, loader, expected):
    try:
        expand(local, WORKBOOK, loader)
    except (ValueError, FileNotFoundError) as error:
        assert expected.lower() in str(error).lower(), error
    else:
        raise AssertionError(f"Expected rejection: {expected}")


def main():
    shared = ROOT / "PROJECTS_EXCEL/workbook/cfgDependencies.xlsx"
    actual = shared_rows(shared)
    source = rows(ROOT / "PROJECTS_EXCEL/config/cfgDependencies/wsConfig.txt")
    assert actual == source, "Shared workbook differs from its editable source"
    assert len({row[1].casefold() for row in actual}) == len(actual)
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with ZipFile(shared) as package:
        styles = ET.fromstring(package.read("xl/styles.xml"))
        assert styles.find("m:tableStyles/m:tableStyle[@name='TableStyleDark_my']", ns) is not None

    count = 0
    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        path = ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder" / filename
        resolved = load_config(path)
        if filename == "wsConfig.txt":
            assert not any(key.startswith("PersonalEventBuilder::test.") for key in resolved)
            assert "PersonalEventBuilder::text.TestScenarioItems" not in resolved
        previous = check_output(["git", "show", f"HEAD:PROJECTS_EXCEL/config/PersonalEventBuilder/{filename}"], cwd=ROOT).decode("utf-8-sig")
        previous_rows = [line.split("\t") for line in previous.splitlines() if line]
        shared_keys = {row[1] for row in actual}
        for row in previous_rows:
            if row[0].casefold() == "importcfg":
                row[1] = "Import::workbooks.__cfgDependencies"
            elif row[0].casefold() == "ref" and row[2] in shared_keys:
                row[2] = "__cfgDependencies." + row[2]
            elif row[0] == "file" and row[1] in shared_keys:
                row[1] = "__cfgDependencies." + row[1]
        previous = expand(previous_rows, WORKBOOK, shared_rows)
        assert resolved["ThisWorkbook::id"] == "PersonalEventBuilder"
        for kind, key, value in rows(path):
            if kind != "ref":
                continue
            assert value in resolved, (key, value)
            if key in previous:
                expected = previous[key]
                assert resolved[key] == expected, (filename, key, resolved[key], expected)
            count += 1
        if filename.endswith("test.txt"):
            for key in ("filePersonnelShpo", "fileEjos", "fileMovementHistory", "fileMovementExport"):
                assert "tests/PersonalEventBuilder/Dependencies/test_" in Path(resolved["__cfgDependencies." + key]).as_posix(), key
            assert resolved["__cfgDependencies.fileMovementHistory"] != load_config(path.with_name("wsConfig.txt"))["__cfgDependencies.fileMovementHistory"]

    imports = [["importCfg", "Import::workbooks.__Shared", "../cfgDependencies.xlsx"]]
    loader = lambda _: [["file", "Data", "Dependencies/source.xlsx"], ["value", "Range", "A1:C20"]]
    for overrides in (
        [["file", "__Shared.Data", "test/source.xlsx"]] + imports,
        imports + [["file", "__Shared.Data", "test/source.xlsx"]],
    ):
        result = expand(overrides + [["ref", "Consumer", "__Shared.Data"]], WORKBOOK, loader)
        assert result["Consumer"] == str((WORKBOOK / "test/source.xlsx").resolve())
    result = expand(imports + [["ref", "Consumer", "__Shared.Data"]], WORKBOOK, loader)
    assert result["Consumer"] == str((shared.parent / "Dependencies/source.xlsx").resolve())
    padded = expand(imports + [["  ref", "Consumer", "__Shared.Data"]], WORKBOOK, loader)
    assert padded["Consumer"] == result["Consumer"]
    reject([["ref", "A", "missing"]], loader, "not found")
    reject([["ref", "A", "B"], ["ref", "B", "A"]], loader, "cycle")
    reject(imports + [["ref", "A", "Import::workbooks.__Shared"]], loader, "import directive")
    reject([["value", "A", ""], ["ref", "B", "A"]], loader, "empty referenced")
    reject(imports, lambda _: [["value", "A", "1"], ["value", "a", "2"]], "duplicate shared")
    reject([["value", "A", "1"], ["value", "a", "2"]], loader, "duplicate local")
    reject(imports, lambda _: [["importCfg", "Other", "nested.xlsx"]], "nested")
    reject([["importCfg", "Import::workbooks.__Shared", "missing.xlsx"]], shared_rows, "missing.xlsx")
    reject([["importCfg", "Shared", "data.xlsx"]], loader, "import variable")
    reject(imports + [["ref", "A", "Data"]], loader, "not found")
    other = [["importCfg", "Import::workbooks.__Other", "../Other/cfg.xlsx"]]
    multi_loader = lambda _: loader(_) + [["ref", "Alias", "Data"]]
    result = expand(imports + other, WORKBOOK, multi_loader)
    assert result["__Shared.Range"] == result["__Other.Range"] == "A1:C20"
    assert result["__Shared.Alias"] == result["__Shared.Data"]
    assert result["__Other.Alias"] == result["__Other.Data"]
    assert result["__Shared.Data"] != result["__Other.Data"]
    reject(imports + other + [["importCfg", "Import::workbooks.__Shared", "third.xlsx"]], loader, "duplicate")

    runtime = (ROOT / "PROJECTS_EXCEL/vba/common/modules/ex_ConfigImports.vba").read_text(encoding="utf-8-sig")
    refresh = (ROOT / "PROJECTS_EXCEL/WorkbookUpdater/ex_ConfigRefresh.vba").read_text(encoding="utf-8-sig")
    assert "service.Backend = QueryAuto" in runtime and 'source.TableName = "tbConfig"' in runtime
    assert refresh.index("ex_ConfigImports.fn_TryExpand") < refresh.index("changed = True")
    assert "targetWorkbook.Path, expandedRecords" in refresh
    print(f"PASS: {len(actual)} shared keys, {count} resolved PEB references, migration parity, native table styles, TEST isolation, overrides, missing files/keys, cycles and duplicate rejection. VBA not executed.")


if __name__ == "__main__":
    main()