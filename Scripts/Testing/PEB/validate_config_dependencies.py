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
        previous = check_output(["git", "show", f"HEAD:PROJECTS_EXCEL/config/PersonalEventBuilder/{filename}"], cwd=ROOT).decode("utf-8-sig")
        previous = {line.split("\t")[1]: line.split("\t")[2] for line in previous.splitlines() if line}
        for kind, key, value in rows(path):
            if kind != "ref":
                continue
            assert value in resolved, (key, value)
            if key in previous:
                expected = previous[key]
                if key.endswith(".path"):
                    expected = str((WORKBOOK / expected.replace('\\', '/')).resolve())
                assert resolved[key] == expected, (filename, key, resolved[key], expected)
            count += 1
        if filename.endswith("test.txt"):
            for key in ("filePersonnelShpo", "fileEjos", "fileMovementHistory", "fileMovementExport"):
                assert "tests/PersonalEventBuilder/Dependencies/test_" in Path(resolved[key]).as_posix(), key
            assert resolved["fileMovementHistory"] != load_config(path.with_name("wsConfig.txt"))["fileMovementHistory"]

    imports = [["importCfg", "Shared", "../cfgDependencies.xlsx"]]
    loader = lambda _: [["file", "Data", "Dependencies/source.xlsx"], ["value", "Range", "A1:C20"]]
    for overrides in (
        [["file", "Data", "test/source.xlsx"]] + imports,
        imports + [["file", "Data", "test/source.xlsx"]],
    ):
        result = expand(overrides + [["ref", "Consumer", "Data"]], WORKBOOK, loader)
        assert result["Consumer"] == str((WORKBOOK / "test/source.xlsx").resolve())
    result = expand(imports + [["ref", "Consumer", "Data"]], WORKBOOK, loader)
    assert result["Consumer"] == str((shared.parent / "Dependencies/source.xlsx").resolve())
    reject([["ref", "A", "missing"]], loader, "not found")
    reject([["ref", "A", "B"], ["ref", "B", "A"]], loader, "cycle")
    reject(imports + [["ref", "A", "Shared"]], loader, "import directive")
    reject([["value", "A", ""], ["ref", "B", "A"]], loader, "empty referenced")
    reject(imports, lambda _: [["value", "A", "1"], ["value", "a", "2"]], "duplicate shared")
    reject([["value", "A", "1"], ["value", "a", "2"]], loader, "duplicate local")
    reject(imports, lambda _: [["importCfg", "Other", "nested.xlsx"]], "nested")
    reject([["importCfg", "Shared", "missing.xlsx"]], shared_rows, "missing.xlsx")

    runtime = (ROOT / "PROJECTS_EXCEL/vba/common/modules/ex_ConfigImports.vba").read_text(encoding="utf-8-sig")
    refresh = (ROOT / "PROJECTS_EXCEL/WorkbookUpdater/ex_ConfigRefresh.vba").read_text(encoding="utf-8-sig")
    assert "service.Backend = QueryAuto" in runtime and 'source.TableName = "tbConfig"' in runtime
    assert refresh.index("ex_ConfigImports.fn_TryExpand") < refresh.index("changed = True")
    assert "targetWorkbook.Path, expandedRecords" in refresh
    print(f"PASS: {len(actual)} shared keys, {count} resolved PEB references, migration parity, native table styles, TEST isolation, overrides, missing files/keys, cycles and duplicate rejection. VBA not executed.")


if __name__ == "__main__":
    main()