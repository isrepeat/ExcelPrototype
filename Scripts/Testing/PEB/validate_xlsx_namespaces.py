"""Check XLSX compatibility prefixes; optionally restore missing declarations from Git HEAD."""

import argparse
import hashlib
import io
import os
from pathlib import Path
import re
import subprocess
import tempfile
from xml.etree import ElementTree as ET
import zipfile


ROOT = Path(__file__).resolve().parents[3]
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def missing_prefixes(data):
    scopes = []
    pending = {}
    missing = set()
    for event, node in ET.iterparse(io.BytesIO(data), events=("start-ns", "start", "end")):
        if event == "start-ns":
            prefix, uri = node
            pending[prefix] = uri
        elif event == "start":
            scope = dict(scopes[-1]) if scopes else {}
            scope.update(pending)
            pending.clear()
            scopes.append(scope)
            for attribute in ("Ignorable", "MustUnderstand"):
                missing.update(set(node.get(f"{{{MC}}}{attribute}", "").split()) - scope.keys())
            for attribute in ("ProcessContent", "PreserveElements", "PreserveAttributes"):
                for name in node.get(f"{{{MC}}}{attribute}", "").split():
                    if ":" in name and name.split(":", 1)[0] not in scope:
                        missing.add(name.split(":", 1)[0])
        else:
            scopes.pop()
    return missing


def audit(payloads):
    result = {}
    for name, data in payloads.items():
        if name.endswith((".xml", ".rels")):
            missing = missing_prefixes(data)
            if missing:
                result[name] = missing
    return result


def repair(path, original, payloads, problems):
    relative = path.relative_to(ROOT).as_posix()
    baseline = subprocess.check_output(["git", "show", "HEAD:" + relative], cwd=ROOT)
    replacements = {}
    with zipfile.ZipFile(io.BytesIO(baseline)) as reference:
        for name, prefixes in problems.items():
            declarations = dict(node for _, node in ET.iterparse(io.BytesIO(reference.read(name)), events=("start-ns",)))
            additions = b""
            for prefix in sorted(prefixes):
                if prefix not in declarations:
                    raise ValueError(f"No reference namespace for {name}: {prefix}")
                uri = declarations[prefix]
                if any(char in uri for char in ('"', '&', '<')):
                    raise ValueError("Reference URI requires escaping")
                additions += f' xmlns:{prefix}="{uri}"'.encode("utf-8")
            data = payloads[name]
            root = re.search(rb"<[A-Za-z_][^>]*>", data)
            if root is None or root.group().endswith(b"/>"):
                raise ValueError(f"Unsupported XML root: {name}")
            offset = root.end() - 1
            replacements[name] = data[:offset] + additions + data[offset:]
            assert replacements[name][:offset] + replacements[name][offset + len(additions):] == data
    expected = payloads | replacements
    if audit(expected):
        raise ValueError("Namespace repair did not resolve all errors")
    backup = Path(tempfile.gettempdir()) / (path.stem + ".before-namespaces-" + hashlib.sha256(original).hexdigest()[:12] + ".xlsx")
    if backup.exists():
        assert backup.read_bytes() == original
    else:
        backup.write_bytes(original)
    descriptor, temporary = tempfile.mkstemp(prefix=path.stem + ".repair-", suffix=".xlsx", dir=path.parent)
    os.close(descriptor)
    try:
        with zipfile.ZipFile(io.BytesIO(original)) as source, zipfile.ZipFile(temporary, "w") as target:
            target.comment = source.comment
            for entry in source.infolist():
                target.writestr(entry, expected[entry.filename])
        with zipfile.ZipFile(temporary) as checked:
            assert checked.testzip() is None
            assert checked.namelist() == list(payloads)
            assert all(checked.read(name) == data for name, data in expected.items())
        if path.read_bytes() != original:
            raise ValueError("Workbook changed during repair")
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)
    print(f"REPAIRED: {relative}: {len(replacements)} XML parts; backup: {backup}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repair", action="store_true")
    parser.add_argument("paths", nargs="*", type=Path)
    args = parser.parse_args()
    paths = args.paths or sorted((ROOT / "PROJECTS_EXCEL/tests/PersonalEventBuilder").rglob("*.xlsx"))
    failures = 0
    for path in paths:
        path = path.resolve()
        original = path.read_bytes()
        with zipfile.ZipFile(io.BytesIO(original)) as workbook:
            if workbook.testzip() is not None:
                raise ValueError(f"Corrupt ZIP archive: {path}")
            payloads = {name: workbook.read(name) for name in workbook.namelist()}
        problems = audit(payloads)
        for name, prefixes in problems.items():
            print(f"MISSING: {path.name}/{name}: {', '.join(sorted(prefixes))}")
        if problems and args.repair:
            repair(path, original, payloads, problems)
        elif problems:
            failures += 1
    if failures:
        raise SystemExit(f"FAIL: {failures} workbooks have unresolved compatibility prefixes")
    print(f"PASS: XML compatibility namespace checks for {len(paths)} workbooks.")


if __name__ == "__main__":
    main()