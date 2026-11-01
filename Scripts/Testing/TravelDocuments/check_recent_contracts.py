"""Audit changed VBA procedures and module/class contracts without executing Office or VBA."""

from argparse import ArgumentParser
import difflib
from pathlib import Path
import re
import subprocess

from lxml import etree as ET


ROOT = Path(__file__).resolve().parents[3]
PROCEDURE = re.compile(r"^(?:Public|Private|Friend) (?:Sub|Function|Property (?:Get|Let|Set)) (\w+).*?^End (?:Sub|Function|Property)", re.M | re.S)
SECTION = re.compile(r"^' //\n' // (\w+)\n' //\n", re.M)


def git(*arguments):
    return subprocess.check_output(["git", "-c", "core.quotepath=false", *arguments], cwd=ROOT, encoding="utf-8", stderr=subprocess.PIPE)


def audit(base):
    names = set(git("diff", base, "--name-only").splitlines())
    names.update(git("ls-files", "--others", "--exclude-standard").splitlines())
    errors = []
    checked_procedures = 0
    checked_vba = 0
    config = {}
    for line in (ROOT / "PROJECTS_EXCEL/config/TravelDocumentsBuilder/wsConfig.txt").read_text(encoding="utf-8-sig").splitlines():
        if line.strip():
            kind, key, value = line.split("\t")
            if key in config:
                errors.append(f"Duplicate configuration key: {key}")
            config[key] = value

    for name in sorted(names):
        path = ROOT / name
        if not path.is_file():
            continue
        if path.suffix.lower() in (".vba", ".py", ".xml", ".xaml", ".md", ".txt") or path.name == ".gitignore":
            data = path.read_bytes()
            if not data or data[-1] in (9, 10, 13, 32):
                errors.append(f"Invalid final byte: {name}")
            text = data.decode("utf-8-sig")
            if re.search(r"[ \t]+$", text, re.M):
                errors.append(f"Trailing whitespace: {name}")
        if path.suffix.lower() in (".xml", ".xaml"):
            priority = "name type dataContext itemsSource row column rowSpan columnSpan orientation style value label caption command".split()
            for node in ET.parse(str(path)).iter():
                if not isinstance(node.tag, str):
                    continue
                attributes = list(node.attrib)
                expected = sorted(attributes, key=lambda key: priority.index(key) if key in priority else len(priority))
                if attributes != expected:
                    errors.append(f"XML attribute order: {name}:{node.sourceline}")
        if path.suffix.lower() != ".vba":
            continue
        checked_vba += 1
        source = path.read_text(encoding="utf-8-sig")
        source_lines = source.splitlines()
        for number, line in enumerate(source_lines[:-1], 1):
            if line.rstrip().endswith(" _") and re.match(r"\s*(?:End\b|Else\b|ElseIf\b|Next\b|Loop\b)", source_lines[number]):
                errors.append(f"Control-flow boundary inside a continued statement: {name}:{number}")
        try:
            original = git("show", base + ":" + name)
        except subprocess.CalledProcessError:
            original = ""
        changed = set()
        for kind, _, _, first, last in difflib.SequenceMatcher(a=original.splitlines(), b=source.splitlines()).get_opcodes():
            if kind != "equal":
                changed.update(range(first + 1, last + 1))
        for number, line in enumerate(source.splitlines(), 1):
            if number in changed and line.lstrip().startswith("'") and re.search(r"[А-Яа-яЁёІіЇїЄє]", line):
                errors.append(f"Non-English changed code comment: {name}:{number}")
        procedures = list(PROCEDURE.finditer(source))
        ordinary_names = [match[1].lower() for match in procedures if not match[0].startswith(("Public Property", "Private Property", "Friend Property"))]
        if len(ordinary_names) != len(set(ordinary_names)):
            errors.append(f"Duplicate VBA procedure names: {name}")

        if name.endswith(".cls.vba"):
            fields = re.findall(r"^Private (\w+) As", source, re.M)
            if fields[:2] != ["m_isInitialized", "m_isDisposed"]:
                errors.append(f"Lifecycle flags are not the first two fields: {name}")
            sections = list(SECTION.finditer(source))
            lifecycle = [section for section in sections if section[1] == "Lifecycle"]
            if len(lifecycle) != 1:
                errors.append(f"Expected one Lifecycle section: {name}")
            else:
                start = lifecycle[0].end()
                end = next((section.start() for section in sections if section.start() > start), len(source))
                callbacks = re.findall(r"^Private (?:Sub|Function) (\w+)", source[start:end], re.M)
                if callbacks != ["Class_Initialize", "Class_Terminate"]:
                    errors.append(f"Invalid Lifecycle callbacks: {name}")
            api = next((section for section in sections if section[1] == "API"), None)
            public_api = re.findall(r"^Public (?:Sub|Function) (\w+)", source[api.end():] if api else "", re.M)
            if public_api[:2] != ["Initialize", "Dispose"]:
                errors.append(f"Initialize/Dispose must lead the API section: {name}")
            if "namespace " in source:
                errors.append(f"Namespace comments in a class: {name}")
        else:
            if not re.search(r"^Public Sub fn_Module_Dispose\(\)", source, re.M):
                errors.append(f"Missing fn_Module_Dispose: {name}")
            scope = None
            for line in source.splitlines():
                opening = re.fullmatch(r"' namespace (\w+) \{", line)
                closing = re.fullmatch(r"' } // namespace (\w+)", line)
                if opening:
                    if scope is not None:
                        errors.append(f"Nested namespace block: {name}")
                    scope = opening[1]
                elif closing:
                    if scope != closing[1]:
                        errors.append(f"Mismatched namespace block: {name}")
                    scope = None
                elif re.match(r"Public (?:Sub|Function) ", line) and scope not in ("Lifecycle", "API"):
                    errors.append(f"Public procedure outside its namespace: {name}: {line}")
                elif re.match(r"Private (?:Sub|Function) ", line) and scope != "Private":
                    errors.append(f"Private procedure outside its namespace: {name}: {line}")
            if scope is not None:
                errors.append(f"Unclosed namespace block: {name}")

        for match in procedures:
            first = source[:match.start()].count("\n") + 1
            last = first + match[0].count("\n")
            if not any(first <= line <= last for line in changed):
                continue
            checked_procedures += 1
            lines = match[0].splitlines()
            signature_end = 0
            while lines[signature_end].endswith(" _"):
                signature_end += 1
            if signature_end == 0 and "," in lines[0]:
                errors.append(f"Parameters must occupy separate lines: {name}: {match[1]}")
            executable_seen = False
            for line in lines[signature_end + 1:-1]:
                stripped = line.strip()
                if not stripped or stripped.startswith(("'", "#")):
                    continue
                if stripped.startswith("Dim "):
                    if executable_seen:
                        errors.append(f"Declaration after executable code: {name}: {match[1]}")
                else:
                    executable_seen = True
                if re.match(r"If .* Then (?!_$).+", stripped):
                    errors.append(f"Single-line If: {name}: {match[1]}")
                if re.match(r"Case .*: .+", stripped):
                    errors.append(f"Compressed Select Case: {name}: {match[1]}")

        if path.name.startswith(("obj_WordTable", "ex_WordDslResources", "ex_WordTableXml")):
            if re.search(r"\b(?:MsgBox|fn_ShowMessage|fn_StatusBar_Show)\b", source):
                errors.append(f"Technical DSL errors must be handled at the export boundary: {name}")
            if "ex_WordDslDiagnostics" in source:
                errors.append(f"Separate DSL diagnostic mechanism: {name}")

    if errors:
        raise AssertionError("\n".join(errors))
    print(f"PASS: {checked_vba} VBA files; {checked_procedures} changed procedures; lifecycle, namespaces, formatting, diagnostics and file contracts.")
    print("Excel and VBA were not executed.")


if __name__ == "__main__":
    parser = ArgumentParser(description=__doc__)
    parser.add_argument("--base", default="ee2a848^")
    audit(parser.parse_args().base)