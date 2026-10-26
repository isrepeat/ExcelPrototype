"""Check root/block/export hierarchy, cleanup cases and template migration without VBA."""

from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import re

from add_word_root_blocks import FOLDER, TEMPLATE_NAME, NS, add_roots
from event_routes import ROOT
from validate_movement_batch import check_blocks
from validate_movement_planner import check_declarations, procedure

MARKER = re.compile(r"\{\\(root|block|export):([A-Za-z0-9_]+)_(Begin|End)\}")


def index(text):
    stack, sections, seen = [], [], set()
    levels = {"root": 0, "block": 1, "export": 2}
    for match in MARKER.finditer(text):
        kind, name, side = match.groups()
        if side == "Begin":
            if (kind, name) in seen or stack and levels[kind] <= levels[stack[-1][0]]:
                raise ValueError("Duplicate or nested section")
            if kind == "export" and stack and stack[-1][0] == "block" and stack[-1][1] != name:
                raise ValueError("Block/export identifiers differ")
            seen.add((kind, name))
            stack.append((kind, name, match))
        else:
            if not stack or stack[-1][:2] != (kind, name):
                raise ValueError("Crossing or unpaired section")
            _, _, opening = stack.pop()
            sections.append((kind, name, opening, match))
    if stack:
        raise ValueError("Unpaired opening")
    return sections


def cleanup(text):
    sections = index(text)
    exports = [entry for entry in sections if entry[0] == "export"]
    removals = []
    for kind, name, opening, closing in sections:
        children = [entry for entry in exports if opening.end() <= entry[2].start() < closing.start()]
        populated = any(re.sub(r"[\s\u00a0]+", "", text[first.end():last.start()]) for _, _, first, last in children)
        if kind in ("root", "block") and not populated:
            removals.append((opening.start(), closing.end()))
        else:
            removals.extend(((opening.start(), opening.end()), (closing.start(), closing.end())))
    merged = []
    for start, end in sorted(removals):
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    for start, end in reversed(merged):
        text = text[:start] + text[end:]
    return text


def block(name, content=""):
    return (f"{{\\block:{name}_Begin}}\n{name} heading\n"
            f"{{\\export:{name}_Begin}}{content}{{\\export:{name}_End}}\n{{\\block:{name}_End}}")


def root(name, content):
    return f"{{\\root:{name}_Begin}}\n{name} heading\n{content}\n{{\\root:{name}_End}}"


def main():
    assert cleanup("before" + root("Empty", block("A") + block("B", "\n\u00a0 ")) + "after") == "beforeafter"
    assert cleanup(root("NoChildren", "Static heading only")) == ""
    mixed = cleanup(root("Keep", block("A") + block("B", "Person")))
    assert "Keep heading" in mixed and "B heading" in mixed and "Person" in mixed
    assert "A heading" not in mixed and "{\\" not in mixed
    multiple = cleanup(root("Drop", block("A")) + root("Keep", block("B", "Person")))
    assert "Drop" not in multiple and "Keep heading" in multiple
    assert "Person" in cleanup(block("Legacy", "Person"))
    independent = cleanup(root("MedicalCompany", block("MedicalLeave")) + block("Transition", "Independent event"))
    assert "MedicalCompany" not in independent and "Independent event" in independent
    assert cleanup(block("Legacy")) == ""
    for invalid in (root("R", root("N", block("A"))),
                    "{\\root:R_Begin}" + block("A"),
                    "{\\root:R_Begin}{\\block:A_Begin}{\\root:R_End}{\\block:A_End}",
                    root("R", block("A")) + root("R", block("B"))):
        try:
            index(invalid)
        except ValueError:
            pass
        else:
            raise AssertionError("Malformed hierarchy accepted")
    path = FOLDER / TEMPLATE_NAME
    with ZipFile(path) as package:
        xml = package.read("word/document.xml").decode("utf-8")
    assert add_roots(xml) == xml, "Migration must be idempotent"
    body = ET.fromstring(xml).find("w:body", NS)
    texts = ["".join(node.itertext()) for node in body.findall("w:p", NS)]
    sections = index("\n".join(texts))
    roots = [entry for entry in sections if entry[0] == "root"]
    assert len(roots) == 6
    assert cleanup("\n".join(texts)).count("Вважати такими") == 0
    for kind, name, first, last in sections:
        if kind == "block":
            expected = 0 if name.startswith("MoveFrom") else 1
            assert sum(start.end() <= first.start() and last.end() <= end.start() for _, _, start, end in roots) == expected, name
    final_child = texts.index("{\\block:ToMedicalCompanyVacationForCuring_End}")
    assert texts[final_child + 1] == "{\\root:OrderSection6_End}"
    assert texts[final_child + 2] == "{\\block:MoveFromHospitalToVacationForCuring_Begin}"
    source = (ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder/obj_PEB_WordAnchors.cls.vba").read_text(encoding="utf-8-sig")
    check_declarations(source)
    check_blocks(source)
    assert "private_IndexRoots" in procedure(source, "Initialize")
    plan = procedure(source, "private_PlanCleanup")
    assert plan.index("For Each root In m_roots.Keys") < plan.index("If skipped.Exists(identifier) Then")
    assert "children.Add identifier" in plan and "skipped(identifier) = True" in plan
    assert "Word root intersects a child section" in procedure(source, "private_IndexRoots")
    print("PASS: empty/mixed roots, legacy blocks, invalid nesting, six template roots and VBA structure. Word/VBA not executed.")


if __name__ == "__main__":
    main()