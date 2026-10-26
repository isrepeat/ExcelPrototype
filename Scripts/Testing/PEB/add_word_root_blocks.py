"""Wrap top-level order headings while preserving existing OOXML and ZIP parts."""

from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import argparse
import re
import uuid

ROOT = Path(__file__).resolve().parents[3]
FOLDER = ROOT / "PROJECTS_EXCEL/workbook/PersonalEventBuilder"
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
TEMPLATE_NAME = "{yyyy}-{mm}-{dd} №{OrderNum} ({Signatory}) (Александр).docx"


def add_roots(xml):
    paragraphs = list(re.finditer(r"<w:p(?:\s[^>]*?)?(?:\s*/>|>.*?</w:p>)", xml, re.DOTALL))
    body = ET.fromstring(xml).find("w:body", NS)
    nodes = body.findall("w:p", NS)
    if len(nodes) != len(paragraphs):
        raise ValueError("Nested paragraphs require an explicit root-block migration")
    texts = ["".join(node.itertext()) for node in nodes]
    final_child = texts.index("{\\block:ToMedicalCompanyVacationForCuring_End}")
    if any("\\root:" in text for text in texts):
        # Correct the initial migration: status transitions are independent sections.
        closing = texts.index("{\\root:OrderSection6_End}")
        if closing == final_child + 1:
            return xml
        if closing < final_child:
            raise ValueError("Section six ends before its final child")
        marker = paragraphs[closing].group(0)
        offset = paragraphs[final_child].end()
        changed = xml[:paragraphs[closing].start()] + xml[paragraphs[closing].end():]
        changed = changed[:offset] + marker + changed[offset:]
        ET.fromstring(changed)
        return changed
    headings = [i for i, node in enumerate(nodes)
                if (level := node.find("w:pPr/w:numPr/w:ilvl", NS)) is not None
                and level.get(f"{{{NS['w']}}}val") == "0"]
    if not headings:
        raise ValueError("No top-level numbered headings found")
    prototype = next(i for i, text in enumerate(texts) if re.fullmatch(r"\{\\block:[^{}]+_Begin\}", text))
    marker = paragraphs[prototype].group(0)

    def boundary(number, suffix):
        result = marker.replace(texts[prototype], f"{{\\root:OrderSection{number}_{suffix}}}")
        return re.sub(r'(w14:(?:paraId|textId)=")[^"]+(\")',
                      lambda match: match[1] + uuid.uuid4().hex[:8].upper() + match[2], result)

    insertions = {}
    for number, start in enumerate(headings, 1):
        end = paragraphs[headings[number]].start() if number < len(headings) else paragraphs[final_child].end()
        insertions[paragraphs[start].start()] = insertions.get(paragraphs[start].start(), "") + boundary(number, "Begin")
        insertions[end] = insertions.get(end, "") + boundary(number, "End")
    changed = xml
    for offset, addition in sorted(insertions.items(), reverse=True):
        changed = changed[:offset] + addition + changed[offset:]
    ET.fromstring(changed)
    # Removing the generated paragraphs must recover every original XML byte.
    restored = changed
    for addition in insertions.values():
        restored = restored.replace(addition, "", 1)
    if restored != xml:
        raise ValueError("Migration modified existing XML")
    return changed


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--template", type=Path, default=FOLDER / TEMPLATE_NAME)
    args = parser.parse_args()
    path = args.template
    original = path.read_bytes()
    with ZipFile(path) as package:
        entries = package.infolist()
        payloads = {entry.filename: package.read(entry.filename) for entry in entries}
        comment = package.comment
    changed = add_roots(payloads["word/document.xml"].decode("utf-8")).encode("utf-8")
    if changed == payloads["word/document.xml"]:
        print("PASS: root markers already present; template unchanged")
        return
    backup = path.parent / ".backup" / "root-blocks" / path.name
    backup.parent.mkdir(parents=True, exist_ok=True)
    if not backup.exists():
        backup.write_bytes(original)
    temporary = path.with_suffix(".roots.tmp.docx")
    try:
        with ZipFile(temporary, "w") as target:
            target.comment = comment
            for entry in entries:
                target.writestr(entry, changed if entry.filename == "word/document.xml" else payloads[entry.filename])
        with ZipFile(temporary) as checked:
            assert checked.testzip() is None
            assert all(checked.read(name) == data for name, data in payloads.items() if name != "word/document.xml")
        if path.read_bytes() != original:
            raise ValueError("Template changed during migration")
        temporary.replace(path)
    finally:
        if temporary.exists():
            temporary.unlink()
    print("PASS: root markers inserted; existing paragraphs and other ZIP parts unchanged; backup retained")


if __name__ == "__main__":
    main()