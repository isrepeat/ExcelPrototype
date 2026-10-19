"""Extend the order template without rewriting unrelated OOXML or ZIP parts."""

from pathlib import Path
import re
import uuid
from xml.etree import ElementTree as ET
from zipfile import ZipFile


ROOT = Path(__file__).resolve().parents[3]
FOLDER = ROOT / "PROJECTS_EXCEL/workbook/PersonalEventBuilder"
SOURCE = "MoveFromHospitalToVacationForCuring"
TARGETS = ("MoveFromAnnualVacationToVacationForCuring", "MoveFromFamilyVacationToVacationForCuring",
           "MoveFromVacationForCuringToAnnualVacation", "MoveFromMedicalCompanyForTreatmentToVacationForCuring",
           "MoveFromMedicalCompanyVacationForCuringToHospital", "MoveFromMedicalCompanyVacationForCuringToVacationForCuring",
           "MoveFromOutpatientVlkToVacationForCuring")
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def main():
    paths = [path for path in FOLDER.glob("*.docx") if not path.name.startswith("~$")]
    if len(paths) != 1:
        raise ValueError("Expected one order template")
    path = paths[0]
    original = path.read_bytes()
    with ZipFile(path) as package:
        entries = package.infolist()
        payloads = {entry.filename: package.read(entry.filename) for entry in entries}
        comment = package.comment
    xml = payloads["word/document.xml"].decode("utf-8")
    paragraphs = list(re.finditer(r"<w:p(?:\s[^>]*?)?(?:\s*/>|>.*?</w:p>)", xml, re.DOTALL))
    nodes = ET.fromstring(xml).find("w:body", NS).findall("w:p", NS)
    texts = ["".join(node.itertext()) for node in nodes]
    assert len(nodes) == len(paragraphs)
    begin = texts.index("{\\export:" + SOURCE + "_Begin}")
    end = texts.index("{\\export:" + SOURCE + "_End}")
    assert end == begin + 1
    source = xml[paragraphs[begin].start():paragraphs[end].end()]
    additions = []
    for target in TARGETS:
        marker = "{\\export:" + target + "_Begin}"
        if marker in texts:
            assert texts.count(marker) == 1
            assert texts.count("{\\export:" + target + "_End}") == 1
            continue
        addition = source.replace(SOURCE, target)
        addition = re.sub(r'(w14:(?:paraId|textId)=")[^"]+(\")',
                          lambda match: match[1] + uuid.uuid4().hex[:8].upper() + match[2], addition)
        additions.append(addition)
    if not additions:
        print("PASS: medical-leave anchors already present")
        return
    offset = paragraphs[end].end()
    addition = "".join(additions)
    changed = xml[:offset] + addition + xml[offset:]
    assert changed[:offset] + changed[offset + len(addition):] == xml
    ET.fromstring(changed)
    temporary = path.with_suffix(".anchors.tmp.docx")
    try:
        with ZipFile(temporary, "w") as target:
            target.comment = comment
            for entry in entries:
                target.writestr(entry, changed.encode("utf-8") if entry.filename == "word/document.xml" else payloads[entry.filename])
        with ZipFile(temporary) as checked:
            assert checked.testzip() is None
            assert all(checked.read(name) == data for name, data in payloads.items() if name != "word/document.xml")
            body = ET.fromstring(checked.read("word/document.xml")).find("w:body", NS)
            joined = ["".join(node.itertext()) for node in body.findall("w:p", NS)]
            for target in TARGETS:
                assert joined.count("{\\export:" + target + "_Begin}") == 1
                assert joined.count("{\\export:" + target + "_End}") == 1
        if path.read_bytes() != original:
            raise ValueError("Order template changed during patch")
        temporary.replace(path)
    finally:
        if temporary.exists():
            temporary.unlink()
    print("PASS: missing anchor pairs added; original content and other ZIP parts preserved")


if __name__ == "__main__":
    main()