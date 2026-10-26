"""Add the standalone hospital-to-annual-vacation continuation DOCX section."""

from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import re
import uuid

from add_word_root_blocks import FOLDER, TEMPLATE_NAME, NS

SOURCE = "MoveFromHospitalToVacationForCuring"
EVENT = "MoveFromHospitalToAnnualVacationContinuation"
LEGACY_EVENT = "MoveFromHospitalToAnnualVacation"


def main():
    path = FOLDER / TEMPLATE_NAME
    original = path.read_bytes()
    with ZipFile(path) as package:
        entries = package.infolist()
        payloads = {entry.filename: package.read(entry.filename) for entry in entries}
        comment = package.comment
    xml = payloads["word/document.xml"].decode("utf-8")
    renamed = LEGACY_EVENT + "_Begin}" in xml
    if renamed:
        xml = xml.replace(LEGACY_EVENT + "_", EVENT + "_")
    nodes = ET.fromstring(xml).find("w:body", NS).findall("w:p", NS)
    texts = ["".join(node.itertext()) for node in nodes]
    if "{\\block:" + EVENT + "_Begin}" in texts:
        assert all(texts.count("{\\" + kind + ":" + EVENT + "_" + side + "}") == 1
                   for kind in ("block", "export") for side in ("Begin", "End"))
        if not renamed:
            print("PASS: continuation anchors already present")
            return
        changed = xml
    else:
        paragraphs = list(re.finditer(r"<w:p(?:\s[^>]*?)?(?:\s*/>|>.*?</w:p>)", xml, re.DOTALL))
        assert len(paragraphs) == len(nodes)
        begin = texts.index("{\\block:" + SOURCE + "_Begin}")
        end = texts.index("{\\block:" + SOURCE + "_End}")
        source = xml[paragraphs[begin].start():paragraphs[end].end()]
        addition = source.replace(SOURCE, EVENT)
        addition = re.sub(r'(w14:(?:paraId|textId)=")[^"]+(\")',
                          lambda match: match[1] + uuid.uuid4().hex[:8].upper() + match[2], addition)
        offset = paragraphs[end].end()
        changed = xml[:offset] + addition + xml[offset:]
        assert changed[:offset] + changed[offset + len(addition):] == xml
    ET.fromstring(changed)
    backup = path.parent / ".backup" / "annual-continuation" / path.name
    backup.parent.mkdir(parents=True, exist_ok=True)
    if not backup.exists():
        backup.write_bytes(original)
    temporary = path.with_suffix(".continuation.tmp.docx")
    try:
        with ZipFile(temporary, "w") as target:
            target.comment = comment
            for entry in entries:
                target.writestr(entry, changed.encode("utf-8") if entry.filename == "word/document.xml" else payloads[entry.filename])
        with ZipFile(temporary) as checked:
            assert checked.testzip() is None
            assert all(checked.read(name) == data for name, data in payloads.items() if name != "word/document.xml")
        if path.read_bytes() != original:
            raise ValueError("Template changed during migration")
        temporary.replace(path)
    finally:
        if temporary.exists():
            temporary.unlink()
    print("PASS: standalone continuation section added; original XML and other ZIP parts preserved")


if __name__ == "__main__":
    main()