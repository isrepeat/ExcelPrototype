"""Add a business-trip section using the original template's departure formatting."""

import hashlib
from pathlib import Path
import re
import tempfile
from xml.etree import ElementTree as ET
import zipfile
from xml.sax.saxutils import escape
import uuid


ROOT = Path(__file__).resolve().parents[3]
PATH = ROOT / "PROJECTS_EXCEL/workbook/PersonalEventBuilder/Template_10.docx"
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def main():
    original = PATH.read_bytes()
    with zipfile.ZipFile(PATH) as package:
        payloads = {entry.filename: package.read(entry.filename) for entry in package.infolist()}
        entries = package.infolist()
        comment = package.comment
    xml = payloads["word/document.xml"].decode("utf-8")
    if "{\\export:ToBusinessTrip_Begin}" in xml:
        raise ValueError("Business-trip section already exists")
    paragraphs = list(re.finditer(r"<w:p(?:\s[^>]*?)?(?:\s*/>|>.*?</w:p>)", xml, re.DOTALL))
    # Find text via the complete document so all original namespace bindings remain in scope.
    nodes = ET.fromstring(xml).find("w:body", NS).findall("w:p", NS)
    texts = ["".join(node.itertext()) for node in nodes]
    assert len(nodes) == len(paragraphs)
    begin = next(index for index, value in enumerate(texts) if "{\\export:ToAnnualMainVacation_Begin}" in value)
    end = next(index for index, value in enumerate(texts) if "{\\export:ToAnnualMainVacation_End}" in value)
    heading = next(index for index in range(begin - 1, -1, -1) if texts[index].strip())
    source = xml[paragraphs[heading].start():paragraphs[end].end()]
    source_heading = paragraphs[heading].group()
    caption = iter(["У відрядження:"])
    replacement = re.sub(r"(<w:t(?:\s[^>]*)?>).*?(</w:t>)", lambda match: match[1] + escape(next(caption, "")) + match[2], source_heading, flags=re.DOTALL)
    assert source_heading != replacement
    addition = replacement + source[len(source_heading):]
    addition = addition.replace("ToAnnualMainVacation", "ToBusinessTrip")
    addition = re.sub(r'(w14:(?:paraId|textId)=")[^"]+(")', lambda match: match[1] + uuid.uuid4().hex[:8].upper() + match[2], addition)
    offset = paragraphs[heading].start()
    changed = (xml[:offset] + addition + xml[offset:]).encode("utf-8")
    ET.fromstring(changed)
    assert changed.decode("utf-8").replace(addition, "", 1) == xml
    backup = Path(tempfile.gettempdir()) / ("Template_10.before-business-trip-" + hashlib.sha256(original).hexdigest()[:12] + ".docx")
    if not backup.exists():
        backup.write_bytes(original)
    temporary = PATH.with_name(PATH.stem + ".business-trip.tmp.docx")
    try:
        with zipfile.ZipFile(temporary, "w") as target:
            target.comment = comment
            for entry in entries:
                target.writestr(entry, changed if entry.filename == "word/document.xml" else payloads[entry.filename])
        with zipfile.ZipFile(temporary) as checked:
            assert checked.testzip() is None
            assert all(checked.read(name) == data for name, data in payloads.items() if name != "word/document.xml")
        if PATH.read_bytes() != original:
            raise ValueError("Template changed during patch")
        temporary.replace(PATH)
    finally:
        if temporary.exists():
            temporary.unlink()
    print("PASS: original template extended; existing XML and all other package parts preserved.")


if __name__ == "__main__":
    main()