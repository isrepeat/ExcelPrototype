"""Restore custom Excel table styles omitted by the artifact-tool exporter."""

from pathlib import Path
from zipfile import ZipFile
import sys
from copy import deepcopy
from xml.etree import ElementTree as ET


original, output = map(Path, sys.argv[1:])
with ZipFile(original) as source, ZipFile(output) as exported:
    styles = ET.fromstring(source.read("xl/styles.xml"))
    parts = [(entry, exported.read(entry.filename)) for entry in exported.infolist()]
for index, (entry, content) in enumerate(parts):
    if entry.filename != "xl/styles.xml":
        continue
    result = ET.fromstring(content)
    namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    for name in ("dxfs", "tableStyles"):
        block = styles.find(namespace + name)
        if block is not None:
            previous = result.find(namespace + name)
            position = len(result) if previous is None else list(result).index(previous)
            if previous is not None:
                result.remove(previous)
            result.insert(position, deepcopy(block))
    parts[index] = entry, ET.tostring(result, encoding="utf-8", xml_declaration=True)
staged = output.with_suffix(".styles.xlsx")
with ZipFile(staged, "w") as package:
    for entry, content in parts:
        package.writestr(entry, content)
staged.replace(output)