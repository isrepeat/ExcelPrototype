"""Resolve alias-qualified references and relative imports in fixture checks."""

from pathlib import Path
import re
from xml.etree import ElementTree as ET


NS = {"p": "urn:excelprototype:profiles"}


def import_path(routes_path, alias, root=None):
    routes_path = Path(routes_path)
    if root is None:
        root = ET.parse(routes_path).getroot()
    import_lists = root.findall("p:imports", NS)
    if len(import_lists) != 1 or import_lists[0].attrib:
        raise ValueError("Expected one XML imports element")
    imports = {}
    for node in import_lists[0]:
        if node.tag != f"{{{NS['p']}}}import" or len(node):
            raise ValueError(f"Invalid XML import element: {node.tag}")
        name, path = node.get("alias", ""), node.get("path", "")
        if set(node.attrib) != {"alias", "path"} or not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]*", name) or name in imports:
            raise ValueError(f"Invalid or duplicate XML import alias: {name}")
        if not path or ":" in path or path.startswith(("/", "\\")):
            raise ValueError(f"XML import requires a relative path: {path}")
        imports[name] = path.replace("\\", "/")
    if alias not in imports:
        raise ValueError(f"Unknown XML import alias: {alias}")
    return (routes_path.parent / imports[alias]).resolve()


def reference_id(reference, alias):
    if not re.fullmatch(rf"{alias}:[A-Za-z][A-Za-z0-9_.]*", reference):
        raise ValueError(f"Expected {alias}:ID reference: {reference}")
    return reference.split(":", 1)[1]