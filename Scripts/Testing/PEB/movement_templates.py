"""Validate reusable text DSL templates without table-write instructions."""

from pathlib import Path
from xml.etree import ElementTree as ET
import re

NS = {"p": "urn:excelprototype:profiles"}


def load_movement(settings, workbook, xml_text=None):
    def require(condition, message):
        if not condition:
            raise ValueError("Result text templates: " + message)
    def attributes(node, allowed, required):
        require(set(node.attrib) <= set(allowed), "Unknown text-template attribute")
        require(all(node.get(key, "").strip() for key in required), "Missing text-template attribute")
    text = xml_text if xml_text is not None else (Path(workbook) / settings["export.Movement.templates.path"]).read_text(encoding="utf-8-sig")
    require("<!DOCTYPE" not in text.upper() and "<!ENTITY" not in text.upper(), "DTD/entities prohibited")
    root = ET.fromstring(text)
    require(root.tag == f"{{{NS['p']}}}resultTemplates" and root.get("version") == "1", "Invalid root/version")
    attributes(root, {"version"}, {"version"})
    require(len(root.findall("p:sharedTemplates", NS)) == 1, "Expected one sharedTemplates section")
    require(all(child.tag in {f"{{{NS['p']}}}sharedTemplates", f"{{{NS['p']}}}template"} for child in root), "Unknown result element")
    known = set(settings["export.fields"].split("|")) | set(settings["export.shared.inputs.scalars"].split("|")) | {
        "EventName", "SourceEventName", "SourceDestination", "TicketDocument", "FoodSupportDate",
    }
    for section, key in (("lookups", "output"), ("aliases", "name"), ("constants", "name")):
        for number in range(1, int(settings[f"export.shared.{section}.count"]) + 1):
            known.add(settings[f"export.shared.{section}.{number}.{key}"])
    for name in settings["export.shared.inputs.collections"].split("|"):
        known.add(name)
        known.update(settings[f"export.shared.inputs.{name}.fields"].split("|"))
    shared, templates = {}, {}
    column_texts = {}
    for node in root.findall("p:sharedTemplates/p:sharedTemplate", NS):
        attributes(node, {"id"}, {"id"})
        name = node.get("id")
        require(name not in shared, "Duplicate shared template: " + name)
        require([child.tag for child in node] == [f"{{{NS['p']}}}text"], "Expected shared text: " + name)
        shared[name] = node
    for node in root.findall("p:template", NS):
        attributes(node, {"name"}, {"name"})
        name = node.get("name")
        require(re.fullmatch(r"[A-Za-z][A-Za-z0-9_]*", name), "Invalid template name")
        require(name not in templates, "Duplicate template: " + name)
        require(len(node) > 0 and all(child.tag == f"{{{NS['p']}}}column" for child in node), "Expected result columns: " + name)
        texts, names = {}, set()
        for column in node:
            attributes(column, {"configKey"}, {"configKey"})
            config_key = column.get("configKey").strip()
            require(config_key in settings and settings[config_key].strip(), "Missing configuration key: " + config_key)
            column_name = settings[config_key]
            require(column_name.casefold() not in names, "Duplicate result column: " + column_name)
            names.add(column_name.casefold())
            require([child.tag for child in column] == [f"{{{NS['p']}}}text"], "Expected one column text: " + column_name)
            text = column.find("p:text", NS).text or ""
            require(text.strip(), "Empty result column text: " + column_name)
            texts[column_name] = text
        templates[name] = node
        column_texts[name] = texts
    def validate_text(text, trail=()):
        require(len(trail) <= int(settings["export.dsl.maxDepth"]), "Include-depth limit exceeded")
        for name in re.findall(r"\{#include\s+([^}]+)\}", text):
            name = name.strip()
            require(name in shared and name not in trail, "Unknown/cyclic include: " + name)
            validate_text(shared[name].find("p:text", NS).text or "", (*trail, name))
        for field in re.findall(r"\[([A-Z][A-Za-z0-9_]*)\]", text):
            require(field in known, "Unknown placeholder: " + field)
    for name, node in shared.items():
        validate_text(node.find("p:text", NS).text or "", (name,))
    for texts in column_texts.values():
        for body in texts.values():
            validate_text(body)
    require(templates, "No text templates")
    return {"root": root, "templates": templates, "shared": shared, "columns": column_texts}