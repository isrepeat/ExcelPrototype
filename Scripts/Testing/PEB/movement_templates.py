"""Validate reusable text DSL templates without table-write instructions."""

from pathlib import Path
from xml.etree import ElementTree as ET
import re
from dsl_resources import load_resources
from template_rejections import validate_rejections

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
    require(all(child.tag in {f"{{{NS['p']}}}{name}" for name in ("resources", "sharedTemplates", "template")} for child in root), "Unknown result element")
    known = set(settings["export.fields"].split("|")) | set(settings["export.shared.inputs.scalars"].split("|")) | {
        "EventName", "SourceEventName", "SourceDestination", "TicketDocument", "FoodSupportDate",
        "SourceStatus", "TargetStatus", "Operation", "TvoReplacedPersonName", "TvoIsChain",
    }
    for section, key in (("lookups", "output"), ("aliases", "name"), ("constants", "name")):
        for number in range(1, int(settings[f"export.shared.{section}.count"]) + 1):
            known.add(settings[f"export.shared.{section}.{number}.{key}"])
    for name in settings["export.shared.inputs.collections"].split("|"):
        known.add(name)
        known.update(settings[f"export.shared.inputs.{name}.fields"].split("|"))
    shared, templates = {}, {}
    load_resources(root, known)
    column_texts = {}
    tvo_texts = {}
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
        validate_rejections(node, settings, known | {"VacationAdditionalDays"})
        require(node.findall("p:column", NS) and all(child.tag in {f"{{{NS['p']}}}column", f"{{{NS['p']}}}reject"} for child in node), "Expected result columns: " + name)
        texts, names = {}, set()
        duty_texts, duty_names = {}, set()
        for column in node.findall("p:column", NS):
            attributes(column, {"configKey", "target"}, {"configKey"})
            target = column.get("target", "Absence")
            require(target in {"Absence", "TVO"}, "Unknown result text target: " + target)
            selected_texts = duty_texts if target == "TVO" else texts
            selected_names = duty_names if target == "TVO" else names
            config_key = column.get("configKey").strip()
            require(config_key in settings and settings[config_key].strip(), "Missing configuration key: " + config_key)
            column_name = settings[config_key]
            require(column_name.casefold() not in selected_names, "Duplicate result column: " + column_name)
            selected_names.add(column_name.casefold())
            require([child.tag for child in column] == [f"{{{NS['p']}}}text"], "Expected one column text: " + column_name)
            text = column.find("p:text", NS).text or ""
            require(text.strip(), "Empty result column text: " + column_name)
            selected_texts[column_name] = text
        require(texts, "Expected absence columns: " + name)
        templates[name] = node
        column_texts[name] = texts
        tvo_texts[name] = duty_texts
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
    for texts in tvo_texts.values():
        for body in texts.values():
            validate_text(body)
    require(templates, "No text templates")
    return {"root": root, "templates": templates, "shared": shared, "columns": column_texts, "tvo_columns": tvo_texts}