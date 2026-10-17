"""Read and validate the shared business-event registry for offline fixture tools."""

from pathlib import Path
from xml.etree import ElementTree as ET
import re

ROOT = Path(__file__).resolve().parents[3]
WORKBOOK = ROOT / "PROJECTS_EXCEL/workbook/PersonalEventBuilder"
NS = {"p": "urn:excelprototype:profiles"}


def parse_config(path):
    result = {}
    for number, line in enumerate(Path(path).read_text(encoding="utf-8-sig").splitlines(), 1):
        if not line:
            continue
        parts = line.split("\t")
        if len(parts) != 3:
            raise ValueError(f"Invalid config row: {path}:{number}")
        key = parts[1].split("::", 1)[-1]
        if key in result:
            raise ValueError(f"Duplicate config key: {key}")
        result[key] = parts[2]
    return result


def load_registry(settings, xml_text=None):
    path = WORKBOOK / settings["events.path"]
    text = xml_text if xml_text is not None else path.read_text(encoding="utf-8-sig")
    if "<!DOCTYPE" in text.upper() or "<!ENTITY" in text.upper():
        raise ValueError("Event registry DTD/entities are prohibited")
    root = ET.fromstring(text)
    def require(condition, message):
        if not condition:
            raise ValueError(f"Event routes XML: {message}")
    def attributes(node, allowed, required):
        require(set(node.attrib) <= set(allowed.split("|")), f"Unknown attribute on {node.tag}")
        for key in required.split("|"):
            require(bool(node.get(key, "").strip()), f"Missing {node.tag}/@{key}")
    def index(nodes, key):
        result = {}
        for node in nodes:
            identifier = node.get(key, "")
            require(re.fullmatch(r"[A-Za-z][A-Za-z0-9_]*", identifier), f"Invalid identifier: {identifier}")
            require(identifier not in result, f"Duplicate identifier: {identifier}")
            result[identifier] = node
        return result
    require(root.tag == f"{{{NS['p']}}}eventRoutes", "Expected eventRoutes root")
    attributes(root, "version|bindingField", "version|bindingField")
    require(root.get("version") == "1", "Unsupported version")
    fields = set(settings["export.fields"].split("|"))
    require(root.get("bindingField") in fields, "Unknown binding field")
    for tag in ("movementStatuses", "wordAnchors", "events"):
        require(len(root.findall(f"p:{tag}", NS)) == 1, f"Expected one {tag}")
    require({child.tag for child in root} == {f"{{{NS['p']}}}{tag}" for tag in ("movementStatuses", "wordAnchors", "events")}, "Unknown root element")
    statuses = index(root.findall("p:movementStatuses/p:status", NS), "id")
    anchors = index(root.findall("p:wordAnchors/p:anchor", NS), "id")
    events = index(root.findall("p:events/p:event", NS), "id")
    require(statuses and anchors and events, "Empty registry section")
    for node in statuses.values():
        attributes(node, "id|name", "id|name")
    require(len({node.get('name') for node in statuses.values()}) == len(statuses), "Duplicate Movement status name")
    for node in anchors.values():
        attributes(node, "id|bookmark", "id|bookmark")
    templates = index(ET.parse(WORKBOOK / settings["export.Word.templates.path"]).getroot().findall("p:template", NS), "name")
    for identifier, event in events.items():
        attributes(event, "id|caption|label|direction|form|visible", "id|caption|label|direction|form|visible")
        require(event.get("direction") in ("Arrival", "Departure", "StatusChange"), f"Invalid direction: {identifier}")
        require(event.get("form") in ("Leave", "Treatment"), f"Invalid form: {identifier}")
        require(event.get("visible") in ("true", "false"), f"Invalid visibility: {identifier}")
        require([child.tag for child in event] == [f"{{{NS['p']}}}movement", f"{{{NS['p']}}}word"], f"Expected movement and word: {identifier}")
        movement, word = event
        attributes(movement, "operation|profile|sourceStatus|targetStatus|destinationField|destinationValue|sourceDestination|required", "operation|profile|destinationField|required")
        operation = movement.get("operation")
        require(operation in ("Departure", "Arrival", "Transition"), f"Unknown operation: {identifier}")
        expected = {"Departure": "append", "Arrival": "close", "Transition": "transition"}[operation]
        profile = "export.Movement.profiles." + movement.get("profile")
        require(settings.get(profile + ".operation") == expected, f"Invalid Movement profile/operation: {identifier}")
        for suffix in ("fields.count", "identity.count"):
            require(settings.get(profile + "." + suffix, "").isdigit(), f"Missing profile mapping: {identifier}/{suffix}")
        if operation == "Transition":
            require(settings.get("export.Movement.profiles." + settings.get(profile + ".closeProfile", "") + ".operation") == "close", f"Invalid close profile: {identifier}")
        for key in ("sourceStatus", "targetStatus"):
            value = movement.get(key)
            require(not value or value in statuses, f"Unknown status: {identifier}/{value}")
        require(operation == "Departure" or movement.get("sourceStatus"), f"Missing source status: {identifier}")
        require(operation == "Arrival" or movement.get("targetStatus"), f"Missing target status: {identifier}")
        require(operation != "Departure" or not movement.get("sourceStatus"), f"Unexpected departure source: {identifier}")
        require(operation != "Arrival" or not movement.get("targetStatus"), f"Unexpected arrival target: {identifier}")
        require(set(movement.get("required").split("|")) <= fields, f"Unknown required field: {identifier}")
        require(movement.get("destinationField") in fields, f"Unknown destination field: {identifier}")
        attributes(word, "policy|profile|template|anchor", "policy|profile|template|anchor")
        require(word.get("policy") in ("Render", "Omit"), f"Unknown Word policy: {identifier}")
        require(word.get("template") in templates, f"Unknown Word template: {identifier}")
        require(word.get("anchor") in anchors, f"Unknown Word anchor: {identifier}")
        profile = "export.Word." + word.get("profile")
        for suffix in ("required", "previewRequired"):
            require(bool(settings.get(profile + "." + suffix)), f"Unknown Word profile: {identifier}")
        require(settings.get(profile + ".lookups.count", "").isdigit(), f"Invalid Word profile: {identifier}")
    return {"events": events, "statuses": statuses, "anchors": anchors, "bindingField": root.get("bindingField")}


def derived_settings(settings):
    """Expose an in-memory compatibility view; never persist derived keys to wsConfig."""
    registry = load_registry(settings)
    result = dict(settings)
    events = registry["events"]
    result["events.ids"] = result["test.registry.ModeId.values"] = "|".join(events)
    for identifier, event in events.items():
        result[f"test.registry.ModeId.{identifier}.runtimeMode"] = identifier
        for key in ("caption", "form", "direction"):
            result[f"modes.{identifier}.{key}"] = event.get(key)
        result[f"export.events.{identifier}"] = event.get("label")
        movement, word = event
        for key, value in word.attrib.items():
            result[f"events.{identifier}.word.{key}"] = value
        for key, value in movement.attrib.items():
            result[f"events.{identifier}.movement.{key}"] = value
        result[f"export.Movement.{identifier}.profile"] = movement.get("profile")
        result[f"export.Movement.{identifier}.required"] = movement.get("required")
        result[f"export.Movement.{identifier}.destinationField"] = movement.get("destinationField")
    return result


if __name__ == "__main__":
    import json
    import sys
    settings = parse_config(sys.argv[-1])
    if "--orders" in sys.argv:
        import datetime
        import openpyxl
        from openpyxl.utils.cell import range_boundaries
        from openpyxl.utils.datetime import to_excel
        book = openpyxl.load_workbook(WORKBOOK / settings["orders.source.path"], data_only=True)
        sheet = book[settings["orders.source.sheet"]]
        left, top, right, bottom = range_boundaries(settings["orders.source.range"])
        headers = [sheet.cell(top, column).value for column in range(left, right + 1)]
        dates = left + headers.index(settings["orders.dateHeader"])
        numbers = left + headers.index(settings["orders.numberHeader"])
        result = {}
        for row in range(top + 1, min(bottom, sheet.max_row) + 1):
            date, number = sheet.cell(row, dates).value, sheet.cell(row, numbers).value
            if date is None and number is None:
                continue
            if isinstance(date, str):
                date = datetime.datetime.strptime(date, "%d.%m.%Y")
            serial = int(to_excel(date, book.epoch)) if isinstance(date, (datetime.datetime, datetime.date)) else int(date)
            values = result.setdefault(str(serial), [])
            if str(number) not in values:
                values.append(str(number))
        print(json.dumps(result, ensure_ascii=False))
    else:
        print(json.dumps(derived_settings(settings), ensure_ascii=False))