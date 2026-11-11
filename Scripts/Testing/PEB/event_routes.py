"""Read and validate the shared business-event registry for offline fixture tools."""

from pathlib import Path
from xml.etree import ElementTree as ET
import re

from movement_templates import load_movement
from copy import deepcopy
from movement_plans import load_native_plans, transition_rules
from config_dependencies import load_config
from dsl_resources import load_resources
from template_rejections import validate_rejections
from xml_imports import import_path, reference_id

ROOT = Path(__file__).resolve().parents[3]
WORKBOOK = ROOT / "PROJECTS_EXCEL/workbook/PersonalEventBuilder"
NS = {"p": "urn:excelprototype:profiles"}


def parse_config(path):
    return {key.split("::", 1)[-1]: value for key, value in load_config(path).items()
            if key.startswith(("PersonalEventBuilder::", "ThisWorkbook::"))}


def validate_word_template(template):
    def require(condition, message):
        if not condition:
            raise ValueError("Word template XML: " + message)
    def attributes(node, allowed, required=()):
        require(set(node.attrib) <= set(allowed), "Unknown attribute")
        require(all(node.get(key, "").strip() for key in required), "Missing attribute")
    def text(node):
        attributes(node, ())
        require(not list(node), "Text must contain DSL, not XML elements")
    def text_block(node):
        require(len(node) == 1 and node[0].tag == f"{{{NS['p']}}}text", "Expected one text element")
        text(node[0])
    def group(node):
        attributes(node, ("by", "branchBy"), ("by",))
        require(all(key.strip() for key in node.get("by").split(";")), "Empty grouping key")
        require(all(child.tag in {f"{{{NS['p']}}}{tag}" for tag in ("aggregate", "single", "multiple")} for child in node), "Unknown group element")
        require(len(node.findall("p:single", NS)) == len(node.findall("p:multiple", NS)) == 1, "Expected single and multiple")
        single = node.find("p:single", NS)
        attributes(single, ())
        if single.find("p:text", NS) is not None:
            text_block(single)
        else:
            branch(single)
        multiple = node.find("p:multiple", NS)
        attributes(multiple, ())
        branch(multiple)
        names = set()
        for aggregate in node.findall("p:aggregate", NS):
            keys = ("name", "field", "operation", "pattern", "separator", "rangeSeparator")
            attributes(aggregate, keys, keys)
            require(not list(aggregate), "Aggregate cannot contain elements")
            require(aggregate.get("name") not in names, "Duplicate aggregate")
            names.add(aggregate.get("name"))
            require(aggregate.get("operation") == "numberRanges", "Unknown aggregate operation")
    def branch(multiple):
        require(len(multiple.findall("p:item", NS)) + len(multiple.findall("p:group", NS)) == 1, "Expected one item or nested group")
        for child in multiple:
            if child.tag == f"{{{NS['p']}}}group":
                group(child)
            else:
                require(child.tag in {f"{{{NS['p']}}}{tag}" for tag in ("header", "item", "footer")}, "Unknown multiple element")
                require(sum(other.tag == child.tag for other in multiple) == 1, "Duplicate group block")
                common = ("common",) if child.tag.endswith(("}header", "}footer")) else ()
                attributes(child, common, common)
                text_block(child)
    attributes(template, ("name", "hash", "identityBy", "orderBy"), ("name",))
    validate_rejections(template)
    bodies = template.findall("p:text", NS) + template.findall("p:group", NS)
    require(len(bodies) == 1, "Expected exactly one text or group body")
    body_seen = False
    for child in template:
        if child.tag == f"{{{NS['p']}}}reject":
            require(not body_seen, "Rejection rules must precede the body")
        elif child.tag == f"{{{NS['p']}}}derived":
            require(not body_seen, "Derived fields must precede the body")
            attributes(child, ("name",), ("name",))
            text_block(child)
        else:
            require(child in bodies, "Unknown template element")
            body_seen = True
    if bodies[0].tag == f"{{{NS['p']}}}text":
        require(not {"identityBy", "orderBy"} & template.attrib.keys(), "State keys require a group")
        text(bodies[0])
    else:
        require(bool(template.get("identityBy")) and bool(template.get("orderBy")), "Missing state keys")
        group(bodies[0])


def load_registry(settings, xml_text=None, movement_text=None, word_text=None):
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
    sections = ("imports", "events", "transitions") if root.find("p:transitions", NS) is not None else ("imports", "events")
    for tag in sections:
        require(len(root.findall(f"p:{tag}", NS)) == 1, f"Expected one {tag}")
    require({child.tag for child in root} == {f"{{{NS['p']}}}{tag}" for tag in sections}, "Unknown root element")
    events = index(root.findall("p:events/p:event", NS), "id")
    for event in events.values():
        for alias in ("word", "movement"):
            exporter = event.find("p:" + alias, NS)
            if exporter is not None:
                exporter.set("template", reference_id(exporter.get("template", ""), alias))
    transitions = root.find("p:transitions", NS)
    transition_templates = {}
    for identifier, _, _, _, _, caption, route, section_anchor in transition_rules(root):
        movement_template = reference_id(route.get("movementTemplate") or transitions.attrib["movementTemplate"], "movement")
        transition_templates[identifier] = movement_template
        if not route.get("form"):
            continue
        require(identifier not in events, f"Duplicate transition event: {identifier}")
        label = caption
        if route.get("labelSourceCaption"):
            label = f'{route.get("labelSourceCaption")} → {caption.split(" → ", 1)[1]}'
        event = ET.Element(f"{{{NS['p']}}}event", id=identifier, caption=caption,
                           label=label, direction=route.get("direction", "StatusChange"),
                           form=route.get("form"))
        if route.get("legacy"):
            event.set("legacy", route.get("legacy"))
        if section_anchor:
            event.set("wordSectionAnchor", section_anchor)
        ET.SubElement(event, f"{{{NS['p']}}}movement", template=movement_template)
        word_template = reference_id(route.attrib["wordTemplate"], "word") if "wordTemplate" in route.attrib else identifier
        ET.SubElement(event, f"{{{NS['p']}}}word", template=word_template)
        events[identifier] = event
    require(events, "Empty event registry")
    has_movement = any(event.find("p:movement", NS) is not None for event in events.values())
    movement_registry = load_movement(settings, WORKBOOK, movement_text, event_root=root) if has_movement else {
        "statuses": {}, "values": {}, "templates": {},
    }
    if has_movement:
        movement_registry.update(load_native_plans(settings, event_root=root))
        for identifier, (base, caption, template, canonical) in movement_registry["generated"].items():
            require(identifier not in events and base in events, f"Invalid generated event: {identifier}")
            event = ET.Element(f"{{{NS['p']}}}event", id=identifier, caption=caption, label=caption,
                               direction="StatusChange", form=events[base].get("form"))
            if identifier != canonical:
                event.set("legacy", "true")
            section_anchor = movement_registry["section_anchors"][identifier]
            if section_anchor:
                event.set("wordSectionAnchor", section_anchor)
            ET.SubElement(event, f"{{{NS['p']}}}movement", template=transition_templates[canonical])
            if events[base].find("p:word", NS) is not None:
                ET.SubElement(event, f"{{{NS['p']}}}word", template=template)
            events[identifier] = event
    statuses = {identifier: ET.Element(f"{{{NS['p']}}}status", id=identifier, name=name)
                for identifier, name in movement_registry["statuses"].items()}
    for node in statuses.values():
        attributes(node, "id|name", "id|name")
    require(len({node.get('name') for node in statuses.values()}) == len(statuses), "Duplicate Movement status name")
    templates = {}
    if any(event.find("p:word", NS) is not None for event in events.values()):
        text = word_text if word_text is not None else import_path(path, "word", root).read_text(encoding="utf-8-sig")
        require("<!DOCTYPE" not in text.upper() and "<!ENTITY" not in text.upper(), "Word DTD/entities are prohibited")
        word_root = ET.fromstring(text)
        require(word_root.tag == f"{{{NS['p']}}}wordResultTemplates" and word_root.get("version") == "2", "Invalid Word root/version")
        require(all(child.tag in {f"{{{NS['p']}}}{name}" for name in ("resources", "documentFilepath", "sharedTemplates", "template", "transitionLayouts")} for child in word_root),
                "Unknown Word template section")
        resources = load_resources(word_root, set(settings["export.Word.context.xmlKeys"].split("|")) | {"ModeId"})
        templates = index(word_root.findall("p:template", NS), "name")
        shared = {
            node.attrib["id"] for node in word_root.findall("p:sharedTemplates/p:sharedTemplate", NS)
        }
        for identifier, _, source, target_id, target, _, route, _ in transition_rules(root):
            if route.get("form") or events[target["departure"]].find("p:word", NS) is None:
                continue
            template = templates[f"TransitionTo{target_id}"]
            word_block = reference_id(source["wordFrom"], "word")
            require(word_block in shared, f"Unknown source Word block: {identifier}/{word_block}")
            text = "".join(template.itertext())
            if 'lookup:"TransitionFemaleBlocks"' in text:
                entry = resources["TransitionFemaleBlocks"]["entries"].get(word_block)
                require(entry is not None and entry["default"] in shared,
                        f"Missing female source Word block: {identifier}/{word_block}")
        layouts = index(word_root.findall("p:transitionLayouts/p:layout", NS), "id")
        if layouts:
            for identifier, _, source, _, target, _, route, _ in transition_rules(root):
                layout_id = route.get("wordLayout")
                if not layout_id:
                    continue
                layout_id = reference_id(layout_id, "word")
                require(layout_id in layouts and identifier not in templates, f"Invalid Word transition layout: {identifier}")
                template = ET.Element(f"{{{NS['p']}}}template", name=identifier)
                template.extend(deepcopy(child) for child in layouts[layout_id])
                text_node = template.find("p:text", NS)
                require(text_node is not None, f"Missing Word transition text: {layout_id}")
                for token, key, status in (("@source@", "wordFrom", source), ("@target@", "wordTo", target)):
                    if token in text_node.text:
                        block = route.get(key) or status.get(key)
                        require(bool(block), f"Missing Word transition block: {identifier}/{key}")
                        text_node.text = text_node.text.replace(token, reference_id(block, "word"))
                templates[identifier] = template
        for template in templates.values():
            validate_word_template(template)
            validate_rejections(template, settings,
                                set(settings["export.Word.context.xmlKeys"].split("|")) | {"VacationAdditionalDays"})
    for identifier, event in events.items():
        attributes(event, "id|caption|label|direction|form|legacy|wordSectionAnchor", "id|caption|label|direction|form")
        require(event.get("legacy") in (None, "true"), "Invalid legacy event marker")
        require(event.get("direction") in ("Arrival", "Departure", "StatusChange"), f"Invalid direction: {identifier}")
        require(f"forms.{event.get('form')}.fields.count" in settings, f"Invalid form: {identifier}")
        tags = [child.tag for child in event]
        allowed = {f"{{{NS['p']}}}movement", f"{{{NS['p']}}}word"}
        require(tags and len(tags) == len(set(tags)) and set(tags) <= allowed,
                f"Expected one or both independent exporter routes: {identifier}")
        movement = event.find("p:movement", NS)
        word = event.find("p:word", NS)
        if movement is not None:
            attributes(movement, "template", "template")
            require(not list(movement), f"Unexpected Movement route children: {identifier}")
            name = movement.get("template")
            require(name in movement_registry["templates"], f"Unknown Movement template: {identifier}/{name}")
            require(identifier in movement_registry["events"], f"Unknown native Movement event: {identifier}")
            plan = movement_registry["events"][identifier]
            prefixes = ["profiles." + plan["profile"] + ".fields"]
            values = movement_registry["values"]
            if plan["operation"] == "Transition":
                prefixes.append("profiles." + values["profiles." + plan["profile"] + ".closeProfile"] + ".fields")
            allowed_columns, required_columns = set(), set()
            for prefix in prefixes:
                for index in range(1, int(values[prefix + ".count"]) + 1):
                    header = values[f"{prefix}.{index}.header"].casefold()
                    allowed_columns.add(header)
                    if values[f"{prefix}.{index}.field"] == "MovementBasis":
                        required_columns.add(header)
            provided_columns = {column.casefold() for column in movement_registry["columns"][name]}
            require(provided_columns <= allowed_columns, f"Text column is not writable: {identifier}")
            require(required_columns <= provided_columns, f"Missing required text column: {identifier}")
            movement.attrib.update(plan)
        if word is not None:
            attributes(word, "template", "template")
            require(not list(word), f"Unexpected Word route children: {identifier}")
            require(word.get("template") in templates, f"Unknown Word template: {identifier}")
            profile = "export.Word." + word.get("template")
            for suffix in ("required", "previewRequired"):
                require(bool(settings.get(profile + "." + suffix)), f"Unknown Word preparation settings: {identifier}")
            require(settings.get(profile + ".lookups.count", "").isdigit(), f"Invalid Word preparation settings: {identifier}")
    return {"events": events, "statuses": statuses, "bindingField": root.get("bindingField"),
            "movement": movement_registry, "word_templates": templates}


def derived_settings(settings):
    """Expose an in-memory compatibility view; never persist derived keys to wsConfig."""
    registry = load_registry(settings)
    result = dict(settings)
    result.update({"export.Movement." + key: value for key, value in registry["movement"]["values"].items()})
    events = registry["events"]
    result["events.ids"] = result["test.registry.ModeId.values"] = "|".join(events)
    for identifier, event in events.items():
        result[f"test.registry.ModeId.{identifier}.runtimeMode"] = identifier
        for key in ("caption", "form", "direction"):
            result[f"modes.{identifier}.{key}"] = event.get(key)
        result[f"export.events.{identifier}"] = event.get("label")
        movement = event.find("p:movement", NS)
        word = event.find("p:word", NS)
        result[f"events.{identifier}.word.policy"] = "Render" if word is not None else "Omit"
        if word is not None:
            for key, value in word.attrib.items():
                result[f"events.{identifier}.word.{key}"] = value
        if movement is not None:
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