"""Read native plan declarations and configured values for offline checks only."""

from pathlib import Path
import re
from xml.etree import ElementTree as ET


def transition_rules(root):
    ns = {"p": "urn:excelprototype:profiles"}
    transitions = root.find("p:transitions", ns)
    if transitions is None:
        return
    default_section_anchor = transitions.get("wordSectionAnchor")
    statuses = {}
    for node in root.findall("p:transitions/p:statuses/p:status", ns):
        status_id = node.attrib["id"]
        if status_id in statuses:
            raise ValueError(f"Duplicate transition status: {status_id}")
        statuses[status_id] = node.attrib

    def caption(status):
        return statuses[status["captionRef"]]["caption"] if "captionRef" in status else status["caption"]

    seen = set()
    for route in transition_nodes(transitions, ns):
        source_id = route.attrib["source"]
        if source_id not in statuses:
            raise ValueError(f"Unknown transition source: {source_id}")
        source = statuses[source_id]
        source_caption = source.get("sourceCaption")
        if source_caption is None:
            text = caption(source)
            source_caption = text[:1].upper() + text[1:]
        for target_id in route.attrib["targets"].split("|"):
            if target_id not in statuses:
                raise ValueError(f"Unknown transition target: {target_id}")
            target = statuses[target_id]
            identifier = f"MoveFrom{source_id}To{target_id}"
            if identifier in seen:
                raise ValueError(f"Duplicate transition: {identifier}")
            seen.add(identifier)
            target_caption = route.get("targetCaption", caption(target))
            section_anchor = route.get("wordSectionAnchor") or default_section_anchor
            yield identifier, source_id, source, target_id, target, f"{source_caption} → {target_caption}", route, section_anchor


def transition_nodes(transitions, ns):
    sources = set()
    for parent in transitions.findall("p:routes/p:transition", ns):
        source = parent.attrib["source"]
        if source in sources:
            raise ValueError(f"Duplicate transition source: {source}")
        sources.add(source)
        if not len(parent):
            yield parent
            continue
        if set(parent.attrib) != {"source"}:
            raise ValueError(f"Grouped transition may only declare its source: {source}")
        for child in parent:
            if child.tag != f"{{{ns['p']}}}route" or len(child) or "source" in child.attrib:
                raise ValueError(f"Invalid child route for transition source: {source}")
            yield ET.Element(child.tag, dict(child.attrib, source=source))


def load_native_plans(settings, source=None, event_root=None):
    path = Path(__file__).resolve().parents[3] / "PROJECTS_EXCEL/vba/PersonalEventBuilder/ex_PEB_MovementRegistry.vba"
    source = path.read_text(encoding="utf-8-sig") if source is None else source
    tvo_enabled = bool(re.search(r'(?mi)^#Const TVO_DEPRECATED = True$', source))
    source = re.sub(r'(?ms)^#If TVO_DEPRECATED Then\n(.*?)^#End If$',
                    lambda match: match[1] if tvo_enabled else '', source)
    def require(condition, message):
        if not condition:
            raise ValueError("Native Movement plan: " + message)
    def strings(text):
        return [value.replace('""', '"') for value in re.findall(r'"((?:[^"]|"")*)"', text)]
    def configured_value(match):
        key = match.group(1)
        require(key in settings and settings[key], "Missing configuration key: " + key)
        return '"' + settings[key].replace('"', '""') + '"'
    source = re.sub(r'ex_PEB_Export\.fn_Config\("([^"]+)"\)', configured_value, source)
    def body(name):
        match = re.search(r"Private Sub " + name + r"\(\)(.*?)\nEnd Sub", source, re.S)
        require(match, "Missing native builder: " + name)
        return re.sub(r"\s+_\r?\n\s*", " ", match.group(1))
    values = {}
    profiles = {}
    for match in re.finditer(r"(?m)^\s*private_AddProfile (.+)$", body("private_BuildProfiles")):
        args = strings(match.group(1))
        require(len(args) == 6, "Invalid profile declaration")
        name, *items = args
        require(name not in profiles, "Duplicate profile: " + name)
        profiles[name] = dict(zip(("operation", "openHeader", "closeProfile", "arrivalDateField", "departureDateHeader"), items))
        for key, value in profiles[name].items():
            if value:
                values[f"profiles.{name}.{key}"] = value
    arrays = {}
    def expand_arrays(match):
        if match[1]:
            arrays[match[1]] = match[2]
            return match[0]
        return f'private_AddMappings {match[3]}, VBA.Array({arrays["fields"]}), VBA.Array({arrays["headers"]})'
    profile_body = re.sub(
        r'(?m)^\s*(fields|headers) = VBA.Array\((.*?)\)|private_AddMappings ("[^"]+"), fields, headers',
        expand_arrays, body("private_BuildProfiles"))
    for match in re.finditer(r'private_AddMappings ("(?:[^"]|"")*"),\s*VBA.Array\((.*?)\),\s*VBA.Array\((.*?)\)', profile_body):
        prefix = strings(match.group(1))[0]
        fields, headers = strings(match.group(2)), strings(match.group(3))
        require(fields and len(fields) == len(headers), "Invalid mapping: " + prefix)
        require(len({header.casefold() for header in headers}) == len(headers), "Duplicate native mapping header")
        values[prefix + ".count"] = str(len(fields))
        for number, (field, header) in enumerate(zip(fields, headers), 1):
            values[f"{prefix}.{number}.field"] = field
            values[f"{prefix}.{number}.header"] = header
    for match in re.finditer(r'private_AddTvoMappings "([^"]+)"', profile_body):
        prefix = match[1]
        count = int(values[prefix + '.count'])
        for number, field in enumerate(('TvoName', 'TvoPersonId', 'TvoPositionCode'), count + 1):
            values[f'{prefix}.{number}.field'] = field
            values[f'{prefix}.{number}.header'] = settings['returnTvo.headers.' + field]
        values[prefix + '.count'] = str(count + 3)
    for match in re.finditer(r'm_values.Add ("profiles\.[^"]+\.guard.openHeader"), ("[^"]+")', body("private_BuildProfiles")):
        key, value = strings(match.group(0))
        values[key] = value
    identifiers = strings(re.search(r"For Each identifier In VBA.Array\((.*?)\)", source).group(1))
    statuses = {identifier: settings["export.Movement.status." + identifier] for identifier in identifiers}
    require(len(set(statuses.values())) == len(statuses), "Duplicate status name")
    events = {}
    form_fields = set(settings["export.fields"].split("|"))
    for match in re.finditer(r"(?m)^\s*private_AddEvent (.+)$", body("private_BuildEvents")):
        args = strings(match.group(1))
        require(len(args) == 9, "Invalid event declaration")
        identifier, *items = args
        require(identifier not in events, "Duplicate event: " + identifier)
        plan = dict(zip(("operation", "profile", "sourceStatus", "targetStatus", "destinationField",
                         "destinationValue", "sourceDestination", "required"), items))
        require(plan["profile"] in profiles, "Unknown profile: " + identifier)
        expected = {"Arrival": "close", "Departure": "append", "Transition": "transition"}.get(plan["operation"])
        require(profiles[plan["profile"]]["operation"] == expected, "Operation mismatch: " + identifier)
        require(set(plan["required"].split("|")) | {plan["destinationField"]} <= form_fields, "Unknown form input: " + identifier)
        require(all(not plan[key] or plan[key] in statuses for key in ("sourceStatus", "targetStatus")), "Unknown status: " + identifier)
        events[identifier] = plan
    require(events and profiles, "Empty native declarations")
    generated = {}
    section_anchors = {}
    if event_root is None:
        event_root = ET.parse(path.parents[2] / "workbook/PersonalEventBuilder" / settings["events.path"]).getroot()
    for identifier, source_id, source, target_id, target, caption, route, section_anchor in transition_rules(event_root):
        if route.get("form"):
            continue
        source_status = source.get("movementStatus", source_id)
        require(source_status in statuses, f"Unknown transition source: {source_status}")
        legacy_departure_ids = route.get("legacyDepartureIds", "false")
        require(legacy_departure_ids in ("true", "false"), f"Invalid legacyDepartureIds: {source_id}")
        require(bool(source.get("wordFrom")), f"Missing transition Word block: {source_id}")
        base = target["departure"]
        require(base in events and events[base]["operation"] == "Departure", f"Unknown transition departure: {base}")
        require(identifier not in events, f"Duplicate event: {identifier}")
        events[identifier] = dict(events[base], operation="Transition", profile="Transition", sourceStatus=source_status)
        generated[identifier] = (base, caption, f"TransitionTo{target_id}", identifier)
        section_anchors[identifier] = section_anchor
        legacy_id = f"MoveFrom{source_id}{base}"
        if legacy_departure_ids == "true" and legacy_id != identifier:
            require(legacy_id not in events, f"Duplicate legacy event: {legacy_id}")
            events[legacy_id] = events[identifier]
            generated[legacy_id] = generated[identifier]
            section_anchors[legacy_id] = section_anchor
    return {"events": events, "values": values, "statuses": statuses, "profiles": profiles,
            "generated": generated, "section_anchors": section_anchors}