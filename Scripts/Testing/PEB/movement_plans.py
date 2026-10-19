"""Read native plan declarations and configured values for offline checks only."""

from pathlib import Path
import re


def load_native_plans(settings, source=None):
    path = Path(__file__).resolve().parents[3] / "PROJECTS_EXCEL/vba/PersonalEventBuilder/ex_PEB_MovementRegistry.vba"
    source = path.read_text(encoding="utf-8-sig") if source is None else source
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
    for match in re.finditer(r'private_AddMappings ("(?:[^"]|"")*"),\s*VBA.Array\((.*?)\),\s*VBA.Array\((.*?)\)', body("private_BuildProfiles")):
        prefix = strings(match.group(1))[0]
        fields, headers = strings(match.group(2)), strings(match.group(3))
        require(fields and len(fields) == len(headers), "Invalid mapping: " + prefix)
        require(len({header.casefold() for header in headers}) == len(headers), "Duplicate native mapping header")
        values[prefix + ".count"] = str(len(fields))
        for number, (field, header) in enumerate(zip(fields, headers), 1):
            values[f"{prefix}.{number}.field"] = field
            values[f"{prefix}.{number}.header"] = header
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
    return {"events": events, "values": values, "statuses": statuses, "profiles": profiles}