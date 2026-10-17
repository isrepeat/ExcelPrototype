"""Emit the one-time XML registry migration as a visible apply_patch patch."""

from pathlib import Path
import difflib
from xml.sax.saxutils import quoteattr

ROOT = Path(__file__).resolve().parents[3]
CONFIG = ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder"


def settings(path):
    return {parts[1].split("::", 1)[-1]: parts[2] for line in path.read_text(encoding="utf-8-sig").splitlines() for parts in [line.split("\t")]}


def element(tag, attributes, indent, close=True):
    items = list(attributes.items())
    lines = [" " * indent + f"<{tag} {items[0][0]}={quoteattr(items[0][1])}"]
    lines += [" " * (indent + len(tag) + 2) + f"{key}={quoteattr(value)}" for key, value in items[1:]]
    lines[-1] += "/>" if close else ">"
    return lines


if __name__ == "__main__":
    source = settings(CONFIG / "wsConfig.test.txt")
    if "events.ids" not in source:
        raise SystemExit("Migration already applied; edit the XML registry directly.")
    events = source["events.ids"].split("|")
    names = {
        "Стаціонарне лікування": "HospitalTreatment",
        "Щорічна відпустка": "AnnualVacation",
        "Відпустка для лікування": "MedicalLeave",
        "Відпустка за сімейними обставинами": "FamilyVacation",
        "Відпустка у зв'язку з вагітністю та пологами": "MaternityLeave",
        "Відрядження": "BusinessTrip",
        "Амбулаторне ВЛК": "OutpatientVlk",
        "ВЛК за межами": "ExternalVlk",
    }
    visible = {"FromHospital", "FromAnnualMainVacation", "FromVacationForCuring", "FromVacationForFamily", "ToHospital", "ToAnnualMainVacation", "ToVacationForCuring", "ToVacationForFamily"}
    xml = ['<?xml version="1.0" encoding="utf-8"?>', '<eventRoutes xmlns="urn:excelprototype:profiles"', '             version="1"', '             bindingField="ModeId">', '  <movementStatuses>']
    for name, identifier in names.items():
        xml += element("status", {"id": identifier, "name": name}, 4)
    xml += ['  </movementStatuses>', '  <wordAnchors>', '    <anchor id="EventBody"', '            bookmark="EventBody"/>', '  </wordAnchors>', '  <events>']
    for event in events:
        prefix = f"events.{event}"
        plan = source[prefix + ".movement.plan"]
        operation = source[f"export.Movement.profiles.{plan}.operation"].capitalize()
        operation = {"Append": "Departure", "Close": "Arrival", "Transition": "Transition"}[operation]
        event_name = source[prefix + ".movement.event"]
        movement = {"operation": operation, "profile": plan}
        if operation == "Arrival":
            movement["sourceStatus"] = names[event_name]
        else:
            if operation == "Transition":
                movement["sourceStatus"] = names[source[prefix + ".movement.sourceEvent"]]
            movement["targetStatus"] = names[event_name]
        movement["destinationField"] = source[f"export.Movement.{event}.destinationField"]
        if source[prefix + ".movement.destinationPolicy"] == "Constant":
            movement["destinationValue"] = source[prefix + ".movement.destinationValue"]
        if source[prefix + ".movement.sourceDestination"] != "None":
            movement["sourceDestination"] = source[prefix + ".movement.sourceDestination"]
        movement["required"] = source[f"export.Movement.{event}.required"]
        xml += element("event", {"id": event, "caption": source[f"modes.{event}.caption"], "label": source[f"export.events.{event}"], "direction": source[f"modes.{event}.direction"], "form": source[f"modes.{event}.form"], "visible": str(event in visible).lower()}, 4, False)
        xml += element("movement", movement, 6)
        xml += element("word", {"policy": source[prefix + ".word.policy"], "profile": source[prefix + ".word.profile"], "template": source[prefix + ".word.template"], "anchor": "EventBody"}, 6)
        xml += ["    </event>"]
    xml += ["  </events>", "</eventRoutes>"]
    patch = ["*** Begin Patch", "*** Add File: " + (ROOT / "PROJECTS_EXCEL/workbook/PersonalEventBuilder/PersonalEventBuilderEventRoutes.xml").as_posix()]
    patch += ["+" + line for line in xml]
    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        path = CONFIG / filename
        original = path.read_text(encoding="utf-8-sig").splitlines()
        def keep(line):
            key = line.split("\t")[1].split("::", 1)[-1]
            if key.startswith(("events.", "export.events.", "test.registry.ModeId.")):
                return key == "test.registry.ModeId.rangeName"
            if key.startswith("modes."):
                return key == "modes.default"
            if key.startswith("export.Movement.") and key.split(".")[2] in events:
                return False
            if key.startswith("export.Word.anchors.") or key == "export.Word.bookmark":
                return False
            if key.startswith("export.Word.") and key.endswith(".templateName"):
                return False
            if key.startswith("export.Word.MoveFromOutpatientVlkToVacationForCuring."):
                return False
            return True
        updated = [line for line in original if keep(line)]
        updated.insert(10, "value\tPersonalEventBuilder::events.path\tPersonalEventBuilderEventRoutes.xml")
        updated = [line.replace("Джерело списків: wsConfig.test.txt. Змінюйте конфігурацію, а не цей лист.", "Події: PersonalEventBuilderEventRoutes.xml; інші значення: TEST-конфіг.") for line in updated]
        diff = list(difflib.unified_diff(original, updated, n=3))[2:]
        patch += ["*** Update File: " + path.as_posix()]
        patch += ["@@" if line.startswith("@@") else line for line in diff]
    print("\n".join(patch + ["*** End Patch"]))