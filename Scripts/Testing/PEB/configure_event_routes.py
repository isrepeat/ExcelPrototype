"""Emit an apply_patch migration for business-event routing; never write configuration."""

from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[3]
CONFIG = ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder"
XML = ROOT / "PROJECTS_EXCEL/workbook/PersonalEventBuilder/PersonalEventBuilderWordResultTemplates.xml"
MODES = [node.get("name") for node in ET.parse(XML).getroot() if node.tag.endswith("}template")]
EXTRA = "MoveFromOutpatientVlkToVacationForCuring"
EVENTS = MODES + [EXTRA]
NAMES = {
    "Hospital": "Стаціонарне лікування",
    "AnnualMainVacation": "Щорічна відпустка",
    "AnnualVacation": "Щорічна відпустка",
    "VacationForFamily": "Відпустка за сімейними обставинами",
    "FamilyVacation": "Відпустка за сімейними обставинами",
    "VacationForCuring": "Відпустка для лікування",
    "MaternityLeave": "Відпустка у зв'язку з вагітністю та пологами",
    "BusinessTrip": "Відрядження",
    "OutpatientExaminationOrVlk": "Амбулаторне ВЛК",
    "OutpatientVlk": "Амбулаторне ВЛК",
    "Vlk": "ВЛК за межами",
    "ExternalVlk": "ВЛК за межами",
    "MedicalCompanyForTreatment": "Стаціонарне лікування",
    "MedicalCompanyVacationForCuring": "Відпустка для лікування",
}


def route(event):
    if event.startswith("MoveFrom"):
        source, target = event[8:].split("To", 1)
        plan = "Transition"
    else:
        source = None
        target = event[4:] if event.startswith("From") else event[2:]
        plan = "Arrival" if source is None and event.startswith("From") else "Departure"
    return source, target, plan


patch = ["*** Begin Patch"]
for filename in ("wsConfig.txt", "wsConfig.test.txt"):
    file = CONFIG / filename
    lines = file.read_text(encoding="utf-8-sig").splitlines()
    settings = {row.split("\t")[1].split("::")[-1]: row.split("\t")[2] for row in lines if row}
    additions = {}

    def add(key, value):
        if key not in settings:
            additions[key] = value

    file_start = len(patch)
    patch.append(f"*** Update File: {file.as_posix()}")
    changes = {"test.registry.ModeId.values": "|".join(EVENTS), "test.registry.Action.values": "ExportMovement|PreviewWord|ExportEvent", "test.registry.Outcome.values": "Success|Error|Pending|Omitted"}
    for key, value in changes.items():
        if key in settings and settings[key] != value:
            old = next(line for line in lines if line.split("\t")[1] == f"PersonalEventBuilder::{key}")
            patch.extend(("@@", "-" + old, f"+value\tPersonalEventBuilder::{key}\t{value}"))
    add("events.bindingField", "ModeId")
    add("events.ids", "|".join(EVENTS))
    add("text.ExportWordOmitted", "Для цієї події пункт Word не формується.")
    for event in EVENTS:
        source, target, plan = route(event)
        medical_source = (source or "").startswith("MedicalCompany") or event == "FromMedicalCompanyForTreatment"
        if medical_source:
            plan += "MedicalCompany"
            for key in (f"events.{event}.movement.plan", f"export.Movement.{event}.profile"):
                if key in settings and settings[key] != plan:
                    old = next(line for line in lines if line.split("\t")[1] == f"PersonalEventBuilder::{key}")
                    patch.extend(("@@", "-" + old, f"+value\tPersonalEventBuilder::{key}\t{plan}"))
        template = "ToVacationForCuring" if event == EXTRA else event
        profile = template
        medical_company = target.startswith("MedicalCompany")
        destination = "HospitalShortName" if target in ("Hospital", "ExternalVlk", "Vlk", "OutpatientExaminationOrVlk") else "Destination"
        add(f"events.{event}.movement.plan", plan)
        add(f"events.{event}.movement.event", NAMES[target])
        add(f"events.{event}.movement.sourceEvent", NAMES[source] if source else "None")
        add(f"events.{event}.movement.sourceDestination", "Медична рота А7383" if medical_source else "None")
        add(f"events.{event}.movement.destinationPolicy", "Constant" if medical_company else "Field")
        if medical_company:
            add(f"events.{event}.movement.destinationValue", "Медична рота А7383")
        add(f"events.{event}.word.policy", "Omit" if event in ("FromOutpatientExaminationOrVlk", "ToOutpatientExaminationOrVlk") else "Render")
        add(f"events.{event}.word.profile", profile)
        add(f"events.{event}.word.template", template)
        add(f"events.{event}.word.anchor", template)
        add(f"export.Word.anchors.{template}.bookmark", settings["export.Word.bookmark"])
        add(f"export.Movement.{event}.profile", plan)
        add(f"export.Movement.{event}.destinationField", destination)
        required_mode = "ToHospital" if destination == "HospitalShortName" else "ToVacationForCuring"
        add(f"export.Movement.{event}.required", settings[f"export.Movement.{required_mode}.required"])
        if event == EXTRA:
            add(f"export.events.{event}", "Амбулаторне ВЛК → відпустка для лікування")
            for key, value in settings.items():
                if key.startswith("export.Word.ToVacationForCuring."):
                    add(key.replace(".ToVacationForCuring.", f".{event}."), value)
            add(f"modes.{event}.caption", "Амбулаторне ВЛК → відпустка для лікування")
            add(f"modes.{event}.direction", "Departure")
            add(f"modes.{event}.form", "Leave")
            if "test.registry.ModeId.values" in settings:
                add(f"test.registry.ModeId.{event}.runtimeMode", event)
    for key, value in settings.items():
        if key.startswith("export.Movement.profiles.Departure.") and ".guard." not in key:
            add(key.replace(".Departure.", ".Transition."), "transition" if key.endswith(".operation") else value)
        if key.startswith("export.Movement.profiles.Arrival.join."):
            add(key.replace(".Arrival.", ".Transition."), "SourceEventName" if key.endswith(".2.field") else value)
    for suffix in ("openHeader", "departureDateHeader", "arrivalDateField"):
        add(f"export.Movement.profiles.Transition.{suffix}", settings[f"export.Movement.profiles.Arrival.{suffix}"])
    add("export.Movement.profiles.Transition.closeProfile", "Arrival")
    for base in ("Arrival", "Transition"):
        profile = base + "MedicalCompany"
        for key, value in {**settings, **additions}.copy().items():
            if key.startswith(f"export.Movement.profiles.{base}."):
                add(key.replace(f".{base}.", f".{profile}."), "3" if key.endswith(".join.count") else value)
        add(f"export.Movement.profiles.{profile}.join.3.field", "SourceDestination")
        add(f"export.Movement.profiles.{profile}.join.3.header", "Куди")
    if additions:
        patch.extend(("@@", " " + lines[-1]))
        patch.extend(f"+value\tPersonalEventBuilder::{key}\t{value}" for key, value in additions.items())
    if len(patch) == file_start + 1:
        patch.pop()
patch.append("*** End Patch")
print("\n".join(patch))