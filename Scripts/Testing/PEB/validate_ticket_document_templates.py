"""Check ticket naming and duplicate suppression without Office."""

import re
from xml.etree import ElementTree as ET

from event_routes import ROOT, parse_config
from movement_templates import load_movement
from dsl_resources import load_resources, lookup


def main():
    folder = ROOT / "PROJECTS_EXCEL/workbook/PersonalEventBuilder"
    word_path = folder / "PersonalEventBuilderWordResultTemplates.xml"
    movement_path = folder / "PersonalEventBuilderMovementTemplates.xml"
    word = ET.parse(word_path).getroot()
    returned = next(node for node in word.iter() if node.get("id") == "block.ReturnFromVacation")
    return_text = "".join(returned.itertext())
    assert return_text.index("{#include block.OptionalBasisDocuments}") < return_text.index("{#include block.OptionalVacationTicket}")
    movement = ET.parse(movement_path).getroot()
    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        settings = parse_config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder" / filename)
        load_movement(settings, folder)
    pattern = r"^\s*[Вв]\s*/\s*[Кк]\s*$"
    ticket_pattern = r"^\s*(?:[Вв]\s*/\s*[Кк]|[Вв]ідпускний\s+квиток)\s*$"
    for root in (word, movement):
        shared = {node.get("id"): "".join(node.itertext()) for node in root.findall("{*}sharedTemplates/{*}sharedTemplate")}
        ticket = shared["block.OptionalVacationTicket"]
        assert "#let __TicketInDocument" in ticket
        assert "{#if __TicketInDocument}{#else}" in ticket
        assert ticket_pattern in ticket
        documents = shared["block.OptionalBasisDocuments"]
        assert '{#for item in BasisDocuments orderBy:"BasisDocuments"}' in documents
        if root is word:
            assert 'm_RegexIsMatch' not in documents
        else:
            assert '__IsSupportingDocument' in documents
        for value in ('В/к', ' В / К ', 'відпускний квиток', 'Відпускний квиток'):
            assert re.fullmatch(ticket_pattern, value)
        assert not re.fullmatch(ticket_pattern, 'довідка військово-лікарської комісії')
    names = "".join(next(node for node in word.iter() if node.get("id") == "block.DocumentName").itertext())
    assert 'lookup:"DocumentNames"' in names
    resources = load_resources(word, {'SourceStatus', 'ModeId'})
    for value in ("В/к", "в/к", " В / К "):
        assert lookup(resources, "DocumentNames", value, {'SourceStatus': ''}) == "відпускний квиток"
    for value in ("В/к", "в/к", " В / К "):
        assert re.sub(pattern, "відпускний квиток", value) == "відпускний квиток"
        assert re.fullmatch(pattern, value)
    assert re.sub(pattern, "відпускний квиток", "довідка ВЛК") == "довідка ВЛК"
    assert '; відпускний квиток' not in movement_path.read_text(encoding="utf-8-sig")
    assert '; відпускний квиток №' not in word_path.read_text(encoding="utf-8-sig")
    supporting_pattern = r"^\s*(?:[Вв]\s*/\s*[Кк]|[Вв]ідпускний\s+квиток|[Пп]освідчення\s+про\s+відрядження)\s*$"
    movement_shared = {node.get("id"): "".join(node.itertext()) for node in movement.findall("{*}sharedTemplates/{*}sharedTemplate")}
    for name in ("block.OptionalBasisDocuments", "block.OptionalDepartureBasisDocuments"):
        assert supporting_pattern in movement_shared[name]
    for node in movement.findall("{*}template"):
        text = "".join(node.itertext())
        assert "{#include block.OptionalBusinessTripCertificate}" not in text
        assert "{#include block.OptionalVacationTicket}" not in text
        assert "{#include block.OptionalPreviousVacationTicket}" not in text
    word_shared = {node.get("id"): "".join(node.itertext()) for node in word.findall("{*}sharedTemplates/{*}sharedTemplate")}
    assert supporting_pattern in word_shared["block.BusinessTripBasisDocuments"]
    assert '[VacationTicketNo]' in word_shared["block.BusinessTripBasisDocuments"]
    for value in ('В/к', ' В / К ', 'відпускний квиток', 'Посвідчення про відрядження'):
        assert re.fullmatch(supporting_pattern, value)
    assert not re.fullmatch(supporting_pattern, 'м/к')
    for name in ('ToBusinessTrip', 'FromBusinessTrip'):
        node = next(node for node in word.findall('{*}template') if node.get('name') == name)
        assert '{#include block.BusinessTripBasisDocuments}' in ''.join(node.itertext())
    print("PASS: Word ticket expansion, Movement abbreviation and duplicate-suppression contracts. VBA was not executed.")


if __name__ == "__main__":
    main()