"""Check technical DSL structure fixtures without Office or VBA execution."""

from xml.etree import ElementTree as ET

from event_routes import ROOT, validate_word_template


FIXTURE = ROOT / "PROJECTS_EXCEL/tests/PersonalEventBuilder/WordTemplateGrammarCases.xml"


def validate_cases():
    cases = ET.parse(FIXTURE).getroot().findall("grammar/case")
    identifiers = set()
    for case in cases:
        identifier = case.get("id")
        assert identifier and identifier not in identifiers, identifier
        identifiers.add(identifier)
        assert case.get("valid") in ("true", "false"), identifier
        template = ET.fromstring(case.text)
        try:
            validate_word_template(template)
        except ValueError as error:
            assert case.get("valid") == "false", f"{identifier}: {error}"
        else:
            assert case.get("valid") == "true", f"{identifier}: invalid template was accepted"
    assert cases, "No grammar cases found"
    return len(cases)


if __name__ == "__main__":
    count = validate_cases()
    print(f"PASS: {count} DSL structure cases. VBA rendering and resolving were not executed.")