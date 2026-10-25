"""Validate the medical company leave routes without executing VBA."""

from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

from event_routes import ROOT, WORKBOOK, NS, load_registry, parse_config


MODE = 'ToMedicalCompanyVacationForCuring'
BASE = 'ToVacationForCuring'

for profile in ('wsConfig.txt', 'wsConfig.test.txt'):
    settings = parse_config(ROOT / 'PROJECTS_EXCEL/config/PersonalEventBuilder' / profile)
    registry = load_registry(settings)
    event = registry['events'][MODE]
    assert event.get('visible') == 'false'
    assert event.get('form') == 'Leave'
    assert MODE in settings['ui.extendedModes'].split('|')
    plan = event.find('p:movement', NS)
    assert plan.get('targetStatus') == 'MedicalLeave'
    assert plan.get('destinationValue') == 'Медична рота А7383'
    for prefix in ('document.fixed.', 'forms.readOnly.modes.'):
        assert settings[prefix + MODE] == settings[prefix + BASE]
    assert settings['extensions.mode.' + MODE + '.default.1.role'] == 'Vlk'
    assert settings['export.Word.textAnchors.' + MODE] == MODE

word = ET.parse(WORKBOOK / 'PersonalEventBuilderWordResultTemplates.xml')
source = (ROOT / 'PROJECTS_EXCEL/vba/PersonalEventBuilder/ex_PEB_MovementRegistry.vba').read_text(encoding='utf-8-sig')
normalization = source.split('Public Function fn_Destination', 1)[1].split('End Function', 1)[0]
assert 'medicalCompany = ex_PEB_Export.fn_Config("export.Movement.destinations.MedicalCompany")' in normalization
guard = normalization.index('VBA.StrComp(VBA.Trim$(section), medicalCompany, VBA.vbTextCompare)')
assert guard < normalization.index('m_destinationRegex.Execute(section)')
assert 'sections(index) = VBA.vbNullString' not in normalization
templates = word.findall('p:template', NS)
assert len({node.get('hash') for node in templates}) == len(templates)
template = next(node for node in templates if node.get('name') == MODE)
assert 'block.VacationDestinationPeriod' in ''.join(template.itertext())
document = WORKBOOK / '{yyyy}-{mm}-{dd} №{OrderNum} ({Signatory}) (Александр).docx'
with ZipFile(document) as archive:
    xml = ET.fromstring(archive.read('word/document.xml'))
    text = ''.join(xml.itertext())
    w = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    active = None
    root_index = None
    leave_index = None
    for index, paragraph in enumerate(xml.findall('w:body/w:p', w)):
        paragraph_text = ''.join(paragraph.itertext())
        if paragraph_text.startswith('{\\block:'):
            active = None if paragraph_text.endswith('_End}') else paragraph_text
        level = paragraph.find('w:pPr/w:numPr/w:ilvl', w)
        if level is not None and level.get('{' + w['w'] + '}val') == '0':
            assert active is None, ('Root heading inside optional block', active)
        if paragraph_text.startswith('Вважати такими, що вибули зі складу сил 2 корпусу') and 'продовольчого' not in paragraph_text:
            root_index = index
        if paragraph_text == '{\\block:' + MODE + '_Begin}':
            leave_index = index
    assert root_index is not None and leave_index > root_index
    for kind in ('block', 'export'):
        for boundary in ('Begin', 'End'):
            assert text.count('{\\' + kind + ':' + MODE + '_' + boundary + '}') == 1
print('Medical company leave: both configurations, routes, Word template and DOCX anchors passed.')