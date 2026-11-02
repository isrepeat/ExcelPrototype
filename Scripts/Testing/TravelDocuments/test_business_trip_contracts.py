"""Validate business-trip UI, history and mixed-order export contracts without VBA."""

from pathlib import Path
import re
import xml.etree.ElementTree as ET


root = Path(__file__).resolve().parents[3] / 'PROJECTS_EXCEL'
source = root / 'vba/TravelDocumentsBuilder'
config = {}
for line in (root / 'config/TravelDocumentsBuilder/wsConfig.txt').read_text(encoding='utf-8-sig').splitlines():
    kind, key, value = line.split('\t')
    assert key not in config, key
    config[key] = value


def cfg(key):
    return config['TravelDocumentsBuilder::' + key]


ids = cfg('documentTypes.ids').split(',')
assert ids == ['VacationTicket', 'BusinessTripCertificate']
assert cfg('documentTypes.items').split(',') == [cfg('text.' + mode + '.DocumentFormTitle') for mode in ids]
for mode in ids:
    for caption in ['DocumentFormTitle', 'DocumentTargetLabel', 'DepartureDateLabel', 'ReturnDateLabel', 'PurposeLabel', 'VacationDays']:
        assert cfg('text.' + mode + '.' + caption)
    catalog = ET.parse(root / 'workbook/TravelDocumentsBuilder' / cfg('generation.tablePaths.' + mode))
    ns = {'p': 'urn:excelprototype:profiles'}
    assert catalog.find('p:template[@name="' + cfg('generation.tableNames.' + mode) + '"]', ns) is not None

columns = cfg('history.columns').split('|')
snapshots = cfg('history.snapshotFields').split('|')
versions = cfg('history.versions.columns').split('|')
mappings = dict(pair.split('=') for pair in cfg('history.formFields').split('|'))
assert len(columns) == len(set(columns)) == 53
for field in ['WeaponLine']:
    assert field in columns and field in snapshots
    assert 'Snapshot.' + field in versions and mappings[field] == field

page = ET.parse(root / 'ui/TravelDocumentsBuilder/MainPage.xaml')
nodes = {node.attrib['name']: node for node in page.iter() if 'name' in node.attrib}
for field in ['WeaponLine']:
    assert nodes[field].attrib['value'] == '{Binding Path=Form.' + field + '}'
    assert nodes[field + 'Label'].attrib['text'] == '{Binding Path=Text.' + field + '}'
    assert cfg('text.' + field)
assert nodes['TripFields'].attrib['visibility'] == '{Binding Path=Form.TripVisibility}'
assert 'TripUnit' not in nodes and 'TripUnitLabel' not in nodes
assert 'TripPurpose' not in nodes and 'TripPurposeLabel' not in nodes
assert 'TripPurpose' not in columns and 'Snapshot.TripPurpose' not in versions
assert 'TripUnit' not in columns and 'Snapshot.TripUnit' not in versions
assert all(len(name) <= 31 for name in nodes)

generator = (source / 'obj_TDB_VacationGenerator.cls.vba').read_text(encoding='utf-8-sig')
controller = (source / 'obj_TDB_PgMainController.cls.vba').read_text(encoding='utf-8-sig')
history = (source / 'ex_TDB_History.vba').read_text(encoding='utf-8-sig')
exporter = (source / 'obj_TDB_OrderWordExport.cls.vba').read_text(encoding='utf-8-sig')
mode_change = re.search(r'Public Function DocumentTypeChangedHandler\([\s\S]*?^End Function', controller, re.M)[0]
assert 'generation.businessTrip.kind' in mode_change
assert '"DonationDays", "0"' in mode_change
assert '"VlkNo", "VlkDate", "LongTreatmentDays", "TreatmentLeaveDays"' in mode_change
assert 'updates("Form.DocumentType")' in history and 'fn_DocumentId' in history
load = re.search(r'Public Function SelectHistoryHandler\([\s\S]*?^End Function', controller, re.M)[0]
assert '"text.VacationTicket.DocumentFormTitle"' not in load
assert 'm_wordValues("Basis") = m_form("Purpose")' in generator
assert 'm_wordValues("TravelPlace") = m_form("DocumentTarget")' in generator
assert 'MilitaryPlace' not in generator and 'TripUnit' not in generator
assert cfg('generation.businessTrip.indefinite').replace('{from}', '01.11.2026') == 'з 01.11.2026 до окремого розпорядження.'
assert '"{from}", VBA.Format$(departure, "dd.mm.yyyy"))' in generator
assert cfg('generation.businessTrip.indefiniteInput') == 'до о/р'
assert '{phrase}' in cfg('generation.businessTrip.durationError')
assert 'fields(private_Header(9)) = ex_TravelDocumentsBuilder.fn_Config("generation.businessTrip.indefiniteInput")' in generator
assert 'private_Days = ex_TravelDocumentsBuilder.fn_ParseDuration(m_documentId, text)' in generator
assert 'VacationDatesChanged' not in controller
for name in ['DepartureDate', 'VacationDays', 'RoadDays', 'DonationDays']:
    assert 'onChange' not in nodes[name].attrib
duration_source = (source / 'ex_TravelDocumentsBuilder.vba').read_text(encoding='utf-8-sig')
assert 'If VBA.Len(value) = 0 Or value Like "*[!0-9]*" Then' in duration_source
assert 'If VBA.CDbl(value) <= 0 Or VBA.CDbl(value) > VBA.CDbl(maximum) Then' in duration_source
assert 'private_IsIndefinite = m_documentId = "BusinessTripCertificate" And days = 0' in generator
assert 'If m_documentId = "VacationTicket" And days = 0 Then' in generator
assert generator.count('private_ValidateDuration vacationDays') == 2
assert generator.count('vacationDays = private_Days("VacationDays", False)') == 2
assert 'fields(private_Header(11)) = VBA.vbNullString' in generator
assert 'VBA.CStr(root("PlannedReturnDate"))' in generator
assert 'If dateValue = 0 Then' in generator
assert 'm_wordValues("WeaponLine") = m_form("WeaponLine")' in generator
assert 'm_wordTables.Exists(m_documentId)' in generator
assert 'Set m_wordAnchors = Nothing' in generator
assert 'frozenTemplates(documentId)' in exporter
assert 'templateStamps(path)' in exporter
assert 'word.filename.prefix.VacationTicket' not in exporter
assert exporter.index('Set templatePaths =') < exporter.index('For Each file In m_fileSystem.GetFolder(folder).Files')
for key in re.findall(r'ex_TravelDocumentsBuilder.fn_Config\("([^"]+)"\)', generator + controller + exporter):
    assert 'TravelDocumentsBuilder::' + key in config, key

print('Business-trip modes, labels, schema, saved fields and mixed-order export contracts: OK')