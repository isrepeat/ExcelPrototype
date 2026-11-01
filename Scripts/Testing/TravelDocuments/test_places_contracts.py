"""Check the saved places source, address projection and UI route without VBA."""

from pathlib import Path
import xml.etree.ElementTree as ET
import openpyxl


root = Path(__file__).resolve().parents[3] / 'PROJECTS_EXCEL'
config = {}
for line in (root / 'config/TravelDocumentsBuilder/wsConfig.txt').read_text(encoding='utf-8-sig').splitlines():
    kind, key, value = line.split('\t')
    assert key not in config
    config[key] = value


def cfg(key):
    return config['TravelDocumentsBuilder::lookup.Places.' + key]


path = (root / 'workbook/TravelDocumentsBuilder' / cfg('source.path')).resolve()
book = openpyxl.load_workbook(path, data_only=True)
sheet = book[cfg('source.sheet')]
table = sheet.tables[cfg('source.table')]
headers = [column.name for column in table.tableColumns]
assert headers == [cfg('columns.' + field) for field in cfg('sourceFields').split('|')]
assert cfg('fields').split('|') == ['Address', 'Region', 'Type', 'Name', 'District']
spans = [int(value) for value in cfg('columnSpans').split('|')]
assert spans == [2, 1, 1, 1, 1]
assert int(cfg('columnOffset')) == 1
assert sum(spans) + int(cfg('columnOffset')) <= int(config['TravelDocumentsBuilder::candidates.layout.columns'])
workspace = int(config['TravelDocumentsBuilder::candidates.layout.columns'])
assert int(cfg('columnOffset')) + sum(spans) == 7
assert 2 + sum(spans) == workspace
assert 3 + sum(spans) > workspace
controller = (root / 'vba/TravelDocumentsBuilder/obj_TDB_PgMainController.cls.vba').read_text(encoding='utf-8-sig')
assert 'columnOffset + rawTable.ColumnCount > maxColumns' in controller
assert 'ex_WindowsUi.fn_ShowMessage message, VBA.vbExclamation' in controller
place_fields = {'PlaceRegion', 'PlaceType', 'PlaceName', 'PlaceDistrict'}
assert place_fields <= set(config['TravelDocumentsBuilder::history.columns'].split('|'))
assert place_fields <= set(config['TravelDocumentsBuilder::history.snapshotFields'].split('|'))
assert {'Snapshot.' + field for field in place_fields} <= set(config['TravelDocumentsBuilder::history.versions.columns'].split('|'))
assert {field + '=' + field for field in place_fields} <= set(config['TravelDocumentsBuilder::history.formFields'].split('|'))
generator = (root / 'vba/TravelDocumentsBuilder/obj_TDB_VacationGenerator.cls.vba').read_text(encoding='utf-8-sig')
assert 'For Each key In ex_TDB_Places.fn_FormFields()' in generator
assert 'metadata.Add VBA.CStr(key), m_form(key)' in generator
rows = list(sheet[table.ref])[1:]
addresses = []
for row in rows:
    values = dict(zip(cfg('sourceFields').split('|'), [cell.value for cell in row]))
    if 'без' in str(values['Name']).casefold():
        addresses.append(cfg('addressTemplate').format(**values))
assert 'сел. Безлюдівка, Харківська обл., Харківський р-н' in addresses
assert 'сел. Безпалівка, Харківська обл., Чугуївський р-н' in addresses
layout = ET.parse(root / 'ui/TravelDocumentsBuilder/MainPage.xaml')
field = next(node for node in layout.iter() if node.attrib.get('name') == 'DocumentTarget')
assert field.attrib['style'] == 'lookupKeyField'
assert field.attrib['onChange'] == '{Binding Path=Commands.PlaceChanged}'
loader = (root / 'vba/TravelDocumentsBuilder/obj_TDB_DependencyLoader.cls.vba').read_text(encoding='utf-8-sig')
assert 'For Each sourceMode In VBA.Array(False, True)' in loader
assert 'Set source = ex_TDB_Places.fn_Source(VBA.CBool(sourceMode))' in loader
print('Places table schema, address examples and UI route: OK')