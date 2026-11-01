"""Check foreign lookup, nullable source columns and history contracts without VBA."""

from pathlib import Path
import re
import xml.etree.ElementTree as ET
import openpyxl


root = Path(__file__).resolve().parents[3] / 'PROJECTS_EXCEL'
sources = root / 'vba/TravelDocumentsBuilder'
config = {}
for line in (root / 'config/TravelDocumentsBuilder/wsConfig.txt').read_text(encoding='utf-8-sig').splitlines():
    kind, key, value = line.split('\t')
    assert key not in config, key
    config[key] = value


def cfg(key):
    return config['TravelDocumentsBuilder::' + key]


prefix = 'lookup.PlacesAbroad.'
book = openpyxl.load_workbook(
    (root / 'workbook/TravelDocumentsBuilder' / cfg(prefix + 'source.path')).resolve(),
    data_only=True,
)
sheet = book[cfg(prefix + 'source.sheet')]
table = sheet.tables[cfg(prefix + 'source.table')]
headers = [column.name for column in table.tableColumns]
aliases = cfg(prefix + 'sourceFields').split('|')
assert len(aliases) == 7
assert {cfg(prefix + 'columns.' + alias) for alias in aliases} <= set(headers)
assert 'Address' not in aliases
rows = [dict(zip(headers, [cell.value for cell in row])) for row in list(sheet[table.ref])[1:]]
assert rows
required = cfg(prefix + 'requiredFields').split('|')
mapping = dict(pair.split('=') for pair in cfg(prefix + 'formFields').split('|'))
for row in rows:
    for alias, target in mapping.items():
        if target in required:
            assert row[cfg(prefix + 'columns.' + alias)], (target, row)
# The actual catalog exercises country-only, region-only and nullable type/region cases.
assert any(row['Тип'] == 'держава' and not row['Регіон (UA)'] for row in rows)
assert any(row['Тип'] == 'регіон' and row['Місце (UA)'] == row['Регіон (UA)'] for row in rows)
assert any(not row['Тип'] for row in rows)
assert any(row['Місце (UA)'] == 'Лінц' and row['Місце (EN)'] == 'Linz' for row in rows)
book.close()
assert sum(map(int, cfg(prefix + 'columnSpans').split('|'))) + int(cfg(prefix + 'columnOffset')) <= int(cfg('candidates.layout.columns'))

places = (sources / 'ex_TDB_Places.vba').read_text(encoding='utf-8-sig')
controller = (sources / 'obj_TDB_PgMainController.cls.vba').read_text(encoding='utf-8-sig')
page = (sources / 'obj_TDB_PgMain.cls.vba').read_text(encoding='utf-8-sig')
chain = (sources / 'ex_TDB_PlaceChain.vba').read_text(encoding='utf-8-sig')
generator = (sources / 'obj_TDB_VacationGenerator.cls.vba').read_text(encoding='utf-8-sig')
loader = (sources / 'obj_TDB_DependencyLoader.cls.vba').read_text(encoding='utf-8-sig')
for key in re.findall(r'ex_TravelDocumentsBuilder.fn_Config\("([^"]+)"\)', places + chain + controller):
    assert 'TravelDocumentsBuilder::' + key in config, key
assert 'For Each sourceMode In VBA.Array(False, True)' in loader
assert 'Set tableSource = fn_Source(isAbroad)' in places
assert 'If VBA.IsNull(value) Or VBA.IsEmpty(value) Then value = VBA.vbNullString' in places
assert 'fields("Region") <> fields("Name")' in places
assert 'fields("RegionEn") <> fields("NameEn")' in places
assert 'ua = VBA.CStr(fields("Country"))' in places
assert 'en = VBA.CStr(fields("CountryEn"))' in places
assert 'address = fn_Address(fields, isAbroad)' in places
assert 'ex_TDB_Places.fn_Search(VBA.CStr(query), isAbroad)' in controller
change = re.search(r'Public Function VacationAbroadChangedHandler\([\s\S]*?^End Function', controller, re.M)[0]
assert 'Me.PlaceChangedHandler()' in change
assert 'SetObject("Data", "PlaceChain"' not in change
assert 'SetObject("Data", "PlaceChain"' not in controller[controller.index('Public Function PlaceChangedHandler'):controller.index('Public Function SelectPersonnelHandler')]

fields = set(cfg('placeChain.fields').split('|'))
assert fields <= set(cfg('history.columns').split('|'))
assert fields <= set(cfg('history.snapshotFields').split('|'))
assert {'Snapshot.' + field for field in fields} <= set(cfg('history.versions.columns').split('|'))
assert {field + '=' + field for field in fields} <= set(cfg('history.formFields').split('|'))
assert 'ex_TDB_Places.fn_FormFields()' in generator
assert 'row("PlaceAbroad") = values("VacationAbroad")' in chain
assert 'address = ex_TDB_Places.fn_RowAddress(row)' in chain
assert 'ElseIf isAbroad Then\n        row("Destination") = address' in chain
assert 'If Not isAbroad Then' in chain  # Ukrainian type constraints do not apply abroad.
assert 'If ex_TDB_Places.fn_IsAbroad(VBA.CStr(first("PlaceAbroad"))) Then' in generator

layout = ET.parse(root / 'ui/TravelDocumentsBuilder/MainPage.xaml')
nodes = {node.attrib['name']: node for node in layout.iter() if 'name' in node.attrib}
assert nodes['VacationAbroad'].attrib['onChange'] == '{Binding Path=Commands.VacationAbroadChanged}'
assert '"VacationAbroadChanged", "VacationAbroadChangedHandler"' in page
assert nodes['ForeignPlaceFields'].attrib['visibility'] == '{Binding Path=Form.ForeignPlaceFieldsVisibility}'
for field in ('PlaceCountry', 'PlaceCountryEn', 'PlaceRegionEn', 'PlaceNameEn'):
    assert nodes[field].attrib['value'] == '{Binding Path=Form.' + field + '}'
    assert cfg('text.' + field)
for node in layout.iter():
    name = node.attrib.get('name', '')
    assert len(name) <= 31, name
    if node.tag.endswith(('button', 'select')):
        assert len('btn_' + name) <= 31, name
print('Foreign source schema, nullable values, switching, manual input and history contracts: OK')