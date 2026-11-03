"""Check the editable enumeration and blocking Excel validation without opening Excel."""

import argparse
import importlib.util
from pathlib import Path
import re


ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('upgrade', Path(__file__).with_name('Upgrade-TDBBasisWorkbooks.py'))
upgrade = importlib.util.module_from_spec(spec)
spec.loader.exec_module(upgrade)
parser = argparse.ArgumentParser()
parser.add_argument('--institutions', type=Path, default=ROOT / 'workbook/Dependencies/Установи.xlsx')
parser.add_argument('--original', type=Path)
args = parser.parse_args()
book = upgrade.Package(args.institutions)
assert 'Переліки' in book.sheets()
_, table = book.tables()['tbEnumInstitutionGroups']
assert [node.get('name') for node in table.find(f'{{{upgrade.S}}}tableColumns')] == ['Група установи', 'Назва']
sheet = upgrade.parse(book.parts[book.sheets()['Переліки']])
start, end = table.get('ref').split(':')
first = int(''.join(filter(str.isdigit, start)))
last = int(''.join(filter(str.isdigit, end)))
allowed = []
for row in sheet.find(f'{{{upgrade.S}}}sheetData'):
    if first < int(row.get('r')) <= last:
        cell = next(cell for cell in row if upgrade.index(cell.get('r')) == upgrade.index(start))
        allowed.append(book.text(cell))
assert all(allowed) and len(allowed) == len(set(allowed))
assert set(allowed) == {'Unit', 'Hospital', 'TCC', 'ChiefCommand', 'GeneralStaff', 'LandCommand', 'Court'}
validated = []
for name, path in book.sheets().items():
    if name == 'Переліки':
        continue
    root = upgrade.parse(book.parts[path])
    cells = root.findall('.//s:c', upgrade.NS)
    headers = [cell for cell in cells if book.text(cell) == 'Група установи']
    for header in headers:
        address = re.match('[A-Z]+', header.get('r'))[0]
        first_data = int(''.join(filter(str.isdigit, header.get('r')))) + 1
        target = f'{address}{first_data}:{address}1048576'
        rules = [rule for rule in root.findall('s:dataValidations/s:dataValidation', upgrade.NS) if target in rule.get('sqref', '').split()]
        assert len(rules) == 1, name
        rule = rules[0]
        assert rule.get('type') == 'list'
        assert rule.get('errorStyle', 'stop') == 'stop' and rule.get('showErrorMessage') == '1'
        assert rule.get('showDropDown', '0') == '0' and rule.get('allowBlank', '0') == '0'
        assert rule.find(f'{{{upgrade.S}}}formula1').text == 'INDIRECT("tbEnumInstitutionGroups[Група установи]")'
        for cell in cells:
            if upgrade.index(cell.get('r')) == upgrade.index(header.get('r')) and int(''.join(filter(str.isdigit, cell.get('r')))) >= first_data:
                assert not book.text(cell) or book.text(cell) in allowed
        validated.append(name)
assert len(validated) == 4, validated
if args.original:
    original = upgrade.Package(args.original)
    for path in original.sheets().values():
        old = upgrade.parse(original.parts[path])
        new = upgrade.parse(book.parts[path])
        a = old.find(f'{{{upgrade.S}}}sheetData')
        b = new.find(f'{{{upgrade.S}}}sheetData')
        assert upgrade.ET.tostring(a) == upgrade.ET.tostring(b), path
        old_rules = old.findall('s:dataValidations/s:dataValidation', upgrade.NS)
        new_rules = new.findall('s:dataValidations/s:dataValidation', upgrade.NS)
        for rule in old_rules:
            assert any(upgrade.ET.tostring(rule) == upgrade.ET.tostring(other) for other in new_rules), path
    assert original.parts['xl/styles.xml'] == book.parts['xl/styles.xml']
print(f'Enumeration: {len(allowed)} groups; blocking validation: {len(validated)} sheets; original data and styles preserved')