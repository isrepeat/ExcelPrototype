"""Validate the single-source medical-unit catalog and shared query mapping offline."""

from pathlib import Path
import re
import openpyxl

root = Path(__file__).resolve().parents[3] / 'PROJECTS_EXCEL'
config = {}
for line in (root / 'config/cfgDependencies/wsConfig.txt').read_text(encoding='utf-8-sig').splitlines():
    kind, key, value = line.split('\t')
    config[key] = value

def cfg(key):
    return config['common::medicalCatalog.' + key]

book = openpyxl.load_workbook(root / 'workbook/Dependencies/Установи.xlsx', data_only=True)

def table_rows(name):
    matches = [(sheet, table) for sheet in book for table in sheet.tables.values() if table.name == name]
    assert len(matches) == 1
    sheet, table = matches[0]
    cells = [[cell.value for cell in row] for row in sheet[table.ref]]
    return [dict(zip(cells[0], row)) for row in cells[1:]]

institutions = table_rows(cfg('tables.institutions'))
units = table_rows(cfg('tables.units'))
headers = cfg('headers').split('|')
unit_headers = cfg('unitHeaders').split('|')
assert len(headers) == len(unit_headers)
medical = [row for row in units if row[cfg('flag.header')] == cfg('flag.value')]
assert medical and institutions
assert all(not re.fullmatch(r'[АA]\d{4}', str(row['Позначення'])) for row in institutions)
normalize = lambda value: str(value).upper().replace('A', 'А')
codes = [normalize(row['Військова частина']) for row in units]
assert codes == sorted(codes) and len(codes) == len(set(codes))
catalog = [{key: row[key] for key in headers} for row in institutions]
catalog += [dict(zip(headers, [row[key] for key in unit_headers])) for row in medical]
for row in medical:
    for field in ['Назва', 'Родовий', 'Знахідний', 'Давальний']:
        assert row[field], (row['Військова частина'], field)
    matches = [result for result in catalog if normalize(result['Позначення']) == normalize(row['Військова частина'])]
    assert len(matches) == 1 and matches[0]['Родовий'] == row['Родовий']
assert not any(normalize(row['Позначення']) == normalize('А7383') for row in catalog)
assert config['source.Institutions.range'] == '@MedicalInstitutions'
service = (root / 'vba/common/[2] classes/obj_TableQueryService.cls.vba').read_text(encoding='utf-8-sig')
assert service.index('ex_TableSnapshots.fn_TryQuery') < service.index('ex_MedicalInstitutions.fn_TryQuery')
assert 'fn_TryApply(combined, query, output, diagnostic)' in (root / 'vba/common/modules/ex_MedicalInstitutions.vba').read_text(encoding='utf-8-sig')
print(f'Medical catalog: {len(institutions)} independent institutions + {len(medical)} flagged units; mappings, unique keys and declensions: OK')