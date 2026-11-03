"""Check reference integrity, document examples, history preservation and UI contracts."""

import argparse
import importlib.util
from pathlib import Path
import re
import xml.etree.ElementTree as ET

from BasisConfigFixtures import configured_rows, enrich


ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('upgrade', Path(__file__).with_name('Upgrade-TDBBasisWorkbooks.py'))
upgrade = importlib.util.module_from_spec(spec)
spec.loader.exec_module(upgrade)
parser = argparse.ArgumentParser()
parser.add_argument('--institutions', type=Path, default=ROOT / 'workbook/Dependencies/Установи.xlsx')
parser.add_argument('--history', type=Path, default=ROOT / 'workbook/TravelDocumentsBuilder/TicketsHistory.xlsx')
parser.add_argument('--original-institutions', type=Path)
parser.add_argument('--original-history', type=Path)
args = parser.parse_args()
book = upgrade.Package(args.institutions)
for _, table in book.tables().values():
    first, last = table.get('ref').split(':')
    assert upgrade.index(last) - upgrade.index(first) + 1 == len(table.find(f'{{{upgrade.S}}}tableColumns')), table.get('name')
assert re.search(rb'<Types(?:\s|>)', book.parts['[Content_Types].xml']), 'OPC content types must use their default namespace'
assert re.search(rb'<Relationships(?:\s|>)', book.parts['xl/_rels/workbook.xml.rels']), 'OPC relationships must use their default namespace'
config = {}
for line in (ROOT / 'config/TravelDocumentsBuilder/wsConfig.txt').read_text(encoding='utf-8-sig').splitlines():
    if not line:
        continue
    fields = line.split('\t')
    assert len(fields) == 3
    assert fields[1] not in config, fields[1]
    config[fields[1]] = fields[2]


def cfg(key):
    return config['TravelDocumentsBuilder::' + key]


def records(package, table_name):
    _, table = package.tables()[table_name]
    start, end = table.get('ref').split(':')
    first_row = int(''.join(filter(str.isdigit, start)))
    last_row = int(''.join(filter(str.isdigit, end)))
    first_column = upgrade.index(start)
    headers = [node.get('name') for node in table.find(f'{{{upgrade.S}}}tableColumns')]
    for sheet_path in package.sheets().values():
        rels = package.parts.get('xl/worksheets/_rels/' + Path(sheet_path).name + '.rels', b'')
        table_path = package.tables()[table_name][0]
        if Path(table_path).name.encode() not in rels:
            continue
        sheet = upgrade.parse(package.parts[sheet_path])
        result = []
        for row in sheet.find(f'{{{upgrade.S}}}sheetData'):
            if not first_row < int(row.get('r')) <= last_row:
                continue
            cells = {upgrade.index(cell.get('r')): cell for cell in row}
            values = {header: package.text(cells.get(first_column + index)) for index, header in enumerate(headers)}
            if any(values.values()):
                result.append(values)
        return result
    raise AssertionError(table_name)


model_config = {key.removeprefix('TravelDocumentsBuilder::'): value for key, value in config.items()}
officials = {row['Id']: row for row in enrich(model_config, 'Officials', records(book, 'tbBasisOfficials'))}
documents = {row['Id']: row for row in enrich(model_config, 'Documents', records(book, 'tbBasisDocumentTypes'))}
official_rules = {(row['Група установи'], row['Id посадової особи']) for row in configured_rows(model_config, 'OfficialRules')}
document_rules = {(row['Група установи'], row['Id документа']) for row in configured_rows(model_config, 'DocumentRules')}
assert len(officials) == len(records(book, 'tbBasisOfficials')), 'Duplicate official ID'
assert len(documents) == len(records(book, 'tbBasisDocumentTypes')), 'Duplicate document ID'
assert len({row['Назва'].casefold() for row in officials.values()}) == len(officials)
assert len({row['Назва'].casefold() for row in documents.values()}) == len(documents)
assert all(identifier in officials for _, identifier in official_rules)
assert all(identifier in documents for _, identifier in document_rules)
assert ('Unit', 'commander') in official_rules
assert ('Hospital', 'commander') not in official_rules
assert ('Hospital', 'order') not in document_rules
assert ('Hospital', 'instruction') not in document_rules
assert ('Hospital', 'letter') in document_rules
assert ('ChiefCommand', 'commander_in_chief') in official_rules
assert ('GeneralStaff', 'commander_in_chief') not in official_rules
assert ('Court', 'court_order') in document_rules
overrides = dict(entry.split('~') for entry in cfg('basisChain.model.officialGroupOverrides').split('|'))
assert ('Unit', 'commander') in official_rules
assert (overrides.get('Unit', 'Hospital'), 'commander') in official_rules
assert (overrides.get('Hospital', 'Hospital'), 'chief') in official_rules
assert ('Hospital', 'order') not in document_rules
editor = set(cfg('basisChain.editorFields').split('|'))
persisted = set(cfg('basisChain.fields').split('|'))
dictionary = {row['Код']: row for row in configured_rows(model_config, 'Fields')}
bindings = {code: target.removeprefix('Form.') for code, target in (entry.split('=') for entry in cfg('basisChain.model.bindings').split('|'))}
for document in documents.values():
    visible = set(document['Поля'].split('|'))
    required = set(filter(None, re.split(r'[|&/()\s]+', document['Обов’язкові поля'])))
    assert {bindings[field] for field in visible} <= editor
    assert required <= visible
    assert 'Шаблон' not in document
    assert document['Родовий']
for source in ['Officials', 'Documents']:
    prefix = 'basisChain.references.' + source
    headers = {node.get('name') for node in book.tables()[cfg(prefix + '.source.table')][1].find(f'{{{upgrade.S}}}tableColumns')}
    assert prefix + '.source' in cfg('snapshots.sources').split('|')
    for field in cfg(prefix + '.fields').split('|'):
        assert cfg(prefix + '.columns.' + field) in headers


history = upgrade.Package(args.history)
for table_name, key in [('tbTickets', 'history.columns'), ('tbTicketRevisions', 'history.versions.columns')]:
    columns = [node.get('name') for node in history.tables()[table_name][1].find(f'{{{upgrade.S}}}tableColumns')]
    assert len(columns) == len(set(columns))
    assert set(cfg(key).split('|')) <= set(columns)
assert persisted <= set(cfg('history.columns').split('|'))
assert persisted <= set(cfg('history.snapshotFields').split('|'))
assert {'Snapshot.' + field for field in persisted} <= set(cfg('history.versions.columns').split('|'))
assert 'FoodDepartureDate' not in cfg('history.columns')
if args.original_history:
    original = upgrade.Package(args.original_history)
    for table_name in ['tbTickets', 'tbTicketRevisions']:
        before = records(original, table_name)
        after = records(history, table_name)
        assert len(before) == len(after)
        for old, new in zip(before, after):
            assert all(new[key] == value for key, value in old.items()), 'Historical values changed'
if args.original_institutions:
    original = upgrade.Package(args.original_institutions)
    for table_name in ['ЛЗК', 'tbMilitaries', 'tbPlacesUkraine', 'tbPlacesAbroad']:
        assert records(original, table_name) == records(book, table_name)
    for name in original.sheets().values():
        if name.endswith(('sheet1.xml', 'sheet2.xml', 'sheet3.xml', 'sheet4.xml')):
            assert original.parts[name] == book.parts[name], 'Unrelated worksheet changed'

page = ET.parse(ROOT / 'ui/TravelDocumentsBuilder/MainPage.xaml')
nodes = {node.get('name'): node for node in page.iter() if node.get('name')}
ordered_names = [node.get('name') for node in page.iter() if node.get('name')]
assert ordered_names.index('BasisIssuer') < ordered_names.index('BasisActing') < ordered_names.index('BasisNumber') < ordered_names.index('BasisDate')
for field in editor:
    assert nodes[field].get('value') == '{Binding Path=Form.' + field + '}'
    assert nodes[field].get('onChange') == '{Binding Path=Commands.BasisDraftChanged}'
    assert any(node.get('visibility') == '{Binding Path=Form.' + field + 'Visibility}' for node in page.iter())
assert nodes['BasisType'].get('onChange') == '{Binding Path=Commands.BasisTypeChanged}'
assert nodes['BasisPreview'].get('readOnly') == 'true'
assert len({node.get('name') for node in page.iter() if node.get('name')}) == sum(bool(node.get('name')) for node in page.iter())
source = (ROOT / 'vba/TravelDocumentsBuilder/ex_TDB_BasisChain.vba').read_text(encoding='utf-8-sig')
assert 'row("BasisRenderedText") = ex_TDB_BasisConfig.fn_Render' in source
assert 'fn_Format = VBA.CStr(row("BasisRenderedText"))' in source
assert 'private_Validate copy' in source, 'Read validation must not mutate persisted rows'
assert 'basisChain.roles.' not in source
assert 'basisChain.items.Types.' not in source
assert 'basisChain.issuerForms.' not in source
for text in [source, (ROOT / 'vba/TravelDocumentsBuilder/obj_TDB_PgMainController.cls.vba').read_text(encoding='utf-8-sig')]:
    assert max(map(len, text.splitlines())) < 1024
print('Reference integrity, example phrases, dependent UI, frozen text and history preservation: OK')