"""Check v1-to-v2 workbook migration and saved-text compatibility without Excel."""

import argparse
from datetime import datetime, timedelta
import importlib.util
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('upgrade', Path(__file__).with_name('Upgrade-TDBBasisWorkbooks.py'))
upgrade = importlib.util.module_from_spec(spec)
spec.loader.exec_module(upgrade)
parser = argparse.ArgumentParser()
parser.add_argument('--original', type=Path, required=True)
parser.add_argument('--updated', type=Path, required=True)
args = parser.parse_args()
before = upgrade.Package(args.original)
after = upgrade.Package(args.updated)
config = {parts[1].removeprefix('TravelDocumentsBuilder::'): parts[2]
          for line in (ROOT / 'config/TravelDocumentsBuilder/wsConfig.txt').read_text(encoding='utf-8-sig').splitlines()
          if line for parts in [line.split('\t')]}


def records(book, name):
    path, table = book.tables()[name]
    headers = [node.get('name') for node in table.find(f'{{{upgrade.S}}}tableColumns')]
    start, end = table.get('ref').split(':')
    first = int(''.join(filter(str.isdigit, start)))
    last = int(''.join(filter(str.isdigit, end)))
    for sheet_path in book.sheets().values():
        rels = book.parts.get('xl/worksheets/_rels/' + Path(sheet_path).name + '.rels', b'')
        if Path(path).name.encode() not in rels:
            continue
        result = []
        for row in upgrade.parse(book.parts[sheet_path]).find(f'{{{upgrade.S}}}sheetData'):
            if first < int(row.get('r')) <= last:
                cells = {upgrade.index(cell.get('r')): cell for cell in row}
                data = {header: book.text(cells.get(upgrade.index(start) + offset)) for offset, header in enumerate(headers)}
                if any(data.values()):
                    result.append(data)
        return result
    raise AssertionError(name)


assert set(after.tables()) == set(before.tables()) - {'tbBasisFragments', 'tbBasisFields', 'tbBasisContract', 'tbBasisDocumentRules', 'tbBasisOfficialRules'}
assert 'Фрагменти тексту' not in after.sheets()
assert before.parts['xl/styles.xml'] == after.parts['xl/styles.xml']
for name in after.tables():
    if name == 'tbBasisContract':
        continue
    old = records(before, name)
    new = records(after, name)
    assert len(old) == len(new), name
    for a, b in zip(old, new):
        if name in {'tbBasisDocumentTypes', 'tbBasisOfficials'}:
            a = {key: a[key] for key in ['Назва', 'Родовий']}
        assert a == b, name
assert 'Контракт' not in after.sheets() and 'Поля документів' not in after.sheets()

history = upgrade.Package(ROOT / 'workbook/TravelDocumentsBuilder/TicketsHistory.xlsx')
saved = records(history, 'tbTickets')
officials = {row['Назва'].casefold(): row[config['basisChain.references.Officials.columns.Text']]
             for row in records(after, 'tbBasisOfficials')}
date_fields = set(config['history.dateFields'].split('|'))


def normalized(row):
    row = dict(row)
    for field in date_fields:
        value = row.get(field, '')
        if value.replace('.', '', 1).isdigit() and float(value) >= 2:
            row[field] = (datetime(1899, 12, 30) + timedelta(days=float(value))).strftime('%d.%m.%Y')
    return row


def legacy_text(row):
    row = normalized(row)
    if row.get('BasisRenderedText'):
        return row['BasisRenderedText']
    text = row.get('BasisTemplate') or config['basisChain.documentTemplate']
    for field in config['basisChain.fields'].split('|'):
        text = text.replace('{Value.' + field + '}', row.get(field, ''))
    for field in config['basisChain.fields'].split('|'):
        value = row.get(field, '')
        if field in {'BasisIssuer', 'BasisIssuerText'}:
            value = row.get('BasisIssuerText') or officials.get(row.get('BasisIssuer', '').casefold(), '')
        text = text.replace('{' + field + '}', value)
    assert '{' not in text and '}' not in text, 'Unexpected legacy formatting branch'
    return ' '.join(text.split())


roots = 0
for event in saved:
    if event.get('RowType') != 'Event' or event.get('VacationKind') != config['generation.businessTrip.kind']:
        continue
    chain = []
    parent = event['RowId']
    while True:
        children = [row for row in saved if row.get('RowType') == 'Basis' and row.get('ParentRowId') == parent]
        assert len(children) <= 1
        if not children:
            break
        chain.append(children[0])
        parent = children[0]['RowId']
    ending = ''.join(config['basisChain.separator'] + ' ' + legacy_text(row) for row in chain) + config['basisChain.terminator']
    assert event['Basis'].endswith(ending) and len(event['Basis']) > len(ending), 'Saved basis/child mismatch'
    roots += 1
print(f'Migration preserves all retained table values and styles; {roots} saved business-trip basis chains remain readable')