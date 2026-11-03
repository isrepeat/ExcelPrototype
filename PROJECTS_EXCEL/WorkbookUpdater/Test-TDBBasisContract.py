"""Validate the public contract, negative fixtures and authored text without Excel/VBA."""

import argparse
import copy
import importlib.util
from pathlib import Path
import re


ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('upgrade', Path(__file__).with_name('Upgrade-TDBBasisWorkbooks.py'))
upgrade = importlib.util.module_from_spec(spec)
spec.loader.exec_module(upgrade)
parser = argparse.ArgumentParser()
parser.add_argument('--institutions', type=Path, default=ROOT / 'workbook/Dependencies/Установи.xlsx')
args = parser.parse_args()
book = upgrade.Package(args.institutions)
config = {parts[1].removeprefix('TravelDocumentsBuilder::'): parts[2]
          for line in (ROOT / 'config/TravelDocumentsBuilder/wsConfig.txt').read_text(encoding='utf-8-sig').splitlines()
          if line for parts in [line.split('\t')]}
bindings = {code: target.removeprefix('Form.') for code, target in
            (entry.split('=') for entry in config['basisChain.contract.bindings'].split('|'))}
binding_types = {target.removeprefix('Form.'): kind for target, kind in
                 (entry.split('=') for entry in config['basisChain.contract.bindingTypes'].split('|'))}


def rows(name, sheet_name):
    _, table = book.tables()[name]
    headers = [node.get('name') for node in table.find(f'{{{upgrade.S}}}tableColumns')]
    start, end = table.get('ref').split(':')
    first = int(''.join(filter(str.isdigit, start)))
    last = int(''.join(filter(str.isdigit, end)))
    result = []
    sheet = upgrade.parse(book.parts[book.sheets()[sheet_name]])
    for row in sheet.find(f'{{{upgrade.S}}}sheetData'):
        if first < int(row.get('r')) <= last:
            cells = {upgrade.index(cell.get('r')): cell for cell in row}
            result.append({key: book.text(cells.get(upgrade.index(start) + index)) for index, key in enumerate(headers)})
    return result


fields = rows('tbBasisFields', 'Поля документів')
fragments = rows('tbBasisFragments', 'Фрагменти тексту')
metadata = rows('tbBasisContract', 'Контракт')
documents = rows('tbBasisDocumentTypes', 'Типи документів')


def tokens(text):
    pattern = r'\{([A-Za-z][A-Za-z0-9_]*(?:\.[A-Za-z][A-Za-z0-9_]*)*)\}'
    assert not re.search('[{}]', re.sub(pattern, '', text)), text
    return re.findall(pattern, text)


def matches(row, value):
    return (row['Умова'] == 'nonempty' and bool(value) or
            row['Умова'] == 'equals' and value == row['Значення'] or
            row['Умова'] == 'except' and bool(value) and value != row['Значення'])


def validate(field_rows, fragment_rows, meta_rows, mapping):
    assert len(meta_rows) == 1
    assert meta_rows[0]['Версія'] == config['basisChain.contract.version']
    assert set(meta_rows[0]['Можливості'].split('|')) == set(config['basisChain.contract.features'].split('|'))
    dictionary = {row['Код']: row for row in field_rows}
    assert len(dictionary) == len(field_rows)
    assert set(dictionary) == set(mapping)
    assert len(set(mapping.values())) == len(mapping)
    editor = set(config['basisChain.editorFields'].split('|'))
    dates = set(config['basisChain.dateFields'].split('|'))
    for code, row in dictionary.items():
        assert re.fullmatch(r'[A-Za-z][A-Za-z0-9_]*(?:\.[A-Za-z][A-Za-z0-9_]*)*', code)
        assert row['Назва'] and row['Опис']
        assert row['Тип'] in {'text', 'date', 'choice'}
        assert row['Тип'] == binding_types[mapping[code]]
        assert row['Використання'] in {'input', 'computed'}
        assert (row['Тип'] == 'date') == (mapping[code] in dates)
        if row['Використання'] == 'input':
            assert mapping[code] in editor
    groups = {}
    for row in fragment_rows:
        assert row['Код'] not in dictionary
        assert row['Назва'] and row['Опис']
        assert row['Умова'] in {'default', 'nonempty', 'equals', 'except'}
        if row['Умова'] == 'default':
            assert not row['Поле'] and not row['Значення']
        else:
            assert row['Поле'] in dictionary
        if row['Умова'] == 'nonempty':
            assert not row['Значення']
        groups.setdefault(row['Код'], []).append(row)
    for code, cases in groups.items():
        assert len({row['Поле'] for row in cases if row['Умова'] != 'default'}) <= 1
        assert sum(row['Умова'] == 'default' for row in cases) <= 1
        assert len({(row['Умова'], row['Поле'], row['Значення']) for row in cases}) == len(cases)
        values = {row['Значення'] for row in cases} | {''}
        other = '_'
        while other in values:
            other += '_'
        for value in values | {other}:
            assert sum(matches(row, value) for row in cases) <= 1, code

    def check(text, trail=()):
        for token in tokens(text):
            assert token in dictionary or token in groups, token
            if token in groups:
                assert token not in trail, token
                for row in groups[token]:
                    check(row['Шаблон'], trail + (token,))

    for code in groups:
        check('{' + code + '}')
    for document in documents:
        visible = set(document['Поля'].split('|'))
        required = set(re.findall(r'[A-Za-z][A-Za-z0-9_.]*', document['Обов’язкові поля']))
        assert required <= visible <= set(dictionary)
        assert all(dictionary[code]['Використання'] == 'input' for code in visible)
        check(document['Шаблон'])
        assert 'Basis' not in document['Шаблон'] + document['Поля'] + document['Обов’язкові поля']
    return groups


groups = validate(fields, fragments, metadata, bindings)


def render(text, values):
    def replace(match):
        code = match[1]
        if code not in groups:
            return values.get(code, '')
        selected = [row for row in groups[code] if row['Умова'] != 'default' and matches(row, values.get(row['Поле'], ''))]
        if not selected:
            selected = [row for row in groups[code] if row['Умова'] == 'default']
        return render(selected[0]['Шаблон'], values) if selected else ''
    return re.sub(r'\{([^}]+)\}', replace, text)


assert render('{document.number}', {'document.number': '123'}) == '123'
assert render('{document.numberText}', {'document.number': '123'}) == '№ 123'
assert render('{document.numberText}{document.dateText}{registration.text}{document.formText}{issuer.actingText}', {}) == ''
order = next(row for row in documents if row['Id'] == 'order')
values = {'document.type': 'наказ', 'document.type.genitive': 'наказу', 'document.number': '12',
          'document.date': '01.10.2026', 'issuer.role.genitive': 'командира', 'institution.genitive': 'військової частини А0001',
          'document.form': 'Витяг', 'document.scope': 'по особовому складу'}
assert ' '.join(render(order['Шаблон'], values).split()) == 'витяг із наказу командира військової частини А0001 (по особовому складу) № 12 від 01.10.2026'
assert render('{document.scopeText}', {'document.scope': 'Без уточнення'}) == ''
assert render('{issuer.actingText}', {'issuer.acting': 'Без ТВО'}) == ''
assert render('{issuer.actingText}', {'issuer.acting': 'Тимчасове виконання обов’язків'}) == 'тимчасово виконуючого обов’язки'

fixtures = []
for version, features in [('2', metadata[0]['Можливості']), ('1', metadata[0]['Можливості'] + '|unsupported.v1')]:
    broken = copy.deepcopy(metadata)
    broken[0].update({'Версія': version, 'Можливості': features})
    fixtures.append((fields, fragments, broken, bindings))
missing = dict(bindings)
missing.pop('document.number')
fixtures.append((fields, fragments, metadata, missing))
wrong_mapping = dict(bindings)
wrong_mapping['document.number'], wrong_mapping['issuer.role'] = wrong_mapping['issuer.role'], wrong_mapping['document.number']
fixtures.append((fields, fragments, metadata, wrong_mapping))
wrong_type = copy.deepcopy(fields)
next(row for row in wrong_type if row['Код'] == 'document.date')['Тип'] = 'text'
fixtures.append((wrong_type, fragments, metadata, bindings))
for template in ['{unknown.field}', '{document.numberText}']:
    broken = copy.deepcopy(fragments)
    broken[0]['Шаблон'] = template
    fixtures.append((fields, broken, metadata, bindings))
fixtures.append((fields, fragments + [copy.deepcopy(fragments[0])], metadata, bindings))
for fixture in fixtures:
    try:
        validate(*fixture)
    except AssertionError:
        pass
    else:
        raise AssertionError('Invalid contract accepted')

loader = (ROOT / 'vba/TravelDocumentsBuilder/obj_TDB_DependencyLoader.cls.vba').read_text(encoding='utf-8-sig')
assert loader.index('ex_TDB_BasisContract.fn_Validate snapshot') < loader.index('TryBuild = True')
chain = (ROOT / 'vba/TravelDocumentsBuilder/ex_TDB_BasisChain.vba').read_text(encoding='utf-8-sig')
assert 'ex_TDB_BasisContract.fn_AdaptDefinition result' in chain
assert 'row("BasisTemplate") = ex_TDB_BasisContract.fn_Template' in chain
assert '"{Value." & field & "}"' in chain
assert 'ex_TDB_BasisContract.fn_Caption' in chain
print(f'Public contract: {len(fields)} fields, {len(fragments)} fragment cases, {len(documents)} profiles; {len(fixtures)} negative fixtures: OK')