"""Check configuration, XML fixtures and historical-text integration without running VBA."""

import argparse
import copy
import importlib.util
from pathlib import Path
import shutil
import tempfile

from BasisTextFixtures import XmlText, value_maps
from BasisConfigFixtures import configured_rows, enrich


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
            (entry.split('=') for entry in config['basisChain.model.bindings'].split('|'))}
binding_types = {target.removeprefix('Form.'): kind for target, kind in
                 (entry.split('=') for entry in config['basisChain.model.bindingTypes'].split('|'))}


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


fields = configured_rows(config, 'Fields')
metadata = []
documents = enrich(config, 'Documents', rows('tbBasisDocumentTypes', 'Типи документів'))
assert 'tbBasisFragments' not in book.tables() and 'Фрагменти тексту' not in book.sheets()
assert all('Шаблон' not in row for row in documents)


def validate(field_rows, meta_rows, mapping):
    dictionary = {row['Код']: row for row in field_rows}
    assert len(dictionary) == len(field_rows)
    assert set(dictionary) == set(mapping)
    assert len(set(mapping.values())) == len(mapping)
    editor = set(config['basisChain.editorFields'].split('|'))
    dates = set(config['basisChain.dateFields'].split('|'))
    for code, row in dictionary.items():
        assert row['Назва'] and row['Опис']
        assert row['Тип'] == binding_types[mapping[code]]
        assert row['Використання'] in {'input', 'computed'}
        assert (row['Тип'] == 'date') == (mapping[code] in dates)
        if row['Використання'] == 'input':
            assert mapping[code] in editor
    for document in documents:
        visible = set(document['Поля'].split('|'))
        assert visible <= set(dictionary)
        assert all(dictionary[code]['Використання'] == 'input' for code in visible)
    return dictionary


dictionary = validate(fields, metadata, bindings)
profiles = {row['Id']: config['basisChain.textTemplates.profiles.' + row['Id']] for row in documents}
assert len(profiles) == len(documents)
folder = ROOT / 'workbook/TravelDocumentsBuilder' / config['basisChain.textTemplates.folder']
maps = value_maps(config)
own_order = config['basisChain.textTemplates.ownOrder']
engine = XmlText(folder, dictionary, maps, profiles, own_order)


def phrase(identifier, expected, **values):
    document = next(row for row in documents if row['Id'] == identifier)
    data = dict.fromkeys(dictionary, '')
    data.update({'document.type': document['Назва'], 'document.type.genitive': document['Родовий']})
    data.update(values)
    actual = engine.render(profiles[identifier], data)
    assert actual == expected, (actual, expected)


phrase('notification', 'повідомлення тимчасово виконуючого обов’язки командира військової частини А0001 № 12/34 від 01.10.2026 (вх. № 56 від 02.10.2026)',
       **{'issuer.acting': 'Тимчасове виконання обов’язків', 'issuer.role.genitive': 'командира',
          'institution.genitive': 'військової частини А0001', 'document.number': '12/34', 'document.date': '01.10.2026',
          'registration.number': '56', 'registration.date': '02.10.2026'})
phrase('order', 'витяг із наказу командира військової частини А0001 (по особовому складу) № 12 від 01.10.2026',
       **{'document.form': 'Витяг', 'issuer.role.genitive': 'командира', 'institution.genitive': 'військової частини А0001',
          'document.scope': 'по особовому складу', 'document.number': '12', 'document.date': '01.10.2026'})
phrase('letter', 'копія листа госпіталю № 12 від 01.10.2026',
       **{'document.form': 'Копія', 'institution.genitive': 'госпіталю', 'document.number': '12', 'document.date': '01.10.2026'})
phrase('report', 'рапорт капітана Прикладенка А.Б. (вх. № 12 від 01.10.2026)',
       **{'document.author': 'капітана Прикладенка А.Б.', 'registration.number': '12', 'registration.date': '01.10.2026'})
phrase('court_order', 'ухвала районного суду від 01.10.2026 у справі № 1/2 провадження № 3/4',
       **{'institution.genitive': 'районного суду', 'document.date': '01.10.2026', 'court.caseNumber': '1/2', 'court.proceedingNumber': '3/4'})
phrase('personnel_notification', 'повідомлення про прийняття кадрових рішень наказом начальника Генерального штабу Збройних Сил України (по особовому складу) від 01.10.2026 № 12',
       **{'reference.issuer': 'начальника Генерального штабу Збройних Сил України', 'reference.scope': 'по особовому складу',
          'reference.date': '01.10.2026', 'reference.number': '12'})
for identifier in profiles:
    phrase(identifier, next(row['Назва'] for row in documents if row['Id'] == identifier) +
           (' від 01.10.2026 у справі № 1/2' if identifier == 'court_order' else
            ' наказом автора від 01.10.2026 № 12' if identifier == 'personnel_notification' else
            ' № 12 від 01.10.2026'),
           **({'document.date': '01.10.2026', 'court.caseNumber': '1/2'} if identifier == 'court_order' else
              {'reference.issuer': 'автора', 'reference.date': '01.10.2026', 'reference.number': '12'} if identifier == 'personnel_notification' else
              {'document.number': '12', 'document.date': '01.10.2026', 'document.scope': 'Без уточнення', 'issuer.acting': 'Без ТВО'}))
assert engine.render(own_order, {'ownUnit.genitive': 'військової частини А7383', 'order.number': '123', 'order.date': '01.10.2026'}, normalize=False) == 'наказ командира військової частини А7383 (по стройовій частині) від 01.10.2026 № 123'

fixtures = []
bad = copy.deepcopy(fields)
bad[0]['Тип'] = 'wrong'
fixtures.append((bad, metadata, bindings))
bad = copy.deepcopy(fields)
bad.append(dict(bad[0]))
fixtures.append((bad, metadata, bindings))
bad = dict(bindings)
bad.pop('document.number')
fixtures.append((fields, metadata, bad))
bad = dict(bindings)
bad['document.number'], bad['issuer.role'] = bad['issuer.role'], bad['document.number']
fixtures.append((fields, metadata, bad))
for fixture in fixtures:
    try:
        validate(*fixture)
    except AssertionError:
        pass
    else:
        raise AssertionError('Invalid contract accepted')

invalid_xml = [
    ('Common.xml', '<textTemplate version="1"/>'),
    ('Common.xml', '<textTemplate version="1"><value field="unknown.field"/></textTemplate>'),
    ('Common.xml', '<textTemplate version="2"><text>x</text></textTemplate>'),
    ('Common.xml', '<textTemplate version="1"><include file="../escape.xml"/></textTemplate>'),
    ('Common.xml', '<textTemplate version="1"><include file="Missing.xml"/></textTemplate>'),
    ('DocumentTitle.xml', '<textTemplate version="1"><include file="Common.xml"/></textTemplate>'),
    ('DocumentTitle.xml', '<textTemplate version="1"><choose field="document.form"><case value="copy"/><case value="copy"/></choose></textTemplate>'),
    ('DocumentTitle.xml', '<textTemplate version="1"><choose field="document.form"><case value="copi"/></choose></textTemplate>'),
    ('Common.xml', '<textTemplate version="1"><if field="document.number"><text>x</text></if></textTemplate>'),
    ('Common.xml', '<textTemplate version="1">unexpected<value field="document.number"/></textTemplate>'),
    ('OwnOrder.xml', '<textTemplate version="1"><value field="document.type"/></textTemplate>'),
    ('Common.xml', '<!DOCTYPE x [<!ENTITY text "x">]><textTemplate version="1"><text>&text;</text></textTemplate>'),
]
for file, text in invalid_xml:
    with tempfile.TemporaryDirectory(prefix='tdb-xml-') as temporary:
        destination = Path(temporary)
        for asset in folder.glob('*.xml'):
            shutil.copy2(asset, destination / asset.name)
        (destination / file).write_text(text, encoding='utf-8')
        try:
            XmlText(destination, dictionary, maps, profiles, own_order)
        except (ValueError, FileNotFoundError):
            pass
        else:
            raise AssertionError(file + ': invalid XML accepted')

with tempfile.TemporaryDirectory(prefix='tdb-xml-cache-') as temporary:
    destination = Path(temporary)
    for asset in folder.glob('*.xml'):
        shutil.copy2(asset, destination / asset.name)
    cached = XmlText(destination, dictionary, maps, profiles, own_order)
    context = dict.fromkeys(dictionary, '')
    context.update({'document.type': 'наказ', 'document.number': '123'})
    approved = cached.render(profiles['order'], context)
    path = destination / 'Common.xml'
    path.write_text(path.read_text(encoding='utf-8').replace('</textTemplate>', '<text>changed</text></textTemplate>'), encoding='utf-8')
    assert cached.render(profiles['order'], context) == approved
    refreshed = XmlText(destination, dictionary, maps, profiles, own_order)
    assert refreshed.render(profiles['order'], context) != approved

loader = (ROOT / 'vba/TravelDocumentsBuilder/obj_TDB_DependencyLoader.cls.vba').read_text(encoding='utf-8-sig')
assert loader.index('ex_TDB_BasisConfig.fn_Validate snapshot') < loader.index('TryBuild = True')
chain = (ROOT / 'vba/TravelDocumentsBuilder/ex_TDB_BasisChain.vba').read_text(encoding='utf-8-sig')
assert 'row("BasisRenderedText") = ex_TDB_BasisConfig.fn_Render' in chain
assert 'fn_Format = VBA.CStr(row("BasisRenderedText"))' in chain
assert 'Public Function fn_SavedOwnText(' in chain
generator = (ROOT / 'vba/TravelDocumentsBuilder/obj_TDB_VacationGenerator.cls.vba').read_text(encoding='utf-8-sig')
assert 'ex_TDB_BasisChain.fn_SavedOwnText(m_basis, VBA.CStr(root("Basis")))' in generator
renderer = (ROOT / 'vba/TravelDocumentsBuilder/ex_TDB_BasisText.vba').read_text(encoding='utf-8-sig')
assert 'document.setProperty "ProhibitDTD", True' in renderer
assert 'document.resolveExternals = False' in renderer
assert 'private_Children(repository(file), repository, values)' in renderer
assert not any(key.startswith('basisChain.contract.') for key in config)
assert 'Поля документів' not in book.sheets() and 'Контракт' not in book.sheets()
print(f'XML configuration: {len(fields)} fields, {len(profiles)} profiles; {len(fixtures)} contract and {len(invalid_xml)} XML negative fixtures: OK')