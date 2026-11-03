"""Remove migrated application metadata from the institutions source workbook."""

import argparse
from datetime import datetime
import importlib.util
from pathlib import Path
import posixpath
import shutil


spec = importlib.util.spec_from_file_location('upgrade', Path(__file__).with_name('Upgrade-TDBBasisWorkbooks.py'))
u = importlib.util.module_from_spec(spec)
spec.loader.exec_module(u)


def rows(book, sheet_name, table_name):
    _, table = book.tables()[table_name]
    headers = [node.get('name') for node in table.find(f'{{{u.S}}}tableColumns')]
    start, end = table.get('ref').split(':')
    first, last = [int(''.join(filter(str.isdigit, address))) for address in (start, end)]
    result = []
    for row in u.parse(book.parts[book.sheets()[sheet_name]]).find(f'{{{u.S}}}sheetData'):
        if first < int(row.get('r')) <= last:
            cells = {u.index(cell.get('r')): cell for cell in row}
            result.append({key: book.text(cells.get(u.index(start) + offset)) for offset, key in enumerate(headers)})
    return result


def drop_sheet(book, name):
    if name not in book.sheets():
        return
    path = book.sheets()[name]
    removed = {path, posixpath.dirname(path) + '/_rels/' + posixpath.basename(path) + '.rels'}
    if list_path := book.parts.get(next(part for part in removed if part.endswith('.rels'))):
        for relation in u.parse(list_path):
            target = relation.get('Target')
            removed.add(target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.dirname(path) + '/' + target))
    workbook = u.parse(book.parts['xl/workbook.xml'])
    sheets = workbook.find(f'{{{u.S}}}sheets')
    node = next(node for node in sheets if node.get('name') == name)
    position = list(sheets).index(node)
    identifier = node.get(f'{{{u.R}}}id')
    sheets.remove(node)
    relations = u.parse(book.parts['xl/_rels/workbook.xml.rels'])
    relations.remove(next(node for node in relations if node.get('Id') == identifier))
    for view in workbook.findall('s:bookViews/s:workbookView', u.NS):
        for attr in ('activeTab', 'firstSheet'):
            if attr in view.attrib:
                previous = int(view.get(attr))
                view.set(attr, str(min(previous - (previous > position), len(sheets) - 1)))
    definitions = workbook.find(f'{{{u.S}}}definedNames')
    for item in list(definitions) if definitions is not None else []:
        if item.get('localSheetId') == str(position):
            definitions.remove(item)
        elif 'localSheetId' in item.attrib and int(item.get('localSheetId')) > position:
            item.set('localSheetId', str(int(item.get('localSheetId')) - 1))
    types = u.parse(book.parts['[Content_Types].xml'])
    for item in list(types):
        if item.get('PartName', '').lstrip('/') in removed:
            types.remove(item)
    for part in removed:
        book.parts.pop(part, None)
    book.parts['xl/workbook.xml'] = u.xml(workbook)
    book.parts['xl/_rels/workbook.xml.rels'] = u.xml(relations)
    book.parts['[Content_Types].xml'] = u.xml(types)


def keep_columns(book, sheet_name, table_name, keep):
    path, table = book.tables()[table_name]
    columns = table.find(f'{{{u.S}}}tableColumns')
    headers = [node.get('name') for node in columns]
    start, end = table.get('ref').split(':')
    first = u.index(start)
    removed = {first + offset for offset, name in enumerate(headers) if name not in keep}
    sheet_path = book.sheets()[sheet_name]
    sheet = u.parse(book.parts[sheet_path])
    if sheet.findall('.//s:f', u.NS) or sheet.find(f'{{{u.S}}}mergeCells') is not None:
        raise ValueError(f'Review formulas or merges before changing {sheet_name}')

    def address(value):
        old = u.index(value)
        return u.column(old - sum(number < old for number in removed)) + ''.join(filter(str.isdigit, value))

    for row in sheet.find(f'{{{u.S}}}sheetData'):
        row.attrib.pop('spans', None)
        if int(row.get('r')) < int(''.join(filter(str.isdigit, start))):
            continue
        for cell in list(row):
            if u.index(cell.get('r')) in removed:
                row.remove(cell)
            else:
                cell.set('r', address(cell.get('r')))
    cols = sheet.find(f'{{{u.S}}}cols')
    for col in list(cols) if cols is not None else []:
        low, high = int(col.get('min')), int(col.get('max'))
        remaining = [value for value in range(low, high + 1) if value not in removed]
        if not remaining:
            cols.remove(col)
        else:
            col.set('min', str(remaining[0] - sum(value < remaining[0] for value in removed)))
            col.set('max', str(remaining[-1] - sum(value < remaining[-1] for value in removed)))
    for node in list(columns):
        if node.get('name') not in keep:
            columns.remove(node)
    columns.set('count', str(len(columns)))
    for number, node in enumerate(columns, 1):
        node.set('id', str(number))
    table.set('ref', address(start) + ':' + u.column(first + len(keep) - 1) + ''.join(filter(str.isdigit, end)))
    auto_filter = table.find(f'{{{u.S}}}autoFilter')
    if auto_filter is not None:
        auto_filter.set('ref', table.get('ref'))
        for item in list(auto_filter):
            if item.tag.endswith('filterColumn'):
                previous = first + int(item.get('colId'))
                if previous in removed:
                    auto_filter.remove(item)
                else:
                    item.set('colId', str(previous - first - sum(value < previous for value in removed)))
    dimension = sheet.find(f'{{{u.S}}}dimension')
    if dimension is not None:
        dimension.set('ref', ':'.join(address(value) for value in dimension.get('ref').split(':')))
    book.parts[path] = u.xml(table)
    book.parts[sheet_path] = u.xml(sheet)


def migrate(book, config):
    if 'Поля документів' in book.sheets():
        for row in rows(book, 'Поля документів', 'tbBasisFields'):
            prefix = 'basisChain.model.Fields.' + row['Код'] + '.'
            for key, header in {'Name': 'Назва', 'Description': 'Опис', 'Type': 'Тип', 'Usage': 'Використання'}.items():
                assert config[prefix + key] == row[header], (prefix + key, 'field settings have not been migrated')
    for kind, sheet, table in [('Documents', 'Типи документів', 'tbBasisDocumentTypes'),
                               ('Officials', 'Посадові особи', 'tbBasisOfficials')]:
        for row in rows(book, sheet, table):
            if 'Id' not in row:
                continue
            prefix = 'basisChain.model.' + kind + '.' + row['Id'] + '.'
            for key, header in {'Name': 'Назва', **({'Forms': 'Форми', 'Required': 'Обов’язкові поля', 'Fields': 'Поля'} if kind == 'Documents' else {})}.items():
                assert config[prefix + key] == row[header], (prefix + key, 'metadata has not been migrated')
    for kind, sheet, table, header in [('DocumentRules', 'Документи установ', 'tbBasisDocumentRules', 'Id документа'),
                                       ('OfficialRules', 'Посади установ', 'tbBasisOfficialRules', 'Id посадової особи')]:
        if sheet in book.sheets():
            actual = {(row['Група установи'], row[header]) for row in rows(book, sheet, table)}
            prefix = 'basisChain.model.' + kind
            expected = {(group, identifier) for group in config[prefix + '.keys'].split('|')
                        for identifier in config[prefix + '.' + group].split('|')}
            assert actual == expected, (kind, 'rules have not been migrated')
    keep_columns(book, 'Типи документів', 'tbBasisDocumentTypes', ['Назва', 'Родовий'])
    keep_columns(book, 'Посадові особи', 'tbBasisOfficials', ['Назва', 'Родовий'])
    for name in ['Контракт', 'Поля документів', 'Фрагменти тексту', 'Документи установ', 'Посади установ']:
        drop_sheet(book, name)


def main():
    root = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--institutions', type=Path, default=root / 'workbook/Dependencies/Установи.xlsx')
    parser.add_argument('--output-dir', type=Path, required=True)
    parser.add_argument('--apply', action='store_true')
    args = parser.parse_args()
    config = {parts[1].split('::', 1)[1]: parts[2] for line in
              (root / 'config/TravelDocumentsBuilder/wsConfig.txt').read_text(encoding='utf-8-sig').splitlines()
              if line for parts in [line.split('\t')] if parts[0] == 'value'}
    book = u.Package(args.institutions)
    original = u.Package(args.institutions)
    retained = {name: rows(original, name, table) for name, table in
                [('Типи документів', 'tbBasisDocumentTypes'), ('Посадові особи', 'tbBasisOfficials')]}
    migrate(book, config)
    for name, table in [('Типи документів', 'tbBasisDocumentTypes'), ('Посадові особи', 'tbBasisOfficials')]:
        assert rows(book, name, table) == [{key: row[key] for key in ['Назва', 'Родовий']} for row in retained[name]]
    args.output_dir.mkdir(parents=True, exist_ok=True)
    output = args.output_dir / args.institutions.name
    assert output.resolve() != args.institutions.resolve()
    book.save(output)
    if args.apply:
        backup = args.output_dir / (args.institutions.stem + '.' + datetime.now().strftime('%Y%m%d-%H%M%S') + '.backup.xlsx')
        shutil.copy2(args.institutions, backup)
        shutil.copy2(output, args.institutions)
    print(f'Source-only workbook verified: {len(book.sheets())} sheets; all retained names and declensions unchanged.')


if __name__ == '__main__':
    main()