"""Upgrade reference and history schemas while preserving unrelated XLSX package parts."""

import argparse
import copy
from datetime import datetime
import io
import math
from pathlib import Path
import posixpath
import re
import shutil
import xml.etree.ElementTree as ET
import zipfile


S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'
C = 'http://schemas.openxmlformats.org/package/2006/content-types'
NS = {'s': S, 'r': R}
ET.register_namespace('', S)
ET.register_namespace('r', R)
NAMESPACES = {}


def parse(data):
    namespaces = {}
    for _, pair in ET.iterparse(io.BytesIO(data), events=['start-ns']):
        namespaces[pair[0]] = pair[1]
        if pair[0] and not pair[0].startswith('ns'):
            ET.register_namespace(*pair)
    root = ET.fromstring(data)
    NAMESPACES[id(root)] = namespaces
    return root


def xml(root):
    namespaces = NAMESPACES.get(id(root), {})
    for prefix, namespace in namespaces.items():
        if prefix and not prefix.startswith('ns'):
            ET.register_namespace(prefix, namespace)
    ET.register_namespace('', root.tag.split('}', 1)[0][1:])
    used = {name.split('}', 1)[0][1:] for node in root.iter() for name in [node.tag, *node.attrib] if name.startswith('{')}
    ignorable = root.get('{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable', '').split()
    for prefix in ignorable:
        if prefix in namespaces and namespaces[prefix] not in used:
            root.set('xmlns:' + prefix, namespaces[prefix])
    return ET.tostring(root, encoding='utf-8', xml_declaration=True)


def column(number):
    result = ''
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def index(address):
    result = 0
    for char in address:
        if char.isalpha():
            result = result * 26 + ord(char.upper()) - 64
    return result


class Package:
    def __init__(self, path):
        self.path = Path(path)
        with zipfile.ZipFile(path) as archive:
            self.parts = {name: archive.read(name) for name in archive.namelist()}
        self.strings = []
        if 'xl/sharedStrings.xml' in self.parts:
            self.strings = [''.join(node.itertext()) for node in parse(self.parts['xl/sharedStrings.xml'])]

    def text(self, cell):
        if cell is None:
            return ''
        value = cell.find(f'{{{S}}}v')
        if cell.get('t') == 's':
            return self.strings[int(value.text)]
        if cell.get('t') == 'inlineStr':
            return ''.join(cell.find(f'{{{S}}}is').itertext())
        return value.text if value is not None and value.text else ''

    def sheets(self):
        relationships = {node.get('Id'): node.get('Target') for node in parse(self.parts['xl/_rels/workbook.xml.rels'])}
        result = {}
        for node in parse(self.parts['xl/workbook.xml']).find(f'{{{S}}}sheets'):
            target = relationships[node.get(f'{{{R}}}id')]
            target = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
            result[node.get('name')] = target
        return result

    def tables(self):
        return {parse(data).get('name'): (name, parse(data)) for name, data in self.parts.items()
                if name.startswith('xl/tables/') and name.endswith('.xml')}

    def save(self, destination):
        destination = Path(destination)
        destination.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(destination, 'w', zipfile.ZIP_DEFLATED) as archive:
            for name, data in self.parts.items():
                archive.writestr(name, data)


def string_cell(address, value, style=None):
    cell = ET.Element(f'{{{S}}}c', {'r': address, 't': 'inlineStr'})
    if style is not None:
        cell.set('s', str(style))
    text = ET.SubElement(ET.SubElement(cell, f'{{{S}}}is'), f'{{{S}}}t')
    text.text = value
    return cell


def import_styles(target, source):
    original = parse(target.parts['xl/styles.xml'])
    incoming = parse(source.parts['xl/styles.xml'])
    offsets = {}
    number_formats = {}
    formats = original.find(f'{{{S}}}numFmts')
    if formats is None:
        formats = ET.Element(f'{{{S}}}numFmts', {'count': '0'})
        original.insert(0, formats)
    next_format = max([163] + [int(node.get('numFmtId')) for node in formats]) + 1
    source_formats = incoming.find(f'{{{S}}}numFmts')
    for node in source_formats if source_formats is not None else []:
        number_formats[node.get('numFmtId')] = str(next_format)
        duplicate = copy.deepcopy(node)
        duplicate.set('numFmtId', str(next_format))
        formats.append(duplicate)
        next_format += 1
    formats.set('count', str(len(formats)))
    for name in ['fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs']:
        base = original.find(f'{{{S}}}{name}')
        extra = incoming.find(f'{{{S}}}{name}')
        offsets[name] = len(base)
        for node in extra:
            duplicate = copy.deepcopy(node)
            if name in ['cellStyleXfs', 'cellXfs']:
                for field, group in [('fontId', 'fonts'), ('fillId', 'fills'), ('borderId', 'borders'), ('xfId', 'cellStyleXfs')]:
                    if field in duplicate.attrib:
                        duplicate.set(field, str(int(duplicate.get(field)) + offsets[group]))
                if duplicate.get('numFmtId') in number_formats:
                    duplicate.set('numFmtId', number_formats[duplicate.get('numFmtId')])
            base.append(duplicate)
        base.set('count', str(len(base)))
    target.parts['xl/styles.xml'] = xml(original)
    return offsets['cellXfs']


def rename_rule_sheets(book):
    aliases = {'Правила': 'Посади установ', 'Правила документів': 'Документи установ'}
    paths = book.sheets()
    workbook = parse(book.parts['xl/workbook.xml'])
    changed = False
    for node in workbook.find(f'{{{S}}}sheets'):
        previous = node.get('name')
        if previous not in aliases:
            continue
        name = aliases[previous]
        if name in paths:
            raise ValueError(f'Both legacy and current rule sheets exist: {previous}, {name}')
        node.set('name', name)
        sheet_path = paths[previous]
        sheet = parse(book.parts[sheet_path])
        cell = next(cell for cell in sheet.findall('.//s:c', NS) if cell.get('r') == 'B2')
        for child in list(cell):
            cell.remove(child)
        cell.set('t', 'inlineStr')
        ET.SubElement(ET.SubElement(cell, f'{{{S}}}is'), f'{{{S}}}t').text = name
        book.parts[sheet_path] = xml(sheet)
        changed = True
    if changed:
        book.parts['xl/workbook.xml'] = xml(workbook)


def style_reference_sheets(target):
    # Reuse the workbook's native custom style instead of approximating its palette.
    template = parse(target.parts[target.sheets()['Органи управління']])
    template_cells = {cell.get('r'): cell for cell in template.findall('.//s:c', NS)}
    template_rows = {int(row.get('r')): row for row in template.find(f'{{{S}}}sheetData')}
    template_table = target.tables()['tbCommands'][1]
    for table_name, sheet_name in [('tbBasisDocumentTypes', 'Типи документів'),
                                   ('tbBasisDocumentRules', 'Документи установ'),
                                   ('tbCommands', 'Органи управління'), ('tbCourts', 'Суди'),
                                   ('tbBasisFields', 'Поля документів'), ('tbBasisFragments', 'Фрагменти тексту'),
                                   ('tbBasisContract', 'Контракт'), ('tbEnumInstitutionGroups', 'Переліки')]:
        if table_name not in target.tables():
            continue
        table_path, table = target.tables()[table_name]
        sheet_path = target.sheets()[sheet_name]
        sheet = parse(target.parts[sheet_path])
        data = sheet.find(f'{{{S}}}sheetData')
        rows = {int(row.get('r')): row for row in data}
        end = table.get('ref').split(':')[1]
        last_column = index(end)
        last_row = int(''.join(filter(str.isdigit, end)))
        widths = {}
        for col in sheet.findall('s:cols/s:col', NS):
            for number in range(int(col.get('min')), min(int(col.get('max')), last_column) + 1):
                widths[number] = float(col.get('width', '20'))
        for number in range(1, max(last_row + 2, 40) + 1):
            row = rows.get(number)
            if row is None:
                row = ET.SubElement(data, f'{{{S}}}row', {'r': str(number)})
            cells = {index(cell.get('r')): cell for cell in row}
            height = float(template_rows.get(number if number < 6 else 6).get('ht', '18.95'))
            for col in range(1, last_column + 2):
                cell = cells.get(col)
                if cell is None:
                    cell = ET.SubElement(row, f'{{{S}}}c', {'r': column(col) + str(number)})
                exemplar = 'A1'
                if number == 2 and col == 2:
                    exemplar = 'B2'
                elif number == 5 and 2 <= col <= last_column:
                    exemplar = 'B5'
                elif 6 <= number <= last_row and 2 <= col <= last_column:
                    exemplar = 'C6' if col == 3 else 'B6'
                    lines = math.ceil(len(target.text(cell)) / max(1, widths.get(col, 20) - 2))
                    height = max(height, lines * 14.25 + 4)
                cell.set('s', template_cells[exemplar].get('s'))
            row.set('ht', str(height))
            row.set('customHeight', '1')
            row[:] = sorted(row, key=lambda cell: index(cell.get('r')))
        data[:] = sorted(data, key=lambda row: int(row.get('r')))
        info = table.find(f'{{{S}}}tableStyleInfo')
        if info is not None:
            table.remove(info)
        table.append(copy.deepcopy(template_table.find(f'{{{S}}}tableStyleInfo')))
        auto_filter = table.find(f'{{{S}}}autoFilter')
        if auto_filter is None:
            auto_filter = ET.Element(f'{{{S}}}autoFilter')
            table.insert(0, auto_filter)
        auto_filter.set('ref', table.get('ref'))
        for node in table.find(f'{{{S}}}tableColumns'):
            for field in ['hiddenButton', 'showFilterButton']:
                node.attrib.pop(field, None)
        target.parts[sheet_path] = xml(sheet)
        target.parts[table_path] = xml(table)


def upgrade_references(target, source):
    style_offset = import_styles(target, source)
    workbook = parse(target.parts['xl/workbook.xml'])
    sheets = workbook.find(f'{{{S}}}sheets')
    relationships = parse(target.parts['xl/_rels/workbook.xml.rels'])
    content_types = parse(target.parts['[Content_Types].xml'])
    target_sheets = target.sheets()
    target_tables = target.tables()
    next_sheet = max(int(node.get('sheetId')) for node in sheets) + 1
    next_table = max(int(table.get('id')) for _, table in target_tables.values()) + 1
    existing_ids = {node.get('Id') for node in relationships}
    for sheet_name, source_path in source.sheets().items():
        incoming = parse(source.parts[source_path])
        for cell in incoming.findall('.//s:c', NS):
            if cell.get('t') == 's':
                replacement = string_cell(cell.get('r'), source.text(cell), cell.get('s'))
                cell.clear()
                cell.attrib.update(replacement.attrib)
                cell.extend(list(replacement))
            if 's' in cell.attrib:
                cell.set('s', str(int(cell.get('s')) + style_offset))
        source_rels = posixpath.dirname(source_path) + '/_rels/' + posixpath.basename(source_path) + '.rels'
        rels = parse(source.parts[source_rels])
        relation = next(node for node in rels if node.get('Type').endswith('/table'))
        source_table_path = posixpath.normpath(posixpath.dirname(source_path) + '/' + relation.get('Target'))
        if relation.get('Target').startswith('/'):
            source_table_path = relation.get('Target').lstrip('/')
        table = parse(source.parts[source_table_path])
        name = table.get('name')
        if name in target_tables:
            target_path = target_sheets[sheet_name]
            sheet = parse(target.parts[target_path])
            data = sheet.find(f'{{{S}}}sheetData')
            table_path, existing_table = target_tables[name]
            start, end = existing_table.get('ref').split(':')
            first_row = int(''.join(filter(str.isdigit, start)))
            last_row = int(''.join(filter(str.isdigit, end)))
            first_column = index(start)
            width = len(existing_table.find(f'{{{S}}}tableColumns'))
            target_headers = [node.get('name') for node in existing_table.find(f'{{{S}}}tableColumns')]
            source_headers = [node.get('name') for node in table.find(f'{{{S}}}tableColumns')]
            source_offsets = {header: offset for offset, header in enumerate(source_headers)}
            rows_by_number = {int(row.get('r')): row for row in data}
            last_existing_row = rows_by_number.get(last_row)
            existing_styles = {index(cell.get('r')): cell.get('s') for cell in last_existing_row} if last_existing_row is not None else {}
            row_height = last_existing_row.get('ht', '30') if last_existing_row is not None else '30'
            keys = set()
            key_offsets = ([target_headers.index('Назва')] if name in ['tbCommands', 'tbCourts'] else
                           [0, 3, 4, 5] if name == 'tbBasisFragments' else [0, 1] if name.endswith('Rules') else [0])
            for number in range(first_row + 1, last_row + 1):
                row = rows_by_number.get(number)
                cells = {index(cell.get('r')): cell for cell in row} if row is not None else {}
                keys.add(tuple(target.text(cells.get(first_column + offset)) for offset in key_offsets))
            for row in incoming.find(f'{{{S}}}sheetData'):
                if int(row.get('r')) <= 5:
                    continue
                if name == 'tbBasisContract':
                    # The explicit architecture migration updates the single metadata record.
                    continue
                cells = {index(cell.get('r')): cell for cell in row}
                values = [source.text(cells.get(2 + source_offsets[header])) if header in source_offsets else '' for header in target_headers]
                key = tuple(values[offset] for offset in key_offsets)
                if not values[0] or key in keys:
                    continue
                keys.add(key)
                last_row += 1
                destination_row = rows_by_number.get(last_row)
                if destination_row is None:
                    destination_row = ET.SubElement(data, f'{{{S}}}row', {'r': str(last_row), 'ht': row_height, 'customHeight': '1'})
                    rows_by_number[last_row] = destination_row
                for offset, value in enumerate(values):
                    old = next((cell for cell in destination_row if index(cell.get('r')) == first_column + offset), None)
                    if old is not None:
                        destination_row.remove(old)
                    source_cell = cells.get(2 + source_offsets[target_headers[offset]]) if target_headers[offset] in source_offsets else None
                    style = existing_styles.get(first_column + offset, source_cell.get('s') if source_cell is not None else None)
                    destination_row.append(string_cell(column(first_column + offset) + str(last_row), value, style))
                destination_row[:] = sorted(destination_row, key=lambda cell: index(cell.get('r')))
            data[:] = sorted(data, key=lambda row: int(row.get('r')))
            new_ref = start + ':' + column(first_column + width - 1) + str(last_row)
            existing_table.set('ref', new_ref)
            auto_filter = existing_table.find(f'{{{S}}}autoFilter')
            if auto_filter is not None:
                auto_filter.set('ref', new_ref)
            validations = sheet.find(f'{{{S}}}dataValidations')
            for validation in validations if validations is not None else []:
                old_ref = validation.get('sqref', '')
                if re.fullmatch(r'[A-Z]+' + str(first_row + 1) + r':[A-Z]+' + ''.join(filter(str.isdigit, end)), old_ref):
                    validation.set('sqref', old_ref.split(':')[0] + ':' + column(index(old_ref.split(':')[1])) + str(last_row))
            dimension = sheet.find(f'{{{S}}}dimension')
            if dimension is not None:
                dimension.set('ref', 'A1:' + column(max(width + 1, 8)) + str(max(last_row, 40)))
            target.parts[target_path] = xml(sheet)
            target.parts[table_path] = xml(existing_table)
        else:
            target_path = f'xl/worksheets/sheet{next_sheet}.xml'
            while target_path in target.parts:
                next_sheet += 1
                target_path = f'xl/worksheets/sheet{next_sheet}.xml'
            table_path = f'xl/tables/table{next_table}.xml'
            table.set('id', str(next_table))
            relation.set('Target', f'../tables/table{next_table}.xml')
            next_table += 1
            relationship_id = f'rIdBasis{next_sheet}'
            assert relationship_id not in existing_ids
            ET.SubElement(sheets, f'{{{S}}}sheet', {'name': sheet_name, 'sheetId': str(next_sheet), f'{{{R}}}id': relationship_id})
            ET.SubElement(relationships, f'{{{P}}}Relationship', {'Id': relationship_id, 'Type': R + '/worksheet', 'Target': '/' + target_path})
            ET.SubElement(content_types, f'{{{C}}}Override', {'PartName': '/' + target_path, 'ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'})
            ET.SubElement(content_types, f'{{{C}}}Override', {'PartName': '/' + table_path, 'ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'})
            target.parts[target_path] = xml(incoming)
            target.parts[posixpath.dirname(target_path) + '/_rels/' + posixpath.basename(target_path) + '.rels'] = xml(rels)
            target.parts[table_path] = xml(table)
            next_sheet += 1
    target.parts['xl/workbook.xml'] = xml(workbook)
    target.parts['xl/_rels/workbook.xml.rels'] = xml(relationships)
    target.parts['[Content_Types].xml'] = xml(content_types)
    style_reference_sheets(target)


def upgrade_history(book, config):
    for name, key in [('tbTickets', 'history.columns'), ('tbTicketRevisions', 'history.versions.columns')]:
        table_path, table = book.tables()[name]
        expected = config[key].split('|')
        columns = table.find(f'{{{S}}}tableColumns')
        existing = [node.get('name') for node in columns]
        missing = [field for field in expected if field not in existing]
        # Preserve retired columns and all historical values; runtime ignores surplus columns.
        if not missing:
            continue
        sheet_path = next(path for path in book.sheets().values() if any(
            posixpath.basename(table_path).encode() in book.parts.get(posixpath.dirname(path) + '/_rels/' + posixpath.basename(path) + '.rels', b'')
            for _ in [0]))
        sheet = parse(book.parts[sheet_path])
        start, end = table.get('ref').split(':')
        row_number = int(''.join(filter(str.isdigit, start)))
        first_column = index(start)
        header = next(row for row in sheet.find(f'{{{S}}}sheetData') if int(row.get('r')) == row_number)
        style = header[-1].get('s')
        for field in missing:
            number = first_column + len(columns)
            header.append(string_cell(column(number) + str(row_number), field, style))
            ET.SubElement(columns, f'{{{S}}}tableColumn', {'id': str(max(int(node.get('id')) for node in columns) + 1), 'name': field})
        columns.set('count', str(len(columns)))
        new_ref = start + ':' + column(first_column + len(columns) - 1) + ''.join(filter(str.isdigit, end))
        table.set('ref', new_ref)
        auto_filter = table.find(f'{{{S}}}autoFilter')
        if auto_filter is not None:
            auto_filter.set('ref', new_ref)
        dimension = sheet.find(f'{{{S}}}dimension')
        if dimension is not None:
            dimension.set('ref', 'A1:' + new_ref.split(':')[1])
        book.parts[sheet_path] = xml(sheet)
        book.parts[table_path] = xml(table)


def explicit_requirement(specification):
    """Make the legacy grouping explicit without changing its Boolean meaning."""
    if '(' in specification or ')' in specification:
        return specification
    if not specification.strip():
        return specification

    def group(parts, operation):
        if any(not part for part in parts):
            raise ValueError(f'Invalid required-field expression: {specification}')
        return parts[0] if len(parts) == 1 else '(' + operation.join(parts) + ')'

    requirements = []
    for requirement in specification.split('|'):
        alternatives = []
        for alternative in requirement.split('/'):
            fields = [field.strip() for field in alternative.split('&')]
            if not all(re.fullmatch(r'[A-Za-z][A-Za-z0-9_.]*', field) for field in fields):
                raise ValueError(f'Invalid required-field expression: {specification}')
            alternatives.append(group(fields, '&'))
        requirements.append(group(alternatives, '/'))
    return group(requirements, '&')


def explicit_reference_requirements(book):
    _, table = book.tables()['tbBasisDocumentTypes']
    headers = [node.get('name') for node in table.find(f'{{{S}}}tableColumns')]
    start, end = table.get('ref').split(':')
    required_column = index(start) + headers.index('Обов’язкові поля')
    first_row = int(''.join(filter(str.isdigit, start)))
    last_row = int(''.join(filter(str.isdigit, end)))
    sheet_path = book.sheets()['Типи документів']
    sheet = parse(book.parts[sheet_path])
    for row in sheet.find(f'{{{S}}}sheetData'):
        if not first_row < int(row.get('r')) <= last_row:
            continue
        cell = next((cell for cell in row if index(cell.get('r')) == required_column), None)
        if cell is None:
            continue
        previous = book.text(cell)
        value = explicit_requirement(previous)
        if value != previous:
            for child in list(cell):
                cell.remove(child)
            cell.set('t', 'inlineStr')
            ET.SubElement(ET.SubElement(cell, f'{{{S}}}is'), f'{{{S}}}t').text = value
    book.parts[sheet_path] = xml(sheet)


def institution_group_validation(book):
    table_name = 'tbEnumInstitutionGroups'
    sheet_name = 'Переліки'
    default_groups = [('Unit', 'Військова частина'), ('Hospital', 'Лікувальний заклад'), ('TCC', 'ТЦК та СП'),
                      ('ChiefCommand', 'Збройні Сили України'), ('GeneralStaff', 'Генеральний штаб'),
                      ('LandCommand', 'Командування Сухопутних військ'), ('Court', 'Суд')]
    if table_name not in book.tables():
        if sheet_name in book.sheets():
            raise ValueError('The enumeration sheet exists but its institution-group table is missing')
        template = parse(book.parts[book.sheets()['Органи управління']])
        style_cells = {cell.get('r'): cell.get('s') for cell in template.findall('.//s:c', NS)}
        template_table = book.tables()['tbCommands'][1]
        workbook = parse(book.parts['xl/workbook.xml'])
        sheets = workbook.find(f'{{{S}}}sheets')
        relationships = parse(book.parts['xl/_rels/workbook.xml.rels'])
        content_types = parse(book.parts['[Content_Types].xml'])
        number = max(int(node.get('sheetId')) for node in sheets) + 1
        while f'xl/worksheets/sheet{number}.xml' in book.parts:
            number += 1
        table_number = max(int(table.get('id')) for _, table in book.tables().values()) + 1
        sheet_path = f'xl/worksheets/sheet{number}.xml'
        table_path = f'xl/tables/table{table_number}.xml'
        relation_id = f'rIdEnums{number}'
        assert relation_id not in {node.get('Id') for node in relationships}
        ET.SubElement(sheets, f'{{{S}}}sheet', {'name': sheet_name, 'sheetId': str(number), f'{{{R}}}id': relation_id})
        ET.SubElement(relationships, f'{{{P}}}Relationship', {'Id': relation_id, 'Type': R + '/worksheet', 'Target': '/' + sheet_path})
        for part, kind in [(sheet_path, 'worksheet'), (table_path, 'table')]:
            ET.SubElement(content_types, f'{{{C}}}Override', {'PartName': '/' + part, 'ContentType': f'application/vnd.openxmlformats-officedocument.spreadsheetml.{kind}+xml'})
        sheet = ET.Element(f'{{{S}}}worksheet')
        ET.SubElement(sheet, f'{{{S}}}dimension', {'ref': 'A1:D40'})
        ET.SubElement(ET.SubElement(sheet, f'{{{S}}}sheetViews'), f'{{{S}}}sheetView', {'workbookViewId': '0'})
        ET.SubElement(sheet, f'{{{S}}}sheetFormatPr', {'defaultRowHeight': '18.95'})
        columns = ET.SubElement(sheet, f'{{{S}}}cols')
        for col, width in [(1, 2), (2, 28), (3, 64), (4, 2)]:
            ET.SubElement(columns, f'{{{S}}}col', {'min': str(col), 'max': str(col), 'width': str(width), 'customWidth': '1'})
        data = ET.SubElement(sheet, f'{{{S}}}sheetData')
        for row_number in range(1, 41):
            height = 39.95 if row_number == 5 else 27 if row_number == 2 else 23.1 if 6 <= row_number <= 12 else 18.95
            row = ET.SubElement(data, f'{{{S}}}row', {'r': str(row_number), 'ht': str(height), 'customHeight': '1'})
            for col in range(1, 5):
                value = ''
                exemplar = 'A1'
                if row_number == 2 and col == 2:
                    value, exemplar = sheet_name, 'B2'
                elif row_number == 5 and col in [2, 3]:
                    value, exemplar = ['Група установи', 'Назва'][col - 2], 'B5'
                elif 6 <= row_number <= 12 and col in [2, 3]:
                    value, exemplar = default_groups[row_number - 6][col - 2], 'B6' if col == 2 else 'C6'
                row.append(string_cell(column(col) + str(row_number), value, style_cells[exemplar]))
        table = ET.Element(f'{{{S}}}table', {'id': str(table_number), 'name': table_name, 'displayName': table_name, 'ref': 'B5:C12', 'totalsRowShown': '0'})
        ET.SubElement(table, f'{{{S}}}autoFilter', {'ref': 'B5:C12'})
        table_columns = ET.SubElement(table, f'{{{S}}}tableColumns', {'count': '2'})
        for col, name in enumerate(['Група установи', 'Назва'], 1):
            ET.SubElement(table_columns, f'{{{S}}}tableColumn', {'id': str(col), 'name': name})
        table.append(copy.deepcopy(template_table.find(f'{{{S}}}tableStyleInfo')))
        ET.SubElement(ET.SubElement(sheet, f'{{{S}}}tableParts', {'count': '1'}), f'{{{S}}}tablePart', {f'{{{R}}}id': 'rIdEnumTable'})
        rels = ET.Element(f'{{{P}}}Relationships')
        ET.SubElement(rels, f'{{{P}}}Relationship', {'Id': 'rIdEnumTable', 'Type': R + '/table', 'Target': f'../tables/table{table_number}.xml'})
        book.parts[sheet_path] = xml(sheet)
        book.parts[table_path] = xml(table)
        book.parts[f'xl/worksheets/_rels/sheet{number}.xml.rels'] = xml(rels)
        book.parts['xl/workbook.xml'] = xml(workbook)
        book.parts['xl/_rels/workbook.xml.rels'] = xml(relationships)
        book.parts['[Content_Types].xml'] = xml(content_types)
    _, enum_table = book.tables()[table_name]
    headers = [node.get('name') for node in enum_table.find(f'{{{S}}}tableColumns')]
    start, end = enum_table.get('ref').split(':')
    code_column = index(start) + headers.index('Група установи')
    first = int(''.join(filter(str.isdigit, start)))
    last = int(''.join(filter(str.isdigit, end)))
    enum_sheet = parse(book.parts[book.sheets()[sheet_name]])
    allowed = []
    for row in enum_sheet.find(f'{{{S}}}sheetData'):
        if first < int(row.get('r')) <= last:
            cell = next((cell for cell in row if index(cell.get('r')) == code_column), None)
            value = book.text(cell)
            if not value or value != value.strip() or value in allowed:
                raise ValueError(f'Blank, duplicate or untrimmed institution-group enumeration: {value}')
            allowed.append(value)
    if not allowed:
        raise ValueError('The institution-group enumeration is empty')
    for name, path in book.sheets().items():
        if name == sheet_name:
            continue
        sheet = parse(book.parts[path])
        group_columns = {}
        for cell in sheet.findall('.//s:c', NS):
            if book.text(cell) == 'Група установи':
                group_columns[index(cell.get('r'))] = int(''.join(filter(str.isdigit, cell.get('r')))) + 1
        if not group_columns:
            continue
        for cell in sheet.findall('.//s:c', NS):
            col = index(cell.get('r'))
            row = int(''.join(filter(str.isdigit, cell.get('r'))))
            if col in group_columns and row >= group_columns[col]:
                value = book.text(cell)
                if value and value not in allowed:
                    raise ValueError(f'Existing institution group is not in the enumeration: {name}!{cell.get("r")} = {value}')
        validations = sheet.find(f'{{{S}}}dataValidations')
        if validations is None:
            validations = ET.Element(f'{{{S}}}dataValidations')
            following = {'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'drawing', 'legacyDrawing', 'tableParts', 'extLst'}
            offset = next((i for i, node in enumerate(sheet) if node.tag.split('}')[-1] in following), len(sheet))
            sheet.insert(offset, validations)
        for col, first_row in group_columns.items():
            address = column(col)
            target_range = f'{address}{first_row}:{address}1048576'
            for rule in list(validations):
                remaining = [ref for ref in rule.get('sqref', '').split() if not re.fullmatch(address + r'\d+:' + address + r'\d+', ref)]
                if remaining:
                    rule.set('sqref', ' '.join(remaining))
                else:
                    validations.remove(rule)
            rule = ET.SubElement(validations, f'{{{S}}}dataValidation', {
                'type': 'list', 'errorStyle': 'stop', 'allowBlank': '0', 'showDropDown': '0',
                'showInputMessage': '1', 'showErrorMessage': '1', 'sqref': target_range,
                'promptTitle': 'Група установи', 'prompt': 'Оберіть групу з переліку на аркуші «Переліки».',
                'errorTitle': 'Невідома група установи', 'error': 'Оберіть значення зі списку. Дозволені групи зберігаються на аркуші «Переліки».'})
            ET.SubElement(rule, f'{{{S}}}formula1').text = 'INDIRECT("tbEnumInstitutionGroups[Група установи]")'
        validations.set('count', str(len(validations)))
        book.parts[path] = xml(sheet)
    workbook = parse(book.parts['xl/workbook.xml'])
    sheets = workbook.find(f'{{{S}}}sheets')
    enum_node = next(node for node in sheets if node.get('name') == sheet_name)
    sheets.remove(enum_node)
    sheets.append(enum_node)
    book.parts['xl/workbook.xml'] = xml(workbook)


def xml_basis_architecture(book, config):
    # Keep the legacy command-line switch as an alias for the source-only migration.
    import importlib.util
    spec = importlib.util.spec_from_file_location('source_only', Path(__file__).with_name('SourceOnly-TDBInstitutions.py'))
    migration = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(migration)
    migration.migrate(book, config)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--references', type=Path)
    parser.add_argument('--institutions', type=Path, required=True)
    parser.add_argument('--history', type=Path)
    parser.add_argument('--output-dir', type=Path, required=True)
    parser.add_argument('--apply', action='store_true')
    parser.add_argument('--explicit-required-only', action='store_true')
    parser.add_argument('--rename-rule-sheets-only', action='store_true')
    parser.add_argument('--group-enums-only', action='store_true')
    parser.add_argument('--xml-basis-only', action='store_true')
    args = parser.parse_args()
    focused = args.explicit_required_only or args.rename_rule_sheets_only or args.group_enums_only or args.xml_basis_only
    if sum([args.explicit_required_only, args.rename_rule_sheets_only, args.group_enums_only, args.xml_basis_only]) > 1:
        parser.error('Choose only one focused upgrade mode')
    if not focused and args.references is None:
        parser.error('--references is required unless a focused upgrade mode is specified')
    if focused and args.history:
        parser.error('Focused upgrade modes do not alter history snapshots')
    args.output_dir.mkdir(parents=True, exist_ok=True)
    config_path = Path(__file__).resolve().parents[1] / 'config/TravelDocumentsBuilder/wsConfig.txt'
    config = {fields[1].split('::', 1)[1]: fields[2] for line in config_path.read_text(encoding='utf-8-sig').splitlines()
              if line.strip() for fields in [line.split('\t')] if fields[0] == 'value'}
    institutions = Package(args.institutions)
    if args.xml_basis_only:
        xml_basis_architecture(institutions, config)
    elif args.group_enums_only:
        institution_group_validation(institutions)
    elif args.rename_rule_sheets_only:
        rename_rule_sheets(institutions)
    elif args.explicit_required_only:
        explicit_reference_requirements(institutions)
    else:
        rename_rule_sheets(institutions)
        xml_basis_architecture(institutions, config)
        upgrade_references(institutions, Package(args.references))
        institution_group_validation(institutions)
    outputs = [(institutions, args.institutions)]
    if args.history:
        history = Package(args.history)
        upgrade_history(history, config)
        outputs.append((history, args.history))
    timestamp = datetime.now().strftime('%Y%m%d-%H%M%S')
    for book, original in outputs:
        output = args.output_dir / original.name
        assert output.resolve() != original.resolve(), 'Output must differ from the working workbook'
        book.save(output)
        if args.apply:
            backup = args.output_dir / (original.stem + '.' + timestamp + '.backup.xlsx')
            shutil.copy2(original, backup)
            shutil.copy2(output, original)
        print(f'Upgraded: {original.name}; tables: {len(book.tables())}')


if __name__ == '__main__':
    main()