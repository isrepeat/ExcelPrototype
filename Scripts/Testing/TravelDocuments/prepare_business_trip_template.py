"""Prepare the business-trip DSL catalog entry and its two-page Word scaffold."""

from argparse import ArgumentParser
from copy import deepcopy
import difflib
from pathlib import Path
import re
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree as ET

from prepare_table_example import element, format_xml, word_value


ROOT = Path(__file__).resolve().parents[3]
FOLDER = ROOT / 'PROJECTS_EXCEL/workbook/TravelDocumentsBuilder'
SOURCE = FOLDER / 'Шаблони/old/Відрядження (Шаблон).docx'
VACATION = FOLDER / 'Шаблони/Відпустка (DSL).docx'
TARGET = FOLDER / 'Шаблони/Відрядження (DSL).docx'
CATALOG = FOLDER / 'BusinessTripTable.xml'
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'urn:excelprototype:profiles'
NS = {'w': W, 'p': P}


def read_package(path):
    with ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def source_text(node):
    return ''.join(node.xpath('.//w:t/text()', namespaces=NS))


def trip_template():
    document = ET.fromstring(read_package(SOURCE)['word/document.xml'])
    original = document.find('.//w:tbl', NS)
    template = ET.Element(f'{{{P}}}template', name='BusinessTrip')
    table = element(template, 'table', style='Ticket')
    columns = element(table, 'columns')
    for index, item in enumerate(original.findall('w:tblGrid/w:gridCol', NS), 1):
        element(columns, 'column', name=f'C{index}', width=f'{int(item.get(f"{{{W}}}w")) / 20:g}pt')
    active = {}
    logo = ET.parse(str(CATALOG)).find('.//p:template[@name="BusinessTrip"]/p:table/p:row/p:cell', NS)
    for index, old_row in enumerate(original.findall('w:tr', NS)):
        row = element(table, 'row', allowBreakAcrossPages='false')
        height = old_row.find('w:trPr/w:trHeight', NS)
        if height is not None:
            row.set('height', f'{int(height.get(f"{{{W}}}val")) / 20:g}pt')
            row.set('heightRule', 'atLeast')
        column = 0
        for old_cell in old_row.findall('w:tc', NS):
            span = int(word_value(old_cell, 'w:tcPr/w:gridSpan', '1'))
            merge = old_cell.find('w:tcPr/w:vMerge', NS)
            if merge is not None and merge.get(f'{{{W}}}val') != 'restart':
                previous = active[column]
                previous.set('rowSpan', str(int(previous.get('rowSpan', '1')) + 1))
                column += span
                continue
            style = 'Centered'
            if index < 6:
                style = 'Stamp' if column == 5 else 'HeaderCell'
            elif index in (6, 7, 8):
                style = 'TicketDetails'
            elif index == 10:
                style = 'Title'
            elif index in (13, 15, 17, 19, 22):
                style = 'Caption'
            elif index == 24 or (index in (12, 16) and column == 0):
                style = 'Body'
            if old_cell.find("w:tcPr/w:tcBorders/w:bottom[@w:val='single']", NS) is not None:
                style = 'UnderlinedLeft' if style == 'Body' else 'Underlined'
            cell = element(row, 'cell', style=style)
            if span > 1:
                cell.set('columnSpan', str(span))
            active[column] = cell
            if index == 0 and column == 0:
                for paragraph in logo:
                    cell.append(deepcopy(paragraph))
            else:
                for old_paragraph in old_cell.findall('w:p', NS):
                    text = source_text(old_paragraph)
                    text = re.sub(r'<([A-Za-z][A-Za-z0-9]*)></\1>', r'{[\1]}', text)
                    if index == 10:
                        text = text.upper()
                    elif index == 21:
                        text = '{[Purpose]}'
                    elif index == 24:
                        text = 'Підстава відрядження: {[Basis]}'
                    text = text.replace('Термін відрядження:', 'Термін відрядження: ')
                    paragraph = element(cell, 'paragraph')
                    element(paragraph, 'text').text = text
            column += span
    spacer = element(table, 'row', allowBreakAcrossPages='false', height='12pt', heightRule='exact')
    cell = element(spacer, 'cell', columnSpan='6', style='Body')
    element(cell, 'text')
    validity = element(table, 'row', allowBreakAcrossPages='false')
    cell = element(validity, 'cell', columnSpan='6', style='Body')
    element(cell, 'text').text = 'Дійсно в разі пред’явлення документа, який засвідчує особу.'
    optional = element(table, 'if', test='[WeaponLine]')
    row = element(optional, 'row', allowBreakAcrossPages='false')
    cell = element(row, 'cell', columnSpan='6', style='Body')
    element(cell, 'text').text = '{[WeaponLine]}'
    return template


def emit_patch():
    original = CATALOG.read_text(encoding='utf-8-sig')
    catalog = ET.fromstring(original.encode('utf-8'))
    for template in catalog.findall('p:template[@name="BusinessTrip"]', NS):
        catalog.remove(template)
    catalog.append(trip_template())
    updated = format_xml(catalog)
    print(f'*** Begin Patch\n*** Update File: {CATALOG.as_posix()}')
    diff = list(difflib.unified_diff(original.splitlines(), updated.splitlines(), n=3))[2:]
    print('\n'.join('@@' if line.startswith('@@') else line for line in diff))
    print('*** End Patch')


def build_scaffold():
    package = read_package(SOURCE)
    vacation = read_package(VACATION)
    document = ET.fromstring(package['word/document.xml'])
    body = document.find('w:body', NS)
    original = list(body)
    # Replace the generated front table and move its signature to the first-page footer.
    for child in original[:15]:
        body.remove(child)
    for index, marker in enumerate((r'{\exportTable:BusinessTrip_Begin}', r'{\exportTable:BusinessTrip_End}')):
        paragraph = ET.Element(f'{{{W}}}p')
        run = ET.SubElement(paragraph, f'{{{W}}}r')
        ET.SubElement(run, f'{{{W}}}t').text = marker
        body.insert(index, paragraph)
    continuation = original[15]
    properties = continuation.find('w:pPr', NS)
    if properties is None:
        properties = ET.Element(f'{{{W}}}pPr')
        continuation.insert(0, properties)
    ET.SubElement(properties, f'{{{W}}}pageBreakBefore')
    # Reuse the exact vacation page geometry and first-page signature footer.
    old_section = body.find('w:sectPr', NS)
    body.remove(old_section)
    section = deepcopy(ET.fromstring(vacation['word/document.xml']).find('.//w:sectPr', NS))
    section.find('w:footerReference', NS).set(f'{{{R}}}id', 'rIdTripFooter')
    body.append(section)
    package['word/footerTrip.xml'] = vacation['word/footer1.xml']
    relationships = ET.fromstring(package['word/_rels/document.xml.rels'])
    ET.SubElement(relationships, '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship',
                  Id='rIdTripFooter', Type=R + '/footer', Target='footerTrip.xml')
    package['word/_rels/document.xml.rels'] = ET.tostring(relationships, encoding='UTF-8', xml_declaration=True, standalone=True)
    types = ET.fromstring(package['[Content_Types].xml'])
    ET.SubElement(types, '{http://schemas.openxmlformats.org/package/2006/content-types}Override',
                  PartName='/word/footerTrip.xml', ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml')
    package['[Content_Types].xml'] = ET.tostring(types, encoding='UTF-8', xml_declaration=True, standalone=True)
    package['word/document.xml'] = ET.tostring(document, encoding='UTF-8', xml_declaration=True, standalone=True)
    with ZipFile(TARGET, 'w', ZIP_DEFLATED) as archive:
        for name, content in package.items():
            archive.writestr(name, content)
    print(TARGET)


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--emit-patch', action='store_true')
    parser.add_argument('--build-scaffold', action='store_true')
    arguments = parser.parse_args()
    if arguments.emit_patch:
        emit_patch()
    if arguments.build_scaffold:
        build_scaffold()