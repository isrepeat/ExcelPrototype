"""Check the business-trip catalog, source preservation and first-page footer offline."""

from hashlib import sha256
from pathlib import Path
import re
from tempfile import TemporaryDirectory

from lxml import etree as ET

from prepare_business_trip_template import CATALOG, SOURCE, TARGET, VACATION, NS, read_package, source_text
from word_table_preview import grid, marker_bounds, preview, styles


assert sha256(SOURCE.read_bytes()).hexdigest() == '27489141aee906e15351d32cf016c360603980fad87031235ca8a9dfbf93a054'
catalog = ET.parse(str(CATALOG)).getroot()
template = catalog.find('p:template[@name="BusinessTrip"]', NS)
assert template is not None
table = template.find('p:table', NS)
fields = set(re.findall(r'\{\[([^]]+)\]\}', ''.join(table.itertext())))
assert fields == {'TicketDate', 'TicketNum', 'PersonalLine', 'Position', 'TravelPlace', 'Duration', 'Basis', 'WeaponLine'}
context = {field: field for field in fields}
columns, rows, cells, occupied = grid(table, context)
assert len(columns) == 6 and len(rows) == 24
empty_weapon = dict(context, WeaponLine='')
assert len(grid(table, empty_weapon)[1]) == 23
assert table.find('p:if[@test="[WeaponLine]"]/p:row/p:cell/p:text', NS).text == 'Слідує зі зброєю: {[WeaponLine]}.'
assert abs(sum(columns) - 474.95) < 0.01
assert styles(catalog)[table.get('style')]['borders'] == 'none'
assert '16.09.2026' not in ''.join(table.itertext())
assert 'Додаток 29' in ''.join(table.itertext())

original = read_package(SOURCE)
scaffold = read_package(TARGET)
vacation = read_package(VACATION)
changed = {name for name in original if original[name] != scaffold[name]}
assert changed == {'word/document.xml', 'word/_rels/document.xml.rels', '[Content_Types].xml'}
assert set(scaffold) - set(original) == {'word/footerTrip.xml'}
assert scaffold['word/footerTrip.xml'] == vacation['word/footer1.xml']
source_body = ET.fromstring(original['word/document.xml']).find('w:body', NS)
body = ET.fromstring(scaffold['word/document.xml']).find('w:body', NS)
assert marker_bounds(body, 'BusinessTrip') == [0, 1]
assert ET.tostring(body.find('w:tbl', NS)) == ET.tostring(source_body.findall('w:tbl', NS)[1])
continuation = next(node for node in body if source_text(node) == 'Продовження додатка 29')
assert continuation.find('w:pPr/w:pageBreakBefore', NS) is not None
assert body.find('w:sectPr/w:titlePg', NS) is not None
assert body.find('w:sectPr/w:footerReference', NS).get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id') == 'rIdTripFooter'
assert 'Олег ЧЕРКАШИН' not in source_text(body)
assert '<WeaponLine>' not in source_text(body)
assert '{WeaponLine}' not in source_text(body)

context.update(WeaponLine='АК-47', SignatoryPosition='Командир військової частини А7383',
               SignatoryRank='полковник', SignatoryName='Іван ПРИКЛАД')
with TemporaryDirectory() as folder:
    path = preview(CATALOG, TARGET, Path(folder) / 'trip.docx', context, 'BusinessTrip')
    package = read_package(path)
    output = ET.fromstring(package['word/document.xml'])
    assert len(output.findall('.//w:body/w:tbl', NS)) == 2
    assert '\\exportTable:' not in source_text(output)
    assert 'Слідує зі зброєю: АК-47.' in source_text(output)
    footer = source_text(ET.fromstring(package['word/footerTrip.xml']))
    assert 'Іван ПРИКЛАД' in footer and '{Signatory' not in footer
    empty_path = preview(CATALOG, TARGET, Path(folder) / 'empty.docx', dict(context, WeaponLine=''), 'BusinessTrip')
    empty_document = ET.fromstring(read_package(empty_path)['word/document.xml'])
    assert 'Слідує зі зброєю:' not in source_text(empty_document)

print('Business-trip grid, fields, source preservation, footer and page break: OK')