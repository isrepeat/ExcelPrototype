"""Check the business-trip catalog, source preservation and signature rows offline."""

from hashlib import sha256
import re

from lxml import etree as ET

from prepare_business_trip_template import CATALOG, SOURCE, TARGET, VACATION, NS, read_package, source_text
from word_table_preview import grid, marker_bounds, styles


assert sha256(SOURCE.read_bytes()).hexdigest() == '27489141aee906e15351d32cf016c360603980fad87031235ca8a9dfbf93a054'
catalog = ET.parse(str(CATALOG)).getroot()
template = catalog.find('p:template[@name="BusinessTrip"]', NS)
assert template is not None
table = template.find('p:table', NS)
fields = set(re.findall(r'\{\[([^]]+)\]\}', ''.join(table.itertext())))
assert {'TicketDateDay', 'TicketNum', 'RankDative', 'FIODative', 'TravelPlace', 'TripDays', 'Basis', 'WeaponLine',
        'SignatoryPosition', 'SignatoryRank', 'SignatoryName'} <= fields
context = dict(IsAwolTrip=False, WeaponLine='АК-47')
columns, rows, cells, occupied = grid(table, context)
assert len(columns) == 6 and len(rows) == 29
empty_weapon = dict(context, WeaponLine='')
assert len(grid(table, empty_weapon)[1]) == 28
for index, field in ((-4, 'SignatoryPosition'), (-3, 'SignatoryRank'), (-3, 'SignatoryName')):
    assert '{[' + field + ']}' in ''.join(rows[index][0].itertext())
assert styles(catalog)['SignatureName']['horizontal'] == 'right'
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

footer = ET.fromstring(scaffold['word/footerTrip.xml'])
assert not footer.findall('.//w:tbl', NS)
assert not source_text(footer).strip()

print('Business-trip grid, signature rows, source preservation, footer and page break: OK')