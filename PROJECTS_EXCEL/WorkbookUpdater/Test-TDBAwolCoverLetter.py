"""Validate the AWOL letter assets and configuration without executing VBA."""

from decimal import Decimal
from pathlib import Path
import re
import unittest
import xml.etree.ElementTree as ET
import zipfile


ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / 'workbook/TravelDocumentsBuilder'
NS = {'p': 'urn:excelprototype:profiles', 'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
PREFIX = 'generation.businessTrip.companion.'


def configuration():
    result = {}
    for line in (ROOT / 'config/TravelDocumentsBuilder/wsConfig.txt').read_text(encoding='utf-8-sig').splitlines():
        if line.strip():
            kind, key, value = line.split('\t')
            if kind == 'value':
                key = key.removeprefix('TravelDocumentsBuilder::')
                assert key not in result, key
                result[key] = value
    return result


class AwolCoverLetterTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.config = configuration()
        cls.document_id = cls.config[PREFIX + 'documentId']
        cls.catalog = ET.parse(WORKBOOK / cls.config['generation.tablePaths.' + cls.document_id]).getroot()
        cls.template_name = cls.config['generation.tableNames.' + cls.document_id]

    def test_scaffold_has_exact_markers_and_no_macros(self):
        scaffold = WORKBOOK / self.config['generation.templates.' + self.document_id]
        with zipfile.ZipFile(scaffold) as package:
            self.assertFalse(any('vbaProject' in name for name in package.namelist()))
            document = ET.fromstring(package.read('word/document.xml'))
        paragraphs = [''.join(p.itertext()) for p in document.findall('w:body/w:p', NS)]
        self.assertEqual(paragraphs, [
            '{\\exportTable:' + self.template_name + '_Begin}',
            '{\\exportTable:' + self.template_name + '_End}',
        ])
        page = document.find('w:body/w:sectPr/w:pgSz', NS)
        self.assertEqual(page.get('{' + NS['w'] + '}w'), '11906')
        self.assertEqual(page.get('{' + NS['w'] + '}h'), '16838')

    def test_dsl_fields_are_supplied(self):
        fields = {'CompanionDestinationUnit', 'CompanionHasCopy', 'CompanionSenderInstitution',
                  'CompanionNotification', 'CompanionRank', 'CompanionName',
                  'CompanionTicketNumber', 'CompanionOrderDate', 'TicketDateDay',
                  'TicketDateMonth', 'TicketDateYear',
                  'SignatoryPosition', 'SignatoryRank', 'SignatoryName',
                  'SignatoryHasUnit', 'SignatoryUnitSuffix'}
        referenced = set(re.findall(r'\[(\w+)\]', ''.join(self.catalog.itertext())))
        self.assertEqual(referenced, fields)
        for type_id in self.config[PREFIX + 'notificationTypes'].split('|'):
            self.assertIn('basisChain.model.Documents.' + type_id + '.Name', self.config)

    def test_table_grid_styles_and_images(self):
        template = self.catalog.find("p:template[@name='" + self.template_name + "']", NS)
        table = template.find('p:table', NS)
        styles = {node.get('name') for node in self.catalog.find('p:styles', NS)}
        columns = table.find('p:columns', NS)
        self.assertEqual(sum(Decimal(node.get('width').removesuffix('pt')) for node in columns), Decimal('474.95'))
        spans = [0] * len(columns)
        for row in table.findall('p:row', NS):
            index = 0
            for cell in row:
                while index < len(spans) and spans[index]:
                    index += 1
                end = index + int(cell.get('columnSpan', '1'))
                self.assertLessEqual(end, len(spans))
                self.assertFalse(any(spans[index:end]))
                spans[index:end] = [int(cell.get('rowSpan', '1'))] * (end - index)
                index = end
            self.assertTrue(all(spans))
            spans = [span - 1 for span in spans]
        self.assertFalse(any(spans))
        for node in self.catalog.iter():
            for attribute in ['style', 'basedOn']:
                if attribute in node.attrib:
                    self.assertIn(node.get(attribute), styles)
        for image in self.catalog.findall('.//p:image', NS):
            self.assertTrue((WORKBOOK / self.config['word.tables.images.' + image.get('source')]).is_file())

    def test_both_filenames_are_managed_and_distinct(self):
        identifier = '12345678-1234-1234-1234-123456789ABC'
        pattern = re.compile(self.config['word.filename.pattern'], re.IGNORECASE)
        names = []
        for document_id in ['BusinessTripCertificate', self.document_id]:
            name = self.config['word.filename.prefix.' + document_id] + '_sample__' + identifier + '.docx'
            self.assertEqual(pattern.fullmatch(name).group(1), identifier)
            names.append(name)
        self.assertNotEqual(*names)


if __name__ == '__main__':
    unittest.main()