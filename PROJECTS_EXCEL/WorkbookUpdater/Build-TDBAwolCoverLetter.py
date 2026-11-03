"""Build the empty Word scaffold; letter content comes from the DSL and wsConfig."""

from pathlib import Path

from docx import Document
from docx.shared import Mm, Pt


def build():
    projects = Path(__file__).resolve().parents[1]
    config = {}
    for line in (projects / 'config/TravelDocumentsBuilder/wsConfig.txt').read_text(encoding='utf-8-sig').splitlines():
        if line.strip():
            kind, key, value = line.split('\t')
            if kind == 'value':
                config[key] = value
    prefix = 'TravelDocumentsBuilder::'
    document_id = config[prefix + 'generation.businessTrip.companion.documentId']
    template_name = config[prefix + 'generation.tableNames.' + document_id]
    output = projects / 'workbook/TravelDocumentsBuilder' / config[prefix + 'generation.templates.' + document_id]
    document = Document()
    section = document.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.top_margin = Mm(20)
    section.bottom_margin = Mm(20)
    section.left_margin = Mm(30)
    section.right_margin = Mm(20)
    normal = document.styles['Normal']
    normal.font.name = 'Times New Roman'
    normal.font.size = Pt(14)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing = 1
    document.add_paragraph('{\\exportTable:' + template_name + '_Begin}')
    document.add_paragraph('{\\exportTable:' + template_name + '_End}')
    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(output)
    print('Created AWOL cover letter scaffold.')


if __name__ == '__main__':
    build()