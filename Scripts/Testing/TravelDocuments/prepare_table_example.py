"""Extract the existing ticket geometry into a reviewable DSL patch and a new DOCX scaffold."""

from argparse import ArgumentParser
import difflib
from pathlib import Path
import re
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree as ET


ROOT = Path(__file__).resolve().parents[3]
FOLDER = ROOT / "PROJECTS_EXCEL/workbook/TravelDocumentsBuilder"
SOURCE = FOLDER / "Шаблони/old/Відпустка (Шаблон).docx"
TARGET = FOLDER / "Шаблони/Відпустка (DSL).docx"
CATALOG = FOLDER / "VacationTicketTable.xml"
ASSET = FOLDER / "Шаблони/TicketEmblem.png"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
P = "urn:excelprototype:profiles"
NS = {"w": W}


def word_value(node, path, default=None):
    item = node.find(path, NS)
    return item.get(f"{{{W}}}val", default) if item is not None else default


def element(parent, tag, **attributes):
    return ET.SubElement(parent, f"{{{P}}}{tag}", {key: str(value) for key, value in attributes.items()})


def source_package():
    with ZipFile(SOURCE) as archive:
        package = {name: archive.read(name) for name in archive.namelist()}
    return package, ET.fromstring(package["word/document.xml"])


def catalog():
    _, document = source_package()
    original = document.find(".//w:tbl", NS)
    root = ET.Element(f"{{{P}}}wordTableTemplates", nsmap={None: P}, version="1")
    styles = element(root, "styles")
    widths = [int(node.get(f"{{{W}}}w")) for node in original.findall("w:tblGrid/w:gridCol", NS)]
    element(styles, "tableStyle", name="Ticket", width=f"{sum(widths) / 20:g}pt", alignment="left", layout="fixed", borders="none", cellPadding="0pt", cellPaddingLeft="5.4pt", cellPaddingRight="5.4pt")
    element(styles, "cellStyle", name="Body", fontName="Times New Roman", fontSize="16", horizontal="left", vertical="center", spaceBefore="0pt", spaceAfter="0pt")
    element(styles, "cellStyle", name="Centered", basedOn="Body", horizontal="center")
    element(styles, "cellStyle", name="HeaderCell", basedOn="Centered", vertical="top")
    element(styles, "cellStyle", name="Caption", basedOn="Centered", fontSize="12")
    element(styles, "cellStyle", name="Title", basedOn="Centered", fontBold="true")
    element(styles, "cellStyle", name="Stamp", basedOn="Body", fontSize="14", vertical="top")
    element(styles, "cellStyle", name="Logo", basedOn="Centered", fontSize="11", fontBold="true")
    element(styles, "cellStyle", name="LogoCode", basedOn="Centered", fontSize="9")
    element(styles, "cellStyle", name="Underlined", basedOn="Centered", borderBottom="single")
    element(styles, "cellStyle", name="UnderlinedLeft", basedOn="Body", borderBottom="single")
    template = element(root, "template", name="Ticket")
    table = element(template, "table", style="Ticket")
    columns = element(table, "columns")
    for index, width in enumerate(widths, 1):
        element(columns, "column", name=f"C{index}", width=f"{width / 20:g}pt")
    active = {}
    for index, old_row in enumerate(original.findall("w:tr", NS), 1):
        row = element(table, "row", allowBreakAcrossPages="false")
        height = old_row.find("w:trPr/w:trHeight", NS)
        if height is not None:
            row.set("height", f"{int(height.get(f'{{{W}}}val')) / 20:g}pt")
            row.set("heightRule", "exact" if height.get(f"{{{W}}}hRule") == "exact" else "atLeast")
        column = 0
        for old_cell in old_row.findall("w:tc", NS):
            span = int(word_value(old_cell, "w:tcPr/w:gridSpan", "1"))
            merge = old_cell.find("w:tcPr/w:vMerge", NS)
            if merge is not None and merge.get(f"{{{W}}}val") != "restart":
                previous = active[column]
                previous.set("rowSpan", str(int(previous.get("rowSpan", "1")) + 1))
                column += span
                continue
            cell_style = "Centered"
            if index == 1 and column in (0, 4):
                cell_style = "HeaderCell"
            if index <= 6 and column >= 9:
                cell_style = "Stamp"
            elif index in (14, 16, 18, 20, 23, 27, 29):
                cell_style = "Caption"
            elif index == 11:
                cell_style = "Title"
            elif index in (15, 19, 22, 24, 28, 30) and column == 0:
                cell_style = "Body"
            if old_cell.find("w:tcPr/w:tcBorders/w:bottom[@w:val='single']", NS) is not None:
                cell_style = "UnderlinedLeft" if cell_style == "Body" else "Underlined"
            cell = element(row, "cell", style=cell_style)
            if span > 1:
                cell.set("columnSpan", str(span))
            active[column] = cell
            if index == 1 and column == 0:
                paragraph = element(cell, "paragraph", style="Logo")
                element(paragraph, "image", source="Emblem", width="7mm", height="9mm")
                paragraph = element(cell, "paragraph", style="Logo")
                element(paragraph, "text").text = "МІНІСТЕРСТВО ОБОРОНИ{#newline}УКРАЇНИ{#newline}ЗБРОЙНІ СИЛИ УКРАЇНИ{#newline}ВІЙСЬКОВА ЧАСТИНА{#newline}А7383"
                paragraph = element(cell, "paragraph", style="LogoCode")
                element(paragraph, "text").text = "Код ЄДРПОУ 26630678"
            else:
                for old_paragraph in old_cell.findall("w:p", NS):
                    text = "".join(old_paragraph.xpath("./w:r/w:t/text()", namespaces=NS))
                    text = re.sub(r"\{([A-Za-z][A-Za-z0-9]*)\}", r"{[\1]}", text)
                    if index == 11:
                        text = text.upper()
                    paragraph = element(cell, "paragraph")
                    element(paragraph, "text").text = text
            column += span
    row = element(table, "row", allowBreakAcrossPages="false", height="12pt", heightRule="exact")
    cell = element(row, "cell", columnSpan="11", style="Body")
    element(cell, "text")
    row = element(table, "row", allowBreakAcrossPages="false")
    cell = element(row, "cell", columnSpan="11", style="Body")
    element(cell, "text").text = "Дійсний у разі пред’явлення документа, який засвідчує особу."
    optional = element(table, "if", test="[VacationAbroadText]")
    row = element(optional, "row", allowBreakAcrossPages="false")
    cell = element(row, "cell", columnSpan="11", style="Body")
    element(cell, "text").text = "{[VacationAbroadText]}"
    return format_xml(root)


def format_xml(root):
    priority = "name type dataContext itemsSource row column rowSpan columnSpan orientation style value label caption command".split()
    def render(node, depth):
        indent = "  " * depth
        tag = ET.QName(node).localname
        attributes = list(node.attrib.items())
        attributes.sort(key=lambda item: priority.index(item[0]) if item[0] in priority else len(priority))
        if depth == 0:
            attributes.insert(0, ("xmlns", P))
        start = indent + "<" + tag
        for index, (key, value) in enumerate(attributes):
            escaped = value.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;")
            start += (" " if index == 0 else "\n" + " " * (len(indent) + len(tag) + 2)) + f'{key}="{escaped}"'
        if len(node):
            return start + ">\n" + "\n".join(render(child, depth + 1) for child in node) + "\n" + indent + f"</{tag}>"
        if node.text:
            text = node.text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            return start + ">" + text + f"</{tag}>"
        return start + "/>"
    return render(root, 0)


def build_scaffold():
    package, document = source_package()
    body = document.find("w:body", NS)
    original = body.find("w:tbl", NS)
    index = body.index(original)
    body.remove(original)
    for paragraph in list(body):
        if ET.QName(paragraph).localname != "p":
            continue
        text = "".join(paragraph.xpath(".//w:t/text()", namespaces=NS))
        if text.startswith("Дійсний у разі пред’явлення документа") or text == "{VacationAbroadText}":
            body.remove(paragraph)
    for offset, marker in enumerate((r"{\exportTable:Ticket_Begin}", r"{\exportTable:Ticket_End}")):
        paragraph = ET.Element(f"{{{W}}}p")
        run = ET.SubElement(paragraph, f"{{{W}}}r")
        ET.SubElement(run, f"{{{W}}}t").text = marker
        body.insert(index + offset, paragraph)
    package["word/document.xml"] = ET.tostring(document, encoding="UTF-8", xml_declaration=True, standalone=True)
    with ZipFile(TARGET, "w", ZIP_DEFLATED) as archive:
        for name, content in package.items():
            archive.writestr(name, content)
    ASSET.write_bytes(package["word/media/image1.png"])
    print(f"Created {TARGET}")
    print(f"Extracted {ASSET}")


if __name__ == "__main__":
    parser = ArgumentParser(description=__doc__)
    parser.add_argument("--emit-patch", action="store_true")
    parser.add_argument("--build-scaffold", action="store_true")
    arguments = parser.parse_args()
    if arguments.emit_patch:
        body = catalog()
        if CATALOG.exists():
            print(f"*** Begin Patch\n*** Update File: {CATALOG.as_posix()}")
            diff = list(difflib.unified_diff(CATALOG.read_text(encoding="utf-8-sig").splitlines(), body.splitlines(), n=3))[2:]
            print("\n".join("@@" if line.startswith("@@") else line for line in diff))
        else:
            print(f"*** Begin Patch\n*** Add File: {CATALOG.as_posix()}")
            print("\n".join("+" + line for line in body.splitlines()))
        print("*** End Patch")
    if arguments.build_scaffold:
        build_scaffold()