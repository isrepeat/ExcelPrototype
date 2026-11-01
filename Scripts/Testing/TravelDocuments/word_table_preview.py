"""Static table-layout checks and an OOXML preview; this does not execute the VBA renderer."""

import base64
from pathlib import Path
import re
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree as ET

P = "urn:excelprototype:profiles"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"p": P, "w": W}


def points(value):
    match = re.fullmatch(r"(\d+(?:\.\d+)?)(mm|cm|pt)", value)
    if not match:
        raise ValueError(f"Invalid length: {value}")
    return float(match[1]) * {"mm": 72 / 25.4, "cm": 72 / 2.54, "pt": 1}[match[2]]


def styles(root):
    definitions = {}
    for node in root.findall("p:styles/*", NS):
        if node.get("name") in definitions:
            raise ValueError("Duplicate style")
        definitions[node.get("name")] = node

    def resolve(name, chain=()):
        if name in chain or name not in definitions:
            raise ValueError(f"Missing style or inheritance cycle: {name}")
        node = definitions[name]
        result = resolve(node.get("basedOn"), chain + (name,)) if node.get("basedOn") else {}
        if node.get("basedOn") and definitions[node.get("basedOn")].tag != node.tag:
            raise ValueError("Style inheritance kind mismatch")
        result.update({key: value for key, value in node.attrib.items() if key not in ("name", "basedOn")})
        return result

    return {name: resolve(name) for name in definitions}


def expand(parent, context, limit=512, depth=0):
    if depth > 32:
        raise ValueError("Expansion depth exceeded")
    rows = []
    for node in parent:
        tag = ET.QName(node).localname
        if tag == "row":
            rows.append((node, context))
        elif tag == "if":
            expression = node.get("test")
            if not re.fullmatch(r"\[\w+\]", expression):
                raise ValueError("The offline layout checker accepts a single-field condition")
            if context[expression[1:-1]]:
                rows.extend(expand(node, context, limit, depth + 1))
        elif tag == "forEach":
            items = context[node.get("itemsSource")]
            for index, item in enumerate(items, 1):
                alias = node.get("item")
                child = {key: value for key, value in context.items() if not key.startswith(alias + ".")}
                child.update({alias + "." + key: value for key, value in item.items()})
                child.update({alias + ".__index": index, alias + ".__first": index == 1, alias + ".__last": index == len(items)})
                rows.extend(expand(node, child, limit, depth + 1))
        if len(rows) > limit:
            raise ValueError("Row limit exceeded")
    return rows


def grid(table, context):
    columns = [points(node.get("width")) for node in table.findall("p:columns/p:column", NS)]
    if not columns or len(columns) > 63 or any(width <= 0 for width in columns):
        raise ValueError("Invalid columns")
    rows = expand(table, context)
    if not rows:
        raise ValueError("Empty table")
    occupied = {}
    cells = []
    for row_index, (row, values) in enumerate(rows):
        column = 0
        for node in row:
            while (row_index, column) in occupied:
                column += 1
            column_span = int(node.get("columnSpan", "1"))
            row_span = int(node.get("rowSpan", "1"))
            if column_span < 1 or row_span < 1 or column + column_span > len(columns) or row_index + row_span > len(rows):
                raise ValueError("Cell span exceeds grid")
            record = (node, values, row_index, column, row_span, column_span)
            for r in range(row_index, row_index + row_span):
                if rows[r][0].get("header", "false") != row.get("header", "false"):
                    raise ValueError("Merge crosses header boundary")
                for c in range(column, column + column_span):
                    if (r, c) in occupied:
                        raise ValueError("Overlapping cells")
                    occupied[r, c] = record
            cells.append(record)
            column += column_span
        if any((row_index, c) not in occupied for c in range(len(columns))):
            raise ValueError("Incomplete row")
    return columns, rows, cells, occupied


def marker_bounds(body, name):
    markers = [f"{{\\exportTable:{name}_{suffix}}}" for suffix in ("Begin", "End")]
    positions = []
    for marker in markers:
        all_matches = body.xpath(".//w:p[.//w:t]", namespaces=NS)
        matches = [node for node in all_matches if marker in "".join(node.xpath(".//w:t/text()", namespaces=NS))]
        if len(matches) != 1 or matches[0].getparent() is not body:
            raise ValueError("Markers must be unique top-level paragraphs")
        if "".join(matches[0].xpath(".//w:t/text()", namespaces=NS)) != marker:
            raise ValueError("Marker must occupy the paragraph")
        positions.append(body.index(matches[0]))
    if positions[0] >= positions[1]:
        raise ValueError("Reversed markers")
    return positions


def node(parent, tag, **attributes):
    return ET.SubElement(parent, f"{{{W}}}{tag}", {f"{{{W}}}{key}": str(value) for key, value in attributes.items()})


def render_text(text, context):
    text = text.replace("{#newline}", "\n").replace("{#space}", " ")
    text = re.sub(r"\{\[([^]]+)\]\}", lambda match: str(context[match[1]]), text)
    text = re.sub(r"\{(\w+)\.\[([^]]+)\]\}", lambda match: str(context[match[1] + "." + match[2]]), text)
    text = re.sub(r"\{(\w+\.__\w+)\}", lambda match: str(context[match[1]]), text)
    if "{" in text or "}" in text:
        raise ValueError("The offline preview does not implement the full text DSL")
    return text


def preview(catalog_path, scaffold_path, output_path, context):
    root = ET.parse(str(catalog_path)).getroot()
    catalog_styles = styles(root)
    template = root.find("p:template", NS)
    table = template.find("p:table", NS)
    columns, rows, cells, occupied = grid(table, context)
    table_style = catalog_styles[table.get("style")]
    if abs(points(table_style["width"]) - sum(columns)) > 0.1:
        raise ValueError("Table width differs from grid")
    with ZipFile(scaffold_path) as archive:
        package = {name: archive.read(name) for name in archive.namelist()}
    document = ET.fromstring(package["word/document.xml"])
    body = document.find("w:body", NS)
    begin, end = marker_bounds(body, template.get("name"))
    with ZipFile(scaffold_path.parent / "Відпустка (Шаблон).docx") as original:
        drawing_template = ET.fromstring(original.read("word/document.xml"))
        image_id = drawing_template.xpath(".//*[local-name()='blip']/@*[local-name()='embed']")[0]
    result = ET.Element(f"{{{W}}}tbl")
    properties = node(result, "tblPr")
    node(properties, "tblW", w=round(sum(columns) * 20), type="dxa")
    node(properties, "jc", val=table_style.get("alignment", "left"))
    borders = node(properties, "tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node(borders, side, val="nil")
    node(properties, "tblLayout", type="fixed")
    margins = node(properties, "tblCellMar")
    for side in ("top", "left", "bottom", "right"):
        padding = table_style.get("cellPadding" + side.title(), table_style.get("cellPadding", "0pt"))
        node(margins, side, w=round(points(padding) * 20), type="dxa")
    table_grid = node(result, "tblGrid")
    for width in columns:
        node(table_grid, "gridCol", w=round(width * 20))
    for row_index, (source_row, _) in enumerate(rows):
        row = node(result, "tr")
        pr = node(row, "trPr")
        if source_row.get("allowBreakAcrossPages", "false") == "false":
            node(pr, "cantSplit")
        if source_row.get("header") == "true":
            node(pr, "tblHeader")
        if source_row.get("height"):
            node(pr, "trHeight", val=round(points(source_row.get("height")) * 20), hRule=source_row.get("heightRule"))
        column = 0
        while column < len(columns):
            cell, values, start_row, start_column, row_span, column_span = occupied[row_index, column]
            tc = node(row, "tc")
            pr = node(tc, "tcPr")
            node(pr, "tcW", w=round(sum(columns[column:column + column_span]) * 20), type="dxa")
            if column_span > 1:
                node(pr, "gridSpan", val=column_span)
            if row_span > 1:
                node(pr, "vMerge", **({"val": "restart"} if start_row == row_index else {}))
            cell_style = catalog_styles.get(cell.get("style"), {})
            cell_borders = node(pr, "tcBorders")
            for side in ("top", "left", "bottom", "right"):
                kind = cell_style.get("borderBottom", table_style.get("borders", "none")) if side == "bottom" else table_style.get("borders", "none")
                node(cell_borders, side, **({"val": "single", "sz": 6, "color": "000000"} if kind == "single" else {"val": "nil"}))
            node(pr, "vAlign", val=cell_style.get("vertical", "center"))
            if start_row != row_index:
                node(tc, "p")
            else:
                for paragraph in cell:
                    style = dict(cell_style)
                    style.update(catalog_styles.get(paragraph.get("style"), {}))
                    image = paragraph.find("p:image", NS)
                    text_node = paragraph if ET.QName(paragraph).localname == "text" else paragraph.find("p:text", NS)
                    text = render_text(text_node.text or "", values) if text_node is not None else ""
                    for line in text.split("\n"):
                        p = node(tc, "p")
                        ppr = node(p, "pPr")
                        node(ppr, "spacing", before=0, after=0)
                        node(ppr, "jc", val=style.get("horizontal", "left"))
                        run = node(p, "r")
                        rpr = node(run, "rPr")
                        node(rpr, "rFonts", ascii=style.get("fontName", "Times New Roman"), hAnsi=style.get("fontName", "Times New Roman"))
                        node(rpr, "sz", val=round(float(style.get("fontSize", 12)) * 2))
                        for field, tag in (("fontBold", "b"), ("fontItalic", "i")):
                            if field in style:
                                node(rpr, tag, val="1" if style[field] == "true" else "0")
                        if image is not None:
                            cx = round(points(image.get("width")) * 12700)
                            cy = round(points(image.get("height")) * 12700)
                            inline = ET.fromstring(f"""<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                                <wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="100" name="Emblem"/><wp:cNvGraphicFramePr/>
                                <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>
                                <pic:nvPicPr><pic:cNvPr id="0" name="Emblem"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="{image_id}"/>
                                <a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/>
                                </a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline>""")
                            node(run, "drawing").append(inline)
                        else:
                            t = node(run, "t")
                            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                            t.text = line
            column += column_span
    for item in list(body)[begin:end + 1]:
        body.remove(item)
    body.insert(begin, result)
    package["word/document.xml"] = ET.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True)
    for name in list(package):
        if name == "word/document.xml" or re.fullmatch(r"word/(header|footer)\d+\.xml", name):
            story = ET.fromstring(package[name])
            for paragraph in story.findall(".//w:p", NS):
                texts = paragraph.findall(".//w:t", NS)
                for key, value in context.items():
                    marker = "{" + key + "}"
                    joined = "".join(item.text or "" for item in texts)
                    while marker in joined:
                        start = joined.index(marker)
                        end = start + len(marker)
                        offset = 0
                        written = False
                        for item in texts:
                            content = item.text or ""
                            limit = offset + len(content)
                            if offset < end and limit > start:
                                item.text = content[:max(0, start - offset)] + (str(value) if not written else "") + content[max(0, end - offset):]
                                written = True
                            offset = limit
                        joined = "".join(item.text or "" for item in texts)
            package[name] = ET.tostring(story, xml_declaration=True, encoding="UTF-8", standalone=True)
    with ZipFile(output_path, "w", ZIP_DEFLATED) as archive:
        for name, content in package.items():
            archive.writestr(name, content)
    return output_path


def flat_opc(preview_path):
    package_ns = "http://schemas.microsoft.com/office/2006/xmlPackage"
    relationship_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    package = ET.Element(f"{{{package_ns}}}package", nsmap={"pkg": package_ns})
    def part(name, content_type, content=None):
        item = ET.SubElement(package, f"{{{package_ns}}}part", {f"{{{package_ns}}}name": name, f"{{{package_ns}}}contentType": content_type})
        data = ET.SubElement(item, f"{{{package_ns}}}xmlData" if content is not None else f"{{{package_ns}}}binaryData")
        if content is not None:
            data.append(content)
        return data
    relationships = ET.Element(f"{{{relationship_ns}}}Relationships")
    ET.SubElement(relationships, f"{{{relationship_ns}}}Relationship", Id="rId1", Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", Target="word/document.xml")
    part("/_rels/.rels", "application/vnd.openxmlformats-package.relationships+xml", relationships)
    with ZipFile(preview_path) as archive:
        source = ET.fromstring(archive.read("word/document.xml"))
        document = ET.Element(f"{{{W}}}document", nsmap={"w": W})
        body = node(document, "body")
        body.append(source.find(".//w:tbl", NS))
        node(body, "p")
        part("/word/document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", document)
        ids = set(document.xpath(".//*[local-name()='blip']/@*[local-name()='embed']"))
        relationships = ET.Element(f"{{{relationship_ns}}}Relationships")
        for original in ET.fromstring(archive.read("word/_rels/document.xml.rels")):
            if original.get("Id") in ids:
                relationships.append(original)
                target = original.get("Target")
                data = part("/word/" + target, "image/png")
                data.text = base64.b64encode(archive.read("word/" + target)).decode("ascii")
        part("/word/_rels/document.xml.rels", "application/vnd.openxmlformats-package.relationships+xml", relationships)
    return ET.tostring(package, encoding="unicode")


if __name__ == "__main__":
    import sys
    from prepare_table_example import CATALOG, TARGET

    if sys.argv[1] == "--flat-opc":
        print(flat_opc(Path(sys.argv[2])))
        raise SystemExit(0)
    output = Path(sys.argv[1])
    sample = {
        "TicketDate": '"01" листопада 2026 р.', "TicketNum": "2026/309/1",
        "PersonalLine": "старший солдат ПРИКЛАД Олександр Іванович",
        "VacationKind": "у щорічну основну відпустку", "VacationPlace": "м. Київ",
        "VacationDuration": "15 календарних днів", "DateFrom": '"02" листопада 2026 р.',
        "DateTo": '"16" листопада 2026 р.', "PersonalInitials": "старший солдат ПРИКЛАД О.І.",
        "DateArrival": '"17" листопада 2026 р.', "VacationAbroadText": "", "FooterSpacer": "",
        "SignatoryPosition": "Командир військової частини А7383", "SignatoryRank": "полковник",
        "SignatoryName": "ПРИКЛАД Іван Іванович",
    }
    print(preview(CATALOG, TARGET, output, sample))