"""Чтение фикстур и сохранение неизменённых частей OOXML после редактирования."""

import copy
import datetime
import json
import posixpath
import sys
import zipfile
from lxml import etree as ET

import openpyxl
from openpyxl.utils.cell import get_column_letter
from openpyxl.utils.datetime import to_excel

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
DOCREL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"m": MAIN, "r": REL}


def dump(path):
    book = openpyxl.load_workbook(path, data_only=False)
    result = {}
    for sheet in book:
        rows = []
        for row in sheet.values:
            rows.append([
                to_excel(value, book.epoch) if isinstance(value, (datetime.datetime, datetime.date)) else value
                for value in row
            ])
        result[sheet.title] = {
            "rows": rows,
            "tables": [{"name": table.name, "ref": table.ref, "style": table.tableStyleInfo.name}
                       for table in sheet.tables.values()],
        }
    print(json.dumps(result, ensure_ascii=False))


def xml(data, name):
    return ET.fromstring(data[name])


def serialize(node):
    return ET.tostring(node, encoding="utf-8", xml_declaration=True)


def rel_path(part):
    folder, name = posixpath.split(part)
    return f"{folder}/_rels/{name}.rels"


def target(part, value):
    return value.lstrip("/") if value.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(part), value))


def sheet_parts(data):
    relationships = {item.get("Id"): target("xl/workbook.xml", item.get("Target"))
                     for item in xml(data, "xl/_rels/workbook.xml.rels")}
    return {item.get("name"): relationships[item.get(f"{{{DOCREL}}}id")]
            for item in xml(data, "xl/workbook.xml").find("m:sheets", NS)}


def merge_styles(original, generated):
    base = xml(original, "xl/styles.xml")
    added = xml(generated, "xl/styles.xml")
    offsets = {}
    for tag in ("fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "dxfs"):
        node = base.find(f"m:{tag}", NS)
        if node is None:
            node = ET.SubElement(base, f"{{{MAIN}}}{tag}")
        offsets[tag] = len(node)
    formats = base.find("m:numFmts", NS)
    if formats is None:
        formats = ET.Element(f"{{{MAIN}}}numFmts")
        base.insert(0, formats)
    existing_formats = {item.get("formatCode"): int(item.get("numFmtId")) for item in formats}
    next_format = max([163] + list(existing_formats.values())) + 1
    format_ids = {}
    for item in added.findall("m:numFmts/m:numFmt", NS):
        code = item.get("formatCode")
        old_id = int(item.get("numFmtId"))
        if code not in existing_formats:
            clone = copy.deepcopy(item)
            clone.set("numFmtId", str(next_format))
            existing_formats[code] = next_format
            formats.append(clone)
            next_format += 1
        format_ids[old_id] = existing_formats[code]
    formats.set("count", str(len(formats)))
    for tag in offsets:
        node = base.find(f"m:{tag}", NS)
        for item in added.findall(f"m:{tag}/*", NS):
            clone = copy.deepcopy(item)
            if tag in ("cellStyleXfs", "cellXfs"):
                for attribute, collection in (("fontId", "fonts"), ("fillId", "fills"), ("borderId", "borders")):
                    if clone.get(attribute) is not None:
                        clone.set(attribute, str(int(clone.get(attribute)) + offsets[collection]))
                if clone.get("xfId") is not None:
                    clone.set("xfId", str(int(clone.get("xfId")) + offsets["cellStyleXfs"]))
                if clone.get("numFmtId") is not None:
                    old_id = int(clone.get("numFmtId"))
                    clone.set("numFmtId", str(format_ids.get(old_id, old_id)))
            node.append(clone)
        node.set("count", str(len(node)))
    original["xl/styles.xml"] = serialize(base)
    return offsets


def column_index(address):
    index = 0
    for character in address:
        if not character.isalpha():
            break
        index = index * 26 + ord(character) - ord("A") + 1
    return index


def preserve_existing_cells(original, destination, edited, name):
    baseline = xml(original, destination)
    data = baseline.find("m:sheetData", NS)
    new_data = edited.find("m:sheetData", NS)
    old_rows = {int(row.get("r")): row for row in data}
    if name in ("Тест_Сценарії", "Тест_ЄЖОС"):
        last_row = max(old_rows)
        for row in new_data:
            if int(row.get("r")) > last_row:
                data.append(copy.deepcopy(row))
    else:
        step_column = 3 if name == "Тест_ВхідніДані" else 4
        for row in new_data:
            index = int(row.get("r"))
            for cell in row:
                column = column_index(cell.get("r"))
                if column != step_column and not (name == "Тест_Результати" and column > 22):
                    continue
                if index == 1 and column == step_column:
                    continue
                if index not in old_rows:
                    old_rows[index] = ET.SubElement(data, f"{{{MAIN}}}row", {"r": str(index)})
                target_row = old_rows[index]
                existing = next((item for item in target_row if item.get("r") == cell.get("r")), None)
                if existing is not None:
                    target_row.remove(existing)
                target_row.append(copy.deepcopy(cell))
        for row in data:
            row[:] = sorted(row, key=lambda cell: column_index(cell.get("r")))
    cells = data.findall("m:row/m:c", NS)
    last_column = max(column_index(cell.get("r")) for cell in cells)
    last_row = max(int(row.get("r")) for row in data)
    baseline.find("m:dimension", NS).set("ref", f"A1:{get_column_letter(last_column)}{last_row}")
    if name == "Тест_Результати":
        columns = baseline.find("m:cols", NS)
        for column in edited.findall("m:cols/m:col", NS):
            if int(column.get("min")) > 22:
                columns.append(copy.deepcopy(column))
    old_parts = baseline.find("m:tableParts", NS)
    old_parts[:] = [copy.deepcopy(item) for item in edited.find("m:tableParts", NS)]
    return baseline


def assemble(original_path, generated_path, output_path, names):
    with zipfile.ZipFile(original_path) as archive:
        original = {name: archive.read(name) for name in archive.namelist()}
    with zipfile.ZipFile(generated_path) as archive:
        generated = {name: archive.read(name) for name in archive.namelist()}
    baseline = dict(original)
    original_sheets = sheet_parts(original)
    generated_sheets = sheet_parts(generated)
    offsets = merge_styles(original, generated)
    strings = xml(generated, "xl/sharedStrings.xml").findall("m:si", NS) if "xl/sharedStrings.xml" in generated else []
    book = xml(original, "xl/workbook.xml")
    sheets = book.find("m:sheets", NS)
    relationships = xml(original, "xl/_rels/workbook.xml.rels")
    types = xml(original, "[Content_Types].xml")
    sheet_id = max(int(item.get("sheetId")) for item in sheets) + 1
    table_id = max(int(xml(original, name).get("id")) for name in original
                   if name.startswith("xl/tables/") and name.endswith(".xml")) + 1
    modified = {"xl/styles.xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml"}
    insertion_index = list(sheets).index(next(item for item in sheets if item.get("name") == "Тест_Кроки")) + 1
    for name in names:
        source_part = generated_sheets[name]
        destination = original_sheets.get(name)
        if destination is None:
            destination = f"xl/worksheets/sheet{sheet_id}.xml"
            relationship_id = f"rIdExpected{sheet_id}"
            sheet = ET.Element(f"{{{MAIN}}}sheet", {"name": name, "sheetId": str(sheet_id), f"{{{DOCREL}}}id": relationship_id})
            sheets.insert(insertion_index, sheet)
            insertion_index += 1
            ET.SubElement(relationships, f"{{{REL}}}Relationship", {
                "Id": relationship_id, "Type": f"{DOCREL}/worksheet", "Target": f"worksheets/sheet{sheet_id}.xml"})
            ET.SubElement(types, f"{{{CT}}}Override", {
                "PartName": f"/{destination}", "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"})
            sheet_id += 1
        node = xml(generated, source_part)
        for cell in node.findall(".//m:c", NS):
            if cell.get("s") is not None:
                cell.set("s", str(int(cell.get("s")) + offsets["cellXfs"]))
            if cell.get("t") == "s":
                value = cell.find("m:v", NS)
                string = copy.deepcopy(strings[int(value.text)])
                string.tag = f"{{{MAIN}}}is"
                cell.remove(value)
                cell.append(string)
                cell.set("t", "inlineStr")
        for rule in node.findall(".//m:cfRule", NS):
            if rule.get("dxfId") is not None:
                rule.set("dxfId", str(int(rule.get("dxfId")) + offsets["dxfs"]))
        if name in original_sheets and name != "Тест_Кроки":
            node = preserve_existing_cells(original, destination, node, name)
        new_relationships = xml(generated, rel_path(source_part))
        original_tables = {}
        if rel_path(destination) in original:
            for relationship in xml(original, rel_path(destination)):
                if relationship.get("Type").endswith("/table"):
                    table_part = target(destination, relationship.get("Target"))
                    original_tables[xml(original, table_part).get("name")] = table_part
        for relationship in new_relationships:
            if not relationship.get("Type").endswith("/table"):
                raise ValueError(f"Unsupported relationship on edited sheet: {name}")
            table = xml(generated, target(source_part, relationship.get("Target")))
            table_part = original_tables.get(table.get("name"))
            if table_part:
                old_table = xml(original, table_part)
                table.set("id", old_table.get("id"))
                table.find("m:tableStyleInfo", NS).set("name", old_table.find("m:tableStyleInfo", NS).get("name"))
            else:
                table.find("m:tableStyleInfo", NS).set("name", "TableStyleDark_my")
                table_part = f"xl/tables/table{table_id}.xml"
                table.set("id", str(table_id))
                table_id += 1
                ET.SubElement(types, f"{{{CT}}}Override", {
                    "PartName": f"/{table_part}", "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"})
            relationship.set("Target", posixpath.relpath(table_part, posixpath.dirname(destination)))
            original[table_part] = serialize(table)
            modified.add(table_part)
        original[destination] = serialize(node)
        original[rel_path(destination)] = serialize(new_relationships)
        modified.update((destination, rel_path(destination)))
    original["xl/workbook.xml"] = serialize(book)
    original["xl/_rels/workbook.xml.rels"] = serialize(relationships)
    original["[Content_Types].xml"] = serialize(types)
    for name, content in baseline.items():
        if name not in modified and original[name] != content:
            raise AssertionError(f"Untouched workbook part changed: {name}")
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in original.items():
            archive.writestr(name, content)
    print(f"Preserved {len(baseline.keys() - modified)} untouched OOXML parts; updated {len(names)} sheets.")


if __name__ == "__main__":
    if sys.argv[1] == "dump":
        dump(sys.argv[2])
    elif sys.argv[1] == "assemble":
        assemble(*sys.argv[2:5], sys.argv[5:])
    else:
        raise SystemExit("Expected dump or assemble")