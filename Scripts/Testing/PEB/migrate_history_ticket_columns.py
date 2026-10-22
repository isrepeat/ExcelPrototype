"""Unify legacy ticket columns; preview by default, --apply requires a closed workbook."""

import argparse
from copy import deepcopy
from datetime import datetime
from io import BytesIO
import json
import os
from pathlib import Path
import posixpath
import re
import shutil
import tempfile
from zipfile import ZipFile

from lxml import etree as ET

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def column_number(name):
    result = 0
    for character in name:
        result = result * 26 + ord(character) - 64
    return result


def column_name(number):
    result = ""
    while number:
        number, digit = divmod(number - 1, 26)
        result = chr(65 + digit) + result
    return result


def coordinate(address):
    match = re.fullmatch(r"([A-Z]+)(\d+)", address)
    if not match:
        raise ValueError(f"Unsupported cell address: {address}")
    return column_number(match[1]), int(match[2])


def value(cell, strings):
    if cell is None:
        return ""
    node = cell.find(f"{{{NS}}}v")
    if cell.get("t") == "s" and node is not None:
        return strings[int(node.text)]
    if cell.get("t") == "inlineStr":
        return "".join(cell.find(f"{{{NS}}}is").itertext())
    return node.text or "" if node is not None else ""


def text_cell(text):
    cell = ET.Element(f"{{{NS}}}c", t="inlineStr")
    inline = ET.SubElement(cell, f"{{{NS}}}is")
    ET.SubElement(inline, f"{{{NS}}}t").text = text
    return cell


def migrate_rows(rows, prefix, strings):
    """Preserve a previous primary document as a child in its original revision."""
    result = []
    converted = 0
    children = 0
    maximum = {}
    for row in rows:
        root = value(row.get("EventRowId" if prefix else "RowId"), strings)
        change = value(row.get("ChangeId"), strings) if prefix else ""
        group = (root, change) if prefix else ("", "")
        order = value(row.get(prefix + "RowOrder"), strings)
        maximum[group] = max(maximum.get(group, 0), int(order or 0))
    existing_ids = {value(row.get(prefix + "RowId"), strings) for row in rows}
    for original in rows:
        row = deepcopy(original)
        number = row.get(prefix + "TicketNumber")
        date = row.get(prefix + "TicketDate")
        if value(number, strings) or value(date, strings):
            if value(row.get(prefix + "RowType"), strings) != "Event":
                raise ValueError("Legacy ticket on a child row requires manual review.")
            kind = value(row.get(prefix + "Document"), strings).strip().lower()
            root = value(row.get(prefix + "RowId"), strings)
            group = (value(row.get("EventRowId"), strings), value(row.get("ChangeId"), strings)) if prefix else ("", "")
            if kind == "в/к":
                for name, old in (("DocumentNumber", number), ("DocumentDate", date)):
                    current = value(row.get(prefix + name), strings)
                    legacy = value(old, strings)
                    if current and legacy and current != legacy:
                        raise ValueError(f"Conflicting ticket values in {root}: {name}")
            elif any(value(row.get(prefix + name), strings) for name in ("Document", "DocumentNumber", "DocumentDate")):
                identifier = root + "/legacy-document"
                if identifier in existing_ids:
                    raise ValueError(f"Migration child ID already exists: {identifier}")
                child = deepcopy(row)
                for name in list(child):
                    if name.startswith(prefix) and (prefix or name not in ("EventRowId", "ChangeId")):
                        field = name[len(prefix):]
                        if field not in ("Document", "DocumentNumber", "DocumentDate", "AcceptedOrderNumber", "AcceptedOrderDate", "OrderNumber", "OrderDate"):
                            child[name] = text_cell("")
                maximum[group] = maximum.get(group, 0) + 1
                for name, text in {"RowId": identifier, "ParentRowId": root, "RowType": "Document",
                                   "RowOrder": str(maximum[group]), "ExtensionRole": ""}.items():
                    child[prefix + name] = text_cell(text)
                # Revision metadata (ChangeId, EventRowId, Revision, Status, ...) remains intact.
                if not prefix:
                    child["Revision"] = text_cell("0")
                    child["LastChangeId"] = text_cell("")
                    if "Status" in row:
                        child["Status"] = deepcopy(row["Status"])
                result.append(child)
                children += 1
            row[prefix + "Document"] = text_cell("В/к")
            for name, old in (("DocumentNumber", number), ("DocumentDate", date)):
                if value(old, strings):
                    row[prefix + name] = deepcopy(old)
            converted += 1
        result.append(row)
    return result, converted, children


def migrate_package(payload):
    report = {}
    with ZipFile(BytesIO(payload)) as archive:
        parts = {info.filename: archive.read(info.filename) for info in archive.infolist()}
        strings = []
        if "xl/sharedStrings.xml" in parts:
            strings = ["".join(node.itertext()) for node in ET.fromstring(parts["xl/sharedStrings.xml"]).findall(f"{{{NS}}}si")]
        for sheet_path in list(parts):
            if not re.fullmatch(r"xl/worksheets/sheet\d+\.xml", sheet_path):
                continue
            rel_path = posixpath.join(posixpath.dirname(sheet_path), "_rels", posixpath.basename(sheet_path) + ".rels")
            if rel_path not in parts:
                continue
            relations = ET.fromstring(parts[rel_path])
            for relation in relations:
                if not relation.get("Type", "").endswith("/table"):
                    continue
                target = relation.get("Target")
                table_path = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(sheet_path), target))
                table = ET.fromstring(parts[table_path])
                name = table.get("name")
                if name not in ("tbEventsHistory", "tbEventRevisions"):
                    continue
                prefix = "Snapshot." if name == "tbEventRevisions" else ""
                columns = table.find(f"{{{NS}}}tableColumns")
                headers = [node.get("name") for node in columns]
                old_names = {prefix + "TicketNumber", prefix + "TicketDate"}
                if not old_names & set(headers):
                    report[name] = {"already_migrated": True}
                    continue
                if not old_names <= set(headers):
                    raise ValueError(f"Partially migrated schema in {name}")
                sheet = ET.fromstring(parts[sheet_path])
                if sheet.findall(f".//{{{NS}}}f"):
                    raise ValueError(f"Formula-bearing history sheet needs manual migration: {name}")
                if sheet.find(f"{{{NS}}}mergeCells") is not None:
                    raise ValueError(f"Merged history sheet needs manual migration: {name}")
                if len(sheet.findall(f"{{{NS}}}tableParts/{{{NS}}}tablePart")) != 1:
                    raise ValueError("Only one table per migrated sheet is supported.")
                first, last = table.get("ref").split(":")
                start_column, start_row = coordinate(first)
                end_column, end_row = coordinate(last)
                data = sheet.find(f"{{{NS}}}sheetData")
                source_rows = {int(node.get("r")): node for node in data}
                rows = []
                for index in range(start_row + 1, end_row + 1):
                    node = source_rows.get(index)
                    cells = {coordinate(cell.get("r"))[0]: cell for cell in node} if node is not None else {}
                    rows.append({header: deepcopy(cells.get(start_column + offset))
                                 for offset, header in enumerate(headers) if cells.get(start_column + offset) is not None})
                migrated, converted, children = migrate_rows(rows, prefix, strings)
                removed = sorted(start_column + headers.index(key) for key in old_names)
                retained = [header for header in headers if header not in old_names]
                new_end_row = start_row + len(migrated)
                for index in range(end_row + 1, new_end_row + 1):
                    node = source_rows.get(index)
                    if node is not None and any(value(cell, strings) for cell in node):
                        raise ValueError(f"Cannot expand {name}: cells below the table contain data.")
                original_header = {coordinate(cell.get("r"))[0]: cell for cell in source_rows[start_row]}
                header_row = {header: deepcopy(original_header.get(start_column + headers.index(header), text_cell(header))) for header in retained}
                for node in data:
                    for cell in list(node):
                        column, row_index = coordinate(cell.get("r"))
                        if column in removed:
                            node.remove(cell)
                        else:
                            cell.set("r", column_name(column - sum(item < column for item in removed)) + str(row_index))
                for index, row in enumerate([header_row] + migrated, start_row):
                    node = source_rows.get(index)
                    if node is None:
                        node = ET.SubElement(data, f"{{{NS}}}row", r=str(index))
                    for cell in list(node):
                        if start_column <= coordinate(cell.get("r"))[0] <= end_column - 2:
                            node.remove(cell)
                    for offset, header in enumerate(retained):
                        cell = deepcopy(row.get(header, text_cell("")))
                        cell.set("r", column_name(start_column + offset) + str(index))
                        node.append(cell)
                    node[:] = sorted(node, key=lambda cell: coordinate(cell.get("r"))[0])
                data[:] = sorted(data, key=lambda node: int(node.get("r")))
                for node in list(columns):
                    if node.get("name") in old_names:
                        columns.remove(node)
                for index, node in enumerate(columns, 1):
                    node.set("id", str(index))
                columns.set("count", str(len(retained)))
                ref = first + ":" + column_name(end_column - 2) + str(new_end_row)
                table.set("ref", ref)
                auto_filter = table.find(f"{{{NS}}}autoFilter")
                if auto_filter is not None:
                    auto_filter.set("ref", ref)
                    for child in list(auto_filter):
                        auto_filter.remove(child)
                for node in list(table.findall(f"{{{NS}}}sortState")):
                    table.remove(node)
                width_nodes = sheet.find(f"{{{NS}}}cols")
                if width_nodes is not None:
                    for node in list(width_nodes):
                        low, high = int(node.get("min")), int(node.get("max"))
                        kept = [item for item in range(low, high + 1) if item not in removed]
                        if not kept:
                            width_nodes.remove(node)
                        else:
                            node.set("min", str(kept[0] - sum(item < kept[0] for item in removed)))
                            node.set("max", str(kept[-1] - sum(item < kept[-1] for item in removed)))
                dimension = sheet.find(f"{{{NS}}}dimension")
                if dimension is not None:
                    dimension.set("ref", f"A1:{column_name(end_column - 2)}{max(new_end_row, max(source_rows))}")
                parts[sheet_path] = ET.tostring(sheet, xml_declaration=True, encoding="UTF-8", standalone=True)
                parts[table_path] = ET.tostring(table, xml_declaration=True, encoding="UTF-8", standalone=True)
                report[name] = {"converted": converted, "preserved_document_children": children, "removed_columns": sorted(old_names)}
        if not {"tbEventsHistory", "tbEventRevisions"} <= report.keys():
            raise ValueError("Both history and revision tables are required.")
        output = BytesIO()
        with ZipFile(output, "w") as target:
            for info in archive.infolist():
                target.writestr(info, parts[info.filename])
    return output.getvalue(), report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args()
    payload, report = migrate_package(args.workbook.read_bytes())
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if not args.apply:
        print("Preview only; workbook was not modified.")
        return
    if all(item.get("already_migrated") for item in report.values()):
        print("Already migrated; workbook was not modified.")
        return
    backup = args.workbook.with_name(args.workbook.stem + ".before-ticket-unification-" + datetime.now().strftime("%Y%m%d-%H%M%S-%f") + args.workbook.suffix)
    shutil.copy2(args.workbook, backup)
    descriptor, temporary = tempfile.mkstemp(suffix=".xlsx", dir=args.workbook.parent)
    try:
        with os.fdopen(descriptor, "wb") as stream:
            stream.write(payload)
        os.replace(temporary, args.workbook)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)
    print(f"Updated {args.workbook}; backup: {backup}")


if __name__ == "__main__":
    main()