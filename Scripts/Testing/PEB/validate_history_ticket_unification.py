"""Check canonical ticket storage and lossless migration in memory, without Office."""

from io import BytesIO
from zipfile import ZipFile

from lxml import etree as ET

from event_routes import ROOT, parse_config
from migrate_history_ticket_columns import NS, migrate_package, migrate_rows, text_cell, value


def main():
    service = (ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder/obj_PEB_ExportService.cls.vba").read_text(encoding="utf-8-sig")
    word = service.split("Public Function TryPrepareHistory(", 1)[1].split("End Function", 1)[0]
    assert word.index('fn_Apply(values, "Word")') < word.index("fn_TicketDocumentValues values, True")
    assert word.index("fn_TicketDocumentValues values, True") < word.index("private_AddAliases values")
    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        settings = parse_config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder" / filename)
        fields = dict(item.split("=") for item in settings["history.fields"].split("|"))
        assert not {"TicketNumber", "TicketDate"} & fields.keys()
        assert {"Document", "DocumentNumber", "DocumentDate"} <= fields.keys()
        assert not {"TicketNumber", "TicketDate"} & set(settings["export.fields"].split("|"))
        assert settings["history.list.columns.7.field"] == "DocumentNumber"
    for prefix in ("", "Snapshot."):
        original = {prefix + key: text_cell(text) for key, text in {
            "RowId": "event-1", "ParentRowId": "", "RowType": "Event", "RowOrder": "1",
            "Document": "м/к", "DocumentNumber": "medical-12", "DocumentDate": "01.10.2026",
            "TicketNumber": "2026/283/8", "TicketDate": "30.09.2026"}.items()}
        if prefix:
            original.update({"EventRowId": text_cell("event-1"), "ChangeId": text_cell("event-1/R1")})
        rows, converted, children = migrate_rows([original], prefix, [])
        assert (converted, children, len(rows)) == (1, 1, 2)
        child, root = rows
        assert value(root[prefix + "Document"], []) == "В/к"
        assert value(root[prefix + "DocumentNumber"], []) == "2026/283/8"
        assert value(child[prefix + "DocumentNumber"], []) == "medical-12"
        assert value(child[prefix + "ParentRowId"], []) == "event-1"
        if prefix:
            assert value(child["EventRowId"], []) == "event-1"
            assert value(child["ChangeId"], []) == "event-1/R1"
        else:
            assert value(child["Revision"], []) == "0"
            assert value(child["LastChangeId"], []) == ""
    path = ROOT / "PROJECTS_EXCEL/workbook/PersonalEventBuilder/EventsHistory.xlsx"
    before = path.read_bytes()
    after, report = migrate_package(before)
    again, repeated = migrate_package(after)
    assert all(row.get("already_migrated") for row in repeated.values())
    with ZipFile(BytesIO(before)) as old, ZipFile(BytesIO(after)) as new:
        assert set(old.namelist()) == set(new.namelist())
        for name in old.namelist():
            if not name.startswith(("xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml", "xl/tables/table1.xml", "xl/tables/table2.xml")):
                assert old.read(name) == new.read(name), name
        for name in new.namelist():
            if name.startswith("xl/tables/") and name.endswith(".xml"):
                table = ET.fromstring(new.read(name))
                names = [node.get("name") for node in table.findall(f"{{{NS}}}tableColumns/{{{NS}}}tableColumn")]
                assert not {"TicketNumber", "TicketDate", "Snapshot.TicketNumber", "Snapshot.TicketDate"} & set(names)
    assert path.read_bytes() == before
    print("PASS: canonical schema, preserved documents/revisions, idempotent in-memory migration; working workbook unchanged.")


if __name__ == "__main__":
    main()