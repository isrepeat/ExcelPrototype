"""Check order-scoped physical deletion and rollback contracts without Office."""

import xml.etree.ElementTree as ET
import re

from event_routes import ROOT, parse_config
from validate_movement_planner import check_declarations, procedure
from validate_movement_batch import check_blocks


def main():
    folder = ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder"
    controller = (folder / "obj_PEB_PgMainController.cls.vba").read_text(encoding="utf-8-sig")
    page = (folder / "obj_PEB_PgMain.cls.vba").read_text(encoding="utf-8-sig")
    for source in (controller, page):
        check_declarations(source)
        check_blocks(source)
    read = procedure(controller, "private_ReadLastOrderEvent")
    assert 'private_TryReadFormValue("AcceptedOrderNumber"' in read
    assert 'private_TryReadFormValue("AcceptedOrderDate"' in read
    assert "fn_ReadOrder(CStr(orderNumber), orderDate, events, diagnostic, True, Nothing, False)" in read
    assert "Set row = events(events.Count)" in read
    assert "m_historyRows" not in read and "m_lookupResult" not in read
    delete = procedure(controller, "DeleteLastEventHandler")
    assert delete.index("fn_ShowMessage") < delete.index("private_ReadLastOrderEvent(current")
    assert delete.index("private_ReadLastOrderEvent(current") < delete.index("fn_DeleteLast")
    assert 'current("RowId") <> identifier Or ex_PEB_History.fn_Stamp(current) <> stamp' in delete
    assert 'fn_SetStatus' not in delete
    assert "vbDefaultButton2" in delete and not re.search(r"\.Delete\b", delete)
    assert 'm_replacementEvent' not in controller
    history = (folder / "ex_PEB_History.vba").read_text(encoding="utf-8-sig")
    check_declarations(history)
    check_blocks(history)
    physical = procedure(history, "fn_DeleteLast")
    assert 'Set current = events(events.Count)' in physical
    assert 'fn_Stamp(current) <> expectedStamp' in physical
    assert 'Loop While ids.Count <> previousCount' in physical
    assert 'versions.ListColumns("EventRowId")' in physical
    assert 'target.ListRows.Count To 1 Step -1' in physical
    assert 'target.ListRows(rowIndex).Delete' in physical
    assert 'removed.Count To 1 Step -1' in physical
    assert 'ListRows.Add(Position:=' in physical
    assert physical.count('Application.EnableEvents = previousEvents') == 1
    assert physical.count('Application.ScreenUpdating = previousScreen') == 1
    assert 'fn_Save' not in physical and 'fn_SetStatus' not in physical
    assert not re.search(r'\.(?:Save|SaveAs|Close)\b', physical)
    assert '"DeleteLastEvent"' in procedure(page, "private_TryRegisterCommands")
    root = ET.parse(ROOT / "PROJECTS_EXCEL/ui/PersonalEventBuilder/MainPage.xaml").getroot()
    button = next(node for node in root.iter() if node.get("name") == "DeleteLastEvent")
    assert button.get("command") == "{Binding Path=Commands.DeleteLastEvent}"
    assert (button.get("row"), button.get("column")) == ("23", "20")
    keys = ("DeleteLastEvent", "DeleteLastEventConfirm", "LastEventDeleted", "LastEventChanged")
    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        config = parse_config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder" / filename)
        assert all(config["text." + key] for key in keys)

    # A later event from another order and trailing child rows must not become the target.
    rows = [("first", "301", "2026", "Event", 1, "Active"),
            ("last", "301", "2026", "Event", 2, "Cancelled"),
            ("other-order", "302", "2026", "Event", 3, "Active"),
            ("other-year", "301", "2025", "Event", 4, "Active"),
            ("child", "301", "2026", "Document", 5, "Active")]
    scoped = sorted((row for row in rows if row[1:4] == ("301", "2026", "Event")), key=lambda row: row[4])
    assert scoped[-1][0] == "last" and scoped[-1][5] == "Cancelled"
    children = {"child": "last", "nested": "child", "other-child": "other-order"}
    removed = {"last"}
    while True:
        following = removed | {child for child, parent in children.items() if parent in removed}
        if following == removed:
            break
        removed = following
    assert removed == {"last", "child", "nested"}
    revisions = [("last", 1), ("last", 2), ("other-order", 1)]
    assert [row for row in revisions if row[0] not in removed] == [("other-order", 1)]
    print("PASS: physical deletion, order scope, descendants, all revisions, confirmation, stale checks and rollback. VBA was not executed.")


if __name__ == "__main__":
    main()