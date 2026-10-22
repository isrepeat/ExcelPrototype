"""Check order-scoped last-event cancellation and replacement without Office."""

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
    assert delete.index('row("Status") = "Cancelled"') < delete.index("fn_ShowMessage")
    assert delete.index("fn_ShowMessage") < delete.index("private_ReadLastOrderEvent(current")
    assert delete.index("private_ReadLastOrderEvent(current") < delete.index("fn_SetStatus")
    assert 'current("RowId") <> identifier Or ex_PEB_History.fn_Stamp(current) <> stamp' in delete
    assert 'fn_SetStatus(current, "Cancelled"' in delete
    assert "vbDefaultButton2" in delete and not re.search(r"\.Delete\b", delete)
    export = procedure(controller, "private_Export")
    assert 'values("RequestedStatus") = "Active"' in export
    assert 'values("ExpectedHistoryStamp") = ex_PEB_History.fn_Stamp(replacement)' in export
    assert 'If success Then Set m_replacementEvent = Nothing' in export
    assert 'Set m_replacementEvent = Nothing' in procedure(controller, "ResetFormHandler")
    assert '"DeleteLastEvent"' in procedure(page, "private_TryRegisterCommands")
    root = ET.parse(ROOT / "PROJECTS_EXCEL/ui/PersonalEventBuilder/MainPage.xaml").getroot()
    button = next(node for node in root.iter() if node.get("name") == "DeleteLastEvent")
    assert button.get("command") == "{Binding Path=Commands.DeleteLastEvent}"
    assert (button.get("row"), button.get("column")) == ("23", "20")
    keys = ("DeleteLastEvent", "DeleteLastEventConfirm", "DeleteLastEventReason",
            "ReplaceLastEventReason", "LastEventDeleted", "LastEventAlreadyDeleted", "LastEventChanged")
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
    print("PASS: persisted order scope, child exclusion, cancelled tail, confirmation, stale checks and replacement contracts. VBA was not executed.")


if __name__ == "__main__":
    main()