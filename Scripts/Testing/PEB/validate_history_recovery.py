"""Check recovery payload parsing and write-ahead source ordering without Office."""

from pathlib import Path
from unittest.mock import patch

from event_routes import ROOT
from inspect_history_recovery import inspect
from validate_movement_batch import check_module_contract
from validate_movement_planner import procedure


def main():
    folder = ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder"
    recovery = (folder / "ex_PEB_HistoryRecovery.vba").read_text(encoding="utf-8-sig")
    history = (folder / "ex_PEB_History.vba").read_text(encoding="utf-8-sig")
    check_module_contract(recovery, "ex_PEB_HistoryRecovery")
    save = procedure(history, "fn_Save")
    assert save.index("ex_PEB_HistoryRecovery.fn_Prepare") < save.index("target.Range.Value2 = data")
    assert save.index("target.Range.Value2 = data") < save.index('recoveryToken, "COMMIT"')
    assert save.index('recoveryToken, "COMMIT"') < save.index("committed = True")
    assert "transaction(\"OriginalRows\") + 1" in recovery
    assert "files.BuildPath(book.Path, book.Name" in recovery
    text = (
        '<Recovery version="1" operation="a" state="PREPARE" workbook="EventsHistory.xlsx" time="t" user="u">'
        '<Row role="After" sheet="EventsHistory" table="tbEventsHistory">'
        '<Field name="RowId" type="8">E-1</Field>'
        '<Field name="Basis" type="8">Підстава &amp; текст&#13;&#10;рядок</Field>'
        '</Row></Recovery>\n'
        '<Recovery version="1" operation="a" state="COMMIT"/>\n'
        '<Recovery version="1" operation="b" state="PREPARE" workbook="x" time="t" user="u"/>\n'
        '<Recovery version="1" operation="broken"'
    )
    with patch.object(Path, "read_text", return_value=text):
        result = inspect(Path("unused"))
    assert len(result["committed"]) == 1
    assert result["committed"][0]["rows"][0]["fields"]["Basis"]["value"] == "Підстава & текст\r\nрядок"
    assert result["pending"][0]["operation"] == "b"
    assert result["invalid_lines"] == [4]
    print("PASS: recovery ordering, complete row payloads, Unicode/newlines, pending operations and truncated tail. Office/VBA not executed.")


if __name__ == "__main__":
    main()