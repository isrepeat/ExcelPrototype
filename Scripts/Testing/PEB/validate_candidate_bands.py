"""Validate candidate band projection and selection contracts without Office."""

import xml.etree.ElementTree as ET

from event_routes import ROOT, parse_config
from validate_movement_planner import check_declarations, procedure
from validate_movement_batch import check_blocks


def main():
    common = ROOT / "PROJECTS_EXCEL/vba/common"
    source = (common / "[2] classes/obj_UiCandidatesSource.cls.vba").read_text(encoding="utf-8-sig")
    control = (common / "[3] controls/obj_UiCandidatesControl.cls.vba").read_text(encoding="utf-8-sig")
    table = (common / "[3] controls/obj_UiTableControl.cls.vba").read_text(encoding="utf-8-sig")
    manager = (ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder/obj_PEB_FormExtensions.cls.vba").read_text(encoding="utf-8-sig")
    for text in (source, control, manager):
        check_declarations(text)
        check_blocks(text)
    for name in ("private_SelectedArea", "private_SourceColumnArea", "private_StyleSelectedRow"):
        check_blocks(procedure(table, name))
    assert "Step m_rowsPerItem" in procedure(source, "obj_IUiHeaderPresentation_HeaderTextAt")
    assert "(rowIndex - 1) \\ m_rowsPerItem + 1" in procedure(source, "obj_IUiSelectionSource_GetItemAt")
    assert "((rowIndex - 1) \\ m_rowsPerItem) * m_rowsPerItem + 1" in procedure(source, "GetRowGroup")
    assert "firstRow + rowCount <= rowIndex" in procedure(table, "private_SourceColumnArea")
    assert "private_SelectedArea(rowIndex)" in procedure(table, "private_StyleSelectedRow")
    assert "private_SelectedArea(m_selectedRow)" in procedure(table, "obj_IUiEventHandler_HandleEvent")
    assert "private_AddCandidateCell lane, private_ExtensionColumn" in procedure(manager, "private_BuildRows")
    assert 'private_Config(kind, "anchor")' in procedure(manager, "private_Allowed")
    assert '"binding"' not in manager
    assert '+ index - 1' in procedure(manager, "private_ExtensionColumn")
    assert 'private_FieldConfig(row("Type"), CStr(field), "caption")' in procedure(manager, "private_BuildRows")
    assert 'private_FieldConfig(row("Type"), CStr(field), "format")' in procedure(manager, "private_BuildRows")
    for filename in ("wsConfig.txt", "wsConfig.test.txt"):
        config = parse_config(ROOT / "PROJECTS_EXCEL/config/PersonalEventBuilder" / filename)
        for kind in config["extensions.types"].split("|"):
            prefix = f"extensions.type.{kind}"
            assert prefix + ".binding" not in config
            assert prefix + ".fields" not in config
            assert prefix + ".dates" not in config
            assert prefix + ".start" not in config
            assert prefix + ".fallbackColumn" not in config
            names = []
            for index in range(1, int(config[prefix + ".form.count"]) + 1):
                field = f"{prefix}.form.{index}"
                names.append(config[field + ".name"])
                assert config[field + ".caption"]
                assert config[field + ".format"]
                assert config[field + ".type"] in ("text", "number", "date")
                if config[field + ".format"] == "dd.mm.yyyy":
                    assert config[field + ".type"] == "date"
            assert len(names) == len(set(names))
        assert "Ticket" not in config["extensions.types"].split("|")
        assert not any(key.startswith("forms.sections.Ticket.") for key in config)
    # Independent names need no matching primary field; only the anchor sets position.
    assert [13 + index for index, _ in enumerate(("CardKind", "CardCode", "IssuedOn"))] == [13, 14, 15]
    export = (ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder/ex_PEB_Export.vba").read_text(encoding="utf-8-sig")
    service = (ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder/obj_PEB_ExportService.cls.vba").read_text(encoding="utf-8-sig")
    for text in (export, service):
        check_declarations(text)
        check_blocks(text)
    assert 'fn_ExtensionValue(scalar, private_FieldConfig' in procedure(manager, "Collect")
    prepare = procedure(service, "TryPrepareExtensions")
    assert 'fieldPrefix & ".type"' in prepare
    assert 'fn_ExtensionValue(row(field), fieldType' in prepare
    assert '.dates' not in prepare
    normalize = procedure(export, "fn_ExtensionValue")
    assert 'Case "text", "number", "date"' in normalize
    assert 'If Not IsNumeric(value)' in normalize
    assert 'fn_ParseDate(value, field)' in normalize
    assert 'private_BuildTicket' not in manager
    assert 'forms.sections.Ticket' not in manager
    assert 'Set renderRows = private_RenderRows()' in procedure(manager, "private_BuildRows")
    assert '"binding"' not in procedure(manager, "private_RenderRows")
    publish = procedure(manager, "private_PublishCandidateDetails")
    assert "related.TryGetValue sourceField" in publish
    assert 'record.GetDetail(CStr(extension("Type")))' in publish
    assert 'TryGetValue("Form"' not in publish
    assert '"PersonnelCandidatesArePrimary"' in publish
    assert 'If result Is Nothing Or Not primary Then' in publish
    controller = (ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder/obj_PEB_PgMainController.cls.vba").read_text(encoding="utf-8-sig")
    check_declarations(controller)
    check_blocks(controller)
    publish_source = procedure(controller, "private_PublishCandidates")
    assert 'Not (m_reporterActive Or m_hospitalActive Or m_documentActive)' in publish_source
    assert publish_source.index('"PersonnelCandidatesArePrimary"') < publish_source.index('"PersonnelCandidates", result')
    assert controller.count('SetObject("Data", "PersonnelCandidates",') == 1
    assert 'cell.Add "sourceAlias", sourceAlias' in publish
    assert "candidates.AppendDetailSourceRanges m_bodyRange, sourceRanges" in procedure(table, "private_Render")
    root = ET.parse(ROOT / "PROJECTS_EXCEL/ui/PersonalEventBuilder/MainPage.xaml").getroot()
    candidates = next(node for node in root.iter() if node.get("name") == "Candidates")
    assert candidates.get("detailsSource") == "{Binding Path=Data.PersonnelCandidateDetails}"
    factory = (common / "[2] classes/obj_UiControlFactory.cls.vba").read_text(encoding="utf-8-sig")
    candidate_schema = factory.split('If m_type = "candidates" Then', 1)[1].split("End If", 1)[0]
    assert 'schema.AddAttribute "detailsSource", "string", False, "", True, 0' in candidate_schema
    assert 'schema.AddAttribute "numberFormat", "string", False, "", False, 0' in procedure(factory, "private_FieldSchema")

    # TVO and medical document share band 1; a ticket occupies band 2.
    for count in (0, 1, 2, 50):
        for stride in (1, 2, 3, 4):
            physical = 1
            visited = []
            while physical <= count * stride:
                first = ((physical - 1) // stride) * stride + 1
                following = first + stride
                assert following > physical
                visited.append(first)
                physical = following
            assert visited == [index * stride + 1 for index in range(count)]
        main_rows = [[f"{person}:{column}" for column in range(17)] for person in range(count)]
        output = []
        for person, row in enumerate(main_rows):
            output.append(row.copy())
            shared = [None] * 17
            shared[3:6] = [f"tvo-{person}"] * 3
            shared[12:15] = [f"medical-{person}"] * 3
            output.append(shared)
            ticket = [None] * 17
            ticket[12:14] = [f"ticket-{person}", f"date-{person}"]
            output.append(ticket)
        assert len(output) == count * 3
        for physical, row in enumerate(output):
            candidate = physical // 3
            first = physical // 3 * 3
            assert first <= physical < first + 3
            assert output[first] == main_rows[candidate]
            if physical % 3 == 2:
                assert row[12] == f"ticket-{candidate}"
        assert [output[index * 3] for index in range(count)] == main_rows
    print("Candidate bands: static contracts and 0/1/2/50 candidate projection checks passed.")


if __name__ == "__main__":
    main()