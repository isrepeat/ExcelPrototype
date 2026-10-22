"""Guard mode-switch retention without executing Office."""

import re

from event_routes import ROOT
from validate_movement_planner import check_declarations, procedure
from validate_movement_batch import check_blocks


def main():
    folder = ROOT / "PROJECTS_EXCEL/vba/PersonalEventBuilder"
    manager = (folder / "obj_PEB_FormExtensions.cls.vba").read_text(encoding="utf-8-sig")
    controller = (folder / "obj_PEB_PgMainController.cls.vba").read_text(encoding="utf-8-sig")
    for source in (manager, controller):
        check_declarations(source)
        check_blocks(source)
        names = re.findall(r"(?im)^(?:Public|Private) (?:Function|Sub) (\w+)", source)
        assert len(names) == len(set(names))
    build = procedure(manager, "Build")
    assert 'private_PublishMarkup("ExtensionHeader", header)' in build
    assert 'private_PublishMarkup("ExtensionBody", body)' in build
    assert 'private_PublishMenu key, kinds, options, payload' in build
    assert 'SetObject("Data", key, options)' not in build
    markup = procedure(manager, "private_PublishMarkup")
    assert markup.index("previous.XML = node.XML") < markup.index("SetObject")
    menu = procedure(manager, "private_PublishMenu")
    assert menu.index("m_menuSignatures(name) = signature Then Exit Sub") < menu.index("SetObject")
    assert menu.index("Then Exit Sub") < menu.index("private_MenuCommand")
    assert "m_menuSignatures.RemoveAll" in procedure(manager, "InvalidateCache")
    details = procedure(manager, "private_PublishDetails")
    assert details.index("m_detailsSignature = signature Then Exit Function") < details.index("SetObject")
    assert 'cell("sourceAlias")' in details and 'cell("lane")' in details
    assert 'cell("column")' in details and 'VarType(value)' in details
    load = procedure(controller, "private_LoadMode")
    fast = load.split("' These modes share", 1)[1].split("private_LoadMode = True", 1)[0]
    assert 'projectedResult.RowCount > 0' in fast
    assert 'private_TryEnrichResult(projectedResult, diagnostic, False)' in fast
    enrich = procedure(controller, "private_TryEnrichResult")
    assert 'Optional ByVal refreshLegend As Boolean = True' in enrich
    assert 'If refreshLegend Then' in enrich
    common = ROOT / "PROJECTS_EXCEL/vba/common"
    form = (common / "[3] controls/obj_UiFormControl.cls.vba").read_text(encoding="utf-8-sig")
    assert 'Set child = m_fields(key)' in procedure(form, "TryRebuildFields")
    render = (common / "[2] classes/obj_UiRenderContext.cls.vba").read_text(encoding="utf-8-sig")
    assert 'If retainedForms.Exists(key) Then GoTo NextOldArea' in render
    print("Mode retention: publication guards, invalidation and partial form reuse checks passed.")


if __name__ == "__main__":
    main()