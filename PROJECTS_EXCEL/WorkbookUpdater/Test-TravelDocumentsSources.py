"""Validate form, configuration, registry and DOCX contracts without running Excel or VBA."""

from pathlib import Path, PureWindowsPath
import re
import xml.etree.ElementTree as ET
import zipfile


ROOT = Path(__file__).resolve().parents[1]
PREFIX = "TravelDocumentsBuilder::"
config = {}
for line in (ROOT / "config/TravelDocumentsBuilder/wsConfig.txt").read_text(encoding="utf-8-sig").splitlines():
    if not line.strip():
        continue
    fields = line.split("\t")
    assert len(fields) == 3, f"Invalid config row: {line}"
    assert fields[1] not in config, f"Duplicate config key: {fields[1]}"
    config[fields[1]] = fields[2].strip()


def cfg(key):
    assert PREFIX + key in config, f"Missing config: {key}"
    return config[PREFIX + key]


sources = ROOT / "vba/TravelDocumentsBuilder"
page = (sources / "obj_TDB_PgMain.cls.vba").read_text(encoding="utf-8-sig")
controller = (sources / "obj_TDB_PgMainController.cls.vba").read_text(encoding="utf-8-sig")
generator = (sources / "obj_TDB_VacationGenerator.cls.vba").read_text(encoding="utf-8-sig")
all_sources = page + controller + generator

# Planned return is persisted as a calculated date value, not formula text in snapshots.
write_row = re.search(r"Private Sub private_WriteRow\([\s\S]*?^End Sub", generator, re.M)[0]
assert 'ByVal plannedReturn As Date' in write_row
assert 'ex_TDB_History.fn_Write table, row, fields' in write_row
assert '.Formula =' not in write_row

# Every form value written to the registry must first be captured from the binding context.
read_form = re.search(r"Private Sub private_ReadForm\(\)[\s\S]*?^End Sub", generator, re.M)[0]
read_fields = re.search(r"For Each key In VBA.Array\(([\s\S]*?)\)\n", read_form)[1]
read_fields = set(re.findall(r'"(\w+)"', read_fields))
assert set(re.findall(r'm_form\("(\w+)"\)', write_row)) <= read_fields
assert 'PersonPosition' in read_fields

# Both tables persist literal text, including dates and revision metadata.
history_source = (sources / "ex_TDB_History.vba").read_text(encoding="utf-8-sig")
write = re.search(r"Public Sub fn_Write\([\s\S]*?^End Sub", history_source, re.M)[0]
assert 'target.Range.NumberFormat = "@"' in write
assert 'cell.Value2 = private_FieldText(VBA.CStr(field), row(field))' in write
assert '.Formula' not in write and 'dd.mm.yyyy' not in write
assert 'registry.arrivalFormula' not in generator and PREFIX + 'generation.registry.arrivalFormula' not in config
assert not re.search(r'\.Formula\s*=|\.Calculate\b', generator)
assert 'snapshot(field) = private_FieldText(' in history_source
assert 'tickets.DataBodyRange.Value2' in history_source and 'revisions.DataBodyRange.Value2' in history_source
assert 'table.DataBodyRange.Value2 = data' in history_source
assert 'If formulas.Exists(key) Then cell.Formula = formulas(key)' in history_source
assert 'If VBA.Format$(parsed, "dd.mm.yyyy") <> text Then GoTo InvalidDate' in history_source
assert 'fn_ReturnDate = VBA.DateAdd("d", vacationDays + roadDays + donationDays, departure)' in history_source
assert 'ex_TDB_History.fn_ReturnDate' in controller
assert 'departureDate, vacationDays, roadDays, ticketNo, finalReturn' in generator
assert 'finalReturn, mainReturn, donationDays, 0, ticketNo, finalReturn' in generator
assert 'finalReturn = ex_TDB_History.fn_Date(root("PlannedReturnDate"))' in generator
assert 'ex_TDB_History.fn_Date(root("OrderDate"))' in generator
assert '{value}' in cfg('history.text.InvalidDate')

recovery = (sources / "ex_TDB_HistoryRecovery.vba").read_text(encoding="utf-8-sig")
assert cfg("history.recovery.suffix") == ".recovery.xml.log"
assert "{details}" in cfg("history.text.RecoveryFailed")
assert 'root.setAttribute "state", "PREPARE"' in recovery
assert 'For Each row In previous' in recovery and 'For Each row In group' in recovery
assert '"RevisionsCount"' in recovery and '"Revision"' in recovery
assert 'files.OpenTextFile(path, 8, True, -1)' in recovery
prepare_at = generator.index('recoveryToken = ex_TDB_HistoryRecovery.fn_Prepare')
mark_at = generator.index('ex_TDB_HistoryRecovery.fn_Mark')
assert generator.index('ex_TDB_History.fn_Commit') < prepare_at < mark_at < generator.index('    tableChanged = False')
assert '"COMMIT"' in generator[mark_at:mark_at + 110]

# Keep TDB section boundaries consistent with the PEB module and class conventions.
for source in sources.glob("*.vba"):
    source_text = source.read_text(encoding="utf-8-sig")
    if source.name.endswith(".cls.vba"):
        assert "namespace " not in source_text, source
        sections = list(re.finditer(r"^' //\n' // (\w+)\n' //\n", source_text, re.M))
        section_names = [section[1] for section in sections]
        assert section_names.count("Lifecycle") == 1, source
        assert section_names[0] == "Lifecycle", source
        fields = re.findall(r"^Private (\w+) As \w+", source_text, re.M)
        assert fields[:2] == ["m_isInitialized", "m_isDisposed"], source
        for index, section in enumerate(sections):
            end = sections[index + 1].start() if index + 1 < len(sections) else len(source_text)
            procedures = re.findall(r"^(Public|Private) (?:Sub|Function|Property \w+) (\w+)", source_text[section.end():end], re.M)
            if section[1] == "Lifecycle":
                assert procedures == [("Private", "Class_Initialize"), ("Private", "Class_Terminate")], source
            elif section[1] == "API":
                assert procedures[:2] == [("Public", "Initialize"), ("Public", "Dispose")], source
            elif section[1] == "Private":
                assert all(access == "Private" and name.startswith("private_") for access, name in procedures), source
    else:
        sections = re.findall(r"^' namespace (\w+) \{", source_text, re.M)
        endings = re.findall(r"^' \} // namespace (\w+)", source_text, re.M)
        assert sections == endings, source
        if source.name == "ThisWorkbook.vba":
            assert sections == ["Private"], source
        else:
            assert sections[:2] == ["Lifecycle", "API"], source
            assert "Public Sub fn_Module_Dispose()" in source_text, source

# Dialogs must use the shared Unicode API; routine feedback belongs in page status.
for source in sources.glob("*.vba"):
    source_text = source.read_text(encoding="utf-8-sig")
    assert not re.search(r"\b(?:VBA\.)?MsgBox\s*(?:\(|\s)", source_text, re.I), source
    assert not re.search(r"\b(?:MessageBox[AW]?|WScript\.Shell)\b", source_text, re.I), source
runtime_handlers = controller[controller.index("Public Function SearchPersonnelHandler()"):]
assert runtime_handlers.count("ex_WindowsUi.fn_ShowMessage") == 1
assert 'If Not private_SetActionStatus Then ex_WindowsUi.fn_ShowMessage' in runtime_handlers

# Check required keys and option lists against configuration.
key_assignment = re.search(r"    keys = (.*?)\n    For Each key", generator, re.S)[1]
required = "".join(re.findall(r'"([^"]*)"', key_assignment)).split(",")
for key in required:
    cfg("generation." + key)
for kind in cfg("generation.kind.ids").split(","):
    for part in ("input", "word", "event"):
        cfg(f"generation.kind.{kind}.{part}")
kinds = [cfg(f"generation.kind.{kind}.input") for kind in cfg("generation.kind.ids").split(",")]
assert cfg("generation.items.VacationKinds").split("|") == kinds
assert cfg("generation.defaults.VacationKind") in kinds
assert len(cfg("generation.months").split("|")) == 12
assert cfg("generation.defaults.VacationAbroad") in cfg("generation.items.AbroadOptions").split(",")
assert cfg("generation.defaults.VacationStatus") in cfg("generation.items.StatusOptions").split(",")

# Check command routes, unique names and initial form values.
markup = ET.parse(ROOT / "ui/TravelDocumentsBuilder/MainPage.xaml")
registered = dict(re.findall(r'bindingContext, "(\w+)", "(\w+)"', page))
names = []
for node in markup.iter():
    name = node.attrib.get("name")
    if name and node.tag.startswith("{urn:excelprototype:controls}"):
        assert len(name) <= 25, f"Control name too long: {name}"
        names.append(name)
        if node.tag.endswith("button"):
            assert len("btn_" + name) <= 31
    for attribute, raw in node.attrib.items():
        binding = re.search(r"Path=(\w+)\.(\w+)", raw)
        if not binding:
            continue
        category, key = binding.groups()
        if category == "Commands":
            assert key in registered, f"Unregistered command: {key}"
            assert f"Public Function {registered[key]}" in controller, f"Missing handler: {key}"
        elif category == "Form":
            derived = any(key == field + suffix for field in cfg("forms.readOnly.fields").split("|") for suffix in ("ReadOnly", "Style", "LabelStyle"))
            if derived:
                suffix = next(suffix for field in cfg("forms.readOnly.fields").split("|") for suffix in ("ReadOnly", "Style", "LabelStyle") if key == field + suffix)
                assert re.search(r'bindingContext\.SetValue\("Form", VBA\.CStr\(key\) & "' + suffix + r'",', page), f"Derived form binding not registered: {key}"
            assert f'"{key}"' in page or derived, f"Form binding not initialized: {key}"
        elif category == "Text":
            if key not in ("PersonnelStatus", "CandidateCount"):
                assert f'"{key}"' in all_sources, f"Unknown text binding: {key}"
        elif category == "Data":
            assert f'"{key}"' in page + controller, f"Unknown data binding: {key}"
assert len(names) == len(set(names)), "Duplicate control names"
assert "TvoLookup" not in names
for name in ("TvoName", "TvoIpn", "TvoPosition"):
    assert name in names, f"Missing TVO control: {name}"
assert int(cfg("lookup.Tvo.apply.count")) == 3
assert {
    cfg(f"lookup.Tvo.apply.{index}.target") for index in range(1, 4)
} == {"Form.TvoName", "Form.TvoIpn", "Form.TvoPosition"}
# Validate contracts against actual files rather than a copied test schema.
workbook_folder = ROOT / "workbook/TravelDocumentsBuilder"
with zipfile.ZipFile(workbook_folder / cfg("history.path")) as archive:
    namespace = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    tables = [
        ET.fromstring(archive.read(name))
        for name in archive.namelist()
        if re.fullmatch(r"xl/tables/table\d+\.xml", name)
    ]
    table = next(table for table in tables if table.attrib["name"] == cfg("history.table"))
    headers = [column.attrib["name"] for column in table.find("s:tableColumns", namespace)]
    assert len(headers) == len(set(headers)) == 28
    assert all(re.fullmatch(r"[A-Za-z][A-Za-z0-9]*", field) for field in headers)
    assert not {"LastChangeId", "RowOrder", "Purpose", "ReportNo", "PersonUnit"} & set(headers)
    assert set(headers) == set(cfg("history.columns").split("|")), "Registry schema mismatch"
    versions = next(table for table in tables if table.attrib["name"] == cfg("history.versions.table"))
    version_headers = [column.attrib["name"] for column in versions.find("s:tableColumns", namespace)]
    assert len(version_headers) == 35
    assert version_headers == cfg("history.versions.columns").split("|")
    snapshot_fields = cfg("history.snapshotFields").split("|")
    assert [field.removeprefix("Snapshot.") for field in version_headers[9:]] == snapshot_fields
    assert set(snapshot_fields) == set(headers) - {"Revision", "Status"}
    assert not {"ТВО.ПІБ", "ТВО.ІПН", "ТВО.Посада", "Статус"} & set(headers)
    assert set(cfg("history.dateFields").split("|")) <= set(headers)
    assert set(cfg("generation.registry.fieldOrder").split("|")) <= set(headers)
    for setting in ("formFields", "dateFormFields", "actorFields"):
        for mapping in cfg("history." + setting).split("|"):
            assert mapping.split("=")[1] in headers, mapping
    for field in re.findall(r'metadata\.Add "(\w+)"', generator):
        assert field in headers, f"Unknown saved metadata field: {field}"
    assert 'Purpose=Basis' in cfg("history.formFields")
    assert 'ReportNo=ReportNumber' in cfg("history.formFields")
    assert 'm_form("PersonPosition"), eventText' in generator
    assert 'm_form("Purpose"), plannedReturn' in generator
    assert 'basis = m_form("ReportNo")' not in generator
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    assert any(sheet.attrib["name"] == cfg("history.sheet")
               for sheet in workbook.find("s:sheets", namespace)), "Registry sheet missing"

template = workbook_folder / cfg("generation.templates.VacationTicket").replace("\\", "/")
with zipfile.ZipFile(template) as archive:
    document = ET.fromstring(archive.read("word/document.xml"))
    text = "".join(document.itertext())
    placeholders = re.search(r"    names = VBA.Array\((.*?)\)\n    values", generator, re.S)[1]
    markers = re.findall(r'"(\w+)"', placeholders)
    assert len(markers) == 12
    for marker in markers:
        if marker != "FooterSpacer":
            assert "{" + marker + "}" in text, f"Template placeholder missing: {marker}"

# Validate explicit dependency ranges and source headers against saved XLSX files.
def source_headers(prefix):
    path = (workbook_folder / cfg(prefix + ".path").replace("\\", "/")).resolve()
    match = re.fullmatch(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", cfg(prefix + ".range"))
    assert match, f"Explicit range required: {prefix}"
    with zipfile.ZipFile(path) as archive:
        shared = archive.read("xl/sharedStrings.xml") if "xl/sharedStrings.xml" in archive.namelist() else None
        strings = ["".join(item.itertext()) for item in ET.fromstring(shared)] if shared else []
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        sheet = next(item for item in workbook.find("s:sheets", namespace)
                     if item.attrib["name"] == cfg(prefix + ".sheet"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relationship_id = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
        target = next(item.attrib["Target"] for item in relationships if item.attrib["Id"] == relationship_id)
        target = target.lstrip("/") if target.startswith("/") else "xl/" + target
        data = ET.fromstring(archive.read(target))
        def column_number(text):
            result = 0
            for char in text:
                result = result * 26 + ord(char) - ord("A") + 1
            return result
        first, last = column_number(match[1]), column_number(match[3])
        actual = []
        for cell in data.findall(".//s:c", namespace):
            address = re.fullmatch(r"([A-Z]+)(\d+)", cell.attrib["r"])
            if address[2] != match[2] or not first <= column_number(address[1]) <= last:
                continue
            value = cell.find("s:v", namespace)
            if cell.attrib.get("t") == "s" and value is not None:
                actual.append(strings[int(value.text)])
            elif cell.attrib.get("t") == "inlineStr":
                actual.append("".join(cell.find("s:is", namespace).itertext()))
            elif value is not None:
                actual.append(value.text)
        return actual

assert cfg("snapshots.sources").split("|") == ["lookup.Personnel.source", "orders.source"]
personnel_headers = source_headers("lookup.Personnel.source")
for index in range(1, int(cfg("lookup.Personnel.columns.count")) + 1):
    header = cfg(f"lookup.Personnel.columns.{index}.header")
    assert personnel_headers.count(header) == 1, f"Missing or ambiguous personnel header: {header}"
order_headers = source_headers("orders.source")
for header in (cfg("orders.dateHeader"), cfg("orders.numberHeader")):
    assert order_headers.count(header) == 1, f"Missing or ambiguous order header: {header}"

# Guard the migration boundaries: prepared reads, explicit loading and live registry writes.
loader = (sources / "obj_TDB_DependencyLoader.cls.vba").read_text(encoding="utf-8-sig")
entry = (sources / "ex_TravelDocumentsBuilder.vba").read_text(encoding="utf-8-sig")
workbook_source = (sources / "ThisWorkbook.vba").read_text(encoding="utf-8-sig")
assert "Implements obj_ITableSnapshotLoader" in loader
assert "snapshot.TryLoad(source, diagnostic)" in loader
assert "lookup.PrepareSnapshot(snapshot, diagnostic)" in loader
assert "orders.PrepareSnapshot(snapshot, diagnostic)" in loader
assert loader.count("private_FileStamp(") == 3
assert "ex_TableSnapshots.fn_Register dependencyLoader" in controller
assert "ex_TableSnapshots.fn_Refresh(diagnostic)" in controller
assert "ex_TableDataRuntime.fn_ResetCaches" not in controller
assert "ex_TableSnapshots.fn_Module_Dispose" in entry
assert "obj_OrderResolver" in generator and "m_orderResolver.TrySearch" in generator
assert "ADODB.Connection" not in generator and "generation.orders." not in generator
assert "matches.Count = 0 And Not allowMissing" in generator
assert "If matches.Count = 0 Then Exit Sub" in generator
assert "ex_TDB_History.fn_Open ticketsTable, revisions" in generator
assert "table.ListRows.Add" in generator
assert cfg("candidates.selectionMode") in ("Focus", "Confirm")
candidate_table = next(node for node in markup.iter() if node.attrib.get("name") == "Candidates")
assert candidate_table.attrib["selectionMode"] == "{Binding Path=Settings.CandidateSelectionMode}"
assert "ThisWorkbook::ui.confirmSelectionKey" in config
assert "fn_CancelPending" in workbook_source and "fn_ForgetContext" in workbook_source
assert "Cancel = True" in workbook_source
for key in re.findall(r'(?<!ex_TDB_History\.)fn_Config\("([^"]+)"\)', page + controller + generator + loader + entry + workbook_source):
    cfg(key)
for key in ("OrderSourceError", "OrderInvalidYear", "OrderInvalidInput", "OrderInvalidDate",
            "OrderYearConflict", "OrderDataInvalid", "OrderDuplicate"):
    cfg("text." + key)
assert config["ThisWorkbook::ui.pageMappings"] == "MainPage=MainPage.xaml"
assert (ROOT / "ui" / cfg("ui.folder") / "MainPage.xaml").is_file()
keys = list(config)
assert keys == sorted(keys, key=lambda key: tuple((len(part), part) for part in key.replace("::", ".").split("."))), "Configuration staircase order"
history = (sources / "ex_TDB_History.vba").read_text(encoding="utf-8-sig")
for key in re.findall(r'ex_TDB_History.fn_Config\("([^"]+)"\)', page + controller + generator):
    cfg("history." + key)
for key in re.findall(r'(?<![\w.])fn_Config\("([^"]+)"\)', history):
    cfg("history." + key)
for key in re.findall(r'(?:ex_TDB_History\.)?fn_Fail "([^"]+)"', history + generator + controller):
    cfg("history.text." + key)
for mapping in cfg("history.listFields").split("|"):
    field, caption = mapping.split("=")
    assert field in headers
    cfg("history.text." + caption)
for text in (history, generator, controller):
    assert "ExportJournal" not in text and "ex_PEB_" not in text
assert "private_FindMain" not in generator and "private_FindDonation" not in generator
assert 'identifier = m_form("HistoryRowId")' in generator
assert "ex_TDB_History.fn_AssertCurrent ticketsTable, revisions, transaction" in generator
assert generator.index("ex_TDB_History.fn_AssertCurrent") < generator.index('private_LogStage "WriteRegistry"')
assert "PublishDocument" not in generator
assert "ex_TDB_History.fn_Rollback ticketsTable, revisions, transaction" in generator
assert 'private_CheckRevision revisions, group' in history
assert 'fn_Stamp(group) <> expectedStamp' in history
assert 'If changed Then revision = revision + 1' in history
assert 'If Not changed Then Exit Sub' in history
assert 'snapshot.Add "Snapshot." & field, row(field)' in history
assert 'metadata("DutyPositionCode") = positionCode' in generator
assert 'positionCode = metadata("PositionCode")' in generator
assert 'metadata("RowId") = ex_TDB_History.fn_NewId()' in generator
assert "CoCreateGuid" in history and "Scriptlet.TypeLib" not in history
assert 'If cell.HasFormula Then' in history and 'private_Formulas(tickets)' in history
assert '"HistoryRowId", "HistoryStamp", "ChangeReason"' in page
assert "AddTvoHandler" in controller and "RemoveTvoHandler" in controller
assert "SelectHistoryHandler" in controller and "ex_TDB_History.fn_Load" in controller
assert "ex_TDB_History.fn_List(m_acceptedOrderNumber, m_acceptedOrderDate)" in controller
assert 'private_Text(row(fn_Config("orderNumberField"))) = orderNumber' in history
assert 'fn_Date(row("OrderDate")) = orderDate' in history
assert 'display.Add "HistoryStamp", fn_Stamp(fn_Group(rows, VBA.CStr(key)))' in history
assert "m_historyOrderDate <> m_acceptedOrderDate" in controller
assert "m_previousAutoFill" in controller
assert 'SetValue("Form", "AutoFillPersonnel", False)' in controller
assert 'updates("Form.AutoFillPersonnel") = m_previousAutoFill' in controller
assert "TryGetNearby(matches, diagnostic)" in controller and "TrySearch(VBA.CStr(query), VBA.CStr(year)" in controller
assert 'matches.Count = 1 And VBA.Len' in controller
nodes = {node.get("name"): node for node in markup.iter() if node.get("name")}
for field in cfg("forms.readOnly.fields").split("|"):
    assert nodes[field].get("readOnly") == "{Binding Path=Form." + field + "ReadOnly}"
    assert nodes[field].get("style") == "{Binding Path=Form." + field + "Style}"
assert nodes["AutoFillPersonnel"].get("onChange") == "{Binding Path=Commands.AutoFillChanged}"
assert nodes["OrderQuery"].get("editable") == "true"
assert nodes["History"].get("closeCommand") == "{Binding Path=Commands.CloseHistory}"
peb_markup = ET.parse(ROOT / "ui/PersonalEventBuilder/MainPage.xaml")
for name in ("personnelFormField", "currentOrderField", "lookupKeyField"):
    expected = next(node for node in peb_markup.iter() if node.tag.endswith("controlStyle") and node.get("name") == name)
    assert nodes[name].attrib == expected.attrib, name
for style in ("personnelFormField", "currentOrderField", "formField"):
    selector = f"type=input;part=editor;tags=readOnly {style}"
    actual = next(node for node in markup.iter() if node.get("selector") == selector)
    expected = next(node for node in peb_markup.iter() if node.get("selector") == selector)
    assert actual.get("styles") == expected.get("styles"), selector
assert nodes["TvoName"].get("onChange") == "{Binding Path=Commands.TvoNameChanged}"
exporter = (sources / "obj_TDB_OrderWordExport.cls.vba").read_text(encoding="utf-8-sig")
save_body = re.search(r"Public Function TrySave\(.*?End Function", generator, re.S)[0]
render_body = re.search(r"Public Function TryRender\(.*?End Function", generator, re.S)[0]
assert "private_GenerateWord" not in save_body and "MoveFile" not in save_body
assert "private_GenerateWord" in render_body
assert 'row("DocumentRevision") = revision' not in history
assert 'If Not private_IsDocumentField(VBA.CStr(field)) Then' in history
assert "obj_TDB_OrderWordExport" in controller and "vacationGenerator.TrySave" in controller
assert "ex_TDB_History.fn_OrderGroups" in exporter
assert "private_CheckRevision revisions, group" in history
assert "ex_TDB_History.fn_Commit" not in exporter and "ExportJournal" not in exporter
assert exporter.index("m_vacationGenerator.TryRender") < exporter.index("m_fileSystem.MoveFile VBA.CStr(path), name")
assert exporter.count("ex_TDB_History.fn_AssertCurrent") == 2

# Every write path must re-protect both sheets, including partial unlock and rollback failures.
protection = (sources / "obj_TDB_HistoryProtection.cls.vba").read_text(encoding="utf-8-sig")
assert 'Private Const SHEET_PASSWORD As String = "TDB.Password"' in protection
assert PREFIX + "history.protection.password" not in config
assert 'Password:=SHEET_PASSWORD' in protection
for key in ("ProtectionFailed", "ProtectionUnavailable", "ProtectionNotApplied"):
    cfg("history.text." + key)
assert 'sheet.Cells.Locked = True' in protection
assert 'UserInterfaceOnly:=False, AllowFiltering:=True' in protection
assert 'sheet.EnableSelection = xlNoRestrictions' in protection
assert protection.count('For Each sheet In m_sheets') == 2
assert 'failure = private_ProtectSheet(sheet)' in protection
assert 'If Not Me.TryProtect(diagnostic) Then failures = failures & diagnostic' in protection
assert 'tickets.Parent.ProtectContents Or revisions.Parent.ProtectContents' not in history
assert 'protection.TryProtect(diagnostic)' in history
save_body = re.search(r'Public Function TrySave\([\s\S]*?^End Function', generator, re.M)[0]
save_success, save_failure = save_body.split('\nEH:', 1)
assert save_success.index('protection.Unprotect') < save_success.index('    tableChanged = True')
assert save_success.index('protection.TryProtect') < save_success.index('ex_TDB_HistoryRecovery.fn_Mark')
assert save_failure.index('protection.Unprotect') < save_failure.index('ex_TDB_History.fn_Rollback') < save_failure.index('protection.TryProtect')
assert exporter.index('protection.Unprotect') < exporter.index('    historyChanged = True')
assert exporter.index('protection.TryProtect') < exporter.index('    committed = True')
assert 'protection.TryProtect' in exporter[exporter.index('\nCleanup:'):]
assert 'If root("Status") = "Active" Then' in exporter
assert 'metadata.Add "DocumentPath", VBA.vbNullString' in exporter
assert 'metadata.Add "DocumentRevision", root("Revision")' in exporter
assert "For Each path In owned" in exporter and "If m_namePattern.Test(file.Name) Or m_legacyNamePattern.Test(file.Name) Then" in exporter
assert exporter.index('m_fileSystem.MoveFile m_fileSystem.BuildPath(stage, staged(path)), name') < exporter.index("published.Add name")
assert 'staged.Add identifier, VBA.CStr(wanted.Count) & ".docx"' in exporter
assert 'm_fileSystem.BuildPath(backupFolder, VBA.CStr(backups.Count + 1) & ".docx")' in exporter
assert 'TryRender(root, group, frozenTemplate, activePath, diagnostic)' in exporter
assert 'VBA.Len(frozenTemplate)' in exporter
long_folder = PureWindowsPath("C:/" + "r" * 92) / "TDB_20261026_309"
stage_folder = long_folder / "._TDB-stage-T-01234567-89AB-4CDE-8F01-23456789ABCD"
long_name = "TDB_" + "x" * 72 + "__T-01234567-89AB-4CDE-8F01-23456789ABCD.docx"
assert len(str(long_folder / long_name)) <= int(cfg("word.maxPath"))
assert len(str(stage_folder / long_name)) > 255
assert len(str(stage_folder / "1.docx")) <= int(cfg("word.maxPath"))
assert len(str(stage_folder / "backup" / "1.docx")) <= int(cfg("word.maxPath"))
assert "If stageCreated And VBA.Len(rollback) = 0 Then" in exporter
assert "GetParentFolderName(m_fileSystem.GetAbsolutePathName(stage))" in exporter
assert "ex_TDB_History.fn_Rollback tickets, revisions, transaction" in exporter
assert generator.count('VBA.CreateObject("Word.Application")') == 1
assert 'If m_wordSession Is Nothing Then' in generator
assert 'word.ScreenUpdating = False' in generator
assert 'word.Documents.Open(' in generator
assert 'Visible:=False' in generator
assert 'word.Documents.Add(templatePath' not in generator
assert generator.count('For Each story In document.StoryRanges') == 1
assert 'searchRange.Find.Execute(marker, True, False, False, False, False, True, 0, False)' in generator
index_body = re.search(r'Private Sub private_IndexWordMarkers\([\s\S]*?^End Sub', generator, re.M)[0]
render_word = re.search(r'Private Sub private_GenerateWord\([\s\S]*?^End Sub', generator, re.M)[0]
assert '.Find.' not in render_word
assert 'If m_wordAnchors Is Nothing Then' in render_word
assert 'private_IndexWordMarkers document, names' in render_word
assert 'document.SaveAs2 m_wordTemplatePath, 12' in render_word
assert 'private_OpenWordDocument(word, m_wordTemplatePath)' in render_word
assert 'document.Bookmarks(VBA.CStr(bookmarkName)).Range' in render_word
assert 'document.Bookmarks.Add bookmarkName, searchRange' in index_body
assert 'Do While searchStart < storyEnd' in index_body
assert 'searchRange.SetRange searchStart, storyEnd' in index_body
assert index_body.index('matchEnd = searchRange.End') < index_body.index('document.Bookmarks.Add')
consume_at = index_body.index('searchRange.Text = VBA.Space$(matchEnd - matchStart)')
assert consume_at < index_body.index('document.Bookmarks.Add') < index_body.index('searchStart = matchEnd')
assert 'matchEnd <= searchStart' in index_body
assert 'searchStart = matchEnd' in index_body
assert 'remainingRange' not in index_body
assert 'occurrences.Count = 0 And names(index) <> "FooterSpacer"' in index_body
assert 'Set m_wordAnchors = Nothing' in generator
assert exporter.index('Application.ScreenUpdating = False') < exporter.index('m_vacationGenerator.TryRender')
assert exporter.index('previousWindow.Activate') < exporter.index('Application.ScreenUpdating = previousUpdating')
assert exporter.index('previousSheet.Activate') < exporter.index('Application.ScreenUpdating = previousUpdating')
for stage in ('Document.Prepare', 'Session.Start', 'Document.Create', 'Document.ReplaceMarkers',
              'Document.Save', 'Document.Close', 'Document.Total', 'Document.FailedTotal',
              'Template.IndexMarkers', 'Template.SaveAnchors'):
    assert '"' + stage + '"' in generator, stage
for stage in ('Prepare', 'RenderAll', 'Verify', 'PublishFiles', 'MetadataAndProtection', 'SessionClose', 'Cleanup', 'Total'):
    assert '"TDB.Word.Export.' + stage + '"' in exporter, stage
assert 'ex_Core.fn_Diagnostic_BeginBatch' in exporter
assert 'If diagnosticBatchEntered Then ex_Core.fn_Diagnostic_EndBatch' in exporter
assert exporter.index('m_vacationGenerator.Dispose', exporter.index('\nCleanup:')) < exporter.index('"TDB.Word.Export.Total"')
assert 'orderRecord.TryGetValue("OrderNumber", value)' in controller
assert "CreateDocumentHandler" not in controller and "UpdateDocumentHandler" not in controller
for key in re.findall(r'private_Config\("([^"]+)"\)', exporter):
    cfg("word." + key)
for key in re.findall(r'private_Fail "([^"]+)"', exporter):
    cfg("word.text." + key)
preflight = re.search(r'For Each key In VBA.Split\((.*?), ","\)', exporter, re.S)[1]
for key in "".join(re.findall(r'"([^"]*)"', preflight)).split(","):
    cfg("word." + key)
assert cfg("word.folder") == "TDB_{date}_{number}"
assert "{TicketNo}" in cfg("word.filename.label") and "{FIO}" in cfg("word.filename.label")
pattern = re.compile(cfg("word.filename.pattern"), re.I)
legacy_pattern = re.compile(cfg("word.filename.legacyPattern"), re.I)
identifier = "T-01234567-89AB-4CDE-8F01-23456789ABCD"
for label in ("old_name", "new_name", "2026_121_1_ПІБ"):
    match = pattern.fullmatch(f"TDB_{label}__{identifier}.docx")
    assert match and match[1] == identifier
    assert legacy_pattern.fullmatch(f"TDB_{identifier}__{label}.docx")[1] == identifier
for unmanaged in ("В.к. 2026_121_1.docx", "notes.docx", f"~$TDB_{identifier}__name.docx",
                  f"TDB_{identifier}__name.docx.bak", "TDB_bad-id__name.docx"):
    assert pattern.fullmatch(unmanaged) is None, unmanaged
assert "SaveHistory" in nodes and "ExportWord" in nodes
print(f"PASS: TicketsHistory schemas, separate saving, order-wide WORD synchronization, RowId file ownership, revisions, {len(names)} controls and source contracts.")
print("Excel, Word and VBA were not executed.")