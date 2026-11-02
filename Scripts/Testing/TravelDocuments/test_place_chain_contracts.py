"""Validate place-chain persistence and UI contracts without executing VBA."""

from pathlib import Path
import re
import xml.etree.ElementTree as ET


root = Path(__file__).resolve().parents[3] / 'PROJECTS_EXCEL'
sources = root / 'vba/TravelDocumentsBuilder'
config = {}
for line in (root / 'config/TravelDocumentsBuilder/wsConfig.txt').read_text(encoding='utf-8-sig').splitlines():
    kind, key, value = line.split('\t')
    assert key not in config, key
    config[key] = value

chain = (sources / 'ex_TDB_PlaceChain.vba').read_text(encoding='utf-8-sig')
controller = (sources / 'obj_TDB_PgMainController.cls.vba').read_text(encoding='utf-8-sig')
generator = (sources / 'obj_TDB_VacationGenerator.cls.vba').read_text(encoding='utf-8-sig')
history = (sources / 'ex_TDB_History.vba').read_text(encoding='utf-8-sig')
page = (sources / 'obj_TDB_PgMain.cls.vba').read_text(encoding='utf-8-sig')

for key in re.findall(r'ex_TravelDocumentsBuilder.fn_Config\("([^"]+)"\)', chain + controller):
    assert 'TravelDocumentsBuilder::' + key in config, key

# Each Place child retains structured domestic and foreign fields.
assert len(config['TravelDocumentsBuilder::history.columns'].split('|')) == 53
assert len(config['TravelDocumentsBuilder::history.versions.columns'].split('|')) == 60
for field in ('Destination', 'PlaceRegion', 'PlaceType', 'PlaceName', 'PlaceDistrict'):
    assert field in config['TravelDocumentsBuilder::history.snapshotFields'].split('|')

graph = re.search(r'Private Sub private_ValidateGraph\([\s\S]*?^End Sub', history, re.M)[0]
assert 'row("RowType") <> "Place"' in graph
assert 'parent("RowType") <> "Event" And parent("RowType") <> "Place"' in graph
assert 'links.Exists(link)' in graph and 'seen.Exists(parent("RowId"))' in graph
assert 'Set places = ex_TDB_PlaceChain.fn_FromGroup(group)' in history
assert 'If places.Count > 0 Then ex_TDB_PlaceChain.fn_ClearDraft bindingContext' in history
load_single = re.search(r'If places.Count = 1 Then[\s\S]*?End If', history)[0]
assert 'updates("Form.DocumentTarget") = row("Destination")' in load_single
assert 'For Each placeField In ex_TDB_Places.fn_FormFields()' in load_single
assert 'updates("Form." & placeField) = row(placeField)' in load_single
assert 'Set places = New Collection' in load_single

# Save includes the current draft, preserves loaded child IDs, and chains each next parent.
assert 'If hasDraft Then result.Add fn_FromValues(values)' in chain
assert 'Set m_places = ex_TDB_PlaceChain.fn_Build(m_bindingContext, m_form)' in generator
assert 'm_form("DocumentTarget") = ex_TDB_PlaceChain.fn_Join(m_places)' in generator
assert 'historyRow("RowType") = "TVO" Or historyRow("RowType") = "Place"' in generator
write = re.search(r'For Each metadata In m_places[\s\S]*?Next metadata', generator)[0]
assert 'If VBA.Len(metadata("RowId")) = 0 Then' in write
assert 'metadata("ParentRowId") = parentId' in write
assert 'parentId = metadata("RowId")' in write
assert 'metadata("Status") = "Cancelled"' in write
assert 'm_wordValues.Add "PlaceChain", placeChain' in generator

# Single-place storage omits children after removing any previous route children.
storage = generator[generator.index("    ' Replace only this event's ordered TVO and place children"):generator.index('    private_LogStage "CommitHistoryRevision"')]
assert storage.index('row.Delete') < storage.index('If m_places.Count > 1 Then')
assert re.search(r'If m_places.Count > 1 Then\s+For Each metadata In m_places[\s\S]*?Next metadata\s+End If', storage)
assert 'For Each key In ex_TDB_Places.fn_FormFields()\n        metadata.Add VBA.CStr(key), m_form(key)' in generator

# Loading/export reconstruct an Event-only place without copying Event's identity.
restore = re.search(r'Public Function fn_FromGroup\([\s\S]*?^End Function', chain, re.M)[0]
single = restore[restore.index('    Set root = group(1)'):]
assert 'If result.Count = 0 And VBA.Len(VBA.Trim$(VBA.CStr(root("PlaceName")))) > 0 Then' in single
assert 'Set row = ex_TDB_History.fn_NewRow()' in single
assert 'VBA.Split("Destination|" & ex_TravelDocumentsBuilder.fn_Config("placeChain.fields"), "|")' in single
assert 'row(field) = root(field)' in single
assert 'RowId' not in single and 'ParentRowId' not in single
assert 'result.Add row' in single
assert restore.index('If result.Count <> count Then') < restore.index('    Set root = group(1)')

layout = ET.parse(root / 'ui/TravelDocumentsBuilder/MainPage.xaml')
for name in ('AddPlace', 'RemovePlace'):
    node = next(node for node in layout.iter() if node.attrib.get('name') == name)
    assert node.attrib['command'] == '{Binding Path=Commands.' + name + '}'
    assert f'"{name}", "{name}Handler"' in page
    assert len('btn_' + name) <= 31
assert 'rowSpan="3"' not in next(
    ET.tostring(node, encoding='unicode') for node in layout.iter()
    if node.attrib.get('name') == 'PlaceChainPanel'
)
assert 'fn_IsMedicalCompany' in generator and 'Dim emptyPlaces As New Collection' in controller
print('Place-chain schema, save/load/export, graph guards and UI contracts: OK')