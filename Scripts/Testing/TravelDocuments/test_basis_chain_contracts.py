"""Check structured basis documents, editable UI and history branch boundaries offline."""

from pathlib import Path
import re
import xml.etree.ElementTree as ET


root = Path(__file__).resolve().parents[3] / 'PROJECTS_EXCEL'
config = {}
for line in (root / 'config/TravelDocumentsBuilder/wsConfig.txt').read_text(encoding='utf-8-sig').splitlines():
    kind, key, value = line.split('\t')
    assert key not in config
    config[key] = value


def cfg(key):
    return config['TravelDocumentsBuilder::' + key]


fields = cfg('basisChain.fields').split('|')
assert fields == ['BasisType', 'BasisInstitution', 'BasisIssuer', 'BasisNumber', 'BasisDate']
for field in fields:
    assert field in cfg('history.columns').split('|')
    assert field in cfg('history.snapshotFields').split('|')
    assert 'Snapshot.' + field in cfg('history.versions.columns').split('|')
assert 'BasisDate' in cfg('history.dateFields').split('|')
assert 'BasisType' not in cfg('history.formFields')

order = cfg('basisChain.orderTemplate').replace('{date}', '11.06.2026').replace('{number}', '170')
extra = cfg('basisChain.documentTemplate')
values = dict(BasisType='повідомлення', BasisIssuer='командира', BasisInstitution='військової частини А7017', BasisNumber='1883/16394', BasisDate='09.06.2026')
for field, value in values.items():
    extra = extra.replace('{' + field + '}', value)
result = order + cfg('basisChain.separator') + ' ' + extra + cfg('basisChain.terminator')
assert result == 'наказ командира військової частини А7383 (по стройовій частині) від 11.06.2026 № 170; повідомлення командира військової частини А7017 № 1883/16394 від 09.06.2026.'
assert '{' not in result

institutions = dict((name, kind) for kind, name in (entry.split('~') for entry in cfg('basisChain.institutions').split('|')))
for institution, kind in institutions.items():
    roles = cfg('basisChain.roles.' + kind).split(',')
    assert roles and len(roles) == len(set(roles))
    if kind == 'TCC':
        assert 'начальника' in roles and 'командира' not in roles
    if kind == 'Unit':
        assert 'командира' in roles and 'заступника командира' in roles
    for issuer in ['', *roles]:
        example = cfg('basisChain.documentTemplate')
        for field, value in dict(values, BasisInstitution=institution, BasisIssuer=issuer).items():
            example = example.replace('{' + field + '}', value)
        example = example.replace('  ', ' ').strip()
        assert '{' not in example and '  ' not in example
        assert institution in example

layout = ET.parse(root / 'ui/TravelDocumentsBuilder/MainPage.xaml')
nodes = {node.attrib['name']: node for node in layout.iter() if 'name' in node.attrib}
assert nodes['BasisInstitution'].attrib['onChange'] == '{Binding Path=Commands.BasisInstitutionChanged}'
for field in fields:
    assert nodes[field].attrib['value'] == '{Binding Path=Form.' + field + '}'
for field in ['BasisType', 'BasisIssuer']:
    assert nodes[field].attrib['editable'] == 'true'
for command in ['AddBasis', 'RemoveBasis']:
    assert nodes[command].attrib['command'] == '{Binding Path=Commands.' + command + '}'
assert nodes['BasisList'].attrib['itemsSource'] == '{Binding Path=Data.BasisView}'
assert all(len(name) + 4 <= 31 for name in ['BasisType', 'BasisIssuer', 'BasisNumber', 'BasisDate', 'AddBasis', 'RemoveBasis'])

sources = root / 'vba/TravelDocumentsBuilder'
chain = (sources / 'ex_TDB_BasisChain.vba').read_text(encoding='utf-8-sig')
history = (sources / 'ex_TDB_History.vba').read_text(encoding='utf-8-sig')
generator = (sources / 'obj_TDB_VacationGenerator.cls.vba').read_text(encoding='utf-8-sig')
controller = (sources / 'obj_TDB_PgMainController.cls.vba').read_text(encoding='utf-8-sig')
graph = re.search(r'Private Sub private_ValidateGraph\([\s\S]*?^End Sub', history, re.M)[0]
assert 'parent("RowType") <> "Event" And parent("RowType") <> "Basis"' in graph
assert 'ex_TDB_BasisChain.fn_Validate row' in graph
assert 'links.Exists(link)' in graph and 'seen.Exists(parent("RowId"))' in graph
assert 'm_form("Purpose") = ex_TDB_BasisChain.fn_Join(m_basis, orderNo, orderDate)' in generator
assert 'metadata("RowType") = "Basis"' in generator
assert 'm_wordValues.Add "BasisDocuments", m_basis' in generator
assert 'Set basisChain = ex_TDB_BasisChain.fn_FromGroup(group)' in history
assert 'updates("Form.Purpose") = VBA.vbNullString' in history
assert 'basisChain.text.Legacy' not in chain + controller + generator
assert 'row("RowId") = fn_Value(bindingContext, "BasisDraftId")' in chain
assert 'If pending Then' in chain and 'result.Add fn_Draft(bindingContext)' in chain
for key in re.findall(r'ex_TravelDocumentsBuilder.fn_Config\("([^"]+)"\)', chain):
    cfg(key)

print('Basis document formatting, controls, schema, history branch and draft contracts: OK')