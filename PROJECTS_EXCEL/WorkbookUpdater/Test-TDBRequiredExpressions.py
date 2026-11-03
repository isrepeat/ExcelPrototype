"""Check explicit grouping against legacy truth tables without executing VBA."""

import argparse
import ast
import importlib.util
import itertools
from pathlib import Path
import re


ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('upgrade', Path(__file__).with_name('Upgrade-TDBBasisWorkbooks.py'))
upgrade = importlib.util.module_from_spec(spec)
spec.loader.exec_module(upgrade)
parser = argparse.ArgumentParser()
parser.add_argument('--original', type=Path, required=True)
parser.add_argument('--updated', type=Path, required=True)
args = parser.parse_args()


def records(package):
    _, table = package.tables()['tbBasisDocumentTypes']
    headers = [node.get('name') for node in table.find(f'{{{upgrade.S}}}tableColumns')]
    start, end = table.get('ref').split(':')
    first = int(''.join(filter(str.isdigit, start)))
    last = int(''.join(filter(str.isdigit, end)))
    sheet = upgrade.parse(package.parts[package.sheets()['Типи документів']])
    result = {}
    for row in sheet.find(f'{{{upgrade.S}}}sheetData'):
        if first < int(row.get('r')) <= last:
            cells = {upgrade.index(cell.get('r')): cell for cell in row}
            values = {header: package.text(cells.get(upgrade.index(start) + offset)) for offset, header in enumerate(headers)}
            result[values['Id']] = values
    return result


def evaluate(node, values):
    if isinstance(node, ast.Expression):
        return evaluate(node.body, values)
    if isinstance(node, ast.Name):
        return values[node.id]
    if isinstance(node, ast.BoolOp):
        results = [evaluate(child, values) for child in node.values]
        return all(results) if isinstance(node.op, ast.And) else any(results)
    raise AssertionError(ast.dump(node))


original = upgrade.Package(args.original)
updated = upgrade.Package(args.updated)
before = records(original)
after = records(updated)
assert before.keys() == after.keys()
registration = '((BasisNumber&BasisDate)/(BasisIncomingNumber&BasisIncomingDate))'
assert upgrade.explicit_requirement('BasisNumber&BasisDate/BasisIncomingNumber&BasisIncomingDate') == registration
assert upgrade.explicit_requirement('BasisAuthor|BasisNumber&BasisDate/BasisIncomingNumber&BasisIncomingDate') == '(BasisAuthor&' + registration + ')'
assert upgrade.explicit_requirement(registration) == registration
assert upgrade.explicit_requirement('') == ''
for malformed in ['A&&B', 'A/', '|A', 'A B', '1Field']:
    try:
        upgrade.explicit_requirement(malformed)
    except ValueError:
        pass
    else:
        raise AssertionError(malformed)
assignments = 0
for identifier, old in before.items():
    new = after[identifier]
    assert all(old[key] == new[key] for key in old if key != 'Обов’язкові поля')
    old_rule = old['Обов’язкові поля']
    new_rule = new['Обов’язкові поля']
    assert new_rule == upgrade.explicit_requirement(old_rule)
    assert new_rule == upgrade.explicit_requirement(new_rule)
    names = sorted(set(re.findall(r'[A-Za-z][A-Za-z0-9_]*', old_rule)))
    tree = ast.parse(new_rule.replace('&', ' and ').replace('|', ' and ').replace('/', ' or '), mode='eval')
    for flags in itertools.product([False, True], repeat=len(names)):
        values = dict(zip(names, flags))
        expected = all(any(all(values[field] for field in alternative.split('&'))
                           for alternative in requirement.split('/')) for requirement in old_rule.split('|'))
        assert evaluate(tree, values) == expected, (identifier, values)
        assignments += 1
assert original.parts.keys() == updated.parts.keys()
changed = [name for name in original.parts if original.parts[name] != updated.parts[name]]
assert changed == [original.sheets()['Типи документів']], changed
old_sheet = upgrade.parse(original.parts[changed[0]])
new_sheet = upgrade.parse(updated.parts[changed[0]])
old_cells = {cell.get('r'): cell for cell in old_sheet.findall('.//s:c', upgrade.NS)}
new_cells = {cell.get('r'): cell for cell in new_sheet.findall('.//s:c', upgrade.NS)}
assert old_cells.keys() == new_cells.keys()
assert all(old_cells[key].get('s') == new_cells[key].get('s') for key in old_cells)
source = (ROOT / 'vba/TravelDocumentsBuilder/ex_TDB_BasisChain.vba').read_text(encoding='utf-8-sig')
assert source.count('private_ValidateRequired row, required, False') == 2
assert 'private_ValidateRequired copy, VBA.CStr(copy("BasisRequiredFields"))' in source
assert 'operation <> nextOperation' in source
assert 'private_ValidateLegacyRequired row, specification' in source
assert 'If Not row.Exists(field) Then' in source
assert 'If depth > 64 Then' in source
print(f'Explicit grouping: {len(after)} document profiles, {assignments} truth assignments; styles and other package parts preserved')