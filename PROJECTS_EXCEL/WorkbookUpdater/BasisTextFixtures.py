"""Independent XML fixture reader for static tests; production uses the VBA renderer."""

from pathlib import Path
import re
import xml.etree.ElementTree as ET


class XmlText:
    def __init__(self, folder, fields, maps, profiles, own_order):
        self.folder = Path(folder)
        self.maps = maps
        self.files = {}
        for file in profiles.values():
            self.load(file, set(fields), ())
        self.load(own_order, {'order.number', 'order.date', 'ownUnit.genitive'}, ())

    def load(self, file, fields, trail):
        if Path(file).name != file or '/' in file or '\\' in file or not file.lower().endswith('.xml'):
            raise ValueError(file)
        if file.lower() in [name.lower() for name in trail] or len(trail) > 64:
            raise ValueError('Include cycle: ' + file)
        if file not in self.files:
            text = (self.folder / file).read_text(encoding='utf-8-sig')
            if '<!DOCTYPE' in text.upper():
                raise ValueError('DTD')
            self.files[file] = ET.fromstring(text)
        root = self.files[file]
        if root.tag != 'textTemplate' or root.get('version') != '1':
            raise ValueError(file)
        self.validate(root, fields, trail + (file,))

    def validate(self, node, fields, trail, parent=None):
        schema = {'textTemplate': {'version'}, 'text': set(), 'otherwise': set(), 'value': {'field'},
                  'if': {'field'}, 'ifFilled': {'field'}, 'choose': {'field'}, 'group': {'prefix', 'suffix'},
                  'include': {'file'}, 'case': {'value'}}
        if node.tag not in schema or not set(node.attrib) <= schema[node.tag]:
            raise ValueError(node.tag)
        if node.tag in {'case', 'otherwise'} and parent != 'choose':
            raise ValueError(node.tag)
        if node.tag == 'textTemplate' and parent is not None:
            raise ValueError(node.tag)
        if node.tag == 'textTemplate' and not len(node):
            raise ValueError('Empty template')
        if node.tag in {'value', 'if', 'ifFilled', 'choose'} and node.get('field') not in fields:
            raise ValueError(node.get('field'))
        if node.tag in {'text', 'value', 'include'} and len(node):
            raise ValueError(node.tag)
        if node.tag != 'text' and (node.text or '').strip():
            raise ValueError('Unexpected text')
        if node.tag == 'include':
            self.load(node.get('file', ''), fields, trail)
        if node.tag == 'choose':
            cases = [child.get('value') for child in node if child.tag == 'case']
            if not cases or any(not value for value in cases) or len(cases) != len(set(cases)):
                raise ValueError('Cases')
            if any(child.tag not in {'case', 'otherwise'} for child in node) or sum(child.tag == 'otherwise' for child in node) > 1:
                raise ValueError('Branches')
            mode, pairs = self.maps.get(node.get('field'), ('open', {}))
            if mode == 'closed' and not set(cases) <= set(pairs.values()):
                raise ValueError('Unknown enum case')
        if node.tag == 'if':
            mode, pairs = self.maps.get(node.get('field'), ('open', {}))
            if mode != 'closed' or set(pairs.values()) != {'true', 'false'}:
                raise ValueError('Non-Boolean field')
        for child in node:
            self.validate(child, fields, trail, node.tag)
            if (child.tail or '').strip():
                raise ValueError('Unexpected tail')

    def normalize(self, values):
        result = dict(values)
        for field, (mode, mapping) in self.maps.items():
            value = result.get(field, '')
            if value in mapping:
                result[field] = mapping[value]
            elif mode == 'closed':
                raise ValueError(field)
        return result

    def render(self, file, values, normalize=True):
        values = self.normalize(values) if normalize else values

        def children(node):
            return ' '.join(text for child in node if (text := render(child)))

        def render(node):
            tag = node.tag
            if tag == 'text':
                return node.text or ''
            if tag == 'value':
                return values[node.get('field')]
            if tag == 'include':
                return children(self.files[node.get('file')])
            if tag == 'group':
                value = children(node)
                return node.get('prefix', '') + value + node.get('suffix', '') if value else ''
            if tag == 'ifFilled':
                return children(node) if values[node.get('field')].strip() else ''
            if tag == 'if':
                value = values[node.get('field')]
                if value not in {'true', 'false'}:
                    raise ValueError(value)
                return children(node) if value == 'true' else ''
            if tag == 'choose':
                branches = [child for child in node if child.tag == 'case' and child.get('value') == values[node.get('field')]]
                if not branches:
                    branches = [child for child in node if child.tag == 'otherwise']
                return children(branches[0]) if branches else ''
            raise ValueError(tag)

        result = re.sub(r'[ \t\r\n]+', ' ', children(self.files[file])).strip()
        if not result:
            raise ValueError('Empty template result')
        return result


def value_maps(config):
    result = {}
    for field in config['basisChain.textTemplates.valueMaps.fields'].split('|'):
        prefix = 'basisChain.textTemplates.valueMaps.' + field
        mode = config[prefix + '.mode']
        pairs = [entry.split('~') for entry in config[prefix + '.items'].split('|')]
        if mode not in {'open', 'closed'} or any(len(pair) != 2 for pair in pairs):
            raise ValueError(field)
        if len({pair[0] for pair in pairs}) != len(pairs):
            raise ValueError('Duplicate map input: ' + field)
        result[field] = mode, dict(pairs)
    return result