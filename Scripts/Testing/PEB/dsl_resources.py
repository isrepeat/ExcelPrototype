"""Offline validation and reference lookup for declarative DSL dictionaries."""

import re

NS = {'p': 'urn:excelprototype:profiles'}


def normalize(value, policy):
    if policy == 'none':
        return value
    value = value.replace('\u00a0', ' ').replace('\u202f', ' ').replace('\t', ' ').strip(' ')
    if policy == 'trim,slashSpaces':
        while ' /' in value or '/ ' in value:
            value = value.replace(' /', '/').replace('/ ', '/')
    return value


def load_resources(root, known):
    def require(condition, message):
        if not condition:
            raise ValueError('DSL resources: ' + message)
    def attributes(node, allowed, required=()):
        require(set(node.attrib) <= set(allowed), 'Unknown attribute')
        require(all(node.get(key, '').strip(' ') for key in required), 'Missing attribute')
    def children(node, allowed):
        require(all(child.tag in {f"{{{NS['p']}}}{name}" for name in allowed} for child in node), 'Unknown element')
        require(not (node.text or '').strip(), 'Unexpected text')
        require(all(not (child.tail or '').strip() for child in node), 'Unexpected text')
    sections = root.findall('p:resources', NS)
    require(len(sections) <= 1, 'Duplicate resources')
    result = {}
    for section in sections:
        attributes(section, ())
        children(section, ('dictionary', 'itemOrder'))
        for node in section.findall('p:dictionary', NS):
            attributes(node, ('name', 'ignoreCase', 'normalize', 'onMissing'),
                       ('name', 'ignoreCase', 'normalize', 'onMissing'))
            children(node, ('entry',))
            name = node.get('name')
            require(name not in result, 'Duplicate dictionary')
            require(node.get('ignoreCase') in ('true', 'false'), 'Invalid ignoreCase')
            require(node.get('normalize') in ('trim', 'trim,slashSpaces'), 'Invalid normalization')
            require(node.get('onMissing') in ('keep', 'error'), 'Invalid onMissing')
            dictionary = dict(node.attrib, entries={})
            def fold(value):
                return value.casefold() if dictionary['ignoreCase'] == 'true' else value
            for entry in node:
                attributes(entry, ('key',), ('key',))
                children(entry, ('variant',))
                key = fold(normalize(entry.get('key'), dictionary['normalize']))
                require(key and key not in dictionary['entries'], 'Duplicate/empty normalized key')
                value = {'variants': {}}
                for variant in entry:
                    attributes(variant, ('field', 'equals', 'value'), ('value',))
                    children(variant, ())
                    if 'field' not in variant.attrib:
                        require('equals' not in variant.attrib and 'default' not in value, 'Invalid default')
                        value['default'] = variant.get('value')
                    else:
                        field = variant.get('field')
                        require(field in known and 'equals' in variant.attrib, 'Invalid condition')
                        pair = (field, fold(variant.get('equals')))
                        require(pair not in value['variants'], 'Duplicate condition')
                        value['variants'][pair] = variant.get('value')
                require('default' in value, 'Missing default')
                dictionary['entries'][key] = value
            require(dictionary['entries'], 'Empty dictionary')
            result[name] = dictionary
        for node in section.findall('p:itemOrder', NS):
            attributes(node, ('name', 'field', 'ignoreCase', 'normalize'), ('name', 'field', 'ignoreCase', 'normalize'))
            children(node, ('item', 'others'))
            name = '@order:' + node.get('name')
            require(name not in result, 'Duplicate item order')
            require(node.get('ignoreCase') in ('true', 'false'), 'Invalid item order ignoreCase')
            require(node.get('normalize') in ('none', 'trim', 'trim,slashSpaces'), 'Invalid item order normalization')
            ranks = {}
            for rank, child in enumerate(node):
                children(child, ())
                if child.tag.endswith('}others'):
                    attributes(child, ())
                    require('@others' not in ranks, 'Duplicate others')
                    ranks['@others'] = rank
                else:
                    attributes(child, ('key', 'aliases'), ('key',))
                    for key in (child.get('key') + '|' + child.get('aliases', '')).strip('|').split('|'):
                        key = normalize(key, node.get('normalize'))
                        if node.get('ignoreCase') == 'true':
                            key = key.casefold()
                        require(key and key.casefold() != '@others' and key not in ranks, 'Duplicate/empty order key')
                        ranks[key] = rank
            require('@others' in ranks, 'Missing others')
            result[name] = dict(node.attrib, ranks=ranks)
    for node in root.findall('.//p:text', NS):
        for name in re.findall(r'orderBy\s*:\s*"([^"]+)"', node.text or ''):
            require('@order:' + name in result, 'Unknown item order')
        for name in re.findall(r'\|lookup\s*:\s*"([^"]+)"', node.text or ''):
            require(name in result, 'Unknown dictionary: ' + name)
    return result


def order_items(resources, name, items):
    order = resources['@order:' + name]
    ranks = order['ranks']
    def rank(item):
        key = normalize(item[order['field']], order['normalize'])
        if order['ignoreCase'] == 'true':
            key = key.casefold()
        return ranks.get(key, ranks['@others'])
    return sorted(items, key=rank)


def lookup(resources, name, value, context):
    dictionary = resources[name]
    def fold(text):
        return text.casefold() if dictionary['ignoreCase'] == 'true' else text
    key = fold(normalize(value, dictionary['normalize']))
    if key not in dictionary['entries']:
        if dictionary['onMissing'] == 'keep':
            return value
        raise ValueError('Unknown dictionary key')
    entry = dictionary['entries'][key]
    matches = [text for (field, expected), text in entry['variants'].items()
               if fold(str(context[field])) == expected]
    if len(matches) > 1:
        raise ValueError('Ambiguous dictionary variants')
    return matches[0] if matches else entry['default']