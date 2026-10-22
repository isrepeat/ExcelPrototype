"""Validate dictionary resources and renderer wiring without Office/VBA."""

from copy import deepcopy
from itertools import permutations
import re
from xml.etree import ElementTree as ET

from event_routes import ROOT
from dsl_resources import load_resources, lookup, order_items
from validate_movement_planner import check_declarations
from validate_movement_batch import check_blocks, check_module_sections
from validate_movement_tvo import preprocess


def main():
    folder = ROOT / 'PROJECTS_EXCEL/workbook/PersonalEventBuilder'
    roots = []
    for filename, default, returned in (
        ('PersonalEventBuilderWordResultTemplates.xml',
         'медична карта стаціонарного хворого', 'виписка із медичної карти стаціонарного хворого'),
        ('PersonalEventBuilderMovementTemplates.xml', 'МКСХ', 'ВМКСХ'),
    ):
        root = ET.parse(folder / filename).getroot()
        roots.append(root)
        resources = load_resources(root, {'SourceStatus', 'ModeId'})
        documents = [{'DocumentNote': name, 'id': index} for index, name in enumerate(
            ('В/к', 'лист', 'Довідка ВЛК', 'інший документ', 'м/к', 'М / К'))]
        for sequence in permutations(documents):
            ordered = order_items(resources, 'BasisDocuments', sequence)
            assert [item for item in ordered if item['id'] in (1, 3)] == [item for item in sequence if item['id'] in (1, 3)]
            assert [item for item in ordered if item['id'] in (4, 5)] == [item for item in sequence if item['id'] in (4, 5)]
            assert [item['id'] for item in ordered][2] == 2
            assert ordered[-1]['id'] == 0
        for value in ('м/к', 'М/К', ' м / к ', '\u00a0М\u202f/\u00a0К\u202f'):
            assert lookup(resources, 'DocumentNames', value, {'SourceStatus': ''}) == default
            assert lookup(resources, 'DocumentNames', value, {'SourceStatus': 'hospitaltreatment'}) == returned
        original = '  Невідомий документ  '
        assert lookup(resources, 'DocumentNames', original, {}) == original
        entry = root.find(".//{*}entry[@key='м/к']")
        reordered = deepcopy(root)
        reversed_entry = reordered.find(".//{*}entry[@key='м/к']")
        reversed_entry[:] = list(reversed(list(reversed_entry)))
        assert lookup(load_resources(reordered, {'SourceStatus'}), 'DocumentNames', 'м/к',
                      {'SourceStatus': 'HospitalTreatment'}) == returned
        assert entry.get('value') is None and all(child.tag.endswith('}variant') for child in entry)
    def reject(change):
        root = deepcopy(roots[0])
        change(root)
        try:
            load_resources(root, {'SourceStatus', 'ModeId'})
        except ValueError:
            return
        raise AssertionError('Malformed resources accepted')
    reject(lambda root: root.find('.//{*}dictionary').append(deepcopy(root.find('.//{*}entry'))))
    reject(lambda root: root.find('.//{*}entry').remove(root.find('.//{*}variant')))
    reject(lambda root: root.find('.//{*}entry').append(deepcopy(root.find('.//{*}variant'))))
    reject(lambda root: root.find('.//{*}variant[@field]').set('field', 'UnknownField'))
    reject(lambda root: root.find('.//{*}dictionary').set('normalize', 'regex'))
    reject(lambda root: root.find('.//{*}dictionary').set('onMissing', 'fallback'))
    reject(lambda root: root.find('.//{*}dictionary').set('ignoreCase', 'yes'))
    reject(lambda root: root.find('.//{*}entry').set('key', '\u00a0'))
    reject(lambda root: root.find('.//{*}itemOrder').remove(root.find('.//{*}others')))
    reject(lambda root: root.find('.//{*}itemOrder').append(deepcopy(root.find('.//{*}others'))))
    reject(lambda root: root.find('.//{*}itemOrder').append(deepcopy(root.find('.//{*}item'))))
    reject(lambda root: root.find('.//{*}itemOrder').attrib.pop('field'))
    reject(lambda root: root.find('.//{*}itemOrder').set('ignoreCase', 'yes'))
    reject(lambda root: root.find('.//{*}itemOrder').set('normalize', 'unknown'))
    generic = ET.fromstring("""
      <root xmlns="urn:excelprototype:profiles">
        <resources>
          <itemOrder name="Priorities" field="Priority" ignoreCase="false" normalize="none">
            <item key="High"/><others/><item key="Low"/>
          </itemOrder>
        </resources>
      </root>""")
    generic_resources = load_resources(generic, set())
    items = [{'Priority': value} for value in ('Low', 'high', ' High ', 'High', 'Other', 'High')]
    ordered = order_items(generic_resources, 'Priorities', items)
    assert ordered == [items[i] for i in (3, 5, 1, 2, 4, 0)]
    assert items[0]['Priority'] == 'Low'
    try:
        order_items(generic_resources, 'Priorities', [{'OtherField': 'High'}])
    except KeyError:
        pass
    else:
        raise AssertionError('Missing item field accepted')
    reject(lambda root: root.find('.//{*}text').__setattr__('text', '{[DocumentNote]|lookup:"Missing"}'))
    root = deepcopy(roots[0])
    variant = ET.SubElement(root.find('.//{*}entry'), '{urn:excelprototype:profiles}variant',
                            {'field': 'ModeId', 'equals': 'FromHospital', 'value': 'other'})
    resources = load_resources(root, {'SourceStatus', 'ModeId'})
    assert lookup(resources, 'DocumentNames', 'м/к',
                  {'SourceStatus': '', 'ModeId': 'FromHospital'}) == 'other'
    assert lookup(resources, 'DocumentNames', 'м/к',
                  {'SourceStatus': '', 'ModeId': 'OtherMode'}) == 'медична карта стаціонарного хворого'
    try:
        lookup(resources, 'DocumentNames', 'м/к',
               {'SourceStatus': 'HospitalTreatment', 'ModeId': 'FromHospital'})
    except ValueError:
        pass
    else:
        raise AssertionError('Ambiguous variants accepted')
    dictionary = root.find('.//{*}dictionary')
    dictionary.set('onMissing', 'error')
    try:
        lookup(load_resources(root, {'SourceStatus', 'ModeId'}), 'DocumentNames', 'missing', {})
    except ValueError:
        pass
    else:
        raise AssertionError('Unknown key accepted with onMissing=error')
    source_folder = ROOT / 'PROJECTS_EXCEL/vba/PersonalEventBuilder'
    for filename in ('ex_PEB_DslResources.vba', 'obj_PEB_DslRenderer.cls.vba',
                     'obj_PEB_WordTemplates.cls.vba', 'ex_PEB_MovementTemplates.vba',
                     'ex_PEB_WordDocument.vba', 'ex_PEB_WordGroups.vba'):
        source = (source_folder / filename).read_text(encoding='utf-8-sig')
        check_declarations(source)
        check_blocks(preprocess(source, True))
        if filename == 'ex_PEB_DslResources.vba':
            check_module_sections(source)
            assert not re.search(r'(?mi)^\s*Dim\s+attribute\s+As\b', source), 'Attribute is reserved in VBA'
            assert not re.search(r'(?mi)^\s*Next\s+attribute\s*$', source), 'Stale loop variable after rename'
    renderer = (source_folder / 'obj_PEB_DslRenderer.cls.vba').read_text(encoding='utf-8-sig')
    sorter = (source_folder / 'ex_PEB_DslResources.vba').read_text(encoding='utf-8-sig')
    sorter = sorter.split('Public Function fn_OrderItems(', 1)[1].split('End Function', 1)[0]
    assert 'DocumentNote' not in sorter and 'itemOrder("Field")' in sorter
    assert 'fn_OrderItems(m_resources, orderName, items)' in renderer
    assert 'Case "lookup"' in renderer and 'fn_Lookup(m_resources, argument' in renderer
    store = (source_folder / 'ex_PEB_WordDocument.vba').read_text(encoding='utf-8-sig')
    assert 'text = resourcesXml' in store and 'VBA.Len(templates.ResourcesXml)' in store
    for root in roots:
        text = root.find(".//{*}sharedTemplate[@id='block.DocumentName']/{*}text").text
        assert 'regexreplace' not in text and 'm_RegexIsMatch' not in text
        assert 'lookup:"DocumentNames"' in text
    print('PASS: dictionary normalization, case, variants, missing keys, malformed resources, '
          'renderer wiring and Word fingerprints. Office/VBA not executed.')


if __name__ == '__main__':
    main()