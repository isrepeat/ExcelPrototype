"""Check departure grouping plans and branch selection without running Office."""

from collections import defaultdict
from xml.etree import ElementTree as ET

from event_routes import ROOT, WORKBOOK, NS, load_registry, parse_config


def key(row, fields):
    return tuple(row[field.strip()] for field in fields.split(';'))


def trace(group, rows):
    buckets = defaultdict(list)
    for row in rows:
        buckets[key(row, group.get('by'))].append(row)
    output = []
    for members in buckets.values():
        branch_by = group.get('branchBy')
        single = len({key(row, branch_by) for row in members}) == 1 if branch_by else len(members) == 1
        branch = group.find('p:single' if single else 'p:multiple', NS)
        for child in branch:
            tag = child.tag.rsplit('}', 1)[-1]
            if tag == 'group':
                output.extend(trace(child, members))
            elif tag in ('header', 'footer'):
                assert len({key(row, child.get('common')) for row in members}) == 1
                output.append((tag, ''.join(child.itertext()), len(members)))
            elif tag == 'item':
                output.extend(('person', row['PersonId'], 1) for row in members)
            elif tag == 'text':
                assert len(members) == 1
                output.append(('person', members[0]['PersonId'], 1))
    return output


root = ET.parse(WORKBOOK / 'PersonalEventBuilderWordResultTemplates.xml')
for filename in ('wsConfig.txt', 'wsConfig.test.txt'):
    load_registry(parse_config(ROOT / 'PROJECTS_EXCEL/config/PersonalEventBuilder' / filename))
for name, place, date in [('ToHospital', 'HospitalAccusative', 'EffectiveReturnDateShort'),
                          ('ToBusinessTrip', 'Destination', 'DateFromShort')]:
    template = root.find(f"p:template[@name='{name}']", NS)
    group = template.find('p:group', NS)
    assert group.get('by') == place and group.get('branchBy') == date
    inline = group.find('p:single/p:header/p:text', NS).text
    assert '[' + place + ']' in inline and '[' + date + ']' in inline
    split = group.find('p:multiple/p:header/p:text', NS).text
    assert '[' + place + ']' in split and '[' + date + ']' not in split
    for places, dates, bases in [(1, 1, 1), (1, 2, 1), (2, 2, 1), (1, 1, 2), (1, 2, 2)]:
        rows = []
        for p in range(places):
            for d in range(dates):
                for b in range(bases):
                    for person in range(2):
                        rows.append({place: str(p), date: str(d), 'PersonId': str(len(rows)),
                                     'GroupBasis': str(b), 'VacationTicketDateShort': '01.10.2026',
                                     'BusinessTripPeriod': str(d), 'BusinessTripDetails': 'term'})
        output = trace(group, rows)
        people = [value for tag, value, _ in output if tag == 'person']
        assert sorted(people) == sorted(row['PersonId'] for row in rows)
        headers = [value for tag, value, _ in output if tag == 'header' and '[' + place + ']' in value]
        assert len(headers) == places
print('PASS: both profiles; one/multiple dates, places, people and separate bases; no duplicate place headers or lost people.')