"""Read source data and configuration separately for offline TDB checks."""


def configured_rows(config, kind):
    prefix = 'basisChain.model.' + kind
    keys = config[prefix + '.keys'].split('|')
    if kind == 'Fields':
        headers = {'Name': 'Назва', 'Description': 'Опис', 'Type': 'Тип', 'Usage': 'Використання'}
        return [{'Код': key, **{header: config[prefix + '.' + key + '.' + field] for field, header in headers.items()}} for key in keys]
    if kind in {'OfficialRules', 'DocumentRules'}:
        header = 'Id документа' if kind == 'DocumentRules' else 'Id посадової особи'
        return [{'Група установи': key, header: identifier} for key in keys for identifier in config[prefix + '.' + key].split('|')]
    headers = {'Name': 'Назва', **({'Forms': 'Форми', 'Required': 'Обов’язкові поля', 'Fields': 'Поля'} if kind == 'Documents' else {})}
    return [{'Id': key, **{header: config[prefix + '.' + key + '.' + field] for field, header in headers.items()}} for key in keys]


def enrich(config, kind, source):
    names = {}
    for row in source:
        name = row['Назва'].casefold()
        assert name not in names, 'Duplicate source name: ' + row['Назва']
        names[name] = row
    result = []
    for row in configured_rows(config, kind):
        result.append({**names[row['Назва'].casefold()], **row})
    return result