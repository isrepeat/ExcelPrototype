import fs from 'node:fs/promises';
import path from 'node:path';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

// Seed data for the reference-workbook upgrade; runtime settings live in wsConfig.
const outputDir = path.resolve(process.argv[2]);
const book = Workbook.create();
const tables = [];
const configText = await fs.readFile(path.resolve(process.cwd(), 'PROJECTS_EXCEL/config/TravelDocumentsBuilder/wsConfig.txt'), 'utf8');
const settings = new Map(configText.replace(/^\uFEFF/, '').split(/\r?\n/).filter(Boolean).map(line => {
  const [, key, value] = line.split('\t');
  return [key.replace('TravelDocumentsBuilder::', ''), value];
}));
const bindingPairs = settings.get('basisChain.contract.bindings').split('|').map(entry => entry.split('='));
const publicCodes = new Map(bindingPairs.map(([code, target]) => [target.replace('Form.', ''), code]));
const legacyTokens = new Map(settings.get('basisChain.contract.legacyTokens').split('|').map(entry => entry.split('=')));
function publicCondition(text) {
  return text.replace(/[A-Za-z][A-Za-z0-9_]*/g, field => {
    if (!publicCodes.has(field)) throw new Error('Unmapped seed field: ' + field);
    return publicCodes.get(field);
  });
}
function publicTemplate(text) {
  return text.replace(/\{([^}]+)\}/g, (_, token) => {
    const code = legacyTokens.get(token) || publicCodes.get(token);
    if (!code) throw new Error('Unmapped seed token: ' + token);
    return '{' + code + '}';
  });
}
function reference(sheetName, tableName, headers, rows, widths) {
  const sheet = book.worksheets.add(sheetName);
  const end = String.fromCharCode(65 + headers.length);
  sheet.getRange(`A1:${end}${rows.length + 6}`).format = {
    fill: '#27251B', font: { name: 'Arial', size: 11, color: '#FFFFFF' },
  };
  sheet.getRange('A:A').format.columnWidth = 2;
  sheet.getRange('B2').values = [[sheetName]];
  sheet.getRange('B2').format.font = { bold: true, size: 16, color: '#FFB347' };
  sheet.getRange(`B5:${end}${rows.length + 5}`).values = [headers, ...rows];
  sheet.getRange(`B5:${end}5`).format = {
    fill: '#625740', font: { bold: true, color: '#FFFFFF' }, wrapText: true, rowHeight: 39.95,
    horizontalAlignment: 'center', verticalAlignment: 'center',
  };
  sheet.getRange(`B6:${end}${rows.length + 5}`).format.wrapText = true;
  sheet.getRange(`B6:${end}${rows.length + 5}`).format.rowHeight = 23.1;
  sheet.getRange(`B6:${end}${rows.length + 5}`).format.verticalAlignment = 'center';
  widths.forEach((width, index) => {
    const column = String.fromCharCode(66 + index);
    sheet.getRange(`${column}:${column}`).format.columnWidth = width;
  });
  const table = sheet.tables.add(`B5:${end}${rows.length + 5}`, true, tableName);
  table.style = 'TableStyleMedium15';
  table.showFilterButton = true;
  table.showBandedColumns = false;
  tables.push({ sheetName, tableName, rows: rows.length });
}

const officials = [
  ['commander', 'Командир', 'командира'],
  ['deputy_commander', 'Заступник командира', 'заступника командира'],
  ['chief', 'Начальник', 'начальника'],
  ['deputy_chief', 'Заступник начальника', 'заступника начальника'],
  ['chief_of_staff', 'Начальник штабу', 'начальника штабу'],
  ['chief_of_staff_deputy', 'Начальник штабу — заступник командира', 'начальника штабу — заступника командира'],
  ['personnel_chief', 'Начальник відділення персоналу штабу', 'начальника відділення персоналу штабу'],
  ['administrative_chief', 'Начальник адміністративної групи штабу', 'начальника адміністративної групи штабу'],
  ['commander_in_chief', 'Головнокомандувач', 'Головнокомандувача'],
  ['forces_commander', 'Командувач', 'командувача'],
];
reference('Посадові особи', 'tbBasisOfficials', ['Id', 'Назва', 'Текст у підставі'], officials, [29, 48, 55]);
const officialRules = [
  ...['commander', 'deputy_commander', 'chief_of_staff', 'chief_of_staff_deputy', 'personnel_chief', 'administrative_chief'].map(id => ['Unit', id]),
  ...['Hospital', 'TCC'].flatMap(group => ['chief', 'deputy_chief'].map(id => [group, id])),
  ['ChiefCommand', 'commander_in_chief'], ['GeneralStaff', 'chief'], ['LandCommand', 'forces_commander'],
];
reference('Посади установ', 'tbBasisOfficialRules', ['Група установи', 'Id посадової особи'], officialRules, [27, 38]);

const common = ['BasisIssuer', 'BasisActing', 'BasisNumber', 'BasisDate', 'BasisIncomingNumber', 'BasisIncomingDate', 'BasisForm'];
const registration = '((BasisNumber&BasisDate)/(BasisIncomingNumber&BasisIncomingDate))';
const ordinary = '{Form} {BasisType} {Acting} {BasisIssuerText} {BasisInstitutionName} {Scope} {Number} {Date} {Incoming}';
const documents = [
  ['notification', 'повідомлення', 'повідомлення', ordinary, registration, common.join('|'), 'Оригінал,Копія'],
  ['letter', 'лист', 'листа', ordinary, registration, common.join('|'), 'Оригінал,Копія'],
  ['order', 'наказ', 'наказу', ordinary, '(BasisNumber&BasisDate)', [...common, 'BasisScope'].join('|'), 'Оригінал,Копія,Витяг'],
  ['instruction', 'розпорядження', 'розпорядження', ordinary, '(BasisNumber&BasisDate)', common.join('|'), 'Оригінал,Копія,Витяг'],
  ['separate_instruction', 'окреме доручення', 'окремого доручення', ordinary, '(BasisNumber&BasisDate)', common.join('|'), 'Оригінал,Копія,Витяг'],
  ['report', 'рапорт', 'рапорту', '{Form} {BasisType} {BasisAuthor} {Number} {Date} {Incoming}', `(BasisAuthor&${registration})`, ['BasisAuthor', 'BasisNumber', 'BasisDate', 'BasisIncomingNumber', 'BasisIncomingDate', 'BasisForm'].join('|'), 'Оригінал,Копія'],
  ['briefing', 'доповідь', 'доповіді', ordinary, registration, common.join('|'), 'Оригінал,Копія'],
  ['travel_certificate', 'посвідчення про відрядження', 'посвідчення про відрядження', '{Form} {BasisType} {Number} {Date} {Incoming}', '(BasisNumber&BasisDate)', ['BasisNumber', 'BasisDate', 'BasisIncomingNumber', 'BasisIncomingDate', 'BasisForm'].join('|'), 'Оригінал,Копія'],
  ['court_order', 'ухвала', 'ухвали', '{Form} {BasisType} {BasisInstitutionName} {Date} у справі № {BasisCaseNumber} {BasisProceedingNumber} {Incoming}', '(BasisDate&BasisCaseNumber)', ['BasisDate', 'BasisCaseNumber', 'BasisProceedingNumber', 'BasisIncomingNumber', 'BasisIncomingDate', 'BasisForm'].join('|'), 'Оригінал,Копія'],
  ['certificate', 'довідка', 'довідки', '{Form} {BasisType} {BasisInstitutionName} {Number} {Date} {Incoming}', '(BasisNumber&BasisDate)', ['BasisNumber', 'BasisDate', 'BasisIncomingNumber', 'BasisIncomingDate', 'BasisForm'].join('|'), 'Оригінал,Копія'],
  ['personnel_notification', 'повідомлення про прийняття кадрових рішень', 'повідомлення про прийняття кадрових рішень', '{Form} {BasisType} наказом {BasisRefIssuer} {RefScope} від {BasisRefDate} № {BasisRefNumber} {Number} {Date} {Incoming}', '(BasisRefIssuer&BasisRefNumber&BasisRefDate)', [...common.filter(f => !['BasisIssuer', 'BasisActing'].includes(f)), 'BasisRefIssuer', 'BasisRefNumber', 'BasisRefDate', 'BasisRefScope'].join('|'), 'Оригінал,Копія'],
];
reference('Типи документів', 'tbBasisDocumentTypes', ['Id', 'Назва', 'Родовий', 'Шаблон', 'Обов’язкові поля', 'Поля', 'Форми'],
  documents.map(([id, name, genitive, template, required, fields, forms]) =>
    [id, name, genitive, publicTemplate(template), publicCondition(required), publicCondition(fields), forms]),
  [25, 38, 38, 80, 52, 75, 28]);
const documentRules = [
  ...['notification', 'letter', 'order', 'instruction', 'separate_instruction', 'report', 'briefing', 'travel_certificate', 'personnel_notification'].map(id => ['Unit', id]),
  ...['notification', 'letter', 'certificate'].map(id => ['Hospital', id]),
  ...['ChiefCommand', 'GeneralStaff', 'LandCommand'].flatMap(group => ['notification', 'letter', 'order', 'instruction', 'separate_instruction', 'personnel_notification'].map(id => [group, id])),
  ['Court', 'court_order'],
];
reference('Документи установ', 'tbBasisDocumentRules', ['Група установи', 'Id документа'], documentRules, [27, 38]);
reference('Органи управління', 'tbCommands', ['Група установи', 'Назва', 'Родовий'], [
  ['ChiefCommand', 'Збройні Сили України', 'Збройних Сил України'],
  ['GeneralStaff', 'Генеральний штаб Збройних Сил України', 'Генерального штабу Збройних Сил України'],
  ['LandCommand', 'Сухопутні війська Збройних Сил України', 'Сухопутних військ Збройних Сил України'],
], [24, 58, 64]);
reference('Суди', 'tbCourts', ['Група установи', 'Назва', 'Родовий'], [
  ['Court', 'Новобаварський районний суд міста Харкова', 'Новобаварського районного суду міста Харкова'],
  ['Court', 'Київський районний суд міста Харкова', 'Київського районного суду міста Харкова'],
], [24, 58, 64]);

const fieldDictionary = [
  [
    "document.type",
    "Тип документа",
    "Назва вибраного типу документа без зміни відмінка.",
    "choice",
    "computed"
  ],
  [
    "document.type.genitive",
    "Тип документа (родовий)",
    "Родовий відмінок назви для копії або витягу.",
    "text",
    "computed"
  ],
  [
    "institution.genitive",
    "Установа (родовий)",
    "Назва вибраної установи в родовому відмінку.",
    "text",
    "computed"
  ],
  [
    "issuer.role",
    "Посадова особа",
    "Посада автора документа в називному відмінку.",
    "choice",
    "input"
  ],
  [
    "issuer.role.genitive",
    "Посада у підставі",
    "Збережений родовий відмінок посади автора.",
    "text",
    "computed"
  ],
  [
    "issuer.acting",
    "Виконання обов’язків",
    "Ознака тимчасового виконання обов’язків.",
    "choice",
    "input"
  ],
  [
    "document.number",
    "Номер документа",
    "Власний номер документа без знака №.",
    "text",
    "input"
  ],
  [
    "document.date",
    "Дата документа",
    "Дата підписання документа у форматі дд.мм.рррр.",
    "date",
    "input"
  ],
  [
    "registration.number",
    "Вхідний номер",
    "Номер реєстрації отриманого документа без знака №.",
    "text",
    "input"
  ],
  [
    "registration.date",
    "Дата реєстрації",
    "Дата вхідної реєстрації у форматі дд.мм.рррр.",
    "date",
    "input"
  ],
  [
    "document.form",
    "Форма документа",
    "Оригінал, копія або витяг відповідно до профілю документа.",
    "choice",
    "input"
  ],
  [
    "document.scope",
    "Різновид наказу",
    "Уточнення: по особовому складу, по стройовій частині тощо.",
    "choice",
    "input"
  ],
  [
    "court.caseNumber",
    "Номер справи",
    "Номер судової справи без службового префікса.",
    "text",
    "input"
  ],
  [
    "court.proceedingNumber",
    "Номер провадження",
    "Номер судового провадження без знака № та слова «провадження».",
    "text",
    "input"
  ],
  [
    "document.author",
    "Автор (у родовому)",
    "Повна фраза про автора: посада, звання та ПІБ у родовому відмінку.",
    "text",
    "input"
  ],
  [
    "reference.issuer",
    "Автор наказу (у родовому)",
    "Повна фраза про автора наказу, на який посилається повідомлення.",
    "text",
    "input"
  ],
  [
    "reference.number",
    "Номер наказу",
    "Номер наказу, на який посилається документ.",
    "text",
    "input"
  ],
  [
    "reference.date",
    "Дата наказу",
    "Дата наказу, на який посилається документ.",
    "date",
    "input"
  ],
  [
    "reference.scope",
    "Уточнення наказу",
    "Різновид наказу, на який посилається документ.",
    "choice",
    "input"
  ]
];
reference('Поля документів', 'tbBasisFields', ['Код', 'Назва', 'Опис', 'Тип', 'Використання'], fieldDictionary, [30, 35, 80, 16, 18]);
const fragments = [
  ['document.numberText', 'Номер із позначкою', 'Порожній номер не додає знак №.', 'nonempty', 'document.number', '', '№ {document.number}'],
  ['document.dateText', 'Дата з прийменником', 'Порожня дата не додає слово «від».', 'nonempty', 'document.date', '', 'від {document.date}'],
  ['registration.text', 'Вхідні реквізити', 'Уся фраза в дужках; номер і дата перевіряються разом.', 'nonempty', 'registration.number', '', '(вх. № {registration.number} від {registration.date})'],
  ['issuer.actingText', 'Тимчасове виконання обов’язків', 'Фраза додається лише для вибраної ознаки ТВО.', 'equals', 'issuer.acting', settings.get('basisChain.acting.option'), settings.get('basisChain.acting.text')],
  ['document.formText', 'Позначка копії', 'Для оригіналу цей фрагмент порожній.', 'equals', 'document.form', settings.get('basisChain.forms.copy'), settings.get('basisChain.forms.copyText')],
  ['document.formText', 'Позначка витягу', 'Для оригіналу цей фрагмент порожній.', 'equals', 'document.form', settings.get('basisChain.forms.extract'), settings.get('basisChain.forms.extractText')],
  ['document.typeText', 'Тип документа для копії', 'Копія використовує родовий відмінок назви документа.', 'equals', 'document.form', settings.get('basisChain.forms.copy'), '{document.type.genitive}'],
  ['document.typeText', 'Тип документа для витягу', 'Витяг використовує родовий відмінок назви документа.', 'equals', 'document.form', settings.get('basisChain.forms.extract'), '{document.type.genitive}'],
  ['document.typeText', 'Тип документа для оригіналу', 'Використовується, якщо жодна спеціальна умова не спрацювала.', 'default', '', '', '{document.type}'],
  ['document.scopeText', 'Уточнення документа', 'Непорожнє уточнення в дужках, крім вибору «Без уточнення».', 'except', 'document.scope', settings.get('basisChain.scope.none'), '({document.scope})'],
  ['reference.scopeText', 'Уточнення наказу-посилання', 'Непорожнє уточнення в дужках, крім вибору «Без уточнення».', 'except', 'reference.scope', settings.get('basisChain.scope.none'), '({reference.scope})'],
  ['court.proceedingText', 'Номер провадження з поясненням', 'Порожній номер не додає слово «провадження» або знак №.', 'nonempty', 'court.proceedingNumber', '', 'провадження № {court.proceedingNumber}'],
];
reference('Фрагменти тексту', 'tbBasisFragments', ['Код', 'Назва', 'Опис', 'Умова', 'Поле', 'Значення', 'Шаблон'], fragments, [31, 36, 75, 16, 28, 42, 70]);
reference('Контракт', 'tbBasisContract', ['Версія', 'Можливості', 'Опис'],
  [[settings.get('basisChain.contract.version'), settings.get('basisChain.contract.features'),
    'Коди полів і фрагментів — публічний контракт. Внутрішні імена TDB зберігаються лише в його конфігурації. Типи: text — текст, date — дата, choice — вибір. Використання: input — введення, computed — підготовлене значення. Умови фрагментів: nonempty — поле заповнено; equals — дорівнює значенню; except — заповнено й не дорівнює значенню; default — запасний варіант. У правилах: & — І, / — АБО, різні оператори групуються дужками.']],
  [12, 48, 120]);

reference('Переліки', 'tbEnumInstitutionGroups', ['Група установи', 'Назва'], [
  ['Unit', 'Військова частина'], ['Hospital', 'Лікувальний заклад'], ['TCC', 'ТЦК та СП'],
  ['ChiefCommand', 'Збройні Сили України'], ['GeneralStaff', 'Генеральний штаб'],
  ['LandCommand', 'Командування Сухопутних військ'], ['Court', 'Суд'],
], [28, 64]);

await fs.mkdir(outputDir, { recursive: true });
book.recalculate();
for (const { sheetName } of tables) {
  const png = await book.render({ sheetName, range: sheetName === 'Типи документів' ? 'B5:D9' : 'B2:D10', scale: 1, format: 'png' });
  await fs.writeFile(path.join(outputDir, `${sheetName}.png`), new Uint8Array(await png.arrayBuffer()));
}
const file = await SpreadsheetFile.exportXlsx(book);
await file.save(path.join(outputDir, 'BasisReferences.xlsx'));
console.log(JSON.stringify(tables));