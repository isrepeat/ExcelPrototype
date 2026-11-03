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
  [
    "Командир",
    "командира"
  ],
  [
    "Заступник командира",
    "заступника командира"
  ],
  [
    "Начальник",
    "начальника"
  ],
  [
    "Заступник начальника",
    "заступника начальника"
  ],
  [
    "Начальник штабу",
    "начальника штабу"
  ],
  [
    "Начальник штабу — заступник командира",
    "начальника штабу — заступника командира"
  ],
  [
    "Начальник відділення персоналу штабу",
    "начальника відділення персоналу штабу"
  ],
  [
    "Начальник адміністративної групи штабу",
    "начальника адміністративної групи штабу"
  ],
  [
    "Головнокомандувач",
    "Головнокомандувача"
  ],
  [
    "Командувач",
    "командувача"
  ]
];
reference('Посадові особи', 'tbBasisOfficials', ['Назва', settings.get('basisChain.references.Officials.columns.Text')], officials, [48, 55]);

reference('Типи документів', 'tbBasisDocumentTypes', ['Назва', 'Родовий'], [
  [
    "повідомлення",
    "повідомлення"
  ],
  [
    "лист",
    "листа"
  ],
  [
    "наказ",
    "наказу"
  ],
  [
    "розпорядження",
    "розпорядження"
  ],
  [
    "окреме доручення",
    "окремого доручення"
  ],
  [
    "рапорт",
    "рапорту"
  ],
  [
    "доповідь",
    "доповіді"
  ],
  [
    "посвідчення про відрядження",
    "посвідчення про відрядження"
  ],
  [
    "ухвала",
    "ухвали"
  ],
  [
    "довідка",
    "довідки"
  ],
  [
    "повідомлення про прийняття кадрових рішень",
    "повідомлення про прийняття кадрових рішень"
  ]
], [48, 55]);
reference('Органи управління', 'tbCommands', ['Група установи', 'Назва', 'Родовий'], [
  ['ChiefCommand', 'Збройні Сили України', 'Збройних Сил України'],
  ['GeneralStaff', 'Генеральний штаб Збройних Сил України', 'Генерального штабу Збройних Сил України'],
  ['LandCommand', 'Сухопутні війська Збройних Сил України', 'Сухопутних військ Збройних Сил України'],
], [24, 58, 64]);
reference('Суди', 'tbCourts', ['Група установи', 'Назва', 'Родовий'], [
  ['Court', 'Новобаварський районний суд міста Харкова', 'Новобаварського районного суду міста Харкова'],
  ['Court', 'Київський районний суд міста Харкова', 'Київського районного суду міста Харкова'],
], [24, 58, 64]);

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