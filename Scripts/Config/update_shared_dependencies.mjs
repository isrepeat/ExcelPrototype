import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { execFileSync } from 'node:child_process';

const require = createRequire(import.meta.url);
const { FileBlob, SpreadsheetFile } = await import(pathToFileURL(require.resolve('@oai/artifact-tool')));
const root = fileURLToPath(new URL('../../PROJECTS_EXCEL/', import.meta.url));
const target = path.join(root, 'workbook/cfgDependencies.xlsx');
const output = path.join(root, 'workbook/.backup/shared-dependencies');
await fs.mkdir(output, { recursive: true });
const templateOption = process.argv.indexOf('--template');
const template = templateOption < 0 ? target : process.argv[templateOption + 1];
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(template));
const sheet = workbook.worksheets.getItem('wsConfig');

if (process.argv.includes('--preview')) {
  const preview = await workbook.render({ sheetName: 'wsConfig', range: 'B2:D13', scale: 1.5 });
  await fs.writeFile(path.join(output, 'before.png'), new Uint8Array(await preview.arrayBuffer()));
  console.log('Preview:', path.join(output, 'before.png'));
} else {
  const source = await fs.readFile(path.join(root, 'config/cfgDependencies/wsConfig.txt'), 'utf8');
  const rows = source.split(/\r?\n/).filter(Boolean).map(line => line.split('\t'));
  const keys = new Set();
  for (const row of rows) {
    if (row.length !== 3 || row.some(value => !value.trim()) || keys.has(row[1])) {
      throw new Error(`Invalid shared configuration row: ${row.join(' | ')}`);
    }
    keys.add(row[1]);
  }
  const table = sheet.tables.items.find(item => item.name === 'tbConfig');
  if (!table) throw new Error('Shared tbConfig table not found');
  const existing = table.rowCount;
  if (rows.length < existing) throw new Error('Refusing to remove shared configuration rows');
  if (rows.length > existing) {
    sheet.getRange(`B${existing + 2}:D${existing + 2}`).copyTo(sheet.getRange(`B${existing + 3}:D${rows.length + 2}`), 'all');
  }
  table.resize(`B2:D${rows.length + 2}`);
  sheet.getRange(`B3:D${rows.length + 2}`).values = rows;
  workbook.recalculate();
  const preview = await workbook.render({ sheetName: 'wsConfig', range: `B${existing + 3}:D${rows.length + 2}`, scale: 1.5 });
  await fs.writeFile(path.join(output, 'after.png'), new Uint8Array(await preview.arrayBuffer()));
  const backup = path.join(output, `cfgDependencies-${Date.now()}.xlsx`);
  await fs.copyFile(target, backup);
  const exported = await SpreadsheetFile.exportXlsx(workbook);
  await exported.save(target);
  const python = process.env.CODEX_PYTHON;
  if (!python) throw new Error('CODEX_PYTHON must point to the bundled Python runtime');
  execFileSync(python, [fileURLToPath(new URL('./preserve_shared_styles.py', import.meta.url)), template === target ? backup : template, target]);
  await fs.rm(`${target}.inspect.ndjson`, { force: true });
  console.log(`Updated ${rows.length} shared configuration rows. Backup: ${backup}`);
  console.log('Preview:', path.join(output, 'after.png'));
}