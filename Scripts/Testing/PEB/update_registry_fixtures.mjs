import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
import { applyFixtureStyle } from "./fixture_styles.mjs";

process.on("uncaughtException", error => { console.error(error.stack); process.exit(1); });
const directory = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(directory, "../../..");
const python = process.env.PEB_PYTHON;
if (!python) throw new Error("Set PEB_PYTHON to the bundled Python executable.");
const bridge = path.join(directory, "xlsx_expectations_bridge.py");
const config = path.join(root, "PROJECTS_EXCEL/config/PersonalEventBuilder/wsConfig.test.txt");
const helper = path.join(directory, "event_routes.py");
const read = args => JSON.parse(execFileSync(python, ["-X", "utf8", ...args], { encoding: "utf8", maxBuffer: 16000000 }));
const settings = read([helper, config]);
const fixture = path.resolve(root, "PROJECTS_EXCEL/workbook/PersonalEventBuilder", settings["test.scenarios.path"]);
const temporary = path.join(os.tmpdir(), "peb-registry-fixtures-20261017");
await fs.mkdir(temporary, { recursive: true });
const source = read([bridge, "dump", fixture]);
const orders = read([helper, "--orders", config]);
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(fixture));
if (process.argv.includes("--inspect")) {
  const preview = await workbook.render({ sheetName: settings["test.scenarios.expectedRows.sheet"], range: "A43:Y47", scale: 1 });
  await fs.writeFile(path.join(temporary, "caption-full.png"), new Uint8Array(await preview.arrayBuffer()));
  process.exit(0);
}
const tables = {};
for (const [sheetName, sheet] of Object.entries(source)) {
  for (const table of sheet.tables) {
    const [start, end] = table.ref.split(":");
    const row = Number(start.match(/\d+/)[0]);
    const bottom = Number(end.match(/\d+/)[0]);
    const headers = sheet.rows[row - 1];
    tables[table.name] = { ...table, sheetName, row, bottom, headers };
  }
}
const summaries = tables[settings["test.scenarios.expectedMovement.table"]];
const expected = tables[settings["test.scenarios.expectedRows.table"]];
if (expected.row - summaries.bottom - 1 < Number(settings["test.expected.minimumGapRows"])) throw new Error("Expected tables need more space before editing; no output was written.");
let corrected = 0;
for (const name of [settings["test.scenarios.history.table"], settings["test.scenarios.expectedRows.table"]]) {
  const table = tables[name];
  const sheet = workbook.worksheets.getItem(table.sheetName);
  for (let row = table.row; row < table.bottom; row++) {
    const values = source[table.sheetName].rows[row];
    const replacements = new Map();
    for (const [dateField, numberField] of [["Вибуття", "Вибуття.Наказ"], ["Прибуття", "Прибуття.Наказ"]]) {
      const old = values[table.headers.indexOf(numberField)];
      if (old === null || old === undefined || old === "") continue;
      if (!["100", "101", "110", "115"].includes(String(old))) continue;
      const date = values[table.headers.indexOf(dateField)];
      if (orders[String(date)]?.length !== 1) throw new Error(`Missing or ambiguous source order for ${table.sheetName}/${row + 1}/${dateField}: ${date}`);
      const actual = orders[String(date)][0];
      // Only retire synthetic historical numbers; current-order backdated arrivals stay intact.
      if (["100", "101", "110", "115"].includes(String(old)) && String(old) !== actual) {
        replacements.set(String(old), actual);
        sheet.getCell(row, table.headers.indexOf(numberField)).values = [[actual]];
        corrected++;
      }
    }
    for (let column = 0; column < table.headers.length; column++) {
      let value = values[column];
      if (typeof value !== "string" || value.startsWith("=")) continue;
      for (const [old, number] of replacements) value = value.replaceAll(`2026/${old}/`, `2026/${number}/`);
      if (value !== values[column]) sheet.getCell(row, column).values = [[value]];
    }
  }
}
const registry = tables[settings["test.registry.table"]];
const kinds = settings["test.registry.kinds"].split("|");
const count = Math.max(...kinds.map(kind => settings[`test.registry.${kind}.values`].split("|").length));
if (count !== registry.bottom - registry.row) throw new Error("Registry size changed; resize it explicitly before this maintenance pass.");
const ids = settings["test.registry.ModeId.values"].split("|");
workbook.worksheets.getItem(registry.sheetName).getRangeByIndexes(registry.row, 0, count, registry.headers.length).values = Array.from({ length: count }, (_, index) => registry.headers.map(header => header === settings["test.registry.runtimeHeader"] ? ids[index] ?? null : settings[`test.registry.${header}.values`].split("|")[index] ?? null));
workbook.worksheets.getItem(registry.sheetName).getCell(registry.bottom + 1, 0).values = [[settings["test.registry.sourceCaption"]]];
workbook.worksheets.getItem(expected.sheetName).getCell(expected.row - 2, 0).values = [[settings["test.expected.rowsCaption"]]];
const captionRange = `A${expected.row - 1}:Y${expected.row - 1}`;
const caption = workbook.worksheets.getItem(expected.sheetName).getRange(captionRange);
caption.merge();
applyFixtureStyle(caption, "expectedRowsCaption");
workbook.recalculate();
const baseline = path.join(temporary, "iteration-original.xlsx");
const generated = path.join(temporary, "generated.xlsx");
await fs.copyFile(fixture, baseline);
await (await SpreadsheetFile.exportXlsx(workbook)).save(generated);
execFileSync(python, ["-X", "utf8", bridge, "assemble", baseline, generated, fixture, "РУХ_Історія", expected.sheetName, registry.sheetName, `--layout-range=${expected.sheetName}!${captionRange}`], { stdio: "inherit" });
execFileSync(python, ["-X", "utf8", bridge, "finalize-registry", fixture, config], { stdio: "inherit" });
const final = await SpreadsheetFile.importXlsx(await FileBlob.load(fixture));
for (const [sheetName, range] of [["РУХ_Історія", "A6:R11"], [expected.sheetName, `A${expected.row - 2}:M${expected.row + 3}`], [registry.sheetName, "A1:J5"]]) {
  const preview = await final.render({ sheetName, range, scale: 1 });
  await fs.writeFile(path.join(temporary, `${sheetName}.png`), new Uint8Array(await preview.arrayBuffer()));
}
console.log(JSON.stringify({ fixture, correctedOrders: corrected, scenarios: tables.tbTestScenarios.bottom - tables.tbTestScenarios.row, temporary }));