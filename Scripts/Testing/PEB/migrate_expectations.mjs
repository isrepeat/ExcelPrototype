import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

process.on("uncaughtException", error => { console.error(error.stack); process.exit(1); });

const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(scriptDirectory, "../../..");
const outputFixture = path.join(root, "PROJECTS_EXCEL/tests/PersonalEventBuilder/TestScenarios.xlsx");
const fixture = process.env.PEB_EXPECTATIONS_SOURCE ?? outputFixture;
const python = process.env.PEB_PYTHON;
if (!python) throw new Error("Set PEB_PYTHON to the bundled Python executable.");
const bridge = path.join(scriptDirectory, "xlsx_expectations_bridge.py");
const workDirectory = process.env.PEB_EXPECTATIONS_WORKDIR ?? path.join(os.tmpdir(), "peb-expectations-20261016");
await fs.mkdir(workDirectory, { recursive: true });
const source = JSON.parse(execFileSync(python, ["-X", "utf8", bridge, "dump", fixture], { encoding: "utf8", maxBuffer: 16000000 }));
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(fixture));

if (process.argv.includes("--inspect")) {
  console.log((await workbook.inspect({ kind: "sheet,table", maxChars: 2500, tableMaxRows: 2, tableMaxCols: 4 })).ndjson);
  const image = await workbook.render({ sheetName: "Тест_Кроки", range: "A1:N5", scale: 1 });
  await fs.writeFile(path.join(workDirectory, "before.png"), new Uint8Array(await image.arrayBuffer()));
  process.exit(0);
}
if (source["Expected_РУХ"] || source["Expected_WORD"]) throw new Error("The workbook is already migrated; expectations must be edited manually.");

const columnName = index => {
  let name = "";
  for (let number = index + 1; number > 0; number = Math.floor((number - 1) / 26)) name = String.fromCharCode(65 + (number - 1) % 26) + name;
  return name;
};
const records = name => {
  const table = source[name].tables[0];
  const lastRow = Number(table.ref.split(":")[1].match(/\d+$/)[0]);
  const headers = source[name].rows[0];
  return source[name].rows.slice(1, lastRow).filter(row => row.some(value => value !== null)).map(row => Object.fromEntries(headers.map((header, index) => [header, row[index]])));
};
const matrix = (headers, rows) => [headers, ...rows.map(row => headers.map(header => row[header] ?? null))];
const oldSteps = records("Тест_Кроки");
const cases = records("Тест_Сценарії");
const history = records("Тест_Історія");
const absence = records("Тест_ЄЖОС");
const nativeHeaders = source["Тест_Історія"].rows[0].slice(3).filter(header => !["Валідація", "Статус потенціально незакритих записів"].includes(header));
const inputHeaders = ["ScenarioId", "StepId", "StepOrder", "Action", ...source["Тест_Кроки"].rows[0].filter(header => !["ScenarioId", "StepId", "ExpectedDecision", "ExpectedCode", "ExpectedOpenCount", "PersonPosition"].includes(header))];
const movement = [];
const snapshots = [];
const steps = [];
const stateByScenario = new Map();
for (const step of oldSteps) {
  const order = step.StepId;
  const id = `${step.ScenarioId}-M${String(order).padStart(2, "0")}`;
  steps.push({ ...step, StepId: id, StepOrder: order, Action: "ExportMovement" });
  movement.push({ StepId: id, ExpectedDecision: step.ExpectedDecision, ExpectedCode: step.ExpectedCode,
    StateCheck: step.ExpectedDecision === "Pending" ? "Pending" : step.ExpectedDecision === "Allow" ? "Snapshot" : "Unchanged",
    ExpectedOpenCount: step.ExpectedOpenCount });
  if (!stateByScenario.has(step.ScenarioId)) stateByScenario.set(step.ScenarioId,
    history.filter(row => row.ScenarioId === step.ScenarioId).map(row => ({ ...row, RecordId: `R${String(row.RowId).padStart(2, "0")}` })));
  const state = stateByScenario.get(step.ScenarioId);
  if (step.ExpectedDecision !== "Allow") continue;
  // Описываем эталонный результат независимо от конфигурации экспортёра.
  if (step.ModeId.startsWith("Departure")) {
    state.push({ RecordId: `R${String(state.length + 1).padStart(2, "0")}`, "ІПН": step["ІПН"], "ПІБ": step["ПІБ"], "Подія": step["Подія"],
      "Вибуття": step.EventDate, "Вибуття.Наказ": step.AcceptedOrderNumber, "Куди": step.Destination, "Вибуття.Підстава": step.IncomingNumber });
  } else {
    const open = state.filter(row => row["ІПН"] === step["ІПН"] && row["Подія"] === step["Подія"] && row["Прибуття"] == null);
    if (open.length !== 1) throw new Error(`Expected one open fixture row: ${id}`);
    Object.assign(open[0], { "Прибуття": step.EventDate, "Прибуття.Наказ": step.AcceptedOrderNumber, "Прибуття.Підстава": step.IncomingNumber });
  }
  snapshots.push(...state.map(row => ({ ...row, StepId: id })));
}

const word = [];
// Текстовые эталоны заданы явно и не вычисляются рабочим DSL-рендерером.
const person = "Лейтенанта БОНДАРЕНКА Олександра Юрійовича, ipn_0001, офіцера інформаційно-аналітичної групи штабу";
const dateHeading = "З 15\u00a0жовтня 2026\u00a0року:";
for (let index = 0; index < 9; index++) {
  const baseline = oldSteps[index < 8 ? index : 0];
  const scenarioId = `W${String(index + 1).padStart(2, "0")}`;
  const stepId = `${scenarioId}-P01`;
  const incoming = `ТЕСТ-${scenarioId}`;
  const input = { ...baseline, ScenarioId: scenarioId, StepId: stepId, StepOrder: 1, Action: "PreviewWord", IncomingNumber: incoming,
    TicketNumber: incoming, HospitalShortName: index === 8 ? "ТЕСТ-ЛЗ-01" : "А0209", Hospital: "Військова частина А0209" };
  steps.push(input);
  cases.push({ ...cases.find(row => row.ScenarioId === baseline.ScenarioId), ScenarioId: scenarioId,
    "Назва": index === 8 ? "WORD: невідомий лікувальний заклад" : `WORD: ${baseline.ModeId}`, RuleStatus: "Defined",
    "Початковий стан": "Порожня ТВ; прев’ю не змінює РУХ", "Операція": "PreviewWord", "Подія / сценарій": baseline["Подія"] });
  const originalAbsence = absence.find(row => row.ScenarioId === baseline.ScenarioId);
  absence.push({ ...originalAbsence, ScenarioId: scenarioId, HasEvent: 0 });
  const basis = `Підстава: рапорт лейтенанта БОНДАРЕНКА О.Ю. (вх.\u00a0№\u00a0${incoming} від\u00a015.10.2026)`;
  const ticket = `; відпускний квиток №\u00a0${incoming} від\u00a015.10.2026`;
  let text;
  if (baseline.ModeId === "ToHospital") text = `У військову частину А0209, з\u00a015\u00a0жовтня 2026\u00a0року:\n\n${person[0].toLowerCase()}${person.slice(1)}.\n\n${basis}.`;
  else if (baseline.ModeId === "FromHospital") text = `${dateHeading}\n\nЗ військової частини А0209:\n\n${person[0].toLowerCase()}${person.slice(1)}.\n\n${basis}.`;
  else if (baseline.ModeId.startsWith("Departure")) text = `${person}, у м. Тестоград, з\u00a015\u00a0жовтня 2026\u00a0року.\n\n${basis}${ticket}.`;
  else text = `${dateHeading}\n\n${person}.\n\n${basis}${ticket}.`;
  word.push({ StepId: stepId, Outcome: index === 8 ? "Error" : "Success", ExpectedText: index === 8 ? null : text,
    ExpectedError: index === 8 ? "Export lookup must return exactly one row: export.Word.ToHospital.lookups.1" : null });
}

const changed = [];
function replaceTable(name, tableName, headers, rows, startRow = 1, clear = true) {
  const sheet = source[name] || changed.includes(name) ? workbook.worksheets.getItem(name) : workbook.worksheets.add(name);
  const originalTable = sheet.tables.items.find(table => table.name === tableName);
  if (originalTable) originalTable.delete();
  if (clear) sheet.getUsedRange().clear({ applyTo: "all" });
  const values = matrix(headers, rows);
  const range = `A${startRow}:${columnName(headers.length - 1)}${startRow + values.length - 1}`;
  sheet.getRange(range).values = values;
  const table = sheet.tables.add(range, true, tableName);
  table.style = "TableStyleDark1";
  const grid = sheet.getRange(range);
  grid.format.font = { name: "Times New Roman", size: 12, color: "#FFFFFF" };
  grid.format.fill = "#303030";
  grid.format.wrapText = true;
  grid.format.verticalAlignment = "center";
  grid.format.rowHeight = 38;
  grid.format.columnWidth = 21;
  sheet.getRange(`A${startRow}:${columnName(headers.length - 1)}${startRow}`).format.font.bold = true;
  sheet.getRange(`A${startRow}:${columnName(headers.length - 1)}${startRow}`).format.fill = "#151515";
  sheet.getRange(`A${startRow}:${columnName(headers.length - 1)}${startRow}`).format.rowHeight = 60;
  sheet.getRange(`A${startRow}:${columnName(headers.length - 1)}${startRow}`).format.horizontalAlignment = "center";
  if (!source[name]) sheet.showGridLines = false;
  for (let column = 0; column < headers.length; column++) {
    const header = headers[column];
    const cells = sheet.getRange(`${columnName(column)}${startRow + 1}:${columnName(column)}${startRow + values.length - 1}`);
    if (["EventDate", "AcceptedOrderDate", "IncomingDate", "TicketDate", "Дата тесту", "Вибуття", "Прибуття", "Прибуття.План"].includes(header)) {
      cells.setNumberFormat("dd\\.mm\\.yyyy");
      cells.format.horizontalAlignment = "center";
    }
    if (["StepOrder", "AcceptedOrderNumber", "ExpectedOpenCount"].includes(header)) cells.format.horizontalAlignment = "center";
    if (["ПІБ", "Назва", "Подія / сценарій", "Початковий стан", "Посада"].includes(header)) cells.format.columnWidth = 32;
  }
  if (!changed.includes(name)) changed.push(name);
  return sheet;
}

replaceTable("Тест_Кроки", "tbTestSteps", inputHeaders, steps);
function appendFixtures(name, rows) {
  const sheet = workbook.worksheets.getItem(name);
  const original = source[name];
  const headers = original.rows[0];
  const table = sheet.tables.items[0];
  const originalRows = records(name);
  const additions = rows.slice(originalRows.length);
  table.rows.add(null, matrix(headers, additions).slice(1));
  for (let index = 0; index < additions.length; index++) {
    const baseline = oldSteps[index < 8 ? index : 0].ScenarioId;
    const sourceRow = originalRows.findIndex(row => row.ScenarioId === baseline) + 2;
    const targetRow = originalRows.length + index + 2;
    sheet.getRange(`A${targetRow}:${columnName(headers.length - 1)}${targetRow}`).copyFrom(sheet.getRange(`A${sourceRow}:${columnName(headers.length - 1)}${sourceRow}`), "all");
    sheet.getRange(`A${targetRow}:${columnName(headers.length - 1)}${targetRow}`).values = matrix(headers, [additions[index]]).slice(1);
  }
  changed.push(name);
}
appendFixtures("Тест_Сценарії", cases);
appendFixtures("Тест_ЄЖОС", absence);
const movementSheet = replaceTable("Expected_РУХ", "tbExpectedMovement", ["StepId", "ExpectedDecision", "ExpectedCode", "StateCheck", "ExpectedOpenCount"], movement);
replaceTable("Expected_РУХ", "tbExpectedMovementRows", ["StepId", "RecordId", ...nativeHeaders], snapshots, movement.length + 5, false);
movementSheet.getRange(`A${movement.length + 3}`).values = [["Очікувані рядки після кроку"]];
movementSheet.getRange(`A${movement.length + 3}:F${movement.length + 3}`).format.font = { name: "Times New Roman", size: 12, bold: true, color: "#202020" };
movementSheet.getRange("C2:C19").format.columnWidth = 28;
const wordSheet = replaceTable("Expected_WORD", "tbExpectedWord", ["StepId", "Outcome", "ExpectedText", "ExpectedError"], word);
wordSheet.getRange("C1:C10").format.columnWidth = 100;
wordSheet.getRange("D1:D10").format.columnWidth = 60;
wordSheet.getRange("A2:D9").format.rowHeight = 175;
wordSheet.getRange("A10:D10").format.rowHeight = 75;
wordSheet.getRange("A2:D10").format.verticalAlignment = "top";
wordSheet.getRange("B2:B10").dataValidation = { rule: { type: "list", values: ["Success", "Error", "Pending"] } };
movementSheet.getRange("D2:D19").dataValidation = { rule: { type: "list", values: ["Snapshot", "Unchanged", "Pending"] } };

// Расширяем схему журнала, сохраняя прежние запуски и ячейки за границами таблицы.
const resultSheet = workbook.worksheets.getItem("Тест_Результати");
const resultTable = source["Тест_Результати"].tables[0];
const resultHeaders = [...source["Тест_Результати"].rows[0], "StepOrder", "Action", "StateCheck", "MovementResult", "WordResult", "ExpectedPreviewText", "PreviewText", "ExpectedPreviewError", "PreviewDiagnostic"];
resultSheet.tables.items.find(table => table.name === resultTable.name).delete();
resultSheet.getRange(`A1:${columnName(resultHeaders.length - 1)}1`).values = [resultHeaders];
const report = resultSheet.tables.add(`A1:${columnName(resultHeaders.length - 1)}${resultTable.ref.match(/\d+$/)[0]}`, true, resultTable.name);
report.style = "TableStyleDark1";
resultSheet.getRange("W1:AE1").format = { fill: "#151515", font: { name: "Times New Roman", size: 12, bold: true, color: "#FFFFFF" }, wrapText: true };
resultSheet.getRange("W1:AE126").format.columnWidth = 24;
changed.push("Тест_Результати");

for (const name of ["Тест_Результати", "Тест_Записи", "Тест_ВхідніДані"]) {
  const rows = source[name].rows;
  const headers = rows[0];
  const scenarioColumn = headers.indexOf("ScenarioId");
  const stepColumn = headers.indexOf("StepId");
  const sheet = workbook.worksheets.getItem(name);
  for (let index = 1; index < rows.length; index++) {
    const row = rows[index];
    if (!row[scenarioColumn] || !Number.isInteger(row[stepColumn])) continue;
    sheet.getCell(index, stepColumn).values = [[`${row[scenarioColumn]}-M${String(row[stepColumn]).padStart(2, "0")}`]];
    if (name === "Тест_Результати") {
      sheet.getRange(`W${index + 1}:AA${index + 1}`).values = [[row[stepColumn], "ExportMovement", null, "Legacy", "NotRun"]];
    }
  }
  if (!changed.includes(name)) changed.push(name);
}

workbook.recalculate();
console.log((await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!", options: { useRegex: true, maxResults: 10 }, maxChars: 1500 })).ndjson);
for (const [name, range, output] of [["Тест_Кроки", "A1:J5", "steps.png"], ["Expected_РУХ", "A1:E8", "movement.png"],
  ["Expected_РУХ", "A23:H28", "rows.png"], ["Expected_WORD", "A1:C3", "word.png"], ["Тест_Сценарії", "A17:J21", "cases.png"],
  ["Тест_ЄЖОС", "A17:E21", "absence.png"], ["Тест_Результати", "W1:AA4", "results.png"]]) {
  const image = await workbook.render({ sheetName: name, range, scale: 1 });
  await fs.writeFile(path.join(workDirectory, output), new Uint8Array(await image.arrayBuffer()));
}
const generated = path.join(workDirectory, "artifact.xlsx");
await (await SpreadsheetFile.exportXlsx(workbook)).save(generated);
const original = path.join(workDirectory, "original.xlsx");
if (path.resolve(fixture) !== path.resolve(original)) await fs.copyFile(fixture, original);
execFileSync(python, ["-X", "utf8", bridge, "assemble", original, generated, outputFixture, ...changed, "--allow-layout-migration"], { stdio: "inherit" });
execFileSync(process.execPath, [path.join(scriptDirectory, "normalize_inputs.mjs"), "--allow-layout-migration"], { stdio: "inherit" });
console.log(`Migrated ${cases.length} scenarios, ${steps.length} steps, ${movement.length} Movement outcomes, ${snapshots.length} expected rows and ${word.length} Word baselines.`);