import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
import { eventCellValue } from "./event_cell_values.mjs";

process.on("uncaughtException", error => { console.error(error.stack); process.exit(1); });
const directory = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(directory, "../../..");
const fixture = process.env.PEB_INPUTS_SOURCE ?? path.join(root, "PROJECTS_EXCEL/tests/PersonalEventBuilder/TestScenarios.xlsx");
const python = process.env.PEB_PYTHON;
if (!python) throw new Error("Set PEB_PYTHON to the bundled Python executable.");
const bridge = path.join(directory, "xlsx_expectations_bridge.py");
const output = process.env.PEB_INPUTS_WORKDIR ?? path.join(os.tmpdir(), "peb-universal-inputs-20261016");
await fs.mkdir(output, { recursive: true });
const source = JSON.parse(execFileSync(python, ["-X", "utf8", bridge, "dump", fixture], { encoding: "utf8", maxBuffer: 16000000 }));
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(fixture));
if (process.argv.includes("--inspect")) {
  const preview = await workbook.render({ sheetName: "Тест_Кроки", range: "A1:N4", scale: 1 });
  await fs.writeFile(path.join(output, "before.png"), new Uint8Array(await preview.arrayBuffer()));
  process.exit(0);
}

const settings = Object.fromEntries((await fs.readFile(path.join(root, "PROJECTS_EXCEL/config/PersonalEventBuilder/wsConfig.test.txt"), "utf8"))
  .split(/\r?\n/).filter(line => line.trim()).map(line => { const fields = line.split("\t"); return [fields[1].split("::").at(-1), fields[2]]; }));
if (process.argv.includes("--movement-enums-only")) {
  const name = settings["test.scenarios.expectedRows.sheet"];
  const table = source[name].tables.find(item => item.name === settings["test.scenarios.expectedRows.table"]);
  if (!source[name].tables.some(item => item.name === "tbMovementEnums")) throw new Error("Missing tbMovementEnums in scenario workbook.");
  const firstRow = Number(table.ref.split(":")[0].match(/\d+/)[0]);
  const lastRow = Number(table.ref.split(":")[1].match(/\d+/)[0]);
  const headers = source[name].rows[firstRow - 1];
  const sheet = workbook.worksheets.getItem(name);
  const formula = '=IF([@Подія]="","",IF(AND(OR([@Подія]="Стаціонарне лікування",[@Подія]="Відпустка для лікування"),[@Куди]="Медична рота А7383"),"Ні",XLOOKUP([@Подія],tbMovementEnums[Подія],tbMovementEnums[За межі])))';
  sheet.getRangeByIndexes(firstRow, headers.indexOf("За межі"), lastRow - firstRow, 1).formulas = Array.from({ length: lastRow - firstRow }, () => [formula]);
  sheet.getRangeByIndexes(firstRow, headers.indexOf("Подія"), lastRow - firstRow, 1).dataValidation = { rule: { type: "list", formula1: "_PEB_MovementEvents" } };
  const baseline = path.join(output, "enums-original.xlsx");
  const generated = path.join(output, "enums-generated.xlsx");
  await fs.copyFile(fixture, baseline);
  await (await SpreadsheetFile.exportXlsx(workbook)).save(generated);
  execFileSync(python, ["-X", "utf8", bridge, "assemble", baseline, generated, fixture, name], { stdio: "inherit" });
  execFileSync(python, ["-X", "utf8", bridge, "finalize-movement-enums", fixture, settings["test.scenarios.expectedRows.table"]], { stdio: "inherit" });
  process.exit(0);
}
if (process.argv.includes("--registry-only")) {
  const kinds = settings["test.registry.kinds"].split("|");
  const rows = Math.max(...kinds.map(kind => settings[`test.registry.${kind}.values`].split("|").length));
  const headers = [...kinds, settings["test.registry.runtimeHeader"]];
  const values = [headers, ...Array.from({ length: rows }, (_, index) => headers.map(kind => {
    if (kind === settings["test.registry.runtimeHeader"]) {
      const mode = settings["test.registry.ModeId.values"].split("|")[index];
      return mode ? settings[`test.registry.ModeId.${mode}.runtimeMode`] : null;
    }
    return settings[`test.registry.${kind}.values`].split("|")[index] ?? null;
  }))];
  const name = settings["test.registry.sheet"];
  const registry = source[name] ? workbook.worksheets.getItem(name) : workbook.worksheets.add(name);
  for (const table of [...registry.tables.items]) table.delete();
  if (source[name]) registry.getUsedRange().clear({ applyTo: "all" });
  const lastColumn = String.fromCharCode(64 + headers.length);
  const range = `A1:${lastColumn}${rows + 1}`;
  registry.getRange(range).values = values;
  registry.tables.add(range, true, settings["test.registry.table"]).style = "TableStyleDark1";
  registry.getRange(range).format = { fill: "#303030", font: { name: "Times New Roman", size: 12, color: "#FFFFFF" }, wrapText: true, verticalAlignment: "center" };
  registry.getRange(range).format.columnWidth = 28;
  registry.getRange(range).format.rowHeight = 38;
  registry.getRange(`A1:${lastColumn}1`).format = { fill: "#151515", font: { name: "Times New Roman", size: 12, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center" };
  registry.getRange(`A${rows + 3}`).values = [[settings["test.registry.sourceCaption"]]];
  registry.showGridLines = false;
  const modes = settings["test.registry.ModeId.values"].split("|");
  const reverse = new Map(modes.map(mode => [settings[`test.registry.ModeId.${mode}.runtimeMode`], mode]));
  const originalSteps = source["Тест_Кроки"].rows;
  const modeIndex = originalSteps[0].indexOf("ModeId");
  const stepSheet = workbook.worksheets.getItem("Тест_Кроки");
  for (let index = 1; index < originalSteps.length; index++) {
    const mode = originalSteps[index][modeIndex];
    if (mode && reverse.has(mode)) stepSheet.getCell(index, modeIndex).values = [[reverse.get(mode)]];
  }
  for (const binding of settings["test.registry.bindings"].split("|")) {
    const [tableName, header, kind] = binding.split(":");
    const sheetName = Object.keys(source).find(name => source[name].tables.some(table => table.name === tableName));
    const table = source[sheetName].tables.find(table => table.name === tableName);
    const row = Number(table.ref.split(":")[0].match(/\d+/)[0]);
    const index = source[sheetName].rows[row - 1].indexOf(header);
    const last = Number(table.ref.split(":")[1].match(/\d+/)[0]);
    workbook.worksheets.getItem(sheetName).getRangeByIndexes(row, index, last - row, 1).dataValidation = {
      rule: { type: "list", formula1: settings[`test.registry.${kind}.rangeName`] }
    };
  }
  workbook.recalculate();
  const preview = await workbook.render({ sheetName: name, range: "A1:E9", scale: 1 });
  await fs.writeFile(path.join(output, "registry.png"), new Uint8Array(await preview.arrayBuffer()));
  const baseline = path.join(output, "registry-original.xlsx");
  const generated = path.join(output, "registry-generated.xlsx");
  await fs.copyFile(fixture, baseline);
  await (await SpreadsheetFile.exportXlsx(workbook)).save(generated);
  execFileSync(python, ["-X", "utf8", bridge, "assemble", baseline, generated, fixture, name], { stdio: "inherit" });
  execFileSync(python, ["-X", "utf8", bridge, "finalize-registry", fixture, path.join(root, "PROJECTS_EXCEL/config/PersonalEventBuilder/wsConfig.test.txt")], { stdio: "inherit" });
  console.log(`Built ${kinds.length} registry groups from TEST configuration using XML template identifiers.`);
  process.exit(0);
}
const hospital = JSON.parse(execFileSync(python, ["-X", "utf8", "-c",
  'import openpyxl,json,sys; s=openpyxl.load_workbook(sys.argv[1],data_only=True)["Лікувальні Заклади"]; r=[list(row[:6]) for row in s.values if row[0]=="КНП МКБЛ №25 ХМР"]; assert len(r)==1 and all(r[0][i] for i in (1,3,4)); print(json.dumps(r[0],ensure_ascii=False))',
  path.join(root, "PROJECTS_EXCEL/workbook/Dependencies/Установи.xlsx")], { encoding: "utf8" }));
const lowerFirst = text => text[0].toLowerCase() + text.slice(1);
function tableRows(sheetName, tableName) {
  const table = source[sheetName].tables.find(item => item.name === tableName);
  const top = Number(table.ref.split(":")[0].match(/\d+/)[0]);
  const bottom = Number(table.ref.split(":")[1].match(/\d+/)[0]);
  const headers = source[sheetName].rows[top - 1];
  return source[sheetName].rows.slice(top, bottom).filter(row => row.some(value => value !== null))
    .map(row => Object.fromEntries(headers.map((header, index) => [header, row[index]])));
}
const headers = settings["test.scenarios.stepHeaders"].split("|");
const steps = tableRows("Тест_Кроки", "tbTestSteps").map(row => ({ ...row,
  PersonId: row.PersonId ?? row["ІПН"], PersonName: row.PersonName ?? row["ПІБ"], Notes: row.Notes ?? row["Примітки"] }));
const stepById = new Map(steps.map(step => [step.StepId, step]));
const runtimeMode = step => {
  const mapped = settings[`test.registry.ModeId.${step.ModeId}.runtimeMode`];
  if (mapped) return mapped;
  if (settings[`export.events.${step.ModeId}`]) return step.ModeId;
  throw new Error(`Unknown fixture mode: ${step.ModeId}`);
};
const outcomes = new Map(tableRows("Expected_WORD", "tbExpectedWord").map(row => [row.StepId, row]));
if (process.argv.includes("--document-fixtures-only")) {
  const shortDate = value => typeof value === "string" ? value : new Date(Date.UTC(1899, 11, 30) + value * 86400000).toISOString().slice(0, 10).split("-").reverse().join(".");
  const basis = (number, date) => `рапорт лейтенанта БОНДАРЕНКА О.Ю. (вх.\u00a0№\u00a0${number} від\u00a0${shortDate(date)}).`;
  const replacements = new Map();
  const scenarioTickets = new Map();
  for (const [index, step] of steps.entries()) {
    const oldIncoming = step.IncomingNumber;
    const oldTicket = step.TicketNumber;
    const marker = oldIncoming?.match(/ТЕСТ-[A-Za-z0-9-]+/)?.[0];
    if (!marker) throw new Error(`Missing fixture document marker: ${step.StepId}`);
    step.IncomingNumber = `1656/${1234 + index}-в (${marker})`;
    if (settings[`modes.${runtimeMode(step)}.form`] === "Leave") {
      const ticket = `${shortDate(step.AcceptedOrderDate).slice(-4)}/${step.AcceptedOrderNumber}/7 (${marker})`;
      if (runtimeMode(step).startsWith("To")) scenarioTickets.set(step.ScenarioId, ticket);
      step.TicketNumber = scenarioTickets.get(step.ScenarioId) ?? ticket;
      step.TicketDate = step.TicketDate ?? step.EventDate;
    }
    const replacement = { basis: basis(step.IncomingNumber, step.IncomingDate), ticket: step.TicketNumber ?? null };
    replacements.set(oldIncoming, replacement);
    replacements.set(basis(oldIncoming, step.IncomingDate), replacement);
    const word = outcomes.get(step.StepId);
    if (word?.Outcome === "Success") {
      word.ExpectedText = word.ExpectedText.replaceAll(`вх.\u00a0№\u00a0${oldIncoming}`, `вх.\u00a0№\u00a0${step.IncomingNumber}`);
      if (oldTicket && step.TicketNumber) word.ExpectedText = word.ExpectedText.replaceAll(`відпускний квиток №\u00a0${oldTicket}`, `відпускний квиток №\u00a0${step.TicketNumber}`);
    }
  }
  const history = tableRows("Тест_Історія", "tbTestHistory");
  const naturalKey = (scenario, row) => JSON.stringify([scenario, ...["ІПН", "Подія", "Вибуття", "Вибуття.Наказ"].map(header => row[header])]);
  const historical = new Map();
  for (const [index, row] of history.entries()) {
    const marker = `ТЕСТ-${row.ScenarioId}-H${String(index + 1).padStart(2, "0")}`;
    row["Вибуття.Підстава"] = basis(`1656/${1200 + index}-в (${marker})`, row["Вибуття"]);
    if (row["Прибуття"]) row["Прибуття.Підстава"] = basis(`1656/${1300 + index}-в (${marker}-A)`, row["Прибуття"]);
    if (["Щорічна відпустка", "Відпустка для лікування", "Відпустка за сімейними обставинами"].includes(row["Подія"])) {
      row["Супровідний документ"] = `${shortDate(row["Вибуття"]).slice(-4)}/${row["Вибуття.Наказ"]}/7 (${marker})`;
    }
    historical.set(naturalKey(row.ScenarioId, row), row);
  }
  const expected = tableRows("Expected_РУХ", "tbExpectedMovementRows");
  for (const row of expected) {
    const previous = historical.get(naturalKey(stepById.get(row.StepId).ScenarioId, row));
    const departure = replacements.get(row["Вибуття.Підстава"]);
    const arrival = replacements.get(row["Прибуття.Підстава"]);
    if (departure) {
      row["Вибуття.Підстава"] = departure.basis;
      row["Супровідний документ"] = departure.ticket;
    } else if (previous) {
      row["Вибуття.Підстава"] = previous["Вибуття.Підстава"];
      row["Супровідний документ"] = previous["Супровідний документ"];
    }
    if (arrival) row["Прибуття.Підстава"] = arrival.basis;
    else if (previous?.["Прибуття"]) row["Прибуття.Підстава"] = previous["Прибуття.Підстава"];
  }
  for (const [name, tableName, records] of [["Тест_Кроки", "tbTestSteps", steps], ["Тест_Історія", "tbTestHistory", history],
    ["Expected_РУХ", "tbExpectedMovementRows", expected], ["Expected_WORD", "tbExpectedWord", [...outcomes.values()]]]) {
    const table = source[name].tables.find(item => item.name === tableName);
    const top = Number(table.ref.split(":")[0].match(/\d+/)[0]);
    const columnCount = table.ref.split(":")[1].replace(/\d/g, "").split("").reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0);
    const columns = source[name].rows[top - 1].slice(0, columnCount);
    workbook.worksheets.getItem(name).getRange(table.ref).values = [columns, ...records.map(row => columns.map(header => row[header] ?? null))];
  }
  const baseline = path.join(output, "documents-original.xlsx");
  const generated = path.join(output, "documents-generated.xlsx");
  await fs.copyFile(fixture, baseline);
  await (await SpreadsheetFile.exportXlsx(workbook)).save(generated);
  execFileSync(python, ["-X", "utf8", bridge, "assemble", baseline, generated, fixture, "Тест_Кроки", "Тест_Історія", "Expected_РУХ", "Expected_WORD"], { stdio: "inherit" });
  console.log("Updated incoming documents, qualified tickets and literal Movement/Word expectations.");
  process.exit(0);
}
for (const step of steps) {
  for (const header of headers.slice(5, -1)) step[header] = eventCellValue(step[header]);
  if (["ToHospital", "FromHospital"].includes(runtimeMode(step))) {
    const negative = outcomes.get(step.StepId)?.Outcome === "Error";
    step.Hospital = negative ? null : hospital[1];
    step.HospitalShortName = negative ? "ТЕСТ-ЛЗ-01" : hospital[0];
    step.Destination = null;
    step.TicketNumber = null;
    step.TicketDate = null;
  } else {
    step.Hospital = null;
    step.HospitalShortName = null;
  }
}
for (const sheetName of ["Тест_Історія", "Expected_РУХ"]) {
  const rows = source[sheetName].rows;
  const headerIndex = rows.findIndex(row => row.includes("Куди") && row.includes("Подія"));
  const eventColumn = rows[headerIndex].indexOf("Подія");
  const destinationColumn = rows[headerIndex].indexOf("Куди");
  for (const row of rows.slice(headerIndex + 1)) {
    for (let index = 0; index < rows[headerIndex].length; index++) {
      if (!["ScenarioId", "RowId", "TargetTable", "StepId", "RecordId"].includes(rows[headerIndex][index])) row[index] = eventCellValue(row[index]);
    }
    if (row[eventColumn] === settings["export.events.ToHospital"]) row[destinationColumn] = hospital[0];
  }
}
for (const row of source.Expected_WORD.rows.slice(1)) {
  const step = stepById.get(row[0]);
  if (!step || row[1] !== "Success" || !["ToHospital", "FromHospital"].includes(runtimeMode(step))) continue;
  const paragraphs = row[2].replace(/\r\n?/g, "\n").split("\n\n");
  if (runtimeMode(step).startsWith("To")) {
    const date = paragraphs[0].match(/, з([\s\S]*)$/);
    if (!date) throw new Error(`Unexpected treatment text baseline: ${step.StepId}`);
    paragraphs[0] = `У ${lowerFirst(hospital[4])}, з${date[1]}`;
  } else {
    if (!paragraphs[1]?.startsWith("З ")) throw new Error(`Unexpected hospital group header: ${step.StepId}`);
    paragraphs[1] = `З ${lowerFirst(hospital[3])}:`;
  }
  row[2] = paragraphs.join("\n\n");
}
const columnName = index => {
  let result = "";
  for (let value = index + 1; value > 0; value = Math.floor((value - 1) / 26)) result = String.fromCharCode(65 + (value - 1) % 26) + result;
  return result;
};
const sheet = workbook.worksheets.getItem("Тест_Кроки");
sheet.tables.items.find(table => table.name === "tbTestSteps").delete();
sheet.getUsedRange().clear({ applyTo: "all" });
const values = [headers, ...steps.map(step => headers.map(header => step[header] ?? null))];
const range = `A1:${columnName(headers.length - 1)}${values.length}`;
sheet.getRange(range).values = values;
sheet.tables.add(range, true, "tbTestSteps").style = "TableStyleDark1";
sheet.getRange(range).format = { fill: "#303030", font: { name: "Times New Roman", size: 12, color: "#FFFFFF" }, wrapText: true, verticalAlignment: "center" };
sheet.getRange(range).format.rowHeight = 62;
sheet.getRange(range).format.columnWidth = 20;
sheet.getRange(`A1:${columnName(headers.length - 1)}1`).format = { fill: "#151515", font: { name: "Times New Roman", size: 12, bold: true, color: "#FFFFFF" }, wrapText: true, horizontalAlignment: "center" };
sheet.getRange("A1:E28").format.columnWidth = 18;
for (let index = 0; index < headers.length; index++) {
  const cells = sheet.getRange(`${columnName(index)}2:${columnName(index)}${values.length}`);
  if (headers[index].endsWith("Date")) {
    cells.setNumberFormat("dd\\.mm\\.yyyy");
    cells.format.horizontalAlignment = "center";
  }
  if (["Hospital", "PersonName", "ReportName", "Document", "Notes"].includes(headers[index])) cells.format.columnWidth = 35;
  if (["Hospital", "Notes"].includes(headers[index])) cells.format.columnWidth = 55;
}
for (const name of ["Тест_Історія", "Expected_РУХ", "Expected_WORD"]) {
  const rows = source[name].rows;
  workbook.worksheets.getItem(name).getRange(`A1:${columnName(rows[0].length - 1)}${rows.length}`).values = rows;
}
for (const name of ["Тест_Історія", "Expected_РУХ"]) {
  const rows = source[name].rows;
  const top = rows.findIndex(row => row.includes("Куди") && row.includes("Подія"));
  const column = columnName(rows[top].indexOf("Куди"));
  const destination = workbook.worksheets.getItem(name).getRange(`${column}${top + 1}:${column}${rows.length}`);
  destination.format.columnWidth = 55;
  destination.format.wrapText = true;
  workbook.worksheets.getItem(name).getRange(`A${top + 2}:${columnName(rows[top].length - 1)}${rows.length}`).format.rowHeight = 80;
}
// В Expected_РУХ верхняя таблица уже нижней, поэтому учитываем полный размер матрицы.
const expectedRows = source["Expected_РУХ"].rows;
const width = Math.max(...expectedRows.map(row => row.length));
workbook.worksheets.getItem("Expected_РУХ").getRange(`A1:${columnName(width - 1)}${expectedRows.length}`).values = expectedRows.map(row => Array.from({ length: width }, (_, index) => row[index] ?? null));
const wordSheet = workbook.worksheets.getItem("Expected_WORD");
wordSheet.getRange("A1:D10").format = { fill: "#303030", font: { name: "Times New Roman", size: 12, color: "#FFFFFF" }, wrapText: true, verticalAlignment: "top" };
wordSheet.getRange("A1:D1").format = { fill: "#151515", font: { name: "Times New Roman", size: 12, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center" };
wordSheet.getRange("A2:D3").format.rowHeight = 210;
workbook.recalculate();
for (const [name, range, filename] of [["Тест_Кроки", "F1:N4", "inputs.png"], ["Тест_Кроки", "O1:Z4", "documents.png"], ["Expected_WORD", "A1:C3", "word.png"]]) {
  const preview = await workbook.render({ sheetName: name, range, scale: 1 });
  await fs.writeFile(path.join(output, filename), new Uint8Array(await preview.arrayBuffer()));
}
const baseline = path.join(output, "original.xlsx");
const generated = path.join(output, "generated.xlsx");
await fs.copyFile(fixture, baseline);
await (await SpreadsheetFile.exportXlsx(workbook)).save(generated);
execFileSync(python, ["-X", "utf8", bridge, "assemble", baseline, generated, fixture, "Тест_Кроки", "Тест_Історія", "Expected_РУХ", "Expected_WORD",
  ...(process.argv.includes("--allow-layout-migration") ? ["--allow-layout-migration"] : [])], { stdio: "inherit" });
console.log(`Updated ${steps.length} steps with ${headers.length} universal columns and hospital ${hospital[0]}.`);
execFileSync(process.execPath, [fileURLToPath(import.meta.url), "--registry-only"], { stdio: "inherit" });