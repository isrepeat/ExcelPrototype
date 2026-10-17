import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
import { fixtureStyle, applyFixtureStyle } from "./fixture_styles.mjs";

process.on("uncaughtException", error => { console.error(error.stack); process.exit(1); });
const directory = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(directory, "../../..");
const input = path.join(root, "PROJECTS_EXCEL/tests/PersonalEventBuilder/TestScenarios.xlsx");
const output = path.join(root, "PROJECTS_EXCEL/tests/PersonalEventBuilder/TestScenarios.EventRoutes.xlsx");
const temporary = path.join(os.tmpdir(), "peb-event-routes-20261016");
await fs.mkdir(temporary, { recursive: true });
const python = process.env.PEB_PYTHON;
if (!python) throw new Error("Set PEB_PYTHON to the bundled Python executable.");
try {
  await fs.access(output);
  throw new Error("The routed scenario workbook already exists; preserve its run archive before creating another migration output.");
} catch (error) {
  if (error.code !== "ENOENT") throw error;
}
const bridge = path.join(directory, "xlsx_expectations_bridge.py");
const configPath = path.join(root, "PROJECTS_EXCEL/config/PersonalEventBuilder/wsConfig.test.txt");
const settings = Object.fromEntries((await fs.readFile(configPath, "utf8")).split(/\r?\n/).filter(Boolean).map(line => { const row = line.split("\t"); return [row[1].split("::").at(-1), row[2]]; }));
const source = JSON.parse(execFileSync(python, ["-X", "utf8", bridge, "dump", input], { encoding: "utf8", maxBuffer: 16000000 }));
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(input));
const letters = text => [...text].reduce((result, char) => result * 26 + char.charCodeAt(0) - 64, 0);
const table = name => {
  const sheetName = Object.keys(source).find(sheet => source[sheet].tables.some(item => item.name === name));
  const info = source[sheetName].tables.find(item => item.name === name);
  const [start, end] = info.ref.split(":");
  const top = Number(start.match(/\d+/)[0]);
  const left = letters(start.match(/[A-Z]+/)[0]);
  const right = letters(end.match(/[A-Z]+/)[0]);
  const headers = source[sheetName].rows[top - 1].slice(left - 1, right);
  const rows = source[sheetName].rows.slice(top, Number(end.match(/\d+/)[0])).map(row => Object.fromEntries(headers.map((header, index) => [header, row[left - 1 + index] ?? null])));
  return { name, sheetName, info, top, left, headers, rows };
};
const cases = table("tbTestScenarios");
const steps = table("tbTestSteps");
const history = table("tbTestHistory");
const absence = table("tbTestAbsence");
const movement = table("tbExpectedMovement");
const rows = table("tbExpectedMovementRows");
const word = table("tbExpectedWord");
const registry = table("tbTestRegistry");
if (cases.rows.some(row => row.ScenarioId === "S17")) throw new Error("S17 already exists; review migration identities before proceeding.");
const clone = value => structuredClone(value);
const stepOf = id => steps.rows.find(row => row.StepId === id);
const today = stepOf("W01-P01").EventDate;
const report = step => `рапорт лейтенанта БОНДАРЕНКА О.Ю. (вх.\u00a0№\u00a0${step.IncomingNumber} від\u00a015.10.2026).`;
const snapshot = (seed, step, recordId) => Object.fromEntries(rows.headers.map(header => [header, header === "StepId" ? step.StepId : header === "RecordId" ? recordId : seed[header] ?? null]));
const close = (seed, step, recordId) => ({ ...snapshot(seed, step, recordId), "Прибуття": today, "Прибуття.Наказ": "298", "Прибуття.Підстава": report(step), "Прибуття.Продовольче": today + 1 });
const departure = (step, recordId) => ({ ...clone(rows.rows.find(row => row.StepId === "S03-M01")), StepId: step.StepId, RecordId: recordId, "Вибуття.Підстава": report(step), "Супровідний документ": null });
for (const mode of ["FromOutpatientExaminationOrVlk", "ToOutpatientExaminationOrVlk"]) {
  const step = steps.rows.find(row => row.ModeId === mode);
  Object.assign(word.rows.find(row => row.StepId === step.StepId), { Outcome: "Omitted", ExpectedText: null, ExpectedError: null });
}
const s15 = stepOf("S15-M01");
s15.Action = "ExportEvent";
Object.assign(cases.rows.find(row => row.ScenarioId === "S15"), { "Операція": "Custom", "Назва": "Щорічна відпустка → лікарня: РУХ і Word", "Початковий стан": "Відкрита щорічна відпустка; закрити її та створити лікування" });
movement.rows.push({ StepId: s15.StepId, ExpectedDecision: "Allow", ExpectedCode: "StatusChanged", StateCheck: "Snapshot", ExpectedOpenCount: 1 });
rows.rows.push(close(history.rows.find(row => row.ScenarioId === "S15"), s15, "R01"));
rows.rows.push({ ...clone(rows.rows.find(row => row.StepId === "S01-M01")), StepId: s15.StepId, RecordId: "R02", "Вибуття.Підстава": report(s15) });
const vlkSeed = { ...clone(history.rows.find(row => row.ScenarioId === "S06")), "Подія": "Амбулаторне ВЛК", "Куди": stepOf("W01-P01").HospitalShortName, "Вибуття.Термін": null, "Вибуття.Дорога": null, "Прибуття.План": null, "Супровідний документ": null };
const enums = table("tbMovementEnums").rows;
const boundary = enums.find(row => row["Подія"] === "Амбулаторне ВЛК");
if (!boundary) throw new Error("Missing native movement enum: Амбулаторне ВЛК");
vlkSeed["За межі"] = boundary["За межі"];
for (const [id, count, decision, code] of [["S17", 1, "Allow", "StatusChanged"], ["S18", 0, "Allow", "DepartureCreated"], ["S19", 0, "Block", "OpenDepartureNotFound"], ["S20", 2, "Block", "AmbiguousOpenDeparture"]]) {
  const step = { ...clone(stepOf(id === "S18" ? "W01-P01" : "W05-P01")), ScenarioId: id, StepId: `${id}-E01`, Action: "ExportEvent", ModeId: id === "S18" ? "ToOutpatientExaminationOrVlk" : "MoveFromOutpatientVlkToVacationForCuring", IncomingNumber: `1656/${1400 + Number(id.slice(1))}-в (ТЕСТ-${id})`, TicketNumber: null, TicketDate: null };
  steps.rows.push(step);
  const label = id === "S18" ? "Амбулаторна ВЛК: РУХ без пункту Word" : `Амбулаторна ВЛК → відпустка: ${count} відкритих записів`;
  cases.rows.push({ ...clone(cases.rows.find(row => row.ScenarioId === "S15")), ScenarioId: id, "Назва": label, "Подія / сценарій": label, "Початковий стан": `${count} відкритих записів амбулаторної ВЛК` });
  absence.rows.push({ ...clone(absence.rows.find(row => row.ScenarioId === "W01")), ScenarioId: id });
  if (id !== "S18") for (let index = 0; index < count; index++) history.rows.push({ ...clone(vlkSeed), ScenarioId: id, RowId: index + 1, "Вибуття": vlkSeed["Вибуття"] + index, "Вибуття.Наказ": String(100 + index) });
  movement.rows.push({ StepId: step.StepId, ExpectedDecision: decision, ExpectedCode: code, StateCheck: decision === "Allow" ? "Snapshot" : "Unchanged", ExpectedOpenCount: id === "S18" || id === "S17" ? 1 : count });
  if (id === "S17") rows.rows.push(close(vlkSeed, step, "R01"), departure(step, "R02"));
  if (id === "S18") rows.rows.push({ ...clone(rows.rows.find(row => row.StepId === "S01-M01")), StepId: step.StepId, "Подія": "Амбулаторне ВЛК", "Вибуття.Підстава": report(step) });
  const baseline = word.rows.find(row => row.StepId === "W05-P01").ExpectedText;
  const old = stepOf("W05-P01");
  const expected = baseline.replace(old.IncomingNumber, step.IncomingNumber).replace(`; відпускний квиток №\u00a0${old.TicketNumber} від\u00a015.10.2026`, "");
  word.rows.push({ StepId: step.StepId, Outcome: id === "S18" ? "Omitted" : "Success", ExpectedText: id === "S18" ? null : expected, ExpectedError: null });
}
const kinds = settings["test.registry.kinds"].split("|");
registry.rows = Array.from({ length: Math.max(...kinds.map(kind => settings[`test.registry.${kind}.values`].split("|").length)) }, (_, index) => Object.fromEntries(registry.headers.map(header => {
  const id = settings["test.registry.ModeId.values"].split("|")[index];
  return [header, header === settings["test.registry.runtimeHeader"] ? id ? settings[`test.registry.ModeId.${id}.runtimeMode`] : null : settings[`test.registry.${header}.values`].split("|")[index] ?? null];
})));
const changed = [cases, steps, history, absence, movement, rows, word, registry];
for (const data of changed) {
  const sheet = workbook.worksheets.getItem(data.sheetName);
  sheet.tables.items.find(item => item.name === data.name).delete();
  const matrix = [data.headers, ...data.rows.map(row => data.headers.map(header => row[header] ?? null))];
  const range = sheet.getRangeByIndexes(data.top - 1, data.left - 1, matrix.length, data.headers.length);
  range.values = matrix.map(row => row.map(value => typeof value === "string" && value.startsWith("=") ? null : value));
  for (let row = 1; row < matrix.length; row++) for (let col = 0; col < data.headers.length; col++) if (typeof matrix[row][col] === "string" && matrix[row][col].startsWith("=")) sheet.getCell(data.top - 1 + row, data.left - 1 + col).formulas = [[matrix[row][col]]];
  sheet.tables.add(range.address, true, data.name).style = fixtureStyle("workbook").tableStyle;
  if (data === word) applyFixtureStyle(sheet.getRangeByIndexes(36, 0, matrix.length - 36, data.headers.length), "wordRows");
}
const generated = path.join(temporary, "generated.xlsx");
await (await SpreadsheetFile.exportXlsx(workbook)).save(generated);
execFileSync(python, ["-X", "utf8", bridge, "assemble", input, generated, output, ...new Set(changed.map(data => data.sheetName))], { stdio: "inherit" });
execFileSync(python, ["-X", "utf8", bridge, "finalize-registry", output, configPath], { stdio: "inherit" });
const finalBook = await SpreadsheetFile.importXlsx(await FileBlob.load(output));
for (const [sheetName, range] of [["Тест_Сценарії", "A50:I55"], ["Expected_WORD", "A36:D40"]]) {
  const preview = await finalBook.render({ sheetName, range, scale: 1 });
  await fs.writeFile(path.join(temporary, `${sheetName}.png`), new Uint8Array(await preview.arrayBuffer()));
}
console.log(JSON.stringify({ output, scenarios: cases.rows.length, steps: steps.rows.length }));