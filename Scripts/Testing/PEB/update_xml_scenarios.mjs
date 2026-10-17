import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
import { applyFixtureStyle, fixtureStyle } from "./fixture_styles.mjs";

process.on("uncaughtException", error => { console.error(error.stack); process.exit(1); });

const directory = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(directory, "../../..");
const configDirectory = path.join(root, "PROJECTS_EXCEL/config/PersonalEventBuilder");
const xml = await fs.readFile(path.join(root, "PROJECTS_EXCEL/workbook/PersonalEventBuilder/PersonalEventBuilderWordResultTemplates.xml"), "utf8");
const modes = [...xml.matchAll(/<template\s+name="([^"]+)"/g)].map(match => match[1]);
const existing = new Set(["FromHospital", "FromVacationForCuring", "FromAnnualMainVacation", "FromVacationForFamily", "ToHospital", "ToAnnualMainVacation", "ToVacationForFamily", "ToVacationForCuring"]);
const baseMode = mode => {
  if (mode.startsWith("From")) return mode === "FromBusinessTrip" ? "FromAnnualMainVacation" : "FromHospital";
  if (mode.endsWith("AnnualVacation")) return "ToAnnualMainVacation";
  if (mode.endsWith("FamilyVacation") || mode === "ToMaternityLeave") return "ToVacationForFamily";
  if (mode.endsWith("VacationForCuring")) return "ToVacationForCuring";
  return "ToHospital";
};
const labels = {
  FromBusinessTrip: "Прибуття з відрядження",
  FromOutpatientExaminationOrVlk: "Прибуття з амбулаторного огляду / ВЛК",
  FromExternalVlk: "Прибуття зі стаціонарної ВЛК",
  FromMedicalCompanyForTreatment: "Прибуття з лікування в медичній роті",
  ToMaternityLeave: "Вибуття у відпустку у зв’язку з вагітністю та пологами",
  ToOutpatientExaminationOrVlk: "Вибуття на амбулаторний огляд / ВЛК",
  ToMedicalCompanyForTreatment: "Вибуття на лікування в медичну роту"
};
const sourceNames = {
  Hospital: "на лікуванні у комунальному некомерційному підприємстві \"Міська клінічна багатопрофільна лікарня № 25\" Харківської міської ради",
  VacationForCuring: "у відпустці для лікування",
  AnnualVacation: "у щорічній основній відпустці",
  FamilyVacation: "у відпустці за сімейними обставинами",
  Vlk: "на медичному огляді військово-лікарської комісії",
  MedicalCompanyForTreatment: "на лікуванні у медичній роті",
  MedicalCompanyVacationForCuring: "у відпустці для лікування в медичній роті",
  ExternalVlk: "на стаціонарному медичному огляді військово-лікарської комісії"
};
const transition = mode => mode.match(/^MoveFrom(.+)To(.+)$/)?.slice(1);
const eventNames = {
  Hospital: "лікування в лікарні", VacationForCuring: "відпустка для лікування",
  AnnualVacation: "щорічна відпустка", FamilyVacation: "відпустка за сімейними обставинами",
  Vlk: "ВЛК", ExternalVlk: "стаціонарна ВЛК", MedicalCompanyForTreatment: "лікування в медичній роті",
  MedicalCompanyVacationForCuring: "відпустка для лікування в медичній роті"
};
const eventLabel = mode => {
  const pair = transition(mode);
  return labels[mode] ?? (pair ? pair.map(name => eventNames[name]).join(" → ") : mode);
};
const configRows = text => text.split(/\r?\n/).filter(Boolean).map(line => line.split("\t"));
const settings = Object.fromEntries(configRows(await fs.readFile(path.join(configDirectory, "wsConfig.test.txt"), "utf8")).map(row => [row[1].split("::").at(-1), row[2]]));

if (process.argv.includes("--label-patch")) {
  let patch = "*** Begin Patch\n";
  for (const filename of ["wsConfig.txt", "wsConfig.test.txt"]) {
    const rows = configRows(await fs.readFile(path.join(configDirectory, filename), "utf8"));
    const changes = rows.flatMap(row => {
      const mode = modes.find(mode => transition(mode) && [`PersonalEventBuilder::modes.${mode}.caption`, `PersonalEventBuilder::export.events.${mode}`].includes(row[1]));
      return mode && row[2] !== eventLabel(mode) ? [`@@\n-${row.join("\t")}\n+${[row[0], row[1], eventLabel(mode)].join("\t")}\n`] : [];
    });
    if (changes.length) patch += `*** Update File: ${path.join(configDirectory, filename).replaceAll("\\", "/")}\n${changes.join("")}`;
  }
  console.log(patch + "*** End Patch");
  process.exit(0);
}

// Emit a reviewable patch; application configuration is never written directly.
if (process.argv.includes("--config-patch")) {
  let patch = "*** Begin Patch\n";
  for (const filename of ["wsConfig.txt", "wsConfig.test.txt"]) {
    const relative = `PROJECTS_EXCEL/config/PersonalEventBuilder/${filename}`;
    const rows = configRows(await fs.readFile(path.join(configDirectory, filename), "utf8"));
    const key = "PersonalEventBuilder::test.registry.ModeId.values";
    const registryRow = rows.find(row => row[1] === key);
    const old = registryRow?.join("\t");
    const additions = [];
    const add = (key, value) => additions.push(`value\tPersonalEventBuilder::${key}\t${value}`);
    for (const mode of modes.filter(mode => !existing.has(mode))) {
      const base = baseMode(mode);
      const [from, to] = transition(mode) ?? [];
      const label = eventLabel(mode);
      if (registryRow) add(`test.registry.ModeId.${mode}.runtimeMode`, mode);
      add(`modes.${mode}.caption`, label);
      add(`modes.${mode}.direction`, from ? "StatusChange" : mode.startsWith("From") ? "Arrival" : "Departure");
      add(`modes.${mode}.form`, settings[`modes.${base}.form`]);
      add(`export.events.${mode}`, label);
      for (const row of rows.filter(row => row[1].startsWith(`PersonalEventBuilder::export.Word.${base}.`))) {
        let value = row[2];
        if (row[1].endsWith(".templateName")) value = mode;
        if (row[1].endsWith(".previewRequired")) {
          if (mode === "FromBusinessTrip") value = value.split("|").filter(field => !["TicketNumber", "TicketDate"].includes(field)).join("|");
          if (mode.includes("MedicalCompany") || ["FromOutpatientExaminationOrVlk", "FromExternalVlk"].includes(mode)) value = value.split("|").filter(field => field !== "HospitalShortName").join("|");
        }
        if (row[1].endsWith(".lookups.count") && (mode.includes("MedicalCompany") || mode === "FromBusinessTrip" || ["FromOutpatientExaminationOrVlk", "FromExternalVlk"].includes(mode))) value = "0";
        if (row[1].includes(".lookups.1.") && (mode.includes("MedicalCompany") || ["FromOutpatientExaminationOrVlk", "FromExternalVlk"].includes(mode))) continue;
        add(row[1].split("::").at(-1).replace(`export.Word.${base}.`, `export.Word.${mode}.`), value);
      }
    }
    add("text.StatusChangeType", "Зміна статусу");
    const last = rows.at(-1).join("\t");
    patch += `*** Update File: ${relative}\n${old ? `@@\n-${old}\n+value\t${key}\t${modes.join("|")}\n` : ""}@@\n ${last}\n${additions.map(line => "+" + line).join("\n")}\n`;
  }
  console.log(patch + "*** End Patch");
  process.exit(0);
}

const python = process.env.PEB_PYTHON;
if (!python) throw new Error("Set PEB_PYTHON to the bundled Python executable.");
const bridge = path.join(directory, "xlsx_expectations_bridge.py");
const fixture = path.join(root, "PROJECTS_EXCEL/tests/PersonalEventBuilder/TestScenarios.xlsx");
const output = path.join(os.tmpdir(), "peb-xml-scenarios-20261016");
await fs.mkdir(output, { recursive: true });
const source = JSON.parse(execFileSync(python, ["-X", "utf8", bridge, "dump", fixture], { encoding: "utf8", maxBuffer: 16000000 }));
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(fixture));
if (process.argv.includes("--inspect")) {
  for (const [sheetName, range] of [["Тест_Сценарії", "A16:I18"], ["Тест_Кроки", "A17:N19"], ["Тест_Історія", "A11:N13"], ["Тест_ЄЖОС", "A25:H28"], ["Expected_РУХ", "A33:Y36"], ["Expected_WORD", "A11:D12"], ["Test_Registry", "A28:J35"]]) {
    const preview = await workbook.render({ sheetName, range, scale: 1 });
    await fs.writeFile(path.join(output, `final-${sheetName}.png`), new Uint8Array(await preview.arrayBuffer()));
  }
  console.log(`Rendered final workbook in ${output}`);
  process.exit(0);
}
const locate = name => {
  const sheetName = Object.keys(source).find(nameOfSheet => source[nameOfSheet].tables.some(table => table.name === name));
  const table = source[sheetName].tables.find(table => table.name === name);
  const [start, end] = table.ref.split(":");
  const row = Number(start.match(/\d+/)[0]);
  const column = [...start.match(/[A-Z]+/)[0]].reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0);
  const right = [...end.match(/[A-Z]+/)[0]].reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0);
  const headers = source[sheetName].rows[row - 1].slice(column - 1, right);
  return { sheetName, table, row, column, headers, rows: source[sheetName].rows.slice(row, Number(end.match(/\d+/)[0])).map(values => Object.fromEntries(headers.map((header, index) => [header, values[column - 1 + index] ?? null]))) };
};
const cases = locate("tbTestScenarios");
const steps = locate("tbTestSteps");
const history = locate("tbTestHistory");
const absence = locate("tbTestAbsence");
const movement = locate("tbExpectedMovement");
const snapshots = locate("tbExpectedMovementRows");
const word = locate("tbExpectedWord");
const registry = locate("tbTestRegistry");
if (cases.rows.some(row => row.ScenarioId === "W10")) throw new Error("Migration already applied; edit the resulting expectations manually.");
const currentDate = steps.rows.find(row => row.ScenarioId === "W01").EventDate;
const hospital = steps.rows.find(row => row.ScenarioId === "W01");
const person = "Лейтенанта БОНДАРЕНКА Олександра Юрійовича, ipn_0001, офіцера інформаційно-аналітичної групи штабу";
const personLower = person[0].toLowerCase() + person.slice(1);
const date = "15\u00a0жовтня 2026\u00a0року";
const hospitalAcc = 'комунальне некомерційне підприємство "Міська клінічна багатопрофільна лікарня № 25" Харківської міської ради';
const basis = step => `Підстава: рапорт лейтенанта БОНДАРЕНКА О.Ю. (вх.\u00a0№\u00a0${step.IncomingNumber} від\u00a015.10.2026)`;
const ticket = step => step.TicketNumber ? `; відпускний квиток №\u00a0${step.TicketNumber} від\u00a015.10.2026` : "";
// Golden texts are explicit expectations, independent of the production DSL evaluator.
const expectedText = step => {
  const mode = step.ModeId;
  const [from, to] = transition(mode) ?? [];
  if (from) {
    let sourcePhrase = sourceNames[from];
    if (from === "MedicalCompanyForTreatment" && to === "VacationForCuring") sourcePhrase = "на лікуванні в медичній роті";
    if (from === "MedicalCompanyVacationForCuring" && to === "VacationForCuring") sourcePhrase = "у відпустці для лікування у медичній роті";
    const targetPhrase = to === "Hospital" ? `госпіталізований у ${hospitalAcc}, з\u00a0${date}.`
      : ["Vlk", "ExternalVlk"].includes(to) ? `вибув на медичний огляд військово-лікарської комісії${to === "Vlk" ? ` у ${hospitalAcc}` : ""}, з\u00a0${date}.`
      : `вибув у ${to === "FamilyVacation" ? "відпустку за сімейними обставинами" : to === "AnnualVacation" ? "щорічну основну відпустку" : "відпустку для лікування"} у м. Тестоград, з\u00a0${date}.`;
    return `NN.\u00a0${person}, який перебуває ${sourcePhrase}, вважати таким, що ${targetPhrase}\n\n${basis(step)}${ticket(step)}.`;
  }
  const bodies = {
    FromBusinessTrip: `З ${date}:\n\nЗ м. Тестоград:\n\n${personLower}.`,
    FromOutpatientExaminationOrVlk: `${person}.`,
    FromExternalVlk: `${person}.`,
    FromMedicalCompanyForTreatment: `З ${date}:\n\n${personLower}.`,
    ToMaternityLeave: `${person}, у м. Тестоград, з\u00a0${date}.`,
    ToOutpatientExaminationOrVlk: `${person}, у ${hospitalAcc}.`,
    ToMedicalCompanyForTreatment: `З\u00a0${date}:\n\n${personLower}.`
  };
  if (!bodies[mode]) throw new Error(`Missing golden expectation: ${mode}`);
  return `${bodies[mode]}\n\n${basis(step)}${ticket(step)}.`;
};
const clone = row => structuredClone(row);
const addWord = (scenarioId, mode, existingStep) => {
  const base = baseMode(mode);
  const reference = steps.rows.find(row => row.Action === "PreviewWord" && row.ModeId === base);
  const step = existingStep ?? { ...clone(reference), ScenarioId: scenarioId, StepId: `${scenarioId}-P01`, StepOrder: 1 };
  Object.assign(step, { Action: "PreviewWord", ModeId: mode, EventDate: currentDate, IncomingDate: currentDate, TicketNumber: null, TicketDate: null, Notes: transition(mode) ? "Перевірка шаблону без зміни продовольчого забезпечення" : null });
  if (!existingStep) step.IncomingNumber = `1656/${1300 + Number(scenarioId.slice(1))}-в (ТЕСТ-${scenarioId})`;
  if (base === "ToAnnualMainVacation") { step.TicketNumber = `2026/298/7 (ТЕСТ-${scenarioId})`; step.TicketDate = currentDate; }
  if (mode === "FromBusinessTrip" || mode.includes("MedicalCompany") || mode.endsWith("VacationForCuring")) step.Destination = "м. Тестоград";
  if (mode.includes("MedicalCompany") && !mode.endsWith("Hospital")) { step.Hospital = null; step.HospitalShortName = null; }
  if (transition(mode) || mode === "ToOutpatientExaminationOrVlk") { step.Hospital = hospital.Hospital; step.HospitalShortName = hospital.HospitalShortName; }
  const scenario = existingStep ? cases.rows.find(row => row.ScenarioId === scenarioId) : { ...clone(cases.rows.find(row => row.ScenarioId === "W01")), ScenarioId: scenarioId };
  Object.assign(scenario, { "Назва": `WORD: ${mode}`, "Подія / сценарій": eventLabel(mode), "Операція": "PreviewWord", "RuleStatus": "Defined", "Початковий стан": existingStep ? "Відкрита щорічна відпустка; лише перевірка Word" : "Порожня ТВ; прев’ю не змінює РУХ" });
  if (!existingStep) { cases.rows.push(scenario); steps.rows.push(step); absence.rows.push({ ...clone(absence.rows.find(row => row.ScenarioId === "W01")), ScenarioId: scenarioId }); }
  word.rows.push({ StepId: step.StepId, Outcome: "Success", ExpectedText: expectedText(step), ExpectedError: null });
};
addWord("S15", "MoveFromAnnualVacationToHospital", steps.rows.find(row => row.StepId === "S15-M01"));
movement.rows = movement.rows.filter(row => row.StepId !== "S15-M01" && row.StepId !== "S16-M01");
word.rows.find(row => row.StepId === "W09-P01").ExpectedError = word.rows.find(row => row.StepId === "W09-P01").ExpectedError.replace("DepartureTreatment", "ToHospital");
let number = 10;
for (const mode of modes.filter(mode => !existing.has(mode) && mode !== "MoveFromAnnualVacationToHospital")) addWord(`W${String(number++).padStart(2, "0")}`, mode);
for (const scenario of cases.rows.filter(row => row["Операція"] === "PreviewWord")) {
  const step = steps.rows.find(row => row.ScenarioId === scenario.ScenarioId);
  scenario["Назва"] = `WORD: ${step.ModeId}${scenario.ScenarioId === "W09" ? " — невідома лікарня" : ""}`;
}
steps.rows = steps.rows.filter(row => row.StepId !== "S16-M01");
const arrival = steps.rows.find(row => row.StepId === "S16-M02");
const departure = steps.rows.find(row => row.StepId === "S16-M03");
Object.assign(arrival, { StepOrder: 1, EventDate: currentDate - 1, IncomingDate: currentDate, TicketNumber: null, TicketDate: null });
Object.assign(departure, { StepOrder: 2, EventDate: currentDate, IncomingDate: currentDate });
Object.assign(cases.rows.find(row => row.ScenarioId === "S16"), { "Назва": "Прибуття заднім числом і вибуття одним наказом", "Подія / сценарій": "Прибуття з відпустки → вибуття на лікування", "Початковий стан": "Відкрита щорічна відпустка з 01.10; один наказ від 15.10", "Операція": "Custom" });
history.rows.push({ ...clone(history.rows.find(row => row.ScenarioId === "S06")), ScenarioId: "S16" });
Object.assign(absence.rows.find(row => row.ScenarioId === "S16"), { ...clone(absence.rows.find(row => row.ScenarioId === "S06")), ScenarioId: "S16" });
snapshots.rows = snapshots.rows.filter(row => !row.StepId.startsWith("S16-"));
const movementBasis = step => basis(step).replace("Підстава: ", "") + ".";
const closed = { ...clone(snapshots.rows.find(row => row.StepId === "S06-M01")), StepId: arrival.StepId, "Прибуття": currentDate - 1, "Прибуття.Наказ": "298", "Прибуття.Підстава": movementBasis(arrival) };
snapshots.rows.push(closed, { ...clone(closed), StepId: departure.StepId });
const newHospital = { ...clone(snapshots.rows.find(row => row.StepId === "S01-M01")), StepId: departure.StepId, RecordId: "R02", "Вибуття.Наказ": "298", "Вибуття.Підстава": movementBasis(departure) };
snapshots.rows.push(newHospital);
for (const expectation of movement.rows.filter(row => row.StepId.startsWith("S16-"))) expectation.ExpectedOpenCount = expectation.StepId === arrival.StepId ? 0 : 1;
registry.rows = Array.from({ length: Math.max(...settings["test.registry.kinds"].split("|").map(kind => settings[`test.registry.${kind}.values`].split("|").length)) }, (_, index) => Object.fromEntries(registry.headers.map(header => [header, header === settings["test.registry.runtimeHeader"] ? modes[index] ?? null : settings[`test.registry.${header}.values`].split("|")[index] ?? null])));
const changed = [cases, steps, history, absence, movement, snapshots, word, registry];
for (const data of changed) {
  const sheet = workbook.worksheets.getItem(data.sheetName);
  sheet.tables.items.find(table => table.name === data.table.name).delete();
  const matrix = [data.headers, ...data.rows.map(row => data.headers.map(header => row[header] ?? null))];
  const range = sheet.getRangeByIndexes(data.row - 1, data.column - 1, matrix.length, data.headers.length);
  range.values = matrix.map(row => row.map(value => typeof value === "string" && value.startsWith("=") ? null : value));
  for (let row = 1; row < matrix.length; row++) for (let column = 0; column < data.headers.length; column++) if (typeof matrix[row][column] === "string" && matrix[row][column].startsWith("=")) sheet.getCell(data.row - 1 + row, data.column - 1 + column).formulas = [[matrix[row][column]]];
  sheet.tables.add(range.address, true, data.table.name).style = fixtureStyle("workbook").tableStyle;
  if (data === word) applyFixtureStyle(sheet.getRangeByIndexes(10, 0, matrix.length - 10, data.headers.length), "wordRows");
  if (data === registry) sheet.getCell(matrix.length + 1, 0).values = [[settings["test.registry.sourceCaption"]]];
}
const baseline = path.join(output, "original.xlsx");
const generated = path.join(output, "generated.xlsx");
await fs.copyFile(fixture, baseline);
await (await SpreadsheetFile.exportXlsx(workbook)).save(generated);
execFileSync(python, ["-X", "utf8", bridge, "assemble", baseline, generated, fixture, ...new Set(changed.map(data => data.sheetName))], { stdio: "inherit" });
execFileSync(python, ["-X", "utf8", bridge, "finalize-registry", fixture, path.join(configDirectory, "wsConfig.test.txt")], { stdio: "inherit" });
for (const [sheetName, range] of [["Тест_Сценарії", "A14:I22"], ["Expected_WORD", "A10:C14"]]) {
  const preview = await workbook.render({ sheetName, range, scale: 1 });
  await fs.writeFile(path.join(output, `${sheetName}.png`), new Uint8Array(await preview.arrayBuffer()));
}
console.log(JSON.stringify({ scenarios: cases.rows.length, steps: steps.rows.length, wordExpectations: word.rows.length, templates: modes.length, output }));