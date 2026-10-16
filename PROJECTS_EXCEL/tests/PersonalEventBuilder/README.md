# PEB dummy sources

## Catalog loading and page lifecycle

Progress includes completed export steps and completed stages of the current step: floor(100 * (completedSteps + completedStages / stageCount) / totalSteps). Each step has eight equally weighted stages: restore the scenario, capture the initial snapshot, prepare export inputs, calculate and capture the definitive pre-export snapshot, execute the export, capture the post-export snapshot, assert the result, and archive the report and working books. Each milestone is reported after its operation finishes. Opening the execution session contributes zero; multi-step scenarios contribute one unit per export step. These are completed-work milestones, not elapsed-time estimates. An archive failure retains the partial percentage instead of reporting completion.

Scenario classification is stored in the Операція and Подія / сценарій columns of Тест_Сценарії. S01-S08 show departure/arrival and the event type separately. S09-S16 are Custom and retain their descriptive names. The UI groups departure cases first, arrival cases second, and Custom cases last; both classification columns are filterable in the Smart table. Scenario identities, inputs and execution rules are unchanged.

Displaying or refreshing the catalog uses the existing ADO query executor against the saved scenario and step worksheets only. It does not create Excel.Application, open working Movement/Absence books or initialize the exporter. Full scenario validation and the private execution session start on the first scheduled runner callback after Run.

The TEST XAML owns TestCatalog through the standard table control with Smart representation and Data.TestCatalog binding. The normal table-plan pipeline handles creation, resizing and removal. The test UI supplies raw data and recreates checkbox overlays after the table update; it does not add or unlist native tables itself. Page refresh goes through the active page Render contract so these overlays and result cells are restored too. Refresh is blocked during an active run.

## Scenario-specific Absence source

Absence fixtures contain leave departures only, including medical leave, but never inpatient treatment. S01, S05 and S11 explicitly use HasEvent=0; S10 and S15 retain only their leave history. Treatment events remain in Movement history and export steps. The Absence fixture has 16 rows, one per scenario.

`Тест_ЄЖОС` / `tbTestAbsence` in `TestScenarios.xlsx` defines the initial Absence source for each scenario. It contains ScenarioId, ІПН, HasEvent and native ЕЖОС column headers. HasEvent=1 inserts an event; a single HasEvent=0 row explicitly defines an empty state. Keep column 66 equal to the scenario participant's full name, including on an empty-state row. Multiple event rows are supported.

Before the first step, the scenario session validates these rows, replaces only that participant's rows in the configured TEST ЕЖОС table and restores Movement history. ЕЖОС has no IPN column, so replacement uses the validated full name. Other participants remain untouched. Subsequent steps retain the same state. The session opens TestScenarios, Movement and Absence books in its private Excel instance and rejects paths outside the test directory, open/locked books and read-only access. Keep all three books closed during catalog loading and execution.

Manual scenario commit saves both working sources. Automated preparation keeps them unsaved until the step archive is saved in TestScenarios; disposal discards pending changes. Saving multiple workbooks is not atomic. The exporter still receives isolated step inputs: restoring ЕЖОС does not yet add end-to-end lookup/UI assertions.

## Separate scenario workbook

`TestScenarios.xlsx` contains four input sheets and native tables: `Тест_Сценарії` / `tbTestScenarios`, `Тест_Історія` / `tbTestHistory`, `Тест_Кроки` / `tbTestSteps`, and `Тест_ЄЖОС` / `tbTestAbsence`. They were moved from `Dependencies/test_РУХ.xlsx` with worksheet XML, table definitions and style catalog preserved byte for byte. Movement retains only its native working sheets.

These tables describe 16 alternative scenarios for the one person already present in the workbook. `tbTestHistory` includes ScenarioId, RowId, TargetTable and every existing column from ТВ. A scenario without history rows means that ТВ starts empty. Computed ТВ columns are intentionally blank in the fixture history; the loader preserves/reapplies the working table's formulas rather than overwriting them with blanks.

`tbTestSteps` contains 18 operations, input values, proposed expected decisions, reason codes and expected open-event counts. S16 is a three-step sequence without resets between steps. S15 has a deliberately unresolved exception expectation. Decisions are test specifications, not evidence of implemented or successfully executed business rules.

The TEST configuration routes Movement export to `Dependencies/test_РУХ.xlsx`, sheet `Відсутні`, native table `ТВ`. Personnel and Absence use `test_ШПО.xlsx` and `test_ЕЖОС.xlsx`; orders and hospitals use the ordinary Dependencies folder. Scenario definitions and diagnostics use `TestScenarios.xlsx`. Production routing remains unchanged.

### Selecting and running tests on the main PEB page

After updating VBA, XAML and configuration, switch to TEST and click `Тестування РУХ`. The selector and live results are rendered on `MainPage`, below the reserved candidate area in columns B:Q. PEB no longer creates a separate test worksheet or stores the full report in the UI workbook. Existing old `PEB_Tests` sheets, if any, are not deleted. The panel uses the existing column widths, preserves the form/candidate areas and displays a bounded window of recent results.

- Check scenarios, use `Вибрати всі` / `Зняти вибір`, then `Запустити обрані`. Catalog entries are read automatically from the existing scenario tables; no Run column or manual workbook editing is needed.
- Each scenario starts from its defined seed history. Steps within that scenario accumulate their changes in order, as in S16. The real Movement preparation/export services run against isolated form bindings and the validated TEST workbook. The main form is preserved.
- Scheduling uses workbook-qualified `Application.OnTime` callbacks: initialize the private session, then execute one test step per callback. The full batch is not a synchronous loop and uses no DoEvents loop. Excel regains control between steps; editing a cell naturally delays OnTime until Excel is ready. A single VBA/COM step still executes synchronously and is not preemptible; opening/saving Excel books can briefly occupy the UI.
- The Excel status bar and PEB action status display the processed-step percentage and counters. `Зупинити` is available in the panel and beside the main test button. Cancellation prevents new export operations; remaining selected steps are archived as SKIP on subsequent short callbacks. A failed assertion skips dependent steps of that scenario while independent scenarios continue; an infrastructure error aborts remaining operations.
- Each persisted step appears immediately in the panel. Pending rules, including S15, are skipped. Manual loading/export and environment switching are guarded during an active run. Reload and workbook close cancel the scheduled callback and release the private workbook/session.

### Durable diagnostics in TestScenarios

Scenario tables remain input definitions; ТВ and other working tables are populated only by the relevant operations. Three permanent archive sheets/native filterable tables are prepared in `TestScenarios.xlsx`. Their dark table style, column formats and widths are stored in the workbook. VBA validates their schema and appends values without creating sheets or changing formatting. An empty last row is reused as the first data row. Missing archive sheets or tables stop the operation with a specific diagnostic.

| Sheet / table | Contents |
| --- | --- |
| Тест_Результати / tbTestResults | RunId, scenario/name, step, mode, event, append/close operation, person, rule status, expected/actual decision/code/open count, PASS/FAIL/SKIP/ERROR, details, milliseconds, timestamp, planned/processed counts and run state |
| Тест_Записи / tbTestRecords | Complete native ТВ column values before and after each step, with scenario/run metadata, Phase, StateAvailable, record index and Created/Closed/Modified/Deleted/Unchanged/Empty markers |
| Тест_ВхідніДані / tbTestInputs | Original fixture fields and prepared export values, stored as field/value rows with input stage and VBA value type |

Filter by RunId, ScenarioId, StepId, ModeId, EventName, Operation, Result or ChangeKind. For an unexpected close, compare the Before and After rows for the same run/scenario/step. Empty initial tables get an explicit Empty marker; unavailable snapshots after preparation errors are marked Unavailable rather than represented as valid empty state. Raw native cell types/values are retained in record snapshots. Formula expressions/protected helper integrity are checked without comparing volatile formula caches.

The automatic borrowed-workbook export does not save independently. Source operation changes and all three diagnostics tables are committed together by the scenario session after each step. A journal save failure is shown as ERROR and ends the job; unsaved source changes are discarded when closing the session. Runtime tables are append-only, schema-checked, and capped by the configured archive row limit. No unrelated sheet data is overwritten. Completed-step reports survive PEB reload/close; a hard interruption can leave the last RunState as Running with CompletedSteps below TotalSteps, which identifies an incomplete run. Keep the test source closed while running; open it after completion/cancellation to inspect/filter the archive.

Assertions compare machine codes, open counts, the exact affected row count, configured export/join fields, unchanged unrelated rows and unmapped fields, and unchanged writable/protected/formula state on rejection. Seeding trims only native test table rows (minimum one blank formula-template row), preventing accumulation of blank rows.

Fixture reason codes `NoOpenDeparture`, `PreviousDepartureClosed` and `MatchingOpenDeparture` map through configured aliases to `DepartureCreated` / `ArrivalClosed`. The report shows normalized expected machine codes. An unchanged duplicate rejection is `NoChange`; unexpected failures remain `ExportError`. Missing business policies are not silently implemented to satisfy tests: S09 should expose a missing open-departure guard as FAIL.

The runner currently covers the eight Movement modes represented by the fixtures. Word preview/document assertions require separate complete inputs and expected text. All runtime operations use VBA; Python validation helpers are development-only and do not participate in a run. Runtime execution still requires explicit authorization under repository rules; static validation does not establish VBA runtime correctness.

### Loading a scenario manually in PEB

After loading the changed VBA, XAML and wsConfig through the normal workbook update workflow:

1. Switch PEB to TEST. Save and close `TestScenarios.xlsx`, `test_РУХ.xlsx` and `test_ЕЖОС.xlsx` before loading scenarios or running tests.
2. Select a ScenarioId in the new selector below the mode panel, then click `Завантажити сценарій`. The loader replaces writable ТВ data with that scenario's `tbTestHistory` rows and fills the PEB form from step 1. Loading again resets the scenario's history. Scenarios without history rows start with no populated event rows.
3. Run the normal `Експорт РУХ` action to exercise the exporter. Loading does not export an event automatically. Compare the actual result and ТВ with the expected decision/code shown in the status and the expected counts in `tbTestSteps`.
4. For S16, export the current step, then click `Наступний крок`. This fills the next step's form without resetting ТВ. Repeat through steps 2 and 3. Advancing does not verify whether the previous step was exported; the operator controls that sequence. An absent or duplicate next step reports an error and preserves the previous form/history.

Scenario loading is guarded by TEST mode and by equality with the configured TEST Movement target; it rejects the configured production path. The loader uses a private Excel instance, suppresses macros/events/link updates, and saves history only after the form update and layout flush succeed. Failures discard unsaved workbook changes and restore the prior form, mode and candidate projection. Workbook locks/read-only state are errors. Source workbook changes are not applied by editing or statically validating the VBA files.

Auto-fill is turned off when loading a step, candidates and preview are cleared, and caches are invalidated. Existing ТВ formulas are reapplied using their R1C1 templates; configured helper columns are never populated from fixture values. The user-cleared workbook currently has no formulas inside ТВ, so those helper columns remain blank rather than receiving invented formulas. Other native tables and scenario sheets are not changed by the loader.

`test.scenarios.fields` maps required step headers to PEB fields. Other form inputs are cleared, except explicit configured defaults. The initial fixtures lack `AcceptedOrderDate`, so `test.scenarios.defaults` explicitly supplies `15.10.2026`; this is test input, not a dependency on the current date or order lookup. Add optional columns named `AcceptedOrderDate`, `PersonRank`, `PositionCode`, `PersonPosition`, `ReportRank`, `ReportName`, `ReportPositionCode`, `IncomingNumber`, `IncomingDate`, `TicketNumber`, `TicketDate`, `Document`, `DocumentDate`, `DocumentNumber` to `tbTestSteps` to fill richer forms. Existing optional columns override defaults, including explicit blank values. Dates accept numeric Excel dates or full `DD.MM.YYYY` text. The loader does not synthesize directory records: Word preview still needs matching identities/positions/hospitals in the other configured test sources.

For the legacy manual selector, add new scenario IDs to the comma-separated `text.TestScenarioItems` configuration for dropdown discovery, or type their IDs directly. The test management page discovers scenarios automatically without configuration edits. Step IDs must start at 1 and be consecutive; each scenario and step must be unique. The automated runner compares `ExpectedDecision`, `ExpectedCode` and `ExpectedOpenCount`; unresolved business rules remain separate work.

Run `Scripts/Testing/PEB/validate_scenario_loader.py` with the bundled Python to validate routing, all scenario steps, native ТВ headers, formula-column protection and UI command contracts without starting Excel.

## Separately generated fixture set

The older generated fixture set below contains only synthetic records. It recreates the consumed source schemas in new workbooks; real source rows, unrelated sheets, formulas, names, links and comments are not copied. Original source files are read only to discover headers and table definitions. The builder verifies that their SHA-256 hashes remain unchanged. This is separate from the user-supplied `test_РУХ.xlsx`, which retains the user's one selected person and original native workbook structure.

## Sources

| Workbook | Sheets / tables |
| --- | --- |
| Dependencies/ШПО.xlsx | ОС (85 columns), АЛФ, Посади, Звання; original native table names retained |
| Dependencies/Накази.xlsx | Накази / tbOrders, starting at D2 |
| Dependencies/Установи.xlsx | Лікувальні Заклади / ЛЗК, starting at A3 |
| Dependencies/ЕЖОС_last.xlsx | 5. Тимчасово відсутні / ВІДСУТНІ, starting at A5 |
| Dependencies/РУХ.xlsx | Відсутні / tbMovement |

The older generated `Dependencies/РУХ.xlsx` schema comes from the configured Movement exporter mappings. It is retained for the original fixture generator/validator and is no longer the TEST export target. Unused source columns are retained but left blank. Names, tax IDs, position codes, hospital and orders in those generated sources are deliberately synthetic. Dates are fixed in October 2026; select order `115 / 15.10.2026` explicitly rather than relying on today's nearby-order list.

## UI environment switch

The button beside the auto-fill checkbox displays the current environment and its destination: `PROD → TEST` or `TEST → PROD`. It loads the complete corresponding configuration and schedules page reinitialization after the click handler returns. Switching resets form values, candidate selection and caches.

- Production configuration: `PROJECTS_EXCEL/config/PersonalEventBuilder/wsConfig.txt`.
- Test configuration: `PROJECTS_EXCEL/config/PersonalEventBuilder/wsConfig.test.txt`.
- Test data paths resolve from the existing `workbook/PersonalEventBuilder` workbook directory into this directory.
- Configuration refresh uses `ThisWorkbook::config.source`, so Ctrl+Alt+U preserves the selected environment.
- Test diagnostics are written to `%TEMP%/PROJECTS_EXCEL/PersonalEventBuilder/TEST`.
- Test Word output is routed into this fixture directory. Movement export is enabled only in the TEST profile to exercise the native test ТВ table; production flags remain unchanged.

No live workbook configuration is changed by generating these files. The UI switch becomes available after the changed VBA and XAML sources and production configuration are loaded through the normal workbook update workflow.

## Manual scenarios

`Scenarios.xlsx` contains:

- `Cases`: eight baseline export modes and eight history variants.
- `FormInput`: form and accepted-order values for each case.
- `PasteTreatment` and `PasteLeave`: values in the exact order of the respective form columns. Copy the row excluding CaseId. Disable auto-fill before pasting a full row.

There are 17 people, including one reporter, and 11 initial Movement rows. Baseline arrivals have a matching open departure; baseline departures have no previous events. Additional people cover closed history, same-type and other-type open departures, multiple open departures, arrival without departure, invalid chronology and a pending exception rule.

These are fixtures, not an automated test runner or implemented business rules. Expected exception decisions are explicitly unresolved. The current exporter does not yet implement every cross-event restriction. Do not interpret a successful preview as a completed history-policy test. Negative scenarios must be evaluated once the corresponding rules are defined and implemented.

Movement export changes the fixture workbook when enabled. Regenerate the fixture set before repeating scenarios that require the initial history.

## Regeneration and validation

Use the bundled Node and Python executables. Create a `node_modules` junction in `Scripts/Testing/PEB` pointing to the bundled Node packages and set `PEB_PYTHON` to the bundled Python executable. Run `Scripts/Testing/PEB/build_fixtures.mjs` with Node. It reads production configuration only to discover the source files and creates six dummy workbooks.

The `--config-patch` option prints an apply_patch patch for creating the complete test configuration from production. Apply that patch with apply_patch; when the test configuration already exists, review and update its differences instead of adding a second copy. Fixture generation does not replace configuration files automatically.

`Scripts/Testing/PEB/validate_fixtures.py` validates the older generated fixture set and requires regeneration of those files first. For the current user-supplied `test_*` copies and scenario loader, use `Scripts/Testing/PEB/validate_scenario_loader.py`. Both validators read XLSX ZIP/XML directly without starting Excel or executing VBA.