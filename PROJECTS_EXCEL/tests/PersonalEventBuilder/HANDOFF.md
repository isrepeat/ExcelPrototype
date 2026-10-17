# PEB test runner handoff

## Current state

The implementation is in source files. The working PEB workbook has not been updated with these changes, and VBA compilation/macros have not been executed. Static validation is not runtime verification. See README.md in this directory for detailed configuration, fixture schemas and manual workflows.

## Architecture

- wsConfig.txt and wsConfig.test.txt are complete PROD and TEST profiles. The environment button loads the other profile and schedules page reinitialization. TEST routes Movement export to Dependencies/test_РУХ.xlsx, sheet Відсутні, native table ТВ. Source paths and an explicit production-path rejection guard isolate test writes.
- Active TestScenarios.EventRoutes.xlsx contains 54 scenarios and 55 stable steps, with independent Movement snapshots and literal Word expectations. The original books remain preserved. Historical synthetic order numbers and linked ticket components are reconciled with the order dictionary; S14 retains exact duplicate identity and S16 retains its backdated arrival under the current order. Journals and Word baselines were not rewritten.
- PersonalEventBuilderEventRoutes.xml owns 35 EventIds and optional independent exporters. Both exporter routes choose only a text template; the Movement operation comes from its native EventId plan. Movement XML contains resultTemplates/sharedTemplates/template/column/text. Each column renders independently into MovementColumnTexts; transitions declare both arrival and departure basis texts. Native writing selects text by header and rejects missing or unwritable text columns before mutation. ex_PEB_MovementPlan owns 35 native business plans and five column-mapping profiles; eight status names are configured under export.Movement.status. Shared providers and optional input schemas are under export.shared. Ref.Form aliases and XML table instructions were removed. Text rendering and transactional table mutation remain separate; route disposal clears both caches.
- obj_PEB_TestScenarios owns a private Excel instance and the validated test workbook. It seeds scenario history, prepares inputs, captures native before/after records and commits diagnostics. Initial history is reset once per scenario; subsequent steps retain preceding changes.
- obj_PEB_TestRunner freezes selected scenarios into a queue and executes ExportMovement, PreviewWord or combined ExportEvent per advance. Movement assertions compare independent snapshots, decisions, codes, counts and mutation integrity. Word assertions compare literal text/diagnostic/omission and require unchanged Movement during Word preparation. Preview display does not control assertions. The visible form is preserved.
- obj_PEB_PgTests and obj_PEB_PgTestsController use the standard page lifecycle and bound controls in ui/PersonalEventBuilder/TestsPage.xaml. ex_PEB_Testing only schedules workbook-qualified Application.OnTime callbacks and holds a borrowed controller reference. Worksheet-to-XAML mappings and the page class are selected through wsConfig; there is no separate test UI folder or manual checkbox renderer.
- The borrowed-workbook Movement exporter does not save independently. The scenario session saves operation changes and diagnostics together after each step. Reload/close cancels callbacks and disposes the private session. Environment switching and manual export/loading are blocked during a run.

## Run workflow after workbook update

1. Import/reload changed VBA, MainPage.xaml and wsConfig through the normal project workflow. Repository rules require explicit authorization before updating the working workbook or compiling/running VBA.
2. Switch to TEST. Save and close TestScenarios.EventRoutes.xlsx, test_РУХ.xlsx and test_ЕЖОС.xlsx in the interactive Excel session.
3. Open the test panel in PEB, select scenarios, and start the run. Cancel prevents further export operations; queued steps are recorded as SKIP.
4. Inspect live results in PEB. After completion/cancellation, open TestScenarios.EventRoutes.xlsx and filter diagnostics by RunId, ScenarioId and StepId.

## Persistent diagnostics

TestScenarios.xlsx contains three permanent native filterable report tables. Formatting is stored in the workbook; the runner only validates their schemas and writes values:

- Тест_Результати/tbTestResults: stable StepId, StepOrder, Action, expected/actual outcome, codes/counts, StateCheck, MovementResult/WordResult, full expected/actual preview texts/diagnostics, PASS/FAIL/SKIP/ERROR, details, timing and progress. Legacy runs are marked Legacy/NotRun for the new component checks.
- Тест_Записи/tbTestRecords: full native ТВ values before and after operations, including Created/Closed/Modified/Deleted/Unchanged markers.
- Тест_ВхідніДані/tbTestInputs: original fixture fields and prepared export values with value types.

The archive is append-only and schema-checked. A save failure ends the run with ERROR; unsaved changes are discarded. A hard interruption may leave the last persisted RunState as Running; compare CompletedSteps with TotalSteps to identify incomplete runs.

## Remaining work and known limits

- Runtime verification is outstanding: compile only with authorization; verify TEST switching, panel layout, selection, execution, cancellation, journal persistence, cleanup on reload/close, and absence of writes to production books.
- Scheduling is cooperative, not parallel. Excel regains control between callbacks. One export/open/save operation remains synchronous and cannot be interrupted midway; measure its runtime before claiming a fully responsive UI.
- Ordinary departures now reject an existing open event (S09); an explicit Transition event is required to close the source and open the target. Runtime verification remains outstanding.
- S15 is a Defined ExportEvent assertion for MoveFromAnnualVacationToHospital. S17-S20 add combined routing assertions for outpatient VLK, medical leave and missing/ambiguous source records.
- TestScenarios.EventRoutes.xlsx is the active fixture: 54 scenarios and 55 steps, with 32 rendered XML templates and two deliberate outpatient-VLK omissions. Configured Movement plans support the business-event registry; fixture coverage remains partial. S16 is a backdated arrival followed by a current-day departure under one order. Actual DOCX layout, formatting and insertion remain separate integration work.
- All scenario inputs, including order dates, rank, position and document details, are stored in tbTestSteps. The loader does not supply configured default data. Word cases require matching test directory records. Existing copied dependencies are not all anonymized synthetic data.
- The user-cleared native ТВ has no formulas inside its helper columns. Existing formulas are preserved/reapplied where present; absent formulas are not invented.
- Verify SaveAs behavior during a pending timer; there is no AfterSave callback rebinding implementation.
- The older Scenarios.xlsx and synthetic-fixture generator are separate from the active native test_РУХ scenario workflow. Do not regenerate or merge them over the current user workbook. validate_fixtures.py targets the older generated set and may report its missing dependencies.

## Development validation

- Scripts/Testing/PEB/validate_scenario_loader.py validates current profiles, native fixture schemas, scenario contracts and source/UI integration without Excel.
- Scripts/PowerShell/Test-PEBWordTemplates.ps1 validates Word template structure without running VBA.
- Scripts/PowerShell/Update-WorkbookModules.ps1 with -Profile PersonalEventBuilder -VbaFolderPath PROJECTS_EXCEL/vba -PlanOnly discovers the modules without importing them.
- Preserve English VBA comments/documentation and configurable UI captions. All text edits use apply_patch; validate final bytes and trailing whitespace before committing.