# PEB test runner handoff

## Current state

The implementation is in source files. The working PEB workbook has not been updated with these changes, and VBA compilation/macros have not been executed. Static validation is not runtime verification. See README.md in this directory for detailed configuration, fixture schemas and manual workflows.

## Architecture

- wsConfig.txt and wsConfig.test.txt are complete PROD and TEST profiles. The environment button loads the other profile and schedules page reinitialization. TEST routes Movement export to Dependencies/test_РУХ.xlsx, sheet Відсутні, native table ТВ. Source paths and an explicit production-path rejection guard isolate test writes.
- TestScenarios.xlsx contains 25 scenarios and 27 steps: the original 16 Movement scenarios/18 steps plus nine independent Word previews. StepId is a stable globally unique string; StepOrder controls sequencing. Inputs remain in Тест_Кроки, Movement outcomes and 14 full expected rows are in Expected_РУХ, and eight literal preview texts plus one lookup error are in Expected_WORD. Original history/Absence data and journal values are preserved; historical numeric StepId values are migrated.
- obj_PEB_TestScenarios owns a private Excel instance and the validated test workbook. It seeds scenario history, prepares inputs, captures native before/after records and commits diagnostics. Initial history is reset once per scenario; subsequent steps retain preceding changes.
- obj_PEB_TestRunner freezes selected scenarios into a queue and executes ExportMovement or PreviewWord per advance. Movement assertions compare independent full snapshots, decisions, codes, counts and mutation integrity. Word assertions compare exact text/diagnostic with newline normalization only and require unchanged Movement state. Linked expectation and archive schemas are validated before source mutation. The visible PEB form is preserved.
- obj_PEB_TestUi renders scenario selection and recent results on MainPage below the candidate area. ex_PEB_Testing schedules workbook-qualified Application.OnTime callbacks. The panel provides selection, run, refresh and cancel controls. Status-bar progress and results update after steps.
- The borrowed-workbook Movement exporter does not save independently. The scenario session saves operation changes and diagnostics together after each step. Reload/close cancels callbacks and disposes the private session. Environment switching and manual export/loading are blocked during a run.

## Run workflow after workbook update

1. Import/reload changed VBA, MainPage.xaml and wsConfig through the normal project workflow. Repository rules require explicit authorization before updating the working workbook or compiling/running VBA.
2. Switch to TEST. Save and close TestScenarios.xlsx, test_РУХ.xlsx and test_ЕЖОС.xlsx in the interactive Excel session.
3. Open the test panel in PEB, select scenarios, and start the run. Cancel prevents further export operations; queued steps are recorded as SKIP.
4. Inspect live results in PEB. After completion/cancellation, open TestScenarios.xlsx and filter the persistent diagnostics by RunId, ScenarioId and StepId.

## Persistent diagnostics

TestScenarios.xlsx contains three permanent native filterable report tables. Formatting is stored in the workbook; the runner only validates their schemas and writes values:

- Тест_Результати/tbTestResults: stable StepId, StepOrder, Action, expected/actual outcome, codes/counts, StateCheck, MovementResult/WordResult, full expected/actual preview texts/diagnostics, PASS/FAIL/SKIP/ERROR, details, timing and progress. Legacy runs are marked Legacy/NotRun for the new component checks.
- Тест_Записи/tbTestRecords: full native ТВ values before and after operations, including Created/Closed/Modified/Deleted/Unchanged markers.
- Тест_ВхідніДані/tbTestInputs: original fixture fields and prepared export values with value types.

The archive is append-only and schema-checked. A save failure ends the run with ERROR; unsaved changes are discarded. A hard interruption may leave the last persisted RunState as Running; compare CompletedSteps with TotalSteps to identify incomplete runs.

## Remaining work and known limits

- Runtime verification is outstanding: compile only with authorization; verify TEST switching, panel layout, selection, execution, cancellation, journal persistence, cleanup on reload/close, and absence of writes to production books.
- Scheduling is cooperative, not parallel. Excel regains control between callbacks. One export/open/save operation remains synchronous and cannot be interrupted midway; measure its runtime before claiming a fully responsive UI.
- S09 intentionally exposes the missing general guard against a new departure while another departure remains open. It should currently FAIL. Decide and implement the business policy separately.
- S15 is an unresolved exception policy and remains SKIP/PendingRule. Clarify allowed exceptions before adding assertions.
- Coverage includes all eight Movement modes and eight Word text templates, plus an unknown-hospital error. Actual DOCX layout, formatting and anchored insertion remain separate integration work.
- All scenario inputs, including order dates, rank, position and document details, are stored in tbTestSteps. The loader does not supply configured default data. Word cases require matching test directory records. Existing copied dependencies are not all anonymized synthetic data.
- The user-cleared native ТВ has no formulas inside its helper columns. Existing formulas are preserved/reapplied where present; absent formulas are not invented.
- Verify SaveAs behavior during a pending timer; there is no AfterSave callback rebinding implementation.
- The older Scenarios.xlsx and synthetic-fixture generator are separate from the active native test_РУХ scenario workflow. Do not regenerate or merge them over the current user workbook. validate_fixtures.py targets the older generated set and may report its missing dependencies.

## Development validation

- Scripts/Testing/PEB/validate_scenario_loader.py validates current profiles, native fixture schemas, scenario contracts and source/UI integration without Excel.
- Scripts/PowerShell/Test-PEBWordTemplates.ps1 validates Word template structure without running VBA.
- Scripts/PowerShell/Update-WorkbookModules.ps1 with -Profile PersonalEventBuilder -VbaFolderPath PROJECTS_EXCEL/vba -PlanOnly discovers the modules without importing them.
- Preserve English VBA comments/documentation and configurable UI captions. All text edits use apply_patch; validate final bytes and trailing whitespace before committing.