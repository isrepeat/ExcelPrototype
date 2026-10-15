# PEB test runner handoff

## Current state

The implementation is in source files. The working PEB workbook has not been updated with these changes, and VBA compilation/macros have not been executed. Static validation is not runtime verification. See README.md in this directory for detailed configuration, fixture schemas and manual workflows.

## Architecture

- wsConfig.txt and wsConfig.test.txt are complete PROD and TEST profiles. The environment button loads the other profile and schedules page reinitialization. TEST routes Movement export to Dependencies/test_РУХ.xlsx, sheet Відсутні, native table ТВ. Source paths and an explicit production-path rejection guard isolate test writes.
- TestScenarios.xlsx contains Тест_Сценарії/tbTestScenarios, Тест_Історія/tbTestHistory, Тест_Кроки/tbTestSteps and Тест_ЄЖОС/tbTestAbsence. There are 16 scenarios, 18 steps, 11 Movement history records and 16 Absence fixture rows. Absence contains leave departures only; S01/S05/S11 use explicit empty states, while treatment is tested in Movement. Definitions and reports are separate from the native test_РУХ and test_ЕЖОС working books. Native sheet/table XML and styles were retained during the move.
- obj_PEB_TestScenarios owns a private Excel instance and the validated test workbook. It seeds scenario history, prepares inputs, captures native before/after records and commits diagnostics. Initial history is reset once per scenario; subsequent steps retain preceding changes.
- obj_PEB_TestRunner freezes selected scenarios into a queue and executes one step per advance. It uses isolated form bindings and the existing PEB export preparation and Movement exporter. The visible PEB form is preserved. Assertions compare decisions, reason codes, open-event counts, affected rows, mapped values and preservation of unrelated data/formulas.
- obj_PEB_TestUi renders scenario selection and recent results on MainPage below the candidate area. ex_PEB_Testing schedules workbook-qualified Application.OnTime callbacks. The panel provides selection, run, refresh and cancel controls. Status-bar progress and results update after steps.
- The borrowed-workbook Movement exporter does not save independently. The scenario session saves operation changes and diagnostics together after each step. Reload/close cancels callbacks and disposes the private session. Environment switching and manual export/loading are blocked during a run.

## Run workflow after workbook update

1. Import/reload changed VBA, MainPage.xaml and wsConfig through the normal project workflow. Repository rules require explicit authorization before updating the working workbook or compiling/running VBA.
2. Switch to TEST. Save and close TestScenarios.xlsx, test_РУХ.xlsx and test_ЕЖОС.xlsx in the interactive Excel session.
3. Open the test panel in PEB, select scenarios, and start the run. Cancel prevents further export operations; queued steps are recorded as SKIP.
4. Inspect live results in PEB. After completion/cancellation, open TestScenarios.xlsx and filter the persistent diagnostics by RunId, ScenarioId and StepId.

## Persistent diagnostics

TestScenarios.xlsx contains three permanent native filterable report tables. Formatting is stored in the workbook; the runner only validates their schemas and writes values:

- Тест_Результати/tbTestResults: expected/actual outcome, codes and counts, PASS/FAIL/SKIP/ERROR, details, elapsed time and run progress.
- Тест_Записи/tbTestRecords: full native ТВ values before and after operations, including Created/Closed/Modified/Deleted/Unchanged markers.
- Тест_ВхідніДані/tbTestInputs: original fixture fields and prepared export values with value types.

The archive is append-only and schema-checked. A save failure ends the run with ERROR; unsaved changes are discarded. A hard interruption may leave the last persisted RunState as Running; compare CompletedSteps with TotalSteps to identify incomplete runs.

## Remaining work and known limits

- Runtime verification is outstanding: compile only with authorization; verify TEST switching, panel layout, selection, execution, cancellation, journal persistence, cleanup on reload/close, and absence of writes to production books.
- Scheduling is cooperative, not parallel. Excel regains control between callbacks. One export/open/save operation remains synchronous and cannot be interrupted midway; measure its runtime before claiming a fully responsive UI.
- S09 intentionally exposes the missing general guard against a new departure while another departure remains open. It should currently FAIL. Decide and implement the business policy separately.
- S15 is an unresolved exception policy and remains SKIP/PendingRule. Clarify allowed exceptions before adding assertions.
- Coverage currently targets the eight Movement export modes. Word preview/document output and additional modes need separate fixtures and expected output assertions.
- The initial step schema lacks AcceptedOrderDate; configuration explicitly supplies 15.10.2026. Richer form/Word cases need optional input columns and matching test directory records. Existing copied dependencies are not all anonymized synthetic data.
- The user-cleared native ТВ has no formulas inside its helper columns. Existing formulas are preserved/reapplied where present; absent formulas are not invented.
- Verify SaveAs behavior during a pending timer; there is no AfterSave callback rebinding implementation.
- The older Scenarios.xlsx and synthetic-fixture generator are separate from the active native test_РУХ scenario workflow. Do not regenerate or merge them over the current user workbook. validate_fixtures.py targets the older generated set and may report its missing dependencies.

## Development validation

- Scripts/Testing/PEB/validate_scenario_loader.py validates current profiles, native fixture schemas, scenario contracts and source/UI integration without Excel.
- Scripts/PowerShell/Test-PEBWordTemplates.ps1 validates Word template structure without running VBA.
- Scripts/PowerShell/Update-WorkbookModules.ps1 with -Profile PersonalEventBuilder -PlanOnly discovers the new modules without importing them.
- Preserve English VBA comments/documentation and configurable UI captions. All text edits use apply_patch; validate final bytes and trailing whitespace before committing.