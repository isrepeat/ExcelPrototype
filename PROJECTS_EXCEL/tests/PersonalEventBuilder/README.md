# PEB dummy sources

## Business events and independent export routes

`PersonalEventBuilderEventRoutes.xml` is the shared TEST/PROD registry of 35 business EventIds. Each event may declare Word, Movement or both. Both exporters select only a text template. Movement operations are defined exclusively by EventId in VBA. Movement result XML contains only reusable text DSL. `ex_PEB_MovementPlan` owns event behavior and native column mappings; status names and shared data providers are configured in wsConfig. An absent Word route produces no Word text and needs no placeholder template. DOCX insertion uses the existing Begin/End markers.

Movement operations are Arrival, Departure and Transition, with configured medical-company profile variants. Transition validates one matching open source record and a nonduplicate target departure, closes the source and appends the target under one save. Missing/ambiguous sources and invalid dates fail before writes. A failure rolls back mutations in borrowed TEST workbooks; owned workbooks close without saving. Medical-company profiles explicitly filter the source destination; other ambiguous matching sources are rejected rather than choosing arbitrarily. Ordinary Departure rejects an existing open event and requires an explicit Transition instead.

Word policy Omit returns Outcome=Omitted and empty text without creating a document. Both directions of outpatient VLK use this policy. MoveFromOutpatientVlkToVacationForCuring closes outpatient VLK and opens medical leave in Movement, while Word uses the ordinary ToVacationForCuring template and anchor. Configured routing replaces ad hoc TemplateName overrides.

The active scenario path is now TestScenarios.EventRoutes.xlsx, generated alongside the preserved original TestScenarios.xlsx. It contains 54 scenarios and 55 steps. ExportEvent requires expectations in both Expected_РУХ and Expected_WORD under the same StepId. Word comparison remains active even when preview display is disabled. S15 verifies annual leave to hospital in both exporters; S17 verifies outpatient VLK to medical leave, S18 verifies outpatient VLK with an omitted Word paragraph, and S19/S20 verify missing/ambiguous source records. W11 and W15 now explicitly expect Omitted. These fixtures do not imply complete Movement coverage of every transition or runtime verification.

`ex_PEB_EventRoutes` resolves text references and obtains Movement operations from compiled event plans. `ex_PEB_MovementTemplates` caches text DSL; `ex_PEB_MovementPlan` caches native mappings/events. Route disposal clears both. Offline `event_routes.py` combines text definitions and native source contracts into an in-memory compatibility view. Its `movement_plans.py` helper does not execute VBA or derive expected workbook rows. Run validate_movement_templates.py and validate_word_routes.py for source/XML contracts. Fixture workbooks remain unchanged; import or runtime validation requires the authorized workbook workflow.

Historical orders now match the source dictionary: 284/01.10.2026, 285/02.10.2026, 254/01.09.2026 and 263/10.09.2026. S13 deliberately retains the future departure date 20.10 with order 303. S14 uses order 298 in both history and input so its duplicate identity remains exact. S16's backdated arrival on 14.10 remains under order 298 dated 15.10; it must not be replaced with order 297. Historical ticket order components were updated with their owning records; literal Word baselines and append-only journals remain unchanged.

Expected_РУХ currently has its summary in A1:E22 and native expected rows in A45:Y63, with a merged caption on row 44. test.expected.minimumGapRows reserves caption space. The OOXML bridge rejects overlapping tables or insufficient spacing before writing; VBA checks the same reserve before loading expectations. Explicit --layout-range changes can update a caption's XAML style/merge without resetting other cells or dimensions. Shrinking tables restores the surrounding worksheet background, and assembly invalidates the stale calculation chain.
Run validate_scenario_loader.py --inputs-only to validate the active fixture while retaining historical audit error values. Without this flag those values are still rejected. The flag does not clear or rewrite any journal cell.

## Catalog loading and page lifecycle

The TEST page uses the same page lifecycle as PROD: `obj_PEB_PgTests` initializes `obj_PageBase`, registers bindings and commands, creates `obj_PEB_PgTestsController`, loads data and calls `obj_PageBase.Render`. The controller updates bound data and handles commands; it never creates controls or Shapes. Master checkboxes and row checkboxes are declared in `ui/PersonalEventBuilder/TestsPage.xaml`, using ordinary controls and table cell templates. Button labels remain in wsConfig.

Shared control styles are explicitly declared by the page root's `styleSources` attribute. PROD uses `styleSources="CommonControlStyles.xaml"`; TEST uses `styleSources="../PersonalEventBuilder/CommonControlStyles.xaml"`. Paths resolve relative to the page's UI folder. Multiple files are separated by semicolons and loaded from left to right, followed by the page's own controlStyle definitions; later definitions replace earlier styles with the same name. With no attribute, only local styles are loaded. Declared missing, malformed or duplicate sources stop page rendering with a diagnostic; there is no implicit filename lookup. Imported files use a styleCatalog or styles root in the profiles namespace; style-pipeline rules remain on the page. Run `Scripts/PowerShell/Test-UiStyleSources.ps1` for static reference and precedence checks without Office.

Offline workbook migration scripts read their visual presets from `ui/PersonalEventBuilder/FixtureStyles.xaml` through `fixture_styles.mjs`. The adapter applies named styles without embedding colors, fonts or dimensions. The OOXML bridge still preserves existing workbook formatting; text-number protection (`@`) remains a data-type operation. Changing these XAML presets does not rewrite existing fixture workbooks automatically.

Both environments use `ui.folder=PersonalEventBuilder`. `ui.pageClass` selects the registered page class (`obj_PEB_PgMain` in PROD and `obj_PEB_PgTests` in TEST). `ThisWorkbook::ui.pageMappings` maps worksheet names to XAML paths relative to that UI folder: PROD uses `MainPage=MainPage.xaml`, TEST uses `MainPage=TestsPage.xaml`. Multiple mappings are separated by `|`, for example `MainPage=MainPage.xaml|Tests=TestsPage.xaml`. An explicit map is authoritative: unmapped worksheets are not rendered. Duplicate or nonexistent sheet names, missing files and paths outside the UI folder stop rendering with a configuration diagnostic. Workbooks without this key retain the existing `<sheet name>.xaml` convention. Shared style paths are resolved relative to each mapped XAML file.

`ex_PEB_Testing` retains catalog access and asynchronous runner scheduling; it borrows the active test controller through attach/detach and does not create a second UI. Import the new page/controller classes and refreshed wsConfig together before reinitializing the workbook. No XLSX migration or sheet rename is required.

Progress includes completed export steps and completed stages of the current step: floor(100 * (completedSteps + completedStages / stageCount) / totalSteps). Each step has eight equally weighted stages: restore the scenario, capture the initial snapshot, prepare export inputs, calculate and capture the definitive pre-export snapshot, execute the export, capture the post-export snapshot, assert the result, and archive the report and working books. Each milestone is reported after its operation finishes. Opening the execution session contributes zero; multi-step scenarios contribute one unit per export step. These are completed-work milestones, not elapsed-time estimates. An archive failure retains the partial percentage instead of reporting completion.

Scenario classification is stored in the Операція and Подія / сценарій columns of Тест_Сценарії. S01-S08 show departure/arrival and the event type separately. S09-S20 are Custom. W01-W34 are PreviewWord cases. The UI groups departure, arrival, Custom and PreviewWord cases in that order; both classification columns are filterable in the Smart table.

Displaying or refreshing the catalog uses the existing ADO query executor against the saved scenario and step worksheets only. It does not create Excel.Application, open working Movement/Absence books or initialize the exporter. Full scenario validation and the private execution session start on the first scheduled runner callback after Run.

The TEST XAML owns TestCatalog through the standard table control with Smart representation and Data.TestCatalog binding. The normal table-plan pipeline handles creation, resizing and removal. The test UI supplies raw data and recreates checkbox overlays after the table update; it does not add or unlist native tables itself. Page refresh goes through the active page Render contract so these overlays and result cells are restored too. Refresh is blocked during an active run.

## Scenario-specific Absence source

Absence fixtures contain leave departures only, including medical leave, but never inpatient treatment. S01, S05 and S11 explicitly use HasEvent=0; S10 and S15 retain only their leave history. Treatment events remain in Movement history and export steps. The 16 existing Absence rows are preserved; W01-W09 add nine explicit empty states.

`Тест_ЄЖОС` / `tbTestAbsence` in `TestScenarios.xlsx` defines the initial Absence source for each scenario. It contains ScenarioId, ІПН, HasEvent and native ЕЖОС column headers. HasEvent=1 inserts an event; a single HasEvent=0 row explicitly defines an empty state. Keep column 66 equal to the scenario participant's full name, including on an empty-state row. Multiple event rows are supported.

Before the first step, the scenario session validates these rows, replaces only that participant's rows in the configured TEST ЕЖОС table and restores Movement history. ЕЖОС has no IPN column, so replacement uses the validated full name. Other participants remain untouched. Subsequent steps retain the same state. The session opens TestScenarios, Movement and Absence books in its private Excel instance and rejects paths outside the test directory, open/locked books and read-only access. Keep all three books closed during catalog loading and execution.

Manual scenario commit saves both working sources. Automated preparation keeps them unsaved until the step archive is saved in TestScenarios; disposal discards pending changes. Saving multiple workbooks is not atomic. The exporter still receives isolated step inputs: restoring ЕЖОС does not yet add end-to-end lookup/UI assertions.

## Separate scenario workbook

### Configuration-backed reserved words

The XML is the source of ModeId values. Other reserved values (Action, RuleStatus, ExpectedDecision, StateCheck, Outcome, ExpectedCode, Operation and Result) remain in test.registry.* in the TEST configuration. Test_Registry/tbTestRegistry is a generated copy. RuntimeModeId is retained for schema compatibility and repeats the same EventId; no alias registry exists in configuration. Dropdowns reference _PEB_* names, not literal lists. Historical journals remain untouched.

ModeId transports 35 XML business-event identities, including MoveFromOutpatientVlkToVacationForCuring. ModeId validation and dropdown synchronization read the same XML as production routing. Each event independently selects its Movement plan and Word representation. Historical journals retain their original identifiers.

To change dropdown choices, edit the TEST configuration and import it into PEB through the normal configuration workflow. With TestScenarios closed, loading a scenario or initializing a test run refreshes the registry values, named ranges and dropdown bindings from the active configuration and saves the scenario workbook. Reopen it after the session closes to see the updated choices. Registry sheet edits are overwritten on refresh. Missing or duplicate values and unregistered input words produce explicit errors. Adding a list option does not create an executor or an export profile for a new action/mode.

For development without Office, update_registry_fixtures.mjs maintains the active workbook and normalize_inputs.mjs --registry-only refreshes the registry using XML-backed derived settings. Derived compatibility keys exist only in offline tool memory and are never written to wsConfig. validate_event_routes.py tests invalid XML references, duplicates, layout collisions, historical orders and S14/S16 semantics. validate_scenario_loader.py --inputs-only checks all active scenarios while retaining historical audit error cells.

`TestScenarios.xlsx` separates scenario inputs, expected outputs and actual run archives. The four input sheets remain `Тест_Сценарії` / `tbTestScenarios`, `Тест_Історія` / `tbTestHistory`, `Тест_Кроки` / `tbTestSteps`, and `Тест_ЄЖОС` / `tbTestAbsence`. Expectations are stored in `Expected_РУХ` and `Expected_WORD`; Movement retains only its native working sheets.

These tables describe 54 scenarios for the existing participant: 15 Movement-only scenarios, 34 Word-only scenarios and five combined scenarios. `tbTestHistory` includes ScenarioId, RowId and every existing column from ТВ. A scenario without history rows means that ТВ starts empty. Computed ТВ columns are intentionally blank in the fixture history; the loader preserves/reapplies the working table's formulas rather than overwriting them with blanks.

`tbTestSteps` contains 55 actions and their input values, with no Expected columns. StepId is a globally unique stable string such as `S16-M02` or `W01-P01`; StepOrder is the consecutive execution order within ScenarioId. Editing StepOrder or sorting worksheet rows must not change StepId. Action is ExportMovement, PreviewWord or ExportEvent. Each scenario represents one order with at most two events. S16 starts with annual leave from 01.10.2026: S16-M02 returns on 14.10, then S16-M03 departs for treatment on 15.10, both under order 298 dated 15.10.2026. The retired S16-M01 remains only in historical archives. S15 closes annual leave, creates hospital treatment and checks its Word text as Defined/ExportEvent. Decisions and texts are specifications, not evidence of successful runtime execution.

### Independent expectations

`Expected_РУХ` contains two native tables:

- `tbExpectedMovement`: StepId, ExpectedDecision, ExpectedCode, StateCheck, ExpectedOpenCount. Allow requires Snapshot; Block and NoChange require Unchanged; Pending requires Pending.
- `tbExpectedMovementRows`: StepId, RecordId and every writable native ТВ column. The 14 rows describe complete post-step snapshots for successful Movement operations. RecordId labels a logical row within a scenario and remains the same when an arrival closes that departure. A blank cell means an expected empty value. Helper columns configured in test.scenarios.formulaColumns are excluded explicitly.

The snapshot covers the complete native TEST ТВ table, consistent with the existing scenario-history reset. Rows match by the configured natural key ІПН/Подія/Вибуття/Вибуття.Наказ, independently of worksheet order. Duplicate keys, missing/extra rows, cell values and text/numeric type differences fail the assertion. Unchanged compares complete before/after signatures, including protected/formula integrity; rejected steps do not repeat snapshot rows. Operational checks still verify the affected row count and preservation of unmapped fields.

tbMovementEnums on Expected_РУХ provides the editable event list and default За межі values. The expected-row Подія dropdown uses the dynamic name _PEB_MovementEvents, referring to tbMovementEnums[Подія]. За межі is a calculated column: treatment or curing leave directed to Медична рота А7383 yields Ні; other populated events use XLOOKUP against the enumeration; an empty event yields an empty result. The expected sheet is recalculated before the runner reads its values. normalize_inputs.mjs --movement-enums-only refreshes these bindings without replacing sheet formatting. This expectation formula does not change production Movement formulas or initial history.

Movement departure/arrival basis fields contain the report text rendered from the shared IncomingBasis and OptionalBasisDocuments blocks, without the Підстава: label or vacation-ticket phrase. The same common Word dictionaries and aliases supply grammatical forms. Departure writes the qualified vacation-ticket number to Супровідний документ; arrival preserves that existing field. Tickets entered as a sequence are qualified using AcceptedOrderDate/AcceptedOrderNumber; already qualified numbers are retained. Word retains its vacation-ticket phrase. Fixture incoming numbers use 1656/<sequence>-в (ТЕСТ-...), and tickets use <orderYear>/<orderNumber>/7 (ТЕСТ-...). Test markers remain annotations within the numbers. normalize_inputs.mjs --document-fixtures-only migrates these fixture values and independent literal expectations while preserving layout and formulas.

Pre-existing history uses the reserved placeholders Тестовий рапорт про вибуття and Тестовий рапорт про прибуття for unknown earlier reports. Expected snapshots retain those placeholders for historical fields; reports created by actual test steps use the complete report text and annotated incoming number. An arrival updates its arrival basis while preserving the historical departure placeholder. Events created by preceding steps of the same scenario retain their real test report texts.

Each step's AcceptedOrderNumber/AcceptedOrderDate is resolved by the same obj_OrderResolver.TrySearch and orders.source configuration used by the UI, before seeding history or exporting. Invalid input, missing pairs, ambiguous results and source errors stop preparation. Identical duplicate pairs follow the resolver's existing deduplication policy; conflicting dates for a number within a year are rejected. Accepted fields are taken from the resolved record. test.fixtures.baseDate anchors the fixtures at 15.10.2026; chronology offsets remain unchanged. normalize_inputs.mjs --order-fixtures-only verifies the source and rebases fixture orders, tickets and expectations to its unique order for the anchor date (currently 298). It stops rather than selecting an arbitrary order when the date has zero or multiple candidates. Historical order numbers are not inferred from event dates, which are not historical order dates.

For every Movement departure and arrival, food support takes effect on MAX(EventDate, AcceptedOrderDate + 1 day), independently of За межі. The offset is configured by export.Movement.foodSupport.minimumOrderOffsetDays and must be at least one. Departure writes the date to Вибуття.Продовольче; arrival writes Прибуття.Продовольче and preserves the departure date. normalize_inputs.mjs --food-fixtures-only updates literal expectations for reports created by test steps, leaving unknown historical values intact. Expected food support cells use dd.mm.yyyy without changing their other formatting. S16 covers a backdated arrival and a current-day departure; both food support dates are 16.10.2026.

Word context covers all current XML root placeholders and metadata-row placeholders. The shared food-support calculation also provides RequiresFoodSupportEnroll/Remove and both short food-support dates. Additional hospital cases use the institution dictionary. Vacation DurationDays accepts PrototypeNew syntax such as 15+2 (дорога) +1 (донація крові), derives inclusive DateToShort, date-range flags and grammatical duration phrases. Optional certificate/previous-ticket values come from explicit extra inputs. The XML loader rejects undeclared placeholder names, and preparation verifies the complete declared context contract.

Extra Word data is a dictionary passed as the final wordInputs argument of obj_PEB_ExportService.TryPrepare, or stored in Data.WordContext. Allowed scalar keys are declared by export.Word.context.scalarInputs; Boolean flags and metadata collections have separate typed schemas. MetaDocumentTables rows contain DocumentNote/DocNo/DocDateShort. MetaTvoTables rows contain PersonRank/PersonId/PositionCode and are enriched into the grammatical forms used by the XML. MetaTables rows contain __sectionTitle/FIO/DocumentNote. Previous ticket numbers are supplied as recorded, without inferring an unknown historical order. Unknown keys, wrong types, invalid dates and incomplete metadata rows fail preparation. TemplateName may explicitly select another existing XML template; changing it requires explicit food-support transition flags. Current UI mode registration is unchanged. The XML documentFilepath is rendered into XmlDocumentFilepath; export.Word.output remains the configured output route and can reference that value.

Each template has export.Word.context.templates.<TemplateName>.requiredKeys, checked for nonempty values before rendering. Optional values have explicit defaults; mandatory hospital/report data cannot disappear into blank text. Extended inputs are supplied programmatically through the dictionary; this change adds no new form controls. For example, pass DurationDays="15+2 (дорога)", VlkNo="123", VlkDateShort="15.10.2026" in wordInputs when preparing the corresponding preview. Date inputs accept full date strings, VBA dates or Excel serial dates.

Grouping behavior is unchanged pending a separate decision. This context extension does not append to an existing DOCX or introduce batch grouping. VBA context/rendering behavior requires explicit runtime verification; the development validator audits all 34 XML templates, dictionary schemas and context-provider coverage without executing VBA.

If a copied test Movement workbook retains [2]!ТВ or [2]!Список references to the original workbook, xlsx_expectations_bridge.py localize-movement-formulas <path> rewrites those references to the local tables in cell and calculated-column formulas. Other external links, data, styles and dimensions are retained, and Excel recalculates the repaired formulas on opening.

`Expected_WORD` / `tbExpectedWord` contains StepId, Outcome, ExpectedText and ExpectedError. Success requires a complete literal text and no error; Error requires an exact preparation diagnostic and no text; Pending is skipped. The current Word service exposes diagnostic text rather than a structured error code, so an error baseline includes that diagnostic. Comparison is binary and normalizes only CRLF/CR to LF. Spaces, nonbreaking spaces, punctuation and blank lines remain significant. Word failure details identify the first differing line/column; complete expected/actual texts and diagnostics are archived.

W01-W34 and S15 cover 32 rendered XML templates and two explicit outpatient-VLK omissions; W09 expects the unknown-hospital lookup diagnostic. S17-S20 cover the shared medical-leave template and omission alongside Movement assertions. StatusChange Word baselines cover the branch without food support changes; alternative food support policies and optional-document branches are not covered. Positive treatment inputs and Word text baselines use КНП МКБЛ №25 ХМР from the institution dictionary. Hospital stores its full name; HospitalShortName stores the exact lookup key. W09 alone retains ТЕСТ-ЛЗ-01 as an unknown key. Text baselines are authored literals, independent of the production DSL renderer. PreviewWord does not create DOCX files or invoke Word and must leave Movement unchanged. DOCX layout, styles and insertion anchors remain outside these assertions.

Before any test writes working data, the execution catalog validates stable IDs, step order, expectation links, required native columns, snapshot record identities and archive schemas. Missing/orphan/duplicate expectations are fixture errors, not silently skipped checks. Runtime archives never update Expected tables.

The TEST configuration routes Movement export to `Dependencies/test_РУХ.xlsx`, sheet `Відсутні`, native table `ТВ`. Personnel and Absence use `test_ШПО.xlsx` and `test_ЕЖОС.xlsx`; orders and hospitals use the ordinary Dependencies folder. Active definitions and diagnostics use `TestScenarios.EventRoutes.xlsx`. Original TestScenarios.xlsx and any intermediate fixed copy are preserved archives, not maintenance inputs. Production workbook routing remains unchanged.

### Selecting and running tests on the main PEB page

After updating VBA, XAML and configuration, switch to TEST and click `Тестування РУХ`. The selector and live results are rendered on `MainPage`, below the reserved candidate area in columns B:Q. PEB no longer creates a separate test worksheet or stores the full report in the UI workbook. Existing old `PEB_Tests` sheets, if any, are not deleted. The panel uses the existing column widths, preserves the form/candidate areas and displays a bounded window of recent results.

- Check scenarios, use `Вибрати всі` / `Зняти вибір`, then `Запустити обрані`. Catalog entries are read automatically from the existing scenario tables; no Run column or manual workbook editing is needed.
- Each scenario starts from its defined seed history. Steps within that scenario accumulate their changes in order, as in S16. The real Movement preparation/export services run against isolated form bindings and the validated TEST workbook. The main form is preserved.
- Scheduling uses workbook-qualified `Application.OnTime` callbacks: initialize the private session, then execute one test step per callback. The full batch is not a synchronous loop and uses no DoEvents loop. Excel regains control between steps; editing a cell naturally delays OnTime until Excel is ready. A single VBA/COM step still executes synchronously and is not preemptible; opening/saving Excel books can briefly occupy the UI.
- The Excel status bar and PEB action status display the processed-step percentage and counters. `Зупинити` is available in the panel and beside the main test button. Cancellation prevents new export operations; remaining selected steps are archived as SKIP on subsequent short callbacks. A failed assertion skips dependent steps of that scenario while independent scenarios continue; an infrastructure error aborts remaining operations.
- Each persisted step appears immediately in the panel. Pending rules are skipped; the current catalog contains no PendingRule cases. Manual loading/export and environment switching are guarded during an active run. Reload and workbook close cancel the scheduled callback and release the private workbook/session.
- The results table displays every result in the current run. Its XAML cell template uses PreviewLink buttons with the source row record. The document Items template creates one block per result from Data.TestPreviews. Empty runs create no preview controls; a new run removes the previous controls. Movement steps retain diagnostic-only previews; standalone PreviewWord steps assert their text/error against Expected_WORD. Preparation diagnostics appear in preview blocks without replacing the archived actual text.

### Durable diagnostics in TestScenarios

Scenario tables remain input definitions; ТВ and other working tables are populated only by the relevant operations. Three permanent archive sheets/native filterable tables are prepared in `TestScenarios.xlsx`. Their dark table style, column formats and widths are stored in the workbook. VBA validates their schema and appends values without creating sheets or changing formatting. An empty last row is reused as the first data row. Missing archive sheets or tables stop the operation with a specific diagnostic.

| Sheet / table | Contents |
| --- | --- |
| Тест_Результати / tbTestResults | Run/scenario metadata, stable StepId and StepOrder, Action, operation, expected/actual decision/code/open count, StateCheck, MovementResult, WordResult, complete expected/actual preview text and diagnostics, PASS/FAIL/SKIP/ERROR, details, timing and run progress |
| Тест_Записи / tbTestRecords | Complete native ТВ column values before and after each step, with scenario/run metadata, Phase, StateAvailable, record index and Created/Closed/Modified/Deleted/Unchanged/Empty markers |
| Тест_ВхідніДані / tbTestInputs | Original fixture fields and prepared export values, stored as field/value rows with input stage and VBA value type |

Filter by RunId, ScenarioId, StepId, ModeId, EventName, Operation, Result or ChangeKind. For an unexpected close, compare the Before and After rows for the same run/scenario/step. Empty initial tables get an explicit Empty marker; unavailable snapshots after preparation errors are marked Unavailable rather than represented as valid empty state. Raw native cell types/values are retained in record snapshots. Formula expressions/protected helper integrity are checked without comparing volatile formula caches.

Word preparation includes the currently configured empty MetaTables/MetaDocumentTables/MetaTvoTables collections. The input archive records these as Collection(0) with vbObject type; nonempty collections or other objects require an explicit serializer and stop archival rather than losing their contents silently.

The automatic borrowed-workbook export does not save independently. Source operation changes and all three diagnostics tables are committed together by the scenario session after each step. A journal save failure is shown as ERROR and ends the job; unsaved source changes are discarded when closing the session. Runtime tables are append-only, schema-checked, and capped by the configured archive row limit. No unrelated sheet data is overwritten. Completed-step reports survive PEB reload/close; a hard interruption can leave the last RunState as Running with CompletedSteps below TotalSteps, which identifies an incomplete run. Keep the test source closed while running; open it after completion/cancellation to inspect/filter the archive.

Assertions compare machine codes, open counts, the exact affected row count, configured export/join fields, unchanged unrelated rows and unmapped fields, and unchanged writable/protected/formula state on rejection. Seeding trims only native test table rows (minimum one blank formula-template row), preventing accumulation of blank rows.

Fixture reason codes `NoOpenDeparture`, `PreviousDepartureClosed` and `MatchingOpenDeparture` map through configured aliases to `DepartureCreated` / `ArrivalClosed`. The report shows normalized expected machine codes. An unchanged duplicate rejection is `NoChange`; unexpected failures remain `ExportError`. S09 verifies that an ordinary departure cannot bypass an existing open event; explicit Transition routes handle status changes.

The runner covers the eight Movement modes and independent text preview assertions. Historical journal rows retain their data with numeric StepId migrated to stable IDs; their MovementResult is Legacy and WordResult is NotRun because old runs did not execute these new assertions. All runtime operations use VBA; development helpers do not participate in a run. Runtime execution still requires explicit authorization under repository rules; static validation does not establish VBA runtime correctness.

### Loading a scenario manually in PEB

After loading the changed VBA, XAML and wsConfig through the normal workbook update workflow:

1. Switch PEB to TEST. Save and close `TestScenarios.xlsx`, `test_РУХ.xlsx` and `test_ЕЖОС.xlsx` before loading scenarios or running tests.
2. Select a ScenarioId in the new selector below the mode panel, then click `Завантажити сценарій`. The loader replaces writable ТВ data with that scenario's `tbTestHistory` rows and fills the PEB form from step 1. Loading again resets the scenario's history. Scenarios without history rows start with no populated event rows.
3. Run the normal `Експорт РУХ` action for a Movement step or the Word preview action for a PreviewWord step. Loading does not execute the action automatically. Compare with the linked Expected_РУХ or Expected_WORD rows. The loaded status displays StepOrder, stable StepId and Action.
4. For S16, export the arrival step, then click `Наступний крок` to fill the departure step without resetting ТВ. There are two steps under one order. Advancing does not verify whether the previous step was exported; the operator controls that sequence. An absent or duplicate next step reports an error and preserves the previous form/history.

Scenario loading is guarded by TEST mode and by equality with the configured TEST Movement target; it rejects the configured production path. The loader uses a private Excel instance, suppresses macros/events/link updates, and saves history only after the form update and layout flush succeed. Failures discard unsaved workbook changes and restore the prior form, mode and candidate projection. Workbook locks/read-only state are errors. Source workbook changes are not applied by editing or statically validating the VBA files.

Auto-fill is turned off when loading a step, candidates and preview are cleared, and caches are invalidated. Existing ТВ formulas are reapplied using their R1C1 templates; configured helper columns are never populated from fixture values. The user-cleared workbook currently has no formulas inside ТВ, so those helper columns remain blank rather than receiving invented formulas. Other native tables and scenario sheets are not changed by the loader.

`tbTestSteps` has a universal English column schema, declared in test.scenarios.stepHeaders. Five metadata columns come first: ScenarioId, StepId, StepOrder, Action, ModeId. AcceptedOrderNumber/AcceptedOrderDate correspond to the order selector above the form. Form inputs then follow the visible form order as closely as possible: Destination/Hospital/HospitalShortName; PersonRank/PersonName/PersonId/PositionCode; ReportRank/ReportName/ReportPositionCode; IncomingNumber/IncomingDate; Document/DocumentDate/DocumentNumber; EventDate; TicketNumber/TicketDate. Notes is a final human-readable comment. PersonPosition remains excluded. Fill only the inputs relevant to the action; unused cells are blank. All columns exist, including reporter/document fields, and are read by name through test.scenarios.fields. ModeId selects the event name from configuration, so no duplicate event column is entered. Dates accept numeric Excel dates or full DD.MM.YYYY text. Word preview uses PositionCode for the position dictionary. The Movement destination file, worksheet and table come from export.Movement configuration; scenarios do not repeat their destination.

export.Movement.<ModeId>.destinationField selects the source of the native Куди value. Treatment uses HospitalShortName, while leave modes use Destination. ToHospital explicitly requires HospitalShortName. Hospital retains the full institution name. This rule is part of the production preparation service as well as tests; a successful preview does not implicitly populate a missing Movement destination.

For the legacy manual selector, add new scenario IDs to the comma-separated `text.TestScenarioItems` configuration for dropdown discovery, or type their IDs directly. The test management page discovers scenarios automatically. StepOrder starts at 1 and is consecutive within a scenario; StepId is globally unique and does not change when steps are reordered. The manual selector advances by StepOrder. Add matching expectation rows for every new action; unresolved business rules remain separate work.

Run `Scripts/Testing/PEB/validate_scenario_loader.py` with the bundled Python to validate routing, all scenario steps, native ТВ headers, formula-column protection and UI command contracts without starting Excel.

`Scripts/Testing/PEB/migrate_expectations.mjs` is a one-time migration for the legacy workbook schema. Set PEB_PYTHON to the bundled Python executable and use the bundled Node/packages. It refuses to migrate a workbook that already has Expected sheets; maintain approved baselines manually afterward. The bridge preserves untouched OOXML parts and existing custom table styles. An optional original-workbook path passed to the validator checks preservation of original inputs, history and journal values.

The legacy migration finishes by running normalize_inputs.mjs to produce the universal input schema and consistent hospital fixtures. The normalizer is a development migration, never part of a runtime test run; it deliberately updates treatment fixture inputs and hospital text baselines. For existing table schemas, the OOXML bridge updates cell contents in the original sheets, preserving cell styles, row heights, column widths, table formatting, views and other sheet objects. Schema replacement requires explicit --allow-layout-migration; the one-time legacy migration passes that flag. PEB_INPUTS_SOURCE can point to a temporary workbook for isolated checks. Preserving formatting does not restore formatting lost in earlier edits.

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

## Runner performance

The runner keeps fixture inputs and Word templates in memory for one run. Start a new run to pick up edited inputs or templates. Interactive previews continue to reload templates per operation. Each runner callback shares one configuration snapshot, including nested export operations. Registry contents, named ranges and dropdowns are checked before synchronization; unchanged registries are not rewritten.

Movement snapshots read values and formulas in bulk and retain formula, type and unrelated-row integrity checks. Journals are appended in batches. Modified dependency workbooks are saved after each step; the scenario journal is saved at scenario boundaries and explicitly flushed on completion, cancellation or a reported error. A process crash can lose journal entries from the current unfinished scenario.

The configuration defaults are `test.runner.tickSeconds=0`, `test.runner.batchMilliseconds=100`, and `test.runner.uiRefreshMilliseconds=100`. Batches yield between steps and keep cancellation checks inside steps. The time budget cannot interrupt a synchronous Excel operation.

The catalog has independent execution and preview checkboxes. The master checkbox above each column selects or clears that entire column; individual changes update its master state. New scenarios have previews enabled by default. Both choices are frozen when a run starts. With preview disabled, Movement steps skip the additional Word preparation and no preview document block is rendered. Dedicated `PreviewWord` steps still validate the expected text or diagnostic, but their document blocks are also omitted when preview is disabled. Mixed runs show only enabled preview blocks, and result links use the filtered block indices.

Selection changes mutate the existing catalog records and raw table values, then invalidate only values (`TestCatalog`, `layoutChanged=False`). The shared renderer refreshes existing cell-template bindings and changed cells without recreating Shapes, styles or click routes. Changed source identity, schema or row records use the regular render path. `Control.Table.RefreshTemplateBindings` records the fast update; a checkbox click should not produce `Control.Table.Render` or batches of `Control.Checkbox.Render` entries.

Performance logs include `Test.OpenMovement`, `Test.OpenScenarios`, `Test.OpenAbsence`, `Test.SyncRegistry`, `Test.Catalog`, `Test.Prepare`, `Test.SeedMovement`, `Test.Calculate`, `Test.Snapshot`, `Test.ArchiveBatch`, `Test.Save` and `Test.Archive`. These timings overlap where an outer stage includes inner stages; do not sum all entries to estimate total duration. Runtime performance must be measured after explicitly importing the updated VBA and configuration into the working workbook.

## Regeneration and validation

Use the bundled Node and Python executables. Create a `node_modules` junction in `Scripts/Testing/PEB` pointing to the bundled Node packages and set `PEB_PYTHON` to the bundled Python executable. Run `Scripts/Testing/PEB/build_fixtures.mjs` with Node. It reads production configuration only to discover the source files and creates six dummy workbooks.

The `--config-patch` option prints an apply_patch patch for creating the complete test configuration from production. Apply that patch with apply_patch; when the test configuration already exists, review and update its differences instead of adding a second copy. Fixture generation does not replace configuration files automatically.

`Scripts/Testing/PEB/validate_fixtures.py` validates the older generated fixture set and requires regeneration of those files first. For the current user-supplied `test_*` copies and scenario loader, use `Scripts/Testing/PEB/validate_scenario_loader.py`. Both validators read XLSX ZIP/XML directly without starting Excel or executing VBA.