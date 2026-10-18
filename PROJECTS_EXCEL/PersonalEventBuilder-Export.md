# PersonalEventBuilder preview and export

The current page prepares one event from its editable form. `PreviewWord` displays generated text in H4:Q10, below the accepted order and above the personnel form. Every click captures fresh data and loads the XML templates again. It does not open Word or resolve a document output path. A mode switch or a failed new preview clears the old text.

Word export remains disabled by `export.Word.enabled=false`; its button is absent from the page. The retained command and exporter enforce the same switch. Movement is also disabled until its destination is configured. Runtime validation and source import into the working workbook require explicit user authorization.

## Template layers

1. `workbook/PersonalEventBuilder/PersonalEventBuilderWordResultTemplates.xml` defines point text, reusable DSL blocks and group headers. The transferred wording remains in this XML, without duplicate point bodies in wsConfig.
2. A DOCX document template defines the document layout and insertion anchors. This layer is not used for worksheet preview.

`PersonalEventBuilderEventRoutes.xml` selects independent Word and Movement text templates. Both exporter routes contain only `template`; Movement operations come exclusively from the native plan selected by EventId. Events may declare either exporter or both; missing tags disable that exporter, while bad references, duplicate tags and empty events are errors. The unused EventBody registry is gone; actual DOCX insertion still uses the configured Begin/End markers.

`PersonalEventBuilderMovementTemplates.xml` is a text-only `resultTemplates` document: local `sharedTemplates` and named `template` elements containing one or more `column configKey="export.Movement.headers.ArrivalBasis"` blocks, each with its own `text` child. Text uses the same placeholder/condition/include/formatter DSL as Word. The XML contains no row search, operation, status, input schema or table-write commands. Texts are final cell values without adding or stripping a prefix. Vacation ticket wording remains local and independently editable.

A single captured event context contains form values, configured lookup results, aliases and computed values. Text uses `{[Name]}` and native code uses the same `Name` key. The experimental `Ref.Form.*` aliases were removed. Optional scalar/date/metadata inputs are declared under `export.shared.inputs` in wsConfig; the existing Data.WordContext snapshot remains the input transport for both exporters.

Each Movement template declares separate nested columns. Arrival references `export.Movement.headers.ArrivalBasis`, departure references `export.Movement.headers.DepartureBasis`, and transition declares both. The loader resolves each required `configKey` through wsConfig before validating writable columns and duplicate resolved headers. Shared blocks remain text-only and may be included by either column. `fn_Texts` supplies column-to-DSL definitions; preparation renders each independently into MovementColumnTexts. Native `fn_WriteValue` selects the rendered text by the actual write header, including separate closing and opening writes during transitions. Native validation rejects columns outside the operation's write set and missing required basis columns before touching the workbook. There is no single shared MovementBasis value; that name remains only an internal native mapping marker. Initial transition texts were duplicated unchanged to preserve behavior while allowing independent edits.

`ex_PEB_MovementTemplates` validates and caches text definitions and local includes only. `ex_PEB_MovementPlan` caches compiled declarations for five column-mapping profiles and 35 business event plans; eight native status names come from `export.Movement.status.<StatusId>`. `obj_PEB_ExptrMovement` executes append/close/transition, row matching, duplicate and chronology checks, guarded writes and rollback in VBA. The scenario verifier reads the same native mappings. Route disposal clears both caches. Adding a business event or changing table behavior requires a code change; changing basis wording only requires XML.

`obj_PEB_WordTemplates` loads the XML, checks root/version and duplicate names, and supplies the selected point and shared blocks. It disables DTD and external resolution. `obj_PEB_DslRenderer` replaces `obj_PEB_Template` and renders DSL independently of Word automation. `obj_PEB_ExportService` captures form data, performs configured lookups, maps aliases, selects the XML template and publishes its generated text through the existing preview binding.

## Configuration

Source paths, schemas, shared data providers, environment settings and user-facing messages are imported from `config/PersonalEventBuilder/wsConfig.txt` into wsConfig/tbConfig. Movement header names are required settings under `export.Movement.headers.<HeaderId>`; the fixed medical-company destination is `export.Movement.destinations.MedicalCompany`. Native plans reference these settings through `fn_Config`, with no hardcoded destination/header fallback. Update the selected TEST/PROD configuration together with VBA. XML columns reference these same keys through their required `configKey` attribute; literal header names are rejected. Output text lives in exporter XML files; native code defines the operation mappings. Paths resolve relative to the workbook directory.

- `export.Word.templates.path`: XML path, currently `./PersonalEventBuilderWordResultTemplates.xml`.
- `events.path`: common event routes, including the selected Word and Movement template names.
- `export.Movement.templates.path`: Movement XML path, currently `./PersonalEventBuilderMovementTemplates.xml`.
- `export.Word.<ModeId>.previewRequired`: required form values for preview. Accepting an order is not required just to display a point.
- `export.Word.<ModeId>.required`: required values for the retained DOCX export.
- `export.fields` and `export.dateFields`: fields to capture and dates to validate. Text dates require `DD.MM.YYYY`; impossible dates and omitted years are rejected.
- `export.shared.aliases`: numbered `name`, `kind`, `source` mappings for both exporters. Kinds are `value`, `shortDate` and `joinLowerFirst`.
- `export.shared.constants`: explicit initial values for optional inputs. Kinds are `text`, `emptyText`, `boolean` and `emptyCollection`; text and Boolean constants also require `value`.
- `export.dsl.maxDepth`: include/block recursion limit from 1 to 64.
- `export.dsl.months.<1..12>`: Ukrainian month names for escaped DSL date patterns.

Four personnel lookups read rank genitive, person accusative, position genitive and abbreviated person genitive from the configured SHPO dictionary. Three reporter lookups run when the configured ReportName condition identifies an external reporter. An empty or self reporter skips these lookups and lets XML select the self-report branch. Treatment modes additionally resolve hospital genitive or accusative from the institution dictionary.

Each lookup under `export.shared.lookups.<index>` or `export.Word.<ModeId>.lookups.<index>` has `path`, `sheet`, `range`, `input`, `keyHeader`, `valueHeader`, `output`, `when.field` and `when.pattern`. Shared lookups supply both exporters; mode-specific Word lookups remain separate. Its parent `.count` selects how many run. A matched condition requires exactly one nonempty result; duplicate/missing results stop preparation. Paths, sheets, ranges and headers occur only in configuration. Open sources use the existing Excel query backend; closed sources use ADO.

## Mode mappings

| Current mode | XML template |
| --- | --- |
| FromHospital | FromHospital |
| FromAnnualMainVacation | FromAnnualMainVacation |
| FromVacationForCuring | FromVacationForCuring |
| FromVacationForFamily | FromVacationForFamily |
| ToHospital | ToHospital |
| ToAnnualMainVacation | ToAnnualMainVacation |
| ToVacationForCuring | ToVacationForCuring |
| ToVacationForFamily | ToVacationForFamily |

Word template XML uses version `2`. Former `groupHeader`, `groupBy` and `groupOrder` declarations have been folded into ordinary template text without changing their single-event DSL. Persistent grouping is declared explicitly through `group`.

## Supported DSL

The renderer supports scalar tokens (`{Name}`, `{[Alias]}`), shared `#include`, nested `#if/#else/#endif`, `#not/#and/#or`, `#for/#endfor`, `{#newline}` and `{#space}`. Loop items are dictionaries in a Collection, with `item.__first`, `item.__last`, `item.__index` and `item.[Alias]`.

Formatters are `upper`, `lower`, `upperFirstLetter`, `lowerFirstLetter`, `regexreplace`, `dateoffset`, `dateformat`, `nbspAfterNumberSign` and `nbspBeforeInitials`. Quoted pipeline arguments preserve commas and pipes inside regexes. Dates accept VBA patterns or the original escaped tokens `\dd`, `\d`, `\mm`, `\m`, `\yy`, `\yyyy`, `\month`.

Use `{[HospitalAccusative]|lowerFirstLetter|nbspAfterNumberSign}` to bind the number sign to a following digit with U+00A0, including input without a space. Use `{[FIOInitialsGenitive]|nbspBeforeInitials}` or `{[ReporterGenitive]|nbspBeforeInitials}` to bind a surname to dotted initials and preserve spaced initials as one nonbreaking name. These opt-in formatters only change rendered values; source data and grouping keys remain unchanged. Spaces and tabs are normalized locally; line breaks remain intact.

Use `{[Destination]|nbspAfterSettlementAbbr}` to bind settlement prefixes `м.`, `с.` and `сел.` to the following name with U+00A0. Matching is case-insensitive and accepts missing spaces, tabs and existing nonbreaking spaces. Token boundaries prevent matching suffixes inside words; line breaks are preserved.

Both plain and braced `#^` and `#_` layout markers are processed before inserting data. Final normalization preserves line breaks and nonbreaking spaces. User values are inserted without reparsing them as DSL.

`#let` only accepts internal variable names starting with `__` and the two transferred operations `$ex_Helpers.m_RegexIsMatch` and `$ex_Helpers.fn_IsDateOnOrAfter`. The renderer evaluates those operations itself; it never dispatches arbitrary VBA named by XML. Unknown aliases, includes and operations stop preview with a diagnostic.

## Remaining scope

The current form explicitly supplies empty MetaTables, MetaDocumentTables and MetaTvoTables. Vacation duration, additional days, end date and VLK-specific fields are also explicitly empty because their inputs are not available. These declarations are configured, not automatic fallbacks for missing DSL fields.

Whole-order aggregation, anchored insertion into an existing order, food support calculations, TVO chains, status changes, HTML table rendering and export deletion/undo require separate input fields and business rules.

The XML reader retains `documentFilepath`; ordinary DOCX exports use `export.Word.output` and refuse an existing output. `documentMode` selects `blank` or `template`; template mode requires `documentTemplate` and the configured paired text anchors.

## Persistent Word grouping

Each `template` contains zero or more `derived` declarations followed by exactly one body: `text` for ordinary output or `group` for persistent grouping. The first-level tags are therefore `derived`, `text` and `group`; mixing `text` with `group`, declaring derived values after the body, or using unknown tags/attributes is rejected. Derived values are evaluated in declaration order for each event and can reference earlier derived values. They are available to both body forms. Grouped templates also require `identityBy` and `orderBy`; ordinary templates cannot declare these state keys.

`ToBusinessTrip` is a Word-only departure route. Its root `group` contains the actual single-person wording inside `single/text`; no outer text or intermediate `WordText` placeholder is needed. Preview and persistent DOCX rebuilding call the same renderer. Preview evaluates one event through `single`; after a matching second event is added, rebuilding selects `multiple`.

`by` is a composite equality key; groups and people retain their first-seen order. `multiple` renders one header, each item and one footer. Nested groups are supported inside `multiple`. Shared header/footer blocks declare `common` scalar fields, which must agree across every member. `GroupBasis` is derived from the entire incoming-basis block, so reports by different authors do not merge merely because their incoming numbers match.

`numberRanges` uses three explicit regex captures: prefix, integer and suffix. It sorts each series numerically, compresses consecutive integers using full endpoint numbers, preserves gaps, suffixes and zero-padding widths, and rejects duplicate or malformed document numbers. Distinct document dates are separate grouping keys. The first implementation accepts integers within VBA Long range and does not normalize destination synonyms; destinations must match exactly.

`ex_PEB_WordState` stores configured scalar snapshots, the order key, template/shared-template definitions and the last generated section text in a Custom XML part with namespace `urn:excelprototype:word-events:v1` and payload version `2`. `export.Word.grouping.stateFields` declares raw inputs; rendered per-person `WordText` is not stored. The namespace stays stable so older stores are detected: version `1` requires migration and is not silently reinterpreted. The store is part of the DOCX and does not require a sidecar file. Identical retries are idempotent; a conflicting event with the same identity, a different order, manual section edits, duplicate stores, incompatible versions or changed template definitions stop export. Existing nonempty sections without stored data require migration. This version adds events; replacing or deleting an existing event is deliberately rejected rather than inferred.

Grouped exports use `export.Word.grouping.output`, which identifies an order rather than a person. On repeat export the output is copied to a temporary file, opened, checked and rebuilt from stored events. Only after saving and closing Word is the original replaced through `ReplaceFileW`; a changed output stamp aborts publication. Blank mode creates the managed anchor pair. Template mode requires a `ToBusinessTrip` Begin/End pair in the selected DOCX; it never invents a location in an existing template. Ordinary exports retain their previous output behavior. Interactive preview still shows the new person's single point; the persisted grouped result is assembled during DOCX export.

`Scripts/Testing/PEB/validate_word_grouping.py` checks configuration, persistence boundaries, 12 grammar cases and workbook scenario definitions without Office. Both UI suites use `TestScenarios.EventRoutes.xlsx`: the shared catalog selects Basic or Grouping through its Suite column. The 16 grouping scenarios use 34 steps in `tbGroupingSteps` and 25 expectations in `tbGroupingExpected`; their mappings are configured under `test.scenarios`. After explicit deployment and authorization, run `ex_PEB_GroupingTests.fn_Run` for the real Word scenarios for business scenarios only. Technical structure cases run separately through `Scripts/Testing/PEB/validate_word_grammar.py`, also called by `validate_word_grouping.py`; this validates the Python structural contract and does not execute VBA rendering or resolution. The business suite uses temporary DOCX files, saves/closes/reopens after every add, compares the first stored point with single-event preview, verifies preserved outside content and persisted event counts, and checks single-to-group conversion, ranges, split keys, retries and rejected mutations. It also checks the payload version and absence of persisted `WordText` fields. It does not edit the scenario workbook or current order files.

The disabled Movement exporter retains configurable append/close profiles, field/identity/join mappings, chronology validation and duplicate checks. It uses a separate Excel instance and refuses an open, locked, protected or 1904-date-system destination. Its path/sheet/table values are placeholders until the user provides a destination.

## Static validation

Run `Scripts/Testing/PEB/validate_movement_templates.py` for TEST/PROD text DSL, native event/column contracts and VBA source structure. `movement_plans.py` reads literal VBA declarations solely for offline validation and fixture tooling; production never parses VBA source. The migration preserves 167 mapping properties, 35 event plans and the layout-normalized contents of all 35 basis texts. No Office code executes. Deploy the new native module, text XML, routes and selected configuration together through the authorized workbook workflow.

Run `Scripts/Testing/PEB/validate_word_routes.py` for TEST/PROD optional exporter routes, standalone Word-only and Movement-only registries, shared-template routing, and rejection of missing references, duplicate tags, empty events and retired sections. Word omission is an internal result of an absent word route.

Run `Scripts/PowerShell/Test-PEBWordTemplates.ps1`. It checks tab-separated configuration, namespace/version, duplicate template names, exclusive template bodies, include references/cycles, DSL block nesting, allowed #let functions and declared context fields across all template text blocks without opening Office. Use the Python validators above for full grouping grammar and context-provider contracts.

Static checks do not establish VBA runtime correctness. Sources have not been imported into the working workbook; VBA compilation, Word automation, ADO lookups and Movement writes have not been executed.