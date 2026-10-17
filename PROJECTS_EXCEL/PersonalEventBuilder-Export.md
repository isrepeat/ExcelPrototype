# PersonalEventBuilder preview and export

The current page prepares one event from its editable form. `PreviewWord` displays generated text in H4:Q10, below the accepted order and above the personnel form. Every click captures fresh data and loads the XML templates again. It does not open Word or resolve a document output path. A mode switch or a failed new preview clears the old text.

Word export remains disabled by `export.Word.enabled=false`; its button is absent from the page. The retained command and exporter enforce the same switch. Movement is also disabled until its destination is configured. Runtime validation and source import into the working workbook require explicit user authorization.

## Template layers

1. `workbook/PersonalEventBuilder/PersonalEventBuilderWordResultTemplates.xml` defines point text, reusable DSL blocks and group headers. The transferred wording remains in this XML, without duplicate point bodies in wsConfig.
2. A DOCX document template defines the document layout and insertion anchors. This layer is not used for worksheet preview.

`PersonalEventBuilderEventRoutes.xml` selects independent Word and Movement text templates. Both exporter routes contain only `template`; Movement operations come exclusively from the native plan selected by EventId. Events may declare either exporter or both; missing tags disable that exporter, while bad references, duplicate tags and empty events are errors. The unused EventBody registry is gone; actual DOCX insertion still uses the configured Begin/End markers.

`PersonalEventBuilderMovementTemplates.xml` is a text-only `resultTemplates` document: local `sharedTemplates` and named `template` elements containing one or more `column name="NativeHeader"` blocks, each with its own `text` child. Text uses the same placeholder/condition/include/formatter DSL as Word. The XML contains no row search, operation, status, input schema or table-write commands. Texts are final cell values without adding or stripping a prefix. Vacation ticket wording remains local and independently editable.

A single captured event context contains form values, configured lookup results, aliases and computed values. Text uses `{[Name]}` and native code uses the same `Name` key. The experimental `Ref.Form.*` aliases were removed. Optional scalar/date/metadata inputs are declared under `export.shared.inputs` in wsConfig; the existing Data.WordContext snapshot remains the input transport for both exporters.

Each Movement template declares separate nested columns. Arrival uses `Прибуття.Підстава`, departure uses `Вибуття.Підстава`, and transition declares both. Shared blocks remain text-only and may be included by either column. `fn_Texts` supplies column-to-DSL definitions; preparation renders each independently into MovementColumnTexts. Native `fn_WriteValue` selects the rendered text by the actual write header, including separate closing and opening writes during transitions. Native validation rejects columns outside the operation's write set and missing required basis columns before touching the workbook. There is no single shared MovementBasis value; that name remains only an internal native mapping marker. Initial transition texts were duplicated unchanged to preserve behavior while allowing independent edits.

`ex_PEB_MovementTemplates` validates and caches text definitions and local includes only. `ex_PEB_MovementPlan` caches compiled declarations for five column-mapping profiles and 35 business event plans; eight native status names come from `export.Movement.status.<StatusId>`. `obj_PEB_ExptrMovement` executes append/close/transition, row matching, duplicate and chronology checks, guarded writes and rollback in VBA. The scenario verifier reads the same native mappings. Route disposal clears both caches. Adding a business event or changing table behavior requires a code change; changing basis wording only requires XML.

`obj_PEB_WordTemplates` loads the XML, checks root/version and duplicate names, and supplies the selected point and shared blocks. It disables DTD and external resolution. `obj_PEB_DslRenderer` replaces `obj_PEB_Template` and renders DSL independently of Word automation. `obj_PEB_ExportService` captures form data, performs configured lookups, maps aliases, selects the XML template and publishes its generated text through the existing preview binding.

## Configuration

Source paths, schemas, shared data providers, environment settings and user-facing messages are imported from `config/PersonalEventBuilder/wsConfig.txt` into wsConfig/tbConfig. Output plans and Movement column mappings live in their exporter XML files. Paths resolve relative to the workbook directory.

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

For one event, `groupHeader` blocks render in `groupBy` order before the point body. This is not whole-order grouping or sorting. The other 26 XML templates remain in the file without mode buttons.

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

The XML reader retains `documentFilepath`; the disabled DOCX exporter still uses `export.Word.output`. Its `documentMode` explicitly selects `blank` or `template`; template mode also requires `documentTemplate` and `bookmark`. DOCX export writes a temporary file and publishes it without overwriting an existing output.

The disabled Movement exporter retains configurable append/close profiles, field/identity/join mappings, chronology validation and duplicate checks. It uses a separate Excel instance and refuses an open, locked, protected or 1904-date-system destination. Its path/sheet/table values are placeholders until the user provides a destination.

## Static validation

Run `Scripts/Testing/PEB/validate_movement_templates.py` for TEST/PROD text DSL, native event/column contracts and VBA source structure. `movement_plans.py` reads literal VBA declarations solely for offline validation and fixture tooling; production never parses VBA source. The migration preserves 167 mapping properties, 35 event plans and the layout-normalized contents of all 35 basis texts. No Office code executes. Deploy the new native module, text XML, routes and selected configuration together through the authorized workbook workflow.

Run `Scripts/Testing/PEB/validate_word_routes.py` for TEST/PROD optional exporter routes, standalone Word-only and Movement-only registries, shared-template routing, and rejection of missing references, duplicate tags, empty events and retired sections. Word omission is an internal result of an absent word route.

Run `Scripts/PowerShell/Test-PEBWordTemplates.ps1`. It checks tab-separated configuration, namespace/version, duplicate template names, include references/cycles, group-header order, DSL block nesting, allowed #let functions and alias/lookup coverage for all eight configured modes without opening Office.

Static checks do not establish VBA runtime correctness. Sources have not been imported into the working workbook; VBA compilation, Word automation, ADO lookups and Movement writes have not been executed.