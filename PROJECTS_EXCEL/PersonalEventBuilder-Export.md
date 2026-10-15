# PersonalEventBuilder preview and export

The current page prepares one event from its editable form. `PreviewWord` displays generated text in H4:Q10, below the accepted order and above the personnel form. Every click captures fresh data and loads the XML templates again. It does not open Word or resolve a document output path. A mode switch or a failed new preview clears the old text.

Word export remains disabled by `export.Word.enabled=false`; its button is absent from the page. The retained command and exporter enforce the same switch. Movement is also disabled until its destination is configured. Runtime validation and source import into the working workbook require explicit user authorization.

## Template layers

1. `workbook/PersonalEventBuilder/PersonalEventBuilderWordResultTemplates.xml` defines point text, reusable DSL blocks and group headers. The transferred wording remains in this XML, without duplicate point bodies in wsConfig.
2. A DOCX document template defines the document layout and insertion anchors. This layer is not used for worksheet preview.

`obj_PEB_WordTemplates` loads the XML, checks root/version and duplicate names, and supplies the selected point and shared blocks. It disables DTD and external resolution. `obj_PEB_DslRenderer` replaces `obj_PEB_Template` and renders DSL independently of Word automation. `obj_PEB_ExportService` captures form data, performs configured lookups, maps aliases, selects the XML template and publishes its generated text through the existing preview binding.

## Configuration

All runtime settings, source schemas and alias mappings are imported from `config/PersonalEventBuilder/wsConfig.txt` into wsConfig/tbConfig. Paths resolve relative to the workbook directory.

- `export.Word.templates.path`: XML path, currently `./PersonalEventBuilderWordResultTemplates.xml`.
- `export.Word.<ModeId>.templateName`: exact XML template name.
- `export.Word.<ModeId>.previewRequired`: required form values for preview. Accepting an order is not required just to display a point.
- `export.Word.<ModeId>.required`: required values for the retained DOCX export.
- `export.fields` and `export.dateFields`: fields to capture and dates to validate. Text dates require `DD.MM.YYYY`; impossible dates and omitted years are rejected.
- `export.Word.aliases`: numbered `name`, `kind`, `source` mappings. Kinds are `value`, `shortDate` and `joinLowerFirst`.
- `export.Word.constants`: explicit declarations for unsupported optional inputs. Kinds are `text`, `emptyText`, `boolean` and `emptyCollection`; text and Boolean constants also require `value`.
- `export.dsl.maxDepth`: include/block recursion limit from 1 to 64.
- `export.dsl.months.<1..12>`: Ukrainian month names for escaped DSL date patterns.

Four personnel lookups read rank genitive, person accusative, position genitive and abbreviated person genitive from the configured SHPO dictionary. Three reporter lookups run when the configured ReportName condition identifies an external reporter. An empty or self reporter skips these lookups and lets XML select the self-report branch. Treatment modes additionally resolve hospital genitive or accusative from the institution dictionary.

Each lookup under `export.Word.lookups.<index>` or `export.Word.<ModeId>.lookups.<index>` has `path`, `sheet`, `range`, `input`, `keyHeader`, `valueHeader`, `output`, `when.field` and `when.pattern`. Its parent `.count` selects how many run. A matched condition requires exactly one nonempty result; duplicate/missing results stop preview. Paths, sheets, ranges and headers occur only in configuration. Open sources use the existing Excel query backend; closed sources use ADO.

## Mode mappings

| Current mode | XML template |
| --- | --- |
| ArrivalTreatment | FromHospital |
| ArrivalAnnualLeave | FromAnnualMainVacation |
| ArrivalMedicalLeave | FromVacationForCuring |
| ArrivalFamilyLeave | FromVacationForFamily |
| DepartureTreatment | ToHospital |
| DepartureAnnualLeave | ToAnnualMainVacation |
| DepartureMedicalLeave | ToVacationForCuring |
| DepartureFamilyLeave | ToVacationForFamily |

For one event, `groupHeader` blocks render in `groupBy` order before the point body. This is not whole-order grouping or sorting. The other 26 XML templates remain in the file without mode buttons.

## Supported DSL

The renderer supports scalar tokens (`{Name}`, `{[Alias]}`), shared `#include`, nested `#if/#else/#endif`, `#not/#and/#or`, `#for/#endfor`, `{#newline}` and `{#space}`. Loop items are dictionaries in a Collection, with `item.__first`, `item.__last`, `item.__index` and `item.[Alias]`.

Formatters are `upper`, `lower`, `upperFirstLetter`, `lowerFirstLetter`, `regexreplace`, `dateoffset` and `dateformat`. Quoted pipeline arguments preserve commas and pipes inside regexes. Dates accept VBA patterns or the original escaped tokens `\dd`, `\d`, `\mm`, `\m`, `\yy`, `\yyyy`, `\month`.

Both plain and braced `#^` and `#_` layout markers are processed before inserting data. Final normalization preserves line breaks and nonbreaking spaces. User values are inserted without reparsing them as DSL.

`#let` only accepts internal variable names starting with `__` and the two transferred operations `$ex_Helpers.m_RegexIsMatch` and `$ex_Helpers.fn_IsDateOnOrAfter`. The renderer evaluates those operations itself; it never dispatches arbitrary VBA named by XML. Unknown aliases, includes and operations stop preview with a diagnostic.

## Remaining scope

The current form explicitly supplies empty MetaTables, MetaDocumentTables and MetaTvoTables. Vacation duration, additional days, end date and VLK-specific fields are also explicitly empty because their inputs are not available. These declarations are configured, not automatic fallbacks for missing DSL fields.

Whole-order aggregation, anchored insertion into an existing order, food support calculations, TVO chains, status changes, HTML table rendering and export deletion/undo require separate input fields and business rules.

The XML reader retains `documentFilepath`; the disabled DOCX exporter still uses `export.Word.output`. Its `documentMode` explicitly selects `blank` or `template`; template mode also requires `documentTemplate` and `bookmark`. DOCX export writes a temporary file and publishes it without overwriting an existing output.

The disabled Movement exporter retains configurable append/close profiles, field/identity/join mappings, chronology validation and duplicate checks. It uses a separate Excel instance and refuses an open, locked, protected or 1904-date-system destination. Its path/sheet/table values are placeholders until the user provides a destination.

## Static validation

Run `Scripts/PowerShell/Test-PEBWordTemplates.ps1`. It checks tab-separated configuration, namespace/version, duplicate template names, include references/cycles, group-header order, DSL block nesting, allowed #let functions and alias/lookup coverage for all eight configured modes without opening Office.

Static checks do not establish VBA runtime correctness. Sources have not been imported into the working workbook; VBA compilation, Word automation, ADO lookups and Movement writes have not been executed.