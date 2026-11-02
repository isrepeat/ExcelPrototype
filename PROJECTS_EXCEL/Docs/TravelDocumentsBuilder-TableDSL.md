# TravelDocumentsBuilder Word table DSL

The vacation ticket's first table is generated from `workbook/TravelDocumentsBuilder/VacationTicketTable.xml`. The Word scaffold is `workbook/TravelDocumentsBuilder/Шаблони/Відпустка (DSL).docx`. The original DOCX is retained in `Шаблони/old`. The validity notice and conditional `VacationAbroadText` row belong to the generated `Ticket` table; the scaffold preserves the page break, reverse side, headers and footers.

## Word insertion boundaries

Place the following markers in two separate paragraphs in the main document body:

```text
{\exportTable:Ticket_Begin}
{\exportTable:Ticket_End}
```

`Ticket` is the XML template name. Each marker must occur exactly once, outside an existing table. The begin marker must precede the end marker. The generator replaces both marker paragraphs and everything between them. It leaves other document content in place. Markers may be split across Word runs.

Signatory placeholders in the footer are still replaced through the existing bookmark mechanism. Table text uses `{[Field]}` syntax. Both the validity notice and `{[VacationAbroadText]}` are inside `Ticket`; an empty `VacationAbroadText` omits its entire row.

## Configuration

All runtime settings are in `config/TravelDocumentsBuilder/wsConfig.txt` and must be loaded into the workbook's `wsConfig` table by the normal update workflow.

| Key after `TravelDocumentsBuilder::` | Meaning |
| --- | --- |
| `word.tables.mode` | `table` enables the DSL. `markers` selects the original scalar-only renderer. |
| `generation.tablePaths.<DocumentId>` | XML catalog path for the saved event's document type, relative to the application workbook. |
| `generation.tableNames.<DocumentId>` | Required template name: `Ticket` or `BusinessTrip`. |
| `word.tables.maxDepth` | Maximum structure, inheritance and text nesting depth. |
| `word.tables.maxRows` | Maximum expanded row count. |
| `word.tables.maxColumns` | Maximum grid column count, up to Word's limit of 63. |
| `word.tables.images.Emblem` | PNG asset path, relative to the application workbook. |
| `word.tables.text.Failed` | Short user-facing failure message; technical details go to the common log. |
| `generation.templates.VacationTicket` | Word scaffold used for vacation tickets. |

The checked-in configuration selects the new table renderer and scaffold. To use the original renderer, explicitly change both `word.tables.mode` to `markers` and `generation.templates.VacationTicket` back to `Шаблони\Відпустка (Шаблон).docx`. There is no automatic fallback when a catalog, asset, style, field or marker is missing.

Technical DSL validation errors use `Err.Raise` and are returned to the export boundary, following the PEB renderer pattern. The exporter records the original number, source and description through `ex_Core.fn_Diagnostic_WriteError`, with the stage, operation ID and document path. Log output, timing, buffering and flushing use the existing `ex_Core.fn_Diagnostic_*` mechanism shared with PEB. Only the short message shown to the user comes from `word.tables.text.Failed`; internal validation text is not shown in the status block. Stateless standard modules expose `fn_Module_Dispose` for the workbook lifecycle contract.

## Structure and styles

```xml
<wordTableTemplates xmlns="urn:excelprototype:profiles"
                    version="1">
  <styles>
    <tableStyle name="Ticket"
                width="170mm"
                alignment="left"
                layout="fixed"
                borders="none"
                cellPadding="0pt"/>
    <cellStyle name="Body"
               fontName="Times New Roman"
               fontSize="12"
               horizontal="left"
               vertical="center"
               spaceBefore="0pt"
               spaceAfter="0pt"/>
    <cellStyle name="Heading"
               basedOn="Body"
               fontBold="true"
               horizontal="center"/>
  </styles>
  <template name="Ticket">
    <table style="Ticket">
      <columns>
        <column name="Label"
                width="45mm"/>
        <column name="Value"
                width="125mm"/>
      </columns>
      <row>
        <cell columnSpan="2"
              style="Heading">
          <text>ВІДПУСКНИЙ КВИТОК № {[TicketNum]}</text>
        </cell>
      </row>
      <row>
        <cell style="Body">
          <text>Виданий</text>
        </cell>
        <cell style="Body">
          <text>{[PersonalLine]}</text>
        </cell>
      </row>
    </table>
  </template>
</wordTableTemplates>
```

Lengths accept `mm`, `cm` and `pt` with a dot as decimal separator. Font size is a number of points. Column widths must add up to the table width. The original ticket uses an 11-column grid to reproduce its irregular cell boundaries.

`rowSpan` and `columnSpan` define rectangular merged cells. Cells are placed left to right in the next unoccupied grid position. Continuations of a vertically merged cell are omitted from later DSL rows. Overlaps, gaps and out-of-bounds spans are rejected.

Rows accept `header="true"` for consecutive repeated header rows, `allowBreakAcrossPages="true|false"`, and `height` with `heightRule="atLeast|exact"`. Omit height for automatic sizing, or explicitly use `heightRule="auto"`. Use exact heights only for intentional empty spacers. Cells cannot merge across a repeated-header boundary.

Table styles support width, alignment (`left`, `center`, `right`), fixed layout, borders (`none`, `single`) and cell padding. `cellPadding` sets a uniform base; `cellPaddingTop`, `cellPaddingLeft`, `cellPaddingBottom` and `cellPaddingRight` override individual sides. The ticket retains the original 5.4-point left/right margins and zero top/bottom margins. Cell styles support font name/size, bold/italic, horizontal and vertical alignment, paragraph spacing, minimum line spacing, and `borderBottom="single|none"`. `basedOn` inherits a style of the same kind; cycles are rejected. Paragraph styles use the same cell-style catalog and override the containing cell's text properties.

Cell borders are emitted explicitly on all four sides, matching Word's No Border operation rather than relying only on inherited table borders. `none` becomes `w:val="nil"`; `borderBottom="single"` retains the printed underline. Word's editing grid remains a separate, nonprinting view setting.

A cell contains either one `text` or a sequence of `paragraph` elements. An empty cell is valid. A paragraph contains one `text` or one `image`:

```xml
<paragraph style="Heading">
  <image source="Emblem"
         width="7mm"
         height="9mm"/>
</paragraph>
```

Images are embedded into the resulting DOCX rather than linked to local files. The current application adapter supplies the `Emblem` asset. The generic renderer accepts an explicit dictionary of additional PNG assets from a caller.

## Text DSL and data context

Text uses the existing PEB syntax: `{[Field]}`, formatter pipelines, `{#if ...}`, `{#for item in Collection}`, `{#include block.Name}`, `#let`, and layout tokens such as `{#newline}`. `sharedTemplates` and `resources` follow the PEB catalog structure. The independent `obj_WordTableText` engine receives nesting limits and month names explicitly and has no PEB configuration dependency.

The application prepares all saved event fields plus these existing Word aliases: `TicketDate`, `TicketNum`, `PersonalLine`, `VacationKind`, `VacationPlace`, `VacationDuration`, `DateFrom`, `DateTo`, `PersonalInitials`, `DateArrival`, `VacationAbroadText`, `FooterSpacer`, `SignatoryPosition`, `SignatoryRank`, `SignatoryName`. Existing alias values take precedence over same-named saved event fields. Dates in the Word aliases are already display strings; use saved date fields such as `DepartureDate` for date formatters.

`TvoChain` contains the saved TVO rows for the event, with fields such as `Rank`, `FullName` and `PositionCode`. Unlike PEB, collections such as `BasisDocuments` must be prepared explicitly by the caller; the shared renderer does not perform profile-specific lookups.

## Dynamic rows

Structural `if` and `forEach` produce table rows. A text `#for` only produces content inside a cell.

```xml
<if test="[ReportNumber]">
  <row>
    <cell style="Body">
      <text>Рапорт</text>
    </cell>
    <cell style="Body">
      <text>№ {[ReportNumber]}</text>
    </cell>
  </row>
</if>
<forEach itemsSource="TvoChain"
         item="tvo">
  <row>
    <cell style="Body">
      <text>{tvo.__index}</text>
    </cell>
    <cell style="Body">
      <text>{tvo.[Rank]} {tvo.[FullName]}</text>
    </cell>
  </row>
</forEach>
```

A loop may contain several rows or nested structural blocks. Each item must be a dictionary in a VBA `Collection`. Loop contexts expose `alias.__index` (one-based), `alias.__first` and `alias.__last`. An empty collection produces no repeated rows; an entirely empty table is rejected. Missing collections and fields cause explicit errors. Structural conditions use the same expression parser as text conditions.

The provided ticket catalog reproduces the original table without adding a TVO section to the printed form. Dynamic rows are available for future template changes.

## Ordered vacation places

The place field is a draft. The plus button appends a validated draft to the ordered place chain and clears the draft fields. The minus button moves the final chain item back into the draft for editing; it refuses to overwrite a nonempty draft. Saving also includes a nonempty validated draft, so the final place does not require a separate plus click. The chain limit is configured by `placeChain.maxItems`. Medical-company modes clear the chain and retain their fixed destination.

A single place is stored directly on the `Event` row in `tbTickets`, with its `Destination` and the structured fields configured by `placeChain.fields`; no `Place` child is written. Routes with two or more places use one `RowType=Place` child per place. The first child references the event through `ParentRowId`; each next child references the previous place. Revision snapshots and rollback include these rows, and editing preserves child IDs while the route remains multiple. Place links cannot cross into TVO or donation branches. The event's `Destination` joins ordered place addresses with a semicolon and a space; scalar place fields on the event retain the first place for compatibility. Saving a route reduced to one place removes its old Place children and keeps the surviving place on Event.

`VacationAbroad` switches the draft lookup between `lookup.Places` (`tbPlacesUkraine`) and `lookup.PlacesAbroad` (`tbPlacesAbroad`). Both named tables are loaded into the same dependency snapshot. Switching clears the draft metadata and replaces candidates and headers without clearing the ordered chain. Overseas addresses are assembled from the seven structured source columns; the workbook's `Повна адреса` column is not queried. Both language variants include a nonempty region unless it repeats the place name. Country-only entries do not repeat the country. Manual overseas input requires country and place names in both languages; region and type are optional. The event is marked abroad when any saved chain item is abroad.

Add `PlaceAbroad`, `PlaceCountry`, `PlaceCountryEn`, `PlaceRegionEn` and `PlaceNameEn` to `tbTickets`, and the matching `Snapshot.` columns to `tbTicketRevisions`. These fields supplement `PlaceRegion`, `PlaceType`, `PlaceName` and `PlaceDistrict`; foreign entries do not require a Ukrainian district.

History loading restores multiple places into the ordered chain and clears the draft to avoid duplication on the next save. A single place, whether stored on Event or as an older Place child, is restored into the regular destination field and its structured draft fields; the chain block stays hidden. Existing one-child routes collapse to Event-only storage on their next save. Older events without structured place fields retain their draft for completion. Word reconstructs a single in-memory place from Event's structured fields and exposes the same ordered `PlaceChain` collection for either storage form, using the saved aggregate destination. A reconstructed item has no child ID and never reuses Event's identity. Export does not look up current place dictionaries.

## Business-trip scaffold

`Шаблони/Відрядження (DSL).docx` is the business-trip scaffold. Its unique top-level markers are `{\exportTable:BusinessTrip_Begin}` and `{\exportTable:BusinessTrip_End}`. The separate `BusinessTripTable.xml` catalog contains the matching `BusinessTrip` template. It preserves the original six-column front grid without the duplicate receiving-unit rows and appendix 29, using the same font styles, emblem and first-page signature footer as vacation tickets. The validity notice and optional weapon line follow the original grid inside the DSL table. The reverse-side form is retained from `Відрядження (Шаблон).docx`; an explicit page break before its continuation heading keeps both reverse headings together on page two. The original DOCX is unchanged.

The front table requires `TicketDate`, `TicketNum`, `PersonalLine`, `Position`, `TravelPlace`, `Duration`, `Basis` and `WeaponLine`. Basis is supplied by the existing form input; the sample course purpose and its separate form field are removed. `WeaponLine` is the complete optional line supplied by the form: an empty string omits the entire row; a nonempty string is printed verbatim without a fixed prefix or punctuation. The first-page footer uses `{SignatoryPosition}`, `{SignatoryRank}` and `{SignatoryName}`. The existing destination input and ordered place chain supply TravelPlace for towns, military units and institutions; there is no separate receiving-unit input or duplicate destination row.

Both document types are enabled. The form shows the optional `WeaponLine` (complete sentence) for business trips. The existing `Purpose` input is the optional basis, stored as `Basis`, and has no separate purpose input. Main-form captions have no required-field stars. `VacationDays` is displayed as trip days; road days are added to the inclusive duration, and trip days must be positive. Donation days are zero and VLK fields are cleared in this mode. Vacation controls retain their previous behavior.

The persisted `VacationKind` field is also the event-kind discriminator: configured vacation-kind inputs identify `VacationTicket`; the exact configured `generation.businessTrip.kind` identifies `BusinessTripCertificate`. Unknown kinds fail explicitly. This preserves the identity of existing vacation events without silently assigning an unknown event to a document type. History loading derives the selector from the saved event kind and restores the weapon field. Type changes during editing are captured by ordinary history revisions. The trip's printed end date is one day before `PlannedReturnDate`: a one-day trip beginning 01.11.2026 prints 01.11.2026 through 01.11.2026 and returns 02.11.2026. The existing vacation calculation is unchanged.

Before using the updated sources, append `WeaponLine` to `tbTickets`, and `Snapshot.WeaponLine` to `tbTicketRevisions`. Keep the existing rows and values; the added cells for historical vacation records are blank. These fields are included in the configured form mappings, content snapshots and rollback schema. The source change does not alter a working workbook or import VBA. Both tables must match the new configuration before either saving or exporting; a schema mismatch stops the operation before any write.

Order export renders both vacation and business-trip events in one transaction. Each active root selects its own configured DOCX and XML by saved kind, independently of the current form selector. All required DOCX sources are copied to staging before rendering, and every source stamp is rechecked before publication. XML catalogs are cached once per document type for the export; scalar bookmark anchors are reset when the frozen DOCX changes. Output names use `Вк` or `Пв` as configured. Existing order-wide rollback, cancellation cleanup and metadata updates apply to both types. Business trips require table mode; the old scalar-only renderer is not supported for them.

`prepare_business_trip_template.py --emit-patch` prints a reviewable patch for the catalog entry. `--build-scaffold` recreates only the new DOCX, preserving every unaffected source package part. Apply catalog text changes through `apply_patch`. Run `test_business_trip_template.py` for offline grid, source-package, footer and pagination-boundary checks. `word_table_preview.preview` accepts an optional template name for synthetic layout checks of either catalog entry.

## Saved grammatical forms

Saving an event resolves grammatical forms from the loaded SHPO dictionaries, using the same source contracts as PersonalEventBuilder: rank by `Звання`, person by `ІПН`, and position by `Код`. The sources, headers and mappings are configured under `declensions` in wsConfig. Missing, ambiguous or empty results stop saving with a field-specific diagnostic; nominative text is not silently substituted.

TicketsHistory requires six additional columns in `tbTickets`: `RankGenitive`, `RankDative`, `FIOAccusative`, `FIODative`, `FIOInitialsGenitive`, and `PositionGenitive`. Add the corresponding `Snapshot.<field>` columns to `tbTicketRevisions`. They are independent of the four VLK columns. Saving records these forms in the event and revision snapshots. Existing events must be opened and saved again to populate their forms before Word export.

Word export uses the saved forms without consulting current dictionaries. `PersonalLine` combines `RankDative` and `FIODative`; all six saved fields are also available as table DSL placeholders. The existing nominative `PersonalInitials` remains appropriate for the return acknowledgement.

## Rendering and verification

The catalog, shared text blocks and PNG assets are loaded once per Word export, on first use. Saving an event does not load the Word table catalog. For each document, the renderer expands conditions/loops, renders cell text, validates styles and the complete grid, then constructs a Flat OPC package. Word receives the table through one `Range.InsertXML` call, avoiding repeated row insertion, cell merging and per-cell COM formatting.

The original scalar-marker bookmarks are indexed once in the frozen template and reused across the order, including for the table renderer. The frozen template retains the table boundary markers; each output document gets its own generated table.

The existing order-wide staging, rollback and history synchronization remain responsible for publishing files. No working workbook is changed by source validation.

Run `PROJECTS_EXCEL/WorkbookUpdater/Test-TravelDocumentsSources.py` and `Scripts/Testing/TravelDocuments/test_word_table_dsl.py` with the bundled Python runtime. The latter checks the original ticket grid, dynamic expansion cases, invalid spans, styles, insertion markers, preservation of other DOCX parts, and text-engine parity. These are static/offline checks, not VBA execution.

`check_recent_contracts.py --base ee2a848^` audits the two UI/DSL commits together with working changes. It checks module lifecycle entry points, namespace boundaries, class lifecycle/API sections, changed procedure formatting, the technical/user message boundary, XML syntax and final bytes. A different base revision can be supplied for a later change set.

`prepare_table_example.py --emit-patch` prints a reviewable catalog patch from the original DOCX. `--build-scaffold` regenerates only the new scaffold and extracted PNG. All text edits must still be applied through `apply_patch`.

`word_table_preview.py OUTPUT.docx` creates a synthetic-data layout preview without executing VBA. Its text evaluator intentionally supports only the simple substitutions used by the provided ticket; it is not a replacement for the VBA DSL. `--flat-opc PREVIEW.docx` emits the first table as a Flat OPC package for a direct Word API smoke check.