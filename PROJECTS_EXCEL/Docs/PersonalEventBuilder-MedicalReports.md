# Medical report validation

The Validation button reads saved medical reports without opening an Excel workbook or changing dependency snapshots. It requires an accepted order. The scope is the inclusive order-date window of three days before and after the order. Missing daily files and monthly folders are skipped. Hospitalization dates are filtered by the accepted order date. Archive-wide name searches are outside this validation workflow.

Settings live under `PersonalEventBuilder::validation.MedicalReports` in both configuration files:

- `root`: report archive root.
- `folder`: date-based folder path relative to the root; default `{yyyy}\{yyyy.mm}`.
- `patterns`: pipe-separated VBA Like filename patterns. Supported date tokens are `{dd.mm.yyyy}` and `{yyyymmdd}`. Use date tokens to select the report day, rather than the filesystem modification date. Temporary `~$` files are excluded.
- `windowDays`: inclusive daily search radius, default 3.
- `source.sheet` / `source.range`: sheet and range with its header row, default `Аркуш1` / `A3:Q20000`, including the tax ID in column Q.
- `fields` / `headers`: source projection and external headers. The projection must include `PersonName`, `PersonId` and `Dates`.
- `display` / `text.Columns`: result fields and independently configurable captions. `Sources` contains report paths.
- `text`: controls, completion messages and diagnostics.

Only the first date in the hospitalization/transfer cell is relevant. A later matching transfer date cannot include a row. Text must start with a valid `dd.mm.yyyy` date after whitespace; native Excel date values and serial numbers are also accepted. Original date lists, institutions and document fields remain visible.

Identical business values are merged across files, retaining source paths. Changed units, ranks, date lists, institutions or document values remain separate rows. EJOS validation is not performed.

Active root hospital events in live EventsHistory are matched by tax ID, event date and the accepted order number/date. Native Movement event plans identify hospital departures and transitions; arrival events, other event types, child rows and cancelled records cannot mark a row as recorded. Recorded rows are green unless an identity warning already requires orange review. The Notes column still confirms the recorded event in either case.

For people without a matching history event, the latest Movement period starting on or before the order date is inspected by tax ID. An unclosed hospital-treatment or medical-leave period marks the row red and adds its departure date; medical leave also includes the planned return date when available. This preserves the conflict conditions from PrototypeNew while using the current Movement period ordering. Closed periods and future departures do not cause red warnings. Identity diagnostics remain visible when a red conflict takes precedence.

Each extracted row is checked against the loaded personnel snapshot using the tax ID. The report name is never used to silently select another tax ID. The result retains the original name and ID, displays the personnel name for that ID, and explains its identity status. Missing, malformed, unknown or ambiguous IDs and name mismatches are highlighted in orange. A missing report name remains visible as a mismatch rather than aborting the report. Names are compared after case, whitespace and typographic-apostrophe normalization; spelling differences are flagged without automatic correction. IDs must contain ten digits after whitespace and an optional Excel text prefix are removed; zeros are never invented.

The generic `obj_IUiRowTags` interface exposes optional per-row style tags to table controls. Lookup results opt in through the optional `rowTagField` initialization argument; existing callers keep their original presentation.

Validation failures are displayed in the action status without a modal message. An empty order dropdown invokes the optional select `onEmpty` command. PersonalEventBuilder reports that dependencies must be loaded when no snapshot exists, or that no orders were found when a snapshot is already available. Clicking the arrow does not implicitly load or replace dependency files.

The result is hidden when another order is accepted. Read or schema errors stop the operation and hide previous results. Files are read through the existing read-only ACE query backend; ACE header substitutions for dots and line breaks are normalized locally. No VBA compilation or workbook import is part of source validation.

## Performance diagnostics

Validation captures one nested, invocation-local configuration snapshot and releases it on success or failure. Tax ID checks use a ten-digit Like pattern without COM allocation. One invocation-local date expression is reused across every report. Personnel IDs are scanned from the loaded two-column array; name sets are built only for IDs in the extracted reports, retaining every distinct matching reference name. Movement arrays are scanned only for relevant people not already recorded in history, and that query is skipped when no such IDs remain. Report files and live history are still reread on each validation; no stale result cache is introduced.

The existing diagnostic logger emits `PERF` entries with `Stage=MedicalValidation.*` and `ElapsedMs`:

- `Config`, `Discover`: configuration capture and report discovery, including file count.
- `File.Query`, `File.Filter`, `File.Total`: each report query, date filtering/merging and total read; query/filter entries identify the report filename.
- `Reports`: aggregate report loading with file and extracted-row counts.
- `Personnel.Query`, `Personnel.Index`, `Personnel.Compare`, `Identity`: loaded personnel projection, restricted name-index construction, identity diagnostics and aggregate identity work.
- `State.People`, `History.Read`, `History.Match`: requested IDs and live history loading/matching.
- `Movement.Query`, `Movement.Index`, `Movement.Skip`, `Movement`: loaded Movement projection, relevant latest-period selection, empty-request skip and aggregate Movement work.
- `State.Notes`, `State.Total`, `State`: notes/tags and aggregate state checks.
- `Result`, `Load`: UI result construction and complete data preparation, excluding rendering.
- `Render`, `Total`: publication/rendering and the complete validation handler. `Total` includes data preparation and rendering; ordinary handled failures log `Success=False`.

Aggregate stages contain their child stages; do not sum every entry. Compare `Total` across runs and use the child stages to locate remaining costs. Runtime improvements must be measured after the updated sources are explicitly imported and validation is invoked; static source checks do not execute VBA.