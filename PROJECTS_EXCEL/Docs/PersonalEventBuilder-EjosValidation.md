# EJOS validation

The existing Validation button now loads two independent results. The EJOS table is placed immediately before the medical report table. A failed source is hidden and its diagnostic is displayed in the action status; a successful companion result is still published. No validation messages use modal dialogs.

The EJOS source is the registered `lookup.Absence.source` dependency snapshot. Validation does not reopen or refresh that workbook. Use the existing load/refresh action to adopt source changes. The current source range in shared configuration remains authoritative.

Settings live under `PersonalEventBuilder::validation.Ejos` in PROD and TEST. `columns` contains the sixteen source headers in order: `65`, `66`, `67`, `68`, `69`, `70`, `71 (А)`, `71 (Б)`, `72`, `73`, `74`, `75 (А)`, `75 (Б)`, `76`, `Стовпець1`, `Стовпець2`. The display prepends a diagnostic `ValidationNote` column with the configured `text.NoteHeader` caption, for seventeen columns in total. `fields` maps source roles, `kinds` maps absence descriptions to native Movement target-status identifiers, `cancelledToken` controls cancellation detection, and `text` contains all diagnostics. Source values are preserved without deduplication.

Only rows whose `71 (А)` date equals the accepted order date are displayed. Date cells accept native Excel dates, serial numbers or text starting with a valid `dd.mm.yyyy` date. A nonempty `71 (Б)` must agree with the selected order number; an empty number is permitted because regular annual-leave rows may omit it.

Identity is resolved against loaded personnel by normalized full name. A unique tax ID is required; column `67` is a position code, not a tax ID. The code may disambiguate identical names only when their name/code pair resolves to exactly one tax ID. A missing/different code is flagged even when the full name resolves uniquely. A code pointing to another name cannot silently identify or mark a person as recorded. Case, whitespace and typographic apostrophes are normalized; spelling errors and unresolved ambiguous names require review.

The live EventsHistory API checks active root events for the resolved tax ID, selected order number/date, departure date from `70`, and configured target status corresponding to `68`. An arrival, cancelled event, different departure date or different absence type cannot produce a green row. The query permits an empty history result and retains separate source rows.

Color precedence:

1. Red: the configured cancellation token appears in `65` or `69`. A matching active history event adds an explicit cancelled-but-recorded warning.
2. Orange: unresolved/ambiguous identity, name/code inconsistency, unknown absence type, missing/invalid departure, invalid or earlier return, conflicting order number, missing destination, duplicate source departure, or a mismatching current-order history departure.
3. Green: a matching active history event with no source/identity warnings.
4. Normal: no matching recorded event and no detected inconsistency.

Every row displays its explanations directly in the first notes column, including orange inconsistencies, cancellation warnings and recorded/not-recorded state. Row navigation has no selection command or status update. All sixteen source columns remain after the note. Each table can be closed independently; switching the accepted order hides both previous results.

`EjosValidation.Source`, `.Filter`, `.Identity`, `.History`, `.Compare`, `.Result` and `.Load` report elapsed milliseconds through the existing diagnostic logger. For compatibility, the existing `MedicalValidation.Total` handler timing now covers both validations and their combined rendering. Child timings must not be added to their containing totals.

Run `WorkbookUpdater/Test-PEBEjosSources.ps1` and the existing medical-source validator for static contracts. These checks do not import sources, compile VBA or execute workbook macros.