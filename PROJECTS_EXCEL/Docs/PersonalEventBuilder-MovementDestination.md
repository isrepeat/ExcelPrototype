# Movement destination formatting

The Movement exporter formats the configured Destination column (`Куди`) before calculating exported field hashes. The rule applies to both native mapped values and explicit Movement column-template values. Original EventsHistory data, form input, and Word text are not changed.

Each semicolon-separated section is inspected independently. If a section contains one distinct unit code, its place description is replaced by that code. A code consists of four digits with an optional single Latin or Ukrainian/Cyrillic letter prefix; the prefix is preserved and uppercased. Spacing between a letter and its digits is removed. Numeric codes such as `3019` are kept as four digits. Sections without a code retain their text; semicolon spacing is standardized when at least one section has an unambiguous unit code.

Examples:

- `військова частина А7383, м. Київ; м. Полтава` becomes `А7383; м. Полтава`.
- `м. Полтава; частина 3019, м. Київ` becomes `м. Полтава; 3019`.
- `частина А7383; частина 3019` becomes `А7383; 3019`.

Longer digit identifiers and common dotted/slash/hyphen date forms do not match. Repeated mentions of the same unit in one section are reduced to one code. A section containing multiple different codes remains unchanged to avoid selecting an arbitrary unit. Four-digit standalone numeric tokens are treated as unit numbers; ambiguous non-unit numbering should be reviewed in the source destination.

The pattern is stored in `PersonalEventBuilder::export.Movement.destination.unitPattern` in both wsConfig files. Capture group 1 is the preceding boundary and capture group 2 is the unit code; the code relies on that contract. A reusable VBScript RegExp instance refreshes when the configured pattern changes and is released with the Movement registry.

Destination comparisons for existing closing links and source-period selection use the same formatter on both sides. This allows legacy descriptions such as `Медична рота А7383` to match normalized destination output while retaining existing field-conflict and write-verification checks.

Run `WorkbookUpdater/Test-PEBMovementDestination.ps1` for representative section examples and native VBScript RegExp compatibility. It does not execute or compile VBA. The change takes effect after the user updates the workbook/configuration and reloads dependencies; exports still use the existing preview/conflict policy for changes to previously exported rows.