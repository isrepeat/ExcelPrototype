# PeopleRegistry

The generator reads the configured RUH workbook and never writes to it. It has four worksheets: a single-person worksheet form, ID history, a separate origin dictionary, and configuration (wsConfig). There is no input table, local person registry, internal person key, separate ID-alias table or journal.

## Files and reload

- Workbook: PROJECTS_EXCEL/workbook/PeopleRegistry/PeopleRegistry.xlsm.
- Sources: PROJECTS_EXCEL/vba/PeopleRegistry/.
- Configuration: PROJECTS_EXCEL/config/PeopleRegistry/wsConfig.txt.
- Import profile: PeopleRegistry in PROJECTS_EXCEL/vba/modules.json.

Ctrl+Alt+R uses the existing WorkbookUpdater. The profile includes the shared ADO executor and its transitive query/data dependencies, Unicode message helper, source reader, lightweight runtime lifecycle, generator module and ThisWorkbook events. No page/control framework is loaded. The updater requires ThisWorkbook::uiPath even though this profile has no UI files. Native preparation and maintenance scripts import the same profile, including class modules.

## Source

Registration validates required form inputs before reading history or opening ADO. History, aliases and configuration are read into arrays once per operation; config lookup and text tokenization do not access Excel per character or per history cell. There is no RUH result cache: every valid operation executes fresh ADO queries. Source.FetchSize configures temporary recordset fetch batching (4096 rows), not cross-operation storage. The shared executor's optional fetchSize parameter defaults to 1 for unchanged callers. New registration appends only the new history row; replacement retains the complete alias update and transactional rollback. Repeated registration of the last saved person avoids rewriting unchanged history. Saving the workbook remains synchronous.

Source.Path identifies the saved file to inspect. Sheet, range and column mappings for the main list and arrivals are configured explicitly. Each operation uses common obj_AdoQueryExecutor with explicit ranges and selected columns. ACE reads the saved file with HDR=YES, IMEX=1 and ReadOnly=True, without launching Excel or resolving ListObject metadata. Connections and recordsets close on success and failure. Source.List.Range is B5:C20000 and Source.Arrivals.Range is C5:F20000: row 5 contains headers, and the configured boundary includes worksheet row 20000. Increase both ranges in configuration if the source grows beyond that boundary. The Microsoft.ACE.OLEDB.12.0 provider must be available to Excel's architecture.

Save RUH before checking: unsaved edits in an open source workbook are not included. All nonempty source IDs are reserved. Duplicate source IDs with different normalized names cannot be resolved automatically. Names provide candidates and never establish identity by themselves. Source/schema errors block generation and clear stale successful output.

## Single history

tbIDHistory on the configured history worksheet starts at A3 and has five columns:

1. Previous/issued ID (Попередній ІД).
2. Current ID (ІД).
3. Full name.
4. Origin.
5. Issue timestamp as immutable sortable ISO text.

A current ID has a self-reference. A historical ID points directly to that current ID. Every ID is unique, every target exists as a current row, and names agree within a group. No chains or generated internal keys are needed.

## Origin abbreviations

The editable tbOriginAliases dictionary sits on the configured Sheet.Aliases worksheet at Layout.AliasesCell (A3). Its two columns contain the full origin name and short code. Native layout updates move the existing table without recreating its data and retain the dark banded style. The generator reloads and validates the dictionary before each operation. Matching ignores case, ordinary whitespace and configured token punctuation, and normalizes common apostrophe variants. Different full-name variants may share one code, but one normalized name cannot have conflicting codes. Unmapped origins retain the existing token-generation behavior. Changing the dictionary affects future automatic IDs; existing IDs are renamed only through an explicit migration, retaining old IDs as direct aliases.

When a temporary ID is replaced by a tax ID, all rows pointing to the temporary ID are updated to point directly to the new tax ID. The new ID receives its own row. Existing IDs and issue timestamps remain reserved and retained. If the old ID exists only in RUH, its history row is added only when saving a proposed replacement.

Normal generation returns existing source IDs without copying source people wholesale into history. History also contains the explicitly authorized import from the TP and TV tables. The 2026-10-21 abbreviation migration replaced 155 current IDs, initially retained their old forms as direct aliases, and updated 246 constant cells in RUH. A subsequent authorized cleanup removed those 155 obsolete rows after confirming no remaining RUH references, leaving 5849 current IDs. Future explicit replacements still retain aliases. Ordinary generation leaves RUH unchanged; transfer future newly issued IDs to RUH manually.

## Operations

The upper form exposes Зареєструвати, bound to SaveIDs. The lower replacement form exposes Замінити ІД, bound to ReplaceID. Both operations share the merged result area. PreviewIDs and ClearResults remain callable maintenance procedures without worksheet buttons. Each action button is vertically centered within its own two-field block.

Form.Cells maps full name, origin and combined result to three cells. Form.ReplaceCells maps the separate previous-ID and new-ID inputs below the registration form. Form.ResultTemplate combines ID, status and diagnostic with line breaks. Form.Headers, Form.Labels and their Replace counterparts define captions and positions. All inputs are ordinary unmerged cells in column D; only the combined result in column J is merged vertically. Native layout updates preserve both forms' inputs.

- Preview checks RUH and proposes an ID in memory without updating history.
- Save rechecks RUH, writes history and saves the workbook. The module remembers the last successfully saved ID together with its normalized name and origin for repeat operations in the same session. This prevents repeated Save clicks from issuing another ID. Clear, reopening or module reload resets that memory.
- Clear empties the form, including inputs, results, status and diagnostic. History is retained.
- Workbook_Open clears all three form values after successful initialization. Module reload initialization alone preserves the cells; history and the abbreviation dictionary are never cleared.

Validation errors are shown in the form. The original history and input identifiers are restored in memory after a save failure; do not use IDs until successful persistence is confirmed. Ordinary outcomes are shown inline without completion dialogs.

For a new person, supply full name and origin. Temporary IDs use surname, remaining name initials, underscore and origin; collision suffixes start at _2 and account for source and history IDs. Comparisons ignore case and normalize ordinary input whitespace. Same-name candidates require explicit confirmation before creating a separate person; a matching name alone does not establish identity.

The result cell is output only. Save rechecks the source and generates the ID again rather than trusting a manually edited result.

To replace an ID, enter a current or historical ID in the lower previous-ID input and the desired identifier in its new-ID input. ReplaceID resolves the person from history or the source, validates both inputs, reads RUH through ADO and rejects identifiers already reserved in history or RUH. Supplying the current ID as the new value is a no-op. Other changes require Unicode Yes/No confirmation with the name, current ID and new ID; No is the default. Only after Yes does the operation flatten all previous aliases to the new ID, append its canonical row and save PeopleRegistry. RUH is never modified. On success the previous-ID input becomes the new current ID and the new-ID input clears. Save failure restores the entered identifiers and uses the shared history rollback. ClearResults clears both forms and the shared output.

## Maintenance and verification

The history sheet has a state-aware lock button. Unlocking requires an explicitly entered password through Excel's Unicode Application.InputBox; Cancel, an empty input or an incorrect password leaves the sheet protected. Locking applies the HISTORY_PASSWORD constant from ex_PeopleRegistry.vba without prompting. Native scripts read that same constant; the password is not stored in wsConfig. The input dialog is ordinary text input, not a masked password field. The caption follows the actual protection state after initialization, manual toggling and history saves. Reopening, reload and subsequent programmatic history writes restore protection. Manual toggling does not save unrelated workbook edits automatically.

The history worksheet is protected against accidental editing. All cells are locked; existing table filters, selection and copying remain available. Manual edits, formatting, row/column insertion/deletion, sorting and table resizing are blocked. This is an accidental-edit guard, not confidential storage. Initialization restores protection with UserInterfaceOnly enabled. Saving temporarily unprotects the sheet because ListObject resizing requires it, then restores protection before saving the workbook and after any rollback. Native preparation and layout maintenance apply the same settings through Set-HistoryProtection.ps1. Source/configuration reload does not leave the sheet editable.

Prepare-Workbook.ps1 creates only a new workbook and refuses to overwrite an existing one. Update-WorkbookLayout.ps1 updates layout, configuration and imported source while preserving existing history and inputs. Set-GeneratorForm.ps1 authors the native worksheet form and preserves the populated legacy input row; migration stops if more than one person is populated, so no extra input row is silently discarded. A complete workbook backup is saved before layout updates. Neither script executes or compiles VBA.

Configuration visibility is preserved on updates. All captions and messages come from configuration; Unicode dialogs use the shared MessageBoxW helper. The configured Calibri font, neutral charcoal backgrounds, black borders and orange action text apply to every worksheet and the form. Buttons have flat gray fills; secondary instructions use muted gray text. Native table styles provide alternating dark gray rows. All theme colors are configured as HEX values.

Workbook structure, source/configuration readback, preserved input data, schema/profile bindings and final bytes are checked statically or through native Excel authoring. Direct read-only ACE queries for the configured ranges matched all selected saved XML values in 6172 list rows and 1455 arrivals rows, including textual IDs; the source file hash remained unchanged. Generation macros, VBA compilation and the full reload cycle have not been executed.

The new-person checkbox is no longer used. If no identifier resolves to an existing person, the operation creates a separate record. Same-name candidates from RUH or history require explicit Yes confirmation through the shared Unicode dialog on both Preview and Save; No cancels without changing history. An existing supplied ID resolves to its current identifier rather than creating a duplicate. Form migrations discard only the obsolete confirmation field and preserve remaining inputs and results.

The configuration owns geometry: Layout.Input.ColumnWidths and RowHeights, per-sheet column widths and default row heights, optional table row auto-fit, form labels and field addresses, field heights, print/view ranges, alignment, zoom and initial selection/scroll position. Theme font-size keys define typography. Layout.ButtonCells anchors registration/replacement to their first fields. Layout.HistoryButtonCell, HistoryButtonOffsetX/Y, HistoryButtonWidth and HistoryButtonHeight define the history lock button independently. Initialization and native Set-WorkbookLayout.ps1 apply the same values; direct manual layout changes are intentionally reset on reload. The lock button is 180 by 34 points at G1 with a 12-point vertical offset, matching the reference placement.