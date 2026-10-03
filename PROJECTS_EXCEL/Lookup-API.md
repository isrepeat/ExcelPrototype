# Lookup and selectable tables

The PersonalEventBuilder page searches the `ОС` worksheet in `workbook/Dependencies/ШПО.xlsx`. Enter at least two characters of a name or tax ID and commit the cell edit with Enter. Click any candidate data cell to apply the record to the personnel form. Double-click also accepts a record, including when the same cell is already active. Headers are not selectable; selection spanning multiple rows is ignored.

The page displays name, tax ID, rank, position and unit. The position code is retained in the record without being displayed. Search is case-insensitive, normalizes outer spaces and NBSP, and treats wildcard characters literally. Results are ordered by name and tax ID. Up to 50 records are displayed, with a message when more matches exist. Editing the query clears the previously selected identity and dependent form fields.

## Profile

All lookup settings are stored in `config/PersonalEventBuilder/wsConfig.txt` and read from the workbook's `wsConfig` / `tbConfig`. Initialize the profile with the configuration prefix `PersonalEventBuilder::lookup.Personnel`, not a file path. Source paths are relative to `ThisWorkbook.Path`.

Keys under the prefix include `key`, `minChars`, `maxResults`, `source.path`, `source.sheet`, and `source.range` (or `source.table`). `range=auto` discovers the last populated row and the last header in row 1. Each section has an explicit `count` and numbered entries:

- `columns.1.name`, `columns.1.header`, optional `columns.1.captionKey` referencing a text key in wsConfig.
- `display.count=0` uses source columns with configured captions. A nonzero count uses `display.1.name` and `display.1.caption`, allowing extension fields in the projection.
- `search.1.name`, `search.1.match`.
- `apply.1.field`, `apply.1.target`, `apply.1.policy`, `apply.1.required`.

Every required key must exist. The internal DOM is only an in-memory parsing representation built from these keys; it is not read from or saved to an XML configuration file. Status text, labels and captions also come from wsConfig. Import the updated wsConfig text along with the sources when workbook updates are authorized.

Display fields absent from a record appear blank. This allows extensions to contribute columns independently of the main source. Full records remain available when the display hides fields or changes their order. Required application fields must exist; optional extension maps can use `required="false"`.

Application targets use `Source.Path`, rather than worksheet addresses. Policies are `overwrite`, `fillIfEmpty` and `ignoreEmpty`. All targets must already exist as scalar dictionary entries in BindingContext. `TryApplyValues` validates the entire batch before changing data and sends notifications after all values have been stored. This is a coordinated data update, not a transaction around arbitrary event subscribers.

## Data and ownership

| Component | Responsibility |
| --- | --- |
| `obj_LookupProfile` | Profile, source, search fields, display and application mappings |
| `obj_LookupService` | Cached source query, search, extension, append and related details |
| `obj_LookupRecord` | Entity key, named scalar fields and named detail collections |
| `obj_LookupResult` | Display projection and row-to-record mapping |
| `obj_IUiSelectionSource` | `GetItemAt(rowIndex)` returns the object behind a displayed row |
| `obj_ILookupRowSelector` | Custom resolution of multiple extension matches |

Concrete classes require explicit `Initialize`, reject repeated initialization, and have idempotent `Dispose`. The controller owns its profile and service. Results retain record references without disposing records that may still be selected elsewhere. The service does not own the profile passed to it. Record field dictionaries and result record collections are returned as copies; contained records are shared objects.

The first search reads the source through `QueryExcel`, retaining Excel values without ACE inference. The resulting in-memory snapshot is reused until `InvalidateCache` or disposal. The `Оновити ШПО` button invalidates the snapshot and searches again. Reopening the page also reloads it. If a source is already open in the current Excel instance, unsaved values are read; otherwise an owned hidden read-only instance is used and closed. No source workbook is saved or edited. Duplicate nonempty keys fail with a diagnostic. Empty-key rows are not candidates.

## Composition

`TryExtend(candidates, extension, joinField, duplicatePolicy, overwrite, output, diagnostic, selector)` performs an indexed left join by the named field. Policies `error`, `first` and `last` are explicit; `first` and `last` refer to extension collection order. With `custom`, supply `obj_ILookupRowSelector` to implement date/status rules. Index 0 means no extension is accepted. Negative or out-of-range indexes fail. The output contains copies, preserving existing related details; a failed merge does not modify the original records.

`TryAppend(candidates, incoming, duplicatePolicy, output, diagnostic)` adds independent candidates by record key, preserving first-seen order. Policies are `error`, `keep` and `replace`. Incoming records must use the same semantic schema and identity convention. No source-specific logic is built into the engine.

`TryAttachDetails(candidates, details, joinField, detailName, diagnostic)` associates all matching detail records with each candidate. For example, multiple documents remain a related collection rather than duplicating the person in the candidate table. Consumers can access it with `record.GetDetail("Documents")`; a detail renderer or page-specific command can decide how to display or use it. The initial personnel page does not load additional document sources or render nested detail rows.

`TryProject(records, result, diagnostic)` rebuilds the table projection after composition. It does not truncate a prepared collection; callers control its size. Secondary data can be prepared by another configured lookup service or another source-specific adapter, using the same record contract.

## Table selection

```xml
<controls:table name="Candidates"
                itemsSource="{Binding Path=Data.PersonnelCandidates}"
                selectedItem="{Binding Path=Data.SelectedPersonnel}"
                onSelect="{Binding Path=Commands.SelectPersonnel}"
                selectedStyle="candidateSelected"/>
```

`selectedItem` and `onSelect` are optional. A selectable source must implement both `obj_IUiTableSource` and `obj_IUiSelectionSource`. Ordinary read-only tables continue to accept their existing sources. Selection routes cover body ranges, excluding titles and headers, and are replaced on render and removed on disposal.

The table writes the actual record to the object binding and invokes `obj_UiCommand.ExecuteWithPayload`. Existing parameterless commands still use `Execute`. Selection uses the row mapping from the data model, not the value or column of the clicked cell. `Workbook_SheetSelectionChange` and `Workbook_SheetBeforeDoubleClick` dispatch through `ex_UiBindings.fn_HandleSelection`, which restores Excel event state and flushes pending layout. Selection output bindings do not themselves trigger table re-layout.

## Verification and deployment

`WorkbookUpdater/Test-LookupSources.py` performs read-only checks of source declarations, interface members, helper visibility, profile/source headers, mappings, UI commands, object names and final file bytes. It uses Python's standard library and never opens Excel or invokes VBA. It also verifies unique nonempty keys for named personnel rows in the current source without printing personal records.

`WorkbookUpdater/Test-UiElements.ps1` has been adapted to initialize the new page bindings and check the updated button layout. It imports sources and runs VBA, so do not run it when compilation/execution is prohibited.

The sources have not been recompiled or executed after the user requested stopping VBA compilation. The working workbook has not been updated. Import the project through the existing updater only when runtime validation and workbook updates are explicitly authorized. Static checks do not establish VBA runtime correctness.

Private helper procedures are called directly. Qualifying them through `Me` caused the observed VBA member-resolution error; making them public just to support that syntax would unnecessarily widen the class API.