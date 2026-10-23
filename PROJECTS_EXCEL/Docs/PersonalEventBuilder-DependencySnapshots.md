# Dependency snapshots

PersonalEventBuilder uses a session snapshot of saved dependency files. Only the load / refresh button loads or replaces the complete snapshot. Page initialization and module updates do not read dependency tables; order options remain empty until loading succeeds. Requests before loading show the configured instruction to press the button, without falling back to disk reads. External edits and exports do not alter an active generation. Refresh explicitly to adopt newer dependency data.

EventsHistory is excluded. Event save, edit, conflict detection and revision operations continue reading its live Excel table, including unsaved changes. Export destinations also remain live write targets.

## Ownership and query routing

`ex_TableSnapshots` owns the active `obj_TableSnapshotSet` and a registered `obj_ITableSnapshotLoader`. The common table layer depends on that interface, never on a mode-specific implementation. PersonalEventBuilder registers `obj_PEB_DependencyLoader` during page initialization.

`obj_TableQueryService.TryExecute` and `ex_TableDataRuntime.fn_TryQueryLoaded` consult the provider before selecting a disk or Excel backend. Only explicitly registered dependency sources are served by the snapshot. Raw loading bypasses this routing through the optional `useSnapshot=False` argument.

Source identity includes the absolute workbook path, sheet, range and table name. Aliases resolving to the same source share one loaded table. The return-ticket and period-validation ranges use the same configuration reference.

The snapshot owns immutable source tables, equality indexes and prepared objects. Query results are independent projections and can be disposed by callers. Shared templates, lookup projections, prepared orders, absence indexes and the return-ticket provider are borrowed; consumers must not dispose shared instances. Local services retain only generation-aware references. UI form state and event-specific history snapshots are not global dependency caches.

## Refresh transaction

1. Read all sources listed in `snapshots.sources` using the saved-file backend.
2. Build the configured `snapshots.indexes` and `snapshots.profiles` projections.
3. Prepare orders, absence data, return tickets, hospital metadata, reporter TVO positions and TVO chains. Record individual TVO issues separately without discarding otherwise usable dependencies.
4. Load and validate WORD and Movement DSL templates, dictionaries and item ordering resources.
5. Check that input file modification metadata and sizes stayed unchanged during loading.
6. Publish the staged snapshot and advance its generation only after every step succeeds.

A failure leaves the previous generation active. Staged queries use the staged snapshot directly and never mix its tables with the previous generation. Source metadata is checked during refresh, not on each preview or history save.

Page reinitialization, environment switching and explicit runtime cache resets invalidate the generation. Loading must then be requested again using the button. Existing UI services check generation changes before using prepared references.

## Index semantics

Text equality indexes normalize whitespace and use case-insensitive comparison. Every matching physical row is retained, including duplicates. The index supplies a candidate superset; the ordinary query filter still applies normalization, case, conjunctions, ordering and limits. OR groups are not used to narrow candidates unsafely. Duplicate and missing-value decisions remain with the relevant consumer.

Column lookup accepts the dot/hash header aliases exposed differently by Excel and ACE, while continuing to reject ambiguous columns.

Prepared TVO chains are keyed by position code, including vacant posts without a root name or tax ID. Actor resolution still requires an unambiguous person, and cycle detection uses source row identity. Source notes, missing actors, cycles and invalid root positions are retained in `Tvo.ChainErrors` and logged as `DEPENDENCY_TVO_ISSUE`; they do not prevent loading unrelated dependencies. `ex_PEB_Export.fn_TvoChain(positionCode)` rejects affected positions and returns copies for valid chains. Refresh reports the issue count.

## Diagnostics and validation

`Dependencies.Table` reports source loading; `Dependencies.Refresh` reports complete preparation and success. Existing `Lookup.*`, `ReturnTvo.*` and `Preview.*` stages remain available for comparing cold loads and repeated operations.

The Excel status bar reports the current preparation step and total step count, including table reads, indexes, lookup profiles, derived data and templates. This is synchronous stage progress, not an estimate of remaining time. A blocking ACE query does not report intermediate progress, and the loader does not pump arbitrary UI events or expose cancellation.

Run `Scripts/Testing/PEB/validate_preview_cache.py` for snapshot routing, publication, source coverage, ownership and source-structure contracts. History, return-ticket, TVO and DSL validators cover the associated existing rules. These checks do not compile or execute VBA; Excel runtime validation is a separate, explicitly requested step.