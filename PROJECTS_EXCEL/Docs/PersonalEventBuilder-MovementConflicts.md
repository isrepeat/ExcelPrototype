# Movement basis conflicts

## Current export policy

Both profiles set export.Movement.manualChanges.policy to Overwrite. Linked events overwrite their mapped export fields from EventsHistory even after manual target edits. Previous accepted exceptions and pending choices are ignored, and the conflict table stays hidden. Formulas, link matching and final period validation still guard the export. Only fields owned by the linked event are overwritten; other Movement columns remain unchanged. Field snapshots continue to record the generated/applied state. The Resolve policy and its UI remain implemented for later use; the workflow below describes that inactive policy.

The export compares actual target fields with the last successful per-field export state. Only DepartureBasis and ArrivalBasis support explicit conflict choices. Changes to semantic fields, formulas, event links, or a link hash inconsistent with the stored state remain blocking errors. EventsHistory and Word data are never rewritten by a conflict choice.

The third UI table sits below HistoryEvents. Its columns identify the event/side/field, the current Movement value and the newly generated value. Ctrl+Enter on either value selects it; the first column is not selectable. Each row keeps its own selection. The status displays the resolved count. Pending choices are held by the page controller until export succeeds or the page is disposed. Retrying after an event, target, current value, generated value or link hash changes invalidates the corresponding selection.

Selecting Movement preserves the target text as an export exception. Selecting EventsHistory replaces the basis with newly generated text. A later export with unchanged generated and applied values preserves an accepted exception. A changed generated basis is exported normally when the target has not been changed again; a subsequent manual target change requires a fresh choice.

ExportFieldState/tbExportFieldState and ExportConflicts/tbExportConflicts must exist before export. The application validates their exact configured columns and never creates workbook structure. New state rows are appended only when revision, generated or applied values change; the last row for an event/target/side/field is authoritative. GeneratedValue and AppliedValue in the state table use the existing value-token encoding (s: text, n: invariant number) to preserve type and avoid locale-dependent comparisons. Conflict journal values are human-readable text.

Field snapshots and selected decisions are appended within HistoryRevisions.fn_RecordExport, before the same workbook save as ExportJournal. Failure removes all newly appended metadata rows, and the Movement writer rolls back its target changes. State tables are read once per export operation and are released afterwards. A repeated unchanged export does not append redundant field snapshots.

Legacy exports without field snapshots still use their existing hash. If it matches actual fields, state can be recorded without a conflict. If it does not, exact attribution of the old manual edit is impossible: explicit basis acceptance is offered only while all other mapped fields already equal the generated fields. This initializes a baseline rather than reconstructing the historic target text.

## Manual verification in workbook copies

1. Import the changed sources and config manually, then export a closing event. Check typed snapshots and ExportJournal.
2. Manually change its Movement basis, update its incoming report number through the form, save the event and export. No Movement changes should be committed; the third table should show both texts.
3. Confirm the Movement value with Ctrl+Enter, then export. The manual text should remain; one decision should be journaled. Repeat export and verify no new conflict or duplicate field state.
4. Repeat using the EventsHistory value. Its generated basis should replace the manual text without changing the event RowId.
5. Change the target text again after selecting a value but before retrying. The old selection must not authorize the new value.
6. Test two people and a transition with both sides edited. All exposed conflicts must be chosen before target writes. Selecting another row must not clear a prior row's choice.
7. Confirm that formula edits and semantic-field edits are still rejected. Missing metadata tables must stop before Movement writes.
8. Make the history workbook unavailable for writing before export and verify no target changes survive. Runtime rollback and Ctrl+Enter behavior require interactive Excel verification; static checks do not establish them.