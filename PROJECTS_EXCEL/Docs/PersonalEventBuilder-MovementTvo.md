# Movement TVO table

History TVO children retain their actor's own `PositionCode`. During history loading, `DutyPositionCode` and `DutyPersonName` come from the immediate parent, including nested replacements. The Movement planner copies these rows and uses `DutyPositionCode` as the replaced post without altering the stored child or its Word context. Existing registry and return rows already contain the replaced post. Assignment comparison uses the same normalized metadata as insertion.

New departure and transition segments render their text columns from `PersonalEventBuilderMovementTemplates.xml`. A `column` without `target`, or with `target="Absence"`, belongs to the absence table. A `column target="TVO"` belongs to the TVO table and is rendered once per assignment before planning the cell changes. `configKey` resolves the destination header through wsConfig; text content, conditions and reasons live in XML. Missing or invalid TVO templates stop the operation before writes. These fields participate in previews, conflict stamps, verification and rollback. Previously exported rows are not automatically repaired.

```xml
<column configKey="returnTvo.table.headers.Duration"
        target="TVO">
  <text>Інше</text>
</column>
<column configKey="returnTvo.table.headers.Information"
        target="TVO">
  <text>{#include block.TvoInformation}</text>
</column>
```

Each departure/transition template declares both columns. The shared `block.TvoInformation` uses `TvoReplacedPersonName` for the immediate replaced person, `PersonName` for the original absent person, and `TvoIsChain` to select the chain wording. The `TvoDepartureReasons` XML dictionary supplies the reason by `TargetStatus`. All other prepared Movement fields remain available. wsConfig retains only header mappings; `returnTvo.table.values` and `returnTvo.reasons` are no longer used.

An admission (`Доступ.Наказ`) is distinct from acceptance of duties. A row without `Прийняття.Наказ` is not an executing assignment, even when it has an admission order or a provisional date. Such a registry row restores no duties and suppresses legacy fallback for that represented position. Departure/transition metadata must not promote that admitted person into a new accepted segment or fill acceptance/handover fields. Accepted segments may still continue through transitions, and accepted assignments with a missing acceptance date remain errors. `MOVEMENT_TVO_EVENT`, `MOVEMENT_TVO_NOT_STARTED` and `MOVEMENT_TVO_ADMISSION_RETAINED` log the event context and admission decisions.

Movement export plans changes to both the configured absence ListObject and the `ТВО` ListObject in the same open target workbook. `returnTvo.table.name` and `returnTvo.table.headers` configure the TVO destination in PROD and TEST. The worksheet is shared with `reporter.tvo.source.sheet`. No source workbook is saved, opened or closed by the writer.

On arrival, return-duty restoration first follows assignments whose replaced position matches the returning person's absence-period position. Each actor's stored own position leads to the next replacement in the chain. The assignment must have started by the event date; an actual handover ends its interval. A handover already recorded on that event's date and order remains readable for repeat arrival previews. Assignments closed before the relevant absence period do not suppress legacy restoration. Cycles, overlapping active assignments and incomplete identity or position data stop processing.

When a corresponding TVO registry record exists, it takes priority over the old absence columns. Only an unrepresented absence falls back to `ТВО.ПІБ`, `ТВО.ІПН` and `ТВО.Посада`. Existing explicit history children and frozen return snapshots retain their established preview semantics. The Movement writer independently closes the applicable live registry chain, including when an older event snapshot contains different return-duty children.

Arrivals fill actual handover date, order and basis in the applicable TVO rows. Status transitions close the previous assignment segments and open successor segments; if the old period exists only in legacy columns, the new segment is created directly in the TVO table. Departures with duties create new assignment rows. New rows contain the actor's name, tax ID, rank and own personnel position, replaced position, acceptance/access order, acceptance date and basis, and an available planned end date. A planned end date does not close an assignment. New exports never write the three legacy absence columns.

The planner models the whole order before any writes. It validates destination headers, formulas, free expansion ranges and assignment overlaps. Already-applied events do not create duplicate TVO segments. Full table value/formula stamps detect changes between planning and writing; number formats do not participate in conflict or content verification. Actual TVO cell changes appear in the same pre-write diff as absence changes. Buttons above the diff select the absence or TVO worksheet using the actual configured worksheet names.

The writer snapshots TVO cell values and formats and any blank extension area. TVO writes and verification occur before export journaling; failures restore both tables, their original dimensions and changed cell values/formats. Restoration failures are reported. The writer does not save the target workbook.

Revisions that would rewrite an active TVO chain, and cancellation of a represented TVO assignment, are currently rejected for manual review rather than silently changing registry history. Automatic cancellation/revision reconciliation of existing TVO intervals requires a separate design for persistent event identity in this table.

Static checks: `validate_movement_tvo.py`, `validate_return_tvo.py`, `validate_movement_diff.py`, configuration/snapshot checks, VBA structure/size checks and UI style/XML checks. These checks do not execute Office or verify Excel runtime behavior. Runtime acceptance must cover a new departure, a transition continuing an assignment, a legacy-only transition, an arrival with a registry chain, a legacy-only arrival, repeated export, conflicting assignments and a forced failure after TVO writes to verify joint rollback.