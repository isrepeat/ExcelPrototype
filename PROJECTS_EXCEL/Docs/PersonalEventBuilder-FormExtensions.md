# Form extensions

Extensions use `Data.FormColumnMap`, the same viewport map as the main form. A type is available only when every field in `extensions.type.<type>.fields` belongs to the active form and is inside the current viewport. Disabled fixed-column placeholders do not enable extensions. There are no fallback extension coordinates.

Each extension occupies two rows: field captions and values. Its title and remove button occupy free cells alongside those rows. Field positions are never adjusted to make room for service cells. Extensions with disjoint field and service cells share a row pair; overlapping extensions use subsequent row pairs. Insufficient service space is reported as a configuration error.

Horizontal scrolling rebuilds both extension menus and rows using the new viewport. An extension is hidden as a whole when any of its fields leaves the viewport. Its values remain in the current event draft and are restored when all fields become visible again. Open extension candidate tables close before a viewport or mode change. Candidate tables align to their extension's mapped form.

Switching to a mode that lacks an extension's fields hides that extension and excludes it from export. Its draft values are retained until the event is reset or another event is loaded. Saving another mode detaches inactive drafts from obsolete history identifiers so they can be saved again after returning to a compatible mode. Scrolling alone does not exclude an otherwise active extension from export or required-field validation.

## Mode defaults

`extensions.modes` lists modes with explicit extension defaults. Each listed mode defines:

- `extensions.mode.<mode>.defaults`: `Type=Role` entries separated by `|`.
- `extensions.mode.<mode>.required`: role identifiers separated by `|`, or `-` for optional defaults.
- `extensions.mode.<mode>.primaryCaption`: the primary document field caption.

Each role defines `extensions.role.<role>.caption` and `.required` (field identifiers separated by `|`). Defaults are created once per event and mode. Returning to a mode reuses its existing rows. Required rows display the configured required marker instead of a remove button and enforce their role's required fields during collection.

`MoveFromHospitalToVacationForCuring` uses the primary caption `Виписка` and a required `Document=Vlk` extension captioned `Довідка ВЛК`. The additional document requires its name, number and date. Defaults do not fabricate document values.

## History compatibility

`ExtensionRole` is persisted through `history.fields` and revision snapshots. Reading an older history without this column returns an empty role. The normal save path adds the optional history column and its snapshot counterpart. Legacy untyped saved document rows may be assigned the configured default role when loading that mode. Empty role metadata does not change existing movement-link stamps.

Source-only checks do not execute VBA, compile projects, import modules, or update workbooks. Deploy the VBA sources and both applicable configuration profiles together before checking the layout in Excel.