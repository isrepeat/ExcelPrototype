# Layout Grid v2 (Draft)

Status: implemented baseline for DevUI integration.

## Goals

- Layout geometry is declared in UI XML.
- Visual styling remains outside layout (button styles / style pipeline).
- No cross-node references by layout ids are required.
- Main layout primitives: `grid`, `stackPanel`, `border`, `control`.

## Supported Layout Nodes

- `grid`
- `stackPanel`
- `border`
- `control`

No `zone`, no `viewHost`, no `spacer` node.
Spacing is emulated via `border` with fixed `width`/`height` and no children.

## Styles Block

Inside `<styles>`, layout engine reads:

- `controlStyle`

`controlStyle` fields:

- `name` (required)
- `backColor`
- `textColor`
- `borderColor`
- `borderWeight`
- `fontName`
- `fontSize`
- `fontBold`

## Result Layout Controls

For result pages, `control` supports:

- `type="button"`
- `type="dropdownButton"`
- `type="input"` (cell-based input bound to config/onChange)
- `type="label"`
- `type="table"`
- `type="itemsPanel"`

`itemsPanel` dynamic rendering attributes:

- `itemsSource` (must resolve to `Collection`)
- `itemTemplate` (template name from `/uiDefinition/templates/template[@name='...']`)
- `itemsSourceFilter` (optional regex; applied to `MetaInfo.Name`)
- `itemsSourceFilterBind` (optional binding path for filter target, e.g. `{Binding TableRef}`)
- `orientation="vertical|horizontal"`

If `itemsSourceFilterBind` is omitted, engine uses legacy target `MetaInfo.Name`.

Template contract:

- A template must contain exactly one root node:
  `stackPanel | border | control`.

Shape control notes (`button`, `dropdownButton`):

- `name` is required.
- `macro`, `visible`, `placement` are supported.
- `button`/`dropdownButton`: `caption` (or `text`), optional `style`.

Input control notes (`input`):

- `inputName` (optional logical key; defaults to `name` or `inputConfigKey`).
- `inputConfigKey` (optional config key for default value + write-back on change).
- `onChange` / `onChangeMacro` (optional macro to run on sheet change).
- `inputPrimary="true"` to bind this input as primary search cell.
- Visuals are not configured via inline control attributes (`backColor`, `textColor`, etc.).
  Use `style="..."` (style declared in `/uiDefinition/styles`) or StylePipeline.

## Sizing

For layout nodes:

- `width="N"`, `height="N"` where `N` is integer tracks (cells).
- `width="auto"`, `height="auto"` are supported for containers (`stackPanel`, `border`).

For `control`:

- `width` and `height` must be numeric.
- `auto` is not supported for control size.

If container size is `auto` and visible content is empty, resolved size is `0`.

`*` is intentionally not supported in this baseline.

## Coordinates

- `grid@anchorCell` sets origin, default `A1`.
- Root layout nodes may define `at="rNcM"` (example: `r2c17`).
- Nodes placed inside `stackPanel` / `border` flow must not define `at`.

## stackPanel

Required attribute:

- `orientation="vertical|horizontal"`

Behavior:

- `vertical`: children are placed top-to-bottom.
- `horizontal`: children are placed left-to-right.
- No built-in `gap` support.

## border

- Can be an empty spacer.
- Can host child layout nodes.
- Children inside `border` are overlaid at the same top-left corner.

## grid columns

Optional inside `grid`:

```xml
<columns>
  <col i="1" width="20" lockWidth="true"/>
</columns>
```

Current baseline applies `i` + `width`.
`lockWidth` is kept as declarative metadata for later lock enforcement extensions.
For layout nodes (`stackPanel`, `border`, `control`) only `spanCells`/`spanRows` are valid.
Legacy `width`/`height` on layout nodes are rejected (no fallback).

## Control discovery

Controls are discovered from both paths:

- `/uiDefinition/controls/control` (legacy)
- `/uiDefinition/layout//control` (new)

## Example

```xml
<layout>
  <stableZone startCol="H" minBufferWidth="5"/>

  <grid sheet="Dev" anchorCell="A1">
    <stackPanel at="r2c17" spanCells="2" spanRows="23" orientation="vertical">
      <control name="btnClear" type="button" spanCells="2" spanRows="2"/>
      <border spanCells="2" spanRows="1"/>
      <control name="btnUpdateUI" type="button" spanCells="2" spanRows="2"/>
    </stackPanel>
  </grid>
</layout>
```

Result-page dynamic example:

```xml
<layout>
  <grid anchorCell="A1">
    <columns>
      <col i="1" width="20"/>
    </columns>

    <control type="itemsPanel"
             at="r1c1"
             itemsSource="ResultTables"
             itemsSourceFilter="^(Main|Events)"
             itemsSourceFilterBind="{Binding TableRef}"
             itemTemplate="ResultTableTemplate"
             orientation="vertical"
             spanCells="20"
             spanRows="auto"/>
  </grid>
</layout>

<templates>
  <template name="ResultTableTemplate">
    <stackPanel orientation="vertical" spanCells="20" spanRows="auto">
      <control type="label" text="{Binding TableRef}" spanCells="20" spanRows="1"/>
      <control type="table" itemsSource="{Binding Rows}" spanCells="20" spanRows="auto"/>
      <border spanCells="20" spanRows="1"/>
    </stackPanel>
  </template>
</templates>
```
