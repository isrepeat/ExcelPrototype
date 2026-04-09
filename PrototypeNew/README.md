# PrototypeNew

Clean-slate sandbox for rebuilding UI loading with object controls.

## Goal
No hardcoded control names in VBA. UI page is described in XML, and runtime builds controls via interface objects.

## Current scope
- Read Dev layout from `PrototypeNew/ui/DevUI.xml`.
- For each declared control, use `type` as root and auto-resolve:
	- control UI: `PrototypeNew/vba/controls/obj_<Type>ControlUI.xml`
	- VM class: `obj_<Type>ControlVM`
- Build controls through `obj_IControl` + `ex_ControlFactory`.
- Render controls through object VM classes (currently `obj_ButtonControlVM`, `obj_LabelControlVM`, `obj_TableListControlVM`, `obj_TableSingleControlVM`, `obj_BannerControlVM`).
- Button controls are rendered as Excel `Shape` objects (not Forms buttons) for richer visual customization.
  - shape bounds are derived from layout cell span (`row/col start/end`) instead of runtime point attributes.
- Table controls render list data directly into worksheet cells (faster path for bulk row output).
- Label controls render text directly into worksheet cells and can be used inside nested list templates.

## Shared control attributes
All controls support these common layout attributes via shared contract validation:
- `name`
- `type`
- `style`
- `spanCells`
- `spanRows`

Common attribute checks are centralized in `ex_ControlAttributeContracts` to avoid per-control duplication.

## Entry macro
Run:
- `ex_Actions.m_LoadPrototypeNewUi`

This resolves `ThisWorkbook.ActiveSheet` and calls `ex_SheetRenderer.m_RenderWorksheet`.

## DevTools Import Rules
`DevTools.dev_UpdateCodeFast` scans only root `vba\` and imports recursively (max depth `4`).

File classification is name-based (not folder-based):
- standard module: `ex_<Name>.vba` or `ex_<Name>.utf8.vba`
- class module: `obj_<Name>.cls.vba` or `obj_<Name>.cls.utf8.vba`
- worksheet module: `ws_<SheetName>.vba` or `ws_<SheetName>.utf8.vba`
- workbook module: `ThisWorkbook.vba` or `ThisWorkbook.utf8.vba`

Files that do not match these patterns are ignored by importer.

## Test helpers
Run from `ex_Test`:
- `m_TEST_RenderDevUI`
	- renders `ui\DevUI.xml` on active worksheet.
- `m_TEST_RegisterDemoListItems`
	- registers demo collection under itemsSource key `Test.People`.
- `m_TEST_RenderDevListUI`
	- registers demo table collections and renders nested-list table demo `ui\DevListUI.xml`.
- `m_TEST_RenderDevTableListUI`
	- alias for table list demo render.
- `m_TEST_RenderDevPrimitiveTableUI`
	- renders `ui\DevPrimitiveTableUI.xml` with table-like nested list templates built from primitive controls.
- `m_TEST_RenderDevListTableSingleUI`
	- registers demo 20 tables and renders `ui\DevListTableSingleUI.xml` (`List + itemsSourceTemplate + TableSingle` per item).
- `m_TEST_RenderDevTablePartStylesUI`
	- renders `ui\DevTablePartStylesUI.xml` with `controlPart` selector rules for `TableList` sections.
- `m_TEST_SetDemoTableItemsMany`
	- updates `RuntimeItems.Test.Tables` with 20 tables; if a page was rendered already, triggers full page rerender.
- `m_TEST_SetDemoTableItemsSingle`
	- updates `RuntimeItems.Test.Tables` with a single merged table; triggers full page rerender.
- `m_TEST_InsertDemoBanner`
	- updates `RuntimeItems.Test.Banner` and inserts a `Banner` control before table list with full rerender.
- `m_TEST_NoOp`
	- empty click handler for display-only test controls.

## Binding
Binding expressions use:
- `{Binding Path=<path>}`
- `{Binding Module=<module>; Method=<method>}`

If attribute value is not in `{Binding ...}` format, runtime treats it as literal value.

### Path resolution
- `Path=.` returns current source object.
- `Path=A.B.C` traverses nested members.
- Class members are resolved via `CallByName` (public properties/getters).
- `Scripting.Dictionary` is resolved by key name.
- VBA `Collection` supports:
	- `Count` (example: `Path=Rows.Count`)
	- 1-based numeric index (example: `Path=Rows.1`)

### Where bindings are used
- Text/value attrs (`caption`, `text`, `header`, `message`, etc.).
- Macro attrs (for example `onClick`).
- `visibility` on controls.
- `itemVisibility` on `TableList` (per-item filter before render).
- `itemsSource` and `objectSource` (resolved from runtime maps or by expression).
- For any `{Binding Path=...}` expression, optional conditional args are supported:
	- `Op`
	- `Value`
	- `TrueAs`
	- `FalseAs`

### Visibility expression format
- Simple:
	- `visibility="true"`
	- `visibility="{Binding Path=Visible}"`
- Conditional:
	- `visibility="{Binding Path=Status; Op=eq; Value=Active}"`
	- `itemVisibility="{Binding Path=RowCount; Op=gt; Value=3; TrueAs=Visible; FalseAs=Collapsed}"`

Supported `Op`:
- `eq`, `ne`, `gt`, `ge`, `lt`, `le`, `isTrue`, `isFalse`

`TrueAs` / `FalseAs` map condition result to final text token.
Default tokens when mapping is omitted: `True` / `False`.

### VBA registration example
```vb
Dim tables As Collection
Dim banner As obj_Banner

Set tables = m_TEST_BuildDemoTableItems()
Call ex_ListItemsSourceRuntime.m_SetItemsSource("RuntimeItems.Test.Tables", tables, True)

Set banner = New obj_Banner
banner.Header = "Data Source Updated"
banner.Message = "Banner inserted from objectSource"
banner.Visible = True
Call ex_ObjectSourceRuntime.m_SetObjectSource("RuntimeObjects.Test.Banner", banner, True)
```

### XML examples
```xml
<control type="Button"
         onClick="{Binding Module=ex_Test;Method=m_TEST_RenderDevUI}"/>

<itemControl objectSource="RuntimeObjects.Test.Banner"
             objectSourceTemplate="bannerObjectTpl"/>

<control type="TableList"
         itemsSource="RuntimeItems.Test.Tables"
         itemVisibility="{Binding Path=RowCount; Op=gt; Value=3; TrueAs=Visible; FalseAs=Collapsed}"/>

<control type="Label"
         text="{Binding Path=Rows.1.CellCount}"/>
```

## Runtime API
- `ex_SheetRenderer.m_RenderWorksheet(ws, wsUiPath)`
	- `ws` required: target worksheet.
	- `wsUiPath` optional: relative UI path for this worksheet page.
	- if `wsUiPath` is empty, runtime auto-loads `PrototypeNew/ui/<SheetName>UI.xml`.
	- delegates XML layout parsing/rendering to `ex_XmlLayoutEngine`.
	- runs page style pass via `ex_StylePipelineEngine.m_ApplyPageStyles` after successful layout render.
	- style pass behavior: `controlStyle` + only stage `name="default"`.

- `ex_SheetRenderer.m_ApplyWorksheetStyleStage(ws, stageName, wsUiPath)`
	- explicit stage apply for post-processing flow.
	- loads page UI and applies only the requested `stylePipelineStage` by name.

- `ex_XmlLayoutEngine.m_RenderPageLayout(wb, ws, wsUiDoc)`
	- parses and renders layout tags: `grid`, `stackPanel`, `control`, `list`.
	- `grid@sheet` is ignored; render target worksheet is always the `ws` argument.
	- computes control bounds and routes them to `ex_ControlRenderer`.

- `list` node contract
	- can be used as child of `grid` or `stackPanel`.
	- required attributes:
		- `itemsSource`
		- `itemsSourceTemplate`
	- `itemsSourceTemplate` points to `/uiDefinition/templates/template[@name='<name>']` and template must contain exactly one visual root node (`control|stackPanel|grid|list`).
	- optional `orientation` (`vertical` default, `horizontal`).

- `control type="TableList"` contract
	- required attributes:
		- `itemsSource`
	- optional attributes:
		- `itemVisibility` (visibility binding evaluated per itemsSource entry before render)
	- uses runtime data from `ex_ListItemsSourceRuntime` (same source map as `list`).
	- `itemsSource` entries must be table model objects:
		- `obj_TableDynamic`, or
		- `obj_Table` (auto-converted to dynamic model at render time).
	- rows inside table models must be `obj_Row`.
	- column count is dynamic per table item and validated against control `spanCells`.
	- renders rows directly into worksheet cells in control span area, without nested template expansion.

- `control type="TableSingle"` contract
	- required attributes:
		- `itemsSource`
	- expects one table model item in collection (or takes first item if collection has many entries):
		- `obj_TableDynamic`, or
		- `obj_Table` (auto-converted to dynamic model at render time).
	- intended for `list` item templates with object binding:
		- `itemsSource="{Binding Path=.}"`.
	- renders one table block (section, header, rows, spacer) into allocated span.

- `control type="Banner"` contract
	- required attributes:
		- `itemsSource`
	- expects first item in collection with fields:
		- `Header`
		- `Message`
		- `Visible` (`true/false`, optional; if absent inferred by non-empty header/message)
	- when `Visible=false`, banner rows are collapsed (`Hidden=True`); when visible, header/message block is rendered.

- table object model classes
	- `obj_Table`
		- fixed-size table; call `m_Init(rowCount, columnCount)` on creation.
		- supports `m_SetColumn`, `m_SetRow`, `m_SetCell`.
	- `obj_TableDynamic`
		- dynamic table; supports `m_AddColumn(obj_Column)` and `m_AddRow(obj_Row)`.
	- `obj_Column`
		- column metadata (`Name`, `Position`).
	- `obj_Row`
		- row cell container (`m_AddCell`, `m_SetCell`, `m_GetCell`).

- `ex_ListItemsSourceRuntime`
	- runtime source registry for `list` items.
	- `m_SetItemsSource(key, itemsCollection, notifyChange)` registers list data for key-based `itemsSource`.
	- when `notifyChange=True` and key is not internal runtime key, runtime triggers full rerender of last rendered page.
	- `m_ResetItemsSources()` clears runtime list sources.
	- `itemsSource` can be:
		- key registered through `m_SetItemsSource`,
		- binding expression resolved against runtime items source map,
		- inline scalar list (`a|b|c` or `a;b;c`).

- `ex_XmlLayoutEngine.m_RenderTemplateChildren(wb, ws, templateControlNode, depth, layoutRowStart, layoutColStart, layoutRowEnd, layoutColEnd)`
	- recursively renders composite controls declared inside control templates.

- `ex_ControlRenderer.m_RenderControl(wb, ws, layoutControlNode, recursionDepth, layoutRowStart, layoutColStart, layoutRowEnd, layoutColEnd)`
	- validates control attributes against each control's contract (`obj_IControl.SupportsAttribute`).
	- loads control template `obj_<Type>ControlUI.xml`, applies allowed overrides from page UI, and renders control VM class.
	- passes worksheet row/column bounds to controls rendered from worksheet-span layout path.
	- triggers recursive rendering for template child controls via `ex_XmlLayoutEngine`.

- `ex_StylePipelineEngine.m_ApplyPageStyles(ws, wsUiDoc)`
	- styles are fully declared in `uiDefinition/styles`.
	- pass order:
		1) apply `controlStyle` declarations to controls by `style` key.
		2) auto-apply only `stylePipelineStage name="default"` to worksheet cells/ranges.
	- `stylePipelineStage name="default"` is required in each page UI file.

- `ex_StylePipelineEngine.m_ApplyPageStyleStage(ws, wsUiDoc, stageName)`
	- explicit stage execution API.
	- use for post-processing stages (analogue of old Prototype banners stage).
	- `stylePipelineStage` supports targets: `row`, `column`, `cell`, `range`, `usedrange`, `sheet`, `controlPart`.
	- supported rule selector keys: `col`, `row`, `address`, `type`, `name`, `part`.
	- `controlPart` target contract:
		- required selector keys: `type`, `part`
		- optional selector key: `name`
		- currently registered control type: `tablelist`
		- supported table parts: `section`, `header`, `rows`, `spacer`
	- button visuals should be styled via `controlStyle`; pipeline rules are intended for sheet cell formatting.
