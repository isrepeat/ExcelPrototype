# PrototypeNew

Clean-slate sandbox for rebuilding UI loading with object controls.

## Goal
No hardcoded control names in VBA. UI page is described in XML, and runtime builds controls via interface objects.

## Current scope
- Read Dev layout from `PrototypeNew/ui/DevUI.xml`.
- For each declared control, use `type` as root and auto-resolve:
	- control UI: `PrototypeNew/vba/classes/obj_<Type>ControlUI.xml`
	- ViewModel: `obj_<Type>ControlViewModel`
- Build controls through `obj_IControl` + `ex_ControlFactory`.
- Render controls through object ViewModels (currently `obj_ButtonControlViewModel`, `obj_TableControlViewModel`).
- Button controls are rendered as Excel `Shape` objects (not Forms buttons) for richer visual customization.
- Table controls render list data directly into worksheet cells (faster path for bulk row output).

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
- `m_TEST_NoOp`
	- empty click handler for display-only test controls.

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

- `control type="Table"` contract
	- required attributes:
		- `itemsSource`
	- uses runtime data from `ex_ListItemsSourceRuntime` (same source map as `list`).
	- `itemsSource` entries must be table model objects:
		- `obj_TableDynamic`, or
		- `obj_Table` (auto-converted to dynamic model at render time).
	- rows inside table models must be `obj_Row`.
	- column count is dynamic per table item and validated against control `spanCells`.
	- renders rows directly into worksheet cells in control span area, without nested template expansion.

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
	- `m_SetItemsSource(key, itemsCollection)` registers list data for key-based `itemsSource`.
	- `m_ResetItemsSources()` clears runtime list sources.
	- `itemsSource` can be:
		- key registered through `m_SetItemsSource`,
		- binding expression resolved against runtime items source map,
		- inline scalar list (`a|b|c` or `a;b;c`).

- `ex_XmlLayoutEngine.m_RenderTemplateChildren(wb, ws, templateControlNode, left, top, width, height, depth)`
	- recursively renders composite controls declared inside control templates.

- `ex_ControlRenderer.m_RenderControl(wb, ws, layoutControlNode, left, top, width, height, recursionDepth, layoutRowStart, layoutColStart, layoutRowEnd, layoutColEnd)`
	- validates control attributes against each control's contract (`obj_IControl.SupportsAttribute`).
	- loads control template `obj_<Type>ControlUI.xml`, applies allowed overrides from page UI, and renders control ViewModel.
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
	- `stylePipelineStage` supports only cell targets: `row`, `column`, `cell`, `range`, `usedrange`, `sheet`.
	- supported rule selector keys: `col`, `row`, `address`.
	- button visuals should be styled via `controlStyle`; pipeline rules are intended for sheet cell formatting.
