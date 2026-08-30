# PrototypeNew

Clean-slate sandbox for rebuilding UI loading with object controls.

## Goal
No hardcoded control names in VBA. UI page is described in XML, and runtime builds controls via interface objects.

## Current scope
- Read Main layout from `PrototypeNew/ui/MainUI.xml`.
- For each declared control, use `type` to resolve its VM class `obj_<Type>ControlVM`.
- The page XML is the single UI source; controls and their optional template children are cloned directly from the loaded page DOM.
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
- `spanColls`
- `spanRows`

Common attribute checks are centralized in `ex_ControlAttributeContracts` to avoid per-control duplication.

## Entry macro
Run:
- `ThisWorkbook.m_ResetWorkbookAndCreateMainPage`

This resets workbook sheets, creates `Main` page and renders it via `rt_PageManager`.

## Settings.xml (XML flags)
- External file: `Settings.xml` рядом с `.xlsm` (папка `ThisWorkbook.Path`).
- Если файла нет, `ex_Core.m_Settings_*` автоматически создает шаблон `Settings.xml`.
- Current flag:
  - `<EnableLogging>true|false</EnableLogging>`
- Cache:
  - используется file-cache в `ex_Core` (кэш текста + `DateLastModified`).
  - `GlobalRuntimeSource='settings'` резолвится в snapshot-объект настроек через `ex_Core.fn_Settings_TryGetObjectSource`.
  - при изменении даты файла `Settings.xml` перечитывается автоматически.
- Если запрошенного флага нет в `Settings.xml`, он автоматически добавляется со значением по умолчанию.
- Main page button:
  - `Toggle Logging` переключает `EnableLogging` в `Settings.xml`.

## Роутинг Кнопок И Обновление Модулей

### Почему роуты это runtime-состояние
Кнопки рендерятся как Excel Shape. Клик по Shape может вызвать только макрос по имени через `OnAction`, а не метод объекта напрямую.
Из-за этого роутинг разбит на стабильные runtime-слои:

1. `OnAction` указывает на стабильный bridge-макрос (`rt_Bridge.fn_OnShapeClick`).
2. Bridge делегирует диспетчеризацию в `rt_Router`.
3. `rt_Router` хранит runtime-таблицы:
	- имя shape -> запись маршрута
	- ключ контрола -> объект VM
4. Запись маршрута содержит, какой метод VM нужно вызвать.

Такая схема позволяет XML-контролам (`onClick="{Binding Module=...;Method=...}"`) работать без хардкода макроса под каждую кнопку.

### Полный путь клика

1. `ex_SheetRenderer.m_RenderWorksheet` сбрасывает runtime-состояние кликов (`rt_Router.m_ResetRouter`).
2. Во время рендера контролов создаются shape-кнопки и регистрируются маршруты через dispatcher/router.
3. Пользователь кликает shape.
4. Excel запускает макрос из `OnAction` (`rt_Bridge.fn_OnShapeClick`).
5. Bridge вызывает `rt_Router.fn_OnShapeClick`.
6. Router берет `Application.Caller`, находит маршрут, находит VM и вызывает целевой метод.

### Безопасное обновление runtime

Для UI-триггеров используется только `rt_CoreActions`. Он ставит один точный
`OnTime` callback, позволяет текущему click-dispatch полностью завершиться и
лишь затем вызывает `ex_Core`. Перед импортом `rt_Lifecycle` отменяет внешние
callbacks и освобождает runtime-граф; после успешного импорта на следующем тике
создаётся новая страница Main и заново регистрируются routes.

Стандартные модули и классы обновляются in-place. Удаляются только stale-компоненты,
исходных файлов которых больше нет. Это сохраняет COM/type identity VBA-классов и
не оставляет Shape привязанными к старым instance.

`ex_Core` остаётся автономным initial installer. Если в VBA-проекте нет ни одного
обязательного runtime-компонента, Full Update выполняет первичный импорт без
вызова lifecycle, после чего запускает обычную инициализацию Main. Если runtime
установлен лишь частично, операция показывает отсутствующие компоненты и
останавливается. Автоматического восстановления частичной установки и скрытых
retry нет.

### Открытие и закрытие книги

Единственная event-точка запуска — `ThisWorkbook.Workbook_Open`; `Auto_Open` и
дополнительные bootstrap-макросы не используются. Cold open и восстановление
после Update Code вызывают один `rt_Lifecycle.fn_InitializeRuntime`. При закрытии сначала
фиксируется решение пользователя о сохранении, затем `rt_Lifecycle` выполняет
две фазы: отменяет все зарегистрированные `OnTime` callbacks и только после
этого освобождает COM-ресурсы, страницы, controllers и runtime registries.
Владелец каждого callback хранит его точное время и имя макроса. Ошибка отмены
останавливает закрытие до начала dispose, поэтому закрытая книга не может быть
повторно открыта Excel ради оставшейся отложенной задачи.

## DevTools Import Rules
`ex_Core.fn_Dev_UpdateCodeByDate` scans only root `vba\` and imports recursively (max depth `4`).

File classification is name-based (not folder-based):
- standard module: `ex_<Name>.vba` or `ex_<Name>.utf8.vba`
- class module: `obj_<Name>.cls.vba` or `obj_<Name>.cls.utf8.vba`
- worksheet module: `ws_<SheetName>.vba` or `ws_<SheetName>.utf8.vba`
- workbook module: `ThisWorkbook.vba` or `ThisWorkbook.utf8.vba`

Files that do not match these patterns are ignored by importer.

## Test helpers
Run from `ex_Test`:
- `fn_TEST_RenderDevUI`
	- renders `ui\MainUI.xml` on active worksheet.
- `fn_TEST_RegisterDemoListItems`
	- registers demo collection under itemsSource key `Test.People`.
- `m_TEST_RenderDevListUI`
	- registers demo table collections and renders nested-list table demo `ui\Dev\DevTableListUI.xml`.
- `fn_TEST_RenderDevTableListUI`
	- alias for table list demo render.
- `fn_TEST_RenderDevPrimitiveTableUI`
	- renders `ui\Dev\DevPrimitiveTableUI.xml` with table-like nested list templates built from primitive controls.
- `fn_TEST_RenderDevListTableSingleUI`
	- registers demo 20 tables and renders `ui\Dev\DevListTableSingleUI.xml` (`List + itemsSourceTemplate + TableSingle` per item).
- `fn_TEST_RenderDevTablePartStylesUI`
	- renders `ui\Dev\DevTablePartStylesUI.xml` with `controlPart` selector rules for `TableList` sections.
- `fn_TEST_SetDemoTableItemsMany`
	- updates `RuntimeItems.Test.Tables` with 20 tables; if a page was rendered already, triggers full page rerender.
- `fn_TEST_SetDemoTableItemsSingle`
	- updates `RuntimeItems.Test.Tables` with a single merged table; triggers full page rerender.
- `fn_TEST_InsertDemoBanner`
	- updates `RuntimeItems.Test.Banner` and inserts a `Banner` control before table list with full rerender.
- `fn_TEST_NoOp`
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
Dim pageBase As obj_PageBase

Set tables = fn_TEST_BuildDemoTableItems()
Call pageBase.RuntimeSources.SetItemsSource("RuntimeItems.Test.Tables", tables)

Set banner = New obj_Banner
banner.Header = "Data Source Updated"
banner.Message = "Banner inserted from objectSource"
banner.Visible = True
Call pageBase.RuntimeSources.SetObjectSource("RuntimeObjects.Test.Banner", banner)
```

### XML examples
```xml
<control type="Button"
         onClick="{Binding Module=ex_Test;Method=fn_TEST_RenderDevUI}"/>

<itemControl objectSource="{PageRuntimeSource='RuntimeObjects.Test.Banner'}"
             objectSourceTemplate="bannerObjectTpl"/>

<control type="TableList"
         itemsSource="{PageRuntimeSource='RuntimeItems.Test.Tables'}"
         itemVisibility="{Binding Path=RowCount; Op=gt; Value=3; TrueAs=Visible; FalseAs=Collapsed}"/>

<control type="TableList"
         itemsSource="{PageRuntimeSource='RuntimeItems.Test.Tables'}"/>

<control type="Button"
         dataContext="{GlobalRuntimeSource='settings'}"
         caption="{Binding Path=EnableLogging;TrueAs=Disable Logging;FalseAs=Enable Logging}"
         onClick="{Binding Module=ex_Core;Method=fn_Dev_ToggleLogging}"/>

<control type="Label"
         text="{Binding Path=Rows.1.CellCount}"/>
```

## Runtime API
- `ex_SheetRenderer.m_RenderWorksheet(ws, wsUiPath)`
	- `ws` required: target worksheet.
	- `wsUiPath` optional: relative UI path for this worksheet page.
	- if `wsUiPath` is empty, runtime auto-loads `PrototypeNew/ui/<SheetName>UI.xml`.
	- delegates XML layout parsing/rendering to `ex_XmlLayoutEngine`.
	- runs page style pass via `ex_StylePipelineEngine.fn_ApplyPageStyles` after successful layout render.
	- style pass behavior: `controlStyle` + only stage `name="default"`.

- `ex_SheetRenderer.m_ApplyWorksheetStyleStage(ws, stageName, wsUiPath)`
	- explicit stage apply for post-processing flow.
	- loads page UI and applies only the requested `stylePipelineStage` by name.

- `ex_XmlLayoutEngine.m_RenderPageLayout(wb, ws, wsUiDoc)`
	- parses and renders layout tags: `grid`, `stackPanel`, `control`, `list`.
	- `grid@sheet` is ignored; render target worksheet is always the `ws` argument.
	- computes control bounds and routes them to `ex_LayoutControlRenderer`.

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
	- uses runtime data from `pageBase.RuntimeSources` (same source map as `list`).
	- `itemsSource` entries must be table model objects:
		- `obj_TableDynamic`, or
		- `obj_Table` (auto-converted to dynamic model at render time).
	- rows inside table models must be `obj_Row`.
	- column count is dynamic per table item and validated against control `spanColls`.
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

	- `obj_PageRuntimeSources` (owned by each `obj_PageBase`)
		- page-local runtime source registry for `itemsSource` / `objectSource`.
		- `SetItemsSource(key, itemsCollection)` registers list data for key-based `itemsSource`.
		- `SetObjectSource(key, sourceObject)` registers data for key-based `objectSource`.
		- `ResetItemsSources()` / `ResetObjectSources()` clear runtime maps.
		- `TryResolveItemsSource()` / `TryResolveObjectSource()` resolve binding/key lookups against page-local maps.
		- source text can be:
			- explicit runtime-source expression, e.g. `{PageRuntimeSource='RuntimeItems.Test.Tables'}`
			- explicit global runtime-source expression, e.g. `{GlobalRuntimeSource='settings'}`
			- binding expression that resolves directly to object/collection, e.g. `{Binding Path=.}` or `{Binding Path=Rows}`

- `ex_XmlLayoutEngine.fn_RenderTemplateChildren(wb, ws, templateControlNode, depth, layoutRowStart, layoutColStart, layoutRowEnd, layoutColEnd)`
	- recursively renders composite controls declared inside control templates.

- `ex_LayoutControlRenderer.fn_Render(renderCtx, layoutControlNode, layoutRowStart, layoutColStart, layoutRowEnd, layoutColEnd)`
	- validates control attributes against each control's contract (`obj_IControl.SupportsAttribute`).
	- clones the control node directly from page UI, validates VM attributes, and renders the control VM class without per-control XML file I/O.
	- passes worksheet row/column bounds to controls rendered from worksheet-span layout path.
	- triggers recursive rendering for template child controls via `ex_XmlLayoutEngine`.

- `ex_StylePipelineEngine.fn_ApplyPageStyles(ws, wsUiDoc)`
	- styles are fully declared in `uiDefinition/styles`.
	- pass order:
		1) apply `controlStyle` declarations to controls by `style` key.
		2) auto-apply only `stylePipelineStage name="default"` to worksheet cells/ranges.
	- `stylePipelineStage name="default"` is required in each page UI file.

- `ex_StylePipelineEngine.fn_ApplyPageStyleStage(ws, wsUiDoc, stageName)`
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
