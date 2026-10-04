# Style pipeline

The page owns a compiled `obj_UiStylePipeline`. `BeginPage` reads common and local named styles and compiles rule selectors and declarations once. No XML serialization or declaration parsing occurs per control. `BeginRender` clears only runtime regions; full rendering registers the new geometry and applies the `default` stage after all controls render. Named stages can also be applied explicitly through `Styles.ApplyStage`.

Rules run in XML order across layers. `enabled=false` on a stage, layer or rule disables that rule. Named control styles and direct attributes form the initial appearance; subsequent pipeline declarations overwrite it. Selected-row overlays are replayed last. To retain a named style after a broad sheet rule, reference it in a later rule through `style`, optionally adding overrides in `styles`.

Supported targets: `sheet`, `usedRange`, `row`, `column`, `cell`, `range`, `control`, `controlPart`, `layoutContainer`, `layoutBound`, `inlinePart`. Range selectors accept numeric `row=2:8`, `col=3:5` or Excel `address=B2:D8`. A cell needs row and col, or address; a range needs address. Runtime selectors accept type, name, style, tags, element, elementDepth and part. Sheet filtering is optional. Unknown keys, malformed declarations, duplicate selector keys and invalid enabled values stop initialization.

Whole-row and whole-column addresses are clipped to the page scope before visual formatting. Structural column width still applies to entire worksheet columns. This avoids filling a million cells when a rule uses address=A:A. Runtime regions replace existing registrations with the same identity instead of accumulating during repeated binding refreshes. Stripe formatting batches up to 64 rows per Excel operation.

Grid and stackPanel register layout bounds. Controls register their control/cell/shape regions. Table and tableList register section (title), header and rows. TableList children retain the tableList owner type and name for selectors. Column aliases for table header/data regions currently use the displayed column header. `RegisterPart` also accepts explicit columnAlias, sourceAlias and sourceAliasTemplate for providers that have this metadata. An alias is an exact match against a registered region, not a worksheet address.

```xml
<stylePipelineStage name="default">
  <layer name="table-formatting">
    <rule target="controlPart"
          selector="type=table;part=header"
          styles="{backColor:#355D8C;fontColor:#FFFFFF;fontBold:true;}"/>
    <rule target="controlPart"
          selector="type=table;part=rows"
          styles="{backColor:#2D2D2D;fontColor:#FFFFFF;}"/>
    <rule target="controlPart"
          selector="type=table;part=row;tags=even"
          styles="{backColor:#383838;}"/>
  </layer>
</stylePipelineStage>
```

Таблица регистрирует каждую строку данных как `part=row` с метками `tableRow odd` или `tableRow even`. Нумерация начинается с единицы, исключает заголовок и название и начинается заново для каждой таблицы в tableList. Область всего тела сохраняет `part=rows`. Движок стилей не вычисляет чётность: он выбирает области по меткам.

`tags=even` требует наличия метки even; `tags=tableRow even` требует обеих меток. Сравнение без учёта регистра, по целым словам, разделённым пробельными символами. Метки области дополняют атрибут `tags` её элемента. `RegisterPart` принимает необязательный параметр `tags`, поэтому другие контролы также могут регистрировать семантические метки.

`element=stackPanel` выбирает имя элемента разметки; `elementDepth=0` или `elementDepth=3:10` задаёт глубину или диапазон глубины среди родительских XML-элементов. Прежние имена `tag` и `tagDepth` не поддерживаются. Селектор `parity` заменён метками строк.

Supported declarations include colors, fonts, borders, horizontal/vertical alignment, overflow, width/columnWidth, minWidth/maxWidth, autoFitColumns, rowHeight (number or auto), cellType (text/date/general), sheet zoom and gridLines. Width uses Excel ColumnWidth units; it changes worksheet columns shared by neighboring controls. Structural properties should normally be assigned by range/column rules.

`ResolveInlineStyle` returns merged declarations for a text renderer; it does not automatically turn a plain cell into rich text. Source aliases and itemBanner/rowBanner regions require registration by a data-aware control. PROJECTS_EXCEL currently has no banner-row or rich-text table renderer, so those PrototypeNew renderers were not copied into the universal style engine. Full page rendering is the supported path; retained partial-style replay and layout-bound translation from PrototypeNew are intentionally unnecessary for the current full-render policy.

Run `WorkbookUpdater/Test-StylePipelineSources.py` and `WorkbookUpdater/Test-LookupSources.py` for read-only declaration/configuration checks. They do not establish VBA runtime correctness; compilation, imports and macro execution require separate explicit authorization.