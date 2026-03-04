# ExcelPrototype

## Настройка Codespace (алиасы + логи)

Если работаешь через расширение и оно работает нормально, этот шаг можно пропустить.
Используй его как fallback для CLI-сценария (когда расширение недоступно/нестабильно), чтобы писать локальные логи сессии в терминале.
```bash
bash .env/Codex/setup_codexlog.sh
```
Использование:
```bash
codexlog
```
Это запускает Codex с логированием терминала в `.env/Codex/logs/codex-YYYY-MM-DD_HH-MM.log`.

## Git-алиасы для репозитория

```bash
bash .env/git/setup_git_aliases.sh
bash .env/git/remove_git_aliases.sh
```

Проверить, какие алиасы активны и откуда они взялись:
```bash
git config --show-origin --get-regexp '^alias\.'
```


## PostProcess DSL

`postProcessScript` задается в профиле, внутри `Prototype/config/modes/*/*Profiles.xml`.

Пример размещения:
```xml
<postProcessScript>
    // DSL-код
</postProcessScript>
```

### Синтаксис и конструкции

Поддерживаемые инструкции:

1. `for (...) { ... }`
2. `if (...) { ... }`
3. `callMacro("Module.Proc", ...);`
4. `let varName = callMacro("Module.Proc", ...);`

Комментарии:

1. Однострочные: `// comment`
2. Многострочные: `/* comment */`

### Циклы

Итерация по строкам таблицы:
```text
for (row in Events.Sheet[EventsOut].rows) {
    ...
}
```

Итерация по колонкам строки:
```text
for (col in row.columns) {
    ...
}
```

`col` (объект колонки) поддерживает:

1. `col.alias` (или `col.name`)
2. `col.mapKey`
3. `col.value`

### Условия

Поддерживаются операторы сравнения:

1. `==`
2. `!=`
3. `gt`, `lt`
4. `gte`, `lte`

И логические связки:

1. `and`, `or`

Важно:

1. Правая часть сравнения должна быть строковым литералом в кавычках.
2. Для `gt/lt/gte/lte` сравнение числовое, если обе стороны парсятся как число; иначе строковое (без учета регистра).
3. Пример: `if (row.column[EventNum] gt "20000" and row.column[Date] != "") { ... }`

### Переменные

Переменная создается только через `let` в связке с `callMacro` (и имеют строковый тип):
```text
let key = callMacro("ex_ResultRuntimeAdapter.m_BuildMapKey", "Events", "EventsOut", "Date");
```

Использование переменной:

1. В `if`: `if (key == "Events.Sheet[EventsOut].Map[Date]") { ... }`
2. В `callMacro` аргументах: `callMacro("Some.Proc", key);`
3. В шаблонах строк: `"value={key}"`

Ограничения на имя переменной (`let` и переменная цикла `for`):

1. Только идентификатор: буквы/цифры/`_`.
2. Первый символ: буква или `_` (цифра и спецсимволы запрещены).
3. Нельзя использовать ключевые слова: `if`, `for`, `callMacro`, `let`, `in`, `and`, `or`, `gt`, `lt`, `gte`, `lte`.

### Шаблоны строк

В строковых аргументах `callMacro` поддерживается подстановка `{...}`:
```text
callMacro("ex_PostProcessActions.m_AppendPostProcessFooterText", "Rows: {Events.Sheet[EventsOut].count}");
```

### Ссылки на данные

Поддерживаемые ссылки:

1. `row.column[FieldAlias]`
2. `Source.Sheet[TableAlias].row[N].column[FieldAlias]`
3. `Source.Sheet[TableAlias].lastRow.column[FieldAlias]`
4. `Source.Sheet[TableAlias].prevRow.column[FieldAlias]`
5. `Source.Sheet[TableAlias].count`
6. `Source.Sheet[TableAlias].rowCount`

### callMacro

Формат:
```text
callMacro("Module.Proc", arg1, arg2, ...);
```

Ограничение: максимум 5 аргументов.

Поддерживаемые аргументы:

1. Строковый литерал `"text"`
2. Переменная `myVar`
3. Текущая переменная строки `row`
4. Ссылки на строки: `Source.Sheet[T].row[0]`, `.lastRow`, `.prevRow`
5. Ссылки на ячейки: `...row[0].column[Date]`, `.lastRow.column[Date]`

### Пример полного скрипта

```text
/* Подсветить нужные записи и отметить regex-совпадения в Note */
for (row in Events.Sheet[EventsOut].rows) {
    if (row.column[Date] == "12.07.2025") {
        callMacro("ex_PostProcessActions.m_HighlightRow", row, "#FF0000");
    }

    let hasOrder = callMacro("ex_PostProcessActions.m_RowCellRegexIsMatch", row, "Note", "№\\s*[0-9]+");
    if (hasOrder == "True") {
        callMacro("ex_PostProcessActions.m_HighlightRowCell", row, "Note", "#404040");
        callMacro("ex_PostProcessActions.m_EmphasizeRowCellTextByRegex", row, "Note", "№\\s*[0-9]+", "#FF0000", "false");
    }
}

let mapKey = callMacro("ex_ResultRuntimeAdapter.m_BuildMapKey", "Events", "EventsOut", "Date");
if (mapKey != "") {
    callMacro("ex_PostProcessActions.m_AppendPostProcessFooterText", "Map key: {mapKey}");
}
```

## StylePipeline (page-based, без stage/workflow)

Источник конфигурации:

1. `Prototype/config/StylePipeline.xml` - слои и правила для конкретной страницы (`sheetPipeline page="..."`).
2. `Prototype/config/SheetStyles.xml` - отдельные настройки `outputStyles`/панелей (не участвует в `StylePipeline`).

Точка входа из VBA:

```vb
ex_OutputFormattingPipeline.m_ApplySheetPipeline ws
```

Полная форма:

```vb
ex_OutputFormattingPipeline.m_ApplySheetPipeline ws, resultFieldRanges, cfgStyles, rowKindRanges, activeModeKey
```

### Как движок собирает и применяет pipeline

1. Собирает runtime pipeline:
   `inline layer` (из `cfgStyles`) -> XML layers из `StylePipeline.xml` для `ws.Name`.
2. Сортирует слои по `priority` по возрастанию.
3. При одинаковом `priority` сохраняется порядок добавления.
4. Применяет только `enabled="true"` слои.
5. Применяет правила по `target` (`sheet/usedRange/range/row/column/cell`).
6. Если в `selector` есть `mode=...`, правило применяется только при совпадении active mode.

Важно:

1. `target="sheet"` применяется к расширенной used-области: `usedScope + 50%` по ширине и `usedScope + 100%` по высоте.
2. `target="usedRange"` применяется к текущей используемой области листа (`Worksheet.UsedRange`) — это прежняя логика `sheet`.
3. Расчет расширения `target="sheet"` ограничивается физическими границами листа Excel.
4. Для `target="row"` + `selector="kind=..."` нужен `rowKindRanges`; если контекст не передан, правило просто пропускается.

### Формат XML

```xml
<?xml version="1.0" standalone="yes"?>
<stylePipeline xmlns="urn:excelprototype:profiles" version="1">
  <sheetPipeline page="Dev">
    <layer id="dev-base" priority="100" enabled="true">
      <rule target="range" selector="address=A:XFD" styles="{ backColor:#202020; fontColor:#EBEBEB; }"/>
    </layer>
  </sheetPipeline>
</stylePipeline>
```

`selector`:

1. Формат: `key=value;key2=value2`.
2. Разделитель пар: `;`.
3. Разделитель key/value: `=` или `:`.

### Поддерживаемые target и selector

`target="sheet"`:

1. Без selector.
2. Пример: базовая тема листа с запасом.

`target="usedRange"`:

1. Без selector.
2. Пример: стили только в пределах текущей используемой области.

`target="range"`:

1. Обязателен `selector address=...`.
2. Пример: `address=A:XFD`, `address=A1:D20`.

`target="row"`:

1. `kind=...` (семантические группы из `rowKindRanges`), опционально `col=...`.
2. `row=...` (числа), опционально `col=...`.
3. `address=...`.

`target="column"`:

1. По данным результата (`resultFieldRanges`) через:
   `mapKey=...`, `source=...`, `table=...`, `field=...`, опционально `row=...`, `col=...`.
2. По диапазону: `address=...`.
3. По span: `col=...`, опционально `row=...`.

`target="cell"`:

1. `address=...` или
2. `row=...;col=...` (берется верхняя левая ячейка span).

Span-форматы:

1. `row=5` или `row=5:12` (также можно `5-12`).
2. `col=1:4`, `col=A:D`, `col=3`, `col=AA`.

Pattern-матчинг для `mapKey/source/table/field`:

1. Точное совпадение.
2. Wildcards: `*`, `?` (VBA `Like`, без учета регистра).

### Поддерживаемые style declarations

Список свойств:

1. `width`, `minWidth`, `maxWidth`, `autoFitColumns`
2. `overflow` (`wrap|clip|shrink`)
3. `autoHeight`, `rowHeight`, `mergeColumns`
4. `fontName`, `fontSize`, `fontBold`
5. `backColor`, `fontColor`
6. `borderColor`, `borderWeight` (`hairline|thin|medium|thick`)
7. `horizontal` (`left|center|right|fill|justify|distributed|general`)
8. `vertical` (`top|center|bottom|justify|distributed`)

Примечания:

1. Цвета: hex (`#RRGGBB`) и другие форматы, которые поддерживает `ex_XmlCore.m_TryParseColor`.
2. `width/minWidth/maxWidth` принимают положительные числа, для `width` допустим суффикс `px`.
3. `borderColor/borderWeight` применяются к внешним и внутренним границам диапазона.

### `rowKindRanges`: общий контракт

Это общий механизм движка, не специфичный для PersonalCard.

Формат:

1. `Dictionary(kindName -> Collection(rowEntry))`.
2. `rowEntry` - объект с `RowStart`/`RowEnd` или номер строки.

Пример:

```vb
Dim rowKindRanges As Object
Dim headerRows As Collection
Dim rowItem As Object

Set rowKindRanges = CreateObject("Scripting.Dictionary")
rowKindRanges.CompareMode = 1

Set headerRows = New Collection
Set rowItem = CreateObject("Scripting.Dictionary")
rowItem("RowStart") = 10
rowItem("RowEnd") = 12
headerRows.Add rowItem

Set rowKindRanges("header") = headerRows

ex_OutputFormattingPipeline.m_ApplySheetPipeline ws, resultFieldRanges, cfgStyles, rowKindRanges
```

### Примеры правил (покрытие всех сценариев)

#### 1) Базовая тема страницы (sheet + range)

```xml
<sheetPipeline page="Dev">
  <layer id="dev-base" priority="100" enabled="true">
    <rule target="sheet" styles="{ fontName:'Segoe UI'; fontSize:10; }"/>
    <rule target="range" selector="address=A:XFD" styles="{ backColor:#202020; fontColor:#EBEBEB; }"/>
  </layer>
</sheetPipeline>
```

#### 2) Row kind стили (row + kind)

```xml
<layer id="timeline-rows" priority="200" enabled="true">
  <rule target="row" selector="kind=header;col=1:8" styles="{ fontBold:true; backColor:#1F3A93; }"/>
  <rule target="row" selector="kind=section;col=1:8" styles="{ backColor:#2A2A2A; }"/>
  <rule target="row" selector="kind=content" styles="{ overflow:wrap; autoHeight:true; }"/>
</layer>
```

#### 3) Row by span/address

```xml
<layer id="manual-rows" priority="210" enabled="true">
  <rule target="row" selector="row=2:2;col=1:4" styles="{ fontBold:true; }"/>
  <rule target="row" selector="address=A20:D20" styles="{ backColor:#333333; }"/>
</layer>
```

#### 4) Column by mapKey/source/table/field

```xml
<layer id="map-columns" priority="300" enabled="true">
  <rule target="column" selector="mapKey=Daily.Sheet[DailyEvents].Map[DocNote]" styles="{ width:60; overflow:wrap; }"/>
  <rule target="column" selector="source=Daily;table=DailyEvents;field=Doc*" styles="{ fontColor:#FFD966; }"/>
</layer>
```

#### 5) Column by address/col span

```xml
<layer id="layout-columns" priority="310" enabled="true">
  <rule target="column" selector="address=B:D" styles="{ minWidth:12; maxWidth:30; }"/>
  <rule target="column" selector="col=1:2;row=1:200" styles="{ horizontal:center; }"/>
</layer>
```

#### 6) Cell target

```xml
<layer id="cells" priority="320" enabled="true">
  <rule target="cell" selector="address=A1" styles="{ fontBold:true; backColor:#404040; }"/>
  <rule target="cell" selector="row=3;col=2" styles="{ fontColor:#00FF00; }"/>
</layer>
```

#### 7) Границы (borderColor/borderWeight)

```xml
<layer id="borders" priority="330" enabled="true">
  <rule target="range" selector="address=A1:D30" styles="{ borderColor:#505050; borderWeight:thin; }"/>
</layer>
```

#### 8) Фильтр по mode

```xml
<layer id="mode-specific" priority="340" enabled="true">
  <rule target="range" selector="mode=timeline;address=A:XFD" styles="{ backColor:#1E1E1E; }"/>
  <rule target="range" selector="mode=comparing;address=A:XFD" styles="{ backColor:#FFFFFF; fontColor:#000000; }"/>
</layer>
```

#### 9) Inline styles из cfgStyles (автоматический layer)

Если в `cfgStyles(mapKey)` лежит style block:

```text
{ width:45; overflow:wrap; autoHeight:true; backColor:#252525; }
```

то движок автоматически добавит правило `target="column" selector="mapKey=<этот mapKey>"`.

### Рекомендуемый паттерн вызовов

`Startup/Dev` (базовая тема листа):

```vb
ex_OutputFormattingPipeline.m_ApplySheetPipeline ws_Dev
```

`Result page` (нужны mapKey + cfg styles + row kinds):

```vb
ex_OutputFormattingPipeline.m_ApplySheetPipeline wsOut, resultFieldRanges, cfgStyles, rowKindRanges
```

`TablesComparing` (`g_Result`, статусы по row kinds):

```vb
ex_OutputFormattingPipeline.m_ApplySheetPipeline wsResult, Nothing, Nothing, rowKindRanges, "TablesComparing"
```

Для `g_PersonTimeline` удобнее использовать обертку:

```vb
ex_OutputFormattingPipeline.m_ApplyTimelineStyleLayers wsOut, headerRows, sectionRows, resultFieldRanges, cfgStyles, partialMatchRowRanges
```
