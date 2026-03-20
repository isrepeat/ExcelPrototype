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

## ResultTemplatesParser

Модуль: `Prototype/vba/modules/common/ex_ResultTemplatesParser.bas`

### Что делает

Парсер закрывает задачу финальной сборки текстового блока по XML-шаблону в `postProcessScript`:

1. Загружает `<template id="..."><text><![CDATA[...]]></text></template>`.
2. Подставляет бизнес-плейсхолдеры (`{Hospital}`, `{FIO}`, ...).
3. Применяет опциональные форматтеры (`{FIO|upper}`, `{Rank|genitive}`, ...).
4. Выполняет финальный проход по зарезервированным токенам даты (`{#dd}`, `{#dd+N}`, `{#dd-N}`).

### Публичные методы

1. `m_GetTemplateText(templateId, resultTemplatesRelPath)` - берет `text` по `template/@id`, нормализует переводы строк.
2. `resultTemplatesRelPath` обязателен:
   - если путь не передан или передан пустым, функция возвращает ошибку.
3. `m_ReplacePlaceholder(sourceText, placeholderName, replacementText)` - заменяет:
   - простой токен `{Name}`
   - токен с форматтером `{Name|formatter}`
   - цепочку форматтеров `{Name|action1|action2}`
   - форматтеры с аргументами `{Name|truncate:20|replace:foo,bar}`
4. `m_ResolveTemplate(sourceText, [baseDateText])` - делает финальный проход:
   - `{#dd}` = день базовой даты
   - `{#dd+N}` / `{#dd-N}` = день со смещением
   - `{#_}` / `#_` = убрать перенос строки вокруг токена (склеить строки)
   - если `baseDateText` не передан, используется `Date` (текущий день)
   - условные блоки `{#if ...}...{#endif}`

### Форматтеры плейсхолдеров

Поддерживаются:

1. `upper`
2. `lower`
3. `capitalize`
4. `firstchar`
5. `upperFirstWord`
6. `upperFirstLetter`
7. `genitive`
8. `accusative`
9. `dative` (укр. давальний: "кому/чому")
10. `lowerFirstWord`
11. `lowerFirstLetter`
12. `truncate:N`
13. `replace:from,to`
14. `calendarDaysUa` (например, `1 календарний день`, `2 календарних дня`, `5 календарних днів`)
15. `surnameInitials` (из `Прізвище Ім'я По батькові` в `Прізвище І.П.`)
16. `fioSurname` (только фамилия из ФИО, в исходном регистре)
17. `fioInitials` (только инициалы из ФИО, например `І.П.`)

Примечание: `genitive`, `accusative`, `dative` реализованы в `ex_MorphUaLite` и ориентированы на украинские формы.

Правила pipeline:

1. Форматтеры применяются слева направо: `{Field|accusative|lowerFirstLetter}`.
2. Пробелы вокруг `|` игнорируются.
3. Для `replace` обязательны оба аргумента (`from,to`), `from` не может быть пустым.

Если форматтер неизвестен, модуль добавляет диагностическую строку в начало результата.

Поддержка в `postProcessScript`:

1. В строковых аргументах `callMacro` можно использовать тот же синтаксис форматтера:
   - `{row.column[Rank]|accusative}`
   - `{row.column[FIO]|genitive}`
2. Форматирование выполняется через `ex_ResultTemplatesParser.m_FormatValue`.

### Условные блоки в шаблоне

Поддерживается синтаксис:

1. `{#if ReturnToDutyLine}...{#endif}` - условие по значению плейсхолдера.
2. `{#if IsAssignDuty}...{#endif}` - условие по флагу (`"true"` / `"false"` строкой).
3. `{#if AdditionalWayDays == 1}...{#endif}` - числовое сравнение (`==`, `!=`, `>`, `<`, `>=`, `<=`).
4. Поддержан префикс отрицания `#not`, например: `{#if #not IsAssignDuty}`.

Правила вычисления условия:

1. `"false"` (без учета регистра) -> `false`
2. пустая строка -> `false`
3. `"true"` -> `true`
4. любая другая непустая строка -> `true`
5. Для числовых сравнений обе части должны парситься как числа.

Дополнительно:

1. Поддержаны вложенные `if`-блоки.
2. Если плейсхолдер заменяется через `m_ReplacePlaceholder`, выражения в `{#if ...}` обновляются тем же значением.
3. Это позволяет делать сравнение плейсхолдера с числовым литералом после подстановки, например:
   - шаблон: `{#if AdditionalWayDays == 2}...{#endif}`
   - подстановка: `m_ReplacePlaceholder(..., "AdditionalWayDays", "2")`
   - итог: условие вычисляется как `true`.

### Диагностика в тексте результата

При ошибках форматирования/резолва модуль не прерывает сборку шаблона, а добавляет первую строку вида:

`[TEMPLATE ERROR] <operation>: [<source> #<number>] <description>`

Эта строка вставляется в самое начало итогового текста.

### Поддержка неразрывного пробела

Внутренний trim/поиск первого непробельного символа учитывает:

1. обычный пробел
2. tab/newline
3. `NBSP` (`U+00A0`)
4. `NARROW NBSP` (`U+202F`)

Это важно для шаблонов, скопированных из Word/почты, где вместо обычных пробелов часто попадает `NBSP`.

### Рекомендуемый pipeline в DSL

```text
let resultTemplatesRelPath = "config\\modes\\PersonalCard\\PersonalCardResultTemplates.xml";
let txt = callMacro("ex_ResultTemplatesParser.m_GetTemplateText", "HospitalBrown", resultTemplatesRelPath);
txt = callMacro("ex_ResultTemplatesParser.m_ReplacePlaceholder", txt, "Hospital", "{row.column[Hospital]}");
txt = callMacro("ex_ResultTemplatesParser.m_ReplacePlaceholder", txt, "FIO", "{row.column[FIO]}");
txt = callMacro("ex_ResultTemplatesParser.m_ResolveTemplate", txt);
callMacro("ex_PostProcessActions.m_AppendToSinglePostProcessFooterText", txt, "\n\n");
```

## StylePipeline (page-based, universal apply)

Источник конфигурации:

1. `Prototype/config/StylePipeline.xml` - декларативные `layer/rule` по страницам (`sheetPipeline page="..."`).
2. `Prototype/config/modes/*/*UI.xml` - конфигурация control panel по режимам; палитры/базовые параметры подставляются кодом и runtime layers.

Единая точка входа из VBA:

```vb
ex_OutputFormattingPipeline.m_ApplySheetPipeline ws
```

Полная форма:

```vb
ex_OutputFormattingPipeline.m_ApplySheetPipeline _
    ws, _
    resultFieldRanges, _
    cfgStyles, _
    rowKindRanges, _
    activeModeKey, _
    autoHeightOnly, _
    runtimeLayers
```

Где:

1. `resultFieldRanges` - карта выходных полей (`MapKey/ColumnIndex/RowStart/RowEnd`) для `target="column"` по `mapKey/source/table/field`.
2. `cfgStyles` - словарь inline style блоков из конфига (автоматически превращаются в inline-layer).
3. `rowKindRanges` - словарь семантических строк для `target="row"` + `selector kind=...`.
4. `activeModeKey` - ключ режима для `selector mode=...` (если пусто, берется активный профиль).
5. `autoHeightOnly` - применить только декларации `autoHeight` (используется в спец-сценариях).
6. `runtimeLayers` - `Collection(obj_StyleLayer)`, добавляемые в pipeline на лету.

### Как теперь формируется pipeline

1. Базовый pipeline строится в движке так:
   `inline layer (cfgStyles)` -> `XML layers (StylePipeline.xml, page = ws.Name)`.
2. Затем вызывающий модуль может добавить `runtimeLayers`.
3. Все слои сортируются по `priority` (возрастание), при равном приоритете сохраняется порядок добавления.
4. Применяются только `enabled=true`.

Ключевая архитектурная договоренность:

1. В `ex_OutputFormattingPipeline` остается только универсальный apply API.
2. Специфичный контекст страницы (какие `rowKindRanges` собрать, какие runtime слои добавить) живет в модуле страницы (`ex_PersonTimeline`, `ex_TableComparing`, и т.д.).

### `runtimeLayers`: зачем и как использовать

`runtimeLayers` нужны, когда стиль зависит от runtime-геометрии/данных и это неудобно держать в статичном XML.

Примеры:

1. Контрольная панель (координаты и размеры известны только после рендера).
2. Warning banners с динамическими диапазонами.

Пример интеграции:

```vb
Dim runtimeLayers As Collection
Dim runtimeLayer As obj_StyleLayer

Set runtimeLayers = New Collection

Set runtimeLayer = ex_OutputPanel.m_CreateRuntimeLayer(wsOut, outputStyle, "runtime-control-panel", 800)
If Not runtimeLayer Is Nothing Then runtimeLayers.Add runtimeLayer

Set runtimeLayer = ex_Messaging.m_CreateWarningBannersRuntimeLayer(wsOut, pendingWarningBanners, "runtime-warning-banners", 850)
If Not runtimeLayer Is Nothing Then runtimeLayers.Add runtimeLayer

ex_OutputFormattingPipeline.m_ApplySheetPipeline wsOut, resultFieldRanges, cfgStyles, rowKindRanges, vbNullString, False, runtimeLayers
```

### AutoFit / autoHeight поведение

1. `autoHeight` теперь применяется отложенно: движок собирает финальное состояние по строкам на всем проходе pipeline и выполняет `Rows.AutoFit` один раз в конце.
2. Принцип: `last rule wins` (`autoHeight:true` может быть отключен поздним `autoHeight:false`).
3. `autoFitColumns` сейчас применяется в момент обработки правила (не отложенно).

### Scope и target

Важно:

1. `target="sheet"` применяется как ограниченный колонками диапазон `A:AZ` (аналогично `target="range" selector="address=A:AZ"`).
2. `target="usedRange"` применяется к `Worksheet.UsedRange`.
3. `target="row"` + `selector kind=...` требует `rowKindRanges`; если он не передан, такие правила просто пропускаются.

Поддерживаемые `target`: `sheet`, `usedRange`, `range`, `row`, `column`, `cell`.

### XML пример

```xml
<?xml version="1.0" standalone="yes"?>
<stylePipeline xmlns="urn:excelprototype:profiles" version="1">
  <sheetPipeline page="Dev">
    <layer id="dev-base" priority="100" enabled="true">
      <rule target="sheet" styles="{ backColor:#202020; fontColor:#EBEBEB; }"/>
    </layer>
    <layer id="dev-grid" priority="110" enabled="true">
      <rule target="range" selector="address=A1:D200" styles="{ borderColor:#505050; borderWeight:thin; }"/>
    </layer>
  </sheetPipeline>
</stylePipeline>
```

### `selector` формат

1. Формат: `key=value;key2=value2`.
2. Разделитель пар: `;`.
3. Разделитель key/value: `=` или `:`.
4. `row` span: `5`, `5:12`, `5-12`.
5. `col` span: `1:4`, `A:D`, `3`, `AA`.

### Поддерживаемые declarations

1. `width`, `minWidth`, `maxWidth`, `autoFitColumns`
2. `overflow` (`wrap|clip|shrink`)
3. `autoHeight`, `rowHeight`, `mergeColumns`
4. `fontName`, `fontSize`, `fontBold`
5. `backColor`, `fontColor`
6. `borderColor`, `borderWeight` (`hairline|thin|medium|thick`)
7. `horizontal` (`left|center|right|fill|justify|distributed|general`)
8. `vertical` (`top|center|bottom|justify|distributed`)

### `rowKindRanges` контракт (общий)

Формат:

1. `Dictionary(kindName -> Collection(rowEntry))`.
2. `rowEntry` может быть:
   `Long` (одна строка) или объект с `RowStart/RowEnd`.

Пример:

```vb
Dim rowKindRanges As Object
Dim headerRanges As Collection
Dim rowEntry As Object

Set rowKindRanges = CreateObject("Scripting.Dictionary")
rowKindRanges.CompareMode = 1

Set headerRanges = New Collection
Set rowEntry = CreateObject("Scripting.Dictionary")
rowEntry("RowStart") = 10
rowEntry("RowEnd") = 12
headerRanges.Add rowEntry

Set rowKindRanges("header") = headerRanges
```

### Рекомендуемые вызовы

`Dev`:

```vb
ex_OutputFormattingPipeline.m_ApplySheetPipeline ws_Dev
```

`Result page` (данные + kinds + runtime):

```vb
ex_OutputFormattingPipeline.m_ApplySheetPipeline wsOut, resultFieldRanges, cfgStyles, rowKindRanges, vbNullString, False, runtimeLayers
```

`TablesComparing`:

```vb
ex_OutputFormattingPipeline.m_ApplySheetPipeline wsResult, Nothing, Nothing, rowKindRanges, "TablesComparing"
```

## Result Page Zoom (profiles)

Для профилей режимов (PersonalCard/TablesComparing) можно задать дефолтный zoom результата атрибутом профиля:

```xml
<profile name="Test2" resultZoom="115">
```

Поведение:

1. Если результатный лист создается впервые, применяется `resultZoom` активного профиля.
2. Пока лист жив (не удален), сохраняется текущий zoom листа; in-memory cache используется как fallback.
3. Повторный Search/Run не переустанавливает профильный zoom для уже существующей страницы.
4. Логика общая и используется как в `ex_PersonTimeline`, так и в `ex_TableComparing` через `ex_SheetViewZoom`.

## Output Layout (gaps between result tables)

Отступы между таблицами результата теперь настраиваются через `Output.*` в профиле (а не через `StylePipeline`).

Пример:

```xml
<v key="Output.Sheets">StateMain; EventsOut; EventsIn; DailyEvents</v>
<v key="Output.Layout.Gap.Default">1</v>
<v key="Output.Layout.Gap.AfterType[Events]">1</v>
<v key="Output.Layout.Gap.Between[EventsOut->EventsIn]">0</v>
```

Поддерживаемые ключи:

1. `Output.Layout.Gap.Between[AliasA->AliasB]` - самый точный приоритет для конкретной пары таблиц.
2. `Output.Layout.Gap.BetweenType[State->Events]` - правило по типам (`State`, `Events`).
3. `Output.Layout.Gap.After[AliasA]` - отступ после конкретной таблицы.
4. `Output.Layout.Gap.AfterType[Events]` - отступ после типа таблицы.
5. `Output.Layout.Gap.Default` - общий дефолт.

Приоритет применения: `Between` -> `BetweenType` -> `After` -> `AfterType` -> `Default` -> встроенный fallback `1`.

Важные детали:

1. Значения gap должны быть целыми `>= 0`, иначе рендер завершится с ошибкой валидации ключа.
2. Gap применяется только между реально отрисованными таблицами с учетом режима (`StateTableOnly`/`EventsTableOnly`).
3. `StylePipeline` управляет визуальным стилем строк/ячеек, но не структурным количеством пустых строк между таблицами.
