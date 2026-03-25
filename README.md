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

Подробное описание синтаксиса, правил и примеров Пвынесено в отдельный файл: [DSL.md](DSL.md).

## Mode Pipeline (PreProcess -> Mode -> ResultLayout -> PostProcess)

Основной оркестратор: `Prototype/vba/modules/actions/ex_ModePipeline.bas`, метод `m_RunModePipeline`.

### Этапы конвейера
Плюс добавил нормализацию ; после удаления, чтобы список Query.TableRefs оставался валидным.



1. `prepare-input`
   - Если входной объект не передан, создается `obj_ScriptIOPayload`.
2. `prepare-default-query-tabrefs` (для `PersonalCard` и `MultiSources`)
   - Если в input нет `Query.TableRefs`, pipeline формирует его автоматически из `Source.*.SheetAliases`.
3. `run-preprocess`
   - Запускается `ex_PreProcessPipeline.m_Run`.
   - Источник скрипта: `Input.PreProcessScript` (`<preProcessScript .../>` в профиле).
4. `run-mode-executor`
   - Вызывается mode-метод `m_RunMode(cfg, modeInput, preProcessContext)`.
   - Mode должен вернуть объект-словарь результата.
5. `run-result-layout`
   - Запускается `ex_ResultLayoutPipeline.m_Run`.
   - Источник скрипта: `ResultLayout.Script` (`<resultLayoutScript .../>`).
   - Для `PersonalCard`/`MultiSources` скрипт `ResultLayout` обязателен.
6. `apply-result-layout-styles`
   - Применяется `ex_OutputFormattingPipeline.m_ApplySheetPipeline` с картой колонок и row kinds из `ResultLayout`.
7. `run-postprocess`
   - Запускается `ex_PostProcessPipeline.m_Run`.
   - Используется `PostProcess.Script.Implicit` (`<postProcessScript execution="Implicit" .../>`).

### Контракт объектов между этапами

1. Вход/выход скриптов (`ScriptIO payload`)
   - Тип: `obj_ScriptIOPayload`.
   - Хранит пары `key -> value` (строка или объект).
   - В DSL доступ через:
     - `callMacroObject("ex_ScriptIO.m_GetInput")`
     - `callMacroObject("ex_ScriptIO.m_CreateOutput")`
     - `callMacro("ex_ScriptIO.m_SetString", ...)`
     - `callMacro("ex_ScriptIO.m_SetObject", ...)`

2. Контекст pre-process (`preProcessContext`)
   - Тип: `Dictionary`.
   - Поля:
     - `HasScript` (`"true"`/`"false"`)
     - `Output` (объект, который пойдет в mode как `modeInput`)

3. Результат mode (`modeResult`)
   - Тип: `Dictionary`.
   - Обязательные поля:
     - `Output` (объект для следующих этапов)
     - `Worksheet` (`Worksheet`)
     - `ResultTables` (`Collection` из `obj_ResultTable`)

4. `ResultTables`
   - `obj_ResultTable`:
     - `TableRef` (например, `Events.Sheet[EventsOut]`)
     - `Rows` (`Collection` из `obj_ResultRow`)
     - `FieldMapByAlias` (alias -> mapKey)
   - `obj_ResultRow`:
     - значения колонок по alias/mapKey
     - `Kind`
     - `RowAnchorName`

### Что должны возвращать скрипты

1. `preProcessScript`
   - Должен вызвать `ex_ScriptIO.m_CreateOutput()` и заполнить output.
   - Если скрипт не создал output, pipeline завершится ошибкой.
   - Если pre-process скрипта нет, в mode уходит исходный input (fallback).

2. `resultLayoutScript`
   - Отдельный output не возвращает.
   - Работает через мутацию входного объекта:
     - читает `__ResultTables`
     - может изменить их через `ex_TableLayoutActions`
     - строит финальный лист через `ex_ResultLayoutActions`
     - записывает служебные ключи layout обратно в input.

3. `postProcessScript` (Implicit/Explicit)
   - Отдельный output не возвращает.
   - Применяет side-effects к листу/строкам/ячейкам и использует runtime-объекты (`ResultTables`, input-контекст).

### Служебные ключи input (runtime)

1. Бизнес-ключи (пример): `CommonKey`, `BaseDate`, `Query.TableRefs`, `KeysCollection`.
2. Служебные ключи pipeline/layout:
   - `__UseResultLayoutScript`
   - `__ResultTables`
   - `__ResultLayoutWorksheet`
   - `__ResultLayoutSheetName`
   - `__ResultLayoutRowKinds`
   - `__ResultLayoutFieldRanges`
3. Дополнительные mode-ключи (пример): `__Batch`, `__ResultTableRefs`.

Примечание: ключи с префиксом `__` считаются внутренними runtime-ключами.

### Поведение `Query.TableRefs`

1. Единый источник выбора таблиц для SQL-запросов.
2. Формат: `Source.Sheet[Table]; Source2.Sheet[Table2]`.
3. Парсер принимает также короткие alias таблиц, но рекомендуемый формат: полный `Source.Sheet[Table]`.
4. Если `Query.TableRefs` не задан:
   - для `PersonalCard`/`MultiSources` pipeline строит дефолт из `Source.*.SheetAliases`.
5. Если подключен `preProcessScript`:
   - для `PersonalCard`/`MultiSources` скрипт должен явно прокинуть `Query.TableRefs` в output (или изменить его).
6. Кнопка режима `State | Events | Timeline` удалена из логики выполнения.
   - Отбор таблиц теперь делается через `Query.TableRefs` в pre-process.

### Кэширование скриптов/DSL

1. `ex_ScriptSourceLoader`
   - кэш текста скрипта по контексту `(mode, profile, scriptKey, profilesFilePath)`;
   - инвалидация по `DateLastModified/Size` файла профилей и include-файлов.
2. `ex_ScriptDSL`
   - кэш распарсенных блоков по `(scriptKey + scriptText)`;
   - кэш валидации по сигнатуре доступных таблиц/полей.

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

Поддержка в `postProcessScript` вынесена в [DSL.md](DSL.md).

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

См. пример в [DSL.md](DSL.md).

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
2. Специфичный контекст страницы (какие `rowKindRanges` собрать, какие runtime слои добавить) живет в модуле страницы (`ex_ModePersonalCard`, `ex_ModeTablesComparing`, и т.д.).

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
4. Логика общая и используется как в `ex_ModePersonalCard`, так и в `ex_ModeTablesComparing` через `ex_SheetViewZoom`.

## Output Layout (gaps between result tables)

Отступы между таблицами результата теперь настраиваются через `Output.*` в профиле (а не через `StylePipeline`).

Пример:

```xml
<v key="Query.TableRefs">Main.Sheet[StateMain]; Events.Sheet[EventsOut]; Events.Sheet[EventsIn]; Daily.Sheet[DailyEvents]</v>
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
2. Gap применяется только между реально отрисованными таблицами.
3. `StylePipeline` управляет визуальным стилем строк/ячеек, но не структурным количеством пустых строк между таблицами.
