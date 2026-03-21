# DSL (PostProcess)

Этот документ содержит актуальное описание DSL-движка для `postProcessScript`.

## Где задается скрипт

`postProcessScript` задается в профиле внутри `Prototype/config/modes/*/*Profiles.xml`.

Пример:

```xml
<postProcessScript>
    // DSL-код
</postProcessScript>
```

## Зарезервированные слова

`if`, `else`, `for`, `let`, `in`, `callMacro`, `callMacroObject`, `break`, `continue`, `return`, `and`, `or`, `gt`, `lt`, `gte`, `lte`.

## Поддерживаемые конструкции

1. `for (...) { ... }`
2. `if (...) { ... }`
3. `callMacro("Module.Proc", ...);`
4. `callMacroObject("Module.Proc", ...);`
5. `let <varName> = ...;`
6. `<varName> = ...;`

Комментарии:

1. `// <comment>`
2. `/* <comment> */`

## Типы доступа по контекстам

В DSL доступ к данным зависит от контекста. Один и тот же токен может быть допустим в `if`, но недопустим в `for`, и наоборот.

### 1) Доступ в `let` и `assign`

`let` и `assign` поддерживают три вида правой части:

1. Строковое выражение
2. `callMacro(...)`
3. `callMacroObject(...)`

Примеры:

```js
let <a> = "<text>";
let <b> = <row>.column[<FieldAlias>];
let <c> = "<prefix>: " + <b>;
let <obj> = callMacroObject("<Module>.<Proc>", <arg1>);
```

Ключевое отличие типов:

1. Переменная из строкового выражения или `callMacro(...)` считается строковой.
2. Переменная из `callMacroObject(...)` считается объектной.
3. `assign` требует совпадения типа с исходным `let` (строка к строке, объект к объекту).

То есть это ошибка:

```js
let <x> = callMacroObject("<Module>.<Proc>");
<x> = "<text>";
```

И наоборот:

```js
let <y> = "<text>";
<y> = callMacroObject("<Module>.<Proc>");
```

### 2) Доступ в `if`

В `if` поддерживается сравнение:

```js
if (<leftToken> == "<literal>") { ... }
if (<leftToken> != <rightToken>) { ... }
if (<leftToken> gt "100" and <otherToken> != "") { ... }
```

Особенности:

1. Левая часть должна быть поддерживаемым токеном.
2. Правая часть может быть либо строковым литералом, либо токеном.
3. Для `gt/lt/gte/lte` сравнение числовое, если обе стороны парсятся как числа; иначе строковое.

### 3) Доступ в `for` (target)

`for` поддерживает только три формы target:

1. `<Source>.Sheet[<Table>].rows`
2. `<rowVar>.columns`
3. `<scopeVar>.<member>`

Примеры:

```js
for (let <row> in <Source>.Sheet[<Table>].rows) { ... }
for (let <col> in <row>.columns) { ... }
for (let <item> in <scopeVar>.<member>) { ... }
```

Неподдерживаемые формы:

1. `for (let <x> in <scopeVar>.<a>.<b>) { ... }`
2. `for (let <x> in <collectionVar>) { ... }`
3. `for (let <x> in <scopeVar>.<Source>.Sheet[<Table>].rows) { ... }`

## Детально про `<scopeVar>.<member>`

Это специальная форма для member-итерации. Она не равна обычному доступу к полю объекта.

### Вариант A: `<scopeVar>` это строка результата (`row`)

Движок ищет в строке alias с именем `<member>`.

Пример:

1. У строки `<ownerRow>` есть поле `Projects` со значением `Data.Sheet[Projects]`.
2. Тогда target `<ownerRow>.Projects` резолвится как строки `Data.Sheet[Projects]`.

DSL:

```js
for (let <projectRow> in <ownerRow>.Projects) {
    callMacro("<Module>.<Proc>", <projectRow>);
}
```

Чтобы это работало, в runtime-строке должен существовать alias `Projects` с корректным table ref.

Практический вариант: если строка содержит alias `SrcTableRows = Data.Sheet[Projects]`, то можно писать:

```js
for (let <projectRow> in <ownerRow>.SrcTableRows) {
    // ...
}
```

### Вариант B: `<scopeVar>` это объектный контейнер

Если `<scopeVar>` указывает на объект-контейнер, то `<member>` должен резолвиться в типизированное значение, которое может быть:

1. Таблицей (`obj_ResultTable`)
2. Коллекцией объектов (`Collection`)
3. Текстовым table ref (`<Source>.Sheet[<Table>]`)

### Важное ограничение по коллекциям

Если `<scopeVar>.<member>` резолвится в `Collection`, элементы должны быть объектами (обычно строки результата). Коллекция примитивных строк (`"<text>"`) не является штатным сценарием для `for`.

## Матрица: где что использовать

1. Нужен обход строк таблицы: используйте `<Source>.Sheet[<Table>].rows`.
2. Нужен обход колонок текущей строки: используйте `<rowVar>.columns`.
3. Нужен обход «связанной таблицы» от строки-контекста: используйте `<scopeVar>.<member>`, где `<member>` хранит table ref.
4. Нужен доступ к скаляру: используйте `if`/строковые выражения с токенами.
5. Нужна объектная переменная: `let <varName> = callMacroObject(...)`.

## Циклы

### 1) По строкам таблицы

```js
for (let <row> in <Source>.Sheet[<Table>].rows) {
    // ...
}
```

### 2) По колонкам строки

```js
for (let <col> in <row>.columns) {
    // ...
}
```

Для `<col>` доступны члены:

1. `<col>.alias` (или `<col>.name`)
2. `<col>.mapKey`
3. `<col>.value`

### 3) По member-цели

```js
for (let <row> in <scopeVar>.<member>) {
    // ...
}
```

Частый кейс в ReportCreation:

```js
let <input> = callMacroObject("ex_ScriptIO.m_GetInput");
let <batch> = callMacroObject("ex_ScriptIO.m_GetObject", <input>, "__Batch");

for (let <keyResult> in <batch>.keysResults) {
    for (let <stateRow> in <keyResult>.Main) {
        if (<stateRow>.FIO == <keyResult>.Key) {
            callMacro("ex_Helpers.m_EmphasizeRowCellTextByRegex", <stateRow>, "FIO", "^\\S+", "#FFD966", "false");
        }
    }
}
```

## Условия

Поддерживаются операторы:

1. `==`
2. `!=`
3. `gt`, `lt`
4. `gte`, `lte`

Логические связки:

1. `and`
2. `or`

Пример:

```js
if (<row>.column[EventNum] gt "20000" and <row>.column[Date] != "") {
    callMacro("ex_PostProcessActions.m_HighlightRow", <row>, "#FF0000");
}
```

## Переменные

Создание:

```js
let <mapKey> = callMacro("ex_ResultRuntimeAdapter.m_BuildMapKey", "Events", "EventsOut", "Date");
```

Использование:

1. В `if`: `if (<mapKey> != "") { ... }`
2. В аргументах `callMacro`: `callMacro("<Module>.<Proc>", <mapKey>);`
3. В строковых шаблонах: `"value={<mapKey>}"`

Ограничения для имен переменных (`let` и loop var):

1. Только идентификатор (`[A-Za-z_][A-Za-z0-9_]*`).
2. Нельзя использовать зарезервированные слова.

## Шаблоны строк

В строковых аргументах поддерживается подстановка `{...}`.

```js
callMacro("ex_PostProcessActions.m_AppendPostProcessFooterText", "Rows: {<Source>.Sheet[<Table>].count}");
```

## Ссылки на данные

Поддерживаемые ссылки:

1. `<row>.column[<FieldAlias>]`
2. `<Source>.Sheet[<TableAlias>].row[<N>].column[<FieldAlias>]`
3. `<Source>.Sheet[<TableAlias>].lastRow.column[<FieldAlias>]`
4. `<Source>.Sheet[<TableAlias>].prevRow.column[<FieldAlias>]`
5. `<Source>.Sheet[<TableAlias>].count`
6. `<Source>.Sheet[<TableAlias>].rowCount`

## callMacro / callMacroObject

Форматы:

```js
callMacro("<Module>.<Proc>", <arg1>, <arg2>, ...);
callMacroObject("<Module>.<Proc>", <arg1>, <arg2>, ...);
```

Ограничение: максимум 5 аргументов.

Типы аргументов:

1. Строковый литерал: `"<text>"`
2. Переменная: `<varName>`
3. Текущая переменная строки: `<row>`
4. Ссылки на строки: `<Source>.Sheet[<T>].row[0]`, `.lastRow`, `.prevRow`
5. Ссылки на ячейки: `...row[0].column[<Field>]`, `.lastRow.column[<Field>]`

## Полный пример

```js
/* Подсветить нужные записи и отметить regex-совпадения в Note */
for (let <row> in Events.Sheet[EventsOut].rows) {
    if (<row>.column[Date] == "12.07.2025") {
        callMacro("ex_PostProcessActions.m_HighlightRow", <row>, "#FF0000");
    }

    let <hasOrder> = callMacro("ex_PostProcessActions.m_RowCellRegexIsMatch", <row>, "Note", "№\\s*[0-9]+");
    if (<hasOrder> == "True") {
        callMacro("ex_PostProcessActions.m_HighlightRowCell", <row>, "Note", "#404040");
        callMacro("ex_PostProcessActions.m_EmphasizeRowCellTextByRegex", <row>, "Note", "№\\s*[0-9]+", "#FF0000", "false");
    }
}

let <mapKey> = callMacro("ex_ResultRuntimeAdapter.m_BuildMapKey", "Events", "EventsOut", "Date");
if (<mapKey> != "") {
    callMacro("ex_PostProcessActions.m_AppendPostProcessFooterText", "Map key: {<mapKey>}");
}
```

## Интеграция с ResultTemplatesParser

Пример pipeline в DSL:

```js
let <resultTemplatesRelPath> = "config\\modes\\PersonalCard\\PersonalCardResultTemplates.xml";
let <txt> = callMacro("ex_ResultTemplatesParser.m_GetTemplateText", "HospitalBrown", <resultTemplatesRelPath>);
<txt> = callMacro("ex_ResultTemplatesParser.m_ReplacePlaceholder", <txt>, "Hospital", "{<row>.column[Hospital]}");
<txt> = callMacro("ex_ResultTemplatesParser.m_ReplacePlaceholder", <txt>, "FIO", "{<row>.column[FIO]}");
<txt> = callMacro("ex_ResultTemplatesParser.m_ResolveTemplate", <txt>);
callMacro("ex_PostProcessActions.m_AppendToSinglePostProcessFooterText", <txt>, "\n\n");
```
