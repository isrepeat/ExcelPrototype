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
