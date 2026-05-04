# PersonalCard SQL Pipeline (PrototypeNew)

## 1) Что было в Prototype

Старый `Prototype` уже имел рабочий конвейер:

1. `PreProcess` (`ex_PreProcessPipeline`)
2. `Mode executor` (`ex_ModePersonalCard.m_RunMode`)
3. `ResultLayout` (`ex_ResultLayoutPipeline`)
4. `PostProcess` (`ex_PostProcessPipeline`)

Ключевые свойства старой реализации:

- Конфиг профиля содержал `Source.*`, `*.Sheet[...]`, `Map[...]`, `Query.TableRefs`.
- SQL к Excel выполнялся через ADO (`ADODB.Connection` + `Recordset`).
- Был preflight/валидация конфигурации и детальные ошибки.
- Были кэши для сопоставления tableRef/headers/DSL-планов.
- `fetchAdditionalData` (DSL) расширял строки виртуальными данными.

## 2) Что уже есть в PrototypeNew и можно переиспользовать

- Внешние profile-файлы в `modes/<Mode>/<Mode>Profiles.xml`.
- Runtime источники (`obj_PageRuntimeSources`) для передачи данных в UI.
- Table-рендер контрола (`TableList`, `TableSingle`) уже готов.
- Config runtime уже умеет редактировать и сохранять профиль (`obj_ConfigControlVM` + `ex_PageMainActions`).

Это значит, что нужно добавить серверную часть данных (query/transform pipeline), а UI-рендер уже в хорошей форме.

## 3) Предложенная улучшенная архитектура для PrototypeNew

### Stage A: Profile -> QueryPlan

Новый модуль, например `ex_QueryPlanBuilder`:

- Принимает плоский список `obj_ConfigEntry`.
- Строит typed-план:
  - `Sources` (путь, resolver, sheet aliases)
  - `Tables` (type, key, fields/map, sort, markers/range)
  - `Execution` (`Query.TableRefs`, strict flags)
- Делает fail-fast валидацию и возвращает понятные ошибки `MsgBox`.

### Stage B: QueryPlan -> SourceSessions

Новый модуль, например `ex_SourceSessionManager`:

- Открывает подключения к источникам (с кэшем по `SourceAlias + resolvedPath + mtime`).
- Дает единый API `GetConnection(sourceAlias)`.
- Гарантирует `Dispose`/close всех соединений в `Finally`.

### Stage C: SQL + Fetch DSL

Новый модуль, например `ex_QueryExecutor`:

- Для каждого `TableRef`:
  - резолвит ADO table object/range,
  - строит `SELECT` (по mapped headers),
  - выполняет exact/partial match стратегии,
  - нормализует типы (даты/empty/null).
- Если есть `*.FetchDsl`, применяет post-fetch virtual rows/fields.

Улучшение относительно Prototype:

- единый контракт на результат каждой таблицы (`obj_TableDynamic` + metadata),
- единый object для diagnostics (`QueryDiagnostics`) с SQL текстом, row count, elapsed.

### Stage D: Domain Projection

Новый модуль, например `ex_PersonalCardProjector`:

- Преобразует query-result в UI-модели:
  - `Collection(obj_TableDynamic)`
  - optional `obj_Banner`
- Проставляет row kinds / section banners централизованно.

### Stage E: Runtime publish + render

- Публикация в `pageBase.RuntimeSources`:
  - `RuntimeItems.PersonalCard.Tables`
  - `RuntimeObjects.PersonalCard.Banner` (опционально)
- Ререндер страницы стандартным `rt_PageManager.m_RenderPage`.

## 4) Какие контракты зафиксировать сразу

1. Ключи runtime:
- `RuntimeItems.PersonalCard.Tables`
- `RuntimeObjects.PersonalCard.Banner`
- `RuntimeObjects.PersonalCard.Diagnostics`

2. Ключи профиля (минимум):
- `Source.<Alias>.FilePath`
- `Source.<Alias>.FileResolver` (optional)
- `Source.<Alias>.SheetAliases`
- `<Alias>.Sheet[<Table>].Type`
- `<Alias>.Sheet[<Table>].Key`
- `<Alias>.Sheet[<Table>].FieldsAliases`
- `<Alias>.Sheet[<Table>].Map[Field]`
- `Query.TableRefs`
- `<Alias>.Sheet[<Table>].FetchDsl` (optional)

3. Типы ошибок:
- `CONFIG_ERROR`
- `SOURCE_ERROR`
- `SQL_ERROR`
- `DSL_ERROR`
- `RENDER_ERROR`

## 5) Почему это лучше старой версии

- Разделение ответственности: profile parse, SQL, projection, render отдельно.
- Меньше глобального состояния и проще тестировать каждый stage отдельно.
- Переиспользование существующего UI runtime из `PrototypeNew` без смешивания с SQL логикой.
- Прозрачная диагностика: легче понять, где именно сломалось.

## 6) Рекомендуемый следующий шаг в коде

1. Сделать `ex_QueryPlanBuilder` + typed DTO-классы плана.
2. Добавить первый `ex_QueryExecutor` только для exact-match по `CommonKey`.
3. Отдать результат в `RuntimeItems.PersonalCard.Tables` и отрендерить через существующий `TableList`.
4. После стабилизации добавить `FetchDsl` stage.
