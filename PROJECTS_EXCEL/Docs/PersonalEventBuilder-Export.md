# Подготовка и экспорт PersonalEventBuilder

## EventsHistory export pipeline

The main page now separates `SaveHistory` (form to Excel) from `ExportWord` (saved Excel events to Word). The legacy per-event Word exporter remains for existing callers; the main page no longer uses its Custom XML event store. Movement export remains independent.

`history.path`, `history.sheet`, `history.table` and `history.fields` configure the history workbook and field-to-header mappings. Column positions are irrelevant. PROD uses `EventsHistory.xlsx`; TEST requires a separate `EventsHistory.test.xlsx` and never substitutes the production file. Missing mapped columns are appended on the first successful save. Existing demonstration records are examples, not migrated production events: new required snapshot columns, especially `Signatory`, must be populated before generating those records.

All exported row cells use the Excel text format. Dates serialize as `dd.mm.yyyy`, Boolean flags as `true` or `false`, and identifiers retain their literal text. History generation parses these values explicitly. It uses saved personnel, institution and report-author snapshots without rereading those external directories. Document naming and the Word template still follow the current configuration; the signatory comes from the saved event.

Saving new events matches `history.identityFields` (order number/date, personnel ID, event ID and event date). Identical rows return `SKIP`; changed rows update immediately and retain their `RowId`, `RowOrder` and existing child rows. Export results and failures appear in the page status without modal dialogs. A history workbook with unsaved edits, read-only access, protection or the 1904 date system is rejected.

`ShowHistory` displays the accepted order's Event rows in a separate compact table below personnel candidates. The markup binds its `visibility` to `Form.HistoryVisibility`, initialized to `Collapsed`; only a successful explicit ShowHistory action changes it to `Visible`. Its projection is configured by `history.list.columns.*`. Confirm a row with Ctrl+Enter to highlight it, load its stored fields and optional scalar context into the form, and switch the save caption to Update. Loading preserves the candidate list and its visibility, disables autofill and remembers its previous preference. The user can enable autofill again and select a different candidate while editing the same event; exiting editing restores the previous preference. Editing keeps the selected `RowId` even when the person, event date or accepted order changes. Updates reject a conflicting event identity or a stored row changed since selection. Saving refreshes the list and retains selection. New event (the reset action) clears the edit identity, optional context, highlight and displayed history rows, then collapses the table without deleting stored events. Closing the table only hides it and does not end editing. The edit mode is displayed separately from the action status; child rows remain untouched.

Shared control markup supports `visibility="Visible|Collapsed"`, including bindings. A collapsed control measures zero rows/columns and skips arrangement and rendering. Hiding an active control disposes its old routes before its area is reused. Before recreating a control, its original local row/column are restored: arrangement writes absolute sheet coordinates into the mutable node, which must not become offsets on reconfiguration. Visibility changes use partial layout refresh. ShowHistory explicitly flushes after publishing its rows and visibility. Deploy both `obj_UiControlFactory` and `obj_UiControlElement` with this markup.

Workspace candidates and order events share the vertical RecordPanels stack. Both start collapsed and expose a close button in the worksheet column immediately to their left through `closeCommand` and `closeCaption`. Closing only changes visibility; data and edit identity remain available. A primary or secondary lookup reopens candidates, and ShowHistory reopens order events. New event collapses both lists. The wrapper includes the close-button column in its render area and removes the button's shape and route on collapse or disposal. The generated `btn_Close_Candidates` and `btn_Close_HistoryEvents` names fit the 31-character limit. Deploy the MainPage markup, both wsConfig profiles, page/controller classes and shared control factory/element together.

`Event` rows have no parent. `Report` and `Document` rows refer directly to an event. `TVO` rows may refer to an event or another TVO row; the parent identifies whose duties are assigned. `RowOrder` is a positive integer ordering siblings; TVO chains render parent-first. Duplicate IDs, unknown types, missing parents and cycles stop generation. Additional reports use `ReportName`, author snapshot columns, `IncomingNumber` and `IncomingDate`; documents use `Document`, `DocumentNumber` and `DocumentDate`; TVO uses personnel snapshots including `RankDative` and `FIODative`. Additional rows are currently entered in the table. Nonempty collection inputs supplied programmatically through the form context are rejected rather than silently discarded.

Word generation selects the accepted order, prepares all sections in memory, accumulates ordinary events and applies the existing `ToBusinessTrip` grouping. Its common basis includes additional reports and documents. TVO remains attached to each person inside a grouped section. Each managed anchor present in the template is replaced, including empty sections; a populated event requires its configured anchor. The document is rebuilt from the template, saved to a temporary DOCX and published only after closing Word and checking that an existing output has not changed. Manual edits in the previous output are not carried forward.

Source edits do not deploy VBA or refresh workbook configuration. Deploy the new `ex_PEB_History` module, the modified classes, UI, both configurations and Word templates together before runtime verification. The existing runtime tests have not been migrated to this pipeline.

**Навигация:** [Главная](../../README.md) → [PROJECTS_EXCEL](../README.md) → [Документация](README.md) → Подготовка и экспорт PersonalEventBuilder

## Содержание

- [Уровни шаблонов](#уровни-шаблонов)
- [Конфигурация](#конфигурация)
- [Сопоставления режимов](#сопоставления-режимов)
- [Поддерживаемый DSL](#поддерживаемый-dsl)
- [Границы реализации](#границы-реализации)
- [Постоянная группировка Word](#постоянная-группировка-word)
- [Статические проверки](#статические-проверки)


Синтаксис Word/Movement, вычисляемые поля, группировка и маршруты событий описаны в [руководстве по DSL](PersonalEventBuilder-TemplateDSL.md).

Страница готовит одно событие по редактируемой форме. `PreviewWord` выводит текст в H4:Q10, под принятым приказом и над кадровой формой. Каждый щелчок снимает свежие данные и заново читает XML-шаблоны. Word не открывается, путь документа не определяется. Смена режима или ошибка нового превью очищает старый текст.

Экспорт Word отключён через `export.Word.enabled=false`; кнопки на странице нет, сохранённая команда и экспортёр проверяют тот же флаг. Movement отключён до настройки назначения. Импорт в рабочую книгу и runtime-проверки требуют явного разрешения.

## Уровни шаблонов

1. `workbook/PersonalEventBuilder/PersonalEventBuilderWordResultTemplates.xml` задаёт текст пунктов, повторяемые DSL-блоки и группировку. Тексты не дублируются в wsConfig.
2. Шаблон DOCX задаёт оформление документа и якоря вставки. В превью листа он не участвует.

`PersonalEventBuilderEventRoutes.xml` выбирает независимые текстовые шаблоны Word и Movement. У обоих маршрутов есть только `template`; операции Movement определяет нативный план по EventId. Событие может объявлять один или оба экспортёра. Отсутствие тега отключает соответствующий экспорт; неверные ссылки, повторы тегов и событие без маршрутов — ошибки. Неиспользуемый реестр EventBody удалён; вставка DOCX использует настроенные маркеры Begin/End.

`PersonalEventBuilderMovementTemplates.xml` — текстовый документ `resultTemplates` с локальными `sharedTemplates` и именованными `template`, содержащими блоки `column configKey="export.Movement.headers.ArrivalBasis"`, каждый со своим `text`. Поддерживается тот же DSL подстановок, условий, включений и форматтеров. В XML нет поиска строк, операций, статусов, схем входа и команд записи таблицы. Тексты — окончательные значения ячеек без добавления или удаления префикса. Формулировка про билет остаётся локальной и редактируется независимо.

Единый снимок контекста содержит форму, результаты настроенных поисков, алиасы и вычисляемые значения. DSL использует `{[Name]}`, нативный код — ключ `Name`. Экспериментальные `Ref.Form.*` удалены. Необязательные скаляры, даты и метаданные объявлены в `export.shared.inputs`; `Data.WordContext` остаётся транспортом входа для обоих экспортёров.

Колонки Movement объявляются раздельно: прибытие ссылается на `export.Movement.headers.ArrivalBasis`, выбытие — на `export.Movement.headers.DepartureBasis`, переход — на обе. Загрузчик разрешает обязательные `configKey` через wsConfig, затем проверяет разрешённые колонки и повторы итоговых заголовков. Общие блоки остаются текстовыми. `fn_Texts` возвращает определения колонка → DSL; подготовка независимо формирует `MovementColumnTexts`. `fn_WriteValue` выбирает текст по фактическому заголовку записи, в том числе отдельно при закрытии и открытии перехода. Недоступные колонки и отсутствие обязательного основания отклоняются до изменения книги. Единого значения MovementBasis больше нет: это лишь внутренний маркер нативного сопоставления. Начальные тексты переходов продублированы без изменения поведения, чтобы редактировать их независимо.

`ex_PEB_MovementTemplates` проверяет и кеширует текстовые определения и локальные include. `ex_PEB_MovementPlan` хранит пять профилей колонок и 35 планов событий; восемь статусов берутся из `export.Movement.status.<StatusId>`. `obj_PEB_ExptrMovement` выполняет добавление, закрытие, переход, сопоставление строк, проверку дублей/хронологии, защищённую запись и откат. Проверка сценариев читает те же сопоставления. Dispose маршрутов очищает оба кеша. Новое бизнес-событие или изменение поведения таблицы требует кода; новая формулировка основания — только XML.

`obj_PEB_WordTemplates` загружает XML, проверяет корень, версию и повторы имён, предоставляет шаблон и общие блоки. DTD и внешние ссылки отключены. `obj_PEB_DslRenderer` заменяет `obj_PEB_Template` и работает независимо от автоматизации Word. `obj_PEB_ExportService` снимает форму, выполняет поиски, сопоставляет алиасы, выбирает XML-шаблон и публикует текст через существующую привязку превью.

## Конфигурация

Пути, схемы, общие поставщики данных, окружения и сообщения импортируются из `config/PersonalEventBuilder/wsConfig.txt` в wsConfig/tbConfig. Имена колонок Movement обязательны в `export.Movement.headers.<HeaderId>`; назначение медицинской роты — в `export.Movement.destinations.MedicalCompany`. Планы читают их через `fn_Config`, без зашитых запасных значений. Обновляйте выбранный TEST/PROD-конфиг вместе с VBA. XML-колонки обязаны ссылаться на эти ключи через `configKey`; буквальные заголовки отклоняются. Выходной текст хранится в XML экспортёров, операции задаёт нативный код. Пути считаются от каталога книги.

- `export.Word.templates.path` — путь XML, сейчас `./PersonalEventBuilderWordResultTemplates.xml`.
- `events.path` — общий реестр событий и выбранных шаблонов.
- `export.Movement.templates.path` — XML Movement, сейчас `./PersonalEventBuilderMovementTemplates.xml`.
- `export.Word.<ModeId>.previewRequired` — обязательные поля превью. Для простого показа пункта принимать приказ не нужно.
- `export.Word.<ModeId>.required` — обязательные данные экспорта DOCX.
- `export.fields`, `export.dateFields` — снимаемые поля и проверяемые даты. Строковый формат — `DD.MM.YYYY`; невозможные даты и пропущенный год запрещены.
- `export.shared.aliases` — нумерованные сопоставления `name`, `kind`, `source`; виды: `value`, `shortDate`, `joinLowerFirst`.
- `export.shared.constants` — явные начальные значения необязательных полей: `text`, `emptyText`, `boolean`, `emptyCollection`. Для text и boolean нужен `value`.
- `export.dsl.maxDepth` — глубина рекурсии включений/блоков от 1 до 64.
- `export.dsl.months.<1..12>` — украинские названия месяцев для токенов даты.

Четыре кадровых поиска получают родительный звания, винительный ФИО, родительный должности и сокращённое ФИО в родительном из настроенного ШПО. Три поиска автора выполняются, когда условие ReportName определяет внешнего автора. Для пустого автора или «сам» они пропускаются; XML выбирает личный рапорт. Лечение дополнительно получает родительный или винительный больницы из справочника учреждений.

У каждого поиска в `export.shared.lookups.<index>` или `export.Word.<ModeId>.lookups.<index>` есть `path`, `sheet`, `range`, `input`, `keyHeader`, `valueHeader`, `output`, `when.field`, `when.pattern`; родительский `.count` задаёт число поисков. Общие результаты доступны обоим экспортёрам, специальные Word-поиски независимы. При совпавшем условии требуется ровно один непустой результат; отсутствие и неоднозначность останавливают подготовку. Пути и заголовки находятся только в конфиге. Открытые источники читаются Excel-механизмом запросов, закрытые — ADO.

## Сопоставления режимов

| Режим | XML-шаблон |
| --- | --- |
| FromHospital | FromHospital |
| FromAnnualMainVacation | FromAnnualMainVacation |
| FromVacationForCuring | FromVacationForCuring |
| FromVacationForFamily | FromVacationForFamily |
| ToHospital | ToHospital |
| ToAnnualMainVacation | ToAnnualMainVacation |
| ToVacationForCuring | ToVacationForCuring |
| ToVacationForFamily | ToVacationForFamily |

Word XML имеет версию `2`. Старые `groupHeader`, `groupBy`, `groupOrder` включены в обычный текст без изменения одиночного DSL. Постоянная группировка явно объявляется через `group`.

## Поддерживаемый DSL

Поддерживаются скаляры `{Name}`, `{[Alias]}`, общие `#include`, вложенные `#if/#else/#endif`, `#not/#and/#or`, `#for/#endfor`, `{#newline}`, `{#space}`. Элементы циклов — словари в Collection; доступны `item.__first`, `item.__last`, `item.__index`, `item.[Alias]`.

Форматтеры: `upper`, `lower`, `upperFirstLetter`, `lowerFirstLetter`, `regexreplace`, `dateoffset`, `dateformat`, `nbspAfterNumberSign`, `nbspBeforeInitials`, `nbspAfterSettlementAbbr`. Кавычки сохраняют запятые и вертикальные черты внутри regex-параметров. Даты принимают VBA-формат или токены `\dd`, `\d`, `\mm`, `\m`, `\yy`, `\yyyy`, `\month`.

`{[HospitalAccusative]|lowerFirstLetter|nbspAfterNumberSign}` связывает № со следующей цифрой через U+00A0, в том числе если исходно пробела нет. `{[FIOInitialsGenitive]|nbspBeforeInitials}` и `{[ReporterGenitive]|nbspBeforeInitials}` связывают фамилию с инициалами и раздельные инициалы между собой. Форматтеры применяются явно и меняют только выведенный текст; исходники и ключи группировки сохраняются. Пробелы и табуляция нормализуются локально, переносы остаются.

`{[Destination]|nbspAfterSettlementAbbr}` связывает `м.`, `с.`, `сел.` с названием через U+00A0. Сравнение без регистра принимает отсутствующие пробелы, табуляцию и готовые NBSP. Граница токена исключает совпадения внутри слов; переносы сохраняются.

Маркеры `#^`, `#_` и их варианты в скобках обрабатываются до данных. Итоговая нормализация сохраняет переносы и NBSP. Входные значения не разбираются повторно как DSL.

`#let` допускает внутренние имена с `__` и две операции: `$ex_Helpers.m_RegexIsMatch`, `$ex_Helpers.fn_IsDateOnOrAfter`. Обработчик вычисляет их сам и не вызывает произвольный VBA из XML. Неизвестные алиасы, включения и операции останавливают превью с диагностикой.

## Границы реализации

Текущая форма явно передаёт пустые MetaTables, MetaDocumentTables, MetaTvoTables. Продолжительность отпуска, дополнительные дни, дата окончания и специальные поля ВЛК также явно пусты, когда их ввод недоступен. Эти объявления настраиваются; они не являются молчаливым замещением отсутствующего поля DSL.

Сборка целого приказа, дополнительные варианты вставки, правила продовольственного обеспечения, цепочки ТВО, изменения статуса, HTML-таблицы и удаление/отмена экспорта требуют соответствующих входных полей и бизнес-правил. Реализованная постоянная группировка командировок описана ниже.

The reader preserves `documentFilepath`. In `blank` mode, regular export uses `export.Word.output` and rejects an existing file. In `template` mode, PROD uses `export.Word.documentTemplate`: `./{yyyy}-{mm}-{dd} №{OrderNum} ({Signatory}) (Александр).docx`. The output filename comes from the template filename with the configured `export.Word.documentFields` replacements. Date and number come from `AcceptedOrderDate` and `AcceptedOrderNumber`; the signatory is read from `orders.signatoryHeader` in the same configured orders source, matching both date and number. Missing or conflicting signatories stop export. Identical duplicate source rows are allowed.

On first export, the exporter creates a document from the template and replaces configured placeholders in all Word stories, including headers and footers. `{month}` uses the configured Ukrainian month name. The template remains unchanged. An exact existing output filename is reused without repeating header replacements; grouped sections retain embedded state and validation. Other anchored sections use their existing insertion and verification rules. Both creation and updates are published through a temporary file, with a source timestamp check before replacing an existing output. TEST retains its separate `blank` profile and test template settings.

## Постоянная группировка Word

У `template` сначала идут ноль или несколько `derived`, затем одно тело: `text` либо `group`. Смешивание тел, derived после тела и неизвестные теги/атрибуты отклоняются. Derived вычисляются по порядку для каждого события и могут использовать предыдущие значения; доступны обоим видам тела. Групповой шаблон требует `identityBy` и `orderBy`; обычный не может их объявлять.

`ToBusinessTrip` — маршрут выбытия только для Word. Одиночный текст находится в `group/single/text`; внешний text и промежуточный WordText не нужны. Превью и перестроение DOCX вызывают один обработчик: для одного события single, после добавления совпадающего второго — multiple.

`by` — составной ключ равенства; группы и люди сохраняют порядок первого появления. В multiple выводятся шапка, элементы и основание; внутри допустима вложенная group. У общей шапки/основания `common` перечисляет скаляры, которые должны совпадать. `GroupBasis` вычисляется из полного блока основания, поэтому рапорты разных авторов не объединяются только из-за одинакового входящего номера.

`numberRanges` использует три группы regex: префикс, целое и суффикс. Он сортирует серии численно, сжимает последовательные номера с полными концами диапазона, сохраняет пропуски, суффиксы и ширину ведущих нулей, отклоняет дубли и некорректные номера. Разные даты документов входят в разные группы. Число ограничено VBA Long; синонимы назначения не нормализуются и должны совпадать точно.

`ex_PEB_WordState` сохраняет настроенные скалярные снимки, ключ приказа, определения шаблонов/общих блоков и последний текст раздела в Custom XML: namespace `urn:excelprototype:word-events:v1`, версия данных `2`. `export.Word.grouping.stateFields` задаёт исходные поля; индивидуальный WordText не хранится. Namespace сохранён для обнаружения старых данных: версия 1 требует миграции. Отдельного файла состояния нет. Одинаковая повторная отправка идемпотентна; конфликт данных той же идентичности, другой приказ, ручные правки, дубли хранилища, несовместимая версия или изменённые шаблоны останавливают экспорт. Непустой старый раздел без данных требует миграции. Сейчас поддерживается добавление событий; замена и удаление не угадываются.

Путь группового документа — `export.Word.grouping.output`, привязанный к приказу, а не человеку. Повторный экспорт копирует файл во временный, проверяет и перестраивает из сохранённых событий. Только после сохранения и закрытия Word оригинал заменяется через `ReplaceFileW`; изменение метки исходного файла отменяет публикацию. Blank создаёт пару якорей; template требует ToBusinessTrip Begin/End в DOCX и не придумывает место вставки. Обычный экспорт сохраняет прежнее поведение. Превью показывает нового человека, общий результат собирается при экспорте.

`Scripts/Testing/PEB/validate_word_grouping.py` без Office проверяет конфигурацию, границы сохранения, 12 примеров грамматики и сценарии книги. Обе UI-группы читают `TestScenarios.EventRoutes.xlsx`, выбирая Basic или Grouping по Suite. У 16 групповых сценариев 34 шага в `tbGroupingSteps` и 25 ожиданий в `tbGroupingExpected`; связи настроены в test.scenarios. После разрешённого развёртывания `ex_PEB_GroupingTests.fn_Run` выполняет бизнес-сценарии Word. Техническая структура проверяется отдельным `validate_word_grammar.py`, также вызываемым групповой проверкой; VBA-рендеринг и разрешение значений не выполняются.

Runtime-набор использует временные DOCX: после каждого добавления сохраняет, закрывает и открывает файл; сравнивает первый пункт с одиночным превью, проверяет внешний текст, число событий, переход single → multiple, диапазоны, разделение ключей, повторы и отказ от запрещённых изменений. Проверяются версия данных и отсутствие сохранённого WordText. Книга сценариев и текущие пользовательские приказы не изменяются.

The Movement exporter retains append/close profiles, field/identity/join mappings, chronology checks and duplicate detection. PROD requires exactly one open workbook containing the configured table (`ТВ`) in the current Excel instance. It saves successful exports and leaves the workbook open; it never opens or closes the target. Missing or ambiguous targets produce a warning. TEST restricts lookup to the configured test path and keeps runner-owned workbook handling. Read-only or protected targets and the 1904 date system are rejected.

## Статические проверки

Word export reports `WordExportAction` through the configured `text.ExportWordSuccess` template: `INSERT` creates a document, fills an empty anchored section, or adds a grouped event; `UPDATE` replaces changed text in a non-grouped section; `SKIP` means identical text, an identical grouped event, or a route without a Word output. An unchanged existing output is closed without saving or replacing the original. Non-grouped sections are owned by their route, not by a personnel identity: exporting another person to the same section replaces its contents. Only the configured `ToBusinessTrip` template currently accumulates grouped events. Its identity is `PersonId; DateFromShort; EventId`; changing fields outside that identity is a conflict rather than an update, while changing an identity field adds a new event. Grouping keys control presentation and are independent of event identity. These statuses do not change the existing insertion/grouping policy or Movement exports.

PEB imports common dependencies using `importCfg / Import::workbooks.__cfgDependencies / ../cfgDependencies.xlsx`. Imported keys belong to that variable: `ref / PersonalEventBuilder::orders.numberHeader / __cfgDependencies.source.Orders.numberHeader`. Each workbook import declares its own `__variable`, so different workbooks can contain identical raw keys without collisions. References inside an imported workbook resolve within its variable unless they explicitly name another `__variable`. The shared `wsConfig/tbConfig` table retains its original unqualified `Kind / Key / Value` rows. `file` paths resolve relative to the workbook owning the row; explicit TEST overrides use qualified keys such as `__cfgDependencies.fileMovementHistory` and paths relative to the PEB workbook. Local overrides remain independent of row order. Missing imports or keys, invalid or duplicate variables, duplicate keys within a variable, cycles and nested imports stop refresh before `tbConfig` changes. Unqualified local references can still address local keys; imported keys require a variable. No fallback to PROD sources is performed.

`ex_ConfigRefresh` expands imports and references into effective values in the local `tbConfig` table. The existing lookup, export and controller snapshot readers continue using that table. Changes to the shared workbook take effect on configuration refresh, including when switching PROD/TEST. Refresh is required after deploying this source migration; importing VBA alone does not refresh a running workbook. Runtime VBA/import checks remain a separate authorized step.

The shared workbook adds institutions, EJOS, Movement history and export targets, and current PEB source sheets/ranges. Saved `РУХ_last.xlsx` history and the open Movement export target have separate keys. TEST overrides `filePersonnelShpo`, `fileEjos`, `fileMovementHistory` and `fileMovementExport`; production addresses are not copied into the TEST overrides. `config/cfgDependencies/wsConfig.txt` is the editable source for the common workbook. `Scripts/Config/update_shared_dependencies.mjs` updates it with artifact-tool, preserves the original table formatting, and makes a backup. It requires the bundled Node runtime, `NODE_PATH` pointing to bundled packages, and `CODEX_PYTHON` pointing to bundled Python. `Scripts/Testing/PEB/validate_config_dependencies.py` verifies resolved configuration parity and import error cases without Office.

Ctrl+Alt+U uses the configuration service in WorkbookUpdater.xlam for both PEB and plain configuration workbooks. PERSONAL captures the active workbook; PEB retains only a compatibility proxy for its guarded hotkey and PROD/TEST switching. Both resolve the add-in via ThisWorkbook::vbaPath. The common workbook provides ThisWorkbook::id, ThisWorkbook::config.source, ThisWorkbook::vbaPath and profile-specific success/failure messages. Plain configurations preserve their original kinds and relative file values. The add-in packages the common import resolver with a standalone Excel reader, reuses an already open shared workbook, and closes only a workbook it opened itself. Ctrl+Alt+R calls the same service before reading the reload plan or requesting runtime shutdown; a failed refresh cancels the reload. Initial installation refreshes through the add-in before runtime initialization. Deploy PEB sources and rebuild WorkbookUpdater.xlam together. Neither deployment nor runtime validation is performed by source-only checks.

Return ticket autofill follows the PrototypeNew contract for `FromVacationForCuring`, `FromAnnualMainVacation`, `FromVacationForFamily` and `FromBusinessTrip`. `lookup.ReturnTickets` configures the saved Movement source (PROD: Desktop `РУХ_last.xlsx`; TEST: `Dependencies/test_РУХ.xlsx`), explicit range and headers. The provider reads a temporary disk copy with ADO, indexes the last physical row containing both `Супровідний документ` and `Вибуття.Наказ`, and matches personnel by IPN. The configured temporary IPN marker uses FIO instead. Ticket numbers are copied literally and are never qualified using the current UI order. Ticket dates are resolved from the same row's departure-order number using a saved copy of the configured orders directory. As in PrototypeNew, this lookup uses the current system year's orders. Missing history or a missing departure-order date leaves both ticket fields blank; conflicting dates stop enrichment. Other modes retain their existing EJOS/current-order ticket calculation.

Parsed snapshots are cached by source path, modification time and size; saving the original or explicitly refreshing candidates rebuilds the cache. Unsaved live workbook changes do not affect the snapshot. Temporary files are removed after ADO closes. Mode switches refresh candidate tickets and the selected record; switching between ticket sources clears autofilled ticket fields before loading so a failed lookup cannot leave a ticket from the previous source. `Scripts/Testing/PEB/validate_return_tickets.py` checks the static wiring and offline saved/unsaved workbook contract without running VBA.

`Scripts/Testing/PEB/validate_movement_templates.py` проверяет TEST/PROD DSL, события, колонки и структуру VBA. `movement_plans.py` читает буквальные VBA-объявления только для offline-проверки и подготовки данных; production не разбирает исходники VBA. Миграция сохраняет 167 свойств сопоставлений, 35 планов и нормализованное содержимое 35 оснований. Office не запускается. Нативный модуль, XML, маршруты и выбранный конфиг развёртываются вместе штатным разрешённым обновлением.

`Scripts/Testing/PEB/validate_word_routes.py` проверяет необязательные экспортёры TEST/PROD, независимые реестры только Word/Movement, общие шаблоны и отказ от отсутствующих ссылок, дублей, пустых событий и удалённых разделов. Пропуск Word — внутренний результат отсутствия word-маршрута.

`Scripts/PowerShell/Test-PEBWordTemplates.ps1` без Office проверяет TSV-конфиг, namespace/версию, имена, единственность тела, ссылки/циклы include, вложенность DSL, разрешённые функции #let и объявленные поля всех текстовых блоков. Для полной грамматики группировки и контрактов поставщиков контекста используйте также Python-валидаторы.

Статические проверки не подтверждают runtime-поведение. Импорт исходников в рабочую книгу, компиляция VBA, автоматизация Word, поиски ADO и запись Movement этой проверкой не выполняются.