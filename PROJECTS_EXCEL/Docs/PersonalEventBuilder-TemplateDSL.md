# DSL шаблонов PersonalEventBuilder

**Навигация:** [Главная](../../README.md) → [PROJECTS_EXCEL](../README.md) → [Документация](README.md) → DSL шаблонов

## Содержание

- [1. Файлы и их назначение](#1-файлы-и-их-назначение)
- [2. Простой текст и подстановка полей](#2-простой-текст-и-подстановка-полей)
- [3. Форматтеры и неразрывные пробелы](#3-форматтеры-и-неразрывные-пробелы)
- [4. Условия](#4-условия)
- [5. Многострочная запись без лишних пробелов](#5-многострочная-запись-без-лишних-пробелов)
- [6. Общие блоки и include](#6-общие-блоки-и-include)
- [7. Циклы и локальные вычисления](#7-циклы-и-локальные-вычисления)
- [8. Шаблоны колонок Movement](#8-шаблоны-колонок-movement)
- [9. Вычисляемые значения derived](#9-вычисляемые-значения-derived)
- [10. Группировка по одной дате](#10-группировка-по-одной-дате)
- [11. Несколько ключей и вложенная группировка](#11-несколько-ключей-и-вложенная-группировка)
- [12. Командировка: общее основание и диапазоны](#12-командировка-общее-основание-и-диапазоны)
- [13. Второй человек в существующем DOCX](#13-второй-человек-в-существующем-docx)
- [14. Маршруты PersonalEventBuilderEventRoutes.xml](#14-маршруты-personaleventbuildereventroutesxml)
- [15. Проверка и тестирование](#15-проверка-и-тестирование)

Руководство описывает текущую реализацию Word и Movement: от подстановки одного поля до перестроения группового пункта в существующем DOCX. Имена тегов, полей и форматтеров приведены в том виде, в котором их понимает код. Примеры — фрагменты XML, если явно не показан корневой элемент.

## 1. Файлы и их назначение

| Файл | Назначение |
| --- | --- |
| `workbook/PersonalEventBuilder/PersonalEventBuilderEventRoutes.xml` | Связывает EventId с независимыми шаблонами Word и Movement. |
| `workbook/PersonalEventBuilder/PersonalEventBuilderWordResultTemplates.xml` | Текст Word, общие блоки, вычисляемые значения и группировка. |
| `workbook/PersonalEventBuilder/PersonalEventBuilderMovementTemplates.xml` | Тексты отдельных колонок Movement. |
| `config/PersonalEventBuilder/wsConfig.txt` и `wsConfig.test.txt` | Пути, имена колонок, подготовка контекста и параметры экспорта. |

Пути в таблице указаны относительно `PROJECTS_EXCEL`. Пространство имён XML — `urn:excelprototype:profiles`. Текстовый DSL общий для двух экспортёров. Структурные элементы `derived` и `group` поддерживаются в Word; в Movement текст формируется отдельно для каждой колонки.

## 2. Простой текст и подстановка полей

Минимальный Word-файл с одним текстовым шаблоном:

```xml
<wordResultTemplates xmlns="urn:excelprototype:profiles"
                     version="2"
                     generationMode="anchored">
  <template name="Example">
    <text>Відрядити військовослужбовця.</text>
  </template>
</wordResultTemplates>
```

Чтобы вывести подготовленное значение, используйте `{[FieldName]}`:

```xml
<template name="Example">
  <text>Відрядити {[FIOAccusative]} у {[Destination]}.</text>
</template>
```

При `FIOAccusative = БОНДАРЕНКА Олександра` и `Destination = м. Київ` получится `Відрядити БОНДАРЕНКА Олександра у м. Київ.`

Фигурные скобки обозначают подстановку, квадратные — поле или алиас. Обычные переменные записываются также как `{Caption}`. Подстановка не создаёт поле и не ищет колонку Excel сама: значение должно быть подготовлено источниками, алиасами или входными параметрами экспортёра. Отсутствующий ключ вызывает ошибку; для необязательного поля нужно предусмотреть пустое значение.

Если текст содержит символы, значимые для XML, используйте CDATA. Она защищает XML, но не отключает обработку DSL:

```xml
<text><![CDATA[Відрядити {[FIOAccusative]}.]]></text>
```

## 3. Форматтеры и неразрывные пробелы

Форматтеры записываются через `|` и выполняются слева направо:

```text
{[PositionGenitive]|lowerFirstLetter}
{[HospitalAccusative]|lowerFirstLetter|nbspAfterNumberSign}
{[ReporterGenitive]|nbspBeforeInitials}
{[Destination]|nbspAfterSettlementAbbr}
```

| Форматтер | Результат |
| --- | --- |
| `upper`, `lower` | Меняет регистр всего значения. |
| `upperFirstLetter`, `lowerFirstLetter` | Меняет регистр первого символа. |
| `nbspAfterNumberSign` | Ставит U+00A0 между `№` и следующей цифрой. |
| `nbspBeforeInitials` | Связывает фамилию с инициалами и инициалы между собой. |
| `nbspAfterSettlementAbbr` | Связывает `м.`, `с.`, `сел.` с названием населённого пункта. |
| `dateoffset:"+1"` | Добавляет целое число дней к дате. |
| `dateformat:"dd.mm.yyyy"` | Форматирует дату по формату VBA. |
| `dateformat:"\dd \month \yyyy року"` | Использует токены даты, включая украинское название месяца. |
| `regexreplace:"pattern","replacement"` | Заменяет совпадения по правилам VBScript.RegExp; группы подставляются как `$1`, `$2`. |

Здесь `⍽` обозначает неразрывный пробел U+00A0; в документ выводится сам пробел:

```text
лікарня № 25 → лікарня №⍽25
БОНДАРЕНКА О. Ю. → БОНДАРЕНКА⍽О.⍽Ю.
м. Київ → м.⍽Київ
```

Форматтер обрабатывает целиком пришедшую строку. Для больницы не нужно отдельно извлекать номер из полного названия.

```text
{[DateFromShort]|dateoffset:"+1"|dateformat:"dd.mm.yyyy"}
{[Destination]|regexreplace:"^[Мм]едична([ \t]+)рота","медичну$1роту"}
```

Неизвестный форматтер, некорректная дата, дробное смещение дней или неверные параметры вызывают ошибку.

## 4. Условия

```text
{#if [VacationTicketNo]}Посвідчення № {[VacationTicketNo]}{#else}Посвідчення не зазначено{#endif}
```

Поддерживаются `#not`, `#and`, `#or`. Для сложных сочетаний удобнее явно вложить условия:

```text
{#if [VacationTicketNo]}{#if [VacationTicketDateShort]}Посвідчення № {[VacationTicketNo]} від {[VacationTicketDateShort]}{#endif}{#endif}
```

Условия позволяют убрать необязательную фразу при пустом значении или ложном флаге. Нужные ключи контекста при этом должны существовать. Переносы и отступы внутри текста выводятся буквально, если их не убрать токенами оформления.

## 5. Многострочная запись без лишних пробелов

| Токен | Действие |
| --- | --- |
| `{#newline}` | Выводит перенос строки. |
| `{#space}` | Выводит обычный пробел явно. |
| `#^` или `{#^}` | Удаляет следующий перенос исходного текста с прилегающими отступами; сам маркер тоже удаляется. |
| `#_` или `{#_}` | Удаляет маркер и следующие за ним пробелы/табуляцию отступа. |

```xml
<text><![CDATA[#^
#_    Відрядити {[FIOAccusative]}#^
#_    {#space}у {[Destination]|nbspAfterSettlementAbbr}.#^
]]></text>
```

Получится одно предложение. `{#space}` между частями нужен, иначе слова склеятся. Разметка обрабатывается до подстановки данных: текст, пришедший из источника, эти токены не изменяют. Нужные разрывы задавайте через `{#newline}`.

## 6. Общие блоки и include

Повторяющийся текст вынесите в `sharedTemplates`: это может быть строка человека или основание.

```xml
<sharedTemplates>
  <sharedTemplate id="block.Person">
    <text>{[RankGenitive]} {[FIOAccusative]}</text>
  </sharedTemplate>
</sharedTemplates>
<template name="Example">
  <text>Відрядити {#include block.Person}.</text>
</template>
```

`{#include block.Person}` обрабатывает общий блок в текущем контексте. Идентификатор блока должен существовать и быть уникальным. Глубина вложенности и включений ограничена обработчиком.

`sharedTemplate` хранит текст, а не объявляет глобальные структурные переменные. `derived` располагается в использующем его Word-шаблоне перед основным телом. Включение общего блока не добавляет `derived` во все шаблоны.

Каталоги общих блоков Word и Movement раздельные. Например, Word-блок `block.IncomingBasis` содержит `Підстава:`, а Movement-блок с тем же именем формирует основание без этой подписи. Определять блок нужно в файле, где он используется.

### Resource dictionaries

Word and Movement templates can declare one optional `resources` section directly under their root element. Resources are parsed and indexed once when the template file is loaded. Each file owns its dictionaries, so Word and Movement can use the same dictionary name with different output phrases.

```xml
<resources>
  <dictionary name="DocumentNames"
              ignoreCase="true"
              normalize="trim,slashSpaces"
              onMissing="keep">
    <entry key="м/к">
      <variant value="медична карта стаціонарного хворого"/>
      <variant value="виписка із медичної карти стаціонарного хворого"
               field="SourceStatus"
               equals="HospitalTreatment"/>
    </entry>
  </dictionary>
</resources>
```

Use the dictionary as a formatter:

```text
{[__DocumentName]|lookup:"DocumentNames"|lowerFirstLetter}
```

All alternatives are sibling `variant` elements. Each entry requires exactly one unconditional variant. A conditional variant supplies both `field` and `equals`; the field must be declared in the template context. `SourceStatus` and `ModeId` can select status-specific or mode-specific phrases. One matching condition overrides the unconditional variant regardless of XML order. Multiple matching conditions are an error.

`ignoreCase` controls key and condition-value comparison. `normalize="trim"` replaces nonbreaking spaces, narrow nonbreaking spaces and tabs with ordinary spaces and trims outer spaces. `trim,slashSpaces` also removes spaces adjacent to slash characters, making ` М / К ` equivalent to `м/к` in a case-insensitive dictionary. It does not remove spaces inside full document names.

`onMissing="keep"` preserves an unknown input verbatim; `onMissing="error"` rejects it. Unknown dictionary names, duplicate normalized keys, missing or duplicate unconditional variants, duplicate conditions and unsupported attributes are errors. Lookup uses indexed keys and conditional values rather than regular expressions. WORD event snapshots retain dictionary condition fields, and the template fingerprint includes resource XML.

### Item ordering

An `itemOrder` resource sorts any collection of dictionaries by a configured field. Document preparation remains separate from sorting. Both export files declare:

```xml
<itemOrder name="BasisDocuments"
           field="DocumentNote"
           ignoreCase="true"
           normalize="trim,slashSpaces">
  <item key="м/к"/>
  <item key="Довідка ВЛК"/>
  <others/>
  <item key="в/к"/>
</itemOrder>
```

Place this element inside `resources`. Each order requires `name`, `field`, `ignoreCase`, `normalize` and exactly one `others` placeholder. `field` names a required scalar field on each collection item, rather than a root context field. `ignoreCase` selects case-sensitive or case-insensitive matching; `normalize` accepts `none`, `trim` or `trim,slashSpaces`. Optional `aliases` contains pipe-separated equivalent keys. Duplicate normalized keys, duplicate aliases, missing item fields and unknown order names are errors. The key `@others` is reserved internally.

The same mechanism supports non-document collections: an order with `field="Priority"` can be used by `{#for item in Tasks orderBy:"TaskPriorities"}`. Neither the source collection nor the UI is reordered.

```text
{#for item in BasisDocuments orderBy:"BasisDocuments"}
  {item.[DocumentNote]} № {item.[DocNo]} від {item.[DocDateShort]}
{#endfor}
```

`BasisDocuments` combines the main document, additional document forms and legacy VLK scalar fields without changing UI or history order. Sorting is stable: documents with the same known type retain their relative order, and unknown documents retain their order inside `others`. Current templates place medical cards first, VLK certificates next, unknown documents next and vacation tickets last. Blank document rows are omitted by the output condition.

## 7. Циклы и локальные вычисления

Quoted string assignments are supported in Word and Movement: `#let __DocumentName = "{[DocumentNote]}";`. The right-hand string is rendered in the current context and stored as a scalar; doubled quotes escape a literal quote. No arbitrary VBA expression is executed. Includes share the current context, while loop iterations use isolated child contexts. Assign `"{item.[DocumentNote]}"` inside a document loop to reuse the same `block.DocumentName` for each child without overwriting the main document's variable.

The event context exposes `SourceStatus`, `TargetStatus` and `Operation` independently of the exporter. Statuses use stable plan identifiers such as `HospitalTreatment` or `MedicalLeave`; operations are `Arrival`, `Departure`, `Transition`. Use the existing `#let` RegexIsMatch function to test these identifiers and combine the resulting Boolean variables with `#and` / `#or`. Domain-specific flags and document phrases do not need to be generated in VBA.

Цикл работает с заранее подготовленной коллекцией словарей:

```text
{#for item in MetaDocumentTables}{#if #not item.__first}; {#endif}{item.[DocumentNote]} № {item.[DocNo]}{#endfor}
```

`item.[DocNo]` — поле текущей строки; `item.__first` — признак первой строки. Значения внешнего контекста доступны внутри цикла. Каждая итерация получает дочерний контекст, поэтому поля одной строки не подменяют следующую. `#for` не читает произвольные листы Excel: коллекцию готовит приложение.

Для локального вычисления поддерживается `#let` с обязательной точкой с запятой. Так существующий блок основания определяет, относится ли рапорт к самому участнику:

```text
#let __IsRaportFromSelf = $ex_Helpers.m_RegexIsMatch("{[ReportPerson]}", "^\s*$|^\s*[Сс][Аа][Мм]\s*$");
{#if __IsRaportFromSelf}{[FIOInitialsGenitive]}{#else}{[ReporterGenitive]}{#endif}
```

Это поддерживаемое выражение DSL, а не возможность вставлять произвольные операторы VBA. Для именованной строки, участвующей в группировке Word, используйте `derived`.

## 8. Шаблоны колонок Movement

Небольшой файл Movement:

```xml
<resultTemplates xmlns="urn:excelprototype:profiles"
                 version="1">
  <template name="ToBusinessTrip">
    <column configKey="export.Movement.headers.DepartureBasis">
      <text>рапорт {[ReporterGenitive]|nbspBeforeInitials}, вх. № {[IncomingNo]} від {[IncomingDateShort]}.</text>
    </column>
    <column configKey="export.Movement.headers.TicketDocument">
      <text>{[VacationTicketNo]}</text>
    </column>
  </template>
</resultTemplates>
```

`column/@configKey` ссылается на ключ wsConfig, значение которого — настоящее имя колонки назначения. Кириллические имена колонок остаются в конфигурации. В этом примере предполагается, что зарегистрированная операция допускает запись обеих колонок.

XML определяет текст колонок. Поиск строк, добавление, закрытие записей и разрешённые колонки определяет `ex_PEB_MovementRegistry` по EventId. Произвольная новая колонка в XML сама по себе не начинает записываться.

Для прибытия и выбытия можно выбрать разные шаблоны. Переход может задавать разные тексты закрываемой записи и нового выбытия в пределах возможностей VBA-операции. Группировки Word и хранения событий DOCX в Movement нет; внутри `text` доступны общие подстановки, условия, циклы, включения и форматтеры.

## 9. Вычисляемые значения derived

`derived` формирует именованную строку отдельно для каждого события до группировки:

```xml
<template name="Example">
  <derived name="HospitalCaption">
    <text>{[HospitalAccusative]|nbspAfterNumberSign}</text>
  </derived>
  <derived name="Caption">
    <text>У {[HospitalCaption]}</text>
  </derived>
  <text>{[Caption]} направити {[FIOAccusative]}.</text>
</template>
```

Можно объявить несколько `derived` или ни одного. Все объявления располагаются до основного тела и вычисляются по порядку: следующий может использовать предыдущий. Имя не должно совпадать с входным полем или ранее объявленным значением. Ссылаться на следующее объявление и создавать циклические зависимости не следует.

Обработчик копирует контекст события и добавляет значения в эту копию. Область доступности — текущий шаблон для текущего события: последующие `derived`, основной текст, включаемые блоки и ключи группировки. Общий входной контекст и другие шаблоны не меняются. При перестроении DOCX значения вычисляются заново из сохранённых исходных данных.

На первом уровне Word-шаблона допустимы `derived*`, а затем ровно одно тело: `text` либо `group`. У обычного `text` нельзя указывать `identityBy` и `orderBy`; у группового шаблона оба обязательны. Одновременно объявлять внешний `text` и `group` нельзя: одиночный пункт располагается в `group/single/text`.

## 10. Группировка по одной дате

Пример структуры группового шаблона:

```xml
<template name="ExampleByDate"
          identityBy="PersonId; DateFromShort; EventId"
          orderBy="AcceptedOrderNumber; OrderDateShort">
  <group by="DateFromShort">
    <single>
      <text>{[FIOAccusative]} — з {[DateFromShort]}.</text>
    </single>
    <multiple>
      <header common="DateFromShort">
        <text>З {[DateFromShort]}:</text>
      </header>
      <item>
        <text>{[FIOAccusative]}.</text>
      </item>
    </multiple>
  </group>
</template>
```

Каждая уникальная дата образует отдельную группу. Для одного человека используется `single`, для нескольких — `multiple`. Обработчик разделяет выводимые блоки и строки людей абзацными интервалами.

| Исходные события | Структура результата |
| --- | --- |
| А: 15.10; Б: 15.10 | Одна шапка даты и два человека. |
| А: 15.10; Б: 16.10 | Два одиночных пункта. |
| А/Б: 15.10; В/Г: 16.10 | Две шапки, под каждой по два человека. |

`common` перечисляет поля, которые обязаны совпадать у всех участников перед выводом общей шапки или основания. Несовпадение вызывает ошибку, а не создаёт другую группу. Поля, по которым группы должны разделяться, перечисляйте в `by`.

Добавление нового группируемого события требует не только XML: нужны маршрут, подготовленные поля, договорённость о сохранении исходных данных и якоря DOCX. Пример не регистрирует новый EventId автоматически.

## 11. Несколько ключей и вложенная группировка

```xml
<group by="HospitalShortName; EffectiveReturnDateShort">
  <!-- Вставьте single и multiple по образцу предыдущего примера. -->
</group>
```

Это фрагмент ключа, а не готовая группа: ветки `single` и `multiple` обязательны. Одна больница и одна дата объединяются; та же больница в другую дату образует другой пункт. Если в контексте есть стабильный идентификатор учреждения, используйте его вместо изменяемого отображаемого названия. Разные написания автоматически не отождествляются.

В `multiple` должен быть ровно один `item` либо одна вложенная `group`. Перед ним и после него допустимы `header` и `footer`. Вложенная группа нужна для иерархии вроде общей даты с подзаголовками больниц. У каждого уровня снова должны быть обе ветки. Если отдельный уровень заголовков не нужен, проще использовать составной `by`.

| Атрибут | Назначение |
| --- | --- |
| `identityBy` | Идентичность события для обнаружения повторов и конфликтов. |
| `orderBy` | Принадлежность одному приказу. Несмотря на имя, это не сортировка. |
| `group/@by` | Поля, по которым события распределяются в группы. |
| `header/@common`, `footer/@common` | Проверка равенства полей общей части. |

Атрибутов `sort`, `groupOrder` и направления сортировки ключей в текущем XML нет. Сортировка номеров в агрегате не сортирует список людей.

## 12. Командировка: общее основание и диапазоны

Структура рабочего `ToBusinessTrip` ниже сохраняет его ключи, но для краткости упрощает оформление дат. Используемые общие блоки объявлены в Word-файле.

```xml
<template name="ToBusinessTrip"
          identityBy="PersonId; DateFromShort; EventId"
          orderBy="AcceptedOrderNumber; OrderDateShort">
  <derived name="GroupBasis">
    <text>{#include block.IncomingBasis}</text>
  </derived>
  <group by="DateFromShort; Destination; GroupBasis; VacationTicketDateShort">
    <aggregate name="CertificateNumbers"
               field="VacationTicketNo"
               operation="numberRanges"
               pattern="^(.*[/])?([0-9]+)(.*)$"
               separator=", "
               rangeSeparator="–"/>
    <single>
      <text><![CDATA[Відрядити {#include block.PersonLineLower} у {[Destination]|nbspAfterSettlementAbbr}, з {[DateFromShort]}.{#newline}{#newline}{[GroupBasis]}{#include block.OptionalBusinessTripCertificate}.]]></text>
    </single>
    <multiple>
      <header common="DateFromShort; Destination">
        <text>Відрядити у {[Destination]|nbspAfterSettlementAbbr}, з {[DateFromShort]}:</text>
      </header>
      <item>
        <text>{#include block.PersonLineLower}.</text>
      </item>
      <footer common="GroupBasis; VacationTicketDateShort">
        <text>{[GroupBasis]}; посвідчення про відрядження № {[CertificateNumbers]} від {[VacationTicketDateShort]}.</text>
      </footer>
    </multiple>
  </group>
</template>
```

`GroupBasis` входит в `by`, потому что одинаковые место и дата не означают одинаковый рапорт. Если у А и Б входящие номера рапортов 100 и 101, строки основания отличаются — люди остаются в отдельных группах. Если совпадает всё отформированное основание, его можно вывести один раз в общем `footer`. Личные рапорты разных людей могут различаться даже при совпадении номеров.

`numberRanges` требует три группы регулярного выражения: префикс, целая числовая часть и суффикс. Выражение должно совпадать со всем номером. Агрегат разделяет совместимые серии, сортирует числа и объединяет последовательные номера. Концы диапазона выводятся полностью:

```text
2026/298/103, 2026/298/101, 2026/298/102
→ 2026/298/101–2026/298/103

2026/298/101, 2026/298/103
→ 2026/298/101, 2026/298/103
```

Разные префиксы и суффиксы не объединяются в один диапазон. Ширина чисел с ведущими нулями участвует в разделении серий. Повтор полного номера, несовпадение с выражением или выход числа за диапазон VBA Long вызывают ошибку.

Агрегаты вычисляются до выбора `single`/`multiple` и доступны контексту общей ветки. Каждый `item` обрабатывается с собственными полями человека.

## 13. Второй человек в существующем DOCX

Группировка работает между последовательными экспортами:

1. Первый экспорт создаёт одиночный пункт и сохраняет исходные данные события в Custom XML внутри DOCX.
2. Следующий экспорт читает эти данные и проверяет приказ, идентичность события, совместимость шаблона и целостность управляемого раздела.
3. Значения `derived` вычисляются заново; события распределяются по `by`. Если новый человек подходит в существующую группу, прежний одиночный пункт заменяется шапкой, списком людей и общим основанием.
4. Документ сохраняется через временный файл и публикуется после проверок. Содержимое вне управляемого раздела должно сохраниться.

Версия данных — 2, namespace — `urn:excelprototype:word-events:v1`. В шаблоне DOCX нужна настроенная пара маркеров Begin/End для раздела; для командировок используется идентификатор `ToBusinessTrip`. Эти маркеры относятся к экспорту, а не к командам текстового DSL.

Повтор того же события с теми же данными не создаёт дубль. Другая информация для той же идентичности, другой приказ, ручные правки раздела, несовместимые данные или изменённые определения шаблонов вызывают отказ. Для старых документов после изменения схемы может потребоваться миграция. Удаление или замена сохранённых событий автоматически не выполняются.

Интерактивное превью показывает одиночный пункт поступившего человека. Общий результат собирается при экспорте. Шрифт, абзацы и подсветка ФИО/основания относятся к DOCX и коду экспортёра; вставка `<b>` внутрь `text` не задаёт оформление Word.

## 14. Маршруты PersonalEventBuilderEventRoutes.xml

```xml
<eventRoutes xmlns="urn:excelprototype:profiles"
             version="1"
             bindingField="ModeId">
  <events>
    <event label="Відрядження"
           caption="У відрядження"
           id="ToBusinessTrip"
           direction="Departure"
           form="Leave"
           visible="true">
      <movement template="ToBusinessTrip"/>
      <word template="ToBusinessTrip"/>
    </event>
  </events>
</eventRoutes>
```

| Атрибут или элемент | Назначение |
| --- | --- |
| `bindingField` | Поле выбора бизнес-события; сейчас ModeId. |
| `id` | Стабильный EventId, используемый подготовкой данных и операциями VBA. |
| `label`, `caption` | Метаданные названия события и его представления в UI. |
| `direction` | Arrival, Departure или StatusChange. |
| `form` | Leave либо Treatment. |
| `visible` | Предлагается ли событие в UI; скрытые маршруты могут существовать. |
| `movement/@template` | Имя шаблона Movement. |
| `word/@template` | Имя шаблона Word. |

У события должен быть один или оба независимых маршрута, каждый не более одного раза. Наличие тега определяет участие экспортёра. Если Word не нужен, тег `word` отсутствует; пустой Word-шаблон создавать не нужно. Аналогично работает событие только для Word.

У `word` и `movement` поддерживается только атрибут `template`. Старые `policy`, `profile`, `anchor`, `operation` и `omittedTemplates` не относятся к этой схеме. Операция Movement определяется в VBA по EventId; подготовка Word и привязка якорей — конфигурацией. Ссылки на шаблоны должны существовать в соответствующих каталогах.

Новый EventId в XML сам по себе не реализует новую операцию Movement, новые источники данных или место вставки Word.

## Export rejection rules

Word and Movement templates accept optional `reject` elements before their body or columns. Without rules, export remains allowed. Each rule checks one prepared input field with exactly one of `equals` or `matches`; any matching rule rejects the current export with the configured `messageKey`. For `equals`, numeric values compare numerically; other values compare as case-sensitive text. `matches` searches the field using a VBScript regular expression; invalid patterns are configuration errors. Missing fields or messages are errors. Rules use prepared inputs, not template-local `derived` fields.

The Word template `ToAnnualMainVacation` now includes:

```xml
<reject field="Destination"
        matches="\(\s*[A-Za-zÀ-ÖØ-öø-ÿĀ-ž][A-Za-zÀ-ÖØ-öø-ÿĀ-ž0-9\s,.'’/-]*\)"
        messageKey="text.ExportRejected"/>
```

The rule detects a potential foreign-vacation marker in `Destination`: a parenthesized phrase written in Latin letters, such as `(Zurich, Switzerland)` or `(Zurich, Schweiz)`. Accented Latin letters, spaces, line breaks, digits and common place-name punctuation are supported. Parenthesized Cyrillic text, digits alone and Latin text without parentheses do not match. The number of additional days no longer affects this rule. The rule applies to the annual-main-vacation Word template; Movement remains allowed unless its own template contains a matching rule. Store each rejection message in both PROD and TEST wsConfig files. Rules apply to history exports too.

`VacationAdditionalDays` remains available as the sum of all additional duration operands, including road days, for independent rules using `equals`.

All 17 Word templates for departure or transition into vacation include `block.OptionalForeignVacationPermission` before the basis. It uses the same `Destination` pattern and conditionally adds a separate line, `Дозволено виїзд за кордон.`. Domestic destinations add neither the sentence nor its line break. This includes annual, family, maternity and medical leave, as well as transitions into these vacations. The annual-main-vacation rejection rule remains in effect independently: a matching rejected event still cannot be exported to Word.

Word rejection skips only the matching event: preparation sets `WordPolicy=Omit`, retains `WordRejectionReason`, and succeeds. Batch generation continues with other events and journals rejected items as `RejectedOmitted`, never `Generated`. If all events are omitted, generation succeeds without opening Word or creating a document. Word preview preparation also omits rejected events; pure template rendering does not enforce export policy. Invalid configuration and missing rule fields remain errors. Movement rejection still stops its current operation. Existing documents are not modified merely by adding a rule.

## 15. Проверка и тестирование

Команды запускаются из корня репозитория:

```text
python Scripts/Testing/PEB/validate_word_grammar.py
python Scripts/Testing/PEB/validate_export_rejections.py
python Scripts/Testing/PEB/validate_word_routes.py
python Scripts/Testing/PEB/validate_movement_templates.py
python Scripts/Testing/PEB/validate_word_grouping.py
```

`validate_word_grammar.py` проверяет 12 искусственных примеров структуры через Python-валидатор. Он не выполняет VBA-подстановки, форматтеры и вычисления `derived`. Проверки маршрутов и шаблонов читают реальные определения и конфигурацию. Проверка группировки дополнительно проверяет сценарии книги и связи с исходным кодом; она также запускает проверку грамматики.

Runtime-сценарии выполняются отдельно. Базовые сравнивают результаты с ожиданиями книги. G01–G16 проверяют реальный экспорт Word: превращение одиночного пункта в группу, разделение по датам/местам/основаниям, серии и пропуски номеров, повторную отправку, конфликты и сохранение внешнего текста.

Общий каталог находится в `TestScenarios.EventRoutes.xlsx`; шаги группировки — на `TestGrouping_Steps`, ожидания — на `Expected_WORD_Grouping`. Импорт исходников и запуск макросов требуют отдельного разрешения пользователя. Успех статических скриптов не подтверждает оформление или поведение Word во время исполнения.

При ошибке определите её уровень: XML, отсутствующее поле, выражение DSL, ключ группы, проверка `common`, возможности операции Movement или якоря/состояние DOCX. Исправляйте соответствующий контракт, не подменяя ошибку пустым результатом.