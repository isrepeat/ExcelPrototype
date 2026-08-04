# Supporting Document Builder

Экспериментальный режим построения сопроводительных документов на базе PEB.

Режим имеет отдельные `obj_PageSDB` и
`obj_PageSDBCtrl`. Имя controller задаётся обязательным ключом
`SupportingDocumentBuilder.ControllerClass` в профиле. Общий для PEB и SDB
парсер шаблонов расположен в `[2] classes/obj_WordResultTplParser.cls.vba`.

Файл `SupportingDocumentBuilderTemplate.docx` должен содержать обычные WORD-якоря
экспортёра для используемых `template@id`, например:

```text
{\export:BusinessTripCertificate_Begin}
{\export:BusinessTripCertificate_End}
```

Сейчас режим содержит один документ `ПОСВІДЧЕННЯ ПРО ВІДРЯДЖЕННЯ` и ручную
форму без Lookup. `<text>` рендерится по правилам PEB. `<table>` показывается в Excel preview как
заполненная HTML-разметка и при экспорте преобразуется в настоящий `Word.Table`.
Preview, содержащее таблицу, нельзя использовать как редактируемый источник
экспорта.

Минимальный синтаксис таблицы:

```xml
<table wordStyle="builtin:tableGrid">
  <tr>
    <th>ПІБ</th>
    <th>Посада</th>
  </tr>
  <tr for="item in MetaTables">
    <td>{item.[FIO]}</td>
    <td>{item.[PositionName]}</td>
  </tr>
</table>
```

Поддерживаемые структурные элементы: `table`, `caption`, `thead`, `tbody`,
`tfoot`, `tr`, `th`, `td`. В первой версии объединения `rowspan`/`colspan` не
преобразуются в объединённые ячейки Word.
