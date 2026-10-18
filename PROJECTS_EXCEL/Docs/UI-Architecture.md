# Архитектура UI в PROJECTS_EXCEL

**Навигация:** [Главная](../../README.md) → [PROJECTS_EXCEL](../README.md) → [Документация](README.md) → Архитектура UI в PROJECTS_EXCEL

## Содержание

- [1. Общая модель](#1-общая-модель)
- [2. Два разных контекста](#2-два-разных-контекста)
- [3. Контракт элемента](#3-контракт-элемента)
- [4. Фабрики тегов и контролов](#4-фабрики-тегов-и-контролов)
- [5. Полный запуск страницы: стек вызовов](#5-полный-запуск-страницы-стек-вызовов)
- [6. Расположение в контейнерах](#6-расположение-в-контейнерах)
- [7. Контрол Form и поля](#7-контрол-form-и-поля)
- [8. От кнопки в XML до Shape](#8-от-кнопки-в-xml-до-shape)
- [9. Правила Binding](#9-правила-binding)
- [10. Обновление данных](#10-обновление-данных)
- [11. Маршрутизация событий Shapes](#11-маршрутизация-событий-shapes)
- [12. Перерисовка и пересчёт расположения](#12-перерисовка-и-пересчёт-расположения)
- [13. Валидация формы](#13-валидация-формы)
- [14. Стили](#14-стили)
- [15. Жизненный цикл и владение ресурсами](#15-жизненный-цикл-и-владение-ресурсами)
- [16. Добавление новых реализаций](#16-добавление-новых-реализаций)
- [17. Границы текущей реализации](#17-границы-текущей-реализации)
- [18. Проверка](#18-проверка)
- [19. Разделение тегов и контролов в исходниках](#19-разделение-тегов-и-контролов-в-исходниках)
- [20. Вызовы составных элементов через интерфейсы](#20-вызовы-составных-элементов-через-интерфейсы)
- [21. Контракт проверки разметки](#21-контракт-проверки-разметки)
- [22. Вложенные схемы Page](#22-вложенные-схемы-page)
- [23. Теги контролов в XML namespace](#23-теги-контролов-в-xml-namespace)
- [24. Одиночная таблица и список таблиц](#24-одиночная-таблица-и-список-таблиц)
- [25. Контекст данных и источник элементов](#25-контекст-данных-и-источник-элементов)
- [26. Grid со строками и столбцами](#26-grid-со-строками-и-столбцами)
- [27. Явное значение поля](#27-явное-значение-поля)
- [28. Раскладка демонстрационной страницы](#28-раскладка-демонстрационной-страницы)
- [29. Скомпилированный stylePipeline](#29-скомпилированный-stylepipeline)
- [30. Фиксированная рабочая область и ширина полей](#30-фиксированная-рабочая-область-и-ширина-полей)
- [31. Режимы событий PEB](#31-режимы-событий-peb)
- [32. Каталоги полей и раскладки форм](#32-каталоги-полей-и-раскладки-форм)
- [Размер Shape и выравнивание внутри ячеек](#размер-shape-и-выравнивание-внутри-ячеек)
- [33. Зависимости кадрового поиска и автозаполнение](#33-зависимости-кадрового-поиска-и-автозаполнение)
- [34. Шаблоны ячеек таблицы](#34-шаблоны-ячеек-таблицы)


Документ описывает текущую реализацию: превращение XML-разметки в объекты VBA, расположение контролов на листе, привязки данных, обработку событий и освобождение ресурсов. Цепочки вызовов соответствуют исходному коду. Ограничения текущей реализации обозначены отдельно.

## 1. Общая модель

У страницы есть три связанных представления:

1. XML-документ: теги, атрибуты, привязки и стили.
2. Сохраняемое дерево элементов: объекты VBA, живущие до освобождения страницы.
3. Представление в Excel: значения и оформление ячеек, объединённые диапазоны и Shapes.

Изменение значения может обновить существующий контрол без повторного чтения XML. Если изменение влияет на расположение, дерево повторно измеряется и отрисовывается. Сохранение объектов дерева не означает, что каждый Shape обязательно сохраняется между проходами Render.

~~~text
XAML-файл
  → obj_UiPageDefinition: исходный XML
  → obj_UiRenderContext.Build: дерево obj_IUiElement
  → Measure → Arrange → Render
  → ячейки и Shapes на Worksheet
~~~

Общие модули запускают страницу, находят фабрику или передают событие. Логика формы, поля, кнопки и таблицы принадлежит соответствующим классам.

## 2. Два разных контекста

### obj_UiBindingContext: данные и уведомления

Хранит именованные источники и событие ValueChanged. Text, Resources, Commands, Data и Form — обычные имена времени выполнения, а не встроенные типы.

SetValue и SetObject создают источник при необходимости и добавляют именованное значение или объект. TryGetValue читает значение, TrySetPathValue записывает значение по пути.

~~~vba
Dim bindings As obj_UiBindingContext

Set bindings = New obj_UiBindingContext
If Not bindings.Initialize() Then Exit Sub

bindings.SetValue "Form", "EventName", "Конференция"
bindings.SetValue "Text", "Title", "Создание события"
bindings.SetValue "Resources", "PrimaryButton", "primaryButton"
~~~

В Commands через SetObject помещаются объекты команд. В Data можно поместить источник таблицы или другой объект данных.

Контекст передаётся в рендер. Несколько страниц могут использовать один BindingContext. Отдельный контекст рендера на каждом листе не означает автоматического копирования данных.

### obj_UiRenderContext: конкретная страница

Принадлежит Worksheet и содержит всё необходимое для построения и обслуживания его UI.

| Объект / свойство | Назначение |
| --- | --- |
| TargetWorksheet | Лист назначения |
| PageDefinition | Исходный документ страницы |
| BindingContext | Данные, команды и уведомления |
| m_root | Корневой элемент сохраняемого дерева |
| Router | Маршруты событий ячеек и Shapes |
| Styles | Каталог стилей страницы |
| m_forms | Формы по имени для ValidateForm |
| m_elements | Именованные элементы для InvalidateVisual |
| m_needsMeasure | Признак необходимости пересчитать расположение |

~~~text
ex_UiRuntime
  └─ контексты листов
      └─ obj_UiRenderContext
          ├─ obj_UiPageDefinition
          ├─ obj_UiStyleCatalog
          ├─ obj_UiEventRouter
          ├─ реестры форм и элементов
          ├─ ссылка на obj_UiBindingContext
          └─ корневой obj_IUiElement
              └─ дочерние элементы
~~~

ex_UiRuntime индексирует контексты по имени листа. fn_TryGetContext сначала проверяет принадлежность листа ThisWorkbook. Реестры содержат ссылки на объекты дерева, а не копии контролов. Повторяющиеся непустые имена элементов и форм вызывают ошибку.

## 3. Контракт элемента

Каждый визуальный тег представлен объектом obj_IUiElement.

| Метод | Ответственность |
| --- | --- |
| Configure(definition, context, source, diagnostic) | Прочитать описание, получить контекст и наследуемый источник, создать внутренние объекты |
| Measure(rows, columns, diagnostic) | Вернуть размер в строках и столбцах |
| Arrange(row, column, diagnostic) | Назначить положение относительно переданного начала |
| Render(diagnostic) | Применить представление к Excel |
| Validate(errors) | Проверить данные и добавить сообщения в Collection |
| Dispose() | Освободить детей, ссылки и подписки |

Перед Configure документ проверяется общим валидатором схем. Configure выполняется при построении дерева. Measure, Arrange и Render могут вызываться многократно для тех же объектов.

Размеры выражаются в ячейках, а не в пикселях. row и column — смещения с единицы относительно начала родителя. rowSpan и columnSpan задают размер обычного контрола. Проверяемые атрибуты расположения должны быть положительными целыми числами.

Boolean-результат сообщает об успехе, diagnostic передаёт ошибку вверх. Отображение сообщения относится к вызывающему слою; валидация формы сама диалоги не открывает.

## 4. Фабрики тегов и контролов

Есть два уровня создания: по имени тега и по атрибуту type.

### Фабрика визуального тега

ex_UiElementFactory сопоставляет имя тега объекту obj_IUiElementFactory. Встроенные регистрации создаются при первом fn_Create.

| Тег | Объект |
| --- | --- |
| page | obj_UiPageElement |
| grid | obj_UiGridElement |
| stackPanel | obj_UiStackPanelElement с orientation horizontal или vertical |
| control | obj_UiControlElement |

obj_UiElementFactory создаёт объект, вызывает Configure и возвращает obj_IUiElement. При неудачной настройке освобождает созданный объект.

styles — метаданные оформления. Панель пропускает этот узел при создании визуальных детей, а каталог стилей читает его отдельно.

field обрабатывается самой формой. Это специальное описание поля внутри контрола Form, а не универсальный тег общей фабрики.

### Фабрика типа контрола

obj_UiControlElement передаёт type в ex_UiControlFactory. Реестр содержит объекты obj_IUiControlFactory.

| type | Реализация |
| --- | --- |
| Label | obj_UiLabelControl |
| Button | obj_UiButtonControl |
| Table | obj_UiTableControl |
| Input | obj_UiFieldControl |
| Select | obj_UiFieldControl |
| Form | obj_UiFormControl |

~~~text
<control type="Button">
  → ex_UiElementFactory.fn_Create
  → obj_UiElementFactory.Create
  → New obj_UiControlElement
  → obj_IUiElement.Configure
      → ex_UiControlFactory.fn_Create
      → obj_UiControlFactory.Create
      → New obj_UiButtonControl
      → obj_IUiControl.Initialize
      → obj_IUiControl.Configure
~~~

obj_IUiElement и obj_IUiControl имеют одинаковые сигнатуры Configure, Measure, Arrange, Render, Validate и Dispose. У obj_IUiControl дополнительно есть Initialize. Тег obj_UiControlElement делегирует этот цикл конкретному контролу; старый контракт, возвращавший Range из Measure, удалён. Диапазоны остаются внутренней деталью листовых контролов. Конкретный контрол создаётся один раз для элемента и повторно используется при проходах рендера.

Адаптер хранит изолированную копию описания. В ней нормализует привязки и меняет рассчитанные координаты. Исходный DOM страницы не изменяется.

## 5. Полный запуск страницы: стек вызовов

По умолчанию runtime ищет `<sheet name>.xaml` в выбранном каталоге UI. `ThisWorkbook::ui.pageMappings` задаёт явную карту, например `MainPage=MainPage.xaml|Tests=TestsPage.xaml`. Пути относительны каталогу UI и могут включать подпапки. При наличии ключа несопоставленные листы пропускаются; пустые записи, повторные/отсутствующие листы, отсутствующие XAML и выход за каталог — ошибки конфигурации. Относительные импорты стилей считаются от каталога выбранного XAML.

PersonalEventBuilder хранит разметку PROD/TEST в `ui/PersonalEventBuilder`. PROD сопоставляет `MainPage=MainPage.xaml` и выбирает `ui.pageClass=obj_PEB_PgMain`; TEST — `MainPage=TestsPage.xaml` и `ui.pageClass=obj_PEB_PgTests`. Обе страницы реализуют `obj_IPage`, инициализируют `obj_PageBase`, регистрируют привязки и команды контроллера, загружают данные до общего рендера. `obj_PEB_PgTestsController` предоставляет каталог, выбор запуска/превью, прогресс и отчёты. Все контролы, включая чекбоксы таблицы, принадлежат XAML-дереву; отдельного тестового UI или накладок после рендера нет.

~~~text
ex_UiRenderer
  → ex_UiRuntime.fn_RenderPages / fn_RenderActivePage
    → private_RenderPage(worksheet, ...)
      → ex_UiPageLoader.fn_TryLoad
        → obj_UiPageDefinition
      → New obj_UiRenderContext
      → Initialize(worksheet, definition, uiFolder, bindings)
      → Build
        → ex_UiElementFactory.fn_Create(documentElement, context, "", diagnostic)
          → рекурсивное Configure контейнеров и детей
      → Dispose предыдущего контекста листа, если он есть
      → сохранить новый контекст в реестр runtime
      → Styles.BeginPage
      → private_ClearUi
      → Styles.ApplyPagePipeline
      → RenderTree
        → root.Measure
        → root.Arrange(1, 1)
        → root.Render
~~~

fn_RenderPages сначала освобождает существующие контексты и обходит листы книги. Лист без соответствующего файла пропускается. fn_RenderActivePage работает с активным Worksheet этой книги.

Build создаёт дерево, но ещё не рисует его. Затем runtime загружает стили, очищает прежнее представление и применяет оформление страницы.

RenderTree измеряет дерево, назначает координаты, затем выполняет Render. На этапе Render отключает Application.EnableEvents, чтобы собственная запись значений не стала пользовательским редактированием. Прежнее состояние событий восстанавливается и при обработанной ошибке.

Ошибка Build или Render передаётся вверх. Результат проверяется вызывающими слоями.

## 6. Расположение в контейнерах

### grid и page

obj_UiGridElement передаёт всем детям одну базовую точку. Каждый ребёнок добавляет свои row и column. Размер контейнера определяется максимумом размеров детей с учётом смещений.

Это размещение по координатам ячеек, без определений строк и столбцов WPF Grid.

~~~xml
<grid>
  <control type="Label"
           name="Title"
           text="Событие"
           row="1"
           column="2"
           columnSpan="4"/>
  <control type="Input"
           name="EventNameInput"
           value="{Binding Path=Form.EventName}"
           row="3"
           column="2"
           columnSpan="4"/>
</grid>
~~~

При начале контейнера в A1 заголовок занимает B1:E1, редактор — B3:E3.

### stackPanel

При orientation="vertical" высоты детей складываются, ширина берётся максимальная. Следующий ребёнок размещается ниже предыдущего.

При orientation="horizontal" складываются ширины, высота берётся максимальная. Следующий ребёнок размещается справа.

ArrangePanel повторно измеряет детей, передаёт каждому текущую точку и сдвигает её на размер ребёнка. Поэтому Measure может вызываться несколько раз за один проход страницы.

## 7. Контрол Form и поля

Контрол type="Form" требует name, source и orientation. Отдельного визуального тега form больше нет. Непосредственные дети — field и control. Вложенный stackPanel не нужен: obj_UiFormControl использует obj_UiStackPanelElement внутри как алгоритм последовательного расположения.

dataContext="{Binding Path=Form}" выбирает источник коротких привязок детей. Атрибут не создаёт источник и не заполняет данные.

~~~xml
<control type="Form" name="EventDraftForm"
      dataContext="{Binding Path=Form}"
      row="7"
      column="2"
      orientation="vertical"
      labelPosition="left">
  <field name="EventName"
       value="{Binding Path=EventName}"
         label="Название"
         type="text"
         required="true"/>
  <field name="Notes"
       value="{Binding Path=Notes}"
         label="Заметки"
         type="text"
         rowSpan="3"/>
  <control type="Button"
           name="SubmitEventDraft"
           caption="Сохранить"
           command="{Binding Path=Commands.SubmitFormCommand}"/>
</control>
~~~

Поле obj_UiFormField создаёт внутреннюю панель, подпись и редактор. Описания подписи и редактора — отдельные копии XML-узла. Они не добавляются в документ страницы.

~~~text
obj_UiFormControl
  └─ внутренняя вертикальная obj_UiStackPanelElement
      ├─ obj_UiFormField: EventName
      │   └─ горизонтальная obj_UiStackPanelElement
      │       ├─ obj_UiControlElement → Label
      │       └─ obj_UiControlElement → Input
      ├─ obj_UiFormField: Notes
      │   └─ obj_UiStackPanelElement → Label + Input
      └─ obj_UiControlElement → Button
~~~

~~~text
obj_UiFormControl.obj_IUiControl_Configure
  → New obj_UiFormField
  → obj_IUiElement.Configure(fieldNode, context, source, diagnostic)
      → проверить имя, type, labelPosition, label и Boolean-атрибуты
      → определить value; по умолчанию "{Binding Path=имяПоля}"
      → разобрать источник и путь
      → создать внутреннюю панель
      → создать описание подписи и адаптер Label
      → создать описание редактора и адаптер Input / Select
      → добавить детей в панель поля
  → добавить поле в панель формы
~~~

labelPosition="left" создаёт горизонтальную композицию, "top" — вертикальную. Каждая подпись и каждый редактор занимают одну ячейку. Форма и field не поддерживают rowSpan, columnSpan и labelColumnSpan; такие атрибуты вызывают ошибку конфигурации. Общий размер формы определяется композицией её детей.

Оформление и labelPosition читаются сначала с поля, затем с формы, затем используются значения по умолчанию. required проверяется для самого поля; не следует считать все атрибуты автоматически наследуемыми.

В примере форма начинается в B7. Подпись занимает B7:C7, редактор — D7:G7. Следующее поле начинается на строке 8 и занимает три строки.

Типы поля: text, select и checkbox. У text перенос включён по умолчанию; отдельного multiline нет. Имя поля должно содержать от 1 до 25 символов, label обязателен.

Checkbox Shape создаёт и обслуживает obj_UiFieldControl. obj_UiCellBinding только переносит значения и обслуживает подписки; он не создаёт checkbox и не проверяет формы.

## 8. От кнопки в XML до Shape

После построения дерева кнопка проходит общий цикл:

~~~text
RenderTree
  → obj_IUiElement.Measure
    → obj_UiControlElement.obj_IUiElement_Measure
      → Button.obj_IUiControl_Measure(rows, columns, diagnostic)
      → внутреннее измерение диапазона Excel
  → obj_IUiElement.Arrange
    → obj_UiControlElement.obj_IUiElement_Arrange
      → Button.obj_IUiControl_Arrange(row, column, diagnostic)
      → obj_UiControlBase.ArrangePosition
      → внутреннее обновление описания и диапазона
  → obj_IUiElement.Render
    → obj_UiControlElement.obj_IUiElement_Render
      → Button.Render(diagnostic)
        → оформить диапазон и создать Shape
        → разрешить привязку команды
        → зарегистрировать обработчик Shape в context.Router
~~~

При повторном Render прежний Shape кнопки удаляется и создаётся актуальное представление.

Таблица проходит тот же контракт адаптера. Её Measure учитывает данные источника; общему runtime не требуется отдельный алгоритм измерения таблиц.

## 9. Правила Binding

Квалифицированные пути:

~~~xml
style="{Binding Path=Resources.PrimaryButton}"
command="{Binding Path=Commands.GenerateTablesCommand}"
value="{Binding Path=Form.EventName}"
~~~

Внутри dataContext="{Binding Path=Form}" доступен короткий путь:

~~~xml
value="{Binding Path=EventName}"
~~~

Старая явная запись также поддерживается:

~~~xml
value="{Binding Source=Form; Path=EventName}"
~~~

Наследуемый source передаётся аргументом при построении дерева. Контейнер может изменить его своим атрибутом source. Адаптер преобразует привязку в явную запись Source / Path в собственной копии описания.

Для пути с точкой парсер учитывает контекст: если первый сегмент зарегистрирован как источник, выбирается этот источник. Иначе при наличии наследуемого source путь остаётся относительным. Person.Name внутри Form может быть вложенным путём Form, если Person не зарегистрирован отдельным источником.

Имена источников нужно настроить до Build. Запасного источника Form или Text нет. Синтаксически разобранная привязка сама по себе не гарантирует существование значения: чтение выполняется соответствующим контролом или слоем привязок.

## 10. Обновление данных

### Из кода в UI

Изменение должно пройти через API контекста:

~~~vba
bindings.SetValue "Form", "EventName", "Новое название"
~~~

~~~text
BindingContext.SetValue / SetObject / TrySetPathValue
  → изменить данные
  → RaiseEvent ValueChanged(sourceName, bindingPath)
    → obj_UiCellBinding: обновить связанную ячейку
    → obj_UiControlElement: обработать привязанные атрибуты
        → obj_IUiBindingTarget.RefreshBindings для Label / Button
        → либо context.InvalidateMeasure для пересчёта расположения
~~~

У кнопки обновление привязок включает повторное разрешение команды. У редактора подписка обновляет значение ячейки.

Произвольное присваивание свойства объекта, помещённого в контекст, не создаёт ValueChanged автоматически. Нужно использовать методы контекста. Наблюдения за любыми VBA-свойствами, словарями и коллекциями здесь нет.

### Из UI в данные

Редактируемый Input поддерживает двустороннюю передачу значения. При readOnly обратная запись отключена; попытка редактирования возвращает представление к значению источника.

~~~text
ThisWorkbook: изменение ячеек
  → ex_UiBindings.fn_HandleCellChange(target)
    → ex_UiRuntime.fn_TryGetContext(target.Parent, context)
    → context.Router.DispatchCells(target)
      → обработчик зарегистрированной ячейки
      → obj_UiFieldControl.obj_IUiEventHandler_HandleEvent("change", ...)
        → obj_UiCellBinding.HandleCellChange
          → BindingContext.TrySetPathValue
          → ValueChanged
          → команда onChange, если задана
    → context.FlushLayout
~~~

Запись из кода и пользовательское редактирование — разные причины изменения. Само ValueChanged не следует считать пользовательским событием onChange.

## 11. Маршрутизация событий Shapes

Router сопоставляет имя Shape или адрес ячейки обработчику obj_IUiEventHandler. Получатель знает своё поведение; мост не выбирает его по типу контрола.

~~~text
Щелчок по Shape
  → ex_UiBridge.fn_OnShapeClick
    → ex_RuntimeLifecycle.fn_TryEnter
    → Application.Caller: имя Shape
    → найти RenderContext активного листа
    → context.Router.DispatchShape
      → handler.HandleEvent("click", shapeName)
        → команда кнопки / действие Select / обработчик Checkbox
    → context.FlushLayout
    → ex_RuntimeLifecycle.fn_Leave
~~~

При смене выделения Router.Broadcast передаёт dismiss обработчикам Shapes. Select закрывает выпадающее представление; остальные получатели игнорируют ненужное событие.

Маршруты принадлежат странице и освобождаются с её контекстом.

## 12. Перерисовка и пересчёт расположения

InvalidateVisual(name, diagnostic) находит именованный элемент и вызывает Render. Дерево не строится заново, Measure и Arrange этим методом не выполняются. Метод подходит, когда положение и размер остаются актуальными.

InvalidateMeasure помечает расположение как требующее пересчёта. FlushLayout выполняет отложенную работу.

~~~text
InvalidateMeasure
  → m_needsMeasure = True

FlushLayout
  → если рендер уже идёт или пересчёт не нужен: завершить
  → отключить события Excel
  → сбросить и заново инициализировать маршруты
  → снять объединения и очистить содержимое прежнего измеренного диапазона
  → RenderTree: Measure → Arrange → Render
  → восстановить события Excel
~~~

После событий Shapes и ячеек FlushLayout вызывается автоматически. Код вне этих маршрутов должен вызвать его сам, если изменение требует пересчёта:

~~~vba
Dim context As obj_UiRenderContext
Dim diagnostic As String

If ex_UiRuntime.fn_TryGetContext(ThisWorkbook.Worksheets("MainPage"), context) Then
    context.InvalidateMeasure
    If Not context.FlushLayout(diagnostic) Then
        ' Обработать diagnostic в вызывающем слое.
    End If
End If
~~~

Сейчас пересчитывается вся страница. Пересчёт отдельного поддерева и планировщик зависимостей пока не реализованы.

## 13. Валидация формы

~~~vba
Dim errors As Collection
Set errors = New Collection

If Not context.ValidateForm("EventDraftForm", errors) Then
    ' Показать ошибки и отменить сохранение на уровне приложения.
End If
~~~

~~~text
RenderContext.ValidateForm
  → найти зарегистрированный obj_IUiControl по имени
  → Form.Validate
  → obj_IUiElement.Validate
  → Field.Validate
  → прочитать значение через BindingContext
  → добавить ошибку required при незаполненном поле
~~~

Панель проверяет всех детей, а не только до первой ошибки. Обычные контролы без правил проверки возвращают успех. Проверка не сохраняет данные и не показывает диалоги: это делает команда приложения.

Произвольные пользовательские валидаторы отдельным контрактом пока не представлены.

## 14. Стили

Страница владеет obj_UiStyleCatalog. BeginPage загружает общие каталоги, явно перечисленные в `styleSources` корня страницы, затем локальные определения. Например, `styleSources="CommonControlStyles.xaml"` импортирует соседний файл; runtime не предполагает имя каталога. ApplyPagePipeline применяет правила страницы; контролы используют тот же каталог.

styles не занимает место в layout. Разрешение style через Resources означает получение значения из BindingContext; именованный стиль затем обслуживается каталогом страницы.

Inline-значения styles в примерах пишутся многострочно:

~~~xml
<rule target="sheet"
      styles="{
        backColor:#1F2329;
        fontColor:#E5E7EB;
        rowHeight:21;
        overflow:wrap;
      }"/>
~~~

## 15. Жизненный цикл и владение ресурсами

Новые классы имеют Class_Initialize, Class_Terminate и публичный Dispose. Class_Terminate вызывает Me.Dispose. Интерфейсный obj_IUiElement_Dispose также делегирует в публичный метод.

Class_Initialize может быть пустым: настройка с аргументами выполняется в Initialize или Configure.

~~~text
Замена страницы / завершение runtime
  → RenderContext.Dispose
    → Router.Dispose: отсоединить обработчики
    → root.Dispose: освободить детей и подписки
    → Dispose элементов реестра: очистить также частично построенные объекты
    → Styles.Dispose
    → освободить реестры и ссылку на BindingContext
    → PageDefinition.Dispose
    → освободить Worksheet и остальные ссылки
~~~

Объект может быть доступен и через дерево, и через реестр, поэтому повторное Dispose безопасно. Панель освобождает детей; адаптер — контрол и WithEvents-ссылку; форма и поле — внутренние панели.

Class_Terminate не заменяет явный Dispose: взаимные ссылки VBA-объектов могут помешать автоматическому уничтожению. Общий BindingContext освобождает его владелец после контекстов страниц.

## 16. Добавление новых реализаций

Для нового визуального тега нужны obj_IUiElement и фабрика obj_IUiElementFactory. Регистрация выполняется через ex_UiElementFactory.fn_Register. Контейнеры создают его через общий fn_Create.

Для нового type внутри control нужны obj_IUiControl и obj_IUiControlFactory, зарегистрированная через ex_UiControlFactory.fn_Register. Контракт дерева обеспечит адаптер.

Для событий контрол или отдельный объект действия реализует obj_IUiEventHandler и регистрирует маршрут в Router. Для непосредственного обновления связанных свойств можно реализовать obj_IUiBindingTarget.

Оба реестра создают встроенные фабрики лениво и сохраняют предварительно зарегистрированные реализации. Поэтому новый тег или тип и замена стандартной фабрики могут быть зарегистрированы до первого построения страницы.

Новый контрол самостоятельно владеет Shapes, подписками и дополнительными объектами. Его Dispose освобождает их. Общему мосту не требуется ветка с названием нового контрола.

## 17. Границы текущей реализации

- Рендер работает в координатах ячеек Excel; это не полный набор возможностей WPF/XAML.
- Сохраняется дерево объектов, но при полном пересчёте обновляется вся страница.
- Теги и контролы имеют одинаковый протокол рендера, но разные интерфейсы: тег control соединяет дерево элементов с конкретным контролом.
- Реактивность основана на ValueChanged от BindingContext, а не на наблюдении произвольных объектов VBA.
- Контрол Form принимает field и control; field создаёт сам контрол, а общая фабрика тегов его не интерпретирует.
- Проверка обязательных полей есть, отдельного расширяемого контракта валидаторов пока нет.
- Measure используется также при Arrange; реализация не гарантирует один вызов измерения за проход.
- Render не является транзакцией Excel: ошибка после начала записи может оставить частично изменённое представление.

## 18. Проверка

Из корня репозитория:

~~~powershell
./PROJECTS_EXCEL/WorkbookUpdater/Test-UiElements.ps1
~~~

Сценарий импортирует реальные VBA-исходники в новую временную книгу. Проверяет текущую страницу, неизменность XML, расположение, реактивные привязки, маршруты событий, обязательные поля, Checkbox, readOnly, независимость контекстов страниц, повторный рендер и освобождение подписок. Пользовательские книги не изменяются.

Test-Updater.ps1 отдельно проверяет установку и обновление VBA-модулей. Успех установщика не заменяет проверку поведения UI.

## 19. Разделение тегов и контролов в исходниках

Классы тегов находятся в vba/common/[4] elements: obj_UiPageElement, obj_UiGridElement, obj_UiStackPanelElement и obj_UiControlElement. Каждый реализует obj_IUiElement; контейнеры также реализуют obj_IUiContainer.AddChild. Общего переключателя режима grid/stack больше нет: каждый класс отвечает за свой алгоритм.

Готовые контролы находятся в vba/common/[3] controls и реализуют obj_IUiControl. Form — такой же зарегистрированный тип, как Button или Table. Его внутреннее дерево использует obj_UiStackPanelElement как кирпичик расположения. obj_UiFormField — внутренний элемент формы; его описание создаёт Form.

~~~text
<page>                              obj_UiPageElement
  <stackPanel>                      obj_UiStackPanelElement
    <control type="Form">           obj_UiControlElement
      внутренний контрол            obj_UiFormControl
        внутренняя панель           obj_UiStackPanelElement
          поле                      obj_UiFormField
            подпись и редактор      obj_UiControlElement → Label / Input
~~~

Контекст хранит формы через obj_IUiControl, а не через конкретный класс. RegisterForm и ValidateForm не знают внутреннего устройства Form и вызывают общий Validate. События по-прежнему идут через obj_IUiEventHandler; прежний HandleCellChange удалён из контракта контролов.

Для листовых контролов obj_UiControlBase предоставляет ConfigurePosition, SetPosition, ArrangePosition и GetSize. Measure возвращает числовые размеры, а Arrange назначает окончательные координаты. Работа с Range скрыта внутри контролов. Нормализация Binding остаётся у тега control; поведение рендера и валидации принадлежит выбранному контролу.

Пример регистрации дополнительного имени для имеющейся реализации:

~~~vba
Dim tagFactory As New obj_UiElementFactory
Dim controlFactory As New obj_UiControlFactory

tagFactory.Initialize "stackpanel"
ex_UiElementFactory.fn_Register "flow", tagFactory

controlFactory.Initialize "form"
ex_UiControlFactory.fn_Register "DraftForm", controlFactory
~~~

Теперь flow создаёт стандартную панель, а control type="DraftForm" — стандартную форму. Для собственного поведения вместо этих фабрик регистрируются классы, реализующие соответствующий интерфейс фабрики. Runtime и мост событий менять не требуется.

## 20. Вызовы составных элементов через интерфейсы

Page, Form и FormField хранят внутренние панели как obj_IUiElement и вызывают только Configure, Measure, Arrange, Render, Validate и Dispose. Добавление детей выполняется через obj_IUiContainer.AddChild. Публичные ConfigurePanel, MeasurePanel, ArrangePanel, RenderPanel, ValidatePanel, DisposePanel и ConfigureField удалены; внутренние алгоритмы имеют префикс private_.

Чтобы собрать контейнер вручную, владелец создаёт изолированное описание без детей через cloneNode(False), задаёт внутренние row и column равными 1, вызывает интерфейсный Configure и добавляет детей через obj_IUiContainer. Дополнительный режим buildChildren не требуется. Form учитывает своё смещение один раз, а внутренняя панель работает от переданного начала. Поле получает настройки владельца из parentNode своего описания; настройка самого поля выполняется через obj_IUiElement.Configure. Исходный документ страницы сохраняется неизменным.

## 21. Контракт проверки разметки

До создания дерева RenderContext.Build вызывает obj_UiMarkupValidator.Validate. При ошибках Build возвращает False, собирает сообщения и не вызывает фабрику визуальных объектов.

~~~text
Загрузка XML
  → obj_UiMarkupValidator
    → схема page из фабрики тегов
    → проверка атрибутов, детей и текстового содержимого
    → для control: схема выбранного type из фабрики контролов
    → рекурсивная проверка детей
  → при успехе Build → Configure → Measure → Arrange → Render
~~~

Фабрики реализуют obj_IUiMarkupSchemaProvider.GetSchema. Схема получается из тех же реестров, которые создают объекты. Фабрика без провайдера схемы отвергается с диагностикой; автоматической permissive-схемы нет. Зарегистрированные дополнительные имена используют схему своей фабрики.

obj_UiMarkupSchema описывает:
- атрибуты: имя, тип, обязательность, допустимые значения, поддержку Binding и максимальную длину;
- альтернативные обязательные атрибуты через RequireAnyAttribute, например text или caption;
- локальных детей со схемой и минимальным/максимальным количеством;
- разрешение детей из реестра визуальных тегов;
- необходимость разрешить type через реестр контролов.

Неизвестные атрибуты отвергаются. Имена атрибутов проверяются с учётом регистра. Объявления XML namespace разрешены отдельно. Проверяются положительные и неотрицательные целые, числовые значения, Boolean true/false и перечисления. Для привязок проверяется синтаксис без чтения данных. Существование источника или объекта остаётся задачей настройки и разрешения Binding.

Page разрешает ровно один grid и не более одного локального styles; другие непосредственные дети запрещены. Grid и StackPanel разрешают визуальные элементы, кроме page; styles внутри них запрещён. Листовые контролы не разрешают детей. Form разрешает только локальный field и зарегистрированный control. Специализированные field и узлы styles имеют локальные схемы, не требуют регистрации как визуальные теги.

Пробельный текст, комментарии XML и инструкции обработки не создают визуальных детей. Содержательный текст или CDATA между тегами отвергаются.

Ошибки возвращаются как Collection объектов obj_UiMarkupDiagnostic с Path, Member, Message и Describe. Путь содержит имя тега, порядковый номер и name, если он задан. Валидатор собирает несколько ошибок за проход и не открывает диалогов.

~~~text
/page/grid[1]/control[1][@name='Title']: Unsupported attribute. [columnSapn]
~~~

Это проверка структуры страницы, а не проверка пользовательских данных. required у field задаёт правило последующей ValidateForm; проверка разметки проверяет только корректность самого атрибута required. Внешний CommonControlStyles.xaml пока загружается каталогом стилей отдельно и этим проходом страницы не проверяется.

Тест Test-UiElements.ps1 включает отрицательные примеры для неизвестных атрибутов и типов, неверных чисел и Boolean, неправильной вложенности, лишнего текста, повторного styles, отсутствующих обязательных атрибутов и некорректных Binding. Положительные примеры проверяют текущую страницу и зарегистрированные расширения.

## 22. Вложенные схемы Page

Схема Page содержит AddChild "grid", Nothing, 1, 1 и AddChild "styles", metadata, 0, 1. Первый аргумент обозначает дочерний тег, второй — объект правил для этого узла, последние два — минимальное и максимальное количество непосредственных детей. Nothing для grid означает получение схемы из зарегистрированной фабрики Grid. metadata — локальная схема, возвращаемая private_StylesSchema: она разрешает controlStyle и stylePipelineStage, а вложенные схемы описывают layer и rule с их атрибутами. Это объект правил, а не содержимое XML или данные стилей.

При встрече styles валидатор берёт metadata из правила Page, передаёт её вместе с XML-узлом styles в рекурсивную проверку и повторяет тот же алгоритм для его детей. XML остаётся неизменным; привязки Binding здесь не используются. Ограничение name для самой Page снято, поскольку она не создаёт Shape с префиксом btn_.

## 23. Теги контролов в XML namespace

В актуальной разметке общий control с атрибутом type заменён тегами controls:form, controls:button, controls:label, controls:input, controls:select и controls:table. В XML используется одно двоеточие. Префикс controls связан с urn:excelprototype:controls; стандартные page, grid, stackPanel, field и styles используют urn:excelprototype:profiles.

~~~xml
<page xmlns="urn:excelprototype:profiles"
      xmlns:controls="urn:excelprototype:controls">
  <grid>
    <controls:form name="Draft" dataContext="{Binding Path=Form}" orientation="vertical">
      <field name="Title"
       value="{Binding Path=Title}" label="Название" type="text"/>
      <controls:button name="Save" caption="Сохранить"/>
    </controls:form>
  </grid>
</page>
~~~

Фабрика выбирается по namespaceURI и localName. Префикс можно заменить любым другим, связанным с тем же URI. URI сравнивается с учётом регистра. Имена зарегистрированных тегов сохраняют прежнюю нечувствительность к регистру. fn_Register у фабрик принимает необязательный namespaceUri; по умолчанию используются URI соответствующего реестра.

Для namespace контролов ex_UiElementFactory создаёт obj_UiControlElement. Тот записывает localName во внутренний атрибут type изолированной копии описания, чтобы конкретные контролы и селекторы стилей могли использовать своё внутреннее представление. Исходная разметка type не содержит. Общая схема control и механизм DispatchControls удалены: схема запрашивается непосредственно у фабрики именованного контрола.

Правила AddChild содержат namespaceUri, по умолчанию profiles. Валидатор считает детей и разрешает локальные схемы по паре URI и имени. Form разрешает field из profiles и зарегистрированные теги namespace controls. Старый control type и неизвестные namespace отвергаются. Предыдущие примеры с control type в этом документе описывают этап до миграции; для текущей разметки используется синтаксис этого раздела.

## 24. Одиночная таблица и список таблиц

`controls:table` (`obj_UiTableControl`) отрисовывает ровно одну таблицу из источника `obj_IUiTableSource`. Источник `obj_UiRawTable` возвращает одну таблицу; источник с нулём или несколькими таблицами отклоняется при измерении.

`controls:tableList` (`obj_UiTableListControl`) принимает тот же контракт источника, включая `obj_UiRawTableList`. Атрибут `gapRows` определяет число пустых строк между таблицами и поддерживается только списком. Пустой список занимает одну ячейку и ничего не выводит.

Представления Raw/Smart, индексные политики XAML, автоматические имена Excel и преобразование готовых таблиц описаны в [руководстве представлений](TablePresentation-API.md).

```xml
<controls:table name="EventTable"
                itemsSource="{Binding Path=Data.EventTable}"
                showHeaders="true"/>

<controls:tableList name="GeneratedTables"
                    itemsSource="{Binding Path=Data.Tables}"
                    gapRows="1"
                    showHeaders="true"/>
```

Стек вызовов списка: `obj_IUiControl.Measure` → разрешение источника → `GetTable(index)` → фабрика одиночного `table` → `obj_IUiTableTarget.SetTable` → `obj_IUiControl.Configure` → `Measure`. Список хранит дочерние контролы и их высоты. `Arrange` размещает их сверху вниз с учётом `gapRows`; `Render` и `Validate` делегируются каждому дочернему контролу. При повторном измерении старые дочерние контролы освобождаются и список строится заново из текущего источника. `Dispose` освобождает детей, но не уничтожает принадлежащие приложению таблицы данных.

`obj_IUiTableTarget` — отдельный интерфейс передачи одной таблицы внутреннему контролу. Он позволяет обходиться без временных источников в общем binding-контексте. Все операции жизненного цикла и рендера проходят через `obj_IUiControl`.

Каждый дочерний `table` самостоятельно пишет название, заголовки и значения и применяет стиль через каталог страницы. `style` и `showHeaders` списка наследуются его дочерними таблицами; отдельные шаблоны для каждого элемента списка пока не предусмотрены. Исходный XML не изменяется: список создаёт копии определения для своих детей.

## 25. Контекст данных и источник элементов

Атрибут `source` удалён из схем визуальных тегов и контролов: валидатор отклоняет старую разметку. Вместо него используются `dataContext` и `itemsSource`.

`dataContext` задаёт объект, относительно которого разрешаются короткие пути привязок. Значение задаётся выражением Binding, например `{Binding Path=Form}` или `{Binding Path=Data.Draft}`. Контекст наследуется через `page`, `grid`, `stackPanel` и составные контролы. Собственный `dataContext` заменяет унаследованный для элемента и его детей. Форме достаточно унаследованного контекста; обязательного локального атрибута нет. При построении дерева проверяется, что заданный контекст существует и является объектом.

```xml
<grid dataContext="{Binding Path=Data.Draft}">
    <controls:form name="DraftForm"
                   orientation="vertical">
        <field name="EventName"
               label="Название"
               type="text"
               value="{Binding Path=EventName}"/>
    </controls:form>
</grid>
```

Здесь `EventName` разрешается как источник `Data`, путь `Draft.EventName`. Вложенный путь `Person.Name` разрешается как `Draft.Person.Name`. Если первый сегмент пути совпадает с зарегистрированным runtime-источником (например, `Commands` или `Resources`), путь считается абсолютным. Привязки команд и ресурсов поэтому работают внутри формы независимо от её контекста.

`itemsSource` задаёт данные конкретного контрола и не меняет контекст его детей. Для `table` требуется одна таблица, для `tableList` допускается коллекция, для `select` — источник вариантов выбора. Контролы по-прежнему сохраняют собственные контракты этих данных.

Разрешение контекста централизовано в `ex_UiElementFactory.fn_Create` → `ex_UiBindingRuntime.fn_TryDataContext`, до вызова `Configure`. Внутри текущего контракта Configure унаследованный контекст передаётся строкой полного пути; runtime-парсер разделяет имя зарегистрированного источника и вложенный путь. Это не создаёт дополнительных runtime-источников. Чтение объекта зарегистрированного источника без вложенного пути поддерживает `BindingContext.TryGetValue`.

При изменении поля обратная привязка записывает полный разрешённый путь, а уведомления BindingContext обновляют отображаемые значения. Прямые изменения произвольного VBA-объекта без уведомления BindingContext автоматически не наблюдаются. Смена самого атрибута dataContext в XML требует повторного построения дерева.

`Source` внутри внутреннего нормализованного выражения `{Binding Source=Data; Path=Draft.EventName}` обозначает имя runtime-источника и не является атрибутом разметки. Более ранние описания архитектуры с атрибутом `source` следует читать как историю реализации; актуальные правила приведены в этом разделе.

## 26. Grid со строками и столбцами

При наличии `grid.rowDefinitions` или `grid.columnDefinitions` контейнер переключается на сетку. Определения принадлежат пространству имён profiles. `rowDefinition` и `columnDefinition` имеют обязательный атрибут `size`: `auto`, положительное целое число, `*` или положительный целый вес, например `2*`. Неописанная ось содержит одну строку или один столбец `auto`.

```xml
<grid>
    <grid.columnDefinitions>
        <columnDefinition size="auto"/>
        <columnDefinition size="1"/>
        <columnDefinition size="auto"/>
    </grid.columnDefinitions>
    <grid.rowDefinitions>
        <rowDefinition size="auto"/>
    </grid.rowDefinitions>
    <controls:form name="Draft"
                   dataContext="{Binding Path=Form}"
                   orientation="vertical"/>
    <controls:button name="Update"
                     column="3"
                     caption="Обновить"/>
</grid>
```

Дочерние `row` и `column` — индексы треков начиная с 1. `rowSpan` и `columnSpan` — число занимаемых треков. Они извлекаются контейнером и не используются как смещения ребёнка в ячейках Excel. Выход за определения отклоняется при Configure. У самой сетки координаты задают положение в родителе.

Расчёт каждой оси выделен в `obj_UiGridAxis`. Measure сбрасывает размеры, измеряет детей, сначала увеличивает одиночные `auto`-треки, затем распределяет недостающий размер детей с несколькими треками между их `auto`-треками. Фиксированные размеры не увеличиваются. Пустые `auto`-треки занимают минимум одну ячейку Excel. Звёздочные треки делят оставшееся место с целочисленным округлением и минимумом в одну ячейку на трек. Для них требуется соответствующий `rowSpan` или `columnSpan` на самой сетке; недостаточная область и отсутствие её размера возвращают диагностику.

Arrange использует суммы размеров предшествующих треков для вычисления координат ребёнка. Интерфейс `obj_IUiLayoutSlot.SetSize` передаёт размер выделенной области адаптеру контрола. Label, Button, Input и Select заполняют эту область. Form и таблицы сохраняют размер содержимого. Порядок соседних элементов в XML не задаёт их позиции.

Сетка без определений сохраняет прежний режим координат относительно её начала, чтобы существующие страницы продолжали работать. `stackPanel` остаётся последовательным контейнером. Эта реализация использует целые ячейки Excel, не пиксели; она не реализует полный WinUI-контракт ограниченного Measure и перерасчёт переноса текста по доступной ширине. Вложенная сетка со звёздочными треками должна иметь собственный явно заданный размер; размеры её ячейки родителя автоматически не становятся ограничением Measure.

## 27. Явное значение поля

У каждого `field` обязателен атрибут `value` с выражением привязки, например `value="{Binding Path=Notes}"`. Имя поля служит идентификатором и не задаёт путь данных. Неявное создание Binding по имени удалено; отсутствие value отклоняется валидатором до рендера. Источник и значение должны быть зарегистрированы приложением явно.

## 28. Раскладка демонстрационной страницы

Слева вертикальная панель содержит форму, пустую строку и список таблиц. Справа кнопки Update page, Reset и Generate tables расположены вертикально, каждая занимает две строки и два столбца Excel. Auto-столбец левой области учитывает ширину таблиц и формы; между областями остаются два столбца. Reset очищает EventName и Notes, возвращает Category к Meeting, очищает коллекцию таблиц и запускает обновление страницы.

## 29. Скомпилированный stylePipeline

`obj_UiStylePipeline` один раз компилирует правила при BeginPage и хранит runtime-области. Полный RenderTree очищает их и применяет default-правила листа по измеренной странице до контролов с их именованными стилями и атрибутами. CompleteRender применяет остальные default-правила по XML-порядку и последними восстанавливает накладки выделения. Стадии, слои и правила учитывают enabled. Явный Styles.ApplyStage сохраняет XML-порядок всех целей. Таблицы регистрируют section, header, rows и отдельные строки с tableRow odd/even; grid и stackPanel — границы раскладки. Полосы задаются controlPart с part=row;tags=even. Элемент выбирается через element, глубина — elementDepth. Синтаксис и границы совместимости с PrototypeNew описаны в StylePipeline-API.md.

## 30. Фиксированная рабочая область и ширина полей

PEB резервирует B:Q как фиксированную рабочую область из 16 колонок шириной 16. R отделяет её от панели действий S:T. Ширины и абсолютные позиции заданы в MainPage.xaml; изменение полей формы не меняет колонки листа и расположение боковой панели.

`columnSpan` формы резервирует её полную ширину. Измеренные поля должны помещаться; переполнение останавливает раскладку с диагностикой. `columnSpan` поля задаёт ширину в колонках листа: по умолчанию 1 при верхней подписи. Подпись и редактор занимают весь span. При левой подписи явная ширина включает одну колонку подписи и требует минимум две; прежняя неявная раскладка остаётся двухколоночной.

Горизонтальная форма `columnSpan="16"` может содержать шестнадцать одноколоночных полей или разные ширины с суммой до 16. Короткая форма оставляет резерв пустым. Для более широкой формы нужно явно изменить раскладку страницы.

`controls:candidates` читает ширины из `alignToForm` и проецирует колонки источника в соответствующие позиции. Отсутствующие поля пусты; алиасы и записи выбора сохраняются. `obj_IUiColumnLayout` передаёт общему renderer положительный span начальной колонки и нули перекрытых. Ячейки заголовка/тела объединяются по span, стили регистрируются на всю ширину. Smart-таблицы не допускают объединённых span. Сопоставление кандидатов остаётся в `obj_UiCandidatesControl`; стандартная таблица только выводит готовую раскладку.

Ширина выровненной таблицы равна сумме span, а не резерву формы: отпуск и его кандидаты занимают 14 колонок, лечение — 16. Необязательные `columnsBefore="Header A|Header B"` и `columnsAfter="Header C"` добавляют колонки вне выровненной части. Каждая занимает одну колонку листа; имя должно совпадать с отображаемым заголовком и не дублировать основные/дополнительные колонки. Цвета источника и выбор записей сохраняются.

Заголовки кандидатов изначально пусты и остаются пустыми для колонок без значений во всём результате. obj_UiCandidatesSource реализует obj_IUiHeaderPresentation, отделяющий видимую подпись от структурного заголовка. Пустые строки, пробелы, Empty и Null не раскрывают заголовок; ноль, False и ошибки — раскрывают. Необязательный контракт меняет только текст, сохраняя алиасы, стили, span и выбор. Источники без контракта используют обычные заголовки.

`workspaceFirstColumn` и `workspaceLastColumn` задают включительные границы относительно начала контейнера; при дополнительных колонках оба обязательны. В PEB это 1 и 16 (B:Q). До измерения и размещения проверяется вхождение всей таблицы. Отсутствие заголовка, повторы или нехватка места дают ошибку до записи ячеек. Форма остаётся закреплённой; columnsBefore расширяет влево, columnsAfter — вправо. Отпуск допускает две дополнительные колонки справа, лечение — ни одной; слева места нет в обоих вариантах. Дополнительное пространство или сдвиг формы задаются в разметке; автоматически область и панель не перемещаются.

Необязательный атрибут страницы `zoom="80%"` или `zoom="80"` задаёт целый масштаб 10–400. После успешной полной отрисовки он имеет приоритет над стилями и применяется к окнам книги с целевым листом. Без атрибута страница масштаб не переопределяет.

## 31. Режимы событий PEB

Поля приказа занимают две верхние строки. Ниже кнопки режимов образуют две колонки: прибытие слева, выбытие справа. Каждая резервирует пять позиций, четыре заполнены. Режимы различают лечение, ежегодный, лечебный и семейный отпуск. Активная кнопка использует activeModeButton. wsConfig задаёт режим по умолчанию, подписи, направление, вид формы, сопоставления полей и алиасы кандидатов. Направление хранится в состоянии страницы, а не в видимом Type.

Контроллер хранит Form.ModeId, Form.ModeDirection и Form.ModeForm независимо от видимых полей. Режимы используют общий PersonalEventBuilder/MainPage.xaml. EventFields определяет каждое поле один раз; EventLayouts содержит Leave (14 fieldRef) и Treatment (16). PersonnelForm привязывает layout к Form.ModeForm. Смена режима очищает ключ кадрового поиска и объединение зависимостей старой/новой формы, кандидатов и выбор, затем перестраивает страницу. Посторонние поля и принятый приказ сохраняются. Обычное обновление использует прежний binding-контекст, не очищая поля. Рабочая область, панели и масштаб заданы один раз. Используется полная перестройка, не частичный рендер.

Колонки кандидатов проецируются из настроенных полей формы: ручные поля остаются пустыми, цвета сохраняются. Выбор применяет только настроенные сопоставления; вручную введённые рапорт, входящий документ и больница не стираются. Event.Type берётся из направления. Event.Date использует Absence.DepartureDate только при выбытии; прибытие вводится вручную до настройки источника. Последняя запись отсутствия выбирается прежними правилами; режимы пока не фильтруют категории отсутствия. Описываемый здесь этап UI не включает реализацию экспорта.

## 32. Каталоги полей и раскладки форм

Кнопки используют горизонтальный ModesPanel с вертикальными ArrivalModes и DepartureModes. У каждого columnSpan=4 и horizontalAlignment=stretch; дети без собственного span получают ширину из клонированного определения. Явная ширина сохраняется и отклоняется при превышении родителя. Stretch требует вертикальной ориентации и явного columnSpan. rowSpan=5 резервирует пять строк без кнопок-заполнителей. Старые stack без stretch сохраняют размеры. Позиции следуют порядку элементов; отдельные row/column не нужны.

На уровне страницы fieldCatalog имеет уникальное имя и обычные определения полей. formLayouts требует уникальное имя и ссылку на каталог, содержит именованные formLayout с упорядоченными fieldRef. Каждый fieldRef задаёт имя поля и необязательное переопределение columnSpan; прочие свойства редактора принадлежат каталогу. Определения и ссылки используют profiles namespace и не создают визуальные контролы.

controls:form выбирает группу через layoutSource, вариант — через layout (литерал или Binding). Оба атрибута нужны вместе; форма с выбранной раскладкой не содержит inline-полей и контролов. Обычные inline-формы сохраняются. До построения дерева схемы проверяют каталоги/раскладки, ex_UiFormLayouts — имена и ссылки, включая неактивные варианты. Настройка клонирует выбранные поля по порядку ссылок и применяет span, не меняя исходный XML. GetDefinition возвращает эту форму, поэтому выравнивание кандидатов использует те же поля. Отсутствие каталога, варианта или поля останавливает рендер; молчаливого выбора по умолчанию нет.

## Размер Shape и выравнивание внутри ячеек

`controls:button` и заголовок `controls:select` поддерживают статические `shapeWidth`, `shapeHeight`, `shapeHorizontalAlignment`, `shapeVerticalAlignment`. Размер — положительная дробь до единицы: `1`, `1/5`, `1/2`. Числитель и знаменатель целые; ноль, отрицательные и значения больше единицы запрещены. Выравнивание — `left|center|right` и `top|center|bottom`. Атрибуты размещают Shape, а не текст.

`shapeWidth="1/5" shapeHeight="1/2" shapeHorizontalAlignment="left" shapeVerticalAlignment="top"` помещает кнопку 20×20 в верхний левый угол области 100×40. Центрирование по обеим осям даёт смещения 40 и 10. Для объединённых ячеек и диапазонов берётся вся выделенная контролу область.

Без атрибутов обычная кнопка занимает всю область. Редактируемый select по умолчанию показывает стрелку в правой десятой части на полную высоту; нередактируемый заполняет область. Геометрия считается при создании Shape и после итоговых ширин/высот от стилей страницы. Масштаб окна учитывает Excel. После ручного изменения ячеек нужен новый проход раскладки.

## 33. Зависимости кадрового поиска и автозаполнение

Форма задаёт lookup.key и список lookup.dependents через `|` в wsConfig. Изменение ключа перед поиском очищает зависимые поля и выбор; удаление ключа очищает кандидатов после фиксации ячейки Excel. Смена режима очищает обе группы, включая ключ. Рапорт, больница и документы не входят в кадровые зависимости. Редактор ключа, как и поля разрешения приказа, использует фон lookupKeyField.

Form.AutoFillPersonnel получает personnel.autoFill.default и сохраняется между режимами. Зависимые поля каталога привязывают readOnly к этому Boolean. Включённое автозаполнение применяет кандидата и запрещает ручной ввод через binding; выключенное сохраняет значения, допускает редактирование и пропускает применение кандидата. Повторное включение применяет выбранную запись. Чекбокс может вызвать onChange после изменения; PEB перестраивает страницу для прав редакторов без повторного поиска. Универсальные input/form readOnly принимают Boolean-привязки и литералы.

## 34. Шаблоны ячеек таблицы

controls:table допускает одну обёртку controls:table.cellTemplates с controls:cellTemplate. Каждому нужен columnAlias; mode — replace по умолчанию либо overlay; внутри ровно один обычный контрол. Цель — существующий алиас/заголовок, новая колонка не создаётся. Повторная, отсутствующая или неоднозначная цель останавливает рендер.

Для каждой строки создаётся локальный контекст: Item — запись obj_IUiRowRecords либо источника выбора, иначе словарь заголовков/алиасов; Cell.Value — исходное значение ячейки. Используйте caption="{Binding Path=Cell.Value}" или caption="{Binding Path=Item.ScenarioName}", не меняя DataContext страницы. Text и Commands сохраняют область страницы. obj_UiRawTable.Initialize/InitializeEmpty принимают aliases после title, отделяя подписи от стабильных полей. SetRecords прикрепляет оригинальные записи, включая скрытые поля. Начальные шаблоны допускают button, label, checkbox; редакторы и контейнеры отклоняются.

Кнопка шаблона получает запись через obj_UiCommand.ExecuteWithPayload и следует области ячейки с placement=moveAndSize, если не задан другой. Replace скрывает текст ячейки Shape-контрола форматом отображения, сохраняя значение; overlay оставляет текст и требует Shape. Контролы, маршруты, локальные контексты и прежние форматы освобождаются при повторном рендере/Dispose. Для пустого источника контролы не создаются. Результаты PEB используют PreviewLink и определяют документ по ScenarioId/StepId записи, без координат колонок и разбора имён Shape.