# Архитектура UI в PROJECTS_EXCEL

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

Configure выполняется при построении дерева. Measure, Arrange и Render могут вызываться многократно для тех же объектов.

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

Для листа MainPage runtime ищет MainPage.xaml в папке UI выбранного проекта.

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

source="Form" выбирает источник коротких привязок детей. Атрибут не создаёт источник и не заполняет данные.

~~~xml
<control type="Form" name="EventDraftForm"
      source="Form"
      row="7"
      column="2"
      orientation="vertical"
      labelPosition="left"
      labelColumnSpan="2"
      fieldWidth="4">
  <field name="EventName"
         label="Название"
         type="text"
         required="true"/>
  <field name="Notes"
         label="Заметки"
         type="text"
         height="3"/>
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

labelPosition="left" создаёт горизонтальную композицию, "top" — вертикальную. labelColumnSpan задаёт число столбцов подписи, fieldWidth — число столбцов редактора, height — число строк редактора.

Размеры и оформление читаются сначала с поля, затем с формы, затем используются значения по умолчанию. required проверяется для самого поля; не следует считать все атрибуты автоматически наследуемыми.

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

Внутри source="Form" доступен короткий путь:

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

Каталог obj_UiStyleCatalog принадлежит странице. BeginPage загружает общие стили из CommonControlStyles.xaml и описание текущей страницы. ApplyPagePipeline применяет правила оформления страницы; конкретные контролы применяют свои стили через каталог контекста.

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