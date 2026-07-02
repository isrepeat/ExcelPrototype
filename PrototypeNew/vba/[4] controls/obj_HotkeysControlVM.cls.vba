VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_HotkeysControlVM"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Implements obj_IControl

' Hotkeys control:
' - itemsSource ожидается как Collection(obj_ConfigEntry)
'   * Key   = текст/идентификатор действия, который попадет в колонку Action
'   * Value = пользовательская комбинация клавиш, которая попадет в колонку Hotkey
'   * Attr сейчас не используется, но оставлен совместимым с ConfigTable-подходом
' - контрол сам рендерит Apply-row + Excel ListObject из 2 колонок: Action / Hotkey
' - кнопка Apply является shape над колонкой Hotkey, как отдельная команда всей таблицы
' - при Apply читается именно текущая таблица на листе, а не исходный itemsSource,
'   потому что пользователь редактирует комбинации прямо в ячейках Excel
' - фактическую регистрацию Application.OnKey делает rt_HotkeyRuntime через PageBase
' - строки Action/Hotkey являются данными страницы, поэтому snapshot-ит их страница.
'   Контрол только применяет текущие строки и синхронизирует их обратно в itemsSource.
Private Const HOTKEY_COL_COUNT As Long = 2
Private Const DEFAULT_ACTION_METHOD As String = "RuntimeHandleHotkeyAction"

Private m_ControlBase As obj_ControlBase
Private m_ControlName As String
Private m_Page As obj_IPage
' m_RuntimeControlKey нужен для маршрута клика по Apply-shape на сам Hotkeys VM.
Private m_RuntimeControlKey As String
' m_ActionControlKey нужен для маршрутов хоткеев на внешний action context
' (обычно controller страницы), чтобы нажатие хоткея вызывало не VM, а бизнес-действие.
Private m_ActionControlKey As String
Private m_RuntimeTableName As String
Private m_ItemsSourceRaw As String
Private m_TableNameRaw As String
Private m_ActionMethodRaw As String
Private m_ActionMethodName As String
Private m_ControlLayout As obj_ControlLayout
' Переиспользуем ConfigTableViewItem как легкий adapter над Collection(obj_ConfigEntry):
' это дает совместимый формат строк без отдельной модели HotkeyRow.
Private m_ConfigTableViewItem As obj_ConfigTableViewItem
' Контекст, на котором будет вызван actionMethod при срабатывании хоткея.
' В XML обычно приходит через dataContext="{PageRuntimeSource='...Controller'}".
Private m_ActionCallbackContext As Object
Private m_IsConfigured As Boolean
Private m_IsDisposed As Boolean

Private Sub Class_Terminate()
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    obj_IControl_Dispose
    On Error GoTo 0
End Sub

' //
' // Interface
' //
Private Function obj_IControl_Initialize(ByVal page As obj_IPage) As Boolean
    m_IsDisposed = False
    m_IsConfigured = False
    Set m_Page = page
    obj_IControl_Initialize = True
End Function

Private Sub obj_IControl_Dispose()
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Set m_ControlBase = Nothing
    Set m_ControlLayout = Nothing
    Set m_ConfigTableViewItem = Nothing
    Set m_ActionCallbackContext = Nothing
    Set m_Page = Nothing
    m_ControlName = VBA.vbNullString
    m_ItemsSourceRaw = VBA.vbNullString
    m_TableNameRaw = VBA.vbNullString
    m_ActionMethodRaw = VBA.vbNullString
    m_ActionMethodName = VBA.vbNullString
    m_RuntimeControlKey = VBA.vbNullString
    m_ActionControlKey = VBA.vbNullString
    m_RuntimeTableName = VBA.vbNullString
    m_IsConfigured = False
    On Error GoTo 0
End Sub

Private Sub obj_IControl_Configure(ByVal controlNode As Object)
    Dim pageBase As obj_PageBase
    Dim resolvedItems As Collection
    Dim configTable As obj_ConfigTable
    Dim dataContext As Object
    Dim resolvedActionMethod As Variant

    m_IsConfigured = False
    Set m_ControlLayout = Nothing
    Set m_ConfigTableViewItem = Nothing
    Set m_ControlBase = Nothing
    Set m_ActionCallbackContext = Nothing
    m_RuntimeControlKey = VBA.vbNullString
    m_ActionControlKey = VBA.vbNullString
    m_RuntimeTableName = VBA.vbNullString

    Set pageBase = m_Page.GetPageBase()
    Set m_ControlBase = New obj_ControlBase
    If Not m_ControlBase.Initialize(m_Page) Then Exit Sub
    If Not m_ControlBase.Configure(pageBase, controlNode, "Hotkeys", "hotkeys", m_ControlName) Then Exit Sub

    ' Configure только подготавливает runtime contract:
    ' читает XML-атрибуты, резолвит itemsSource/dataContext/actionMethod и layout bounds.
    ' Никаких Excel-объектов здесь не создаем — это делает Render.
    m_ItemsSourceRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "itemsSource")))
    If VBA.Len(m_ItemsSourceRaw) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Hotkeys: itemsSource is not specified for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    m_TableNameRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "tableName")))
    ' actionMethod — имя метода на dataContext, который будет вызван при хоткее.
    ' Для PrsnlEvntBuilder это RuntimeHandleHotkeyAction(actionText).
    m_ActionMethodRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "actionMethod")))
    If VBA.Len(m_ActionMethodRaw) = 0 Then m_ActionMethodRaw = DEFAULT_ACTION_METHOD

    ' dataContext отделяет универсальный контрол от конкретной страницы:
    ' Hotkeys VM знает, как зарегистрировать маршруты, но не знает бизнес-логику actions.
    Set dataContext = m_ControlBase.DataContext
    If dataContext Is Nothing Then Set dataContext = m_Page
    Set m_ActionCallbackContext = dataContext

    If Not ex_BindingRuntime.fn_TryResolveValueBinding(m_ActionMethodRaw, dataContext, resolvedActionMethod) Then Exit Sub
    If VBA.IsObject(resolvedActionMethod) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Hotkeys: actionMethod binding must resolve to scalar method name for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If
    m_ActionMethodName = VBA.Trim$(VBA.CStr(resolvedActionMethod))
    If VBA.Len(m_ActionMethodName) = 0 Then m_ActionMethodName = DEFAULT_ACTION_METHOD

    Set m_ControlLayout = New obj_ControlLayout
    If Not m_ControlLayout.TryReadFromNode(controlNode, "Hotkeys", m_ControlName, "style") Then Exit Sub
    If (m_ControlLayout.ColEnd - m_ControlLayout.ColStart + 1) < HOTKEY_COL_COUNT Then
        VBA.MsgBox "PrototypeNew: hotkeys control '" & m_ControlName & "' requires at least 2 columns.", VBA.vbExclamation, "PrototypeNew / Hotkeys layout"
        Exit Sub
    End If

    Set pageBase = m_ControlBase.PageBase
    If pageBase Is Nothing Then Exit Sub
    ' Два ключа регистрируются в PageBase:
    ' 1) m_RuntimeControlKey -> Me, чтобы shape click попал в RuntimeHandleApplyClick.
    ' 2) m_ActionControlKey  -> controller/dataContext, чтобы hotkey попал в actionMethod.
    m_RuntimeControlKey = "hotkeys|" & VBA.LCase$(VBA.Trim$(m_ControlLayout.LayoutSheetName & "|" & m_ControlName))
    m_ActionControlKey = m_RuntimeControlKey & "|action-target"
    If Not ex_RuntimeSourceResolver.fn_TryResolveItemsSource(pageBase.RuntimeSources, m_ItemsSourceRaw, resolvedItems) Then Exit Sub
    If Not private_TryBuildConfigTable(resolvedItems, configTable) Then Exit Sub

    Set m_ConfigTableViewItem = New obj_ConfigTableViewItem
    If Not m_ConfigTableViewItem.Initialize(configTable) Then Exit Sub

    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render()
    Dim ws As Worksheet
    Dim boundsRange As Range
    Dim writeRange As Range
    Dim valueBlock As Variant
    Dim rowsToWrite As Long
    Dim maxRows As Long
    Dim idx As Long
    Dim rowOut As Long
    Dim entryItems As list__obj_ConfigEntryViewItem
    Dim entryViewItem As obj_ConfigEntryViewItem
    Dim configEntry As obj_ConfigEntry
    Dim tableObj As ListObject
    Dim targetTableName As String
    Dim pageBase As obj_PageBase
    Dim totalRowsToClear As Long

    ' Render строит физическое представление на листе:
    ' Apply shape row -> ListObject(Action/Hotkey) -> controlPart registrations -> shape route.
    If Not m_IsConfigured Then Exit Sub

    Set pageBase = m_ControlBase.PageBase
    If pageBase Is Nothing Then Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Sub
    Set ws = private_GetWorksheetByName(pageBase, m_ControlLayout.LayoutSheetName)
    If ws Is Nothing Then Exit Sub
    If m_ConfigTableViewItem Is Nothing Then Exit Sub
    If Not m_ConfigTableViewItem.TryResyncEntryItemsFromModel() Then Exit Sub
    Set entryItems = m_ConfigTableViewItem.EntryItems
    If entryItems Is Nothing Then Exit Sub

    ' Контрол не рендерит частично. Если строк itemsSource больше, чем layout bounds,
    ' показываем ошибку, иначе пользователь получил бы неполный конфиг хоткеев.
    maxRows = m_ControlLayout.RowEnd - m_ControlLayout.RowStart + 1
    rowsToWrite = 1 + entryItems.Count
    If rowsToWrite < 2 Then rowsToWrite = 2
    totalRowsToClear = 1 + rowsToWrite
    If totalRowsToClear > maxRows Then
        VBA.MsgBox "PrototypeNew: hotkeys control '" & m_ControlName & "' does not fit into allocated bounds.", VBA.vbExclamation, "PrototypeNew / Hotkeys layout"
        Exit Sub
    End If

    ' Пишем всю таблицу одним Value2-блоком: так быстрее и меньше шансов оставить
    ' полуобновленный диапазон при ошибке.
    ReDim valueBlock(1 To rowsToWrite, 1 To HOTKEY_COL_COUNT)
    valueBlock(1, 1) = "Action"
    valueBlock(1, 2) = "Hotkey"

    For idx = 1 To entryItems.Count
        Set entryViewItem = entryItems.Item(idx)
        If entryViewItem Is Nothing Then GoTo ContinueItem
        Set configEntry = entryViewItem.Model
        If configEntry Is Nothing Then GoTo ContinueItem

        rowOut = idx + 1
        valueBlock(rowOut, 1) = configEntry.Key
        valueBlock(rowOut, 2) = configEntry.Value

ContinueItem:
    Next idx

    Set boundsRange = ws.Range( _
        ws.Cells(m_ControlLayout.RowStart, m_ControlLayout.ColStart), _
        ws.Cells(m_ControlLayout.RowEnd, m_ControlLayout.ColStart + HOTKEY_COL_COUNT - 1))

    ' Перед созданием нового ListObject надо удалить старые таблицы в этом диапазоне:
    ' Excel не разрешает пересекающиеся ListObject и упадет на ws.ListObjects.Add.
    If Not private_TryDeleteIntersectingTables(ws, boundsRange) Then Exit Sub
    boundsRange.UnMerge
    boundsRange.ClearContents

    Set writeRange = ws.Range( _
        ws.Cells(m_ControlLayout.RowStart + 1, m_ControlLayout.ColStart), _
        ws.Cells(m_ControlLayout.RowStart + rowsToWrite, m_ControlLayout.ColStart + HOTKEY_COL_COUNT - 1))
    writeRange.NumberFormat = "@"
    writeRange.Value2 = valueBlock

    On Error GoTo EH_TABLE
    Set tableObj = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=writeRange, XlListObjectHasHeaders:=xlYes)
    On Error GoTo 0

    targetTableName = private_BuildTableName(ws)
    If VBA.Len(targetTableName) > 0 Then
        On Error Resume Next
        tableObj.Name = targetTableName
        On Error GoTo 0
    End If

    On Error Resume Next
    tableObj.TableStyle = "TableStyleMedium4"
    tableObj.ShowAutoFilter = False
    On Error GoTo 0
    m_RuntimeTableName = VBA.Trim$(tableObj.Name)

    ' controlPart registrations дают style pipeline возможность красить отдельные
    ' колонки Hotkeys-контрола через selector type=hotkeys;part=...
    If Not private_RegisterColumnPart(ws, "action", writeRange.Columns(1)) Then Exit Sub
    If Not private_RegisterColumnPart(ws, "hotkey", writeRange.Columns(2)) Then Exit Sub
    ' Apply — это shape над таблицей в колонке Hotkey. Excel cells сами по себе
    ' не умеют OnAction, поэтому используем тот же bridge-паттерн, что и Button.
    If Not private_TryRenderApplyButton(ws, ws.Cells(m_ControlLayout.RowStart, m_ControlLayout.ColStart + 1)) Then Exit Sub
    If Not private_TryRegisterRuntimeControl() Then Exit Sub
    Exit Sub

EH_TABLE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "Hotkeys: failed to create table for control '" & m_ControlName & "': " & Err.Description
#End If
End Sub

Private Function obj_IControl_Measure( _
    ByVal controlNode As Object, _
    ByRef outSpanRows As Long, _
    ByRef outSpanColls As Long, _
    Optional ByVal dataContext As Object _
) As Boolean
    obj_IControl_Measure = private_TryMeasureNode(controlNode, outSpanRows, outSpanColls)
End Function

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(attrName))
        Case "itemssource", "tablename", "actionmethod"
            obj_IControl_SupportsAttribute = True
    End Select
End Function

Private Function obj_IControl_IsConfigured() As Boolean
    obj_IControl_IsConfigured = m_IsConfigured
End Function

' //
' // API
' //
Public Function RuntimeHandleApplyClick() As Boolean
    RuntimeHandleApplyClick = Me.RuntimeApplyCurrentRows(True)
End Function

Public Function RuntimeApplyCurrentRows(Optional ByVal showStatus As Boolean = False) As Boolean
    Dim tableObj As ListObject
    Dim dataRange As Range
    Dim rowIndex As Long
    Dim actionText As String
    Dim hotkeyText As String
    Dim appliedRows As Collection
    Dim registeredCount As Long
    Dim actionValidationError As String

    ' Apply — единственный сценарий, где rendered table с листа становится моделью.
    ' Это позволяет пользователю менять Hotkey-ячейки без обновления RuntimeSources.
    If Not private_TryResolveRenderedTableObject(tableObj) Then Exit Function
    Set dataRange = tableObj.DataBodyRange
    If dataRange Is Nothing Then Exit Function

    ' Action-колонка является read-only по смыслу:
    ' пользователь настраивает только Hotkey, а Action задается страницей при Configure.
    ' Поэтому перед сохранением сравниваем лист с m_ConfigTableViewItem —
    ' последней нормальной таблицей VM.
    ' Важный нюанс: порядок Action тоже защищен. Сейчас action text используется как
    ' аргумент бизнес-метода, а строка таблицы задает, какая комбинация относится к
    ' какому action. Если разрешить перестановку, Apply начнет менять семантику строк
    ' неявно. Для reorder нужен отдельный явный сценарий в модели, а не ручная правка UI.
    actionValidationError = VBA.vbNullString
    If Not private_TryValidateActionColumn(dataRange, actionValidationError) Then
        VBA.MsgBox _
            "PrototypeNew: hotkeys Action column was changed and cannot be applied." & VBA.vbCrLf & _
            actionValidationError & VBA.vbCrLf & _
            "The hotkeys table will be reset from the previous valid configuration.", _
            VBA.vbExclamation, _
            "PrototypeNew / Hotkeys"
        If Not private_TryResetRenderedTableFromCurrentConfig() Then Exit Function
        Exit Function
    End If

    Set appliedRows = New Collection

    For rowIndex = 1 To dataRange.Rows.Count
        ' Контракт rendered table:
        ' Col 1 = action argument, Col 2 = пользовательский hotkey text.
        actionText = VBA.Trim$(VBA.CStr(dataRange.Cells(rowIndex, 1).Value2))
        hotkeyText = VBA.Trim$(VBA.CStr(dataRange.Cells(rowIndex, 2).Value2))
        If VBA.Len(actionText) = 0 Then GoTo ContinueRow
        ' Сохраняем строку даже с пустым Hotkey: action остается в конфиге,
        ' но физическая OnKey-привязка для нее не создается.
        If Not private_AddHotkeyEntry(appliedRows, actionText, hotkeyText) Then Exit Function

ContinueRow:
    Next rowIndex

    If appliedRows.Count = 0 Then
        VBA.MsgBox "PrototypeNew: hotkeys control '" & m_ControlName & "' has no action rows to apply.", VBA.vbExclamation, "PrototypeNew / Hotkeys"
        Exit Function
    End If

    If Not private_TryRegisterHotkeyRows(appliedRows, registeredCount) Then Exit Function
    ' После успешной валидации новые строки целиком заменяют модель:
    ' RuntimeItems нужен следующему render, m_ConfigTableViewItem — текущей VM.
    If Not private_TryStoreHotkeyRows(appliedRows, False) Then Exit Function
    If Not private_TrySetCurrentConfigRows(appliedRows) Then Exit Function

    If showStatus Then rt_Messaging.fn_ShowStatusBarSuccess "Hotkeys applied: " & VBA.CStr(registeredCount), 3
    RuntimeApplyCurrentRows = True
End Function

Public Function RuntimeRegisterBoundRows(Optional ByVal showStatus As Boolean = False) As Boolean
    Dim currentRows As Collection
    Dim registeredCount As Long

    ' Регистрация после Render/Restore работает от текущего ConfigTableViewItem, а не от листа.
    ' Configure уже построил его из bound itemsSource, поэтому это last-known-good модель VM.
    ' UI может быть временно пустым при rerender, но модель страницы от этого не меняется.
    If Not private_TryGetCurrentConfigRows(currentRows) Then Exit Function
    If Not private_TryRegisterHotkeyRows(currentRows, registeredCount) Then Exit Function

    If showStatus Then rt_Messaging.fn_ShowStatusBarSuccess "Hotkeys registered: " & VBA.CStr(registeredCount), 3
    RuntimeRegisterBoundRows = True
End Function

' //
' // Internal
' //
Private Function private_TryRenderApplyButton(ByVal ws As Worksheet, ByVal targetCell As Range) As Boolean
    Dim shp As Shape
    Dim shapeName As String
    Dim callbackMacroRef As String
    Dim pageBase As obj_PageBase
    Dim metaMap As Object

    If ws Is Nothing Then Exit Function
    If targetCell Is Nothing Then Exit Function

    ' Shape сохраняется/переиспользуется по стабильному имени, чтобы repeated render
    ' не плодил новые кнопки.
    shapeName = "btn_" & m_ControlName & "_Apply"
    Set shp = private_GetUiShapeByName(ws, shapeName)
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, targetCell.Left, targetCell.Top, targetCell.Width, targetCell.Height)
        shp.Name = shapeName
    Else
        shp.Left = targetCell.Left
        shp.Top = targetCell.Top
        shp.Width = targetCell.Width
        shp.Height = targetCell.Height
    End If

    shp.Placement = xlMoveAndSize
    shp.TextFrame2.TextRange.Text = "Apply"
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame.Characters.Text = "Apply"
    shp.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp.TextFrame.VerticalAlignment = xlVAlignCenter
    shp.Fill.ForeColor.RGB = VBA.RGB(46, 125, 50)
    shp.Line.ForeColor.RGB = VBA.RGB(27, 94, 32)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = VBA.RGB(0, 0, 0)
    shp.TextFrame.Characters.Font.Color = VBA.RGB(0, 0, 0)

    ' Shape.OnAction принимает только строку макроса. Поэтому все клики идут в bridge,
    ' а bridge уже по активному листу и shapeName находит нужный PageBase route.
    callbackMacroRef = "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!rt_Bridge.fn_OnShapeClick"
    shp.OnAction = callbackMacroRef

    ' Метаданные нужны retained-render cleanup-у: если контрол исчезнет из layout,
    ' PageBase сможет определить, что shape был runtime-частью именно этого контрола.
    Set metaMap = VBA.CreateObject("Scripting.Dictionary")
    metaMap.CompareMode = 1
    metaMap("pn.control") = m_ControlName
    metaMap("pn.style") = VBA.Trim$(m_ControlLayout.StyleName)
    If Not ex_ShapeMetaRuntime.fn_TrySetShapeMetaValues(shp, metaMap) Then Exit Function

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    ' Route клика по Apply: shapeName -> this VM -> RuntimeHandleApplyClick.
    If Not pageBase.RegisterControl(m_RuntimeControlKey, Me) Then Exit Function
    If Not pageBase.RegisterShapeRoute(shp.Name, m_RuntimeControlKey, "RuntimeHandleApplyClick", False) Then Exit Function

    private_TryRenderApplyButton = True
End Function

Private Function private_TryRegisterRuntimeControl() As Boolean
    Dim pageBase As obj_PageBase

    ' Повторная регистрация тем же ключом безопасна: PageBase хранит последний VM.
    ' Она нужна, чтобы TryGetRegisteredControlByName и cleanup видели контрол.
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    private_TryRegisterRuntimeControl = pageBase.RegisterControl(m_RuntimeControlKey, Me)
End Function

Private Function private_RegisterColumnPart( _
    ByVal ws As Worksheet, _
    ByVal partName As String, _
    ByVal columnRange As Range _
) As Boolean
    If ws Is Nothing Then Exit Function
    If columnRange Is Nothing Then Exit Function

    private_RegisterColumnPart = ex_ControlPartsRuntime.fn_RegisterControlPart( _
        ws, _
        "hotkeys", _
        m_ControlName, _
        VBA.LCase$(VBA.Trim$(partName)), _
        columnRange)
End Function

Private Function private_TryParseHotkeyInput( _
    ByVal hotkeyText As String, _
    ByRef outHotkeyKey As String, _
    Optional ByRef outErrorText As String = VBA.vbNullString _
) As Boolean
    Dim normalized As String
    Dim rx As Object
    Dim matches As Object
    Dim matchObj As Object
    Dim modifier1 As String
    Dim modifier2 As String
    Dim keyText As String
    Dim keyToken As String

    outHotkeyKey = VBA.vbNullString
    outErrorText = VBA.vbNullString

    normalized = VBA.UCase$(VBA.Trim$(hotkeyText))
    normalized = VBA.Replace$(normalized, " ", VBA.vbNullString)
    If VBA.Len(normalized) = 0 Then
        outErrorText = "Hotkey is empty."
        Exit Function
    End If

    Set rx = VBA.CreateObject("VBScript.RegExp")
    rx.IgnoreCase = True
    rx.Global = False

    ' UX-контракт хоткеев намеренно уже, чем полный OnKey:
    ' один модификатор (CTRL/ALT/SHIFT) или CTRL+SHIFT / CTRL+ALT + key.
    ' SHIFT+ALT не включен, пока нет явного сценария.
    rx.Pattern = "^(CTRL)\+(SHIFT|ALT)\+([^+]+)$"
    If rx.Test(normalized) Then
        Set matches = rx.Execute(normalized)
        Set matchObj = matches(0)
        modifier1 = VBA.UCase$(VBA.CStr(matchObj.SubMatches(0)))
        modifier2 = VBA.UCase$(VBA.CStr(matchObj.SubMatches(1)))
        keyText = VBA.UCase$(VBA.CStr(matchObj.SubMatches(2)))
        GoTo BuildToken
    End If

    ' Одномодификаторный шаблон: CTRL+<key>, ALT+<key>, SHIFT+<key>.
    rx.Pattern = "^(CTRL|ALT|SHIFT)\+([^+]+)$"
    If rx.Test(normalized) Then
        Set matches = rx.Execute(normalized)
        Set matchObj = matches(0)
        modifier1 = VBA.UCase$(VBA.CStr(matchObj.SubMatches(0)))
        modifier2 = VBA.vbNullString
        keyText = VBA.UCase$(VBA.CStr(matchObj.SubMatches(1)))
        GoTo BuildToken
    End If

    outErrorText = "Expected CTRL+<key>, ALT+<key>, SHIFT+<key>, CTRL+SHIFT+<key>, or CTRL+ALT+<key>."
    Exit Function

BuildToken:
    If Not private_TryMapKeyPart(keyText, keyToken, outErrorText) Then Exit Function
    keyToken = private_NormalizeOnKeyLetterToken(keyToken)

    outHotkeyKey = private_ModifierToOnKeyPrefix(modifier1)
    If VBA.Len(modifier2) > 0 Then outHotkeyKey = outHotkeyKey & private_ModifierToOnKeyPrefix(modifier2)
    outHotkeyKey = outHotkeyKey & keyToken

    private_TryParseHotkeyInput = (VBA.Len(outHotkeyKey) > 0)
End Function

Private Function private_ModifierToOnKeyPrefix(ByVal modifierText As String) As String
    Select Case VBA.UCase$(VBA.Trim$(modifierText))
        Case "CTRL"
            private_ModifierToOnKeyPrefix = "^"
        Case "SHIFT"
            private_ModifierToOnKeyPrefix = "+"
        Case "ALT"
            private_ModifierToOnKeyPrefix = "%"
    End Select
End Function

Private Function private_NormalizeOnKeyLetterToken(ByVal keyToken As String) As String
    ' Application.OnKey трактует заглавные буквы как нажатые с Shift.
    ' Поэтому буквы держим в lowercase, а Shift выражаем только префиксом "+".
    If VBA.Len(keyToken) = 1 Then
        If keyToken Like "[A-Z]" Then
            private_NormalizeOnKeyLetterToken = VBA.LCase$(keyToken)
            Exit Function
        End If
    End If

    private_NormalizeOnKeyLetterToken = keyToken
End Function

Private Function private_TryMapKeyPart( _
    ByVal keyText As String, _
    ByRef outKeyToken As String, _
    Optional ByRef outErrorText As String = VBA.vbNullString _
) As Boolean
    Dim fNumberText As String
    Dim fNumber As Long

    outKeyToken = VBA.vbNullString
    outErrorText = VBA.vbNullString
    keyText = VBA.UCase$(VBA.Trim$(keyText))

    Select Case keyText
        Case VBA.vbNullString
            outErrorText = "Key is empty."
            Exit Function

        Case "CTRL", "ALT", "SHIFT"
            outErrorText = "Key cannot be a modifier."
            Exit Function

        Case "ENTER", "RETURN"
            ' Application.OnKey использует "~" для основного Enter/Return.
            ' "{ENTER}" относится к Enter на numeric keypad и не ловит Ctrl+Enter надежно.
            outKeyToken = "~"

        Case "NUMENTER", "NUMPADENTER"
            outKeyToken = "{ENTER}"

        Case "ESC", "ESCAPE"
            outKeyToken = "{ESC}"

        Case "TAB"
            outKeyToken = "{TAB}"

        Case "BACKSPACE", "BKSP"
            outKeyToken = "{BACKSPACE}"

        Case "DELETE", "DEL"
            outKeyToken = "{DELETE}"

        Case "SPACE"
            outKeyToken = " "

        Case "UP", "DOWN", "LEFT", "RIGHT", "HOME", "END", "PGUP", "PGDN"
            outKeyToken = "{" & keyText & "}"

        Case Else
            If VBA.Len(keyText) = 1 Then
                outKeyToken = keyText
            ElseIf VBA.Left$(keyText, 1) = "F" And VBA.Len(keyText) <= 3 Then
                fNumberText = VBA.Mid$(keyText, 2)
                If VBA.IsNumeric(fNumberText) Then
                    fNumber = VBA.CLng(fNumberText)
                    If fNumber >= 1 And fNumber <= 24 Then outKeyToken = "{" & keyText & "}"
                End If
            End If
    End Select

    If VBA.Len(outKeyToken) = 0 Then
        outErrorText = "Unsupported key '" & keyText & "'. Use a single character, ENTER, NUMENTER, TAB, ESC, DELETE, BACKSPACE, SPACE, arrows, HOME/END, PGUP/PGDN, or F1-F24."
        Exit Function
    End If

    private_TryMapKeyPart = True
End Function

Private Function private_AddHotkeyEntry( _
    ByVal rows As Collection, _
    ByVal actionText As String, _
    ByVal hotkeyText As String _
) As Boolean
    Dim configEntry As obj_ConfigEntry

    If rows Is Nothing Then Exit Function
    actionText = VBA.Trim$(actionText)
    If VBA.Len(actionText) = 0 Then Exit Function

    Set configEntry = New obj_ConfigEntry
    configEntry.Attr = VBA.vbNullString
    configEntry.Key = actionText
    configEntry.Value = VBA.Trim$(hotkeyText)
    rows.Add configEntry

    private_AddHotkeyEntry = True
End Function

Private Function private_TryGetCurrentConfigRows(ByRef outRows As Collection) As Boolean
    Dim entryItems As list__obj_ConfigEntryViewItem
    Dim entryViewItem As obj_ConfigEntryViewItem
    Dim configEntry As obj_ConfigEntry
    Dim rowIndex As Long

    ' m_ConfigTableViewItem — last-known-good состояние Hotkeys VM.
    ' Оно строится в Configure из itemsSource, а после успешного Apply заменяется
    ' целиком новой валидной таблицей. Именно по нему проверяем Action и из него
    ' восстанавливаем UI, если пользователь случайно изменил защищенную колонку.
    Set outRows = Nothing
    If m_ConfigTableViewItem Is Nothing Then Exit Function
    If Not m_ConfigTableViewItem.TryResyncEntryItemsFromModel() Then Exit Function
    Set entryItems = m_ConfigTableViewItem.EntryItems
    If entryItems Is Nothing Then Exit Function

    Set outRows = New Collection
    For rowIndex = 1 To entryItems.Count
        Set entryViewItem = entryItems.Item(rowIndex)
        If entryViewItem Is Nothing Then GoTo ContinueRow
        Set configEntry = entryViewItem.Model
        If configEntry Is Nothing Then GoTo ContinueRow
        If Not private_AddHotkeyEntry(outRows, configEntry.Key, configEntry.Value) Then Exit Function

ContinueRow:
    Next rowIndex

    If outRows.Count = 0 Then Exit Function
    private_TryGetCurrentConfigRows = True
End Function

Private Function private_TryValidateActionColumn( _
    ByVal dataRange As Range, _
    ByRef outError As String _
) As Boolean
    Dim expectedActions As Collection
    Dim rowIndex As Long
    Dim expectedAction As String
    Dim actualAction As String

    outError = VBA.vbNullString
    If dataRange Is Nothing Then Exit Function
    If Not private_TryBuildExpectedActionRows(expectedActions) Then Exit Function

    ' Action — не пользовательское поле. Оно является contract-ключом страницы,
    ' поэтому Apply принимает только тот же набор actions, который уже есть
    ' в текущей нормальной ConfigTable VM.
    ' Сравнение идет по строкам, а не как unordered set: порядок защищен так же,
    ' как и сам текст Action.
    If dataRange.Rows.Count <> expectedActions.Count Then
        outError = "Expected action rows: " & VBA.CStr(expectedActions.Count) & _
            "; actual table rows: " & VBA.CStr(dataRange.Rows.Count) & "."
        Exit Function
    End If

    For rowIndex = 1 To expectedActions.Count
        expectedAction = VBA.Trim$(VBA.CStr(expectedActions.Item(rowIndex)))
        actualAction = VBA.Trim$(VBA.CStr(dataRange.Cells(rowIndex, 1).Value2))
        If VBA.StrComp(actualAction, expectedAction, VBA.vbBinaryCompare) <> 0 Then
            outError = "Row " & VBA.CStr(rowIndex) & ": expected action '" & expectedAction & _
                "', actual action '" & actualAction & "'."
            Exit Function
        End If
    Next rowIndex

    private_TryValidateActionColumn = True
End Function

Private Function private_TryBuildExpectedActionRows(ByRef outActions As Collection) As Boolean
    Dim currentRows As Collection
    Dim rowItem As Variant
    Dim configEntry As obj_ConfigEntry
    Dim actionText As String

    Set outActions = New Collection
    If Not private_TryGetCurrentConfigRows(currentRows) Then Exit Function

    For Each rowItem In currentRows
        Set configEntry = Nothing
        If Not VBA.IsObject(rowItem) Then GoTo ContinueRow
        Set configEntry = rowItem
        If configEntry Is Nothing Then GoTo ContinueRow

        actionText = VBA.Trim$(VBA.CStr(configEntry.Key))
        If VBA.Len(actionText) = 0 Then GoTo ContinueRow
        outActions.Add actionText

ContinueRow:
    Next rowItem

    If outActions.Count = 0 Then Exit Function
    private_TryBuildExpectedActionRows = True
End Function

Private Function private_TryResetRenderedTableFromCurrentConfig() As Boolean
    ' RuntimeItems не меняем: m_ConfigTableViewItem уже содержит предыдущую
    ' валидную конфигурацию. Просто повторно рендерим контрол из нее.
    If m_ConfigTableViewItem Is Nothing Then Exit Function
    obj_IControl_Render
    private_TryResetRenderedTableFromCurrentConfig = True
End Function

Private Function private_TrySetCurrentConfigRows(ByVal rows As Collection) As Boolean
    Dim configTable As obj_ConfigTable
    Dim nextConfigTableViewItem As obj_ConfigTableViewItem

    If rows Is Nothing Then Exit Function
    If Not private_TryBuildConfigTable(rows, configTable) Then Exit Function

    Set nextConfigTableViewItem = New obj_ConfigTableViewItem
    If Not nextConfigTableViewItem.Initialize(configTable) Then Exit Function

    ' Заменяем VM-state только целиком и только после успешной валидации/сохранения.
    ' Частичного обновления строк здесь нет: либо вся таблица стала новой нормальной,
    ' либо остается предыдущий m_ConfigTableViewItem.
    Set m_ConfigTableViewItem = nextConfigTableViewItem

    private_TrySetCurrentConfigRows = True
End Function

Private Function private_TryRegisterHotkeyRows( _
    ByVal hotkeyRows As Collection, _
    ByRef outRegisteredCount As Long _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim seenHotkeys As Object
    Dim routeRows As Collection
    Dim rowItem As Variant
    Dim configEntry As obj_ConfigEntry
    Dim routeEntry As obj_ConfigEntry
    Dim actionText As String
    Dim hotkeyText As String
    Dim hotkeyKey As String
    Dim parseError As String

    outRegisteredCount = 0
    If hotkeyRows Is Nothing Then Exit Function
    If m_Page Is Nothing Then Exit Function

    Set seenHotkeys = ex_Helpers.fn_CreateDictionaryTextCompare()
    Set routeRows = New Collection

    For Each rowItem In hotkeyRows
        Set configEntry = Nothing
        If Not VBA.IsObject(rowItem) Then GoTo ContinueValidateRow
        Set configEntry = rowItem
        If configEntry Is Nothing Then GoTo ContinueValidateRow

        actionText = VBA.Trim$(VBA.CStr(configEntry.Key))
        hotkeyText = VBA.Trim$(VBA.CStr(configEntry.Value))
        If VBA.Len(actionText) = 0 Then GoTo ContinueValidateRow
        If VBA.Len(hotkeyText) = 0 Then GoTo ContinueValidateRow

        ' Парсер живет в контроле, потому что именно Hotkeys VM определяет UX-контракт
        ' пользовательского ввода. Runtime получает уже готовый Application.OnKey token.
        parseError = VBA.vbNullString
        If Not private_TryParseHotkeyInput(hotkeyText, hotkeyKey, parseError) Then
            VBA.MsgBox "PrototypeNew: unsupported hotkey '" & hotkeyText & "' for action '" & actionText & "'. " & parseError, VBA.vbExclamation, "PrototypeNew / Hotkeys"
            Exit Function
        End If
        If seenHotkeys.Exists(hotkeyKey) Then
            VBA.MsgBox "PrototypeNew: duplicate hotkey '" & hotkeyText & "'. Set different combinations before applying.", VBA.vbExclamation, "PrototypeNew / Hotkeys"
            Exit Function
        End If
        seenHotkeys(hotkeyKey) = True
        If Not private_AddRouteEntry(routeRows, hotkeyKey, actionText) Then Exit Function

ContinueValidateRow:
    Next rowItem

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    ' Любая регистрация заменяет все hotkey routes текущей страницы.
    ' Shape/cell routes не трогаются, очищаются только hotkey routes.
    If Not pageBase.ResetHotkeyRoutes() Then Exit Function

    ' Регистрируем controller/dataContext как target для будущих hotkey dispatch.
    ' PageBase потом вызовет m_ActionMethodName на этом объекте через CallByName.
    If Not pageBase.RegisterControl(m_ActionControlKey, m_ActionCallbackContext) Then Exit Function

    For Each rowItem In routeRows
        Set routeEntry = Nothing
        If Not VBA.IsObject(rowItem) Then GoTo ContinueRegisterRow
        Set routeEntry = rowItem
        If routeEntry Is Nothing Then GoTo ContinueRegisterRow

        ' Регистрируем route: hotkey -> action target -> actionMethod(actionText).
        ' Сам Application.OnKey будет установлен внутри rt_HotkeyRuntime.
        If Not pageBase.RegisterHotkeyRouteByKey(routeEntry.Key, m_ActionControlKey, m_ActionMethodName, True, routeEntry.Value) Then Exit Function
        outRegisteredCount = outRegisteredCount + 1

ContinueRegisterRow:
    Next rowItem

    private_TryRegisterHotkeyRows = True
End Function

Private Function private_AddRouteEntry( _
    ByVal rows As Collection, _
    ByVal hotkeyKey As String, _
    ByVal actionText As String _
) As Boolean
    Dim configEntry As obj_ConfigEntry

    If rows Is Nothing Then Exit Function
    actionText = VBA.Trim$(actionText)
    If VBA.Len(hotkeyKey) = 0 Then Exit Function
    If VBA.Len(actionText) = 0 Then Exit Function

    Set configEntry = New obj_ConfigEntry
    configEntry.Attr = VBA.vbNullString
    configEntry.Key = hotkeyKey
    configEntry.Value = actionText
    rows.Add configEntry

    private_AddRouteEntry = True
End Function

Private Function private_TryStoreHotkeyRows( _
    ByVal rows As Collection, _
    ByVal notifyChange As Boolean _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim runtimeSources As obj_PageRuntimeSources
    Dim sourceKey As String

    If rows Is Nothing Then Exit Function
    If Not private_TryResolvePageItemsSourceKey(sourceKey) Then Exit Function

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set runtimeSources = pageBase.RuntimeSources
    If runtimeSources Is Nothing Then Exit Function

    ' После Apply rendered table становится новым источником истины для следующего render.
    ' Иначе PrepareRuntime/Render снова увидят старые default-строки.
    If Not runtimeSources.RemoveItemsSource(sourceKey) Then Exit Function
    If Not runtimeSources.SetItemsSource(sourceKey, rows, notifyChange) Then Exit Function

    private_TryStoreHotkeyRows = True
End Function

Private Function private_TryResolvePageItemsSourceKey(ByRef outSourceKey As String) As Boolean
    Dim rawSource As String
    Dim expressionBody As String
    Dim eqPos As Long
    Dim argName As String
    Dim argValue As String
    Dim quoteChar As String

    outSourceKey = VBA.vbNullString
    rawSource = VBA.Trim$(m_ItemsSourceRaw)
    If VBA.Len(rawSource) < 3 Then Exit Function
    If VBA.Left$(rawSource, 1) <> "{" Then Exit Function
    If VBA.Right$(rawSource, 1) <> "}" Then Exit Function

    expressionBody = VBA.Trim$(VBA.Mid$(rawSource, 2, VBA.Len(rawSource) - 2))
    eqPos = VBA.InStr(1, expressionBody, "=", VBA.vbBinaryCompare)
    If eqPos <= 1 Then Exit Function

    argName = VBA.Trim$(VBA.Left$(expressionBody, eqPos - 1))
    If VBA.StrComp(argName, "PageRuntimeSource", VBA.vbTextCompare) <> 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Hotkeys: writable itemsSource must be PageRuntimeSource for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    argValue = VBA.Trim$(VBA.Mid$(expressionBody, eqPos + 1))
    quoteChar = VBA.Left$(argValue, 1)
    If VBA.Len(argValue) >= 2 And (quoteChar = """" Or quoteChar = "'") And VBA.Right$(argValue, 1) = quoteChar Then
        argValue = VBA.Mid$(argValue, 2, VBA.Len(argValue) - 2)
    End If

    outSourceKey = VBA.LCase$(VBA.Trim$(argValue))
    If VBA.Len(outSourceKey) = 0 Then Exit Function
    private_TryResolvePageItemsSourceKey = True
End Function

Private Function private_TryBuildConfigTable(ByVal sourceItems As Collection, ByRef outTable As obj_ConfigTable) As Boolean
    Dim sourceConfigEntry As obj_ConfigEntry

    Set outTable = Nothing
    If sourceItems Is Nothing Then Exit Function

    ' Hotkeys использует тот же простой источник строк, что и Config:
    ' Collection(obj_ConfigEntry). Это намеренно, чтобы не плодить модель ради 2 полей.
    Set outTable = New obj_ConfigTable
    For Each sourceConfigEntry In sourceItems
        If sourceConfigEntry Is Nothing Then GoTo ContinueSourceItem
        If Not outTable.AddItem(sourceConfigEntry) Then Exit Function
ContinueSourceItem:
    Next sourceConfigEntry

    private_TryBuildConfigTable = True
End Function

Private Function private_TryMeasureNode( _
    ByVal controlNode As Object, _
    ByRef outSpanRows As Long, _
    ByRef outSpanColls As Long _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim itemsSourceRaw As String
    Dim resolvedItems As Collection
    Dim configTable As obj_ConfigTable

    ' Measure должен вернуть высоту до Render, чтобы layout engine правильно
    ' разместил следующие элементы. Высота = Apply-row + header + rows itemsSource.
    outSpanRows = 3
    outSpanColls = HOTKEY_COL_COUNT
    If controlNode Is Nothing Then Exit Function
    If m_Page Is Nothing Then Exit Function

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If pageBase.RuntimeSources Is Nothing Then Exit Function

    itemsSourceRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "itemsSource")))
    If VBA.Len(itemsSourceRaw) = 0 Then Exit Function

    If Not ex_RuntimeSourceResolver.fn_TryResolveItemsSource(pageBase.RuntimeSources, itemsSourceRaw, resolvedItems) Then Exit Function
    If Not private_TryBuildConfigTable(resolvedItems, configTable) Then Exit Function

    If Not configTable Is Nothing Then
        outSpanRows = 2 + configTable.Count
        If outSpanRows < 3 Then outSpanRows = 3
    End If
    private_TryMeasureNode = True
End Function

Private Function private_TryResolveRenderedTableObject(ByRef outTableObj As ListObject) As Boolean
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet

    Set outTableObj = Nothing
    If Not m_IsConfigured Then Exit Function
    If VBA.Len(VBA.Trim$(m_RuntimeTableName)) = 0 Then Exit Function

    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    Set ws = private_GetWorksheetByName(pageBase, m_ControlLayout.LayoutSheetName)
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set outTableObj = ws.ListObjects(m_RuntimeTableName)
    On Error GoTo 0
    private_TryResolveRenderedTableObject = Not outTableObj Is Nothing
End Function

Private Function private_TryDeleteIntersectingTables(ByVal ws As Worksheet, ByVal boundsRange As Range) As Boolean
    Dim idx As Long
    Dim tableObj As ListObject

    If ws Is Nothing Then Exit Function
    If boundsRange Is Nothing Then Exit Function

    On Error GoTo EH_DELETE
    For idx = ws.ListObjects.Count To 1 Step -1
        Set tableObj = ws.ListObjects(idx)
        If Not tableObj Is Nothing Then
            If Not Application.Intersect(tableObj.Range, boundsRange) Is Nothing Then
                tableObj.Delete
            End If
        End If
    Next idx

    private_TryDeleteIntersectingTables = True
    Exit Function

EH_DELETE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "Hotkeys: failed to delete intersecting tables for control '" & m_ControlName & "': " & Err.Description
#End If
End Function

Private Function private_BuildTableName(ByVal ws As Worksheet) As String
    Dim baseName As String

    baseName = VBA.Trim$(m_TableNameRaw)
    If VBA.Len(baseName) = 0 Then baseName = "hotkeys" & m_ControlName
    baseName = private_SanitizeTableName(baseName)
    If VBA.Len(baseName) = 0 Then baseName = "hotkeysTable"
    private_BuildTableName = private_BuildUniqueTableName(ws, baseName)
End Function

Private Function private_BuildUniqueTableName(ByVal ws As Worksheet, ByVal baseName As String) As String
    Dim candidate As String
    Dim suffixIndex As Long

    If ws Is Nothing Then Exit Function
    baseName = VBA.Left$(VBA.Trim$(baseName), 240)
    If VBA.Len(baseName) = 0 Then baseName = "hotkeysTable"

    candidate = baseName
    suffixIndex = 1
    Do While private_TableNameExists(ws, candidate)
        suffixIndex = suffixIndex + 1
        candidate = VBA.Left$(baseName, 240 - VBA.Len(VBA.CStr(suffixIndex))) & VBA.CStr(suffixIndex)
    Loop

    private_BuildUniqueTableName = candidate
End Function

Private Function private_TableNameExists(ByVal ws As Worksheet, ByVal tableName As String) As Boolean
    Dim tableObj As ListObject

    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Set tableObj = ws.ListObjects(tableName)
    private_TableNameExists = Not tableObj Is Nothing
    On Error GoTo 0
End Function

Private Function private_SanitizeTableName(ByVal valueText As String) As String
    Dim i As Long
    Dim ch As String
    Dim result As String

    valueText = VBA.Trim$(valueText)
    For i = 1 To VBA.Len(valueText)
        ch = VBA.Mid$(valueText, i, 1)
        If ch Like "[A-Za-z0-9_]" Then
            result = result & ch
        Else
            result = result & "_"
        End If
    Next i
    If VBA.Len(result) = 0 Then Exit Function
    If VBA.Mid$(result, 1, 1) Like "[0-9]" Then result = "t_" & result
    private_SanitizeTableName = result
End Function

Private Function private_GetUiShapeByName(ByVal ws As Worksheet, ByVal shapeName As String) As Shape
    If ws Is Nothing Then Exit Function
    shapeName = VBA.Trim$(shapeName)
    If VBA.Len(shapeName) = 0 Then Exit Function

    On Error Resume Next
    Set private_GetUiShapeByName = ws.Shapes(shapeName)
    On Error GoTo 0
End Function

Private Function private_GetWorksheetByName(ByVal page As obj_PageBase, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    If page Is Nothing Then Exit Function
    Set ws = page.Worksheet
    If ws Is Nothing Then Exit Function

    sheetName = VBA.LCase$(VBA.Trim$(sheetName))
    If VBA.Len(sheetName) > 0 Then
        If VBA.StrComp(VBA.LCase$(VBA.Trim$(ws.Name)), sheetName, VBA.vbTextCompare) <> 0 Then Exit Function
    End If

    Set private_GetWorksheetByName = ws
End Function
