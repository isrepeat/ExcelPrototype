VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_SelectControlVM"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False
Private m_IsDisposed As Boolean
Implements obj_IControl
Implements obj_ISerializable

Private Const DEFAULT_PLACEHOLDER As String = "Choose option"
Private Const DEFAULT_ITEM_HEIGHT As Double = 18#
Private Const DEFAULT_ITEM_MARGIN As Double = 2#
Private Const SHAPE_NAME_MAX_LEN As Long = 30
Private Const SHAPE_CONTROL_HASH_LEN As Long = 8
Private Const SHAPE_SUFFIX_HASH_LEN As Long = 8

Private m_ControlBase As obj_ControlBase
' Общие layout-параметры контрола (лист, границы в ячейках, style).
Private m_ControlLayout As obj_ControlLayout
Private m_ControlName As String
' Значения, считанные из XML (raw-конфиг).
Private m_ItemsSourceRaw As String
Private m_PlaceholderText As String
Private m_OnChangeRaw As String
Private m_OnChangeMacroRef As String
Private m_DropDownOpenedRaw As String
Private m_DropDownOpenedMacroRef As String
Private m_SelectedIdRaw As String
Private m_ItemStyleName As String
Private m_PanelStyleName As String
Private m_ItemHeight As Double
Private m_ItemMargin As Double

' Буферы данных после разрешения itemsSource.
' Эти коллекции нужны для первичного рендера.
Private m_Items As Collection
Private m_ItemCaptions As Collection
Private m_ItemIds As Collection
Private m_ItemActionMacros As Collection
Private m_ItemRawItems As Collection
Private m_SelectedIndex As Long
Private m_SelectStateKey As String

' UI-буферы уже отрисованного select.
' Используются при кликах через page-owned карты (без повторного configure/render).
Private m_UiHeaderShapeName As String
Private m_UiDropdownPanelShapeName As String
Private m_UiOptionShapeNames As Collection
Private m_UiOptionCaptions As Collection
Private m_UiOptionIds As Collection
Private m_UiOptionActionMacros As Collection
Private m_UiOptionRawItems As Collection
Private m_IsDropdownExpanded As Boolean
Private m_IsConfigured As Boolean
Private m_Page As obj_IPage
Private m_CallbackContext As Object
Private m_ShapeNameByXmlSuffix As Object
Private m_XmlSuffixByShapeName As Object
Private m_SemanticPartsSignature As String
Private m_StyledSelectionSignature As String
Private m_IsControlRenderPass As Boolean

Private Sub Class_Initialize()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Initialize"
#End If
End Sub

Private Sub Class_Terminate()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Class_Terminate"
#End If
    If m_IsDisposed Then Exit Sub
    On Error Resume Next
    obj_IControl_Dispose
    On Error GoTo 0
End Sub

' //
' // Interface
' //
Private Function obj_IControl_Initialize(ByVal page As obj_IPage) As Boolean
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Initialize"
#End If
    m_IsDisposed = False
    Set m_Page = page
    obj_IControl_Initialize = True
End Function

Private Sub obj_IControl_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:" & VBA.TypeName(Me) & ".Dispose"
#End If
    If m_IsDisposed Then Exit Sub
    m_IsDisposed = True
    On Error Resume Next
    Err.Clear
    Err.Clear
    Err.Clear
    Err.Clear
    Err.Clear
    Err.Clear
    Err.Clear
    Set m_ControlBase = Nothing
    Set m_ControlLayout = Nothing
    Set m_Items = Nothing
    Set m_ItemCaptions = Nothing
    Set m_ItemIds = Nothing
    Set m_ItemActionMacros = Nothing
    Set m_ItemRawItems = Nothing
    Set m_UiOptionShapeNames = Nothing
    Set m_UiOptionCaptions = Nothing
    Set m_UiOptionIds = Nothing
    Set m_UiOptionActionMacros = Nothing
    Set m_UiOptionRawItems = Nothing
    Set m_CallbackContext = Nothing
    Set m_ShapeNameByXmlSuffix = Nothing
    Set m_XmlSuffixByShapeName = Nothing
    Set m_Page = Nothing
    On Error GoTo 0
End Sub

Private Sub obj_IControl_Configure(ByVal controlNode As Object)
    Dim pageBase As obj_PageBase
    Dim dataContext As Object
    Dim selectedIdText As String

    ' Полный reset состояния: важно при повторной конфигурации того же VM.
    m_IsConfigured = False
    Set m_ControlLayout = Nothing
    Set m_Items = Nothing
    Set m_ItemCaptions = Nothing
    Set m_ItemIds = Nothing
    Set m_ItemActionMacros = Nothing
    Set m_ItemRawItems = Nothing
    m_UiHeaderShapeName = VBA.vbNullString
    m_UiDropdownPanelShapeName = VBA.vbNullString
    Set m_UiOptionShapeNames = Nothing
    Set m_UiOptionCaptions = Nothing
    Set m_UiOptionIds = Nothing
    Set m_UiOptionActionMacros = Nothing
    Set m_UiOptionRawItems = Nothing
    Set m_ControlBase = Nothing
    Set m_CallbackContext = Nothing
    Set m_ShapeNameByXmlSuffix = Nothing
    Set m_XmlSuffixByShapeName = Nothing
    m_IsDropdownExpanded = False
    m_SelectedIndex = 0
    m_SemanticPartsSignature = VBA.vbNullString
    m_StyledSelectionSignature = VBA.vbNullString
    m_IsControlRenderPass = False

    If m_Page Is Nothing Then Exit Sub
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Sub
    Set m_ControlBase = New obj_ControlBase
    If Not m_ControlBase.Initialize(m_Page) Then Exit Sub
    If Not m_ControlBase.Configure(pageBase, controlNode, "Select", "select", m_ControlName) Then Exit Sub

    m_ItemsSourceRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "itemsSource")))
    If VBA.Len(m_ItemsSourceRaw) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Select: itemsSource is not specified for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    m_SelectedIdRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "selectedId")))

    m_PlaceholderText = VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "placeholder"))
    If VBA.Len(VBA.Trim$(m_PlaceholderText)) = 0 Then m_PlaceholderText = DEFAULT_PLACEHOLDER

    m_ItemStyleName = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "itemStyle")))
    m_PanelStyleName = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "panelStyle")))

    If Not private_TryReadPositiveDoubleAttr(controlNode, "itemHeight", DEFAULT_ITEM_HEIGHT, m_ItemHeight) Then Exit Sub
    If Not private_TryReadNonNegativeDoubleAttr(controlNode, "itemMargin", DEFAULT_ITEM_MARGIN, m_ItemMargin) Then Exit Sub

    Set dataContext = m_ControlBase.DataContext
    If dataContext Is Nothing Then Set dataContext = m_Page
    Set m_CallbackContext = dataContext

    m_OnChangeRaw = VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "onChange"))
    m_OnChangeMacroRef = VBA.vbNullString
    If VBA.Len(VBA.Trim$(m_OnChangeRaw)) > 0 Then
        If Not private_TryResolveCallbackRef(m_OnChangeRaw, m_CallbackContext, m_OnChangeMacroRef) Then Exit Sub
    End If

    m_DropDownOpenedRaw = VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "dropDownOpened"))
    m_DropDownOpenedMacroRef = VBA.vbNullString
    If VBA.Len(VBA.Trim$(m_DropDownOpenedRaw)) > 0 Then
        If Not private_TryResolveCallbackRef(m_DropDownOpenedRaw, m_CallbackContext, m_DropDownOpenedMacroRef) Then Exit Sub
    End If

    ' 2) Читаем общий layout (лист + границы + style).
    Set m_ControlLayout = New obj_ControlLayout
    If Not m_ControlLayout.TryReadFromNode(controlNode, "Select", m_ControlName, "headerStyle") Then Exit Sub

    ' 3) Разрешаем itemsSource в runtime-коллекцию и готовим буферы.
    Set pageBase = m_ControlBase.PageBase
    If pageBase Is Nothing Then Exit Sub
    If Not ex_RuntimeSourceResolver.fn_TryResolveItemsSource(pageBase.RuntimeSources, m_ItemsSourceRaw, m_Items) Then Exit Sub
    If Not private_TryBuildItemBuffers() Then Exit Sub

    ' 4) Определяем начальный выбранный элемент. Пустой selectedId означает
    '    настоящее placeholder-состояние, а не неявный выбор первого item.
    m_SelectStateKey = VBA.LCase$(m_ControlLayout.LayoutSheetName & "|" & m_ControlName)
    If Not private_TryResolveSelectedIdText(selectedIdText) Then Exit Sub
    m_SelectedIndex = private_FindSelectedIndexById(selectedIdText)

    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render()
    Dim ws As Worksheet
    Dim headerRange As Range
    Dim headerShape As Shape
    Dim panelShape As Shape
    Dim itemShapes As Collection
    Dim itemCaptions As Collection
    Dim itemIds As Collection
    Dim itemActions As Collection
    Dim itemRawItems As Collection
    Dim callbackMacroRef As String
    Dim renderItemCount As Long
    Dim selectedIndexRendered As Long
    Dim headerLeft As Double
    Dim headerTop As Double
    Dim headerWidth As Double
    Dim headerHeight As Double
    Dim panelLeft As Double
    Dim panelTop As Double
    Dim panelWidth As Double
    Dim panelHeight As Double
    Dim pageBase As obj_PageBase
    Dim uiStateInitialized As Boolean

    If Not m_IsConfigured Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Select: control '" & m_ControlName & "' is not configured."
#End If
        Exit Sub
    End If

    Set pageBase = Nothing
    If Not m_ControlBase Is Nothing Then Set pageBase = m_ControlBase.PageBase
    If pageBase Is Nothing Then
        Set pageBase = m_Page.GetPageBase()
    End If
    If pageBase Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Select: page is not specified for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    Set ws = pageBase.Worksheet
    If ws Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Select: page worksheet is not specified for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    If m_ItemCaptions Is Nothing Or m_ItemIds Is Nothing Or m_ItemActionMacros Is Nothing Or m_ItemRawItems Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Select: item metadata is not configured for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    ' Shape.OnAction вызывает только модульный макрос,
    ' поэтому указываем стабильный bridge rt_Bridge.fn_OnShapeClick.
    callbackMacroRef = private_GetRuntimeCallbackMacroRef()
    If VBA.Len(callbackMacroRef) = 0 Then Exit Sub

    If Not private_TryBuildHeaderRange(ws, headerRange) Then Exit Sub
    renderItemCount = private_GetRenderItemCount()

    headerLeft = headerRange.Left
    headerTop = headerRange.Top
    headerWidth = headerRange.Width
    headerHeight = headerRange.Height

    panelLeft = headerLeft
    panelTop = headerTop + headerHeight
    panelWidth = headerWidth
    panelHeight = private_CalcPanelHeight(renderItemCount)

    ' В lazy-режиме item-shape создаются только при открытии dropdown.
    ' На page-render гарантированно очищаем item-shape, чтобы закрытый select
    ' не оставлял "висячие" опции от предыдущего открытия.
    private_DeleteStaleItemShapes ws, 0

    Set headerShape = private_CreateShapeByRange(ws, headerRange, "header", callbackMacroRef)
    If headerShape Is Nothing Then Exit Sub
    ' Встроенная палитра нужна только Select без XML-стиля. Иначе локальный
    ' rerender не должен затирать уже применённый controlStyle.
    If VBA.Len(VBA.Trim$(m_ControlLayout.StyleName)) = 0 Then
        private_ApplyHeaderVisualDefaults headerShape
    End If

    If renderItemCount > 0 Then
        Set panelShape = private_CreateShapeByBounds(ws, panelLeft, panelTop, panelWidth, panelHeight, "panel", VBA.vbNullString)
        If panelShape Is Nothing Then Exit Sub
        ' Новый page-render всегда начинает с закрытого dropdown. Скрываем
        ' переиспользованный panel до любых последующих configure/style шагов.
        panelShape.Visible = msoFalse
        If VBA.Len(VBA.Trim$(m_PanelStyleName)) = 0 Then
            private_ApplyPanelVisualDefaults panelShape
        End If
    Else
        Set panelShape = private_CreateShapeByBounds(ws, headerLeft, headerTop, headerWidth, headerHeight, "panel", VBA.vbNullString)
        If panelShape Is Nothing Then Exit Sub
        panelShape.Visible = msoFalse
    End If

    ' Ленивая стратегия:
    ' на page-render создаем только header/panel, а item-shape и item-routes
    ' материализуем в момент раскрытия dropdown (HandleHeaderClick).
    Set itemShapes = New Collection
    Set itemCaptions = m_ItemCaptions
    Set itemIds = m_ItemIds
    Set itemActions = m_ItemActionMacros
    Set itemRawItems = m_ItemRawItems

    selectedIndexRendered = m_SelectedIndex
    If selectedIndexRendered <= 0 Or selectedIndexRendered > renderItemCount Then selectedIndexRendered = 0

    ' Синхронизируем runtime-буферы VM с только что созданными shape.
    m_IsControlRenderPass = True
    uiStateInitialized = private_InitializeUiState( _
        headerShapeName:=headerShape.Name, _
        panelShapeName:=panelShape.Name, _
        itemShapeNames:=itemShapes, _
        itemCaptions:=itemCaptions, _
        itemIds:=itemIds, _
        itemActionMacros:=itemActions, _
        itemRawItems:=itemRawItems, _
        selectedIndex:=selectedIndexRendered)
    m_IsControlRenderPass = False
    If Not uiStateInitialized Then Exit Sub

    If Not private_TryBindUiRoutes(ws, headerShape.Name, itemShapes) Then Exit Sub
End Sub

Private Function obj_IControl_Measure( _
    ByVal controlNode As Object, _
    ByRef outSpanRows As Long, _
    ByRef outSpanColls As Long, _
    Optional ByVal dataContext As Object _
) As Boolean
    outSpanRows = 1
    outSpanColls = 1
    obj_IControl_Measure = True
End Function

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(attrName))
        Case "itemssource", "placeholder", "onchange", "dropdownopened", "selectedid", _
             "headerstyle", _
             "itemstyle", "panelstyle", "itemheight", "itemmargin"
            obj_IControl_SupportsAttribute = True
    End Select
End Function

Private Function obj_IControl_IsConfigured() As Boolean
    obj_IControl_IsConfigured = m_IsConfigured
End Function

Private Function obj_ISerializable_GetSerializableTypeRoot() As String
    obj_ISerializable_GetSerializableTypeRoot = "select"
End Function

Private Function obj_ISerializable_TrySerializeSnapshot(ByRef outSnapshotXml As String) As Boolean
    obj_ISerializable_TrySerializeSnapshot = Me.TrySerializeSnapshot(outSnapshotXml)
End Function

Private Function obj_ISerializable_TryDeserializeSnapshot(ByVal snapshotXml As String) As Boolean
    obj_ISerializable_TryDeserializeSnapshot = Me.TryDeserializeSnapshot(snapshotXml)
End Function

Private Function obj_ISerializable_TryRestoreState() As Boolean
    obj_ISerializable_TryRestoreState = True
End Function

' //
' // API
' //
Public Function GetSelectedId() As String
    If Not m_UiOptionIds Is Nothing Then
        If m_SelectedIndex <= 0 Or m_SelectedIndex > m_UiOptionIds.Count Then Exit Function
        GetSelectedId = VBA.CStr(m_UiOptionIds(m_SelectedIndex))
        Exit Function
    End If

    If m_ItemIds Is Nothing Then Exit Function
    If m_SelectedIndex <= 0 Or m_SelectedIndex > m_ItemIds.Count Then Exit Function
    GetSelectedId = VBA.CStr(m_ItemIds(m_SelectedIndex))
End Function

Public Function GetSelectedCaption() As String
    If Not m_UiOptionCaptions Is Nothing Then
        If m_SelectedIndex <= 0 Or m_SelectedIndex > m_UiOptionCaptions.Count Then Exit Function
        GetSelectedCaption = VBA.CStr(m_UiOptionCaptions(m_SelectedIndex))
        Exit Function
    End If

    If m_ItemCaptions Is Nothing Then Exit Function
    If m_SelectedIndex <= 0 Or m_SelectedIndex > m_ItemCaptions.Count Then Exit Function
    GetSelectedCaption = VBA.CStr(m_ItemCaptions(m_SelectedIndex))
End Function

Public Function GetControlKey() As String
    GetControlKey = VBA.CStr(m_ControlName)
End Function

Public Function HasSelectedOption() As Boolean
    If m_UiOptionCaptions Is Nothing Then Exit Function
    HasSelectedOption = (m_SelectedIndex > 0 And m_SelectedIndex <= m_UiOptionCaptions.Count)
End Function

Public Function GetSelectedOptionIndex() As Long
    GetSelectedOptionIndex = m_SelectedIndex
End Function

Public Function GetSelectedOptionCaption() As String
    GetSelectedOptionCaption = Me.GetSelectedCaption()
End Function

Public Function GetSelectedOptionId() As String
    GetSelectedOptionId = Me.GetSelectedId()
End Function

' Программно синхронизирует выбранный item без запуска пользовательского
' onChange. Используется страницами, когда связанное состояние меняется
' другим контролом, например основной кнопкой рядом с dropdown.
Public Function TrySelectId(ByVal selectedId As String) As Boolean
    Dim selectedIndex As Long

    selectedId = VBA.Trim$(selectedId)
    If VBA.Len(selectedId) = 0 Then Exit Function
    selectedIndex = private_FindSelectedIndexById(selectedId)
    If selectedIndex <= 0 Then Exit Function

    m_SelectedIndex = selectedIndex
    m_IsDropdownExpanded = False
    If Not private_TryPersistSelectedId(selectedId) Then Exit Function
    If Not private_ApplyUiStateToShapes() Then Exit Function
    TrySelectId = True
End Function

' Возвращает Select в placeholder-состояние. По умолчанию программный сброс
' не имитирует пользовательский выбор; при необходимости потребитель может
' явно запросить onChange через notifyChange=True.
Public Function ResetSelection(Optional ByVal notifyChange As Boolean = False) As Boolean
    m_SelectedIndex = 0
    m_IsDropdownExpanded = False
    If Not private_TryPersistSelectedId(VBA.vbNullString) Then Exit Function
    If Not private_ApplyUiStateToShapes() Then Exit Function

    If notifyChange And VBA.Len(VBA.Trim$(m_OnChangeMacroRef)) > 0 Then
        If Not private_RunOptionMacro(m_OnChangeMacroRef) Then Exit Function
    End If
    ResetSelection = True
End Function

Public Function HandleHeaderClick() As Boolean
    ' При раскрытии dropdown можем выполнить callback DropDownOpened
    ' (например, чтобы обновить runtime-список перед показом опций).
    If Not m_IsDropdownExpanded Then
        If VBA.Len(VBA.Trim$(m_DropDownOpenedMacroRef)) > 0 Then
            If Not private_RunOptionMacro(m_DropDownOpenedMacroRef) Then Exit Function
            ' Важный нюанс:
            ' callback обычно обновляет только RuntimeSources (данные),
            ' но текущие shape/буферы select пока еще содержат старые item-ы.
            ' Поэтому сразу делаем локальный refresh + rerender этого контрола,
            ' чтобы в уже открывающемся dropdown показать актуальный список.
            If Not private_TryRefreshItemsAndRerenderAfterDropDownOpenedCallback() Then Exit Function
        End If

        ' Перед раскрытием гарантируем, что item-shape реально существуют
        ' и имеют routes для HandleOptionClick.
        If Not private_TryEnsureDropdownItemsReady() Then Exit Function
    End If

    ' Клик по header переключает open/close dropdown panel.
    m_IsDropdownExpanded = (Not m_IsDropdownExpanded)
    HandleHeaderClick = private_ApplyUiStateToShapes()
End Function

Public Function HandleOptionClick(ByVal itemIndex As Long) As Boolean
    Dim selectedId As String
    Dim itemMacro As String

    If m_UiOptionCaptions Is Nothing Then Exit Function
    If itemIndex <= 0 Or itemIndex > m_UiOptionCaptions.Count Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Select: selected item index is out of range for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    m_SelectedIndex = itemIndex
    m_IsDropdownExpanded = False

    If Not private_ApplyUiStateToShapes() Then Exit Function

    ' При выборе item:
    ' 1) фиксируем selectedId в CustomXMLPart
    ' 2) запускаем item macro и onChange.
    selectedId = Me.GetSelectedId()
    If Not private_TryPersistSelectedId(selectedId) Then Exit Function

    itemMacro = private_GetOptionCollectionText(m_UiOptionActionMacros, itemIndex)
    If VBA.Len(itemMacro) > 0 Then
        If Not private_RunOptionMacro(itemMacro) Then Exit Function
    End If

    If VBA.Len(m_OnChangeMacroRef) > 0 Then
        If Not private_RunOptionMacro(m_OnChangeMacroRef) Then Exit Function
    End If

    HandleOptionClick = True
End Function

Public Function CollapseDropdown() As Boolean
    ' Явное закрытие dropdown (используется dispatcher-ом при клике по другим shape).
    If Not m_IsDropdownExpanded Then
        CollapseDropdown = True
        Exit Function
    End If

    m_IsDropdownExpanded = False
    CollapseDropdown = private_ApplyUiStateToShapes()
End Function

Public Function HandleGlobalClick(ByVal clickedControlKey As String) As Boolean
    clickedControlKey = VBA.LCase$(VBA.Trim$(clickedControlKey))

    If clickedControlKey = VBA.LCase$(VBA.Trim$(m_SelectStateKey)) Then
        HandleGlobalClick = True
        Exit Function
    End If

    HandleGlobalClick = Me.CollapseDropdown()
End Function

Public Function TrySerializeSnapshot(ByRef outSnapshotXml As String) As Boolean
    Dim i As Long
    Dim selectedIndexText As String
    Dim optionTags As Collection
    Dim optionStates As Collection

    outSnapshotXml = VBA.vbNullString

    If m_ControlLayout Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(m_ControlName)) = 0 Then Exit Function
    If VBA.Len(VBA.Trim$(m_UiHeaderShapeName)) = 0 Then Exit Function
    If VBA.Len(VBA.Trim$(m_UiDropdownPanelShapeName)) = 0 Then Exit Function
    If m_UiOptionShapeNames Is Nothing Then Exit Function
    If m_UiOptionCaptions Is Nothing Then Exit Function
    If m_UiOptionIds Is Nothing Then Exit Function
    If m_UiOptionActionMacros Is Nothing Then Exit Function
    If m_UiOptionRawItems Is Nothing Then Exit Function
    If m_UiOptionCaptions.Count <> m_UiOptionIds.Count Then Exit Function
    If m_UiOptionCaptions.Count <> m_UiOptionActionMacros.Count Then Exit Function
    If m_UiOptionCaptions.Count <> m_UiOptionRawItems.Count Then Exit Function

    selectedIndexText = VBA.CStr(m_SelectedIndex)
    outSnapshotXml = "<select version=""4"""
    outSnapshotXml = outSnapshotXml & " controlName=""" & ex_Helpers.fn_EscapeXmlAttr(m_ControlName) & """"
    outSnapshotXml = outSnapshotXml & " itemsSource=""" & ex_Helpers.fn_EscapeXmlAttr(m_ItemsSourceRaw) & """"
    outSnapshotXml = outSnapshotXml & " selectedIdRaw=""" & ex_Helpers.fn_EscapeXmlAttr(m_SelectedIdRaw) & """"
    outSnapshotXml = outSnapshotXml & " sheet=""" & ex_Helpers.fn_EscapeXmlAttr(m_ControlLayout.LayoutSheetName) & """"
    outSnapshotXml = outSnapshotXml & " rowStart=""" & VBA.CStr(m_ControlLayout.RowStart) & """"
    outSnapshotXml = outSnapshotXml & " colStart=""" & VBA.CStr(m_ControlLayout.ColStart) & """"
    outSnapshotXml = outSnapshotXml & " rowEnd=""" & VBA.CStr(m_ControlLayout.RowEnd) & """"
    outSnapshotXml = outSnapshotXml & " colEnd=""" & VBA.CStr(m_ControlLayout.ColEnd) & """"
    outSnapshotXml = outSnapshotXml & " selectKey=""" & ex_Helpers.fn_EscapeXmlAttr(m_SelectStateKey) & """"
    outSnapshotXml = outSnapshotXml & " onChangeRaw=""" & ex_Helpers.fn_EscapeXmlAttr(m_OnChangeRaw) & """"
    outSnapshotXml = outSnapshotXml & " onChange=""" & ex_Helpers.fn_EscapeXmlAttr(m_OnChangeMacroRef) & """"
    outSnapshotXml = outSnapshotXml & " dropDownOpenedRaw=""" & ex_Helpers.fn_EscapeXmlAttr(m_DropDownOpenedRaw) & """"
    outSnapshotXml = outSnapshotXml & " dropDownOpened=""" & ex_Helpers.fn_EscapeXmlAttr(m_DropDownOpenedMacroRef) & """"
    outSnapshotXml = outSnapshotXml & " selectedIndex=""" & ex_Helpers.fn_EscapeXmlAttr(selectedIndexText) & """"
    outSnapshotXml = outSnapshotXml & " isConfigured=""" & VBA.IIf(m_IsConfigured, "true", "false") & """"
    outSnapshotXml = outSnapshotXml & ">"
    outSnapshotXml = outSnapshotXml & "<header shape=""" & ex_Helpers.fn_EscapeXmlAttr(m_UiHeaderShapeName) & """ />"
    outSnapshotXml = outSnapshotXml & "<panel shape=""" & ex_Helpers.fn_EscapeXmlAttr(m_UiDropdownPanelShapeName) & """ />"

    For i = 1 To m_UiOptionIds.Count
        If Not private_TryGetOptionSemantics(i, optionTags, optionStates) Then Exit Function
        outSnapshotXml = outSnapshotXml & _
            "<item" & _
            " caption=""" & ex_Helpers.fn_EscapeXmlAttr(VBA.CStr(m_UiOptionCaptions(i))) & """" & _
            " id=""" & ex_Helpers.fn_EscapeXmlAttr(VBA.CStr(m_UiOptionIds(i))) & """" & _
            " action=""" & ex_Helpers.fn_EscapeXmlAttr(VBA.CStr(m_UiOptionActionMacros(i))) & """" & _
            " tags=""" & ex_Helpers.fn_EscapeXmlAttr(private_CollectionSignature(optionTags)) & """" & _
            " states=""" & ex_Helpers.fn_EscapeXmlAttr(private_CollectionSignature(optionStates)) & """" & _
                " rawValue=""" & ex_Helpers.fn_EscapeXmlAttr(ex_Helpers.fn_GetSnapshotRawValueText(m_UiOptionRawItems, i, VBA.CStr(m_UiOptionIds(i)))) & """" & _
            " />"
    Next i

    outSnapshotXml = outSnapshotXml & "</select>"
    TrySerializeSnapshot = True
End Function

Public Function TryDeserializeSnapshot(ByVal snapshotXml As String) As Boolean
    Dim dom As Object
    Dim root As Object
    Dim ws As Worksheet
    Dim itemNodes As Object
    Dim itemNode As Object
    Dim headerNode As Object
    Dim panelNode As Object
    Dim itemShapeNames As Collection
    Dim itemCaptions As Collection
    Dim itemIds As Collection
    Dim itemActionMacros As Collection
    Dim itemRawItems As Collection
    Dim selectedIndex As Long
    Dim isDropdownExpanded As Boolean
    Dim headerShapeName As String
    Dim panelShapeName As String
    Dim layoutSheetName As String
    Dim layoutRowStart As Long
    Dim layoutColStart As Long
    Dim layoutRowEnd As Long
    Dim layoutColEnd As Long
    Dim layoutStyle As String
    Dim isConfiguredAttr As String
    Dim i As Long
    Dim optionObj As obj_SelectOption
    Dim pageBase As obj_PageBase
    Dim currentControlNode As Object
    Dim escapedControlName As String
    Dim restoredPanelShape As Shape

    snapshotXml = VBA.Trim$(snapshotXml)
    If VBA.Len(snapshotXml) = 0 Then Exit Function

    If Not ex_Core.fn_CustomXmlPartStore_TryLoadDomFromXml(snapshotXml, dom) Then Exit Function
    Set root = dom.DocumentElement
    If root Is Nothing Then Exit Function
    If VBA.LCase$(VBA.CStr(root.baseName)) <> "select" Then Exit Function

    m_ControlName = VBA.Trim$(VBA.CStr(root.getAttribute("controlName")))
    m_ItemsSourceRaw = VBA.Trim$(VBA.CStr(root.getAttribute("itemsSource")))
    m_SelectedIdRaw = VBA.Trim$(VBA.CStr(root.getAttribute("selectedIdRaw")))
    m_SelectStateKey = VBA.LCase$(VBA.Trim$(VBA.CStr(root.getAttribute("selectKey"))))
    m_OnChangeRaw = VBA.CStr(root.getAttribute("onChangeRaw"))
    m_OnChangeMacroRef = VBA.Trim$(VBA.CStr(root.getAttribute("onChange")))
    m_DropDownOpenedRaw = VBA.CStr(root.getAttribute("dropDownOpenedRaw"))
    m_DropDownOpenedMacroRef = VBA.Trim$(VBA.CStr(root.getAttribute("dropDownOpened")))
    ' Open/closed — transient UI-state: после reload Select всегда закрыт.
    isDropdownExpanded = False
    isConfiguredAttr = VBA.LCase$(VBA.Trim$(VBA.CStr(root.getAttribute("isConfigured"))))
    layoutSheetName = VBA.Trim$(VBA.CStr(root.getAttribute("sheet")))
    layoutRowStart = ex_Helpers.fn_ReadSnapshotLongAttr(root, "rowStart", 1)
    layoutColStart = ex_Helpers.fn_ReadSnapshotLongAttr(root, "colStart", 1)
    layoutRowEnd = ex_Helpers.fn_ReadSnapshotLongAttr(root, "rowEnd", layoutRowStart)
    layoutColEnd = ex_Helpers.fn_ReadSnapshotLongAttr(root, "colEnd", layoutColStart)
    If VBA.Len(m_ControlName) = 0 Then Exit Function
    If VBA.Len(m_SelectStateKey) = 0 Then
        m_SelectStateKey = VBA.LCase$(VBA.Trim$(layoutSheetName) & "|" & m_ControlName)
    End If
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If pageBase.XmlDom Is Nothing Then Exit Function
    escapedControlName = ex_XmlCore.fn_XPathLiteral(m_ControlName)
    Set currentControlNode = pageBase.XmlDom.selectSingleNode( _
        "/p:page//p:control[@name=" & escapedControlName & "] | " & _
        "/p:uiDefinition/p:layout//p:control[@name=" & escapedControlName & "]")
    If currentControlNode Is Nothing Then Exit Function

    ' Snapshot хранит runtime-состояние, но не является источником UI-конфига.
    ' После reload стили и geometry читаются из актуального XML так же, как при
    ' полном render, поэтому восстановленный Select не получает старую палитру.
    layoutStyle = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText( _
        currentControlNode, "headerStyle")))
    m_ItemStyleName = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText( _
        currentControlNode, "itemStyle")))
    m_PanelStyleName = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText( _
        currentControlNode, "panelStyle")))
    m_PlaceholderText = VBA.CStr(ex_XmlCore.fn_NodeAttrText( _
        currentControlNode, "placeholder"))
    If VBA.Len(VBA.Trim$(m_PlaceholderText)) = 0 Then m_PlaceholderText = DEFAULT_PLACEHOLDER
    If Not private_TryReadPositiveDoubleAttr( _
        currentControlNode, "itemHeight", DEFAULT_ITEM_HEIGHT, m_ItemHeight) Then Exit Function
    If Not private_TryReadNonNegativeDoubleAttr( _
        currentControlNode, "itemMargin", DEFAULT_ITEM_MARGIN, m_ItemMargin) Then Exit Function

    Set m_ControlLayout = New obj_ControlLayout
    If Not m_ControlLayout.TryReadFromRuntimeValues( _
        "Select", _
        m_ControlName, _
        layoutSheetName, _
        layoutRowStart, _
        layoutColStart, _
        layoutRowEnd, _
        layoutColEnd, _
        layoutStyle) Then Exit Function

    Set ws = ex_HelpersSheet.fn_GetRuntimeWorksheetByName(m_ControlLayout.LayoutSheetName)
    If ws Is Nothing Then Exit Function
    If Not private_TryRestoreCallbackContextFromPage() Then Exit Function

    ' Lazy item Shape не восстанавливаются из snapshot. Удаляем старый overlay,
    ' чтобы следующее раскрытие построило его из актуального XML/style pipeline.
    private_DeleteStaleItemShapes ws, 0

    Set headerNode = root.selectSingleNode("*[local-name()='header']")
    If headerNode Is Nothing Then Exit Function
    headerShapeName = VBA.Trim$(VBA.CStr(headerNode.getAttribute("shape")))
    If VBA.Len(headerShapeName) = 0 Then Exit Function

    Set panelNode = root.selectSingleNode("*[local-name()='panel']")
    If panelNode Is Nothing Then Exit Function
    panelShapeName = VBA.Trim$(VBA.CStr(panelNode.getAttribute("shape")))
    If VBA.Len(panelShapeName) = 0 Then Exit Function
    Set restoredPanelShape = private_GetUiShapeByName(ws, panelShapeName)
    If Not restoredPanelShape Is Nothing Then restoredPanelShape.Visible = msoFalse

    Set itemShapeNames = New Collection
    Set itemCaptions = New Collection
    Set itemIds = New Collection
    Set itemActionMacros = New Collection
    Set itemRawItems = New Collection

    Set itemNodes = root.selectNodes("*[local-name()='item']")
    If Not itemNodes Is Nothing Then
        For Each itemNode In itemNodes
            itemCaptions.Add VBA.CStr(itemNode.getAttribute("caption"))
            itemIds.Add VBA.CStr(itemNode.getAttribute("id"))
            itemActionMacros.Add VBA.CStr(itemNode.getAttribute("action"))

            Set optionObj = New obj_SelectOption
            optionObj.Id = VBA.CStr(itemNode.getAttribute("id"))
            optionObj.Caption = VBA.CStr(itemNode.getAttribute("caption"))
            optionObj.OnSelect = VBA.CStr(itemNode.getAttribute("action"))
            If Not private_TryRestoreOptionSemantics( _
                optionObj, _
                VBA.CStr(itemNode.getAttribute("tags")), _
                VBA.CStr(itemNode.getAttribute("states"))) Then Exit Function
            itemRawItems.Add optionObj
        Next itemNode
    End If

    ' Lazy-render snapshot может не содержать item-узлы (shape еще не были материализованы).
    ' В этом случае восстанавливаем item-данные из текущего runtime itemsSource.
    If itemCaptions.Count = 0 Then
        Set pageBase = m_Page.GetPageBase()
        If Not pageBase Is Nothing Then
            If VBA.Len(VBA.Trim$(m_ItemsSourceRaw)) > 0 Then
                If ex_RuntimeSourceResolver.fn_TryResolveItemsSource(pageBase.RuntimeSources, m_ItemsSourceRaw, m_Items) Then
                    If private_TryBuildItemBuffers() Then
                        Set itemCaptions = m_ItemCaptions
                        Set itemIds = m_ItemIds
                        Set itemActionMacros = m_ItemActionMacros
                        Set itemRawItems = m_ItemRawItems
                        Set itemShapeNames = New Collection
                    End If
                End If
            End If
        End If
    End If

    If VBA.IsNumeric(VBA.CStr(root.getAttribute("selectedIndex"))) Then
        selectedIndex = VBA.CLng(root.getAttribute("selectedIndex"))
    Else
        selectedIndex = 0
    End If

    If Not private_TrySyncRestoredShapeStyleMetadata( _
        ws, headerShapeName, panelShapeName, itemShapeNames) Then Exit Function
    ' Restore меняет runtime buffers и style metadata поверх уже созданного VM;
    ' прежняя semantic signature больше не доказывает актуальность Shape.
    m_SemanticPartsSignature = VBA.vbNullString
    m_StyledSelectionSignature = VBA.vbNullString

    If Not private_InitializeUiState( _
        headerShapeName:=headerShapeName, _
        panelShapeName:=panelShapeName, _
        itemShapeNames:=itemShapeNames, _
        itemCaptions:=itemCaptions, _
        itemIds:=itemIds, _
        itemActionMacros:=itemActionMacros, _
        itemRawItems:=itemRawItems, _
        selectedIndex:=selectedIndex) Then Exit Function

    If Not private_SyncStaticBuffersFromUiState() Then Exit Function

    m_IsDropdownExpanded = isDropdownExpanded
    If Not private_ApplyUiStateToShapes() Then Exit Function

    If Not private_TryBindUiRoutes(ws, headerShapeName, itemShapeNames) Then Exit Function

    If isConfiguredAttr = "false" Or isConfiguredAttr = "0" Then
        m_IsConfigured = False
    Else
        m_IsConfigured = True
    End If
    TryDeserializeSnapshot = True
End Function

Private Function private_TrySyncRestoredShapeStyleMetadata( _
    ByVal ws As Worksheet, _
    ByVal headerShapeName As String, _
    ByVal panelShapeName As String, _
    ByVal itemShapeNames As Collection _
) As Boolean
    Dim shp As Shape
    Dim i As Long

    If ws Is Nothing Or itemShapeNames Is Nothing Then Exit Function
    Set shp = private_GetUiShapeByName(ws, headerShapeName)
    If shp Is Nothing Then Exit Function
    If Not private_TrySetShapeStyleMetadata(shp, m_ControlLayout.StyleName) Then Exit Function

    Set shp = private_GetUiShapeByName(ws, panelShapeName)
    If Not shp Is Nothing Then
        If Not private_TrySetShapeStyleMetadata(shp, m_PanelStyleName) Then Exit Function
    End If
    For i = 1 To itemShapeNames.Count
        Set shp = private_GetUiShapeByName(ws, VBA.CStr(itemShapeNames(i)))
        If shp Is Nothing Then GoTo ContinueItem
        If Not private_TrySetShapeStyleMetadata(shp, m_ItemStyleName) Then Exit Function
ContinueItem:
    Next i
    private_TrySyncRestoredShapeStyleMetadata = True
End Function

Private Function private_TrySetShapeStyleMetadata( _
    ByVal shp As Shape, _
    ByVal styleName As String _
) As Boolean
    Dim metaMap As Object

    If shp Is Nothing Then Exit Function
    Set metaMap = VBA.CreateObject("Scripting.Dictionary")
    metaMap.CompareMode = 1
    metaMap("pn.style") = VBA.Trim$(styleName)
    metaMap("pn.appliedStyleSignature") = VBA.vbNullString
    metaMap("pn.appliedPartStyleSignature") = VBA.vbNullString
    private_TrySetShapeStyleMetadata = _
        ex_ShapeMetaRuntime.fn_TrySetShapeMetaValues(shp, metaMap)
End Function

' //
' // Internal
' //
Private Function private_TryBindUiRoutes( _
    ByVal ws As Worksheet, _
    ByVal headerShapeName As String, _
    ByVal itemShapeNames As Collection _
) As Boolean
    Dim callbackMacroRef As String
    Dim selectId As String
    Dim pageBase As obj_PageBase
    Dim headerShape As Shape
    Dim itemShape As Shape
    Dim boundItemShapeNames As Collection
    Dim boundItemIndexes As Collection
    Dim boundCount As Long
    Dim i As Long

    If ws Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "select:bind-routes worksheet-missing control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "'"
#End If
        Exit Function
    End If
    If itemShapeNames Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "select:bind-routes item-shapes-missing control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "'"
#End If
        Exit Function
    End If

    callbackMacroRef = private_GetRuntimeCallbackMacroRef()
    If VBA.Len(callbackMacroRef) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "select:bind-routes callback-macro-empty control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "'"
#End If
        Exit Function
    End If

    Set headerShape = private_GetUiShapeByName(ws, headerShapeName)
    If headerShape Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "select:bind-routes header-shape-missing control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' shape='" & VBA.Replace$(VBA.Trim$(headerShapeName), "'", "''") & "'"
#End If
        Exit Function
    End If
    If Not private_TrySetShapeOnAction(headerShape, callbackMacroRef) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "select:bind-routes header-onaction-failed control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' shape='" & VBA.Replace$(VBA.Trim$(headerShape.Name), "'", "''") & "'"
#End If
        Exit Function
    End If

    Set boundItemShapeNames = New Collection
    Set boundItemIndexes = New Collection

    For i = 1 To itemShapeNames.Count
        Set itemShape = private_GetUiShapeByName(ws, VBA.CStr(itemShapeNames(i)))
        If itemShape Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogInfo "select:bind-routes item-shape-missing-skip control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' index=" & VBA.CStr(i)
#End If
            GoTo ContinueBindItem
        End If
        If Not private_TrySetShapeOnAction(itemShape, callbackMacroRef) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogInfo "select:bind-routes item-onaction-failed-skip control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' shape='" & VBA.Replace$(VBA.Trim$(itemShape.Name), "'", "''") & "'"
#End If
            GoTo ContinueBindItem
        End If
        boundItemShapeNames.Add itemShape.Name
        boundItemIndexes.Add VBA.CLng(i)
ContinueBindItem:
    Next i

    selectId = VBA.LCase$(VBA.Trim$(m_SelectStateKey))
    If VBA.Len(selectId) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "select:bind-routes select-id-empty control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "'"
#End If
        Exit Function
    End If

    Set pageBase = m_Page.GetPageBase()
    If Not pageBase.RegisterControl(selectId, Me) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "select:bind-routes register-control-failed control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' key='" & VBA.Replace$(selectId, "'", "''") & "'"
#End If
        Exit Function
    End If
    If Not pageBase.RegisterShapeRoute(headerShape.Name, selectId, "HandleHeaderClick", False) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "select:bind-routes register-header-route-failed control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' shape='" & VBA.Replace$(VBA.Trim$(headerShape.Name), "'", "''") & "'"
#End If
        Exit Function
    End If

    For i = 1 To boundItemShapeNames.Count
        If Not pageBase.RegisterShapeRoute(VBA.CStr(boundItemShapeNames(i)), selectId, "HandleOptionClick", True, VBA.CLng(boundItemIndexes(i))) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "select:bind-routes register-item-route-failed control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' index=" & VBA.CStr(i)
#End If
            Exit Function
        End If
    Next i

    boundCount = boundItemShapeNames.Count
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "select:bind-routes ok control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' items=" & VBA.CStr(boundCount) & " macro='" & VBA.Replace$(callbackMacroRef, "'", "''") & "'"
#End If
    private_TryBindUiRoutes = True
End Function

Private Function private_TrySetShapeOnAction(ByVal shp As Shape, ByVal callbackMacroRef As String) As Boolean
    If shp Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(callbackMacroRef)) = 0 Then Exit Function

    On Error Resume Next
    shp.OnAction = callbackMacroRef
    If Err.Number <> 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "select:set-onaction-failed control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' shape='" & VBA.Replace$(VBA.Trim$(shp.Name), "'", "''") & "' err='" & VBA.Replace$(Err.Description, "'", "''") & "'"
#End If
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    private_TrySetShapeOnAction = True
End Function

Private Function private_SyncStaticBuffersFromUiState() As Boolean
    Dim i As Long

    If m_UiOptionCaptions Is Nothing Then Exit Function
    If m_UiOptionIds Is Nothing Then Exit Function
    If m_UiOptionActionMacros Is Nothing Then Exit Function
    If m_UiOptionRawItems Is Nothing Then Exit Function

    Set m_Items = New Collection
    Set m_ItemCaptions = New Collection
    Set m_ItemIds = New Collection
    Set m_ItemActionMacros = New Collection
    Set m_ItemRawItems = New Collection

    For i = 1 To m_UiOptionCaptions.Count
        m_ItemCaptions.Add VBA.CStr(m_UiOptionCaptions(i))
        m_ItemIds.Add VBA.CStr(m_UiOptionIds(i))
        m_ItemActionMacros.Add VBA.CStr(m_UiOptionActionMacros(i))
        m_ItemRawItems.Add m_UiOptionRawItems(i)
        m_Items.Add m_UiOptionRawItems(i)
    Next i

    private_SyncStaticBuffersFromUiState = True
End Function

Private Function private_InitializeUiState( _
    ByVal headerShapeName As String, _
    ByVal panelShapeName As String, _
    ByVal itemShapeNames As Collection, _
    ByVal itemCaptions As Collection, _
    ByVal itemIds As Collection, _
    ByVal itemActionMacros As Collection, _
    ByVal itemRawItems As Collection, _
    ByVal selectedIndex As Long _
) As Boolean
    Dim selectedId As String

    ' UI-state хранит shape-имена и коллекции, с которыми работает click handler.
    ' Это позволяет обрабатывать клики без повторного рендера.
    If itemShapeNames Is Nothing Or itemCaptions Is Nothing Or itemIds Is Nothing Or itemActionMacros Is Nothing Or itemRawItems Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Select: runtime item metadata collection is not specified for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    m_UiHeaderShapeName = VBA.CStr(headerShapeName)
    m_UiDropdownPanelShapeName = VBA.CStr(panelShapeName)
    Set m_UiOptionShapeNames = itemShapeNames
    Set m_UiOptionCaptions = itemCaptions
    Set m_UiOptionIds = itemIds
    Set m_UiOptionActionMacros = itemActionMacros
    Set m_UiOptionRawItems = itemRawItems
    m_IsDropdownExpanded = False

    m_SelectedIndex = selectedIndex
    If m_SelectedIndex < 0 Then m_SelectedIndex = 0
    If m_SelectedIndex > m_UiOptionCaptions.Count Then m_SelectedIndex = 0

    ' Первичная отрисовка состояния (текст header + видимость panel/items).
    If Not private_ApplyUiStateToShapes() Then Exit Function

    If Me.HasSelectedOption() Then
        ' Сразу синхронизируем persistent state.
        selectedId = Me.GetSelectedId()
        If Not private_TryPersistSelectedId(selectedId) Then Exit Function
    End If

    private_InitializeUiState = True
End Function

Private Function private_ApplyUiStateToShapes() As Boolean
    Dim ws As Worksheet
    Dim headerShape As Shape
    Dim panelShape As Shape
    Dim itemShape As Shape
    Dim i As Long
    Dim headerText As String

    If VBA.Len(m_UiHeaderShapeName) = 0 Then Exit Function
    If m_UiOptionShapeNames Is Nothing Then Exit Function
    If m_UiOptionCaptions Is Nothing Then Exit Function

    Set ws = ex_HelpersSheet.fn_GetRuntimeWorksheetByName(m_ControlLayout.LayoutSheetName)
    If ws Is Nothing Then Exit Function

    Set headerShape = private_GetUiShapeByName(ws, m_UiHeaderShapeName)
    If headerShape Is Nothing Then Exit Function

    If Not private_TryRefreshSemanticParts(ws, headerShape) Then Exit Function

    Set panelShape = private_GetUiShapeByName(ws, m_UiDropdownPanelShapeName)

    If m_IsDropdownExpanded Then
        Call private_ReanchorDropdownToHeader(ws, headerShape, panelShape)
        Set panelShape = private_GetUiShapeByName(ws, m_UiDropdownPanelShapeName)
    End If

    ' Header показывает либо выбранный caption, либо placeholder.
    If Me.HasSelectedOption() Then
        headerText = Me.GetSelectedCaption()
    Else
        headerText = m_PlaceholderText
    End If

    ' Схлопываем только когда реально нет данных опций.
    ' При lazy-render shape-коллекция может быть пустой до первого открытия,
    ' но сами данные (caption/id) уже есть.
    If m_UiOptionCaptions.Count = 0 Then
        m_IsDropdownExpanded = False
    End If

    If m_IsDropdownExpanded Then
        headerText = headerText & " ^"
    Else
        headerText = headerText & " v"
    End If

    private_SetShapeText headerShape, headerText
    private_AlignHeaderText headerShape

    ' Open/close panel + item-shapes.
    If Not panelShape Is Nothing Then
        panelShape.Visible = VBA.IIf(m_IsDropdownExpanded, msoTrue, msoFalse)
        If m_IsDropdownExpanded Then panelShape.ZOrder msoBringToFront
    End If

    For i = 1 To m_UiOptionShapeNames.Count
        Set itemShape = private_GetUiShapeByName(ws, VBA.CStr(m_UiOptionShapeNames(i)))
        If itemShape Is Nothing Then GoTo ContinueItem

        private_AlignItemText itemShape

        ' Сначала нормализуем шрифт всех item. Если менять Visible в этом же
        ' проходе, Excel показывает промежуточный кадр с base fontBold у ещё
        ' не обработанных элементов.
        On Error Resume Next
        itemShape.TextFrame.Characters.Font.Bold = (i = m_SelectedIndex)
        itemShape.TextFrame2.TextRange.Font.Bold = (i = m_SelectedIndex)
        On Error GoTo 0

ContinueItem:
    Next i

    ' Видимость меняется отдельной фазой уже после завершения форматирования.
    For i = 1 To m_UiOptionShapeNames.Count
        Set itemShape = private_GetUiShapeByName(ws, VBA.CStr(m_UiOptionShapeNames(i)))
        If itemShape Is Nothing Then GoTo ContinueVisibility
        itemShape.Visible = VBA.IIf(m_IsDropdownExpanded, msoTrue, msoFalse)
        If m_IsDropdownExpanded Then itemShape.ZOrder msoBringToFront
ContinueVisibility:
    Next i

    private_ApplyUiStateToShapes = True
End Function

Private Function private_TryRefreshSemanticParts( _
    ByVal ws As Worksheet, _
    ByVal headerShape As Shape _
) As Boolean
    Dim semanticSignature As String
    Dim selectionSignature As String
    Dim selectionChanged As Boolean
    Dim selectedTags As Collection
    Dim selectedStates As Collection
    Dim optionTags As Collection
    Dim optionStates As Collection
    Dim panelShape As Shape
    Dim itemShape As Shape
    Dim registeredHeaders As Collection
    Dim i As Long

    If ws Is Nothing Or headerShape Is Nothing Then Exit Function
    If m_UiOptionShapeNames Is Nothing Then Exit Function

    If Me.HasSelectedOption() Then
        If Not private_TryGetOptionSemantics(m_SelectedIndex, selectedTags, selectedStates) Then Exit Function
    End If
    semanticSignature = private_BuildSemanticPartsSignature(selectedTags, selectedStates)
    selectionSignature = private_BuildSelectionSemanticSignature(selectedTags, selectedStates)
    selectionChanged = (VBA.StrComp( _
        m_StyledSelectionSignature, selectionSignature, VBA.vbBinaryCompare) <> 0)
    If VBA.StrComp(m_SemanticPartsSignature, semanticSignature, VBA.vbBinaryCompare) = 0 Then
        If Not ex_ControlPartsRuntime.fn_TryResolveControlPartShapes( _
            ws, "select", m_ControlName, "header", registeredHeaders) Then Exit Function
        If Not registeredHeaders Is Nothing Then
            If registeredHeaders.Count > 0 Then
                private_TryRefreshSemanticParts = True
                Exit Function
            End If
        End If
    End If

    ' Dynamic parts принадлежат только этому Select: старые state/tag buckets
    ' удаляются перед публикацией текущего semantic snapshot.
    If Not ex_ControlPartsRuntime.fn_RemoveControlPartsByControl( _
        ws.Name, m_ControlName) Then Exit Function

    If Not private_RegisterShapePart(ws, headerShape, "header") Then Exit Function
    If Me.HasSelectedOption() Then
        If Not private_RegisterShapePart(ws, headerShape, "state-selected") Then Exit Function
        If Not private_RegisterShapeCollectionParts(ws, headerShape, "tag-", selectedTags) Then Exit Function
        If Not private_RegisterShapeCollectionParts(ws, headerShape, "state-", selectedStates) Then Exit Function
    Else
        If Not private_RegisterShapePart(ws, headerShape, "state-placeholder") Then Exit Function
    End If
    If m_IsDropdownExpanded Then
        If Not private_RegisterShapePart(ws, headerShape, "state-open") Then Exit Function
    Else
        If Not private_RegisterShapePart(ws, headerShape, "state-closed") Then Exit Function
    End If

    Set panelShape = private_GetUiShapeByName(ws, m_UiDropdownPanelShapeName)
    If Not panelShape Is Nothing Then
        If Not private_RegisterShapePart(ws, panelShape, "panel") Then Exit Function
    End If

    For i = 1 To m_UiOptionShapeNames.Count
        Set itemShape = private_GetUiShapeByName(ws, VBA.CStr(m_UiOptionShapeNames(i)))
        If itemShape Is Nothing Then GoTo ContinueItem
        If Not private_RegisterShapePart(ws, itemShape, "item") Then Exit Function
        If Not private_TryGetOptionSemantics(i, optionTags, optionStates) Then Exit Function
        ' Item parts имеют отдельный namespace, чтобы header tag/state rules
        ' не перекрашивали строки раскрытого списка.
        If Not private_RegisterShapeCollectionParts(ws, itemShape, "item-tag-", optionTags) Then Exit Function
        If Not private_RegisterShapeCollectionParts(ws, itemShape, "item-state-", optionStates) Then Exit Function
        If i = m_SelectedIndex Then
            If Not private_RegisterShapePart(ws, itemShape, "item-selected") Then Exit Function
        End If
ContinueItem:
    Next i

    ' Полный render позже выполнит единый page style-pass. При интерактивной
    ' смене состояния обновляем только Shape и rules этого Select.
    If Not m_IsControlRenderPass Then
        If Not private_TryApplySemanticStateStyles( _
            ws, headerShape, panelShape, selectionChanged) Then Exit Function
    End If
    m_SemanticPartsSignature = semanticSignature
    m_StyledSelectionSignature = selectionSignature
    private_TryRefreshSemanticParts = True
End Function

Private Function private_BuildSelectionSemanticSignature( _
    ByVal selectedTags As Collection, _
    ByVal selectedStates As Collection _
) As String
    private_BuildSelectionSemanticSignature = _
        VBA.IIf(Me.HasSelectedOption(), "selected", "placeholder") & _
        "|index=" & VBA.CStr(m_SelectedIndex) & _
        "|shapes=" & VBA.CStr(m_UiOptionShapeNames.Count) & _
        "|tags=" & private_CollectionSignature(selectedTags) & _
        "|states=" & private_CollectionSignature(selectedStates)
End Function

Private Function private_BuildSemanticPartsSignature( _
    ByVal selectedTags As Collection, _
    ByVal selectedStates As Collection _
) As String
    Dim optionTags As Collection
    Dim optionStates As Collection
    Dim optionSignature As String
    Dim i As Long

    If Not m_UiOptionRawItems Is Nothing Then
        For i = 1 To m_UiOptionRawItems.Count
            If Not private_TryGetOptionSemantics(i, optionTags, optionStates) Then Exit Function
            optionSignature = optionSignature & "|item" & VBA.CStr(i) & "=" & _
                private_CollectionSignature(optionTags) & ":" & _
                private_CollectionSignature(optionStates)
        Next i
    End If
    private_BuildSemanticPartsSignature = _
        VBA.IIf(Me.HasSelectedOption(), "selected", "placeholder") & _
        "|" & VBA.IIf(m_IsDropdownExpanded, "open", "closed") & _
        "|index=" & VBA.CStr(m_SelectedIndex) & _
        "|shapes=" & VBA.CStr(m_UiOptionShapeNames.Count) & _
        "|tags=" & private_CollectionSignature(selectedTags) & _
        "|states=" & private_CollectionSignature(selectedStates) & _
        optionSignature
End Function

Private Function private_CollectionSignature(ByVal values As Collection) As String
    Dim valueItem As Variant

    If values Is Nothing Then Exit Function
    For Each valueItem In values
        If VBA.Len(private_CollectionSignature) > 0 Then private_CollectionSignature = private_CollectionSignature & ","
        private_CollectionSignature = private_CollectionSignature & _
            VBA.LCase$(VBA.Trim$(VBA.CStr(valueItem)))
    Next valueItem
End Function

Private Function private_TryGetOptionSemantics( _
    ByVal itemIndex As Long, _
    ByRef outTags As Collection, _
    ByRef outStates As Collection _
) As Boolean
    If m_UiOptionRawItems Is Nothing Then Exit Function
    If itemIndex <= 0 Or itemIndex > m_UiOptionRawItems.Count Then Exit Function

    private_TryGetOptionSemantics = private_TryGetRawItemSemantics( _
        m_UiOptionRawItems(itemIndex), outTags, outStates)
End Function

Private Function private_TryGetRawItemSemantics( _
    ByVal rawValue As Variant, _
    ByRef outTags As Collection, _
    ByRef outStates As Collection _
) As Boolean
    Dim contractItem As obj_IButtonGroupItem
    Dim rawItem As Object
    Dim semanticValues As Collection

    Set outTags = New Collection
    Set outStates = New Collection
    If Not VBA.IsObject(rawValue) Then
        private_TryGetRawItemSemantics = True
        Exit Function
    End If
    Set rawItem = rawValue
    On Error Resume Next
    Set contractItem = rawItem
    On Error GoTo 0
    If contractItem Is Nothing Then
        private_TryGetRawItemSemantics = True
        Exit Function
    End If

    Set semanticValues = contractItem.Tags
    If Not semanticValues Is Nothing Then Set outTags = semanticValues
    Set semanticValues = contractItem.States
    If Not semanticValues Is Nothing Then Set outStates = semanticValues
    private_TryGetRawItemSemantics = True
End Function

Private Function private_TryRestoreOptionSemantics( _
    ByVal optionObj As obj_SelectOption, _
    ByVal tagsText As String, _
    ByVal statesText As String _
) As Boolean
    Dim valueItem As Variant
    Dim valueText As String

    If optionObj Is Nothing Then Exit Function
    For Each valueItem In VBA.Split(tagsText, ",")
        valueText = VBA.LCase$(VBA.Trim$(VBA.CStr(valueItem)))
        If VBA.Len(valueText) > 0 Then
            If Not optionObj.AddTag(valueText) Then Exit Function
        End If
    Next valueItem
    For Each valueItem In VBA.Split(statesText, ",")
        valueText = VBA.LCase$(VBA.Trim$(VBA.CStr(valueItem)))
        If VBA.Len(valueText) > 0 Then
            If Not optionObj.SetState(valueText, True) Then Exit Function
        End If
    Next valueItem
    private_TryRestoreOptionSemantics = True
End Function

Private Function private_RegisterShapeCollectionParts( _
    ByVal ws As Worksheet, _
    ByVal shp As Shape, _
    ByVal partPrefix As String, _
    ByVal values As Collection _
) As Boolean
    Dim valueItem As Variant

    If values Is Nothing Then
        private_RegisterShapeCollectionParts = True
        Exit Function
    End If
    For Each valueItem In values
        If Not private_RegisterShapePart( _
            ws, shp, partPrefix & VBA.LCase$(VBA.Trim$(VBA.CStr(valueItem)))) Then Exit Function
    Next valueItem
    private_RegisterShapeCollectionParts = True
End Function

Private Function private_RegisterShapePart( _
    ByVal ws As Worksheet, _
    ByVal shp As Shape, _
    ByVal partName As String _
) As Boolean
    Dim partRange As Range

    If ws Is Nothing Or shp Is Nothing Then Exit Function
    partName = VBA.LCase$(VBA.Trim$(partName))
    If VBA.Len(partName) = 0 Then Exit Function
    On Error Resume Next
    Set partRange = ws.Range(shp.TopLeftCell, shp.BottomRightCell)
    On Error GoTo 0
    If partRange Is Nothing Then Exit Function

    private_RegisterShapePart = ex_ControlPartsRuntime.fn_RegisterControlPart( _
        ws, "select", m_ControlName, partName, partRange, shp)
End Function

Private Function private_TryApplySemanticStateStyles( _
    ByVal ws As Worksheet, _
    ByVal headerShape As Shape, _
    ByVal panelShape As Shape, _
    ByVal selectionChanged As Boolean _
) As Boolean
    Dim pageBase As obj_PageBase
    Dim itemShape As Shape
    Dim i As Long

    If ws Is Nothing Or headerShape Is Nothing Then Exit Function
    If Not m_ControlBase Is Nothing Then Set pageBase = m_ControlBase.PageBase
    If pageBase Is Nothing Then Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Function
    If pageBase.XmlDom Is Nothing Then Exit Function

    If Not ex_StylePipelineEngine.fn_ApplyControlStyleToShape( _
        headerShape, pageBase.XmlDom, True) Then Exit Function
    If selectionChanged Then
        If Not panelShape Is Nothing Then
            If Not ex_StylePipelineEngine.fn_ApplyControlStyleToShape( _
                panelShape, pageBase.XmlDom, True) Then Exit Function
        End If
        For i = 1 To m_UiOptionShapeNames.Count
            Set itemShape = private_GetUiShapeByName(ws, VBA.CStr(m_UiOptionShapeNames(i)))
            If itemShape Is Nothing Then GoTo ContinueItem
            If Not ex_StylePipelineEngine.fn_ApplyControlStyleToShape( _
                itemShape, pageBase.XmlDom, True) Then Exit Function
ContinueItem:
        Next i
    End If

    private_TryApplySemanticStateStyles = _
        ex_StylePipelineEngine.fn_ApplyControlPartStylesForControl( _
            ws, pageBase.XmlDom, m_ControlName, True)
End Function

Private Sub private_AlignHeaderText(ByVal shp As Shape)
    If shp Is Nothing Then Exit Sub
    On Error Resume Next
    shp.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp.TextFrame.VerticalAlignment = xlVAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    On Error GoTo 0
End Sub

Private Sub private_AlignItemText(ByVal shp As Shape)
    If shp Is Nothing Then Exit Sub
    On Error Resume Next
    shp.TextFrame.HorizontalAlignment = xlHAlignLeft
    shp.TextFrame.VerticalAlignment = xlVAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    On Error GoTo 0
End Sub

Private Function private_ReanchorDropdownToHeader( _
    ByVal ws As Worksheet, _
    ByVal headerShape As Shape, _
    ByVal panelShape As Shape _
) As Boolean
    Dim itemShape As Shape
    Dim itemTop As Double
    Dim itemWidth As Double
    Dim panelLeft As Double
    Dim panelTop As Double
    Dim panelWidth As Double
    Dim panelHeight As Double
    Dim i As Long
    Dim renderItemCount As Long

    If ws Is Nothing Then Exit Function
    If headerShape Is Nothing Then Exit Function

    If m_UiOptionShapeNames Is Nothing Then
        private_ReanchorDropdownToHeader = True
        Exit Function
    End If

    renderItemCount = m_UiOptionShapeNames.Count
    panelLeft = headerShape.Left
    panelTop = headerShape.Top + headerShape.Height
    panelWidth = headerShape.Width
    panelHeight = private_CalcPanelHeight(renderItemCount)

    If panelShape Is Nothing Then
        Set panelShape = private_GetUiShapeByName(ws, m_UiDropdownPanelShapeName)
    End If

    If Not panelShape Is Nothing Then
        On Error Resume Next
        panelShape.Left = panelLeft
        panelShape.Top = panelTop
        panelShape.Width = panelWidth
        If renderItemCount > 0 Then
            panelShape.Height = panelHeight
        Else
            panelShape.Height = headerShape.Height
        End If
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    End If

    itemWidth = panelWidth
    For i = 1 To renderItemCount
        Set itemShape = private_GetUiShapeByName(ws, VBA.CStr(m_UiOptionShapeNames(i)))
        If itemShape Is Nothing Then GoTo ContinueItem

        itemTop = panelTop + VBA.CDbl(i - 1) * (m_ItemHeight + m_ItemMargin)
        On Error Resume Next
        itemShape.Left = panelLeft
        itemShape.Top = itemTop
        itemShape.Width = itemWidth
        itemShape.Height = m_ItemHeight
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0

ContinueItem:
    Next i

    private_ReanchorDropdownToHeader = True
End Function

Private Function private_GetOptionCollectionText(ByVal values As Collection, ByVal idx As Long) As String
    If values Is Nothing Then Exit Function
    If idx <= 0 Or idx > values.Count Then Exit Function
    private_GetOptionCollectionText = VBA.Trim$(VBA.CStr(values(idx)))
End Function

Private Function private_RunOptionMacro(ByVal macroRef As String) As Boolean
    macroRef = VBA.Trim$(macroRef)
    If VBA.Len(macroRef) = 0 Then
        private_RunOptionMacro = True
        Exit Function
    End If

    private_RunOptionMacro = rt_Bridge.fn_RunCallback(macroRef, m_CallbackContext)
    If private_RunOptionMacro Then Exit Function
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "Select: failed to execute callback '" & macroRef & "' for control '" & m_ControlName & "'."
#End If
End Function

Private Function private_TryRefreshItemsAndRerenderAfterDropDownOpenedCallback() As Boolean
    Dim pageBase As obj_PageBase
    Dim currentSelectedId As String
    Dim selectedIndexRefreshed As Long
    Dim previousUiCount As Long
    Dim resolvedCount As Long

    ' Этот метод намеренно перерисовывает только текущий Select-контрол.
    ' Полный page rerender здесь не нужен: нам важно обновить список "на месте"
    ' в том же клике по header, до фактического открытия dropdown.
    Set pageBase = Nothing
    If Not m_ControlBase Is Nothing Then Set pageBase = m_ControlBase.PageBase
    If pageBase Is Nothing Then
        Set pageBase = m_Page.GetPageBase()
    End If
    If pageBase Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Select: page base is not resolved for DropDownOpened refresh in control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    currentSelectedId = VBA.Trim$(Me.GetSelectedId())
    If VBA.Len(currentSelectedId) = 0 Then currentSelectedId = VBA.Trim$(m_SelectedIdRaw)
    previousUiCount = 0
    If Not m_UiOptionIds Is Nothing Then previousUiCount = m_UiOptionIds.Count

    ' 1) Перечитываем itemsSource после DropDownOpened callback.
    If Not ex_RuntimeSourceResolver.fn_TryResolveItemsSource(pageBase.RuntimeSources, m_ItemsSourceRaw, m_Items) Then Exit Function
    ' 2) Пересобираем плоские буферы caption/id/action/raw.
    If Not private_TryBuildItemBuffers() Then Exit Function
    resolvedCount = 0
    If Not m_ItemIds Is Nothing Then resolvedCount = m_ItemIds.Count


    ' 3) Сохраняем выбор пользователя (по id), если элемент все еще существует.
    selectedIndexRefreshed = private_FindSelectedIndexById(currentSelectedId)
    m_SelectedIndex = selectedIndexRefreshed

    ' 4) Если состав опций не изменился, не пересоздаем shapes/routes:
    ' это самый дорогой шаг при открытии dropdown.
    If private_AreResolvedItemBuffersEqualToUiState() Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogInfo "select:dropdown-refresh skip-rerender control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "'"
#End If
        Set m_UiOptionCaptions = m_ItemCaptions
        Set m_UiOptionIds = m_ItemIds
        Set m_UiOptionActionMacros = m_ItemActionMacros
        Set m_UiOptionRawItems = m_ItemRawItems

        private_TryRefreshItemsAndRerenderAfterDropDownOpenedCallback = private_ApplyUiStateToShapes()
        Exit Function
    End If

    ' 5) Локальный rerender Select (header/panel/items + routes),
    ' только когда источник действительно изменился.
    Call obj_IControl_Render
    private_TryRefreshItemsAndRerenderAfterDropDownOpenedCallback = True
End Function

Private Function private_TryEnsureDropdownItemsReady() As Boolean
    Dim ws As Worksheet
    Dim headerShape As Shape
    Dim panelShape As Shape
    Dim itemShape As Shape
    Dim itemShapeNames As Collection
    Dim callbackMacroRef As String
    Dim renderItemCount As Long
    Dim panelLeft As Double
    Dim panelTop As Double
    Dim panelWidth As Double
    Dim panelHeight As Double
    Dim itemTop As Double
    Dim i As Long

    If m_ControlLayout Is Nothing Then Exit Function
    If m_UiOptionCaptions Is Nothing Then Exit Function
    If m_UiOptionIds Is Nothing Then Exit Function
    If m_UiOptionActionMacros Is Nothing Then Exit Function
    If m_UiOptionRawItems Is Nothing Then Exit Function

    Set ws = ex_HelpersSheet.fn_GetRuntimeWorksheetByName(m_ControlLayout.LayoutSheetName)
    If ws Is Nothing Then Exit Function

    Set headerShape = private_GetUiShapeByName(ws, m_UiHeaderShapeName)
    If headerShape Is Nothing Then Exit Function

    Set panelShape = private_GetUiShapeByName(ws, m_UiDropdownPanelShapeName)

    renderItemCount = m_UiOptionCaptions.Count
    callbackMacroRef = private_GetRuntimeCallbackMacroRef()
    If VBA.Len(callbackMacroRef) = 0 Then Exit Function

    panelLeft = headerShape.Left
    panelTop = headerShape.Top + headerShape.Height
    panelWidth = headerShape.Width
    panelHeight = private_CalcPanelHeight(renderItemCount)

    If panelShape Is Nothing Then
        If renderItemCount > 0 Then
            Set panelShape = private_CreateShapeByBounds(ws, panelLeft, panelTop, panelWidth, panelHeight, "panel", VBA.vbNullString)
        Else
            Set panelShape = private_CreateShapeByBounds(ws, headerShape.Left, headerShape.Top, headerShape.Width, headerShape.Height, "panel", VBA.vbNullString)
        End If
        If panelShape Is Nothing Then Exit Function
    Else
        On Error Resume Next
        panelShape.Left = panelLeft
        panelShape.Top = panelTop
        panelShape.Width = panelWidth
        If renderItemCount > 0 Then
            panelShape.Height = panelHeight
        Else
            panelShape.Height = headerShape.Height
        End If
        On Error GoTo 0
    End If

    If renderItemCount > 0 Then
        If VBA.Len(VBA.Trim$(m_PanelStyleName)) = 0 Then
            private_ApplyPanelVisualDefaults panelShape
        End If
    Else
        panelShape.Visible = msoFalse
    End If

    private_DeleteStaleItemShapes ws, renderItemCount

    Set itemShapeNames = New Collection
    For i = 1 To renderItemCount
        itemTop = panelTop + VBA.CDbl(i - 1) * (m_ItemHeight + m_ItemMargin)
        Set itemShape = private_CreateShapeByBounds(ws, panelLeft, itemTop, panelWidth, m_ItemHeight, "item" & VBA.CStr(i), callbackMacroRef)
        If itemShape Is Nothing Then Exit Function

        ' Lazy item создаётся скрытым: базовый/semantic styles и шрифт должны
        ' примениться до первого видимого кадра раскрытого списка.
        itemShape.Visible = msoFalse
        private_SetShapeText itemShape, VBA.CStr(m_UiOptionCaptions(i))
        If VBA.Len(VBA.Trim$(m_ItemStyleName)) = 0 Then
            private_ApplyItemVisualDefaults itemShape
        End If
        itemShapeNames.Add itemShape.Name
    Next i

    Set m_UiOptionShapeNames = itemShapeNames
    If Not private_TryBindUiRoutes(ws, m_UiHeaderShapeName, itemShapeNames) Then Exit Function
    private_TryEnsureDropdownItemsReady = True
End Function

Private Function private_AreResolvedItemBuffersEqualToUiState() As Boolean
    Dim i As Long
    Dim sourceTags As Collection
    Dim sourceStates As Collection
    Dim uiTags As Collection
    Dim uiStates As Collection

    If m_ItemCaptions Is Nothing Then Exit Function
    If m_ItemIds Is Nothing Then Exit Function
    If m_ItemActionMacros Is Nothing Then Exit Function

    If m_UiOptionCaptions Is Nothing Then Exit Function
    If m_UiOptionIds Is Nothing Then Exit Function
    If m_UiOptionActionMacros Is Nothing Then Exit Function

    If m_ItemCaptions.Count <> m_UiOptionCaptions.Count Then Exit Function
    If m_ItemIds.Count <> m_UiOptionIds.Count Then Exit Function
    If m_ItemActionMacros.Count <> m_UiOptionActionMacros.Count Then Exit Function

    For i = 1 To m_ItemIds.Count
        If VBA.StrComp(VBA.CStr(m_ItemIds(i)), VBA.CStr(m_UiOptionIds(i)), VBA.vbBinaryCompare) <> 0 Then Exit Function
        If VBA.StrComp(VBA.CStr(m_ItemCaptions(i)), VBA.CStr(m_UiOptionCaptions(i)), VBA.vbBinaryCompare) <> 0 Then Exit Function
        If VBA.StrComp(VBA.CStr(m_ItemActionMacros(i)), VBA.CStr(m_UiOptionActionMacros(i)), VBA.vbBinaryCompare) <> 0 Then Exit Function
        If Not private_TryGetRawItemSemantics(m_ItemRawItems(i), sourceTags, sourceStates) Then Exit Function
        If Not private_TryGetRawItemSemantics(m_UiOptionRawItems(i), uiTags, uiStates) Then Exit Function
        If VBA.StrComp(private_CollectionSignature(sourceTags), private_CollectionSignature(uiTags), VBA.vbBinaryCompare) <> 0 Then Exit Function
        If VBA.StrComp(private_CollectionSignature(sourceStates), private_CollectionSignature(uiStates), VBA.vbBinaryCompare) <> 0 Then Exit Function
    Next i

    private_AreResolvedItemBuffersEqualToUiState = True
End Function

Private Function private_TryBuildItemBuffers() As Boolean
    Dim itemRaw As Variant
    Dim itemCaption As String
    Dim itemId As String
    Dim itemAction As String

    If m_Items Is Nothing Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Select: itemsSource resolved to Nothing for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    ' Нормализуем itemsSource в плоские буферы:
    ' Caption / Id / ItemAction / RawItem.
    Set m_ItemCaptions = New Collection
    Set m_ItemIds = New Collection
    Set m_ItemActionMacros = New Collection
    Set m_ItemRawItems = New Collection

    For Each itemRaw In m_Items
        itemCaption = VBA.vbNullString
        itemId = VBA.vbNullString
        itemAction = VBA.vbNullString

        If Not private_TryResolveItemMetadata(itemRaw, itemCaption, itemId, itemAction) Then Exit Function

        m_ItemCaptions.Add itemCaption
        m_ItemIds.Add itemId
        m_ItemActionMacros.Add itemAction
        m_ItemRawItems.Add itemRaw
    Next itemRaw

    private_TryBuildItemBuffers = True
End Function

Private Function private_TryResolveItemMetadata( _
    ByVal itemRaw As Variant, _
    ByRef outCaption As String, _
    ByRef outId As String, _
    ByRef outItemActionMacro As String _
) As Boolean
    Dim itemObj As Object
    Dim actionRaw As String

    ' Scalar item: Caption=Id=значение, action пустой.
    If Not VBA.IsObject(itemRaw) Then
        outCaption = VBA.CStr(itemRaw)
        outId = VBA.CStr(itemRaw)
        outItemActionMacro = VBA.vbNullString
        private_TryResolveItemMetadata = True
        Exit Function
    End If

    Set itemObj = itemRaw

    ' Object item: сначала fast-path для obj_SelectOption,
    ' затем универсальный путь (CallByName/Dictionary).
    Select Case VBA.LCase$(VBA.TypeName(itemObj))
        Case "obj_selectoption"
            outCaption = VBA.CStr(itemObj.Caption)
            outId = VBA.CStr(itemObj.Id)
            actionRaw = VBA.Trim$(VBA.CStr(itemObj.OnSelect))

        Case Else
            If Not private_TryReadObjectMemberText(itemObj, "Caption", True, outCaption) Then Exit Function
            If Not private_TryReadObjectMemberText(itemObj, "Id", True, outId) Then Exit Function
            If Not private_TryReadObjectMemberText(itemObj, "OnSelect", False, actionRaw) Then Exit Function
            actionRaw = VBA.Trim$(actionRaw)
    End Select

    If VBA.Len(VBA.Trim$(outId)) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Select: item Id is empty in control '" & m_ControlName & "'."
#End If
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(outCaption)) = 0 Then outCaption = outId

    outItemActionMacro = VBA.vbNullString
    If VBA.Len(actionRaw) > 0 Then
        If VBA.InStr(1, actionRaw, ".", VBA.vbBinaryCompare) > 0 Or VBA.InStr(1, actionRaw, "!", VBA.vbBinaryCompare) > 0 Then
            outItemActionMacro = private_QualifyMacroName(actionRaw)
        Else
            outItemActionMacro = actionRaw
        End If
    End If

    private_TryResolveItemMetadata = True
End Function

Private Function private_TryReadObjectMemberText( _
    ByVal sourceObject As Object, _
    ByVal memberName As String, _
    ByVal isRequired As Boolean, _
    ByRef outText As String _
) As Boolean
    Dim dictObj As Object
    Dim scalarValue As Variant

    outText = VBA.vbNullString
    If sourceObject Is Nothing Then
        If isRequired Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "Select: item object is Nothing while reading member '" & memberName & "'."
#End If
            Exit Function
        End If
        private_TryReadObjectMemberText = True
        Exit Function
    End If

    Set dictObj = private_AsDictionary(sourceObject)
    If Not dictObj Is Nothing Then
        If Not dictObj.Exists(memberName) Then
            If isRequired Then
#If LOGGING_DEBUG_ENABLED Then
                ex_Core.fn_Diagnostic_LogError "Select: member '" & memberName & "' was not found on dictionary item."
#End If
                Exit Function
            End If

            private_TryReadObjectMemberText = True
            Exit Function
        End If

        scalarValue = dictObj.Item(memberName)
        If VBA.IsObject(scalarValue) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "Select: member '" & memberName & "' must resolve to scalar value."
#End If
            Exit Function
        End If

        outText = VBA.CStr(scalarValue)
        private_TryReadObjectMemberText = True
        Exit Function
    End If

    On Error Resume Next
    scalarValue = VBA.CallByName(sourceObject, memberName, VbGet)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0

        If isRequired Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "Select: member '" & memberName & "' was not found on object '" & VBA.TypeName(sourceObject) & "'."
#End If
            Exit Function
        End If

        private_TryReadObjectMemberText = True
        Exit Function
    End If
    On Error GoTo 0

    If VBA.IsObject(scalarValue) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Select: member '" & memberName & "' on object '" & VBA.TypeName(sourceObject) & "' must resolve to scalar value."
#End If
        Exit Function
    End If

    outText = VBA.CStr(scalarValue)
    private_TryReadObjectMemberText = True
End Function

Private Function private_AsDictionary(ByVal sourceObject As Object) As Object
    If sourceObject Is Nothing Then Exit Function
    If VBA.LCase$(VBA.TypeName(sourceObject)) <> "dictionary" Then Exit Function
    Set private_AsDictionary = sourceObject
End Function

Private Function private_TryResolveSelectedIdText(ByRef outSelectedIdText As String) As Boolean
    Dim storedSelectedId As String

    ' Порядок получения selectedId:
    ' 1) selectedId в XML
    ' 2) сохраненное состояние в CustomXMLPart (obj_SelectControlVMStatic)
    outSelectedIdText = VBA.Trim$(m_SelectedIdRaw)
    If VBA.Len(outSelectedIdText) > 0 Then
        private_TryResolveSelectedIdText = True
        Exit Function
    End If

    If Not private_TryLoadStoredSelectedId(storedSelectedId) Then Exit Function
    outSelectedIdText = VBA.Trim$(storedSelectedId)

    private_TryResolveSelectedIdText = True
End Function

Private Function private_TryPersistSelectedId(ByVal selectedId As String) As Boolean
    Dim selectControlVMStatic As obj_SelectControlVMStatic

    selectedId = VBA.Trim$(selectedId)
    If VBA.Len(VBA.Trim$(m_SelectStateKey)) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Select: state key is empty for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    Set selectControlVMStatic = New obj_SelectControlVMStatic
    private_TryPersistSelectedId = selectControlVMStatic.SetSelectedId(m_SelectStateKey, selectedId)
End Function

Private Function private_TryLoadStoredSelectedId(ByRef outSelectedId As String) As Boolean
    Dim selectControlVMStatic As obj_SelectControlVMStatic

    If VBA.Len(VBA.Trim$(m_SelectStateKey)) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Select: state key is empty for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    Set selectControlVMStatic = New obj_SelectControlVMStatic
    private_TryLoadStoredSelectedId = selectControlVMStatic.TryGetSelectedId(m_SelectStateKey, outSelectedId)
End Function

Private Function private_FindSelectedIndexById(ByVal selectedIdText As String) As Long
    Dim i As Long

    selectedIdText = VBA.LCase$(VBA.Trim$(selectedIdText))
    If VBA.Len(selectedIdText) = 0 Then Exit Function
    If m_ItemIds Is Nothing Then Exit Function

    For i = 1 To m_ItemIds.Count
        If VBA.LCase$(VBA.Trim$(VBA.CStr(m_ItemIds(i)))) = selectedIdText Then
            private_FindSelectedIndexById = i
            Exit Function
        End If
    Next i
End Function

Private Function private_GetRenderItemCount() As Long
    If m_ItemCaptions Is Nothing Then Exit Function
    private_GetRenderItemCount = m_ItemCaptions.Count
End Function

Private Function private_TryBuildHeaderRange(ByVal ws As Worksheet, ByRef outRange As Range) As Boolean
    If ws Is Nothing Then Exit Function

    On Error GoTo EH_RANGE
    Set outRange = ws.Range(ws.Cells(m_ControlLayout.RowStart, m_ControlLayout.ColStart), ws.Cells(m_ControlLayout.RowEnd, m_ControlLayout.ColEnd))
    private_TryBuildHeaderRange = True
    Exit Function

EH_RANGE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "Select: failed to resolve header range for control '" & m_ControlName & "'."
#End If
End Function

Private Function private_CalcPanelHeight(ByVal renderItemCount As Long) As Double
    If renderItemCount <= 0 Then Exit Function

    private_CalcPanelHeight = VBA.CDbl(renderItemCount) * m_ItemHeight
    If renderItemCount > 1 Then
        private_CalcPanelHeight = private_CalcPanelHeight + VBA.CDbl(renderItemCount - 1) * m_ItemMargin
    End If
End Function

Private Function private_CreateShapeByRange( _
    ByVal ws As Worksheet, _
    ByVal targetRange As Range, _
    ByVal suffix As String, _
    ByVal onActionMacroRef As String _
) As Shape
    Dim shapeName As String
    Dim shp As Shape
    Dim shapeRole As String
    Dim shapeStyle As String
    Dim metaMap As Object

    If ws Is Nothing Then Exit Function
    If targetRange Is Nothing Then Exit Function

    If targetRange.Width <= 0# Or targetRange.Height <= 0# Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Select: target range has non-positive width/height for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    shapeName = private_BuildShapeName(suffix)

    Set shp = private_GetUiShapeByName(ws, shapeName)
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, targetRange.Left, targetRange.Top, targetRange.Width, targetRange.Height)
        shp.Name = shapeName
    Else
        shp.Left = targetRange.Left
        shp.Top = targetRange.Top
        shp.Width = targetRange.Width
        shp.Height = targetRange.Height
    End If

    If VBA.Len(VBA.Trim$(onActionMacroRef)) > 0 Then
        If Not private_TryAssignShapeOnActionIfChanged(shp, onActionMacroRef) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "Select: failed to bind click action for shape '" & shapeName & "' in control '" & m_ControlName & "'."
#End If
            Exit Function
        End If
    End If

    shapeRole = VBA.LCase$(VBA.Trim$(suffix))
    shapeStyle = VBA.vbNullString
    Select Case shapeRole
        Case "header"
            shapeStyle = VBA.Trim$(m_ControlLayout.StyleName)
        Case "panel"
            shapeStyle = VBA.Trim$(m_PanelStyleName)
        Case Else
            shapeRole = "item"
            shapeStyle = VBA.Trim$(m_ItemStyleName)
    End Select

    On Error Resume Next
    shp.Placement = xlMoveAndSize
    Err.Clear
    On Error GoTo EH_SHAPE

    Set metaMap = VBA.CreateObject("Scripting.Dictionary")
    metaMap.CompareMode = 1
    metaMap("pn.control") = m_ControlName
    metaMap("pn.role") = shapeRole
    If VBA.Len(shapeStyle) > 0 Then
        metaMap("pn.style") = shapeStyle
    Else
        metaMap("pn.style") = VBA.vbNullString
    End If
    If Not ex_ShapeMetaRuntime.fn_TrySetShapeMetaValues(shp, metaMap) Then Exit Function

    Set private_CreateShapeByRange = shp
    Exit Function

EH_SHAPE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "Select: failed to create shape '" & shapeName & "' for control '" & m_ControlName & "': " & Err.Description
#End If
End Function

Private Function private_CreateShapeByBounds( _
    ByVal ws As Worksheet, _
    ByVal shapeLeft As Double, _
    ByVal shapeTop As Double, _
    ByVal shapeWidth As Double, _
    ByVal shapeHeight As Double, _
    ByVal suffix As String, _
    ByVal onActionMacroRef As String _
) As Shape
    Dim shapeName As String
    Dim shp As Shape
    Dim shapeRole As String
    Dim shapeStyle As String
    Dim metaMap As Object

    If ws Is Nothing Then Exit Function
    If shapeWidth <= 0# Or shapeHeight <= 0# Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Select: target bounds have non-positive width/height for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    shapeName = private_BuildShapeName(suffix)

    Set shp = private_GetUiShapeByName(ws, shapeName)
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, shapeLeft, shapeTop, shapeWidth, shapeHeight)
        shp.Name = shapeName
    Else
        shp.Left = shapeLeft
        shp.Top = shapeTop
        shp.Width = shapeWidth
        shp.Height = shapeHeight
    End If

    If VBA.Len(VBA.Trim$(onActionMacroRef)) > 0 Then
        If Not private_TryAssignShapeOnActionIfChanged(shp, onActionMacroRef) Then
#If LOGGING_DEBUG_ENABLED Then
            ex_Core.fn_Diagnostic_LogError "Select: failed to bind click action for shape '" & shapeName & "' in control '" & m_ControlName & "'."
#End If
            Exit Function
        End If
    End If

    shapeRole = VBA.LCase$(VBA.Trim$(suffix))
    shapeStyle = VBA.vbNullString
    Select Case shapeRole
        Case "header"
            shapeStyle = VBA.Trim$(m_ControlLayout.StyleName)
        Case "panel"
            shapeStyle = VBA.Trim$(m_PanelStyleName)
        Case Else
            shapeRole = "item"
            shapeStyle = VBA.Trim$(m_ItemStyleName)
    End Select

    On Error Resume Next
    shp.Placement = xlMoveAndSize
    Err.Clear
    On Error GoTo EH_SHAPE

    Set metaMap = VBA.CreateObject("Scripting.Dictionary")
    metaMap.CompareMode = 1
    metaMap("pn.control") = m_ControlName
    metaMap("pn.role") = shapeRole
    If VBA.Len(shapeStyle) > 0 Then
        metaMap("pn.style") = shapeStyle
    Else
        metaMap("pn.style") = VBA.vbNullString
    End If
    If Not ex_ShapeMetaRuntime.fn_TrySetShapeMetaValues(shp, metaMap) Then Exit Function

    Set private_CreateShapeByBounds = shp
    Exit Function

EH_SHAPE:
#If LOGGING_DEBUG_ENABLED Then
    ex_Core.fn_Diagnostic_LogError "Select: failed to create floating shape '" & shapeName & "' for control '" & m_ControlName & "': " & Err.Description
#End If
End Function

Private Sub private_ApplyHeaderVisualDefaults(ByVal shp As Shape)
    If shp Is Nothing Then Exit Sub

    shp.Fill.Visible = msoTrue
    shp.Fill.ForeColor.RGB = VBA.RGB(67, 142, 32)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = VBA.RGB(40, 93, 20)
    shp.Line.Weight = 1#

    On Error Resume Next
    shp.TextFrame.Characters.Font.Color = VBA.RGB(10, 10, 10)
    shp.TextFrame.Characters.Font.Bold = False
    shp.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp.TextFrame.VerticalAlignment = xlVAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = VBA.RGB(10, 10, 10)
    shp.TextFrame2.TextRange.Font.Bold = False
    On Error GoTo 0
End Sub

Private Sub private_ApplyPanelVisualDefaults(ByVal shp As Shape)
    If shp Is Nothing Then Exit Sub

    shp.Fill.Visible = msoTrue
    shp.Fill.ForeColor.RGB = VBA.RGB(43, 49, 56)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = VBA.RGB(28, 32, 36)
    shp.Line.Weight = 0.75
End Sub

Private Sub private_ApplyItemVisualDefaults(ByVal shp As Shape)
    If shp Is Nothing Then Exit Sub

    shp.Fill.Visible = msoTrue
    shp.Fill.ForeColor.RGB = VBA.RGB(59, 66, 74)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = VBA.RGB(40, 45, 50)
    shp.Line.Weight = 0.75

    On Error Resume Next
    shp.TextFrame.Characters.Font.Color = VBA.RGB(245, 245, 245)
    shp.TextFrame.HorizontalAlignment = xlHAlignLeft
    shp.TextFrame.VerticalAlignment = xlVAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = VBA.RGB(245, 245, 245)
    On Error GoTo 0
End Sub

Private Sub private_SetShapeText(ByVal shp As Shape, ByVal textValue As String)
    If shp Is Nothing Then Exit Sub

    On Error Resume Next
    shp.TextFrame2.TextRange.Text = textValue
    shp.TextFrame.Characters.Text = textValue
    On Error GoTo 0
End Sub

Private Sub private_DeleteControlShapes(ByVal ws As Worksheet)
    Dim i As Long
    Dim shp As Shape
    Dim controlMeta As String
    Dim deletedCount As Long

    If ws Is Nothing Then Exit Sub

    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        controlMeta = VBA.LCase$(VBA.Trim$(ex_ShapeMetaRuntime.fn_GetShapeMetaValue(shp, "pn.control", VBA.vbNullString)))
        If VBA.Len(controlMeta) = 0 Then GoTo ContinueShape
        If controlMeta = VBA.LCase$(VBA.Trim$(m_ControlName)) Then
            deletedCount = deletedCount + 1
            shp.Delete
        End If
ContinueShape:
    Next i

#If LOGGING_DEBUG_ENABLED Then
    If deletedCount > 0 Then ex_Core.fn_Diagnostic_LogInfo "select:delete-control-shapes control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' count=" & VBA.CStr(deletedCount)
#End If
End Sub

Private Sub private_DeleteStaleItemShapes(ByVal ws As Worksheet, ByVal keepItemCount As Long)
    Dim idx As Long
    Dim staleShape As Shape

    If ws Is Nothing Then Exit Sub
    If keepItemCount < 0 Then keepItemCount = 0

    ' Shape-имена itemN формируются последовательно, поэтому удаляем хвост
    ' до первого пропуска (быстрый O(число удалений), без обхода всех Shapes).
    idx = keepItemCount + 1
    Do
        Set staleShape = private_GetUiShapeByName(ws, private_BuildShapeName("item" & VBA.CStr(idx)))
        If staleShape Is Nothing Then Exit Do

        On Error Resume Next
        staleShape.Delete
        On Error GoTo 0
        idx = idx + 1
    Loop
End Sub

Private Function private_TryAssignShapeOnActionIfChanged(ByVal shp As Shape, ByVal macroRef As String) As Boolean
    Dim currentMacroRef As String

    If shp Is Nothing Then Exit Function
    macroRef = VBA.Trim$(macroRef)
    If VBA.Len(macroRef) = 0 Then
        private_TryAssignShapeOnActionIfChanged = True
        Exit Function
    End If

    On Error Resume Next
    currentMacroRef = VBA.Trim$(VBA.CStr(shp.OnAction))
    If Err.Number <> 0 Then
        Err.Clear
        currentMacroRef = VBA.vbNullString
    End If
    On Error GoTo 0

    If VBA.StrComp(currentMacroRef, macroRef, VBA.vbBinaryCompare) <> 0 Then
        On Error GoTo EH_SET
        shp.OnAction = macroRef
        On Error GoTo 0
    End If

    private_TryAssignShapeOnActionIfChanged = True
    Exit Function

EH_SET:
    On Error GoTo 0
End Function

Private Function private_BuildShapeName(ByVal suffix As String) As String
    Dim normalizedSuffix As String
    Dim normalizedControlName As String
    Dim maxControlAliasLen As Long
    Dim maxSuffixAliasLen As Long
    Dim controlAlias As String
    Dim suffixAlias As String
    Dim shapeName As String

    normalizedSuffix = private_NormalizeNamePart(suffix)
    normalizedControlName = private_NormalizeNamePart(m_ControlName)
    maxControlAliasLen = SHAPE_CONTROL_HASH_LEN
    maxSuffixAliasLen = SHAPE_SUFFIX_HASH_LEN

    If VBA.Len("sel__") + maxControlAliasLen + maxSuffixAliasLen > SHAPE_NAME_MAX_LEN Then
        maxSuffixAliasLen = SHAPE_NAME_MAX_LEN - VBA.Len("sel__") - maxControlAliasLen
    End If

    controlAlias = private_BuildShortControlAlias(normalizedControlName, maxControlAliasLen)
    suffixAlias = private_BuildShortControlAlias(normalizedSuffix, maxSuffixAliasLen)

    shapeName = "sel_" & controlAlias & "_" & suffixAlias
    private_RememberShapeNameMapping normalizedSuffix, shapeName

    private_BuildShapeName = shapeName
End Function

Private Function private_NormalizeNamePart(ByVal rawText As String) As String
    Dim i As Long
    Dim ch As String
    Dim outText As String

    rawText = VBA.Trim$(rawText)
    If VBA.Len(rawText) = 0 Then
        private_NormalizeNamePart = "x"
        Exit Function
    End If

    For i = 1 To VBA.Len(rawText)
        ch = VBA.Mid$(rawText, i, 1)
        If (ch >= "A" And ch <= "Z") Or _
           (ch >= "a" And ch <= "z") Or _
           (ch >= "0" And ch <= "9") Or _
           ch = "_" Then
            outText = outText & ch
        Else
            outText = outText & "_"
        End If
    Next i

    If VBA.Len(outText) = 0 Then outText = "x"
    private_NormalizeNamePart = VBA.Left$(outText, 120)
End Function

Private Function private_BuildShortControlAlias(ByVal normalizedControlName As String, ByVal maxLen As Long) As String
    Dim hashText As String

    normalizedControlName = VBA.LCase$(VBA.Trim$(normalizedControlName))
    If VBA.Len(normalizedControlName) = 0 Then normalizedControlName = "x"

    If maxLen <= 0 Then
        private_BuildShortControlAlias = "x"
        Exit Function
    End If

    If VBA.Len(normalizedControlName) <= maxLen Then
        private_BuildShortControlAlias = normalizedControlName
        Exit Function
    End If

    hashText = private_ComputeStableHexHash(normalizedControlName)
    private_BuildShortControlAlias = VBA.Left$(hashText, maxLen)
End Function

Private Sub private_RememberShapeNameMapping(ByVal xmlSuffix As String, ByVal shapeName As String)
    Dim suffixKey As String
    Dim shapeKey As String

    suffixKey = VBA.LCase$(VBA.Trim$(xmlSuffix))
    shapeKey = VBA.LCase$(VBA.Trim$(shapeName))
    If VBA.Len(suffixKey) = 0 Then Exit Sub
    If VBA.Len(shapeKey) = 0 Then Exit Sub

    If m_ShapeNameByXmlSuffix Is Nothing Then
        Set m_ShapeNameByXmlSuffix = VBA.CreateObject("Scripting.Dictionary")
        m_ShapeNameByXmlSuffix.CompareMode = 1
    End If
    If m_XmlSuffixByShapeName Is Nothing Then
        Set m_XmlSuffixByShapeName = VBA.CreateObject("Scripting.Dictionary")
        m_XmlSuffixByShapeName.CompareMode = 1
    End If

    m_ShapeNameByXmlSuffix(suffixKey) = shapeName
    m_XmlSuffixByShapeName(shapeKey) = xmlSuffix
End Sub

Private Function private_ComputeStableHexHash(ByVal sourceText As String) As String
    Dim i As Long
    Dim hashValue As Double
    Dim codePoint As Long
    Dim workValue As Double
    Dim digitValue As Long
    Dim hexChars As String
    Dim outHex As String
    Dim modBase As Double

    sourceText = VBA.CStr(sourceText)
    hashValue = 2166136261#
    modBase = 4294967296#

    For i = 1 To VBA.Len(sourceText)
        codePoint = VBA.AscW(VBA.Mid$(sourceText, i, 1))
        If codePoint < 0 Then codePoint = codePoint + 65536

        ' Не используем VBA.Mod: он приводит к целочисленной арифметике
        ' и может дать Overflow на промежуточных значениях.
        hashValue = (hashValue * 16777619#) + VBA.CDbl(codePoint)
        hashValue = hashValue - modBase * VBA.Int(hashValue / modBase)
    Next i

    workValue = hashValue
    hexChars = "0123456789abcdef"
    outHex = VBA.vbNullString

    For i = 1 To SHAPE_CONTROL_HASH_LEN
        digitValue = VBA.CLng(workValue - (16# * VBA.Int(workValue / 16#)))
        outHex = VBA.Mid$(hexChars, digitValue + 1, 1) & outHex
        workValue = VBA.Int(workValue / 16#)
    Next i

    private_ComputeStableHexHash = outHex
End Function

Private Function private_GetRuntimeCallbackMacroRef() As String
    private_GetRuntimeCallbackMacroRef = private_QualifyMacroName("rt_Bridge.fn_OnShapeClick")
End Function

Private Function private_TryResolveCallbackRef( _
    ByVal rawText As String, _
    ByVal dataContext As Object, _
    ByRef outCallbackRef As String _
) As Boolean
    Dim resolvedValue As Variant

    outCallbackRef = VBA.vbNullString
    rawText = VBA.Trim$(rawText)
    If VBA.Len(rawText) = 0 Then
        private_TryResolveCallbackRef = True
        Exit Function
    End If

    If Not ex_BindingRuntime.fn_TryResolveValueBinding(rawText, dataContext, resolvedValue) Then Exit Function
    If VBA.IsObject(resolvedValue) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Select: callback binding must resolve to scalar value for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    outCallbackRef = VBA.Trim$(VBA.CStr(resolvedValue))
    If VBA.Len(outCallbackRef) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Select: callback binding resolved to empty value for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    private_TryResolveCallbackRef = True
End Function

Private Function private_TryRestoreCallbackContextFromPage() As Boolean
    Dim callbackContext As Object

    Set m_CallbackContext = Nothing
    If Not m_Page.TryGetController(callbackContext) Then Exit Function
    If callbackContext Is Nothing Then
        private_TryRestoreCallbackContextFromPage = True
        Exit Function
    End If

    Set m_CallbackContext = callbackContext
    private_TryRestoreCallbackContextFromPage = True
End Function

Private Function private_QualifyMacroName(ByVal macroName As String) As String
    Dim wbName As String

    macroName = VBA.Trim$(macroName)
    If VBA.Len(macroName) = 0 Then Exit Function
    If VBA.InStr(1, macroName, "!", VBA.vbBinaryCompare) > 0 Then
        private_QualifyMacroName = macroName
        Exit Function
    End If

    wbName = ThisWorkbook.Name
    wbName = VBA.Replace$(wbName, "'", "''")
    private_QualifyMacroName = "'" & wbName & "'!" & macroName
End Function

Private Function private_TryReadPositiveDoubleAttr( _
    ByVal controlNode As Object, _
    ByVal attrName As String, _
    ByVal defaultValue As Double, _
    ByRef outValue As Double _
) As Boolean
    Dim rawText As String

    rawText = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, attrName)))
    If VBA.Len(rawText) = 0 Then
        outValue = defaultValue
        private_TryReadPositiveDoubleAttr = True
        Exit Function
    End If

    If Not private_TryParseFlexibleDouble(rawText, outValue) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Select: attribute '" & attrName & "' must be numeric for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    If outValue <= 0# Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Select: attribute '" & attrName & "' must be greater than zero for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    private_TryReadPositiveDoubleAttr = True
End Function

Private Function private_TryReadNonNegativeDoubleAttr( _
    ByVal controlNode As Object, _
    ByVal attrName As String, _
    ByVal defaultValue As Double, _
    ByRef outValue As Double _
) As Boolean
    Dim rawText As String

    rawText = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, attrName)))
    If VBA.Len(rawText) = 0 Then
        outValue = defaultValue
        private_TryReadNonNegativeDoubleAttr = True
        Exit Function
    End If

    If Not private_TryParseFlexibleDouble(rawText, outValue) Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Select: attribute '" & attrName & "' must be numeric for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    If outValue < 0# Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "Select: attribute '" & attrName & "' must be greater or equal to zero for control '" & m_ControlName & "'."
#End If
        Exit Function
    End If

    private_TryReadNonNegativeDoubleAttr = True
End Function

Private Function private_TryParseFlexibleDouble(ByVal rawText As String, ByRef outValue As Double) As Boolean
    Dim normalized As String
    Dim decimalSep As String

    rawText = VBA.Trim$(rawText)
    If VBA.Len(rawText) = 0 Then Exit Function

    decimalSep = VBA.CStr(Application.International(xlDecimalSeparator))
    normalized = rawText

    If decimalSep = "," Then
        normalized = VBA.Replace$(normalized, ".", ",")
    Else
        normalized = VBA.Replace$(normalized, ",", ".")
    End If

    If Not VBA.IsNumeric(normalized) Then Exit Function
    outValue = VBA.CDbl(normalized)
    private_TryParseFlexibleDouble = True
End Function

Private Function private_GetUiShapeByName(ByVal ws As Worksheet, ByVal shapeName As String) As Shape
    If ws Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(shapeName)) = 0 Then Exit Function

    On Error Resume Next
    Set private_GetUiShapeByName = ws.Shapes(shapeName)
    On Error GoTo 0
End Function
