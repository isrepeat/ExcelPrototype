VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_SelectControlVM"
Option Explicit
Implements obj_IControl
Implements obj_ISerializable

Private Const DEFAULT_PLACEHOLDER As String = "Choose option"
Private Const DEFAULT_ITEM_HEIGHT As Double = 18#
Private Const DEFAULT_ITEM_MARGIN As Double = 2#

Private m_Base As obj_ControlBase
' Общие layout-параметры контрола (лист, границы в ячейках, style).
Private m_Layout As obj_ControlLayout
Private m_ControlName As String
' Значения, считанные из XML (raw-конфиг).
Private m_ItemsSourceRaw As String
Private m_PlaceholderText As String
Private m_OnChangeRaw As String
Private m_OnChangeMacroRef As String
Private m_SelectedIdRaw As String
Private m_SelectedItemSourceRaw As String
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

' Runtime-буферы уже отрисованного select.
' Используются при кликах через page-owned runtime карты (без повторного configure/render).
Private m_RuntimeHeaderShapeName As String
Private m_RuntimePanelShapeName As String
Private m_RuntimeItemShapeNames As Collection
Private m_RuntimeItemCaptions As Collection
Private m_RuntimeItemIds As Collection
Private m_RuntimeItemActionMacros As Collection
Private m_RuntimeItemRawItems As Collection
Private m_RuntimeIsOpen As Boolean
Private m_IsConfigured As Boolean
Private m_Page As obj_PageBase

' //
' // Interface
' //
Private Sub obj_IControl_Configure(ByVal page As obj_PageBase, ByVal controlNode As Object)
    Dim currentPage As obj_PageBase
    Dim selectedIdText As String
    Dim dataContext As Object

    ' Полный reset состояния: важно при повторной конфигурации того же VM.
    m_IsConfigured = False
    Set m_Layout = Nothing
    Set m_Items = Nothing
    Set m_ItemCaptions = Nothing
    Set m_ItemIds = Nothing
    Set m_ItemActionMacros = Nothing
    Set m_ItemRawItems = Nothing
    m_RuntimeHeaderShapeName = VBA.vbNullString
    m_RuntimePanelShapeName = VBA.vbNullString
    Set m_RuntimeItemShapeNames = Nothing
    Set m_RuntimeItemCaptions = Nothing
    Set m_RuntimeItemIds = Nothing
    Set m_RuntimeItemActionMacros = Nothing
    Set m_RuntimeItemRawItems = Nothing
    Set m_Base = Nothing
    Set m_Page = Nothing
    m_RuntimeIsOpen = False
    m_SelectedIndex = 0

    Set m_Base = New obj_ControlBase
    If Not m_Base.Configure(page, controlNode, "Select", "select", m_ControlName) Then Exit Sub
    Set m_Page = m_Base.PageBase

    m_ItemsSourceRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "itemsSource")))
    If VBA.Len(m_ItemsSourceRaw) = 0 Then
        VBA.MsgBox "Select: itemsSource is not specified for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    m_SelectedIdRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "selectedId")))
    m_SelectedItemSourceRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "selectedItem")))
    If VBA.Len(m_SelectedItemSourceRaw) = 0 Then
        m_SelectedItemSourceRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "selectedItemSource")))
    End If

    m_PlaceholderText = VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "placeholder"))
    If VBA.Len(VBA.Trim$(m_PlaceholderText)) = 0 Then m_PlaceholderText = DEFAULT_PLACEHOLDER

    m_ItemStyleName = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "itemStyle")))
    m_PanelStyleName = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "panelStyle")))

    If Not private_TryReadPositiveDoubleAttr(controlNode, "itemHeight", DEFAULT_ITEM_HEIGHT, m_ItemHeight) Then Exit Sub
    If Not private_TryReadNonNegativeDoubleAttr(controlNode, "itemMargin", DEFAULT_ITEM_MARGIN, m_ItemMargin) Then Exit Sub

    m_OnChangeRaw = VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, "onChange"))
    m_OnChangeMacroRef = VBA.vbNullString
    If VBA.Len(VBA.Trim$(m_OnChangeRaw)) > 0 Then
        Set dataContext = m_Base.DataContext
        If dataContext Is Nothing Then Set dataContext = Me
        If Not ex_BindingRuntime.m_TryResolveMacroBinding(m_OnChangeRaw, dataContext, m_OnChangeMacroRef) Then Exit Sub
    End If

    ' 2) Читаем общий layout (лист + границы + style).
    Set m_Layout = New obj_ControlLayout
    If Not m_Layout.TryReadFromNode(controlNode, "Select", m_ControlName, "style") Then Exit Sub

    ' 3) Разрешаем itemsSource в runtime-коллекцию и готовим буферы.
    Set currentPage = m_Base.PageBase
    If currentPage Is Nothing Then Exit Sub
    If Not ex_RuntimeSourceResolver.m_TryResolveItemsSource(currentPage.RuntimeSources, m_ItemsSourceRaw, m_Items) Then Exit Sub
    If Not private_TryBuildItemBuffers() Then Exit Sub

    ' 4) Определяем начальный выбранный элемент:
    '    selectedId из XML -> state store -> fallback на первый item.
    m_SelectStateKey = VBA.LCase$(m_Layout.LayoutSheetName & "|" & m_ControlName)
    If Not private_TryResolveSelectedIdText(selectedIdText) Then Exit Sub
    m_SelectedIndex = private_FindSelectedIndexById(selectedIdText)

    If m_SelectedIndex = 0 And Not m_ItemIds Is Nothing Then
        If m_ItemIds.Count > 0 Then m_SelectedIndex = 1
    End If

    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render()
    Dim ws As Worksheet
    Dim headerRange As Range
    Dim headerShape As Shape
    Dim panelShape As Shape
    Dim itemShape As Shape
    Dim itemShapes As Collection
    Dim itemCaptions As Collection
    Dim itemIds As Collection
    Dim itemActions As Collection
    Dim itemRawItems As Collection
    Dim callbackMacroRef As String
    Dim renderItemCount As Long
    Dim i As Long
    Dim selectedIndexRendered As Long
    Dim headerLeft As Double
    Dim headerTop As Double
    Dim headerWidth As Double
    Dim headerHeight As Double
    Dim panelLeft As Double
    Dim panelTop As Double
    Dim panelWidth As Double
    Dim panelHeight As Double
    Dim itemLeft As Double
    Dim itemTop As Double
    Dim itemWidth As Double
    Dim page As obj_PageBase

    If Not m_IsConfigured Then
        VBA.MsgBox "Select: control '" & m_ControlName & "' is not configured.", VBA.vbExclamation
        Exit Sub
    End If

    Set page = Nothing
    If Not m_Base Is Nothing Then Set page = m_Base.PageBase
    If page Is Nothing Then Set page = m_Page
    If page Is Nothing Then
        VBA.MsgBox "Select: page is not specified for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If
    Set m_Page = page

    Set ws = page.Worksheet
    If ws Is Nothing Then
        VBA.MsgBox "Select: page worksheet is not specified for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    If m_ItemCaptions Is Nothing Or m_ItemIds Is Nothing Or m_ItemActionMacros Is Nothing Or m_ItemRawItems Is Nothing Then
        VBA.MsgBox "Select: item metadata is not configured for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Sub
    End If

    ' Shape.OnAction вызывает только модульный макрос,
    ' поэтому указываем стабильный bridge rt_Bridge.m_OnShapeClick.
    callbackMacroRef = private_GetRuntimeCallbackMacroRef()
    If VBA.Len(callbackMacroRef) = 0 Then Exit Sub

    If Not private_TryBuildHeaderRange(ws, headerRange) Then Exit Sub
    renderItemCount = private_GetRenderItemCount()

    ' Перерисовка "с нуля": удаляем предыдущие shape контрола.
    private_DeleteControlShapes ws

    headerLeft = headerRange.Left
    headerTop = headerRange.Top
    headerWidth = headerRange.Width
    headerHeight = headerRange.Height

    panelLeft = headerLeft
    panelTop = headerTop + headerHeight
    panelWidth = headerWidth
    panelHeight = private_CalcPanelHeight(renderItemCount)

    Set headerShape = private_CreateShapeByRange(ws, headerRange, "header", callbackMacroRef)
    If headerShape Is Nothing Then Exit Sub
    private_ApplyHeaderVisualDefaults headerShape

    If renderItemCount > 0 Then
        Set panelShape = private_CreateShapeByBounds(ws, panelLeft, panelTop, panelWidth, panelHeight, "panel", VBA.vbNullString)
        If panelShape Is Nothing Then Exit Sub
        private_ApplyPanelVisualDefaults panelShape
    Else
        Set panelShape = private_CreateShapeByBounds(ws, headerLeft, headerTop, headerWidth, headerHeight, "panel", VBA.vbNullString)
        If panelShape Is Nothing Then Exit Sub
        panelShape.Visible = msoFalse
    End If

    Set itemShapes = New Collection
    Set itemCaptions = New Collection
    Set itemIds = New Collection
    Set itemActions = New Collection
    Set itemRawItems = New Collection

    itemLeft = headerLeft
    itemWidth = headerWidth
    For i = 1 To renderItemCount
        itemTop = panelTop + VBA.CDbl(i - 1) * (m_ItemHeight + m_ItemMargin)
        Set itemShape = private_CreateShapeByBounds(ws, itemLeft, itemTop, itemWidth, m_ItemHeight, "item" & VBA.CStr(i), callbackMacroRef)
        If itemShape Is Nothing Then Exit Sub
        private_SetShapeText itemShape, VBA.CStr(m_ItemCaptions(i))
        private_ApplyItemVisualDefaults itemShape

        itemShapes.Add itemShape.Name
        itemCaptions.Add VBA.CStr(m_ItemCaptions(i))
        itemIds.Add VBA.CStr(m_ItemIds(i))
        itemActions.Add VBA.CStr(m_ItemActionMacros(i))
        itemRawItems.Add m_ItemRawItems(i)
    Next i

    selectedIndexRendered = m_SelectedIndex
    If selectedIndexRendered <= 0 Or selectedIndexRendered > renderItemCount Then selectedIndexRendered = 0

    ' Синхронизируем runtime-буферы VM с только что созданными shape.
    If Not private_InitializeRuntimeState( _
        headerShapeName:=headerShape.Name, _
        panelShapeName:=panelShape.Name, _
        itemShapeNames:=itemShapes, _
        itemCaptions:=itemCaptions, _
        itemIds:=itemIds, _
        itemActionMacros:=itemActions, _
        itemRawItems:=itemRawItems, _
        selectedIndex:=selectedIndexRendered) Then Exit Sub

    If Not private_TryBindRuntimeRoutes(ws, headerShape.Name, itemShapes) Then Exit Sub
End Sub

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(attrName))
        Case "itemssource", "placeholder", "onchange", "selectedid", "selecteditem", "selecteditemsource", _
             "style", _
             "itemstyle", "panelstyle", "itemheight", "itemmargin"
            obj_IControl_SupportsAttribute = True
    End Select
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

' //
' // API
' //
Public Function GetSelectedId() As String
    If Not m_RuntimeItemIds Is Nothing Then
        If m_SelectedIndex <= 0 Or m_SelectedIndex > m_RuntimeItemIds.Count Then Exit Function
        GetSelectedId = VBA.CStr(m_RuntimeItemIds(m_SelectedIndex))
        Exit Function
    End If

    If m_ItemIds Is Nothing Then Exit Function
    If m_SelectedIndex <= 0 Or m_SelectedIndex > m_ItemIds.Count Then Exit Function
    GetSelectedId = VBA.CStr(m_ItemIds(m_SelectedIndex))
End Function

Public Function GetSelectedCaption() As String
    If Not m_RuntimeItemCaptions Is Nothing Then
        If m_SelectedIndex <= 0 Or m_SelectedIndex > m_RuntimeItemCaptions.Count Then Exit Function
        GetSelectedCaption = VBA.CStr(m_RuntimeItemCaptions(m_SelectedIndex))
        Exit Function
    End If

    If m_ItemCaptions Is Nothing Then Exit Function
    If m_SelectedIndex <= 0 Or m_SelectedIndex > m_ItemCaptions.Count Then Exit Function
    GetSelectedCaption = VBA.CStr(m_ItemCaptions(m_SelectedIndex))
End Function

Public Function GetControlKey() As String
    GetControlKey = VBA.CStr(m_ControlName)
End Function

Public Function RuntimeHasSelectedItem() As Boolean
    If m_RuntimeItemCaptions Is Nothing Then Exit Function
    RuntimeHasSelectedItem = (m_SelectedIndex > 0 And m_SelectedIndex <= m_RuntimeItemCaptions.Count)
End Function

Public Function RuntimeGetSelectedIndex() As Long
    RuntimeGetSelectedIndex = m_SelectedIndex
End Function

Public Function RuntimeGetSelectedCaption() As String
    RuntimeGetSelectedCaption = Me.GetSelectedCaption()
End Function

Public Function RuntimeGetSelectedId() As String
    RuntimeGetSelectedId = Me.GetSelectedId()
End Function

Public Function RuntimeHandleHeaderClick() As Boolean
    ' Клик по header только переключает open/close dropdown panel.
    m_RuntimeIsOpen = (Not m_RuntimeIsOpen)
    RuntimeHandleHeaderClick = private_ApplyRuntimeVisualState()
End Function

Public Function RuntimeHandleItemClick(ByVal itemIndex As Long) As Boolean
    Dim selectedId As String
    Dim itemMacro As String

    If m_RuntimeItemCaptions Is Nothing Then Exit Function
    If itemIndex <= 0 Or itemIndex > m_RuntimeItemCaptions.Count Then
        VBA.MsgBox "Select: selected item index is out of range for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Function
    End If

    m_SelectedIndex = itemIndex
    m_RuntimeIsOpen = False

    If Not private_ApplyRuntimeVisualState() Then Exit Function

    ' При выборе item:
    ' 1) фиксируем selectedId в CustomXMLPart
    ' 2) публикуем selectedItem в objectSource
    ' 3) запускаем item macro и onChange.
    selectedId = Me.GetSelectedId()
    If Not private_TryPersistSelectedId(selectedId) Then Exit Function
    If Not private_TryPublishRuntimeSelectedItem(itemIndex) Then Exit Function

    itemMacro = private_GetRuntimeCollectionText(m_RuntimeItemActionMacros, itemIndex)
    If VBA.Len(itemMacro) > 0 Then
        If Not private_RunRuntimeMacro(itemMacro) Then Exit Function
    End If

    If VBA.Len(m_OnChangeMacroRef) > 0 Then
        If Not private_RunRuntimeMacro(m_OnChangeMacroRef) Then Exit Function
    End If

    RuntimeHandleItemClick = True
End Function

Public Function RuntimeCloseDropdown() As Boolean
    ' Явное закрытие dropdown (используется dispatcher-ом при клике по другим shape).
    If Not m_RuntimeIsOpen Then
        RuntimeCloseDropdown = True
        Exit Function
    End If

    m_RuntimeIsOpen = False
    RuntimeCloseDropdown = private_ApplyRuntimeVisualState()
End Function

Public Function RuntimeOnGlobalClick(ByVal clickedControlKey As String) As Boolean
    clickedControlKey = VBA.LCase$(VBA.Trim$(clickedControlKey))

    If clickedControlKey = VBA.LCase$(VBA.Trim$(m_SelectStateKey)) Then
        RuntimeOnGlobalClick = True
        Exit Function
    End If

    RuntimeOnGlobalClick = Me.RuntimeCloseDropdown()
End Function

Public Function TrySerializeSnapshot(ByRef outSnapshotXml As String) As Boolean
    Dim i As Long
    Dim selectedIndexText As String
    Dim runtimeIsOpenText As String

    outSnapshotXml = VBA.vbNullString

    If m_Layout Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(m_ControlName)) = 0 Then Exit Function
    If VBA.Len(VBA.Trim$(m_RuntimeHeaderShapeName)) = 0 Then Exit Function
    If VBA.Len(VBA.Trim$(m_RuntimePanelShapeName)) = 0 Then Exit Function
    If m_RuntimeItemShapeNames Is Nothing Then Exit Function
    If m_RuntimeItemCaptions Is Nothing Then Exit Function
    If m_RuntimeItemIds Is Nothing Then Exit Function
    If m_RuntimeItemActionMacros Is Nothing Then Exit Function
    If m_RuntimeItemRawItems Is Nothing Then Exit Function

    selectedIndexText = VBA.CStr(m_SelectedIndex)
    runtimeIsOpenText = VBA.IIf(m_RuntimeIsOpen, "true", "false")

    outSnapshotXml = "<select version=""2"""
    outSnapshotXml = outSnapshotXml & " controlName=""" & ex_Helpers.m_EscapeXmlAttr(m_ControlName) & """"
    outSnapshotXml = outSnapshotXml & " itemsSource=""" & ex_Helpers.m_EscapeXmlAttr(m_ItemsSourceRaw) & """"
    outSnapshotXml = outSnapshotXml & " selectedIdRaw=""" & ex_Helpers.m_EscapeXmlAttr(m_SelectedIdRaw) & """"
    outSnapshotXml = outSnapshotXml & " sheet=""" & ex_Helpers.m_EscapeXmlAttr(m_Layout.LayoutSheetName) & """"
    outSnapshotXml = outSnapshotXml & " rowStart=""" & VBA.CStr(m_Layout.RowStart) & """"
    outSnapshotXml = outSnapshotXml & " colStart=""" & VBA.CStr(m_Layout.ColStart) & """"
    outSnapshotXml = outSnapshotXml & " rowEnd=""" & VBA.CStr(m_Layout.RowEnd) & """"
    outSnapshotXml = outSnapshotXml & " colEnd=""" & VBA.CStr(m_Layout.ColEnd) & """"
    outSnapshotXml = outSnapshotXml & " style=""" & ex_Helpers.m_EscapeXmlAttr(m_Layout.StyleName) & """"
    outSnapshotXml = outSnapshotXml & " selectKey=""" & ex_Helpers.m_EscapeXmlAttr(m_SelectStateKey) & """"
    outSnapshotXml = outSnapshotXml & " placeholder=""" & ex_Helpers.m_EscapeXmlAttr(m_PlaceholderText) & """"
    outSnapshotXml = outSnapshotXml & " onChangeRaw=""" & ex_Helpers.m_EscapeXmlAttr(m_OnChangeRaw) & """"
    outSnapshotXml = outSnapshotXml & " onChange=""" & ex_Helpers.m_EscapeXmlAttr(m_OnChangeMacroRef) & """"
    outSnapshotXml = outSnapshotXml & " selectedItemSource=""" & ex_Helpers.m_EscapeXmlAttr(m_SelectedItemSourceRaw) & """"
    outSnapshotXml = outSnapshotXml & " itemStyle=""" & ex_Helpers.m_EscapeXmlAttr(m_ItemStyleName) & """"
    outSnapshotXml = outSnapshotXml & " panelStyle=""" & ex_Helpers.m_EscapeXmlAttr(m_PanelStyleName) & """"
    outSnapshotXml = outSnapshotXml & " itemHeight=""" & VBA.CStr(m_ItemHeight) & """"
    outSnapshotXml = outSnapshotXml & " itemMargin=""" & VBA.CStr(m_ItemMargin) & """"
    outSnapshotXml = outSnapshotXml & " selectedIndex=""" & ex_Helpers.m_EscapeXmlAttr(selectedIndexText) & """"
    outSnapshotXml = outSnapshotXml & " isOpen=""" & runtimeIsOpenText & """"
    outSnapshotXml = outSnapshotXml & " isConfigured=""" & VBA.IIf(m_IsConfigured, "true", "false") & """"
    outSnapshotXml = outSnapshotXml & ">"
    outSnapshotXml = outSnapshotXml & "<header shape=""" & ex_Helpers.m_EscapeXmlAttr(m_RuntimeHeaderShapeName) & """ />"
    outSnapshotXml = outSnapshotXml & "<panel shape=""" & ex_Helpers.m_EscapeXmlAttr(m_RuntimePanelShapeName) & """ />"

    For i = 1 To m_RuntimeItemShapeNames.Count
        outSnapshotXml = outSnapshotXml & _
            "<item" & _
            " shape=""" & ex_Helpers.m_EscapeXmlAttr(VBA.CStr(m_RuntimeItemShapeNames(i))) & """" & _
            " caption=""" & ex_Helpers.m_EscapeXmlAttr(VBA.CStr(m_RuntimeItemCaptions(i))) & """" & _
            " id=""" & ex_Helpers.m_EscapeXmlAttr(VBA.CStr(m_RuntimeItemIds(i))) & """" & _
            " action=""" & ex_Helpers.m_EscapeXmlAttr(VBA.CStr(m_RuntimeItemActionMacros(i))) & """" & _
                " rawValue=""" & ex_Helpers.m_EscapeXmlAttr(ex_Helpers.m_GetSnapshotRawValueText(m_RuntimeItemRawItems, i, VBA.CStr(m_RuntimeItemIds(i)))) & """" & _
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
    Dim runtimeIsOpen As Boolean
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
    Dim rawObj As Object

    snapshotXml = VBA.Trim$(snapshotXml)
    If VBA.Len(snapshotXml) = 0 Then Exit Function

    If Not ex_Core.m_CustomXmlPartStore_TryLoadDomFromXml(snapshotXml, dom) Then Exit Function
    Set root = dom.DocumentElement
    If root Is Nothing Then Exit Function
    If VBA.LCase$(VBA.CStr(root.baseName)) <> "select" Then Exit Function

    m_ControlName = VBA.Trim$(VBA.CStr(root.getAttribute("controlName")))
    m_ItemsSourceRaw = VBA.Trim$(VBA.CStr(root.getAttribute("itemsSource")))
    m_SelectedIdRaw = VBA.Trim$(VBA.CStr(root.getAttribute("selectedIdRaw")))
    m_SelectStateKey = VBA.LCase$(VBA.Trim$(VBA.CStr(root.getAttribute("selectKey"))))
    m_PlaceholderText = VBA.CStr(root.getAttribute("placeholder"))
    m_OnChangeRaw = VBA.CStr(root.getAttribute("onChangeRaw"))
    m_OnChangeMacroRef = VBA.Trim$(VBA.CStr(root.getAttribute("onChange")))
    m_SelectedItemSourceRaw = VBA.Trim$(VBA.CStr(root.getAttribute("selectedItemSource")))
    m_ItemStyleName = VBA.Trim$(VBA.CStr(root.getAttribute("itemStyle")))
    m_PanelStyleName = VBA.Trim$(VBA.CStr(root.getAttribute("panelStyle")))
    m_ItemHeight = ex_Helpers.m_ReadSnapshotDoubleAttr(root, "itemHeight", DEFAULT_ITEM_HEIGHT)
    m_ItemMargin = ex_Helpers.m_ReadSnapshotDoubleAttr(root, "itemMargin", DEFAULT_ITEM_MARGIN)
    runtimeIsOpen = ex_Helpers.m_ReadSnapshotBooleanAttr(root, "isOpen", False)
    isConfiguredAttr = VBA.LCase$(VBA.Trim$(VBA.CStr(root.getAttribute("isConfigured"))))
    layoutSheetName = VBA.Trim$(VBA.CStr(root.getAttribute("sheet")))
    layoutRowStart = ex_Helpers.m_ReadSnapshotLongAttr(root, "rowStart", 1)
    layoutColStart = ex_Helpers.m_ReadSnapshotLongAttr(root, "colStart", 1)
    layoutRowEnd = ex_Helpers.m_ReadSnapshotLongAttr(root, "rowEnd", layoutRowStart)
    layoutColEnd = ex_Helpers.m_ReadSnapshotLongAttr(root, "colEnd", layoutColStart)
    layoutStyle = VBA.Trim$(VBA.CStr(root.getAttribute("style")))

    If VBA.Len(m_ControlName) = 0 Then Exit Function
    If VBA.Len(m_SelectStateKey) = 0 Then
        m_SelectStateKey = VBA.LCase$(VBA.Trim$(layoutSheetName) & "|" & m_ControlName)
    End If
    If VBA.Len(VBA.Trim$(m_PlaceholderText)) = 0 Then m_PlaceholderText = DEFAULT_PLACEHOLDER
    If m_ItemHeight <= 0# Then m_ItemHeight = DEFAULT_ITEM_HEIGHT
    If m_ItemMargin < 0# Then m_ItemMargin = DEFAULT_ITEM_MARGIN

    Set m_Layout = New obj_ControlLayout
    If Not m_Layout.TryReadFromRuntimeValues( _
        "Select", _
        m_ControlName, _
        layoutSheetName, _
        layoutRowStart, _
        layoutColStart, _
        layoutRowEnd, _
        layoutColEnd, _
        layoutStyle) Then Exit Function

    Set ws = ex_HelpersSheet.m_GetRuntimeWorksheetByName(m_Layout.LayoutSheetName)
    If ws Is Nothing Then Exit Function
    If Not ex_HelpersSheet.m_TryGetPageBaseByWorksheetName(m_Layout.LayoutSheetName, m_Page) Then Exit Function

    Set headerNode = root.selectSingleNode("*[local-name()='header']")
    If headerNode Is Nothing Then Exit Function
    headerShapeName = VBA.Trim$(VBA.CStr(headerNode.getAttribute("shape")))
    If VBA.Len(headerShapeName) = 0 Then Exit Function

    Set panelNode = root.selectSingleNode("*[local-name()='panel']")
    If panelNode Is Nothing Then Exit Function
    panelShapeName = VBA.Trim$(VBA.CStr(panelNode.getAttribute("shape")))
    If VBA.Len(panelShapeName) = 0 Then Exit Function

    Set itemShapeNames = New Collection
    Set itemCaptions = New Collection
    Set itemIds = New Collection
    Set itemActionMacros = New Collection
    Set itemRawItems = New Collection

    Set itemNodes = root.selectNodes("*[local-name()='item']")
    If Not itemNodes Is Nothing Then
        For Each itemNode In itemNodes
            itemShapeNames.Add VBA.CStr(itemNode.getAttribute("shape"))
            itemCaptions.Add VBA.CStr(itemNode.getAttribute("caption"))
            itemIds.Add VBA.CStr(itemNode.getAttribute("id"))
            itemActionMacros.Add VBA.CStr(itemNode.getAttribute("action"))

            Set rawObj = VBA.CreateObject("Scripting.Dictionary")
            rawObj.CompareMode = 1
            rawObj("Id") = VBA.CStr(itemNode.getAttribute("id"))
            rawObj("Caption") = VBA.CStr(itemNode.getAttribute("caption"))
            rawObj("RawValue") = VBA.CStr(itemNode.getAttribute("rawValue"))
            If VBA.Len(VBA.Trim$(VBA.CStr(rawObj("RawValue")))) = 0 Then rawObj("RawValue") = VBA.CStr(itemNode.getAttribute("id"))
            itemRawItems.Add rawObj
        Next itemNode
    End If

    If VBA.IsNumeric(VBA.CStr(root.getAttribute("selectedIndex"))) Then
        selectedIndex = VBA.CLng(root.getAttribute("selectedIndex"))
    Else
        selectedIndex = 0
    End If

    If Not private_InitializeRuntimeState( _
        headerShapeName:=headerShapeName, _
        panelShapeName:=panelShapeName, _
        itemShapeNames:=itemShapeNames, _
        itemCaptions:=itemCaptions, _
        itemIds:=itemIds, _
        itemActionMacros:=itemActionMacros, _
        itemRawItems:=itemRawItems, _
        selectedIndex:=selectedIndex) Then Exit Function

    If Not private_SyncStaticBuffersFromRuntime() Then Exit Function

    m_RuntimeIsOpen = runtimeIsOpen
    If Not private_ApplyRuntimeVisualState() Then Exit Function

    If Not private_TryBindRuntimeRoutes(ws, headerShapeName, itemShapeNames) Then Exit Function

    If isConfiguredAttr = "false" Or isConfiguredAttr = "0" Then
        m_IsConfigured = False
    Else
        m_IsConfigured = True
    End If
    TryDeserializeSnapshot = True
End Function

' //
' // Internal
' //
Private Function private_TryBindRuntimeRoutes( _
    ByVal ws As Worksheet, _
    ByVal headerShapeName As String, _
    ByVal itemShapeNames As Collection _
) As Boolean
    Dim callbackMacroRef As String
    Dim selectId As String
    Dim headerShape As Shape
    Dim itemShape As Shape
    Dim i As Long

    If ws Is Nothing Then
        ex_Core.m_Diagnostic_LogError "select:bind-routes worksheet-missing control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "'"
        Exit Function
    End If
    If itemShapeNames Is Nothing Then
        ex_Core.m_Diagnostic_LogError "select:bind-routes item-shapes-missing control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "'"
        Exit Function
    End If

    callbackMacroRef = private_GetRuntimeCallbackMacroRef()
    If VBA.Len(callbackMacroRef) = 0 Then
        ex_Core.m_Diagnostic_LogError "select:bind-routes callback-macro-empty control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "'"
        Exit Function
    End If

    Set headerShape = private_GetRuntimeShapeByName(ws, headerShapeName)
    If headerShape Is Nothing Then
        ex_Core.m_Diagnostic_LogError "select:bind-routes header-shape-missing control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' shape='" & VBA.Replace$(VBA.Trim$(headerShapeName), "'", "''") & "'"
        Exit Function
    End If
    If Not private_TrySetShapeOnAction(headerShape, callbackMacroRef) Then
        ex_Core.m_Diagnostic_LogError "select:bind-routes header-onaction-failed control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' shape='" & VBA.Replace$(VBA.Trim$(headerShape.Name), "'", "''") & "'"
        Exit Function
    End If

    For i = 1 To itemShapeNames.Count
        Set itemShape = private_GetRuntimeShapeByName(ws, VBA.CStr(itemShapeNames(i)))
        If itemShape Is Nothing Then
            ex_Core.m_Diagnostic_LogError "select:bind-routes item-shape-missing control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' index=" & VBA.CStr(i)
            Exit Function
        End If
        If Not private_TrySetShapeOnAction(itemShape, callbackMacroRef) Then
            ex_Core.m_Diagnostic_LogError "select:bind-routes item-onaction-failed control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' shape='" & VBA.Replace$(VBA.Trim$(itemShape.Name), "'", "''") & "'"
            Exit Function
        End If
    Next i

    selectId = VBA.LCase$(VBA.Trim$(m_SelectStateKey))
    If VBA.Len(selectId) = 0 Then
        ex_Core.m_Diagnostic_LogError "select:bind-routes select-id-empty control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "'"
        Exit Function
    End If

    If m_Page Is Nothing Then
        ex_Core.m_Diagnostic_LogError "select:bind-routes page-base-missing control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "'"
        Exit Function
    End If
    If Not m_Page.RegisterControl(selectId, Me) Then
        ex_Core.m_Diagnostic_LogError "select:bind-routes register-control-failed control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' key='" & VBA.Replace$(selectId, "'", "''") & "'"
        Exit Function
    End If
    If Not m_Page.RegisterShapeRoute(headerShape.Name, selectId, "RuntimeHandleHeaderClick", False) Then
        ex_Core.m_Diagnostic_LogError "select:bind-routes register-header-route-failed control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' shape='" & VBA.Replace$(VBA.Trim$(headerShape.Name), "'", "''") & "'"
        Exit Function
    End If

    For i = 1 To itemShapeNames.Count
        If Not m_Page.RegisterShapeRoute(VBA.CStr(itemShapeNames(i)), selectId, "RuntimeHandleItemClick", True, VBA.CLng(i)) Then
            ex_Core.m_Diagnostic_LogError "select:bind-routes register-item-route-failed control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' index=" & VBA.CStr(i)
            Exit Function
        End If
    Next i

    ex_Core.m_Diagnostic_LogInfo "select:bind-routes ok control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' items=" & VBA.CStr(itemShapeNames.Count) & " macro='" & VBA.Replace$(callbackMacroRef, "'", "''") & "'"
    private_TryBindRuntimeRoutes = True
End Function

Private Function private_TrySetShapeOnAction(ByVal shp As Shape, ByVal callbackMacroRef As String) As Boolean
    If shp Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(callbackMacroRef)) = 0 Then Exit Function

    On Error Resume Next
    shp.OnAction = callbackMacroRef
    If Err.Number <> 0 Then
        ex_Core.m_Diagnostic_LogError "select:set-onaction-failed control='" & VBA.Replace$(VBA.Trim$(m_ControlName), "'", "''") & "' shape='" & VBA.Replace$(VBA.Trim$(shp.Name), "'", "''") & "' err='" & VBA.Replace$(Err.Description, "'", "''") & "'"
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    private_TrySetShapeOnAction = True
End Function

Private Function private_SyncStaticBuffersFromRuntime() As Boolean
    Dim i As Long

    If m_RuntimeItemCaptions Is Nothing Then Exit Function
    If m_RuntimeItemIds Is Nothing Then Exit Function
    If m_RuntimeItemActionMacros Is Nothing Then Exit Function
    If m_RuntimeItemRawItems Is Nothing Then Exit Function

    Set m_Items = New Collection
    Set m_ItemCaptions = New Collection
    Set m_ItemIds = New Collection
    Set m_ItemActionMacros = New Collection
    Set m_ItemRawItems = New Collection

    For i = 1 To m_RuntimeItemCaptions.Count
        m_ItemCaptions.Add VBA.CStr(m_RuntimeItemCaptions(i))
        m_ItemIds.Add VBA.CStr(m_RuntimeItemIds(i))
        m_ItemActionMacros.Add VBA.CStr(m_RuntimeItemActionMacros(i))
        m_ItemRawItems.Add m_RuntimeItemRawItems(i)
        m_Items.Add m_RuntimeItemRawItems(i)
    Next i

    private_SyncStaticBuffersFromRuntime = True
End Function

Private Function private_InitializeRuntimeState( _
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

    ' Runtime-state хранит shape-имена и коллекции, с которыми работает click handler.
    ' Это позволяет обрабатывать клики без повторного рендера.
    If itemShapeNames Is Nothing Or itemCaptions Is Nothing Or itemIds Is Nothing Or itemActionMacros Is Nothing Or itemRawItems Is Nothing Then
        VBA.MsgBox "Select: runtime item metadata collection is not specified for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Function
    End If

    m_RuntimeHeaderShapeName = VBA.CStr(headerShapeName)
    m_RuntimePanelShapeName = VBA.CStr(panelShapeName)
    Set m_RuntimeItemShapeNames = itemShapeNames
    Set m_RuntimeItemCaptions = itemCaptions
    Set m_RuntimeItemIds = itemIds
    Set m_RuntimeItemActionMacros = itemActionMacros
    Set m_RuntimeItemRawItems = itemRawItems
    m_RuntimeIsOpen = False

    m_SelectedIndex = selectedIndex
    If m_SelectedIndex < 0 Then m_SelectedIndex = 0
    If m_SelectedIndex > m_RuntimeItemCaptions.Count Then m_SelectedIndex = 0

    ' Первичная отрисовка состояния (текст header + видимость panel/items).
    If Not private_ApplyRuntimeVisualState() Then Exit Function

    If Me.RuntimeHasSelectedItem() Then
        ' Сразу синхронизируем persistent state и selectedItem-source.
        selectedId = Me.GetSelectedId()
        If Not private_TryPersistSelectedId(selectedId) Then Exit Function
        If Not private_TryPublishRuntimeSelectedItem(m_SelectedIndex) Then Exit Function
    End If

    private_InitializeRuntimeState = True
End Function

Private Function private_ApplyRuntimeVisualState() As Boolean
    Dim ws As Worksheet
    Dim headerShape As Shape
    Dim panelShape As Shape
    Dim itemShape As Shape
    Dim i As Long
    Dim headerText As String

    If VBA.Len(m_RuntimeHeaderShapeName) = 0 Then Exit Function
    If m_RuntimeItemShapeNames Is Nothing Then Exit Function
    If m_RuntimeItemCaptions Is Nothing Then Exit Function

    Set ws = ex_HelpersSheet.m_GetRuntimeWorksheetByName(m_Layout.LayoutSheetName)
    If ws Is Nothing Then Exit Function

    Set headerShape = private_GetRuntimeShapeByName(ws, m_RuntimeHeaderShapeName)
    If headerShape Is Nothing Then Exit Function

    Set panelShape = private_GetRuntimeShapeByName(ws, m_RuntimePanelShapeName)

    ' Header показывает либо выбранный caption, либо placeholder.
    If Me.RuntimeHasSelectedItem() Then
        headerText = Me.GetSelectedCaption()
    Else
        headerText = m_PlaceholderText
    End If

    If m_RuntimeItemShapeNames.Count = 0 Then
        m_RuntimeIsOpen = False
    End If

    If m_RuntimeIsOpen Then
        headerText = headerText & " ^"
    Else
        headerText = headerText & " v"
    End If

    private_SetShapeText headerShape, headerText

    ' Open/close panel + item-shapes.
    If Not panelShape Is Nothing Then
        panelShape.Visible = VBA.IIf(m_RuntimeIsOpen, msoTrue, msoFalse)
        If m_RuntimeIsOpen Then panelShape.ZOrder msoBringToFront
    End If

    For i = 1 To m_RuntimeItemShapeNames.Count
        Set itemShape = private_GetRuntimeShapeByName(ws, VBA.CStr(m_RuntimeItemShapeNames(i)))
        If itemShape Is Nothing Then GoTo ContinueItem

        itemShape.Visible = VBA.IIf(m_RuntimeIsOpen, msoTrue, msoFalse)
        If m_RuntimeIsOpen Then itemShape.ZOrder msoBringToFront

        ' Выделяем выбранный item через Bold.
        On Error Resume Next
        itemShape.TextFrame.Characters.Font.Bold = (i = m_SelectedIndex)
        itemShape.TextFrame2.TextRange.Font.Bold = (i = m_SelectedIndex)
        On Error GoTo 0

ContinueItem:
    Next i

    private_ApplyRuntimeVisualState = True
End Function

Private Function private_TryPublishRuntimeSelectedItem(ByVal itemIndex As Long) As Boolean
    Dim selectedObject As Object
    Dim selectedRaw As Variant

    If VBA.Len(m_SelectedItemSourceRaw) = 0 Then
        private_TryPublishRuntimeSelectedItem = True
        Exit Function
    End If

    If m_RuntimeItemRawItems Is Nothing Then Exit Function
    If itemIndex <= 0 Or itemIndex > m_RuntimeItemRawItems.Count Then Exit Function

    ' Пытаемся отдать исходный объект как selectedItem.
    Set selectedObject = Nothing
    On Error Resume Next
    Set selectedObject = m_RuntimeItemRawItems(itemIndex)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    ' Для scalar item создаем dictionary-обертку.
    If selectedObject Is Nothing Then
        selectedRaw = m_RuntimeItemRawItems(itemIndex)

        Set selectedObject = VBA.CreateObject("Scripting.Dictionary")
        selectedObject.CompareMode = 1
        selectedObject("Id") = VBA.CStr(m_RuntimeItemIds(itemIndex))
        selectedObject("Caption") = VBA.CStr(m_RuntimeItemCaptions(itemIndex))
        selectedObject("RawValue") = VBA.CStr(selectedRaw)
    End If

    If m_Page Is Nothing Then Exit Function
    If Not m_Page.RuntimeSources.SetObjectSource(m_SelectedItemSourceRaw, selectedObject) Then Exit Function
    private_TryPublishRuntimeSelectedItem = True
End Function

Private Function private_GetRuntimeCollectionText(ByVal values As Collection, ByVal idx As Long) As String
    If values Is Nothing Then Exit Function
    If idx <= 0 Or idx > values.Count Then Exit Function
    private_GetRuntimeCollectionText = VBA.Trim$(VBA.CStr(values(idx)))
End Function

Private Function private_RunRuntimeMacro(ByVal macroRef As String) As Boolean
    macroRef = VBA.Trim$(macroRef)
    If VBA.Len(macroRef) = 0 Then
        private_RunRuntimeMacro = True
        Exit Function
    End If

    On Error GoTo EH_RUN
    Application.Run macroRef
    private_RunRuntimeMacro = True
    Exit Function

EH_RUN:
    VBA.MsgBox "Select: failed to execute macro '" & macroRef & "' for control '" & m_ControlName & "': " & Err.Description, VBA.vbExclamation
End Function

Private Function private_TryBuildItemBuffers() As Boolean
    Dim itemRaw As Variant
    Dim itemCaption As String
    Dim itemId As String
    Dim itemAction As String

    If m_Items Is Nothing Then
        VBA.MsgBox "Select: itemsSource resolved to Nothing for control '" & m_ControlName & "'.", VBA.vbExclamation
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
        VBA.MsgBox "Select: item Id is empty in control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Function
    End If
    If VBA.Len(VBA.Trim$(outCaption)) = 0 Then outCaption = outId

    outItemActionMacro = VBA.vbNullString
    If VBA.Len(actionRaw) > 0 Then
        outItemActionMacro = private_QualifyMacroName(actionRaw)
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
            VBA.MsgBox "Select: item object is Nothing while reading member '" & memberName & "'.", VBA.vbExclamation
            Exit Function
        End If
        private_TryReadObjectMemberText = True
        Exit Function
    End If

    Set dictObj = private_AsDictionary(sourceObject)
    If Not dictObj Is Nothing Then
        If Not dictObj.Exists(memberName) Then
            If isRequired Then
                VBA.MsgBox "Select: member '" & memberName & "' was not found on dictionary item.", VBA.vbExclamation
                Exit Function
            End If

            private_TryReadObjectMemberText = True
            Exit Function
        End If

        scalarValue = dictObj.Item(memberName)
        If VBA.IsObject(scalarValue) Then
            VBA.MsgBox "Select: member '" & memberName & "' must resolve to scalar value.", VBA.vbExclamation
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
            VBA.MsgBox "Select: member '" & memberName & "' was not found on object '" & VBA.TypeName(sourceObject) & "'.", VBA.vbExclamation
            Exit Function
        End If

        private_TryReadObjectMemberText = True
        Exit Function
    End If
    On Error GoTo 0

    If VBA.IsObject(scalarValue) Then
        VBA.MsgBox "Select: member '" & memberName & "' on object '" & VBA.TypeName(sourceObject) & "' must resolve to scalar value.", VBA.vbExclamation
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
    ' Порядок получения selectedId:
    ' 1) selectedId в XML
    ' 2) сохраненное состояние в CustomXMLPart (obj_SelectControlVMStatic)
    outSelectedIdText = VBA.Trim$(m_SelectedIdRaw)
    If VBA.Len(outSelectedIdText) > 0 Then
        private_TryResolveSelectedIdText = True
        Exit Function
    End If

    If Not private_TryLoadStoredSelectedId(outSelectedIdText) Then Exit Function
    outSelectedIdText = VBA.Trim$(outSelectedIdText)
    private_TryResolveSelectedIdText = True
End Function

Private Function private_TryPersistSelectedId(ByVal selectedId As String) As Boolean
    Dim selectStatic As obj_SelectControlVMStatic

    selectedId = VBA.Trim$(selectedId)
    If VBA.Len(VBA.Trim$(m_SelectStateKey)) = 0 Then
        VBA.MsgBox "Select: state key is empty for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Function
    End If

    Set selectStatic = New obj_SelectControlVMStatic
    private_TryPersistSelectedId = selectStatic.SetSelectedId(m_SelectStateKey, selectedId)
End Function

Private Function private_TryLoadStoredSelectedId(ByRef outSelectedId As String) As Boolean
    Dim selectStatic As obj_SelectControlVMStatic

    If VBA.Len(VBA.Trim$(m_SelectStateKey)) = 0 Then
        VBA.MsgBox "Select: state key is empty for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Function
    End If

    Set selectStatic = New obj_SelectControlVMStatic
    private_TryLoadStoredSelectedId = selectStatic.TryGetSelectedId(m_SelectStateKey, outSelectedId)
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
    Set outRange = ws.Range(ws.Cells(m_Layout.RowStart, m_Layout.ColStart), ws.Cells(m_Layout.RowEnd, m_Layout.ColEnd))
    private_TryBuildHeaderRange = True
    Exit Function

EH_RANGE:
    VBA.MsgBox "Select: failed to resolve header range for control '" & m_ControlName & "'.", VBA.vbExclamation
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
        VBA.MsgBox "Select: target range has non-positive width/height for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Function
    End If

    shapeName = private_BuildShapeName(suffix)

    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo EH_SHAPE

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, targetRange.Left, targetRange.Top, targetRange.Width, targetRange.Height)
    shp.Name = shapeName

    If VBA.Len(VBA.Trim$(onActionMacroRef)) > 0 Then
        On Error Resume Next
        shp.OnAction = onActionMacroRef
        If Err.Number <> 0 Then
            VBA.MsgBox "Select: failed to bind click action for shape '" & shapeName & "' in control '" & m_ControlName & "'.", VBA.vbExclamation
            Err.Clear
        End If
        On Error GoTo EH_SHAPE
    End If

    shapeRole = VBA.LCase$(VBA.Trim$(suffix))
    shapeStyle = VBA.vbNullString
    Select Case shapeRole
        Case "header"
            shapeStyle = VBA.Trim$(m_Layout.StyleName)
        Case "panel"
            shapeStyle = VBA.Trim$(m_PanelStyleName)
        Case Else
            shapeStyle = VBA.Trim$(m_ItemStyleName)
    End Select

    On Error Resume Next
    If shapeRole = "header" Then
        shp.Placement = xlMoveAndSize
    Else
        shp.Placement = xlFreeFloating
    End If
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
    If Not ex_ShapeMetaRuntime.m_TrySetShapeMetaValues(shp, metaMap) Then Exit Function

    Set private_CreateShapeByRange = shp
    Exit Function

EH_SHAPE:
    VBA.MsgBox "Select: failed to create shape '" & shapeName & "' for control '" & m_ControlName & "': " & Err.Description, VBA.vbExclamation
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
        VBA.MsgBox "Select: target bounds have non-positive width/height for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Function
    End If

    shapeName = private_BuildShapeName(suffix)

    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo EH_SHAPE

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, shapeLeft, shapeTop, shapeWidth, shapeHeight)
    shp.Name = shapeName

    If VBA.Len(VBA.Trim$(onActionMacroRef)) > 0 Then
        On Error Resume Next
        shp.OnAction = onActionMacroRef
        If Err.Number <> 0 Then
            VBA.MsgBox "Select: failed to bind click action for shape '" & shapeName & "' in control '" & m_ControlName & "'.", VBA.vbExclamation
            Err.Clear
        End If
        On Error GoTo EH_SHAPE
    End If

    shapeRole = VBA.LCase$(VBA.Trim$(suffix))
    shapeStyle = VBA.vbNullString
    Select Case shapeRole
        Case "header"
            shapeStyle = VBA.Trim$(m_Layout.StyleName)
        Case "panel"
            shapeStyle = VBA.Trim$(m_PanelStyleName)
        Case Else
            shapeStyle = VBA.Trim$(m_ItemStyleName)
    End Select

    On Error Resume Next
    If shapeRole = "header" Then
        shp.Placement = xlMoveAndSize
    Else
        shp.Placement = xlFreeFloating
    End If
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
    If Not ex_ShapeMetaRuntime.m_TrySetShapeMetaValues(shp, metaMap) Then Exit Function

    Set private_CreateShapeByBounds = shp
    Exit Function

EH_SHAPE:
    VBA.MsgBox "Select: failed to create floating shape '" & shapeName & "' for control '" & m_ControlName & "': " & Err.Description, VBA.vbExclamation
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

    If ws Is Nothing Then Exit Sub

    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        controlMeta = VBA.LCase$(VBA.Trim$(ex_ShapeMetaRuntime.m_GetShapeMetaValue(shp, "pn.control", VBA.vbNullString)))
        If VBA.Len(controlMeta) = 0 Then GoTo ContinueShape
        If controlMeta = VBA.LCase$(VBA.Trim$(m_ControlName)) Then
            shp.Delete
        End If
ContinueShape:
    Next i
End Sub

Private Function private_BuildShapeName(ByVal suffix As String) As String
    private_BuildShapeName = "sel_" & private_NormalizeNamePart(m_ControlName) & "_" & private_NormalizeNamePart(suffix)
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

Private Function private_GetRuntimeCallbackMacroRef() As String
    private_GetRuntimeCallbackMacroRef = private_QualifyMacroName("rt_Bridge.m_OnShapeClick")
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

    rawText = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, attrName)))
    If VBA.Len(rawText) = 0 Then
        outValue = defaultValue
        private_TryReadPositiveDoubleAttr = True
        Exit Function
    End If

    If Not private_TryParseFlexibleDouble(rawText, outValue) Then
        VBA.MsgBox "Select: attribute '" & attrName & "' must be numeric for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Function
    End If

    If outValue <= 0# Then
        VBA.MsgBox "Select: attribute '" & attrName & "' must be greater than zero for control '" & m_ControlName & "'.", VBA.vbExclamation
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

    rawText = VBA.Trim$(VBA.CStr(ex_XmlCore.m_NodeAttrText(controlNode, attrName)))
    If VBA.Len(rawText) = 0 Then
        outValue = defaultValue
        private_TryReadNonNegativeDoubleAttr = True
        Exit Function
    End If

    If Not private_TryParseFlexibleDouble(rawText, outValue) Then
        VBA.MsgBox "Select: attribute '" & attrName & "' must be numeric for control '" & m_ControlName & "'.", VBA.vbExclamation
        Exit Function
    End If

    If outValue < 0# Then
        VBA.MsgBox "Select: attribute '" & attrName & "' must be greater or equal to zero for control '" & m_ControlName & "'.", VBA.vbExclamation
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

Private Function private_GetRuntimeShapeByName(ByVal ws As Worksheet, ByVal shapeName As String) As Shape
    If ws Is Nothing Then Exit Function
    If VBA.Len(VBA.Trim$(shapeName)) = 0 Then Exit Function

    On Error Resume Next
    Set private_GetRuntimeShapeByName = ws.Shapes(shapeName)
    On Error GoTo 0
End Function
