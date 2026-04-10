VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_SelectControlVM"
Option Explicit
Implements obj_IControl

Private Const DEFAULT_PLACEHOLDER As String = "Choose option"
Private Const DEFAULT_ITEM_HEIGHT As Double = 18#
Private Const DEFAULT_ITEM_MARGIN As Double = 2#

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
' Используются при кликах через ex_ShapeClickDispatcher (без повторного configure/render).
Private m_RuntimeHeaderShapeName As String
Private m_RuntimePanelShapeName As String
Private m_RuntimeItemShapeNames As Collection
Private m_RuntimeItemCaptions As Collection
Private m_RuntimeItemIds As Collection
Private m_RuntimeItemActionMacros As Collection
Private m_RuntimeItemRawItems As Collection
Private m_RuntimeIsOpen As Boolean
Private m_IsConfigured As Boolean

' //
' // Interface
' //
Private Sub obj_IControl_Configure(ByVal controlNode As Object)
    Dim selectedIdText As String

    ' Полный reset состояния: важно при повторной конфигурации того же VM.
    m_IsConfigured = False
    Set m_Layout = Nothing
    Set m_Items = Nothing
    Set m_ItemCaptions = Nothing
    Set m_ItemIds = Nothing
    Set m_ItemActionMacros = Nothing
    Set m_ItemRawItems = Nothing
    m_RuntimeHeaderShapeName = vbNullString
    m_RuntimePanelShapeName = vbNullString
    Set m_RuntimeItemShapeNames = Nothing
    Set m_RuntimeItemCaptions = Nothing
    Set m_RuntimeItemIds = Nothing
    Set m_RuntimeItemActionMacros = Nothing
    Set m_RuntimeItemRawItems = Nothing
    m_RuntimeIsOpen = False
    m_SelectedIndex = 0

    If controlNode Is Nothing Then
        MsgBox "Select: control node is not specified.", vbExclamation
        Exit Sub
    End If

    ' 1) Читаем базовые атрибуты из XML.
    m_ControlName = Trim$(ex_XmlCore.m_NodeAttrText(controlNode, "name"))
    If Len(m_ControlName) = 0 Then m_ControlName = "select"

    m_ItemsSourceRaw = Trim$(CStr(ex_XmlCore.m_NodeAttrText(controlNode, "itemsSource")))
    If Len(m_ItemsSourceRaw) = 0 Then
        MsgBox "Select: itemsSource is not specified for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    m_SelectedIdRaw = Trim$(CStr(ex_XmlCore.m_NodeAttrText(controlNode, "selectedId")))
    m_SelectedItemSourceRaw = Trim$(CStr(ex_XmlCore.m_NodeAttrText(controlNode, "selectedItem")))
    If Len(m_SelectedItemSourceRaw) = 0 Then
        m_SelectedItemSourceRaw = Trim$(CStr(ex_XmlCore.m_NodeAttrText(controlNode, "selectedItemSource")))
    End If

    m_PlaceholderText = CStr(ex_XmlCore.m_NodeAttrText(controlNode, "placeholder"))
    If Len(Trim$(m_PlaceholderText)) = 0 Then m_PlaceholderText = DEFAULT_PLACEHOLDER

    m_ItemStyleName = Trim$(CStr(ex_XmlCore.m_NodeAttrText(controlNode, "itemStyle")))
    m_PanelStyleName = Trim$(CStr(ex_XmlCore.m_NodeAttrText(controlNode, "panelStyle")))

    If Not mp_TryReadPositiveDoubleAttr(controlNode, "itemHeight", DEFAULT_ITEM_HEIGHT, m_ItemHeight) Then Exit Sub
    If Not mp_TryReadNonNegativeDoubleAttr(controlNode, "itemMargin", DEFAULT_ITEM_MARGIN, m_ItemMargin) Then Exit Sub

    m_OnChangeRaw = CStr(ex_XmlCore.m_NodeAttrText(controlNode, "onChange"))
    m_OnChangeMacroRef = vbNullString
    If Len(Trim$(m_OnChangeRaw)) > 0 Then
        If Not ex_BindingRuntime.m_TryResolveMacroBinding(m_OnChangeRaw, Me, m_OnChangeMacroRef) Then Exit Sub
    End If

    ' 2) Читаем общий layout (лист + границы + style).
    Set m_Layout = New obj_ControlLayout
    If Not m_Layout.m_TryReadFromNode(controlNode, "Select", m_ControlName, "style") Then Exit Sub

    ' 3) Разрешаем itemsSource в runtime-коллекцию и готовим буферы.
    If Not ex_ListItemsSourceRuntime.m_TryResolveItemsSource(m_ItemsSourceRaw, m_Items) Then Exit Sub
    If Not mp_TryBuildItemBuffers() Then Exit Sub

    ' 4) Определяем начальный выбранный элемент:
    '    selectedId из XML -> state store -> fallback на первый item.
    m_SelectStateKey = LCase$(m_Layout.LayoutSheet & "|" & m_ControlName)
    If Not mp_TryResolveSelectedIdText(selectedIdText) Then Exit Sub
    m_SelectedIndex = mp_FindSelectedIndexById(selectedIdText)

    If m_SelectedIndex = 0 And Not m_ItemIds Is Nothing Then
        If m_ItemIds.Count > 0 Then m_SelectedIndex = 1
    End If

    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render(ByVal wb As Workbook)
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
    Dim selectId As String
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

    If Not m_IsConfigured Then
        MsgBox "Select: control '" & m_ControlName & "' is not configured.", vbExclamation
        Exit Sub
    End If

    If wb Is Nothing Then Set wb = ThisWorkbook
    If wb Is Nothing Then
        MsgBox "Select: workbook is not specified for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    Set ws = mp_GetWorksheetByName(wb, m_Layout.LayoutSheet)
    If ws Is Nothing Then
        MsgBox "Select: sheet '" & m_Layout.LayoutSheet & "' was not found for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    If m_ItemCaptions Is Nothing Or m_ItemIds Is Nothing Or m_ItemActionMacros Is Nothing Or m_ItemRawItems Is Nothing Then
        MsgBox "Select: item metadata is not configured for control '" & m_ControlName & "'.", vbExclamation
        Exit Sub
    End If

    ' Shape.OnAction вызывает только модульный макрос,
    ' поэтому указываем dispatcher ex_ShapeClickDispatcher.m_OnShapeClick.
    callbackMacroRef = mp_GetRuntimeCallbackMacroRef()
    If Len(callbackMacroRef) = 0 Then Exit Sub

    If Not mp_TryBuildHeaderRange(ws, headerRange) Then Exit Sub
    renderItemCount = mp_GetRenderItemCount()

    ' Перерисовка "с нуля": удаляем предыдущие shape контрола.
    mp_DeleteControlShapes ws

    headerLeft = headerRange.Left
    headerTop = headerRange.Top
    headerWidth = headerRange.Width
    headerHeight = headerRange.Height

    panelLeft = headerLeft
    panelTop = headerTop + headerHeight
    panelWidth = headerWidth
    panelHeight = mp_CalcPanelHeight(renderItemCount)

    Set headerShape = mp_CreateShapeByRange(ws, headerRange, "header", callbackMacroRef)
    If headerShape Is Nothing Then Exit Sub
    mp_ApplyHeaderVisualDefaults headerShape

    If renderItemCount > 0 Then
        Set panelShape = mp_CreateShapeByBounds(ws, panelLeft, panelTop, panelWidth, panelHeight, "panel", vbNullString)
        If panelShape Is Nothing Then Exit Sub
        mp_ApplyPanelVisualDefaults panelShape
    Else
        Set panelShape = mp_CreateShapeByBounds(ws, headerLeft, headerTop, headerWidth, headerHeight, "panel", vbNullString)
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
        itemTop = panelTop + CDbl(i - 1) * (m_ItemHeight + m_ItemMargin)
        Set itemShape = mp_CreateShapeByBounds(ws, itemLeft, itemTop, itemWidth, m_ItemHeight, "item" & CStr(i), callbackMacroRef)
        If itemShape Is Nothing Then Exit Sub
        mp_SetShapeText itemShape, CStr(m_ItemCaptions(i))
        mp_ApplyItemVisualDefaults itemShape

        itemShapes.Add itemShape.Name
        itemCaptions.Add CStr(m_ItemCaptions(i))
        itemIds.Add CStr(m_ItemIds(i))
        itemActions.Add CStr(m_ItemActionMacros(i))
        itemRawItems.Add m_ItemRawItems(i)
    Next i

    selectedIndexRendered = m_SelectedIndex
    If selectedIndexRendered <= 0 Or selectedIndexRendered > renderItemCount Then selectedIndexRendered = 0

    ' Синхронизируем runtime-буферы VM с только что созданными shape.
    If Not mp_InitializeRuntimeState( _
        headerShapeName:=headerShape.Name, _
        panelShapeName:=panelShape.Name, _
        itemShapeNames:=itemShapes, _
        itemCaptions:=itemCaptions, _
        itemIds:=itemIds, _
        itemActionMacros:=itemActions, _
        itemRawItems:=itemRawItems, _
        selectedIndex:=selectedIndexRendered) Then Exit Sub

    selectId = m_SelectStateKey
    ' Регистрируем этот VM и маршруты shape-кликов в универсальном dispatcher-е.
    If Not ex_ShapeClickDispatcher.m_RegisterControl(selectId, Me) Then Exit Sub
    If Not ex_ShapeClickDispatcher.m_RegisterShapeRoute(headerShape.Name, selectId, "m_RuntimeHandleHeaderClick", False) Then Exit Sub

    For i = 1 To itemShapes.Count
        If Not ex_ShapeClickDispatcher.m_RegisterShapeRoute(CStr(itemShapes(i)), selectId, "m_RuntimeHandleItemClick", True, CLng(i)) Then Exit Sub
    Next i

    If m_RuntimeHasSelectedItem() Then
        ex_ShapeClickDispatcher.m_SetSelectContextFromVm selectId, Me
    End If
End Sub

Private Function obj_IControl_SupportsAttribute(ByVal attrName As String) As Boolean
    Select Case LCase$(Trim$(attrName))
        Case "itemssource", "placeholder", "onchange", "selectedid", "selecteditem", "selecteditemsource", _
             "style", _
             "itemstyle", "panelstyle", "itemheight", "itemmargin"
            obj_IControl_SupportsAttribute = True
    End Select
End Function

' //
' // API
' //
Public Function m_GetSelectedId() As String
    If Not m_RuntimeItemIds Is Nothing Then
        If m_SelectedIndex <= 0 Or m_SelectedIndex > m_RuntimeItemIds.Count Then Exit Function
        m_GetSelectedId = CStr(m_RuntimeItemIds(m_SelectedIndex))
        Exit Function
    End If

    If m_ItemIds Is Nothing Then Exit Function
    If m_SelectedIndex <= 0 Or m_SelectedIndex > m_ItemIds.Count Then Exit Function
    m_GetSelectedId = CStr(m_ItemIds(m_SelectedIndex))
End Function

Public Function m_GetSelectedCaption() As String
    If Not m_RuntimeItemCaptions Is Nothing Then
        If m_SelectedIndex <= 0 Or m_SelectedIndex > m_RuntimeItemCaptions.Count Then Exit Function
        m_GetSelectedCaption = CStr(m_RuntimeItemCaptions(m_SelectedIndex))
        Exit Function
    End If

    If m_ItemCaptions Is Nothing Then Exit Function
    If m_SelectedIndex <= 0 Or m_SelectedIndex > m_ItemCaptions.Count Then Exit Function
    m_GetSelectedCaption = CStr(m_ItemCaptions(m_SelectedIndex))
End Function

Public Function m_GetControlKey() As String
    m_GetControlKey = CStr(m_ControlName)
End Function

Public Function m_RuntimeHasSelectedItem() As Boolean
    If m_RuntimeItemCaptions Is Nothing Then Exit Function
    m_RuntimeHasSelectedItem = (m_SelectedIndex > 0 And m_SelectedIndex <= m_RuntimeItemCaptions.Count)
End Function

Public Function m_RuntimeGetSelectedIndex() As Long
    m_RuntimeGetSelectedIndex = m_SelectedIndex
End Function

Public Function m_RuntimeGetSelectedCaption() As String
    m_RuntimeGetSelectedCaption = m_GetSelectedCaption()
End Function

Public Function m_RuntimeGetSelectedId() As String
    m_RuntimeGetSelectedId = m_GetSelectedId()
End Function

Public Function m_RuntimeHandleHeaderClick() As Boolean
    ' Клик по header только переключает open/close dropdown panel.
    m_RuntimeIsOpen = (Not m_RuntimeIsOpen)
    m_RuntimeHandleHeaderClick = mp_ApplyRuntimeVisualState()
End Function

Public Function m_RuntimeHandleItemClick(ByVal itemIndex As Long) As Boolean
    Dim selectedId As String
    Dim itemMacro As String

    If m_RuntimeItemCaptions Is Nothing Then Exit Function
    If itemIndex <= 0 Or itemIndex > m_RuntimeItemCaptions.Count Then
        MsgBox "Select: selected item index is out of range for control '" & m_ControlName & "'.", vbExclamation
        Exit Function
    End If

    m_SelectedIndex = itemIndex
    m_RuntimeIsOpen = False

    If Not mp_ApplyRuntimeVisualState() Then Exit Function

    ' При выборе item:
    ' 1) фиксируем selectedId в CustomXMLPart
    ' 2) публикуем selectedItem в objectSource
    ' 3) запускаем item macro и onChange.
    selectedId = m_GetSelectedId()
    If Not mp_TryPersistSelectedId(selectedId) Then Exit Function
    If Not mp_TryPublishRuntimeSelectedItem(itemIndex) Then Exit Function

    itemMacro = mp_GetRuntimeCollectionText(m_RuntimeItemActionMacros, itemIndex)
    If Len(itemMacro) > 0 Then
        If Not mp_RunRuntimeMacro(itemMacro) Then Exit Function
    End If

    If Len(m_OnChangeMacroRef) > 0 Then
        If Not mp_RunRuntimeMacro(m_OnChangeMacroRef) Then Exit Function
    End If

    m_RuntimeHandleItemClick = True
End Function

Public Function m_RuntimeCloseDropdown() As Boolean
    ' Явное закрытие dropdown (используется dispatcher-ом при клике по другим shape).
    If Not m_RuntimeIsOpen Then
        m_RuntimeCloseDropdown = True
        Exit Function
    End If

    m_RuntimeIsOpen = False
    m_RuntimeCloseDropdown = mp_ApplyRuntimeVisualState()
End Function

' //
' // Internal
' //
Private Function mp_InitializeRuntimeState( _
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
        MsgBox "Select: runtime item metadata collection is not specified for control '" & m_ControlName & "'.", vbExclamation
        Exit Function
    End If

    m_RuntimeHeaderShapeName = CStr(headerShapeName)
    m_RuntimePanelShapeName = CStr(panelShapeName)
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
    If Not mp_ApplyRuntimeVisualState() Then Exit Function

    If m_RuntimeHasSelectedItem() Then
        ' Сразу синхронизируем persistent state и selectedItem-source.
        selectedId = m_GetSelectedId()
        If Not mp_TryPersistSelectedId(selectedId) Then Exit Function
        If Not mp_TryPublishRuntimeSelectedItem(m_SelectedIndex) Then Exit Function
    End If

    mp_InitializeRuntimeState = True
End Function

Private Function mp_ApplyRuntimeVisualState() As Boolean
    Dim ws As Worksheet
    Dim headerShape As Shape
    Dim panelShape As Shape
    Dim itemShape As Shape
    Dim i As Long
    Dim headerText As String

    If Len(m_RuntimeHeaderShapeName) = 0 Then Exit Function
    If m_RuntimeItemShapeNames Is Nothing Then Exit Function
    If m_RuntimeItemCaptions Is Nothing Then Exit Function

    Set ws = mp_GetRuntimeWorksheetByName(m_Layout.LayoutSheet)
    If ws Is Nothing Then Exit Function

    Set headerShape = mp_GetRuntimeShapeByName(ws, m_RuntimeHeaderShapeName)
    If headerShape Is Nothing Then Exit Function

    Set panelShape = mp_GetRuntimeShapeByName(ws, m_RuntimePanelShapeName)

    ' Header показывает либо выбранный caption, либо placeholder.
    If m_RuntimeHasSelectedItem() Then
        headerText = m_GetSelectedCaption()
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

    mp_SetShapeText headerShape, headerText

    ' Open/close panel + item-shapes.
    If Not panelShape Is Nothing Then
        panelShape.Visible = IIf(m_RuntimeIsOpen, msoTrue, msoFalse)
        If m_RuntimeIsOpen Then panelShape.ZOrder msoBringToFront
    End If

    For i = 1 To m_RuntimeItemShapeNames.Count
        Set itemShape = mp_GetRuntimeShapeByName(ws, CStr(m_RuntimeItemShapeNames(i)))
        If itemShape Is Nothing Then GoTo ContinueItem

        itemShape.Visible = IIf(m_RuntimeIsOpen, msoTrue, msoFalse)
        If m_RuntimeIsOpen Then itemShape.ZOrder msoBringToFront

        ' Выделяем выбранный item через Bold.
        On Error Resume Next
        itemShape.TextFrame.Characters.Font.Bold = (i = m_SelectedIndex)
        itemShape.TextFrame2.TextRange.Font.Bold = (i = m_SelectedIndex)
        On Error GoTo 0

ContinueItem:
    Next i

    mp_ApplyRuntimeVisualState = True
End Function

Private Function mp_TryPublishRuntimeSelectedItem(ByVal itemIndex As Long) As Boolean
    Dim selectedObject As Object
    Dim selectedRaw As Variant

    If Len(m_SelectedItemSourceRaw) = 0 Then
        mp_TryPublishRuntimeSelectedItem = True
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

        Set selectedObject = CreateObject("Scripting.Dictionary")
        selectedObject.CompareMode = 1
        selectedObject("Id") = CStr(m_RuntimeItemIds(itemIndex))
        selectedObject("Caption") = CStr(m_RuntimeItemCaptions(itemIndex))
        selectedObject("RawValue") = CStr(selectedRaw)
    End If

    If Not ex_ObjectSourceRuntime.m_SetObjectSource(m_SelectedItemSourceRaw, selectedObject, False) Then Exit Function
    mp_TryPublishRuntimeSelectedItem = True
End Function

Private Function mp_GetRuntimeCollectionText(ByVal values As Collection, ByVal idx As Long) As String
    If values Is Nothing Then Exit Function
    If idx <= 0 Or idx > values.Count Then Exit Function
    mp_GetRuntimeCollectionText = Trim$(CStr(values(idx)))
End Function

Private Function mp_RunRuntimeMacro(ByVal macroRef As String) As Boolean
    macroRef = Trim$(macroRef)
    If Len(macroRef) = 0 Then
        mp_RunRuntimeMacro = True
        Exit Function
    End If

    On Error GoTo EH_RUN
    Application.Run macroRef
    mp_RunRuntimeMacro = True
    Exit Function

EH_RUN:
    MsgBox "Select: failed to execute macro '" & macroRef & "' for control '" & m_ControlName & "': " & Err.Description, vbExclamation
End Function

Private Function mp_TryBuildItemBuffers() As Boolean
    Dim itemRaw As Variant
    Dim itemCaption As String
    Dim itemId As String
    Dim itemAction As String

    If m_Items Is Nothing Then
        MsgBox "Select: itemsSource resolved to Nothing for control '" & m_ControlName & "'.", vbExclamation
        Exit Function
    End If

    ' Нормализуем itemsSource в плоские буферы:
    ' Caption / Id / ItemAction / RawItem.
    Set m_ItemCaptions = New Collection
    Set m_ItemIds = New Collection
    Set m_ItemActionMacros = New Collection
    Set m_ItemRawItems = New Collection

    For Each itemRaw In m_Items
        itemCaption = vbNullString
        itemId = vbNullString
        itemAction = vbNullString

        If Not mp_TryResolveItemMetadata(itemRaw, itemCaption, itemId, itemAction) Then Exit Function

        m_ItemCaptions.Add itemCaption
        m_ItemIds.Add itemId
        m_ItemActionMacros.Add itemAction
        m_ItemRawItems.Add itemRaw
    Next itemRaw

    mp_TryBuildItemBuffers = True
End Function

Private Function mp_TryResolveItemMetadata( _
    ByVal itemRaw As Variant, _
    ByRef outCaption As String, _
    ByRef outId As String, _
    ByRef outItemActionMacro As String _
) As Boolean
    Dim itemObj As Object
    Dim actionRaw As String

    ' Scalar item: Caption=Id=значение, action пустой.
    If Not IsObject(itemRaw) Then
        outCaption = CStr(itemRaw)
        outId = CStr(itemRaw)
        outItemActionMacro = vbNullString
        mp_TryResolveItemMetadata = True
        Exit Function
    End If

    Set itemObj = itemRaw

    ' Object item: сначала fast-path для obj_SelectOption,
    ' затем универсальный путь (CallByName/Dictionary).
    Select Case LCase$(TypeName(itemObj))
        Case "obj_selectoption"
            outCaption = CStr(itemObj.Caption)
            outId = CStr(itemObj.Id)
            actionRaw = Trim$(CStr(itemObj.OnSelect))

        Case Else
            If Not mp_TryReadObjectMemberText(itemObj, "Caption", True, outCaption) Then Exit Function
            If Not mp_TryReadObjectMemberText(itemObj, "Id", True, outId) Then Exit Function
            If Not mp_TryReadObjectMemberText(itemObj, "OnSelect", False, actionRaw) Then Exit Function
            actionRaw = Trim$(actionRaw)
    End Select

    If Len(Trim$(outId)) = 0 Then
        MsgBox "Select: item Id is empty in control '" & m_ControlName & "'.", vbExclamation
        Exit Function
    End If
    If Len(Trim$(outCaption)) = 0 Then outCaption = outId

    outItemActionMacro = vbNullString
    If Len(actionRaw) > 0 Then
        outItemActionMacro = mp_QualifyMacroName(actionRaw)
    End If

    mp_TryResolveItemMetadata = True
End Function

Private Function mp_TryReadObjectMemberText( _
    ByVal sourceObject As Object, _
    ByVal memberName As String, _
    ByVal isRequired As Boolean, _
    ByRef outText As String _
) As Boolean
    Dim dictObj As Object
    Dim scalarValue As Variant

    outText = vbNullString
    If sourceObject Is Nothing Then
        If isRequired Then
            MsgBox "Select: item object is Nothing while reading member '" & memberName & "'.", vbExclamation
            Exit Function
        End If
        mp_TryReadObjectMemberText = True
        Exit Function
    End If

    Set dictObj = mp_AsDictionary(sourceObject)
    If Not dictObj Is Nothing Then
        If Not dictObj.Exists(memberName) Then
            If isRequired Then
                MsgBox "Select: member '" & memberName & "' was not found on dictionary item.", vbExclamation
                Exit Function
            End If

            mp_TryReadObjectMemberText = True
            Exit Function
        End If

        scalarValue = dictObj.Item(memberName)
        If IsObject(scalarValue) Then
            MsgBox "Select: member '" & memberName & "' must resolve to scalar value.", vbExclamation
            Exit Function
        End If

        outText = CStr(scalarValue)
        mp_TryReadObjectMemberText = True
        Exit Function
    End If

    On Error Resume Next
    scalarValue = CallByName(sourceObject, memberName, VbGet)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0

        If isRequired Then
            MsgBox "Select: member '" & memberName & "' was not found on object '" & TypeName(sourceObject) & "'.", vbExclamation
            Exit Function
        End If

        mp_TryReadObjectMemberText = True
        Exit Function
    End If
    On Error GoTo 0

    If IsObject(scalarValue) Then
        MsgBox "Select: member '" & memberName & "' on object '" & TypeName(sourceObject) & "' must resolve to scalar value.", vbExclamation
        Exit Function
    End If

    outText = CStr(scalarValue)
    mp_TryReadObjectMemberText = True
End Function

Private Function mp_AsDictionary(ByVal sourceObject As Object) As Object
    If sourceObject Is Nothing Then Exit Function
    If LCase$(TypeName(sourceObject)) <> "dictionary" Then Exit Function
    Set mp_AsDictionary = sourceObject
End Function

Private Function mp_TryResolveSelectedIdText(ByRef outSelectedIdText As String) As Boolean
    ' Порядок получения selectedId:
    ' 1) selectedId в XML
    ' 2) сохраненное состояние в CustomXMLPart (obj_SelectControlVMStatic)
    outSelectedIdText = Trim$(m_SelectedIdRaw)
    If Len(outSelectedIdText) > 0 Then
        mp_TryResolveSelectedIdText = True
        Exit Function
    End If

    If Not mp_TryLoadStoredSelectedId(outSelectedIdText) Then Exit Function
    outSelectedIdText = Trim$(outSelectedIdText)
    mp_TryResolveSelectedIdText = True
End Function

Private Function mp_TryPersistSelectedId(ByVal selectedId As String) As Boolean
    Dim selectStatic As obj_SelectControlVMStatic

    selectedId = Trim$(selectedId)
    If Len(Trim$(m_SelectStateKey)) = 0 Then
        MsgBox "Select: state key is empty for control '" & m_ControlName & "'.", vbExclamation
        Exit Function
    End If

    Set selectStatic = New obj_SelectControlVMStatic
    mp_TryPersistSelectedId = selectStatic.m_SetSelectedId(m_SelectStateKey, selectedId)
End Function

Private Function mp_TryLoadStoredSelectedId(ByRef outSelectedId As String) As Boolean
    Dim selectStatic As obj_SelectControlVMStatic

    If Len(Trim$(m_SelectStateKey)) = 0 Then
        MsgBox "Select: state key is empty for control '" & m_ControlName & "'.", vbExclamation
        Exit Function
    End If

    Set selectStatic = New obj_SelectControlVMStatic
    mp_TryLoadStoredSelectedId = selectStatic.m_TryGetSelectedId(m_SelectStateKey, outSelectedId)
End Function

Private Function mp_FindSelectedIndexById(ByVal selectedIdText As String) As Long
    Dim i As Long

    selectedIdText = LCase$(Trim$(selectedIdText))
    If Len(selectedIdText) = 0 Then Exit Function
    If m_ItemIds Is Nothing Then Exit Function

    For i = 1 To m_ItemIds.Count
        If LCase$(Trim$(CStr(m_ItemIds(i)))) = selectedIdText Then
            mp_FindSelectedIndexById = i
            Exit Function
        End If
    Next i
End Function

Private Function mp_GetRenderItemCount() As Long
    If m_ItemCaptions Is Nothing Then Exit Function
    mp_GetRenderItemCount = m_ItemCaptions.Count
End Function

Private Function mp_TryBuildHeaderRange(ByVal ws As Worksheet, ByRef outRange As Range) As Boolean
    If ws Is Nothing Then Exit Function

    On Error GoTo EH_RANGE
    Set outRange = ws.Range(ws.Cells(m_Layout.RowStart, m_Layout.ColStart), ws.Cells(m_Layout.RowEnd, m_Layout.ColEnd))
    mp_TryBuildHeaderRange = True
    Exit Function

EH_RANGE:
    MsgBox "Select: failed to resolve header range for control '" & m_ControlName & "'.", vbExclamation
End Function

Private Function mp_CalcPanelHeight(ByVal renderItemCount As Long) As Double
    If renderItemCount <= 0 Then Exit Function

    mp_CalcPanelHeight = CDbl(renderItemCount) * m_ItemHeight
    If renderItemCount > 1 Then
        mp_CalcPanelHeight = mp_CalcPanelHeight + CDbl(renderItemCount - 1) * m_ItemMargin
    End If
End Function

Private Function mp_CreateShapeByRange( _
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
        MsgBox "Select: target range has non-positive width/height for control '" & m_ControlName & "'.", vbExclamation
        Exit Function
    End If

    shapeName = mp_BuildShapeName(suffix)

    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo EH_SHAPE

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, targetRange.Left, targetRange.Top, targetRange.Width, targetRange.Height)
    shp.Name = shapeName

    If Len(Trim$(onActionMacroRef)) > 0 Then
        On Error Resume Next
        shp.OnAction = onActionMacroRef
        If Err.Number <> 0 Then
            MsgBox "Select: failed to bind click action for shape '" & shapeName & "' in control '" & m_ControlName & "'.", vbExclamation
            Err.Clear
        End If
        On Error GoTo EH_SHAPE
    End If

    shapeRole = LCase$(Trim$(suffix))
    shapeStyle = vbNullString
    Select Case shapeRole
        Case "header"
            shapeStyle = Trim$(m_Layout.StyleName)
        Case "panel"
            shapeStyle = Trim$(m_PanelStyleName)
        Case Else
            shapeStyle = Trim$(m_ItemStyleName)
    End Select

    On Error Resume Next
    If shapeRole = "header" Then
        shp.Placement = xlMoveAndSize
    Else
        shp.Placement = xlFreeFloating
    End If
    Err.Clear
    On Error GoTo EH_SHAPE

    Set metaMap = CreateObject("Scripting.Dictionary")
    metaMap.CompareMode = 1
    metaMap("pn.control") = m_ControlName
    metaMap("pn.role") = shapeRole
    If Len(shapeStyle) > 0 Then
        metaMap("pn.style") = shapeStyle
    Else
        metaMap("pn.style") = vbNullString
    End If
    If Not ex_ShapeMetaRuntime.m_TrySetShapeMetaValues(shp, metaMap) Then Exit Function

    Set mp_CreateShapeByRange = shp
    Exit Function

EH_SHAPE:
    MsgBox "Select: failed to create shape '" & shapeName & "' for control '" & m_ControlName & "': " & Err.Description, vbExclamation
End Function

Private Function mp_CreateShapeByBounds( _
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
        MsgBox "Select: target bounds have non-positive width/height for control '" & m_ControlName & "'.", vbExclamation
        Exit Function
    End If

    shapeName = mp_BuildShapeName(suffix)

    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo EH_SHAPE

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, shapeLeft, shapeTop, shapeWidth, shapeHeight)
    shp.Name = shapeName

    If Len(Trim$(onActionMacroRef)) > 0 Then
        On Error Resume Next
        shp.OnAction = onActionMacroRef
        If Err.Number <> 0 Then
            MsgBox "Select: failed to bind click action for shape '" & shapeName & "' in control '" & m_ControlName & "'.", vbExclamation
            Err.Clear
        End If
        On Error GoTo EH_SHAPE
    End If

    shapeRole = LCase$(Trim$(suffix))
    shapeStyle = vbNullString
    Select Case shapeRole
        Case "header"
            shapeStyle = Trim$(m_Layout.StyleName)
        Case "panel"
            shapeStyle = Trim$(m_PanelStyleName)
        Case Else
            shapeStyle = Trim$(m_ItemStyleName)
    End Select

    On Error Resume Next
    If shapeRole = "header" Then
        shp.Placement = xlMoveAndSize
    Else
        shp.Placement = xlFreeFloating
    End If
    Err.Clear
    On Error GoTo EH_SHAPE

    Set metaMap = CreateObject("Scripting.Dictionary")
    metaMap.CompareMode = 1
    metaMap("pn.control") = m_ControlName
    metaMap("pn.role") = shapeRole
    If Len(shapeStyle) > 0 Then
        metaMap("pn.style") = shapeStyle
    Else
        metaMap("pn.style") = vbNullString
    End If
    If Not ex_ShapeMetaRuntime.m_TrySetShapeMetaValues(shp, metaMap) Then Exit Function

    Set mp_CreateShapeByBounds = shp
    Exit Function

EH_SHAPE:
    MsgBox "Select: failed to create floating shape '" & shapeName & "' for control '" & m_ControlName & "': " & Err.Description, vbExclamation
End Function

Private Sub mp_ApplyHeaderVisualDefaults(ByVal shp As Shape)
    If shp Is Nothing Then Exit Sub

    shp.Fill.Visible = msoTrue
    shp.Fill.ForeColor.RGB = RGB(67, 142, 32)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(40, 93, 20)
    shp.Line.Weight = 1#

    On Error Resume Next
    shp.TextFrame.Characters.Font.Color = RGB(10, 10, 10)
    shp.TextFrame.Characters.Font.Bold = False
    shp.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp.TextFrame.VerticalAlignment = xlVAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(10, 10, 10)
    shp.TextFrame2.TextRange.Font.Bold = False
    On Error GoTo 0
End Sub

Private Sub mp_ApplyPanelVisualDefaults(ByVal shp As Shape)
    If shp Is Nothing Then Exit Sub

    shp.Fill.Visible = msoTrue
    shp.Fill.ForeColor.RGB = RGB(43, 49, 56)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(28, 32, 36)
    shp.Line.Weight = 0.75
End Sub

Private Sub mp_ApplyItemVisualDefaults(ByVal shp As Shape)
    If shp Is Nothing Then Exit Sub

    shp.Fill.Visible = msoTrue
    shp.Fill.ForeColor.RGB = RGB(59, 66, 74)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(40, 45, 50)
    shp.Line.Weight = 0.75

    On Error Resume Next
    shp.TextFrame.Characters.Font.Color = RGB(245, 245, 245)
    shp.TextFrame.HorizontalAlignment = xlHAlignLeft
    shp.TextFrame.VerticalAlignment = xlVAlignCenter
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(245, 245, 245)
    On Error GoTo 0
End Sub

Private Sub mp_SetShapeText(ByVal shp As Shape, ByVal textValue As String)
    If shp Is Nothing Then Exit Sub

    On Error Resume Next
    shp.TextFrame2.TextRange.Text = textValue
    shp.TextFrame.Characters.Text = textValue
    On Error GoTo 0
End Sub

Private Sub mp_DeleteControlShapes(ByVal ws As Worksheet)
    Dim i As Long
    Dim shp As Shape
    Dim controlMeta As String

    If ws Is Nothing Then Exit Sub

    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        controlMeta = LCase$(Trim$(ex_ShapeMetaRuntime.m_GetShapeMetaValue(shp, "pn.control", vbNullString)))
        If Len(controlMeta) = 0 Then GoTo ContinueShape
        If controlMeta = LCase$(Trim$(m_ControlName)) Then
            shp.Delete
        End If
ContinueShape:
    Next i
End Sub

Private Function mp_BuildShapeName(ByVal suffix As String) As String
    mp_BuildShapeName = "sel_" & mp_NormalizeNamePart(m_ControlName) & "_" & mp_NormalizeNamePart(suffix)
End Function

Private Function mp_NormalizeNamePart(ByVal rawText As String) As String
    Dim i As Long
    Dim ch As String
    Dim outText As String

    rawText = Trim$(rawText)
    If Len(rawText) = 0 Then
        mp_NormalizeNamePart = "x"
        Exit Function
    End If

    For i = 1 To Len(rawText)
        ch = Mid$(rawText, i, 1)
        If (ch >= "A" And ch <= "Z") Or _
           (ch >= "a" And ch <= "z") Or _
           (ch >= "0" And ch <= "9") Or _
           ch = "_" Then
            outText = outText & ch
        Else
            outText = outText & "_"
        End If
    Next i

    If Len(outText) = 0 Then outText = "x"
    mp_NormalizeNamePart = Left$(outText, 120)
End Function

Private Function mp_GetRuntimeCallbackMacroRef() As String
    mp_GetRuntimeCallbackMacroRef = mp_QualifyMacroName("ex_ShapeClickDispatcher.m_OnShapeClick")
End Function

Private Function mp_QualifyMacroName(ByVal macroName As String) As String
    Dim wbName As String

    macroName = Trim$(macroName)
    If Len(macroName) = 0 Then Exit Function
    If InStr(1, macroName, "!", vbBinaryCompare) > 0 Then
        mp_QualifyMacroName = macroName
        Exit Function
    End If

    wbName = ThisWorkbook.Name
    wbName = Replace$(wbName, "'", "''")
    mp_QualifyMacroName = "'" & wbName & "'!" & macroName
End Function

Private Function mp_TryReadPositiveDoubleAttr( _
    ByVal controlNode As Object, _
    ByVal attrName As String, _
    ByVal defaultValue As Double, _
    ByRef outValue As Double _
) As Boolean
    Dim rawText As String

    rawText = Trim$(CStr(ex_XmlCore.m_NodeAttrText(controlNode, attrName)))
    If Len(rawText) = 0 Then
        outValue = defaultValue
        mp_TryReadPositiveDoubleAttr = True
        Exit Function
    End If

    If Not mp_TryParseFlexibleDouble(rawText, outValue) Then
        MsgBox "Select: attribute '" & attrName & "' must be numeric for control '" & m_ControlName & "'.", vbExclamation
        Exit Function
    End If

    If outValue <= 0# Then
        MsgBox "Select: attribute '" & attrName & "' must be greater than zero for control '" & m_ControlName & "'.", vbExclamation
        Exit Function
    End If

    mp_TryReadPositiveDoubleAttr = True
End Function

Private Function mp_TryReadNonNegativeDoubleAttr( _
    ByVal controlNode As Object, _
    ByVal attrName As String, _
    ByVal defaultValue As Double, _
    ByRef outValue As Double _
) As Boolean
    Dim rawText As String

    rawText = Trim$(CStr(ex_XmlCore.m_NodeAttrText(controlNode, attrName)))
    If Len(rawText) = 0 Then
        outValue = defaultValue
        mp_TryReadNonNegativeDoubleAttr = True
        Exit Function
    End If

    If Not mp_TryParseFlexibleDouble(rawText, outValue) Then
        MsgBox "Select: attribute '" & attrName & "' must be numeric for control '" & m_ControlName & "'.", vbExclamation
        Exit Function
    End If

    If outValue < 0# Then
        MsgBox "Select: attribute '" & attrName & "' must be greater or equal to zero for control '" & m_ControlName & "'.", vbExclamation
        Exit Function
    End If

    mp_TryReadNonNegativeDoubleAttr = True
End Function

Private Function mp_TryParseFlexibleDouble(ByVal rawText As String, ByRef outValue As Double) As Boolean
    Dim normalized As String
    Dim decimalSep As String

    rawText = Trim$(rawText)
    If Len(rawText) = 0 Then Exit Function

    decimalSep = CStr(Application.International(xlDecimalSeparator))
    normalized = rawText

    If decimalSep = "," Then
        normalized = Replace$(normalized, ".", ",")
    Else
        normalized = Replace$(normalized, ",", ".")
    End If

    If Not IsNumeric(normalized) Then Exit Function
    outValue = CDbl(normalized)
    mp_TryParseFlexibleDouble = True
End Function

Private Function mp_GetWorksheetByName(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set mp_GetWorksheetByName = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function mp_GetRuntimeWorksheetByName(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set mp_GetRuntimeWorksheetByName = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function mp_GetRuntimeShapeByName(ByVal ws As Worksheet, ByVal shapeName As String) As Shape
    If ws Is Nothing Then Exit Function
    If Len(Trim$(shapeName)) = 0 Then Exit Function

    On Error Resume Next
    Set mp_GetRuntimeShapeByName = ws.Shapes(shapeName)
    On Error GoTo 0
End Function
