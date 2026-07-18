VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_ButtonGroupControlVM"
Option Explicit
#Const LOGGING_DEBUG_ENABLED = True
#Const LOGGING_VERBOSE_ENABLED = False

Implements obj_IControl

Private Const DEFAULT_COLUMNS As Long = 2
Private Const DEFAULT_CAPTION_PATH As String = "Caption"
Private Const DEFAULT_ID_PATH As String = "Id"
Private Const DEFAULT_STYLE_PATH As String = "StyleName"
Private Const DEFAULT_ITEM_SPAN_ROWS As Long = 1
Private Const FLOW_ROW As String = "row"
Private Const FLOW_COLUMN As String = "column"

Private m_ControlBase As obj_ControlBase
Private m_ControlLayout As obj_ControlLayout
Private m_Page As obj_IPage
Private m_ControlName As String
Private m_ItemsSourceRaw As String
Private m_OnClickRaw As String
Private m_OnClickMacroRef As String
Private m_OnClickCallbackContext As Object
Private m_Items As Collection
Private m_RuntimeControlKey As String
Private m_Columns As Long
Private m_CaptionPath As String
Private m_IdPath As String
Private m_StylePath As String
Private m_ShapePrefix As String
Private m_FlowDirection As String
Private m_ItemSpanRows As Long
Private m_IsConfigured As Boolean
Private m_IsDisposed As Boolean

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
    Set m_Page = Nothing
    Set m_OnClickCallbackContext = Nothing
    Set m_Items = Nothing
    On Error GoTo 0
End Sub

Private Sub obj_IControl_Configure(ByVal controlNode As Object)
    Dim pageBase As obj_PageBase
    Dim dataContext As Object
    Dim callbackContext As Object
    Dim onClickResolved As Variant
    Dim columnsRaw As String

    m_IsConfigured = False
    Set m_ControlBase = Nothing
    Set m_ControlLayout = Nothing
    Set m_OnClickCallbackContext = Nothing
    Set m_Items = Nothing
    m_RuntimeControlKey = VBA.vbNullString
    m_OnClickMacroRef = VBA.vbNullString
    m_Columns = DEFAULT_COLUMNS
    m_CaptionPath = DEFAULT_CAPTION_PATH
    m_IdPath = DEFAULT_ID_PATH
    m_StylePath = DEFAULT_STYLE_PATH
    m_ShapePrefix = VBA.vbNullString
    m_FlowDirection = FLOW_ROW
    m_ItemSpanRows = DEFAULT_ITEM_SPAN_ROWS

    If m_Page Is Nothing Then Exit Sub
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Sub

    Set m_ControlBase = New obj_ControlBase
    If Not m_ControlBase.Initialize(m_Page) Then Exit Sub
    If Not m_ControlBase.Configure(pageBase, controlNode, "ButtonGroup", "buttongroup", m_ControlName) Then Exit Sub

    m_ItemsSourceRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "itemsSource")))
    If VBA.Len(m_ItemsSourceRaw) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "ButtonGroup: itemsSource is required for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If
    If Not ex_RuntimeSourceResolver.fn_TryResolveItemsSource(pageBase.RuntimeSources, m_ItemsSourceRaw, m_Items) Then Exit Sub
    If m_Items Is Nothing Then Exit Sub

    Set dataContext = m_ControlBase.DataContext
    If dataContext Is Nothing Then Set dataContext = m_Page

    m_OnClickRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "onClick")))
    If VBA.Len(m_OnClickRaw) = 0 Then
#If LOGGING_DEBUG_ENABLED Then
        ex_Core.fn_Diagnostic_LogError "ButtonGroup: onClick is required for control '" & m_ControlName & "'."
#End If
        Exit Sub
    End If

    Set callbackContext = dataContext
    ' Для ButtonGroup callback target можно указать прямо в onClick binding:
    '   onClick="{Binding DataContext={PageRuntimeSource='RuntimeObjects.Page.Controller'}; Method=SelectItem}"
    ' Это держит метод и его target в одном выражении.
    If Not ex_BindingRuntime.fn_TryResolveBindingSourceObject(m_OnClickRaw, pageBase.RuntimeSources, dataContext, callbackContext) Then Exit Sub

    If Not ex_BindingRuntime.fn_TryResolveValueBinding(m_OnClickRaw, callbackContext, onClickResolved) Then Exit Sub
    If VBA.IsObject(onClickResolved) Then Exit Sub
    m_OnClickMacroRef = VBA.Trim$(VBA.CStr(onClickResolved))
    If VBA.Len(m_OnClickMacroRef) = 0 Then Exit Sub
    Set m_OnClickCallbackContext = callbackContext

    columnsRaw = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "columns")))
    If VBA.Len(columnsRaw) > 0 And VBA.IsNumeric(columnsRaw) Then
        m_Columns = VBA.CLng(columnsRaw)
        If m_Columns <= 0 Then m_Columns = DEFAULT_COLUMNS
    End If

    m_CaptionPath = private_ReadPathAttr(controlNode, "captionPath", DEFAULT_CAPTION_PATH)
    m_IdPath = private_ReadPathAttr(controlNode, "idPath", DEFAULT_ID_PATH)
    m_StylePath = private_ReadPathAttr(controlNode, "stylePath", DEFAULT_STYLE_PATH)
    m_ShapePrefix = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "shapePrefix")))
    m_FlowDirection = private_ReadFlowDirection(controlNode)
    ' itemSpanRows задает высоту одной кнопки в layout-строках.
    m_ItemSpanRows = private_ReadPositiveLongAttr(controlNode, "itemSpanRows", DEFAULT_ITEM_SPAN_ROWS)

    Set m_ControlLayout = New obj_ControlLayout
    If Not m_ControlLayout.TryReadFromNode(controlNode, "ButtonGroup", m_ControlName, "style") Then Exit Sub
    m_RuntimeControlKey = "buttongroup|" & VBA.LCase$(VBA.Trim$(m_ControlLayout.LayoutSheetName & "|" & m_ControlName))

    m_IsConfigured = True
End Sub

Private Sub obj_IControl_Render()
    Dim pageBase As obj_PageBase
    Dim ws As Worksheet
    Dim itemCount As Long
    Dim itemIndex As Long
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim rowStart As Long
    Dim rowEnd As Long
    Dim colStart As Long
    Dim colEnd As Long
    Dim targetRange As Range
    Dim itemObj As Variant
    Dim captionText As String
    Dim itemId As String
    Dim styleName As String
    Dim shp As Shape
    Dim shapeName As String
    Dim macroRef As String
    Dim flattenedItems As Collection
    Dim rowsPerColumn As Long
    Dim buttonSpanCols As Long
    Dim renderSignature As String
    Dim previousRenderSignature As String
    Dim visualUnchanged As Boolean

    If Not m_IsConfigured Then Exit Sub
    If m_Page Is Nothing Then Exit Sub
    Set pageBase = m_Page.GetPageBase()
    If pageBase Is Nothing Then Exit Sub
    Set ws = pageBase.Worksheet
    If ws Is Nothing Then Exit Sub
    If m_Items Is Nothing Then Exit Sub

    Set flattenedItems = private_FlattenItems(m_Items)
    itemCount = flattenedItems.Count
    If itemCount <= 0 Then
        If Not pageBase.RegisterControl(m_RuntimeControlKey, Me) Then Exit Sub
        private_DeleteExtraShapes ws, 1
        Exit Sub
    End If

    macroRef = private_GetRuntimeCallbackMacroRef()
    If VBA.Len(macroRef) = 0 Then Exit Sub

    If Not pageBase.RegisterControl(m_RuntimeControlKey, Me) Then Exit Sub

    rowsPerColumn = private_RowsPerColumn(itemCount)
    buttonSpanCols = private_ButtonSpanCols()
    itemIndex = 0
    For Each itemObj In flattenedItems
        itemIndex = itemIndex + 1
        If Not private_TryReadItem(itemObj, captionText, itemId, styleName) Then GoTo ContinueItem

        If private_IsColumnFlow() Then
            rowIndex = (itemIndex - 1) Mod rowsPerColumn
            colIndex = (itemIndex - 1) \ rowsPerColumn
        Else
            rowIndex = (itemIndex - 1) \ m_Columns
            colIndex = (itemIndex - 1) Mod m_Columns
        End If
        If colIndex >= m_Columns Then Exit For

        rowStart = m_ControlLayout.RowStart + rowIndex * m_ItemSpanRows
        rowEnd = rowStart + m_ItemSpanRows - 1
        colStart = m_ControlLayout.ColStart + colIndex * buttonSpanCols
        colEnd = colStart + buttonSpanCols - 1
        If rowStart > m_ControlLayout.RowEnd Then Exit For
        If rowEnd > m_ControlLayout.RowEnd Then rowEnd = m_ControlLayout.RowEnd
        If colEnd > m_ControlLayout.ColEnd Then colEnd = m_ControlLayout.ColEnd

        On Error Resume Next
        Set targetRange = ws.Range(ws.Cells(rowStart, colStart), ws.Cells(rowEnd, colEnd))
        On Error GoTo 0
        If targetRange Is Nothing Then GoTo ContinueItem

        shapeName = private_BuildShapeName(itemIndex)
        Set shp = private_GetOrCreateShape(ws, shapeName, targetRange)
        If shp Is Nothing Then GoTo ContinueItem

        If VBA.Len(styleName) = 0 Then styleName = m_ControlLayout.StyleName
        ' У ButtonGroup один VM управляет десятками Shapes. Сравнение короткой
        ' signature в памяти дешевле повторной записи caption/alignment/meta
        ' через Excel COM для каждого неизменившегося элемента группы.
        renderSignature = private_BuildItemVisualSignature( _
            itemIndex, rowStart, colStart, rowEnd, colEnd, captionText, styleName)
        previousRenderSignature = ex_ShapeMetaRuntime.fn_GetShapeMetaValue( _
            shp, "pn.renderSignature", VBA.vbNullString)
        visualUnchanged = (VBA.StrComp(previousRenderSignature, renderSignature, VBA.vbBinaryCompare) = 0)
        If Not visualUnchanged Then
            private_ApplyShapeContent shp, captionText
            If Not private_SetShapeMeta(shp, styleName, renderSignature) Then GoTo ContinueItem
        End If
        ' OnAction и route относятся к runtime-состоянию текущего прохода,
        ' поэтому проверяем/регистрируем их даже при неизменившемся Shape.
        If Not private_AssignShapeOnAction(shp, macroRef) Then GoTo ContinueItem
        If Not pageBase.RegisterShapeRoute(shp.Name, m_RuntimeControlKey, "RuntimeHandleClick", True, itemId) Then GoTo ContinueItem

ContinueItem:
        Set targetRange = Nothing
        Set shp = Nothing
    Next itemObj

    private_DeleteExtraShapes ws, itemIndex + 1
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
        Case "itemssource", "onclick", "columns", "captionpath", "idpath", "stylepath", "shapeprefix", "flow", "itemspanrows"
            obj_IControl_SupportsAttribute = True
    End Select
End Function

Private Function obj_IControl_IsConfigured() As Boolean
    obj_IControl_IsConfigured = m_IsConfigured
End Function

Public Function RuntimeHandleClick(Optional ByVal itemId As Variant) As Boolean
    If Not rt_Bridge.fn_RunCallback(m_OnClickMacroRef, m_OnClickCallbackContext, itemId) Then Exit Function
    RuntimeHandleClick = True
End Function

Private Function private_ReadPathAttr(ByVal controlNode As Object, ByVal attrName As String, ByVal defaultValue As String) As String
    Dim attrValue As String

    attrValue = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, attrName)))
    If VBA.Len(attrValue) = 0 Then attrValue = defaultValue
    private_ReadPathAttr = attrValue
End Function

Private Function private_ReadFlowDirection(ByVal controlNode As Object) As String
    Dim flowText As String

    flowText = VBA.LCase$(VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, "flow"))))
    Select Case flowText
        Case FLOW_COLUMN
            private_ReadFlowDirection = FLOW_COLUMN
        Case Else
            private_ReadFlowDirection = FLOW_ROW
    End Select
End Function

Private Function private_ReadPositiveLongAttr( _
    ByVal controlNode As Object, _
    ByVal attrName As String, _
    ByVal defaultValue As Long _
) As Long
    Dim attrValue As String

    private_ReadPositiveLongAttr = defaultValue
    attrValue = VBA.Trim$(VBA.CStr(ex_XmlCore.fn_NodeAttrText(controlNode, attrName)))
    If VBA.Len(attrValue) = 0 Then Exit Function
    If Not VBA.IsNumeric(attrValue) Then Exit Function

    private_ReadPositiveLongAttr = VBA.CLng(attrValue)
    If private_ReadPositiveLongAttr <= 0 Then private_ReadPositiveLongAttr = defaultValue
End Function

Private Function private_ButtonSpanCols() As Long
    Dim availableCols As Long

    availableCols = m_ControlLayout.ColEnd - m_ControlLayout.ColStart + 1
    If availableCols <= 0 Or m_Columns <= 0 Then
        private_ButtonSpanCols = 1
    Else
        private_ButtonSpanCols = availableCols \ m_Columns
        If private_ButtonSpanCols <= 0 Then private_ButtonSpanCols = 1
    End If
End Function

Private Function private_RowsPerColumn(ByVal itemCount As Long) As Long
    Dim availableRows As Long
    Dim availableItemRows As Long

    availableRows = m_ControlLayout.RowEnd - m_ControlLayout.RowStart + 1
    If availableRows <= 0 Then availableRows = 1
    availableItemRows = availableRows \ m_ItemSpanRows
    If availableItemRows <= 0 Then availableItemRows = 1
    If itemCount <= 0 Then
        private_RowsPerColumn = availableItemRows
    ElseIf Not private_IsColumnFlow() Then
        private_RowsPerColumn = availableItemRows
    Else
        private_RowsPerColumn = (itemCount + m_Columns - 1) \ m_Columns
        If private_RowsPerColumn <= 0 Then private_RowsPerColumn = 1
        If private_RowsPerColumn > availableItemRows Then private_RowsPerColumn = availableItemRows
    End If
End Function

Private Function private_IsColumnFlow() As Boolean
    private_IsColumnFlow = (VBA.StrComp(m_FlowDirection, FLOW_COLUMN, VBA.vbTextCompare) = 0)
End Function

Private Function private_FlattenItems(ByVal sourceItems As Collection) As Collection
    Dim flattened As Collection
    Dim itemObj As Variant
    Dim rowItems As Collection
    Dim rowItem As Variant

    Set flattened = New Collection
    If sourceItems Is Nothing Then
        Set private_FlattenItems = flattened
        Exit Function
    End If

    For Each itemObj In sourceItems
        Set rowItems = private_TryGetNestedItems(itemObj)
        If rowItems Is Nothing Then
            flattened.Add itemObj
        Else
            For Each rowItem In rowItems
                flattened.Add rowItem
            Next rowItem
        End If
    Next itemObj

    Set private_FlattenItems = flattened
End Function

Private Function private_TryGetNestedItems(ByVal itemObj As Variant) As Collection
    Dim valueObj As Object

    If Not VBA.IsObject(itemObj) Then Exit Function
    Set valueObj = itemObj

    On Error Resume Next
    If VBA.TypeName(valueObj) = "Dictionary" Or VBA.TypeName(valueObj) = "Scripting.Dictionary" Then
        If valueObj.Exists("Items") Then Set private_TryGetNestedItems = valueObj("Items")
    Else
        Set private_TryGetNestedItems = VBA.CallByName(valueObj, "Items", VbGet)
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function private_TryReadItem( _
    ByVal itemObj As Variant, _
    ByRef outCaption As String, _
    ByRef outId As String, _
    ByRef outStyleName As String _
) As Boolean
    outCaption = private_ReadPathText(itemObj, m_CaptionPath)
    outId = private_ReadPathText(itemObj, m_IdPath)
    outStyleName = private_ReadPathText(itemObj, m_StylePath)

    If Not VBA.IsObject(itemObj) Then
        If VBA.Len(VBA.Trim$(outCaption)) = 0 Then outCaption = VBA.Trim$(VBA.CStr(itemObj))
    End If
    If VBA.Len(VBA.Trim$(outId)) = 0 Then outId = outCaption

    private_TryReadItem = (VBA.Len(VBA.Trim$(outCaption)) > 0)
End Function

Private Function private_ReadPathText(ByVal itemObj As Variant, ByVal pathText As String) As String
    Dim currentValue As Variant

    pathText = VBA.Trim$(pathText)
    If VBA.Len(pathText) = 0 Then Exit Function
    If pathText = "." Then
        If VBA.IsObject(itemObj) Then Exit Function
        private_ReadPathText = VBA.Trim$(VBA.CStr(itemObj))
        Exit Function
    End If

    If Not private_TryReadPathValue(itemObj, pathText, currentValue) Then Exit Function
    If VBA.IsObject(currentValue) Then Exit Function
    private_ReadPathText = VBA.Trim$(VBA.CStr(currentValue))
End Function

Private Function private_TryReadPathValue(ByVal itemObj As Variant, ByVal pathText As String, ByRef outValue As Variant) As Boolean
    Dim pathParts() As String
    Dim pathIndex As Long
    Dim currentValue As Variant
    Dim nextValue As Variant

    pathParts = VBA.Split(pathText, ".")
    If VBA.IsObject(itemObj) Then
        Set currentValue = itemObj
    Else
        currentValue = itemObj
    End If

    For pathIndex = LBound(pathParts) To UBound(pathParts)
        If Not private_TryReadMemberValue(currentValue, VBA.Trim$(pathParts(pathIndex)), nextValue) Then Exit Function
        If VBA.IsObject(nextValue) Then
            Set currentValue = nextValue
        Else
            currentValue = nextValue
        End If
    Next pathIndex

    If VBA.IsObject(currentValue) Then
        Set outValue = currentValue
    Else
        outValue = currentValue
    End If
    private_TryReadPathValue = True
End Function

Private Function private_TryReadMemberValue(ByVal itemObj As Variant, ByVal memberName As String, ByRef outValue As Variant) As Boolean
    Dim valueObj As Object
    Dim memberValue As Variant

    memberName = VBA.Trim$(memberName)
    If VBA.Len(memberName) = 0 Then Exit Function
    If Not VBA.IsObject(itemObj) Then Exit Function
    Set valueObj = itemObj

    On Error Resume Next
    If VBA.TypeName(valueObj) = "Dictionary" Or VBA.TypeName(valueObj) = "Scripting.Dictionary" Then
        If Not valueObj.Exists(memberName) Then
            On Error GoTo 0
            Exit Function
        End If

        Set outValue = valueObj(memberName)
        If Err.Number = 0 Then
            private_TryReadMemberValue = True
            On Error GoTo 0
            Exit Function
        End If
        Err.Clear

        outValue = valueObj(memberName)
    Else
        Set outValue = VBA.CallByName(valueObj, memberName, VbGet)
        If Err.Number = 0 Then
            private_TryReadMemberValue = True
            On Error GoTo 0
            Exit Function
        End If
        Err.Clear

        memberValue = VBA.CallByName(valueObj, memberName, VbGet)
        If VBA.IsObject(memberValue) Then
            Set outValue = memberValue
        Else
            outValue = memberValue
        End If
    End If
    private_TryReadMemberValue = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function private_BuildShapeName(ByVal itemIndex As Long) As String
    private_BuildShapeName = private_ShapeNamePrefix() & VBA.CStr(itemIndex)
End Function

Private Function private_ShapeNamePrefix() As String
    If VBA.Len(VBA.Trim$(m_ShapePrefix)) > 0 Then
        private_ShapeNamePrefix = VBA.Trim$(m_ShapePrefix)
        Exit Function
    End If

    private_ShapeNamePrefix = "btn_BG" & private_Hash4(m_ControlName) & "_"
End Function

Private Function private_Hash4(ByVal valueText As String) As String
    Dim hashValue As Long
    Dim charIndex As Long

    hashValue = 5381
    For charIndex = 1 To VBA.Len(valueText)
        hashValue = ((hashValue * 33) Xor VBA.AscW(VBA.Mid$(valueText, charIndex, 1))) And &H7FFF&
    Next charIndex
    private_Hash4 = VBA.Right$("0000" & VBA.Hex$(hashValue), 4)
End Function

Private Function private_GetOrCreateShape(ByVal ws As Worksheet, ByVal shapeName As String, ByVal targetRange As Range) As Shape
    Dim shp As Shape

    If ws Is Nothing Then Exit Function
    If targetRange Is Nothing Then Exit Function

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, targetRange.Left, targetRange.Top, targetRange.Width, targetRange.Height)
        shp.Name = shapeName
    Else
        ' В retained render shape обычно уже на нужном месте. Избегаем лишних
        ' COM-записей в Excel object model: они заметно дороже обычных сравнений.
        If Not private_DoublesClose(shp.Left, targetRange.Left) Then shp.Left = targetRange.Left
        If Not private_DoublesClose(shp.Top, targetRange.Top) Then shp.Top = targetRange.Top
        If Not private_DoublesClose(shp.Width, targetRange.Width) Then shp.Width = targetRange.Width
        If Not private_DoublesClose(shp.Height, targetRange.Height) Then shp.Height = targetRange.Height
    End If
    If shp.Placement <> xlMoveAndSize Then shp.Placement = xlMoveAndSize

    Set private_GetOrCreateShape = shp
End Function

Private Sub private_ApplyShapeContent(ByVal shp As Shape, ByVal captionText As String)
    Dim currentText As String

    If shp Is Nothing Then Exit Sub

    On Error Resume Next
    currentText = VBA.CStr(shp.TextFrame2.TextRange.Text)
    If VBA.StrComp(currentText, captionText, VBA.vbBinaryCompare) <> 0 Then
        shp.TextFrame2.TextRange.Text = captionText
        shp.TextFrame.Characters.Text = captionText
    End If
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp.TextFrame.VerticalAlignment = xlVAlignCenter
    On Error GoTo 0
End Sub

Private Function private_DoublesClose(ByVal leftValue As Double, ByVal rightValue As Double) As Boolean
    private_DoublesClose = (VBA.Abs(leftValue - rightValue) < 0.05)
End Function

Private Function private_SetShapeMeta( _
    ByVal shp As Shape, _
    ByVal styleName As String, _
    Optional ByVal renderSignature As String = VBA.vbNullString _
) As Boolean
    Dim metaMap As Object

    If shp Is Nothing Then Exit Function
    Set metaMap = VBA.CreateObject("Scripting.Dictionary")
    metaMap.CompareMode = 1
    metaMap("pn.control") = m_ControlName
    If VBA.Len(VBA.Trim$(styleName)) > 0 Then metaMap("pn.style") = VBA.Trim$(styleName)
    If VBA.Len(renderSignature) > 0 Then metaMap("pn.renderSignature") = renderSignature
    private_SetShapeMeta = ex_ShapeMetaRuntime.fn_TrySetShapeMetaValues(shp, metaMap)
End Function

Private Function private_BuildItemVisualSignature( _
    ByVal itemIndex As Long, _
    ByVal rowStart As Long, _
    ByVal colStart As Long, _
    ByVal rowEnd As Long, _
    ByVal colEnd As Long, _
    ByVal captionText As String, _
    ByVal styleName As String _
) As String
    ' itemIndex нужен, чтобы перестановка элементов считалась изменением даже
    ' при совпадающих caption. Bounds фиксируют фактическое место элемента.
    private_BuildItemVisualSignature = VBA.LCase$(VBA.Trim$(m_ControlName)) & "|" & _
        VBA.CStr(itemIndex) & "|" & _
        VBA.CStr(rowStart) & ":" & VBA.CStr(colStart) & ":" & _
        VBA.CStr(rowEnd) & ":" & VBA.CStr(colEnd) & "|" & _
        VBA.Trim$(styleName) & "|" & captionText
End Function

Private Function private_AssignShapeOnAction(ByVal shp As Shape, ByVal macroRef As String) As Boolean
    If shp Is Nothing Then Exit Function
    macroRef = VBA.Trim$(macroRef)
    If VBA.Len(macroRef) = 0 Then Exit Function

    On Error Resume Next
    If VBA.StrComp(VBA.Trim$(shp.OnAction), macroRef, VBA.vbBinaryCompare) <> 0 Then shp.OnAction = macroRef
    private_AssignShapeOnAction = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub private_DeleteExtraShapes(ByVal ws As Worksheet, ByVal firstExtraIndex As Long)
    Dim shapeIndex As Long
    Dim shapeName As String

    If ws Is Nothing Then Exit Sub
    If firstExtraIndex <= 0 Then firstExtraIndex = 1

    shapeIndex = firstExtraIndex
    Do
        shapeName = private_BuildShapeName(shapeIndex)
        On Error Resume Next
        ws.Shapes(shapeName).Delete
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Do
        End If
        Err.Clear
        On Error GoTo 0
        shapeIndex = shapeIndex + 1
    Loop
End Sub

Private Function private_GetRuntimeCallbackMacroRef() As String
    private_GetRuntimeCallbackMacroRef = private_QualifyMacroName("rt_Bridge.fn_OnShapeClick")
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
