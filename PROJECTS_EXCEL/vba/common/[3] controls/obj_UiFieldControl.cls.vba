VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_UiFieldControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiControl

Private Const SELECT_SHAPE_PREFIX As String = "sel_"
Private m_uiControlBase As obj_UiControlBase
Private m_targetRange As Range
Private m_sourceName As String
Private m_bindingPath As String
Private m_items As String
Private m_itemsSourceRaw As String
Private m_changeCommandRaw As String
Private m_changeCommand As obj_UiCommand
Private m_isSelect As Boolean
Private m_isDisposed As Boolean

Private Sub Class_Initialize()
    Set m_uiControlBase = New obj_UiControlBase
End Sub

Private Sub Class_Terminate()
    obj_IUiControl_Dispose
End Sub

' //
' // Interface
' //
Private Function obj_IUiControl_Initialize() As Boolean
    m_isDisposed = False
    If m_uiControlBase Is Nothing Then Set m_uiControlBase = New obj_UiControlBase
    obj_IUiControl_Initialize = m_uiControlBase.Initialize()
End Function

Private Sub obj_IUiControl_Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    If Not m_uiControlBase Is Nothing Then m_uiControlBase.Dispose
    Set m_uiControlBase = Nothing
    Set m_targetRange = Nothing
    m_sourceName = VBA.vbNullString
    m_bindingPath = VBA.vbNullString
    m_items = VBA.vbNullString
    m_itemsSourceRaw = VBA.vbNullString
    m_changeCommandRaw = VBA.vbNullString
    Set m_changeCommand = Nothing
End Sub

Private Function obj_IUiControl_Configure(ByVal controlNode As Object) As Boolean
    Dim controlType As String
    Dim inputType As String
    Dim rawValue As String
    Dim defaultSource As String

    If Not m_uiControlBase.Configure(controlNode) Then Exit Function
    controlType = VBA.LCase$(private_ReadAttribute(controlNode, "type"))
    inputType = VBA.LCase$(private_ReadAttribute(controlNode, "inputType"))
    m_isSelect = (controlType = "select" Or inputType = "select")
    m_items = private_ReadAttribute(controlNode, "items")
    m_itemsSourceRaw = private_ReadAttribute(controlNode, "itemsSource")
    If m_isSelect And VBA.Len(VBA.Trim$(m_items)) = 0 And _
       VBA.Len(VBA.Trim$(m_itemsSourceRaw)) = 0 Then Exit Function
    rawValue = private_ReadAttribute(controlNode, "value")
    defaultSource = private_GetFormSource(controlNode)
    If Not private_TryParseBinding(rawValue, defaultSource, m_sourceName, m_bindingPath) Then
        ex_WindowsUi.fn_ShowMessage "Invalid field binding. Specify Source or a qualified Path; " & _
            "an unqualified Path requires source on the nearest form: " & rawValue, _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    m_changeCommandRaw = private_ReadAttribute(controlNode, "onChange")
    obj_IUiControl_Configure = True
End Function

Private Function obj_IUiControl_Measure(ByVal uiRenderContext As obj_UiRenderContext) As Range
    If m_uiControlBase Is Nothing Then Exit Function
    Set m_targetRange = m_uiControlBase.Measure(uiRenderContext)
    Set obj_IUiControl_Measure = m_targetRange
End Function

Private Function obj_IUiControl_Render(ByVal uiRenderContext As obj_UiRenderContext) As Boolean
    Dim value As Variant
    Dim sourceObject As Object
    Dim isObject As Boolean
    Dim uiCellBinding As obj_UiCellBinding
    Dim targetCell As Range
    Dim selectItems As Collection

    If m_targetRange Is Nothing Then Set m_targetRange = obj_IUiControl_Measure(uiRenderContext)
    If m_targetRange Is Nothing Then Exit Function
    If m_targetRange.Cells.CountLarge > 1 Then m_targetRange.Merge
    If Not uiRenderContext.BindingContext.TryGetValue( _
            m_sourceName, m_bindingPath, value, sourceObject, isObject) Then Exit Function
    If isObject Then Exit Function
    If m_isSelect Then
        On Error Resume Next
        m_targetRange.Cells(1, 1).Validation.Delete
        On Error GoTo 0
    End If
    If VBA.Len(m_changeCommandRaw) > 0 Then
        If Not ex_UiBindingRuntime.fn_TryResolveCommand( _
                m_changeCommandRaw, uiRenderContext.BindingContext, m_changeCommand) Then Exit Function
    End If
    m_targetRange.Value2 = value
    If m_isSelect Then
        m_targetRange.NumberFormat = ";;;"
        If Not private_TryResolveItems(uiRenderContext, selectItems) Then Exit Function
    Else
        ex_StylePipeline.fn_ApplyControlStyle m_targetRange, Nothing, _
            m_uiControlBase.ControlNode, uiRenderContext.BindingContext
        On Error Resume Next
        m_targetRange.Cells(1, 1).Validation.Delete
        On Error GoTo 0
    End If

    Set targetCell = m_targetRange.Cells(1, 1)
    Set uiCellBinding = New obj_UiCellBinding
    If Not uiCellBinding.Initialize( _
            targetCell.Parent.Name, targetCell.Address(False, False), _
            uiRenderContext.BindingContext, m_sourceName, m_bindingPath, m_changeCommand) Then Exit Function
    If Not ex_UiBindings.fn_RegisterCellBinding(uiCellBinding) Then Exit Function
    If m_isSelect Then
        If Not private_RenderSelectShapes( _
                uiRenderContext, selectItems, VBA.CStr(value), uiCellBinding) Then Exit Function
    End If
    obj_IUiControl_Render = True
End Function

Private Function obj_IUiControl_HandleCellChange(ByVal target As Range) As Boolean
    obj_IUiControl_HandleCellChange = False
End Function

' //
' // Private
' //
Private Function private_TryResolveItems( _
    ByVal uiRenderContext As obj_UiRenderContext, _
    ByRef outItems As Collection _
) As Boolean
    Dim value As Variant
    Dim sourceObject As Object
    Dim isObject As Boolean
    Dim item As Variant

    Set outItems = New Collection
    If VBA.Len(VBA.Trim$(m_itemsSourceRaw)) = 0 Then
        private_AddDelimitedItems m_items, outItems
        private_TryResolveItems = (outItems.Count > 0)
        Exit Function
    End If
    If Not ex_UiBindingRuntime.fn_TryResolveValue( _
            m_itemsSourceRaw, uiRenderContext.BindingContext, value, sourceObject, isObject) Then Exit Function
    If isObject Then
        If TypeName(sourceObject) <> "Collection" Then Exit Function
        For Each item In sourceObject
            outItems.Add VBA.CStr(item)
        Next item
    ElseIf VBA.IsArray(value) Then
        For Each item In value
            outItems.Add VBA.CStr(item)
        Next item
    Else
        private_AddDelimitedItems VBA.CStr(value), outItems
    End If
    private_TryResolveItems = (outItems.Count > 0)
End Function

Private Sub private_AddDelimitedItems(ByVal listText As String, ByVal outItems As Collection)
    Dim item As Variant
    Dim itemText As String

    For Each item In VBA.Split(listText, ",")
        itemText = VBA.Trim$(VBA.CStr(item))
        If VBA.Len(itemText) > 0 Then outItems.Add itemText
    Next item
End Sub

Private Function private_RenderSelectShapes( _
    ByVal uiRenderContext As obj_UiRenderContext, _
    ByVal items As Collection, _
    ByVal selectedValue As String, _
    ByVal uiCellBinding As obj_UiCellBinding _
) As Boolean
    Dim targetWorksheet As Worksheet
    Dim headerShape As Shape
    Dim panelShape As Shape
    Dim itemShape As Shape
    Dim itemShapeNames As Collection
    Dim shapeNames As Collection
    Dim uiSelectAction As obj_UiSelectShapeAction
    Dim controlId As Long
    Dim headerShapeName As String
    Dim panelShapeName As String
    Dim itemShapeName As String
    Dim itemText As String
    Dim arrowText As String
    Dim itemHeight As Double
    Dim itemMargin As Double
    Dim itemTop As Double
    Dim itemIndex As Long

    If items Is Nothing Or uiCellBinding Is Nothing Then Exit Function
    If items.Count = 0 Then Exit Function
    Set targetWorksheet = uiRenderContext.TargetWorksheet
    controlId = ex_UiBindings.fn_NextSelectControlId()
    headerShapeName = SELECT_SHAPE_PREFIX & "h_" & VBA.CStr(controlId)
    panelShapeName = SELECT_SHAPE_PREFIX & "p_" & VBA.CStr(controlId)
    itemHeight = m_targetRange.Height
    itemMargin = 0
    If VBA.IsNumeric(private_ReadAttribute(m_uiControlBase.ControlNode, "itemHeight")) Then
        itemHeight = VBA.CDbl(private_ReadAttribute(m_uiControlBase.ControlNode, "itemHeight"))
    End If
    If VBA.IsNumeric(private_ReadAttribute(m_uiControlBase.ControlNode, "itemMargin")) Then
        itemMargin = VBA.CDbl(private_ReadAttribute(m_uiControlBase.ControlNode, "itemMargin"))
    End If
    If itemHeight <= 0 Or itemMargin < 0 Then Exit Function

    Set headerShape = targetWorksheet.Shapes.AddShape( _
        msoShapeRectangle, m_targetRange.Left, m_targetRange.Top, _
        m_targetRange.Width, m_targetRange.Height)
    headerShape.Name = headerShapeName
    private_ApplyDefaultShapeStyle headerShape
    arrowText = VBA.ChrW(&H25BC)
    headerShape.TextFrame2.TextRange.Text = selectedValue & " " & arrowText
    headerShape.TextFrame2.MarginLeft = 5
    headerShape.TextFrame2.MarginRight = 5
    headerShape.TextFrame2.MarginTop = 0
    headerShape.TextFrame2.MarginBottom = 0
    ex_StylePipeline.fn_ApplyControlStyle Nothing, headerShape, _
        m_uiControlBase.ControlNode, uiRenderContext.BindingContext
    headerShape.OnAction = "ex_UiBridge.fn_OnShapeClick"

    Set panelShape = targetWorksheet.Shapes.AddShape( _
        msoShapeRectangle, m_targetRange.Left, m_targetRange.Top + m_targetRange.Height, _
        m_targetRange.Width, items.Count * itemHeight + (items.Count - 1) * itemMargin)
    panelShape.Name = panelShapeName
    private_ApplyDefaultShapeStyle panelShape
    ex_StylePipeline.fn_ApplyControlPartStyle panelShape, _
        m_uiControlBase.ControlNode, uiRenderContext.BindingContext, "panelStyle"
    panelShape.OnAction = "ex_UiBridge.fn_OnShapeClick"
    panelShape.Visible = msoFalse

    Set itemShapeNames = New Collection
    Set shapeNames = New Collection
    shapeNames.Add headerShapeName
    shapeNames.Add panelShapeName
    For itemIndex = 1 To items.Count
        itemShapeName = SELECT_SHAPE_PREFIX & "i_" & VBA.CStr(controlId) & "_" & VBA.CStr(itemIndex)
        itemTop = panelShape.Top + (itemIndex - 1) * (itemHeight + itemMargin)
        Set itemShape = targetWorksheet.Shapes.AddShape( _
            msoShapeRectangle, panelShape.Left, itemTop, panelShape.Width, itemHeight)
        itemShape.Name = itemShapeName
        private_ApplyDefaultShapeStyle itemShape
        itemText = VBA.CStr(items(itemIndex))
        itemShape.TextFrame2.TextRange.Text = itemText
        itemShape.TextFrame2.MarginLeft = 5
        itemShape.TextFrame2.MarginRight = 5
        itemShape.TextFrame2.MarginTop = 0
        itemShape.TextFrame2.MarginBottom = 0
        ex_StylePipeline.fn_ApplyControlPartStyle itemShape, _
            m_uiControlBase.ControlNode, uiRenderContext.BindingContext, "itemStyle"
        itemShape.OnAction = "ex_UiBridge.fn_OnShapeClick"
        itemShape.Visible = msoFalse
        itemShapeNames.Add itemShapeName
        shapeNames.Add itemShapeName
    Next itemIndex

    Set uiSelectAction = New obj_UiSelectShapeAction
    If Not uiSelectAction.Initialize( _
            controlId, m_targetRange.Cells(1, 1), headerShapeName, panelShapeName, _
            itemShapeNames, shapeNames, items, selectedValue, uiCellBinding) Then Exit Function
    If Not ex_UiBindings.fn_RegisterSelectControl(uiSelectAction) Then Exit Function
    private_RenderSelectShapes = True
End Function

Private Sub private_ApplyDefaultShapeStyle(ByVal targetShape As Shape)
    On Error Resume Next
    targetShape.Fill.Solid
    targetShape.Fill.ForeColor.RGB = VBA.RGB(255, 255, 255)
    targetShape.Line.Visible = msoTrue
    targetShape.Line.ForeColor.RGB = VBA.RGB(100, 116, 139)
    targetShape.Line.Weight = 0.75
    targetShape.TextFrame2.TextRange.Font.Name = "Calibri"
    targetShape.TextFrame2.TextRange.Font.Size = 11
    targetShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = VBA.RGB(17, 24, 39)
    targetShape.TextFrame2.VerticalAnchor = msoAnchorMiddle
    targetShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    On Error GoTo 0
End Sub

Private Function private_TryParseBinding( _
    ByVal rawBinding As String, _
    ByVal defaultSource As String, _
    ByRef outSourceName As String, _
    ByRef outBindingPath As String _
) As Boolean
    private_TryParseBinding = ex_UiBindingRuntime.fn_TryParseBinding( _
        rawBinding, defaultSource, outSourceName, outBindingPath)
End Function

Private Function private_GetFormSource(ByVal controlNode As Object) As String
    Dim currentNode As Object
    Dim nodeName As String

    Set currentNode = controlNode.parentNode
    Do While Not currentNode Is Nothing
        nodeName = VBA.LCase$(VBA.CStr(currentNode.baseName))
        If nodeName = "form" Then
            private_GetFormSource = private_ReadAttribute(currentNode, "source")
            Exit Function
        End If
        Set currentNode = currentNode.parentNode
    Loop
End Function

Private Function private_ReadAttribute(ByVal node As Object, ByVal attributeName As String) As String
    Dim value As Variant
    value = node.getAttribute(attributeName)
    If Not VBA.IsNull(value) And Not VBA.IsEmpty(value) Then private_ReadAttribute = VBA.CStr(value)
End Function