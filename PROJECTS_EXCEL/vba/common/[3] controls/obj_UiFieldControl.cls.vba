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
Implements obj_IUiEventHandler

Private Const SELECT_SHAPE_PREFIX As String = "sel_"
Private m_renderContext As obj_UiRenderContext
Private m_uiControlBase As obj_UiControlBase
Private m_targetRange As Range
Private m_sourceName As String
Private m_bindingPath As String
Private m_items As String
Private m_itemsSourceRaw As String
Private m_changeCommandRaw As String
Private m_changeCommand As obj_UiCommand
Private m_isSelect As Boolean
Private m_isCheckbox As Boolean
Private m_readOnly As Boolean
Private WithEvents m_cellBinding As obj_UiCellBinding
Private m_selectAction As obj_UiSelectShapeAction
Private m_checkboxShape As Shape
Private m_isDisposed As Boolean

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
    Set m_uiControlBase = New obj_UiControlBase
End Sub

Private Sub Class_Terminate()
    Me.Dispose
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
    Me.Dispose
End Sub

Private Function obj_IUiControl_Configure( _
    ByVal definition As Object, _
    ByVal context As obj_UiRenderContext, _
    ByVal source As String, _
    ByRef diagnostic As String _
) As Boolean
    Set m_renderContext = context
    m_uiControlBase.ConfigurePosition definition
    obj_IUiControl_Configure = private_Configure(definition)
    If Not obj_IUiControl_Configure Then diagnostic = "Cannot configure control: " & ex_UiElementFactory.fn_Attribute(definition, "name")
End Function

Private Function obj_IUiControl_Measure( _
    ByRef rows As Long, _
    ByRef columns As Long, _
    ByRef diagnostic As String _
) As Boolean
    Dim target As Range

    m_uiControlBase.SetPosition 1, 1
    If Not private_Configure(m_uiControlBase.ControlNode) Then Exit Function
    Set target = private_Measure(m_renderContext)
    If target Is Nothing Then
        diagnostic = "Cannot measure control: " & m_uiControlBase.ControlName
        Exit Function
    End If
    m_uiControlBase.GetSize target, rows, columns
    obj_IUiControl_Measure = True
End Function

Private Function obj_IUiControl_Arrange( _
    ByVal row As Long, _
    ByVal column As Long, _
    ByRef diagnostic As String _
) As Boolean
    Dim target As Range

    m_uiControlBase.ArrangePosition row, column
    If Not private_Configure(m_uiControlBase.ControlNode) Then Exit Function
    Set target = private_Measure(m_renderContext)
    obj_IUiControl_Arrange = Not target Is Nothing
    If target Is Nothing Then diagnostic = "Cannot arrange control: " & m_uiControlBase.ControlName
End Function

Private Function obj_IUiControl_Render(ByRef diagnostic As String) As Boolean
    obj_IUiControl_Render = private_Render(m_renderContext)
    If Not obj_IUiControl_Render Then diagnostic = "Cannot render control: " & m_uiControlBase.ControlName
End Function

Private Function obj_IUiControl_Validate(ByVal errors As Collection) As Boolean
    obj_IUiControl_Validate = True
End Function

Private Function obj_IUiEventHandler_HandleEvent( _
    ByVal kind As String, _
    ByVal payload As Variant _
) As Boolean
    Dim target As Range
    Dim previousEvents As Boolean

    If kind <> "click" And kind <> "change" Then Exit Function
    If kind = "click" Then
        If m_checkboxShape Is Nothing Then Exit Function
        If m_readOnly Then
            obj_IUiEventHandler_HandleEvent = True
            Exit Function
        End If
        Set target = m_targetRange.Cells(1, 1)
        previousEvents = Application.EnableEvents
        On Error GoTo EH
        Application.EnableEvents = False
        target.Value2 = (m_checkboxShape.ControlFormat.Value = xlOn)
        Application.EnableEvents = previousEvents
    Else
        Set target = payload
        If m_isCheckbox And VBA.VarType(target.Value2) <> VBA.vbBoolean Then
            m_cellBinding.SetTwoWay False
            m_cellBinding.HandleCellChange target
            m_cellBinding.SetTwoWay Not m_readOnly
            Exit Function
        End If
    End If
    obj_IUiEventHandler_HandleEvent = m_cellBinding.HandleCellChange(target)
    Exit Function
EH:
    Application.EnableEvents = previousEvents
    VBA.Err.Raise VBA.Err.Number, "Input.HandleEvent", VBA.Err.Description
End Function

' //
' // API
' //
Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    If Not m_cellBinding Is Nothing Then m_cellBinding.Dispose
    If Not m_selectAction Is Nothing Then m_selectAction.Dispose
    Set m_cellBinding = Nothing
    Set m_selectAction = Nothing
    Set m_checkboxShape = Nothing
    If Not m_uiControlBase Is Nothing Then m_uiControlBase.Dispose
    Set m_uiControlBase = Nothing
    Set m_renderContext = Nothing
    Set m_targetRange = Nothing
    m_sourceName = VBA.vbNullString
    m_bindingPath = VBA.vbNullString
    m_items = VBA.vbNullString
    m_itemsSourceRaw = VBA.vbNullString
    m_changeCommandRaw = VBA.vbNullString
    Set m_changeCommand = Nothing
End Sub

' //
' // Private
' //
Private Function private_Configure(ByVal controlNode As Object) As Boolean
    Dim controlType As String
    Dim inputType As String
    Dim rawValue As String
    Dim defaultSource As String

    If Not m_uiControlBase.Configure(controlNode) Then Exit Function
    controlType = VBA.LCase$(private_ReadAttribute(controlNode, "type"))
    inputType = VBA.LCase$(private_ReadAttribute(controlNode, "inputType"))
    m_isSelect = (controlType = "select" Or inputType = "select")
    m_isCheckbox = (inputType = "checkbox")
    m_readOnly = (VBA.LCase$(private_ReadAttribute(controlNode, "readOnly")) = "true")
    m_items = private_ReadAttribute(controlNode, "items")
    m_itemsSourceRaw = private_ReadAttribute(controlNode, "itemsSource")
    If m_isSelect And VBA.Len(VBA.Trim$(m_items)) = 0 And _
       VBA.Len(VBA.Trim$(m_itemsSourceRaw)) = 0 Then Exit Function
    rawValue = private_ReadAttribute(controlNode, "value")
    defaultSource = VBA.vbNullString
    If Not private_TryParseBinding(rawValue, defaultSource, m_sourceName, m_bindingPath) Then
        ex_WindowsUi.fn_ShowMessage "Invalid field binding. Specify Source or a qualified Path; " & _
            "an unqualified Path requires source on the nearest form: " & rawValue, _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    m_changeCommandRaw = private_ReadAttribute(controlNode, "onChange")
    private_Configure = True
End Function

Private Function private_Measure(ByVal uiRenderContext As obj_UiRenderContext) As Range
    If m_uiControlBase Is Nothing Then Exit Function
    Set m_targetRange = m_uiControlBase.Measure(uiRenderContext)
    Set private_Measure = m_targetRange
End Function

Private Function private_Render(ByVal uiRenderContext As obj_UiRenderContext) As Boolean
    Dim value As Variant
    Dim sourceObject As Object
    Dim isObject As Boolean
    Dim uiCellBinding As obj_UiCellBinding
    Dim targetCell As Range
    Dim selectItems As Collection

    Dim shapeNameToRemove As Variant

    If Not m_cellBinding Is Nothing Then m_cellBinding.Dispose
    Set m_cellBinding = Nothing
    If Not m_selectAction Is Nothing Then
        For Each shapeNameToRemove In m_selectAction.ShapeNames
            uiRenderContext.Router.UnregisterShape VBA.CStr(shapeNameToRemove)
            uiRenderContext.TargetWorksheet.Shapes(VBA.CStr(shapeNameToRemove)).Delete
        Next shapeNameToRemove
        m_selectAction.Dispose
        Set m_selectAction = Nothing
    End If
    If Not m_checkboxShape Is Nothing Then
        uiRenderContext.Router.UnregisterShape m_checkboxShape.Name
        m_checkboxShape.Delete
        Set m_checkboxShape = Nothing
    End If
    If m_targetRange Is Nothing Then Set m_targetRange = private_Measure(uiRenderContext)
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
    If m_isCheckbox And VBA.VarType(value) <> VBA.vbBoolean Then
        ex_WindowsUi.fn_ShowMessage "A checkbox binding requires a Boolean value.", VBA.vbExclamation, "Field"
        Exit Function
    End If
    m_targetRange.Value2 = value
    m_targetRange.Locked = m_readOnly
    If Not m_isSelect And Not m_isCheckbox Then m_targetRange.WrapText = True
    If m_isSelect Then
        m_targetRange.NumberFormat = ";;;"
        If Not private_TryResolveItems(uiRenderContext, selectItems) Then Exit Function
    Else
        uiRenderContext.Styles.ApplyControlStyle m_targetRange, Nothing, _
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
    uiCellBinding.SetTwoWay Not m_readOnly
    Set m_cellBinding = uiCellBinding
    uiRenderContext.Router.RegisterCell targetCell, Me
    If m_isSelect Then
        If Not private_RenderSelectShapes( _
                uiRenderContext, selectItems, VBA.CStr(value), uiCellBinding) Then Exit Function
    End If
    If m_isCheckbox Then
        Set m_checkboxShape = uiRenderContext.TargetWorksheet.Shapes.AddFormControl( _
            xlCheckBox, targetCell.Left, targetCell.Top, 20, 18)
        m_checkboxShape.Name = "chk_" & VBA.CStr(uiRenderContext.Router.NextId())
        m_checkboxShape.TextFrame.Characters.Text = VBA.vbNullString
        m_checkboxShape.ControlFormat.Enabled = Not m_readOnly
        m_checkboxShape.ControlFormat.Value = VBA.IIf(VBA.CBool(value), xlOn, xlOff)
        m_checkboxShape.OnAction = "ex_UiBridge.fn_OnShapeClick"
        m_targetRange.NumberFormat = ";;;"
        uiRenderContext.Router.RegisterShape m_checkboxShape.Name, Me
    End If
    private_Render = True
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
    controlId = uiRenderContext.Router.NextId()
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
    uiRenderContext.Styles.ApplyControlStyle Nothing, headerShape, _
        m_uiControlBase.ControlNode, uiRenderContext.BindingContext
    headerShape.OnAction = "ex_UiBridge.fn_OnShapeClick"

    Set panelShape = targetWorksheet.Shapes.AddShape( _
        msoShapeRectangle, m_targetRange.Left, m_targetRange.Top + m_targetRange.Height, _
        m_targetRange.Width, items.Count * itemHeight + (items.Count - 1) * itemMargin)
    panelShape.Name = panelShapeName
    private_ApplyDefaultShapeStyle panelShape
    uiRenderContext.Styles.ApplyControlPartStyle panelShape, _
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
        uiRenderContext.Styles.ApplyControlPartStyle itemShape, _
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
    Set m_selectAction = uiSelectAction
    Dim routeName As Variant

    For Each routeName In uiSelectAction.ShapeNames
        uiRenderContext.Router.RegisterShape VBA.CStr(routeName), uiSelectAction
    Next routeName
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

Private Function private_ReadAttribute( _
    ByVal node As Object, _
    ByVal attributeName As String _
) As String
    Dim value As Variant

    value = node.getAttribute(attributeName)
    If Not VBA.IsNull(value) And Not VBA.IsEmpty(value) Then private_ReadAttribute = VBA.CStr(value)
End Function

Private Sub m_cellBinding_ValueRefreshed()
    private_cellBinding_ValueRefreshed
End Sub

Private Sub private_cellBinding_ValueRefreshed()
    If Not m_checkboxShape Is Nothing Then _
        m_checkboxShape.ControlFormat.Value = VBA.IIf(VBA.CBool(m_targetRange.Cells(1, 1).Value2), xlOn, xlOff)
End Sub