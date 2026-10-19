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

#Const ENABLE_LOGGING_VERBOSE = False

Implements obj_IUiControl
Implements obj_IUiEventHandler
Implements obj_IUiCellRefreshListener
Implements obj_IUiBindingTarget

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_appliedStyleName As String
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
Private m_editable As Boolean
Private m_isCheckbox As Boolean
Private m_readOnly As Boolean
Private m_appliedReadOnly As Boolean
Private m_cellBinding As obj_UiCellBinding
Private m_selectAction As obj_UiSelectShapeAction
Private m_checkboxShape As Shape

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
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    If m_uiControlBase Is Nothing Then Set m_uiControlBase = New obj_UiControlBase
    obj_IUiControl_Initialize = m_uiControlBase.Initialize()
    m_isInitialized = obj_IUiControl_Initialize
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

Private Function obj_IUiBindingTarget_RefreshBindings( _
    ByVal context As obj_UiRenderContext _
) As Boolean
    Dim nextStyle As String

    If Not private_Configure(m_uiControlBase.ControlNode) Then Exit Function
    If m_targetRange Is Nothing Or m_cellBinding Is Nothing Then Exit Function
    m_targetRange.Locked = m_readOnly
    m_appliedReadOnly = m_readOnly
    m_cellBinding.SetTwoWay Not m_readOnly
    If Not m_checkboxShape Is Nothing Then m_checkboxShape.ControlFormat.Enabled = Not m_readOnly
    If Not ex_UiBindingRuntime.fn_TryResolveText( _
        ex_UiElementFactory.fn_Attribute(m_uiControlBase.ControlNode, "style"), context.BindingContext, nextStyle) Then Exit Function
    If Not context.Styles.TryApplyFontOnly(m_targetRange, m_appliedStyleName, nextStyle) Then
        context.Styles.RestoreControlArea m_targetRange, m_uiControlBase.ControlNode, context.BindingContext
    End If
    m_appliedStyleName = nextStyle
    private_RegisterStatePart
    obj_IUiBindingTarget_RefreshBindings = True
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
    Dim shapeName As Variant

    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    If Not m_renderContext Is Nothing Then
        If Not m_targetRange Is Nothing Then m_renderContext.Router.UnregisterCell m_targetRange.Cells(1, 1)
        m_renderContext.Router.UnregisterSelection m_uiControlBase.ControlName
        If Not m_selectAction Is Nothing Then
            For Each shapeName In m_selectAction.ShapeNames
                m_renderContext.Router.UnregisterShape VBA.CStr(shapeName)
                m_renderContext.UnregisterShapeLayout VBA.CStr(shapeName)
                m_renderContext.TargetWorksheet.Shapes(VBA.CStr(shapeName)).Delete
            Next shapeName
        End If
        If Not m_checkboxShape Is Nothing Then
            m_renderContext.Router.UnregisterShape m_checkboxShape.Name
            m_checkboxShape.Delete
        End If
    End If
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

Public Function RefreshReadOnly() As Boolean
    Dim raw As String
    Dim resolved As String
    Dim readOnly As Boolean

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If m_targetRange Is Nothing Or m_cellBinding Is Nothing Then Exit Function
    raw = private_ReadAttribute(m_uiControlBase.ControlNode, "readOnly")
    If VBA.Len(raw) = 0 Then raw = "false"
    If Not ex_UiBindingRuntime.fn_TryResolveText(raw, m_renderContext.BindingContext, resolved) Then Exit Function
    resolved = VBA.LCase$(resolved)
    If resolved <> "true" And resolved <> "false" Then VBA.Err.Raise 5, , "Invalid input readOnly value."
    readOnly = (resolved = "true")
    If readOnly <> m_appliedReadOnly Then
        m_targetRange.Locked = readOnly
        m_cellBinding.SetTwoWay Not readOnly
        If Not m_checkboxShape Is Nothing Then m_checkboxShape.ControlFormat.Enabled = Not readOnly
        m_appliedReadOnly = readOnly
    End If
    m_readOnly = readOnly
    If Not m_renderContext.Styles.TryApplyFontOnly(m_targetRange, m_appliedStyleName, m_appliedStyleName) Then
        m_renderContext.Styles.RestoreControlArea m_targetRange, m_uiControlBase.ControlNode, m_renderContext.BindingContext
    End If
    private_RegisterStatePart
    RefreshReadOnly = True
End Function

Public Function CreateDropdownShapes( _
    ByVal controlId As Long, _
    ByVal panelShapeName As String, _
    ByVal itemShapeNames As Collection, _
    ByVal shapeNames As Collection, _
    ByVal items As Collection, _
    ByVal uiSelectShapeAction As obj_UiSelectShapeAction _
) As Boolean
    Dim uiRenderContext As obj_UiRenderContext
    Dim targetWorksheet As Worksheet
    Dim panelShape As Shape
    Dim itemShape As Shape
    Dim itemShapeName As String
    Dim itemText As String
    Dim itemHeight As Double
    Dim itemMargin As Double
    Dim itemTop As Double
    Dim itemIndex As Long
    Dim highlightValue As Variant
    Dim highlightObject As Object
    Dim highlightIsObject As Boolean
    Dim highlightRaw As String
    Dim highlightItem As Variant
    Dim isHighlighted As Boolean
    Dim startedAt As Double
    Dim routeName As Variant

    If m_isDisposed Or m_renderContext Is Nothing Then Exit Function
    startedAt = VBA.Timer
    Set uiRenderContext = m_renderContext
    highlightRaw = private_ReadAttribute(m_uiControlBase.ControlNode, "highlightValue")
    If VBA.Len(highlightRaw) > 0 Then
        If Not ex_UiBindingRuntime.fn_TryResolveValue( _
                highlightRaw, uiRenderContext.BindingContext, highlightValue, highlightObject, highlightIsObject) Then Exit Function
        If highlightIsObject Then
            If VBA.TypeName(highlightObject) <> "Collection" Then Exit Function
        Else
            If VBA.IsNull(highlightValue) Or VBA.IsError(highlightValue) Then Exit Function
        End If
    End If
    Set targetWorksheet = uiRenderContext.TargetWorksheet
    itemHeight = m_targetRange.Height
    itemMargin = 0
    If VBA.IsNumeric(private_ReadAttribute(m_uiControlBase.ControlNode, "itemHeight")) Then
        itemHeight = VBA.CDbl(private_ReadAttribute(m_uiControlBase.ControlNode, "itemHeight"))
    End If
    If VBA.IsNumeric(private_ReadAttribute(m_uiControlBase.ControlNode, "itemMargin")) Then
        itemMargin = VBA.CDbl(private_ReadAttribute(m_uiControlBase.ControlNode, "itemMargin"))
    End If
    If itemHeight <= 0 Or itemMargin < 0 Then Exit Function

    Set panelShape = targetWorksheet.Shapes.AddShape( _
        msoShapeRectangle, m_targetRange.Left, m_targetRange.Top + m_targetRange.Height, _
        m_targetRange.Width, items.Count * itemHeight + (items.Count - 1) * itemMargin)
    panelShape.Name = panelShapeName
    shapeNames.Add panelShapeName
    panelShape.Placement = xlMoveAndSize
    private_ApplyDefaultShapeStyle panelShape
    uiRenderContext.Styles.ApplyControlPartStyle panelShape, _
        m_uiControlBase.ControlNode, uiRenderContext.BindingContext, "panelStyle"
    panelShape.OnAction = "ex_UiBridge.fn_OnShapeClick"
    panelShape.Visible = msoFalse

    For itemIndex = 1 To items.Count
        itemShapeName = private_ShapeName(SELECT_SHAPE_PREFIX & "i_" & VBA.CStr(controlId) & "_" & VBA.CStr(itemIndex))
        itemTop = panelShape.Top + (itemIndex - 1) * (itemHeight + itemMargin)
        Set itemShape = targetWorksheet.Shapes.AddShape( _
            msoShapeRectangle, panelShape.Left, itemTop, panelShape.Width, itemHeight)
        itemShape.Name = itemShapeName
        shapeNames.Add itemShapeName
        itemShape.Placement = xlMoveAndSize
        private_ApplyDefaultShapeStyle itemShape
        itemText = VBA.CStr(items(itemIndex))
        itemShape.TextFrame2.TextRange.Text = itemText
        itemShape.TextFrame2.MarginLeft = 5
        itemShape.TextFrame2.MarginRight = 5
        itemShape.TextFrame2.MarginTop = 0
        itemShape.TextFrame2.MarginBottom = 0
        uiRenderContext.Styles.ApplyControlPartStyle itemShape, _
            m_uiControlBase.ControlNode, uiRenderContext.BindingContext, "itemStyle"
        If VBA.Len(highlightRaw) > 0 Then
            isHighlighted = False
            If highlightIsObject Then
                For Each highlightItem In highlightObject
                    If itemText = VBA.CStr(highlightItem) Then
                        isHighlighted = True
                        Exit For
                    End If
                Next highlightItem
            Else
                isHighlighted = itemText = VBA.CStr(highlightValue)
            End If
            If isHighlighted Then
                uiRenderContext.Styles.ApplyControlPartStyle itemShape, _
                    m_uiControlBase.ControlNode, uiRenderContext.BindingContext, "highlightStyle"
            End If
        End If
        itemShape.OnAction = "ex_UiBridge.fn_OnShapeClick"
        itemShape.Visible = msoFalse
        itemShapeNames.Add itemShapeName
    Next itemIndex

    For Each routeName In shapeNames
        uiRenderContext.Router.RegisterShape VBA.CStr(routeName), uiSelectShapeAction
    Next routeName
    CreateDropdownShapes = True
    ex_Core.fn_Diagnostic_WritePerf "Control.Select.CreateDropdown | Name=" & m_uiControlBase.ControlName & _
        " | Items=" & VBA.CStr(items.Count), startedAt
End Function

' //
' // Private
' //
Private Sub private_RegisterStatePart(Optional ByVal applyStyles As Boolean = True)
    Dim tags As String

    If private_ReadAttribute(m_uiControlBase.ControlNode, "partStylesEnabled") = "false" Then Exit Sub
    If m_targetRange Is Nothing Then Exit Sub
    tags = m_appliedStyleName
    If m_readOnly Then tags = tags & " readOnly"
    m_renderContext.Styles.RegisterPart m_uiControlBase.ControlNode, m_targetRange, "editor", tags:=tags
    If applyStyles Then m_renderContext.Styles.ApplyPartStyles m_targetRange
End Sub

Private Function private_Configure(ByVal controlNode As Object) As Boolean
    Dim controlType As String
    Dim inputType As String
    Dim rawValue As String
    Dim defaultSource As String
    Dim readOnlyRaw As String
    Dim readOnlyValue As String

    If Not m_uiControlBase.Configure(controlNode) Then Exit Function
    controlType = VBA.LCase$(private_ReadAttribute(controlNode, "type"))
    inputType = VBA.LCase$(private_ReadAttribute(controlNode, "inputType"))
    m_isSelect = (controlType = "select" Or inputType = "select")
    m_isCheckbox = (inputType = "checkbox")
    readOnlyRaw = private_ReadAttribute(controlNode, "readOnly")
    If VBA.Len(readOnlyRaw) = 0 Then readOnlyRaw = "false"
    If Not ex_UiBindingRuntime.fn_TryResolveText(readOnlyRaw, m_renderContext.BindingContext, readOnlyValue) Then Exit Function
    readOnlyValue = VBA.LCase$(readOnlyValue)
    If readOnlyValue <> "true" And readOnlyValue <> "false" Then Err.Raise 5, , "Invalid input readOnly value."
    m_readOnly = (readOnlyValue = "true")
    m_editable = m_isSelect And (VBA.LCase$(private_ReadAttribute(controlNode, "editable")) = "true")
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
    Dim startedAt As Double
    Dim value As Variant
    Dim sourceObject As Object
    Dim isObject As Boolean
    Dim uiCellBinding As obj_UiCellBinding
    Dim targetCell As Range
    Dim selectItems As Collection
    Dim focusCommand As obj_UiCommand

    Dim shapeNameToRemove As Variant

    startedAt = VBA.Timer
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_FIELD_RENDER_STARTED | Name=" & m_uiControlBase.ControlName
#End If
    If Not m_cellBinding Is Nothing Then m_cellBinding.Dispose
    Set m_cellBinding = Nothing
    If Not m_selectAction Is Nothing Then
        For Each shapeNameToRemove In m_selectAction.ShapeNames
            uiRenderContext.Router.UnregisterShape VBA.CStr(shapeNameToRemove)
            uiRenderContext.UnregisterShapeLayout VBA.CStr(shapeNameToRemove)
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
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_FIELD_RANGE_PREPARE_STARTED | Name=" & m_uiControlBase.ControlName
#End If
    If m_targetRange Is Nothing Then Set m_targetRange = private_Measure(uiRenderContext)
    If m_targetRange Is Nothing Then Exit Function
    If m_targetRange.Cells.CountLarge > 1 Then m_targetRange.Merge
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_FIELD_RANGE_PREPARE_COMPLETED | Name=" & m_uiControlBase.ControlName
#End If
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_FIELD_VALUE_RESOLVE_STARTED | Name=" & m_uiControlBase.ControlName
#End If
    If Not uiRenderContext.BindingContext.TryGetValue( _
            m_sourceName, m_bindingPath, value, sourceObject, isObject) Then Exit Function
    If isObject Then Exit Function
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_FIELD_VALUE_RESOLVE_COMPLETED | Name=" & m_uiControlBase.ControlName
#End If
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
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_FIELD_VALUE_WRITE_STARTED | Name=" & m_uiControlBase.ControlName
#End If
    ' Apply the configured cell format before Excel receives the bound value.
    If Not m_isSelect Or m_editable Then
        uiRenderContext.Styles.ApplyControlStyle m_targetRange, Nothing, _
            m_uiControlBase.ControlNode, uiRenderContext.BindingContext
        If Not ex_UiBindingRuntime.fn_TryResolveText( _
            ex_UiElementFactory.fn_Attribute(m_uiControlBase.ControlNode, "style"), _
            uiRenderContext.BindingContext, m_appliedStyleName) Then Exit Function
    End If
    m_targetRange.Value2 = value
    m_targetRange.Locked = m_readOnly
    m_appliedReadOnly = m_readOnly
    If Not m_isSelect And Not m_isCheckbox Then m_targetRange.WrapText = True
    private_RegisterStatePart False
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_FIELD_VALUE_WRITE_COMPLETED | Name=" & m_uiControlBase.ControlName
#End If
    If m_isSelect Then
        If m_editable Then
            m_targetRange.NumberFormat = "@"
        Else
            m_targetRange.NumberFormat = ";;;"
        End If
        If Not private_TryResolveItems(uiRenderContext, selectItems) Then Exit Function
    Else
        On Error Resume Next
        m_targetRange.Cells(1, 1).Validation.Delete
        On Error GoTo 0
    End If

#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_FIELD_BINDING_NEW_STARTED | Name=" & m_uiControlBase.ControlName
#End If
    Set targetCell = m_targetRange.Cells(1, 1)
    Set uiCellBinding = New obj_UiCellBinding
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_FIELD_BINDING_NEW_COMPLETED | Name=" & m_uiControlBase.ControlName
#End If
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_FIELD_BINDING_INITIALIZE_STARTED | Name=" & m_uiControlBase.ControlName
#End If
    If Not uiCellBinding.Initialize( _
            targetCell.Parent.Name, targetCell.Address(False, False), _
            uiRenderContext.BindingContext, m_sourceName, m_bindingPath, m_changeCommand) Then Exit Function
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_FIELD_BINDING_INITIALIZE_COMPLETED | Name=" & m_uiControlBase.ControlName
#End If
    uiCellBinding.SetTwoWay Not m_readOnly
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_FIELD_REFRESH_SUBSCRIBE_STARTED | Name=" & m_uiControlBase.ControlName
#End If
    Set m_cellBinding = uiCellBinding
    If VBA.Len(private_ReadAttribute(m_uiControlBase.ControlNode, "onFocus")) > 0 Then
        If Not ex_UiBindingRuntime.fn_TryResolveCommand( _
                private_ReadAttribute(m_uiControlBase.ControlNode, "onFocus"), _
                uiRenderContext.BindingContext, focusCommand) Then Exit Function
        uiRenderContext.Router.RegisterSelection m_uiControlBase.ControlName, m_targetRange, focusCommand
    End If
    m_cellBinding.SubscribeRefresh Me
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_FIELD_REFRESH_SUBSCRIBE_COMPLETED | Name=" & m_uiControlBase.ControlName
#End If
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_FIELD_ROUTER_REGISTER_STARTED | Name=" & m_uiControlBase.ControlName
#End If
    uiRenderContext.Router.RegisterCell targetCell, Me
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_FIELD_ROUTER_REGISTER_COMPLETED | Name=" & m_uiControlBase.ControlName
#End If
    If m_isSelect Then
        If Not private_RenderSelectShapes( _
                uiRenderContext, selectItems, VBA.CStr(value), uiCellBinding) Then Exit Function
    End If
    If m_isCheckbox Then
        Set m_checkboxShape = uiRenderContext.TargetWorksheet.Shapes.AddFormControl( _
            xlCheckBox, targetCell.Left, targetCell.Top, 20, 18)
        m_checkboxShape.Name = private_ShapeName("chk_" & VBA.CStr(uiRenderContext.Router.NextId()))
        m_checkboxShape.TextFrame.Characters.Text = VBA.vbNullString
        m_checkboxShape.ControlFormat.Enabled = Not m_readOnly
        m_checkboxShape.ControlFormat.Value = VBA.IIf(VBA.CBool(value), xlOn, xlOff)
        m_checkboxShape.OnAction = "ex_UiBridge.fn_OnShapeClick"
        m_targetRange.NumberFormat = ";;;"
        uiRenderContext.Router.RegisterShape m_checkboxShape.Name, Me
    End If
    private_Render = True
    ex_Core.fn_Diagnostic_WritePerf "Control.Field.Render | Name=" & m_uiControlBase.ControlName, startedAt
#If ENABLE_LOGGING_VERBOSE Then
    ex_Core.fn_Diagnostic_WriteLog "UI_FIELD_RENDER_COMPLETED | Name=" & m_uiControlBase.ControlName
#End If
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
    ' An empty bound collection is valid while lookup results are pending.
    private_TryResolveItems = True
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
    Dim itemShapeNames As Collection
    Dim shapeNames As Collection
    Dim uiSelectAction As obj_UiSelectShapeAction
    Dim controlId As Long
    Dim headerShapeName As String
    Dim panelShapeName As String
    Dim arrowText As String
    Dim placeholder As String
    Dim headerText As String
    Dim startedAt As Double

    startedAt = VBA.Timer
    If items Is Nothing Or uiCellBinding Is Nothing Then Exit Function
    Set targetWorksheet = uiRenderContext.TargetWorksheet
    controlId = uiRenderContext.Router.NextId()
    headerShapeName = private_ShapeName(SELECT_SHAPE_PREFIX & "h_" & VBA.CStr(controlId))
    panelShapeName = private_ShapeName(SELECT_SHAPE_PREFIX & "p_" & VBA.CStr(controlId))

    Set headerShape = targetWorksheet.Shapes.AddShape( _
        msoShapeRectangle, m_targetRange.Left, m_targetRange.Top, _
        m_targetRange.Width, m_targetRange.Height)
    headerShape.Name = headerShapeName
    If m_editable Then
        uiRenderContext.RegisterShapeLayout headerShape, m_targetRange, _
            m_uiControlBase.ControlNode, "1/10", "right"
    Else
        uiRenderContext.RegisterShapeLayout headerShape, m_targetRange, m_uiControlBase.ControlNode
    End If
    private_ApplyDefaultShapeStyle headerShape
    arrowText = VBA.ChrW(&H25BC)
    If Not ex_UiBindingRuntime.fn_TryResolveText( _
            private_ReadAttribute(m_uiControlBase.ControlNode, "placeholder"), uiRenderContext.BindingContext, placeholder) Then Exit Function
    headerText = selectedValue
    If VBA.Len(headerText) = 0 Then headerText = placeholder
    headerShape.TextFrame2.TextRange.Text = headerText
    If items.Count > 0 Then headerShape.TextFrame2.TextRange.Text = headerText & " " & arrowText
    headerShape.TextFrame2.MarginLeft = 5
    headerShape.TextFrame2.MarginRight = 5
    headerShape.TextFrame2.MarginTop = 0
    headerShape.TextFrame2.MarginBottom = 0
    uiRenderContext.Styles.ApplyControlStyle Nothing, headerShape, _
        m_uiControlBase.ControlNode, uiRenderContext.BindingContext
    If m_editable Then
        headerShape.TextFrame2.TextRange.Text = arrowText
        headerShape.TextFrame2.MarginLeft = 0
        headerShape.TextFrame2.MarginRight = 0
        headerShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End If
    headerShape.OnAction = "ex_UiBridge.fn_OnShapeClick"

    Set itemShapeNames = New Collection
    Set shapeNames = New Collection
    shapeNames.Add headerShapeName

    Set uiSelectAction = New obj_UiSelectShapeAction
    If Not uiSelectAction.Initialize( _
            controlId, m_targetRange.Cells(1, 1), headerShapeName, panelShapeName, _
            itemShapeNames, shapeNames, items, selectedValue, uiCellBinding, m_editable, Me, placeholder) Then Exit Function
    Set m_selectAction = uiSelectAction
    Dim routeName As Variant

    For Each routeName In uiSelectAction.ShapeNames
        uiRenderContext.Router.RegisterShape VBA.CStr(routeName), uiSelectAction
    Next routeName
    private_RenderSelectShapes = True
    ex_Core.fn_Diagnostic_WritePerf "Control.Select.CreateShapes | Name=" & m_uiControlBase.ControlName & _
        " | Items=" & VBA.CStr(items.Count), startedAt
End Function

Private Sub private_ApplyDefaultShapeStyle(ByVal targetShape As Shape)
    Dim node As Object

    Set node = m_uiControlBase.ControlNode.cloneNode(False)
    node.setAttribute "baseStyle", "selectShapeBase"
    m_renderContext.Styles.ApplyControlPartStyle targetShape, node, m_renderContext.BindingContext, "baseStyle"
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

Private Sub obj_IUiCellRefreshListener_OnValueRefreshed()
    private_cellBinding_ValueRefreshed
End Sub

Private Sub private_cellBinding_ValueRefreshed()
    If m_isDisposed Then Exit Sub
    If Not m_checkboxShape Is Nothing Then _
        m_checkboxShape.ControlFormat.Value = VBA.IIf(VBA.CBool(m_targetRange.Cells(1, 1).Value2), xlOn, xlOff)
End Sub

Private Function private_ShapeName(ByVal name As String) As String
    If VBA.Len(name) > 31 Then
        VBA.Err.Raise VBA.vbObjectError + 2240, "Input.ShapeName", _
            "The generated Shape name exceeds 31 characters: " & name
    End If
    private_ShapeName = name
End Function