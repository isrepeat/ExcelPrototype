VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_UiCheckboxControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiControl
Implements obj_IUiBindingTarget
Implements obj_IUiEventHandler

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private Const CHECKBOX_SHAPE_PREFIX As String = "chk_"
Private m_renderedCaption As String
Private m_renderContext As obj_UiRenderContext
Private m_uiControlBase As obj_UiControlBase
Private m_targetRange As Range
Private m_buttonShape As Shape

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
    obj_IUiControl_Initialize = Me.Initialize()
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
    obj_IUiBindingTarget_RefreshBindings = private_RefreshCaption()
End Function

Private Function obj_IUiEventHandler_HandleEvent( _
    ByVal kind As String, _
    ByVal payload As Variant _
) As Boolean
    Dim checked As Variant
    Dim resolvedObject As Object
    Dim isObject As Boolean
    Dim sourceName As String
    Dim bindingPath As String
    Dim rawBinding As String
    Dim changeCommand As obj_UiCommand
    Dim changeBinding As String

    If Not m_isInitialized Or m_isDisposed Or kind <> "click" Then Exit Function
    rawBinding = private_ReadAttribute(m_uiControlBase.ControlNode, "value")
    If Not ex_UiBindingRuntime.fn_TryResolveValue( _
            rawBinding, m_renderContext.BindingContext, checked, resolvedObject, isObject) Then Exit Function
    If isObject Or VBA.VarType(checked) <> VBA.vbBoolean Then Exit Function
    If Not ex_UiBindingRuntime.fn_TryParseBinding( _
            rawBinding, VBA.vbNullString, sourceName, bindingPath, m_renderContext.BindingContext) Then Exit Function
    If Not m_renderContext.BindingContext.TrySetPathValue( _
            sourceName, bindingPath, Not VBA.CBool(checked)) Then Exit Function
    obj_IUiEventHandler_HandleEvent = True
    changeBinding = private_ReadAttribute(m_uiControlBase.ControlNode, "onChange")
    If VBA.Len(changeBinding) > 0 Then
        If Not ex_UiBindingRuntime.fn_TryResolveCommand(changeBinding, m_renderContext.BindingContext, changeCommand) Then
            obj_IUiEventHandler_HandleEvent = False
            Exit Function
        End If
        obj_IUiEventHandler_HandleEvent = changeCommand.Execute()
    End If
End Function

Private Function private_RefreshCaption() As Boolean
    Dim caption As String
    Dim glyph As String
    Dim glyphAttribute As String
    Dim checked As Variant
    Dim resolvedObject As Object
    Dim isObject As Boolean

    If m_buttonShape Is Nothing Then Exit Function
    If Not m_uiControlBase.TryGetCaption(m_renderContext.BindingContext, caption) Then Exit Function
    If Not ex_UiBindingRuntime.fn_TryResolveValue( _
            private_ReadAttribute(m_uiControlBase.ControlNode, "value"), _
            m_renderContext.BindingContext, checked, resolvedObject, isObject) Then Exit Function
    If isObject Or VBA.VarType(checked) <> VBA.vbBoolean Then Exit Function
    If VBA.CBool(checked) Then
        glyphAttribute = "checkedGlyph"
    Else
        glyphAttribute = "uncheckedGlyph"
    End If
    If Not ex_UiBindingRuntime.fn_TryResolveText( _
            private_ReadAttribute(m_uiControlBase.ControlNode, glyphAttribute), _
            m_renderContext.BindingContext, glyph) Then Exit Function
    If VBA.Len(glyph) = 0 Then Exit Function
    caption = glyph & "  " & caption
    If caption <> m_renderedCaption Then
        m_buttonShape.TextFrame2.TextRange.Text = caption
        m_buttonShape.TextFrame2.WordWrap = msoFalse
        m_buttonShape.TextFrame2.AutoSize = msoAutoSizeNone
        m_renderedCaption = caption
    End If
    private_RefreshCaption = True
End Function

' //
' // API
' //
Public Function Initialize() As Boolean
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    If m_uiControlBase Is Nothing Then
        Set m_uiControlBase = New obj_UiControlBase
    End If
    Initialize = m_uiControlBase.Initialize()
    m_isInitialized = Initialize
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    Set m_buttonShape = Nothing
    If Not m_uiControlBase Is Nothing Then m_uiControlBase.Dispose
    Set m_uiControlBase = Nothing
    Set m_renderContext = Nothing
    Set m_targetRange = Nothing
End Sub

' //
' // Private
' //
Private Function private_Configure(ByVal controlNode As Object) As Boolean
    Dim checked As Variant
    Dim resolvedObject As Object
    Dim isObject As Boolean
    Dim sourceName As String
    Dim bindingPath As String
    Dim rawBinding As String

    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If Not m_uiControlBase.Configure(controlNode) Then Exit Function
    rawBinding = private_ReadAttribute(controlNode, "value")
    If Not ex_UiBindingRuntime.fn_TryParseBinding( _
            rawBinding, VBA.vbNullString, sourceName, bindingPath, m_renderContext.BindingContext) Then Exit Function
    If VBA.Len(bindingPath) = 0 Then Exit Function
    If Not ex_UiBindingRuntime.fn_TryResolveValue( _
            rawBinding, m_renderContext.BindingContext, checked, resolvedObject, isObject) Then Exit Function
    private_Configure = Not isObject And VBA.VarType(checked) = VBA.vbBoolean
End Function

Private Function private_Render(ByVal uiRenderContext As obj_UiRenderContext) As Boolean
    Dim startedAt As Double
    Dim targetRange As Range

    startedAt = VBA.Timer
    Dim shapeName As String
    Dim captionText As String

    Set targetRange = m_targetRange
    If targetRange Is Nothing Then Set targetRange = private_Measure(uiRenderContext)
    shapeName = m_uiControlBase.ShapeName(CHECKBOX_SHAPE_PREFIX)
    If VBA.Len(shapeName) = 0 Then Exit Function
    If Not m_uiControlBase.TryGetCaption(uiRenderContext.BindingContext, captionText) Then Exit Function
    If Not m_buttonShape Is Nothing Then m_buttonShape.Delete
    Set m_buttonShape = uiRenderContext.TargetWorksheet.Shapes.AddShape( _
        msoShapeRectangle, targetRange.Left, targetRange.Top, targetRange.Width, targetRange.Height)
    m_buttonShape.Name = shapeName
    m_renderedCaption = VBA.vbNullString
    m_buttonShape.Placement = xlMoveAndSize
    m_buttonShape.TextFrame2.TextRange.Text = captionText
    uiRenderContext.Styles.ApplyControlStyle targetRange, m_buttonShape, m_uiControlBase.ControlNode, uiRenderContext.BindingContext
    m_buttonShape.OnAction = "ex_UiBridge.fn_OnShapeClick"
    uiRenderContext.Router.RegisterShape shapeName, Me
    If Not private_RefreshCaption() Then Exit Function
    private_Render = True
    ex_Core.fn_Diagnostic_WritePerf "Control.Checkbox.Render", startedAt
End Function

Private Function private_Measure(ByVal uiRenderContext As obj_UiRenderContext) As Range
    If m_uiControlBase Is Nothing Then Exit Function
    Set m_targetRange = m_uiControlBase.Measure(uiRenderContext)
    Set private_Measure = m_targetRange
End Function

' //
' // Private
' //
Private Function private_ReadAttribute( _
    ByVal node As Object, _
    ByVal attributeName As String _
) As String
    Dim attributeValue As Variant

    attributeValue = node.getAttribute(attributeName)
    If VBA.IsNull(attributeValue) Or VBA.IsEmpty(attributeValue) Then Exit Function
    private_ReadAttribute = VBA.CStr(attributeValue)
End Function