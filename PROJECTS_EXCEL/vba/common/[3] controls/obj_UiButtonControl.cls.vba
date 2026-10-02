VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_UiButtonControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiControl
Implements obj_IUiBindingTarget

Private Const BUTTON_SHAPE_PREFIX As String = "btn_"
Private m_renderContext As obj_UiRenderContext
Private m_uiControlBase As obj_UiControlBase
Private m_targetRange As Range
Private m_buttonShape As Shape
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

Private Function obj_IUiBindingTarget_RefreshBindings( _
    ByVal context As obj_UiRenderContext _
) As Boolean
    Dim caption As String
    Dim uiCommand As obj_UiCommand

    If m_buttonShape Is Nothing Then Exit Function
    If Not m_uiControlBase.TryGetCaption(context.BindingContext, caption) Then Exit Function
    m_buttonShape.TextFrame2.TextRange.Text = caption
    context.Styles.ApplyControlStyle m_targetRange, m_buttonShape, _
        m_uiControlBase.ControlNode, context.BindingContext
    If Not ex_UiBindingRuntime.fn_TryResolveCommand(private_ReadAttribute( _
            m_uiControlBase.ControlNode, "command"), context.BindingContext, uiCommand) Then Exit Function
    context.Router.RegisterShape m_buttonShape.Name, uiCommand
    obj_IUiBindingTarget_RefreshBindings = True
End Function

' //
' // API
' //
Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
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
    private_Configure = m_uiControlBase.Configure(controlNode)
End Function

Private Function private_Render(ByVal uiRenderContext As obj_UiRenderContext) As Boolean
    Dim startedAt As Double
    Dim targetRange As Range

    startedAt = VBA.Timer
    Dim shapeName As String
    Dim captionText As String
    Dim uiCommand As obj_UiCommand

    Set targetRange = m_targetRange
    If targetRange Is Nothing Then Set targetRange = private_Measure(uiRenderContext)
    shapeName = m_uiControlBase.ShapeName(BUTTON_SHAPE_PREFIX)
    If VBA.Len(shapeName) = 0 Then Exit Function
    If Not m_uiControlBase.TryGetCaption(uiRenderContext.BindingContext, captionText) Then Exit Function
    If Not ex_UiBindingRuntime.fn_TryResolveCommand(private_ReadAttribute( _
            m_uiControlBase.ControlNode, "command"), uiRenderContext.BindingContext, uiCommand) Then Exit Function
    If Not m_buttonShape Is Nothing Then m_buttonShape.Delete
    Set m_buttonShape = uiRenderContext.TargetWorksheet.Shapes.AddShape( _
        msoShapeRoundedRectangle, targetRange.Left, targetRange.Top, targetRange.Width, targetRange.Height)
    m_buttonShape.Name = shapeName
    m_buttonShape.TextFrame2.TextRange.Text = captionText
    uiRenderContext.Styles.ApplyControlStyle targetRange, m_buttonShape, m_uiControlBase.ControlNode, uiRenderContext.BindingContext
    m_buttonShape.OnAction = "ex_UiBridge.fn_OnShapeClick"
    uiRenderContext.Router.RegisterShape shapeName, uiCommand
    private_Render = True
    ex_Core.fn_Diagnostic_WritePerf "Control.Button.Render", startedAt
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