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
    Set m_buttonShape = Nothing
    If Not m_uiControlBase Is Nothing Then m_uiControlBase.Dispose
    Set m_uiControlBase = Nothing
    Set m_targetRange = Nothing
End Sub

Private Function obj_IUiControl_Configure(ByVal controlNode As Object) As Boolean
    obj_IUiControl_Configure = m_uiControlBase.Configure(controlNode)
End Function

Private Function obj_IUiControl_Render(ByVal uiRenderContext As obj_UiRenderContext) As Boolean
    Dim startedAt As Double
    Dim targetRange As Range

    startedAt = VBA.Timer
    Dim shapeName As String
    Dim captionText As String
    Dim uiCommand As obj_UiCommand

    Set targetRange = m_targetRange
    If targetRange Is Nothing Then Set targetRange = obj_IUiControl_Measure(uiRenderContext)
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
    obj_IUiControl_Render = True
    ex_Core.fn_Diagnostic_WritePerf "Control.Button.Render", startedAt
End Function

Private Function obj_IUiControl_Measure(ByVal uiRenderContext As obj_UiRenderContext) As Range
    If m_uiControlBase Is Nothing Then Exit Function
    Set m_targetRange = m_uiControlBase.Measure(uiRenderContext)
    Set obj_IUiControl_Measure = m_targetRange
End Function

Private Function obj_IUiControl_HandleCellChange(ByVal target As Range) As Boolean
    obj_IUiControl_HandleCellChange = False
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

' //
' // Interface
' //
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