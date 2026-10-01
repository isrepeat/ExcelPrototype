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

Private Const BUTTON_SHAPE_PREFIX As String = "btn_"
Private m_uiControlBase As obj_UiControlBase
Private m_targetRange As Range
Private m_buttonShape As Shape

Private Sub Class_Initialize()
    Set m_uiControlBase = New obj_UiControlBase
End Sub

Private Sub Class_Terminate()
    Me.obj_IUiControl_Dispose
End Sub

' //
' // Interface
' //
Private Function obj_IUiControl_Initialize() As Boolean
    If m_uiControlBase Is Nothing Then Set m_uiControlBase = New obj_UiControlBase
    obj_IUiControl_Initialize = m_uiControlBase.Initialize()
End Function

Private Sub obj_IUiControl_Dispose()
    Set m_buttonShape = Nothing
    m_uiControlBase.Dispose
    Set m_uiControlBase = Nothing
    Set m_targetRange = Nothing
End Sub

Private Function obj_IUiControl_Configure(ByVal controlNode As Object) As Boolean
    obj_IUiControl_Configure = m_uiControlBase.Configure(controlNode)
End Function

Private Function obj_IUiControl_Render(ByVal uiRenderContext As obj_UiRenderContext) As Boolean
    Dim targetRange As Range
    Dim shapeName As String
    Dim captionText As String
    Dim uiCommand As obj_UiCommand

    Set targetRange = m_targetRange
    If targetRange Is Nothing Then Set targetRange = Me.obj_IUiControl_Measure(uiRenderContext)
    shapeName = m_uiControlBase.ShapeName(BUTTON_SHAPE_PREFIX)
    If VBA.Len(shapeName) = 0 Then Exit Function
    If Not m_uiControlBase.TryGetCaption(uiRenderContext.BindingContext, captionText) Then Exit Function
    If Not ex_UiBindingRuntime.fn_TryResolveCommand(private_ReadAttribute( _
            m_uiControlBase.ControlNode, "command"), uiRenderContext.BindingContext, uiCommand) Then Exit Function
    Set m_buttonShape = uiRenderContext.TargetWorksheet.Shapes.AddShape( _
        msoShapeRoundedRectangle, targetRange.Left, targetRange.Top, targetRange.Width, targetRange.Height)
    m_buttonShape.Name = shapeName
    m_buttonShape.TextFrame2.TextRange.Text = captionText
    ex_StylePipeline.fn_ApplyControlStyle targetRange, m_buttonShape, m_uiControlBase.ControlNode, uiRenderContext.BindingContext
    m_buttonShape.OnAction = "ex_UiBridge.fn_OnShapeClick"
    ex_UiBindings.fn_Register shapeName, uiCommand
    obj_IUiControl_Render = True
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
Private Function private_ReadAttribute(ByVal node As Object, ByVal attributeName As String) As String
    Dim attributeValue As Variant
    attributeValue = node.getAttribute(attributeName)
    If VBA.IsNull(attributeValue) Or VBA.IsEmpty(attributeValue) Then Exit Function
    private_ReadAttribute = VBA.CStr(attributeValue)
End Function