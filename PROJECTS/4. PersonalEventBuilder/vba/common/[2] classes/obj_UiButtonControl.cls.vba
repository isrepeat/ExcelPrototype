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

Private Const BUTTON_SHAPE_PREFIX As String = "btn_"
Private m_uiControlBase As obj_UiControlBase
Private m_buttonShape As Shape

Private Sub Class_Initialize()
    Set m_uiControlBase = New obj_UiControlBase
End Sub

Public Function fn_Configure(ByVal controlNode As Object) As Boolean
    fn_Configure = m_uiControlBase.fn_Configure(controlNode)
End Function

Public Function fn_Render(ByVal uiRenderContext As obj_UiRenderContext) As Boolean
    Dim targetRange As Range
    Dim shapeName As String
    Dim captionText As String
    Dim uiCommand As obj_UiCommand

    Set targetRange = m_uiControlBase.fn_TargetRange(uiRenderContext)
    shapeName = m_uiControlBase.fn_ShapeName(BUTTON_SHAPE_PREFIX)
    If VBA.Len(shapeName) = 0 Then Exit Function
    If Not m_uiControlBase.fn_TryGetCaption(uiRenderContext.fn_BindingContext, captionText) Then Exit Function
    If Not ex_UiBindingRuntime.fn_TryResolveCommand(private_ReadAttribute( _
            m_uiControlBase.fn_ControlNode, "command"), uiRenderContext.fn_BindingContext, uiCommand) Then Exit Function
    Set m_buttonShape = uiRenderContext.fn_TargetWorksheet.Shapes.AddShape( _
        msoShapeRoundedRectangle, targetRange.Left, targetRange.Top, targetRange.Width, targetRange.Height)
    m_buttonShape.Name = shapeName
    m_buttonShape.TextFrame2.TextRange.Text = captionText
    ex_StylePipeline.fn_ApplyControlStyle targetRange, m_buttonShape, m_uiControlBase.fn_ControlNode, uiRenderContext.fn_BindingContext
    m_buttonShape.OnAction = "ex_UiBridge.fn_OnShapeClick"
    ex_UiBindings.fn_Register shapeName, uiCommand
    fn_Render = True
End Function

Public Sub fn_Dispose()
    Set m_buttonShape = Nothing
    m_uiControlBase.fn_Dispose
    Set m_uiControlBase = Nothing
End Sub

Private Function private_ReadAttribute(ByVal node As Object, ByVal attributeName As String) As String
    Dim attributeValue As Variant
    attributeValue = node.getAttribute(attributeName)
    If VBA.IsNull(attributeValue) Or VBA.IsEmpty(attributeValue) Then Exit Function
    private_ReadAttribute = VBA.CStr(attributeValue)
End Function