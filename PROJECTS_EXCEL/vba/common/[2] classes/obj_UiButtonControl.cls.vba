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

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize() As Boolean
    If m_uiControlBase Is Nothing Then Set m_uiControlBase = New obj_UiControlBase
    Initialize = m_uiControlBase.Initialize()
End Function

Public Sub Dispose()
    Set m_buttonShape = Nothing
    m_uiControlBase.Dispose
    Set m_uiControlBase = Nothing
End Sub

Public Function Configure(ByVal controlNode As Object) As Boolean
    Configure = m_uiControlBase.Configure(controlNode)
End Function

Public Function Render(ByVal uiRenderContext As obj_UiRenderContext) As Boolean
    Dim targetRange As Range
    Dim shapeName As String
    Dim captionText As String
    Dim uiCommand As obj_UiCommand

    Set targetRange = m_uiControlBase.TargetRange(uiRenderContext)
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
    Render = True
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