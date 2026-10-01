VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_UiLabelControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiControl

Private m_uiControlBase As obj_UiControlBase
Private m_targetRange As Range

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
    m_uiControlBase.Dispose
    Set m_uiControlBase = Nothing
    Set m_targetRange = Nothing
End Sub

Private Function obj_IUiControl_Configure(ByVal controlNode As Object) As Boolean
    obj_IUiControl_Configure = m_uiControlBase.Configure(controlNode)
End Function

Private Function obj_IUiControl_Render(ByVal uiRenderContext As obj_UiRenderContext) As Boolean
    Dim targetRange As Range
    Dim captionText As String

    Set targetRange = m_targetRange
    If targetRange Is Nothing Then Set targetRange = Me.obj_IUiControl_Measure(uiRenderContext)
    If Not m_uiControlBase.TryGetCaption(uiRenderContext.BindingContext, captionText) Then Exit Function
    targetRange.Merge
    targetRange.Value2 = captionText
    ex_StylePipeline.fn_ApplyControlStyle targetRange, Nothing, m_uiControlBase.ControlNode, uiRenderContext.BindingContext
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