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
Implements obj_IUiBindingTarget

Private m_uiControlBase As obj_UiControlBase
Private m_targetRange As Range
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
    Dim captionText As String

    Set targetRange = m_targetRange
    If targetRange Is Nothing Then Set targetRange = obj_IUiControl_Measure(uiRenderContext)
    If Not m_uiControlBase.TryGetCaption(uiRenderContext.BindingContext, captionText) Then Exit Function
    targetRange.Merge
    targetRange.Value2 = captionText
    uiRenderContext.Styles.ApplyControlStyle targetRange, Nothing, m_uiControlBase.ControlNode, uiRenderContext.BindingContext
    obj_IUiControl_Render = True
    ex_Core.fn_Diagnostic_WritePerf "Control.Label.Render", startedAt
End Function

Private Function obj_IUiControl_Measure(ByVal uiRenderContext As obj_UiRenderContext) As Range
    If m_uiControlBase Is Nothing Then Exit Function
    Set m_targetRange = m_uiControlBase.Measure(uiRenderContext)
    Set obj_IUiControl_Measure = m_targetRange
End Function

Private Function obj_IUiControl_HandleCellChange(ByVal target As Range) As Boolean
    obj_IUiControl_HandleCellChange = False
End Function

Private Function obj_IUiBindingTarget_RefreshBindings( _
    ByVal context As obj_UiRenderContext _
) As Boolean
    obj_IUiBindingTarget_RefreshBindings = obj_IUiControl_Render(context)
End Function