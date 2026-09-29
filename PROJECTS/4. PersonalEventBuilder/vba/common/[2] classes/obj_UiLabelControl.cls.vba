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

Private m_uiControlBase As obj_UiControlBase

Private Sub Class_Initialize()
    Set m_uiControlBase = New obj_UiControlBase
End Sub

Public Function fn_Configure(ByVal controlNode As Object) As Boolean
    fn_Configure = m_uiControlBase.fn_Configure(controlNode)
End Function

Public Function fn_Render(ByVal uiRenderContext As obj_UiRenderContext) As Boolean
    Dim targetRange As Range

    Set targetRange = m_uiControlBase.fn_TargetRange(uiRenderContext)
    targetRange.Merge
    targetRange.Value2 = m_uiControlBase.fn_Caption
    ex_StylePipeline.fn_ApplyControlStyle targetRange, Nothing, m_uiControlBase.fn_ControlNode
    fn_Render = True
End Function

Public Sub fn_Dispose()
    m_uiControlBase.fn_Dispose
    Set m_uiControlBase = Nothing
End Sub