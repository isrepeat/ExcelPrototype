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
    m_uiControlBase.Dispose
    Set m_uiControlBase = Nothing
End Sub

Public Function Configure(ByVal controlNode As Object) As Boolean
    Configure = m_uiControlBase.Configure(controlNode)
End Function

Public Function Render(ByVal uiRenderContext As obj_UiRenderContext) As Boolean
    Dim targetRange As Range
    Dim captionText As String

    Set targetRange = m_uiControlBase.TargetRange(uiRenderContext)
    If Not m_uiControlBase.TryGetCaption(uiRenderContext.BindingContext, captionText) Then Exit Function
    targetRange.Merge
    targetRange.Value2 = captionText
    ex_StylePipeline.fn_ApplyControlStyle targetRange, Nothing, m_uiControlBase.ControlNode, uiRenderContext.BindingContext
    Render = True
End Function