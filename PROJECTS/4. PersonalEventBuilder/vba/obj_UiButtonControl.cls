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
    Dim callbackName As String

    Set targetRange = m_uiControlBase.fn_TargetRange(uiRenderContext)
    shapeName = m_uiControlBase.fn_ShapeName(BUTTON_SHAPE_PREFIX)
    If VBA.Len(shapeName) = 0 Then Exit Function

    Set m_buttonShape = uiRenderContext.fn_TargetWorksheet.Shapes.AddShape( _
        msoShapeRoundedRectangle, targetRange.Left, targetRange.Top, _
        targetRange.Width, targetRange.Height)
    m_buttonShape.Name = shapeName
    m_buttonShape.TextFrame2.TextRange.Text = m_uiControlBase.fn_Caption
    ex_StylePipeline.fn_ApplyControlStyle targetRange, m_buttonShape, _
        m_uiControlBase.fn_ControlNode
    m_buttonShape.OnAction = "ex_UiBridge.fn_OnShapeClick"

    callbackName = private_ReadCallback(m_uiControlBase.fn_ControlNode)
    If VBA.Len(callbackName) = 0 Then
        VBA.MsgBox "Button binding is required: " & shapeName, _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    ex_UiBindings.fn_Register shapeName, callbackName
    fn_Render = True
End Function

Public Sub fn_Dispose()
    Set m_buttonShape = Nothing
    m_uiControlBase.fn_Dispose
    Set m_uiControlBase = Nothing
End Sub

Private Function private_ReadCallback(ByVal controlNode As Object) As String
    Dim bindingText As String
    Dim moduleName As String
    Dim methodName As String

    bindingText = private_ReadAttribute(controlNode, "onClick")
    moduleName = private_ReadBindingValue(bindingText, "Module=")
    methodName = private_ReadBindingValue(bindingText, "Method=")
    If VBA.Len(moduleName) > 0 And VBA.Len(methodName) > 0 Then _
        private_ReadCallback = moduleName & "." & methodName
End Function

Private Function private_ReadBindingValue(ByVal bindingText As String, ByVal keyName As String) As String
    Dim valueStart As Long
    Dim valueEnd As Long

    valueStart = VBA.InStr(1, bindingText, keyName, VBA.vbTextCompare)
    If valueStart = 0 Then Exit Function
    valueStart = valueStart + VBA.Len(keyName)
    valueEnd = VBA.InStr(valueStart, bindingText, ";")
    If valueEnd = 0 Then valueEnd = VBA.InStr(valueStart, bindingText, "}")
    private_ReadBindingValue = VBA.Trim$(VBA.Mid$(bindingText, valueStart, valueEnd - valueStart))
End Function

Private Function private_ReadAttribute(ByVal node As Object, ByVal attributeName As String) As String
    Dim attributeValue As Variant

    attributeValue = node.getAttribute(attributeName)
    If VBA.IsNull(attributeValue) Or VBA.IsEmpty(attributeValue) Then Exit Function
    private_ReadAttribute = VBA.CStr(attributeValue)
End Function