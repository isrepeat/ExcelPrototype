VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_UiControlBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_controlNode As Object
Private m_controlName As String

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Configure(ByVal controlNode As Object) As Boolean
    If controlNode Is Nothing Then Exit Function
    Set m_controlNode = controlNode
    m_controlName = private_ReadAttribute(controlNode, "name")
    If VBA.Len(m_controlName) = 0 Then
        VBA.MsgBox "A control name is required.", VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    fn_Configure = True
End Function

Public Function fn_TargetRange(ByVal uiRenderContext As obj_UiRenderContext) As Range
    Dim targetWorksheet As Worksheet
    Dim rowNo As Long
    Dim columnNo As Long
    Dim rowSpan As Long
    Dim columnSpan As Long

    If uiRenderContext Is Nothing Then Exit Function
    Set targetWorksheet = uiRenderContext.fn_TargetWorksheet
    rowNo = private_ReadLong("row", 1)
    columnNo = private_ReadLong("column", 1)
    rowSpan = private_ReadLong("rowSpan", 1)
    columnSpan = private_ReadLong("columnSpan", 1)
    Set fn_TargetRange = targetWorksheet.Range(targetWorksheet.Cells(rowNo, columnNo), _
        targetWorksheet.Cells(rowNo + rowSpan - 1, columnNo + columnSpan - 1))
End Function

Public Function fn_TryGetCaption(ByVal uiBindingContext As obj_UiBindingContext, ByRef outCaption As String) As Boolean
    Dim rawCaption As String

    rawCaption = private_ReadAttribute(m_controlNode, "text")
    If VBA.Len(rawCaption) = 0 Then rawCaption = private_ReadAttribute(m_controlNode, "caption")
    If VBA.Len(rawCaption) = 0 Then
        VBA.MsgBox "A control caption is required: " & m_controlName, VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    fn_TryGetCaption = ex_UiBindingRuntime.fn_TryResolveText(rawCaption, uiBindingContext, outCaption)
End Function

Public Property Get fn_ControlNode() As Object
    Set fn_ControlNode = m_controlNode
End Property

Public Property Get fn_ControlName() As String
    fn_ControlName = m_controlName
End Property

Public Function fn_ShapeName(ByVal prefix As String) As String
    fn_ShapeName = prefix & m_controlName
    If VBA.Len(fn_ShapeName) <= 31 Then Exit Function
    VBA.MsgBox "The generated Shape name exceeds 31 characters: " & fn_ShapeName, _
        VBA.vbExclamation, "PersonalEventBuilder"
    fn_ShapeName = VBA.vbNullString
End Function

Public Sub fn_Dispose()
    Set m_controlNode = Nothing
    m_controlName = VBA.vbNullString
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_ReadLong(ByVal attributeName As String, ByVal defaultValue As Long) As Long
    Dim valueText As String
    valueText = private_ReadAttribute(m_controlNode, attributeName)
    If VBA.IsNumeric(valueText) Then private_ReadLong = VBA.CLng(valueText) Else private_ReadLong = defaultValue
End Function

Private Function private_ReadAttribute(ByVal node As Object, ByVal attributeName As String) As String
    Dim attributeValue As Variant
    attributeValue = node.getAttribute(attributeName)
    If VBA.IsNull(attributeValue) Or VBA.IsEmpty(attributeValue) Then Exit Function
    private_ReadAttribute = VBA.CStr(attributeValue)
End Function