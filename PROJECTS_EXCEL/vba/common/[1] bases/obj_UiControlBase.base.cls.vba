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

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_controlNode As Object
Private m_rowOffset As Long
Private m_columnOffset As Long
Private m_controlName As String

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Properties
' //
Public Property Get ControlNode() As Object
    Set ControlNode = m_controlNode
End Property

Public Property Get ControlName() As String
    ControlName = m_controlName
End Property

' //
' // API
' //
Public Function Initialize() As Boolean
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    Initialize = True
    m_isInitialized = Initialize
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    Set m_controlNode = Nothing
    m_controlName = VBA.vbNullString
End Sub

Public Function Configure(ByVal controlNode As Object) As Boolean
    If controlNode Is Nothing Then Exit Function
    Set m_controlNode = controlNode
    m_controlName = private_ReadAttribute(controlNode, "name")
    If VBA.Len(m_controlName) = 0 Then
        ex_WindowsUi.fn_ShowMessage "A control name is required.", VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    Configure = True
End Function

Public Sub ConfigurePosition(ByVal definition As Object)
    m_rowOffset = ex_UiElementFactory.fn_Long(definition, "row", 1)
    m_columnOffset = ex_UiElementFactory.fn_Long(definition, "column", 1)
End Sub

Public Sub SetPosition(ByVal row As Long, ByVal column As Long)
    m_controlNode.setAttribute "row", VBA.CStr(row)
    m_controlNode.setAttribute "column", VBA.CStr(column)
End Sub

Public Sub ArrangePosition(ByVal row As Long, ByVal column As Long)
    Me.SetPosition row + m_rowOffset - 1, column + m_columnOffset - 1
End Sub

Public Sub GetSize(ByVal target As Range, ByRef rows As Long, ByRef columns As Long)
    rows = target.Rows.Count + m_rowOffset - 1
    columns = target.Columns.Count + m_columnOffset - 1
End Sub

Public Function Measure(ByVal uiRenderContext As obj_UiRenderContext) As Range
    Dim targetWorksheet As Worksheet
    Dim rowNo As Long
    Dim columnNo As Long
    Dim rowSpan As Long
    Dim columnSpan As Long

    If uiRenderContext Is Nothing Then Exit Function
    Set targetWorksheet = uiRenderContext.TargetWorksheet
    rowNo = private_ReadLong("row", 1)
    columnNo = private_ReadLong("column", 1)
    rowSpan = private_ReadLong("rowSpan", 1)
    columnSpan = private_ReadLong("columnSpan", 1)
    If rowNo <= 0 Or columnNo <= 0 Or rowSpan <= 0 Or columnSpan <= 0 Then Exit Function
    Set Measure = targetWorksheet.Range(targetWorksheet.Cells(rowNo, columnNo), _
        targetWorksheet.Cells(rowNo + rowSpan - 1, columnNo + columnSpan - 1))
End Function

Public Function TryGetCaption(ByVal uiBindingContext As obj_UiBindingContext, ByRef outCaption As String) As Boolean
    Dim rawCaption As String

    rawCaption = private_ReadAttribute(m_controlNode, "text")
    If VBA.Len(rawCaption) = 0 Then rawCaption = private_ReadAttribute(m_controlNode, "caption")
    If VBA.Len(rawCaption) = 0 Then
        ex_WindowsUi.fn_ShowMessage "A control caption is required: " & m_controlName, VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    TryGetCaption = ex_UiBindingRuntime.fn_TryResolveText(rawCaption, uiBindingContext, outCaption)
End Function

Public Function ShapeName(ByVal prefix As String) As String
    ShapeName = prefix & m_controlName
    If VBA.Len(ShapeName) <= 31 Then Exit Function
    ex_WindowsUi.fn_ShowMessage "The generated Shape name exceeds 31 characters: " & ShapeName, _
        VBA.vbExclamation, "PersonalEventBuilder"
    ShapeName = VBA.vbNullString
End Function

' //
' // Private
' //
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