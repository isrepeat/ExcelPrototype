VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_UiFieldControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiControl

Private m_uiControlBase As obj_UiControlBase
Private m_targetRange As Range
Private m_sourceName As String
Private m_bindingPath As String
Private m_items As String
Private m_itemsSourceRaw As String
Private m_changeCommandRaw As String
Private m_changeCommand As obj_UiCommand
Private m_isSelect As Boolean
Private m_isDisposed As Boolean

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
    m_sourceName = VBA.vbNullString
    m_bindingPath = VBA.vbNullString
    m_items = VBA.vbNullString
    m_itemsSourceRaw = VBA.vbNullString
    m_changeCommandRaw = VBA.vbNullString
    Set m_changeCommand = Nothing
End Sub

Private Function obj_IUiControl_Configure(ByVal controlNode As Object) As Boolean
    Dim controlType As String
    Dim inputType As String
    Dim rawValue As String
    Dim defaultSource As String

    If Not m_uiControlBase.Configure(controlNode) Then Exit Function
    controlType = VBA.LCase$(private_ReadAttribute(controlNode, "type"))
    inputType = VBA.LCase$(private_ReadAttribute(controlNode, "inputType"))
    m_isSelect = (controlType = "select" Or inputType = "select")
    m_items = private_ReadAttribute(controlNode, "items")
    m_itemsSourceRaw = private_ReadAttribute(controlNode, "itemsSource")
    If m_isSelect And VBA.Len(VBA.Trim$(m_items)) = 0 And _
       VBA.Len(VBA.Trim$(m_itemsSourceRaw)) = 0 Then Exit Function
    rawValue = private_ReadAttribute(controlNode, "value")
    defaultSource = private_GetFormSource(controlNode)
    If Not private_TryParseBinding(rawValue, defaultSource, m_sourceName, m_bindingPath) Then Exit Function
    m_changeCommandRaw = private_ReadAttribute(controlNode, "onChange")
    obj_IUiControl_Configure = True
End Function

Private Function obj_IUiControl_Measure(ByVal uiRenderContext As obj_UiRenderContext) As Range
    If m_uiControlBase Is Nothing Then Exit Function
    Set m_targetRange = m_uiControlBase.Measure(uiRenderContext)
    Set obj_IUiControl_Measure = m_targetRange
End Function

Private Function obj_IUiControl_Render(ByVal uiRenderContext As obj_UiRenderContext) As Boolean
    Dim value As Variant
    Dim sourceObject As Object
    Dim isObject As Boolean
    Dim uiCellBinding As obj_UiCellBinding
    Dim targetCell As Range
    Dim listText As String

    If m_targetRange Is Nothing Then Set m_targetRange = obj_IUiControl_Measure(uiRenderContext)
    If m_targetRange Is Nothing Then Exit Function
    If m_targetRange.Cells.CountLarge > 1 Then m_targetRange.Merge
    If Not uiRenderContext.BindingContext.TryGetValue( _
            m_sourceName, m_bindingPath, value, sourceObject, isObject) Then Exit Function
    If isObject Then Exit Function
    If VBA.Len(m_changeCommandRaw) > 0 Then
        If Not ex_UiBindingRuntime.fn_TryResolveCommand( _
                m_changeCommandRaw, uiRenderContext.BindingContext, m_changeCommand) Then Exit Function
    End If
    m_targetRange.Value2 = value
    ex_StylePipeline.fn_ApplyControlStyle m_targetRange, Nothing, _
        m_uiControlBase.ControlNode, uiRenderContext.BindingContext
    If m_isSelect Then
        If Not private_TryResolveItems(uiRenderContext, listText) Then Exit Function
        If Not private_ApplyListValidation(m_targetRange.Cells(1, 1), listText) Then Exit Function
    Else
        On Error Resume Next
        m_targetRange.Cells(1, 1).Validation.Delete
        On Error GoTo 0
    End If

    Set targetCell = m_targetRange.Cells(1, 1)
    Set uiCellBinding = New obj_UiCellBinding
    If Not uiCellBinding.Initialize( _
            targetCell.Parent.Name, targetCell.Address(False, False), _
            uiRenderContext.BindingContext, m_sourceName, m_bindingPath, m_changeCommand) Then Exit Function
    If Not ex_UiBindings.fn_RegisterCellBinding(uiCellBinding) Then Exit Function
    obj_IUiControl_Render = True
End Function

Private Function obj_IUiControl_HandleCellChange(ByVal target As Range) As Boolean
    obj_IUiControl_HandleCellChange = False
End Function

' //
' // Private
' //
Private Function private_ApplyListValidation( _
    ByVal targetCell As Range, _
    ByVal listText As String _
) As Boolean
    listText = VBA.Trim$(listText)
    If VBA.Len(listText) = 0 Or VBA.Len(listText) > 255 Then Exit Function
    On Error GoTo EH
    targetCell.Validation.Delete
    targetCell.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:=listText
    targetCell.Validation.IgnoreBlank = True
    targetCell.Validation.InCellDropdown = True
    private_ApplyListValidation = True
    Exit Function
EH:
    ex_Core.fn_Diagnostic_WriteLog "UI_FIELD_VALIDATION_ERROR | Name=" & _
        m_uiControlBase.ControlName & " | Description=" & VBA.Err.Description
End Function

Private Function private_TryResolveItems( _
    ByVal uiRenderContext As obj_UiRenderContext, _
    ByRef outListText As String _
) As Boolean
    Dim value As Variant
    Dim sourceObject As Object
    Dim isObject As Boolean
    Dim item As Variant

    outListText = m_items
    If VBA.Len(VBA.Trim$(m_itemsSourceRaw)) = 0 Then
        private_TryResolveItems = (VBA.Len(VBA.Trim$(outListText)) > 0)
        Exit Function
    End If
    If Not ex_UiBindingRuntime.fn_TryResolveValue( _
            m_itemsSourceRaw, uiRenderContext.BindingContext, value, sourceObject, isObject) Then Exit Function
    If isObject Then
        If TypeName(sourceObject) <> "Collection" Then Exit Function
        For Each item In sourceObject
            If VBA.Len(outListText) > 0 Then outListText = outListText & ","
            outListText = outListText & VBA.CStr(item)
        Next item
    ElseIf VBA.IsArray(value) Then
        For Each item In value
            If VBA.Len(outListText) > 0 Then outListText = outListText & ","
            outListText = outListText & VBA.CStr(item)
        Next item
    Else
        outListText = VBA.CStr(value)
    End If
    private_TryResolveItems = (VBA.Len(VBA.Trim$(outListText)) > 0)
End Function

Private Function private_TryParseBinding( _
    ByVal rawBinding As String, _
    ByVal defaultSource As String, _
    ByRef outSourceName As String, _
    ByRef outBindingPath As String _
) As Boolean
    Dim body As String
    Dim argument As Variant
    Dim separatorPosition As Long
    Dim argumentName As String

    rawBinding = VBA.Trim$(rawBinding)
    If VBA.Left$(rawBinding, 9) <> "{Binding " Or VBA.Right$(rawBinding, 1) <> "}" Then Exit Function
    body = VBA.Mid$(rawBinding, 10, VBA.Len(rawBinding) - 10)
    For Each argument In VBA.Split(body, ";")
        separatorPosition = VBA.InStr(1, VBA.CStr(argument), "=", VBA.vbBinaryCompare)
        If separatorPosition <= 0 Then GoTo ContinueArgument
        argumentName = VBA.Trim$(VBA.Left$(VBA.CStr(argument), separatorPosition - 1))
        Select Case VBA.LCase$(argumentName)
            Case "source"
                outSourceName = VBA.Trim$(VBA.Mid$(VBA.CStr(argument), separatorPosition + 1))
            Case "path"
                outBindingPath = VBA.Trim$(VBA.Mid$(VBA.CStr(argument), separatorPosition + 1))
        End Select
ContinueArgument:
    Next argument
    If VBA.Len(outSourceName) = 0 Then outSourceName = defaultSource
    If VBA.Len(outSourceName) = 0 Then outSourceName = "Form"
    private_TryParseBinding = (VBA.Len(outBindingPath) > 0)
End Function

Private Function private_GetFormSource(ByVal controlNode As Object) As String
    Dim currentNode As Object
    Dim nodeName As String

    Set currentNode = controlNode.parentNode
    Do While Not currentNode Is Nothing
        nodeName = VBA.LCase$(VBA.CStr(currentNode.baseName))
        If nodeName = "form" Then
            private_GetFormSource = private_ReadAttribute(currentNode, "source")
            Exit Function
        End If
        Set currentNode = currentNode.parentNode
    Loop
End Function

Private Function private_ReadAttribute(ByVal node As Object, ByVal attributeName As String) As String
    Dim value As Variant
    value = node.getAttribute(attributeName)
    If Not VBA.IsNull(value) And Not VBA.IsEmpty(value) Then private_ReadAttribute = VBA.CStr(value)
End Function