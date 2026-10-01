Attribute VB_Name = "ex_UiControlFactory"
Option Explicit

' --------------------------------------
' namespace Lifecycle {
' --------------------------------------
Public Sub fn_Module_Dispose()
End Sub
' --------------------------------------
' } // namespace Lifecycle
' --------------------------------------

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Create(ByVal controlNode As Object) As Object
    Dim controlType As String
    Dim uiControl As Object

    controlType = private_ReadAttribute(controlNode, "type")
    Select Case VBA.LCase$(controlType)
        Case "label"
            Set uiControl = New obj_UiLabelControl
        Case "button"
            Set uiControl = New obj_UiButtonControl
        Case Else
            VBA.MsgBox "Unsupported control type: " & controlType, _
                VBA.vbExclamation, "PersonalEventBuilder"
            Exit Function
    End Select
    If Not uiControl.Initialize() Then Exit Function
    Set fn_Create = uiControl
End Function
' --------------------------------------
' } // namespace API
' --------------------------------------

Private Function private_ReadAttribute(ByVal node As Object, ByVal attributeName As String) As String
    Dim attributeValue As Variant

    attributeValue = node.getAttribute(attributeName)
    If VBA.IsNull(attributeValue) Or VBA.IsEmpty(attributeValue) Then Exit Function
    private_ReadAttribute = VBA.CStr(attributeValue)
End Function