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

    controlType = private_ReadAttribute(controlNode, "type")
    Select Case VBA.LCase$(controlType)
        Case "label"
            Set fn_Create = New obj_UiLabelControl
        Case "button"
            Set fn_Create = New obj_UiButtonControl
        Case Else
            VBA.MsgBox "Unsupported control type: " & controlType, _
                VBA.vbExclamation, "PersonalEventBuilder"
    End Select
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