Attribute VB_Name = "ex_ControlAttributeContracts"
Option Explicit
#Const LOGGING_VERBOSE_ENABLED = False

Public Sub fn_Module_Dispose()
#If LOGGING_VERBOSE_ENABLED Then
    ex_Core.fn_Diagnostic_LogInfo "lifecycle:ex_ControlAttributeContracts.fn_Module_Dispose"
#End If
End Sub

' //
' // API
' //
Public Function fn_IsCommonControlAttribute(ByVal attrName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(attrName))
        Case "name", "type", "style", "spancolls", "spanrows", "visibility", "datacontext"
            fn_IsCommonControlAttribute = True
    End Select
End Function


Public Function fn_IsSupportedControlAttribute(ByVal control As obj_IControl, ByVal attrName As String) As Boolean
    If fn_IsCommonControlAttribute(attrName) Then
        fn_IsSupportedControlAttribute = True
        Exit Function
    End If

    If control Is Nothing Then Exit Function
    fn_IsSupportedControlAttribute = control.SupportsAttribute(attrName)
End Function
