Attribute VB_Name = "ex_ControlAttributeContracts"
Option Explicit

' //
' // API
' //
Public Function m_IsCommonControlAttribute(ByVal attrName As String) As Boolean
    Select Case VBA.LCase$(VBA.Trim$(attrName))
        Case "name", "type", "style", "spancells", "spanrows", "visibility"
            m_IsCommonControlAttribute = True
    End Select
End Function


Public Function m_IsSupportedControlAttribute(ByVal control As obj_IControl, ByVal attrName As String) As Boolean
    If m_IsCommonControlAttribute(attrName) Then
        m_IsSupportedControlAttribute = True
        Exit Function
    End If

    If control Is Nothing Then Exit Function
    m_IsSupportedControlAttribute = control.SupportsAttribute(attrName)
End Function
