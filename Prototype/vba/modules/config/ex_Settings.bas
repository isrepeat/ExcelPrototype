Attribute VB_Name = "ex_Settings"
Option Explicit

' =============================================================================
' Public API: Bool flags
' =============================================================================

Public Function m_GetBoolFlag(ByVal flagName As String, ByVal defaultValue As Boolean) As Boolean
    On Error GoTo NoProp
    m_GetBoolFlag = CBool(ThisWorkbook.CustomDocumentProperties(flagName).Value)
    Exit Function
NoProp:
    m_GetBoolFlag = defaultValue
End Function

Public Sub m_SetBoolFlag(ByVal flagName As String, ByVal valueBool As Boolean)
    On Error GoTo AddProp
    ThisWorkbook.CustomDocumentProperties(flagName).Value = valueBool
    Exit Sub
AddProp:
    ThisWorkbook.CustomDocumentProperties.Add _
        Name:=flagName, _
        LinkToContent:=False, _
        Type:=msoPropertyTypeBoolean, _
        Value:=valueBool
End Sub
