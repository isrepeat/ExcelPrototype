Attribute VB_Name = "ex_State"
Option Explicit

Public Const STATE_ACTIVE_MODE As String = "Settings.ActiveMode"
Public Const STATE_ACTIVE_PROFILE As String = "Settings.ActiveProfile"

Public Function m_GetText(ByVal propName As String, Optional ByVal defaultValue As String = vbNullString) As String
    On Error GoTo EH
    m_GetText = CStr(ThisWorkbook.CustomDocumentProperties(propName).Value)
    Exit Function
EH:
    m_GetText = defaultValue
End Function

Public Sub m_SetText(ByVal propName As String, ByVal valueText As String)
    On Error GoTo AddProp
    ThisWorkbook.CustomDocumentProperties(propName).Value = CStr(valueText)
    Exit Sub
AddProp:
    ThisWorkbook.CustomDocumentProperties.Add _
        Name:=propName, _
        LinkToContent:=False, _
        Type:=msoPropertyTypeString, _
        Value:=CStr(valueText)
End Sub
