VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_UiCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_callbackName As String

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Initialize(ByVal callbackName As String) As Boolean
    m_callbackName = VBA.Trim$(callbackName)
    If VBA.Len(m_callbackName) = 0 Then
        VBA.MsgBox "A command callback is required.", VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    fn_Initialize = True
End Function

Public Function fn_Execute() As Boolean
    On Error GoTo EH
    Application.Run "'" & VBA.Replace$(ThisWorkbook.Name, "'", "''") & "'!" & m_callbackName
    fn_Execute = True
    Exit Function
EH:
    VBA.MsgBox "The command cannot be executed: " & m_callbackName & _
        " | " & VBA.Err.Description, VBA.vbExclamation, "PersonalEventBuilder"
End Function

Public Property Get fn_CallbackName() As String
    fn_CallbackName = m_callbackName
End Property

Public Sub fn_Dispose()
    m_callbackName = VBA.vbNullString
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------