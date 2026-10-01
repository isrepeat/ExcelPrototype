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

Private m_target As Object
Private m_methodName As String

' --------------------------------------
' namespace API {
' --------------------------------------
Public Function fn_Initialize(ByVal target As Object, ByVal methodName As String) As Boolean
    Set m_target = target
    m_methodName = VBA.Trim$(methodName)
    If m_target Is Nothing Or VBA.Len(m_methodName) = 0 Then
        VBA.MsgBox "A command target and method are required.", VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    fn_Initialize = True
End Function

Public Function fn_Execute() As Boolean
    Dim callbackResult As Variant

    On Error GoTo EH
    callbackResult = VBA.CallByName(m_target, m_methodName, VbMethod)
    fn_Execute = CBool(callbackResult)
    Exit Function
EH:
    VBA.MsgBox "The command cannot be executed: " & m_methodName & _
        " | " & VBA.Err.Description, VBA.vbExclamation, "PersonalEventBuilder"
End Function

Public Property Get fn_CallbackName() As String
    fn_CallbackName = VBA.TypeName(m_target) & "." & m_methodName
End Property

Public Sub fn_Dispose()
    Set m_target = Nothing
    m_methodName = VBA.vbNullString
End Sub
' --------------------------------------
' } // namespace API
' --------------------------------------