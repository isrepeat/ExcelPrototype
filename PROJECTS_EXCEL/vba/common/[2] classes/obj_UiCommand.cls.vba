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
Private m_isDisposed As Boolean

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Properties
' //
Public Property Get CallbackName() As String
    CallbackName = VBA.TypeName(m_target) & "." & m_methodName
End Property

' //
' // API
' //
Public Function Initialize(ByVal target As Object, ByVal methodName As String) As Boolean
    m_isDisposed = False
    Set m_target = target
    m_methodName = VBA.Trim$(methodName)
    If m_target Is Nothing Or VBA.Len(m_methodName) = 0 Then
        ex_WindowsUi.fn_ShowMessage "A command target and method are required.", VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    Initialize = True
End Function

Public Sub Dispose()
    If m_isDisposed Then Exit Sub
    m_isDisposed = True
    Set m_target = Nothing
    m_methodName = VBA.vbNullString
End Sub

Public Function Execute() As Boolean
    Dim callbackResult As Variant

    On Error GoTo EH
    callbackResult = VBA.CallByName(m_target, m_methodName, VbMethod)
    Execute = CBool(callbackResult)
    Exit Function
EH:
    ex_WindowsUi.fn_ShowMessage "The command cannot be executed: " & m_methodName & _
        " | " & VBA.Err.Description, VBA.vbExclamation, "PersonalEventBuilder"
End Function