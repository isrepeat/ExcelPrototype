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

Implements obj_IUiEventHandler

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_target As Object
Private m_methodName As String
Private m_payload As Object

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // Interface
' //
Private Function obj_IUiEventHandler_HandleEvent( _
    ByVal kind As String, _
    ByVal payload As Variant _
) As Boolean
    If kind = "click" Or kind = "selection" Then obj_IUiEventHandler_HandleEvent = Me.Execute()
End Function

' //
' // Properties
' //
Public Property Get CallbackName() As String
    CallbackName = VBA.TypeName(m_target) & "." & m_methodName
End Property

' //
' // API
' //
Public Function Initialize( _
    ByVal target As Object, _
    ByVal methodName As String, _
    Optional ByVal payload As Object = Nothing _
) As Boolean
    If m_isDisposed Or m_isInitialized Then
        Exit Function
    End If
    Set m_target = target
    Set m_payload = payload
    m_methodName = VBA.Trim$(methodName)
    If m_target Is Nothing Or VBA.Len(m_methodName) = 0 Then
        ex_WindowsUi.fn_ShowMessage "A command target and method are required.", VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    Initialize = True
    m_isInitialized = Initialize
End Function

Public Sub Dispose()
    If m_isDisposed Then
        Exit Sub
    End If
    m_isDisposed = True
    m_isInitialized = False
    Set m_target = Nothing
    Set m_payload = Nothing
    m_methodName = VBA.vbNullString
End Sub

Public Function ExecuteWithPayload(ByVal payload As Object) As Boolean
    Dim callbackResult As Variant

    On Error GoTo EH
    If Not m_isInitialized Or m_isDisposed Then Exit Function
    If m_payload Is Nothing Then
        callbackResult = VBA.CallByName(m_target, m_methodName, VbMethod, payload)
    Else
        callbackResult = VBA.CallByName(m_target, m_methodName, VbMethod, m_payload, payload)
    End If
    ExecuteWithPayload = VBA.CBool(callbackResult)
    Exit Function
EH:
    ex_WindowsUi.fn_ShowMessage "The command cannot be executed: " & m_methodName & _
        " | " & VBA.Err.Description, VBA.vbExclamation, "Command"
End Function

Public Function Execute() As Boolean
    Dim callbackResult As Variant

    On Error GoTo EH
    If m_payload Is Nothing Then
        callbackResult = VBA.CallByName(m_target, m_methodName, VbMethod)
    Else
        callbackResult = VBA.CallByName(m_target, m_methodName, VbMethod, m_payload)
    End If
    Execute = VBA.CBool(callbackResult)
    Exit Function
EH:
    ex_WindowsUi.fn_ShowMessage "The command cannot be executed: " & m_methodName & _
        " | " & VBA.Err.Description, VBA.vbExclamation, "PersonalEventBuilder"
End Function