VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiPanelVisibility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiBindingListener

Private m_isInitialized As Boolean
Private m_isDisposed As Boolean

Private m_context As obj_UiRenderContext
Private m_expression As String
Private m_bindingSource As String
Private m_bindingPath As String
Private m_collapsed As Boolean

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
Private Sub obj_IUiBindingListener_OnValueChanged( _
    ByVal sourceName As String, _
    ByVal bindingPath As String _
)
    Dim collapsed As Boolean
    Dim diagnostic As String

    If Not m_isInitialized Or m_isDisposed Then
        Exit Sub
    End If
    If StrComp(sourceName, m_bindingSource, vbTextCompare) <> 0 Then
        Exit Sub
    End If
    If Len(bindingPath) > 0 Then
        If StrComp(bindingPath, m_bindingPath, vbTextCompare) <> 0 And _
                StrComp(Left$(m_bindingPath, Len(bindingPath) + 1), bindingPath & ".", vbTextCompare) <> 0 Then
            Exit Sub
        End If
    End If
    If Not private_TryResolve(collapsed, diagnostic) Then
        Err.Raise 5, , diagnostic
    End If
    If collapsed <> m_collapsed Then
        m_context.InvalidateMeasure
    End If
End Sub

' //
' // API
' //
Public Function Initialize( _
    ByVal definition As Object, _
    ByVal context As obj_UiRenderContext, _
    ByVal source As String, _
    ByRef diagnostic As String _
) As Boolean
    On Error GoTo EH
    If m_isInitialized Or m_isDisposed Then
        diagnostic = "Visibility state is already initialized or disposed."
        Exit Function
    End If
    Set m_context = context
    m_expression = ex_UiElementFactory.fn_Attribute(definition, "visibility")
    If StrComp(Left$(Trim$(m_expression), 9), "{Binding ", vbTextCompare) = 0 Then
        If Not ex_UiBindingRuntime.fn_TryParseBinding( _
                m_expression, source, m_bindingSource, m_bindingPath, context.BindingContext) Then
            diagnostic = "Cannot parse container visibility binding: " & m_expression
            Exit Function
        End If
        m_expression = "{Binding Source=" & m_bindingSource & "; Path=" & m_bindingPath & "}"
    End If
    If Not private_TryResolve(m_collapsed, diagnostic) Then
        Exit Function
    End If
    If Len(m_bindingSource) > 0 Then
        m_context.BindingContext.Subscribe Me
    End If
    m_isInitialized = True
    Initialize = True
    Exit Function
EH:
    diagnostic = Err.Description
    Me.Dispose
End Function

Public Sub Dispose()
    Dim bindings As obj_UiBindingContext

    If m_isDisposed Then
        Exit Sub
    End If
    m_isInitialized = False
    m_isDisposed = True
    If Not m_context Is Nothing Then
        Set bindings = m_context.BindingContext
        If Not bindings Is Nothing Then
            bindings.Unsubscribe Me
        End If
    End If
    Set m_context = Nothing
End Sub

Public Function TryMeasure( _
    ByRef collapsed As Boolean, _
    ByRef diagnostic As String _
) As Boolean
    If Not m_isInitialized Or m_isDisposed Then
        diagnostic = "Container visibility state is unavailable."
        Exit Function
    End If
    If Not private_TryResolve(collapsed, diagnostic) Then
        Exit Function
    End If
    m_collapsed = collapsed
    TryMeasure = True
End Function

' //
' // Private
' //
Private Function private_TryResolve( _
    ByRef collapsed As Boolean, _
    ByRef diagnostic As String _
) As Boolean
    Dim value As String

    collapsed = False
    If Len(m_expression) = 0 Then
        private_TryResolve = True
        Exit Function
    End If
    If Not ex_UiBindingRuntime.fn_TryResolveText(m_expression, m_context.BindingContext, value) Then
        diagnostic = "Cannot resolve container visibility."
        Exit Function
    End If
    Select Case value
        Case "Visible"
            collapsed = False
        Case "Collapsed"
            collapsed = True
        Case Else
            diagnostic = "Invalid container visibility: " & value
            Exit Function
    End Select
    private_TryResolve = True
End Function