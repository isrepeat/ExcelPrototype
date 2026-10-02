VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiElementFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiElementFactory

Private m_kind As String

' //
' // Lifecycle
' //
Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //
Public Sub Dispose()
    m_kind = VBA.vbNullString
End Sub

Public Sub Initialize(ByVal kind As String)
    m_kind = kind
End Sub

Public Function Create( _
    ByVal definition As Object, _
    ByVal context As obj_UiRenderContext, _
    ByVal source As String, _
    ByRef diagnostic As String _
) As obj_IUiElement
    Dim element As obj_IUiElement
    Dim panel As obj_UiPanel

    On Error GoTo EH_CONFIGURE
    Select Case m_kind
        Case "page", "grid"
            Set panel = New obj_UiPanel
            panel.Initialize "grid"
            Set element = panel
        Case "stackpanel"
            Set panel = New obj_UiPanel
            panel.Initialize "stack"
            Set element = panel
        Case "form"
            Set element = New obj_UiForm
        Case "control"
            Set element = New obj_UiControlElement
    End Select
    If element Is Nothing Then Exit Function
    If Not element.Configure(definition, context, source, diagnostic) Then
        element.Dispose
        Exit Function
    End If
    If m_kind <> "control" Then context.RegisterElement ex_UiElementFactory.fn_Attribute(definition, "name"), element
    Set Create = element
    Exit Function
EH_CONFIGURE:
    diagnostic = VBA.Err.Description
    If Not element Is Nothing Then element.Dispose
End Function

' //
' // Interface
' //
Private Function obj_IUiElementFactory_Create( _
    ByVal definition As Object, _
    ByVal context As obj_UiRenderContext, _
    ByVal source As String, _
    ByRef diagnostic As String _
) As obj_IUiElement
    Set obj_IUiElementFactory_Create = Me.Create(definition, context, source, diagnostic)
End Function