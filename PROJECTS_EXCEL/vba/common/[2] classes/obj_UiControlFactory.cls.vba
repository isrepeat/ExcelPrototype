VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "obj_UiControlFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements obj_IUiControlFactory

Private m_type As String

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
    m_type = VBA.vbNullString
End Sub

Public Sub Initialize(ByVal controlType As String)
    m_type = VBA.LCase$(controlType)
End Sub

' //
' // Interface
' //
Private Function obj_IUiControlFactory_Create() As obj_IUiControl
    Dim control As obj_IUiControl

    Select Case m_type
        Case "label"
            Set control = New obj_UiLabelControl
        Case "button"
            Set control = New obj_UiButtonControl
        Case "table"
            Set control = New obj_UiTableControl
        Case "input", "select"
            Set control = New obj_UiFieldControl
    End Select
    If control Is Nothing Then Exit Function
    If Not control.Initialize() Then Exit Function
    Set obj_IUiControlFactory_Create = control
End Function