VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "obj_PEB_PgMainController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pageBase As obj_PageBase

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

' //
' // API
' //
Public Function Initialize(ByVal pageBase As obj_PageBase) As Boolean
    Set m_pageBase = pageBase
    If m_pageBase Is Nothing Then
        VBA.MsgBox "The PersonalEventBuilder page controller requires a page base.", _
            VBA.vbExclamation, "PersonalEventBuilder"
        Exit Function
    End If
    Initialize = True
End Function

Public Sub Dispose()
    Set m_pageBase = Nothing
End Sub

Public Function HelloWorld() As Boolean
    ex_Core.fn_Diagnostic_WriteLog "HELLO_WORLD_CLICKED"
    VBA.MsgBox "Hello World from PersonalEventBuilder.", VBA.vbInformation, "PersonalEventBuilder"
    HelloWorld = True
End Function

Public Function UpdatePage() As Boolean
    If m_pageBase Is Nothing Then Exit Function
    UpdatePage = m_pageBase.UpdatePage()
End Function